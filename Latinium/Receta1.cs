using System;
using System.Data;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de Receta1.
	/// </summary>
	public class Receta1 : System.Windows.Forms.Form
	{
		private C1.Data.C1DataSet cdsProducto;
		private C1.Data.C1DataView cdvProducto;
		private Infragistics.Win.UltraWinGrid.UltraGrid ultraGrid1;
		private C1.Data.C1DataView cdvArtCodigo;
		private C1.Data.C1DataSet cdsArticulo;
		private C1.Data.C1DataView cdvArticulo;
		private System.Windows.Forms.Label label1;
		private DevExpress.XtraEditors.LookUpEdit cmbArtCodigo;
		private DevExpress.XtraEditors.TextEdit txtIdProducto;
		private DevExpress.XtraEditors.LookUpEdit cdmArticulo;
		private BarraDatos.BarraDato barraDato1;
		private Infragistics.Win.UltraWinGrid.UltraDropDown ultraDropDown1;
		private Infragistics.Win.UltraWinGrid.UltraDropDown ultraDropDown2;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public Receta1()
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
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("_ProductoDetalle", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Borrar");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cantidad");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Costo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idArticulo", -1, "ultraDropDown1");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idProdDetalle");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idProducto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Impuesto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Precio");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Importante");
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(Receta1));
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Articulo", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Alto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn11 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Ancho");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn12 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Aplicacion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn13 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bodega1");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn14 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bodega2");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn15 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bodega3");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn16 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bodega4");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn17 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bodega5");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn18 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bodega6");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn19 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bodega7");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn20 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bodega8");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn21 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bodega9");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn22 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Borrar");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn23 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CantidadMaxima");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn24 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CantidadMinima");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn25 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CantidadPromocion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn26 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CantidadRequerida");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn27 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CantReqDesc");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn28 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CantReqPorc");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn29 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Caracteristica");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn30 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CodAlterno");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn31 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn32 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Color");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn33 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Comision120");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn34 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Comision30");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn35 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Comision60");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn36 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Comision90");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn37 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ComisionValor");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn38 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CostoMaximo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn39 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CostoPromedio");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn40 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CostoUltimo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn41 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CostoUltimo1");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn42 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CostoUltimo2");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn43 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CostoUltimo3");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn44 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Depreciacion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn45 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Descontinuado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn46 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("DescPorcContado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn47 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("DescPorcCre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn48 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("DescuentoContado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn49 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("DescuentoCredito");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn50 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("DistUltimo1");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn51 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("DistUltimo2");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn52 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("DistUltimo3");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn53 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaUltima1");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn54 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaUltima2");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn55 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaUltima3");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn56 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Gramos");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn57 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Ice");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn58 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idArancel");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn59 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idArticulo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn60 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idGrupoArticulo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn61 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idMarca");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn62 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPromocion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn63 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idSubGrupo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn64 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idUnidad");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn65 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Incr1");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn66 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Incr2");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn67 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Incr3");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn68 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Incr4");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn69 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Incr5");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn70 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Incr6");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn71 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Iva");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn72 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Largo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn73 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Notas");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn74 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("NumeroParte");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn75 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Peso");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn76 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Precio1");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn77 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Precio2");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn78 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Precio3");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn79 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Precio4");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn80 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Precio5");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn81 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Precio6");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn82 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("PrecioPromocion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn83 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Residual");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn84 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Responsable");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn85 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Saldo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn86 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("SaldoT");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn87 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("TipoPromocion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn88 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Ubicacion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn89 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("UnidadPeso");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn90 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Articulo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn91 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("BodegaPredef");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn92 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Habilitado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn93 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ValorPeso");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn94 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idClase");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn95 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idSeccion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn96 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ComisionPorcentaje");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn97 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Tiempo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn98 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idTipoGrupo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn99 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ConsumoProm");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn100 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("TiempoReposicion");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand3 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Articulo", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn101 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Alto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn102 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Ancho");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn103 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Aplicacion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn104 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bodega1");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn105 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bodega2");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn106 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bodega3");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn107 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bodega4");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn108 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bodega5");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn109 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bodega6");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn110 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bodega7");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn111 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bodega8");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn112 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bodega9");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn113 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Borrar");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn114 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CantidadMaxima");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn115 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CantidadMinima");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn116 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CantidadPromocion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn117 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CantidadRequerida");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn118 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CantReqDesc");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn119 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CantReqPorc");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn120 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Caracteristica");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn121 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CodAlterno");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn122 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn123 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Color");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn124 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Comision120");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn125 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Comision30");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn126 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Comision60");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn127 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Comision90");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn128 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ComisionValor");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn129 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CostoMaximo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn130 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CostoPromedio");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn131 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CostoUltimo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn132 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CostoUltimo1");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn133 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CostoUltimo2");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn134 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CostoUltimo3");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn135 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Depreciacion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn136 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Descontinuado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn137 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("DescPorcContado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn138 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("DescPorcCre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn139 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("DescuentoContado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn140 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("DescuentoCredito");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn141 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("DistUltimo1");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn142 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("DistUltimo2");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn143 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("DistUltimo3");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn144 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaUltima1");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn145 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaUltima2");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn146 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaUltima3");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn147 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Gramos");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn148 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Ice");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn149 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idArancel");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn150 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idArticulo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn151 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idGrupoArticulo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn152 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idMarca");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn153 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPromocion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn154 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idSubGrupo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn155 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idUnidad");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn156 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Incr1");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn157 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Incr2");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn158 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Incr3");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn159 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Incr4");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn160 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Incr5");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn161 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Incr6");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn162 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Iva");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn163 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Largo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn164 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Notas");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn165 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("NumeroParte");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn166 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Peso");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn167 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Precio1");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn168 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Precio2");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn169 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Precio3");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn170 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Precio4");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn171 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Precio5");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn172 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Precio6");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn173 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("PrecioPromocion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn174 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Residual");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn175 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Responsable");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn176 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Saldo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn177 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("SaldoT");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn178 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("TipoPromocion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn179 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Ubicacion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn180 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("UnidadPeso");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn181 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Articulo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn182 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("BodegaPredef");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn183 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Habilitado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn184 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ValorPeso");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn185 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idClase");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn186 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idSeccion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn187 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ComisionPorcentaje");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn188 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Tiempo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn189 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idTipoGrupo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn190 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ConsumoProm");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn191 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("TiempoReposicion");
			this.cdsProducto = new C1.Data.C1DataSet();
			this.cdvProducto = new C1.Data.C1DataView();
			this.ultraGrid1 = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.cdvArtCodigo = new C1.Data.C1DataView();
			this.cdsArticulo = new C1.Data.C1DataSet();
			this.cdvArticulo = new C1.Data.C1DataView();
			this.label1 = new System.Windows.Forms.Label();
			this.cmbArtCodigo = new DevExpress.XtraEditors.LookUpEdit();
			this.txtIdProducto = new DevExpress.XtraEditors.TextEdit();
			this.cdmArticulo = new DevExpress.XtraEditors.LookUpEdit();
			this.barraDato1 = new BarraDatos.BarraDato();
			this.ultraDropDown1 = new Infragistics.Win.UltraWinGrid.UltraDropDown();
			this.ultraDropDown2 = new Infragistics.Win.UltraWinGrid.UltraDropDown();
			((System.ComponentModel.ISupportInitialize)(this.cdsProducto)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvProducto)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvArtCodigo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsArticulo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvArticulo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbArtCodigo.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdProducto.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdmArticulo.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDropDown1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDropDown2)).BeginInit();
			this.SuspendLayout();
			// 
			// cdsProducto
			// 
			this.cdsProducto.BindingContextCtrl = this;
			this.cdsProducto.DataLibrary = "LibFacturacion";
			this.cdsProducto.DataLibraryUrl = "";
			this.cdsProducto.DataSetDef = "dsProducto";
			this.cdsProducto.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsProducto.Name = "cdsProducto";
			this.cdsProducto.SchemaClassName = "LibFacturacion.DataClass";
			this.cdsProducto.SchemaDef = null;
			this.cdsProducto.BeforeFill += new C1.Data.FillEventHandler(this.cdsProducto_BeforeFill);
			// 
			// cdvProducto
			// 
			this.cdvProducto.BindingContextCtrl = this;
			this.cdvProducto.DataSet = this.cdsProducto;
			this.cdvProducto.Sort = "idProducto";
			this.cdvProducto.TableName = "";
			this.cdvProducto.TableViewName = "Producto";
			// 
			// ultraGrid1
			// 
			this.ultraGrid1.Cursor = System.Windows.Forms.Cursors.Default;
			this.ultraGrid1.DataMember = "_Producto.Producto - ProductoDetalle";
			this.ultraGrid1.DataSource = this.cdsProducto;
			appearance1.BackColor = System.Drawing.Color.White;
			this.ultraGrid1.DisplayLayout.Appearance = appearance1;
			ultraGridBand1.AddButtonCaption = "_Producto";
			ultraGridColumn1.Header.VisiblePosition = 3;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn2.Header.VisiblePosition = 0;
			ultraGridColumn3.Header.VisiblePosition = 4;
			ultraGridColumn3.Hidden = true;
			ultraGridColumn4.Header.Caption = "Artículo";
			ultraGridColumn4.Header.VisiblePosition = 1;
			ultraGridColumn4.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownList;
			ultraGridColumn4.Width = 270;
			ultraGridColumn5.Header.VisiblePosition = 6;
			ultraGridColumn5.Hidden = true;
			ultraGridColumn6.Hidden = true;
			ultraGridColumn7.Header.VisiblePosition = 7;
			ultraGridColumn7.Hidden = true;
			ultraGridColumn8.Header.VisiblePosition = 8;
			ultraGridColumn8.Hidden = true;
			ultraGridColumn9.Header.VisiblePosition = 2;
			ultraGridColumn9.Width = 87;
			ultraGridBand1.Columns.Add(ultraGridColumn1);
			ultraGridBand1.Columns.Add(ultraGridColumn2);
			ultraGridBand1.Columns.Add(ultraGridColumn3);
			ultraGridBand1.Columns.Add(ultraGridColumn4);
			ultraGridBand1.Columns.Add(ultraGridColumn5);
			ultraGridBand1.Columns.Add(ultraGridColumn6);
			ultraGridBand1.Columns.Add(ultraGridColumn7);
			ultraGridBand1.Columns.Add(ultraGridColumn8);
			ultraGridBand1.Columns.Add(ultraGridColumn9);
			this.ultraGrid1.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.ultraGrid1.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.TemplateOnBottom;
			this.ultraGrid1.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.True;
			this.ultraGrid1.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
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
			this.ultraGrid1.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.ultraGrid1.Location = new System.Drawing.Point(16, 48);
			this.ultraGrid1.Name = "ultraGrid1";
			this.ultraGrid1.Size = new System.Drawing.Size(512, 232);
			this.ultraGrid1.TabIndex = 1;
			// 
			// cdvArtCodigo
			// 
			this.cdvArtCodigo.BindingContextCtrl = this;
			this.cdvArtCodigo.DataSet = this.cdsArticulo;
			this.cdvArtCodigo.Sort = "Codigo";
			this.cdvArtCodigo.TableName = "";
			this.cdvArtCodigo.TableViewName = "Articulo";
			// 
			// cdsArticulo
			// 
			this.cdsArticulo.BindingContextCtrl = this;
			this.cdsArticulo.DataLibrary = "LibFacturacion";
			this.cdsArticulo.DataLibraryUrl = "";
			this.cdsArticulo.DataSetDef = "dsArticulo";
			this.cdsArticulo.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsArticulo.Name = "cdsArticulo";
			this.cdsArticulo.SchemaClassName = "LibFacturacion.DataClass";
			this.cdsArticulo.SchemaDef = null;
			this.cdsArticulo.BeforeFill += new C1.Data.FillEventHandler(this.cdsArticulo_BeforeFill);
			// 
			// cdvArticulo
			// 
			this.cdvArticulo.BindingContextCtrl = this;
			this.cdvArticulo.DataSet = this.cdsArticulo;
			this.cdvArticulo.Sort = "Articulo";
			this.cdvArticulo.TableName = "";
			this.cdvArticulo.TableViewName = "Articulo";
			// 
			// label1
			// 
			this.label1.AutoSize = true;
			this.label1.Location = new System.Drawing.Point(16, 16);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(53, 13);
			this.label1.TabIndex = 9;
			this.label1.Text = "Producto:";
			// 
			// cmbArtCodigo
			// 
			this.cmbArtCodigo.DataBindings.Add(new System.Windows.Forms.Binding("EditValue", this.cdsProducto, "_Producto.idArticulo"));
			this.cmbArtCodigo.EditValue = ((System.DBNull)(resources.GetObject("cmbArtCodigo.EditValue")));
			this.cmbArtCodigo.Location = new System.Drawing.Point(304, 16);
			this.cmbArtCodigo.Name = "cmbArtCodigo";
			// 
			// cmbArtCodigo.Properties
			// 
			this.cmbArtCodigo.Properties.Buttons.AddRange(new DevExpress.XtraEditors.Controls.EditorButton[] {
																																																				 new DevExpress.XtraEditors.Controls.EditorButton(DevExpress.XtraEditors.Controls.ButtonPredefines.Combo)});
			this.cmbArtCodigo.Properties.Columns.AddRange(new DevExpress.XtraEditors.Controls.LookUpColumnInfo[] {
																																																						 new DevExpress.XtraEditors.Controls.LookUpColumnInfo("Codigo", "Código", 20, DevExpress.Utils.FormatType.None, "", true, DevExpress.Utils.HorzAlignment.Default),
																																																						 new DevExpress.XtraEditors.Controls.LookUpColumnInfo("Articulo", "Articulo", 47, DevExpress.Utils.FormatType.None, "", true, DevExpress.Utils.HorzAlignment.Near)});
			this.cmbArtCodigo.Properties.DataSource = this.cdvArtCodigo;
			this.cmbArtCodigo.Properties.DisplayMember = "Codigo";
			this.cmbArtCodigo.Properties.NullText = "";
			this.cmbArtCodigo.Properties.ValueMember = "idArticulo";
			this.cmbArtCodigo.Size = new System.Drawing.Size(160, 19);
			this.cmbArtCodigo.TabIndex = 8;
			// 
			// txtIdProducto
			// 
			this.txtIdProducto.DataBindings.Add(new System.Windows.Forms.Binding("EditValue", this.cdsProducto, "_Producto.idProducto"));
			this.txtIdProducto.EditValue = ((System.DBNull)(resources.GetObject("txtIdProducto.EditValue")));
			this.txtIdProducto.Location = new System.Drawing.Point(464, 16);
			this.txtIdProducto.Name = "txtIdProducto";
			this.txtIdProducto.Size = new System.Drawing.Size(56, 19);
			this.txtIdProducto.TabIndex = 7;
			// 
			// cdmArticulo
			// 
			this.cdmArticulo.DataBindings.Add(new System.Windows.Forms.Binding("EditValue", this.cdsProducto, "_Producto.idArticulo"));
			this.cdmArticulo.EditValue = ((System.DBNull)(resources.GetObject("cdmArticulo.EditValue")));
			this.cdmArticulo.Location = new System.Drawing.Point(96, 16);
			this.cdmArticulo.Name = "cdmArticulo";
			// 
			// cdmArticulo.Properties
			// 
			this.cdmArticulo.Properties.Buttons.AddRange(new DevExpress.XtraEditors.Controls.EditorButton[] {
																																																				new DevExpress.XtraEditors.Controls.EditorButton(DevExpress.XtraEditors.Controls.ButtonPredefines.Combo)});
			this.cdmArticulo.Properties.Columns.AddRange(new DevExpress.XtraEditors.Controls.LookUpColumnInfo[] {
																																																						new DevExpress.XtraEditors.Controls.LookUpColumnInfo("Articulo", "Articulo", 47, DevExpress.Utils.FormatType.None, "", true, DevExpress.Utils.HorzAlignment.Near),
																																																						new DevExpress.XtraEditors.Controls.LookUpColumnInfo("Codigo", "Código", 20, DevExpress.Utils.FormatType.None, "", true, DevExpress.Utils.HorzAlignment.Default)});
			this.cdmArticulo.Properties.DataSource = this.cdvArticulo;
			this.cdmArticulo.Properties.DisplayMember = "Articulo";
			this.cdmArticulo.Properties.NullText = "";
			this.cdmArticulo.Properties.PopupWidth = 400;
			this.cdmArticulo.Properties.ValueMember = "idArticulo";
			this.cdmArticulo.Size = new System.Drawing.Size(192, 19);
			this.cdmArticulo.TabIndex = 6;
			// 
			// barraDato1
			// 
			this.barraDato1.BarraMovimiento = true;
			this.barraDato1.DataMember = "_Producto";
			this.barraDato1.DataSource = this.cdsProducto;
			this.barraDato1.Location = new System.Drawing.Point(16, 288);
			this.barraDato1.Name = "barraDato1";
			this.barraDato1.Size = new System.Drawing.Size(304, 24);
			this.barraDato1.TabIndex = 10;
			this.barraDato1.VisibleBorrar = true;
			this.barraDato1.VisibleBuscar = true;
			this.barraDato1.VisibleEditar = true;
			this.barraDato1.VisibleImprimir = false;
			this.barraDato1.VisibleNuevo = true;
			this.barraDato1.Buscar += new System.EventHandler(this.barraDato1_Buscar);
			this.barraDato1.Borrar += new System.EventHandler(this.barraDato1_Borrar);
			this.barraDato1.Crear += new System.EventHandler(this.barraDato1_Crear);
			this.barraDato1.Deshacer += new System.EventHandler(this.barraDato1_Deshacer);
			this.barraDato1.Editar += new System.EventHandler(this.barraDato1_Editar);
			this.barraDato1.Grabar += new System.EventHandler(this.barraDato1_Grabar);
			// 
			// ultraDropDown1
			// 
			this.ultraDropDown1.Cursor = System.Windows.Forms.Cursors.Default;
			this.ultraDropDown1.DataSource = this.cdvArticulo;
			ultraGridColumn10.Header.VisiblePosition = 2;
			ultraGridColumn10.Hidden = true;
			ultraGridColumn11.Header.VisiblePosition = 3;
			ultraGridColumn11.Hidden = true;
			ultraGridColumn12.Header.VisiblePosition = 4;
			ultraGridColumn12.Hidden = true;
			ultraGridColumn13.Header.VisiblePosition = 5;
			ultraGridColumn13.Hidden = true;
			ultraGridColumn14.Header.VisiblePosition = 6;
			ultraGridColumn14.Hidden = true;
			ultraGridColumn15.Header.VisiblePosition = 7;
			ultraGridColumn15.Hidden = true;
			ultraGridColumn16.Header.VisiblePosition = 8;
			ultraGridColumn16.Hidden = true;
			ultraGridColumn17.Header.VisiblePosition = 9;
			ultraGridColumn17.Hidden = true;
			ultraGridColumn18.Header.VisiblePosition = 10;
			ultraGridColumn18.Hidden = true;
			ultraGridColumn19.Header.VisiblePosition = 11;
			ultraGridColumn19.Hidden = true;
			ultraGridColumn20.Header.VisiblePosition = 12;
			ultraGridColumn20.Hidden = true;
			ultraGridColumn21.Header.VisiblePosition = 13;
			ultraGridColumn21.Hidden = true;
			ultraGridColumn22.Header.VisiblePosition = 14;
			ultraGridColumn22.Hidden = true;
			ultraGridColumn23.Header.VisiblePosition = 15;
			ultraGridColumn23.Hidden = true;
			ultraGridColumn24.Header.VisiblePosition = 16;
			ultraGridColumn24.Hidden = true;
			ultraGridColumn25.Header.VisiblePosition = 17;
			ultraGridColumn25.Hidden = true;
			ultraGridColumn26.Header.VisiblePosition = 18;
			ultraGridColumn26.Hidden = true;
			ultraGridColumn27.Header.VisiblePosition = 19;
			ultraGridColumn27.Hidden = true;
			ultraGridColumn28.Header.VisiblePosition = 20;
			ultraGridColumn28.Hidden = true;
			ultraGridColumn29.Header.VisiblePosition = 21;
			ultraGridColumn29.Hidden = true;
			ultraGridColumn30.Header.VisiblePosition = 22;
			ultraGridColumn30.Hidden = true;
			ultraGridColumn31.Header.VisiblePosition = 1;
			ultraGridColumn32.Header.VisiblePosition = 23;
			ultraGridColumn32.Hidden = true;
			ultraGridColumn33.Header.VisiblePosition = 24;
			ultraGridColumn33.Hidden = true;
			ultraGridColumn34.Header.VisiblePosition = 25;
			ultraGridColumn34.Hidden = true;
			ultraGridColumn35.Header.VisiblePosition = 26;
			ultraGridColumn35.Hidden = true;
			ultraGridColumn36.Header.VisiblePosition = 27;
			ultraGridColumn36.Hidden = true;
			ultraGridColumn37.Header.VisiblePosition = 28;
			ultraGridColumn37.Hidden = true;
			ultraGridColumn38.Header.VisiblePosition = 29;
			ultraGridColumn38.Hidden = true;
			ultraGridColumn39.Header.VisiblePosition = 30;
			ultraGridColumn39.Hidden = true;
			ultraGridColumn40.Header.VisiblePosition = 31;
			ultraGridColumn40.Hidden = true;
			ultraGridColumn41.Header.VisiblePosition = 32;
			ultraGridColumn41.Hidden = true;
			ultraGridColumn42.Header.VisiblePosition = 33;
			ultraGridColumn42.Hidden = true;
			ultraGridColumn43.Header.VisiblePosition = 34;
			ultraGridColumn43.Hidden = true;
			ultraGridColumn44.Header.VisiblePosition = 35;
			ultraGridColumn44.Hidden = true;
			ultraGridColumn45.Header.VisiblePosition = 36;
			ultraGridColumn45.Hidden = true;
			ultraGridColumn46.Header.VisiblePosition = 37;
			ultraGridColumn46.Hidden = true;
			ultraGridColumn47.Header.VisiblePosition = 38;
			ultraGridColumn47.Hidden = true;
			ultraGridColumn48.Header.VisiblePosition = 39;
			ultraGridColumn48.Hidden = true;
			ultraGridColumn49.Header.VisiblePosition = 40;
			ultraGridColumn49.Hidden = true;
			ultraGridColumn50.Header.VisiblePosition = 41;
			ultraGridColumn50.Hidden = true;
			ultraGridColumn51.Header.VisiblePosition = 42;
			ultraGridColumn51.Hidden = true;
			ultraGridColumn52.Header.VisiblePosition = 43;
			ultraGridColumn52.Hidden = true;
			ultraGridColumn53.Header.VisiblePosition = 44;
			ultraGridColumn53.Hidden = true;
			ultraGridColumn54.Header.VisiblePosition = 45;
			ultraGridColumn54.Hidden = true;
			ultraGridColumn55.Header.VisiblePosition = 46;
			ultraGridColumn55.Hidden = true;
			ultraGridColumn56.Header.VisiblePosition = 47;
			ultraGridColumn56.Hidden = true;
			ultraGridColumn57.Header.VisiblePosition = 48;
			ultraGridColumn57.Hidden = true;
			ultraGridColumn58.Header.VisiblePosition = 49;
			ultraGridColumn58.Hidden = true;
			ultraGridColumn59.Header.VisiblePosition = 50;
			ultraGridColumn59.Hidden = true;
			ultraGridColumn60.Header.VisiblePosition = 51;
			ultraGridColumn60.Hidden = true;
			ultraGridColumn61.Header.VisiblePosition = 52;
			ultraGridColumn61.Hidden = true;
			ultraGridColumn62.Header.VisiblePosition = 53;
			ultraGridColumn62.Hidden = true;
			ultraGridColumn63.Header.VisiblePosition = 54;
			ultraGridColumn63.Hidden = true;
			ultraGridColumn64.Header.VisiblePosition = 55;
			ultraGridColumn64.Hidden = true;
			ultraGridColumn65.Header.VisiblePosition = 56;
			ultraGridColumn65.Hidden = true;
			ultraGridColumn66.Header.VisiblePosition = 57;
			ultraGridColumn66.Hidden = true;
			ultraGridColumn67.Header.VisiblePosition = 58;
			ultraGridColumn67.Hidden = true;
			ultraGridColumn68.Header.VisiblePosition = 59;
			ultraGridColumn68.Hidden = true;
			ultraGridColumn69.Header.VisiblePosition = 60;
			ultraGridColumn69.Hidden = true;
			ultraGridColumn70.Header.VisiblePosition = 61;
			ultraGridColumn70.Hidden = true;
			ultraGridColumn71.Header.VisiblePosition = 62;
			ultraGridColumn71.Hidden = true;
			ultraGridColumn72.Header.VisiblePosition = 63;
			ultraGridColumn72.Hidden = true;
			ultraGridColumn73.Header.VisiblePosition = 64;
			ultraGridColumn73.Hidden = true;
			ultraGridColumn74.Header.VisiblePosition = 65;
			ultraGridColumn74.Hidden = true;
			ultraGridColumn75.Header.VisiblePosition = 66;
			ultraGridColumn75.Hidden = true;
			ultraGridColumn76.Header.VisiblePosition = 67;
			ultraGridColumn76.Hidden = true;
			ultraGridColumn77.Header.VisiblePosition = 68;
			ultraGridColumn77.Hidden = true;
			ultraGridColumn78.Header.VisiblePosition = 69;
			ultraGridColumn78.Hidden = true;
			ultraGridColumn79.Header.VisiblePosition = 70;
			ultraGridColumn79.Hidden = true;
			ultraGridColumn80.Header.VisiblePosition = 71;
			ultraGridColumn80.Hidden = true;
			ultraGridColumn81.Header.VisiblePosition = 72;
			ultraGridColumn81.Hidden = true;
			ultraGridColumn82.Header.VisiblePosition = 73;
			ultraGridColumn82.Hidden = true;
			ultraGridColumn83.Header.VisiblePosition = 74;
			ultraGridColumn83.Hidden = true;
			ultraGridColumn84.Header.VisiblePosition = 75;
			ultraGridColumn84.Hidden = true;
			ultraGridColumn85.Header.VisiblePosition = 76;
			ultraGridColumn85.Hidden = true;
			ultraGridColumn86.Header.VisiblePosition = 77;
			ultraGridColumn86.Hidden = true;
			ultraGridColumn87.Header.VisiblePosition = 78;
			ultraGridColumn87.Hidden = true;
			ultraGridColumn88.Header.VisiblePosition = 79;
			ultraGridColumn88.Hidden = true;
			ultraGridColumn89.Header.VisiblePosition = 80;
			ultraGridColumn89.Hidden = true;
			ultraGridColumn90.Header.VisiblePosition = 0;
			ultraGridColumn90.Width = 166;
			ultraGridColumn91.Hidden = true;
			ultraGridColumn92.Hidden = true;
			ultraGridColumn93.Hidden = true;
			ultraGridColumn94.Hidden = true;
			ultraGridColumn95.Hidden = true;
			ultraGridColumn96.Hidden = true;
			ultraGridColumn97.Hidden = true;
			ultraGridColumn98.Hidden = true;
			ultraGridBand2.Columns.Add(ultraGridColumn10);
			ultraGridBand2.Columns.Add(ultraGridColumn11);
			ultraGridBand2.Columns.Add(ultraGridColumn12);
			ultraGridBand2.Columns.Add(ultraGridColumn13);
			ultraGridBand2.Columns.Add(ultraGridColumn14);
			ultraGridBand2.Columns.Add(ultraGridColumn15);
			ultraGridBand2.Columns.Add(ultraGridColumn16);
			ultraGridBand2.Columns.Add(ultraGridColumn17);
			ultraGridBand2.Columns.Add(ultraGridColumn18);
			ultraGridBand2.Columns.Add(ultraGridColumn19);
			ultraGridBand2.Columns.Add(ultraGridColumn20);
			ultraGridBand2.Columns.Add(ultraGridColumn21);
			ultraGridBand2.Columns.Add(ultraGridColumn22);
			ultraGridBand2.Columns.Add(ultraGridColumn23);
			ultraGridBand2.Columns.Add(ultraGridColumn24);
			ultraGridBand2.Columns.Add(ultraGridColumn25);
			ultraGridBand2.Columns.Add(ultraGridColumn26);
			ultraGridBand2.Columns.Add(ultraGridColumn27);
			ultraGridBand2.Columns.Add(ultraGridColumn28);
			ultraGridBand2.Columns.Add(ultraGridColumn29);
			ultraGridBand2.Columns.Add(ultraGridColumn30);
			ultraGridBand2.Columns.Add(ultraGridColumn31);
			ultraGridBand2.Columns.Add(ultraGridColumn32);
			ultraGridBand2.Columns.Add(ultraGridColumn33);
			ultraGridBand2.Columns.Add(ultraGridColumn34);
			ultraGridBand2.Columns.Add(ultraGridColumn35);
			ultraGridBand2.Columns.Add(ultraGridColumn36);
			ultraGridBand2.Columns.Add(ultraGridColumn37);
			ultraGridBand2.Columns.Add(ultraGridColumn38);
			ultraGridBand2.Columns.Add(ultraGridColumn39);
			ultraGridBand2.Columns.Add(ultraGridColumn40);
			ultraGridBand2.Columns.Add(ultraGridColumn41);
			ultraGridBand2.Columns.Add(ultraGridColumn42);
			ultraGridBand2.Columns.Add(ultraGridColumn43);
			ultraGridBand2.Columns.Add(ultraGridColumn44);
			ultraGridBand2.Columns.Add(ultraGridColumn45);
			ultraGridBand2.Columns.Add(ultraGridColumn46);
			ultraGridBand2.Columns.Add(ultraGridColumn47);
			ultraGridBand2.Columns.Add(ultraGridColumn48);
			ultraGridBand2.Columns.Add(ultraGridColumn49);
			ultraGridBand2.Columns.Add(ultraGridColumn50);
			ultraGridBand2.Columns.Add(ultraGridColumn51);
			ultraGridBand2.Columns.Add(ultraGridColumn52);
			ultraGridBand2.Columns.Add(ultraGridColumn53);
			ultraGridBand2.Columns.Add(ultraGridColumn54);
			ultraGridBand2.Columns.Add(ultraGridColumn55);
			ultraGridBand2.Columns.Add(ultraGridColumn56);
			ultraGridBand2.Columns.Add(ultraGridColumn57);
			ultraGridBand2.Columns.Add(ultraGridColumn58);
			ultraGridBand2.Columns.Add(ultraGridColumn59);
			ultraGridBand2.Columns.Add(ultraGridColumn60);
			ultraGridBand2.Columns.Add(ultraGridColumn61);
			ultraGridBand2.Columns.Add(ultraGridColumn62);
			ultraGridBand2.Columns.Add(ultraGridColumn63);
			ultraGridBand2.Columns.Add(ultraGridColumn64);
			ultraGridBand2.Columns.Add(ultraGridColumn65);
			ultraGridBand2.Columns.Add(ultraGridColumn66);
			ultraGridBand2.Columns.Add(ultraGridColumn67);
			ultraGridBand2.Columns.Add(ultraGridColumn68);
			ultraGridBand2.Columns.Add(ultraGridColumn69);
			ultraGridBand2.Columns.Add(ultraGridColumn70);
			ultraGridBand2.Columns.Add(ultraGridColumn71);
			ultraGridBand2.Columns.Add(ultraGridColumn72);
			ultraGridBand2.Columns.Add(ultraGridColumn73);
			ultraGridBand2.Columns.Add(ultraGridColumn74);
			ultraGridBand2.Columns.Add(ultraGridColumn75);
			ultraGridBand2.Columns.Add(ultraGridColumn76);
			ultraGridBand2.Columns.Add(ultraGridColumn77);
			ultraGridBand2.Columns.Add(ultraGridColumn78);
			ultraGridBand2.Columns.Add(ultraGridColumn79);
			ultraGridBand2.Columns.Add(ultraGridColumn80);
			ultraGridBand2.Columns.Add(ultraGridColumn81);
			ultraGridBand2.Columns.Add(ultraGridColumn82);
			ultraGridBand2.Columns.Add(ultraGridColumn83);
			ultraGridBand2.Columns.Add(ultraGridColumn84);
			ultraGridBand2.Columns.Add(ultraGridColumn85);
			ultraGridBand2.Columns.Add(ultraGridColumn86);
			ultraGridBand2.Columns.Add(ultraGridColumn87);
			ultraGridBand2.Columns.Add(ultraGridColumn88);
			ultraGridBand2.Columns.Add(ultraGridColumn89);
			ultraGridBand2.Columns.Add(ultraGridColumn90);
			ultraGridBand2.Columns.Add(ultraGridColumn91);
			ultraGridBand2.Columns.Add(ultraGridColumn92);
			ultraGridBand2.Columns.Add(ultraGridColumn93);
			ultraGridBand2.Columns.Add(ultraGridColumn94);
			ultraGridBand2.Columns.Add(ultraGridColumn95);
			ultraGridBand2.Columns.Add(ultraGridColumn96);
			ultraGridBand2.Columns.Add(ultraGridColumn97);
			ultraGridBand2.Columns.Add(ultraGridColumn98);
			ultraGridBand2.Columns.Add(ultraGridColumn99);
			ultraGridBand2.Columns.Add(ultraGridColumn100);
			this.ultraDropDown1.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			this.ultraDropDown1.DisplayMember = "Articulo";
			this.ultraDropDown1.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.ultraDropDown1.Location = new System.Drawing.Point(104, 144);
			this.ultraDropDown1.Name = "ultraDropDown1";
			this.ultraDropDown1.Size = new System.Drawing.Size(288, 80);
			this.ultraDropDown1.TabIndex = 11;
			this.ultraDropDown1.Text = "cmbArticulo";
			this.ultraDropDown1.ValueMember = "idArticulo";
			this.ultraDropDown1.Visible = false;
			// 
			// ultraDropDown2
			// 
			this.ultraDropDown2.Cursor = System.Windows.Forms.Cursors.Default;
			this.ultraDropDown2.DataSource = this.cdvArtCodigo;
			ultraGridColumn101.Header.VisiblePosition = 2;
			ultraGridColumn101.Hidden = true;
			ultraGridColumn102.Header.VisiblePosition = 3;
			ultraGridColumn102.Hidden = true;
			ultraGridColumn103.Header.VisiblePosition = 4;
			ultraGridColumn103.Hidden = true;
			ultraGridColumn104.Header.VisiblePosition = 5;
			ultraGridColumn104.Hidden = true;
			ultraGridColumn105.Header.VisiblePosition = 6;
			ultraGridColumn105.Hidden = true;
			ultraGridColumn106.Header.VisiblePosition = 7;
			ultraGridColumn106.Hidden = true;
			ultraGridColumn107.Header.VisiblePosition = 8;
			ultraGridColumn107.Hidden = true;
			ultraGridColumn108.Header.VisiblePosition = 9;
			ultraGridColumn108.Hidden = true;
			ultraGridColumn109.Header.VisiblePosition = 10;
			ultraGridColumn109.Hidden = true;
			ultraGridColumn110.Header.VisiblePosition = 11;
			ultraGridColumn110.Hidden = true;
			ultraGridColumn111.Header.VisiblePosition = 12;
			ultraGridColumn111.Hidden = true;
			ultraGridColumn112.Header.VisiblePosition = 13;
			ultraGridColumn112.Hidden = true;
			ultraGridColumn113.Header.VisiblePosition = 14;
			ultraGridColumn113.Hidden = true;
			ultraGridColumn114.Header.VisiblePosition = 15;
			ultraGridColumn114.Hidden = true;
			ultraGridColumn115.Header.VisiblePosition = 16;
			ultraGridColumn115.Hidden = true;
			ultraGridColumn116.Header.VisiblePosition = 17;
			ultraGridColumn116.Hidden = true;
			ultraGridColumn117.Header.VisiblePosition = 18;
			ultraGridColumn117.Hidden = true;
			ultraGridColumn118.Header.VisiblePosition = 19;
			ultraGridColumn118.Hidden = true;
			ultraGridColumn119.Header.VisiblePosition = 20;
			ultraGridColumn119.Hidden = true;
			ultraGridColumn120.Header.VisiblePosition = 21;
			ultraGridColumn120.Hidden = true;
			ultraGridColumn121.Header.VisiblePosition = 22;
			ultraGridColumn121.Hidden = true;
			ultraGridColumn122.Header.VisiblePosition = 0;
			ultraGridColumn123.Header.VisiblePosition = 23;
			ultraGridColumn123.Hidden = true;
			ultraGridColumn124.Header.VisiblePosition = 24;
			ultraGridColumn124.Hidden = true;
			ultraGridColumn125.Header.VisiblePosition = 25;
			ultraGridColumn125.Hidden = true;
			ultraGridColumn126.Header.VisiblePosition = 26;
			ultraGridColumn126.Hidden = true;
			ultraGridColumn127.Header.VisiblePosition = 27;
			ultraGridColumn127.Hidden = true;
			ultraGridColumn128.Header.VisiblePosition = 28;
			ultraGridColumn128.Hidden = true;
			ultraGridColumn129.Header.VisiblePosition = 29;
			ultraGridColumn129.Hidden = true;
			ultraGridColumn130.Header.VisiblePosition = 30;
			ultraGridColumn130.Hidden = true;
			ultraGridColumn131.Header.VisiblePosition = 31;
			ultraGridColumn131.Hidden = true;
			ultraGridColumn132.Header.VisiblePosition = 32;
			ultraGridColumn132.Hidden = true;
			ultraGridColumn133.Header.VisiblePosition = 33;
			ultraGridColumn133.Hidden = true;
			ultraGridColumn134.Header.VisiblePosition = 34;
			ultraGridColumn134.Hidden = true;
			ultraGridColumn135.Header.VisiblePosition = 35;
			ultraGridColumn135.Hidden = true;
			ultraGridColumn136.Header.VisiblePosition = 36;
			ultraGridColumn136.Hidden = true;
			ultraGridColumn137.Header.VisiblePosition = 37;
			ultraGridColumn137.Hidden = true;
			ultraGridColumn138.Header.VisiblePosition = 38;
			ultraGridColumn138.Hidden = true;
			ultraGridColumn139.Header.VisiblePosition = 39;
			ultraGridColumn139.Hidden = true;
			ultraGridColumn140.Header.VisiblePosition = 40;
			ultraGridColumn140.Hidden = true;
			ultraGridColumn141.Header.VisiblePosition = 41;
			ultraGridColumn141.Hidden = true;
			ultraGridColumn142.Header.VisiblePosition = 42;
			ultraGridColumn142.Hidden = true;
			ultraGridColumn143.Header.VisiblePosition = 43;
			ultraGridColumn143.Hidden = true;
			ultraGridColumn144.Header.VisiblePosition = 44;
			ultraGridColumn144.Hidden = true;
			ultraGridColumn145.Header.VisiblePosition = 45;
			ultraGridColumn145.Hidden = true;
			ultraGridColumn146.Header.VisiblePosition = 46;
			ultraGridColumn146.Hidden = true;
			ultraGridColumn147.Header.VisiblePosition = 47;
			ultraGridColumn147.Hidden = true;
			ultraGridColumn148.Header.VisiblePosition = 48;
			ultraGridColumn148.Hidden = true;
			ultraGridColumn149.Header.VisiblePosition = 49;
			ultraGridColumn149.Hidden = true;
			ultraGridColumn150.Header.VisiblePosition = 50;
			ultraGridColumn150.Hidden = true;
			ultraGridColumn151.Header.VisiblePosition = 51;
			ultraGridColumn151.Hidden = true;
			ultraGridColumn152.Header.VisiblePosition = 52;
			ultraGridColumn152.Hidden = true;
			ultraGridColumn153.Header.VisiblePosition = 53;
			ultraGridColumn153.Hidden = true;
			ultraGridColumn154.Header.VisiblePosition = 54;
			ultraGridColumn154.Hidden = true;
			ultraGridColumn155.Header.VisiblePosition = 55;
			ultraGridColumn155.Hidden = true;
			ultraGridColumn156.Header.VisiblePosition = 56;
			ultraGridColumn156.Hidden = true;
			ultraGridColumn157.Header.VisiblePosition = 57;
			ultraGridColumn157.Hidden = true;
			ultraGridColumn158.Header.VisiblePosition = 58;
			ultraGridColumn158.Hidden = true;
			ultraGridColumn159.Header.VisiblePosition = 59;
			ultraGridColumn159.Hidden = true;
			ultraGridColumn160.Header.VisiblePosition = 60;
			ultraGridColumn160.Hidden = true;
			ultraGridColumn161.Header.VisiblePosition = 61;
			ultraGridColumn161.Hidden = true;
			ultraGridColumn162.Header.VisiblePosition = 62;
			ultraGridColumn162.Hidden = true;
			ultraGridColumn163.Header.VisiblePosition = 63;
			ultraGridColumn163.Hidden = true;
			ultraGridColumn164.Header.VisiblePosition = 64;
			ultraGridColumn164.Hidden = true;
			ultraGridColumn165.Header.VisiblePosition = 65;
			ultraGridColumn165.Hidden = true;
			ultraGridColumn166.Header.VisiblePosition = 66;
			ultraGridColumn166.Hidden = true;
			ultraGridColumn167.Header.VisiblePosition = 67;
			ultraGridColumn167.Hidden = true;
			ultraGridColumn168.Header.VisiblePosition = 68;
			ultraGridColumn168.Hidden = true;
			ultraGridColumn169.Header.VisiblePosition = 69;
			ultraGridColumn169.Hidden = true;
			ultraGridColumn170.Header.VisiblePosition = 70;
			ultraGridColumn170.Hidden = true;
			ultraGridColumn171.Header.VisiblePosition = 71;
			ultraGridColumn171.Hidden = true;
			ultraGridColumn172.Header.VisiblePosition = 72;
			ultraGridColumn172.Hidden = true;
			ultraGridColumn173.Header.VisiblePosition = 73;
			ultraGridColumn173.Hidden = true;
			ultraGridColumn174.Header.VisiblePosition = 74;
			ultraGridColumn174.Hidden = true;
			ultraGridColumn175.Header.VisiblePosition = 75;
			ultraGridColumn175.Hidden = true;
			ultraGridColumn176.Header.VisiblePosition = 76;
			ultraGridColumn176.Hidden = true;
			ultraGridColumn177.Header.VisiblePosition = 77;
			ultraGridColumn177.Hidden = true;
			ultraGridColumn178.Header.VisiblePosition = 78;
			ultraGridColumn178.Hidden = true;
			ultraGridColumn179.Header.VisiblePosition = 79;
			ultraGridColumn179.Hidden = true;
			ultraGridColumn180.Header.VisiblePosition = 80;
			ultraGridColumn180.Hidden = true;
			ultraGridColumn181.Header.VisiblePosition = 1;
			ultraGridColumn181.Width = 166;
			ultraGridColumn182.Hidden = true;
			ultraGridColumn183.Hidden = true;
			ultraGridColumn184.Hidden = true;
			ultraGridColumn185.Hidden = true;
			ultraGridColumn186.Hidden = true;
			ultraGridColumn187.Hidden = true;
			ultraGridColumn188.Hidden = true;
			ultraGridColumn189.Hidden = true;
			ultraGridBand3.Columns.Add(ultraGridColumn101);
			ultraGridBand3.Columns.Add(ultraGridColumn102);
			ultraGridBand3.Columns.Add(ultraGridColumn103);
			ultraGridBand3.Columns.Add(ultraGridColumn104);
			ultraGridBand3.Columns.Add(ultraGridColumn105);
			ultraGridBand3.Columns.Add(ultraGridColumn106);
			ultraGridBand3.Columns.Add(ultraGridColumn107);
			ultraGridBand3.Columns.Add(ultraGridColumn108);
			ultraGridBand3.Columns.Add(ultraGridColumn109);
			ultraGridBand3.Columns.Add(ultraGridColumn110);
			ultraGridBand3.Columns.Add(ultraGridColumn111);
			ultraGridBand3.Columns.Add(ultraGridColumn112);
			ultraGridBand3.Columns.Add(ultraGridColumn113);
			ultraGridBand3.Columns.Add(ultraGridColumn114);
			ultraGridBand3.Columns.Add(ultraGridColumn115);
			ultraGridBand3.Columns.Add(ultraGridColumn116);
			ultraGridBand3.Columns.Add(ultraGridColumn117);
			ultraGridBand3.Columns.Add(ultraGridColumn118);
			ultraGridBand3.Columns.Add(ultraGridColumn119);
			ultraGridBand3.Columns.Add(ultraGridColumn120);
			ultraGridBand3.Columns.Add(ultraGridColumn121);
			ultraGridBand3.Columns.Add(ultraGridColumn122);
			ultraGridBand3.Columns.Add(ultraGridColumn123);
			ultraGridBand3.Columns.Add(ultraGridColumn124);
			ultraGridBand3.Columns.Add(ultraGridColumn125);
			ultraGridBand3.Columns.Add(ultraGridColumn126);
			ultraGridBand3.Columns.Add(ultraGridColumn127);
			ultraGridBand3.Columns.Add(ultraGridColumn128);
			ultraGridBand3.Columns.Add(ultraGridColumn129);
			ultraGridBand3.Columns.Add(ultraGridColumn130);
			ultraGridBand3.Columns.Add(ultraGridColumn131);
			ultraGridBand3.Columns.Add(ultraGridColumn132);
			ultraGridBand3.Columns.Add(ultraGridColumn133);
			ultraGridBand3.Columns.Add(ultraGridColumn134);
			ultraGridBand3.Columns.Add(ultraGridColumn135);
			ultraGridBand3.Columns.Add(ultraGridColumn136);
			ultraGridBand3.Columns.Add(ultraGridColumn137);
			ultraGridBand3.Columns.Add(ultraGridColumn138);
			ultraGridBand3.Columns.Add(ultraGridColumn139);
			ultraGridBand3.Columns.Add(ultraGridColumn140);
			ultraGridBand3.Columns.Add(ultraGridColumn141);
			ultraGridBand3.Columns.Add(ultraGridColumn142);
			ultraGridBand3.Columns.Add(ultraGridColumn143);
			ultraGridBand3.Columns.Add(ultraGridColumn144);
			ultraGridBand3.Columns.Add(ultraGridColumn145);
			ultraGridBand3.Columns.Add(ultraGridColumn146);
			ultraGridBand3.Columns.Add(ultraGridColumn147);
			ultraGridBand3.Columns.Add(ultraGridColumn148);
			ultraGridBand3.Columns.Add(ultraGridColumn149);
			ultraGridBand3.Columns.Add(ultraGridColumn150);
			ultraGridBand3.Columns.Add(ultraGridColumn151);
			ultraGridBand3.Columns.Add(ultraGridColumn152);
			ultraGridBand3.Columns.Add(ultraGridColumn153);
			ultraGridBand3.Columns.Add(ultraGridColumn154);
			ultraGridBand3.Columns.Add(ultraGridColumn155);
			ultraGridBand3.Columns.Add(ultraGridColumn156);
			ultraGridBand3.Columns.Add(ultraGridColumn157);
			ultraGridBand3.Columns.Add(ultraGridColumn158);
			ultraGridBand3.Columns.Add(ultraGridColumn159);
			ultraGridBand3.Columns.Add(ultraGridColumn160);
			ultraGridBand3.Columns.Add(ultraGridColumn161);
			ultraGridBand3.Columns.Add(ultraGridColumn162);
			ultraGridBand3.Columns.Add(ultraGridColumn163);
			ultraGridBand3.Columns.Add(ultraGridColumn164);
			ultraGridBand3.Columns.Add(ultraGridColumn165);
			ultraGridBand3.Columns.Add(ultraGridColumn166);
			ultraGridBand3.Columns.Add(ultraGridColumn167);
			ultraGridBand3.Columns.Add(ultraGridColumn168);
			ultraGridBand3.Columns.Add(ultraGridColumn169);
			ultraGridBand3.Columns.Add(ultraGridColumn170);
			ultraGridBand3.Columns.Add(ultraGridColumn171);
			ultraGridBand3.Columns.Add(ultraGridColumn172);
			ultraGridBand3.Columns.Add(ultraGridColumn173);
			ultraGridBand3.Columns.Add(ultraGridColumn174);
			ultraGridBand3.Columns.Add(ultraGridColumn175);
			ultraGridBand3.Columns.Add(ultraGridColumn176);
			ultraGridBand3.Columns.Add(ultraGridColumn177);
			ultraGridBand3.Columns.Add(ultraGridColumn178);
			ultraGridBand3.Columns.Add(ultraGridColumn179);
			ultraGridBand3.Columns.Add(ultraGridColumn180);
			ultraGridBand3.Columns.Add(ultraGridColumn181);
			ultraGridBand3.Columns.Add(ultraGridColumn182);
			ultraGridBand3.Columns.Add(ultraGridColumn183);
			ultraGridBand3.Columns.Add(ultraGridColumn184);
			ultraGridBand3.Columns.Add(ultraGridColumn185);
			ultraGridBand3.Columns.Add(ultraGridColumn186);
			ultraGridBand3.Columns.Add(ultraGridColumn187);
			ultraGridBand3.Columns.Add(ultraGridColumn188);
			ultraGridBand3.Columns.Add(ultraGridColumn189);
			ultraGridBand3.Columns.Add(ultraGridColumn190);
			ultraGridBand3.Columns.Add(ultraGridColumn191);
			this.ultraDropDown2.DisplayLayout.BandsSerializer.Add(ultraGridBand3);
			this.ultraDropDown2.DisplayMember = "Codigo";
			this.ultraDropDown2.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.ultraDropDown2.Location = new System.Drawing.Point(248, 208);
			this.ultraDropDown2.Name = "ultraDropDown2";
			this.ultraDropDown2.Size = new System.Drawing.Size(288, 80);
			this.ultraDropDown2.TabIndex = 12;
			this.ultraDropDown2.Text = "cmbArtCodigo";
			this.ultraDropDown2.ValueMember = "idArticulo";
			this.ultraDropDown2.Visible = false;
			// 
			// Receta1
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(544, 326);
			this.Controls.AddRange(new System.Windows.Forms.Control[] {
																																	this.ultraDropDown2,
																																	this.ultraDropDown1,
																																	this.barraDato1,
																																	this.label1,
																																	this.cmbArtCodigo,
																																	this.txtIdProducto,
																																	this.cdmArticulo,
																																	this.ultraGrid1});
			this.Name = "Receta1";
			this.Text = "Receta";
			this.Load += new System.EventHandler(this.Receta1_Load);
			((System.ComponentModel.ISupportInitialize)(this.cdsProducto)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvProducto)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvArtCodigo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsArticulo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvArticulo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbArtCodigo.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdProducto.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdmArticulo.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDropDown1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDropDown2)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void barraDato1_Borrar(object sender, System.EventArgs e)
		{
			if (DialogResult.No == MessageBox.Show(
				"¿Está seguro que desea borrar la Receta?"
				+ ".\nEste proceso no se puede reversar", 
				"Confirmación", MessageBoxButtons.YesNo, MessageBoxIcon.Warning, 
				MessageBoxDefaultButton.Button2)) return;
			
			this.barraDato1.BorraRegistro();
		}

		private void barraDato1_Buscar(object sender, System.EventArgs e)
		{
			BuscaReceta miBusqueda = new BuscaReceta();
			int id, idx;

			if (miBusqueda.ShowDialog() == DialogResult.OK)
			{
				id = int.Parse(miBusqueda.idBusca.Text);
				idx = this.cdvProducto.Find(id);
				if (idx == -1) MessageBox.Show("IdAsiento no encontrado " + id.ToString(), "Error");
				else BindingContext[this.cdsProducto,"_Producto"].Position = idx;

				this.barraDato1.ActualizaPosicion();
			}		
		}

		private void barraDato1_Crear(object sender, System.EventArgs e)
		{
			this.barraDato1.CrearRegistro();
		}

		private void barraDato1_Grabar(object sender, System.EventArgs e)
		{
			string stBusca = "idArticulo = " + this.cmbArtCodigo.EditValue.ToString() +
				" And idProducto <> " + this.txtIdProducto.EditValue.ToString();
			DataRow [] dr = this.cdsProducto.StorageDataSet.Tables["Producto"].Select(stBusca);
			if (dr.Length > 0)
			{
				MessageBox.Show("Receta de Artículo ya existe");
				return;
			}
			this.barraDato1.GrabaRegistro();
		}

		private void barraDato1_Editar(object sender, System.EventArgs e)
		{
			this.barraDato1.EditarRegistro();
		}

		private void barraDato1_Deshacer(object sender, System.EventArgs e)
		{
			this.barraDato1.DeshacerRegistro();
		}

		private void Receta1_Load(object sender, System.EventArgs e)
		{
			this.txtIdProducto.Width = 0;
			this.barraDato1.PosUltima();
		}

		private void cdsProducto_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			this.cdsProducto.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirectorio);
		}

		private void cdsArticulo_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			this.cdsArticulo.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirectorio);		
		}
	}
}
