using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using Infragistics.Win.UltraWinGrid;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de TomaFisica.
	/// </summary>
	public class TomaFisica : DevExpress.XtraEditors.XtraForm
	{
		private Infragistics.Win.UltraWinGrid.UltraGrid ultraGrid1;
		private C1.Data.C1DataSet cdsTomaFisica;
		private BarraDatoSQL.BarraDatoSQL barraDatoSQL1;
		private C1.Data.C1DataSet cdsArticulo;
		private C1.Data.C1DataView cdvArticulo;
		private C1.Data.C1DataView cdvArtCodigo;
		private Infragistics.Win.Misc.UltraLabel ultraLabel1;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo cmbFecha;
		private Infragistics.Win.Misc.UltraLabel ultraLabel2;
		private Infragistics.Win.Misc.UltraLabel ultraLabel3;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtTarjeta;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtBodega;
		private Infragistics.Win.UltraWinGrid.UltraDropDown cmbCodigo;
		private Infragistics.Win.UltraWinGrid.UltraDropDown cmbArticulo;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtIdToma;
		private C1.Data.C1DataSet cdsSeguridad;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public TomaFisica()
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
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("_TmpTomaDetalle", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idTmpTomaDetalle");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idTmpTomaFisica");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo", -1, "cmbCodigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idArticulo", -1, "cmbArticulo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cantidad");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Saldo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Mostrar");
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Articulo", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idArticulo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idArticuloSRI");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idColor");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn11 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idModelo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn12 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idGrupoArticulo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn13 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idSubGrupo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn14 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idArancel");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn15 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idMarca");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn16 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idUnidad");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn17 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPromocion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn18 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idTipoGrupo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn19 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idClase");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn20 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idSeccion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn21 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idSucursal");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn22 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPadre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn23 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaIngreso");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn24 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CodAlterno");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn25 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Saldo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn26 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Peso");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn27 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Caracteristica");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn28 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Aplicacion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn29 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Iva");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn30 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Ice");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn31 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Precio1");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn32 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Precio2");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn33 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Precio3");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn34 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Precio4");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn35 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Precio5");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn36 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Precio6");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn37 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CantidadMinima");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn38 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CantidadMaxima");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn39 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Residual");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn40 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CostoUltimo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn41 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaUltimo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn42 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCompra1");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn43 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCompra2");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn44 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCompra3");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn45 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CostoPromedio");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn46 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Gramos");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn47 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ConsumoProm");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn48 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("TiempoReposicion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn49 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("UnidadPeso");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn50 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Depreciacion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn51 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Incr1");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn52 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Incr2");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn53 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Incr3");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn54 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Incr4");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn55 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Incr5");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn56 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Incr6");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn57 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Ubicacion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn58 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("NumeroParte");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn59 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ComisionValor");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn60 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ComisionPorcentaje");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn61 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CantidadRequerida");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn62 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("PrecioPromocion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn63 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CantidadPromocion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn64 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("DescuentoContado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn65 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("DescuentoCredito");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn66 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("DescPorcContado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn67 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("DescPorcCre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn68 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CantReqDesc");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn69 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CantReqPorc");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn70 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("TipoPromocion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn71 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Responsable");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn72 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Comision30");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn73 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Comision60");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn74 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Comision90");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn75 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Comision120");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn76 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bodega1");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn77 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bodega2");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn78 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bodega3");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn79 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bodega4");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn80 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bodega5");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn81 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bodega6");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn82 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bodega7");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn83 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bodega8");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn84 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bodega9");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn85 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("SaldoT");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn86 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Alto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn87 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Ancho");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn88 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Largo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn89 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Color");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn90 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Descontinuado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn91 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Borrar");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn92 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Habilitado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn93 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ValorPeso");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn94 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("BodegaPredef");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn95 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Tiempo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn96 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cantidad1");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn97 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cantidad2");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn98 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cantidad3");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn99 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cantidad4");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn100 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cantidad5");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn101 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("PaqueteUnidad");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn102 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Editar");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn103 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nuevo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn104 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaEditado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn105 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CambiarPrecio");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn106 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("dCompra");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn107 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("dVenta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn108 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Foto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn109 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Notas");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn110 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Articulo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn111 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn112 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("BodegaPredefCR");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn113 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CodPromocion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn114 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CostoFecha");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn115 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Aportable");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn116 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ExigeLote");
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton1 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand3 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Articulo", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn117 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idArticulo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn118 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idArticuloSRI");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn119 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idColor");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn120 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idModelo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn121 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idGrupoArticulo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn122 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idSubGrupo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn123 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idArancel");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn124 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idMarca");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn125 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idUnidad");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn126 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPromocion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn127 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idTipoGrupo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn128 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idClase");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn129 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idSeccion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn130 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idSucursal");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn131 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPadre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn132 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaIngreso");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn133 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CodAlterno");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn134 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Saldo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn135 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Peso");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn136 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Caracteristica");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn137 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Aplicacion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn138 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Iva");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn139 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Ice");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn140 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Precio1");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn141 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Precio2");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn142 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Precio3");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn143 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Precio4");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn144 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Precio5");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn145 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Precio6");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn146 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CantidadMinima");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn147 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CantidadMaxima");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn148 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Residual");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn149 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CostoUltimo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn150 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaUltimo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn151 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCompra1");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn152 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCompra2");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn153 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCompra3");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn154 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CostoPromedio");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn155 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Gramos");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn156 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ConsumoProm");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn157 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("TiempoReposicion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn158 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("UnidadPeso");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn159 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Depreciacion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn160 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Incr1");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn161 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Incr2");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn162 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Incr3");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn163 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Incr4");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn164 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Incr5");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn165 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Incr6");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn166 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Ubicacion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn167 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("NumeroParte");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn168 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ComisionValor");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn169 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ComisionPorcentaje");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn170 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CantidadRequerida");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn171 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("PrecioPromocion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn172 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CantidadPromocion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn173 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("DescuentoContado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn174 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("DescuentoCredito");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn175 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("DescPorcContado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn176 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("DescPorcCre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn177 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CantReqDesc");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn178 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CantReqPorc");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn179 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("TipoPromocion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn180 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Responsable");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn181 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Comision30");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn182 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Comision60");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn183 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Comision90");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn184 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Comision120");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn185 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bodega1");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn186 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bodega2");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn187 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bodega3");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn188 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bodega4");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn189 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bodega5");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn190 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bodega6");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn191 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bodega7");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn192 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bodega8");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn193 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bodega9");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn194 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("SaldoT");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn195 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Alto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn196 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Ancho");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn197 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Largo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn198 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Color");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn199 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Descontinuado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn200 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Borrar");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn201 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Habilitado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn202 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ValorPeso");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn203 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("BodegaPredef");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn204 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Tiempo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn205 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cantidad1");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn206 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cantidad2");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn207 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cantidad3");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn208 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cantidad4");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn209 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cantidad5");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn210 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("PaqueteUnidad");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn211 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Editar");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn212 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nuevo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn213 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaEditado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn214 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CambiarPrecio");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn215 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("dCompra");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn216 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("dVenta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn217 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Foto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn218 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Notas");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn219 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Articulo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn220 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn221 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("BodegaPredefCR");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn222 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CodPromocion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn223 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CostoFecha");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn224 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Aportable");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn225 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ExigeLote");
			this.ultraGrid1 = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.cdsTomaFisica = new C1.Data.C1DataSet();
			this.barraDatoSQL1 = new BarraDatoSQL.BarraDatoSQL();
			this.txtTarjeta = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.cdsArticulo = new C1.Data.C1DataSet();
			this.cdvArticulo = new C1.Data.C1DataView();
			this.cdvArtCodigo = new C1.Data.C1DataView();
			this.cmbCodigo = new Infragistics.Win.UltraWinGrid.UltraDropDown();
			this.ultraLabel1 = new Infragistics.Win.Misc.UltraLabel();
			this.cmbFecha = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.ultraLabel2 = new Infragistics.Win.Misc.UltraLabel();
			this.ultraLabel3 = new Infragistics.Win.Misc.UltraLabel();
			this.txtBodega = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.cmbArticulo = new Infragistics.Win.UltraWinGrid.UltraDropDown();
			this.txtIdToma = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.cdsSeguridad = new C1.Data.C1DataSet();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsTomaFisica)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtTarjeta)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsArticulo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvArticulo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvArtCodigo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCodigo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbFecha)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtBodega)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbArticulo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdToma)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeguridad)).BeginInit();
			this.SuspendLayout();
			// 
			// ultraGrid1
			// 
			this.ultraGrid1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.ultraGrid1.Cursor = System.Windows.Forms.Cursors.Default;
			this.ultraGrid1.DataMember = "_TmpTomaFisica.TomaFisica";
			this.ultraGrid1.DataSource = this.cdsTomaFisica;
			appearance1.BackColor = System.Drawing.Color.White;
			this.ultraGrid1.DisplayLayout.Appearance = appearance1;
			ultraGridColumn1.Header.VisiblePosition = 5;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn2.Header.VisiblePosition = 4;
			ultraGridColumn2.Hidden = true;
			ultraGridColumn3.Header.VisiblePosition = 1;
			ultraGridColumn3.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownValidate;
			ultraGridColumn3.Width = 131;
			ultraGridColumn4.Header.Caption = "Articulo";
			ultraGridColumn4.Header.VisiblePosition = 0;
			ultraGridColumn4.PromptChar = ' ';
			ultraGridColumn4.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownValidate;
			ultraGridColumn4.Width = 254;
			ultraGridColumn5.Header.VisiblePosition = 2;
			ultraGridColumn6.Header.VisiblePosition = 3;
			ultraGridColumn6.Hidden = true;
			ultraGridColumn7.Header.VisiblePosition = 6;
			ultraGridColumn7.Hidden = true;
			ultraGridBand1.Columns.AddRange(new object[] {
																										 ultraGridColumn1,
																										 ultraGridColumn2,
																										 ultraGridColumn3,
																										 ultraGridColumn4,
																										 ultraGridColumn5,
																										 ultraGridColumn6,
																										 ultraGridColumn7});
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
			this.ultraGrid1.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortSingle;
			appearance4.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance4.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance4.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.RowSelectorAppearance = appearance4;
			appearance5.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance5.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance5.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.SelectedRowAppearance = appearance5;
			this.ultraGrid1.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.ultraGrid1.Location = new System.Drawing.Point(16, 52);
			this.ultraGrid1.Name = "ultraGrid1";
			this.ultraGrid1.Size = new System.Drawing.Size(496, 250);
			this.ultraGrid1.TabIndex = 0;
			this.ultraGrid1.InitializeRow += new Infragistics.Win.UltraWinGrid.InitializeRowEventHandler(this.ultraGrid1_InitializeRow);
			this.ultraGrid1.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.ultraGrid1_KeyPress);
			this.ultraGrid1.KeyDown += new System.Windows.Forms.KeyEventHandler(this.ultraGrid1_KeyDown);
			this.ultraGrid1.AfterCellUpdate += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.ultraGrid1_AfterCellUpdate);
			this.ultraGrid1.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.ultraGrid1_InitializeLayout);
			// 
			// cdsTomaFisica
			// 
			this.cdsTomaFisica.BindingContextCtrl = this;
			this.cdsTomaFisica.DataLibrary = "LibFacturacion";
			this.cdsTomaFisica.DataLibraryUrl = "";
			this.cdsTomaFisica.DataSetDef = "dsTomaFisica";
			this.cdsTomaFisica.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsTomaFisica.Name = "cdsTomaFisica";
			this.cdsTomaFisica.SchemaClassName = "LibFacturacion.DataClass";
			this.cdsTomaFisica.SchemaDef = null;
			this.cdsTomaFisica.BeforeFill += new C1.Data.FillEventHandler(this.cdsTomaFisica_BeforeFill);
			// 
			// barraDatoSQL1
			// 
			this.barraDatoSQL1.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.barraDatoSQL1.BarraMovimiento = true;
			this.barraDatoSQL1.DataMember = "_TmpTomaFisica";
			this.barraDatoSQL1.DataNombreId = "idTmpTomaFisica";
			this.barraDatoSQL1.DataOrden = "";
			this.barraDatoSQL1.DataSource = this.cdsTomaFisica;
			this.barraDatoSQL1.DataTabla = "TmpTomaFisica";
			this.barraDatoSQL1.DataTablaHija = "TmpTomaDetalle";
			this.barraDatoSQL1.Location = new System.Drawing.Point(40, 319);
			this.barraDatoSQL1.Name = "barraDatoSQL1";
			this.barraDatoSQL1.Size = new System.Drawing.Size(304, 26);
			this.barraDatoSQL1.TabIndex = 1;
			this.barraDatoSQL1.VisibleBorrar = true;
			this.barraDatoSQL1.VisibleBuscar = true;
			this.barraDatoSQL1.VisibleEditar = true;
			this.barraDatoSQL1.VisibleImprimir = true;
			this.barraDatoSQL1.VisibleNuevo = true;
			this.barraDatoSQL1.Buscar += new System.EventHandler(this.barraDatoSQL1_Buscar);
			this.barraDatoSQL1.Borrar += new System.EventHandler(this.barraDatoSQL1_Borrar);
			this.barraDatoSQL1.Imprimir += new System.EventHandler(this.barraDatoSQL1_Imprimir);
			this.barraDatoSQL1.Restaurar += new System.EventHandler(this.barraDatoSQL1_Restaurar);
			this.barraDatoSQL1.Refresca += new System.EventHandler(this.barraDatoSQL1_Refresca);
			this.barraDatoSQL1.Crear += new System.EventHandler(this.barraDatoSQL1_Crear);
			this.barraDatoSQL1.Deshacer += new System.EventHandler(this.barraDatoSQL1_Deshacer);
			this.barraDatoSQL1.Impresora += new System.EventHandler(this.barraDatoSQL1_Impresora);
			this.barraDatoSQL1.Editar += new System.EventHandler(this.barraDatoSQL1_Editar);
			this.barraDatoSQL1.Grabar += new System.EventHandler(this.barraDatoSQL1_Grabar);
			// 
			// txtTarjeta
			// 
			this.txtTarjeta.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsTomaFisica, "_TmpTomaFisica.Tarjeta"));
			this.txtTarjeta.Location = new System.Drawing.Point(72, 17);
			this.txtTarjeta.Name = "txtTarjeta";
			this.txtTarjeta.Size = new System.Drawing.Size(100, 22);
			this.txtTarjeta.TabIndex = 2;
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
			// cdvArticulo
			// 
			this.cdvArticulo.BindingContextCtrl = this;
			this.cdvArticulo.DataSet = this.cdsArticulo;
			this.cdvArticulo.Sort = "Articulo";
			this.cdvArticulo.TableName = "";
			this.cdvArticulo.TableViewName = "Articulo";
			// 
			// cdvArtCodigo
			// 
			this.cdvArtCodigo.BindingContextCtrl = this;
			this.cdvArtCodigo.DataSet = this.cdsArticulo;
			this.cdvArtCodigo.Sort = "Codigo";
			this.cdvArtCodigo.TableName = "";
			this.cdvArtCodigo.TableViewName = "Articulo";
			// 
			// cmbCodigo
			// 
			this.cmbCodigo.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbCodigo.DataSource = this.cdvArtCodigo;
			ultraGridColumn8.Header.VisiblePosition = 6;
			ultraGridColumn8.Hidden = true;
			ultraGridColumn9.Header.VisiblePosition = 3;
			ultraGridColumn9.Hidden = true;
			ultraGridColumn10.Header.VisiblePosition = 4;
			ultraGridColumn10.Hidden = true;
			ultraGridColumn11.Header.VisiblePosition = 7;
			ultraGridColumn11.Hidden = true;
			ultraGridColumn12.Header.VisiblePosition = 13;
			ultraGridColumn12.Hidden = true;
			ultraGridColumn13.Header.VisiblePosition = 9;
			ultraGridColumn13.Hidden = true;
			ultraGridColumn14.Header.VisiblePosition = 11;
			ultraGridColumn14.Hidden = true;
			ultraGridColumn15.Header.VisiblePosition = 14;
			ultraGridColumn15.Hidden = true;
			ultraGridColumn16.Header.VisiblePosition = 15;
			ultraGridColumn16.Hidden = true;
			ultraGridColumn17.Header.VisiblePosition = 16;
			ultraGridColumn17.Hidden = true;
			ultraGridColumn18.Header.VisiblePosition = 8;
			ultraGridColumn18.Hidden = true;
			ultraGridColumn19.Header.VisiblePosition = 17;
			ultraGridColumn19.Hidden = true;
			ultraGridColumn20.Header.VisiblePosition = 18;
			ultraGridColumn20.Hidden = true;
			ultraGridColumn21.Header.VisiblePosition = 19;
			ultraGridColumn21.Hidden = true;
			ultraGridColumn22.Header.VisiblePosition = 20;
			ultraGridColumn22.Hidden = true;
			ultraGridColumn23.Header.VisiblePosition = 21;
			ultraGridColumn23.Hidden = true;
			ultraGridColumn24.Header.VisiblePosition = 22;
			ultraGridColumn24.Hidden = true;
			ultraGridColumn25.Header.VisiblePosition = 2;
			ultraGridColumn26.Header.VisiblePosition = 23;
			ultraGridColumn26.Hidden = true;
			ultraGridColumn27.Header.VisiblePosition = 24;
			ultraGridColumn27.Hidden = true;
			ultraGridColumn28.Header.VisiblePosition = 25;
			ultraGridColumn28.Hidden = true;
			ultraGridColumn29.Header.VisiblePosition = 26;
			ultraGridColumn29.Hidden = true;
			ultraGridColumn30.Header.VisiblePosition = 27;
			ultraGridColumn30.Hidden = true;
			ultraGridColumn31.Header.VisiblePosition = 28;
			ultraGridColumn31.Hidden = true;
			ultraGridColumn32.Header.VisiblePosition = 29;
			ultraGridColumn32.Hidden = true;
			ultraGridColumn33.Header.VisiblePosition = 30;
			ultraGridColumn33.Hidden = true;
			ultraGridColumn34.Header.VisiblePosition = 31;
			ultraGridColumn34.Hidden = true;
			ultraGridColumn35.Header.VisiblePosition = 32;
			ultraGridColumn35.Hidden = true;
			ultraGridColumn36.Header.VisiblePosition = 33;
			ultraGridColumn36.Hidden = true;
			ultraGridColumn37.Header.VisiblePosition = 34;
			ultraGridColumn37.Hidden = true;
			ultraGridColumn38.Header.VisiblePosition = 35;
			ultraGridColumn38.Hidden = true;
			ultraGridColumn39.Header.VisiblePosition = 36;
			ultraGridColumn39.Hidden = true;
			ultraGridColumn40.Header.VisiblePosition = 37;
			ultraGridColumn40.Hidden = true;
			ultraGridColumn41.Header.VisiblePosition = 38;
			ultraGridColumn41.Hidden = true;
			ultraGridColumn42.Header.VisiblePosition = 39;
			ultraGridColumn42.Hidden = true;
			ultraGridColumn43.Header.VisiblePosition = 40;
			ultraGridColumn43.Hidden = true;
			ultraGridColumn44.Header.VisiblePosition = 41;
			ultraGridColumn44.Hidden = true;
			ultraGridColumn45.Header.VisiblePosition = 42;
			ultraGridColumn45.Hidden = true;
			ultraGridColumn46.Header.VisiblePosition = 43;
			ultraGridColumn46.Hidden = true;
			ultraGridColumn47.Header.VisiblePosition = 44;
			ultraGridColumn47.Hidden = true;
			ultraGridColumn48.Header.VisiblePosition = 45;
			ultraGridColumn48.Hidden = true;
			ultraGridColumn49.Header.VisiblePosition = 46;
			ultraGridColumn49.Hidden = true;
			ultraGridColumn50.Header.VisiblePosition = 47;
			ultraGridColumn50.Hidden = true;
			ultraGridColumn51.Header.VisiblePosition = 48;
			ultraGridColumn51.Hidden = true;
			ultraGridColumn52.Header.VisiblePosition = 49;
			ultraGridColumn52.Hidden = true;
			ultraGridColumn53.Header.VisiblePosition = 50;
			ultraGridColumn53.Hidden = true;
			ultraGridColumn54.Header.VisiblePosition = 51;
			ultraGridColumn54.Hidden = true;
			ultraGridColumn55.Header.VisiblePosition = 52;
			ultraGridColumn55.Hidden = true;
			ultraGridColumn56.Header.VisiblePosition = 53;
			ultraGridColumn56.Hidden = true;
			ultraGridColumn57.Header.VisiblePosition = 54;
			ultraGridColumn57.Hidden = true;
			ultraGridColumn58.Header.VisiblePosition = 55;
			ultraGridColumn58.Hidden = true;
			ultraGridColumn59.Header.VisiblePosition = 56;
			ultraGridColumn59.Hidden = true;
			ultraGridColumn60.Header.VisiblePosition = 57;
			ultraGridColumn60.Hidden = true;
			ultraGridColumn61.Header.VisiblePosition = 58;
			ultraGridColumn61.Hidden = true;
			ultraGridColumn62.Header.VisiblePosition = 59;
			ultraGridColumn62.Hidden = true;
			ultraGridColumn63.Header.VisiblePosition = 60;
			ultraGridColumn63.Hidden = true;
			ultraGridColumn64.Header.VisiblePosition = 61;
			ultraGridColumn64.Hidden = true;
			ultraGridColumn65.Header.VisiblePosition = 62;
			ultraGridColumn65.Hidden = true;
			ultraGridColumn66.Header.VisiblePosition = 63;
			ultraGridColumn66.Hidden = true;
			ultraGridColumn67.Header.VisiblePosition = 64;
			ultraGridColumn67.Hidden = true;
			ultraGridColumn68.Header.VisiblePosition = 65;
			ultraGridColumn68.Hidden = true;
			ultraGridColumn69.Header.VisiblePosition = 66;
			ultraGridColumn69.Hidden = true;
			ultraGridColumn70.Header.VisiblePosition = 67;
			ultraGridColumn70.Hidden = true;
			ultraGridColumn71.Header.VisiblePosition = 68;
			ultraGridColumn71.Hidden = true;
			ultraGridColumn72.Header.VisiblePosition = 69;
			ultraGridColumn72.Hidden = true;
			ultraGridColumn73.Header.VisiblePosition = 70;
			ultraGridColumn73.Hidden = true;
			ultraGridColumn74.Header.VisiblePosition = 71;
			ultraGridColumn74.Hidden = true;
			ultraGridColumn75.Header.VisiblePosition = 72;
			ultraGridColumn75.Hidden = true;
			ultraGridColumn76.Header.VisiblePosition = 73;
			ultraGridColumn76.Hidden = true;
			ultraGridColumn77.Header.VisiblePosition = 74;
			ultraGridColumn77.Hidden = true;
			ultraGridColumn78.Header.VisiblePosition = 75;
			ultraGridColumn78.Hidden = true;
			ultraGridColumn79.Header.VisiblePosition = 76;
			ultraGridColumn79.Hidden = true;
			ultraGridColumn80.Header.VisiblePosition = 77;
			ultraGridColumn80.Hidden = true;
			ultraGridColumn81.Header.VisiblePosition = 78;
			ultraGridColumn81.Hidden = true;
			ultraGridColumn82.Header.VisiblePosition = 79;
			ultraGridColumn82.Hidden = true;
			ultraGridColumn83.Header.VisiblePosition = 80;
			ultraGridColumn83.Hidden = true;
			ultraGridColumn84.Header.VisiblePosition = 81;
			ultraGridColumn84.Hidden = true;
			ultraGridColumn85.Header.VisiblePosition = 82;
			ultraGridColumn85.Hidden = true;
			ultraGridColumn86.Header.VisiblePosition = 83;
			ultraGridColumn86.Hidden = true;
			ultraGridColumn87.Header.VisiblePosition = 84;
			ultraGridColumn87.Hidden = true;
			ultraGridColumn88.Header.VisiblePosition = 85;
			ultraGridColumn88.Hidden = true;
			ultraGridColumn89.Header.VisiblePosition = 86;
			ultraGridColumn89.Hidden = true;
			ultraGridColumn90.Header.VisiblePosition = 5;
			ultraGridColumn90.Hidden = true;
			ultraGridColumn91.Header.VisiblePosition = 87;
			ultraGridColumn91.Hidden = true;
			ultraGridColumn92.Header.VisiblePosition = 88;
			ultraGridColumn92.Hidden = true;
			ultraGridColumn93.Header.VisiblePosition = 89;
			ultraGridColumn93.Hidden = true;
			ultraGridColumn94.Header.VisiblePosition = 90;
			ultraGridColumn94.Hidden = true;
			ultraGridColumn95.Header.VisiblePosition = 91;
			ultraGridColumn95.Hidden = true;
			ultraGridColumn96.Header.VisiblePosition = 92;
			ultraGridColumn96.Hidden = true;
			ultraGridColumn97.Header.VisiblePosition = 93;
			ultraGridColumn97.Hidden = true;
			ultraGridColumn98.Header.VisiblePosition = 94;
			ultraGridColumn98.Hidden = true;
			ultraGridColumn99.Header.VisiblePosition = 95;
			ultraGridColumn99.Hidden = true;
			ultraGridColumn100.Header.VisiblePosition = 96;
			ultraGridColumn100.Hidden = true;
			ultraGridColumn101.Header.VisiblePosition = 97;
			ultraGridColumn101.Hidden = true;
			ultraGridColumn102.Header.VisiblePosition = 98;
			ultraGridColumn102.Hidden = true;
			ultraGridColumn103.Header.VisiblePosition = 99;
			ultraGridColumn103.Hidden = true;
			ultraGridColumn104.Header.VisiblePosition = 100;
			ultraGridColumn104.Hidden = true;
			ultraGridColumn105.Header.VisiblePosition = 101;
			ultraGridColumn105.Hidden = true;
			ultraGridColumn106.Header.VisiblePosition = 10;
			ultraGridColumn106.Hidden = true;
			ultraGridColumn107.Header.VisiblePosition = 12;
			ultraGridColumn107.Hidden = true;
			ultraGridColumn108.Header.VisiblePosition = 102;
			ultraGridColumn108.Hidden = true;
			ultraGridColumn109.Header.VisiblePosition = 103;
			ultraGridColumn109.Hidden = true;
			ultraGridColumn110.Header.VisiblePosition = 1;
			ultraGridColumn110.Width = 220;
			ultraGridColumn111.Header.VisiblePosition = 0;
			ultraGridColumn111.Width = 118;
			ultraGridColumn112.Header.VisiblePosition = 104;
			ultraGridColumn112.Hidden = true;
			ultraGridColumn113.Header.VisiblePosition = 105;
			ultraGridColumn113.Hidden = true;
			ultraGridColumn114.Header.VisiblePosition = 106;
			ultraGridColumn114.Hidden = true;
			ultraGridColumn115.Header.VisiblePosition = 107;
			ultraGridColumn115.Hidden = true;
			ultraGridColumn116.Header.VisiblePosition = 108;
			ultraGridColumn116.Hidden = true;
			ultraGridBand2.Columns.AddRange(new object[] {
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
																										 ultraGridColumn116});
			this.cmbCodigo.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			this.cmbCodigo.DisplayMember = "Codigo";
			this.cmbCodigo.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbCodigo.Location = new System.Drawing.Point(104, 181);
			this.cmbCodigo.Name = "cmbCodigo";
			this.cmbCodigo.Size = new System.Drawing.Size(360, 103);
			this.cmbCodigo.TabIndex = 3;
			this.cmbCodigo.ValueMember = "idArticulo";
			this.cmbCodigo.Visible = false;
			// 
			// ultraLabel1
			// 
			this.ultraLabel1.AutoSize = true;
			this.ultraLabel1.Location = new System.Drawing.Point(24, 17);
			this.ultraLabel1.Name = "ultraLabel1";
			this.ultraLabel1.Size = new System.Drawing.Size(40, 15);
			this.ultraLabel1.TabIndex = 4;
			this.ultraLabel1.Text = "Tarjeta";
			// 
			// cmbFecha
			// 
			this.cmbFecha.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsTomaFisica, "_TmpTomaFisica.Fecha"));
			dateButton1.Caption = "Today";
			this.cmbFecha.DateButtons.Add(dateButton1);
			this.cmbFecha.Format = "dd/MMM/yyyy";
			this.cmbFecha.Location = new System.Drawing.Point(240, 17);
			this.cmbFecha.Name = "cmbFecha";
			this.cmbFecha.NonAutoSizeHeight = 23;
			this.cmbFecha.Size = new System.Drawing.Size(112, 21);
			this.cmbFecha.TabIndex = 5;
			this.cmbFecha.Value = new System.DateTime(2006, 7, 21, 0, 0, 0, 0);
			// 
			// ultraLabel2
			// 
			this.ultraLabel2.AutoSize = true;
			this.ultraLabel2.Location = new System.Drawing.Point(192, 17);
			this.ultraLabel2.Name = "ultraLabel2";
			this.ultraLabel2.Size = new System.Drawing.Size(34, 15);
			this.ultraLabel2.TabIndex = 6;
			this.ultraLabel2.Text = "Fecha";
			// 
			// ultraLabel3
			// 
			this.ultraLabel3.AutoSize = true;
			this.ultraLabel3.Location = new System.Drawing.Point(376, 17);
			this.ultraLabel3.Name = "ultraLabel3";
			this.ultraLabel3.Size = new System.Drawing.Size(41, 15);
			this.ultraLabel3.TabIndex = 7;
			this.ultraLabel3.Text = "Bodega";
			// 
			// txtBodega
			// 
			this.txtBodega.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsTomaFisica, "_TmpTomaFisica.Bodega"));
			this.txtBodega.Location = new System.Drawing.Point(440, 17);
			this.txtBodega.MaskInput = "nnn";
			this.txtBodega.MinValue = 0;
			this.txtBodega.Name = "txtBodega";
			this.txtBodega.PromptChar = ' ';
			this.txtBodega.Size = new System.Drawing.Size(56, 22);
			this.txtBodega.SpinButtonDisplayStyle = Infragistics.Win.ButtonDisplayStyle.Always;
			this.txtBodega.TabIndex = 8;
			// 
			// cmbArticulo
			// 
			this.cmbArticulo.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbArticulo.DataSource = this.cdvArticulo;
			ultraGridColumn117.Header.VisiblePosition = 6;
			ultraGridColumn117.Hidden = true;
			ultraGridColumn118.Header.VisiblePosition = 3;
			ultraGridColumn118.Hidden = true;
			ultraGridColumn119.Header.VisiblePosition = 4;
			ultraGridColumn119.Hidden = true;
			ultraGridColumn120.Header.VisiblePosition = 7;
			ultraGridColumn120.Hidden = true;
			ultraGridColumn121.Header.VisiblePosition = 13;
			ultraGridColumn121.Hidden = true;
			ultraGridColumn122.Header.VisiblePosition = 9;
			ultraGridColumn122.Hidden = true;
			ultraGridColumn123.Header.VisiblePosition = 11;
			ultraGridColumn123.Hidden = true;
			ultraGridColumn124.Header.VisiblePosition = 14;
			ultraGridColumn124.Hidden = true;
			ultraGridColumn125.Header.VisiblePosition = 15;
			ultraGridColumn125.Hidden = true;
			ultraGridColumn126.Header.VisiblePosition = 16;
			ultraGridColumn126.Hidden = true;
			ultraGridColumn127.Header.VisiblePosition = 8;
			ultraGridColumn127.Hidden = true;
			ultraGridColumn128.Header.VisiblePosition = 17;
			ultraGridColumn128.Hidden = true;
			ultraGridColumn129.Header.VisiblePosition = 18;
			ultraGridColumn129.Hidden = true;
			ultraGridColumn130.Header.VisiblePosition = 19;
			ultraGridColumn130.Hidden = true;
			ultraGridColumn131.Header.VisiblePosition = 20;
			ultraGridColumn131.Hidden = true;
			ultraGridColumn132.Header.VisiblePosition = 21;
			ultraGridColumn132.Hidden = true;
			ultraGridColumn133.Header.VisiblePosition = 22;
			ultraGridColumn133.Hidden = true;
			ultraGridColumn134.Header.VisiblePosition = 2;
			ultraGridColumn135.Header.VisiblePosition = 23;
			ultraGridColumn135.Hidden = true;
			ultraGridColumn136.Header.VisiblePosition = 24;
			ultraGridColumn136.Hidden = true;
			ultraGridColumn137.Header.VisiblePosition = 25;
			ultraGridColumn137.Hidden = true;
			ultraGridColumn138.Header.VisiblePosition = 26;
			ultraGridColumn138.Hidden = true;
			ultraGridColumn139.Header.VisiblePosition = 27;
			ultraGridColumn139.Hidden = true;
			ultraGridColumn140.Header.VisiblePosition = 28;
			ultraGridColumn140.Hidden = true;
			ultraGridColumn141.Header.VisiblePosition = 29;
			ultraGridColumn141.Hidden = true;
			ultraGridColumn142.Header.VisiblePosition = 30;
			ultraGridColumn142.Hidden = true;
			ultraGridColumn143.Header.VisiblePosition = 31;
			ultraGridColumn143.Hidden = true;
			ultraGridColumn144.Header.VisiblePosition = 32;
			ultraGridColumn144.Hidden = true;
			ultraGridColumn145.Header.VisiblePosition = 33;
			ultraGridColumn145.Hidden = true;
			ultraGridColumn146.Header.VisiblePosition = 34;
			ultraGridColumn146.Hidden = true;
			ultraGridColumn147.Header.VisiblePosition = 35;
			ultraGridColumn147.Hidden = true;
			ultraGridColumn148.Header.VisiblePosition = 36;
			ultraGridColumn148.Hidden = true;
			ultraGridColumn149.Header.VisiblePosition = 37;
			ultraGridColumn149.Hidden = true;
			ultraGridColumn150.Header.VisiblePosition = 38;
			ultraGridColumn150.Hidden = true;
			ultraGridColumn151.Header.VisiblePosition = 39;
			ultraGridColumn151.Hidden = true;
			ultraGridColumn152.Header.VisiblePosition = 40;
			ultraGridColumn152.Hidden = true;
			ultraGridColumn153.Header.VisiblePosition = 41;
			ultraGridColumn153.Hidden = true;
			ultraGridColumn154.Header.VisiblePosition = 42;
			ultraGridColumn154.Hidden = true;
			ultraGridColumn155.Header.VisiblePosition = 43;
			ultraGridColumn155.Hidden = true;
			ultraGridColumn156.Header.VisiblePosition = 44;
			ultraGridColumn156.Hidden = true;
			ultraGridColumn157.Header.VisiblePosition = 45;
			ultraGridColumn157.Hidden = true;
			ultraGridColumn158.Header.VisiblePosition = 46;
			ultraGridColumn158.Hidden = true;
			ultraGridColumn159.Header.VisiblePosition = 47;
			ultraGridColumn159.Hidden = true;
			ultraGridColumn160.Header.VisiblePosition = 48;
			ultraGridColumn160.Hidden = true;
			ultraGridColumn161.Header.VisiblePosition = 49;
			ultraGridColumn161.Hidden = true;
			ultraGridColumn162.Header.VisiblePosition = 50;
			ultraGridColumn162.Hidden = true;
			ultraGridColumn163.Header.VisiblePosition = 51;
			ultraGridColumn163.Hidden = true;
			ultraGridColumn164.Header.VisiblePosition = 52;
			ultraGridColumn164.Hidden = true;
			ultraGridColumn165.Header.VisiblePosition = 53;
			ultraGridColumn165.Hidden = true;
			ultraGridColumn166.Header.VisiblePosition = 54;
			ultraGridColumn166.Hidden = true;
			ultraGridColumn167.Header.VisiblePosition = 55;
			ultraGridColumn167.Hidden = true;
			ultraGridColumn168.Header.VisiblePosition = 56;
			ultraGridColumn168.Hidden = true;
			ultraGridColumn169.Header.VisiblePosition = 57;
			ultraGridColumn169.Hidden = true;
			ultraGridColumn170.Header.VisiblePosition = 58;
			ultraGridColumn170.Hidden = true;
			ultraGridColumn171.Header.VisiblePosition = 59;
			ultraGridColumn171.Hidden = true;
			ultraGridColumn172.Header.VisiblePosition = 60;
			ultraGridColumn172.Hidden = true;
			ultraGridColumn173.Header.VisiblePosition = 61;
			ultraGridColumn173.Hidden = true;
			ultraGridColumn174.Header.VisiblePosition = 62;
			ultraGridColumn174.Hidden = true;
			ultraGridColumn175.Header.VisiblePosition = 63;
			ultraGridColumn175.Hidden = true;
			ultraGridColumn176.Header.VisiblePosition = 64;
			ultraGridColumn176.Hidden = true;
			ultraGridColumn177.Header.VisiblePosition = 65;
			ultraGridColumn177.Hidden = true;
			ultraGridColumn178.Header.VisiblePosition = 66;
			ultraGridColumn178.Hidden = true;
			ultraGridColumn179.Header.VisiblePosition = 67;
			ultraGridColumn179.Hidden = true;
			ultraGridColumn180.Header.VisiblePosition = 68;
			ultraGridColumn180.Hidden = true;
			ultraGridColumn181.Header.VisiblePosition = 69;
			ultraGridColumn181.Hidden = true;
			ultraGridColumn182.Header.VisiblePosition = 70;
			ultraGridColumn182.Hidden = true;
			ultraGridColumn183.Header.VisiblePosition = 71;
			ultraGridColumn183.Hidden = true;
			ultraGridColumn184.Header.VisiblePosition = 72;
			ultraGridColumn184.Hidden = true;
			ultraGridColumn185.Header.VisiblePosition = 73;
			ultraGridColumn185.Hidden = true;
			ultraGridColumn186.Header.VisiblePosition = 74;
			ultraGridColumn186.Hidden = true;
			ultraGridColumn187.Header.VisiblePosition = 75;
			ultraGridColumn187.Hidden = true;
			ultraGridColumn188.Header.VisiblePosition = 76;
			ultraGridColumn188.Hidden = true;
			ultraGridColumn189.Header.VisiblePosition = 77;
			ultraGridColumn189.Hidden = true;
			ultraGridColumn190.Header.VisiblePosition = 78;
			ultraGridColumn190.Hidden = true;
			ultraGridColumn191.Header.VisiblePosition = 79;
			ultraGridColumn191.Hidden = true;
			ultraGridColumn192.Header.VisiblePosition = 80;
			ultraGridColumn192.Hidden = true;
			ultraGridColumn193.Header.VisiblePosition = 81;
			ultraGridColumn193.Hidden = true;
			ultraGridColumn194.Header.VisiblePosition = 82;
			ultraGridColumn194.Hidden = true;
			ultraGridColumn195.Header.VisiblePosition = 83;
			ultraGridColumn195.Hidden = true;
			ultraGridColumn196.Header.VisiblePosition = 84;
			ultraGridColumn196.Hidden = true;
			ultraGridColumn197.Header.VisiblePosition = 85;
			ultraGridColumn197.Hidden = true;
			ultraGridColumn198.Header.VisiblePosition = 86;
			ultraGridColumn198.Hidden = true;
			ultraGridColumn199.Header.VisiblePosition = 5;
			ultraGridColumn199.Hidden = true;
			ultraGridColumn200.Header.VisiblePosition = 87;
			ultraGridColumn200.Hidden = true;
			ultraGridColumn201.Header.VisiblePosition = 88;
			ultraGridColumn201.Hidden = true;
			ultraGridColumn202.Header.VisiblePosition = 89;
			ultraGridColumn202.Hidden = true;
			ultraGridColumn203.Header.VisiblePosition = 90;
			ultraGridColumn203.Hidden = true;
			ultraGridColumn204.Header.VisiblePosition = 91;
			ultraGridColumn204.Hidden = true;
			ultraGridColumn205.Header.VisiblePosition = 92;
			ultraGridColumn205.Hidden = true;
			ultraGridColumn206.Header.VisiblePosition = 93;
			ultraGridColumn206.Hidden = true;
			ultraGridColumn207.Header.VisiblePosition = 94;
			ultraGridColumn207.Hidden = true;
			ultraGridColumn208.Header.VisiblePosition = 95;
			ultraGridColumn208.Hidden = true;
			ultraGridColumn209.Header.VisiblePosition = 96;
			ultraGridColumn209.Hidden = true;
			ultraGridColumn210.Header.VisiblePosition = 97;
			ultraGridColumn210.Hidden = true;
			ultraGridColumn211.Header.VisiblePosition = 98;
			ultraGridColumn211.Hidden = true;
			ultraGridColumn212.Header.VisiblePosition = 99;
			ultraGridColumn212.Hidden = true;
			ultraGridColumn213.Header.VisiblePosition = 100;
			ultraGridColumn213.Hidden = true;
			ultraGridColumn214.Header.VisiblePosition = 101;
			ultraGridColumn214.Hidden = true;
			ultraGridColumn215.Header.VisiblePosition = 10;
			ultraGridColumn215.Hidden = true;
			ultraGridColumn216.Header.VisiblePosition = 12;
			ultraGridColumn216.Hidden = true;
			ultraGridColumn217.Header.VisiblePosition = 102;
			ultraGridColumn217.Hidden = true;
			ultraGridColumn218.Header.VisiblePosition = 103;
			ultraGridColumn218.Hidden = true;
			ultraGridColumn219.Header.VisiblePosition = 0;
			ultraGridColumn219.Width = 246;
			ultraGridColumn220.Header.VisiblePosition = 1;
			ultraGridColumn221.Header.VisiblePosition = 104;
			ultraGridColumn221.Hidden = true;
			ultraGridColumn222.Header.VisiblePosition = 105;
			ultraGridColumn222.Hidden = true;
			ultraGridColumn223.Header.VisiblePosition = 106;
			ultraGridColumn223.Hidden = true;
			ultraGridColumn224.Header.VisiblePosition = 107;
			ultraGridColumn224.Hidden = true;
			ultraGridColumn225.Header.VisiblePosition = 108;
			ultraGridColumn225.Hidden = true;
			ultraGridBand3.Columns.AddRange(new object[] {
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
																										 ultraGridColumn190,
																										 ultraGridColumn191,
																										 ultraGridColumn192,
																										 ultraGridColumn193,
																										 ultraGridColumn194,
																										 ultraGridColumn195,
																										 ultraGridColumn196,
																										 ultraGridColumn197,
																										 ultraGridColumn198,
																										 ultraGridColumn199,
																										 ultraGridColumn200,
																										 ultraGridColumn201,
																										 ultraGridColumn202,
																										 ultraGridColumn203,
																										 ultraGridColumn204,
																										 ultraGridColumn205,
																										 ultraGridColumn206,
																										 ultraGridColumn207,
																										 ultraGridColumn208,
																										 ultraGridColumn209,
																										 ultraGridColumn210,
																										 ultraGridColumn211,
																										 ultraGridColumn212,
																										 ultraGridColumn213,
																										 ultraGridColumn214,
																										 ultraGridColumn215,
																										 ultraGridColumn216,
																										 ultraGridColumn217,
																										 ultraGridColumn218,
																										 ultraGridColumn219,
																										 ultraGridColumn220,
																										 ultraGridColumn221,
																										 ultraGridColumn222,
																										 ultraGridColumn223,
																										 ultraGridColumn224,
																										 ultraGridColumn225});
			this.cmbArticulo.DisplayLayout.BandsSerializer.Add(ultraGridBand3);
			this.cmbArticulo.DisplayMember = "Articulo";
			this.cmbArticulo.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbArticulo.Location = new System.Drawing.Point(168, 95);
			this.cmbArticulo.Name = "cmbArticulo";
			this.cmbArticulo.Size = new System.Drawing.Size(304, 95);
			this.cmbArticulo.TabIndex = 9;
			this.cmbArticulo.ValueMember = "idArticulo";
			this.cmbArticulo.Visible = false;
			// 
			// txtIdToma
			// 
			this.txtIdToma.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsTomaFisica, "_TmpTomaFisica.idTmpTomaFisica"));
			this.txtIdToma.Location = new System.Drawing.Point(392, 319);
			this.txtIdToma.Name = "txtIdToma";
			this.txtIdToma.PromptChar = ' ';
			this.txtIdToma.Size = new System.Drawing.Size(100, 22);
			this.txtIdToma.TabIndex = 10;
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
			this.cdsSeguridad.BeforeFill += new C1.Data.FillEventHandler(this.cdsSeguridad_BeforeFill);
			// 
			// TomaFisica
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 14);
			this.ClientSize = new System.Drawing.Size(528, 359);
			this.Controls.Add(this.txtIdToma);
			this.Controls.Add(this.cmbArticulo);
			this.Controls.Add(this.txtBodega);
			this.Controls.Add(this.ultraLabel3);
			this.Controls.Add(this.ultraLabel2);
			this.Controls.Add(this.cmbFecha);
			this.Controls.Add(this.ultraLabel1);
			this.Controls.Add(this.cmbCodigo);
			this.Controls.Add(this.txtTarjeta);
			this.Controls.Add(this.barraDatoSQL1);
			this.Controls.Add(this.ultraGrid1);
			this.Name = "TomaFisica";
			this.Text = "Toma Física";
			this.Closing += new System.ComponentModel.CancelEventHandler(this.TomaFisica_Closing);
			this.Load += new System.EventHandler(this.TomaFisica_Load);
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsTomaFisica)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtTarjeta)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsArticulo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvArticulo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvArtCodigo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCodigo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbFecha)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtBodega)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbArticulo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdToma)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeguridad)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		C1.Data.C1DataRow drToma;
		private void barraDatoSQL1_Refresca(object sender, System.EventArgs e)
		{
			try
			{
				drToma = cdsTomaFisica.TableViews["TmpTomaFisica"].Rows[0];
			}
			catch{};
		}

		private void barraDatoSQL1_Borrar(object sender, System.EventArgs e)
		{
			if (!miAcceso.Eliminar)
			{
				MessageBox.Show("No tiene acceso a Borrar " + this.Text.Trim(),
					"Seguridad", MessageBoxButtons.OK, MessageBoxIcon.Information);
				return;
			}
			if (DialogResult.No == MessageBox.Show(
				"¿Está seguro que desea Anular el Registro?", 
				"Confirmación", MessageBoxButtons.YesNo, MessageBoxIcon.Warning, 
				MessageBoxDefaultButton.Button2)) return;

			int IdTomaF = int.Parse(txtIdToma.Value.ToString());
			string stBorrar = "Delete TmpTomaDetalle Where idTmpTomaFisica = " + IdTomaF.ToString();
			Funcion.EjecutaSQL(cdsTomaFisica, stBorrar, true);
			barraDatoSQL1.BorraRegistro();
		}

		private void barraDatoSQL1_Buscar(object sender, System.EventArgs e)
		{
			using (BuscaTomaFisica miBusqueda = new BuscaTomaFisica())
			{
				if (DialogResult.OK == miBusqueda.ShowDialog())
				{
					if (miBusqueda.idBusca.Text.Trim().Length > 0)
					{
						barraDatoSQL1.IdRegistro = int.Parse(miBusqueda.idBusca.Text);
						barraDatoSQL1.ProximoId(5);
					}
				}
			}
		}

		private void barraDatoSQL1_Crear(object sender, System.EventArgs e)
		{
			if (!miAcceso.Nuevo)
			{
				MessageBox.Show("No tiene Acceso a Crear Registro",
					"Seguridad", MessageBoxButtons.OK, MessageBoxIcon.Information);
				return;
			}

			DateTime dtFecha = DateTime.Today;
			int iBodega = 1;
			if (cmbFecha.Value != System.DBNull.Value)
				dtFecha= (DateTime) cmbFecha.Value;
			if (txtBodega.Value != System.DBNull.Value)
				iBodega = (int) txtBodega.Value;
			barraDatoSQL1.CrearRegistro();
			drToma["Fecha"] = dtFecha;
			drToma["Bodega"] = iBodega;
		}

		private void barraDatoSQL1_Deshacer(object sender, System.EventArgs e)
		{
			barraDatoSQL1.DeshacerRegistro();
		}

		private void barraDatoSQL1_Editar(object sender, System.EventArgs e)
		{
			if (!miAcceso.Editar)
			{
				MessageBox.Show("No tiene Acceso a Editar Registro",
					"Seguridad", MessageBoxButtons.OK, MessageBoxIcon.Information);
				return;
			}

			barraDatoSQL1.EditarRegistro();
		}

		private void barraDatoSQL1_Grabar(object sender, System.EventArgs e)
		{
			barraDatoSQL1.GrabaRegistro();
		}

		private void barraDatoSQL1_Imprimir(object sender, System.EventArgs e)
		{
			if (!miAcceso.Imprimir)
			{
				MessageBox.Show("No tiene Acceso a Imprimir ",
					"Seguridad", MessageBoxButtons.OK, MessageBoxIcon.Information);
				return;
			}
			RepTomaFisica miRep = new RepTomaFisica();
			miRep.MdiParent = this.MdiParent;
			miRep.Show();
		}

		private void barraDatoSQL1_Impresora(object sender, System.EventArgs e)
		{
		
		}

		private void barraDatoSQL1_Restaurar(object sender, System.EventArgs e)
		{
		
		}

		private void cdsTomaFisica_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			this.cdsTomaFisica.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);
		}

		private Acceso miAcceso;
		private void TomaFisica_Load(object sender, System.EventArgs e)
		{
			string stAudita = "Exec AuditaCrear 53, 6, 'Toma Fisica'";
			Funcion.EjecutaSQL(cdsArticulo, stAudita, true);
			miAcceso =new Acceso(cdsSeguridad, "04090LA");

			string stExec = "Exec ArticuloCargaNombre 5";
			cmbArticulo.DataSource = Funcion.dvProcedimiento(cdsArticulo, stExec);
			cmbArticulo.DisplayLayout.Bands[0].Columns["idArticulo"].Hidden = true;

			stExec = "Exec ArticuloCargaCodigo 5";
			cmbCodigo.DataSource = Funcion.dvProcedimiento(cdsArticulo, stExec);
			cmbCodigo.DisplayLayout.Bands[0].Columns["idArticulo"].Hidden = true;

			txtIdToma.Width = 0;
			barraDatoSQL1.PosUltima();
			barraDatoSQL1.HabilitaControles(false);
		}

		private void ultraGrid1_AfterCellUpdate(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			if (e.Cell.IsInEditMode)
				if (e.Cell.Column.ToString() == "Codigo")
					if (e.Cell.OriginalValue != e.Cell.Value)
					{
						e.Cell.Row.Cells["idArticulo"].Value = e.Cell.Row.Cells["Codigo"].Value;
					}
			if (e.Cell.IsInEditMode)
				if (e.Cell.Column.ToString() == "idArticulo")
					if (e.Cell.OriginalValue != e.Cell.Value)
					{
						e.Cell.Row.Cells["Codigo"].Value = e.Cell.Row.Cells["idArticulo"].Value;
					}
		}

		private void cdsArticulo_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsArticulo.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);		
		}

		private void cdsSeguridad_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeguridad.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);
		}

		private void ultraGrid1_InitializeRow(object sender, Infragistics.Win.UltraWinGrid.InitializeRowEventArgs e)
		{
	
		}

		private void TomaFisica_Closing(object sender, System.ComponentModel.CancelEventArgs e)
		{
			string stAudita = "Exec AuditaCrear 53, 7, 'Toma Fisica'";
			Funcion.EjecutaSQL(cdsArticulo, stAudita, true);
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
						ultraGrid1.PerformAction(UltraGridAction.BelowCell, false, false);
						e.Handled = true;
						ultraGrid1.PerformAction(UltraGridAction.EnterEditMode, false, false);
						break;
				}
			}
			catch(System.Exception ex)
			{
				Console.WriteLine(ex.Message);
				//MessageBox.Show(ex.Message, "Tecla");
			}

		}

		private void ultraGrid1_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}

		private void ultraGrid1_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (Funcion.bEjecutaSQL(cdsArticulo, "Exec SeteoGExiste 'FTRCxG'")) // cambia la comilla por guion				
			{		
				try
				{
					if(e.KeyChar == 39 )//&& e.KeyChar = guion)
					{						
						e.Handled = true;
						SendKeys.Send("-");											
					}
				}
				catch(Exception ex)
				{
					MessageBox.Show(ex.Message);
				}
			}

		}


	}
}
