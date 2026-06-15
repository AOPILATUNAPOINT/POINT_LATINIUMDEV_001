using System;
using System.IO;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de RepTomaFisica.
	/// </summary>
	public class RepTomaFisica : DevExpress.XtraEditors.XtraForm
	{
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkTodaBodega;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtBodega;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtTarjeta;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbArticulo;
		private C1.Data.C1DataSet cdsArticulo;
		private C1.Data.C1DataSet cdsGrupoArt;
		private C1.Data.C1DataView cdvArtCodigo;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbGrupo;
		private Infragistics.Win.Misc.UltraButton btReporte;
		private Infragistics.Win.Misc.UltraButton btSalir;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtReporte;
		private Infragistics.Win.Misc.UltraLabel ultraLabel1;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkTodaTarjeta;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkTodoArticulo;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkTodoGrupo;
		private System.Windows.Forms.ErrorProvider errorProvider;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkResumen;
		private C1.Data.C1DataView cdvGrupoArt;
		private System.Windows.Forms.ToolTip toolTip1;
		private Infragistics.Win.Misc.UltraButton btAjuste;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo cmbFecha;
		private Infragistics.Win.Misc.UltraLabel lblFecha;
		private Infragistics.Win.Misc.UltraButton ultraButton1;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbTipo;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkTipo;
		private C1.Data.C1DataView cdvTipo;
		private C1.Data.C1DataSet cdsTipo;
		private System.ComponentModel.IContainer components;

		public RepTomaFisica()
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
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Articulo", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idArticulo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idArticuloSRI");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idColor");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idModelo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idGrupoArticulo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idSubGrupo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idArancel");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idMarca");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idUnidad");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPromocion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn11 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idTipoGrupo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn12 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idClase");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn13 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idSeccion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn14 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idSucursal");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn15 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPadre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn16 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaIngreso");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn17 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CodAlterno");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn18 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Saldo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn19 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Peso");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn20 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Caracteristica");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn21 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Aplicacion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn22 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Iva");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn23 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Ice");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn24 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Precio1");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn25 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Precio2");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn26 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Precio3");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn27 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Precio4");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn28 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Precio5");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn29 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Precio6");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn30 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CantidadMinima");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn31 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CantidadMaxima");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn32 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Residual");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn33 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CostoUltimo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn34 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaUltimo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn35 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCompra1");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn36 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCompra2");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn37 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCompra3");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn38 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CostoPromedio");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn39 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Gramos");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn40 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ConsumoProm");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn41 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("TiempoReposicion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn42 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("UnidadPeso");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn43 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Depreciacion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn44 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Incr1");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn45 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Incr2");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn46 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Incr3");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn47 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Incr4");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn48 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Incr5");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn49 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Incr6");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn50 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Ubicacion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn51 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("NumeroParte");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn52 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ComisionValor");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn53 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ComisionPorcentaje");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn54 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CantidadRequerida");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn55 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("PrecioPromocion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn56 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CantidadPromocion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn57 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("DescuentoContado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn58 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("DescuentoCredito");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn59 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("DescPorcContado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn60 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("DescPorcCre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn61 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CantReqDesc");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn62 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CantReqPorc");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn63 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("TipoPromocion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn64 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Responsable");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn65 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Comision30");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn66 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Comision60");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn67 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Comision90");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn68 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Comision120");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn69 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bodega1");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn70 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bodega2");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn71 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bodega3");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn72 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bodega4");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn73 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bodega5");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn74 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bodega6");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn75 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bodega7");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn76 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bodega8");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn77 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bodega9");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn78 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("SaldoT");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn79 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Alto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn80 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Ancho");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn81 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Largo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn82 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Color");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn83 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Descontinuado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn84 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Borrar");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn85 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Habilitado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn86 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ValorPeso");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn87 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("BodegaPredef");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn88 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Tiempo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn89 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cantidad1");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn90 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cantidad2");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn91 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cantidad3");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn92 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cantidad4");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn93 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cantidad5");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn94 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("PaqueteUnidad");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn95 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Editar");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn96 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nuevo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn97 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaEditado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn98 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CambiarPrecio");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn99 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("dCompra");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn100 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("dVenta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn101 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Foto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn102 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Notas");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn103 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Articulo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn104 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn105 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("BodegaPredefCR");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn106 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CodPromocion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn107 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CostoFecha");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn108 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Aportable");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn109 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ExigeLote");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("ArticuloGrupo", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn110 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idGrupoArticulo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn111 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idTipoGrupo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn112 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Grupo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn113 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCuentas2");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn114 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCuentas3");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn115 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCuentas4");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn116 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Borrar");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn117 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCuentas5");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn118 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCuentas6");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn119 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idArticuloSRI");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn120 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn121 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCuentas7");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn122 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCuentas8");
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton1 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand3 = new Infragistics.Win.UltraWinGrid.UltraGridBand("ArticuloTipo", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn123 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idTipoGrupo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn124 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Tipo");
			this.chkTodaBodega = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.txtBodega = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.chkTodaTarjeta = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.txtTarjeta = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.chkTodoArticulo = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.cmbArticulo = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.cdvArtCodigo = new C1.Data.C1DataView();
			this.cdsArticulo = new C1.Data.C1DataSet();
			this.cdsGrupoArt = new C1.Data.C1DataSet();
			this.cmbGrupo = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.cdvGrupoArt = new C1.Data.C1DataView();
			this.chkTodoGrupo = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.btReporte = new Infragistics.Win.Misc.UltraButton();
			this.btSalir = new Infragistics.Win.Misc.UltraButton();
			this.txtReporte = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.ultraLabel1 = new Infragistics.Win.Misc.UltraLabel();
			this.errorProvider = new System.Windows.Forms.ErrorProvider();
			this.chkResumen = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.toolTip1 = new System.Windows.Forms.ToolTip(this.components);
			this.ultraButton1 = new Infragistics.Win.Misc.UltraButton();
			this.btAjuste = new Infragistics.Win.Misc.UltraButton();
			this.cmbFecha = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.lblFecha = new Infragistics.Win.Misc.UltraLabel();
			this.cmbTipo = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.cdvTipo = new C1.Data.C1DataView();
			this.cdsTipo = new C1.Data.C1DataSet();
			this.chkTipo = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			((System.ComponentModel.ISupportInitialize)(this.txtBodega)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtTarjeta)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbArticulo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvArtCodigo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsArticulo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsGrupoArt)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbGrupo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvGrupoArt)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtReporte)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbFecha)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbTipo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvTipo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsTipo)).BeginInit();
			this.SuspendLayout();
			// 
			// chkTodaBodega
			// 
			this.chkTodaBodega.Checked = true;
			this.chkTodaBodega.CheckState = System.Windows.Forms.CheckState.Checked;
			this.chkTodaBodega.Location = new System.Drawing.Point(32, 26);
			this.chkTodaBodega.Name = "chkTodaBodega";
			this.chkTodaBodega.Size = new System.Drawing.Size(96, 21);
			this.chkTodaBodega.TabIndex = 0;
			this.chkTodaBodega.Text = "Toda Bodega";
			// 
			// txtBodega
			// 
			this.txtBodega.Location = new System.Drawing.Point(160, 26);
			this.txtBodega.MaskInput = "nnn";
			this.txtBodega.MinValue = 0;
			this.txtBodega.Name = "txtBodega";
			this.txtBodega.PromptChar = ' ';
			this.txtBodega.Size = new System.Drawing.Size(100, 22);
			this.txtBodega.SpinButtonDisplayStyle = Infragistics.Win.ButtonDisplayStyle.Always;
			this.txtBodega.TabIndex = 1;
			this.txtBodega.Value = 1;
			// 
			// chkTodaTarjeta
			// 
			this.chkTodaTarjeta.Checked = true;
			this.chkTodaTarjeta.CheckState = System.Windows.Forms.CheckState.Checked;
			this.chkTodaTarjeta.Location = new System.Drawing.Point(32, 60);
			this.chkTodaTarjeta.Name = "chkTodaTarjeta";
			this.chkTodaTarjeta.Size = new System.Drawing.Size(96, 22);
			this.chkTodaTarjeta.TabIndex = 2;
			this.chkTodaTarjeta.Text = "Toda Tarjeta";
			// 
			// txtTarjeta
			// 
			this.txtTarjeta.Location = new System.Drawing.Point(160, 60);
			this.txtTarjeta.Name = "txtTarjeta";
			this.txtTarjeta.Size = new System.Drawing.Size(100, 22);
			this.txtTarjeta.TabIndex = 3;
			// 
			// chkTodoArticulo
			// 
			this.chkTodoArticulo.Checked = true;
			this.chkTodoArticulo.CheckState = System.Windows.Forms.CheckState.Checked;
			this.chkTodoArticulo.Location = new System.Drawing.Point(32, 95);
			this.chkTodoArticulo.Name = "chkTodoArticulo";
			this.chkTodoArticulo.Size = new System.Drawing.Size(96, 21);
			this.chkTodoArticulo.TabIndex = 4;
			this.chkTodoArticulo.Text = "Todo Articulo";
			// 
			// cmbArticulo
			// 
			this.cmbArticulo.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbArticulo.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbArticulo.DataSource = this.cdvArtCodigo;
			ultraGridColumn1.Header.VisiblePosition = 44;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn2.Header.VisiblePosition = 100;
			ultraGridColumn2.Hidden = true;
			ultraGridColumn3.Header.VisiblePosition = 93;
			ultraGridColumn3.Hidden = true;
			ultraGridColumn4.Header.VisiblePosition = 94;
			ultraGridColumn4.Hidden = true;
			ultraGridColumn5.Header.VisiblePosition = 46;
			ultraGridColumn5.Hidden = true;
			ultraGridColumn6.Header.VisiblePosition = 50;
			ultraGridColumn6.Hidden = true;
			ultraGridColumn7.Header.VisiblePosition = 43;
			ultraGridColumn7.Hidden = true;
			ultraGridColumn8.Header.VisiblePosition = 47;
			ultraGridColumn8.Hidden = true;
			ultraGridColumn9.Header.VisiblePosition = 52;
			ultraGridColumn9.Hidden = true;
			ultraGridColumn10.Header.VisiblePosition = 48;
			ultraGridColumn10.Hidden = true;
			ultraGridColumn11.Header.VisiblePosition = 51;
			ultraGridColumn11.Hidden = true;
			ultraGridColumn12.Header.VisiblePosition = 45;
			ultraGridColumn12.Hidden = true;
			ultraGridColumn13.Header.VisiblePosition = 49;
			ultraGridColumn13.Hidden = true;
			ultraGridColumn14.Header.VisiblePosition = 91;
			ultraGridColumn14.Hidden = true;
			ultraGridColumn15.Header.VisiblePosition = 99;
			ultraGridColumn15.Hidden = true;
			ultraGridColumn16.Header.VisiblePosition = 84;
			ultraGridColumn16.Hidden = true;
			ultraGridColumn17.Header.VisiblePosition = 23;
			ultraGridColumn17.Hidden = true;
			ultraGridColumn18.Header.VisiblePosition = 72;
			ultraGridColumn18.Hidden = true;
			ultraGridColumn19.Header.VisiblePosition = 62;
			ultraGridColumn19.Hidden = true;
			ultraGridColumn20.Header.VisiblePosition = 22;
			ultraGridColumn20.Hidden = true;
			ultraGridColumn21.Header.VisiblePosition = 4;
			ultraGridColumn21.Hidden = true;
			ultraGridColumn22.Header.VisiblePosition = 59;
			ultraGridColumn22.Hidden = true;
			ultraGridColumn23.Header.VisiblePosition = 42;
			ultraGridColumn23.Hidden = true;
			ultraGridColumn24.Header.VisiblePosition = 63;
			ultraGridColumn24.Hidden = true;
			ultraGridColumn25.Header.VisiblePosition = 64;
			ultraGridColumn25.Hidden = true;
			ultraGridColumn26.Header.VisiblePosition = 65;
			ultraGridColumn26.Hidden = true;
			ultraGridColumn27.Header.VisiblePosition = 66;
			ultraGridColumn27.Hidden = true;
			ultraGridColumn28.Header.VisiblePosition = 67;
			ultraGridColumn28.Hidden = true;
			ultraGridColumn29.Header.VisiblePosition = 68;
			ultraGridColumn29.Hidden = true;
			ultraGridColumn30.Header.VisiblePosition = 17;
			ultraGridColumn30.Hidden = true;
			ultraGridColumn31.Header.VisiblePosition = 16;
			ultraGridColumn31.Hidden = true;
			ultraGridColumn32.Header.VisiblePosition = 70;
			ultraGridColumn32.Hidden = true;
			ultraGridColumn33.Header.VisiblePosition = 33;
			ultraGridColumn33.Hidden = true;
			ultraGridColumn34.Header.VisiblePosition = 83;
			ultraGridColumn34.Hidden = true;
			ultraGridColumn35.Header.VisiblePosition = 80;
			ultraGridColumn35.Hidden = true;
			ultraGridColumn36.Header.VisiblePosition = 81;
			ultraGridColumn36.Hidden = true;
			ultraGridColumn37.Header.VisiblePosition = 82;
			ultraGridColumn37.Hidden = true;
			ultraGridColumn38.Header.VisiblePosition = 32;
			ultraGridColumn38.Hidden = true;
			ultraGridColumn39.Header.VisiblePosition = 40;
			ultraGridColumn39.Hidden = true;
			ultraGridColumn40.Header.VisiblePosition = 31;
			ultraGridColumn40.Hidden = true;
			ultraGridColumn41.Header.VisiblePosition = 75;
			ultraGridColumn41.Hidden = true;
			ultraGridColumn42.Header.VisiblePosition = 78;
			ultraGridColumn42.Hidden = true;
			ultraGridColumn43.Header.VisiblePosition = 34;
			ultraGridColumn43.Hidden = true;
			ultraGridColumn44.Header.VisiblePosition = 53;
			ultraGridColumn44.Hidden = true;
			ultraGridColumn45.Header.VisiblePosition = 54;
			ultraGridColumn45.Hidden = true;
			ultraGridColumn46.Header.VisiblePosition = 55;
			ultraGridColumn46.Hidden = true;
			ultraGridColumn47.Header.VisiblePosition = 56;
			ultraGridColumn47.Hidden = true;
			ultraGridColumn48.Header.VisiblePosition = 57;
			ultraGridColumn48.Hidden = true;
			ultraGridColumn49.Header.VisiblePosition = 58;
			ultraGridColumn49.Hidden = true;
			ultraGridColumn50.Header.VisiblePosition = 77;
			ultraGridColumn50.Hidden = true;
			ultraGridColumn51.Header.VisiblePosition = 61;
			ultraGridColumn51.Hidden = true;
			ultraGridColumn52.Header.VisiblePosition = 30;
			ultraGridColumn52.Hidden = true;
			ultraGridColumn53.Header.VisiblePosition = 29;
			ultraGridColumn53.Hidden = true;
			ultraGridColumn54.Header.VisiblePosition = 19;
			ultraGridColumn54.Hidden = true;
			ultraGridColumn55.Header.VisiblePosition = 69;
			ultraGridColumn55.Hidden = true;
			ultraGridColumn56.Header.VisiblePosition = 18;
			ultraGridColumn56.Hidden = true;
			ultraGridColumn57.Header.VisiblePosition = 38;
			ultraGridColumn57.Hidden = true;
			ultraGridColumn58.Header.VisiblePosition = 39;
			ultraGridColumn58.Hidden = true;
			ultraGridColumn59.Header.VisiblePosition = 36;
			ultraGridColumn59.Hidden = true;
			ultraGridColumn60.Header.VisiblePosition = 37;
			ultraGridColumn60.Hidden = true;
			ultraGridColumn61.Header.VisiblePosition = 20;
			ultraGridColumn61.Hidden = true;
			ultraGridColumn62.Header.VisiblePosition = 21;
			ultraGridColumn62.Hidden = true;
			ultraGridColumn63.Header.VisiblePosition = 76;
			ultraGridColumn63.Hidden = true;
			ultraGridColumn64.Header.VisiblePosition = 71;
			ultraGridColumn64.Hidden = true;
			ultraGridColumn65.Header.VisiblePosition = 26;
			ultraGridColumn65.Hidden = true;
			ultraGridColumn66.Header.VisiblePosition = 27;
			ultraGridColumn66.Hidden = true;
			ultraGridColumn67.Header.VisiblePosition = 28;
			ultraGridColumn67.Hidden = true;
			ultraGridColumn68.Header.VisiblePosition = 25;
			ultraGridColumn68.Hidden = true;
			ultraGridColumn69.Header.VisiblePosition = 5;
			ultraGridColumn69.Hidden = true;
			ultraGridColumn70.Header.VisiblePosition = 6;
			ultraGridColumn70.Hidden = true;
			ultraGridColumn71.Header.VisiblePosition = 7;
			ultraGridColumn71.Hidden = true;
			ultraGridColumn72.Header.VisiblePosition = 8;
			ultraGridColumn72.Hidden = true;
			ultraGridColumn73.Header.VisiblePosition = 9;
			ultraGridColumn73.Hidden = true;
			ultraGridColumn74.Header.VisiblePosition = 10;
			ultraGridColumn74.Hidden = true;
			ultraGridColumn75.Header.VisiblePosition = 11;
			ultraGridColumn75.Hidden = true;
			ultraGridColumn76.Header.VisiblePosition = 12;
			ultraGridColumn76.Hidden = true;
			ultraGridColumn77.Header.VisiblePosition = 13;
			ultraGridColumn77.Hidden = true;
			ultraGridColumn78.Header.VisiblePosition = 73;
			ultraGridColumn78.Hidden = true;
			ultraGridColumn79.Header.VisiblePosition = 2;
			ultraGridColumn79.Hidden = true;
			ultraGridColumn80.Header.VisiblePosition = 3;
			ultraGridColumn80.Hidden = true;
			ultraGridColumn81.Header.VisiblePosition = 60;
			ultraGridColumn81.Hidden = true;
			ultraGridColumn82.Header.VisiblePosition = 24;
			ultraGridColumn82.Hidden = true;
			ultraGridColumn83.Header.VisiblePosition = 35;
			ultraGridColumn83.Hidden = true;
			ultraGridColumn84.Header.VisiblePosition = 15;
			ultraGridColumn84.Hidden = true;
			ultraGridColumn85.Header.VisiblePosition = 41;
			ultraGridColumn85.Hidden = true;
			ultraGridColumn86.Header.VisiblePosition = 79;
			ultraGridColumn86.Hidden = true;
			ultraGridColumn87.Header.VisiblePosition = 14;
			ultraGridColumn87.Hidden = true;
			ultraGridColumn88.Header.VisiblePosition = 74;
			ultraGridColumn88.Hidden = true;
			ultraGridColumn89.Header.VisiblePosition = 85;
			ultraGridColumn89.Hidden = true;
			ultraGridColumn90.Header.VisiblePosition = 86;
			ultraGridColumn90.Hidden = true;
			ultraGridColumn91.Header.VisiblePosition = 87;
			ultraGridColumn91.Hidden = true;
			ultraGridColumn92.Header.VisiblePosition = 88;
			ultraGridColumn92.Hidden = true;
			ultraGridColumn93.Header.VisiblePosition = 89;
			ultraGridColumn93.Hidden = true;
			ultraGridColumn94.Header.VisiblePosition = 90;
			ultraGridColumn94.Hidden = true;
			ultraGridColumn95.Header.VisiblePosition = 95;
			ultraGridColumn95.Hidden = true;
			ultraGridColumn96.Header.VisiblePosition = 96;
			ultraGridColumn96.Hidden = true;
			ultraGridColumn97.Header.VisiblePosition = 97;
			ultraGridColumn97.Hidden = true;
			ultraGridColumn98.Header.VisiblePosition = 98;
			ultraGridColumn98.Hidden = true;
			ultraGridColumn99.Header.VisiblePosition = 101;
			ultraGridColumn99.Hidden = true;
			ultraGridColumn100.Header.VisiblePosition = 102;
			ultraGridColumn100.Hidden = true;
			ultraGridColumn101.Header.VisiblePosition = 103;
			ultraGridColumn101.Hidden = true;
			ultraGridColumn102.Header.VisiblePosition = 92;
			ultraGridColumn102.Hidden = true;
			ultraGridColumn103.Header.VisiblePosition = 1;
			ultraGridColumn103.Width = 298;
			ultraGridColumn104.Header.VisiblePosition = 0;
			ultraGridColumn105.Header.VisiblePosition = 104;
			ultraGridColumn105.Hidden = true;
			ultraGridColumn106.Header.VisiblePosition = 105;
			ultraGridColumn106.Hidden = true;
			ultraGridColumn107.Header.VisiblePosition = 106;
			ultraGridColumn107.Hidden = true;
			ultraGridColumn108.Header.VisiblePosition = 107;
			ultraGridColumn108.Hidden = true;
			ultraGridColumn109.Header.VisiblePosition = 108;
			ultraGridColumn109.Hidden = true;
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
																										 ultraGridColumn109});
			this.cmbArticulo.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.cmbArticulo.DisplayMember = "Codigo";
			this.cmbArticulo.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbArticulo.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbArticulo.Location = new System.Drawing.Point(160, 95);
			this.cmbArticulo.Name = "cmbArticulo";
			this.cmbArticulo.Size = new System.Drawing.Size(232, 21);
			this.cmbArticulo.TabIndex = 5;
			this.cmbArticulo.ValueMember = "idArticulo";
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
			// 
			// cdsGrupoArt
			// 
			this.cdsGrupoArt.BindingContextCtrl = this;
			this.cdsGrupoArt.DataLibrary = "LibFacturacion";
			this.cdsGrupoArt.DataLibraryUrl = "";
			this.cdsGrupoArt.DataSetDef = "dsArticuloTabla";
			this.cdsGrupoArt.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsGrupoArt.Name = "cdsGrupoArt";
			this.cdsGrupoArt.SchemaClassName = "LibFacturacion.DataClass";
			this.cdsGrupoArt.SchemaDef = null;
			// 
			// cmbGrupo
			// 
			this.cmbGrupo.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbGrupo.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbGrupo.DataSource = this.cdvGrupoArt;
			ultraGridBand2.AddButtonCaption = "Articulo";
			ultraGridColumn110.Header.VisiblePosition = 8;
			ultraGridColumn110.Hidden = true;
			ultraGridColumn111.Header.VisiblePosition = 9;
			ultraGridColumn111.Hidden = true;
			ultraGridColumn112.Header.VisiblePosition = 0;
			ultraGridColumn112.Width = 313;
			ultraGridColumn113.Header.VisiblePosition = 2;
			ultraGridColumn113.Hidden = true;
			ultraGridColumn114.Header.VisiblePosition = 3;
			ultraGridColumn114.Hidden = true;
			ultraGridColumn115.Header.VisiblePosition = 4;
			ultraGridColumn115.Hidden = true;
			ultraGridColumn116.Header.VisiblePosition = 7;
			ultraGridColumn116.Hidden = true;
			ultraGridColumn117.Header.VisiblePosition = 5;
			ultraGridColumn117.Hidden = true;
			ultraGridColumn118.Header.VisiblePosition = 6;
			ultraGridColumn118.Hidden = true;
			ultraGridColumn119.Header.VisiblePosition = 10;
			ultraGridColumn119.Hidden = true;
			ultraGridColumn120.Header.VisiblePosition = 1;
			ultraGridColumn120.Hidden = true;
			ultraGridColumn121.Header.VisiblePosition = 11;
			ultraGridColumn121.Hidden = true;
			ultraGridColumn122.Header.VisiblePosition = 12;
			ultraGridColumn122.Hidden = true;
			ultraGridBand2.Columns.AddRange(new object[] {
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
																										 ultraGridColumn122});
			this.cmbGrupo.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			this.cmbGrupo.DisplayMember = "Grupo";
			this.cmbGrupo.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbGrupo.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbGrupo.Location = new System.Drawing.Point(160, 129);
			this.cmbGrupo.Name = "cmbGrupo";
			this.cmbGrupo.Size = new System.Drawing.Size(232, 21);
			this.cmbGrupo.TabIndex = 7;
			this.cmbGrupo.ValueMember = "idGrupoArticulo";
			// 
			// cdvGrupoArt
			// 
			this.cdvGrupoArt.BindingContextCtrl = this;
			this.cdvGrupoArt.DataSet = this.cdsGrupoArt;
			this.cdvGrupoArt.Sort = "Grupo";
			this.cdvGrupoArt.TableName = "";
			this.cdvGrupoArt.TableViewName = "ArticuloGrupo";
			// 
			// chkTodoGrupo
			// 
			this.chkTodoGrupo.Checked = true;
			this.chkTodoGrupo.CheckState = System.Windows.Forms.CheckState.Checked;
			this.chkTodoGrupo.Location = new System.Drawing.Point(32, 129);
			this.chkTodoGrupo.Name = "chkTodoGrupo";
			this.chkTodoGrupo.Size = new System.Drawing.Size(96, 22);
			this.chkTodoGrupo.TabIndex = 6;
			this.chkTodoGrupo.Text = "Todo Grupo";
			// 
			// btReporte
			// 
			this.btReporte.Location = new System.Drawing.Point(80, 284);
			this.btReporte.Name = "btReporte";
			this.btReporte.Size = new System.Drawing.Size(75, 25);
			this.btReporte.TabIndex = 8;
			this.btReporte.Text = "Reporte";
			this.btReporte.Click += new System.EventHandler(this.btReporte_Click);
			// 
			// btSalir
			// 
			this.btSalir.Location = new System.Drawing.Point(256, 284);
			this.btSalir.Name = "btSalir";
			this.btSalir.Size = new System.Drawing.Size(75, 25);
			this.btSalir.TabIndex = 9;
			this.btSalir.Text = "Salir";
			this.btSalir.Click += new System.EventHandler(this.btSalir_Click);
			// 
			// txtReporte
			// 
			this.txtReporte.Location = new System.Drawing.Point(160, 207);
			this.txtReporte.Name = "txtReporte";
			this.txtReporte.Size = new System.Drawing.Size(100, 22);
			this.txtReporte.TabIndex = 10;
			this.txtReporte.Text = "TomaFisica";
			this.toolTip1.SetToolTip(this.txtReporte, "Validos TomaFisicaTar y TomaFisicaArt");
			// 
			// ultraLabel1
			// 
			this.ultraLabel1.AutoSize = true;
			this.ultraLabel1.Location = new System.Drawing.Point(32, 215);
			this.ultraLabel1.Name = "ultraLabel1";
			this.ultraLabel1.Size = new System.Drawing.Size(43, 15);
			this.ultraLabel1.TabIndex = 11;
			this.ultraLabel1.Text = "Reporte";
			// 
			// errorProvider
			// 
			this.errorProvider.ContainerControl = this;
			// 
			// chkResumen
			// 
			this.chkResumen.Location = new System.Drawing.Point(288, 207);
			this.chkResumen.Name = "chkResumen";
			this.chkResumen.Size = new System.Drawing.Size(72, 21);
			this.chkResumen.TabIndex = 12;
			this.chkResumen.Text = "Resumen";
			// 
			// ultraButton1
			// 
			this.ultraButton1.Location = new System.Drawing.Point(296, 60);
			this.ultraButton1.Name = "ultraButton1";
			this.ultraButton1.Size = new System.Drawing.Size(112, 26);
			this.ultraButton1.TabIndex = 16;
			this.ultraButton1.Text = "Ajuste C.P.B.";
			this.toolTip1.SetToolTip(this.ultraButton1, "Ajuste por Bodega y con costo Promedio por cada bodega");
			this.ultraButton1.Click += new System.EventHandler(this.ultraButton1_Click);
			this.ultraButton1.MouseDown += new System.Windows.Forms.MouseEventHandler(this.ultraButton1_MouseDown);
			// 
			// btAjuste
			// 
			this.btAjuste.Location = new System.Drawing.Point(296, 26);
			this.btAjuste.Name = "btAjuste";
			this.btAjuste.Size = new System.Drawing.Size(112, 26);
			this.btAjuste.TabIndex = 13;
			this.btAjuste.Text = "Ajuste";
			this.btAjuste.Click += new System.EventHandler(this.btAjuste_Click);
			this.btAjuste.MouseDown += new System.Windows.Forms.MouseEventHandler(this.btAjuste_MouseDown);
			// 
			// cmbFecha
			// 
			dateButton1.Caption = "Today";
			this.cmbFecha.DateButtons.Add(dateButton1);
			this.cmbFecha.Format = "dd/MMM/yyyy";
			this.cmbFecha.Location = new System.Drawing.Point(160, 241);
			this.cmbFecha.Name = "cmbFecha";
			this.cmbFecha.NonAutoSizeHeight = 23;
			this.cmbFecha.Size = new System.Drawing.Size(112, 21);
			this.cmbFecha.SpinButtonIncrement = Infragistics.Win.UltraWinSchedule.SpinIncrementUnit.Months;
			this.cmbFecha.SpinButtonsVisible = true;
			this.cmbFecha.TabIndex = 14;
			this.cmbFecha.Value = new System.DateTime(2010, 1, 4, 0, 0, 0, 0);
			// 
			// lblFecha
			// 
			this.lblFecha.AutoSize = true;
			this.lblFecha.Location = new System.Drawing.Point(32, 241);
			this.lblFecha.Name = "lblFecha";
			this.lblFecha.Size = new System.Drawing.Size(34, 15);
			this.lblFecha.TabIndex = 15;
			this.lblFecha.Text = "Fecha";
			// 
			// cmbTipo
			// 
			this.cmbTipo.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbTipo.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbTipo.DataSource = this.cdvTipo;
			this.cmbTipo.DisplayLayout.AutoFitColumns = true;
			ultraGridBand3.AddButtonCaption = "Articulo";
			ultraGridColumn123.Header.VisiblePosition = 1;
			ultraGridColumn123.Hidden = true;
			ultraGridColumn124.Header.VisiblePosition = 0;
			ultraGridColumn124.Width = 213;
			ultraGridBand3.Columns.AddRange(new object[] {
																										 ultraGridColumn123,
																										 ultraGridColumn124});
			this.cmbTipo.DisplayLayout.BandsSerializer.Add(ultraGridBand3);
			this.cmbTipo.DisplayMember = "Tipo";
			this.cmbTipo.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbTipo.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbTipo.Location = new System.Drawing.Point(160, 164);
			this.cmbTipo.Name = "cmbTipo";
			this.cmbTipo.Size = new System.Drawing.Size(232, 21);
			this.cmbTipo.TabIndex = 18;
			this.cmbTipo.ValueMember = "idTipoGrupo";
			// 
			// cdvTipo
			// 
			this.cdvTipo.BindingContextCtrl = this;
			this.cdvTipo.DataSet = this.cdsTipo;
			this.cdvTipo.TableName = "";
			this.cdvTipo.TableViewName = "ArticuloTipo";
			// 
			// cdsTipo
			// 
			this.cdsTipo.BindingContextCtrl = this;
			this.cdsTipo.DataLibrary = "LibFacturacion";
			this.cdsTipo.DataLibraryUrl = "";
			this.cdsTipo.DataSetDef = "dsArticuloTabla";
			this.cdsTipo.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsTipo.Name = "cdsTipo";
			this.cdsTipo.SchemaClassName = "LibFacturacion.DataClass";
			this.cdsTipo.SchemaDef = null;
			// 
			// chkTipo
			// 
			this.chkTipo.Checked = true;
			this.chkTipo.CheckState = System.Windows.Forms.CheckState.Checked;
			this.chkTipo.Location = new System.Drawing.Point(32, 164);
			this.chkTipo.Name = "chkTipo";
			this.chkTipo.Size = new System.Drawing.Size(96, 21);
			this.chkTipo.TabIndex = 17;
			this.chkTipo.Text = "Todo Tipo";
			// 
			// RepTomaFisica
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 14);
			this.ClientSize = new System.Drawing.Size(432, 332);
			this.Controls.Add(this.cmbTipo);
			this.Controls.Add(this.chkTipo);
			this.Controls.Add(this.ultraButton1);
			this.Controls.Add(this.lblFecha);
			this.Controls.Add(this.cmbFecha);
			this.Controls.Add(this.btAjuste);
			this.Controls.Add(this.chkResumen);
			this.Controls.Add(this.ultraLabel1);
			this.Controls.Add(this.txtReporte);
			this.Controls.Add(this.btSalir);
			this.Controls.Add(this.btReporte);
			this.Controls.Add(this.cmbGrupo);
			this.Controls.Add(this.chkTodoGrupo);
			this.Controls.Add(this.cmbArticulo);
			this.Controls.Add(this.chkTodoArticulo);
			this.Controls.Add(this.txtTarjeta);
			this.Controls.Add(this.chkTodaTarjeta);
			this.Controls.Add(this.txtBodega);
			this.Controls.Add(this.chkTodaBodega);
			this.MaximizeBox = false;
			this.Name = "RepTomaFisica";
			this.Text = "Toma Fisica";
			this.Closing += new System.ComponentModel.CancelEventHandler(this.RepTomaFisica_Closing);
			this.Load += new System.EventHandler(this.RepTomaFisica_Load);
			((System.ComponentModel.ISupportInitialize)(this.txtBodega)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtTarjeta)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbArticulo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvArtCodigo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsArticulo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsGrupoArt)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbGrupo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvGrupoArt)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtReporte)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbFecha)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbTipo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvTipo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsTipo)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void btSalir_Click(object sender, System.EventArgs e)
		{
			Close();
		}

		private bool Verifica()
		{
			bool bOk = true;
			errorProvider.SetError(txtBodega, "");
			errorProvider.SetError(txtTarjeta, "");
			errorProvider.SetError(cmbArticulo, "");
			errorProvider.SetError(cmbArticulo, "");

			if (!chkTodaBodega.Checked) 
				if ((int) txtBodega.Value < 1)
				{
					errorProvider.SetError(txtBodega, "Ingrese Bodega");
					bOk = false;
				}
					
			if (!chkTodaTarjeta.Checked)
				if (txtTarjeta.Text.Trim().Length == 0)
				{
					errorProvider.SetError(txtTarjeta, "Ingrese Tarjeta");
					bOk = false;
				}
			if (!chkTodoArticulo.Checked) 
				if (cmbArticulo.Text.Trim().Length == 0)
				{
					errorProvider.SetError(cmbArticulo, "Ingrese Articulo");
					bOk = false;
				}
			if (!chkTodoGrupo.Checked) 
				if (cmbGrupo.Text.Trim().Length == 0)
				{
					errorProvider.SetError(cmbGrupo, "Ingrese Grupo de Articulos");
					bOk = false;
				}
			string stReporte = MenuLatinium.stDirReportes.Trim();
			if (!stReporte.EndsWith("\\")) stReporte += "\\";
			stReporte += txtReporte.Text.Trim() + ".rpt";
			if (!File.Exists(stReporte))
			{
				errorProvider.SetError(txtReporte, "Reporte " + stReporte + " no existe");
				bOk = false;
			}
			return bOk;				
		}

		private void btReporte_Click(object sender, System.EventArgs e)
		{
			if (!Verifica()) return;

			Cursor = Cursors.WaitCursor;
			string stReporte = txtReporte.Text.Trim() + ".RPT";
			Funcion.EjecutaSQL(cdsArticulo, "Update TmpTomaDetalle Set Mostrar = 0", true);
			string stFiltro = "Update TmpTomaDetalle Set Mostrar = 1 "
				+ "From TmpTomaDetalle Inner Join TmpTomaFisica On "
				+ "TmpTomaDetalle.idTmpTomaFisica = TmpTomaFisica.idTmpTomaFisica Where 1=1 ";
			string stTitulo = "";
			string stFiltro1 = "true";
			DateTime dtFecha = (DateTime) cmbFecha.Value;
			if (!chkTodaBodega.Checked) 
			{
//				string stBodega = "Update Articulo Set SaldoT = "
//					+ "(SELECT  IsNull(sum(IsNull(DC.CANTIDAD, 0) * DC.SIGNO), 0) "
//					+ "FROM COMPRA C INNER JOIN DETCOMPRA DC ON C.IDCOMPRA = DC.IDCOMPRA "
//					+ "WHERE (DC.SIGNO=1 OR DC.SIGNO=-1) AND DC.BORRAR = 0 "
//					+ "AND C.BORRAR = 0 AND DC.IDARTICULO = Articulo.idarticulo "
//					+ "And Dc.Bodega = " + txtBodega.Value.ToString().Trim()
//					+ " And c.Fecha < '" + dtFecha.AddDays(1).ToString("yyyyMMdd") + "')"
//					+ " Where idTipoGrupo In (1, 6, 8)";
				string stBodega = string.Format("Exec TomaFisicaBodega {0}, '{1}'", txtBodega.Value, dtFecha.ToString("yyyyMMdd"));
				Funcion.EjecutaSQL(cdsArticulo, stBodega, true);
				stTitulo += "Bodega: " + txtBodega.Value.ToString().Trim() + " ";
				stFiltro += " And TmpTomaFisica.Bodega = " + txtBodega.Value.ToString().Trim();
				txtReporte.Text = "TomaFisicaBodega";
				stReporte = txtReporte.Text.Trim() + ".RPT";
				//				stFiltro1 += " And {TmpTomaFisica.Bodega} = " + txtBodega.Value.ToString().Trim();
			}
			else
			{
				string stBodega = "Update Articulo Set SaldoT = "
					+ "(SELECT  IsNull(sum(IsNull(DC.CANTIDAD, 0) * DC.SIGNO), 0) "
					+ "FROM COMPRA C INNER JOIN DETCOMPRA DC ON C.IDCOMPRA = DC.IDCOMPRA "
					+ "WHERE (DC.SIGNO=1 OR DC.SIGNO=-1) AND DC.BORRAR = 0 "
					+ "AND C.BORRAR = 0 AND DC.IDARTICULO = Articulo.idarticulo And "
          + "c.Fecha < '" + dtFecha.AddDays(1).ToString("yyyyMMdd") + "')"
          + " Where idTipoGrupo In (1, 6, 8)";
				Funcion.EjecutaSQL(cdsArticulo, stBodega, true);
			}

			if (!chkTodaTarjeta.Checked)
			{
				stTitulo += "Tarjeta: " + txtTarjeta.Value.ToString().Trim() + " ";
				stFiltro += " And TmpTomaFisica.Tarjeta = '" + txtTarjeta.Value.ToString().Trim() + "'";
				stFiltro1 += " And {TmpTomaFisica.Tarjeta} = '" + txtTarjeta.Value.ToString().Trim() + "'";
			}
			if (!chkTodoArticulo.Checked) 
			{
				stTitulo += "Articulo: " + cmbArticulo.Text.Trim() + " ";
				stFiltro1 += " And {Articulo.idArticulo} = " + cmbArticulo.Value.ToString();
			}
			if (!chkTodoGrupo.Checked) 
			{
				stTitulo += "Grupo: " + cmbGrupo.Text.Trim() + " ";
				stFiltro1 += " And {Articulo.idGrupoArticulo} = " + cmbGrupo.Value.ToString();
			}
			if (!chkTipo.Checked) 
			{
				stTitulo += "Grupo: " + cmbTipo.Text.Trim() + " ";
				stFiltro1 += " And {Articulo.idTipoGrupo} = " + cmbTipo.Value.ToString();
			}
			Funcion.EjecutaSQL(cdsArticulo, stFiltro, true);
			Reporte miReporte = new Reporte(stReporte, stFiltro1);
			miReporte.MdiParent = MdiParent;
			if (chkResumen.Checked)
			{
				miReporte.Resumen(1);
			}
			miReporte.Titulo(stTitulo);
			miReporte.Show();
			Cursor = Cursors.Default;
		}

		private void btAjuste_Click(object sender, System.EventArgs e)
		{
			int iBodega = int.Parse(txtBodega.Value.ToString());
			if (iBodega == 0)
			{
				MessageBox.Show("Seleccione la Bodega de la cual desea realizar el ajuste");
				return;
			}
			if (DialogResult.No == MessageBox.Show(
				"¿Está seguro que desea Ajustar los saldos de la Bodega " + iBodega.ToString() + "?", 
				"Confirmación", MessageBoxButtons.YesNo, MessageBoxIcon.Warning, 
				MessageBoxDefaultButton.Button2)) return;

			Cursor = Cursors.WaitCursor;
			DateTime dtFecha = (DateTime) cmbFecha.Value;
			string stBodega = "Update Articulo Set SaldoT = "
				+ "(SELECT  IsNull(sum(IsNull(DC.CANTIDAD, 0) * DC.SIGNO), 0) "
				+ "FROM COMPRA C INNER JOIN DETCOMPRA DC ON C.IDCOMPRA = DC.IDCOMPRA "
				+ "WHERE (DC.SIGNO=1 OR DC.SIGNO=-1) AND DC.BORRAR = 0 "
				+ "AND C.BORRAR = 0 AND DC.IDARTICULO = Articulo.idArticulo "
				+ "And Dc.Bodega = " + iBodega.ToString()
				+ " And c.Fecha < '" + dtFecha.AddDays(1).ToString("yyyyMMdd") + "')";
			Funcion.EjecutaSQL(cdsArticulo, stBodega, true);
			stBodega = "Exec TomaFisicaAjuste " + iBodega.ToString();
			Funcion.EjecutaSQL(cdsArticulo, stBodega, true);
			Cursor = Cursors.Default;

			MessageBox.Show("Ajuste Realizado. Use la opcion Herramientas Actualizar Saldos y el boton Saldos y Costos", "Información");
		}

		private void ultraButton1_Click(object sender, System.EventArgs e)
		{
			int iBodega = int.Parse(txtBodega.Value.ToString());
			if (iBodega == 0)
			{
				MessageBox.Show("Seleccione la Bodega de la cual desea realizar el ajuste");
				return;
			}
			if (DialogResult.No == MessageBox.Show(
				"¿Está seguro que desea Ajustar los saldos con costo pormedio por bodega de la Bodega " + iBodega.ToString() + "?", 
				"Confirmación", MessageBoxButtons.YesNo, MessageBoxIcon.Warning, 
				MessageBoxDefaultButton.Button2)) return;

			Cursor = Cursors.WaitCursor;
			DateTime dtFecha = (DateTime) cmbFecha.Value;
			string stBodega = "Update Articulo Set SaldoT = "
				+ "(SELECT  IsNull(sum(IsNull(DC.CANTIDAD, 0) * DC.SIGNO), 0) "
				+ "FROM COMPRA C INNER JOIN DETCOMPRA DC ON C.IDCOMPRA = DC.IDCOMPRA "
				+ "WHERE (DC.SIGNO=1 OR DC.SIGNO=-1) AND DC.BORRAR = 0 "
				+ "AND C.BORRAR = 0 AND DC.IDARTICULO = Articulo.idarticulo "
				+ "And Dc.Bodega = " + iBodega.ToString()
				+ " And c.Fecha < '" + dtFecha.AddDays(1).ToString("yyyyMMdd") + "')";
			Funcion.EjecutaSQL(cdsArticulo, stBodega, true);
			stBodega = "Exec TomaFisicaAjusteBodega " + iBodega.ToString() + ", '"
				+ dtFecha.AddDays(1).ToString("yyyyMMdd") + "'";
			Funcion.EjecutaSQL(cdsArticulo, stBodega, true);
			Cursor = Cursors.Default;
			MessageBox.Show("Ajuste Realizado", "Información");
		}

		private void btAjuste_MouseDown(object sender, System.Windows.Forms.MouseEventArgs e)
		{
			if (e.Button != MouseButtons.Right) return;
			int iBodega = int.Parse(txtBodega.Value.ToString());
			if (iBodega == 0)
			{
				MessageBox.Show("Seleccione la Bodega de la cual desea realizar el ajuste");
				return;
			}
			if (DialogResult.No == MessageBox.Show(
				"¿Está seguro que desea Ajustar los saldos de la Bodega " + iBodega.ToString() + "?", 
				"Confirmación", MessageBoxButtons.YesNo, MessageBoxIcon.Warning, 
				MessageBoxDefaultButton.Button2)) return;

			Cursor = Cursors.WaitCursor;
			DateTime dtFecha = (DateTime) cmbFecha.Value;
			string stBodega = "Update Articulo Set SaldoT = "
				+ "(SELECT  IsNull(sum(IsNull(DC.CANTIDAD, 0) * DC.SIGNO), 0) "
				+ "FROM COMPRA C INNER JOIN DETCOMPRA DC ON C.IDCOMPRA = DC.IDCOMPRA "
				+ "WHERE (DC.SIGNO=1 OR DC.SIGNO=-1) AND DC.BORRAR = 0 "
				+ "AND C.BORRAR = 0 AND DC.IDARTICULO = Articulo.idarticulo "
				+ "And Dc.Bodega = " + iBodega.ToString()
				+ " And c.Fecha < '" + dtFecha.AddDays(1).ToString("yyyyMMdd") + "')";
			Funcion.EjecutaSQL(cdsArticulo, stBodega, true);
			stBodega = "Exec TomaFisicaAjuste " + iBodega.ToString();
			Cursor = Cursors.Default;

			StreamWriter fs = File.AppendText(MenuLatinium.stDirInicio + @"\TomaFisica.Sql");
			fs.WriteLine(stBodega);
			fs.Flush();
			fs.Close();
			string stBase = MenuLatinium.stDirConta;
			string stParametro = "/U infoelect /P Bl45o6$9 /d ";
			stParametro += stBase;
			stParametro += " /S " + MenuLatinium.stDirServidor;
			stParametro += @" /n /t 10000 /i " + MenuLatinium.stDirInicio + @"\TomaFisica.Sql /o "
			+ MenuLatinium.stDirInicio + @"\Salida.txt";

			System.Diagnostics.Process.Start("osql", @stParametro);

			if (DialogResult.Yes == MessageBox.Show(
				"¿Desea Ver Registro de Actividades?", 
				"Información", MessageBoxButtons.YesNo, MessageBoxIcon.Question))
				System.Diagnostics.Process.Start("NotePad.exe", MenuLatinium.stDirInicio + @"\Salida.txt");
		}

		private void ultraButton1_MouseDown(object sender, System.Windows.Forms.MouseEventArgs e)
		{
			if (e.Button != MouseButtons.Right) return;
			int iBodega = int.Parse(txtBodega.Value.ToString());
			if (iBodega == 0)
			{
				MessageBox.Show("Seleccione la Bodega de la cual desea realizar el ajuste");
				return;
			}
			if (DialogResult.No == MessageBox.Show(
				"¿Está seguro que desea Ajustar los saldos con costo pormedio por bodega de la Bodega " + iBodega.ToString() + "?", 
				"Confirmación", MessageBoxButtons.YesNo, MessageBoxIcon.Warning, 
				MessageBoxDefaultButton.Button2)) return;

			Cursor = Cursors.WaitCursor;
			DateTime dtFecha = (DateTime) cmbFecha.Value;
			string stBodega = "Update Articulo Set SaldoT = "
				+ "(SELECT  IsNull(sum(IsNull(DC.CANTIDAD, 0) * DC.SIGNO), 0) "
				+ "FROM COMPRA C INNER JOIN DETCOMPRA DC ON C.IDCOMPRA = DC.IDCOMPRA "
				+ "WHERE (DC.SIGNO=1 OR DC.SIGNO=-1) AND DC.BORRAR = 0 "
				+ "AND C.BORRAR = 0 AND DC.IDARTICULO = Articulo.idarticulo "
				+ "And Dc.Bodega = " + iBodega.ToString()
				+ " And c.Fecha < '" + dtFecha.AddDays(1).ToString("yyyyMMdd") + "')";
			Funcion.EjecutaSQL(cdsArticulo, stBodega, true);

			stBodega = "Exec TomaFisicaAjusteBodega " + iBodega.ToString() + ", '"
				+ dtFecha.AddDays(1).ToString("yyyyMMdd") + "'";
			Cursor = Cursors.Default;
			StreamWriter fs = File.AppendText(MenuLatinium.stDirInicio + @"\TomaFisica.Sql");
			fs.WriteLine(stBodega);
			fs.Flush();
			fs.Close();
			string stBase = MenuLatinium.stDirConta;
			string stParametro = "/U infoelect /P Bl45o6$9 /d ";
			stParametro += stBase;
			stParametro += " /S " + MenuLatinium.stDirServidor;
			stParametro += @" /n /t 10000 /i " + MenuLatinium.stDirInicio + @"\TomaFisica.Sql /o "
			+ MenuLatinium.stDirInicio + @"\Salida.txt";

			System.Diagnostics.Process.Start("osql", @stParametro);

			if (DialogResult.Yes == MessageBox.Show(
				"¿Desea Ver Registro de Actividades?", 
				"Información", MessageBoxButtons.YesNo, MessageBoxIcon.Question))
				System.Diagnostics.Process.Start("NotePad.exe", MenuLatinium.stDirInicio + @"\Salida.txt");
		}

		private void RepTomaFisica_Closing(object sender, System.ComponentModel.CancelEventArgs e)
		{
			string stAudita = "Exec AuditaCrear 53, 7, 'RepTomaFisica'";
			Funcion.EjecutaSQL(cdsArticulo, stAudita, true);
		}

		private void RepTomaFisica_Load(object sender, System.EventArgs e)
		{
			string stAudita = "Exec AuditaCrear 53, 6, 'RepTomaFisica'";
			Funcion.EjecutaSQL(cdsArticulo, stAudita, true);
			cmbFecha.Value = DateTime.Today;
		}
	}
}
