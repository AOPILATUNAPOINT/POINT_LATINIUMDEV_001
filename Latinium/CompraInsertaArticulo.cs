using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de CompraInsertaArticulo.
	/// </summary>
	public class CompraInsertaArticulo : DevExpress.XtraEditors.XtraForm
	{
		private C1.Data.C1DataSet cdsCompraNumero;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbTipo;
		private System.Windows.Forms.Label label2;
		private System.Windows.Forms.Label label1;
		private Infragistics.Win.Misc.UltraButton btCrea;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbArticulo;
		private System.Windows.Forms.Label label3;
		private System.Windows.Forms.Label label4;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtValor;
		private System.Windows.Forms.Label label5;
		private System.Windows.Forms.TextBox txtNota;
		private C1.Data.C1DataSet cdsArticulo;
		private System.Windows.Forms.MainMenu mainMenu1;
		private Infragistics.Win.Misc.UltraButton btBorra;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtCantidad;
		private System.Windows.Forms.Label label6;
		private System.Windows.Forms.Panel panel1;
		private System.Windows.Forms.Label label7;
		private System.Windows.Forms.Panel panel2;
		private System.Windows.Forms.Label label8;
		private System.Windows.Forms.Label label9;
		private C1.Data.C1DataSet cdsGrupoCliente;
		private Infragistics.Win.Misc.UltraButton ultraButton1;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbGrupoCli;
		private C1.Data.C1DataSet cdsCliente;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbCliente;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbGrupoCliD;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbGrupoCliH;
		private System.Windows.Forms.Label label10;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public CompraInsertaArticulo()
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
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Articulo", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn18 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idArticulo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn19 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idArticuloSRI");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn20 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idColor");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn21 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idModelo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn22 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idGrupoArticulo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn23 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idSubGrupo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn24 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idArancel");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn25 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idMarca");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn26 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idUnidad");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn27 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPromocion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn28 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idTipoGrupo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn29 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idClase");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn30 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idSeccion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn31 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idSucursal");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn32 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPadre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn33 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaIngreso");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn34 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CodAlterno");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn35 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Saldo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn36 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Peso");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn37 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Caracteristica");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn38 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Aplicacion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn39 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Iva");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn40 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Ice");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn41 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Precio1");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn42 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Precio2");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn43 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Precio3");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn44 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Precio4");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn45 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Precio5");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn46 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Precio6");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn47 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CantidadMinima");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn48 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CantidadMaxima");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn49 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Residual");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn50 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CostoUltimo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn51 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaUltimo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn52 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCompra1");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn53 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCompra2");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn54 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCompra3");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn55 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CostoPromedio");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn56 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Gramos");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn57 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ConsumoProm");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn58 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("TiempoReposicion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn59 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("UnidadPeso");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn60 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Depreciacion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn61 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Incr1");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn62 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Incr2");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn63 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Incr3");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn64 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Incr4");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn65 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Incr5");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn66 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Incr6");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn67 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Ubicacion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn68 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("NumeroParte");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn69 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ComisionValor");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn70 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ComisionPorcentaje");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn71 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CantidadRequerida");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn72 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("PrecioPromocion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn73 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CantidadPromocion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn74 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("DescuentoContado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn75 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("DescuentoCredito");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn76 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("DescPorcContado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn77 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("DescPorcCre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn78 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CantReqDesc");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn79 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CantReqPorc");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn80 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("TipoPromocion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn81 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Responsable");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn82 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Comision30");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn83 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Comision60");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn84 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Comision90");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn85 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Comision120");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn86 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bodega1");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn87 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bodega2");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn88 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bodega3");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn89 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bodega4");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn90 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bodega5");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn91 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bodega6");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn92 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bodega7");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn93 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bodega8");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn94 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bodega9");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn95 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("SaldoT");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn96 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Alto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn97 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Ancho");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn98 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Largo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn99 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Color");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn100 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Descontinuado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn101 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Borrar");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn102 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Habilitado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn103 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ValorPeso");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn104 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("BodegaPredef");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn105 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Tiempo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn106 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cantidad1");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn107 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cantidad2");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn108 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cantidad3");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn109 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cantidad4");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn110 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cantidad5");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn111 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("PaqueteUnidad");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn112 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Editar");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn113 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nuevo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn114 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaEditado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn115 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CambiarPrecio");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn116 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("dCompra");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn117 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("dVenta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn118 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Foto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn119 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Notas");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn120 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Articulo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn121 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn122 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("BodegaPredefCR");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn123 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CodPromocion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn124 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CostoFecha");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn125 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Aportable");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn126 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ExigeLote");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand3 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Cliente", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn127 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCliente");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn128 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idGrupoCliente");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn129 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCobrador");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn130 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idVendedor");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn131 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idTipoPrecio");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn132 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCiudad");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn133 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idSector");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn134 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idTipoRuc");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn135 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idSucursal");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn136 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPadre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn137 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idFormaPago");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn138 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCredTribut");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn139 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn140 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn141 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Direccion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn142 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Telefono");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn143 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Fax");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn144 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Ruc");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn145 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Observacion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn146 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Representante");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn147 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Email");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn148 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Celular");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn149 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Contacto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn150 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Pais");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn151 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Notas");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn152 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Ambos");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn153 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Proveedor");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn154 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("AutorizaFactura");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn155 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("AutorizaImprenta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn156 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("MaximoCredito");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn157 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("DiasCredito");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn158 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("SerieFactura");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn159 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Fecha");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn160 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaIngreso");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn161 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("DescuentoContado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn162 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("DescuentoCredito");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn163 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Borrar");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn164 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("SaldoCxC");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn165 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("SaldoCxP");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn166 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Editar");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn167 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nuevo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn168 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaEditado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn169 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idRetRenta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn170 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idRetIva");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn171 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaCaducidad");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn172 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("AutorizaRet");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn173 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("SerieRet");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn174 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ContribEspecial");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn175 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CuentaCorr");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn176 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CuentaAh");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn177 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Conyuge");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn178 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ConyDireccion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn179 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ConyTelefono");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn180 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ConyCelular");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn181 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Referencia");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn182 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("RefTelefono");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn183 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("RefDireccion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn184 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Poliza");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn185 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Parroquia");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn186 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Sector");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn187 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("RefLocal");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn188 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("DesconProv");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn189 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("DesconCliente");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn190 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idClientePublicidad");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn191 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("InstPublica");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn192 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("PorComT");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn193 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("NetoValorT");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn194 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ProveMaster");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn195 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Entrega");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn196 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("RangoHasta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn197 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("RangoDesde");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn198 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idBanco");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand4 = new Infragistics.Win.UltraWinGrid.UltraGridBand("ClienteGrupo", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn199 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idGrupoCliente");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn200 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Grupo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn201 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCuenta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn202 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Borrar");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand5 = new Infragistics.Win.UltraWinGrid.UltraGridBand("ClienteGrupo", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn203 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idGrupoCliente");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn204 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Grupo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn205 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCuenta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn206 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Borrar");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand6 = new Infragistics.Win.UltraWinGrid.UltraGridBand("ClienteGrupo", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn207 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idGrupoCliente");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn208 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Grupo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn209 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCuenta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn210 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Borrar");
			this.cdsCompraNumero = new C1.Data.C1DataSet();
			this.cmbTipo = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.cmbArticulo = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.cdsArticulo = new C1.Data.C1DataSet();
			this.label2 = new System.Windows.Forms.Label();
			this.label1 = new System.Windows.Forms.Label();
			this.btCrea = new Infragistics.Win.Misc.UltraButton();
			this.label3 = new System.Windows.Forms.Label();
			this.label4 = new System.Windows.Forms.Label();
			this.txtValor = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label5 = new System.Windows.Forms.Label();
			this.txtNota = new System.Windows.Forms.TextBox();
			this.mainMenu1 = new System.Windows.Forms.MainMenu();
			this.btBorra = new Infragistics.Win.Misc.UltraButton();
			this.txtCantidad = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label6 = new System.Windows.Forms.Label();
			this.panel1 = new System.Windows.Forms.Panel();
			this.cmbCliente = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.cdsCliente = new C1.Data.C1DataSet();
			this.label7 = new System.Windows.Forms.Label();
			this.cmbGrupoCli = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.cdsGrupoCliente = new C1.Data.C1DataSet();
			this.panel2 = new System.Windows.Forms.Panel();
			this.cmbGrupoCliH = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.label8 = new System.Windows.Forms.Label();
			this.label9 = new System.Windows.Forms.Label();
			this.cmbGrupoCliD = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.ultraButton1 = new Infragistics.Win.Misc.UltraButton();
			this.label10 = new System.Windows.Forms.Label();
			((System.ComponentModel.ISupportInitialize)(this.cdsCompraNumero)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbTipo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbArticulo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsArticulo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtValor)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCantidad)).BeginInit();
			this.panel1.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.cmbCliente)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsCliente)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbGrupoCli)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsGrupoCliente)).BeginInit();
			this.panel2.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.cmbGrupoCliH)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbGrupoCliD)).BeginInit();
			this.SuspendLayout();
			// 
			// cdsCompraNumero
			// 
			this.cdsCompraNumero.BindingContextCtrl = this;
			this.cdsCompraNumero.DataLibrary = "LibFacturacion";
			this.cdsCompraNumero.DataLibraryUrl = "";
			this.cdsCompraNumero.DataSetDef = "dsCompraTabla";
			this.cdsCompraNumero.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsCompraNumero.Name = "cdsCompraNumero";
			this.cdsCompraNumero.SchemaClassName = "LibFacturacion.DataClass";
			this.cdsCompraNumero.SchemaDef = null;
			this.cdsCompraNumero.BeforeFill += new C1.Data.FillEventHandler(this.cdsCompraNumero_BeforeFill);
			// 
			// cmbTipo
			// 
			this.cmbTipo.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbTipo.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbTipo.DataMember = "CompraNumero";
			this.cmbTipo.DataSource = this.cdsCompraNumero;
			this.cmbTipo.DisplayLayout.AutoFitColumns = true;
			ultraGridColumn1.Header.VisiblePosition = 5;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn2.Header.VisiblePosition = 0;
			ultraGridColumn2.Width = 125;
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
			ultraGridColumn10.Width = 33;
			ultraGridColumn11.Header.VisiblePosition = 10;
			ultraGridColumn11.Hidden = true;
			ultraGridColumn11.Width = 52;
			ultraGridColumn12.Header.VisiblePosition = 11;
			ultraGridColumn12.Hidden = true;
			ultraGridColumn12.Width = 288;
			ultraGridColumn13.Header.VisiblePosition = 12;
			ultraGridColumn13.Hidden = true;
			ultraGridColumn13.Width = 18;
			ultraGridColumn14.Header.VisiblePosition = 13;
			ultraGridColumn14.Hidden = true;
			ultraGridColumn14.Width = 18;
			ultraGridColumn15.Header.VisiblePosition = 14;
			ultraGridColumn15.Hidden = true;
			ultraGridColumn15.Width = 37;
			ultraGridColumn16.Header.VisiblePosition = 15;
			ultraGridColumn16.Hidden = true;
			ultraGridColumn16.Width = 37;
			ultraGridColumn17.Header.VisiblePosition = 16;
			ultraGridColumn17.Hidden = true;
			ultraGridColumn17.Width = 25;
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
			this.cmbTipo.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.cmbTipo.DisplayMember = "Nombre";
			this.cmbTipo.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbTipo.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbTipo.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbTipo.Location = new System.Drawing.Point(96, 24);
			this.cmbTipo.Name = "cmbTipo";
			this.cmbTipo.Size = new System.Drawing.Size(144, 21);
			this.cmbTipo.TabIndex = 4;
			this.cmbTipo.ValueMember = "idTipoFactura";
			// 
			// cmbArticulo
			// 
			this.cmbArticulo.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbArticulo.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbArticulo.DataMember = "Articulo";
			this.cmbArticulo.DataSource = this.cdsArticulo;
			this.cmbArticulo.DisplayLayout.AutoFitColumns = true;
			ultraGridColumn18.Header.VisiblePosition = 0;
			ultraGridColumn18.Hidden = true;
			ultraGridColumn18.Width = 8;
			ultraGridColumn19.Header.VisiblePosition = 1;
			ultraGridColumn19.Hidden = true;
			ultraGridColumn19.Width = 8;
			ultraGridColumn20.Header.VisiblePosition = 2;
			ultraGridColumn20.Hidden = true;
			ultraGridColumn20.Width = 8;
			ultraGridColumn21.Header.VisiblePosition = 3;
			ultraGridColumn21.Hidden = true;
			ultraGridColumn21.Width = 8;
			ultraGridColumn22.Header.VisiblePosition = 4;
			ultraGridColumn22.Hidden = true;
			ultraGridColumn22.Width = 8;
			ultraGridColumn23.Header.VisiblePosition = 5;
			ultraGridColumn23.Hidden = true;
			ultraGridColumn23.Width = 8;
			ultraGridColumn24.Header.VisiblePosition = 6;
			ultraGridColumn24.Hidden = true;
			ultraGridColumn24.Width = 8;
			ultraGridColumn25.Header.VisiblePosition = 7;
			ultraGridColumn25.Hidden = true;
			ultraGridColumn25.Width = 8;
			ultraGridColumn26.Header.VisiblePosition = 8;
			ultraGridColumn26.Hidden = true;
			ultraGridColumn26.Width = 8;
			ultraGridColumn27.Header.VisiblePosition = 9;
			ultraGridColumn27.Hidden = true;
			ultraGridColumn27.Width = 8;
			ultraGridColumn28.Header.VisiblePosition = 10;
			ultraGridColumn28.Hidden = true;
			ultraGridColumn28.Width = 8;
			ultraGridColumn29.Header.VisiblePosition = 11;
			ultraGridColumn29.Hidden = true;
			ultraGridColumn29.Width = 8;
			ultraGridColumn30.Header.VisiblePosition = 12;
			ultraGridColumn30.Hidden = true;
			ultraGridColumn30.Width = 8;
			ultraGridColumn31.Header.VisiblePosition = 13;
			ultraGridColumn31.Hidden = true;
			ultraGridColumn31.Width = 8;
			ultraGridColumn32.Header.VisiblePosition = 14;
			ultraGridColumn32.Hidden = true;
			ultraGridColumn32.Width = 8;
			ultraGridColumn33.Header.VisiblePosition = 15;
			ultraGridColumn33.Hidden = true;
			ultraGridColumn33.Width = 16;
			ultraGridColumn34.Header.VisiblePosition = 16;
			ultraGridColumn34.Hidden = true;
			ultraGridColumn34.Width = 8;
			ultraGridColumn35.Header.VisiblePosition = 17;
			ultraGridColumn35.Hidden = true;
			ultraGridColumn35.Width = 8;
			ultraGridColumn36.Header.VisiblePosition = 18;
			ultraGridColumn36.Hidden = true;
			ultraGridColumn36.Width = 8;
			ultraGridColumn37.Header.VisiblePosition = 19;
			ultraGridColumn37.Hidden = true;
			ultraGridColumn37.Width = 8;
			ultraGridColumn38.Header.VisiblePosition = 20;
			ultraGridColumn38.Hidden = true;
			ultraGridColumn38.Width = 8;
			ultraGridColumn39.Header.VisiblePosition = 21;
			ultraGridColumn39.Hidden = true;
			ultraGridColumn39.Width = 8;
			ultraGridColumn40.Header.VisiblePosition = 22;
			ultraGridColumn40.Hidden = true;
			ultraGridColumn40.Width = 8;
			ultraGridColumn41.Header.VisiblePosition = 23;
			ultraGridColumn41.Hidden = true;
			ultraGridColumn41.Width = 8;
			ultraGridColumn42.Header.VisiblePosition = 24;
			ultraGridColumn42.Hidden = true;
			ultraGridColumn42.Width = 8;
			ultraGridColumn43.Header.VisiblePosition = 25;
			ultraGridColumn43.Hidden = true;
			ultraGridColumn43.Width = 8;
			ultraGridColumn44.Header.VisiblePosition = 26;
			ultraGridColumn44.Hidden = true;
			ultraGridColumn44.Width = 8;
			ultraGridColumn45.Header.VisiblePosition = 27;
			ultraGridColumn45.Hidden = true;
			ultraGridColumn45.Width = 8;
			ultraGridColumn46.Header.VisiblePosition = 28;
			ultraGridColumn46.Hidden = true;
			ultraGridColumn46.Width = 8;
			ultraGridColumn47.Header.VisiblePosition = 29;
			ultraGridColumn47.Hidden = true;
			ultraGridColumn47.Width = 8;
			ultraGridColumn48.Header.VisiblePosition = 30;
			ultraGridColumn48.Hidden = true;
			ultraGridColumn48.Width = 8;
			ultraGridColumn49.Header.VisiblePosition = 31;
			ultraGridColumn49.Hidden = true;
			ultraGridColumn49.Width = 8;
			ultraGridColumn50.Header.VisiblePosition = 32;
			ultraGridColumn50.Hidden = true;
			ultraGridColumn50.Width = 8;
			ultraGridColumn51.Header.VisiblePosition = 33;
			ultraGridColumn51.Hidden = true;
			ultraGridColumn51.Width = 16;
			ultraGridColumn52.Header.VisiblePosition = 34;
			ultraGridColumn52.Hidden = true;
			ultraGridColumn52.Width = 8;
			ultraGridColumn53.Header.VisiblePosition = 35;
			ultraGridColumn53.Hidden = true;
			ultraGridColumn53.Width = 8;
			ultraGridColumn54.Header.VisiblePosition = 36;
			ultraGridColumn54.Hidden = true;
			ultraGridColumn54.Width = 8;
			ultraGridColumn55.Header.VisiblePosition = 37;
			ultraGridColumn55.Hidden = true;
			ultraGridColumn55.Width = 8;
			ultraGridColumn56.Header.VisiblePosition = 38;
			ultraGridColumn56.Hidden = true;
			ultraGridColumn56.Width = 8;
			ultraGridColumn57.Header.VisiblePosition = 39;
			ultraGridColumn57.Hidden = true;
			ultraGridColumn57.Width = 8;
			ultraGridColumn58.Header.VisiblePosition = 40;
			ultraGridColumn58.Hidden = true;
			ultraGridColumn58.Width = 8;
			ultraGridColumn59.Header.VisiblePosition = 41;
			ultraGridColumn59.Hidden = true;
			ultraGridColumn59.Width = 8;
			ultraGridColumn60.Header.VisiblePosition = 42;
			ultraGridColumn60.Hidden = true;
			ultraGridColumn60.Width = 8;
			ultraGridColumn61.Header.VisiblePosition = 43;
			ultraGridColumn61.Hidden = true;
			ultraGridColumn61.Width = 8;
			ultraGridColumn62.Header.VisiblePosition = 44;
			ultraGridColumn62.Hidden = true;
			ultraGridColumn62.Width = 8;
			ultraGridColumn63.Header.VisiblePosition = 45;
			ultraGridColumn63.Hidden = true;
			ultraGridColumn63.Width = 8;
			ultraGridColumn64.Header.VisiblePosition = 46;
			ultraGridColumn64.Hidden = true;
			ultraGridColumn64.Width = 8;
			ultraGridColumn65.Header.VisiblePosition = 47;
			ultraGridColumn65.Hidden = true;
			ultraGridColumn65.Width = 8;
			ultraGridColumn66.Header.VisiblePosition = 48;
			ultraGridColumn66.Hidden = true;
			ultraGridColumn66.Width = 8;
			ultraGridColumn67.Header.VisiblePosition = 49;
			ultraGridColumn67.Hidden = true;
			ultraGridColumn67.Width = 8;
			ultraGridColumn68.Header.VisiblePosition = 50;
			ultraGridColumn68.Hidden = true;
			ultraGridColumn68.Width = 8;
			ultraGridColumn69.Header.VisiblePosition = 51;
			ultraGridColumn69.Hidden = true;
			ultraGridColumn69.Width = 8;
			ultraGridColumn70.Header.VisiblePosition = 52;
			ultraGridColumn70.Hidden = true;
			ultraGridColumn70.Width = 8;
			ultraGridColumn71.Header.VisiblePosition = 53;
			ultraGridColumn71.Hidden = true;
			ultraGridColumn71.Width = 8;
			ultraGridColumn72.Header.VisiblePosition = 54;
			ultraGridColumn72.Hidden = true;
			ultraGridColumn72.Width = 8;
			ultraGridColumn73.Header.VisiblePosition = 55;
			ultraGridColumn73.Hidden = true;
			ultraGridColumn73.Width = 8;
			ultraGridColumn74.Header.VisiblePosition = 56;
			ultraGridColumn74.Hidden = true;
			ultraGridColumn74.Width = 8;
			ultraGridColumn75.Header.VisiblePosition = 57;
			ultraGridColumn75.Hidden = true;
			ultraGridColumn75.Width = 8;
			ultraGridColumn76.Header.VisiblePosition = 58;
			ultraGridColumn76.Hidden = true;
			ultraGridColumn76.Width = 8;
			ultraGridColumn77.Header.VisiblePosition = 59;
			ultraGridColumn77.Hidden = true;
			ultraGridColumn77.Width = 8;
			ultraGridColumn78.Header.VisiblePosition = 60;
			ultraGridColumn78.Hidden = true;
			ultraGridColumn78.Width = 8;
			ultraGridColumn79.Header.VisiblePosition = 61;
			ultraGridColumn79.Hidden = true;
			ultraGridColumn79.Width = 8;
			ultraGridColumn80.Header.VisiblePosition = 62;
			ultraGridColumn80.Hidden = true;
			ultraGridColumn80.Width = 8;
			ultraGridColumn81.Header.VisiblePosition = 63;
			ultraGridColumn81.Hidden = true;
			ultraGridColumn81.Width = 8;
			ultraGridColumn82.Header.VisiblePosition = 64;
			ultraGridColumn82.Hidden = true;
			ultraGridColumn82.Width = 8;
			ultraGridColumn83.Header.VisiblePosition = 65;
			ultraGridColumn83.Hidden = true;
			ultraGridColumn83.Width = 8;
			ultraGridColumn84.Header.VisiblePosition = 66;
			ultraGridColumn84.Hidden = true;
			ultraGridColumn84.Width = 8;
			ultraGridColumn85.Header.VisiblePosition = 67;
			ultraGridColumn85.Hidden = true;
			ultraGridColumn85.Width = 8;
			ultraGridColumn86.Header.VisiblePosition = 68;
			ultraGridColumn86.Hidden = true;
			ultraGridColumn86.Width = 8;
			ultraGridColumn87.Header.VisiblePosition = 69;
			ultraGridColumn87.Hidden = true;
			ultraGridColumn87.Width = 8;
			ultraGridColumn88.Header.VisiblePosition = 70;
			ultraGridColumn88.Hidden = true;
			ultraGridColumn88.Width = 8;
			ultraGridColumn89.Header.VisiblePosition = 71;
			ultraGridColumn89.Hidden = true;
			ultraGridColumn89.Width = 8;
			ultraGridColumn90.Header.VisiblePosition = 72;
			ultraGridColumn90.Hidden = true;
			ultraGridColumn90.Width = 8;
			ultraGridColumn91.Header.VisiblePosition = 73;
			ultraGridColumn91.Hidden = true;
			ultraGridColumn91.Width = 8;
			ultraGridColumn92.Header.VisiblePosition = 74;
			ultraGridColumn92.Hidden = true;
			ultraGridColumn92.Width = 8;
			ultraGridColumn93.Header.VisiblePosition = 75;
			ultraGridColumn93.Hidden = true;
			ultraGridColumn93.Width = 8;
			ultraGridColumn94.Header.VisiblePosition = 76;
			ultraGridColumn94.Hidden = true;
			ultraGridColumn94.Width = 8;
			ultraGridColumn95.Header.VisiblePosition = 77;
			ultraGridColumn95.Hidden = true;
			ultraGridColumn95.Width = 8;
			ultraGridColumn96.Header.VisiblePosition = 78;
			ultraGridColumn96.Hidden = true;
			ultraGridColumn96.Width = 8;
			ultraGridColumn97.Header.VisiblePosition = 79;
			ultraGridColumn97.Hidden = true;
			ultraGridColumn97.Width = 8;
			ultraGridColumn98.Header.VisiblePosition = 80;
			ultraGridColumn98.Hidden = true;
			ultraGridColumn98.Width = 8;
			ultraGridColumn99.Header.VisiblePosition = 81;
			ultraGridColumn99.Hidden = true;
			ultraGridColumn99.Width = 8;
			ultraGridColumn100.Header.VisiblePosition = 82;
			ultraGridColumn100.Hidden = true;
			ultraGridColumn100.Width = 16;
			ultraGridColumn101.Header.VisiblePosition = 83;
			ultraGridColumn101.Hidden = true;
			ultraGridColumn101.Width = 16;
			ultraGridColumn102.Header.VisiblePosition = 84;
			ultraGridColumn102.Hidden = true;
			ultraGridColumn102.Width = 16;
			ultraGridColumn103.Header.VisiblePosition = 85;
			ultraGridColumn103.Hidden = true;
			ultraGridColumn103.Width = 16;
			ultraGridColumn104.Header.VisiblePosition = 86;
			ultraGridColumn104.Hidden = true;
			ultraGridColumn104.Width = 8;
			ultraGridColumn105.Header.VisiblePosition = 87;
			ultraGridColumn105.Hidden = true;
			ultraGridColumn105.Width = 8;
			ultraGridColumn106.Header.VisiblePosition = 88;
			ultraGridColumn106.Hidden = true;
			ultraGridColumn106.Width = 8;
			ultraGridColumn107.Header.VisiblePosition = 89;
			ultraGridColumn107.Hidden = true;
			ultraGridColumn107.Width = 8;
			ultraGridColumn108.Header.VisiblePosition = 90;
			ultraGridColumn108.Hidden = true;
			ultraGridColumn108.Width = 8;
			ultraGridColumn109.Header.VisiblePosition = 91;
			ultraGridColumn109.Hidden = true;
			ultraGridColumn109.Width = 8;
			ultraGridColumn110.Header.VisiblePosition = 92;
			ultraGridColumn110.Hidden = true;
			ultraGridColumn110.Width = 8;
			ultraGridColumn111.Header.VisiblePosition = 93;
			ultraGridColumn111.Hidden = true;
			ultraGridColumn111.Width = 8;
			ultraGridColumn112.Header.VisiblePosition = 94;
			ultraGridColumn112.Hidden = true;
			ultraGridColumn112.Width = 16;
			ultraGridColumn113.Header.VisiblePosition = 95;
			ultraGridColumn113.Hidden = true;
			ultraGridColumn113.Width = 16;
			ultraGridColumn114.Header.VisiblePosition = 96;
			ultraGridColumn114.Hidden = true;
			ultraGridColumn114.Width = 16;
			ultraGridColumn115.Header.VisiblePosition = 97;
			ultraGridColumn115.Hidden = true;
			ultraGridColumn115.Width = 16;
			ultraGridColumn116.Header.VisiblePosition = 98;
			ultraGridColumn116.Hidden = true;
			ultraGridColumn116.Width = 16;
			ultraGridColumn117.Header.VisiblePosition = 99;
			ultraGridColumn117.Hidden = true;
			ultraGridColumn117.Width = 16;
			ultraGridColumn118.Header.VisiblePosition = 100;
			ultraGridColumn118.Hidden = true;
			ultraGridColumn118.Width = 8;
			ultraGridColumn119.Header.VisiblePosition = 101;
			ultraGridColumn119.Hidden = true;
			ultraGridColumn119.Width = 8;
			ultraGridColumn120.Header.VisiblePosition = 102;
			ultraGridColumn120.Width = 165;
			ultraGridColumn121.Header.VisiblePosition = 103;
			ultraGridColumn121.Width = 165;
			ultraGridColumn122.Header.VisiblePosition = 104;
			ultraGridColumn122.Hidden = true;
			ultraGridColumn122.Width = 8;
			ultraGridColumn123.Header.VisiblePosition = 105;
			ultraGridColumn123.Hidden = true;
			ultraGridColumn123.Width = 8;
			ultraGridColumn124.Header.VisiblePosition = 106;
			ultraGridColumn124.Hidden = true;
			ultraGridColumn124.Width = 8;
			ultraGridColumn125.Header.VisiblePosition = 107;
			ultraGridColumn125.Hidden = true;
			ultraGridColumn125.Width = 16;
			ultraGridColumn126.Header.VisiblePosition = 108;
			ultraGridColumn126.Width = 59;
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
																										 ultraGridColumn126});
			this.cmbArticulo.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			this.cmbArticulo.DisplayMember = "Articulo";
			this.cmbArticulo.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbArticulo.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbArticulo.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbArticulo.Location = new System.Drawing.Point(96, 88);
			this.cmbArticulo.Name = "cmbArticulo";
			this.cmbArticulo.Size = new System.Drawing.Size(408, 21);
			this.cmbArticulo.TabIndex = 3;
			this.cmbArticulo.ValueMember = "idArticulo";
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
			// label2
			// 
			this.label2.AutoSize = true;
			this.label2.Location = new System.Drawing.Point(24, 88);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(42, 16);
			this.label2.TabIndex = 6;
			this.label2.Text = "Artículo";
			// 
			// label1
			// 
			this.label1.AutoSize = true;
			this.label1.Location = new System.Drawing.Point(24, 24);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(26, 16);
			this.label1.TabIndex = 7;
			this.label1.Text = "Tipo";
			// 
			// btCrea
			// 
			this.btCrea.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.btCrea.Location = new System.Drawing.Point(416, 152);
			this.btCrea.Name = "btCrea";
			this.btCrea.Size = new System.Drawing.Size(96, 23);
			this.btCrea.TabIndex = 8;
			this.btCrea.Text = "Crea  Artículo";
			this.btCrea.Click += new System.EventHandler(this.btCrea_Click);
			// 
			// label3
			// 
			this.label3.AutoSize = true;
			this.label3.Location = new System.Drawing.Point(312, 24);
			this.label3.Name = "label3";
			this.label3.Size = new System.Drawing.Size(31, 16);
			this.label3.TabIndex = 9;
			this.label3.Text = "Valor";
			this.label3.Click += new System.EventHandler(this.label3_Click);
			// 
			// label4
			// 
			this.label4.AutoSize = true;
			this.label4.Location = new System.Drawing.Point(24, 56);
			this.label4.Name = "label4";
			this.label4.Size = new System.Drawing.Size(50, 16);
			this.label4.TabIndex = 12;
			this.label4.Text = "Cantidad";
			// 
			// txtValor
			// 
			this.txtValor.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtValor.FormatString = "#,##0.00";
			this.txtValor.Location = new System.Drawing.Point(368, 24);
			this.txtValor.Name = "txtValor";
			this.txtValor.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtValor.PromptChar = ' ';
			this.txtValor.Size = new System.Drawing.Size(136, 21);
			this.txtValor.TabIndex = 2;
			// 
			// label5
			// 
			this.label5.AutoSize = true;
			this.label5.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label5.Location = new System.Drawing.Point(312, 56);
			this.label5.Name = "label5";
			this.label5.Size = new System.Drawing.Size(31, 17);
			this.label5.TabIndex = 23;
			this.label5.Text = "Nota";
			// 
			// txtNota
			// 
			this.txtNota.Location = new System.Drawing.Point(368, 56);
			this.txtNota.Name = "txtNota";
			this.txtNota.Size = new System.Drawing.Size(136, 20);
			this.txtNota.TabIndex = 5;
			this.txtNota.Text = "";
			// 
			// btBorra
			// 
			this.btBorra.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.btBorra.Location = new System.Drawing.Point(416, 192);
			this.btBorra.Name = "btBorra";
			this.btBorra.Size = new System.Drawing.Size(96, 23);
			this.btBorra.TabIndex = 24;
			this.btBorra.Text = "Borra  Artículo";
			this.btBorra.Click += new System.EventHandler(this.btBorra_Click);
			// 
			// txtCantidad
			// 
			this.txtCantidad.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtCantidad.FormatString = "#,##0.00";
			this.txtCantidad.Location = new System.Drawing.Point(96, 56);
			this.txtCantidad.Name = "txtCantidad";
			this.txtCantidad.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtCantidad.PromptChar = ' ';
			this.txtCantidad.Size = new System.Drawing.Size(144, 21);
			this.txtCantidad.TabIndex = 25;
			// 
			// label6
			// 
			this.label6.AutoSize = true;
			this.label6.Location = new System.Drawing.Point(8, 56);
			this.label6.Name = "label6";
			this.label6.Size = new System.Drawing.Size(40, 16);
			this.label6.TabIndex = 27;
			this.label6.Text = "Cliente";
			// 
			// panel1
			// 
			this.panel1.Controls.Add(this.cmbCliente);
			this.panel1.Controls.Add(this.label7);
			this.panel1.Controls.Add(this.cmbGrupoCli);
			this.panel1.Controls.Add(this.label6);
			this.panel1.Location = new System.Drawing.Point(16, 136);
			this.panel1.Name = "panel1";
			this.panel1.Size = new System.Drawing.Size(384, 96);
			this.panel1.TabIndex = 31;
			// 
			// cmbCliente
			// 
			this.cmbCliente.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbCliente.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbCliente.DataMember = "Cliente";
			this.cmbCliente.DataSource = this.cdsCliente;
			this.cmbCliente.DisplayLayout.AutoFitColumns = true;
			ultraGridBand3.AddButtonCaption = "CompraNumero";
			ultraGridColumn127.Header.VisiblePosition = 0;
			ultraGridColumn127.Hidden = true;
			ultraGridColumn127.Width = 8;
			ultraGridColumn128.Header.VisiblePosition = 1;
			ultraGridColumn128.Hidden = true;
			ultraGridColumn128.Width = 33;
			ultraGridColumn129.Header.VisiblePosition = 2;
			ultraGridColumn129.Hidden = true;
			ultraGridColumn129.Width = 8;
			ultraGridColumn130.Header.VisiblePosition = 3;
			ultraGridColumn130.Hidden = true;
			ultraGridColumn130.Width = 8;
			ultraGridColumn131.Header.VisiblePosition = 5;
			ultraGridColumn131.Hidden = true;
			ultraGridColumn131.Width = 8;
			ultraGridColumn132.Header.VisiblePosition = 6;
			ultraGridColumn132.Hidden = true;
			ultraGridColumn132.Width = 8;
			ultraGridColumn133.Header.VisiblePosition = 7;
			ultraGridColumn133.Hidden = true;
			ultraGridColumn133.Width = 8;
			ultraGridColumn134.Header.VisiblePosition = 8;
			ultraGridColumn134.Hidden = true;
			ultraGridColumn134.Width = 8;
			ultraGridColumn135.Header.VisiblePosition = 9;
			ultraGridColumn135.Hidden = true;
			ultraGridColumn135.Width = 8;
			ultraGridColumn136.Header.VisiblePosition = 10;
			ultraGridColumn136.Hidden = true;
			ultraGridColumn136.Width = 8;
			ultraGridColumn137.Header.VisiblePosition = 11;
			ultraGridColumn137.Hidden = true;
			ultraGridColumn137.Width = 8;
			ultraGridColumn138.Header.VisiblePosition = 12;
			ultraGridColumn138.Hidden = true;
			ultraGridColumn138.Width = 8;
			ultraGridColumn139.Header.VisiblePosition = 13;
			ultraGridColumn139.Width = 94;
			ultraGridColumn140.Header.VisiblePosition = 14;
			ultraGridColumn140.Width = 94;
			ultraGridColumn141.Header.VisiblePosition = 15;
			ultraGridColumn141.Hidden = true;
			ultraGridColumn141.Width = 8;
			ultraGridColumn142.Header.VisiblePosition = 16;
			ultraGridColumn142.Hidden = true;
			ultraGridColumn142.Width = 8;
			ultraGridColumn143.Header.VisiblePosition = 17;
			ultraGridColumn143.Hidden = true;
			ultraGridColumn143.Width = 8;
			ultraGridColumn144.Header.VisiblePosition = 18;
			ultraGridColumn144.Hidden = true;
			ultraGridColumn144.Width = 8;
			ultraGridColumn145.Header.VisiblePosition = 19;
			ultraGridColumn145.Hidden = true;
			ultraGridColumn145.Width = 8;
			ultraGridColumn146.Header.VisiblePosition = 20;
			ultraGridColumn146.Hidden = true;
			ultraGridColumn146.Width = 8;
			ultraGridColumn147.Header.VisiblePosition = 21;
			ultraGridColumn147.Hidden = true;
			ultraGridColumn147.Width = 8;
			ultraGridColumn148.Header.VisiblePosition = 22;
			ultraGridColumn148.Hidden = true;
			ultraGridColumn148.Width = 8;
			ultraGridColumn149.Header.VisiblePosition = 23;
			ultraGridColumn149.Hidden = true;
			ultraGridColumn149.Width = 8;
			ultraGridColumn150.Header.VisiblePosition = 24;
			ultraGridColumn150.Hidden = true;
			ultraGridColumn150.Width = 8;
			ultraGridColumn151.Header.VisiblePosition = 25;
			ultraGridColumn151.Hidden = true;
			ultraGridColumn151.Width = 8;
			ultraGridColumn152.Header.VisiblePosition = 26;
			ultraGridColumn152.Hidden = true;
			ultraGridColumn152.Width = 16;
			ultraGridColumn153.Header.VisiblePosition = 27;
			ultraGridColumn153.Hidden = true;
			ultraGridColumn153.Width = 16;
			ultraGridColumn154.Header.VisiblePosition = 28;
			ultraGridColumn154.Hidden = true;
			ultraGridColumn154.Width = 8;
			ultraGridColumn155.Header.VisiblePosition = 29;
			ultraGridColumn155.Hidden = true;
			ultraGridColumn155.Width = 8;
			ultraGridColumn156.Header.VisiblePosition = 30;
			ultraGridColumn156.Hidden = true;
			ultraGridColumn156.Width = 8;
			ultraGridColumn157.Header.VisiblePosition = 31;
			ultraGridColumn157.Hidden = true;
			ultraGridColumn157.Width = 8;
			ultraGridColumn158.Header.VisiblePosition = 32;
			ultraGridColumn158.Hidden = true;
			ultraGridColumn158.Width = 8;
			ultraGridColumn159.Header.VisiblePosition = 33;
			ultraGridColumn159.Hidden = true;
			ultraGridColumn159.Width = 16;
			ultraGridColumn160.Header.VisiblePosition = 34;
			ultraGridColumn160.Hidden = true;
			ultraGridColumn160.Width = 16;
			ultraGridColumn161.Header.VisiblePosition = 35;
			ultraGridColumn161.Hidden = true;
			ultraGridColumn161.Width = 8;
			ultraGridColumn162.Header.VisiblePosition = 36;
			ultraGridColumn162.Hidden = true;
			ultraGridColumn162.Width = 8;
			ultraGridColumn163.Header.VisiblePosition = 4;
			ultraGridColumn163.Hidden = true;
			ultraGridColumn163.Width = 19;
			ultraGridColumn164.Header.VisiblePosition = 37;
			ultraGridColumn164.Hidden = true;
			ultraGridColumn164.Width = 8;
			ultraGridColumn165.Header.VisiblePosition = 38;
			ultraGridColumn165.Hidden = true;
			ultraGridColumn165.Width = 8;
			ultraGridColumn166.Header.VisiblePosition = 39;
			ultraGridColumn166.Hidden = true;
			ultraGridColumn166.Width = 16;
			ultraGridColumn167.Header.VisiblePosition = 40;
			ultraGridColumn167.Hidden = true;
			ultraGridColumn167.Width = 16;
			ultraGridColumn168.Header.VisiblePosition = 41;
			ultraGridColumn168.Hidden = true;
			ultraGridColumn168.Width = 16;
			ultraGridColumn169.Header.VisiblePosition = 42;
			ultraGridColumn169.Hidden = true;
			ultraGridColumn169.Width = 8;
			ultraGridColumn170.Header.VisiblePosition = 43;
			ultraGridColumn170.Hidden = true;
			ultraGridColumn170.Width = 8;
			ultraGridColumn171.Header.VisiblePosition = 44;
			ultraGridColumn171.Hidden = true;
			ultraGridColumn171.Width = 16;
			ultraGridColumn172.Header.VisiblePosition = 45;
			ultraGridColumn172.Hidden = true;
			ultraGridColumn172.Width = 8;
			ultraGridColumn173.Header.VisiblePosition = 46;
			ultraGridColumn173.Hidden = true;
			ultraGridColumn173.Width = 8;
			ultraGridColumn174.Header.VisiblePosition = 47;
			ultraGridColumn174.Hidden = true;
			ultraGridColumn174.Width = 16;
			ultraGridColumn175.Header.VisiblePosition = 48;
			ultraGridColumn175.Hidden = true;
			ultraGridColumn175.Width = 8;
			ultraGridColumn176.Header.VisiblePosition = 49;
			ultraGridColumn176.Hidden = true;
			ultraGridColumn176.Width = 8;
			ultraGridColumn177.Header.VisiblePosition = 50;
			ultraGridColumn177.Hidden = true;
			ultraGridColumn177.Width = 8;
			ultraGridColumn178.Header.VisiblePosition = 51;
			ultraGridColumn178.Hidden = true;
			ultraGridColumn178.Width = 8;
			ultraGridColumn179.Header.VisiblePosition = 52;
			ultraGridColumn179.Hidden = true;
			ultraGridColumn179.Width = 8;
			ultraGridColumn180.Header.VisiblePosition = 53;
			ultraGridColumn180.Hidden = true;
			ultraGridColumn180.Width = 8;
			ultraGridColumn181.Header.VisiblePosition = 54;
			ultraGridColumn181.Hidden = true;
			ultraGridColumn181.Width = 8;
			ultraGridColumn182.Header.VisiblePosition = 55;
			ultraGridColumn182.Hidden = true;
			ultraGridColumn182.Width = 8;
			ultraGridColumn183.Header.VisiblePosition = 56;
			ultraGridColumn183.Hidden = true;
			ultraGridColumn183.Width = 8;
			ultraGridColumn184.Header.VisiblePosition = 57;
			ultraGridColumn184.Hidden = true;
			ultraGridColumn184.Width = 8;
			ultraGridColumn185.Header.VisiblePosition = 58;
			ultraGridColumn185.Hidden = true;
			ultraGridColumn185.Width = 8;
			ultraGridColumn186.Header.VisiblePosition = 59;
			ultraGridColumn186.Hidden = true;
			ultraGridColumn186.Width = 8;
			ultraGridColumn187.Header.VisiblePosition = 60;
			ultraGridColumn187.Hidden = true;
			ultraGridColumn187.Width = 8;
			ultraGridColumn188.Header.VisiblePosition = 61;
			ultraGridColumn188.Hidden = true;
			ultraGridColumn188.Width = 16;
			ultraGridColumn189.Header.VisiblePosition = 62;
			ultraGridColumn189.Hidden = true;
			ultraGridColumn189.Width = 16;
			ultraGridColumn190.Header.VisiblePosition = 63;
			ultraGridColumn190.Hidden = true;
			ultraGridColumn190.Width = 8;
			ultraGridColumn191.Header.VisiblePosition = 64;
			ultraGridColumn191.Hidden = true;
			ultraGridColumn191.Width = 16;
			ultraGridColumn192.Header.VisiblePosition = 65;
			ultraGridColumn192.Hidden = true;
			ultraGridColumn192.Width = 8;
			ultraGridColumn193.Header.VisiblePosition = 66;
			ultraGridColumn193.Hidden = true;
			ultraGridColumn193.Width = 16;
			ultraGridColumn194.Header.VisiblePosition = 67;
			ultraGridColumn194.Hidden = true;
			ultraGridColumn194.Width = 8;
			ultraGridColumn195.Header.VisiblePosition = 68;
			ultraGridColumn195.Hidden = true;
			ultraGridColumn195.Width = 16;
			ultraGridColumn196.Header.VisiblePosition = 69;
			ultraGridColumn196.Hidden = true;
			ultraGridColumn196.Width = 8;
			ultraGridColumn197.Header.VisiblePosition = 70;
			ultraGridColumn197.Hidden = true;
			ultraGridColumn197.Width = 8;
			ultraGridColumn198.Header.VisiblePosition = 71;
			ultraGridColumn198.Width = 49;
			ultraGridBand3.Columns.AddRange(new object[] {
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
																										 ultraGridColumn198});
			this.cmbCliente.DisplayLayout.BandsSerializer.Add(ultraGridBand3);
			this.cmbCliente.DisplayMember = "Nombre";
			this.cmbCliente.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbCliente.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbCliente.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbCliente.Location = new System.Drawing.Point(96, 56);
			this.cmbCliente.Name = "cmbCliente";
			this.cmbCliente.Size = new System.Drawing.Size(256, 21);
			this.cmbCliente.TabIndex = 30;
			this.cmbCliente.ValueMember = "idCliente";
			// 
			// cdsCliente
			// 
			this.cdsCliente.BindingContextCtrl = this;
			this.cdsCliente.DataLibrary = "LibFacturacion";
			this.cdsCliente.DataLibraryUrl = "";
			this.cdsCliente.DataSetDef = "dsCliente";
			this.cdsCliente.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsCliente.Name = "cdsCliente";
			this.cdsCliente.SchemaClassName = "LibFacturacion.DataClass";
			this.cdsCliente.SchemaDef = null;
			this.cdsCliente.BeforeFill += new C1.Data.FillEventHandler(this.cdsCliente_BeforeFill);
			// 
			// label7
			// 
			this.label7.AutoSize = true;
			this.label7.Location = new System.Drawing.Point(8, 16);
			this.label7.Name = "label7";
			this.label7.Size = new System.Drawing.Size(74, 16);
			this.label7.TabIndex = 29;
			this.label7.Text = "Grupo Cliente";
			// 
			// cmbGrupoCli
			// 
			this.cmbGrupoCli.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbGrupoCli.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbGrupoCli.DataMember = "ClienteGrupo";
			this.cmbGrupoCli.DataSource = this.cdsGrupoCliente;
			this.cmbGrupoCli.DisplayLayout.AutoFitColumns = true;
			ultraGridBand4.AddButtonCaption = "CompraNumero";
			ultraGridColumn199.Header.VisiblePosition = 0;
			ultraGridColumn199.Hidden = true;
			ultraGridColumn199.Width = 33;
			ultraGridColumn200.Header.VisiblePosition = 1;
			ultraGridColumn200.Width = 237;
			ultraGridColumn201.Header.VisiblePosition = 2;
			ultraGridColumn201.Hidden = true;
			ultraGridColumn201.Width = 22;
			ultraGridColumn202.Header.VisiblePosition = 3;
			ultraGridColumn202.Hidden = true;
			ultraGridColumn202.Width = 19;
			ultraGridBand4.Columns.AddRange(new object[] {
																										 ultraGridColumn199,
																										 ultraGridColumn200,
																										 ultraGridColumn201,
																										 ultraGridColumn202});
			this.cmbGrupoCli.DisplayLayout.BandsSerializer.Add(ultraGridBand4);
			this.cmbGrupoCli.DisplayMember = "Grupo";
			this.cmbGrupoCli.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbGrupoCli.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbGrupoCli.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbGrupoCli.Location = new System.Drawing.Point(96, 16);
			this.cmbGrupoCli.Name = "cmbGrupoCli";
			this.cmbGrupoCli.Size = new System.Drawing.Size(256, 21);
			this.cmbGrupoCli.TabIndex = 28;
			this.cmbGrupoCli.ValueMember = "idGrupoCliente";
			// 
			// cdsGrupoCliente
			// 
			this.cdsGrupoCliente.BindingContextCtrl = this;
			this.cdsGrupoCliente.DataLibrary = "LibFacturacion";
			this.cdsGrupoCliente.DataLibraryUrl = "";
			this.cdsGrupoCliente.DataSetDef = "dsClienteTabla";
			this.cdsGrupoCliente.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsGrupoCliente.Name = "cdsGrupoCliente";
			this.cdsGrupoCliente.SchemaClassName = "LibFacturacion.DataClass";
			this.cdsGrupoCliente.SchemaDef = null;
			this.cdsGrupoCliente.BeforeFill += new C1.Data.FillEventHandler(this.cdsGrupoCliente_BeforeFill);
			// 
			// panel2
			// 
			this.panel2.Controls.Add(this.cmbGrupoCliH);
			this.panel2.Controls.Add(this.label8);
			this.panel2.Controls.Add(this.label9);
			this.panel2.Controls.Add(this.cmbGrupoCliD);
			this.panel2.Location = new System.Drawing.Point(16, 256);
			this.panel2.Name = "panel2";
			this.panel2.Size = new System.Drawing.Size(384, 96);
			this.panel2.TabIndex = 32;
			this.panel2.Paint += new System.Windows.Forms.PaintEventHandler(this.panel2_Paint);
			// 
			// cmbGrupoCliH
			// 
			this.cmbGrupoCliH.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbGrupoCliH.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbGrupoCliH.DataMember = "ClienteGrupo";
			this.cmbGrupoCliH.DataSource = this.cdsGrupoCliente;
			this.cmbGrupoCliH.DisplayLayout.AutoFitColumns = true;
			ultraGridBand5.AddButtonCaption = "CompraNumero";
			ultraGridColumn203.Header.VisiblePosition = 0;
			ultraGridColumn203.Hidden = true;
			ultraGridColumn203.Width = 33;
			ultraGridColumn204.Header.VisiblePosition = 1;
			ultraGridColumn204.Width = 237;
			ultraGridColumn205.Header.VisiblePosition = 2;
			ultraGridColumn205.Hidden = true;
			ultraGridColumn205.Width = 22;
			ultraGridColumn206.Header.VisiblePosition = 3;
			ultraGridColumn206.Hidden = true;
			ultraGridColumn206.Width = 19;
			ultraGridBand5.Columns.AddRange(new object[] {
																										 ultraGridColumn203,
																										 ultraGridColumn204,
																										 ultraGridColumn205,
																										 ultraGridColumn206});
			this.cmbGrupoCliH.DisplayLayout.BandsSerializer.Add(ultraGridBand5);
			this.cmbGrupoCliH.DisplayMember = "Grupo";
			this.cmbGrupoCliH.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbGrupoCliH.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbGrupoCliH.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbGrupoCliH.Location = new System.Drawing.Point(96, 56);
			this.cmbGrupoCliH.Name = "cmbGrupoCliH";
			this.cmbGrupoCliH.Size = new System.Drawing.Size(256, 21);
			this.cmbGrupoCliH.TabIndex = 31;
			this.cmbGrupoCliH.ValueMember = "idGrupoCliente";
			this.cmbGrupoCliH.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.ultraCombo2_InitializeLayout);
			// 
			// label8
			// 
			this.label8.AutoSize = true;
			this.label8.Location = new System.Drawing.Point(8, 24);
			this.label8.Name = "label8";
			this.label8.Size = new System.Drawing.Size(71, 16);
			this.label8.TabIndex = 29;
			this.label8.Text = "Grupo Desde";
			this.label8.Click += new System.EventHandler(this.label8_Click);
			// 
			// label9
			// 
			this.label9.AutoSize = true;
			this.label9.Location = new System.Drawing.Point(8, 56);
			this.label9.Name = "label9";
			this.label9.Size = new System.Drawing.Size(68, 16);
			this.label9.TabIndex = 27;
			this.label9.Text = "Grupo Hasta";
			// 
			// cmbGrupoCliD
			// 
			this.cmbGrupoCliD.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbGrupoCliD.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbGrupoCliD.DataMember = "ClienteGrupo";
			this.cmbGrupoCliD.DataSource = this.cdsGrupoCliente;
			this.cmbGrupoCliD.DisplayLayout.AutoFitColumns = true;
			ultraGridBand6.AddButtonCaption = "CompraNumero";
			ultraGridColumn207.Header.VisiblePosition = 0;
			ultraGridColumn207.Hidden = true;
			ultraGridColumn207.Width = 33;
			ultraGridColumn208.Header.VisiblePosition = 1;
			ultraGridColumn208.Width = 237;
			ultraGridColumn209.Header.VisiblePosition = 2;
			ultraGridColumn209.Hidden = true;
			ultraGridColumn209.Width = 22;
			ultraGridColumn210.Header.VisiblePosition = 3;
			ultraGridColumn210.Hidden = true;
			ultraGridColumn210.Width = 19;
			ultraGridBand6.Columns.AddRange(new object[] {
																										 ultraGridColumn207,
																										 ultraGridColumn208,
																										 ultraGridColumn209,
																										 ultraGridColumn210});
			this.cmbGrupoCliD.DisplayLayout.BandsSerializer.Add(ultraGridBand6);
			this.cmbGrupoCliD.DisplayMember = "Grupo";
			this.cmbGrupoCliD.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbGrupoCliD.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbGrupoCliD.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbGrupoCliD.Location = new System.Drawing.Point(96, 24);
			this.cmbGrupoCliD.Name = "cmbGrupoCliD";
			this.cmbGrupoCliD.Size = new System.Drawing.Size(256, 21);
			this.cmbGrupoCliD.TabIndex = 30;
			this.cmbGrupoCliD.ValueMember = "idGrupoCliente";
			// 
			// ultraButton1
			// 
			this.ultraButton1.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.ultraButton1.Location = new System.Drawing.Point(416, 280);
			this.ultraButton1.Name = "ultraButton1";
			this.ultraButton1.Size = new System.Drawing.Size(88, 23);
			this.ultraButton1.TabIndex = 33;
			this.ultraButton1.Text = "Cambiar";
			this.ultraButton1.Click += new System.EventHandler(this.ultraButton1_Click);
			// 
			// label10
			// 
			this.label10.AutoSize = true;
			this.label10.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label10.Location = new System.Drawing.Point(24, 248);
			this.label10.Name = "label10";
			this.label10.Size = new System.Drawing.Size(97, 16);
			this.label10.TabIndex = 34;
			this.label10.Text = "Cambio de Grupo";
			// 
			// CompraInsertaArticulo
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(552, 365);
			this.Controls.Add(this.label10);
			this.Controls.Add(this.ultraButton1);
			this.Controls.Add(this.panel1);
			this.Controls.Add(this.txtCantidad);
			this.Controls.Add(this.btBorra);
			this.Controls.Add(this.label5);
			this.Controls.Add(this.txtNota);
			this.Controls.Add(this.label4);
			this.Controls.Add(this.label3);
			this.Controls.Add(this.label1);
			this.Controls.Add(this.label2);
			this.Controls.Add(this.txtValor);
			this.Controls.Add(this.btCrea);
			this.Controls.Add(this.cmbArticulo);
			this.Controls.Add(this.cmbTipo);
			this.Controls.Add(this.panel2);
			this.Menu = this.mainMenu1;
			this.Name = "CompraInsertaArticulo";
			this.Text = "Compra Inserta Artículo";
			this.Load += new System.EventHandler(this.CompraInsertaArticulo_Load);
			((System.ComponentModel.ISupportInitialize)(this.cdsCompraNumero)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbTipo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbArticulo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsArticulo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtValor)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCantidad)).EndInit();
			this.panel1.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.cmbCliente)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsCliente)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbGrupoCli)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsGrupoCliente)).EndInit();
			this.panel2.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.cmbGrupoCliH)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbGrupoCliD)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void btCrea_Click(object sender, System.EventArgs e)
		{
			Cursor = Cursors.WaitCursor;
			if (Funcion.bEjecutaSQL(cdsCompraNumero, "Exec SeteoGExiste 'FLCA'"))
			{
				int iArt = 0;
				int iTipo = 0;
				int iGrupo = 0;
				int iEst = 0;

				if (cmbArticulo.Value != null) iArt = (int) cmbArticulo.Value;			
				if (cmbTipo.Value != null) iTipo = (int) cmbTipo.Value;
				if (cmbGrupoCli.Value != null) iGrupo = (int) cmbGrupoCli.Value;
				if (cmbCliente.Value != null) iEst = (int) cmbCliente.Value;

				string stProcedimiento = "Exec CompraInsertaArticulo " 
					+ iArt.ToString() 
					+ "," + iTipo.ToString()
					+ "," + txtCantidad.Value.ToString()
					+ "," + txtValor.Value.ToString()
					+ "," + iGrupo.ToString()
					+ "," + iEst.ToString()
					+ ",'" + txtNota.Text +"'";

				Funcion.sEscalarSQL(cdsArticulo,stProcedimiento, true);
				MessageBox.Show("Proceso Terminado", "Información", MessageBoxButtons.OK);
				Cursor = Cursors.Default;
				return;
			}
			Cursor = Cursors.Default;
			MessageBox.Show("Proceso Restringido", "Información", MessageBoxButtons.OK);
			return;
		}

		private void label3_Click(object sender, System.EventArgs e)
		{
		
		}

		private void CompraInsertaArticulo_Load(object sender, System.EventArgs e)
		{
		
		}

		private void cdsCompraNumero_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			this.cdsCompraNumero.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirRoles);

		}

		private void cdsArticulo_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			this.cdsArticulo.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirRoles);

		}
		private void btBorra_Click(object sender, System.EventArgs e)
		{
			Cursor = Cursors.WaitCursor;
			if (Funcion.bEjecutaSQL(cdsCompraNumero, "Exec SeteoGExiste 'FLCA'"))
			{
				int iArt = 0;
				int iTipo = 0;
				int iGrupo = 0;
				int iEst = 0;

				if (cmbArticulo.Value != null) iArt = (int) cmbArticulo.Value;			
				if (cmbTipo.Value != null) iTipo = (int) cmbTipo.Value;
				if (cmbGrupoCli.Value != null) iGrupo = (int) cmbGrupoCli.Value;
				if (cmbCliente.Value != null) iEst = (int) cmbCliente.Value;

				string stProcedimiento = "Exec CompraBorraArticulo " 
					+ iArt.ToString() 
					+ "," + iTipo.ToString()
					+ "," + txtCantidad.Value.ToString()
					+ "," + txtValor.Value.ToString()
					+ "," + iGrupo.ToString()
					+ "," + iEst.ToString()
					+ ",'" + txtNota.Text +"'";

				Funcion.sEscalarSQL(cdsArticulo,stProcedimiento, true);
				MessageBox.Show("Proceso Terminado", "Información", MessageBoxButtons.OK);
				Cursor = Cursors.Default;
				return;
			}
			Cursor = Cursors.Default;
			MessageBox.Show("Proceso Restringido", "Información", MessageBoxButtons.OK);
			return;
		}

		private void chGrupoEst_CheckedChanged(object sender, System.EventArgs e)
		{
//			if (chGrupoEst.Checked = true)
//			{
//				label6.Text = "Grupo";
//			}
//			else
//				label6.Text = "Estudiante";

		}

		private void label8_Click(object sender, System.EventArgs e)
		{
		
		}

		private void panel2_Paint(object sender, System.Windows.Forms.PaintEventArgs e)
		{
		
		}

		private void cdsCliente_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			this.cdsCliente.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirRoles);

		}

		private void cdsGrupoCliente_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			this.cdsGrupoCliente.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirRoles);

		}

		private void ultraButton1_Click(object sender, System.EventArgs e)
		{
			Cursor = Cursors.WaitCursor;
			if (Funcion.bEjecutaSQL(cdsCompraNumero, "Exec SeteoGExiste 'FLCACG'"))
			{
				int iGrupoD = 0;
				int iGrupoH = 0;

				if (cmbGrupoCliD.Value != null) iGrupoD = (int) cmbGrupoCliD.Value;
				if (cmbGrupoCliH.Value != null) iGrupoH = (int) cmbGrupoCliH.Value;
				
				string stSelect = "Update Cliente set idGrupoCliente = " 
					+ iGrupoH.ToString() 
					+ " where idGrupoCliente = "
					+ iGrupoD.ToString();

				Funcion.EjecutaSQL(cdsCliente,stSelect);
				MessageBox.Show("Proceso Terminado", "Información", MessageBoxButtons.OK);
				Cursor = Cursors.Default;
				return;
			}
			Cursor = Cursors.Default;
			MessageBox.Show("Proceso Restringido", "Información", MessageBoxButtons.OK);
			return;
		}

		private void ultraCombo2_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}


	}
}
