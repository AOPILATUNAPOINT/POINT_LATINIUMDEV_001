using System;
using System.Data.SqlClient;
using System.Globalization;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using Infragistics.Win.UltraWinGrid;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de CambioConstantes.
	/// </summary>
	public class CambioConstantes : DevExpress.XtraEditors.XtraForm
	{
		private Infragistics.Win.Misc.UltraButton btAceptar;
		private Infragistics.Win.Misc.UltraButton btCancelar;
		private C1.Data.C1DataSet cdsPrecio;
		private C1.Data.C1DataSet cdsArticulo;
		private Infragistics.Win.UltraWinGrid.UltraGrid ultraGrid1;
		private C1.Data.C1DataSet cdsArticuloTabla;
		private System.Windows.Forms.ErrorProvider errorProvider;
		private Infragistics.Win.UltraWinTabControl.UltraTabControl ultraTabControl1;
		private Infragistics.Win.UltraWinTabControl.UltraTabSharedControlsPage ultraTabSharedControlsPage1;
		private Infragistics.Win.UltraWinTabControl.UltraTabPageControl ultraTabPageControl1;
		private Infragistics.Win.UltraWinTabControl.UltraTabPageControl ultraTabPageControl2;
		private Infragistics.Win.Misc.UltraButton btCambiaIVA;
		private Infragistics.Win.Misc.UltraLabel ultraLabel1;
		private Infragistics.Win.Misc.UltraLabel ultraLabel2;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtIvaAntes;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtIvaDespues;
		private Infragistics.Win.Misc.UltraLabel ultraLabel3;
		private Infragistics.Win.Misc.UltraLabel ultraLabel4;
		private Infragistics.Win.UltraWinEditors.UltraComboEditor cmbSigno;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbPrecio1;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtPorcIncr;
		private Infragistics.Win.Misc.UltraLabel ultraLabel5;
		private Infragistics.Win.Misc.UltraButton btIncPorc;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbPrecio2;
		private System.Windows.Forms.Panel panel3;
		private System.Windows.Forms.Panel panel4;
		private System.Windows.Forms.Panel panel5;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbMarca;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbGrupo;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbColor;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkArancel;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkMarca;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbSubGrupo;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbArancel;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkGrupo;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkColor;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbModelo;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkModelo;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkSubGrupo;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbPrecio1F;
		private Infragistics.Win.Misc.UltraLabel ultraLabel6;
		private Infragistics.Win.Misc.UltraButton btCambiaPrecioValor;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtPrecioF;
		private Infragistics.Win.Misc.UltraLabel ultraLabel7;
		private Infragistics.Win.Misc.UltraButton btCosto;
		private Infragistics.Win.Misc.UltraButton btPrecio;
		private Infragistics.Win.Misc.UltraButton btGrabar;
		private Infragistics.Win.UltraWinTabControl.UltraTabPageControl ultraTabPageControl3;
		private Infragistics.Win.Misc.UltraButton btKardex;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource udsKardex;
		private Infragistics.Win.UltraWinGrid.UltraGrid ugKardex;
		private Infragistics.Win.Misc.UltraButton btFactura;
		private Infragistics.Win.Misc.UltraLabel ultraLabel8;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtDescuento;
		private Infragistics.Win.Misc.UltraButton btDescuento;
		private C1.Data.C1DataView cdvGrupo;
		private Infragistics.Win.Misc.UltraLabel ultraLabel10;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtDescCredito;
		private Infragistics.Win.Misc.UltraButton btCliente;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public CambioConstantes()
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
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("ClientePrecio", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idTipoPrecio");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Precio");
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.ValueListItem valueListItem1 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem2 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("ClientePrecio", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idTipoPrecio");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Precio");
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand3 = new Infragistics.Win.UltraWinGrid.UltraGridBand("ClientePrecio", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idTipoPrecio");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Precio");
			Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand4 = new Infragistics.Win.UltraWinGrid.UltraGridBand("ArticuloMarca", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idMarca");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Marca");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.Appearance appearance11 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand5 = new Infragistics.Win.UltraWinGrid.UltraGridBand("ArticuloGrupo", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idGrupoArticulo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn11 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idTipoGrupo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn12 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Grupo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn13 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCuentas2");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn14 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCuentas3");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn15 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCuentas4");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn16 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Borrar");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn17 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCuentas5");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn18 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCuentas6");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn19 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idArticuloSRI");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn20 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn21 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCuentas7");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn22 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCuentas8");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn23 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCuentas9");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn24 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Incr1");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn25 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Incr2");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn26 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Incr3");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn27 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Incr4");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn28 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Incr5");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn29 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Incr6");
			Infragistics.Win.Appearance appearance12 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand6 = new Infragistics.Win.UltraWinGrid.UltraGridBand("ArticuloColor", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn30 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idColor");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn31 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Color");
			Infragistics.Win.Appearance appearance13 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance14 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance15 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand7 = new Infragistics.Win.UltraWinGrid.UltraGridBand("ArticuloSubGrupo", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn32 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idSubGrupo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn33 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("SubGrupo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn34 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.Appearance appearance16 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand8 = new Infragistics.Win.UltraWinGrid.UltraGridBand("ArticuloArancel", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn35 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idArancel");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn36 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn37 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Porcentaje");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn38 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Arancel");
			Infragistics.Win.Appearance appearance17 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance18 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance19 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand9 = new Infragistics.Win.UltraWinGrid.UltraGridBand("ArticuloModelo", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn39 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idModelo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn40 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Modelo");
			Infragistics.Win.Appearance appearance20 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance21 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance22 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance23 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance24 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance25 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance26 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand10 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Articulo", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn41 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idArticulo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn42 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idArticuloSRI");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn43 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idColor");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn44 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idModelo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn45 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idGrupoArticulo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn46 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idSubGrupo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn47 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idArancel");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn48 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idMarca");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn49 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idUnidad");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn50 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPromocion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn51 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idTipoGrupo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn52 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idClase");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn53 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idSeccion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn54 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idSucursal");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn55 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPadre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn56 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaIngreso");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn57 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CodAlterno");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn58 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Saldo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn59 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Peso");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn60 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Caracteristica");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn61 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Aplicacion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn62 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Iva");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn63 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Ice");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn64 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Precio1");
			Infragistics.Win.Appearance appearance27 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn65 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Precio2");
			Infragistics.Win.Appearance appearance28 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn66 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Precio3");
			Infragistics.Win.Appearance appearance29 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn67 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Precio4");
			Infragistics.Win.Appearance appearance30 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn68 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Precio5");
			Infragistics.Win.Appearance appearance31 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn69 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Precio6");
			Infragistics.Win.Appearance appearance32 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn70 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CantidadMinima");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn71 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CantidadMaxima");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn72 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Residual");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn73 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CostoUltimo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn74 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaUltimo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn75 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCompra1");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn76 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCompra2");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn77 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCompra3");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn78 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CostoPromedio");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn79 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Gramos");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn80 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ConsumoProm");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn81 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("TiempoReposicion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn82 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("UnidadPeso");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn83 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Depreciacion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn84 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Incr1");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn85 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Incr2");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn86 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Incr3");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn87 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Incr4");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn88 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Incr5");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn89 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Incr6");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn90 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Ubicacion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn91 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("NumeroParte");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn92 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ComisionValor");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn93 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ComisionPorcentaje");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn94 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CantidadRequerida");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn95 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("PrecioPromocion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn96 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CantidadPromocion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn97 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("DescuentoContado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn98 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("DescuentoCredito");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn99 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("DescPorcContado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn100 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("DescPorcCre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn101 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CantReqDesc");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn102 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CantReqPorc");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn103 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("TipoPromocion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn104 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Responsable");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn105 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Comision30");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn106 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Comision60");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn107 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Comision90");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn108 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Comision120");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn109 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bodega1");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn110 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bodega2");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn111 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bodega3");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn112 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bodega4");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn113 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bodega5");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn114 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bodega6");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn115 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bodega7");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn116 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bodega8");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn117 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bodega9");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn118 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("SaldoT");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn119 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Alto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn120 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Ancho");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn121 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Largo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn122 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Color");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn123 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Descontinuado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn124 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Borrar");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn125 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Habilitado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn126 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ValorPeso");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn127 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("BodegaPredef");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn128 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Tiempo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn129 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cantidad1");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn130 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cantidad2");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn131 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cantidad3");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn132 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cantidad4");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn133 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cantidad5");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn134 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("PaqueteUnidad");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn135 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Editar");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn136 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nuevo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn137 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaEditado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn138 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CambiarPrecio");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn139 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("dCompra");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn140 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("dVenta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn141 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Foto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn142 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Notas");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn143 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Articulo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn144 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn145 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("BodegaPredefCR");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn146 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CodPromocion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn147 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CostoFecha");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn148 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Aportable");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn149 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ExigeLote");
			Infragistics.Win.Appearance appearance33 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance34 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance35 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance36 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance37 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand11 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn150 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Fecha");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn151 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Ingresos");
			Infragistics.Win.Appearance appearance38 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn152 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Egresos");
			Infragistics.Win.Appearance appearance39 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn153 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Numero");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn154 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Precio");
			Infragistics.Win.Appearance appearance40 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn155 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Tipo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn156 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCompra");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn157 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Costo");
			Infragistics.Win.Appearance appearance41 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance42 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance43 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance44 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance45 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Fecha");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Ingresos");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn3 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Egresos");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn4 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Numero");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn5 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Precio");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn6 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Tipo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn7 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCompra");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn8 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Costo");
			Infragistics.Win.UltraWinTabControl.UltraTab ultraTab1 = new Infragistics.Win.UltraWinTabControl.UltraTab();
			Infragistics.Win.UltraWinTabControl.UltraTab ultraTab2 = new Infragistics.Win.UltraWinTabControl.UltraTab();
			Infragistics.Win.UltraWinTabControl.UltraTab ultraTab3 = new Infragistics.Win.UltraWinTabControl.UltraTab();
			this.ultraTabPageControl1 = new Infragistics.Win.UltraWinTabControl.UltraTabPageControl();
			this.panel4 = new System.Windows.Forms.Panel();
			this.ultraLabel7 = new Infragistics.Win.Misc.UltraLabel();
			this.btCambiaPrecioValor = new Infragistics.Win.Misc.UltraButton();
			this.txtPrecioF = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.ultraLabel6 = new Infragistics.Win.Misc.UltraLabel();
			this.cmbPrecio1F = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.cdsPrecio = new C1.Data.C1DataSet();
			this.txtPorcIncr = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.cmbSigno = new Infragistics.Win.UltraWinEditors.UltraComboEditor();
			this.ultraLabel4 = new Infragistics.Win.Misc.UltraLabel();
			this.ultraLabel5 = new Infragistics.Win.Misc.UltraLabel();
			this.cmbPrecio1 = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.btIncPorc = new Infragistics.Win.Misc.UltraButton();
			this.cmbPrecio2 = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.panel3 = new System.Windows.Forms.Panel();
			this.cmbMarca = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.cdsArticuloTabla = new C1.Data.C1DataSet();
			this.cmbGrupo = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.cdvGrupo = new C1.Data.C1DataView();
			this.cmbColor = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.chkArancel = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.chkMarca = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.cmbSubGrupo = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.cmbArancel = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.chkGrupo = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.chkColor = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.cmbModelo = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.chkModelo = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.chkSubGrupo = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.btCancelar = new Infragistics.Win.Misc.UltraButton();
			this.btAceptar = new Infragistics.Win.Misc.UltraButton();
			this.panel5 = new System.Windows.Forms.Panel();
			this.ultraLabel10 = new Infragistics.Win.Misc.UltraLabel();
			this.txtDescCredito = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.btDescuento = new Infragistics.Win.Misc.UltraButton();
			this.ultraLabel8 = new Infragistics.Win.Misc.UltraLabel();
			this.txtDescuento = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.btCambiaIVA = new Infragistics.Win.Misc.UltraButton();
			this.ultraLabel2 = new Infragistics.Win.Misc.UltraLabel();
			this.ultraLabel3 = new Infragistics.Win.Misc.UltraLabel();
			this.txtIvaDespues = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.ultraLabel1 = new Infragistics.Win.Misc.UltraLabel();
			this.txtIvaAntes = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.ultraTabPageControl2 = new Infragistics.Win.UltraWinTabControl.UltraTabPageControl();
			this.btKardex = new Infragistics.Win.Misc.UltraButton();
			this.btGrabar = new Infragistics.Win.Misc.UltraButton();
			this.btPrecio = new Infragistics.Win.Misc.UltraButton();
			this.btCosto = new Infragistics.Win.Misc.UltraButton();
			this.ultraGrid1 = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.cdsArticulo = new C1.Data.C1DataSet();
			this.ultraTabPageControl3 = new Infragistics.Win.UltraWinTabControl.UltraTabPageControl();
			this.btFactura = new Infragistics.Win.Misc.UltraButton();
			this.ugKardex = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.udsKardex = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.errorProvider = new System.Windows.Forms.ErrorProvider();
			this.ultraTabControl1 = new Infragistics.Win.UltraWinTabControl.UltraTabControl();
			this.ultraTabSharedControlsPage1 = new Infragistics.Win.UltraWinTabControl.UltraTabSharedControlsPage();
			this.btCliente = new Infragistics.Win.Misc.UltraButton();
			this.ultraTabPageControl1.SuspendLayout();
			this.panel4.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.txtPrecioF)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbPrecio1F)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsPrecio)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtPorcIncr)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbSigno)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbPrecio1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbPrecio2)).BeginInit();
			this.panel3.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.cmbMarca)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsArticuloTabla)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbGrupo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvGrupo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbColor)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbSubGrupo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbArancel)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbModelo)).BeginInit();
			this.panel5.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.txtDescCredito)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDescuento)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIvaDespues)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIvaAntes)).BeginInit();
			this.ultraTabPageControl2.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsArticulo)).BeginInit();
			this.ultraTabPageControl3.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.ugKardex)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.udsKardex)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraTabControl1)).BeginInit();
			this.ultraTabControl1.SuspendLayout();
			this.SuspendLayout();
			// 
			// ultraTabPageControl1
			// 
			this.ultraTabPageControl1.Controls.Add(this.panel4);
			this.ultraTabPageControl1.Controls.Add(this.panel3);
			this.ultraTabPageControl1.Controls.Add(this.panel5);
			this.ultraTabPageControl1.Location = new System.Drawing.Point(1, 23);
			this.ultraTabPageControl1.Name = "ultraTabPageControl1";
			this.ultraTabPageControl1.Size = new System.Drawing.Size(596, 369);
			// 
			// panel4
			// 
			this.panel4.BackColor = System.Drawing.Color.Transparent;
			this.panel4.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
			this.panel4.Controls.Add(this.btCliente);
			this.panel4.Controls.Add(this.ultraLabel7);
			this.panel4.Controls.Add(this.btCambiaPrecioValor);
			this.panel4.Controls.Add(this.txtPrecioF);
			this.panel4.Controls.Add(this.ultraLabel6);
			this.panel4.Controls.Add(this.cmbPrecio1F);
			this.panel4.Controls.Add(this.txtPorcIncr);
			this.panel4.Controls.Add(this.cmbSigno);
			this.panel4.Controls.Add(this.ultraLabel4);
			this.panel4.Controls.Add(this.ultraLabel5);
			this.panel4.Controls.Add(this.cmbPrecio1);
			this.panel4.Controls.Add(this.btIncPorc);
			this.panel4.Controls.Add(this.cmbPrecio2);
			this.panel4.Location = new System.Drawing.Point(24, 267);
			this.panel4.Name = "panel4";
			this.panel4.Size = new System.Drawing.Size(544, 86);
			this.panel4.TabIndex = 28;
			// 
			// ultraLabel7
			// 
			this.ultraLabel7.AutoSize = true;
			this.ultraLabel7.BackColor = System.Drawing.Color.Transparent;
			this.ultraLabel7.Location = new System.Drawing.Point(224, 43);
			this.ultraLabel7.Name = "ultraLabel7";
			this.ultraLabel7.Size = new System.Drawing.Size(10, 15);
			this.ultraLabel7.TabIndex = 31;
			this.ultraLabel7.Text = "$";
			// 
			// btCambiaPrecioValor
			// 
			this.btCambiaPrecioValor.Location = new System.Drawing.Point(240, 43);
			this.btCambiaPrecioValor.Name = "btCambiaPrecioValor";
			this.btCambiaPrecioValor.Size = new System.Drawing.Size(120, 25);
			this.btCambiaPrecioValor.TabIndex = 30;
			this.btCambiaPrecioValor.Text = "Cambiar Precio $";
			this.btCambiaPrecioValor.Click += new System.EventHandler(this.btCambiaPrecioValor_Click);
			// 
			// txtPrecioF
			// 
			appearance1.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtPrecioF.Appearance = appearance1;
			this.txtPrecioF.FormatString = "0.00";
			this.txtPrecioF.Location = new System.Drawing.Point(152, 43);
			this.txtPrecioF.MaskInput = "nn.nn";
			this.txtPrecioF.Name = "txtPrecioF";
			this.txtPrecioF.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtPrecioF.PromptChar = ' ';
			this.txtPrecioF.Size = new System.Drawing.Size(64, 22);
			this.txtPrecioF.SpinButtonDisplayStyle = Infragistics.Win.ButtonDisplayStyle.Always;
			this.txtPrecioF.TabIndex = 29;
			this.txtPrecioF.Value = 10;
			// 
			// ultraLabel6
			// 
			appearance2.ForeColorDisabled = System.Drawing.Color.Black;
			this.ultraLabel6.Appearance = appearance2;
			this.ultraLabel6.AutoSize = true;
			this.ultraLabel6.BackColor = System.Drawing.Color.Transparent;
			this.ultraLabel6.Location = new System.Drawing.Point(136, 43);
			this.ultraLabel6.Name = "ultraLabel6";
			this.ultraLabel6.Size = new System.Drawing.Size(12, 15);
			this.ultraLabel6.TabIndex = 28;
			this.ultraLabel6.Text = "=";
			// 
			// cmbPrecio1F
			// 
			appearance3.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbPrecio1F.Appearance = appearance3;
			this.cmbPrecio1F.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbPrecio1F.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbPrecio1F.DataMember = "ClientePrecio";
			this.cmbPrecio1F.DataSource = this.cdsPrecio;
			ultraGridColumn1.Header.VisiblePosition = 1;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn2.Header.VisiblePosition = 0;
			ultraGridBand1.Columns.AddRange(new object[] {
																										 ultraGridColumn1,
																										 ultraGridColumn2});
			this.cmbPrecio1F.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.cmbPrecio1F.DisplayMember = "Precio";
			this.cmbPrecio1F.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbPrecio1F.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbPrecio1F.Location = new System.Drawing.Point(16, 43);
			this.cmbPrecio1F.Name = "cmbPrecio1F";
			this.cmbPrecio1F.Size = new System.Drawing.Size(112, 21);
			this.cmbPrecio1F.TabIndex = 27;
			this.cmbPrecio1F.ValueMember = "idTipoPrecio";
			// 
			// cdsPrecio
			// 
			this.cdsPrecio.BindingContextCtrl = this;
			this.cdsPrecio.DataLibrary = "LibFacturacion";
			this.cdsPrecio.DataLibraryUrl = "";
			this.cdsPrecio.DataSetDef = "dsClienteTabla";
			this.cdsPrecio.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsPrecio.Name = "cdsPrecio";
			this.cdsPrecio.SchemaClassName = "LibFacturacion.DataClass";
			this.cdsPrecio.SchemaDef = null;
			this.cdsPrecio.BeforeFill += new C1.Data.FillEventHandler(this.cdsPrecio_BeforeFill);
			// 
			// txtPorcIncr
			// 
			appearance4.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtPorcIncr.Appearance = appearance4;
			this.txtPorcIncr.FormatString = "0.00";
			this.txtPorcIncr.Location = new System.Drawing.Point(320, 9);
			this.txtPorcIncr.MaskInput = "nn.nn";
			this.txtPorcIncr.Name = "txtPorcIncr";
			this.txtPorcIncr.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtPorcIncr.PromptChar = ' ';
			this.txtPorcIncr.Size = new System.Drawing.Size(64, 22);
			this.txtPorcIncr.SpinButtonDisplayStyle = Infragistics.Win.ButtonDisplayStyle.Always;
			this.txtPorcIncr.TabIndex = 24;
			this.txtPorcIncr.Value = 10;
			// 
			// cmbSigno
			// 
			appearance5.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbSigno.Appearance = appearance5;
			this.cmbSigno.DropDownStyle = Infragistics.Win.DropDownStyle.DropDownList;
			valueListItem1.DataValue = "+";
			valueListItem1.DisplayText = "+";
			valueListItem2.DataValue = "-";
			valueListItem2.DisplayText = "-";
			this.cmbSigno.Items.Add(valueListItem1);
			this.cmbSigno.Items.Add(valueListItem2);
			this.cmbSigno.Location = new System.Drawing.Point(272, 9);
			this.cmbSigno.Name = "cmbSigno";
			this.cmbSigno.Size = new System.Drawing.Size(40, 22);
			this.cmbSigno.TabIndex = 23;
			this.cmbSigno.Text = "+";
			// 
			// ultraLabel4
			// 
			appearance6.ForeColorDisabled = System.Drawing.Color.Black;
			this.ultraLabel4.Appearance = appearance6;
			this.ultraLabel4.AutoSize = true;
			this.ultraLabel4.BackColor = System.Drawing.Color.Transparent;
			this.ultraLabel4.Location = new System.Drawing.Point(136, 9);
			this.ultraLabel4.Name = "ultraLabel4";
			this.ultraLabel4.Size = new System.Drawing.Size(12, 15);
			this.ultraLabel4.TabIndex = 22;
			this.ultraLabel4.Text = "=";
			// 
			// ultraLabel5
			// 
			appearance7.ForeColorDisabled = System.Drawing.Color.Black;
			this.ultraLabel5.Appearance = appearance7;
			this.ultraLabel5.AutoSize = true;
			this.ultraLabel5.BackColor = System.Drawing.Color.Transparent;
			this.ultraLabel5.Location = new System.Drawing.Point(392, 17);
			this.ultraLabel5.Name = "ultraLabel5";
			this.ultraLabel5.Size = new System.Drawing.Size(15, 15);
			this.ultraLabel5.TabIndex = 25;
			this.ultraLabel5.Text = "%";
			// 
			// cmbPrecio1
			// 
			appearance8.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbPrecio1.Appearance = appearance8;
			this.cmbPrecio1.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbPrecio1.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbPrecio1.DataMember = "ClientePrecio";
			this.cmbPrecio1.DataSource = this.cdsPrecio;
			ultraGridColumn3.Header.VisiblePosition = 1;
			ultraGridColumn3.Hidden = true;
			ultraGridColumn4.Header.VisiblePosition = 0;
			ultraGridBand2.Columns.AddRange(new object[] {
																										 ultraGridColumn3,
																										 ultraGridColumn4});
			this.cmbPrecio1.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			this.cmbPrecio1.DisplayMember = "Precio";
			this.cmbPrecio1.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbPrecio1.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbPrecio1.Location = new System.Drawing.Point(16, 9);
			this.cmbPrecio1.Name = "cmbPrecio1";
			this.cmbPrecio1.Size = new System.Drawing.Size(112, 21);
			this.cmbPrecio1.TabIndex = 2;
			this.cmbPrecio1.ValueMember = "idTipoPrecio";
			// 
			// btIncPorc
			// 
			this.btIncPorc.Location = new System.Drawing.Point(416, 9);
			this.btIncPorc.Name = "btIncPorc";
			this.btIncPorc.Size = new System.Drawing.Size(120, 24);
			this.btIncPorc.TabIndex = 26;
			this.btIncPorc.Text = "Cambiar Precio %";
			this.btIncPorc.Click += new System.EventHandler(this.btIncPorc_Click);
			// 
			// cmbPrecio2
			// 
			appearance9.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbPrecio2.Appearance = appearance9;
			this.cmbPrecio2.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbPrecio2.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbPrecio2.DataMember = "ClientePrecio";
			this.cmbPrecio2.DataSource = this.cdsPrecio;
			ultraGridColumn5.Header.VisiblePosition = 1;
			ultraGridColumn5.Hidden = true;
			ultraGridColumn6.Header.VisiblePosition = 0;
			ultraGridBand3.Columns.AddRange(new object[] {
																										 ultraGridColumn5,
																										 ultraGridColumn6});
			this.cmbPrecio2.DisplayLayout.BandsSerializer.Add(ultraGridBand3);
			this.cmbPrecio2.DisplayMember = "Precio";
			this.cmbPrecio2.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbPrecio2.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbPrecio2.Location = new System.Drawing.Point(152, 9);
			this.cmbPrecio2.Name = "cmbPrecio2";
			this.cmbPrecio2.Size = new System.Drawing.Size(112, 21);
			this.cmbPrecio2.TabIndex = 21;
			this.cmbPrecio2.ValueMember = "idTipoPrecio";
			// 
			// panel3
			// 
			this.panel3.BackColor = System.Drawing.Color.Transparent;
			this.panel3.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
			this.panel3.Controls.Add(this.cmbMarca);
			this.panel3.Controls.Add(this.cmbGrupo);
			this.panel3.Controls.Add(this.cmbColor);
			this.panel3.Controls.Add(this.chkArancel);
			this.panel3.Controls.Add(this.chkMarca);
			this.panel3.Controls.Add(this.cmbSubGrupo);
			this.panel3.Controls.Add(this.cmbArancel);
			this.panel3.Controls.Add(this.chkGrupo);
			this.panel3.Controls.Add(this.chkColor);
			this.panel3.Controls.Add(this.cmbModelo);
			this.panel3.Controls.Add(this.chkModelo);
			this.panel3.Controls.Add(this.chkSubGrupo);
			this.panel3.Controls.Add(this.btCancelar);
			this.panel3.Controls.Add(this.btAceptar);
			this.panel3.Location = new System.Drawing.Point(24, 9);
			this.panel3.Name = "panel3";
			this.panel3.Size = new System.Drawing.Size(544, 155);
			this.panel3.TabIndex = 28;
			// 
			// cmbMarca
			// 
			appearance10.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbMarca.Appearance = appearance10;
			this.cmbMarca.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbMarca.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbMarca.DataMember = "ArticuloMarca";
			this.cmbMarca.DataSource = this.cdsArticuloTabla;
			ultraGridColumn7.Header.VisiblePosition = 2;
			ultraGridColumn7.Hidden = true;
			ultraGridColumn8.Header.VisiblePosition = 0;
			ultraGridColumn8.Width = 235;
			ultraGridColumn9.Header.VisiblePosition = 1;
			ultraGridColumn9.Hidden = true;
			ultraGridBand4.Columns.AddRange(new object[] {
																										 ultraGridColumn7,
																										 ultraGridColumn8,
																										 ultraGridColumn9});
			this.cmbMarca.DisplayLayout.BandsSerializer.Add(ultraGridBand4);
			this.cmbMarca.DisplayMember = "Marca";
			this.cmbMarca.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbMarca.Enabled = false;
			this.cmbMarca.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbMarca.Location = new System.Drawing.Point(138, 43);
			this.cmbMarca.Name = "cmbMarca";
			this.cmbMarca.Size = new System.Drawing.Size(112, 22);
			this.cmbMarca.TabIndex = 22;
			this.cmbMarca.ValueMember = "idMarca";
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
			this.cdsArticuloTabla.BeforeFill += new C1.Data.FillEventHandler(this.cdsArticuloTabla_BeforeFill);
			// 
			// cmbGrupo
			// 
			appearance11.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbGrupo.Appearance = appearance11;
			this.cmbGrupo.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbGrupo.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbGrupo.DataSource = this.cdvGrupo;
			ultraGridColumn10.Header.VisiblePosition = 6;
			ultraGridColumn10.Hidden = true;
			ultraGridColumn11.Header.VisiblePosition = 7;
			ultraGridColumn11.Hidden = true;
			ultraGridColumn12.Header.VisiblePosition = 0;
			ultraGridColumn12.Width = 345;
			ultraGridColumn13.Header.VisiblePosition = 3;
			ultraGridColumn13.Hidden = true;
			ultraGridColumn14.Header.VisiblePosition = 4;
			ultraGridColumn14.Hidden = true;
			ultraGridColumn15.Header.VisiblePosition = 5;
			ultraGridColumn15.Hidden = true;
			ultraGridColumn16.Header.VisiblePosition = 1;
			ultraGridColumn16.Hidden = true;
			ultraGridColumn17.Header.VisiblePosition = 8;
			ultraGridColumn17.Hidden = true;
			ultraGridColumn18.Header.VisiblePosition = 9;
			ultraGridColumn18.Hidden = true;
			ultraGridColumn19.Header.VisiblePosition = 10;
			ultraGridColumn19.Hidden = true;
			ultraGridColumn20.Header.VisiblePosition = 2;
			ultraGridColumn20.Hidden = true;
			ultraGridColumn21.Header.VisiblePosition = 11;
			ultraGridColumn21.Hidden = true;
			ultraGridColumn22.Header.VisiblePosition = 12;
			ultraGridColumn23.Header.VisiblePosition = 13;
			ultraGridColumn24.Header.VisiblePosition = 14;
			ultraGridColumn25.Header.VisiblePosition = 15;
			ultraGridColumn26.Header.VisiblePosition = 16;
			ultraGridColumn27.Header.VisiblePosition = 17;
			ultraGridColumn28.Header.VisiblePosition = 18;
			ultraGridColumn29.Header.VisiblePosition = 19;
			ultraGridBand5.Columns.AddRange(new object[] {
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
																										 ultraGridColumn29});
			this.cmbGrupo.DisplayLayout.BandsSerializer.Add(ultraGridBand5);
			this.cmbGrupo.DisplayMember = "Grupo";
			this.cmbGrupo.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbGrupo.Enabled = false;
			this.cmbGrupo.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbGrupo.Location = new System.Drawing.Point(138, 9);
			this.cmbGrupo.Name = "cmbGrupo";
			this.cmbGrupo.Size = new System.Drawing.Size(112, 22);
			this.cmbGrupo.TabIndex = 24;
			this.cmbGrupo.ValueMember = "idGrupoArticulo";
			// 
			// cdvGrupo
			// 
			this.cdvGrupo.BindingContextCtrl = this;
			this.cdvGrupo.DataSet = this.cdsArticuloTabla;
			this.cdvGrupo.Sort = "Grupo";
			this.cdvGrupo.TableName = "";
			this.cdvGrupo.TableViewName = "ArticuloGrupo";
			// 
			// cmbColor
			// 
			appearance12.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbColor.Appearance = appearance12;
			this.cmbColor.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbColor.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbColor.DataMember = "ArticuloColor";
			this.cmbColor.DataSource = this.cdsArticuloTabla;
			ultraGridColumn30.Header.VisiblePosition = 1;
			ultraGridColumn30.Hidden = true;
			ultraGridColumn31.Header.VisiblePosition = 0;
			ultraGridColumn31.Width = 221;
			ultraGridBand6.Columns.AddRange(new object[] {
																										 ultraGridColumn30,
																										 ultraGridColumn31});
			this.cmbColor.DisplayLayout.BandsSerializer.Add(ultraGridBand6);
			this.cmbColor.DisplayMember = "Color";
			this.cmbColor.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbColor.Enabled = false;
			this.cmbColor.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbColor.Location = new System.Drawing.Point(138, 78);
			this.cmbColor.Name = "cmbColor";
			this.cmbColor.Size = new System.Drawing.Size(112, 22);
			this.cmbColor.TabIndex = 25;
			this.cmbColor.ValueMember = "idColor";
			// 
			// chkArancel
			// 
			appearance13.ForeColorDisabled = System.Drawing.Color.Black;
			this.chkArancel.Appearance = appearance13;
			this.chkArancel.BackColor = System.Drawing.Color.Transparent;
			this.chkArancel.Checked = true;
			this.chkArancel.CheckState = System.Windows.Forms.CheckState.Checked;
			this.chkArancel.Location = new System.Drawing.Point(290, 78);
			this.chkArancel.Name = "chkArancel";
			this.chkArancel.Size = new System.Drawing.Size(120, 21);
			this.chkArancel.TabIndex = 20;
			this.chkArancel.Text = "Toda Arancel";
			this.chkArancel.CheckedChanged += new System.EventHandler(this.chkArancel_CheckedChanged);
			// 
			// chkMarca
			// 
			appearance14.ForeColorDisabled = System.Drawing.Color.Black;
			this.chkMarca.Appearance = appearance14;
			this.chkMarca.BackColor = System.Drawing.Color.Transparent;
			this.chkMarca.Checked = true;
			this.chkMarca.CheckState = System.Windows.Forms.CheckState.Checked;
			this.chkMarca.Location = new System.Drawing.Point(10, 43);
			this.chkMarca.Name = "chkMarca";
			this.chkMarca.Size = new System.Drawing.Size(120, 22);
			this.chkMarca.TabIndex = 17;
			this.chkMarca.Text = "Toda Marca";
			this.chkMarca.CheckedChanged += new System.EventHandler(this.chkMarca_CheckedChanged);
			// 
			// cmbSubGrupo
			// 
			appearance15.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbSubGrupo.Appearance = appearance15;
			this.cmbSubGrupo.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbSubGrupo.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbSubGrupo.DataMember = "ArticuloSubGrupo";
			this.cmbSubGrupo.DataSource = this.cdsArticuloTabla;
			ultraGridColumn32.Header.VisiblePosition = 2;
			ultraGridColumn32.Hidden = true;
			ultraGridColumn33.Header.VisiblePosition = 0;
			ultraGridColumn33.Width = 269;
			ultraGridColumn34.Header.VisiblePosition = 1;
			ultraGridColumn34.Hidden = true;
			ultraGridBand7.Columns.AddRange(new object[] {
																										 ultraGridColumn32,
																										 ultraGridColumn33,
																										 ultraGridColumn34});
			this.cmbSubGrupo.DisplayLayout.BandsSerializer.Add(ultraGridBand7);
			this.cmbSubGrupo.DisplayMember = "SubGrupo";
			this.cmbSubGrupo.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbSubGrupo.Enabled = false;
			this.cmbSubGrupo.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbSubGrupo.Location = new System.Drawing.Point(418, 9);
			this.cmbSubGrupo.Name = "cmbSubGrupo";
			this.cmbSubGrupo.Size = new System.Drawing.Size(112, 22);
			this.cmbSubGrupo.TabIndex = 23;
			this.cmbSubGrupo.ValueMember = "idSubGrupo";
			// 
			// cmbArancel
			// 
			appearance16.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbArancel.Appearance = appearance16;
			this.cmbArancel.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbArancel.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbArancel.DataMember = "ArticuloArancel";
			this.cmbArancel.DataSource = this.cdsArticuloTabla;
			ultraGridColumn35.Header.VisiblePosition = 2;
			ultraGridColumn35.Hidden = true;
			ultraGridColumn36.Header.VisiblePosition = 1;
			ultraGridColumn36.Hidden = true;
			ultraGridColumn37.Header.VisiblePosition = 3;
			ultraGridColumn37.Hidden = true;
			ultraGridColumn38.Header.VisiblePosition = 0;
			ultraGridColumn38.Width = 212;
			ultraGridBand8.Columns.AddRange(new object[] {
																										 ultraGridColumn35,
																										 ultraGridColumn36,
																										 ultraGridColumn37,
																										 ultraGridColumn38});
			this.cmbArancel.DisplayLayout.BandsSerializer.Add(ultraGridBand8);
			this.cmbArancel.DisplayMember = "Arancel";
			this.cmbArancel.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbArancel.Enabled = false;
			this.cmbArancel.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbArancel.Location = new System.Drawing.Point(418, 78);
			this.cmbArancel.Name = "cmbArancel";
			this.cmbArancel.Size = new System.Drawing.Size(112, 22);
			this.cmbArancel.TabIndex = 26;
			this.cmbArancel.ValueMember = "idArancel";
			// 
			// chkGrupo
			// 
			appearance17.ForeColorDisabled = System.Drawing.Color.Black;
			this.chkGrupo.Appearance = appearance17;
			this.chkGrupo.BackColor = System.Drawing.Color.Transparent;
			this.chkGrupo.Checked = true;
			this.chkGrupo.CheckState = System.Windows.Forms.CheckState.Checked;
			this.chkGrupo.Location = new System.Drawing.Point(10, 9);
			this.chkGrupo.Name = "chkGrupo";
			this.chkGrupo.Size = new System.Drawing.Size(120, 21);
			this.chkGrupo.TabIndex = 15;
			this.chkGrupo.Text = "Todo Grupo";
			this.chkGrupo.CheckedChanged += new System.EventHandler(this.chkGrupo_CheckedChanged);
			// 
			// chkColor
			// 
			appearance18.ForeColorDisabled = System.Drawing.Color.Black;
			this.chkColor.Appearance = appearance18;
			this.chkColor.BackColor = System.Drawing.Color.Transparent;
			this.chkColor.Checked = true;
			this.chkColor.CheckState = System.Windows.Forms.CheckState.Checked;
			this.chkColor.Location = new System.Drawing.Point(10, 78);
			this.chkColor.Name = "chkColor";
			this.chkColor.Size = new System.Drawing.Size(120, 21);
			this.chkColor.TabIndex = 19;
			this.chkColor.Text = "Todo Color";
			this.chkColor.CheckedChanged += new System.EventHandler(this.chkColor_CheckedChanged);
			// 
			// cmbModelo
			// 
			appearance19.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbModelo.Appearance = appearance19;
			this.cmbModelo.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbModelo.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbModelo.DataMember = "ArticuloModelo";
			this.cmbModelo.DataSource = this.cdsArticuloTabla;
			ultraGridColumn39.Header.VisiblePosition = 1;
			ultraGridColumn39.Hidden = true;
			ultraGridColumn40.Header.VisiblePosition = 0;
			ultraGridColumn40.Width = 230;
			ultraGridBand9.Columns.AddRange(new object[] {
																										 ultraGridColumn39,
																										 ultraGridColumn40});
			this.cmbModelo.DisplayLayout.BandsSerializer.Add(ultraGridBand9);
			this.cmbModelo.DisplayMember = "Modelo";
			this.cmbModelo.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbModelo.Enabled = false;
			this.cmbModelo.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbModelo.Location = new System.Drawing.Point(418, 43);
			this.cmbModelo.Name = "cmbModelo";
			this.cmbModelo.Size = new System.Drawing.Size(112, 22);
			this.cmbModelo.TabIndex = 21;
			this.cmbModelo.ValueMember = "idModelo";
			// 
			// chkModelo
			// 
			appearance20.ForeColorDisabled = System.Drawing.Color.Black;
			this.chkModelo.Appearance = appearance20;
			this.chkModelo.BackColor = System.Drawing.Color.Transparent;
			this.chkModelo.Checked = true;
			this.chkModelo.CheckState = System.Windows.Forms.CheckState.Checked;
			this.chkModelo.Location = new System.Drawing.Point(290, 43);
			this.chkModelo.Name = "chkModelo";
			this.chkModelo.Size = new System.Drawing.Size(120, 22);
			this.chkModelo.TabIndex = 18;
			this.chkModelo.Text = "Toda Modelo";
			this.chkModelo.CheckedChanged += new System.EventHandler(this.chkModelo_CheckedChanged);
			// 
			// chkSubGrupo
			// 
			appearance21.ForeColorDisabled = System.Drawing.Color.Black;
			this.chkSubGrupo.Appearance = appearance21;
			this.chkSubGrupo.BackColor = System.Drawing.Color.Transparent;
			this.chkSubGrupo.Checked = true;
			this.chkSubGrupo.CheckState = System.Windows.Forms.CheckState.Checked;
			this.chkSubGrupo.Location = new System.Drawing.Point(290, 9);
			this.chkSubGrupo.Name = "chkSubGrupo";
			this.chkSubGrupo.Size = new System.Drawing.Size(120, 21);
			this.chkSubGrupo.TabIndex = 16;
			this.chkSubGrupo.Text = "Todo SubGrupo";
			this.chkSubGrupo.CheckedChanged += new System.EventHandler(this.chkSubGrupo_CheckedChanged);
			// 
			// btCancelar
			// 
			this.btCancelar.Location = new System.Drawing.Point(352, 121);
			this.btCancelar.Name = "btCancelar";
			this.btCancelar.Size = new System.Drawing.Size(75, 24);
			this.btCancelar.TabIndex = 1;
			this.btCancelar.Text = "&Salir";
			this.btCancelar.Click += new System.EventHandler(this.btCancelar_Click);
			// 
			// btAceptar
			// 
			this.btAceptar.Location = new System.Drawing.Point(72, 121);
			this.btAceptar.Name = "btAceptar";
			this.btAceptar.Size = new System.Drawing.Size(168, 24);
			this.btAceptar.TabIndex = 0;
			this.btAceptar.Text = "&Filtrar  Artículos ->";
			this.btAceptar.Click += new System.EventHandler(this.btAceptar_Click);
			// 
			// panel5
			// 
			this.panel5.BackColor = System.Drawing.Color.Transparent;
			this.panel5.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
			this.panel5.Controls.Add(this.ultraLabel10);
			this.panel5.Controls.Add(this.txtDescCredito);
			this.panel5.Controls.Add(this.btDescuento);
			this.panel5.Controls.Add(this.ultraLabel8);
			this.panel5.Controls.Add(this.txtDescuento);
			this.panel5.Controls.Add(this.btCambiaIVA);
			this.panel5.Controls.Add(this.ultraLabel2);
			this.panel5.Controls.Add(this.ultraLabel3);
			this.panel5.Controls.Add(this.txtIvaDespues);
			this.panel5.Controls.Add(this.ultraLabel1);
			this.panel5.Controls.Add(this.txtIvaAntes);
			this.panel5.Location = new System.Drawing.Point(24, 172);
			this.panel5.Name = "panel5";
			this.panel5.Size = new System.Drawing.Size(544, 78);
			this.panel5.TabIndex = 28;
			// 
			// ultraLabel10
			// 
			this.ultraLabel10.AutoSize = true;
			this.ultraLabel10.BackColor = System.Drawing.Color.Transparent;
			this.ultraLabel10.Location = new System.Drawing.Point(192, 43);
			this.ultraLabel10.Name = "ultraLabel10";
			this.ultraLabel10.Size = new System.Drawing.Size(83, 15);
			this.ultraLabel10.TabIndex = 26;
			this.ultraLabel10.Text = "%     Crédito %";
			// 
			// txtDescCredito
			// 
			appearance22.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtDescCredito.Appearance = appearance22;
			this.txtDescCredito.FormatString = "0.00";
			this.txtDescCredito.Location = new System.Drawing.Point(272, 43);
			this.txtDescCredito.MaskInput = "nn.nn";
			this.txtDescCredito.Name = "txtDescCredito";
			this.txtDescCredito.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtDescCredito.PromptChar = ' ';
			this.txtDescCredito.Size = new System.Drawing.Size(64, 22);
			this.txtDescCredito.SpinButtonDisplayStyle = Infragistics.Win.ButtonDisplayStyle.Always;
			this.txtDescCredito.TabIndex = 25;
			this.txtDescCredito.Value = 5;
			// 
			// btDescuento
			// 
			this.btDescuento.Location = new System.Drawing.Point(376, 43);
			this.btDescuento.Name = "btDescuento";
			this.btDescuento.Size = new System.Drawing.Size(128, 25);
			this.btDescuento.TabIndex = 23;
			this.btDescuento.Text = "Cambiar Descuento";
			this.btDescuento.Click += new System.EventHandler(this.btDescuento_Click);
			// 
			// ultraLabel8
			// 
			this.ultraLabel8.AutoSize = true;
			this.ultraLabel8.BackColor = System.Drawing.Color.Transparent;
			this.ultraLabel8.Location = new System.Drawing.Point(16, 43);
			this.ultraLabel8.Name = "ultraLabel8";
			this.ultraLabel8.Size = new System.Drawing.Size(102, 15);
			this.ultraLabel8.TabIndex = 22;
			this.ultraLabel8.Text = "Descuento Contado";
			// 
			// txtDescuento
			// 
			appearance23.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtDescuento.Appearance = appearance23;
			this.txtDescuento.FormatString = "0.00";
			this.txtDescuento.Location = new System.Drawing.Point(128, 43);
			this.txtDescuento.MaskInput = "nn.nn";
			this.txtDescuento.Name = "txtDescuento";
			this.txtDescuento.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtDescuento.PromptChar = ' ';
			this.txtDescuento.Size = new System.Drawing.Size(64, 22);
			this.txtDescuento.SpinButtonDisplayStyle = Infragistics.Win.ButtonDisplayStyle.Always;
			this.txtDescuento.TabIndex = 21;
			this.txtDescuento.Value = 5;
			// 
			// btCambiaIVA
			// 
			this.btCambiaIVA.Location = new System.Drawing.Point(400, 9);
			this.btCambiaIVA.Name = "btCambiaIVA";
			this.btCambiaIVA.Size = new System.Drawing.Size(88, 24);
			this.btCambiaIVA.TabIndex = 15;
			this.btCambiaIVA.Text = "Cambiar IVA";
			this.btCambiaIVA.Click += new System.EventHandler(this.btCambiaIVA_Click);
			// 
			// ultraLabel2
			// 
			this.ultraLabel2.AutoSize = true;
			this.ultraLabel2.BackColor = System.Drawing.Color.Transparent;
			this.ultraLabel2.Location = new System.Drawing.Point(176, 9);
			this.ultraLabel2.Name = "ultraLabel2";
			this.ultraLabel2.Size = new System.Drawing.Size(32, 15);
			this.ultraLabel2.TabIndex = 19;
			this.ultraLabel2.Text = "%   a";
			// 
			// ultraLabel3
			// 
			this.ultraLabel3.AutoSize = true;
			this.ultraLabel3.BackColor = System.Drawing.Color.Transparent;
			this.ultraLabel3.Location = new System.Drawing.Point(296, 9);
			this.ultraLabel3.Name = "ultraLabel3";
			this.ultraLabel3.Size = new System.Drawing.Size(15, 15);
			this.ultraLabel3.TabIndex = 20;
			this.ultraLabel3.Text = "%";
			// 
			// txtIvaDespues
			// 
			appearance24.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtIvaDespues.Appearance = appearance24;
			this.txtIvaDespues.FormatString = "0.00";
			this.txtIvaDespues.Location = new System.Drawing.Point(216, 9);
			this.txtIvaDespues.MaskInput = "nn.nn";
			this.txtIvaDespues.Name = "txtIvaDespues";
			this.txtIvaDespues.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtIvaDespues.PromptChar = ' ';
			this.txtIvaDespues.Size = new System.Drawing.Size(64, 22);
			this.txtIvaDespues.SpinButtonDisplayStyle = Infragistics.Win.ButtonDisplayStyle.Always;
			this.txtIvaDespues.TabIndex = 18;
			this.txtIvaDespues.Value = 14;
			// 
			// ultraLabel1
			// 
			this.ultraLabel1.AutoSize = true;
			this.ultraLabel1.BackColor = System.Drawing.Color.Transparent;
			this.ultraLabel1.Location = new System.Drawing.Point(16, 9);
			this.ultraLabel1.Name = "ultraLabel1";
			this.ultraLabel1.Size = new System.Drawing.Size(82, 15);
			this.ultraLabel1.TabIndex = 17;
			this.ultraLabel1.Text = "Cambiar IVA de";
			// 
			// txtIvaAntes
			// 
			appearance25.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtIvaAntes.Appearance = appearance25;
			this.txtIvaAntes.FormatString = "0.00";
			this.txtIvaAntes.Location = new System.Drawing.Point(104, 9);
			this.txtIvaAntes.MaskInput = "nn.nn";
			this.txtIvaAntes.Name = "txtIvaAntes";
			this.txtIvaAntes.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtIvaAntes.PromptChar = ' ';
			this.txtIvaAntes.Size = new System.Drawing.Size(64, 22);
			this.txtIvaAntes.SpinButtonDisplayStyle = Infragistics.Win.ButtonDisplayStyle.Always;
			this.txtIvaAntes.TabIndex = 16;
			this.txtIvaAntes.Value = 12;
			// 
			// ultraTabPageControl2
			// 
			this.ultraTabPageControl2.Controls.Add(this.btKardex);
			this.ultraTabPageControl2.Controls.Add(this.btGrabar);
			this.ultraTabPageControl2.Controls.Add(this.btPrecio);
			this.ultraTabPageControl2.Controls.Add(this.btCosto);
			this.ultraTabPageControl2.Controls.Add(this.ultraGrid1);
			this.ultraTabPageControl2.Location = new System.Drawing.Point(-10000, -10769);
			this.ultraTabPageControl2.Name = "ultraTabPageControl2";
			this.ultraTabPageControl2.Size = new System.Drawing.Size(596, 369);
			// 
			// btKardex
			// 
			this.btKardex.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btKardex.Location = new System.Drawing.Point(208, 336);
			this.btKardex.Name = "btKardex";
			this.btKardex.Size = new System.Drawing.Size(64, 26);
			this.btKardex.TabIndex = 19;
			this.btKardex.Text = "Kardex ->";
			this.btKardex.Click += new System.EventHandler(this.btKardex_Click);
			// 
			// btGrabar
			// 
			this.btGrabar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btGrabar.Location = new System.Drawing.Point(144, 336);
			this.btGrabar.Name = "btGrabar";
			this.btGrabar.Size = new System.Drawing.Size(64, 26);
			this.btGrabar.TabIndex = 18;
			this.btGrabar.Text = "Grabar";
			this.btGrabar.Click += new System.EventHandler(this.btGrabar_Click);
			// 
			// btPrecio
			// 
			this.btPrecio.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btPrecio.Location = new System.Drawing.Point(80, 336);
			this.btPrecio.Name = "btPrecio";
			this.btPrecio.Size = new System.Drawing.Size(64, 26);
			this.btPrecio.TabIndex = 17;
			this.btPrecio.Text = "Precio";
			this.btPrecio.Click += new System.EventHandler(this.btPrecio_Click);
			// 
			// btCosto
			// 
			this.btCosto.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btCosto.Location = new System.Drawing.Point(16, 336);
			this.btCosto.Name = "btCosto";
			this.btCosto.Size = new System.Drawing.Size(64, 26);
			this.btCosto.TabIndex = 16;
			this.btCosto.Text = "Costos";
			this.btCosto.Click += new System.EventHandler(this.btCosto_Click);
			// 
			// ultraGrid1
			// 
			this.ultraGrid1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.ultraGrid1.Cursor = System.Windows.Forms.Cursors.Default;
			this.ultraGrid1.DataMember = "Articulo";
			this.ultraGrid1.DataSource = this.cdsArticulo;
			appearance26.BackColor = System.Drawing.Color.White;
			this.ultraGrid1.DisplayLayout.Appearance = appearance26;
			ultraGridColumn41.Header.VisiblePosition = 45;
			ultraGridColumn41.Hidden = true;
			ultraGridColumn42.Header.VisiblePosition = 100;
			ultraGridColumn42.Hidden = true;
			ultraGridColumn43.Header.VisiblePosition = 93;
			ultraGridColumn43.Hidden = true;
			ultraGridColumn44.Header.VisiblePosition = 94;
			ultraGridColumn44.Hidden = true;
			ultraGridColumn45.Header.VisiblePosition = 47;
			ultraGridColumn45.Hidden = true;
			ultraGridColumn46.Header.VisiblePosition = 51;
			ultraGridColumn46.Hidden = true;
			ultraGridColumn47.Header.VisiblePosition = 44;
			ultraGridColumn47.Hidden = true;
			ultraGridColumn48.Header.VisiblePosition = 48;
			ultraGridColumn48.Hidden = true;
			ultraGridColumn49.Header.VisiblePosition = 53;
			ultraGridColumn49.Hidden = true;
			ultraGridColumn50.Header.VisiblePosition = 49;
			ultraGridColumn50.Hidden = true;
			ultraGridColumn51.Header.VisiblePosition = 52;
			ultraGridColumn51.Hidden = true;
			ultraGridColumn52.Header.VisiblePosition = 46;
			ultraGridColumn52.Hidden = true;
			ultraGridColumn53.Header.VisiblePosition = 50;
			ultraGridColumn53.Hidden = true;
			ultraGridColumn54.Header.VisiblePosition = 91;
			ultraGridColumn54.Hidden = true;
			ultraGridColumn55.Header.VisiblePosition = 99;
			ultraGridColumn55.Hidden = true;
			ultraGridColumn56.Header.VisiblePosition = 84;
			ultraGridColumn56.Hidden = true;
			ultraGridColumn57.Header.VisiblePosition = 25;
			ultraGridColumn57.Hidden = true;
			ultraGridColumn58.Header.VisiblePosition = 3;
			ultraGridColumn58.Width = 65;
			ultraGridColumn59.Header.VisiblePosition = 63;
			ultraGridColumn59.Hidden = true;
			ultraGridColumn60.Header.VisiblePosition = 24;
			ultraGridColumn60.Hidden = true;
			ultraGridColumn61.Header.VisiblePosition = 6;
			ultraGridColumn61.Hidden = true;
			ultraGridColumn62.Header.VisiblePosition = 60;
			ultraGridColumn62.Width = 37;
			ultraGridColumn63.Header.VisiblePosition = 43;
			ultraGridColumn63.Hidden = true;
			appearance27.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn64.CellAppearance = appearance27;
			ultraGridColumn64.Format = "#,##0.00";
			ultraGridColumn64.Header.VisiblePosition = 64;
			ultraGridColumn64.Width = 58;
			appearance28.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn65.CellAppearance = appearance28;
			ultraGridColumn65.Format = "#,##0.00";
			ultraGridColumn65.Header.VisiblePosition = 65;
			ultraGridColumn65.Width = 62;
			appearance29.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn66.CellAppearance = appearance29;
			ultraGridColumn66.Format = "#,##0.00";
			ultraGridColumn66.Header.VisiblePosition = 66;
			ultraGridColumn66.Width = 68;
			appearance30.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn67.CellAppearance = appearance30;
			ultraGridColumn67.Format = "#,##0.00";
			ultraGridColumn67.Header.VisiblePosition = 67;
			ultraGridColumn67.Width = 64;
			appearance31.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn68.CellAppearance = appearance31;
			ultraGridColumn68.Format = "#,##0.00";
			ultraGridColumn68.Header.VisiblePosition = 68;
			ultraGridColumn68.Width = 60;
			appearance32.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn69.CellAppearance = appearance32;
			ultraGridColumn69.Format = "#,##0.00";
			ultraGridColumn69.Header.VisiblePosition = 69;
			ultraGridColumn69.Width = 59;
			ultraGridColumn70.Header.VisiblePosition = 19;
			ultraGridColumn70.Hidden = true;
			ultraGridColumn71.Header.VisiblePosition = 18;
			ultraGridColumn71.Hidden = true;
			ultraGridColumn72.Header.VisiblePosition = 71;
			ultraGridColumn72.Hidden = true;
			ultraGridColumn73.Header.VisiblePosition = 34;
			ultraGridColumn73.Hidden = true;
			ultraGridColumn74.Header.VisiblePosition = 83;
			ultraGridColumn74.Hidden = true;
			ultraGridColumn75.Header.VisiblePosition = 80;
			ultraGridColumn75.Hidden = true;
			ultraGridColumn76.Header.VisiblePosition = 81;
			ultraGridColumn76.Hidden = true;
			ultraGridColumn77.Header.VisiblePosition = 82;
			ultraGridColumn77.Hidden = true;
			ultraGridColumn78.Header.VisiblePosition = 2;
			ultraGridColumn78.Width = 63;
			ultraGridColumn79.Header.VisiblePosition = 41;
			ultraGridColumn79.Hidden = true;
			ultraGridColumn80.Header.VisiblePosition = 33;
			ultraGridColumn80.Hidden = true;
			ultraGridColumn81.Header.VisiblePosition = 75;
			ultraGridColumn81.Hidden = true;
			ultraGridColumn82.Header.VisiblePosition = 78;
			ultraGridColumn82.Hidden = true;
			ultraGridColumn83.Header.VisiblePosition = 35;
			ultraGridColumn83.Hidden = true;
			ultraGridColumn84.Header.VisiblePosition = 54;
			ultraGridColumn84.Hidden = true;
			ultraGridColumn85.Header.VisiblePosition = 55;
			ultraGridColumn85.Hidden = true;
			ultraGridColumn86.Header.VisiblePosition = 56;
			ultraGridColumn86.Hidden = true;
			ultraGridColumn87.Header.VisiblePosition = 57;
			ultraGridColumn87.Hidden = true;
			ultraGridColumn88.Header.VisiblePosition = 58;
			ultraGridColumn88.Hidden = true;
			ultraGridColumn89.Header.VisiblePosition = 59;
			ultraGridColumn89.Hidden = true;
			ultraGridColumn90.Header.VisiblePosition = 77;
			ultraGridColumn90.Hidden = true;
			ultraGridColumn91.Header.VisiblePosition = 62;
			ultraGridColumn91.Hidden = true;
			ultraGridColumn92.Header.VisiblePosition = 32;
			ultraGridColumn92.Hidden = true;
			ultraGridColumn93.Header.VisiblePosition = 31;
			ultraGridColumn93.Hidden = true;
			ultraGridColumn94.Header.VisiblePosition = 21;
			ultraGridColumn94.Hidden = true;
			ultraGridColumn95.Header.VisiblePosition = 70;
			ultraGridColumn95.Hidden = true;
			ultraGridColumn96.Header.VisiblePosition = 20;
			ultraGridColumn96.Hidden = true;
			ultraGridColumn97.Header.VisiblePosition = 39;
			ultraGridColumn97.Hidden = true;
			ultraGridColumn98.Header.VisiblePosition = 40;
			ultraGridColumn98.Hidden = true;
			ultraGridColumn99.Header.VisiblePosition = 37;
			ultraGridColumn99.Hidden = true;
			ultraGridColumn100.Header.VisiblePosition = 38;
			ultraGridColumn100.Hidden = true;
			ultraGridColumn101.Header.VisiblePosition = 22;
			ultraGridColumn101.Hidden = true;
			ultraGridColumn102.Header.VisiblePosition = 23;
			ultraGridColumn102.Hidden = true;
			ultraGridColumn103.Header.VisiblePosition = 76;
			ultraGridColumn103.Hidden = true;
			ultraGridColumn104.Header.VisiblePosition = 72;
			ultraGridColumn104.Hidden = true;
			ultraGridColumn105.Header.VisiblePosition = 28;
			ultraGridColumn105.Hidden = true;
			ultraGridColumn106.Header.VisiblePosition = 29;
			ultraGridColumn106.Hidden = true;
			ultraGridColumn107.Header.VisiblePosition = 30;
			ultraGridColumn107.Hidden = true;
			ultraGridColumn108.Header.VisiblePosition = 27;
			ultraGridColumn108.Hidden = true;
			ultraGridColumn109.Header.VisiblePosition = 7;
			ultraGridColumn109.Hidden = true;
			ultraGridColumn110.Header.VisiblePosition = 8;
			ultraGridColumn110.Hidden = true;
			ultraGridColumn111.Header.VisiblePosition = 9;
			ultraGridColumn111.Hidden = true;
			ultraGridColumn112.Header.VisiblePosition = 10;
			ultraGridColumn112.Hidden = true;
			ultraGridColumn113.Header.VisiblePosition = 11;
			ultraGridColumn113.Hidden = true;
			ultraGridColumn114.Header.VisiblePosition = 12;
			ultraGridColumn114.Hidden = true;
			ultraGridColumn115.Header.VisiblePosition = 13;
			ultraGridColumn115.Hidden = true;
			ultraGridColumn116.Header.VisiblePosition = 14;
			ultraGridColumn116.Hidden = true;
			ultraGridColumn117.Header.VisiblePosition = 15;
			ultraGridColumn117.Hidden = true;
			ultraGridColumn118.Header.VisiblePosition = 73;
			ultraGridColumn118.Hidden = true;
			ultraGridColumn119.Header.VisiblePosition = 4;
			ultraGridColumn119.Hidden = true;
			ultraGridColumn120.Header.VisiblePosition = 5;
			ultraGridColumn120.Hidden = true;
			ultraGridColumn121.Header.VisiblePosition = 61;
			ultraGridColumn121.Hidden = true;
			ultraGridColumn122.Header.VisiblePosition = 26;
			ultraGridColumn122.Hidden = true;
			ultraGridColumn123.Header.VisiblePosition = 36;
			ultraGridColumn123.Hidden = true;
			ultraGridColumn124.Header.VisiblePosition = 17;
			ultraGridColumn124.Hidden = true;
			ultraGridColumn125.Header.VisiblePosition = 42;
			ultraGridColumn125.Hidden = true;
			ultraGridColumn126.Header.VisiblePosition = 79;
			ultraGridColumn126.Hidden = true;
			ultraGridColumn127.Header.VisiblePosition = 16;
			ultraGridColumn127.Hidden = true;
			ultraGridColumn128.Header.VisiblePosition = 74;
			ultraGridColumn128.Hidden = true;
			ultraGridColumn129.Header.VisiblePosition = 85;
			ultraGridColumn129.Hidden = true;
			ultraGridColumn130.Header.VisiblePosition = 86;
			ultraGridColumn130.Hidden = true;
			ultraGridColumn131.Header.VisiblePosition = 87;
			ultraGridColumn131.Hidden = true;
			ultraGridColumn132.Header.VisiblePosition = 88;
			ultraGridColumn132.Hidden = true;
			ultraGridColumn133.Header.VisiblePosition = 89;
			ultraGridColumn133.Hidden = true;
			ultraGridColumn134.Header.VisiblePosition = 90;
			ultraGridColumn134.Hidden = true;
			ultraGridColumn135.Header.VisiblePosition = 95;
			ultraGridColumn135.Hidden = true;
			ultraGridColumn136.Header.VisiblePosition = 96;
			ultraGridColumn136.Hidden = true;
			ultraGridColumn137.Header.VisiblePosition = 97;
			ultraGridColumn137.Hidden = true;
			ultraGridColumn138.Header.VisiblePosition = 98;
			ultraGridColumn138.Hidden = true;
			ultraGridColumn139.Header.VisiblePosition = 101;
			ultraGridColumn139.Hidden = true;
			ultraGridColumn140.Header.VisiblePosition = 102;
			ultraGridColumn140.Hidden = true;
			ultraGridColumn141.Header.VisiblePosition = 103;
			ultraGridColumn141.Hidden = true;
			ultraGridColumn142.Header.VisiblePosition = 92;
			ultraGridColumn142.Hidden = true;
			ultraGridColumn143.Header.VisiblePosition = 0;
			ultraGridColumn143.Width = 133;
			ultraGridColumn144.Header.VisiblePosition = 1;
			ultraGridColumn144.Width = 86;
			ultraGridColumn145.Header.VisiblePosition = 104;
			ultraGridColumn146.Header.VisiblePosition = 105;
			ultraGridColumn147.Header.VisiblePosition = 106;
			ultraGridColumn148.Header.VisiblePosition = 107;
			ultraGridColumn149.Header.VisiblePosition = 108;
			ultraGridBand10.Columns.AddRange(new object[] {
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
																											ultraGridColumn149});
			this.ultraGrid1.DisplayLayout.BandsSerializer.Add(ultraGridBand10);
			this.ultraGrid1.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
			this.ultraGrid1.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.True;
			this.ultraGrid1.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance33.BackColor = System.Drawing.Color.Transparent;
			this.ultraGrid1.DisplayLayout.Override.CardAreaAppearance = appearance33;
			appearance34.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance34.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance34.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance34.FontData.BoldAsString = "True";
			appearance34.FontData.Name = "Arial";
			appearance34.FontData.SizeInPoints = 10F;
			appearance34.ForeColor = System.Drawing.Color.White;
			appearance34.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.ultraGrid1.DisplayLayout.Override.HeaderAppearance = appearance34;
			this.ultraGrid1.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti;
			appearance35.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance35.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance35.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.RowSelectorAppearance = appearance35;
			appearance36.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance36.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance36.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.SelectedRowAppearance = appearance36;
			this.ultraGrid1.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.ultraGrid1.Location = new System.Drawing.Point(16, 9);
			this.ultraGrid1.Name = "ultraGrid1";
			this.ultraGrid1.Size = new System.Drawing.Size(559, 318);
			this.ultraGrid1.TabIndex = 15;
			this.ultraGrid1.KeyDown += new System.Windows.Forms.KeyEventHandler(this.ultraGrid1_KeyDown);
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
			// ultraTabPageControl3
			// 
			this.ultraTabPageControl3.Controls.Add(this.btFactura);
			this.ultraTabPageControl3.Controls.Add(this.ugKardex);
			this.ultraTabPageControl3.Location = new System.Drawing.Point(-10000, -10769);
			this.ultraTabPageControl3.Name = "ultraTabPageControl3";
			this.ultraTabPageControl3.Size = new System.Drawing.Size(596, 369);
			// 
			// btFactura
			// 
			this.btFactura.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btFactura.Location = new System.Drawing.Point(16, 336);
			this.btFactura.Name = "btFactura";
			this.btFactura.Size = new System.Drawing.Size(64, 26);
			this.btFactura.TabIndex = 1;
			this.btFactura.Text = "Factura";
			this.btFactura.Click += new System.EventHandler(this.btFactura_Click);
			// 
			// ugKardex
			// 
			this.ugKardex.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.ugKardex.Cursor = System.Windows.Forms.Cursors.Default;
			this.ugKardex.DataSource = this.udsKardex;
			appearance37.BackColor = System.Drawing.Color.White;
			this.ugKardex.DisplayLayout.Appearance = appearance37;
			ultraGridBand11.AddButtonCaption = "_DetCompra";
			ultraGridColumn150.Format = "dd/MMM/yyyy";
			ultraGridColumn150.Header.VisiblePosition = 0;
			appearance38.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn151.CellAppearance = appearance38;
			ultraGridColumn151.Format = "#,##0.00";
			ultraGridColumn151.Header.VisiblePosition = 4;
			appearance39.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn152.CellAppearance = appearance39;
			ultraGridColumn152.Format = "#,##0.00";
			ultraGridColumn152.Header.VisiblePosition = 5;
			ultraGridColumn153.Header.VisiblePosition = 2;
			appearance40.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn154.CellAppearance = appearance40;
			ultraGridColumn154.Format = "#,##0.00";
			ultraGridColumn154.Header.VisiblePosition = 3;
			ultraGridColumn155.Header.VisiblePosition = 1;
			ultraGridColumn155.Width = 36;
			ultraGridColumn156.Header.VisiblePosition = 7;
			ultraGridColumn156.Hidden = true;
			appearance41.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn157.CellAppearance = appearance41;
			ultraGridColumn157.Format = "#,##0.00";
			ultraGridColumn157.Header.VisiblePosition = 6;
			ultraGridBand11.Columns.AddRange(new object[] {
																											ultraGridColumn150,
																											ultraGridColumn151,
																											ultraGridColumn152,
																											ultraGridColumn153,
																											ultraGridColumn154,
																											ultraGridColumn155,
																											ultraGridColumn156,
																											ultraGridColumn157});
			this.ugKardex.DisplayLayout.BandsSerializer.Add(ultraGridBand11);
			this.ugKardex.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
			appearance42.BackColor = System.Drawing.Color.Transparent;
			this.ugKardex.DisplayLayout.Override.CardAreaAppearance = appearance42;
			appearance43.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance43.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance43.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance43.FontData.BoldAsString = "True";
			appearance43.FontData.Name = "Arial";
			appearance43.FontData.SizeInPoints = 10F;
			appearance43.ForeColor = System.Drawing.Color.White;
			appearance43.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.ugKardex.DisplayLayout.Override.HeaderAppearance = appearance43;
			appearance44.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance44.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance44.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ugKardex.DisplayLayout.Override.RowSelectorAppearance = appearance44;
			appearance45.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance45.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance45.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ugKardex.DisplayLayout.Override.SelectedRowAppearance = appearance45;
			this.ugKardex.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.ugKardex.Location = new System.Drawing.Point(16, 17);
			this.ugKardex.Name = "ugKardex";
			this.ugKardex.Size = new System.Drawing.Size(567, 310);
			this.ugKardex.TabIndex = 0;
			this.ugKardex.KeyDown += new System.Windows.Forms.KeyEventHandler(this.ugKardex_KeyDown);
			// 
			// udsKardex
			// 
			ultraDataColumn1.DataType = typeof(System.DateTime);
			ultraDataColumn2.DataType = typeof(System.Double);
			ultraDataColumn3.DataType = typeof(System.Double);
			ultraDataColumn5.DataType = typeof(System.Double);
			ultraDataColumn7.DataType = typeof(int);
			ultraDataColumn8.DataType = typeof(System.Double);
			this.udsKardex.Band.Columns.AddRange(new object[] {
																													ultraDataColumn1,
																													ultraDataColumn2,
																													ultraDataColumn3,
																													ultraDataColumn4,
																													ultraDataColumn5,
																													ultraDataColumn6,
																													ultraDataColumn7,
																													ultraDataColumn8});
			// 
			// errorProvider
			// 
			this.errorProvider.ContainerControl = this;
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
			this.ultraTabControl1.Location = new System.Drawing.Point(16, 17);
			this.ultraTabControl1.Name = "ultraTabControl1";
			this.ultraTabControl1.SharedControlsPage = this.ultraTabSharedControlsPage1;
			this.ultraTabControl1.Size = new System.Drawing.Size(600, 395);
			this.ultraTabControl1.TabIndex = 16;
			ultraTab1.TabPage = this.ultraTabPageControl1;
			ultraTab1.Text = "Opciones";
			ultraTab2.TabPage = this.ultraTabPageControl2;
			ultraTab2.Text = "Articulos";
			ultraTab3.TabPage = this.ultraTabPageControl3;
			ultraTab3.Text = "Kardex";
			this.ultraTabControl1.Tabs.AddRange(new Infragistics.Win.UltraWinTabControl.UltraTab[] {
																																															 ultraTab1,
																																															 ultraTab2,
																																															 ultraTab3});
			// 
			// ultraTabSharedControlsPage1
			// 
			this.ultraTabSharedControlsPage1.Location = new System.Drawing.Point(-10000, -10000);
			this.ultraTabSharedControlsPage1.Name = "ultraTabSharedControlsPage1";
			this.ultraTabSharedControlsPage1.Size = new System.Drawing.Size(596, 369);
			// 
			// btCliente
			// 
			this.btCliente.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.btCliente.Location = new System.Drawing.Point(424, 40);
			this.btCliente.Name = "btCliente";
			this.btCliente.Size = new System.Drawing.Size(24, 24);
			this.btCliente.TabIndex = 98;
			this.btCliente.Text = "...";
			this.btCliente.Click += new System.EventHandler(this.btCliente_Click);
			// 
			// CambioConstantes
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 14);
			this.ClientSize = new System.Drawing.Size(632, 427);
			this.Controls.Add(this.ultraTabControl1);
			this.Name = "CambioConstantes";
			this.Text = "Cambio de Constantes";
			this.Closing += new System.ComponentModel.CancelEventHandler(this.CambioConstantes_Closing);
			this.Load += new System.EventHandler(this.CambioConstantes_Load);
			this.ultraTabPageControl1.ResumeLayout(false);
			this.panel4.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.txtPrecioF)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbPrecio1F)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsPrecio)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtPorcIncr)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbSigno)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbPrecio1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbPrecio2)).EndInit();
			this.panel3.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.cmbMarca)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsArticuloTabla)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbGrupo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvGrupo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbColor)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbSubGrupo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbArancel)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbModelo)).EndInit();
			this.panel5.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.txtDescCredito)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDescuento)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIvaDespues)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIvaAntes)).EndInit();
			this.ultraTabPageControl2.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsArticulo)).EndInit();
			this.ultraTabPageControl3.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.ugKardex)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.udsKardex)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraTabControl1)).EndInit();
			this.ultraTabControl1.ResumeLayout(false);
			this.ResumeLayout(false);

		}
		#endregion

		private void cdsPrecio_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsPrecio.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);
		}

		private void cdsArticuloTabla_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsArticuloTabla.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);
		}

		private void cdsArticulo_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsArticulo.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);
		}

		private void chkGrupo_CheckedChanged(object sender, System.EventArgs e)
		{
			cmbGrupo.Enabled = !chkGrupo.Checked;
			if (!chkGrupo.Checked) cmbGrupo.Select();
		}

		private void chkSubGrupo_CheckedChanged(object sender, System.EventArgs e)
		{
			cmbSubGrupo.Enabled = !chkSubGrupo.Checked;
			if (!chkSubGrupo.Checked) cmbSubGrupo.Select();		
		}

		private void chkMarca_CheckedChanged(object sender, System.EventArgs e)
		{
			cmbMarca.Enabled = !chkMarca.Checked;
			if (!chkMarca.Checked) cmbMarca.Select();
		}

		private void chkModelo_CheckedChanged(object sender, System.EventArgs e)
		{
			cmbModelo.Enabled = !chkModelo.Checked;
			if (!chkModelo.Checked) cmbModelo.Select();
		}

		private void chkColor_CheckedChanged(object sender, System.EventArgs e)
		{
			cmbColor.Enabled = !chkColor.Checked;
			if (!chkColor.Checked) cmbColor.Select();
		}

		private void chkArancel_CheckedChanged(object sender, System.EventArgs e)
		{
			cmbArancel.Enabled = !chkArancel.Checked;
			if (!chkArancel.Checked) cmbArancel.Select();
		}

		private void BorraErrores()
		{
			errorProvider.SetError(cmbGrupo, "");
			errorProvider.SetError(cmbSubGrupo, "");
			errorProvider.SetError(cmbMarca, "");
			errorProvider.SetError(cmbModelo, "");
			errorProvider.SetError(cmbColor, "");
			errorProvider.SetError(cmbArancel, "");
		}

		private bool Verifica()
		{
			bool bOk = true;
			BorraErrores();

			if (!chkArancel.Checked)
				if (cmbArancel.Value == null)
				{
					errorProvider.SetError(cmbArancel, "Ingrese Arancel");
					bOk = false;
				}
			if (!chkGrupo.Checked)
				if (cmbGrupo.Value == null)
				{
					errorProvider.SetError(cmbGrupo, "Ingrese Grupo");
					bOk = false;
				}
			if (!chkSubGrupo.Checked)
				if (cmbSubGrupo.Value == null)
				{
					errorProvider.SetError(cmbSubGrupo, "Ingrese SubGrupo");
					bOk = false;
				}
			if (!chkMarca.Checked)
				if (cmbMarca.Value == null)
				{
					errorProvider.SetError(cmbMarca, "Ingrese Marca");
					bOk = false;
				}
			if (!chkModelo.Checked)
				if (cmbModelo.Value == null)
				{
					errorProvider.SetError(cmbModelo, "Ingrese Modelo");
					bOk = false;
				}
			if (!chkColor.Checked)
				if (cmbColor.Value == null)
				{
					errorProvider.SetError(cmbColor, "Ingrese Color");
					bOk = false;
				}
			return bOk;
		}

		private string Filtro()
		{
			string stFiltro = "";
			string stAnd = "";
			if (!chkGrupo.Checked)
				stFiltro += "idGrupoArticulo = " + cmbGrupo.Value.ToString().Trim();
			
			if (stFiltro.Length > 0) stAnd = " And ";
			if (!chkSubGrupo.Checked)
				stFiltro += stAnd + "idSubGrupo = " + cmbSubGrupo.Value.ToString().Trim();

			if (stFiltro.Length > 0) stAnd = " And ";
			if (!chkMarca.Checked)
				stFiltro += stAnd + "idMarca = " + cmbMarca.Value.ToString().Trim();
			
			if (stFiltro.Length > 0) stAnd = " And ";
			if (!chkModelo.Checked)
				stFiltro += stAnd  + "idModelo = " + cmbModelo.Value.ToString().Trim();
			
			if (stFiltro.Length > 0) stAnd = " And ";
			if (!chkArancel.Checked)
				stFiltro += stAnd + "idArancel = " + cmbArancel.Value.ToString().Trim();
			
			if (stFiltro.Length > 0) stAnd = " And ";
			if (!chkColor.Checked)
				stFiltro += stAnd + "idColor = " + cmbColor.Value.ToString().Trim();
			return stFiltro;
		}

		private void btAceptar_Click(object sender, System.EventArgs e)
		{
			if (!Verifica()) return;
			string stFiltro = Filtro();
		
			if (stFiltro.Length > 0)
			{
				C1.Data.FilterCondition fcPadre = new 
					C1.Data.FilterCondition(cdsArticulo, "Articulo", stFiltro);
				C1.Data.FilterConditions fcTotal = new C1.Data.FilterConditions();
				fcTotal.Add(fcPadre);

				this.cdsArticulo.Fill(fcTotal, false);
			}
			else
				cdsArticulo.Fill();
			this.ultraGrid1.Select();
			
			MessageBox.Show("Articulo Filtrados en Ficha Articulos", "Información",
				MessageBoxButtons.OK, MessageBoxIcon.Information);
		}

		private int CuentaArticulos()
		{
			string stSelect = "Select Count(*) From articulo ";
			string stFiltro = Filtro();
			if (stFiltro.Length > 0) stSelect += " Where " + stFiltro;
			int iCuenta = Funcion.iEscalarSQL(cdsArticulo, stSelect);
			return iCuenta;
		}

		CultureInfo us = new CultureInfo("en-US");
		private void btCambiaIVA_Click(object sender, System.EventArgs e)
		{
			if (!Verifica()) return;
			int iCuenta = CuentaArticulos();
			if (iCuenta == 0) 
			{
				MessageBox.Show("No existen Articulos que cumplan la condicion", "Información",
					MessageBoxButtons.OK, MessageBoxIcon.Information);
				return;
			}
			
			double dIvaAntes = (double) txtIvaAntes.Value;
			double dIvaDespues = (double) txtIvaDespues.Value;

			string stMensaje = "Desea Cambiar IVA de " 
				+ dIvaAntes.ToString("0.00", us) + "% a " 
				+ dIvaDespues.ToString("0.00", us) + "% en "
				+ iCuenta.ToString() + " articulos";
			if (DialogResult.No == MessageBox.Show(stMensaje, "Confirmación",
				MessageBoxButtons.YesNo, MessageBoxIcon.Question,
				MessageBoxDefaultButton.Button2)) return;

			string stFiltro = Filtro();
			if (stFiltro.Length == 0) 
				stFiltro = " Where IsNull(Iva, 0) = " + dIvaAntes.ToString("0.00", us);
			else 
				stFiltro = " Where " + stFiltro + " And IsNull(Iva, 0) = " + dIvaAntes.ToString("0.00", us);
			string stSelect = "UPDATE Articulo SET Iva = " 
				+ dIvaDespues.ToString("0.00", us) + " " + stFiltro;

			Funcion.EjecutaSQL(cdsArticulo, stSelect);
			ultraGrid1.Refresh();
			MessageBox.Show("Articulo cambiados de IVA", "Información",
				MessageBoxButtons.OK, MessageBoxIcon.Information);
		}

		private void btIncPorc_Click(object sender, System.EventArgs e)
		{
			if (!Verifica()) return;
			int iCuenta = CuentaArticulos();
			if (iCuenta == 0) 
			{
				MessageBox.Show("No existen Articulos que cumplan la condicion", "Información",
					MessageBoxButtons.OK, MessageBoxIcon.Information);
				return;
			}

			double dPorcIncr = (double) txtPorcIncr.Value;
			string stIncremento = "Incrementar";
			if (cmbSigno.Text == "-") stIncremento = "Disminuir";
			string stMensaje = "Desea " + stIncremento + " el " 
				+ cmbPrecio1.Text.Trim() + " en " 
				+ dPorcIncr.ToString("0.00", us) + "% con respecto al " 
				+ cmbPrecio2.Text.Trim() + " en "
				+ iCuenta.ToString() + " articulos";

			if (DialogResult.No == MessageBox.Show(stMensaje, "Confirmación",
				MessageBoxButtons.YesNo, MessageBoxIcon.Question,
				MessageBoxDefaultButton.Button2)) return;

			string stFiltro = Filtro();
			if (stFiltro.Length > 0) stFiltro = " Where " + stFiltro;

			if (cmbSigno.Text == "-") dPorcIncr = 1 - dPorcIncr / 100.0;
			else dPorcIncr = 1 + dPorcIncr / 100.0;
			string stSelect = "Select IsNull(DecPrecio, 2) From SeteoF";
			int iNumDec = Funcion.iEscalarSQL(cdsArticulo, stSelect, true);

			string stCampo1 = "Precio" + cmbPrecio1.Value.ToString().Trim();
			string stCampo2 = "Precio" + cmbPrecio2.Value.ToString().Trim();
			stSelect = "UPDATE Articulo SET " + stCampo1 + " = Round("
				+ stCampo2 +  "* " + dPorcIncr.ToString("0.0000", us) 
				+ ", " + iNumDec.ToString() + ")" + stFiltro;

			Funcion.EjecutaSQL(cdsArticulo, stSelect);
			ultraGrid1.Refresh();


			if (cmbSigno.Text == "-") stIncremento = "Disminuidos";
			else stIncremento = "Incrementados";
			MessageBox.Show("Precios de Articulos " + stIncremento, "Información",
				MessageBoxButtons.OK, MessageBoxIcon.Information);		
		}

		private void btCancelar_Click(object sender, System.EventArgs e)
		{
			Close();
		}

		private void CambioConstantes_Load(object sender, System.EventArgs e)
		{
			cmbPrecio1.Value = 1;
			cmbPrecio2.Value = 2;		
			string stAudita = "Exec AuditaCrear 53, 6, 'Constantes'";
			Funcion.EjecutaSQL(cdsArticulo, stAudita, true);
		}

		private void btCambiaPrecioValor_Click(object sender, System.EventArgs e)
		{
			errorProvider.SetError(cmbPrecio1F, "");
			if (cmbPrecio1F.Text.Trim().Length == 0) 
			{
				errorProvider.SetError(cmbPrecio1F, "Ingrese Precio");
				return;
			}
			if (!Verifica()) return;
			int iCuenta = CuentaArticulos();
			if (iCuenta == 0) 
			{
				MessageBox.Show("No existen Artículos que cumplan la condición", "Información",
					MessageBoxButtons.OK, MessageBoxIcon.Information);
				return;
			}

			double dPorcIncr = (double) txtPrecioF.Value;
			string stMensaje = "Desea Cambiar el " 
				+ cmbPrecio1F.Text.Trim() + " de los Articulos a "  
				+ dPorcIncr.ToString("0.00", us) 
				+ " en " + iCuenta.ToString() + " articulos";

			if (DialogResult.No == MessageBox.Show(stMensaje, "Confirmación",
				MessageBoxButtons.YesNo, MessageBoxIcon.Question,
				MessageBoxDefaultButton.Button2)) return;

			string stFiltro = Filtro();
			if (stFiltro.Length > 0) stFiltro = " Where " + stFiltro;
			string stCampo = "Precio" + cmbPrecio1F.Value.ToString().Trim();
			string stSelect = "UPDATE Articulo SET " + stCampo + " = "
				+ dPorcIncr.ToString("0.00", us) + stFiltro;

			Funcion.EjecutaSQL(cdsArticulo, stSelect);
			ultraGrid1.Refresh();


			MessageBox.Show("Precios de Articulos Modificados", "Información",
				MessageBoxButtons.OK, MessageBoxIcon.Information);		
		}

		private void CostoPrecio(int iSigno)
		{
			if (ultraGrid1.ActiveRow ==  null)
			{
				MessageBox.Show("Seleccione un Articulo", "Información");
				return;
			}

			string stIdArt = ultraGrid1.ActiveRow.Cells["idArticulo"].Value.ToString();
			string stArticulo = ultraGrid1.ActiveRow.Cells["Articulo"].Value.ToString();
			string stSelect = "Select MAX(Precio) FROM DetCompra "
				+ " Where idArticulo = " + stIdArt + " And Signo = " + iSigno.ToString();
			double dMaximo = Funcion.dEscalarSQL(cdsArticulo, stSelect, false);
			string stMensaje = stArticulo + "\nMaximo: " + dMaximo.ToString("#,##0.00");

			stSelect = "Select Min(Precio) FROM DetCompra "
				+ " Where idArticulo = " + stIdArt + " And Signo = " + iSigno.ToString();
			double dMinimo = Funcion.dEscalarSQL(cdsArticulo, stSelect, false);
			stMensaje += "\nMinimo: " + dMinimo.ToString("#,##0.00");
			MessageBox.Show(stMensaje, "Información");
		}

		private void btCosto_Click(object sender, System.EventArgs e)
		{
			CostoPrecio(1);
		}

		private void btPrecio_Click(object sender, System.EventArgs e)
		{
			CostoPrecio(-1);		
		}

		private void btGrabar_Click(object sender, System.EventArgs e)
		{
			try
			{
				cdsArticulo.Update();
			}
			catch(System.Exception ex)
			{
				MessageBox.Show(ex.Message, "Grabar Artículo");
			}
		}

		private void btKardex_Click(object sender, System.EventArgs e)
		{
			if (ultraGrid1.ActiveRow ==  null)
			{
				MessageBox.Show("Seleccione un Artículo", "Información");
				return;
			}
			if (DialogResult.No == MessageBox.Show(
				"Desea Generar Kardex de artículo seleccionado", 
				"Confirmación", MessageBoxButtons.YesNo, MessageBoxIcon.Question))
				return;

			Cursor = Cursors.WaitCursor;
			int iNumFilas = -1;
			string stIdArticulo = ultraGrid1.ActiveRow.Cells["idArticulo"].Value.ToString();
			string stSelect = "SELECT DetCompra.Cantidad, DetCompra.Signo, "
				+ "DetCompra.Precio, Compra.Fecha, Compra.Numero, Compra.idCompra, "
				+ "DetCompra.CostoPromedio, CompraNumero.Abrev "
				+ "FROM DetCompra LEFT OUTER JOIN Compra "
				+ "ON DetCompra.idCompra=Compra.idCompra "
				+ "LEFT OUTER JOIN CompraNumero "
				+ "ON Compra.idTipoFactura=CompraNumero.idTipoFactura "
				+ " Where DetCompra.idArticulo = " + stIdArticulo
				+ " And (DetCompra.Signo = 1 Or DetCompra.Signo = -1) "
				+ " Order by Compra.Fecha";

			cdsArticulo.Schema.Connections[0].Open();
			SqlDataReader miReader;
			C1.Data.SchemaObjects.Connection miConeccion = cdsArticulo.Schema.Connections[0];
			SqlCommand cmdCant = new SqlCommand (stSelect
				, (SqlConnection) miConeccion.DbConnection);
			udsKardex.Rows.Clear();
			miReader = cmdCant.ExecuteReader();

			try
			{
				while (miReader.Read())
				{
					double dCant = miReader.GetDouble(0);
					int ISigno = miReader.GetInt32(1);
					double dPrecio = miReader.GetDouble(2);
					DateTime dtFecha = miReader.GetDateTime(3);
					string stNumero = miReader.GetString(4);
					int ICompra = miReader.GetInt32(5);
					double dCosto = miReader.GetDouble(6);
					string stTipo = miReader.GetString(7);
					iNumFilas ++;
					udsKardex.Rows.Add();
					udsKardex.Rows[iNumFilas]["Fecha"]=dtFecha;
//					udsKardex.Rows[iNumFilas]["Tipo"]=stArticulo;
					udsKardex.Rows[iNumFilas]["Numero"]=stNumero;
					udsKardex.Rows[iNumFilas]["Precio"]=dPrecio;
					if (ISigno > 0)
						udsKardex.Rows[iNumFilas]["Ingresos"]=dCant;
					else
						udsKardex.Rows[iNumFilas]["Egresos"]=dCant;
					udsKardex.Rows[iNumFilas]["idCompra"]=ICompra;
					udsKardex.Rows[iNumFilas]["Costo"]=dCosto;
					udsKardex.Rows[iNumFilas]["Tipo"]=stTipo;
				}
			}
			catch(System.Exception ex)
			{
				MessageBox.Show(ex.Message, "Llenar Grilla");
			}
			miReader.Close();
			cdsArticulo.Schema.Connections[0].Close();
			Cursor = Cursors.Default;
			MessageBox.Show("Kardex Generado", "Información");

		}

		private void btFactura_Click(object sender, System.EventArgs e)
		{
			if (ugKardex.ActiveRow ==  null)
			{
				MessageBox.Show("Seleccione la Factura", "Información");
				return;
			}
			Cursor = Cursors.WaitCursor;
			int IdCompra = (int) ugKardex.ActiveRow.Cells["idCompra"].Value;
			string stSelect = "Select idTipoFactura FROM Compra Where idCompra = "
				+ IdCompra.ToString();
			int idTrans = Funcion.iEscalarSQL(cdsArticulo, stSelect);
//			Compra miCompra = new Compra(idTrans, IdCompra);
//			miCompra.MdiParent = this.MdiParent;
//			miCompra.Show();
			Cursor = Cursors.Default;
		}

		private void CambioConstantes_Closing(object sender, System.ComponentModel.CancelEventArgs e)
		{
			string stAudita = "Exec AuditaCrear 53, 7, 'Constantes'";
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
						ultraGrid1.PerformAction(UltraGridAction.NextCellByTab, false, false);
						e.Handled = true;
						ultraGrid1.PerformAction(UltraGridAction.EnterEditMode, false, false);
						if(ultraGrid1.ActiveCell.DroppedDown == false)
							ultraGrid1.ActiveCell.DroppedDown = true;
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

		private void ugKardex_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (ugKardex.ActiveCell == null) return;
			try
			{
				if (ugKardex.ActiveCell.Column.Key.ToString().ToUpper() == "NOTAS") return;
				switch(e.KeyValue)
				{
					case (int) Keys.PageDown:
						if(ugKardex.ActiveCell.DroppedDown == false)
							ugKardex.ActiveCell.DroppedDown = true;
						break;

					case (int) Keys.Up:
						if (ugKardex.ActiveCell.DroppedDown) return;

						ugKardex.PerformAction(UltraGridAction.ExitEditMode, false, false);
						ugKardex.PerformAction(UltraGridAction.AboveCell, false, false);
						e.Handled = true;
						ugKardex.PerformAction(UltraGridAction.EnterEditMode, false, false);
						break;
					case (int) Keys.Left:
						ugKardex.PerformAction(UltraGridAction.ExitEditMode, false, false);
						ugKardex.PerformAction(UltraGridAction.PrevCellByTab, false, false);
						e.Handled = true;
						ugKardex.PerformAction(UltraGridAction.EnterEditMode, false, false);
						break;
					case (int) Keys.Down:
						if (ultraGrid1.ActiveCell.DroppedDown) return;
			
						ugKardex.PerformAction(UltraGridAction.ExitEditMode, false, false);
						ugKardex.PerformAction(UltraGridAction.BelowCell, false, false);
						e.Handled = true;
						ultraGrid1.PerformAction(UltraGridAction.EnterEditMode, false, false);
						break;
					case (int) Keys.Enter:
						ugKardex.PerformAction(UltraGridAction.ExitEditMode, false, false);
						ugKardex.PerformAction(UltraGridAction.NextCellByTab, false, false);
						e.Handled = true;
						ugKardex.PerformAction(UltraGridAction.EnterEditMode, false, false);
						if(ugKardex.ActiveCell.DroppedDown == false)
							ugKardex.ActiveCell.DroppedDown = true;
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

		private void btDescuento_Click(object sender, System.EventArgs e)
		{
			if (!Verifica()) return;
			int iCuenta = CuentaArticulos();
			if (iCuenta == 0) 
			{
				MessageBox.Show("No existen Articulos que cumplan la condicion", "Información",
					MessageBoxButtons.OK, MessageBoxIcon.Information);
				return;
			}
			
			double dDescuento = (double) txtDescuento.Value;
			double dDescCredito = (double) txtDescCredito.Value;

			string stMensaje = "Desea Cambiar Descuento a " 
				+ dDescuento.ToString("0.00", us) + "% en "
				+ iCuenta.ToString() + " articulos";
			if (DialogResult.No == MessageBox.Show(stMensaje, "Confirmación",
				MessageBoxButtons.YesNo, MessageBoxIcon.Question,
				MessageBoxDefaultButton.Button2)) return;

			string stFiltro = Filtro();
			if (stFiltro.Length == 0) 
				stFiltro = " ";
			else 
				stFiltro = " Where " + stFiltro;
			string stSelect = "UPDATE Articulo SET CantReqDesc = 1, TipoPromocion = 3, DescuentoContado = " 
				+ dDescuento.ToString("0.00", us) + ", DescuentoCredito = " 
				+ dDescCredito.ToString("0.00", us) + stFiltro;

			Funcion.EjecutaSQL(cdsArticulo, stSelect);
			ultraGrid1.Refresh();
			MessageBox.Show("Articulo cambiados el descuento", "Información",
				MessageBoxButtons.OK, MessageBoxIcon.Information);
		}

		private void btCliente_Click(object sender, System.EventArgs e)
		{
			using (Point_FormulasPrecio miFpp = new Point_FormulasPrecio())
			{
				miFpp.ShowDialog ();
			} 
						
		}
	}
}
