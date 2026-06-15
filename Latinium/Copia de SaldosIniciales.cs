using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de SaldosIniciales.
	/// </summary>
	public class SaldosIniciales : DevExpress.XtraEditors.XtraForm
	{
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbEmpresa;
		private Infragistics.Win.Misc.UltraLabel ultraLabel1;
		private Infragistics.Win.Misc.UltraButton btCxCyP;
		private Infragistics.Win.Misc.UltraButton btInventario;
		private Infragistics.Win.Misc.UltraButton btActivoFijo;
		private C1.Data.C1DataView cdvEmpresa;
		private C1.Data.C1DataSet cdsEmpresa;
		private System.Windows.Forms.ErrorProvider errorProvider;
		private C1.Data.C1DataSet cdsFacturacion;
		private Infragistics.Win.Misc.UltraButton btBalance;
		private Infragistics.Win.Misc.UltraLabel ultraLabel2;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbArticulo;
		private C1.Data.C1DataView cdvArticulo;
		private C1.Data.C1DataSet cdsContabilidad;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo cmbFecha;
		private Infragistics.Win.Misc.UltraLabel ultraLabel3;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public SaldosIniciales()
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
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Empresa", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Empresa");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idEmpresa");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Principal");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Contabilidad");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Numeracion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Retencion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Roles");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idTipoDB");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Reportes");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Servidor");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn11 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Facturacion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn12 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Fox");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn13 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Ruta");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Articulo", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn14 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idArticulo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn15 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idArticuloSRI");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn16 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idColor");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn17 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idModelo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn18 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idGrupoArticulo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn19 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idSubGrupo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn20 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idArancel");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn21 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idMarca");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn22 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idUnidad");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn23 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPromocion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn24 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idTipoGrupo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn25 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idClase");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn26 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idSeccion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn27 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idSucursal");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn28 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPadre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn29 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaIngreso");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn30 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CodAlterno");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn31 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Saldo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn32 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Peso");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn33 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Caracteristica");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn34 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Aplicacion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn35 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Iva");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn36 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Ice");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn37 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Precio1");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn38 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Precio2");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn39 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Precio3");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn40 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Precio4");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn41 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Precio5");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn42 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Precio6");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn43 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CantidadMinima");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn44 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CantidadMaxima");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn45 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Residual");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn46 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CostoUltimo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn47 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaUltimo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn48 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCompra1");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn49 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCompra2");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn50 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCompra3");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn51 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CostoPromedio");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn52 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Gramos");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn53 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ConsumoProm");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn54 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("TiempoReposicion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn55 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("UnidadPeso");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn56 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Depreciacion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn57 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Incr1");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn58 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Incr2");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn59 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Incr3");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn60 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Incr4");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn61 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Incr5");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn62 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Incr6");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn63 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Ubicacion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn64 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("NumeroParte");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn65 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ComisionValor");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn66 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ComisionPorcentaje");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn67 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CantidadRequerida");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn68 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("PrecioPromocion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn69 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CantidadPromocion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn70 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("DescuentoContado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn71 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("DescuentoCredito");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn72 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("DescPorcContado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn73 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("DescPorcCre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn74 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CantReqDesc");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn75 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CantReqPorc");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn76 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("TipoPromocion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn77 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Responsable");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn78 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Comision30");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn79 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Comision60");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn80 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Comision90");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn81 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Comision120");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn82 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bodega1");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn83 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bodega2");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn84 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bodega3");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn85 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bodega4");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn86 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bodega5");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn87 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bodega6");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn88 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bodega7");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn89 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bodega8");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn90 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bodega9");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn91 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("SaldoT");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn92 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Alto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn93 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Ancho");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn94 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Largo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn95 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Color");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn96 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Descontinuado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn97 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Borrar");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn98 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Habilitado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn99 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ValorPeso");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn100 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("BodegaPredef");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn101 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Tiempo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn102 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cantidad1");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn103 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cantidad2");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn104 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cantidad3");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn105 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cantidad4");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn106 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cantidad5");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn107 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("PaqueteUnidad");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn108 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Editar");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn109 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nuevo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn110 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaEditado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn111 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CambiarPrecio");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn112 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("dCompra");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn113 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("dVenta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn114 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Foto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn115 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Notas");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn116 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Articulo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn117 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn118 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("BodegaPredefCR");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn119 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CodPromocion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn120 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CostoFecha");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn121 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Aportable");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn122 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ExigeLote");
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton1 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			this.btCxCyP = new Infragistics.Win.Misc.UltraButton();
			this.btInventario = new Infragistics.Win.Misc.UltraButton();
			this.btActivoFijo = new Infragistics.Win.Misc.UltraButton();
			this.cmbEmpresa = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.cdvEmpresa = new C1.Data.C1DataView();
			this.cdsEmpresa = new C1.Data.C1DataSet();
			this.ultraLabel1 = new Infragistics.Win.Misc.UltraLabel();
			this.cdsContabilidad = new C1.Data.C1DataSet();
			this.errorProvider = new System.Windows.Forms.ErrorProvider();
			this.cdsFacturacion = new C1.Data.C1DataSet();
			this.btBalance = new Infragistics.Win.Misc.UltraButton();
			this.cmbArticulo = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.cdvArticulo = new C1.Data.C1DataView();
			this.ultraLabel2 = new Infragistics.Win.Misc.UltraLabel();
			this.cmbFecha = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.ultraLabel3 = new Infragistics.Win.Misc.UltraLabel();
			((System.ComponentModel.ISupportInitialize)(this.cmbEmpresa)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvEmpresa)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsEmpresa)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsContabilidad)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsFacturacion)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbArticulo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvArticulo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbFecha)).BeginInit();
			this.SuspendLayout();
			// 
			// btCxCyP
			// 
			this.btCxCyP.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.btCxCyP.Location = new System.Drawing.Point(56, 128);
			this.btCxCyP.Name = "btCxCyP";
			this.btCxCyP.Size = new System.Drawing.Size(112, 24);
			this.btCxCyP.TabIndex = 0;
			this.btCxCyP.Text = "Cx Cobrar y Pagar";
			this.btCxCyP.Click += new System.EventHandler(this.btCxCyP_Click);
			// 
			// btInventario
			// 
			this.btInventario.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.btInventario.Location = new System.Drawing.Point(200, 128);
			this.btInventario.Name = "btInventario";
			this.btInventario.Size = new System.Drawing.Size(120, 24);
			this.btInventario.TabIndex = 1;
			this.btInventario.Text = "Inventario";
			this.btInventario.Click += new System.EventHandler(this.btInventario_Click);
			// 
			// btActivoFijo
			// 
			this.btActivoFijo.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.btActivoFijo.Location = new System.Drawing.Point(56, 168);
			this.btActivoFijo.Name = "btActivoFijo";
			this.btActivoFijo.Size = new System.Drawing.Size(112, 24);
			this.btActivoFijo.TabIndex = 2;
			this.btActivoFijo.Text = "Activos Fijos";
			this.btActivoFijo.Click += new System.EventHandler(this.btActivoFijo_Click);
			// 
			// cmbEmpresa
			// 
			this.cmbEmpresa.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbEmpresa.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbEmpresa.DataSource = this.cdvEmpresa;
			this.cmbEmpresa.DisplayLayout.AutoFitColumns = true;
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn1.Width = 166;
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridColumn2.Hidden = true;
			ultraGridColumn3.Header.VisiblePosition = 2;
			ultraGridColumn3.Hidden = true;
			ultraGridColumn4.Header.VisiblePosition = 3;
			ultraGridColumn4.Hidden = true;
			ultraGridColumn5.Header.VisiblePosition = 4;
			ultraGridColumn5.Hidden = true;
			ultraGridColumn6.Header.VisiblePosition = 5;
			ultraGridColumn6.Hidden = true;
			ultraGridColumn7.Header.VisiblePosition = 6;
			ultraGridColumn7.Hidden = true;
			ultraGridColumn8.Header.VisiblePosition = 7;
			ultraGridColumn8.Hidden = true;
			ultraGridColumn9.Header.VisiblePosition = 8;
			ultraGridColumn9.Hidden = true;
			ultraGridColumn10.Header.VisiblePosition = 9;
			ultraGridColumn10.Hidden = true;
			ultraGridColumn11.Header.VisiblePosition = 10;
			ultraGridColumn11.Hidden = true;
			ultraGridColumn12.Header.VisiblePosition = 11;
			ultraGridColumn12.Hidden = true;
			ultraGridColumn12.Width = 19;
			ultraGridColumn13.Header.VisiblePosition = 12;
			ultraGridColumn13.Hidden = true;
			ultraGridColumn13.Width = 56;
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
																										 ultraGridColumn13});
			this.cmbEmpresa.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.cmbEmpresa.DisplayMember = "";
			this.cmbEmpresa.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbEmpresa.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbEmpresa.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbEmpresa.Location = new System.Drawing.Point(160, 24);
			this.cmbEmpresa.Name = "cmbEmpresa";
			this.cmbEmpresa.Size = new System.Drawing.Size(184, 21);
			this.cmbEmpresa.TabIndex = 21;
			this.cmbEmpresa.ValueMember = "";
			// 
			// cdvEmpresa
			// 
			this.cdvEmpresa.BindingContextCtrl = this;
			this.cdvEmpresa.DataSet = this.cdsEmpresa;
			this.cdvEmpresa.Sort = "Empresa";
			this.cdvEmpresa.TableName = "";
			this.cdvEmpresa.TableViewName = "Empresa";
			// 
			// cdsEmpresa
			// 
			this.cdsEmpresa.BindingContextCtrl = this;
			this.cdsEmpresa.DataLibrary = "LibEmpresa";
			this.cdsEmpresa.DataLibraryUrl = "";
			this.cdsEmpresa.DataSetDef = "dsEmpresa";
			this.cdsEmpresa.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsEmpresa.Name = "cdsEmpresa";
			this.cdsEmpresa.SchemaClassName = "LibEmpresa.DataClass";
			this.cdsEmpresa.SchemaDef = null;
			// 
			// ultraLabel1
			// 
			this.ultraLabel1.AutoSize = true;
			this.ultraLabel1.Location = new System.Drawing.Point(24, 24);
			this.ultraLabel1.Name = "ultraLabel1";
			this.ultraLabel1.Size = new System.Drawing.Size(103, 14);
			this.ultraLabel1.TabIndex = 20;
			this.ultraLabel1.Text = "Copiar Archivos de:";
			// 
			// cdsContabilidad
			// 
			this.cdsContabilidad.BindingContextCtrl = this;
			this.cdsContabilidad.DataLibrary = "LibContabilidad";
			this.cdsContabilidad.DataLibraryUrl = "";
			this.cdsContabilidad.DataSetDef = "dsSeteoC";
			this.cdsContabilidad.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsContabilidad.Name = "cdsContabilidad";
			this.cdsContabilidad.SchemaClassName = "LibContabilidad.DataClass";
			this.cdsContabilidad.SchemaDef = null;
			this.cdsContabilidad.BeforeFill += new C1.Data.FillEventHandler(this.cdContabilidad_BeforeFill);
			// 
			// errorProvider
			// 
			this.errorProvider.ContainerControl = this;
			// 
			// cdsFacturacion
			// 
			this.cdsFacturacion.BindingContextCtrl = this;
			this.cdsFacturacion.DataLibrary = "LibFacturacion";
			this.cdsFacturacion.DataLibraryUrl = "";
			this.cdsFacturacion.DataSetDef = "dsArticulo";
			this.cdsFacturacion.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsFacturacion.Name = "cdsFacturacion";
			this.cdsFacturacion.SchemaClassName = "LibFacturacion.DataClass";
			this.cdsFacturacion.SchemaDef = null;
			this.cdsFacturacion.BeforeFill += new C1.Data.FillEventHandler(this.cdsFacturacion_BeforeFill);
			// 
			// btBalance
			// 
			this.btBalance.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.btBalance.Location = new System.Drawing.Point(200, 168);
			this.btBalance.Name = "btBalance";
			this.btBalance.Size = new System.Drawing.Size(120, 24);
			this.btBalance.TabIndex = 22;
			this.btBalance.Text = "Balance Inicial";
			this.btBalance.Click += new System.EventHandler(this.btBalance_Click);
			// 
			// cmbArticulo
			// 
			this.cmbArticulo.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbArticulo.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbArticulo.DataSource = this.cdvArticulo;
			this.cmbArticulo.DisplayLayout.AutoFitColumns = true;
			ultraGridBand2.AddButtonCaption = "Empresa";
			ultraGridColumn14.Header.VisiblePosition = 44;
			ultraGridColumn14.Hidden = true;
			ultraGridColumn14.Width = 8;
			ultraGridColumn15.Header.VisiblePosition = 100;
			ultraGridColumn15.Hidden = true;
			ultraGridColumn15.Width = 32;
			ultraGridColumn16.Header.VisiblePosition = 93;
			ultraGridColumn16.Hidden = true;
			ultraGridColumn16.Width = 8;
			ultraGridColumn17.Header.VisiblePosition = 94;
			ultraGridColumn17.Hidden = true;
			ultraGridColumn17.Width = 8;
			ultraGridColumn18.Header.VisiblePosition = 46;
			ultraGridColumn18.Hidden = true;
			ultraGridColumn18.Width = 8;
			ultraGridColumn19.Header.VisiblePosition = 50;
			ultraGridColumn19.Hidden = true;
			ultraGridColumn19.Width = 8;
			ultraGridColumn20.Header.VisiblePosition = 43;
			ultraGridColumn20.Hidden = true;
			ultraGridColumn20.Width = 8;
			ultraGridColumn21.Header.VisiblePosition = 47;
			ultraGridColumn21.Hidden = true;
			ultraGridColumn21.Width = 8;
			ultraGridColumn22.Header.VisiblePosition = 52;
			ultraGridColumn22.Hidden = true;
			ultraGridColumn22.Width = 8;
			ultraGridColumn23.Header.VisiblePosition = 48;
			ultraGridColumn23.Hidden = true;
			ultraGridColumn23.Width = 8;
			ultraGridColumn24.Header.VisiblePosition = 51;
			ultraGridColumn24.Hidden = true;
			ultraGridColumn24.Width = 8;
			ultraGridColumn25.Header.VisiblePosition = 45;
			ultraGridColumn25.Hidden = true;
			ultraGridColumn25.Width = 8;
			ultraGridColumn26.Header.VisiblePosition = 49;
			ultraGridColumn26.Hidden = true;
			ultraGridColumn26.Width = 8;
			ultraGridColumn27.Header.VisiblePosition = 91;
			ultraGridColumn27.Hidden = true;
			ultraGridColumn27.Width = 8;
			ultraGridColumn28.Header.VisiblePosition = 99;
			ultraGridColumn28.Hidden = true;
			ultraGridColumn28.Width = 23;
			ultraGridColumn29.Header.VisiblePosition = 84;
			ultraGridColumn29.Hidden = true;
			ultraGridColumn29.Width = 17;
			ultraGridColumn30.Header.VisiblePosition = 23;
			ultraGridColumn30.Hidden = true;
			ultraGridColumn30.Width = 8;
			ultraGridColumn31.Header.VisiblePosition = 72;
			ultraGridColumn31.Hidden = true;
			ultraGridColumn31.Width = 8;
			ultraGridColumn32.Header.VisiblePosition = 62;
			ultraGridColumn32.Hidden = true;
			ultraGridColumn32.Width = 8;
			ultraGridColumn33.Header.VisiblePosition = 22;
			ultraGridColumn33.Hidden = true;
			ultraGridColumn33.Width = 8;
			ultraGridColumn34.Header.VisiblePosition = 4;
			ultraGridColumn34.Hidden = true;
			ultraGridColumn34.Width = 8;
			ultraGridColumn35.Header.VisiblePosition = 59;
			ultraGridColumn35.Hidden = true;
			ultraGridColumn35.Width = 8;
			ultraGridColumn36.Header.VisiblePosition = 42;
			ultraGridColumn36.Hidden = true;
			ultraGridColumn36.Width = 8;
			ultraGridColumn37.Header.VisiblePosition = 63;
			ultraGridColumn37.Hidden = true;
			ultraGridColumn37.Width = 8;
			ultraGridColumn38.Header.VisiblePosition = 64;
			ultraGridColumn38.Hidden = true;
			ultraGridColumn38.Width = 8;
			ultraGridColumn39.Header.VisiblePosition = 65;
			ultraGridColumn39.Hidden = true;
			ultraGridColumn39.Width = 8;
			ultraGridColumn40.Header.VisiblePosition = 66;
			ultraGridColumn40.Hidden = true;
			ultraGridColumn40.Width = 8;
			ultraGridColumn41.Header.VisiblePosition = 67;
			ultraGridColumn41.Hidden = true;
			ultraGridColumn41.Width = 8;
			ultraGridColumn42.Header.VisiblePosition = 68;
			ultraGridColumn42.Hidden = true;
			ultraGridColumn42.Width = 8;
			ultraGridColumn43.Header.VisiblePosition = 17;
			ultraGridColumn43.Hidden = true;
			ultraGridColumn43.Width = 8;
			ultraGridColumn44.Header.VisiblePosition = 16;
			ultraGridColumn44.Hidden = true;
			ultraGridColumn44.Width = 8;
			ultraGridColumn45.Header.VisiblePosition = 70;
			ultraGridColumn45.Hidden = true;
			ultraGridColumn45.Width = 8;
			ultraGridColumn46.Header.VisiblePosition = 33;
			ultraGridColumn46.Hidden = true;
			ultraGridColumn46.Width = 8;
			ultraGridColumn47.Header.VisiblePosition = 83;
			ultraGridColumn47.Hidden = true;
			ultraGridColumn47.Width = 17;
			ultraGridColumn48.Header.VisiblePosition = 80;
			ultraGridColumn48.Hidden = true;
			ultraGridColumn48.Width = 8;
			ultraGridColumn49.Header.VisiblePosition = 81;
			ultraGridColumn49.Hidden = true;
			ultraGridColumn49.Width = 8;
			ultraGridColumn50.Header.VisiblePosition = 82;
			ultraGridColumn50.Hidden = true;
			ultraGridColumn50.Width = 8;
			ultraGridColumn51.Header.VisiblePosition = 32;
			ultraGridColumn51.Hidden = true;
			ultraGridColumn51.Width = 8;
			ultraGridColumn52.Header.VisiblePosition = 40;
			ultraGridColumn52.Hidden = true;
			ultraGridColumn52.Width = 8;
			ultraGridColumn53.Header.VisiblePosition = 31;
			ultraGridColumn53.Hidden = true;
			ultraGridColumn53.Width = 8;
			ultraGridColumn54.Header.VisiblePosition = 75;
			ultraGridColumn54.Hidden = true;
			ultraGridColumn54.Width = 8;
			ultraGridColumn55.Header.VisiblePosition = 78;
			ultraGridColumn55.Hidden = true;
			ultraGridColumn55.Width = 8;
			ultraGridColumn56.Header.VisiblePosition = 34;
			ultraGridColumn56.Hidden = true;
			ultraGridColumn56.Width = 8;
			ultraGridColumn57.Header.VisiblePosition = 53;
			ultraGridColumn57.Hidden = true;
			ultraGridColumn57.Width = 8;
			ultraGridColumn58.Header.VisiblePosition = 54;
			ultraGridColumn58.Hidden = true;
			ultraGridColumn58.Width = 8;
			ultraGridColumn59.Header.VisiblePosition = 55;
			ultraGridColumn59.Hidden = true;
			ultraGridColumn59.Width = 8;
			ultraGridColumn60.Header.VisiblePosition = 56;
			ultraGridColumn60.Hidden = true;
			ultraGridColumn60.Width = 8;
			ultraGridColumn61.Header.VisiblePosition = 57;
			ultraGridColumn61.Hidden = true;
			ultraGridColumn61.Width = 8;
			ultraGridColumn62.Header.VisiblePosition = 58;
			ultraGridColumn62.Hidden = true;
			ultraGridColumn62.Width = 8;
			ultraGridColumn63.Header.VisiblePosition = 77;
			ultraGridColumn63.Hidden = true;
			ultraGridColumn63.Width = 8;
			ultraGridColumn64.Header.VisiblePosition = 61;
			ultraGridColumn64.Hidden = true;
			ultraGridColumn64.Width = 8;
			ultraGridColumn65.Header.VisiblePosition = 30;
			ultraGridColumn65.Hidden = true;
			ultraGridColumn65.Width = 8;
			ultraGridColumn66.Header.VisiblePosition = 29;
			ultraGridColumn66.Hidden = true;
			ultraGridColumn66.Width = 8;
			ultraGridColumn67.Header.VisiblePosition = 19;
			ultraGridColumn67.Hidden = true;
			ultraGridColumn67.Width = 8;
			ultraGridColumn68.Header.VisiblePosition = 69;
			ultraGridColumn68.Hidden = true;
			ultraGridColumn68.Width = 8;
			ultraGridColumn69.Header.VisiblePosition = 18;
			ultraGridColumn69.Hidden = true;
			ultraGridColumn69.Width = 8;
			ultraGridColumn70.Header.VisiblePosition = 38;
			ultraGridColumn70.Hidden = true;
			ultraGridColumn70.Width = 8;
			ultraGridColumn71.Header.VisiblePosition = 39;
			ultraGridColumn71.Hidden = true;
			ultraGridColumn71.Width = 8;
			ultraGridColumn72.Header.VisiblePosition = 36;
			ultraGridColumn72.Hidden = true;
			ultraGridColumn72.Width = 8;
			ultraGridColumn73.Header.VisiblePosition = 37;
			ultraGridColumn73.Hidden = true;
			ultraGridColumn73.Width = 8;
			ultraGridColumn74.Header.VisiblePosition = 20;
			ultraGridColumn74.Hidden = true;
			ultraGridColumn74.Width = 8;
			ultraGridColumn75.Header.VisiblePosition = 21;
			ultraGridColumn75.Hidden = true;
			ultraGridColumn75.Width = 8;
			ultraGridColumn76.Header.VisiblePosition = 76;
			ultraGridColumn76.Hidden = true;
			ultraGridColumn76.Width = 8;
			ultraGridColumn77.Header.VisiblePosition = 71;
			ultraGridColumn77.Hidden = true;
			ultraGridColumn77.Width = 8;
			ultraGridColumn78.Header.VisiblePosition = 26;
			ultraGridColumn78.Hidden = true;
			ultraGridColumn78.Width = 8;
			ultraGridColumn79.Header.VisiblePosition = 27;
			ultraGridColumn79.Hidden = true;
			ultraGridColumn79.Width = 8;
			ultraGridColumn80.Header.VisiblePosition = 28;
			ultraGridColumn80.Hidden = true;
			ultraGridColumn80.Width = 8;
			ultraGridColumn81.Header.VisiblePosition = 25;
			ultraGridColumn81.Hidden = true;
			ultraGridColumn81.Width = 8;
			ultraGridColumn82.Header.VisiblePosition = 5;
			ultraGridColumn82.Hidden = true;
			ultraGridColumn82.Width = 8;
			ultraGridColumn83.Header.VisiblePosition = 6;
			ultraGridColumn83.Hidden = true;
			ultraGridColumn83.Width = 8;
			ultraGridColumn84.Header.VisiblePosition = 7;
			ultraGridColumn84.Hidden = true;
			ultraGridColumn84.Width = 8;
			ultraGridColumn85.Header.VisiblePosition = 8;
			ultraGridColumn85.Hidden = true;
			ultraGridColumn85.Width = 8;
			ultraGridColumn86.Header.VisiblePosition = 9;
			ultraGridColumn86.Hidden = true;
			ultraGridColumn86.Width = 8;
			ultraGridColumn87.Header.VisiblePosition = 10;
			ultraGridColumn87.Hidden = true;
			ultraGridColumn87.Width = 8;
			ultraGridColumn88.Header.VisiblePosition = 11;
			ultraGridColumn88.Hidden = true;
			ultraGridColumn88.Width = 8;
			ultraGridColumn89.Header.VisiblePosition = 12;
			ultraGridColumn89.Hidden = true;
			ultraGridColumn89.Width = 8;
			ultraGridColumn90.Header.VisiblePosition = 13;
			ultraGridColumn90.Hidden = true;
			ultraGridColumn90.Width = 8;
			ultraGridColumn91.Header.VisiblePosition = 73;
			ultraGridColumn91.Hidden = true;
			ultraGridColumn91.Width = 8;
			ultraGridColumn92.Header.VisiblePosition = 1;
			ultraGridColumn92.Hidden = true;
			ultraGridColumn92.Width = 32;
			ultraGridColumn93.Header.VisiblePosition = 3;
			ultraGridColumn93.Hidden = true;
			ultraGridColumn93.Width = 8;
			ultraGridColumn94.Header.VisiblePosition = 60;
			ultraGridColumn94.Hidden = true;
			ultraGridColumn94.Width = 8;
			ultraGridColumn95.Header.VisiblePosition = 24;
			ultraGridColumn95.Hidden = true;
			ultraGridColumn95.Width = 8;
			ultraGridColumn96.Header.VisiblePosition = 35;
			ultraGridColumn96.Hidden = true;
			ultraGridColumn96.Width = 17;
			ultraGridColumn97.Header.VisiblePosition = 15;
			ultraGridColumn97.Hidden = true;
			ultraGridColumn97.Width = 17;
			ultraGridColumn98.Header.VisiblePosition = 41;
			ultraGridColumn98.Hidden = true;
			ultraGridColumn98.Width = 17;
			ultraGridColumn99.Header.VisiblePosition = 79;
			ultraGridColumn99.Hidden = true;
			ultraGridColumn99.Width = 17;
			ultraGridColumn100.Header.VisiblePosition = 14;
			ultraGridColumn100.Hidden = true;
			ultraGridColumn100.Width = 8;
			ultraGridColumn101.Header.VisiblePosition = 74;
			ultraGridColumn101.Hidden = true;
			ultraGridColumn101.Width = 8;
			ultraGridColumn102.Header.VisiblePosition = 85;
			ultraGridColumn102.Hidden = true;
			ultraGridColumn102.Width = 8;
			ultraGridColumn103.Header.VisiblePosition = 86;
			ultraGridColumn103.Hidden = true;
			ultraGridColumn103.Width = 8;
			ultraGridColumn104.Header.VisiblePosition = 87;
			ultraGridColumn104.Hidden = true;
			ultraGridColumn104.Width = 8;
			ultraGridColumn105.Header.VisiblePosition = 88;
			ultraGridColumn105.Hidden = true;
			ultraGridColumn105.Width = 8;
			ultraGridColumn106.Header.VisiblePosition = 89;
			ultraGridColumn106.Hidden = true;
			ultraGridColumn106.Width = 8;
			ultraGridColumn107.Header.VisiblePosition = 90;
			ultraGridColumn107.Hidden = true;
			ultraGridColumn107.Width = 8;
			ultraGridColumn108.Header.VisiblePosition = 95;
			ultraGridColumn108.Hidden = true;
			ultraGridColumn108.Width = 17;
			ultraGridColumn109.Header.VisiblePosition = 96;
			ultraGridColumn109.Hidden = true;
			ultraGridColumn109.Width = 17;
			ultraGridColumn110.Header.VisiblePosition = 97;
			ultraGridColumn110.Hidden = true;
			ultraGridColumn110.Width = 17;
			ultraGridColumn111.Header.VisiblePosition = 98;
			ultraGridColumn111.Hidden = true;
			ultraGridColumn111.Width = 17;
			ultraGridColumn112.Header.VisiblePosition = 101;
			ultraGridColumn112.Hidden = true;
			ultraGridColumn112.Width = 24;
			ultraGridColumn113.Header.VisiblePosition = 102;
			ultraGridColumn113.Hidden = true;
			ultraGridColumn113.Width = 22;
			ultraGridColumn114.Header.VisiblePosition = 103;
			ultraGridColumn114.Hidden = true;
			ultraGridColumn114.Width = 37;
			ultraGridColumn115.Header.VisiblePosition = 92;
			ultraGridColumn115.Hidden = true;
			ultraGridColumn115.Width = 8;
			ultraGridColumn116.Header.VisiblePosition = 0;
			ultraGridColumn116.Width = 39;
			ultraGridColumn117.Header.VisiblePosition = 2;
			ultraGridColumn117.Hidden = true;
			ultraGridColumn117.Width = 8;
			ultraGridColumn118.Header.VisiblePosition = 104;
			ultraGridColumn118.Hidden = true;
			ultraGridColumn118.Width = 46;
			ultraGridColumn119.Header.VisiblePosition = 105;
			ultraGridColumn119.Hidden = true;
			ultraGridColumn119.Width = 46;
			ultraGridColumn120.Header.VisiblePosition = 106;
			ultraGridColumn120.Hidden = true;
			ultraGridColumn120.Width = 43;
			ultraGridColumn121.Header.VisiblePosition = 107;
			ultraGridColumn121.Hidden = true;
			ultraGridColumn121.Width = 35;
			ultraGridColumn122.Header.VisiblePosition = 108;
			ultraGridColumn122.Hidden = true;
			ultraGridColumn122.Width = 49;
			ultraGridBand2.Columns.AddRange(new object[] {
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
																										 ultraGridColumn116,
																										 ultraGridColumn117,
																										 ultraGridColumn118,
																										 ultraGridColumn119,
																										 ultraGridColumn120,
																										 ultraGridColumn121,
																										 ultraGridColumn122});
			this.cmbArticulo.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			this.cmbArticulo.DisplayMember = "Articulo";
			this.cmbArticulo.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbArticulo.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbArticulo.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbArticulo.Location = new System.Drawing.Point(160, 56);
			this.cmbArticulo.Name = "cmbArticulo";
			this.cmbArticulo.Size = new System.Drawing.Size(184, 21);
			this.cmbArticulo.TabIndex = 24;
			this.cmbArticulo.ValueMember = "idArticulo";
			// 
			// cdvArticulo
			// 
			this.cdvArticulo.BindingContextCtrl = this;
			this.cdvArticulo.DataSet = this.cdsFacturacion;
			this.cdvArticulo.RowFilter = "idTipoGrupo = 4";
			this.cdvArticulo.TableName = "";
			this.cdvArticulo.TableViewName = "Articulo";
			// 
			// ultraLabel2
			// 
			this.ultraLabel2.AutoSize = true;
			this.ultraLabel2.Location = new System.Drawing.Point(24, 56);
			this.ultraLabel2.Name = "ultraLabel2";
			this.ultraLabel2.Size = new System.Drawing.Size(124, 14);
			this.ultraLabel2.TabIndex = 23;
			this.ultraLabel2.Text = "Saldos Iniciales de CxC";
			// 
			// cmbFecha
			// 
			this.cmbFecha.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton1.Caption = "Today";
			this.cmbFecha.DateButtons.Add(dateButton1);
			this.cmbFecha.Format = "dd/MMM/yyyy";
			this.cmbFecha.Location = new System.Drawing.Point(160, 88);
			this.cmbFecha.Name = "cmbFecha";
			this.cmbFecha.NonAutoSizeHeight = 16;
			this.cmbFecha.Size = new System.Drawing.Size(184, 21);
			this.cmbFecha.SpinButtonsVisible = true;
			this.cmbFecha.TabIndex = 25;
			this.cmbFecha.Value = new System.DateTime(2007, 12, 26, 0, 0, 0, 0);
			// 
			// ultraLabel3
			// 
			this.ultraLabel3.AutoSize = true;
			this.ultraLabel3.Location = new System.Drawing.Point(24, 88);
			this.ultraLabel3.Name = "ultraLabel3";
			this.ultraLabel3.Size = new System.Drawing.Size(36, 14);
			this.ultraLabel3.TabIndex = 26;
			this.ultraLabel3.Text = "Fecha";
			// 
			// SaldosIniciales
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(368, 214);
			this.Controls.Add(this.ultraLabel3);
			this.Controls.Add(this.cmbFecha);
			this.Controls.Add(this.cmbArticulo);
			this.Controls.Add(this.ultraLabel2);
			this.Controls.Add(this.btBalance);
			this.Controls.Add(this.cmbEmpresa);
			this.Controls.Add(this.ultraLabel1);
			this.Controls.Add(this.btActivoFijo);
			this.Controls.Add(this.btInventario);
			this.Controls.Add(this.btCxCyP);
			this.Name = "SaldosIniciales";
			this.Text = "Saldos Iniciales";
			this.Closing += new System.ComponentModel.CancelEventHandler(this.SaldosIniciales_Closing);
			this.Load += new System.EventHandler(this.SaldosIniciales_Load);
			((System.ComponentModel.ISupportInitialize)(this.cmbEmpresa)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvEmpresa)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsEmpresa)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsContabilidad)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsFacturacion)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbArticulo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvArticulo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbFecha)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void BorraErrores()
		{
			errorProvider.SetError(cmbEmpresa, "");
			errorProvider.SetError(cmbArticulo, "");
		}

		private bool Verifica()
		{
			bool bOk = true;
			BorraErrores();
			if (cmbEmpresa.Text.Trim().Length == 0)
			{
				bOk = false;
				errorProvider.SetError(cmbEmpresa, "Ingrese Empresa");
			}
			if (cmbEmpresa.Text.Trim().ToUpper() == MenuLatinium.stDirConta.Trim().ToUpper())
			{
				bOk = false;
				errorProvider.SetError(cmbEmpresa, "Empresa de Origen y Destino no pueden ser las mismas");
			}
			if (cmbArticulo.Text.Trim().Length == 0)
			{
				bOk = false;
				errorProvider.SetError(cmbArticulo, 
					"Ingrese Articulo que pertenesca a un Grupo de Articulos de Saldos Iniciales");
			}
			return bOk;
		}

		private void btCxCyP_Click(object sender, System.EventArgs e)
		{
			if (!Verifica()) return;

			Cursor = Cursors.WaitCursor;
			DateTime dtFecha = (DateTime) cmbFecha.Value;
			dtFecha = dtFecha.AddDays(1);
			int IdArt = (int) cmbArticulo.Value;
			string stBaseOrigen = cmbEmpresa.ActiveRow.Cells["Facturacion"].Value.ToString();

			// Borrar y Copia datos de Tabla Anterior
			string stSelect = "EXEC SaldoCXCInicial '" + stBaseOrigen + "', '"
				+ dtFecha.ToString("yyyyMMdd") + "', " + IdArt.ToString();
			Funcion.EjecutaSQL(cdsFacturacion, stSelect);

			Cursor = Cursors.Default;
			MessageBox.Show("Copia de Saldos Iniciales Terminada", "Información",
				MessageBoxButtons.OK, MessageBoxIcon.Information);
		}

		private void cdContabilidad_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsContabilidad.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirConta);		
		}

		private void cdsFacturacion_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsFacturacion.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);		
		}

		private void btBalance_Click(object sender, System.EventArgs e)
		{
			if (!Verifica()) return;

			Cursor = Cursors.WaitCursor;
			DateTime dtFecha = (DateTime) cmbFecha.Value;
			string stFecha = dtFecha.ToString("yyyyMMdd");
			string stBaseOrigen = cmbEmpresa.ActiveRow.Cells["Facturacion"].Value.ToString();

			// Borrar y Copia datos de Tabla Anterior
			string stSelect = "EXEC CopiaBalance '" + stFecha + "', '" 
				+ stBaseOrigen + "'";
			Funcion.EjecutaSQL(cdsFacturacion, stSelect);
			Cursor = Cursors.Default;
			MessageBox.Show("Creacion de Asiento Inicial Terminado", "Información",
				MessageBoxButtons.OK, MessageBoxIcon.Information);
		}

		private void btActivoFijo_Click(object sender, System.EventArgs e)
		{
			if (!Verifica()) return;

			Cursor = Cursors.WaitCursor;
			DateTime dtFecha = DateTime.Today;
			dtFecha = dtFecha.AddDays(1);
			string stBaseOrigen = cmbEmpresa.ActiveRow.Cells["Facturacion"].Value.ToString();

			// Borrar y Copia datos de Tabla Anterior
			string stExec = "EXEC CopiaActivoFijo '" + dtFecha.ToString("yyyyMMdd") + "', '" + stBaseOrigen + "'";
			Funcion.EjecutaSQL(cdsFacturacion, stExec, true);
			Cursor = Cursors.Default;
			MessageBox.Show("Copia de Activos Fijos Terminada", "Información",
				MessageBoxButtons.OK, MessageBoxIcon.Information);
		}

		private void btInventario_Click(object sender, System.EventArgs e)
		{
			if (!Verifica()) return;

			Cursor = Cursors.WaitCursor;
			DateTime dtFecha = (DateTime) cmbFecha.Value;
			dtFecha = dtFecha.AddDays(1);
			string stFecha = dtFecha.ToString("yyyyMMdd");
			string stBaseOrigen = cmbEmpresa.ActiveRow.Cells["Facturacion"].Value.ToString();

			// Borrar y Copia datos de Tabla Anterior
			string stSelect = "EXEC CopiaInventario '" + stFecha + "', '" 
				+ stBaseOrigen + "'";
			Funcion.EjecutaSQL(cdsFacturacion, stSelect);
			Cursor = Cursors.Default;
			MessageBox.Show("Copia de Inventario Terminada", "Información",
				MessageBoxButtons.OK, MessageBoxIcon.Information);
		}

		private void SaldosIniciales_Load(object sender, System.EventArgs e)
		{
			cmbFecha.Value = DateTime.Today;
			string stAudita = "Exec AuditaCrear 53, 6, 'Saldos Iniciales'";
			Funcion.EjecutaSQL(cdsContabilidad, stAudita, true);
		}

		private void SaldosIniciales_Closing(object sender, System.ComponentModel.CancelEventArgs e)
		{
			string stAudita = "Exec AuditaCrear 53, 7, 'Saldos Iniciales'";
			Funcion.EjecutaSQL(cdsContabilidad, stAudita, true);
		}
	}
}
