using System;
using System.Data.SqlClient;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de ActualizaSaldo.
	/// </summary>
	public class ActualizaSaldo : DevExpress.XtraEditors.XtraForm
	{
		private C1.Data.C1DataSet cdsCompra;
		private C1.Data.C1DataSet cdsArticulo;
		private DevExpress.XtraEditors.SimpleButton btSaldos;
		private DevExpress.XtraEditors.SimpleButton btCostos;
		private DevExpress.XtraEditors.SimpleButton btAceptar;
		private DevExpress.XtraEditors.SimpleButton btCancelar;
		private DevExpress.XtraEditors.SimpleButton btBorraBlanco;
		private C1.Data.C1DataSet cdsGrupo;
		private DevExpress.XtraEditors.SimpleButton btGrupoArt;
		private DevExpress.XtraEditors.SimpleButton btSaldoCxC;
		private DevExpress.XtraEditors.SimpleButton btSaldoCosto;
		private DevExpress.XtraEditors.SimpleButton btSigno;
		private System.Windows.Forms.Label label1;
		private System.Windows.Forms.HelpProvider helpProvider1;
		private DevExpress.XtraEditors.SimpleButton btActEgreso;
		private DevExpress.XtraEditors.SimpleButton simpleButton1;
		private C1.Data.C1DataSet cdsSeteoF;
		private System.Windows.Forms.Label label2;
		private System.Windows.Forms.Label label3;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo cmbHasta;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo cmbDesde;
		private DevExpress.XtraEditors.SimpleButton btIngreso;
		private DevExpress.XtraEditors.SimpleButton btSaldoCostoB;
		private DevExpress.XtraEditors.SimpleButton btEnvioCostoB;
		private DevExpress.XtraEditors.SimpleButton btnFIFO;
		private C1.Data.C1DataView cdvArticulo;
		private C1.Data.C1DataSet c1DataSet1;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbArt;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkArt;
		private DevExpress.XtraEditors.SimpleButton btSaldoCostoLotes;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public ActualizaSaldo()
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
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton1 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton2 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.Appearance appearance1 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Articulo", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Alto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Ancho");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Aplicacion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bodega1");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bodega2");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bodega3");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bodega4");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bodega5");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bodega6");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bodega7");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn11 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bodega8");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn12 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bodega9");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn13 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("BodegaPredef");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn14 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Borrar");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn15 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CambiarPrecio");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn16 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cantidad1");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn17 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cantidad2");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn18 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cantidad3");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn19 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cantidad4");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn20 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cantidad5");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn21 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CantidadMaxima");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn22 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CantidadMinima");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn23 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CantidadPromocion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn24 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CantidadRequerida");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn25 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CantReqDesc");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn26 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CantReqPorc");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn27 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Caracteristica");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn28 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CodAlterno");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn29 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Color");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn30 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Comision120");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn31 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Comision30");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn32 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Comision60");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn33 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Comision90");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn34 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ComisionPorcentaje");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn35 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ComisionValor");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn36 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ConsumoProm");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn37 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CostoPromedio");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn38 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CostoUltimo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn39 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("dCompra");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn40 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Depreciacion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn41 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Descontinuado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn42 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("DescPorcContado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn43 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("DescPorcCre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn44 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("DescuentoContado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn45 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("DescuentoCredito");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn46 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("dVenta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn47 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Editar");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn48 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaEditado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn49 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaIngreso");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn50 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaUltimo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn51 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Gramos");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn52 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Habilitado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn53 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Ice");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn54 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idArancel");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn55 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idArticulo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn56 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idArticuloSRI");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn57 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idClase");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn58 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idColor");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn59 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCompra1");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn60 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCompra2");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn61 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCompra3");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn62 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idGrupoArticulo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn63 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idMarca");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn64 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idModelo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn65 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPadre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn66 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPromocion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn67 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idSeccion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn68 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idSubGrupo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn69 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idSucursal");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn70 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idTipoGrupo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn71 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idUnidad");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn72 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Incr1");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn73 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Incr2");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn74 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Incr3");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn75 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Incr4");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn76 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Incr5");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn77 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Incr6");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn78 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Iva");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn79 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Largo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn80 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nuevo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn81 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("NumeroParte");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn82 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("PaqueteUnidad");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn83 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Peso");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn84 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Precio1");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn85 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Precio2");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn86 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Precio3");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn87 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Precio4");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn88 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Precio5");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn89 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Precio6");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn90 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("PrecioPromocion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn91 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Residual");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn92 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Responsable");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn93 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Saldo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn94 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("SaldoT");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn95 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Tiempo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn96 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("TiempoReposicion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn97 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("TipoPromocion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn98 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Ubicacion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn99 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("UnidadPeso");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn100 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ValorPeso");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn101 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Foto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn102 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Notas");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn103 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Articulo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn104 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn105 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("BodegaPredefCR");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn106 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CodPromocion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn107 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CostoFecha");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn108 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Aportable");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn109 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ExigeLote");
			this.cdsCompra = new C1.Data.C1DataSet();
			this.cdsArticulo = new C1.Data.C1DataSet();
			this.btSaldos = new DevExpress.XtraEditors.SimpleButton();
			this.btCostos = new DevExpress.XtraEditors.SimpleButton();
			this.btAceptar = new DevExpress.XtraEditors.SimpleButton();
			this.btCancelar = new DevExpress.XtraEditors.SimpleButton();
			this.btBorraBlanco = new DevExpress.XtraEditors.SimpleButton();
			this.cdsGrupo = new C1.Data.C1DataSet();
			this.btGrupoArt = new DevExpress.XtraEditors.SimpleButton();
			this.cmbDesde = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.btSaldoCxC = new DevExpress.XtraEditors.SimpleButton();
			this.btSaldoCosto = new DevExpress.XtraEditors.SimpleButton();
			this.btSigno = new DevExpress.XtraEditors.SimpleButton();
			this.label1 = new System.Windows.Forms.Label();
			this.helpProvider1 = new System.Windows.Forms.HelpProvider();
			this.btActEgreso = new DevExpress.XtraEditors.SimpleButton();
			this.simpleButton1 = new DevExpress.XtraEditors.SimpleButton();
			this.cmbHasta = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.label2 = new System.Windows.Forms.Label();
			this.label3 = new System.Windows.Forms.Label();
			this.btIngreso = new DevExpress.XtraEditors.SimpleButton();
			this.btSaldoCostoB = new DevExpress.XtraEditors.SimpleButton();
			this.btEnvioCostoB = new DevExpress.XtraEditors.SimpleButton();
			this.btnFIFO = new DevExpress.XtraEditors.SimpleButton();
			this.cdvArticulo = new C1.Data.C1DataView();
			this.c1DataSet1 = new C1.Data.C1DataSet();
			this.cmbArt = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.chkArt = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.btSaldoCostoLotes = new DevExpress.XtraEditors.SimpleButton();
			((System.ComponentModel.ISupportInitialize)(this.cdsCompra)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsArticulo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsGrupo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbDesde)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbHasta)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvArticulo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.c1DataSet1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbArt)).BeginInit();
			this.SuspendLayout();
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
			// btSaldos
			// 
			this.btSaldos.Location = new System.Drawing.Point(480, 176);
			this.btSaldos.Name = "btSaldos";
			this.btSaldos.Size = new System.Drawing.Size(75, 25);
			this.btSaldos.TabIndex = 0;
			this.btSaldos.Text = "Saldos";
			this.btSaldos.Visible = false;
			this.btSaldos.Click += new System.EventHandler(this.btSaldos_Click);
			this.btSaldos.MouseDown += new System.Windows.Forms.MouseEventHandler(this.btSaldos_MouseDown);
			// 
			// btCostos
			// 
			this.btCostos.Location = new System.Drawing.Point(480, 216);
			this.btCostos.Name = "btCostos";
			this.btCostos.Size = new System.Drawing.Size(75, 25);
			this.btCostos.TabIndex = 1;
			this.btCostos.Text = "Costos";
			this.btCostos.Visible = false;
			this.btCostos.Click += new System.EventHandler(this.btCostos_Click);
			// 
			// btAceptar
			// 
			this.btAceptar.Location = new System.Drawing.Point(352, 207);
			this.btAceptar.Name = "btAceptar";
			this.btAceptar.Size = new System.Drawing.Size(75, 25);
			this.btAceptar.TabIndex = 2;
			this.btAceptar.Text = "&Aceptar";
			this.btAceptar.Click += new System.EventHandler(this.btAceptar_Click);
			// 
			// btCancelar
			// 
			this.btCancelar.Location = new System.Drawing.Point(352, 241);
			this.btCancelar.Name = "btCancelar";
			this.btCancelar.Size = new System.Drawing.Size(75, 25);
			this.btCancelar.TabIndex = 3;
			this.btCancelar.Text = "&Cancelar";
			this.btCancelar.Click += new System.EventHandler(this.btCancelar_Click);
			// 
			// btBorraBlanco
			// 
			this.btBorraBlanco.Location = new System.Drawing.Point(56, 164);
			this.btBorraBlanco.Name = "btBorraBlanco";
			this.btBorraBlanco.Size = new System.Drawing.Size(96, 24);
			this.btBorraBlanco.TabIndex = 4;
			this.btBorraBlanco.Text = "Borra Blanco";
			this.btBorraBlanco.Click += new System.EventHandler(this.btBorraBlanco_Click);
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
			this.cdsGrupo.BeforeFill += new C1.Data.FillEventHandler(this.cdsGrupo_BeforeFill);
			// 
			// btGrupoArt
			// 
			this.btGrupoArt.Location = new System.Drawing.Point(56, 207);
			this.btGrupoArt.Name = "btGrupoArt";
			this.btGrupoArt.Size = new System.Drawing.Size(96, 25);
			this.btGrupoArt.TabIndex = 5;
			this.btGrupoArt.Text = "Actualiza Grupo";
			this.btGrupoArt.Click += new System.EventHandler(this.btGrupoArt_Click);
			// 
			// cmbDesde
			// 
			this.cmbDesde.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsSeteoF, "SeteoF.FechaInicio"));
			dateButton1.Caption = "Today";
			this.cmbDesde.DateButtons.Add(dateButton1);
			this.cmbDesde.Format = "dd/MMM/yyyy";
			this.cmbDesde.Location = new System.Drawing.Point(112, 86);
			this.cmbDesde.Name = "cmbDesde";
			this.cmbDesde.NonAutoSizeHeight = 23;
			this.cmbDesde.Size = new System.Drawing.Size(112, 21);
			this.cmbDesde.TabIndex = 6;
			this.cmbDesde.Value = new System.DateTime(2009, 6, 21, 0, 0, 0, 0);
			// 
			// cdsSeteoF
			// 
			this.cdsSeteoF.BindingContextCtrl = this;
			this.cdsSeteoF.DataLibrary = "LibFacturacion";
			this.cdsSeteoF.DataLibraryUrl = "";
			this.cdsSeteoF.DataSetDef = "dsSeteoF";
			this.cdsSeteoF.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsSeteoF.Name = "cdsSeteoF";
			this.cdsSeteoF.SchemaClassName = "LibFacturacion.DataClass";
			this.cdsSeteoF.SchemaDef = null;
			this.cdsSeteoF.BeforeFill += new C1.Data.FillEventHandler(this.cdsSeteoF_BeforeFill);
			// 
			// btSaldoCxC
			// 
			this.btSaldoCxC.Location = new System.Drawing.Point(200, 129);
			this.btSaldoCxC.Name = "btSaldoCxC";
			this.btSaldoCxC.Size = new System.Drawing.Size(96, 25);
			this.btSaldoCxC.TabIndex = 7;
			this.btSaldoCxC.Text = "Saldos CxC";
			this.btSaldoCxC.Click += new System.EventHandler(this.btSaldoCxC_Click);
			this.btSaldoCxC.MouseUp += new System.Windows.Forms.MouseEventHandler(this.btSaldoCxC_MouseUp);
			// 
			// btSaldoCosto
			// 
			this.btSaldoCosto.Location = new System.Drawing.Point(56, 129);
			this.btSaldoCosto.Name = "btSaldoCosto";
			this.btSaldoCosto.Size = new System.Drawing.Size(96, 25);
			this.btSaldoCosto.TabIndex = 8;
			this.btSaldoCosto.Text = "Saldos y Costo";
			this.btSaldoCosto.Click += new System.EventHandler(this.btSaldoCosto_Click);
			this.btSaldoCosto.MouseUp += new System.Windows.Forms.MouseEventHandler(this.btSaldoCosto_MouseDown);
			// 
			// btSigno
			// 
			this.btSigno.Location = new System.Drawing.Point(56, 241);
			this.btSigno.Name = "btSigno";
			this.btSigno.Size = new System.Drawing.Size(96, 25);
			this.btSigno.TabIndex = 9;
			this.btSigno.Text = "Actualiza Signo";
			this.btSigno.Click += new System.EventHandler(this.btSigno_Click);
			this.btSigno.MouseUp += new System.Windows.Forms.MouseEventHandler(this.btSaldos_MouseDown);
			// 
			// label1
			// 
			this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label1.ForeColor = System.Drawing.Color.Brown;
			this.label1.Location = new System.Drawing.Point(56, 9);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(392, 69);
			this.label1.TabIndex = 10;
			this.label1.Text = "Todos estos Procesos deben Terminar.  Si Ud. Presiona un botón y el proceso no te" +
				"rmina debe regresar a esta pantalla y presionar otra vez el mismo boton hasta qu" +
				"e el proceso termine";
			// 
			// helpProvider1
			// 
			this.helpProvider1.HelpNamespace = "C:\\Latinium\\Latinium.chm";
			// 
			// btActEgreso
			// 
			this.btActEgreso.Location = new System.Drawing.Point(200, 241);
			this.btActEgreso.Name = "btActEgreso";
			this.btActEgreso.Size = new System.Drawing.Size(96, 25);
			this.btActEgreso.TabIndex = 11;
			this.btActEgreso.Text = "Act. Egreso";
			this.btActEgreso.Click += new System.EventHandler(this.btActEgreso_Click);
			// 
			// simpleButton1
			// 
			this.simpleButton1.Location = new System.Drawing.Point(200, 164);
			this.simpleButton1.Name = "simpleButton1";
			this.simpleButton1.Size = new System.Drawing.Size(96, 24);
			this.simpleButton1.TabIndex = 12;
			this.simpleButton1.Text = "GraphicSource";
			this.simpleButton1.Click += new System.EventHandler(this.simpleButton1_Click);
			// 
			// cmbHasta
			// 
			dateButton2.Caption = "Today";
			this.cmbHasta.DateButtons.Add(dateButton2);
			this.cmbHasta.Format = "dd/MMM/yyyy";
			this.cmbHasta.Location = new System.Drawing.Point(304, 86);
			this.cmbHasta.Name = "cmbHasta";
			this.cmbHasta.NonAutoSizeHeight = 23;
			this.cmbHasta.Size = new System.Drawing.Size(112, 21);
			this.cmbHasta.TabIndex = 13;
			this.cmbHasta.Value = new System.DateTime(2009, 6, 21, 0, 0, 0, 0);
			// 
			// label2
			// 
			this.label2.AutoSize = true;
			this.label2.Location = new System.Drawing.Point(56, 86);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(35, 16);
			this.label2.TabIndex = 14;
			this.label2.Text = "Desde";
			// 
			// label3
			// 
			this.label3.AutoSize = true;
			this.label3.Location = new System.Drawing.Point(248, 86);
			this.label3.Name = "label3";
			this.label3.Size = new System.Drawing.Size(33, 16);
			this.label3.TabIndex = 15;
			this.label3.Text = "Hasta";
			// 
			// btIngreso
			// 
			this.btIngreso.Location = new System.Drawing.Point(200, 207);
			this.btIngreso.Name = "btIngreso";
			this.btIngreso.Size = new System.Drawing.Size(96, 25);
			this.btIngreso.TabIndex = 16;
			this.btIngreso.Text = "Act. Ingreso";
			this.btIngreso.Click += new System.EventHandler(this.btIngreso_Click);
			this.btIngreso.MouseUp += new System.Windows.Forms.MouseEventHandler(this.btIngreso_MouseUp);
			// 
			// btSaldoCostoB
			// 
			this.btSaldoCostoB.Location = new System.Drawing.Point(344, 129);
			this.btSaldoCostoB.Name = "btSaldoCostoB";
			this.btSaldoCostoB.Size = new System.Drawing.Size(96, 25);
			this.btSaldoCostoB.TabIndex = 17;
			this.btSaldoCostoB.Text = "Costo x Bodega";
			this.btSaldoCostoB.Click += new System.EventHandler(this.btSaldoCostoB_Click);
			// 
			// btEnvioCostoB
			// 
			this.btEnvioCostoB.Location = new System.Drawing.Point(344, 164);
			this.btEnvioCostoB.Name = "btEnvioCostoB";
			this.btEnvioCostoB.Size = new System.Drawing.Size(96, 24);
			this.btEnvioCostoB.TabIndex = 18;
			this.btEnvioCostoB.Text = "Envío Costo P.";
			this.btEnvioCostoB.Click += new System.EventHandler(this.btEnvioCostoB_Click);
			// 
			// btnFIFO
			// 
			this.btnFIFO.Location = new System.Drawing.Point(480, 138);
			this.btnFIFO.Name = "btnFIFO";
			this.btnFIFO.Size = new System.Drawing.Size(80, 25);
			this.btnFIFO.TabIndex = 19;
			this.btnFIFO.Text = "FIFO";
			this.btnFIFO.Visible = false;
			this.btnFIFO.Click += new System.EventHandler(this.simpleButton2_Click);
			// 
			// cdvArticulo
			// 
			this.cdvArticulo.BindingContextCtrl = this;
			this.cdvArticulo.DataSet = this.cdsArticulo;
			this.cdvArticulo.Sort = "Articulo";
			this.cdvArticulo.TableName = "Articulo";
			this.cdvArticulo.TableViewName = "";
			// 
			// c1DataSet1
			// 
			this.c1DataSet1.BindingContextCtrl = this;
			this.c1DataSet1.CaseSensitive = false;
			this.c1DataSet1.DataLibrary = "LibFacturacion";
			this.c1DataSet1.DataLibraryUrl = "";
			this.c1DataSet1.DataSetDef = "dsArticulo";
			this.c1DataSet1.FillOnRequest = false;
			this.c1DataSet1.Locale = new System.Globalization.CultureInfo("es-EC");
			this.c1DataSet1.Name = "c1DataSet1";
			this.c1DataSet1.SchemaClassName = "LibFacturacion.DataClass";
			this.c1DataSet1.SchemaDef = null;
			// 
			// cmbArt
			// 
			appearance1.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbArt.Appearance = appearance1;
			this.cmbArt.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbArt.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbArt.DataSource = this.cdvArticulo;
			ultraGridColumn1.Header.VisiblePosition = 80;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn2.Header.VisiblePosition = 81;
			ultraGridColumn2.Hidden = true;
			ultraGridColumn3.Header.VisiblePosition = 22;
			ultraGridColumn3.Hidden = true;
			ultraGridColumn4.Header.VisiblePosition = 70;
			ultraGridColumn4.Hidden = true;
			ultraGridColumn5.Header.VisiblePosition = 71;
			ultraGridColumn5.Hidden = true;
			ultraGridColumn6.Header.VisiblePosition = 72;
			ultraGridColumn6.Hidden = true;
			ultraGridColumn7.Header.VisiblePosition = 73;
			ultraGridColumn7.Hidden = true;
			ultraGridColumn8.Header.VisiblePosition = 74;
			ultraGridColumn8.Hidden = true;
			ultraGridColumn9.Header.VisiblePosition = 75;
			ultraGridColumn9.Hidden = true;
			ultraGridColumn10.Header.VisiblePosition = 76;
			ultraGridColumn10.Hidden = true;
			ultraGridColumn11.Header.VisiblePosition = 77;
			ultraGridColumn11.Hidden = true;
			ultraGridColumn12.Header.VisiblePosition = 78;
			ultraGridColumn12.Hidden = true;
			ultraGridColumn13.Header.VisiblePosition = 87;
			ultraGridColumn13.Hidden = true;
			ultraGridColumn14.Header.VisiblePosition = 1;
			ultraGridColumn14.Hidden = true;
			ultraGridColumn14.Width = 42;
			ultraGridColumn15.Header.VisiblePosition = 98;
			ultraGridColumn15.Hidden = true;
			ultraGridColumn16.Header.VisiblePosition = 89;
			ultraGridColumn16.Hidden = true;
			ultraGridColumn17.Header.VisiblePosition = 90;
			ultraGridColumn17.Hidden = true;
			ultraGridColumn18.Header.VisiblePosition = 91;
			ultraGridColumn18.Hidden = true;
			ultraGridColumn19.Header.VisiblePosition = 92;
			ultraGridColumn19.Hidden = true;
			ultraGridColumn20.Header.VisiblePosition = 93;
			ultraGridColumn20.Hidden = true;
			ultraGridColumn21.Header.VisiblePosition = 32;
			ultraGridColumn21.Hidden = true;
			ultraGridColumn22.Header.VisiblePosition = 31;
			ultraGridColumn22.Hidden = true;
			ultraGridColumn23.Header.VisiblePosition = 57;
			ultraGridColumn23.Hidden = true;
			ultraGridColumn24.Header.VisiblePosition = 55;
			ultraGridColumn24.Hidden = true;
			ultraGridColumn25.Header.VisiblePosition = 62;
			ultraGridColumn25.Hidden = true;
			ultraGridColumn26.Header.VisiblePosition = 63;
			ultraGridColumn26.Hidden = true;
			ultraGridColumn27.Header.VisiblePosition = 21;
			ultraGridColumn27.Hidden = true;
			ultraGridColumn28.Header.VisiblePosition = 18;
			ultraGridColumn28.Hidden = true;
			ultraGridColumn29.Header.VisiblePosition = 83;
			ultraGridColumn29.Hidden = true;
			ultraGridColumn30.Header.VisiblePosition = 69;
			ultraGridColumn30.Hidden = true;
			ultraGridColumn31.Header.VisiblePosition = 66;
			ultraGridColumn31.Hidden = true;
			ultraGridColumn32.Header.VisiblePosition = 67;
			ultraGridColumn32.Hidden = true;
			ultraGridColumn33.Header.VisiblePosition = 68;
			ultraGridColumn33.Hidden = true;
			ultraGridColumn34.Header.VisiblePosition = 54;
			ultraGridColumn34.Hidden = true;
			ultraGridColumn35.Header.VisiblePosition = 53;
			ultraGridColumn35.Hidden = true;
			ultraGridColumn36.Header.VisiblePosition = 41;
			ultraGridColumn36.Hidden = true;
			ultraGridColumn37.Header.VisiblePosition = 39;
			ultraGridColumn37.Hidden = true;
			ultraGridColumn38.Header.VisiblePosition = 34;
			ultraGridColumn38.Hidden = true;
			ultraGridColumn39.Header.VisiblePosition = 99;
			ultraGridColumn39.Hidden = true;
			ultraGridColumn40.Header.VisiblePosition = 44;
			ultraGridColumn40.Hidden = true;
			ultraGridColumn41.Header.VisiblePosition = 84;
			ultraGridColumn41.Hidden = true;
			ultraGridColumn42.Header.VisiblePosition = 60;
			ultraGridColumn42.Hidden = true;
			ultraGridColumn43.Header.VisiblePosition = 61;
			ultraGridColumn43.Hidden = true;
			ultraGridColumn44.Header.VisiblePosition = 58;
			ultraGridColumn44.Hidden = true;
			ultraGridColumn45.Header.VisiblePosition = 59;
			ultraGridColumn45.Hidden = true;
			ultraGridColumn46.Header.VisiblePosition = 100;
			ultraGridColumn46.Hidden = true;
			ultraGridColumn47.Header.VisiblePosition = 95;
			ultraGridColumn47.Hidden = true;
			ultraGridColumn48.Header.VisiblePosition = 97;
			ultraGridColumn48.Hidden = true;
			ultraGridColumn49.Header.VisiblePosition = 17;
			ultraGridColumn49.Hidden = true;
			ultraGridColumn50.Header.VisiblePosition = 35;
			ultraGridColumn50.Hidden = true;
			ultraGridColumn51.Header.VisiblePosition = 40;
			ultraGridColumn51.Hidden = true;
			ultraGridColumn52.Header.VisiblePosition = 85;
			ultraGridColumn52.Hidden = true;
			ultraGridColumn53.Header.VisiblePosition = 24;
			ultraGridColumn53.Hidden = true;
			ultraGridColumn54.Header.VisiblePosition = 7;
			ultraGridColumn54.Hidden = true;
			ultraGridColumn55.Header.VisiblePosition = 0;
			ultraGridColumn55.Hidden = true;
			ultraGridColumn56.Header.VisiblePosition = 12;
			ultraGridColumn56.Hidden = true;
			ultraGridColumn56.Width = 28;
			ultraGridColumn57.Header.VisiblePosition = 13;
			ultraGridColumn57.Hidden = true;
			ultraGridColumn58.Header.VisiblePosition = 2;
			ultraGridColumn58.Hidden = true;
			ultraGridColumn59.Header.VisiblePosition = 36;
			ultraGridColumn59.Hidden = true;
			ultraGridColumn60.Header.VisiblePosition = 37;
			ultraGridColumn60.Hidden = true;
			ultraGridColumn61.Header.VisiblePosition = 38;
			ultraGridColumn61.Hidden = true;
			ultraGridColumn62.Header.VisiblePosition = 6;
			ultraGridColumn62.Hidden = true;
			ultraGridColumn62.Width = 84;
			ultraGridColumn63.Header.VisiblePosition = 8;
			ultraGridColumn63.Hidden = true;
			ultraGridColumn64.Header.VisiblePosition = 4;
			ultraGridColumn64.Hidden = true;
			ultraGridColumn65.Header.VisiblePosition = 16;
			ultraGridColumn65.Hidden = true;
			ultraGridColumn66.Header.VisiblePosition = 11;
			ultraGridColumn66.Hidden = true;
			ultraGridColumn67.Header.VisiblePosition = 14;
			ultraGridColumn67.Hidden = true;
			ultraGridColumn68.Header.VisiblePosition = 5;
			ultraGridColumn68.Hidden = true;
			ultraGridColumn69.Header.VisiblePosition = 15;
			ultraGridColumn69.Hidden = true;
			ultraGridColumn70.Header.VisiblePosition = 9;
			ultraGridColumn70.Hidden = true;
			ultraGridColumn70.Width = 73;
			ultraGridColumn71.Header.VisiblePosition = 10;
			ultraGridColumn71.Hidden = true;
			ultraGridColumn72.Header.VisiblePosition = 45;
			ultraGridColumn72.Hidden = true;
			ultraGridColumn73.Header.VisiblePosition = 46;
			ultraGridColumn73.Hidden = true;
			ultraGridColumn74.Header.VisiblePosition = 47;
			ultraGridColumn74.Hidden = true;
			ultraGridColumn75.Header.VisiblePosition = 48;
			ultraGridColumn75.Hidden = true;
			ultraGridColumn76.Header.VisiblePosition = 49;
			ultraGridColumn76.Hidden = true;
			ultraGridColumn77.Header.VisiblePosition = 50;
			ultraGridColumn77.Hidden = true;
			ultraGridColumn78.Header.VisiblePosition = 23;
			ultraGridColumn78.Hidden = true;
			ultraGridColumn79.Header.VisiblePosition = 82;
			ultraGridColumn79.Hidden = true;
			ultraGridColumn80.Header.VisiblePosition = 96;
			ultraGridColumn80.Hidden = true;
			ultraGridColumn81.Header.VisiblePosition = 52;
			ultraGridColumn81.Hidden = true;
			ultraGridColumn82.Header.VisiblePosition = 94;
			ultraGridColumn82.Hidden = true;
			ultraGridColumn83.Header.VisiblePosition = 20;
			ultraGridColumn83.Hidden = true;
			ultraGridColumn84.Header.VisiblePosition = 25;
			ultraGridColumn84.Hidden = true;
			ultraGridColumn85.Header.VisiblePosition = 26;
			ultraGridColumn85.Hidden = true;
			ultraGridColumn86.Header.VisiblePosition = 27;
			ultraGridColumn86.Hidden = true;
			ultraGridColumn87.Header.VisiblePosition = 28;
			ultraGridColumn87.Hidden = true;
			ultraGridColumn88.Header.VisiblePosition = 29;
			ultraGridColumn88.Hidden = true;
			ultraGridColumn89.Header.VisiblePosition = 30;
			ultraGridColumn89.Hidden = true;
			ultraGridColumn90.Header.VisiblePosition = 56;
			ultraGridColumn90.Hidden = true;
			ultraGridColumn91.Header.VisiblePosition = 33;
			ultraGridColumn91.Hidden = true;
			ultraGridColumn92.Header.VisiblePosition = 65;
			ultraGridColumn92.Hidden = true;
			ultraGridColumn93.Header.VisiblePosition = 19;
			ultraGridColumn93.Hidden = true;
			ultraGridColumn94.Header.VisiblePosition = 79;
			ultraGridColumn94.Hidden = true;
			ultraGridColumn95.Header.VisiblePosition = 88;
			ultraGridColumn95.Hidden = true;
			ultraGridColumn96.Header.VisiblePosition = 42;
			ultraGridColumn96.Hidden = true;
			ultraGridColumn97.Header.VisiblePosition = 64;
			ultraGridColumn97.Hidden = true;
			ultraGridColumn98.Header.VisiblePosition = 51;
			ultraGridColumn98.Hidden = true;
			ultraGridColumn99.Header.VisiblePosition = 43;
			ultraGridColumn99.Hidden = true;
			ultraGridColumn100.Header.VisiblePosition = 86;
			ultraGridColumn100.Hidden = true;
			ultraGridColumn101.Header.VisiblePosition = 101;
			ultraGridColumn101.Hidden = true;
			ultraGridColumn102.Header.VisiblePosition = 102;
			ultraGridColumn102.Hidden = true;
			ultraGridColumn103.Header.VisiblePosition = 103;
			ultraGridColumn104.Header.VisiblePosition = 3;
			ultraGridColumn104.Hidden = true;
			ultraGridColumn104.Width = 89;
			ultraGridColumn105.Header.VisiblePosition = 104;
			ultraGridColumn105.Hidden = true;
			ultraGridColumn106.Header.VisiblePosition = 105;
			ultraGridColumn106.Hidden = true;
			ultraGridColumn107.Header.VisiblePosition = 106;
			ultraGridColumn107.Hidden = true;
			ultraGridColumn108.Header.VisiblePosition = 107;
			ultraGridColumn108.Hidden = true;
			ultraGridColumn109.Header.VisiblePosition = 108;
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
			this.cmbArt.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.cmbArt.DisplayMember = "Articulo";
			this.cmbArt.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbArt.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbArt.Enabled = false;
			this.cmbArt.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbArt.Location = new System.Drawing.Point(464, 103);
			this.cmbArt.Name = "cmbArt";
			this.cmbArt.Size = new System.Drawing.Size(128, 21);
			this.cmbArt.TabIndex = 35;
			this.cmbArt.ValueMember = "idArticulo";
			this.cmbArt.Visible = false;
			// 
			// chkArt
			// 
			this.chkArt.BackColor = System.Drawing.Color.Transparent;
			this.chkArt.CheckAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.chkArt.Checked = true;
			this.chkArt.CheckState = System.Windows.Forms.CheckState.Checked;
			this.chkArt.Location = new System.Drawing.Point(464, 78);
			this.chkArt.Name = "chkArt";
			this.chkArt.Size = new System.Drawing.Size(104, 21);
			this.chkArt.TabIndex = 39;
			this.chkArt.Text = "Artículo";
			this.chkArt.Visible = false;
			this.chkArt.CheckedChanged += new System.EventHandler(this.chkArt_CheckedChanged);
			// 
			// btSaldoCostoLotes
			// 
			this.btSaldoCostoLotes.Location = new System.Drawing.Point(56, 129);
			this.btSaldoCostoLotes.Name = "btSaldoCostoLotes";
			this.btSaldoCostoLotes.Size = new System.Drawing.Size(96, 25);
			this.btSaldoCostoLotes.TabIndex = 40;
			this.btSaldoCostoLotes.Text = "Costo lotes";
			this.btSaldoCostoLotes.Click += new System.EventHandler(this.btSaldoCostoLotes_Click);
			// 
			// ActualizaSaldo
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 14);
			this.ClientSize = new System.Drawing.Size(600, 289);
			this.Controls.Add(this.chkArt);
			this.Controls.Add(this.cmbArt);
			this.Controls.Add(this.btnFIFO);
			this.Controls.Add(this.btEnvioCostoB);
			this.Controls.Add(this.btSaldoCostoB);
			this.Controls.Add(this.btIngreso);
			this.Controls.Add(this.label3);
			this.Controls.Add(this.label2);
			this.Controls.Add(this.cmbHasta);
			this.Controls.Add(this.simpleButton1);
			this.Controls.Add(this.btActEgreso);
			this.Controls.Add(this.label1);
			this.Controls.Add(this.btSigno);
			this.Controls.Add(this.btSaldoCosto);
			this.Controls.Add(this.btSaldoCxC);
			this.Controls.Add(this.cmbDesde);
			this.Controls.Add(this.btGrupoArt);
			this.Controls.Add(this.btBorraBlanco);
			this.Controls.Add(this.btCancelar);
			this.Controls.Add(this.btAceptar);
			this.Controls.Add(this.btCostos);
			this.Controls.Add(this.btSaldos);
			this.Controls.Add(this.btSaldoCostoLotes);
			this.helpProvider1.SetHelpKeyword(this, "Actualizaldosp.htm");
			this.helpProvider1.SetHelpNavigator(this, System.Windows.Forms.HelpNavigator.Topic);
			this.Name = "ActualizaSaldo";
			this.helpProvider1.SetShowHelp(this, true);
			this.Text = "Actualiza Saldo";
			this.Closing += new System.ComponentModel.CancelEventHandler(this.ActualizaSaldo_Closing);
			this.Load += new System.EventHandler(this.ActualizaSaldo_Load);
			((System.ComponentModel.ISupportInitialize)(this.cdsCompra)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsArticulo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsGrupo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbDesde)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbHasta)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvArticulo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.c1DataSet1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbArt)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void btAceptar_Click(object sender, System.EventArgs e)
		{
			SqlCommand cmdActualiza;
			this.cdsCompra.Schema.Connections[0].Open();
			C1.Data.SchemaObjects.Connection miConeccion = 
				this.cdsCompra.Schema.Connections[0];

			// Deshabilita DetCompra
			miConeccion = this.cdsCompra.Schema.Connections[0];
			string stSelect = "ALTER TABLE DetCompra DISABLE TRIGGER SaldoCosto";
			cmdActualiza = new SqlCommand (stSelect
				, (SqlConnection) miConeccion.DbConnection);
			cmdActualiza.Transaction = (SqlTransaction)
				miConeccion.DbTransaction;
			cmdActualiza.ExecuteNonQuery();
			this.cdsCompra.Schema.Connections[0].Close();		

			this.cdsArticulo.Update();
			this.cdsCompra.Update();
			this.Close();
			MessageBox.Show("Datos Grabados", "Información");
		}

		private void btCancelar_Click(object sender, System.EventArgs e)
		{
			this.Close();
		}

		private void btSaldos_Click(object sender, System.EventArgs e)
		{
			double fSaldo = 0;
			this.cdsCompra.Schema.Connections[0].Open();
			SqlDataReader miReader;
			C1.Data.SchemaObjects.Connection miConeccion = cdsCompra.Schema.Connections[0];
			foreach(C1.Data.C1DataRow drArt in this.cdsArticulo.TableViews["Articulo"].Rows)
			{
				SqlCommand cmdCant = new SqlCommand (
					"SELECT Cantidad, Signo FROM DetCompra WHERE idArticulo = " +
					drArt["idArticulo"].ToString() + " AND (Signo = 1 Or Signo = -1) AND Borrar = 0"
					, (SqlConnection) miConeccion.DbConnection);
				miReader = cmdCant.ExecuteReader();
				try
				{
					while (miReader.Read())
					{
						int i=0;
						int iSigno =  miReader.GetInt32(1);
						if (iSigno == 1 || iSigno == 2)	i = 1;
						if (iSigno == -1 || iSigno == -2)	i = -1;
						fSaldo += miReader.GetDouble(0) * i;
					}
					drArt["Saldo"] = fSaldo;
					fSaldo = 0;
				}
				catch
				{
//          drArt["Saldo"] = 0;
				}
				miReader.Close();
			}
			this.cdsCompra.Schema.Connections[0].Close();
			MessageBox.Show("Saldos actualizados");
		}

		private void btCostos_Click(object sender, System.EventArgs e)
		{
			this.Cursor = Cursors.WaitCursor;
			DateTime dtHoraIni = DateTime.Now;
			DateTime dtFecha = (DateTime) this.cmbDesde.Value;
			double fSaldo = 0;
			double fCostoP = 0;
			double fPrecio = 0;
			double fCantidad = 0;
			double fDescuento = 0;
			double fFlete = 0;
			int iSigno = 0;
			int IdDetCompra = 0;
			int IdTipoFactura = 0;
			this.cdsCompra.Schema.Connections[0].Open();
			SqlDataReader miReader;
			C1.Data.SchemaObjects.Connection miConeccion = this.cdsCompra.Schema.Connections[0];
			foreach(C1.Data.C1DataRow drArt in this.cdsArticulo.TableViews["Articulo"].Rows)
			{
				string stSelect = 
					"SELECT DetCompra.Cantidad, DetCompra.Precio, DetCompra.Signo, "+
					"DetCompra.DescuentoArt, DetCompra.Flete, DetCompra.idDetCompra, "+
					"Compra.idTipoFactura FROM Compra INNER JOIN DetCompra ON "+
					"Compra.idCompra = DetCompra.idCompra "+
					"WHERE (((DetCompra.Signo)=1 Or (DetCompra.Signo)=-1)) " +
					"AND DetCompra.Borrar = 0 " +
					"AND Compra.Fecha <= '" + dtFecha.ToString("yyyyMMdd")+
					"' AND idArticulo = " + drArt["idArticulo"].ToString() + 
					" ORDER BY DetCompra.idArticulo, Compra.Fecha, Compra.idCompra;";
				SqlCommand cmdCant = new SqlCommand (stSelect
					, (SqlConnection) miConeccion.DbConnection);

			miReader = cmdCant.ExecuteReader();
				try
				{
					while (miReader.Read())
					{
						fCantidad = miReader.GetDouble(0);
						fPrecio =  miReader.GetDouble(1);
						iSigno =  miReader.GetInt32(2);
						fDescuento = miReader.GetDouble(3);
						fFlete = miReader.GetDouble(4);
						IdDetCompra =  miReader.GetInt32(5);
						IdTipoFactura =  miReader.GetInt32(6);

						// Calculo del Costo Promedio
						if (iSigno == 1 && IdTipoFactura != 5) // Es necesario excluir la devolucion en venta
						{
							if (fSaldo > 0)
							{
								fCostoP = (fCostoP * fSaldo + fCantidad * fPrecio
									+ fFlete - fDescuento) / (fSaldo + fCantidad);
							}
							else
								if (fCantidad > 0)
									fCostoP = fPrecio + (fFlete - fDescuento) / fCantidad;
								else
									fCostoP = fPrecio + (fFlete - fDescuento);
						}
						else // Ingresos no afectan costo promedio
						{
							if (fSaldo <= 0 && IdTipoFactura != 5)
								fCostoP = 0; // Hay un error con devolucion en venta
						}
						// Fin de Calculo de Costo Promedio
						C1.Data.C1DataRow drDet = 
							this.cdsCompra.TableViews["DetCompra"].Rows.Find(IdDetCompra);
						drDet["CostoPromedio"] = fCostoP;
						// Fin de asignacion de costo

						fSaldo += fCantidad * iSigno;
						drDet["Saldo"] = fSaldo;
					}
					// Termina lazo de Detalle de Compra con Articulo
					drArt["CostoPromedio"] = fCostoP;
					fSaldo = 0;
					fCostoP = 0;
				}
				catch(Exception ex)
				{
					MessageBox.Show(ex.Message.ToString());
					return;
				}
				miReader.Close();
			}
			this.cdsCompra.Schema.Connections[0].Close();

//			this.cdsArticulo.Update();
//			this.cdsCompra.Update();
			this.Cursor = Cursors.Default;
			DateTime dtHoraFin = DateTime.Now;

			TimeSpan ts = dtHoraFin.Subtract(dtHoraIni);
			MessageBox.Show("Costo Promedio generado en " +
				ts.Minutes + ":" + ts.Seconds.ToString());
		}

		private void btBorraBlanco_Click(object sender, System.EventArgs e)
		{
			SqlCommand cmdActualiza;
			this.cdsCompra.Schema.Connections[0].Open();
			C1.Data.SchemaObjects.Connection miConeccion = 
				this.cdsCompra.Schema.Connections[0];

			// Borra blancos de Asiento
			miConeccion = this.cdsCompra.Schema.Connections[0];
			string stSelect = "UPDATE Asiento Set Descripcion = RTrim(Descripcion), " +
				"Numero = RTrim(Numero), CodAsiento = RTrim(CodAsiento), " + 
				"Cheque = RTrim(Cheque), Usuario = RTrim(Usuario)";
//			Notas = RTrim(Notas),
			cmdActualiza = new SqlCommand (stSelect
				, (SqlConnection) miConeccion.DbConnection);
			cmdActualiza.Transaction = (SqlTransaction)
				miConeccion.DbTransaction;
			cmdActualiza.ExecuteNonQuery();
			this.cdsCompra.Schema.Connections[0].Close();

			// Borra blancos de Detalle de asientos
			this.cdsCompra.Schema.Connections[0].Open();
			miConeccion = this.cdsCompra.Schema.Connections[0];
			cmdActualiza = new SqlCommand (
				"UPDATE AsientoDetalle Set Detalle = RTrim(Detalle), " +
				"DetIngles = RTrim(DetIngles)" 
				, (SqlConnection) miConeccion.DbConnection);
			cmdActualiza.Transaction = (SqlTransaction)
				miConeccion.DbTransaction;
			cmdActualiza.ExecuteNonQuery();
			this.cdsCompra.Schema.Connections[0].Close();

			// Borra blancos de Cuentas
			this.cdsCompra.Schema.Connections[0].Open();
			miConeccion = this.cdsCompra.Schema.Connections[0];
			cmdActualiza = new SqlCommand (
				"UPDATE Cuenta Set Descripcion = RTrim(Descripcion), " +
				"CodRapido = RTrim(CodRapido), Codigo = RTrim(Codigo), " +
				"CodPresup = RTrim(CodPresup), DescIngles = RTrim(DescIngles)"
				, (SqlConnection) miConeccion.DbConnection);
			cmdActualiza.Transaction = (SqlTransaction)
				miConeccion.DbTransaction;
			cmdActualiza.ExecuteNonQuery();
			this.cdsCompra.Schema.Connections[0].Close();

			// Borra blancos de Articulo
			this.cdsCompra.Schema.Connections[0].Open();
			miConeccion = this.cdsCompra.Schema.Connections[0];
			cmdActualiza = new SqlCommand (
				"UPDATE Articulo Set Articulo = RTrim(Articulo), " +
				"CodAlterno = RTrim(CodAlterno), Codigo = RTrim(Codigo), " +
				"Caracteristica = RTrim(Caracteristica), " +
				"Aplicacion = RTrim(Aplicacion), UnidadPeso = RTrim(UnidadPeso)"
				, (SqlConnection) miConeccion.DbConnection);
			cmdActualiza.Transaction = (SqlTransaction)
				miConeccion.DbTransaction;
			cmdActualiza.ExecuteNonQuery();
			this.cdsCompra.Schema.Connections[0].Close();

			// Borra blancos de Grupo de Articulo
			this.cdsCompra.Schema.Connections[0].Open();
			miConeccion = this.cdsCompra.Schema.Connections[0];
			cmdActualiza = new SqlCommand (
				"UPDATE ArticuloGrupo Set Grupo = RTrim(Grupo), " +
				"Codigo = RTrim(Codigo)" 
				, (SqlConnection) miConeccion.DbConnection);
			cmdActualiza.Transaction = (SqlTransaction)
				miConeccion.DbTransaction;
			cmdActualiza.ExecuteNonQuery();
			this.cdsCompra.Schema.Connections[0].Close();

			// Borra blancos de Unidad de Articulo
			this.cdsCompra.Schema.Connections[0].Open();
			miConeccion = this.cdsCompra.Schema.Connections[0];
			cmdActualiza = new SqlCommand (
				"UPDATE ArticuloUnidad Set Unidades = RTrim(Unidades), " +
				"Codigo = RTrim(Codigo)" 
				, (SqlConnection) miConeccion.DbConnection);
			cmdActualiza.Transaction = (SqlTransaction)
				miConeccion.DbTransaction;
			cmdActualiza.ExecuteNonQuery();
			this.cdsCompra.Schema.Connections[0].Close();
			
			// Borra blancos de Clientes
			this.cdsCompra.Schema.Connections[0].Open();
			miConeccion = this.cdsCompra.Schema.Connections[0];
			cmdActualiza = new SqlCommand (
				"UPDATE Cliente Set Nombre = RTrim(Nombre), " +
				"Codigo = RTrim(Codigo), Direccion = RTrim(Direccion), " +
				"Telefono = RTrim(Telefono), Fax= RTrim(Fax), Ruc = RTrim(Ruc), " +
				"Observacion = RTrim(Observacion), Representante = RTrim(Representante), " +
				"EMail = RTrim(EMail), Celular = RTrim(Celular), " +
				"Contacto = RTrim(Contacto), Foto = RTrim(Foto), " + 
				"Pais = RTrim(Pais), " +
				"AutorizaFactura = RTrim(AutorizaFactura), " +
				"AutorizaImprenta = RTrim(AutorizaImprenta), " +
				"SerieFactura = RTrim(SerieFactura)"
				, (SqlConnection) miConeccion.DbConnection);
//			Notas = RTrim(Notas)
			cmdActualiza.Transaction = (SqlTransaction)
				miConeccion.DbTransaction;
			cmdActualiza.ExecuteNonQuery();
			this.cdsCompra.Schema.Connections[0].Close();

			// Borra blancos de Compra
			this.cdsCompra.Schema.Connections[0].Open();
			miConeccion = this.cdsCompra.Schema.Connections[0];
			cmdActualiza = new SqlCommand (
				"UPDATE Compra Set Usuario = RTrim(Usuario), " +
				"Numero = RTrim(Numero), " +
				"NumCheque = RTrim(NumCheque), NumCuenta= RTrim(NumCuenta), " +
				"Otro = RTrim(Otro), Refrendo = RTrim(Refrendo), " +
				"Comprobante = RTrim(Comprobante), AutImprenta = RTrim(AutImprenta), " +
				"AutFactura = RTrim(AutFactura), SerieFactura = RTrim(SerieFactura), " + 
				"Departamento = RTrim(Departamento)"
				, (SqlConnection) miConeccion.DbConnection);
//			Notas = RTrim(Notas), 
			cmdActualiza.Transaction = (SqlTransaction)
				miConeccion.DbTransaction;
			cmdActualiza.ExecuteNonQuery();
			this.cdsCompra.Schema.Connections[0].Close();

			// Borra blancos de DetCompra
			this.cdsCompra.Schema.Connections[0].Open();
			miConeccion = this.cdsCompra.Schema.Connections[0];
			cmdActualiza = new SqlCommand (
				"UPDATE DetCompra Set Notas = RTrim(Notas), " +
				"RefCodigo = RTrim(RefCodigo), RefNumero = RTrim(RefNumero)" 
				, (SqlConnection) miConeccion.DbConnection);
			cmdActualiza.Transaction = (SqlTransaction)
				miConeccion.DbTransaction;
			cmdActualiza.ExecuteNonQuery();
			this.cdsCompra.Schema.Connections[0].Close();

			// Borra blancos de Pago
			this.cdsCompra.Schema.Connections[0].Open();
			miConeccion = this.cdsCompra.Schema.Connections[0];
			cmdActualiza = new SqlCommand (
				"UPDATE Pago Set Concepto = RTrim(Concepto), " +
				"NumeroIngreso = RTrim(NumeroIngreso), Cheque = RTrim(Cheque), " +
				"Cuenta = RTrim(Cuenta), Nombre= RTrim(Nombre), " +
				"Numero = RTrim(Numero) "
				, (SqlConnection) miConeccion.DbConnection);
//			, Notas = RTrim(Notas)
			cmdActualiza.Transaction = (SqlTransaction)
				miConeccion.DbTransaction;
			cmdActualiza.ExecuteNonQuery();
			this.cdsCompra.Schema.Connections[0].Close();

			// Borra blancos de Retencion
			this.cdsCompra.Schema.Connections[0].Open();
			miConeccion = this.cdsCompra.Schema.Connections[0];
			cmdActualiza = new SqlCommand (
				"UPDATE Retencion Set CodigoSri = RTrim(CodigoSri), " +
				" Cheque = RTrim(Cheque), " +
				"Banco = RTrim(Banco), NumFactura= RTrim(NumFactura), " +
				"Nombre = RTrim(Nombre), Codigo = RTrim(Codigo), " + 
				"CodAsiento = RTrim(CodAsiento), Beneficiario = RTrim(Beneficiario)"
				, (SqlConnection) miConeccion.DbConnection);
//			Notas = RTrim(Notas),
			cmdActualiza.Transaction = (SqlTransaction)
				miConeccion.DbTransaction;
			cmdActualiza.ExecuteNonQuery();
			this.cdsCompra.Schema.Connections[0].Close();


			MessageBox.Show("Finalizado");
		}

		private void btGrupoArt_Click(object sender, System.EventArgs e)
		{
			SqlCommand cmdActualiza;
			this.cdsCompra.Schema.Connections[0].Open();
			C1.Data.SchemaObjects.Connection miConeccion = this.cdsCompra.Schema.Connections[0];
			foreach(C1.Data.C1DataRow drGrupo in cdsGrupo.TableViews["ArticuloGrupo"].Rows)
			{
				cmdActualiza = new SqlCommand (
					"UPDATE Articulo SET idTipoGrupo = " +
					drGrupo["idTipoGrupo"].ToString() +
					" WHERE idGrupoArticulo = " +
					drGrupo["idGrupoArticulo"].ToString()
					, (SqlConnection) miConeccion.DbConnection);

				cmdActualiza.Transaction = (SqlTransaction)
					miConeccion.DbTransaction;
				cmdActualiza.ExecuteNonQuery();
			}
			this.cdsCompra.Schema.Connections[0].Close();
			MessageBox.Show("Grupo de articulos actualizado");
		}

		private void cdsCompra_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			this.cdsCompra.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);
		}

		private void cdsArticulo_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsArticulo.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);		
		}

		private void cdsGrupo_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			this.cdsGrupo.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);
		}

		private void ActualizaSaldo_Load(object sender, System.EventArgs e)
		{
			if (Funcion.bEjecutaSQL(cdsArticulo, "Exec SeteoGExiste 'ACTSFIFO'")) 				
			{
				cdvArticulo.RowFilter = "idTipoGrupo in(1,6)";
				btnFIFO.Visible = true;
				chkArt.Visible = true;
				cmbArt.Visible=true;
			}

			Cursor = Cursors.Default;
			string stAudita = "Exec AuditaCrear 53, 6, 'Act. Saldo'";
			Funcion.EjecutaSQL(cdsArticulo, stAudita, true);
			cmbHasta.Value = DateTime.Today;
			btSaldoCostoB.Visible = false;
			btEnvioCostoB.Visible = false;

			string stSelect = "Select Activado from SeteoG where codigo = 'I' ";
			bool bSelect = Funcion.bEjecutaSQL(cdsCompra, stSelect);
			if (bSelect== true) 
			{
				btSaldoCostoB.Visible = true;
				btEnvioCostoB.Visible = true;
			}

			if (Funcion.bEjecutaSQL(cdsArticulo, "Exec SeteoGExiste 'FPCLOT'")) 				
			{
				btSaldoCosto.Visible = false;
				btSaldoCostoLotes.Visible= true;
			}
		}

		private void btSaldoCxC_Click(object sender, System.EventArgs e)
		{
			this.Cursor = Cursors.WaitCursor;
			string stSelect = "EXEC SaldoCxCyP";
			Funcion.EjecutaSQL(cdsCompra, stSelect);
			this.Cursor = Cursors.Default;
			MessageBox.Show("Saldos de CxC Actualizados");
		}

		private void ActualizaSaldo_Closing(object sender, System.ComponentModel.CancelEventArgs e)
		{
			string stActiva = "Exec TriggerActiva";
			Funcion.EjecutaSQL(cdsArticulo, stActiva, true);
			string stAudita = "Exec AuditaCrear 53, 7, 'Act. Saldo'";
			Funcion.EjecutaSQL(cdsArticulo, stAudita, true);
		}

		private void btSaldoCosto_Click(object sender, System.EventArgs e)
		{
			DateTime dtDesde = (DateTime) cmbDesde.Value;
			DateTime dtHasta = (DateTime) cmbHasta.Value;
			if (DialogResult.No == MessageBox.Show("Desea Actualizar costos desde el " 
				+ dtDesde.ToString("dd/MMM/yyyy")
				+ " hasta el " + dtHasta.ToString("dd/MMM/yyyy"), "Confirmación", 
				MessageBoxButtons.YesNo, MessageBoxIcon.Question)) return;

			this.Cursor = Cursors.WaitCursor;
			DateTime dtHoraIni = DateTime.Now;
			string stSelect = "EXEC ActualizaTodo '" + dtDesde.ToString("yyyyMMdd")
				+ "', '" + dtHasta.AddDays(1).ToString("yyyyMMdd") + "'";
			Funcion.EjecutaSQL(cdsArticulo, stSelect);
			this.Cursor = Cursors.Default;
			DateTime dtHoraFin = DateTime.Now;

			TimeSpan ts = dtHoraFin.Subtract(dtHoraIni);
			MessageBox.Show("Costos y Saldos generados en " +
				ts.Minutes + ":" + ts.Seconds.ToString());
		}

		private void btSigno_Click(object sender, System.EventArgs e)
		{
			if (MessageBox.Show("¿Desea Regenerar Signos?",
				"Confirmación", MessageBoxButtons.YesNo, MessageBoxIcon.Information)
				== DialogResult.No) return;
			this.Cursor = Cursors.WaitCursor;
			DateTime dtHoraIni = DateTime.Now;

			string stSelect = "Update DetCompra Set DetCompra.Signo = CompraNumero.Signo "
				+ " From DetCompra "
				+ " Inner Join Compra "
				+ " On DetCompra.idCompra = Compra.IdCompra "
				+ " inner join CompraNumero "
				+ " On Compra.idTipoFactura = CompraNumero.idTipoFactura "
				+ " Where DetCompra.Signo <> CompraNumero.Signo ";
			Funcion.EjecutaSQL(cdsArticulo, stSelect);
			this.Cursor = Cursors.WaitCursor;
			MessageBox.Show("Terminado", "Información",
				MessageBoxButtons.OK, MessageBoxIcon.Information);
		}

		private void btSaldos_MouseDown(object sender, System.Windows.Forms.MouseEventArgs e)
		{
			if (e.Button != MouseButtons.Right) return;
			string stSelect = "Select COUNT(DetCompra.idDetCompra) "
				+ "from Compra inner join CompraNumero "
				+ "On Compra.idTipoFactura = CompraNumero.idTipoFactura "
				+ "Inner Join DetCompra "
				+ "On DetCompra.idCompra = Compra.IdCompra "
				+ "Where DetCompra.Signo <> CompraNumero.Signo ";
			int iCant = Funcion.iEscalarSQL(cdsCompra, stSelect);
			MessageBox.Show("Número de Registros con Errores en Signo: " 
				+ iCant.ToString(), "Información", 
				MessageBoxButtons.OK, MessageBoxIcon.Information);
		}

		private void btActEgreso_Click(object sender, System.EventArgs e)
		{
			if (DialogResult.No == MessageBox.Show(
				"Desea Actualizar precios de egresos con el costo promedio",
				"Información", MessageBoxButtons.YesNo, MessageBoxIcon.Question)) return;
			DateTime dtFecha = (DateTime) cmbDesde.Value;
			DateTime dtHasta = (DateTime) cmbHasta.Value;
			DateTime dtFecha2 = new DateTime(2009,05,01);
			string stSelect = "Select Activado from SeteoG where codigo = 'I' ";
			bool bSelect = Funcion.bEjecutaSQL(cdsCompra, stSelect);
			if (bSelect == true)
				if (dtFecha < dtFecha2) dtFecha = dtFecha2;

			stSelect = "Update DetCompra Set Precio = CostoPromedio "
				+ "From DetCompra Inner Join Compra "
				+ "On DetCompra.idCompra = Compra.idCompra "
				+ "Where Compra.idTipoFactura = 8 And Compra.Fecha >= '"
				+ dtFecha.ToString("yyyyMMdd")+ "' And Compra.Fecha < '"
				+ dtHasta.AddDays(1).ToString("yyyyMMdd") + "' ";
			Funcion.EjecutaSQL(cdsCompra, stSelect, true);
			string stexec = "Exec FacturaTotalesEgr '" 
				+ dtFecha.ToString("yyyyMMdd")+ "'";
			Funcion.EjecutaSQL(cdsCompra, stexec, true);
			MessageBox.Show("Egresos Finalizado", "Información");

			stSelect = "Select Activado from SeteoG where codigo = 'I' ";
			bSelect = Funcion.bEjecutaSQL(cdsCompra, stSelect);
			if (bSelect == true)
			{
				DateTime dtDesde = (DateTime) cmbDesde.Value;
				dtHasta = (DateTime) cmbHasta.Value;
				if (DialogResult.No == MessageBox.Show("Desea Actualizar costos en ingresos de bodega desde el " 
					+ dtDesde.ToString("dd/MMM/yyyy")
					+ " hasta el " + dtHasta.ToString("dd/MMM/yyyy"), "Confirmaciónn", 
					MessageBoxButtons.YesNo, MessageBoxIcon.Question)) return;

				this.Cursor = Cursors.WaitCursor;
				DateTime dtHoraIni = DateTime.Now;
				stSelect = "EXEC ActualizaIngresosB '" + dtDesde.ToString("yyyyMMdd")
					+ "', '" + dtHasta.AddDays(1).ToString("yyyyMMdd") + "'";
				Funcion.EjecutaSQL(cdsArticulo, stSelect);
				this.Cursor = Cursors.Default;
				DateTime dtHoraFin = DateTime.Now;

				TimeSpan ts = dtHoraFin.Subtract(dtHoraIni);
				MessageBox.Show("Ingresos de Bodega Relacionados Actualizados en " +
					ts.Minutes + ":" + ts.Seconds.ToString());
			}






		}

		private void btSaldoCosto_MouseDown(object sender, System.Windows.Forms.MouseEventArgs e)
		{
			if (e.Button != MouseButtons.Right) return;
			this.Cursor = Cursors.WaitCursor;
			DateTime dtHoraIni = DateTime.Now;
			Funcion.EjecutaSQL(cdsArticulo, "EXEC SaldoBodega");
			this.Cursor = Cursors.Default;
			DateTime dtHoraFin = DateTime.Now;

			TimeSpan ts = dtHoraFin.Subtract(dtHoraIni);
			MessageBox.Show("Saldos de Bodega generados en " +
				ts.Minutes + ":" + ts.Seconds.ToString());
		}

		private void btSaldoCxC_MouseUp(object sender, System.Windows.Forms.MouseEventArgs e)
		{
			if (e.Button == MouseButtons.Right)
			{
				string stBase = MenuLatinium.stDirConta;
				string stParametro = "/U infoelect /P Bl45o6$9 /d ";
				stParametro += stBase;
				stParametro += " /S " + MenuLatinium.stDirServidor;
				stParametro += @" /n /t 1000 /i " + MenuLatinium.stDirInicio + @"\Actualiza.Sql /o "
				+ MenuLatinium.stDirInicio + @"\Salida.txt";

				System.Diagnostics.Process.Start("osql", @stParametro);

				if (DialogResult.Yes == MessageBox.Show(
					"¿Desea Ver Registro de Actividades?", 
					"Información", MessageBoxButtons.YesNo, MessageBoxIcon.Question))
					System.Diagnostics.Process.Start("NotePad.exe", MenuLatinium.stDirInicio + @"\Salida.txt");

			}
		}

		private void simpleButton1_Click(object sender, System.EventArgs e)
		{
			string stBase = MenuLatinium.stDirConta;
			string stParametro = "/U infoelect /P Bl45o6$9 /d ";
			stParametro += stBase;
			stParametro += " /S " + MenuLatinium.stDirServidor;
			stParametro += @" /n /t 1000 /i " + MenuLatinium.stDirInicio + @"\ActualizaGS.Sql /o "
			+ MenuLatinium.stDirInicio + @"\Salida.txt";

			System.Diagnostics.Process.Start("osql", @stParametro);

			if (DialogResult.Yes == MessageBox.Show(
				"¿Desea Ver Registro de Actividades?", 
				"Información", MessageBoxButtons.YesNo, MessageBoxIcon.Question))
				System.Diagnostics.Process.Start("NotePad.exe", MenuLatinium.stDirInicio + @"\Salida.txt");
		}

		private void btIngreso_Click(object sender, System.EventArgs e)
		{
			if (DialogResult.No == MessageBox.Show(
				"Desea Actualizar precios de Ingresos que tienen comprobante con el costo promedio",
				"Información", MessageBoxButtons.YesNo, MessageBoxIcon.Question)) return;
			DateTime dtFecha = (DateTime) cmbDesde.Value;
			DateTime dtHasta = (DateTime) cmbHasta.Value;
			DateTime dtFecha2 = new DateTime(2009,05,01);
			string stSelect = "Select Activado from SeteoG where codigo = 'I' ";
			bool bSelect = Funcion.bEjecutaSQL(cdsCompra, stSelect);
			if (bSelect == true)
				if (dtFecha < dtFecha2) dtFecha = dtFecha2;

			stSelect = "Update DetCompra Set Precio = CostoPromedio "
				+ "From DetCompra Inner Join Compra "
				+ "On DetCompra.idCompra = Compra.idCompra "
				+ "Where Compra.idTipoFactura = 9 And Compra.Fecha >= '"
				+ dtFecha.ToString("yyyyMMdd") + "' And Compra.Fecha < '"
				+ dtHasta.AddDays(1).ToString("yyyyMMdd") + "' "
				+ " And Len(Compra.Comprobante) > 0";
			Funcion.EjecutaSQL(cdsCompra, stSelect, true);
			MessageBox.Show("Finalizado", "Información");
		}

		private void btIngreso_MouseUp(object sender, System.Windows.Forms.MouseEventArgs e)
		{
			if (e.Button != MouseButtons.Right) return;
			if (DialogResult.No == MessageBox.Show(
				"Desea Actualizar precios de Ingresos al costo promedio",
				"Información", MessageBoxButtons.YesNo, MessageBoxIcon.Question)) return;
			DateTime dtFecha = (DateTime) cmbDesde.Value;
			DateTime dtHasta = (DateTime) cmbHasta.Value;
			string stExec = "Exec IngresoActualizaCosto '" + dtFecha.ToString("yyyyMMdd")
				+ "', '" + dtHasta.ToString("yyyyMMdd") + "'" ;
			Funcion.EjecutaSQL(cdsCompra, stExec, true);
			MessageBox.Show("Ingresos de Bodega Recalculados al costo promedio", "Información");
		}

		private void btSaldoCostoB_Click(object sender, System.EventArgs e)
		{
			DateTime dtDesde = (DateTime) cmbDesde.Value;
			DateTime dtHasta = (DateTime) cmbHasta.Value;
			if (DialogResult.No == MessageBox.Show("Desea Actualizar costos desde el " 
				+ dtDesde.ToString("dd/MMM/yyyy")
				+ " hasta el " + dtHasta.ToString("dd/MMM/yyyy"), "Confirmación", 
				MessageBoxButtons.YesNo, MessageBoxIcon.Question)) return;

			this.Cursor = Cursors.WaitCursor;
			DateTime dtHoraIni = DateTime.Now;
			string stSelect = "EXEC ActualizaTodoB '" + dtDesde.ToString("yyyyMMdd")
				+ "', '" + dtHasta.AddDays(1).ToString("yyyyMMdd") + "'";
			Funcion.EjecutaSQL(cdsArticulo, stSelect);
			this.Cursor = Cursors.Default;
			DateTime dtHoraFin = DateTime.Now;

			TimeSpan ts = dtHoraFin.Subtract(dtHoraIni);
			MessageBox.Show("Costos x Bodega Generados en " +
				ts.Minutes + ":" + ts.Seconds.ToString());

		}

		private void btEnvioCostoB_Click(object sender, System.EventArgs e)
		{
			string stSelect = "Select Activado from SeteoG where codigo = 'I' ";
			bool bSelect = Funcion.bEjecutaSQL(cdsCompra, stSelect);
			if (bSelect == false) return;
	
			DateTime dtDesde = (DateTime) cmbDesde.Value;
			DateTime dtHasta = (DateTime) cmbHasta.Value;
			if (DialogResult.No == MessageBox.Show("Desea Enviar Costos " 
				+ dtDesde.ToString("dd/MMM/yyyy")
				+ " hasta el " + dtHasta.ToString("dd/MMM/yyyy"), "Confirmación", 
				MessageBoxButtons.YesNo, MessageBoxIcon.Question)) return;

			this.Cursor = Cursors.WaitCursor;
			DateTime dtHoraIni = DateTime.Now;
			stSelect = "EXEC Agro_Bodega_CC '" + dtDesde.ToString("yyyyMMdd")
				+ "', '" + dtHasta.AddDays(1).ToString("yyyyMMdd") + "'";
			Funcion.EjecutaSQL(cdsArticulo, stSelect);
			this.Cursor = Cursors.Default;
			DateTime dtHoraFin = DateTime.Now;

			TimeSpan ts = dtHoraFin.Subtract(dtHoraIni);
			MessageBox.Show("Costos x Bodega Enviados en " +
				ts.Minutes + ":" + ts.Seconds.ToString());

	
		}

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			this.cdsSeteoF.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);		
		}

		private void simpleButton2_Click(object sender, System.EventArgs e)
		{
			this.Cursor = Cursors.WaitCursor;
			int idArt=0;			
			if (!chkArt.Checked)
			{
				if(cmbArt.Text.Length <= 0){ MessageBox.Show("Seleccione Articulo..."); return;}
				idArt = (int) cmbArt.Value;			
			}
			DateTime dtdesde = (DateTime) cmbDesde.Value;
			

			string stSelect = "EXEC ActualizaSaldoCostoIdArtFIFO  "+idArt+", '"+dtdesde.ToString("yyyyMMdd")+"'";
			Funcion.EjecutaSQL(cdsArticulo, stSelect);
			this.Cursor = Cursors.Default;
		}

		private void chkArt_CheckedChanged(object sender, System.EventArgs e)
		{
			if (chkArt.Checked == true)
				cmbArt.Enabled = false;
			else
				cmbArt.Enabled = true;
		}

		private void btSaldoCostoLotes_Click(object sender, System.EventArgs e)
		{
			DateTime dtDesde = (DateTime) cmbDesde.Value;
			DateTime dtHasta = (DateTime) cmbHasta.Value;
			if (DialogResult.No == MessageBox.Show("Desea Actualizar costos desde el " 
				+ dtDesde.ToString("dd/MMM/yyyy")
				+ " hasta el " + dtHasta.ToString("dd/MMM/yyyy"), "Confirmación", 
				MessageBoxButtons.YesNo, MessageBoxIcon.Question)) return;

			this.Cursor = Cursors.WaitCursor;
			DateTime dtHoraIni = DateTime.Now;
			string stSelect = "EXEC ActualizaTodoLT '" + dtDesde.ToString("yyyyMMdd")
				+ "', '" + dtHasta.AddDays(1).ToString("yyyyMMdd") + "'";
			Funcion.EjecutaSQL(cdsArticulo, stSelect);
			this.Cursor = Cursors.Default;
			DateTime dtHoraFin = DateTime.Now;

			TimeSpan ts = dtHoraFin.Subtract(dtHoraIni);
			MessageBox.Show("Costos x Bodega Generados en " +
				ts.Minutes + ":" + ts.Seconds.ToString());
		}
	}
}
