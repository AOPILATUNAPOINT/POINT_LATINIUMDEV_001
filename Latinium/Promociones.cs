using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using DevExpress.XtraEditors;
using Infragistics.Win.UltraWinGrid;
using System.Data;
using System.Data.SqlClient;
using C1.Data;
using System.Globalization;

namespace Latinium
{
	/// <summary>
	/// Summary description for Promociones.
	/// </summary>
	public class Promociones : DevExpress.XtraEditors.XtraForm
	{
		private System.Windows.Forms.Label label1;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtFin;
		private System.Windows.Forms.Label label6;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtInicio;
		private System.Windows.Forms.Label lblCliente;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtDescripcion;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtIdPromocion;
		private System.ComponentModel.IContainer components;
		private C1.Data.C1DataRow drPromocion;
		private Infragistics.Win.UltraWinGrid.UltraGrid ultraGrid1;
		private C1.Data.C1DataRow drSeteoF;
		private System.Windows.Forms.Label label13;
		private System.Windows.Forms.Label label14;
		private System.Windows.Forms.Label label15;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtNroCuotas;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtCuotaReal;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtCuotaCalculada;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtValorMaximo;
		private System.Windows.Forms.Label label3;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtValorMinimo;
		private System.Windows.Forms.Label label2;
		private System.Windows.Forms.Label label16;
		private Infragistics.Win.UltraWinEditors.UltraComboEditor cmbTipo;
		private Infragistics.Win.UltraWinGrid.UltraGrid uGridTarjetas;
		private Infragistics.Win.UltraWinGrid.UltraGrid uGridCiudades;
		private Infragistics.Win.UltraWinGrid.UltraGrid uGridBodegas;
		private Infragistics.Win.UltraWinGrid.UltraGrid uGridTiposDePrecio;
		private Infragistics.Win.UltraWinGrid.UltraGrid uGridMarcas;
		private Infragistics.Win.UltraWinGrid.UltraGrid uGridSubgrupos;
		private Infragistics.Win.UltraWinGrid.UltraGrid uGridGrupos;
		private Infragistics.Win.UltraWinGrid.UltraGrid uGridBancos;
		private C1.Data.C1DataSet cdsPromocion;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource1;
		private DevExpress.XtraEditors.SimpleButton btNuevo;
		private DevExpress.XtraEditors.SimpleButton btSalir;
		private DevExpress.XtraEditors.SimpleButton btEditar;
		private DevExpress.XtraEditors.SimpleButton btBuscar;
		private DevExpress.XtraEditors.SimpleButton btGrabar;
		private DevExpress.XtraEditors.SimpleButton btCancelar;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource2;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource3;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource4;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource5;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource6;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource7;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource8;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource9;
		private DevExpress.XtraEditors.SimpleButton btFinalizar;
		private System.Windows.Forms.Label lblFinalizado;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtidArticulo;
		private C1.Data.C1DataSet cdsSeteoF;
		private DevExpress.XtraEditors.SimpleButton btnActualizaPrecios;
		private System.Windows.Forms.ImageList imageList1;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkLocales;
		private System.Windows.Forms.GroupBox groupBox1;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkGrupos;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkSubGrupos;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkMarcas;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkVariable;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource10;
		private Infragistics.Win.UltraWinGrid.UltraGrid uGridPlanes;
		private Infragistics.Win.UltraWinGrid.UltraGrid uGridModelos;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource11;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkSumaDescuentos;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkAdicional;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkValor;
		private Infragistics.Win.UltraWinGrid.UltraGrid uGridPlazos;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource12;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource13;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkOpcional;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource14;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkCuotas;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkCupon;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkValidaCargaProducto;
		private System.Windows.Forms.Label label4;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtMargenTotal;
		private Infragistics.Win.UltraWinGrid.UltraGrid uGridCuotas;
				
		public Promociones()
		{
			//
			// Required for Windows Form Designer support
			//
			InitializeComponent();

			//
			// TODO: Add any constructor code after InitializeComponent call
			//
		}

		public Promociones(int idPromocion, bool bInactivarControles)
		{
			InitializeComponent();
			this.Consulta(idPromocion);
			if(bInactivarControles)
			{
				this.btNuevo.Enabled = false;
				this.btBuscar.Enabled = false;
				this.btGrabar.Enabled = false;
				this.btFinalizar.Enabled = false;
				this.btCancelar.Enabled = false;
				this.btnActualizaPrecios.Enabled = false;
				this.btEditar.Enabled = false;
			}
		}

		/// <summary>
		/// Clean up any resources being used.
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
		/// Required method for Designer support - do not modify
		/// the contents of this method with the code editor.
		/// </summary>
		private void InitializeComponent()
		{
			this.components = new System.ComponentModel.Container();
			Infragistics.Win.Appearance appearance1 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton1 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(Promociones));
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton2 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idDetallePromocion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPromocion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idArticulo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Articulo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CostoPromedio");
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cantidad");
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Descuento");
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Precio2");
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("DescuentoTC");
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn11 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Precio3");
			Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn12 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("DescuentoCP");
			Infragistics.Win.Appearance appearance11 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn13 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Precio4");
			Infragistics.Win.Appearance appearance12 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn14 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("PromoPrecio2");
			Infragistics.Win.Appearance appearance13 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn15 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("PromoPrecio3");
			Infragistics.Win.Appearance appearance14 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn16 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("PromoPrecio4");
			Infragistics.Win.Appearance appearance15 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn17 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Tipo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn18 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Principal");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn19 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idArticuloAlterno1");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn20 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CodAlterno1");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn21 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idArticuloAlterno2");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn22 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CodAlterno2");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn23 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idArticuloAlterno3");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn24 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CodAlterno3");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn25 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idArticuloAlterno4");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn26 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CodAlterno4");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn27 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idArticuloAlterno5");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn28 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CodAlterno5");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn29 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("BloqueaDescuento");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn30 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ValorMinimo");
			Infragistics.Win.Appearance appearance16 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn31 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ValorMaximo");
			Infragistics.Win.Appearance appearance17 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn32 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Baratazos");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn33 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Combos");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn34 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Promociones");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn35 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Precio6");
			Infragistics.Win.Appearance appearance18 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance19 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn36 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("DescuentoTCSI");
			Infragistics.Win.Appearance appearance20 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance21 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn37 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("PrecioC3");
			Infragistics.Win.Appearance appearance22 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn38 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("PromoPrecioC3");
			Infragistics.Win.Appearance appearance23 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn39 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("PrecioC6");
			Infragistics.Win.Appearance appearance24 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn40 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("PromoPrecioC6");
			Infragistics.Win.Appearance appearance25 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn41 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("PrecioC9");
			Infragistics.Win.Appearance appearance26 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn42 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("PromoPrecioC9");
			Infragistics.Win.Appearance appearance27 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn43 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("PrecioC12");
			Infragistics.Win.Appearance appearance28 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn44 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("PromoPrecioC12");
			Infragistics.Win.Appearance appearance29 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn45 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("PrecioC15");
			Infragistics.Win.Appearance appearance30 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn46 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("PromoPrecioC15");
			Infragistics.Win.Appearance appearance31 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn47 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("PrecioC18");
			Infragistics.Win.Appearance appearance32 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn48 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("PromoPrecioC18");
			Infragistics.Win.Appearance appearance33 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn49 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("PrecioC21");
			Infragistics.Win.Appearance appearance34 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn50 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("PromoPrecioC21");
			Infragistics.Win.Appearance appearance35 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn51 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("PrecioC24");
			Infragistics.Win.Appearance appearance36 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn52 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("PromoPrecioC24");
			Infragistics.Win.Appearance appearance37 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn53 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("DescuentoCombo");
			Infragistics.Win.Appearance appearance38 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn54 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("WEBPRINCIPAL");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn55 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Aporte");
			Infragistics.Win.Appearance appearance39 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn56 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("UnidadMaxima");
			Infragistics.Win.Appearance appearance40 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn57 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCliente");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn58 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Proveedor");
			Infragistics.Win.Appearance appearance41 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance42 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance43 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings1 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "CostoPromedio", 5, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "CostoPromedio", 5, true);
			Infragistics.Win.Appearance appearance44 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings2 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Precio2", 8, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Precio2", 8, true);
			Infragistics.Win.Appearance appearance45 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings3 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Precio3", 10, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Precio3", 10, true);
			Infragistics.Win.Appearance appearance46 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings4 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Precio4", 12, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Precio4", 12, true);
			Infragistics.Win.Appearance appearance47 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings5 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "PromoPrecio2", 13, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "PromoPrecio2", 13, true);
			Infragistics.Win.Appearance appearance48 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings6 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "PromoPrecio3", 14, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "PromoPrecio3", 14, true);
			Infragistics.Win.Appearance appearance49 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings7 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "PromoPrecio4", 15, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "PromoPrecio4", 15, true);
			Infragistics.Win.Appearance appearance50 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings8 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Precio6", 34, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Precio6", 34, true);
			Infragistics.Win.Appearance appearance51 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance52 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance53 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance54 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance55 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance56 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance57 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance58 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idDetallePromocion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idPromocion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn3 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idArticulo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn4 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Codigo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn5 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Articulo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn6 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("CostoPromedio");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn7 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Cantidad");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn8 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Descuento");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn9 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Precio2");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn10 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("DescuentoTC");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn11 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Precio3");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn12 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("DescuentoCP");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn13 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Precio4");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn14 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("PromoPrecio2");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn15 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("PromoPrecio3");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn16 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("PromoPrecio4");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn17 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Tipo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn18 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Principal");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn19 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idArticuloAlterno1");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn20 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("CodAlterno1");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn21 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idArticuloAlterno2");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn22 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("CodAlterno2");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn23 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idArticuloAlterno3");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn24 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("CodAlterno3");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn25 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idArticuloAlterno4");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn26 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("CodAlterno4");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn27 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idArticuloAlterno5");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn28 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("CodAlterno5");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn29 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("BloqueaDescuento");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn30 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("ValorMinimo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn31 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("ValorMaximo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn32 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Baratazos");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn33 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Combos");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn34 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Promociones");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn35 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Precio6");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn36 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("DescuentoTCSI");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn37 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("PrecioC3");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn38 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("PromoPrecioC3");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn39 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("PrecioC6");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn40 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("PromoPrecioC6");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn41 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("PrecioC9");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn42 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("PromoPrecioC9");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn43 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("PrecioC12");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn44 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("PromoPrecioC12");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn45 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("PrecioC15");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn46 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("PromoPrecioC15");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn47 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("PrecioC18");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn48 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("PromoPrecioC18");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn49 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("PrecioC21");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn50 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("PromoPrecioC21");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn51 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("PrecioC24");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn52 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("PromoPrecioC24");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn53 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("DescuentoCombo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn54 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("WEBPRINCIPAL");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn55 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Aporte");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn56 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("UnidadMaxima");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn57 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCliente");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn58 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Proveedor");
			Infragistics.Win.Appearance appearance59 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance60 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance61 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance62 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance63 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance64 = new Infragistics.Win.Appearance();
			Infragistics.Win.ValueListItem valueListItem1 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem2 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem3 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem4 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.Appearance appearance65 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn59 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPromocionBancoTarjeta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn60 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPromocionBanco");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn61 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idTarjeta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn62 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn63 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Sel");
			Infragistics.Win.Appearance appearance66 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance67 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance68 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance69 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance70 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance71 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance72 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn59 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idPromocionBancoTarjeta");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn60 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idPromocionBanco");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn61 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idTarjeta");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn62 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn63 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Sel");
			Infragistics.Win.Appearance appearance73 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand3 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn64 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPromocionCiudad");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn65 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPromocion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn66 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCiudad");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn67 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Ciudad");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn68 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Sel");
			Infragistics.Win.Appearance appearance74 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance75 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance76 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance77 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance78 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance79 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance80 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn64 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idPromocionCiudad");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn65 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idPromocion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn66 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCiudad");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn67 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Ciudad");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn68 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Sel");
			Infragistics.Win.Appearance appearance81 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand4 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn69 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPromCiudadBodega");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn70 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPromocionCiudad");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn71 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bodega");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn72 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn73 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Sel");
			Infragistics.Win.Appearance appearance82 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance83 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance84 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance85 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance86 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance87 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance88 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn69 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idPromCiudadBodega");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn70 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idPromocionCiudad");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn71 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Bodega");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn72 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn73 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Sel");
			Infragistics.Win.Appearance appearance89 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand5 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn74 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPromoPrecio");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn75 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idTipoPrecio");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn76 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPromocion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn77 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Precio");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn78 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Operador");
			Infragistics.Win.Appearance appearance90 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn79 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("PorcentajeInicial");
			Infragistics.Win.Appearance appearance91 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn80 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Porcentaje");
			Infragistics.Win.Appearance appearance92 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn81 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Sel");
			Infragistics.Win.Appearance appearance93 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance94 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance95 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance96 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance97 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance98 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance99 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn74 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idPromoPrecio");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn75 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idTipoPrecio");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn76 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idPromocion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn77 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Precio");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn78 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Operador");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn79 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("PorcentajeInicial");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn80 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Porcentaje");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn81 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Sel");
			Infragistics.Win.Appearance appearance100 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand6 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn82 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPromocionMarca");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn83 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idMarca");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn84 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Marca");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn85 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Sel");
			Infragistics.Win.Appearance appearance101 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance102 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance103 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance104 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance105 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance106 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance107 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn82 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idPromocionMarca");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn83 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idMarca");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn84 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Marca");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn85 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Sel");
			Infragistics.Win.Appearance appearance108 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand7 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn86 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPromocionSubGrupo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn87 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idSubGrupo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn88 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("SubGrupo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn89 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Sel");
			Infragistics.Win.Appearance appearance109 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance110 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance111 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance112 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance113 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance114 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance115 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn86 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idPromocionSubGrupo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn87 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idSubGrupo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn88 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("SubGrupo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn89 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Sel");
			Infragistics.Win.Appearance appearance116 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand8 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn90 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPromocionGrupo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn91 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPromocion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn92 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idGrupoArticulo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn93 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Grupo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn94 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Sel");
			Infragistics.Win.Appearance appearance117 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance118 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance119 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance120 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance121 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance122 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance123 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn90 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idPromocionGrupo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn91 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idPromocion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn92 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idGrupoArticulo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn93 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Grupo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn94 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Sel");
			Infragistics.Win.Appearance appearance124 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand9 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn95 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPromocionBanco");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn96 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPromocion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn97 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idBanco");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn98 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn99 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Sel");
			Infragistics.Win.Appearance appearance125 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance126 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance127 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance128 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance129 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance130 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance131 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn95 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idPromocionBanco");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn96 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idPromocion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn97 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idBanco");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn98 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn99 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Sel");
			Infragistics.Win.Appearance appearance132 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance133 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance134 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance135 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance136 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance137 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand10 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn100 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPromocionesPlanes");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn101 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPromocionBanco");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn102 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPromocionBancoTarjeta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn103 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPlan");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn104 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPromocion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn105 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Descripcion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn106 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Sel");
			Infragistics.Win.Appearance appearance138 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance139 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance140 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance141 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance142 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance143 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance144 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn100 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idPromocionesPlanes");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn101 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idPromocionBanco");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn102 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idPromocionBancoTarjeta");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn103 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idPlan");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn104 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idPromocion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn105 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Descripcion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn106 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Sel");
			Infragistics.Win.Appearance appearance145 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand11 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn107 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPromocionModelo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn108 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idArticulo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn109 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Modelo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn110 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Sel");
			Infragistics.Win.Appearance appearance146 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance147 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance148 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance149 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance150 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance151 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance152 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn107 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idPromocionModelo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn108 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idArticulo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn109 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Modelo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn110 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Sel");
			Infragistics.Win.Appearance appearance153 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance154 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance155 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance156 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand12 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn111 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPromocionesPlazos");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn112 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPlazo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn113 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPromocion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn114 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Descripcion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn115 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Sel");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn116 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("MesesDeGracia");
			Infragistics.Win.Appearance appearance157 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance158 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance159 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance160 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance161 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance162 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance163 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance164 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn111 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idPromocionesPlazos");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn112 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idPlazo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn113 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idPromocion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn114 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Descripcion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn115 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Sel");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn116 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("MesesDeGracia");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn117 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idPromocionesCuotas");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn118 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idPromocion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn119 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Numero");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn120 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Sel");
			Infragistics.Win.Appearance appearance165 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance166 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand13 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn117 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPromocionesCuotas");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn118 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPromocion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn119 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cuota");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn120 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Sel");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn121 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Porcentaje_Entrada");
			Infragistics.Win.Appearance appearance167 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance168 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance169 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance170 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance171 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance172 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance173 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance174 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn121 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idPromocionesCuotas");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn122 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idPromocion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn123 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Cuota");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn124 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Sel");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn125 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Porcentaje_Entrada");
			Infragistics.Win.Appearance appearance175 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance176 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance177 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance178 = new Infragistics.Win.Appearance();
			this.label1 = new System.Windows.Forms.Label();
			this.dtFin = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.label6 = new System.Windows.Forms.Label();
			this.dtInicio = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.lblCliente = new System.Windows.Forms.Label();
			this.txtDescripcion = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtIdPromocion = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.ultraGrid1 = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource1 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.txtCuotaCalculada = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label13 = new System.Windows.Forms.Label();
			this.txtCuotaReal = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label14 = new System.Windows.Forms.Label();
			this.txtNroCuotas = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label15 = new System.Windows.Forms.Label();
			this.txtValorMaximo = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label3 = new System.Windows.Forms.Label();
			this.txtValorMinimo = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label2 = new System.Windows.Forms.Label();
			this.cmbTipo = new Infragistics.Win.UltraWinEditors.UltraComboEditor();
			this.label16 = new System.Windows.Forms.Label();
			this.uGridTarjetas = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource6 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.uGridCiudades = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource2 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.uGridBodegas = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource3 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.uGridTiposDePrecio = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource4 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.uGridMarcas = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource9 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.uGridSubgrupos = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource8 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.uGridGrupos = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource7 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.uGridBancos = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource5 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.cdsPromocion = new C1.Data.C1DataSet();
			this.btNuevo = new DevExpress.XtraEditors.SimpleButton();
			this.imageList1 = new System.Windows.Forms.ImageList(this.components);
			this.btSalir = new DevExpress.XtraEditors.SimpleButton();
			this.btEditar = new DevExpress.XtraEditors.SimpleButton();
			this.btBuscar = new DevExpress.XtraEditors.SimpleButton();
			this.btGrabar = new DevExpress.XtraEditors.SimpleButton();
			this.btCancelar = new DevExpress.XtraEditors.SimpleButton();
			this.btFinalizar = new DevExpress.XtraEditors.SimpleButton();
			this.lblFinalizado = new System.Windows.Forms.Label();
			this.txtidArticulo = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.btnActualizaPrecios = new DevExpress.XtraEditors.SimpleButton();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.chkLocales = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.groupBox1 = new System.Windows.Forms.GroupBox();
			this.chkGrupos = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.chkSubGrupos = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.chkMarcas = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.chkVariable = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.uGridPlanes = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource10 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.uGridModelos = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource11 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.chkSumaDescuentos = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.chkAdicional = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.chkValor = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.uGridPlazos = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource12 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.ultraDataSource13 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.chkOpcional = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.uGridCuotas = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource14 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.chkCuotas = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.chkCupon = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.chkValidaCargaProducto = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.label4 = new System.Windows.Forms.Label();
			this.txtMargenTotal = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			((System.ComponentModel.ISupportInitialize)(this.dtFin)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtInicio)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDescripcion)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdPromocion)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCuotaCalculada)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCuotaReal)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNroCuotas)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtValorMaximo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtValorMinimo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbTipo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridTarjetas)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource6)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridCiudades)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridBodegas)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource3)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridTiposDePrecio)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource4)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridMarcas)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource9)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridSubgrupos)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource8)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridGrupos)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource7)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridBancos)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource5)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsPromocion)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtidArticulo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridPlanes)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource10)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridModelos)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource11)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridPlazos)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource12)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource13)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridCuotas)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource14)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtMargenTotal)).BeginInit();
			this.SuspendLayout();
			// 
			// label1
			// 
			this.label1.AutoSize = true;
			this.label1.Enabled = false;
			this.label1.Location = new System.Drawing.Point(432, 10);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(33, 16);
			this.label1.TabIndex = 201;
			this.label1.Text = "Hasta";
			this.label1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// dtFin
			// 
			appearance1.FontData.Name = "Tahoma";
			appearance1.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtFin.Appearance = appearance1;
			this.dtFin.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton1.Caption = "Today";
			this.dtFin.DateButtons.Add(dateButton1);
			this.dtFin.Enabled = false;
			this.dtFin.Format = "dd/MM/yyyy HH:mm";
			this.dtFin.Location = new System.Drawing.Point(480, 8);
			this.dtFin.Name = "dtFin";
			this.dtFin.NonAutoSizeHeight = 23;
			this.dtFin.Size = new System.Drawing.Size(144, 21);
			this.dtFin.SpinButtonsVisible = true;
			this.dtFin.TabIndex = 2;
			this.dtFin.Value = ((object)(resources.GetObject("dtFin.Value")));
			this.dtFin.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.dtFin_KeyPress);
			// 
			// label6
			// 
			this.label6.AutoSize = true;
			this.label6.Enabled = false;
			this.label6.Location = new System.Drawing.Point(224, 10);
			this.label6.Name = "label6";
			this.label6.Size = new System.Drawing.Size(35, 16);
			this.label6.TabIndex = 199;
			this.label6.Text = "Desde";
			this.label6.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// dtInicio
			// 
			appearance2.FontData.Name = "Tahoma";
			appearance2.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtInicio.Appearance = appearance2;
			this.dtInicio.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton2.Caption = "Today";
			this.dtInicio.DateButtons.Add(dateButton2);
			this.dtInicio.Enabled = false;
			this.dtInicio.Format = "dd/MM/yyyy HH:mm";
			this.dtInicio.Location = new System.Drawing.Point(272, 8);
			this.dtInicio.Name = "dtInicio";
			this.dtInicio.NonAutoSizeHeight = 23;
			this.dtInicio.Size = new System.Drawing.Size(144, 21);
			this.dtInicio.SpinButtonsVisible = true;
			this.dtInicio.TabIndex = 1;
			this.dtInicio.Value = ((object)(resources.GetObject("dtInicio.Value")));
			this.dtInicio.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.dtInicio_KeyPress);
			this.dtInicio.Validated += new System.EventHandler(this.dtInicio_Validated);
			// 
			// lblCliente
			// 
			this.lblCliente.AutoSize = true;
			this.lblCliente.Enabled = false;
			this.lblCliente.Location = new System.Drawing.Point(640, 8);
			this.lblCliente.Name = "lblCliente";
			this.lblCliente.Size = new System.Drawing.Size(44, 16);
			this.lblCliente.TabIndex = 205;
			this.lblCliente.Text = "Nombre";
			this.lblCliente.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtDescripcion
			// 
			appearance3.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtDescripcion.Appearance = appearance3;
			this.txtDescripcion.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtDescripcion.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtDescripcion.Enabled = false;
			this.txtDescripcion.Location = new System.Drawing.Point(696, 8);
			this.txtDescripcion.Name = "txtDescripcion";
			this.txtDescripcion.Size = new System.Drawing.Size(352, 22);
			this.txtDescripcion.TabIndex = 3;
			this.txtDescripcion.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtDescripcion_KeyPress);
			// 
			// txtIdPromocion
			// 
			this.txtIdPromocion.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.txtIdPromocion.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtIdPromocion.Enabled = false;
			this.txtIdPromocion.Location = new System.Drawing.Point(1056, 8);
			this.txtIdPromocion.Name = "txtIdPromocion";
			this.txtIdPromocion.PromptChar = ' ';
			this.txtIdPromocion.Size = new System.Drawing.Size(16, 22);
			this.txtIdPromocion.TabIndex = 234;
			// 
			// ultraGrid1
			// 
			this.ultraGrid1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.ultraGrid1.Cursor = System.Windows.Forms.Cursors.Default;
			this.ultraGrid1.DataSource = this.ultraDataSource1;
			appearance4.BackColor = System.Drawing.Color.White;
			this.ultraGrid1.DisplayLayout.Appearance = appearance4;
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn1.Width = 127;
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridColumn2.Hidden = true;
			ultraGridColumn2.Width = 97;
			ultraGridColumn3.Header.VisiblePosition = 2;
			ultraGridColumn3.Hidden = true;
			ultraGridColumn3.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.Edit;
			ultraGridColumn3.Width = 91;
			ultraGridColumn4.Header.VisiblePosition = 3;
			ultraGridColumn4.Width = 180;
			ultraGridColumn5.Header.VisiblePosition = 4;
			ultraGridColumn5.Hidden = true;
			ultraGridColumn5.Width = 204;
			ultraGridColumn6.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance5.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn6.CellAppearance = appearance5;
			ultraGridColumn6.Format = "#,##0.00";
			ultraGridColumn6.Header.Caption = "Costo";
			ultraGridColumn6.Header.VisiblePosition = 5;
			ultraGridColumn6.Width = 70;
			appearance6.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn7.CellAppearance = appearance6;
			ultraGridColumn7.Header.Caption = "Cant";
			ultraGridColumn7.Header.VisiblePosition = 6;
			ultraGridColumn7.Width = 40;
			appearance7.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn8.CellAppearance = appearance7;
			ultraGridColumn8.Format = "#,##0.000";
			ultraGridColumn8.Header.Caption = "Dscto";
			ultraGridColumn8.Header.VisiblePosition = 7;
			ultraGridColumn8.PromptChar = ' ';
			ultraGridColumn8.Width = 60;
			ultraGridColumn9.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance8.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn9.CellAppearance = appearance8;
			ultraGridColumn9.Format = "#,##0.00";
			ultraGridColumn9.Header.Caption = "Contado";
			ultraGridColumn9.Header.VisiblePosition = 8;
			ultraGridColumn9.Width = 70;
			appearance9.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn10.CellAppearance = appearance9;
			ultraGridColumn10.Format = "#,##0.000";
			ultraGridColumn10.Header.Caption = "% Dscto TC";
			ultraGridColumn10.Header.VisiblePosition = 9;
			ultraGridColumn10.Width = 70;
			ultraGridColumn11.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance10.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn11.CellAppearance = appearance10;
			ultraGridColumn11.Format = "#,##0.00";
			ultraGridColumn11.Header.Caption = "Tarjeta";
			ultraGridColumn11.Header.VisiblePosition = 10;
			ultraGridColumn11.Width = 70;
			appearance11.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn12.CellAppearance = appearance11;
			ultraGridColumn12.Format = "#,##0.000";
			ultraGridColumn12.Header.Caption = "% Dscto CP";
			ultraGridColumn12.Header.VisiblePosition = 11;
			ultraGridColumn12.Width = 70;
			ultraGridColumn13.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance12.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn13.CellAppearance = appearance12;
			ultraGridColumn13.Format = "#,##0.00";
			ultraGridColumn13.Header.Caption = "Crédito";
			ultraGridColumn13.Header.VisiblePosition = 12;
			ultraGridColumn13.Width = 70;
			ultraGridColumn14.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance13.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn14.CellAppearance = appearance13;
			ultraGridColumn14.Format = "#,##0.00";
			ultraGridColumn14.Header.Caption = "Prom Cont";
			ultraGridColumn14.Header.VisiblePosition = 15;
			ultraGridColumn14.Width = 70;
			ultraGridColumn15.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance14.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn15.CellAppearance = appearance14;
			ultraGridColumn15.Format = "#,##0.00";
			ultraGridColumn15.Header.Caption = "Prom Tarj";
			ultraGridColumn15.Header.VisiblePosition = 16;
			ultraGridColumn15.Width = 70;
			ultraGridColumn16.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance15.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn16.CellAppearance = appearance15;
			ultraGridColumn16.Format = "#,##0.00";
			ultraGridColumn16.Header.Caption = "Prom Créd";
			ultraGridColumn16.Header.VisiblePosition = 17;
			ultraGridColumn16.Width = 70;
			ultraGridColumn17.Header.VisiblePosition = 18;
			ultraGridColumn17.Hidden = true;
			ultraGridColumn18.Header.VisiblePosition = 21;
			ultraGridColumn18.Width = 70;
			ultraGridColumn19.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn19.Header.VisiblePosition = 31;
			ultraGridColumn19.Hidden = true;
			ultraGridColumn20.Header.Caption = "Codigo Alterno 1";
			ultraGridColumn20.Header.VisiblePosition = 32;
			ultraGridColumn20.Width = 180;
			ultraGridColumn21.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn21.Header.VisiblePosition = 33;
			ultraGridColumn21.Hidden = true;
			ultraGridColumn22.Header.Caption = "Codigo Alterno 2";
			ultraGridColumn22.Header.VisiblePosition = 34;
			ultraGridColumn22.Width = 180;
			ultraGridColumn23.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn23.Header.VisiblePosition = 35;
			ultraGridColumn23.Hidden = true;
			ultraGridColumn24.Header.Caption = "Codigo Alterno 3";
			ultraGridColumn24.Header.VisiblePosition = 36;
			ultraGridColumn24.Width = 180;
			ultraGridColumn25.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn25.Header.VisiblePosition = 37;
			ultraGridColumn25.Hidden = true;
			ultraGridColumn26.Header.Caption = "Codigo Alterno 4";
			ultraGridColumn26.Header.VisiblePosition = 38;
			ultraGridColumn26.Width = 180;
			ultraGridColumn27.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn27.Header.VisiblePosition = 39;
			ultraGridColumn27.Hidden = true;
			ultraGridColumn28.Header.Caption = "Codigo Alterno 5";
			ultraGridColumn28.Header.VisiblePosition = 40;
			ultraGridColumn28.Width = 180;
			ultraGridColumn29.Header.Caption = "Bloq Dscto";
			ultraGridColumn29.Header.VisiblePosition = 23;
			ultraGridColumn29.Width = 70;
			appearance16.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn30.CellAppearance = appearance16;
			ultraGridColumn30.Format = "#,##0.00";
			ultraGridColumn30.Header.Caption = "Valor Minimo";
			ultraGridColumn30.Header.VisiblePosition = 19;
			ultraGridColumn30.Width = 70;
			appearance17.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn31.CellAppearance = appearance17;
			ultraGridColumn31.Format = "#,##0.00";
			ultraGridColumn31.Header.Caption = "Valor Maximo";
			ultraGridColumn31.Header.VisiblePosition = 20;
			ultraGridColumn32.Header.VisiblePosition = 24;
			ultraGridColumn32.Width = 70;
			ultraGridColumn33.Header.VisiblePosition = 25;
			ultraGridColumn33.Width = 70;
			ultraGridColumn34.Header.VisiblePosition = 26;
			ultraGridColumn35.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance18.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn35.CellAppearance = appearance18;
			appearance19.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn35.CellButtonAppearance = appearance19;
			ultraGridColumn35.Format = "#,##0.00";
			ultraGridColumn35.Header.Caption = "Tarjeta Sin Intereses";
			ultraGridColumn35.Header.VisiblePosition = 14;
			appearance20.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn36.CellAppearance = appearance20;
			appearance21.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn36.CellButtonAppearance = appearance21;
			ultraGridColumn36.Format = "#,##0.000";
			ultraGridColumn36.Header.Caption = "% Desc TCSI";
			ultraGridColumn36.Header.VisiblePosition = 13;
			ultraGridColumn37.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance22.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn37.CellAppearance = appearance22;
			ultraGridColumn37.Format = "#,##0.00";
			ultraGridColumn37.Header.Caption = "Precio C 3";
			ultraGridColumn37.Header.VisiblePosition = 41;
			ultraGridColumn37.Hidden = true;
			ultraGridColumn37.Width = 100;
			ultraGridColumn38.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance23.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn38.CellAppearance = appearance23;
			ultraGridColumn38.Format = "#,##0.00";
			ultraGridColumn38.Header.Caption = "Promo Precio C 3";
			ultraGridColumn38.Header.VisiblePosition = 42;
			ultraGridColumn38.Hidden = true;
			ultraGridColumn38.Width = 100;
			ultraGridColumn39.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance24.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn39.CellAppearance = appearance24;
			ultraGridColumn39.Format = "#,##0.00";
			ultraGridColumn39.Header.Caption = "Precio C 6";
			ultraGridColumn39.Header.VisiblePosition = 43;
			ultraGridColumn39.Hidden = true;
			ultraGridColumn39.Width = 100;
			ultraGridColumn40.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance25.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn40.CellAppearance = appearance25;
			ultraGridColumn40.Format = "#,##0.00";
			ultraGridColumn40.Header.Caption = "Promo Precio C 6";
			ultraGridColumn40.Header.VisiblePosition = 44;
			ultraGridColumn40.Hidden = true;
			ultraGridColumn40.Width = 100;
			ultraGridColumn41.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance26.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn41.CellAppearance = appearance26;
			ultraGridColumn41.Format = "#,##0.00";
			ultraGridColumn41.Header.Caption = "Precio C 9";
			ultraGridColumn41.Header.VisiblePosition = 45;
			ultraGridColumn41.Hidden = true;
			ultraGridColumn41.Width = 100;
			ultraGridColumn42.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance27.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn42.CellAppearance = appearance27;
			ultraGridColumn42.Format = "#,##0.00";
			ultraGridColumn42.Header.Caption = "Promo Precio C 9";
			ultraGridColumn42.Header.VisiblePosition = 46;
			ultraGridColumn42.Hidden = true;
			ultraGridColumn42.Width = 100;
			ultraGridColumn43.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance28.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn43.CellAppearance = appearance28;
			ultraGridColumn43.Format = "#,##0.00";
			ultraGridColumn43.Header.Caption = "Precio C 12";
			ultraGridColumn43.Header.VisiblePosition = 47;
			ultraGridColumn43.Hidden = true;
			ultraGridColumn43.Width = 100;
			ultraGridColumn44.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance29.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn44.CellAppearance = appearance29;
			ultraGridColumn44.Format = "#,##0.00";
			ultraGridColumn44.Header.Caption = "Promo Precio C 12";
			ultraGridColumn44.Header.VisiblePosition = 48;
			ultraGridColumn44.Hidden = true;
			ultraGridColumn44.Width = 100;
			ultraGridColumn45.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance30.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn45.CellAppearance = appearance30;
			ultraGridColumn45.Format = "#,##0.00";
			ultraGridColumn45.Header.Caption = "Precio C 15";
			ultraGridColumn45.Header.VisiblePosition = 49;
			ultraGridColumn45.Hidden = true;
			ultraGridColumn45.Width = 100;
			ultraGridColumn46.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance31.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn46.CellAppearance = appearance31;
			ultraGridColumn46.Format = "#,##0.00";
			ultraGridColumn46.Header.Caption = "Promo Precio C 15";
			ultraGridColumn46.Header.VisiblePosition = 50;
			ultraGridColumn46.Hidden = true;
			ultraGridColumn46.Width = 100;
			ultraGridColumn47.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance32.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn47.CellAppearance = appearance32;
			ultraGridColumn47.Format = "#,##0.00";
			ultraGridColumn47.Header.Caption = "Precio C 18";
			ultraGridColumn47.Header.VisiblePosition = 51;
			ultraGridColumn47.Hidden = true;
			ultraGridColumn47.Width = 100;
			ultraGridColumn48.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance33.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn48.CellAppearance = appearance33;
			ultraGridColumn48.Format = "#,##0.00";
			ultraGridColumn48.Header.Caption = "Promo Precio C 18";
			ultraGridColumn48.Header.VisiblePosition = 52;
			ultraGridColumn48.Hidden = true;
			ultraGridColumn48.Width = 100;
			ultraGridColumn49.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance34.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn49.CellAppearance = appearance34;
			ultraGridColumn49.Format = "#,##0.00";
			ultraGridColumn49.Header.Caption = "Precio C 21";
			ultraGridColumn49.Header.VisiblePosition = 53;
			ultraGridColumn49.Hidden = true;
			ultraGridColumn49.Width = 100;
			ultraGridColumn50.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance35.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn50.CellAppearance = appearance35;
			ultraGridColumn50.Format = "#,##0.00";
			ultraGridColumn50.Header.Caption = "Promo Precio C 21";
			ultraGridColumn50.Header.VisiblePosition = 54;
			ultraGridColumn50.Hidden = true;
			ultraGridColumn50.Width = 100;
			ultraGridColumn51.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance36.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn51.CellAppearance = appearance36;
			ultraGridColumn51.Format = "#,##0.00";
			ultraGridColumn51.Header.Caption = "Precio C 24";
			ultraGridColumn51.Header.VisiblePosition = 55;
			ultraGridColumn51.Hidden = true;
			ultraGridColumn51.Width = 100;
			ultraGridColumn52.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance37.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn52.CellAppearance = appearance37;
			ultraGridColumn52.Format = "#,##0.00";
			ultraGridColumn52.Header.Caption = "Promo Precio C 24";
			ultraGridColumn52.Header.VisiblePosition = 56;
			ultraGridColumn52.Hidden = true;
			ultraGridColumn52.Width = 100;
			appearance38.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn53.CellAppearance = appearance38;
			ultraGridColumn53.Format = "#,##0.00";
			ultraGridColumn53.Header.VisiblePosition = 57;
			ultraGridColumn54.Header.VisiblePosition = 22;
			appearance39.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn55.CellAppearance = appearance39;
			ultraGridColumn55.Format = "#,##0.00";
			ultraGridColumn55.Header.VisiblePosition = 29;
			ultraGridColumn55.Width = 90;
			appearance40.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn56.CellAppearance = appearance40;
			ultraGridColumn56.Header.Caption = "Unidades Máximas";
			ultraGridColumn56.Header.VisiblePosition = 30;
			ultraGridColumn57.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn57.Header.VisiblePosition = 27;
			ultraGridColumn57.Hidden = true;
			ultraGridColumn58.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn58.Header.VisiblePosition = 28;
			ultraGridColumn58.Width = 180;
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
																										 ultraGridColumn58});
			appearance41.TextHAlign = Infragistics.Win.HAlign.Left;
			ultraGridBand1.Header.Appearance = appearance41;
			ultraGridBand1.Header.Caption = "Lista de Artículos Para Combos";
			ultraGridBand1.HeaderVisible = true;
			appearance42.ForeColor = System.Drawing.Color.Black;
			appearance42.ForeColorDisabled = System.Drawing.Color.Black;
			ultraGridBand1.Override.ActiveCellAppearance = appearance42;
			appearance43.FontData.SizeInPoints = 7.25F;
			ultraGridBand1.Override.HeaderAppearance = appearance43;
			appearance44.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings1.Appearance = appearance44;
			summarySettings1.DisplayFormat = "{0: #,##0.00}";
			summarySettings1.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance45.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings2.Appearance = appearance45;
			summarySettings2.DisplayFormat = "{0: #,##0.00}";
			summarySettings2.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance46.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings3.Appearance = appearance46;
			summarySettings3.DisplayFormat = "{0: #,##0.00}";
			summarySettings3.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance47.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings4.Appearance = appearance47;
			summarySettings4.DisplayFormat = "{0: #,##0.00}";
			summarySettings4.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance48.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings5.Appearance = appearance48;
			summarySettings5.DisplayFormat = "{0: #,##0.00}";
			summarySettings5.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance49.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings6.Appearance = appearance49;
			summarySettings6.DisplayFormat = "{0: #,##0.00}";
			summarySettings6.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance50.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings7.Appearance = appearance50;
			summarySettings7.DisplayFormat = "{0: #,##0.00}";
			summarySettings7.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance51.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings8.Appearance = appearance51;
			summarySettings8.DisplayFormat = "{0: #,##0.00}";
			summarySettings8.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			ultraGridBand1.Summaries.AddRange(new Infragistics.Win.UltraWinGrid.SummarySettings[] {
																																															summarySettings1,
																																															summarySettings2,
																																															summarySettings3,
																																															summarySettings4,
																																															summarySettings5,
																																															summarySettings6,
																																															summarySettings7,
																																															summarySettings8});
			ultraGridBand1.SummaryFooterCaption = "";
			this.ultraGrid1.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			appearance52.ForeColor = System.Drawing.Color.Black;
			appearance52.ForeColorDisabled = System.Drawing.Color.Black;
			this.ultraGrid1.DisplayLayout.Override.ActiveRowAppearance = appearance52;
			this.ultraGrid1.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.TemplateOnBottom;
			this.ultraGrid1.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.True;
			this.ultraGrid1.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance53.BackColor = System.Drawing.Color.Transparent;
			this.ultraGrid1.DisplayLayout.Override.CardAreaAppearance = appearance53;
			appearance54.ForeColor = System.Drawing.Color.Black;
			appearance54.ForeColorDisabled = System.Drawing.Color.Black;
			this.ultraGrid1.DisplayLayout.Override.CellAppearance = appearance54;
			appearance55.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance55.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance55.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance55.FontData.BoldAsString = "True";
			appearance55.FontData.Name = "Arial";
			appearance55.FontData.SizeInPoints = 10F;
			appearance55.ForeColor = System.Drawing.Color.White;
			appearance55.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.ultraGrid1.DisplayLayout.Override.HeaderAppearance = appearance55;
			appearance56.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance56.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance56.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.RowAlternateAppearance = appearance56;
			appearance57.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance57.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance57.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.RowSelectorAppearance = appearance57;
			appearance58.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance58.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance58.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.SelectedRowAppearance = appearance58;
			this.ultraGrid1.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.ultraGrid1.Location = new System.Drawing.Point(8, 552);
			this.ultraGrid1.Name = "ultraGrid1";
			this.ultraGrid1.Size = new System.Drawing.Size(1086, 91);
			this.ultraGrid1.TabIndex = 248;
			this.ultraGrid1.InitializeRow += new Infragistics.Win.UltraWinGrid.InitializeRowEventHandler(this.ultraGrid1_InitializeRow);
			this.ultraGrid1.InitializeLogicalPrintPage += new Infragistics.Win.UltraWinGrid.InitializeLogicalPrintPageEventHandler(this.ultraGrid1_InitializeLogicalPrintPage);
			this.ultraGrid1.AfterRowsDeleted += new System.EventHandler(this.ultraGrid1_AfterRowsDeleted);
			this.ultraGrid1.AfterRowInsert += new Infragistics.Win.UltraWinGrid.RowEventHandler(this.ultraGrid1_AfterRowInsert);
			this.ultraGrid1.BeforeRowsDeleted += new Infragistics.Win.UltraWinGrid.BeforeRowsDeletedEventHandler(this.ultraGrid1_BeforeRowsDeleted);
			this.ultraGrid1.KeyDown += new System.Windows.Forms.KeyEventHandler(this.ultraGrid1_KeyDown);
			this.ultraGrid1.AfterCellUpdate += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.ultraGrid1_AfterCellUpdate);
			this.ultraGrid1.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.ultraGrid1_InitializeLayout);
			// 
			// ultraDataSource1
			// 
			ultraDataColumn1.DataType = typeof(int);
			ultraDataColumn2.DataType = typeof(int);
			ultraDataColumn3.DataType = typeof(int);
			ultraDataColumn6.DataType = typeof(System.Decimal);
			ultraDataColumn7.DataType = typeof(int);
			ultraDataColumn8.DataType = typeof(System.Decimal);
			ultraDataColumn9.DataType = typeof(System.Decimal);
			ultraDataColumn10.DataType = typeof(System.Decimal);
			ultraDataColumn11.DataType = typeof(System.Decimal);
			ultraDataColumn12.DataType = typeof(System.Decimal);
			ultraDataColumn13.DataType = typeof(System.Decimal);
			ultraDataColumn14.DataType = typeof(System.Decimal);
			ultraDataColumn15.DataType = typeof(System.Decimal);
			ultraDataColumn16.DataType = typeof(System.Decimal);
			ultraDataColumn17.DataType = typeof(int);
			ultraDataColumn18.DataType = typeof(bool);
			ultraDataColumn19.DataType = typeof(int);
			ultraDataColumn21.DataType = typeof(int);
			ultraDataColumn23.DataType = typeof(int);
			ultraDataColumn25.DataType = typeof(int);
			ultraDataColumn27.DataType = typeof(int);
			ultraDataColumn29.DataType = typeof(bool);
			ultraDataColumn30.DataType = typeof(System.Decimal);
			ultraDataColumn31.DataType = typeof(System.Decimal);
			ultraDataColumn32.DataType = typeof(bool);
			ultraDataColumn33.DataType = typeof(bool);
			ultraDataColumn34.DataType = typeof(bool);
			ultraDataColumn35.DataType = typeof(System.Decimal);
			ultraDataColumn36.DataType = typeof(System.Decimal);
			ultraDataColumn37.DataType = typeof(System.Decimal);
			ultraDataColumn37.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
			ultraDataColumn38.DataType = typeof(System.Decimal);
			ultraDataColumn38.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
			ultraDataColumn39.DataType = typeof(System.Decimal);
			ultraDataColumn39.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
			ultraDataColumn40.DataType = typeof(System.Decimal);
			ultraDataColumn40.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
			ultraDataColumn41.DataType = typeof(System.Decimal);
			ultraDataColumn41.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
			ultraDataColumn42.DataType = typeof(System.Decimal);
			ultraDataColumn42.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
			ultraDataColumn43.DataType = typeof(System.Decimal);
			ultraDataColumn43.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
			ultraDataColumn44.DataType = typeof(System.Decimal);
			ultraDataColumn44.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
			ultraDataColumn45.DataType = typeof(System.Decimal);
			ultraDataColumn45.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
			ultraDataColumn46.DataType = typeof(System.Decimal);
			ultraDataColumn46.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
			ultraDataColumn47.DataType = typeof(System.Decimal);
			ultraDataColumn47.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
			ultraDataColumn48.DataType = typeof(System.Decimal);
			ultraDataColumn48.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
			ultraDataColumn49.DataType = typeof(System.Decimal);
			ultraDataColumn49.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
			ultraDataColumn50.DataType = typeof(System.Decimal);
			ultraDataColumn50.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
			ultraDataColumn51.DataType = typeof(System.Decimal);
			ultraDataColumn51.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
			ultraDataColumn52.DataType = typeof(System.Decimal);
			ultraDataColumn52.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
			ultraDataColumn53.DataType = typeof(System.Decimal);
			ultraDataColumn54.DataType = typeof(bool);
			ultraDataColumn55.DataType = typeof(System.Decimal);
			ultraDataColumn56.DataType = typeof(int);
			ultraDataColumn57.DataType = typeof(int);
			this.ultraDataSource1.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn1,
																																 ultraDataColumn2,
																																 ultraDataColumn3,
																																 ultraDataColumn4,
																																 ultraDataColumn5,
																																 ultraDataColumn6,
																																 ultraDataColumn7,
																																 ultraDataColumn8,
																																 ultraDataColumn9,
																																 ultraDataColumn10,
																																 ultraDataColumn11,
																																 ultraDataColumn12,
																																 ultraDataColumn13,
																																 ultraDataColumn14,
																																 ultraDataColumn15,
																																 ultraDataColumn16,
																																 ultraDataColumn17,
																																 ultraDataColumn18,
																																 ultraDataColumn19,
																																 ultraDataColumn20,
																																 ultraDataColumn21,
																																 ultraDataColumn22,
																																 ultraDataColumn23,
																																 ultraDataColumn24,
																																 ultraDataColumn25,
																																 ultraDataColumn26,
																																 ultraDataColumn27,
																																 ultraDataColumn28,
																																 ultraDataColumn29,
																																 ultraDataColumn30,
																																 ultraDataColumn31,
																																 ultraDataColumn32,
																																 ultraDataColumn33,
																																 ultraDataColumn34,
																																 ultraDataColumn35,
																																 ultraDataColumn36,
																																 ultraDataColumn37,
																																 ultraDataColumn38,
																																 ultraDataColumn39,
																																 ultraDataColumn40,
																																 ultraDataColumn41,
																																 ultraDataColumn42,
																																 ultraDataColumn43,
																																 ultraDataColumn44,
																																 ultraDataColumn45,
																																 ultraDataColumn46,
																																 ultraDataColumn47,
																																 ultraDataColumn48,
																																 ultraDataColumn49,
																																 ultraDataColumn50,
																																 ultraDataColumn51,
																																 ultraDataColumn52,
																																 ultraDataColumn53,
																																 ultraDataColumn54,
																																 ultraDataColumn55,
																																 ultraDataColumn56,
																																 ultraDataColumn57,
																																 ultraDataColumn58});
			// 
			// txtCuotaCalculada
			// 
			appearance59.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtCuotaCalculada.Appearance = appearance59;
			this.txtCuotaCalculada.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.txtCuotaCalculada.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtCuotaCalculada.Enabled = false;
			this.txtCuotaCalculada.FormatString = "#,##0.00";
			this.txtCuotaCalculada.Location = new System.Drawing.Point(992, 112);
			this.txtCuotaCalculada.Name = "txtCuotaCalculada";
			this.txtCuotaCalculada.NullText = "0.00";
			this.txtCuotaCalculada.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtCuotaCalculada.PromptChar = ' ';
			this.txtCuotaCalculada.Size = new System.Drawing.Size(72, 20);
			this.txtCuotaCalculada.TabIndex = 252;
			this.txtCuotaCalculada.Visible = false;
			// 
			// label13
			// 
			this.label13.AutoSize = true;
			this.label13.BackColor = System.Drawing.Color.Transparent;
			this.label13.Enabled = false;
			this.label13.Location = new System.Drawing.Point(992, 96);
			this.label13.Name = "label13";
			this.label13.Size = new System.Drawing.Size(84, 16);
			this.label13.TabIndex = 255;
			this.label13.Text = "Cuota Calculada";
			this.label13.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.label13.Visible = false;
			// 
			// txtCuotaReal
			// 
			appearance60.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtCuotaReal.Appearance = appearance60;
			this.txtCuotaReal.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.txtCuotaReal.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtCuotaReal.Enabled = false;
			this.txtCuotaReal.FormatString = "#,##0.00";
			this.txtCuotaReal.Location = new System.Drawing.Point(992, 72);
			this.txtCuotaReal.Name = "txtCuotaReal";
			this.txtCuotaReal.NullText = "0.00";
			this.txtCuotaReal.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtCuotaReal.PromptChar = ' ';
			this.txtCuotaReal.Size = new System.Drawing.Size(72, 20);
			this.txtCuotaReal.TabIndex = 251;
			this.txtCuotaReal.Visible = false;
			// 
			// label14
			// 
			this.label14.AutoSize = true;
			this.label14.BackColor = System.Drawing.Color.Transparent;
			this.label14.Enabled = false;
			this.label14.Location = new System.Drawing.Point(992, 56);
			this.label14.Name = "label14";
			this.label14.Size = new System.Drawing.Size(58, 16);
			this.label14.TabIndex = 254;
			this.label14.Text = "Cuota Real";
			this.label14.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.label14.Visible = false;
			// 
			// txtNroCuotas
			// 
			appearance61.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtNroCuotas.Appearance = appearance61;
			this.txtNroCuotas.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.txtNroCuotas.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNroCuotas.Enabled = false;
			this.txtNroCuotas.FormatString = "";
			this.txtNroCuotas.Location = new System.Drawing.Point(872, 120);
			this.txtNroCuotas.MaxValue = 100;
			this.txtNroCuotas.MinValue = 0;
			this.txtNroCuotas.Name = "txtNroCuotas";
			this.txtNroCuotas.PromptChar = ' ';
			this.txtNroCuotas.Size = new System.Drawing.Size(72, 20);
			this.txtNroCuotas.TabIndex = 250;
			this.txtNroCuotas.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtNroCuotas_KeyPress);
			this.txtNroCuotas.ValueChanged += new System.EventHandler(this.txtNroCuotas_ValueChanged);
			// 
			// label15
			// 
			this.label15.AutoSize = true;
			this.label15.BackColor = System.Drawing.Color.Transparent;
			this.label15.Enabled = false;
			this.label15.Location = new System.Drawing.Point(784, 120);
			this.label15.Name = "label15";
			this.label15.Size = new System.Drawing.Size(38, 16);
			this.label15.TabIndex = 253;
			this.label15.Text = "Cuotas";
			this.label15.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtValorMaximo
			// 
			appearance62.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtValorMaximo.Appearance = appearance62;
			this.txtValorMaximo.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.txtValorMaximo.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtValorMaximo.Enabled = false;
			this.txtValorMaximo.FormatString = "#,##0.00";
			this.txtValorMaximo.Location = new System.Drawing.Point(872, 168);
			this.txtValorMaximo.Name = "txtValorMaximo";
			this.txtValorMaximo.NullText = "0.00";
			this.txtValorMaximo.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtValorMaximo.PromptChar = ' ';
			this.txtValorMaximo.Size = new System.Drawing.Size(72, 20);
			this.txtValorMaximo.TabIndex = 268;
			this.txtValorMaximo.Visible = false;
			this.txtValorMaximo.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtValorMaximo_KeyPress);
			// 
			// label3
			// 
			this.label3.AutoSize = true;
			this.label3.BackColor = System.Drawing.Color.Transparent;
			this.label3.Enabled = false;
			this.label3.Location = new System.Drawing.Point(784, 168);
			this.label3.Name = "label3";
			this.label3.Size = new System.Drawing.Size(77, 16);
			this.label3.TabIndex = 271;
			this.label3.Text = "Precio máximo";
			this.label3.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.label3.Visible = false;
			// 
			// txtValorMinimo
			// 
			appearance63.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtValorMinimo.Appearance = appearance63;
			this.txtValorMinimo.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.txtValorMinimo.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtValorMinimo.Enabled = false;
			this.txtValorMinimo.FormatString = "#,##0.00";
			this.txtValorMinimo.Location = new System.Drawing.Point(872, 144);
			this.txtValorMinimo.Name = "txtValorMinimo";
			this.txtValorMinimo.NullText = "0.00";
			this.txtValorMinimo.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtValorMinimo.PromptChar = ' ';
			this.txtValorMinimo.Size = new System.Drawing.Size(72, 20);
			this.txtValorMinimo.TabIndex = 267;
			this.txtValorMinimo.Visible = false;
			this.txtValorMinimo.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtValorMinimo_KeyPress);
			// 
			// label2
			// 
			this.label2.AutoSize = true;
			this.label2.BackColor = System.Drawing.Color.Transparent;
			this.label2.Enabled = false;
			this.label2.Location = new System.Drawing.Point(784, 144);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(75, 16);
			this.label2.TabIndex = 270;
			this.label2.Text = "Precio mínimo";
			this.label2.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.label2.Visible = false;
			// 
			// cmbTipo
			// 
			appearance64.FontData.Name = "Tahoma";
			appearance64.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbTipo.Appearance = appearance64;
			this.cmbTipo.DropDownStyle = Infragistics.Win.DropDownStyle.DropDownList;
			this.cmbTipo.Enabled = false;
			valueListItem1.DataValue = "ValueListItem1";
			valueListItem1.DisplayText = "COMBO";
			valueListItem2.DataValue = "ValueListItem0";
			valueListItem2.DisplayText = "PROMOCIÓN";
			valueListItem3.DataValue = "ValueListItem3";
			valueListItem3.DisplayText = "BARATAZO";
			valueListItem4.DataValue = "ValueListItem3";
			valueListItem4.DisplayText = "COMBO PAYJOY";
			this.cmbTipo.Items.Add(valueListItem1);
			this.cmbTipo.Items.Add(valueListItem2);
			this.cmbTipo.Items.Add(valueListItem3);
			this.cmbTipo.Items.Add(valueListItem4);
			this.cmbTipo.Location = new System.Drawing.Point(64, 7);
			this.cmbTipo.Name = "cmbTipo";
			this.cmbTipo.Size = new System.Drawing.Size(144, 22);
			this.cmbTipo.TabIndex = 277;
			this.cmbTipo.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.cmbTipo_KeyPress);
			this.cmbTipo.ValueChanged += new System.EventHandler(this.cmbTipo_ValueChanged);
			// 
			// label16
			// 
			this.label16.AutoSize = true;
			this.label16.Enabled = false;
			this.label16.Location = new System.Drawing.Point(8, 10);
			this.label16.Name = "label16";
			this.label16.Size = new System.Drawing.Size(26, 16);
			this.label16.TabIndex = 278;
			this.label16.Text = "Tipo";
			this.label16.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// uGridTarjetas
			// 
			this.uGridTarjetas.Cursor = System.Windows.Forms.Cursors.Default;
			this.uGridTarjetas.DataSource = this.ultraDataSource6;
			appearance65.BackColor = System.Drawing.Color.White;
			this.uGridTarjetas.DisplayLayout.Appearance = appearance65;
			this.uGridTarjetas.DisplayLayout.AutoFitColumns = true;
			ultraGridColumn59.Header.VisiblePosition = 0;
			ultraGridColumn59.Hidden = true;
			ultraGridColumn59.Width = 60;
			ultraGridColumn60.Header.VisiblePosition = 1;
			ultraGridColumn60.Hidden = true;
			ultraGridColumn60.Width = 60;
			ultraGridColumn61.Header.VisiblePosition = 2;
			ultraGridColumn61.Hidden = true;
			ultraGridColumn61.Width = 73;
			ultraGridColumn62.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn62.Header.VisiblePosition = 3;
			ultraGridColumn62.Width = 181;
			ultraGridColumn63.Header.Caption = "...";
			ultraGridColumn63.Header.VisiblePosition = 4;
			ultraGridColumn63.Width = 30;
			ultraGridBand2.Columns.AddRange(new object[] {
																										 ultraGridColumn59,
																										 ultraGridColumn60,
																										 ultraGridColumn61,
																										 ultraGridColumn62,
																										 ultraGridColumn63});
			ultraGridBand2.Header.Caption = "TARJETAS";
			ultraGridBand2.HeaderVisible = true;
			ultraGridBand2.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			appearance66.FontData.SizeInPoints = 7.25F;
			ultraGridBand2.Override.HeaderAppearance = appearance66;
			this.uGridTarjetas.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			appearance67.ForeColor = System.Drawing.Color.Black;
			appearance67.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridTarjetas.DisplayLayout.Override.ActiveRowAppearance = appearance67;
			this.uGridTarjetas.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.uGridTarjetas.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance68.BackColor = System.Drawing.Color.Transparent;
			this.uGridTarjetas.DisplayLayout.Override.CardAreaAppearance = appearance68;
			appearance69.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance69.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance69.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance69.FontData.BoldAsString = "True";
			appearance69.FontData.Name = "Arial";
			appearance69.FontData.SizeInPoints = 10F;
			appearance69.ForeColor = System.Drawing.Color.White;
			appearance69.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridTarjetas.DisplayLayout.Override.HeaderAppearance = appearance69;
			appearance70.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance70.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance70.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridTarjetas.DisplayLayout.Override.RowAlternateAppearance = appearance70;
			appearance71.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance71.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance71.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridTarjetas.DisplayLayout.Override.RowSelectorAppearance = appearance71;
			appearance72.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance72.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance72.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridTarjetas.DisplayLayout.Override.SelectedRowAppearance = appearance72;
			this.uGridTarjetas.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridTarjetas.Location = new System.Drawing.Point(248, 208);
			this.uGridTarjetas.Name = "uGridTarjetas";
			this.uGridTarjetas.Size = new System.Drawing.Size(232, 150);
			this.uGridTarjetas.TabIndex = 280;
			this.uGridTarjetas.KeyDown += new System.Windows.Forms.KeyEventHandler(this.uGridTarjetas_KeyDown);
			this.uGridTarjetas.AfterCellUpdate += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.uGridTarjetas_AfterCellUpdate);
			this.uGridTarjetas.Click += new System.EventHandler(this.uGridTarjetas_Click);
			// 
			// ultraDataSource6
			// 
			ultraDataColumn59.DataType = typeof(int);
			ultraDataColumn60.DataType = typeof(int);
			ultraDataColumn61.DataType = typeof(int);
			ultraDataColumn63.DataType = typeof(bool);
			ultraDataColumn63.DefaultValue = false;
			this.ultraDataSource6.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn59,
																																 ultraDataColumn60,
																																 ultraDataColumn61,
																																 ultraDataColumn62,
																																 ultraDataColumn63});
			// 
			// uGridCiudades
			// 
			this.uGridCiudades.Cursor = System.Windows.Forms.Cursors.Hand;
			this.uGridCiudades.DataSource = this.ultraDataSource2;
			appearance73.BackColor = System.Drawing.Color.White;
			this.uGridCiudades.DisplayLayout.Appearance = appearance73;
			this.uGridCiudades.DisplayLayout.AutoFitColumns = true;
			ultraGridColumn64.Header.VisiblePosition = 0;
			ultraGridColumn64.Hidden = true;
			ultraGridColumn64.Width = 71;
			ultraGridColumn65.Header.VisiblePosition = 1;
			ultraGridColumn65.Hidden = true;
			ultraGridColumn65.Width = 66;
			ultraGridColumn66.Header.VisiblePosition = 2;
			ultraGridColumn66.Hidden = true;
			ultraGridColumn66.Width = 70;
			ultraGridColumn67.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn67.Header.VisiblePosition = 3;
			ultraGridColumn67.Width = 180;
			ultraGridColumn68.Header.Caption = "...";
			ultraGridColumn68.Header.VisiblePosition = 4;
			ultraGridColumn68.Width = 31;
			ultraGridBand3.Columns.AddRange(new object[] {
																										 ultraGridColumn64,
																										 ultraGridColumn65,
																										 ultraGridColumn66,
																										 ultraGridColumn67,
																										 ultraGridColumn68});
			ultraGridBand3.Header.Caption = "CIUDADES";
			ultraGridBand3.HeaderVisible = true;
			ultraGridBand3.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			appearance74.FontData.SizeInPoints = 7.25F;
			ultraGridBand3.Override.HeaderAppearance = appearance74;
			this.uGridCiudades.DisplayLayout.BandsSerializer.Add(ultraGridBand3);
			appearance75.ForeColor = System.Drawing.Color.Black;
			appearance75.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridCiudades.DisplayLayout.Override.ActiveRowAppearance = appearance75;
			this.uGridCiudades.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.uGridCiudades.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance76.BackColor = System.Drawing.Color.Transparent;
			this.uGridCiudades.DisplayLayout.Override.CardAreaAppearance = appearance76;
			appearance77.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance77.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance77.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance77.FontData.BoldAsString = "True";
			appearance77.FontData.Name = "Arial";
			appearance77.FontData.SizeInPoints = 10F;
			appearance77.ForeColor = System.Drawing.Color.White;
			appearance77.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridCiudades.DisplayLayout.Override.HeaderAppearance = appearance77;
			appearance78.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance78.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance78.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridCiudades.DisplayLayout.Override.RowAlternateAppearance = appearance78;
			appearance79.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance79.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance79.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridCiudades.DisplayLayout.Override.RowSelectorAppearance = appearance79;
			appearance80.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance80.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance80.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridCiudades.DisplayLayout.Override.SelectedRowAppearance = appearance80;
			this.uGridCiudades.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridCiudades.Location = new System.Drawing.Point(8, 56);
			this.uGridCiudades.Name = "uGridCiudades";
			this.uGridCiudades.Size = new System.Drawing.Size(232, 150);
			this.uGridCiudades.TabIndex = 281;
			this.uGridCiudades.UpdateMode = Infragistics.Win.UltraWinGrid.UpdateMode.OnCellChange;
			this.uGridCiudades.KeyDown += new System.Windows.Forms.KeyEventHandler(this.uGridCiudades_KeyDown);
			this.uGridCiudades.AfterCellUpdate += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.uGridCiudades_AfterCellUpdate);
			this.uGridCiudades.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.uGridCiudades_InitializeLayout);
			this.uGridCiudades.Click += new System.EventHandler(this.uGridCiudades_Click);
			// 
			// ultraDataSource2
			// 
			ultraDataColumn64.DataType = typeof(int);
			ultraDataColumn65.DataType = typeof(int);
			ultraDataColumn66.DataType = typeof(int);
			ultraDataColumn68.DataType = typeof(bool);
			ultraDataColumn68.DefaultValue = false;
			this.ultraDataSource2.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn64,
																																 ultraDataColumn65,
																																 ultraDataColumn66,
																																 ultraDataColumn67,
																																 ultraDataColumn68});
			// 
			// uGridBodegas
			// 
			this.uGridBodegas.Cursor = System.Windows.Forms.Cursors.Default;
			this.uGridBodegas.DataSource = this.ultraDataSource3;
			appearance81.BackColor = System.Drawing.Color.White;
			this.uGridBodegas.DisplayLayout.Appearance = appearance81;
			this.uGridBodegas.DisplayLayout.AutoFitColumns = true;
			ultraGridColumn69.Header.VisiblePosition = 0;
			ultraGridColumn69.Hidden = true;
			ultraGridColumn69.Width = 72;
			ultraGridColumn70.Header.VisiblePosition = 1;
			ultraGridColumn70.Hidden = true;
			ultraGridColumn70.Width = 66;
			ultraGridColumn71.Header.VisiblePosition = 2;
			ultraGridColumn71.Hidden = true;
			ultraGridColumn71.Width = 69;
			ultraGridColumn72.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn72.Header.VisiblePosition = 3;
			ultraGridColumn72.Width = 181;
			ultraGridColumn73.Header.Caption = "...";
			ultraGridColumn73.Header.VisiblePosition = 4;
			ultraGridColumn73.Width = 30;
			ultraGridBand4.Columns.AddRange(new object[] {
																										 ultraGridColumn69,
																										 ultraGridColumn70,
																										 ultraGridColumn71,
																										 ultraGridColumn72,
																										 ultraGridColumn73});
			ultraGridBand4.Header.Caption = "LOCALES";
			ultraGridBand4.HeaderVisible = true;
			ultraGridBand4.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			appearance82.FontData.SizeInPoints = 7.25F;
			ultraGridBand4.Override.HeaderAppearance = appearance82;
			this.uGridBodegas.DisplayLayout.BandsSerializer.Add(ultraGridBand4);
			appearance83.ForeColor = System.Drawing.Color.Black;
			appearance83.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridBodegas.DisplayLayout.Override.ActiveRowAppearance = appearance83;
			this.uGridBodegas.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.uGridBodegas.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance84.BackColor = System.Drawing.Color.Transparent;
			this.uGridBodegas.DisplayLayout.Override.CardAreaAppearance = appearance84;
			appearance85.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance85.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance85.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance85.FontData.BoldAsString = "True";
			appearance85.FontData.Name = "Arial";
			appearance85.FontData.SizeInPoints = 10F;
			appearance85.ForeColor = System.Drawing.Color.White;
			appearance85.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridBodegas.DisplayLayout.Override.HeaderAppearance = appearance85;
			appearance86.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance86.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance86.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridBodegas.DisplayLayout.Override.RowAlternateAppearance = appearance86;
			appearance87.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance87.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance87.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridBodegas.DisplayLayout.Override.RowSelectorAppearance = appearance87;
			appearance88.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance88.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance88.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridBodegas.DisplayLayout.Override.SelectedRowAppearance = appearance88;
			this.uGridBodegas.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridBodegas.Location = new System.Drawing.Point(248, 56);
			this.uGridBodegas.Name = "uGridBodegas";
			this.uGridBodegas.Size = new System.Drawing.Size(232, 150);
			this.uGridBodegas.TabIndex = 282;
			this.uGridBodegas.KeyDown += new System.Windows.Forms.KeyEventHandler(this.uGridBodegas_KeyDown);
			this.uGridBodegas.AfterCellUpdate += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.uGridBodegas_AfterCellUpdate);
			this.uGridBodegas.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.uGridBodegas_InitializeLayout);
			// 
			// ultraDataSource3
			// 
			ultraDataColumn69.DataType = typeof(int);
			ultraDataColumn70.DataType = typeof(int);
			ultraDataColumn71.DataType = typeof(int);
			ultraDataColumn73.DataType = typeof(bool);
			ultraDataColumn73.DefaultValue = false;
			this.ultraDataSource3.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn69,
																																 ultraDataColumn70,
																																 ultraDataColumn71,
																																 ultraDataColumn72,
																																 ultraDataColumn73});
			// 
			// uGridTiposDePrecio
			// 
			this.uGridTiposDePrecio.Cursor = System.Windows.Forms.Cursors.Default;
			this.uGridTiposDePrecio.DataSource = this.ultraDataSource4;
			appearance89.BackColor = System.Drawing.Color.White;
			this.uGridTiposDePrecio.DisplayLayout.Appearance = appearance89;
			this.uGridTiposDePrecio.DisplayLayout.AutoFitColumns = true;
			ultraGridColumn74.Header.VisiblePosition = 0;
			ultraGridColumn74.Hidden = true;
			ultraGridColumn74.Width = 35;
			ultraGridColumn75.Header.VisiblePosition = 1;
			ultraGridColumn75.Hidden = true;
			ultraGridColumn75.Width = 39;
			ultraGridColumn76.Header.VisiblePosition = 2;
			ultraGridColumn76.Hidden = true;
			ultraGridColumn76.Width = 46;
			ultraGridColumn77.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn77.Header.VisiblePosition = 3;
			ultraGridColumn77.Width = 115;
			appearance90.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn78.CellAppearance = appearance90;
			ultraGridColumn78.FieldLen = 1;
			ultraGridColumn78.Header.Caption = "Tipo";
			ultraGridColumn78.Header.VisiblePosition = 4;
			ultraGridColumn78.Width = 39;
			appearance91.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn79.CellAppearance = appearance91;
			ultraGridColumn79.Format = "#,##0.00";
			ultraGridColumn79.Header.Caption = "% D ";
			ultraGridColumn79.Header.VisiblePosition = 5;
			ultraGridColumn79.Width = 40;
			appearance92.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn80.CellAppearance = appearance92;
			ultraGridColumn80.Format = "#,##0.00";
			ultraGridColumn80.Header.Caption = "% H";
			ultraGridColumn80.Header.VisiblePosition = 6;
			ultraGridColumn80.MaskInput = "{LOC}nn,nnn,nnn.nn";
			ultraGridColumn80.Width = 40;
			ultraGridColumn81.Header.Caption = "...";
			ultraGridColumn81.Header.VisiblePosition = 7;
			ultraGridColumn81.Width = 25;
			ultraGridBand5.Columns.AddRange(new object[] {
																										 ultraGridColumn74,
																										 ultraGridColumn75,
																										 ultraGridColumn76,
																										 ultraGridColumn77,
																										 ultraGridColumn78,
																										 ultraGridColumn79,
																										 ultraGridColumn80,
																										 ultraGridColumn81});
			ultraGridBand5.Header.Caption = "PRECIOS";
			ultraGridBand5.HeaderVisible = true;
			ultraGridBand5.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			appearance93.FontData.SizeInPoints = 7.25F;
			ultraGridBand5.Override.HeaderAppearance = appearance93;
			this.uGridTiposDePrecio.DisplayLayout.BandsSerializer.Add(ultraGridBand5);
			appearance94.ForeColor = System.Drawing.Color.Black;
			appearance94.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridTiposDePrecio.DisplayLayout.Override.ActiveRowAppearance = appearance94;
			this.uGridTiposDePrecio.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.uGridTiposDePrecio.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance95.BackColor = System.Drawing.Color.Transparent;
			this.uGridTiposDePrecio.DisplayLayout.Override.CardAreaAppearance = appearance95;
			appearance96.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance96.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance96.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance96.FontData.BoldAsString = "True";
			appearance96.FontData.Name = "Arial";
			appearance96.FontData.SizeInPoints = 10F;
			appearance96.ForeColor = System.Drawing.Color.White;
			appearance96.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridTiposDePrecio.DisplayLayout.Override.HeaderAppearance = appearance96;
			appearance97.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance97.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance97.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridTiposDePrecio.DisplayLayout.Override.RowAlternateAppearance = appearance97;
			appearance98.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance98.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance98.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridTiposDePrecio.DisplayLayout.Override.RowSelectorAppearance = appearance98;
			appearance99.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance99.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance99.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridTiposDePrecio.DisplayLayout.Override.SelectedRowAppearance = appearance99;
			this.uGridTiposDePrecio.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridTiposDePrecio.Location = new System.Drawing.Point(488, 56);
			this.uGridTiposDePrecio.Name = "uGridTiposDePrecio";
			this.uGridTiposDePrecio.Size = new System.Drawing.Size(280, 150);
			this.uGridTiposDePrecio.TabIndex = 283;
			this.uGridTiposDePrecio.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.uGridTiposDePrecio_KeyPress);
			this.uGridTiposDePrecio.KeyDown += new System.Windows.Forms.KeyEventHandler(this.uGridTiposDePrecio_KeyDown);
			this.uGridTiposDePrecio.AfterCellUpdate += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.uGridTiposDePrecio_AfterCellUpdate);
			this.uGridTiposDePrecio.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.uGridTiposDePrecio_InitializeLayout);
			this.uGridTiposDePrecio.Click += new System.EventHandler(this.uGridTiposDePrecio_Click);
			// 
			// ultraDataSource4
			// 
			ultraDataColumn74.DataType = typeof(int);
			ultraDataColumn75.DataType = typeof(int);
			ultraDataColumn76.DataType = typeof(int);
			ultraDataColumn79.DataType = typeof(System.Decimal);
			ultraDataColumn79.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
			ultraDataColumn80.DataType = typeof(System.Decimal);
			ultraDataColumn80.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
			ultraDataColumn81.DataType = typeof(bool);
			ultraDataColumn81.DefaultValue = false;
			this.ultraDataSource4.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn74,
																																 ultraDataColumn75,
																																 ultraDataColumn76,
																																 ultraDataColumn77,
																																 ultraDataColumn78,
																																 ultraDataColumn79,
																																 ultraDataColumn80,
																																 ultraDataColumn81});
			// 
			// uGridMarcas
			// 
			this.uGridMarcas.Cursor = System.Windows.Forms.Cursors.Default;
			this.uGridMarcas.DataSource = this.ultraDataSource9;
			appearance100.BackColor = System.Drawing.Color.White;
			this.uGridMarcas.DisplayLayout.Appearance = appearance100;
			this.uGridMarcas.DisplayLayout.AutoFitColumns = true;
			ultraGridColumn82.Header.VisiblePosition = 0;
			ultraGridColumn82.Hidden = true;
			ultraGridColumn83.Header.VisiblePosition = 1;
			ultraGridColumn83.Hidden = true;
			ultraGridColumn84.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn84.Header.VisiblePosition = 2;
			ultraGridColumn84.Width = 181;
			ultraGridColumn85.Header.Caption = "...";
			ultraGridColumn85.Header.VisiblePosition = 3;
			ultraGridColumn85.Width = 30;
			ultraGridBand6.Columns.AddRange(new object[] {
																										 ultraGridColumn82,
																										 ultraGridColumn83,
																										 ultraGridColumn84,
																										 ultraGridColumn85});
			ultraGridBand6.Header.Caption = "MARCAS DE ARTÍCULOS";
			ultraGridBand6.HeaderVisible = true;
			ultraGridBand6.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			appearance101.FontData.SizeInPoints = 7.25F;
			ultraGridBand6.Override.HeaderAppearance = appearance101;
			this.uGridMarcas.DisplayLayout.BandsSerializer.Add(ultraGridBand6);
			appearance102.ForeColor = System.Drawing.Color.Black;
			appearance102.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridMarcas.DisplayLayout.Override.ActiveRowAppearance = appearance102;
			this.uGridMarcas.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.uGridMarcas.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance103.BackColor = System.Drawing.Color.Transparent;
			this.uGridMarcas.DisplayLayout.Override.CardAreaAppearance = appearance103;
			appearance104.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance104.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance104.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance104.FontData.BoldAsString = "True";
			appearance104.FontData.Name = "Arial";
			appearance104.FontData.SizeInPoints = 10F;
			appearance104.ForeColor = System.Drawing.Color.White;
			appearance104.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridMarcas.DisplayLayout.Override.HeaderAppearance = appearance104;
			appearance105.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance105.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance105.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridMarcas.DisplayLayout.Override.RowAlternateAppearance = appearance105;
			appearance106.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance106.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance106.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridMarcas.DisplayLayout.Override.RowSelectorAppearance = appearance106;
			appearance107.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance107.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance107.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridMarcas.DisplayLayout.Override.SelectedRowAppearance = appearance107;
			this.uGridMarcas.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridMarcas.Location = new System.Drawing.Point(488, 360);
			this.uGridMarcas.Name = "uGridMarcas";
			this.uGridMarcas.Size = new System.Drawing.Size(232, 184);
			this.uGridMarcas.TabIndex = 286;
			this.uGridMarcas.CellChange += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.uGridMarcas_CellChange);
			this.uGridMarcas.KeyDown += new System.Windows.Forms.KeyEventHandler(this.uGridMarcas_KeyDown);
			this.uGridMarcas.AfterCellUpdate += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.uGridMarcas_AfterCellUpdate);
			this.uGridMarcas.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.uGridMarcas_InitializeLayout);
			this.uGridMarcas.Click += new System.EventHandler(this.uGridMarcas_Click);
			// 
			// ultraDataSource9
			// 
			ultraDataColumn82.DataType = typeof(int);
			ultraDataColumn83.DataType = typeof(int);
			ultraDataColumn85.DataType = typeof(bool);
			ultraDataColumn85.DefaultValue = false;
			this.ultraDataSource9.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn82,
																																 ultraDataColumn83,
																																 ultraDataColumn84,
																																 ultraDataColumn85});
			// 
			// uGridSubgrupos
			// 
			this.uGridSubgrupos.Cursor = System.Windows.Forms.Cursors.Default;
			this.uGridSubgrupos.DataSource = this.ultraDataSource8;
			appearance108.BackColor = System.Drawing.Color.White;
			this.uGridSubgrupos.DisplayLayout.Appearance = appearance108;
			this.uGridSubgrupos.DisplayLayout.AutoFitColumns = true;
			ultraGridColumn86.Header.VisiblePosition = 0;
			ultraGridColumn86.Hidden = true;
			ultraGridColumn87.Header.VisiblePosition = 1;
			ultraGridColumn87.Hidden = true;
			ultraGridColumn88.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn88.Header.VisiblePosition = 2;
			ultraGridColumn88.Width = 181;
			ultraGridColumn89.Header.Caption = "...";
			ultraGridColumn89.Header.VisiblePosition = 3;
			ultraGridColumn89.Width = 30;
			ultraGridBand7.Columns.AddRange(new object[] {
																										 ultraGridColumn86,
																										 ultraGridColumn87,
																										 ultraGridColumn88,
																										 ultraGridColumn89});
			ultraGridBand7.Header.Caption = "SUBGRUPOS DE ARTÍCULOS";
			ultraGridBand7.HeaderVisible = true;
			ultraGridBand7.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			appearance109.FontData.SizeInPoints = 7.25F;
			ultraGridBand7.Override.HeaderAppearance = appearance109;
			this.uGridSubgrupos.DisplayLayout.BandsSerializer.Add(ultraGridBand7);
			appearance110.ForeColor = System.Drawing.Color.Black;
			appearance110.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridSubgrupos.DisplayLayout.Override.ActiveRowAppearance = appearance110;
			this.uGridSubgrupos.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.uGridSubgrupos.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance111.BackColor = System.Drawing.Color.Transparent;
			this.uGridSubgrupos.DisplayLayout.Override.CardAreaAppearance = appearance111;
			appearance112.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance112.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance112.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance112.FontData.BoldAsString = "True";
			appearance112.FontData.Name = "Arial";
			appearance112.FontData.SizeInPoints = 10F;
			appearance112.ForeColor = System.Drawing.Color.White;
			appearance112.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridSubgrupos.DisplayLayout.Override.HeaderAppearance = appearance112;
			appearance113.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance113.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance113.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridSubgrupos.DisplayLayout.Override.RowAlternateAppearance = appearance113;
			appearance114.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance114.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance114.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridSubgrupos.DisplayLayout.Override.RowSelectorAppearance = appearance114;
			appearance115.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance115.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance115.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridSubgrupos.DisplayLayout.Override.SelectedRowAppearance = appearance115;
			this.uGridSubgrupos.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridSubgrupos.Location = new System.Drawing.Point(248, 360);
			this.uGridSubgrupos.Name = "uGridSubgrupos";
			this.uGridSubgrupos.Size = new System.Drawing.Size(232, 184);
			this.uGridSubgrupos.TabIndex = 285;
			this.uGridSubgrupos.CellChange += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.uGridSubgrupos_CellChange);
			this.uGridSubgrupos.KeyDown += new System.Windows.Forms.KeyEventHandler(this.uGridSubgrupos_KeyDown);
			this.uGridSubgrupos.AfterCellUpdate += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.uGridSubgrupos_AfterCellUpdate);
			this.uGridSubgrupos.Click += new System.EventHandler(this.uGridSubgrupos_Click);
			// 
			// ultraDataSource8
			// 
			ultraDataColumn86.DataType = typeof(int);
			ultraDataColumn87.DataType = typeof(int);
			ultraDataColumn89.DataType = typeof(bool);
			ultraDataColumn89.DefaultValue = false;
			this.ultraDataSource8.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn86,
																																 ultraDataColumn87,
																																 ultraDataColumn88,
																																 ultraDataColumn89});
			// 
			// uGridGrupos
			// 
			this.uGridGrupos.Cursor = System.Windows.Forms.Cursors.Default;
			this.uGridGrupos.DataSource = this.ultraDataSource7;
			appearance116.BackColor = System.Drawing.Color.White;
			this.uGridGrupos.DisplayLayout.Appearance = appearance116;
			this.uGridGrupos.DisplayLayout.AutoFitColumns = true;
			ultraGridColumn90.Header.VisiblePosition = 0;
			ultraGridColumn90.Hidden = true;
			ultraGridColumn90.Width = 72;
			ultraGridColumn91.Header.VisiblePosition = 1;
			ultraGridColumn91.Hidden = true;
			ultraGridColumn91.Width = 20;
			ultraGridColumn92.Header.VisiblePosition = 2;
			ultraGridColumn92.Hidden = true;
			ultraGridColumn92.Width = 65;
			ultraGridColumn93.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn93.Header.VisiblePosition = 3;
			ultraGridColumn93.Width = 181;
			ultraGridColumn94.Header.Caption = "...";
			ultraGridColumn94.Header.VisiblePosition = 4;
			ultraGridColumn94.Width = 30;
			ultraGridBand8.Columns.AddRange(new object[] {
																										 ultraGridColumn90,
																										 ultraGridColumn91,
																										 ultraGridColumn92,
																										 ultraGridColumn93,
																										 ultraGridColumn94});
			ultraGridBand8.Header.Caption = "GRUPOS DE ARTÍCULOS";
			ultraGridBand8.HeaderVisible = true;
			ultraGridBand8.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			appearance117.FontData.SizeInPoints = 7.25F;
			ultraGridBand8.Override.HeaderAppearance = appearance117;
			this.uGridGrupos.DisplayLayout.BandsSerializer.Add(ultraGridBand8);
			appearance118.ForeColor = System.Drawing.Color.Black;
			appearance118.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridGrupos.DisplayLayout.Override.ActiveRowAppearance = appearance118;
			this.uGridGrupos.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.uGridGrupos.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance119.BackColor = System.Drawing.Color.Transparent;
			this.uGridGrupos.DisplayLayout.Override.CardAreaAppearance = appearance119;
			appearance120.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance120.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance120.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance120.FontData.BoldAsString = "True";
			appearance120.FontData.Name = "Arial";
			appearance120.FontData.SizeInPoints = 10F;
			appearance120.ForeColor = System.Drawing.Color.White;
			appearance120.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridGrupos.DisplayLayout.Override.HeaderAppearance = appearance120;
			appearance121.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance121.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance121.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridGrupos.DisplayLayout.Override.RowAlternateAppearance = appearance121;
			appearance122.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance122.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance122.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridGrupos.DisplayLayout.Override.RowSelectorAppearance = appearance122;
			appearance123.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance123.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance123.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridGrupos.DisplayLayout.Override.SelectedRowAppearance = appearance123;
			this.uGridGrupos.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridGrupos.Location = new System.Drawing.Point(8, 360);
			this.uGridGrupos.Name = "uGridGrupos";
			this.uGridGrupos.Size = new System.Drawing.Size(232, 184);
			this.uGridGrupos.TabIndex = 284;
			this.uGridGrupos.CellChange += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.uGridGrupos_CellChange);
			this.uGridGrupos.KeyDown += new System.Windows.Forms.KeyEventHandler(this.uGridGrupos_KeyDown);
			this.uGridGrupos.AfterCellUpdate += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.uGridGrupos_AfterCellUpdate);
			this.uGridGrupos.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.uGridGrupos_InitializeLayout);
			this.uGridGrupos.Click += new System.EventHandler(this.uGridGrupos_Click);
			// 
			// ultraDataSource7
			// 
			ultraDataColumn90.DataType = typeof(int);
			ultraDataColumn91.DataType = typeof(int);
			ultraDataColumn92.DataType = typeof(int);
			ultraDataColumn94.DataType = typeof(bool);
			ultraDataColumn94.DefaultValue = false;
			this.ultraDataSource7.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn90,
																																 ultraDataColumn91,
																																 ultraDataColumn92,
																																 ultraDataColumn93,
																																 ultraDataColumn94});
			// 
			// uGridBancos
			// 
			this.uGridBancos.Cursor = System.Windows.Forms.Cursors.Default;
			this.uGridBancos.DataSource = this.ultraDataSource5;
			appearance124.BackColor = System.Drawing.Color.White;
			this.uGridBancos.DisplayLayout.Appearance = appearance124;
			this.uGridBancos.DisplayLayout.AutoFitColumns = true;
			ultraGridColumn95.Header.VisiblePosition = 0;
			ultraGridColumn95.Hidden = true;
			ultraGridColumn95.Width = 70;
			ultraGridColumn96.Header.VisiblePosition = 1;
			ultraGridColumn96.Hidden = true;
			ultraGridColumn96.Width = 72;
			ultraGridColumn97.Header.VisiblePosition = 2;
			ultraGridColumn97.Hidden = true;
			ultraGridColumn97.Width = 75;
			ultraGridColumn98.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn98.Header.VisiblePosition = 3;
			ultraGridColumn98.Width = 181;
			ultraGridColumn99.Header.Caption = "...";
			ultraGridColumn99.Header.VisiblePosition = 4;
			ultraGridColumn99.Width = 30;
			ultraGridBand9.Columns.AddRange(new object[] {
																										 ultraGridColumn95,
																										 ultraGridColumn96,
																										 ultraGridColumn97,
																										 ultraGridColumn98,
																										 ultraGridColumn99});
			ultraGridBand9.Header.Caption = "BANCOS";
			ultraGridBand9.HeaderVisible = true;
			ultraGridBand9.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			appearance125.FontData.SizeInPoints = 7.25F;
			ultraGridBand9.Override.HeaderAppearance = appearance125;
			this.uGridBancos.DisplayLayout.BandsSerializer.Add(ultraGridBand9);
			appearance126.ForeColor = System.Drawing.Color.Black;
			appearance126.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridBancos.DisplayLayout.Override.ActiveRowAppearance = appearance126;
			this.uGridBancos.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.uGridBancos.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance127.BackColor = System.Drawing.Color.Transparent;
			this.uGridBancos.DisplayLayout.Override.CardAreaAppearance = appearance127;
			appearance128.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance128.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance128.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance128.FontData.BoldAsString = "True";
			appearance128.FontData.Name = "Arial";
			appearance128.FontData.SizeInPoints = 10F;
			appearance128.ForeColor = System.Drawing.Color.White;
			appearance128.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridBancos.DisplayLayout.Override.HeaderAppearance = appearance128;
			appearance129.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance129.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance129.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridBancos.DisplayLayout.Override.RowAlternateAppearance = appearance129;
			appearance130.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance130.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance130.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridBancos.DisplayLayout.Override.RowSelectorAppearance = appearance130;
			appearance131.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance131.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance131.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridBancos.DisplayLayout.Override.SelectedRowAppearance = appearance131;
			this.uGridBancos.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridBancos.Location = new System.Drawing.Point(8, 208);
			this.uGridBancos.Name = "uGridBancos";
			this.uGridBancos.Size = new System.Drawing.Size(232, 150);
			this.uGridBancos.TabIndex = 287;
			this.uGridBancos.KeyDown += new System.Windows.Forms.KeyEventHandler(this.uGridBancos_KeyDown);
			this.uGridBancos.AfterCellUpdate += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.uGridBancos_AfterCellUpdate);
			this.uGridBancos.Click += new System.EventHandler(this.uGridBancos_Click);
			// 
			// ultraDataSource5
			// 
			ultraDataColumn95.DataType = typeof(int);
			ultraDataColumn96.DataType = typeof(int);
			ultraDataColumn97.DataType = typeof(int);
			ultraDataColumn99.DataType = typeof(bool);
			ultraDataColumn99.DefaultValue = false;
			this.ultraDataSource5.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn95,
																																 ultraDataColumn96,
																																 ultraDataColumn97,
																																 ultraDataColumn98,
																																 ultraDataColumn99});
			// 
			// cdsPromocion
			// 
			this.cdsPromocion.BindingContextCtrl = this;
			this.cdsPromocion.DataLibrary = "LibFacturacion";
			this.cdsPromocion.DataLibraryUrl = "";
			this.cdsPromocion.DataSetDef = "dsPromocion";
			this.cdsPromocion.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsPromocion.Name = "cdsPromocion";
			this.cdsPromocion.SchemaClassName = "LibFacturacion.DataClass";
			this.cdsPromocion.SchemaDef = null;
			this.cdsPromocion.BeforeFill += new C1.Data.FillEventHandler(this.cdsPromocion_BeforeFill);
			// 
			// btNuevo
			// 
			this.btNuevo.Appearance.Options.UseTextOptions = true;
			this.btNuevo.Appearance.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Near;
			this.btNuevo.ImageAlignment = DevExpress.Utils.HorzAlignment.Far;
			this.btNuevo.ImageIndex = 8;
			this.btNuevo.ImageList = this.imageList1;
			this.btNuevo.Location = new System.Drawing.Point(984, 352);
			this.btNuevo.Name = "btNuevo";
			this.btNuevo.Size = new System.Drawing.Size(96, 24);
			this.btNuevo.TabIndex = 290;
			this.btNuevo.Text = "&Nuevo";
			this.btNuevo.Click += new System.EventHandler(this.btNuevo_Click);
			// 
			// imageList1
			// 
			this.imageList1.ImageSize = new System.Drawing.Size(16, 16);
			this.imageList1.ImageStream = ((System.Windows.Forms.ImageListStreamer)(resources.GetObject("imageList1.ImageStream")));
			this.imageList1.TransparentColor = System.Drawing.Color.Transparent;
			// 
			// btSalir
			// 
			this.btSalir.Appearance.Options.UseTextOptions = true;
			this.btSalir.Appearance.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Near;
			this.btSalir.ImageAlignment = DevExpress.Utils.HorzAlignment.Far;
			this.btSalir.ImageIndex = 62;
			this.btSalir.ImageList = this.imageList1;
			this.btSalir.Location = new System.Drawing.Point(984, 520);
			this.btSalir.Name = "btSalir";
			this.btSalir.Size = new System.Drawing.Size(96, 24);
			this.btSalir.TabIndex = 289;
			this.btSalir.Text = "&Salir";
			this.btSalir.Click += new System.EventHandler(this.btSalir_Click);
			// 
			// btEditar
			// 
			this.btEditar.Appearance.Options.UseTextOptions = true;
			this.btEditar.Appearance.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Near;
			this.btEditar.Enabled = false;
			this.btEditar.ImageAlignment = DevExpress.Utils.HorzAlignment.Far;
			this.btEditar.ImageIndex = 56;
			this.btEditar.ImageList = this.imageList1;
			this.btEditar.Location = new System.Drawing.Point(984, 400);
			this.btEditar.Name = "btEditar";
			this.btEditar.Size = new System.Drawing.Size(96, 24);
			this.btEditar.TabIndex = 288;
			this.btEditar.Text = "&Editar";
			this.btEditar.Click += new System.EventHandler(this.btEditar_Click);
			// 
			// btBuscar
			// 
			this.btBuscar.Appearance.Options.UseTextOptions = true;
			this.btBuscar.Appearance.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Near;
			this.btBuscar.ImageAlignment = DevExpress.Utils.HorzAlignment.Far;
			this.btBuscar.ImageIndex = 57;
			this.btBuscar.ImageList = this.imageList1;
			this.btBuscar.Location = new System.Drawing.Point(984, 376);
			this.btBuscar.Name = "btBuscar";
			this.btBuscar.Size = new System.Drawing.Size(96, 24);
			this.btBuscar.TabIndex = 292;
			this.btBuscar.Text = "&Buscar";
			this.btBuscar.Click += new System.EventHandler(this.btBuscar_Click);
			// 
			// btGrabar
			// 
			this.btGrabar.Appearance.Options.UseTextOptions = true;
			this.btGrabar.Appearance.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Near;
			this.btGrabar.Enabled = false;
			this.btGrabar.ImageAlignment = DevExpress.Utils.HorzAlignment.Far;
			this.btGrabar.ImageIndex = 9;
			this.btGrabar.ImageList = this.imageList1;
			this.btGrabar.Location = new System.Drawing.Point(984, 424);
			this.btGrabar.Name = "btGrabar";
			this.btGrabar.Size = new System.Drawing.Size(96, 24);
			this.btGrabar.TabIndex = 291;
			this.btGrabar.Text = "&Grabar";
			this.btGrabar.Click += new System.EventHandler(this.btGrabar_Click);
			// 
			// btCancelar
			// 
			this.btCancelar.Appearance.Options.UseTextOptions = true;
			this.btCancelar.Appearance.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Near;
			this.btCancelar.Enabled = false;
			this.btCancelar.ImageAlignment = DevExpress.Utils.HorzAlignment.Far;
			this.btCancelar.ImageIndex = 61;
			this.btCancelar.ImageList = this.imageList1;
			this.btCancelar.Location = new System.Drawing.Point(984, 472);
			this.btCancelar.Name = "btCancelar";
			this.btCancelar.Size = new System.Drawing.Size(96, 24);
			this.btCancelar.TabIndex = 293;
			this.btCancelar.Text = "&Cancelar";
			this.btCancelar.Click += new System.EventHandler(this.btCancelar_Click);
			// 
			// btFinalizar
			// 
			this.btFinalizar.Appearance.Options.UseTextOptions = true;
			this.btFinalizar.Appearance.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Near;
			this.btFinalizar.Enabled = false;
			this.btFinalizar.Image = ((System.Drawing.Image)(resources.GetObject("btFinalizar.Image")));
			this.btFinalizar.ImageAlignment = DevExpress.Utils.HorzAlignment.Far;
			this.btFinalizar.Location = new System.Drawing.Point(984, 448);
			this.btFinalizar.Name = "btFinalizar";
			this.btFinalizar.Size = new System.Drawing.Size(96, 24);
			this.btFinalizar.TabIndex = 297;
			this.btFinalizar.Text = "&Finalizar";
			this.btFinalizar.Click += new System.EventHandler(this.btFinalizar_Click);
			// 
			// lblFinalizado
			// 
			this.lblFinalizado.AutoSize = true;
			this.lblFinalizado.Font = new System.Drawing.Font("Tahoma", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblFinalizado.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(192)), ((System.Byte)(0)), ((System.Byte)(0)));
			this.lblFinalizado.Location = new System.Drawing.Point(968, 88);
			this.lblFinalizado.Name = "lblFinalizado";
			this.lblFinalizado.Size = new System.Drawing.Size(0, 23);
			this.lblFinalizado.TabIndex = 298;
			this.lblFinalizado.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// txtidArticulo
			// 
			this.txtidArticulo.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.txtidArticulo.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtidArticulo.Enabled = false;
			this.txtidArticulo.Location = new System.Drawing.Point(1072, 8);
			this.txtidArticulo.Name = "txtidArticulo";
			this.txtidArticulo.PromptChar = ' ';
			this.txtidArticulo.Size = new System.Drawing.Size(16, 22);
			this.txtidArticulo.TabIndex = 299;
			// 
			// btnActualizaPrecios
			// 
			this.btnActualizaPrecios.Appearance.Options.UseTextOptions = true;
			this.btnActualizaPrecios.Appearance.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Near;
			this.btnActualizaPrecios.Enabled = false;
			this.btnActualizaPrecios.Image = ((System.Drawing.Image)(resources.GetObject("btnActualizaPrecios.Image")));
			this.btnActualizaPrecios.ImageAlignment = DevExpress.Utils.HorzAlignment.Far;
			this.btnActualizaPrecios.Location = new System.Drawing.Point(984, 496);
			this.btnActualizaPrecios.Name = "btnActualizaPrecios";
			this.btnActualizaPrecios.Size = new System.Drawing.Size(96, 24);
			this.btnActualizaPrecios.TabIndex = 300;
			this.btnActualizaPrecios.Text = "&Precios";
			this.btnActualizaPrecios.Click += new System.EventHandler(this.btnActualizaPrecios_Click);
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
			// chkLocales
			// 
			appearance132.ForeColorDisabled = System.Drawing.Color.Black;
			this.chkLocales.Appearance = appearance132;
			this.chkLocales.BackColor = System.Drawing.Color.Transparent;
			this.chkLocales.CausesValidation = false;
			this.chkLocales.CheckAlign = System.Drawing.ContentAlignment.MiddleCenter;
			this.chkLocales.Enabled = false;
			this.chkLocales.Location = new System.Drawing.Point(440, 80);
			this.chkLocales.Name = "chkLocales";
			this.chkLocales.Size = new System.Drawing.Size(13, 13);
			this.chkLocales.TabIndex = 301;
			this.chkLocales.Text = "Precio Manual";
			this.chkLocales.CheckedChanged += new System.EventHandler(this.chkLocales_CheckedChanged);
			// 
			// groupBox1
			// 
			this.groupBox1.Location = new System.Drawing.Point(2, 40);
			this.groupBox1.Name = "groupBox1";
			this.groupBox1.Size = new System.Drawing.Size(1167, 8);
			this.groupBox1.TabIndex = 303;
			this.groupBox1.TabStop = false;
			// 
			// chkGrupos
			// 
			appearance133.ForeColorDisabled = System.Drawing.Color.Black;
			this.chkGrupos.Appearance = appearance133;
			this.chkGrupos.BackColor = System.Drawing.Color.Transparent;
			this.chkGrupos.CausesValidation = false;
			this.chkGrupos.CheckAlign = System.Drawing.ContentAlignment.MiddleCenter;
			this.chkGrupos.Enabled = false;
			this.chkGrupos.Location = new System.Drawing.Point(201, 384);
			this.chkGrupos.Name = "chkGrupos";
			this.chkGrupos.Size = new System.Drawing.Size(13, 13);
			this.chkGrupos.TabIndex = 305;
			this.chkGrupos.Text = "Precio Manual";
			this.chkGrupos.CheckedChanged += new System.EventHandler(this.chkGrupos_CheckedChanged);
			// 
			// chkSubGrupos
			// 
			appearance134.ForeColorDisabled = System.Drawing.Color.Black;
			this.chkSubGrupos.Appearance = appearance134;
			this.chkSubGrupos.BackColor = System.Drawing.Color.Transparent;
			this.chkSubGrupos.CausesValidation = false;
			this.chkSubGrupos.CheckAlign = System.Drawing.ContentAlignment.MiddleCenter;
			this.chkSubGrupos.Enabled = false;
			this.chkSubGrupos.Location = new System.Drawing.Point(441, 384);
			this.chkSubGrupos.Name = "chkSubGrupos";
			this.chkSubGrupos.Size = new System.Drawing.Size(13, 13);
			this.chkSubGrupos.TabIndex = 306;
			this.chkSubGrupos.Text = "Precio Manual";
			this.chkSubGrupos.CheckedChanged += new System.EventHandler(this.chkSubGrupos_CheckedChanged);
			// 
			// chkMarcas
			// 
			appearance135.ForeColorDisabled = System.Drawing.Color.Black;
			this.chkMarcas.Appearance = appearance135;
			this.chkMarcas.BackColor = System.Drawing.Color.Transparent;
			this.chkMarcas.CausesValidation = false;
			this.chkMarcas.CheckAlign = System.Drawing.ContentAlignment.MiddleCenter;
			this.chkMarcas.Enabled = false;
			this.chkMarcas.Location = new System.Drawing.Point(681, 384);
			this.chkMarcas.Name = "chkMarcas";
			this.chkMarcas.Size = new System.Drawing.Size(13, 13);
			this.chkMarcas.TabIndex = 307;
			this.chkMarcas.Text = "Precio Manual";
			this.chkMarcas.CheckedChanged += new System.EventHandler(this.chkMarcas_CheckedChanged);
			// 
			// chkVariable
			// 
			appearance136.FontData.Name = "Tahoma";
			appearance136.ForeColorDisabled = System.Drawing.Color.Black;
			this.chkVariable.Appearance = appearance136;
			this.chkVariable.BackColor = System.Drawing.Color.Transparent;
			this.chkVariable.CausesValidation = false;
			this.chkVariable.Enabled = false;
			this.chkVariable.Location = new System.Drawing.Point(968, 200);
			this.chkVariable.Name = "chkVariable";
			this.chkVariable.Size = new System.Drawing.Size(136, 21);
			this.chkVariable.TabIndex = 308;
			this.chkVariable.Text = "% Variable";
			// 
			// uGridPlanes
			// 
			this.uGridPlanes.Cursor = System.Windows.Forms.Cursors.Default;
			this.uGridPlanes.DataSource = this.ultraDataSource10;
			appearance137.BackColor = System.Drawing.Color.White;
			this.uGridPlanes.DisplayLayout.Appearance = appearance137;
			this.uGridPlanes.DisplayLayout.AutoFitColumns = true;
			ultraGridColumn100.Header.VisiblePosition = 0;
			ultraGridColumn100.Hidden = true;
			ultraGridColumn100.Width = 64;
			ultraGridColumn101.Header.VisiblePosition = 1;
			ultraGridColumn101.Hidden = true;
			ultraGridColumn101.Width = 68;
			ultraGridColumn102.Header.VisiblePosition = 2;
			ultraGridColumn102.Hidden = true;
			ultraGridColumn102.Width = 48;
			ultraGridColumn103.Header.VisiblePosition = 3;
			ultraGridColumn103.Hidden = true;
			ultraGridColumn103.Width = 30;
			ultraGridColumn104.Header.VisiblePosition = 4;
			ultraGridColumn104.Hidden = true;
			ultraGridColumn104.Width = 41;
			ultraGridColumn105.Header.VisiblePosition = 5;
			ultraGridColumn105.Width = 181;
			ultraGridColumn106.Header.Caption = "...";
			ultraGridColumn106.Header.VisiblePosition = 6;
			ultraGridColumn106.Width = 30;
			ultraGridBand10.Columns.AddRange(new object[] {
																											ultraGridColumn100,
																											ultraGridColumn101,
																											ultraGridColumn102,
																											ultraGridColumn103,
																											ultraGridColumn104,
																											ultraGridColumn105,
																											ultraGridColumn106});
			ultraGridBand10.Header.Caption = "PLANES";
			ultraGridBand10.HeaderVisible = true;
			ultraGridBand10.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			appearance138.FontData.SizeInPoints = 7.25F;
			ultraGridBand10.Override.HeaderAppearance = appearance138;
			this.uGridPlanes.DisplayLayout.BandsSerializer.Add(ultraGridBand10);
			appearance139.ForeColor = System.Drawing.Color.Black;
			appearance139.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridPlanes.DisplayLayout.Override.ActiveRowAppearance = appearance139;
			this.uGridPlanes.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.uGridPlanes.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance140.BackColor = System.Drawing.Color.Transparent;
			this.uGridPlanes.DisplayLayout.Override.CardAreaAppearance = appearance140;
			appearance141.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance141.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance141.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance141.FontData.BoldAsString = "True";
			appearance141.FontData.Name = "Arial";
			appearance141.FontData.SizeInPoints = 10F;
			appearance141.ForeColor = System.Drawing.Color.White;
			appearance141.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridPlanes.DisplayLayout.Override.HeaderAppearance = appearance141;
			appearance142.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance142.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance142.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridPlanes.DisplayLayout.Override.RowAlternateAppearance = appearance142;
			appearance143.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance143.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance143.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridPlanes.DisplayLayout.Override.RowSelectorAppearance = appearance143;
			appearance144.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance144.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance144.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridPlanes.DisplayLayout.Override.SelectedRowAppearance = appearance144;
			this.uGridPlanes.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridPlanes.Location = new System.Drawing.Point(488, 208);
			this.uGridPlanes.Name = "uGridPlanes";
			this.uGridPlanes.Size = new System.Drawing.Size(232, 150);
			this.uGridPlanes.TabIndex = 309;
			this.uGridPlanes.AfterCellUpdate += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.uGridPlanes_AfterCellUpdate);
			this.uGridPlanes.Click += new System.EventHandler(this.uGridPlanes_Click);
			// 
			// ultraDataSource10
			// 
			ultraDataColumn100.DataType = typeof(int);
			ultraDataColumn101.DataType = typeof(int);
			ultraDataColumn102.DataType = typeof(int);
			ultraDataColumn103.DataType = typeof(int);
			ultraDataColumn104.DataType = typeof(int);
			ultraDataColumn106.DataType = typeof(bool);
			ultraDataColumn106.DefaultValue = false;
			this.ultraDataSource10.Band.Columns.AddRange(new object[] {
																																	ultraDataColumn100,
																																	ultraDataColumn101,
																																	ultraDataColumn102,
																																	ultraDataColumn103,
																																	ultraDataColumn104,
																																	ultraDataColumn105,
																																	ultraDataColumn106});
			// 
			// uGridModelos
			// 
			this.uGridModelos.Cursor = System.Windows.Forms.Cursors.Default;
			this.uGridModelos.DataSource = this.ultraDataSource11;
			appearance145.BackColor = System.Drawing.Color.White;
			this.uGridModelos.DisplayLayout.Appearance = appearance145;
			this.uGridModelos.DisplayLayout.AutoFitColumns = true;
			ultraGridColumn107.Header.VisiblePosition = 0;
			ultraGridColumn107.Hidden = true;
			ultraGridColumn107.Width = 103;
			ultraGridColumn108.Header.VisiblePosition = 1;
			ultraGridColumn108.Hidden = true;
			ultraGridColumn108.Width = 55;
			ultraGridColumn109.Header.VisiblePosition = 2;
			ultraGridColumn109.Width = 181;
			ultraGridColumn110.Header.Caption = "...";
			ultraGridColumn110.Header.VisiblePosition = 3;
			ultraGridColumn110.Width = 30;
			ultraGridBand11.Columns.AddRange(new object[] {
																											ultraGridColumn107,
																											ultraGridColumn108,
																											ultraGridColumn109,
																											ultraGridColumn110});
			ultraGridBand11.Header.Caption = "MODELOS DE ARTÍCULOS";
			ultraGridBand11.HeaderVisible = true;
			ultraGridBand11.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			appearance146.FontData.SizeInPoints = 7.25F;
			ultraGridBand11.Override.HeaderAppearance = appearance146;
			this.uGridModelos.DisplayLayout.BandsSerializer.Add(ultraGridBand11);
			appearance147.ForeColor = System.Drawing.Color.Black;
			appearance147.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridModelos.DisplayLayout.Override.ActiveRowAppearance = appearance147;
			this.uGridModelos.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.uGridModelos.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance148.BackColor = System.Drawing.Color.Transparent;
			this.uGridModelos.DisplayLayout.Override.CardAreaAppearance = appearance148;
			appearance149.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance149.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance149.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance149.FontData.BoldAsString = "True";
			appearance149.FontData.Name = "Arial";
			appearance149.FontData.SizeInPoints = 10F;
			appearance149.ForeColor = System.Drawing.Color.White;
			appearance149.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridModelos.DisplayLayout.Override.HeaderAppearance = appearance149;
			appearance150.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance150.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance150.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridModelos.DisplayLayout.Override.RowAlternateAppearance = appearance150;
			appearance151.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance151.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance151.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridModelos.DisplayLayout.Override.RowSelectorAppearance = appearance151;
			appearance152.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance152.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance152.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridModelos.DisplayLayout.Override.SelectedRowAppearance = appearance152;
			this.uGridModelos.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridModelos.Location = new System.Drawing.Point(728, 360);
			this.uGridModelos.Name = "uGridModelos";
			this.uGridModelos.Size = new System.Drawing.Size(232, 184);
			this.uGridModelos.TabIndex = 310;
			this.uGridModelos.CellChange += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.uGridModelos_CellChange);
			this.uGridModelos.KeyDown += new System.Windows.Forms.KeyEventHandler(this.uGridModelos_KeyDown);
			this.uGridModelos.AfterCellUpdate += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.uGridModelos_AfterCellUpdate);
			this.uGridModelos.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.uGridModelos_InitializeLayout);
			// 
			// ultraDataSource11
			// 
			ultraDataColumn107.DataType = typeof(int);
			ultraDataColumn108.DataType = typeof(int);
			ultraDataColumn110.DataType = typeof(bool);
			ultraDataColumn110.DefaultValue = false;
			this.ultraDataSource11.Band.Columns.AddRange(new object[] {
																																	ultraDataColumn107,
																																	ultraDataColumn108,
																																	ultraDataColumn109,
																																	ultraDataColumn110});
			// 
			// chkSumaDescuentos
			// 
			appearance153.FontData.Name = "Tahoma";
			appearance153.ForeColorDisabled = System.Drawing.Color.Black;
			this.chkSumaDescuentos.Appearance = appearance153;
			this.chkSumaDescuentos.BackColor = System.Drawing.Color.Transparent;
			this.chkSumaDescuentos.CausesValidation = false;
			this.chkSumaDescuentos.Enabled = false;
			this.chkSumaDescuentos.Location = new System.Drawing.Point(968, 224);
			this.chkSumaDescuentos.Name = "chkSumaDescuentos";
			this.chkSumaDescuentos.Size = new System.Drawing.Size(136, 21);
			this.chkSumaDescuentos.TabIndex = 311;
			this.chkSumaDescuentos.Text = "Suma Descuentos";
			// 
			// chkAdicional
			// 
			appearance154.FontData.Name = "Tahoma";
			appearance154.ForeColorDisabled = System.Drawing.Color.Black;
			this.chkAdicional.Appearance = appearance154;
			this.chkAdicional.BackColor = System.Drawing.Color.Transparent;
			this.chkAdicional.CausesValidation = false;
			this.chkAdicional.Enabled = false;
			this.chkAdicional.Location = new System.Drawing.Point(968, 248);
			this.chkAdicional.Name = "chkAdicional";
			this.chkAdicional.Size = new System.Drawing.Size(136, 21);
			this.chkAdicional.TabIndex = 313;
			this.chkAdicional.Text = "Adicional";
			// 
			// chkValor
			// 
			appearance155.FontData.Name = "Tahoma";
			appearance155.ForeColorDisabled = System.Drawing.Color.Black;
			this.chkValor.Appearance = appearance155;
			this.chkValor.BackColor = System.Drawing.Color.Transparent;
			this.chkValor.CausesValidation = false;
			this.chkValor.Enabled = false;
			this.chkValor.Location = new System.Drawing.Point(968, 272);
			this.chkValor.Name = "chkValor";
			this.chkValor.Size = new System.Drawing.Size(136, 21);
			this.chkValor.TabIndex = 312;
			this.chkValor.Text = "Valor";
			// 
			// uGridPlazos
			// 
			this.uGridPlazos.Cursor = System.Windows.Forms.Cursors.Default;
			this.uGridPlazos.DataSource = this.ultraDataSource12;
			appearance156.BackColor = System.Drawing.Color.White;
			this.uGridPlazos.DisplayLayout.Appearance = appearance156;
			this.uGridPlazos.DisplayLayout.AutoFitColumns = true;
			ultraGridColumn111.Header.VisiblePosition = 0;
			ultraGridColumn111.Hidden = true;
			ultraGridColumn111.Width = 60;
			ultraGridColumn112.Header.VisiblePosition = 1;
			ultraGridColumn112.Hidden = true;
			ultraGridColumn112.Width = 28;
			ultraGridColumn113.Header.VisiblePosition = 2;
			ultraGridColumn113.Hidden = true;
			ultraGridColumn113.Width = 41;
			ultraGridColumn114.Header.VisiblePosition = 3;
			ultraGridColumn114.Width = 129;
			ultraGridColumn115.Header.Caption = "...";
			ultraGridColumn115.Header.VisiblePosition = 5;
			ultraGridColumn115.Width = 22;
			appearance157.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn116.CellAppearance = appearance157;
			ultraGridColumn116.Format = "#,##0";
			ultraGridColumn116.Header.Caption = "M Gracia";
			ultraGridColumn116.Header.VisiblePosition = 4;
			ultraGridColumn116.Width = 60;
			ultraGridBand12.Columns.AddRange(new object[] {
																											ultraGridColumn111,
																											ultraGridColumn112,
																											ultraGridColumn113,
																											ultraGridColumn114,
																											ultraGridColumn115,
																											ultraGridColumn116});
			ultraGridBand12.Header.Caption = "PLAZOS";
			ultraGridBand12.HeaderVisible = true;
			ultraGridBand12.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			appearance158.FontData.SizeInPoints = 7.25F;
			ultraGridBand12.Override.HeaderAppearance = appearance158;
			this.uGridPlazos.DisplayLayout.BandsSerializer.Add(ultraGridBand12);
			appearance159.ForeColor = System.Drawing.Color.Black;
			appearance159.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridPlazos.DisplayLayout.Override.ActiveRowAppearance = appearance159;
			this.uGridPlazos.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.uGridPlazos.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance160.BackColor = System.Drawing.Color.Transparent;
			this.uGridPlazos.DisplayLayout.Override.CardAreaAppearance = appearance160;
			appearance161.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance161.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance161.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance161.FontData.BoldAsString = "True";
			appearance161.FontData.Name = "Arial";
			appearance161.FontData.SizeInPoints = 10F;
			appearance161.ForeColor = System.Drawing.Color.White;
			appearance161.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridPlazos.DisplayLayout.Override.HeaderAppearance = appearance161;
			appearance162.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance162.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance162.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridPlazos.DisplayLayout.Override.RowAlternateAppearance = appearance162;
			appearance163.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance163.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance163.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridPlazos.DisplayLayout.Override.RowSelectorAppearance = appearance163;
			appearance164.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance164.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance164.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridPlazos.DisplayLayout.Override.SelectedRowAppearance = appearance164;
			this.uGridPlazos.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridPlazos.Location = new System.Drawing.Point(728, 208);
			this.uGridPlazos.Name = "uGridPlazos";
			this.uGridPlazos.Size = new System.Drawing.Size(232, 150);
			this.uGridPlazos.TabIndex = 314;
			this.uGridPlazos.AfterCellUpdate += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.uGridPlazos_AfterCellUpdate);
			// 
			// ultraDataSource12
			// 
			ultraDataColumn111.DataType = typeof(int);
			ultraDataColumn111.DefaultValue = 0;
			ultraDataColumn112.DataType = typeof(int);
			ultraDataColumn112.DefaultValue = 0;
			ultraDataColumn113.DataType = typeof(int);
			ultraDataColumn113.DefaultValue = 0;
			ultraDataColumn115.DataType = typeof(bool);
			ultraDataColumn115.DefaultValue = false;
			ultraDataColumn116.DataType = typeof(int);
			this.ultraDataSource12.Band.Columns.AddRange(new object[] {
																																	ultraDataColumn111,
																																	ultraDataColumn112,
																																	ultraDataColumn113,
																																	ultraDataColumn114,
																																	ultraDataColumn115,
																																	ultraDataColumn116});
			// 
			// ultraDataSource13
			// 
			ultraDataColumn117.DataType = typeof(int);
			ultraDataColumn118.DataType = typeof(int);
			ultraDataColumn119.DataType = typeof(int);
			ultraDataColumn120.DataType = typeof(bool);
			ultraDataColumn120.DefaultValue = false;
			this.ultraDataSource13.Band.Columns.AddRange(new object[] {
																																	ultraDataColumn117,
																																	ultraDataColumn118,
																																	ultraDataColumn119,
																																	ultraDataColumn120});
			// 
			// chkOpcional
			// 
			appearance165.FontData.Name = "Tahoma";
			appearance165.ForeColorDisabled = System.Drawing.Color.Black;
			this.chkOpcional.Appearance = appearance165;
			this.chkOpcional.BackColor = System.Drawing.Color.Transparent;
			this.chkOpcional.CausesValidation = false;
			this.chkOpcional.Enabled = false;
			this.chkOpcional.Location = new System.Drawing.Point(968, 296);
			this.chkOpcional.Name = "chkOpcional";
			this.chkOpcional.Size = new System.Drawing.Size(136, 21);
			this.chkOpcional.TabIndex = 315;
			this.chkOpcional.Text = "Obsequios Opcionales";
			// 
			// uGridCuotas
			// 
			this.uGridCuotas.Cursor = System.Windows.Forms.Cursors.Default;
			this.uGridCuotas.DataSource = this.ultraDataSource14;
			appearance166.BackColor = System.Drawing.Color.White;
			this.uGridCuotas.DisplayLayout.Appearance = appearance166;
			this.uGridCuotas.DisplayLayout.AutoFitColumns = true;
			ultraGridColumn117.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn117.Header.VisiblePosition = 0;
			ultraGridColumn117.Hidden = true;
			ultraGridColumn117.Width = 69;
			ultraGridColumn118.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn118.Header.VisiblePosition = 1;
			ultraGridColumn118.Hidden = true;
			ultraGridColumn118.Width = 44;
			ultraGridColumn119.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn119.Header.VisiblePosition = 2;
			ultraGridColumn119.Width = 90;
			ultraGridColumn120.Header.Caption = "...";
			ultraGridColumn120.Header.VisiblePosition = 4;
			ultraGridColumn120.Width = 19;
			appearance167.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn121.CellAppearance = appearance167;
			ultraGridColumn121.Format = "#,##0.00";
			ultraGridColumn121.Header.Caption = "% Ent.";
			ultraGridColumn121.Header.VisiblePosition = 3;
			ultraGridColumn121.Width = 54;
			ultraGridBand13.Columns.AddRange(new object[] {
																											ultraGridColumn117,
																											ultraGridColumn118,
																											ultraGridColumn119,
																											ultraGridColumn120,
																											ultraGridColumn121});
			ultraGridBand13.Header.Caption = "CUOTAS CREDIPOINT";
			ultraGridBand13.HeaderVisible = true;
			ultraGridBand13.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			appearance168.FontData.SizeInPoints = 7.25F;
			ultraGridBand13.Override.HeaderAppearance = appearance168;
			this.uGridCuotas.DisplayLayout.BandsSerializer.Add(ultraGridBand13);
			appearance169.ForeColor = System.Drawing.Color.Black;
			appearance169.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridCuotas.DisplayLayout.Override.ActiveRowAppearance = appearance169;
			this.uGridCuotas.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.uGridCuotas.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance170.BackColor = System.Drawing.Color.Transparent;
			this.uGridCuotas.DisplayLayout.Override.CardAreaAppearance = appearance170;
			appearance171.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance171.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance171.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance171.FontData.BoldAsString = "True";
			appearance171.FontData.Name = "Arial";
			appearance171.FontData.SizeInPoints = 10F;
			appearance171.ForeColor = System.Drawing.Color.White;
			appearance171.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridCuotas.DisplayLayout.Override.HeaderAppearance = appearance171;
			appearance172.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance172.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance172.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridCuotas.DisplayLayout.Override.RowAlternateAppearance = appearance172;
			appearance173.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance173.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance173.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridCuotas.DisplayLayout.Override.RowSelectorAppearance = appearance173;
			appearance174.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance174.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance174.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridCuotas.DisplayLayout.Override.SelectedRowAppearance = appearance174;
			this.uGridCuotas.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridCuotas.Location = new System.Drawing.Point(776, 56);
			this.uGridCuotas.Name = "uGridCuotas";
			this.uGridCuotas.Size = new System.Drawing.Size(184, 150);
			this.uGridCuotas.TabIndex = 316;
			this.uGridCuotas.KeyDown += new System.Windows.Forms.KeyEventHandler(this.uGridCuotas_KeyDown);
			this.uGridCuotas.AfterCellUpdate += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.uGridCuotas_AfterCellUpdate);
			// 
			// ultraDataSource14
			// 
			ultraDataColumn121.DataType = typeof(int);
			ultraDataColumn122.DataType = typeof(int);
			ultraDataColumn123.DataType = typeof(int);
			ultraDataColumn124.DataType = typeof(bool);
			ultraDataColumn124.DefaultValue = false;
			ultraDataColumn125.DataType = typeof(System.Decimal);
			this.ultraDataSource14.Band.Columns.AddRange(new object[] {
																																	ultraDataColumn121,
																																	ultraDataColumn122,
																																	ultraDataColumn123,
																																	ultraDataColumn124,
																																	ultraDataColumn125});
			// 
			// chkCuotas
			// 
			appearance175.ForeColorDisabled = System.Drawing.Color.Black;
			this.chkCuotas.Appearance = appearance175;
			this.chkCuotas.BackColor = System.Drawing.Color.Transparent;
			this.chkCuotas.CausesValidation = false;
			this.chkCuotas.CheckAlign = System.Drawing.ContentAlignment.MiddleCenter;
			this.chkCuotas.Enabled = false;
			this.chkCuotas.Location = new System.Drawing.Point(944, 80);
			this.chkCuotas.Name = "chkCuotas";
			this.chkCuotas.Size = new System.Drawing.Size(13, 13);
			this.chkCuotas.TabIndex = 317;
			this.chkCuotas.Text = "Precio Manual";
			this.chkCuotas.CheckedChanged += new System.EventHandler(this.chkCuotas_CheckedChanged);
			// 
			// chkCupon
			// 
			appearance176.FontData.Name = "Tahoma";
			appearance176.ForeColorDisabled = System.Drawing.Color.Black;
			this.chkCupon.Appearance = appearance176;
			this.chkCupon.BackColor = System.Drawing.Color.Transparent;
			this.chkCupon.CausesValidation = false;
			this.chkCupon.Enabled = false;
			this.chkCupon.Location = new System.Drawing.Point(968, 176);
			this.chkCupon.Name = "chkCupon";
			this.chkCupon.Size = new System.Drawing.Size(136, 21);
			this.chkCupon.TabIndex = 318;
			this.chkCupon.Text = "Cupon";
			// 
			// chkValidaCargaProducto
			// 
			appearance177.FontData.Name = "Tahoma";
			appearance177.ForeColorDisabled = System.Drawing.Color.Black;
			this.chkValidaCargaProducto.Appearance = appearance177;
			this.chkValidaCargaProducto.BackColor = System.Drawing.Color.Transparent;
			this.chkValidaCargaProducto.CausesValidation = false;
			this.chkValidaCargaProducto.Enabled = false;
			this.chkValidaCargaProducto.Location = new System.Drawing.Point(968, 320);
			this.chkValidaCargaProducto.Name = "chkValidaCargaProducto";
			this.chkValidaCargaProducto.Size = new System.Drawing.Size(136, 21);
			this.chkValidaCargaProducto.TabIndex = 319;
			this.chkValidaCargaProducto.Text = "Producto Obligatorio";
			// 
			// label4
			// 
			this.label4.AutoSize = true;
			this.label4.BackColor = System.Drawing.Color.Transparent;
			this.label4.Enabled = false;
			this.label4.Location = new System.Drawing.Point(992, 136);
			this.label4.Name = "label4";
			this.label4.Size = new System.Drawing.Size(70, 16);
			this.label4.TabIndex = 321;
			this.label4.Text = "Margen Total";
			this.label4.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtMargenTotal
			// 
			appearance178.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtMargenTotal.Appearance = appearance178;
			this.txtMargenTotal.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.txtMargenTotal.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtMargenTotal.Enabled = false;
			this.txtMargenTotal.FormatString = "#,##0.00";
			this.txtMargenTotal.Location = new System.Drawing.Point(992, 152);
			this.txtMargenTotal.Name = "txtMargenTotal";
			this.txtMargenTotal.NullText = "0.00";
			this.txtMargenTotal.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtMargenTotal.PromptChar = ' ';
			this.txtMargenTotal.Size = new System.Drawing.Size(72, 20);
			this.txtMargenTotal.TabIndex = 320;
			// 
			// Promociones
			// 
			this.Appearance.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.Appearance.Options.UseFont = true;
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 14);
			this.ClientSize = new System.Drawing.Size(1098, 643);
			this.Controls.Add(this.label4);
			this.Controls.Add(this.label16);
			this.Controls.Add(this.label3);
			this.Controls.Add(this.label2);
			this.Controls.Add(this.label13);
			this.Controls.Add(this.label14);
			this.Controls.Add(this.label15);
			this.Controls.Add(this.lblCliente);
			this.Controls.Add(this.label1);
			this.Controls.Add(this.label6);
			this.Controls.Add(this.txtMargenTotal);
			this.Controls.Add(this.chkValidaCargaProducto);
			this.Controls.Add(this.chkCupon);
			this.Controls.Add(this.chkCuotas);
			this.Controls.Add(this.uGridCuotas);
			this.Controls.Add(this.chkOpcional);
			this.Controls.Add(this.uGridPlazos);
			this.Controls.Add(this.chkAdicional);
			this.Controls.Add(this.chkValor);
			this.Controls.Add(this.chkSumaDescuentos);
			this.Controls.Add(this.uGridModelos);
			this.Controls.Add(this.uGridPlanes);
			this.Controls.Add(this.chkVariable);
			this.Controls.Add(this.chkMarcas);
			this.Controls.Add(this.chkSubGrupos);
			this.Controls.Add(this.chkGrupos);
			this.Controls.Add(this.groupBox1);
			this.Controls.Add(this.chkLocales);
			this.Controls.Add(this.btnActualizaPrecios);
			this.Controls.Add(this.txtidArticulo);
			this.Controls.Add(this.lblFinalizado);
			this.Controls.Add(this.btFinalizar);
			this.Controls.Add(this.btCancelar);
			this.Controls.Add(this.btBuscar);
			this.Controls.Add(this.btGrabar);
			this.Controls.Add(this.btNuevo);
			this.Controls.Add(this.btSalir);
			this.Controls.Add(this.btEditar);
			this.Controls.Add(this.uGridBancos);
			this.Controls.Add(this.uGridMarcas);
			this.Controls.Add(this.uGridSubgrupos);
			this.Controls.Add(this.uGridGrupos);
			this.Controls.Add(this.uGridTiposDePrecio);
			this.Controls.Add(this.uGridBodegas);
			this.Controls.Add(this.uGridCiudades);
			this.Controls.Add(this.uGridTarjetas);
			this.Controls.Add(this.cmbTipo);
			this.Controls.Add(this.txtValorMaximo);
			this.Controls.Add(this.txtValorMinimo);
			this.Controls.Add(this.txtCuotaCalculada);
			this.Controls.Add(this.txtCuotaReal);
			this.Controls.Add(this.txtNroCuotas);
			this.Controls.Add(this.ultraGrid1);
			this.Controls.Add(this.txtIdPromocion);
			this.Controls.Add(this.txtDescripcion);
			this.Controls.Add(this.dtFin);
			this.Controls.Add(this.dtInicio);
			this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
			this.KeyPreview = true;
			this.MaximizeBox = false;
			this.Name = "Promociones";
			this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
			this.Text = "Promociones";
			this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.Promociones_KeyDown);
			this.Closing += new System.ComponentModel.CancelEventHandler(this.Promociones_Closing);
			this.Load += new System.EventHandler(this.Promociones_Load);
			((System.ComponentModel.ISupportInitialize)(this.dtFin)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtInicio)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDescripcion)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdPromocion)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCuotaCalculada)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCuotaReal)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNroCuotas)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtValorMaximo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtValorMinimo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbTipo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridTarjetas)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource6)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridCiudades)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridBodegas)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource3)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridTiposDePrecio)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource4)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridMarcas)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource9)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridSubgrupos)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource8)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridGrupos)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource7)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridBancos)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource5)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsPromocion)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtidArticulo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridPlanes)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource10)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridModelos)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource11)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridPlazos)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource12)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource13)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridCuotas)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource14)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtMargenTotal)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		CultureInfo us = new CultureInfo("en-US");
		bool bNuevo = false;
		bool bEditar = false;
		bool bModoEdicion = false;
		int iNumIngreso = 0;
		string sArticulo = "";		
		bool bReading = false;
		bool bSelLocales = false;		
		bool bSelSubGrupos = false;
		bool bSelMarcas = false;
		bool bSelCuotas = false;
		bool bComboArticulo = false;
		int idBloqueaTransacciones = 0;
		
		public void EstadoGrid(bool Estado)
		{		
			if (Estado)
			{
				this.uGridCiudades.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
				this.uGridBodegas.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
				this.uGridTiposDePrecio.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
				this.uGridCuotas.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;	
				this.uGridBancos.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
				this.uGridTarjetas.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
				this.uGridPlanes.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
				this.uGridPlazos.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
				this.uGridGrupos.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
				this.uGridSubgrupos.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
				this.uGridMarcas.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;	
				this.uGridModelos.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;					
				this.ultraGrid1.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			}
			else
			{
				this.uGridCiudades.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.False;
				this.uGridBodegas.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.False;
				this.uGridTiposDePrecio.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.False;
				this.uGridCuotas.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.False;	
				this.uGridBancos.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.False;
				this.uGridTarjetas.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.False;
				this.uGridPlanes.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.False;
				this.uGridPlazos.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.False;
				this.uGridGrupos.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.False;
				this.uGridSubgrupos.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.False;
				this.uGridMarcas.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.False;
				this.uGridModelos.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.False;
				this.ultraGrid1.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.False;
			}
		}
		
		private void Promociones_Load(object sender, System.EventArgs e)
		{				
			EstadoGrid(false);		
		}

		private void dtInicio_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (e.KeyChar == 13) this.dtFin.Focus();
		}

		private void dtFin_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (e.KeyChar == 13)
			{
				this.txtDescripcion.Focus();
			}
		}

		private void dtInicio_Validated(object sender, System.EventArgs e)
		{
			this.dtFin.CalendarInfo.MinDate = (DateTime) this.dtInicio.Value;
			this.dtFin.Value = (DateTime) this.dtInicio.Value;
		}

		private void ultraGrid1_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			FuncionesProcedimientos.DesplazamientoGrids(sender, e, this.ultraGrid1);
			
			if (ultraGrid1.ActiveCell == null) return;
			try
			{				
				switch(e.KeyValue)
				{	
						#region Proveedores
					case (int) Keys.F3:
						if (this.ultraGrid1.ActiveCell.Column.Key.ToString().ToUpper() == "PROVEEDOR")
						{
							using (BuscaClientes miBusqueda = new BuscaClientes(this.ultraGrid1.ActiveRow.Cells["Proveedor"].Value.ToString(), 4, 0, 1, 0, false))
							{
								if (DialogResult.OK == miBusqueda.ShowDialog())
								{	
									this.ultraGrid1.ActiveRow.Cells["idCliente"].Value = miBusqueda.idCliente;									
									this.ultraGrid1.ActiveRow.Cells["Proveedor"].Value = miBusqueda.Nombre;
								}
							}
						}						
						break;
						#endregion Proveedores						
				}
			}
			catch(System.Exception ex)
			{
				MessageBox.Show(ex.Message, "Point Technology");
			}
		}
		
		private void ultraGrid1_AfterCellUpdate(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			#region Articulo
			if (bReading) return;
			bReading = false;

//			if (e.Cell.Column.ToString() == "Cantidad")
//			{
//				bReading = true;
//			}

			if (e.Cell.Column.ToString() == "Codigo")
			{	
				bReading = true;
			
				int iCont = 0;
				iCont = Funcion.iEscalarSQL(cdsSeteoF, string.Format("Select Count(*) From Combo c Inner Join Articulo a On a.idArticulo = c.idArticuloCombo Where a.Codigo = '{0}'", e.Cell.Row.Cells["Codigo"].Value.ToString()));
					
				if (iCont > 0 && DialogResult.Yes == MessageBox.Show("Este Artículo esta Asociado a un Combo. ¿Cargar los Artículos del Combo?", "Point Technology", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2))
				{
					int idCombo = Funcion.iEscalarSQL(cdsSeteoF, string.Format("Select idCombo From Combo c Inner Join Articulo a On a.idArticulo = c.idArticuloCombo Where a.Codigo = '{0}'", e.Cell.Row.Cells["Codigo"].Value.ToString()));
						
					int iFila = e.Cell.Row.Index;

					string sSQLInf = string.Format("Select a.idArticulo, a.Codigo, a.Articulo, ROUND(a.CostoUltimo, 2), ROUND(a.Precio2, 2), ROUND(a.Precio3, 2), ROUND(a.Precio4, 2), ROUND(a.Precio6, 2), cd.Descuento, DescuentoTC, DescuentoCP, cd.DescuentoTCSI, cd.Principal, a.PrecioC3,  a.PrecioC6,  a.PrecioC9,  a.PrecioC12,  a.PrecioC15,  a.PrecioC18,  a.PrecioC21,  a.PrecioC24 From Articulo a Inner Join ComboDetalle cd On cd.idArticulo = a.idArticulo Where cd.idCombo = {0} Order By cd.Principal Desc, idDetCombo", idCombo);//a.idArticulo = {0} And iIdArticulo, 
					SqlDataReader drInf = Funcion.rEscalarSQL(cdsSeteoF, sSQLInf, true);
										
					while (drInf.Read())
					{	
						ultraGrid1.Rows[iFila].Cells["idArticulo"].Value = drInf.GetInt32(0);
						string sCodigoArt = drInf.GetString(1);
						string Articulo = drInf.GetString(2);
						decimal dCostoPromedio = drInf.GetDecimal(3);
						decimal dPrecio2 = drInf.GetDecimal(4);
						decimal dPrecio3 = drInf.GetDecimal(5);
						decimal dPrecio4 = drInf.GetDecimal(6);
						decimal dPrecio6 = drInf.GetDecimal(7);
						decimal dDesc = drInf.GetDecimal(8);
						decimal dDescTC = drInf.GetDecimal(9);
						decimal dDescCP = drInf.GetDecimal(10);
						decimal DescuentoTCSI = drInf.GetDecimal(11);
						bool bPrincipal = drInf.GetBoolean(12);
						//if (this.cmbTipo.Text == "COMBO PAYJOY") dDescTC= 100;	
			
						ultraGrid1.Rows[iFila].Cells["Articulo"].Value = Articulo;
						ultraGrid1.Rows[iFila].Cells["CostoPromedio"].Value = dCostoPromedio;
						ultraGrid1.Rows[iFila].Cells["Precio2"].Value = dPrecio2;
						ultraGrid1.Rows[iFila].Cells["Precio3"].Value = dPrecio3;
						ultraGrid1.Rows[iFila].Cells["Precio4"].Value = dPrecio4;		
						ultraGrid1.Rows[iFila].Cells["Precio6"].Value = dPrecio6;		
						ultraGrid1.Rows[iFila].Cells["Principal"].Value = bPrincipal;																									
						ultraGrid1.Rows[iFila].Cells["Codigo"].Value = sCodigoArt;
						ultraGrid1.Rows[iFila].Cells["Descuento"].Value = dDesc;
						ultraGrid1.Rows[iFila].Cells["DescuentoTC"].Value = dDescTC;
						ultraGrid1.Rows[iFila].Cells["DescuentoCP"].Value = dDescCP;
						ultraGrid1.Rows[iFila].Cells["DescuentoTCSI"].Value = DescuentoTCSI;

						e.Cell.Row.Cells["PrecioC3"].Value = drInf.GetDecimal(13);	
						e.Cell.Row.Cells["PrecioC6"].Value = drInf.GetDecimal(14);	
						e.Cell.Row.Cells["PrecioC9"].Value = drInf.GetDecimal(15);	
						e.Cell.Row.Cells["PrecioC12"].Value = drInf.GetDecimal(16);	
						e.Cell.Row.Cells["PrecioC15"].Value = drInf.GetDecimal(17);	
						e.Cell.Row.Cells["PrecioC18"].Value = drInf.GetDecimal(18);	
						e.Cell.Row.Cells["PrecioC21"].Value = drInf.GetDecimal(19);	
						e.Cell.Row.Cells["PrecioC24"].Value = drInf.GetDecimal(20);	

						ultraGrid1.Rows[iFila].Cells["PromoPrecio2"].Value = Math.Round((decimal)ultraGrid1.Rows[iFila].Cells["Precio2"].Value - Math.Round((((decimal)ultraGrid1.Rows[iFila].Cells["Precio2"].Value * (decimal)ultraGrid1.Rows[iFila].Cells["Descuento"].Value) / 100), 3), 2); 
						ultraGrid1.Rows[iFila].Cells["PromoPrecio3"].Value = Math.Round((decimal)ultraGrid1.Rows[iFila].Cells["Precio3"].Value - Math.Round((((decimal)ultraGrid1.Rows[iFila].Cells["Precio3"].Value * (decimal)ultraGrid1.Rows[iFila].Cells["DescuentoTC"].Value) / 100), 3), 2);
						ultraGrid1.Rows[iFila].Cells["PromoPrecio4"].Value = Math.Round((decimal)ultraGrid1.Rows[iFila].Cells["Precio4"].Value - Math.Round((((decimal)ultraGrid1.Rows[iFila].Cells["Precio4"].Value * (decimal)ultraGrid1.Rows[iFila].Cells["DescuentoCP"].Value) / 100), 3), 2);
//          ultraGrid1.Rows[iFila].Cells["Precio6"].Value = Math.Round((decimal)ultraGrid1.Rows[iFila].Cells["Precio6"].Value - Math.Round((((decimal)ultraGrid1.Rows[iFila].Cells["Precio6"].Value * (decimal)ultraGrid1.Rows[iFila].Cells["DescuentoTCSI"].Value) / 100), 2), 2);
						e.Cell.Row.Cells["PromoPrecioC3"].Value = Math.Round((decimal)e.Cell.Row.Cells["PrecioC3"].Value - Math.Round((((decimal)e.Cell.Row.Cells["PrecioC3"].Value * (decimal)e.Cell.Row.Cells["DescuentoCP"].Value) / 100), 3), 2);
						e.Cell.Row.Cells["PromoPrecioC6"].Value = Math.Round((decimal)e.Cell.Row.Cells["PrecioC6"].Value - Math.Round((((decimal)e.Cell.Row.Cells["PrecioC6"].Value * (decimal)e.Cell.Row.Cells["DescuentoCP"].Value) / 100), 3), 2);
						e.Cell.Row.Cells["PromoPrecioC9"].Value = Math.Round((decimal)e.Cell.Row.Cells["PrecioC9"].Value - Math.Round((((decimal)e.Cell.Row.Cells["PrecioC9"].Value * (decimal)e.Cell.Row.Cells["DescuentoCP"].Value) / 100), 3), 2);
						e.Cell.Row.Cells["PromoPrecioC12"].Value = Math.Round((decimal)e.Cell.Row.Cells["PrecioC12"].Value - Math.Round((((decimal)e.Cell.Row.Cells["PrecioC12"].Value * (decimal)e.Cell.Row.Cells["DescuentoCP"].Value) / 100), 3), 2);
						e.Cell.Row.Cells["PromoPrecioC15"].Value = Math.Round((decimal)e.Cell.Row.Cells["PrecioC15"].Value - Math.Round((((decimal)e.Cell.Row.Cells["PrecioC15"].Value * (decimal)e.Cell.Row.Cells["DescuentoCP"].Value) / 100), 3), 2);
						e.Cell.Row.Cells["PromoPrecioC18"].Value = Math.Round((decimal)e.Cell.Row.Cells["PrecioC18"].Value - Math.Round((((decimal)e.Cell.Row.Cells["PrecioC18"].Value * (decimal)e.Cell.Row.Cells["DescuentoCP"].Value) / 100), 3), 2);
						e.Cell.Row.Cells["PromoPrecioC21"].Value = Math.Round((decimal)e.Cell.Row.Cells["PrecioC21"].Value - Math.Round((((decimal)e.Cell.Row.Cells["PrecioC21"].Value * (decimal)e.Cell.Row.Cells["DescuentoCP"].Value) / 100), 3), 2);
						e.Cell.Row.Cells["PromoPrecioC24"].Value = Math.Round((decimal)e.Cell.Row.Cells["PrecioC24"].Value - Math.Round((((decimal)e.Cell.Row.Cells["PrecioC24"].Value * (decimal)e.Cell.Row.Cells["DescuentoCP"].Value) / 100), 3), 2);

						iFila = iFila + 1;					

						this.ultraGrid1.DisplayLayout.Bands[0].AddNew();						
					}
					drInf.Close();												
															
					bReading = false;					
				}
				else
				{										
					string sSQLInf1 = string.Format("Exec PromocionInfoArticulo '{0}'", e.Cell.Row.Cells["Codigo"].Value.ToString());
					SqlDataReader drInf1 = Funcion.rEscalarSQL(cdsSeteoF, sSQLInf1, true);
					drInf1.Read();
					if (drInf1.HasRows)
					{
						e.Cell.Row.Cells["idArticulo"].Value = drInf1.GetInt32(0);						
						e.Cell.Row.Cells["Articulo"].Value = drInf1.GetString(2);
						e.Cell.Row.Cells["CostoPromedio"].Value = drInf1.GetDecimal(3);
						e.Cell.Row.Cells["Precio2"].Value = drInf1.GetDecimal(4);
						e.Cell.Row.Cells["Precio3"].Value = drInf1.GetDecimal(5);
						e.Cell.Row.Cells["Precio4"].Value = drInf1.GetDecimal(6);
						e.Cell.Row.Cells["Precio6"].Value = drInf1.GetDecimal(7);
						decimal dDscto = drInf1.GetDecimal(8);
						e.Cell.Row.Cells["PrecioC3"].Value = drInf1.GetDecimal(9);	
						e.Cell.Row.Cells["PrecioC6"].Value = drInf1.GetDecimal(10);	
						e.Cell.Row.Cells["PrecioC9"].Value = drInf1.GetDecimal(11);	
						e.Cell.Row.Cells["PrecioC12"].Value = drInf1.GetDecimal(12);	
						e.Cell.Row.Cells["PrecioC15"].Value = drInf1.GetDecimal(13);	
						e.Cell.Row.Cells["PrecioC18"].Value = drInf1.GetDecimal(14);	
						e.Cell.Row.Cells["PrecioC21"].Value = drInf1.GetDecimal(15);	
						e.Cell.Row.Cells["PrecioC24"].Value = drInf1.GetDecimal(16);	

						drInf1.Close();
            
						e.Cell.Row.Cells["Descuento"].Value = dDscto;
						e.Cell.Row.Cells["DescuentoTC"].Value = (this.cmbTipo.Text == "COMBO PAYJOY") ? 100 : dDscto;
						e.Cell.Row.Cells["DescuentoCP"].Value = dDscto;
						e.Cell.Row.Cells["DescuentoTCSI"].Value = dDscto;

						e.Cell.Row.Cells["PromoPrecio2"].Value = Math.Round((decimal)e.Cell.Row.Cells["Precio2"].Value - Math.Round((((decimal)e.Cell.Row.Cells["Precio2"].Value * (decimal)e.Cell.Row.Cells["Descuento"].Value) / 100), 2), 2); 
						e.Cell.Row.Cells["PromoPrecio3"].Value = Math.Round((decimal)e.Cell.Row.Cells["Precio3"].Value - Math.Round((((decimal)e.Cell.Row.Cells["Precio3"].Value * (decimal)e.Cell.Row.Cells["DescuentoTC"].Value) / 100), 2), 2);
						e.Cell.Row.Cells["PromoPrecio4"].Value = Math.Round((decimal)e.Cell.Row.Cells["Precio4"].Value - Math.Round((((decimal)e.Cell.Row.Cells["Precio4"].Value * (decimal)e.Cell.Row.Cells["DescuentoCP"].Value) / 100), 2), 2);
						//e.Cell.Row.Cells["PromoPrecio6"].Value = Math.Round((decimal)e.Cell.Row.Cells["Precio6"].Value - Math.Round((((decimal)e.Cell.Row.Cells["Precio6"].Value * (decimal)e.Cell.Row.Cells["DescuentoTCSI"].Value) / 100), 2), 2);

						e.Cell.Row.Cells["PromoPrecioC3"].Value = Math.Round((decimal)e.Cell.Row.Cells["PrecioC3"].Value - Math.Round((((decimal)e.Cell.Row.Cells["PrecioC3"].Value * (decimal)e.Cell.Row.Cells["DescuentoCP"].Value) / 100), 2), 2);
						e.Cell.Row.Cells["PromoPrecioC6"].Value = Math.Round((decimal)e.Cell.Row.Cells["PrecioC6"].Value - Math.Round((((decimal)e.Cell.Row.Cells["PrecioC6"].Value * (decimal)e.Cell.Row.Cells["DescuentoCP"].Value) / 100), 2), 2);
						e.Cell.Row.Cells["PromoPrecioC9"].Value = Math.Round((decimal)e.Cell.Row.Cells["PrecioC9"].Value - Math.Round((((decimal)e.Cell.Row.Cells["PrecioC9"].Value * (decimal)e.Cell.Row.Cells["DescuentoCP"].Value) / 100), 2), 2);
						e.Cell.Row.Cells["PromoPrecioC12"].Value = Math.Round((decimal)e.Cell.Row.Cells["PrecioC12"].Value - Math.Round((((decimal)e.Cell.Row.Cells["PrecioC12"].Value * (decimal)e.Cell.Row.Cells["DescuentoCP"].Value) / 100), 2), 2);
						e.Cell.Row.Cells["PromoPrecioC15"].Value = Math.Round((decimal)e.Cell.Row.Cells["PrecioC15"].Value - Math.Round((((decimal)e.Cell.Row.Cells["PrecioC15"].Value * (decimal)e.Cell.Row.Cells["DescuentoCP"].Value) / 100), 2), 2);
						e.Cell.Row.Cells["PromoPrecioC18"].Value = Math.Round((decimal)e.Cell.Row.Cells["PrecioC18"].Value - Math.Round((((decimal)e.Cell.Row.Cells["PrecioC18"].Value * (decimal)e.Cell.Row.Cells["DescuentoCP"].Value) / 100), 2), 2);
						e.Cell.Row.Cells["PromoPrecioC21"].Value = Math.Round((decimal)e.Cell.Row.Cells["PrecioC21"].Value - Math.Round((((decimal)e.Cell.Row.Cells["PrecioC21"].Value * (decimal)e.Cell.Row.Cells["DescuentoCP"].Value) / 100), 2), 2);
						e.Cell.Row.Cells["PromoPrecioC24"].Value = Math.Round((decimal)e.Cell.Row.Cells["PrecioC24"].Value - Math.Round((((decimal)e.Cell.Row.Cells["PrecioC24"].Value * (decimal)e.Cell.Row.Cells["DescuentoCP"].Value) / 100), 2), 2);
					}
					drInf1.Close();
				}
				bReading = false;
				CalculaCuota();
			}
			#endregion Articulo

			#region Descuentos
			if (e.Cell.Column.ToString() == "Descuento")
			{
				e.Cell.Row.Cells["PromoPrecio2"].Value = Math.Round(Convert.ToDecimal(e.Cell.Row.Cells["Precio2"].Value) - Math.Round(((Convert.ToDecimal(e.Cell.Row.Cells["Precio2"].Value) * Convert.ToDecimal(e.Cell.Row.Cells["Descuento"].Value)) / 100), 2), 2); 
				
				CalculaCuota();
			}

			if (e.Cell.Column.ToString() == "DescuentoTC")
			{
				e.Cell.Row.Cells["PromoPrecio3"].Value = Math.Round(Convert.ToDecimal(e.Cell.Row.Cells["Precio3"].Value) - Math.Round(((Convert.ToDecimal(e.Cell.Row.Cells["Precio3"].Value) * Convert.ToDecimal(e.Cell.Row.Cells["DescuentoTC"].Value)) / 100), 2), 2);

				CalculaCuota();
			}

			if (e.Cell.Column.ToString() == "DescuentoTCSI")
			{
				//e.Cell.Row.Cells["PromoPrecio6"].Value = Math.Round(Convert.ToDecimal(e.Cell.Row.Cells["Precio6"].Value) - Math.Round(((Convert.ToDecimal(e.Cell.Row.Cells["Precio6"].Value) * Convert.ToDecimal(e.Cell.Row.Cells["DescuentoTCSI"].Value)) / 100), 2), 2);

				CalculaCuota();
			}

			if (e.Cell.Column.ToString() == "DescuentoCP")
			{
				e.Cell.Row.Cells["PromoPrecio4"].Value = Math.Round(Convert.ToDecimal(e.Cell.Row.Cells["Precio4"].Value) - Math.Round(((Convert.ToDecimal(e.Cell.Row.Cells["Precio4"].Value) * Convert.ToDecimal(e.Cell.Row.Cells["DescuentoCP"].Value)) / 100), 2), 2);

				e.Cell.Row.Cells["PromoPrecioC3"].Value = Math.Round(Convert.ToDecimal(e.Cell.Row.Cells["PrecioC3"].Value) - Math.Round(((Convert.ToDecimal(e.Cell.Row.Cells["PrecioC3"].Value) * Convert.ToDecimal(e.Cell.Row.Cells["DescuentoCP"].Value)) / 100), 2), 2);
				e.Cell.Row.Cells["PromoPrecioC6"].Value = Math.Round(Convert.ToDecimal(e.Cell.Row.Cells["PrecioC6"].Value) - Math.Round(((Convert.ToDecimal(e.Cell.Row.Cells["PrecioC6"].Value) * Convert.ToDecimal(e.Cell.Row.Cells["DescuentoCP"].Value)) / 100), 2), 2);
				e.Cell.Row.Cells["PromoPrecioC9"].Value = Math.Round(Convert.ToDecimal(e.Cell.Row.Cells["PrecioC9"].Value) - Math.Round(((Convert.ToDecimal(e.Cell.Row.Cells["PrecioC9"].Value) * Convert.ToDecimal(e.Cell.Row.Cells["DescuentoCP"].Value)) / 100), 2), 2);
				e.Cell.Row.Cells["PromoPrecioC12"].Value = Math.Round(Convert.ToDecimal(e.Cell.Row.Cells["PrecioC12"].Value) - Math.Round(((Convert.ToDecimal(e.Cell.Row.Cells["PrecioC12"].Value) * Convert.ToDecimal(e.Cell.Row.Cells["DescuentoCP"].Value)) / 100), 2), 2);
				e.Cell.Row.Cells["PromoPrecioC15"].Value = Math.Round(Convert.ToDecimal(e.Cell.Row.Cells["PrecioC15"].Value) - Math.Round(((Convert.ToDecimal(e.Cell.Row.Cells["PrecioC15"].Value) * Convert.ToDecimal(e.Cell.Row.Cells["DescuentoCP"].Value)) / 100), 2), 2);
				e.Cell.Row.Cells["PromoPrecioC18"].Value = Math.Round(Convert.ToDecimal(e.Cell.Row.Cells["PrecioC18"].Value) - Math.Round(((Convert.ToDecimal(e.Cell.Row.Cells["PrecioC18"].Value) * Convert.ToDecimal(e.Cell.Row.Cells["DescuentoCP"].Value)) / 100), 2), 2);
				e.Cell.Row.Cells["PromoPrecioC21"].Value = Math.Round(Convert.ToDecimal(e.Cell.Row.Cells["PrecioC21"].Value) - Math.Round(((Convert.ToDecimal(e.Cell.Row.Cells["PrecioC21"].Value) * Convert.ToDecimal(e.Cell.Row.Cells["DescuentoCP"].Value)) / 100), 2), 2);
				e.Cell.Row.Cells["PromoPrecioC24"].Value = Math.Round(Convert.ToDecimal(e.Cell.Row.Cells["PrecioC24"].Value) - Math.Round(((Convert.ToDecimal(e.Cell.Row.Cells["PrecioC24"].Value) * Convert.ToDecimal(e.Cell.Row.Cells["DescuentoCP"].Value)) / 100), 2), 2);

				CalculaCuota();
			}
			#endregion Descuentos

			#region CodAlterno1 Articulo
			if (e.Cell.Column.ToString() == "CodAlterno1")
			{
				string sSQL = string.Format("Exec BuscaArticulosDetalleCombos '{0}'", e.Cell.Row.Cells["CodAlterno1"].Value.ToString());
				int iCuenta = Funcion.iEscalarSQL(cdsSeteoF, sSQL);
				if (iCuenta == 1)
				{
					string sSQLArticulo = string.Format("Exec ConsultaInformacionArticulosCombos '{0}'", e.Cell.Row.Cells["CodAlterno1"].Value.ToString());
					SqlDataReader drArticulo = Funcion.rEscalarSQL(cdsSeteoF, sSQLArticulo, true);
					drArticulo.Read();
					if (drArticulo.HasRows)
					{
						e.Cell.Row.Cells["idArticuloAlterno1"].Value = drArticulo.GetInt32(0);
						//e.Cell.Row.Cells["CostoAlt1"].Value = drArticulo.GetDecimal(2);
					}
					drArticulo.Close();
				}
				else if (iCuenta > 1)
				{
					using (BuscaArticulos miBusqueda = new BuscaArticulos(e.Cell.Row.Cells["CodAlterno1"].Value.ToString(), 1, 0))
					{
						if (DialogResult.OK == miBusqueda.ShowDialog())
						{
							e.Cell.Row.Cells["CodAlterno1"].Value = miBusqueda.sCodigo;
						}
					}
				}

				if (e.Cell.Row.Cells["CodAlterno1"].Value.ToString().Length == 0)
				{
					e.Cell.Row.Cells["idArticuloAlterno1"].Value = System.DBNull.Value;
					//e.Cell.Row.Cells["CostoAlt1"].Value = 0.00m;
				}
			}
			#endregion CodAlterno1 Articulo

			#region CodAlterno2 Articulo
			if (e.Cell.Column.ToString() == "CodAlterno2")
			{
				string sSQL = string.Format("Exec BuscaArticulosDetalleCombos '{0}'", e.Cell.Row.Cells["CodAlterno2"].Value.ToString());
				int iCuenta = Funcion.iEscalarSQL(cdsSeteoF, sSQL);
				if (iCuenta == 1)
				{
					string sSQLArticulo = string.Format("Exec ConsultaInformacionArticulosCombos '{0}'", e.Cell.Row.Cells["CodAlterno2"].Value.ToString());
					SqlDataReader drArticulo = Funcion.rEscalarSQL(cdsSeteoF, sSQLArticulo, true);
					drArticulo.Read();
					if (drArticulo.HasRows)
					{
						e.Cell.Row.Cells["idArticuloAlterno2"].Value = drArticulo.GetInt32(0);
						//e.Cell.Row.Cells["CostoAlt2"].Value = drArticulo.GetDecimal(2);
					}
					drArticulo.Close();
				}
				else if (iCuenta > 1)
				{
					using (BuscaArticulos miBusqueda = new BuscaArticulos(e.Cell.Row.Cells["CodAlterno2"].Value.ToString(), 1, 0))
					{
						if (DialogResult.OK == miBusqueda.ShowDialog())
						{
							e.Cell.Row.Cells["CodAlterno2"].Value = miBusqueda.sCodigo;
						}
					}
				}

				if (e.Cell.Row.Cells["CodAlterno2"].Value.ToString().Length == 0) 
				{
					e.Cell.Row.Cells["idArticuloAlterno2"].Value = System.DBNull.Value;
					//e.Cell.Row.Cells["CostoAlt2"].Value = 0.00m;
				}
			}
			#endregion CodAlterno2 Articulo

			#region CodAlterno3 Articulo
			if (e.Cell.Column.ToString() == "CodAlterno3")
			{
				string sSQL = string.Format("Exec BuscaArticulosDetalleCombos '{0}'", e.Cell.Row.Cells["CodAlterno3"].Value.ToString());
				int iCuenta = Funcion.iEscalarSQL(cdsSeteoF, sSQL);
				if (iCuenta == 1)
				{
					string sSQLArticulo = string.Format("Exec ConsultaInformacionArticulosCombos '{0}'", e.Cell.Row.Cells["CodAlterno3"].Value.ToString());
					SqlDataReader drArticulo = Funcion.rEscalarSQL(cdsSeteoF, sSQLArticulo, true);
					drArticulo.Read();
					if (drArticulo.HasRows)
					{
						e.Cell.Row.Cells["idArticuloAlterno3"].Value = drArticulo.GetInt32(0);
						//e.Cell.Row.Cells["CostoAlt3"].Value = drArticulo.GetDecimal(2);
					}
					drArticulo.Close();
				}
				else if (iCuenta > 1)
				{
					using (BuscaArticulos miBusqueda = new BuscaArticulos(e.Cell.Row.Cells["CodAlterno3"].Value.ToString(), 1, 0))
					{
						if (DialogResult.OK == miBusqueda.ShowDialog())
						{
							e.Cell.Row.Cells["CodAlterno3"].Value = miBusqueda.sCodigo;
						}
					}
				}

				if (e.Cell.Row.Cells["CodAlterno3"].Value.ToString().Length == 0)
				{
					e.Cell.Row.Cells["idArticuloAlterno3"].Value = System.DBNull.Value;
					//e.Cell.Row.Cells["CostoAlt3"].Value = 0.00m;
				}
			}
			#endregion CodAlterno3 Articulo

			#region CodAlterno4 Articulo
			if (e.Cell.Column.ToString() == "CodAlterno4")
			{
				string sSQL = string.Format("Exec BuscaArticulosDetalleCombos '{0}'", e.Cell.Row.Cells["CodAlterno4"].Value.ToString());
				int iCuenta = Funcion.iEscalarSQL(cdsSeteoF, sSQL);
				if (iCuenta == 1)
				{
					string sSQLArticulo = string.Format("Exec ConsultaInformacionArticulosCombos '{0}'", e.Cell.Row.Cells["CodAlterno4"].Value.ToString());
					SqlDataReader drArticulo = Funcion.rEscalarSQL(cdsSeteoF, sSQLArticulo, true);
					drArticulo.Read();
					if (drArticulo.HasRows)
					{
						e.Cell.Row.Cells["idArticuloAlterno4"].Value = drArticulo.GetInt32(0);
						//e.Cell.Row.Cells["CostoAlt4"].Value = drArticulo.GetDecimal(2);
					}
					drArticulo.Close();
				}
				else if (iCuenta > 1)
				{
					using (BuscaArticulos miBusqueda = new BuscaArticulos(e.Cell.Row.Cells["CodAlterno4"].Value.ToString(), 1, 0))
					{
						if (DialogResult.OK == miBusqueda.ShowDialog())
						{
							e.Cell.Row.Cells["CodAlterno4"].Value = miBusqueda.sCodigo;
						}
					}
				}

				if (e.Cell.Row.Cells["CodAlterno4"].Value.ToString().Length == 0)
				{
					e.Cell.Row.Cells["idArticuloAlterno4"].Value = System.DBNull.Value;
					//e.Cell.Row.Cells["CostoAlt4"].Value = 0.00m;
				}
			}
			#endregion CodAlterno4 Articulo

			#region CodAlterno5 Articulo
			if (e.Cell.Column.ToString() == "CodAlterno5")
			{
				string sSQL = string.Format("Exec BuscaArticulosDetalleCombos '{0}'", e.Cell.Row.Cells["CodAlterno5"].Value.ToString());
				int iCuenta = Funcion.iEscalarSQL(cdsSeteoF, sSQL);
				if (iCuenta == 1)
				{
					string sSQLArticulo = string.Format("Exec ConsultaInformacionArticulosCombos '{0}'", e.Cell.Row.Cells["CodAlterno5"].Value.ToString());
					SqlDataReader drArticulo = Funcion.rEscalarSQL(cdsSeteoF, sSQLArticulo, true);
					drArticulo.Read();
					if (drArticulo.HasRows)
					{
						e.Cell.Row.Cells["idArticuloAlterno5"].Value = drArticulo.GetInt32(0);
						//e.Cell.Row.Cells["CostoAlt5"].Value = drArticulo.GetDecimal(2);
					}
					drArticulo.Close();
				}
				else if (iCuenta > 1)
				{
					using (BuscaArticulos miBusqueda = new BuscaArticulos(e.Cell.Row.Cells["CodAlterno5"].Value.ToString(), 1, 0))
					{
						if (DialogResult.OK == miBusqueda.ShowDialog())
						{
							e.Cell.Row.Cells["CodAlterno5"].Value = miBusqueda.sCodigo;
						}
					}
				}

				if (e.Cell.Row.Cells["CodAlterno5"].Value.ToString().Length == 0)
				{
					e.Cell.Row.Cells["idArticuloAlterno5"].Value = System.DBNull.Value;
					//e.Cell.Row.Cells["CostoAlt5"].Value = 0.00m;
				}
			}
			#endregion CodAlterno5 Articulo
		}

		private void CalculaCuota()
		{
			decimal dValor = 0;
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in ultraGrid1.Rows.All)
			{
				if (dr.Cells["PromoPrecio4"].Value != System.DBNull.Value) dValor = dValor + Convert.ToDecimal(dr.Cells["PromoPrecio4"].Value);
			}

			if ((int)this.txtNroCuotas.Value > 0) 
			{
				decimal dIVA = Funcion.decEscalarSQL(cdsSeteoF, "Select dbo.RetornaImpuestoIVAVigenteDecimal(GETDATE(), 1)");
				dValor = dValor * dIVA;

				string sSQL1 = string.Format("Select dbo.CalCuotaPromocion ({0}, {1})", dValor, (int)this.txtNroCuotas.Value);
				this.txtCuotaCalculada.Value  = Funcion.decEscalarSQL(cdsSeteoF, sSQL1);
			}
		}

		private void ultraGrid1_InitializeRow(object sender, Infragistics.Win.UltraWinGrid.InitializeRowEventArgs e)
		{
			e.Row.Cells["Cantidad"].Value = 1;
			e.Row.Cells["idCliente"].Value = 0;
		}

		private void ultraGrid1_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
			drSeteoF = cdsSeteoF.TableViews["SeteoF"].Rows[0];
			#region Formato numerico
			CultureInfo cultureENUS = CultureInfo.CreateSpecificCulture("ja-jp");
			string stMoneda = "#,##0.";
			string stMascara = "nnn,nnn,nnn.";
			string stCeros = "0000000000";
			string stNumero = "nnnnnnnnnn";
			string stFormato = "";
			string stInput = "";
			#endregion Formato numerico

			#region Decimales para porcentaje de descuento
			stFormato = stMoneda + stCeros.Substring(0,3);
			e.Layout.Bands[0].Columns["Descuento"].Format = stFormato;
			int iDecPorD = 0;
			if (drSeteoF["decPorcDesc"] != null) iDecPorD = (int) drSeteoF["decPorcDesc"];

			stInput = stMascara + stNumero.Substring(0, 3);
			e.Layout.Bands[0].Columns["Descuento"].MaskInput = stInput;
			stInput = stMascara + stNumero.Substring(0, 3);
			e.Layout.Bands[0].Columns["DescuentoTC"].MaskInput = stInput;
			e.Layout.Bands[0].Columns["DescuentoTCSI"].MaskInput = stInput;
			stInput = stMascara + stNumero.Substring(0, 3);
			e.Layout.Bands[0].Columns["DescuentoCP"].MaskInput = stInput;
			e.Layout.Bands[0].Columns["ValorMinimo"].MaskInput = stInput;
			e.Layout.Bands[0].Columns["ValorMaximo"].MaskInput = stInput;
			#endregion Decimales para porcentaje de descuento
		}

		private void txtNroCuotas_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (e.KeyChar == 13)
			{
				if ((int)this.txtNroCuotas.Value == 0) return;
				if (this.ultraGrid1.Rows.Count == 0) return;
				decimal dValor = 0;
				foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in ultraGrid1.Rows.All)
				{
					dValor = dValor + (decimal)dr.Cells["Precio4"].Value;
				}

				if (dValor == 0) return;
				
				DateTime dtFecha = DateTime.Today;
				string sSQL1 = string.Format("Exec CalculaCuotaInicial {0}, {1}, 0, '{2}'", dValor, (int)this.txtNroCuotas.Value, dtFecha.ToString("yyyyMMdd"));
				SqlDataReader dReader =	Funcion.rEscalarSQL(cdsSeteoF, sSQL1, true);
				dReader.Read();
				this.txtCuotaReal.Value = dReader.GetDecimal(2);
				dReader.Close();
			}
		}

		private void txtValorMinimo_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (e.KeyChar == 13) this.txtValorMaximo.Focus();
		}

		private void cdsPromocion_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsPromocion.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);		
		}

		private void btSalir_Click(object sender, System.EventArgs e)
		{
			Close();
		}

		private void btNuevo_Click(object sender, System.EventArgs e)
		{
			this.btCancelar_Click(sender, e);
			
			this.txtIdPromocion.Value = Funcion.iEscalarSQL(cdsSeteoF, string.Format("Exec GrabaPromociones {0}, 'ACTIVO'", (int)this.txtIdPromocion.Value));

			if (DialogResult.Yes == MessageBox.Show("¿Desea aplicar a todos los Almacenes.?", "Point Technology", MessageBoxButtons.YesNo, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2))
			{
				string sSQL = string.Format("Exec [GrabaPromocionesFull] {0}", (int)this.txtIdPromocion.Value);
				Funcion.EjecutaSQL(cdsSeteoF, sSQL, true);	

				this.uGridCiudades.DataSource = Funcion.dvProcedimiento(cdsSeteoF, string.Format("Exec ListaCiudadesPromocion {0}", (int)this.txtIdPromocion.Value));				
			}
			else
			{
				this.uGridCiudades.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Exec ListaCiudadesPromocion 0");
			
			}
			this.uGridTiposDePrecio.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Exec ListaPreciosPromocion 0");
			this.uGridGrupos.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Exec ListaGruposPromocion 0");

			EstadoGrid(true);
   			
			iNumIngreso = 1;

			this.lblFinalizado.Text = "ACTIVO";

			this.dtInicio.Value = DateTime.Today;
			this.dtFin.Value = DateTime.Today; 

			this.txtNroCuotas.Enabled = true;
			this.cmbTipo.Enabled = true;
			this.dtInicio.Enabled = true;
			this.dtFin.Enabled = true;
			this.txtDescripcion.Enabled = true;
			this.txtMargenTotal.Enabled = true;
			
			this.btNuevo.Enabled = false;
			this.btBuscar.Enabled = false;
			this.btEditar.Enabled = false;
			this.btGrabar.Enabled = true;
			this.btCancelar.Enabled = true;
						
			this.cmbTipo.Focus();

			bModoEdicion = true;
			bNuevo = true;
		}

		private void LimpiaGrids()
		{
			this.uGridCiudades.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Declare @False As Bit = 0 Select null As idPromocionCiudad, null As idPromocion, idCiudad, Ciudad, @False As Sel From ClienteCiudad Where idCiudad = 0");
			this.uGridBodegas.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Declare @False As Bit = 0 Select null As idPromCiudadBodega, null As idPromocionCiudad, Bodega, Nombre, @False As Sel From Bodega Where Bodega = 0");
			
			this.uGridTiposDePrecio.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Declare @False As Bit = 0 Select Null As idPromoPrecio, idTipoPrecio, Null As idPromocion, Precio, Null As Operador, 0.0 As Porcentaje, 0.0 As PorcentajeInicial, @False As Sel From ClientePrecio Where idTipoPrecio = 0");
			this.uGridCuotas.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Select 0 As idPromocionesCuotas, 0 As idPromocion, Cuota, ISNULL (Porcentaje_Entrada, 0) AS Porcentaje_Entrada ,CONVERT(Bit, 0) As Sel From CuotasCrediPOINT Where Cuota = 0");

			this.uGridBancos.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Declare @False As Bit = 0 Select Null As idPromocionBanco, Null As idPromocion, idBanco, Nombre, @False As Sel From Banco Where idBanco = 0");
			this.uGridTarjetas.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Declare @False As Bit = 0 Select Null As idPromocionBancoTarjeta, Null As idPromocionBanco, idTarjeta, Nombre, @False As Sel From TarjetasDeCredito Where idTarjeta = 0");
			this.uGridPlanes.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Select Null As idPromocionesPlanes, NULL As idPromocionBancoTarjeta, Null As idPromocionBanco, idPlan, Null As idPromocion, Descripcion, CONVERT(Bit, 0) As Sel From Planes Where idPlan = 0");
			this.uGridPlazos.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Select Null As idPromocionesPlazos, idPlazo, Null As idPromocion, Descripcion, CONVERT(Bit, 0) As Sel  From Plazos Where idPlazo = 0");

			this.uGridGrupos.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Declare @False As Bit = 0 Select Null As idPromocionGrupo, Null As idPromocion, idGrupoArticulo, Grupo, @False As Sel From ArticuloGrupo Where idGrupoArticulo = 0");
			this.uGridSubgrupos.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Declare @False As Bit = 0 Select Null As idPromocionSubGrupo, idSubGrupo, SubGrupo, @False As Sel From ArticuloSubGrupo Where idSubGrupo = 0");
			this.uGridMarcas.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Declare @False As Bit = 0 Select Null As idPromocionMarca, idMarca, Marca, @False As Sel From ArticuloMarca Where idMarca = 0");
			this.uGridModelos.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Select 0 As idPromocionModelo, idArticulo, Modelo, CONVERT(Bit, 0) As Sel From Articulo Where idArticulo = 0");
			
			this.ultraGrid1.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Exec BuscaDetallePromocion 0");
		}

		private void btCancelar_Click(object sender, System.EventArgs e)
		{
			string sSQLBloqueo = string.Format("Exec RegistraBloqueoTransacciones {0}, 0, {1}, 7", idBloqueaTransacciones, (int)this.txtIdPromocion.Value);
			Funcion.EjecutaSQL(cdsSeteoF, sSQLBloqueo, true);

			if (bNuevo) Funcion.EjecutaSQL(cdsSeteoF, string.Format("Exec CancelaPromocion {0}", (int)this.txtIdPromocion.Value));	

			bModoEdicion = false;	
			bNuevo = false;

			#region Limpiar - Habilitar
			this.txtIdPromocion.Value = 0;
			this.cmbTipo.Text = "";
			this.dtInicio.Value = DateTime.Today;
			this.dtFin.Value = DateTime.Today;
			this.txtDescripcion.Text = "";
			this.txtValorMinimo.Value = 0.00m;
			this.txtValorMaximo.Value = 0.00m;
			this.txtNroCuotas.Value = 0;
			this.txtCuotaReal.Value = 0.00m;
			this.txtCuotaCalculada.Value = 0.00m;
			this.txtidArticulo.Value = 0;
			this.lblFinalizado.Text = "";
			this.txtMargenTotal.Value = 0.00m;

			this.txtIdPromocion.Enabled = false;
			this.cmbTipo.Enabled = false;
			this.dtInicio.Enabled = false;
			this.dtFin.Enabled = false;
			this.txtDescripcion.Enabled = false;
			this.txtValorMinimo.Enabled = false;
			this.txtValorMaximo.Enabled = false;
			this.txtNroCuotas.Enabled = false;
			this.txtCuotaReal.Enabled = false;
			this.txtMargenTotal.Enabled = false;
			this.txtCuotaCalculada.Enabled = false;		
			#endregion Limpiar - Habilitar

			EstadoGrid(false);
			LimpiaGrids();

			this.btCancelar.Enabled = false;
			this.btGrabar.Enabled = false;
			this.btEditar.Enabled = false;
			this.btFinalizar.Enabled = false;
			this.btNuevo.Enabled = true;
			this.btBuscar.Enabled = true;
			this.btnActualizaPrecios.Enabled = false;					

			this.chkCuotas.Enabled = false;
			bSelCuotas = false;
			this.chkCuotas.Checked = false;

			this.chkLocales.Enabled = false;
			bSelLocales = false;
			this.chkLocales.Checked = false;

			this.chkGrupos.Enabled = false;			
			this.chkGrupos.Checked = false;

			this.chkSubGrupos.Enabled = false;
			bSelSubGrupos = false;
			this.chkSubGrupos.Checked = false;

			this.chkMarcas.Enabled = false;
			bSelMarcas = false;
			this.chkMarcas.Checked = false;	
		
			this.chkVariable.Checked = false;
			this.chkVariable.Enabled = false;
			this.chkSumaDescuentos.Checked = false;
			this.chkSumaDescuentos.Enabled = false;
			this.chkValor.Checked = false;
			this.chkValor.Enabled = false;
			this.chkAdicional.Checked = false;
			this.chkAdicional.Enabled = false;
			this.chkOpcional.Checked = false;
			this.chkOpcional.Enabled = false;
			this.chkCupon.Checked = false;
			this.chkCupon.Enabled = false;
			this.chkValidaCargaProducto.Checked = false;
			this.chkValidaCargaProducto.Enabled = false;

			idBloqueaTransacciones = 0;
		}
		
		private void uGridCiudades_Click(object sender, System.EventArgs e)
		{
			if (this.uGridCiudades.Rows.Count > 0) 
			{
				int idPromCiudad = 0;

				if (this.uGridCiudades.ActiveRow.Cells["idPromocionCiudad"].Value != System.DBNull.Value) idPromCiudad = (int)this.uGridCiudades.ActiveRow.Cells["idPromocionCiudad"].Value;

				if ((bool)this.uGridCiudades.ActiveRow.Cells["Sel"].Value)
				{
					string sSQL = string.Format("Exec ListaCiudadesBodegasPromocion {0}, {1}", idPromCiudad, (int)this.uGridCiudades.ActiveRow.Cells["idCiudad"].Value);
					this.uGridBodegas.DataSource = Funcion.dvProcedimiento(cdsSeteoF, sSQL);
					if (bEditar) this.chkLocales.Enabled = true;
					bSelLocales = true;
				}
				else if (!(bool)this.uGridCiudades.ActiveRow.Cells["Sel"].Value)
				{				
					this.uGridBodegas.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Exec ListaCiudadesBodegasPromocion 0, 0");
					this.chkLocales.Enabled = false;
					bSelLocales = false;
					this.chkLocales.Checked = false;
				}
			}
		}
		
		private void uGridTiposDePrecio_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
			#region Formato numerico
			CultureInfo cultureENUS = CultureInfo.CreateSpecificCulture("ja-jp");
			string stMoneda = "#,##0.";
			string stMascara = "nnn,nnn,nnn.";
			string stCeros = "0000000000";
			string stNumero = "nnnnnnnnnn";
			string stFormato = "";
			string stInput = "";
			#endregion Formato numerico

			#region Decimales para porcentaje de descuento
			stFormato = stMoneda + stCeros.Substring(0, 2);
			stInput = stMascara + stNumero.Substring(0, 2);
			e.Layout.Bands[0].Columns["Porcentaje"].Format = stFormato;			
			e.Layout.Bands[0].Columns["Porcentaje"].MaskInput = stInput;
			e.Layout.Bands[0].Columns["PorcentajeInicial"].Format = stFormato;			
			e.Layout.Bands[0].Columns["PorcentajeInicial"].MaskInput = stInput;			
			#endregion Decimales para porcentaje de descuento
		}

		private void uGridTiposDePrecio_Click(object sender, System.EventArgs e)
		{
			if (this.uGridTiposDePrecio.Rows.Count == 0) return;

			if (this.uGridTiposDePrecio.ActiveRow.Cells["idPromoPrecio"].Value == System.DBNull.Value) return;

			this.uGridBancos.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Declare @False As Bit = 0 Select Null As idPromocionBanco, Null As idPromocion, idBanco, Nombre, @False As Sel From Banco Where idBanco = 0");
			this.uGridTarjetas.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Declare @False As Bit = 0 Select Null As idPromocionBancoTarjeta, Null As idPromocionBanco, idTarjeta, Nombre, @False As Sel From TarjetasDeCredito Where idTarjeta = 0");
			this.uGridPlanes.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Select Null As idPromocionesPlanes, NULL As idPromocionBancoTarjeta, Null As idPromocionBanco, idPlan, Null As idPromocion, Descripcion, CONVERT(Bit, 0) As Sel From Planes Where idPlan = 0");
			this.uGridPlazos.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Select Null As idPromocionesPlazos, idPlazo, Null As idPromocion, Descripcion, CONVERT(Bit, 0) As Sel  From Plazos Where idPlazo = 0");
			this.uGridCuotas.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Select 0 As idPromocionesCuotas, 0 As idPromocion, ISNULL (Porcentaje_Entrada, 0 ) AS Porcentaje_Entrada, Cuota, CONVERT(Bit, 0) As Sel From CuotasCrediPOINT Where Cuota = 0");

			this.chkCuotas.Enabled = false;
			bSelCuotas = false;
			this.chkCuotas.Checked = false;

			if ((bool)this.uGridTiposDePrecio.ActiveRow.Cells["Sel"].Value)
			{
				if ((int)this.uGridTiposDePrecio.ActiveRow.Cells["idTipoPrecio"].Value == 3 || (int)this.uGridTiposDePrecio.ActiveRow.Cells["idTipoPrecio"].Value == 92)
				{				
					string sSQL = string.Format("Exec ListaBancosPromocion {0}", (int)this.txtIdPromocion.Value);
					this.uGridBancos.DataSource = Funcion.dvProcedimiento(cdsSeteoF, sSQL);		
			
					this.uGridTarjetas.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Declare @False As Bit = 0 Select Null As idPromocionBancoTarjeta, Null As idPromocionBanco, idTarjeta, Nombre, @False As Sel From TarjetasDeCredito Where idTarjeta = 0");
					this.uGridPlanes.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Select Null As idPromocionesPlanes, NULL As idPromocionBancoTarjeta, Null As idPromocionBanco, idPlan, Null As idPromocion, Descripcion, CONVERT(Bit, 0) As Sel From Planes Where idPlan = 0");
					this.uGridPlazos.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Select Null As idPromocionesPlazos, idPlazo, Null As idPromocion, Descripcion, CONVERT(Bit, 0) As Sel  From Plazos Where idPlazo = 0");
				}							
				else if ((int)this.uGridTiposDePrecio.ActiveRow.Cells["idTipoPrecio"].Value == 4)
				{
					string sSQL = string.Format("Exec ListaCuotasPromocion {0}", (int)this.txtIdPromocion.Value);
					this.uGridCuotas.DataSource = Funcion.dvProcedimiento(cdsSeteoF, sSQL);

					if (bEditar) this.chkCuotas.Enabled = true;
					bSelCuotas = true;
				}
			}
		}

		private void uGridBancos_Click(object sender, System.EventArgs e)
		{
			if (this.uGridBancos.Rows.Count == 0) return;

//			if (this.uGridBancos.ActiveRow.Cells["idPromocionBanco"].Value == System.DBNull.Value) return;

			if ((bool)this.uGridBancos.ActiveRow.Cells["Sel"].Value)
			{
				string sSQL = string.Format("Exec ListaTarjetasBancosPromocion {0}", (int)this.uGridBancos.ActiveRow.Cells["idPromocionBanco"].Value);//, (int)this.uGridBancos.ActiveRow.Cells["idBanco"].Value);
				this.uGridTarjetas.DataSource = Funcion.dvProcedimiento(cdsSeteoF, sSQL);				
			}
			else if (!(bool)this.uGridBancos.ActiveRow.Cells["Sel"].Value)
			{				
				this.uGridTarjetas.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Declare @False As Bit = 0 Select Null As idPromocionBancoTarjeta, Null As idPromocionBanco, idTarjeta, Nombre, @False As Sel From TarjetasDeCredito Where idTarjeta = 0");
				this.uGridPlanes.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Select Null As idPromocionesPlanes, NULL As idPromocionBancoTarjeta, Null As idPromocionBanco, idPlan, Null As idPromocion, Descripcion, CONVERT(Bit, 0) As Sel From Planes Where idPlan = 0");
				this.uGridPlazos.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Select Null As idPromocionesPlazos, idPlazo, Null As idPromocion, Descripcion, CONVERT(Bit, 0) As Sel  From Plazos Where idPlazo = 0");
			}
		}

		private void uGridCiudades_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			FuncionesProcedimientos.DesplazamientoGrids(sender, e, this.uGridCiudades);
		}

		private void uGridBodegas_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			FuncionesProcedimientos.DesplazamientoGrids(sender, e, this.uGridBodegas);
		}

		private void uGridTiposDePrecio_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			FuncionesProcedimientos.DesplazamientoGrids(sender, e, this.uGridTiposDePrecio);
		}

		private void uGridTiposDePrecio_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (uGridTiposDePrecio.ActiveCell.Column.Key.ToString().ToUpper() == "OPERADOR")
			{
				if (Char.IsControl(e.KeyChar)) e.Handled = false; 
				else if (e.KeyChar == 43)	e.Handled = false;
				else if (e.KeyChar == 45) e.Handled = false;
				else e.Handled = true;
			}
		}

		private void uGridBancos_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			FuncionesProcedimientos.DesplazamientoGrids(sender, e, this.uGridBancos);
		}

		private void uGridTarjetas_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			FuncionesProcedimientos.DesplazamientoGrids(sender, e, this.uGridTarjetas);
		}

		private void uGridGrupos_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			FuncionesProcedimientos.DesplazamientoGrids(sender, e, this.uGridGrupos);
		}

		private void uGridSubgrupos_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			FuncionesProcedimientos.DesplazamientoGrids(sender, e, this.uGridSubgrupos);
		}

		private void uGridMarcas_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			FuncionesProcedimientos.DesplazamientoGrids(sender, e, this.uGridMarcas);
		}

		private void btGrabar_Click(object sender, System.EventArgs e)
		{
			

			#region Validación
			if (this.cmbTipo.Text.ToString().Length == 0)
			{
				MessageBox.Show("Seleccione COMBO o PROMOCIÓN.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
				this.cmbTipo.Focus();
				return;
			}
			if (this.txtDescripcion.Text.ToString().Length == 0)
			{
				MessageBox.Show("Escriba un nombre para la promoción.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
				this.txtDescripcion.Focus();
				return;
			}
			if (this.ultraGrid1.Rows.Count == 0 && this.cmbTipo.Text.ToString() == "COMBO")
			{
				MessageBox.Show("Ingrese los productos para el combo.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
				this.cmbTipo.Focus();
				return;
			}

			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridTiposDePrecio.Rows.All)
			{
				if ((bool)dr.Cells["Sel"].Value)
				{
					if (dr.Cells["Operador"].Value.ToString().Length == 0)
					{
						MessageBox.Show("Escriba si es negativo o positivo", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
						this.uGridTiposDePrecio.ActiveRow = dr;
						this.uGridTiposDePrecio.ActiveRow.Selected = true;
						return;
					}
				}
			}				
			if (this.ultraGrid1.Rows.Count > 0)
			{
				bool hayWebPrincipal = false;
				if (this.cmbTipo.Text != "COMBO PAYJOY") 
				{
					foreach (Infragistics.Win.UltraWinGrid.UltraGridRow r in this.ultraGrid1.Rows.All)
					{
						if (r.Cells.Exists("WEBPRINCIPAL"))
							if (r.Cells["WEBPRINCIPAL"].Value != System.DBNull.Value && (bool)r.Cells["WEBPRINCIPAL"].Value) hayWebPrincipal = true;
					}

					if (!hayWebPrincipal)
					{
						System.Windows.Forms.MessageBox.Show("Debe marcar al menos una línea como WEB PRINCIPAL.","Point Technology", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Stop);
						return;
					}
				}
			}
			#endregion Validación

			#region Valida PayJoy
			if (this.cmbTipo.Text == "COMBO PAYJOY")
			{
				bool haySeleccion = false;

				foreach (Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridTiposDePrecio.Rows.All)
				{
					if (dr.Cells["Sel"].Value != null && (bool)dr.Cells["Sel"].Value)
					{
						haySeleccion = true;

						if ((int)dr.Cells["idTipoPrecio"].Value != 93)
						{
							MessageBox.Show("Para Combo PAYJOY debe seleccionar únicamente PRECIO PAYJOY.", 
								"Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
							this.uGridTiposDePrecio.ActiveRow = dr;
							this.uGridTiposDePrecio.ActiveRow.Selected = true;
							return;
						}
					}
				}

				// Si no se seleccionó ninguna fila, mostrar mensaje
				if (!haySeleccion)
				{
					MessageBox.Show("Debe seleccionar al menos un tipo de precio.", 
						"Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
					return;
				}

				#region Validar cien
				if (this.ultraGrid1.Rows.Count > 0)
				{
					foreach (Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.ultraGrid1.Rows.All)
					{
						if ((string)dr.Cells["Codigo"].Value != null )
						{
							if (dr.Cells["DescuentoTC"].Value != null && (decimal)dr.Cells["DescuentoTC"].Value != 100)
							{
								MessageBox.Show("Para Combo PAYJOY el % debe ser 100.", 
									"Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
								this.ultraGrid1.ActiveRow = dr;
								this.ultraGrid1.ActiveRow.Selected = true;
								return;
							}
						}
					}
				}
				#endregion Validar cien

			}
			#endregion Valida PayJoy

			#region Encabezado
			decimal dContado = 0.00m;
			decimal dTarjeta = 0.00m;
			decimal dCredito = 0.00m;

			DateTime dtInicio = (DateTime)this.dtInicio.Value;
			DateTime dtFin = (DateTime)this.dtFin.Value;

			if (this.cmbTipo.Text == "COMBO" || this.cmbTipo.Text == "BARATAZO")
			{
				string sCodigo = "";
				int iCont = 0;
				string sMarca = "PROMO -" + this.cmbTipo.Text.ToString().Substring(0, 5);

				if (bNuevo)
				{
					iCont = Funcion.iEscalarSQL(cdsSeteoF, string.Format("Select COUNT(*) From CompraPromocion Where Tipo = '{0}'", this.cmbTipo.Text));
					iCont++;
					sMarca = sMarca + iCont.ToString() + "-" + dtInicio.ToString("ddMMyy");
					sCodigo = Funcion.sEscalarSQL(cdsSeteoF, string.Format("Select dbo.CodigoArticulo(151, 112, '{0}', 0)", sMarca));
				}
				else
				{
					sCodigo = Funcion.sEscalarSQL(cdsSeteoF, string.Format("Declare @Codigo Varchar(100) = '' If Exists(Select ISNULL(Codigo, '') From Articulo Where idArticulo = {0}) Select @Codigo = ISNULL(Codigo, '') From Articulo Where idArticulo = {0} Else Set @Codigo = '' Select @Codigo", (int)this.txtidArticulo.Value));

					if (sCodigo.ToString().Length == 0) 
					{
//						iCont = Funcion.iEscalarSQL(cdsSeteoF, string.Format("Select COUNT(*) From CompraPromocion Where Tipo = '{0}'", this.cmbTipo.Text));
						iCont = Funcion.iEscalarSQL(cdsSeteoF, "Select COUNT(*) From CompraPromocion");
						iCont++;
						sMarca = sMarca + "-" + dtInicio.ToString("ddMMyy") + "-" + iCont.ToString() + this.txtIdPromocion.Value.ToString();
						sCodigo = Funcion.sEscalarSQL(cdsSeteoF, string.Format("Select dbo.CodigoArticulo(151, 112, '{0}', 0)", sMarca));
					}
				}
       
				if(MenuLatinium.stDirFacturacion == "POINT")
				{
					string sSQLArticulo = string.Format("Exec GuardaArticulos {0}, 1, 26, 151, 0, 112, '', '', 0, '{1}', '', '{2}', 53, 12, 0, 0, 0, 0, 0, 0, 1, 0, 0, 0, 0, 0, 1, 0, 0, {3}, 0, {4}, {5}, 0, 0, 0, 0, 0, 0, 0, NULL, 1, 0, 0, 1, 0, 0, 0, 0, 0, 0, 0, '', 0, 0, 0, 0, 0, 0, '', 0, 0, 0, 0, 0, 0, '', 0, 0",
						(int)this.txtidArticulo.Value, sCodigo, this.txtDescripcion.Text.ToString(), dContado, dTarjeta, dCredito);
					this.txtidArticulo.Value = Funcion.iEscalarSQL(cdsSeteoF, sSQLArticulo);
				}
				else if(MenuLatinium.stDirFacturacion == "CUEROFINO")
				{
					string sSQLArticulo = string.Format("Exec GuardaArticulos {0}, 1, 16, 24, 0, 4, '', '', 0, '{1}', '', '{2}', 53, 12, 0, 0, 0, 0, 0, 0, 1, 0, 0, 0, 0, 0, 1, 0, 0, {3}, 0, {4}, {5}, 0, 0, 0, 0, 0, 0, 0, NULL, 1, 0, 0, 1, 0, 0, 0, 0, 0, 0, 0, '', 0, 0, 0, 0, 0, 0, '', 0, 0, 0, 0, 0, 0, '', 0, 0",
						(int)this.txtidArticulo.Value, sCodigo, this.txtDescripcion.Text.ToString(), dContado, dTarjeta, dCredito);
					this.txtidArticulo.Value = Funcion.iEscalarSQL(cdsSeteoF, sSQLArticulo);
				}
			}

			string sSQL = string.Format("Exec GrabaPromociones {0}, '{1}', '{2}', '{3}', '{4}', {5}, {6}, {7}, {8}, {9}, {10}, '{11}', {12}, {13}, {14}, {15}, {16}, {17}, {18}, 0, {19}, '20000101', '20000101', 0,{20}", 
				(int)this.txtIdPromocion.Value, this.cmbTipo.Text.ToString(), dtInicio.ToString("yyyyMMdd HH:mm"), dtFin.ToString("yyyyMMdd HH:mm"), this.txtDescripcion.Text.ToString().Trim(), //4
				Convert.ToDecimal(this.txtValorMinimo.Value), Convert.ToDecimal(this.txtValorMaximo.Value), //6
				(int)this.txtNroCuotas.Value, Convert.ToDecimal(this.txtCuotaReal.Value), Convert.ToDecimal(this.txtCuotaCalculada.Value), //9
				iNumIngreso, this.lblFinalizado.Text.ToString(), (int)this.txtidArticulo.Value, //12
				(bool)this.chkVariable.Checked, (bool)this.chkSumaDescuentos.Checked, (bool)this.chkValor.Checked, //15
				(bool)this.chkAdicional.Checked, (bool)this.chkOpcional.Checked, (bool)this.chkCupon.Checked, (bool)this.chkValidaCargaProducto.Checked,Convert.ToDecimal(this.txtMargenTotal.Value));
			Funcion.EjecutaSQL(cdsSeteoF, sSQL, true);			
			#endregion Encabezado

			

			#region Detalle
			int idDetalle = 0;
			
			int idArtAlterno1 = 0;
			int idArtAlterno2 = 0;
			int idArtAlterno3 = 0;
			int idArtAlterno4 = 0;
			int idArtAlterno5 = 0;

			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.ultraGrid1.Rows.All)
			{
				if (dr.Cells["idDetallePromocion"].Value == System.DBNull.Value) idDetalle = 0; else idDetalle = (int)dr.Cells["idDetallePromocion"].Value;
				
				idArtAlterno1 = 0;
				idArtAlterno2 = 0;
				idArtAlterno3 = 0;
				idArtAlterno4 = 0;
				idArtAlterno5 = 0;

				if (dr.Cells["idArticuloAlterno1"].Value != System.DBNull.Value) idArtAlterno1 = (int) dr.Cells["idArticuloAlterno1"].Value;
				if (dr.Cells["idArticuloAlterno2"].Value != System.DBNull.Value) idArtAlterno2 = (int) dr.Cells["idArticuloAlterno2"].Value;
				if (dr.Cells["idArticuloAlterno3"].Value != System.DBNull.Value) idArtAlterno3 = (int) dr.Cells["idArticuloAlterno3"].Value;
				if (dr.Cells["idArticuloAlterno4"].Value != System.DBNull.Value) idArtAlterno4 = (int) dr.Cells["idArticuloAlterno4"].Value;
				if (dr.Cells["idArticuloAlterno5"].Value != System.DBNull.Value) idArtAlterno5 = (int) dr.Cells["idArticuloAlterno5"].Value;

				sSQL = string.Format("Exec GuardaDetallePromociones {0}, {1}, {2}, {3}, {4}, {5}, {6}, {7}, {8}, {9}, {10}, {11}, {12}, {13}, {14}, {15}, {16}, {17}, {18}, {19}, {20}, {21}, {22}, {23}, {24}, {25}, {26}, {27}, {28}, {29}, {30}, {31}, {32}, {33}, {34}, {35}, {36}, {37}, {38}, {39}, {40}, {41}, {42}, {43}, {44}, {45}, {46}", 
																					idDetalle, (int)this.txtIdPromocion.Value, 
																					(int)dr.Cells["idArticulo"].Value, (decimal)dr.Cells["CostoPromedio"].Value, (int)dr.Cells["Cantidad"].Value,
																					(decimal)dr.Cells["Precio2"].Value, (decimal)dr.Cells["Precio3"].Value, (decimal)dr.Cells["Precio4"].Value,
																					(decimal)dr.Cells["Descuento"].Value, (decimal)dr.Cells["DescuentoTC"].Value, (decimal)dr.Cells["DescuentoCP"].Value, 
																					(decimal)dr.Cells["PromoPrecio2"].Value, (decimal)dr.Cells["PromoPrecio3"].Value, (decimal)dr.Cells["PromoPrecio4"].Value, (bool)dr.Cells["Principal"].Value,
					idArtAlterno1, idArtAlterno2, idArtAlterno3, idArtAlterno4, idArtAlterno5, (bool)dr.Cells["BloqueaDescuento"].Value, (decimal)dr.Cells["ValorMinimo"].Value, (decimal)dr.Cells["ValorMaximo"].Value,
					(bool)dr.Cells["Baratazos"].Value, (bool)dr.Cells["Combos"].Value, (bool)dr.Cells["Promociones"].Value, (decimal)dr.Cells["Precio6"].Value, (decimal)dr.Cells["DescuentoTCSI"].Value,
					(decimal)dr.Cells["PrecioC3"].Value, (decimal)dr.Cells["PromoPrecioC3"].Value, 
					(decimal)dr.Cells["PrecioC6"].Value, (decimal)dr.Cells["PromoPrecioC6"].Value, 
					(decimal)dr.Cells["PrecioC9"].Value, (decimal)dr.Cells["PromoPrecioC9"].Value, 
					(decimal)dr.Cells["PrecioC12"].Value, (decimal)dr.Cells["PromoPrecioC12"].Value, 
					(decimal)dr.Cells["PrecioC15"].Value, (decimal)dr.Cells["PromoPrecioC15"].Value, 
					(decimal)dr.Cells["PrecioC18"].Value, (decimal)dr.Cells["PromoPrecioC18"].Value, 
					(decimal)dr.Cells["PrecioC21"].Value, (decimal)dr.Cells["PromoPrecioC21"].Value, 
					(decimal)dr.Cells["PrecioC24"].Value, (decimal)dr.Cells["PromoPrecioC24"].Value,
					(bool)dr.Cells["WEBPRINCIPAL"].Value, (decimal)dr.Cells["DescuentoCombo"].Value,
					(decimal)dr.Cells["Aporte"].Value, (int)dr.Cells["UnidadMaxima"].Value, (int)dr.Cells["idCliente"].Value
					);
				
				idDetalle = Funcion.iEscalarSQL(cdsSeteoF, sSQL);

				if (dr.Cells["idDetallePromocion"].Value == DBNull.Value) dr.Cells["idDetallePromocion"].Value = idDetalle; 
			}
			#endregion Detalle 

			

			string sSQLBloqueo = string.Format("Exec RegistraBloqueoTransacciones {0}, 0, {1}, 7", 
				idBloqueaTransacciones, (int)this.txtIdPromocion.Value);
			Funcion.EjecutaSQL(cdsSeteoF, sSQLBloqueo, true);

			this.txtIdPromocion.Enabled = false;
			this.cmbTipo.Enabled = false;
			this.dtInicio.Enabled = false;
			this.dtFin.Enabled = false;
			this.txtDescripcion.Enabled = false;
			this.txtValorMinimo.Enabled = false;
			this.txtValorMaximo.Enabled = false;
			this.txtNroCuotas.Enabled = false;
			this.txtCuotaReal.Enabled = false;
			this.txtCuotaCalculada.Enabled = false;		
			this.txtMargenTotal.Enabled = false;		
			
			EstadoGrid(false);
			
			this.btGrabar.Enabled = false;
			this.btEditar.Enabled = true;
			this.btNuevo.Enabled = true;
			this.btBuscar.Enabled = true;
			this.btCancelar.Enabled = true;
			this.btFinalizar.Enabled = true;
			this.btnActualizaPrecios.Enabled = false;

			this.chkLocales.Enabled = false;
			bSelLocales = false;
			this.chkLocales.Checked = false;

			this.chkCuotas.Enabled = false;
			bSelCuotas = false;
			this.chkCuotas.Checked = false;

			this.chkGrupos.Enabled = false;			
			this.chkGrupos.Checked = false;

			this.chkSubGrupos.Enabled = false;
			bSelSubGrupos = false;
			this.chkSubGrupos.Checked = false;

			this.chkMarcas.Enabled = false;
			bSelMarcas = false;
			this.chkMarcas.Checked = false;			

			this.chkVariable.Enabled = false;	
			this.chkSumaDescuentos.Enabled = false;
			this.chkValor.Enabled = false;
			this.chkAdicional.Enabled = false;
			this.chkOpcional.Enabled = false;
			this.chkCupon.Enabled = false;
			this.chkValidaCargaProducto.Enabled = false;

			bModoEdicion = false;
			bNuevo = false;
		}

		private void cmbTipo_ValueChanged(object sender, System.EventArgs e)
		{
			if (bModoEdicion)
			{
				if (this.cmbTipo.Text.ToString() == "COMBO")
				{
					this.txtValorMinimo.Enabled = false;
					this.txtValorMaximo.Enabled = false;
					this.uGridGrupos.Enabled = false;
					this.uGridSubgrupos.Enabled = false;
					this.uGridMarcas.Enabled = false;
					this.ultraGrid1.Enabled = true;
					this.btnActualizaPrecios.Enabled = true;
					this.chkGrupos.Enabled = false;
					this.chkVariable.Enabled = false;
					this.chkOpcional.Enabled = false;
					this.chkCupon.Enabled = false;
					this.chkValidaCargaProducto.Enabled = false;
				}
				else if (this.cmbTipo.Text.ToString() == "PROMOCIÓN")
				{
					this.txtValorMinimo.Enabled = true;
					this.txtValorMaximo.Enabled = true;
					this.uGridGrupos.Enabled = true;
					this.uGridSubgrupos.Enabled = true;
					this.uGridMarcas.Enabled = true;
					this.ultraGrid1.Enabled = true;
					this.btnActualizaPrecios.Enabled = false;
					this.chkGrupos.Enabled = true;
					this.chkVariable.Enabled = true;
					this.chkSumaDescuentos.Enabled = true;
					this.chkValor.Enabled = true;
					this.chkAdicional.Enabled = true;					
					this.chkOpcional.Enabled = true;
					this.chkCupon.Enabled = true;
					this.chkValidaCargaProducto.Enabled = true;
				}
				else if (this.cmbTipo.Text.ToString() == "ESPECÍFICO")
				{
					this.txtValorMinimo.Enabled = true;
					this.txtValorMaximo.Enabled = true;
					this.uGridGrupos.Enabled = true;
					this.uGridSubgrupos.Enabled = true;
					this.uGridMarcas.Enabled = true;
					this.ultraGrid1.Enabled = true;
					this.btnActualizaPrecios.Enabled = false;
					this.chkVariable.Enabled = false;
					this.chkOpcional.Enabled = false;
					this.chkCupon.Enabled = false;
					this.chkValidaCargaProducto.Enabled = false;
				}
			}
		}

		private void cmbTipo_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (e.KeyChar == 13) this.dtInicio.Focus();
		}

		private void txtDescripcion_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (e.KeyChar == 13) this.txtValorMinimo.Focus();
		}

		private void txtValorMaximo_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
		
		}

		private void Promociones_Closing(object sender, System.ComponentModel.CancelEventArgs e)
		{
			if (bModoEdicion)
			{
				MessageBox.Show("No puede salir mientras edita una Promocion", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				e.Cancel = true;
				return;
			}
		}

		private void uGridCiudades_AfterCellUpdate(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			if (e.Cell.Column.ToString() == "Sel")
			{
				this.uGridCiudades.Update();
				//this.uGridCiudades.UpdateMode = Infragistics.Win.UltraWinGrid.UpdateMode.
				int idCP = 0;
				if (e.Cell.Row.Cells["idPromocionCiudad"].Value != DBNull.Value) idCP = (int)e.Cell.Row.Cells["idPromocionCiudad"].Value; 
							
				string sSQL = string.Format("Exec PromocionesActualizaCiudades {0}, {1}, {2}, {3}, {4}", idCP, (int)this.txtIdPromocion.Value,  (int)e.Cell.Row.Cells["idCiudad"].Value, iNumIngreso, (bool)e.Cell.Row.Cells["Sel"].Value);				  
				idCP = Funcion.iEscalarSQL(cdsSeteoF, sSQL);

				if (e.Cell.Row.Cells["idPromocionCiudad"].Value == DBNull.Value) e.Cell.Row.Cells["idPromocionCiudad"].Value = idCP; 
				else e.Cell.Row.Cells["idPromocionCiudad"].Value = DBNull.Value;
			}
		}

		private void uGridBodegas_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}

		private void uGridBodegas_AfterCellUpdate(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			if (e.Cell.Column.ToString() == "Sel")
			{
				this.uGridBodegas.Update();
			
				int idBP = 0;
				if (e.Cell.Row.Cells["idPromCiudadBodega"].Value != DBNull.Value) idBP = (int)e.Cell.Row.Cells["idPromCiudadBodega"].Value; 
							
				string sSQL = string.Format("Exec PromocionesActualizaBodegas {0}, {1}, {2}, {3}, {4}, {5}", idBP, (int)this.uGridCiudades.ActiveRow.Cells["idPromocionCiudad"].Value, (int)e.Cell.Row.Cells["Bodega"].Value, iNumIngreso, (bool)e.Cell.Row.Cells["Sel"].Value, (int)this.txtIdPromocion.Value);				  
				idBP = Funcion.iEscalarSQL(cdsSeteoF, sSQL);

				if (e.Cell.Row.Cells["idPromCiudadBodega"].Value == DBNull.Value) e.Cell.Row.Cells["idPromCiudadBodega"].Value = idBP; 
				else e.Cell.Row.Cells["idPromCiudadBodega"].Value = DBNull.Value;
			}
		}

		private void uGridTiposDePrecio_AfterCellUpdate(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			if (e.Cell.Column.ToString() == "Sel")
			{
				this.uGridTiposDePrecio.Update();
			
				int idPP = 0;
				if (e.Cell.Row.Cells["idPromoPrecio"].Value != DBNull.Value) idPP = (int)e.Cell.Row.Cells["idPromoPrecio"].Value; 
							
				string sSQL = string.Format("Exec PromocionesActualizaPrecios {0}, {1}, {2}, '{3}', {4}, {5}, {6}, {7}", idPP, 
					(int)this.txtIdPromocion.Value, (int)e.Cell.Row.Cells["idTipoPrecio"].Value, e.Cell.Row.Cells["Operador"].Value.ToString(), (decimal)e.Cell.Row.Cells["Porcentaje"].Value, iNumIngreso, (bool)e.Cell.Row.Cells["Sel"].Value, (decimal)e.Cell.Row.Cells["PorcentajeInicial"].Value);
				idPP = Funcion.iEscalarSQL(cdsSeteoF, sSQL);

				if (e.Cell.Row.Cells["idPromoPrecio"].Value == DBNull.Value) e.Cell.Row.Cells["idPromoPrecio"].Value = idPP; 
				else e.Cell.Row.Cells["idPromoPrecio"].Value = DBNull.Value;
			}

			if (e.Cell.Column.ToString() == "PorcentajeInicial")
			{
				if ((bool)e.Cell.Row.Cells["Sel"].Value)
				{
					int idPP = 0;
					if (e.Cell.Row.Cells["idPromoPrecio"].Value != DBNull.Value) idPP = (int)e.Cell.Row.Cells["idPromoPrecio"].Value; 
							
					string sSQL = string.Format("Update PromocionPrecios Set PorcentajeInicial = {0} Where idPromoPrecio = {1}", (decimal)e.Cell.Row.Cells["PorcentajeInicial"].Value, idPP);
					Funcion.EjecutaSQL(cdsSeteoF, sSQL);
				}
			}

			if (e.Cell.Column.ToString() == "Porcentaje")
			{
				if ((bool)e.Cell.Row.Cells["Sel"].Value)
				{
					int idPP = 0;
					if (e.Cell.Row.Cells["idPromoPrecio"].Value != DBNull.Value) idPP = (int)e.Cell.Row.Cells["idPromoPrecio"].Value; 
							
					string sSQL = string.Format("Update PromocionPrecios Set Porcentaje = {0} Where idPromoPrecio = {1}", (decimal)e.Cell.Row.Cells["Porcentaje"].Value, idPP);
					Funcion.EjecutaSQL(cdsSeteoF, sSQL);
				}
			}
			
			if (e.Cell.Column.ToString() == "Operador")
			{
				if ((bool)e.Cell.Row.Cells["Sel"].Value)
				{
					int idPP = 0;
					if (e.Cell.Row.Cells["idPromoPrecio"].Value != DBNull.Value) idPP = (int)e.Cell.Row.Cells["idPromoPrecio"].Value; 
							
					string sSQL = string.Format("Update PromocionPrecios Set Operador = '{0}' Where idPromoPrecio = {1}", e.Cell.Row.Cells["Operador"].Value.ToString(), idPP);
					Funcion.EjecutaSQL(cdsSeteoF, sSQL);				
				}
			}
		}

		private void uGridBancos_AfterCellUpdate(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			if (e.Cell.Column.ToString() == "Sel")
			{
				this.uGridBancos.Update();

				int idPB = 0;

				if (e.Cell.Row.Cells["idPromocionBanco"].Value != DBNull.Value) idPB = (int)e.Cell.Row.Cells["idPromocionBanco"].Value; 
							
				string sSQL = string.Format("Exec PromocionesActualizaBancos {0}, {1}, {2}, {3}, {4}", idPB, (int)this.txtIdPromocion.Value, (int)e.Cell.Row.Cells["idBanco"].Value, iNumIngreso, (bool)e.Cell.Row.Cells["Sel"].Value);				  
				idPB = Funcion.iEscalarSQL(cdsSeteoF, sSQL);

				if (e.Cell.Row.Cells["idPromocionBanco"].Value == DBNull.Value) e.Cell.Row.Cells["idPromocionBanco"].Value = idPB; 
				else e.Cell.Row.Cells["idPromocionBanco"].Value = DBNull.Value;
			}
		}
		
		private void uGridTarjetas_Click(object sender, System.EventArgs e)
		{
			if (this.uGridTarjetas.Rows.Count == 0) return;

			if ((bool)this.uGridTarjetas.ActiveRow.Cells["Sel"].Value)
			{
				string sSQL = string.Format("Exec ListaPlanesPromocion {0}, {1}, {2}", 
					(int)this.txtIdPromocion.Value, (int)this.uGridBancos.ActiveRow.Cells["idPromocionBanco"].Value, (int)this.uGridTarjetas.ActiveRow.Cells["idPromocionBancoTarjeta"].Value);
				this.uGridPlanes.DataSource = Funcion.dvProcedimiento(cdsSeteoF, sSQL);								
			}
			else if (!(bool)this.uGridTarjetas.ActiveRow.Cells["Sel"].Value)
			{				
				this.uGridPlanes.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Select Null As idPromocionesPlanes, NULL As idPromocionBancoTarjeta, Null As idPromocionBanco, idPlan, Null As idPromocion, Descripcion, CONVERT(Bit, 0) As Sel From Planes Where idPlan = 0");
				this.uGridPlazos.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Select Null As idPromocionesPlazos, idPlazo, Null As idPromocion, Descripcion, CONVERT(Bit, 0) As Sel  From Plazos Where idPlazo = 0");
			}
		}

		private void uGridTarjetas_AfterCellUpdate(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			if (e.Cell.Column.ToString() == "Sel")
			{
				this.uGridTarjetas.Update();
			
				int idTP = 0;
				if (e.Cell.Row.Cells["idPromocionBancoTarjeta"].Value != DBNull.Value) idTP = (int)e.Cell.Row.Cells["idPromocionBancoTarjeta"].Value; 
							
				string sSQL = string.Format("Exec PromocionesActualizaTarjetas {0}, {1}, {2}, {3}, {4}, {5}", 
					idTP, (int)this.uGridBancos.ActiveRow.Cells["idPromocionBanco"].Value, (int)this.txtIdPromocion.Value, (int)e.Cell.Row.Cells["idTarjeta"].Value, iNumIngreso, (bool)e.Cell.Row.Cells["Sel"].Value);				  
				idTP = Funcion.iEscalarSQL(cdsSeteoF, sSQL);

				if (e.Cell.Row.Cells["idPromocionBancoTarjeta"].Value == DBNull.Value) e.Cell.Row.Cells["idPromocionBancoTarjeta"].Value = idTP; 
				else e.Cell.Row.Cells["idPromocionBancoTarjeta"].Value = DBNull.Value;
			}
		}

		private void uGridGrupos_Click(object sender, System.EventArgs e)
		{
			if (this.uGridGrupos.Rows.Count == 0) return;

			if (this.uGridGrupos.ActiveRow.Cells["idPromocionGrupo"].Value == System.DBNull.Value) return;
			
			if ((bool)this.uGridGrupos.ActiveRow.Cells["Sel"].Value)
			{
				int idGrupo = 0;
				if (this.uGridGrupos.ActiveRow != null) idGrupo = (int)this.uGridGrupos.ActiveRow.Cells["idPromocionGrupo"].Value;
				string sSQL = string.Format("Exec ListaSubGruposPromociones {0}, {1}, {2}", idGrupo, (int)this.txtIdPromocion.Value, (int)this.uGridGrupos.ActiveRow.Cells["idGrupoArticulo"].Value);
				this.uGridSubgrupos.DataSource = Funcion.dvProcedimiento(cdsSeteoF, sSQL);
				
				this.chkSubGrupos.Checked = false;
				this.chkMarcas.Checked = false;
				if (bEditar) this.chkSubGrupos.Enabled = true;
				bSelSubGrupos = true;
				
			}
			else if (!(bool)this.uGridGrupos.ActiveRow.Cells["Sel"].Value)
			{				
				this.uGridSubgrupos.DataSource = Funcion.dvProcedimiento(cdsSeteoF, string.Format("Exec ListaSubGruposPromociones -1, {0}, -1", (int)this.txtIdPromocion.Value));

				this.chkSubGrupos.Enabled = false;
				this.chkMarcas.Checked = false;
				bSelSubGrupos = false;
				this.chkSubGrupos.Checked = false;
			}
		}

		private void uGridGrupos_AfterCellUpdate(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			
		}
		
		private void uGridSubgrupos_AfterCellUpdate(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			
		}

		private void uGridSubgrupos_Click(object sender, System.EventArgs e)
		{
			if (this.uGridSubgrupos.Rows.Count == 0) return;

			if (this.uGridSubgrupos.ActiveRow.Cells["idPromocionSubGrupo"].Value == System.DBNull.Value) return;

			if ((bool)this.uGridSubgrupos.ActiveRow.Cells["Sel"].Value)
			{
				int idSubGrupo = 0;
				if (this.uGridSubgrupos.ActiveRow != null) idSubGrupo = (int)this.uGridSubgrupos.ActiveRow.Cells["idPromocionSubGrupo"].Value;
				string sSQL = string.Format("Exec ListaMarcasPromociones {0}, {1}, {2}", idSubGrupo, (int)this.txtIdPromocion.Value, (int)this.uGridSubgrupos.ActiveRow.Cells["idSubGrupo"].Value);
				this.uGridMarcas.DataSource = Funcion.dvProcedimiento(cdsSeteoF, sSQL);
				
				this.chkMarcas.Checked = false;
				if (bEditar) this.chkMarcas.Enabled = true;
				bSelMarcas = true;
			}
			else if (!(bool)this.uGridGrupos.ActiveRow.Cells["Sel"].Value)
			{				
				this.uGridMarcas.DataSource = Funcion.dvProcedimiento(cdsSeteoF, string.Format("Exec ListaMarcasPromociones -1, {0}, -1", (int)this.txtIdPromocion.Value));
				this.chkMarcas.Enabled = false;
				bSelMarcas = false;
				this.chkMarcas.Checked = false;
			}
		}

		private void uGridMarcas_AfterCellUpdate(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			
		}

		private void Consulta(int idPromocion)
		{
			try
			{
				#region Encabezado Campos Individuales
				string sSQL = string.Format("Exec ConsultaIndividualPromociones {0}", idPromocion);
				SqlDataReader drPromocion = Funcion.rEscalarSQL(cdsSeteoF, sSQL, true);
				drPromocion.Read();
				if (drPromocion.HasRows)
				{
					//idPromocion, ISNULL(idArticulo, 0), Tipo, FechaDesde, FechaHasta, Descripcion, PrecioMinimo, PrecioMaximo, Cuotas, CuotaReal, CuotaCalculada, NumIngreso, Estado, Variable 
					this.txtIdPromocion.Value = idPromocion;
					this.txtidArticulo.Value  = drPromocion.GetInt32(1);
					this.cmbTipo.Text = drPromocion.GetValue(2).ToString();
					this.dtInicio.Value = drPromocion.GetDateTime(3);
					this.dtFin.Value = drPromocion.GetDateTime(4);
					this.txtDescripcion.Text = drPromocion.GetValue(5).ToString();
					this.txtValorMinimo.Value = drPromocion.GetDecimal(6);
					this.txtValorMaximo.Value = drPromocion.GetDecimal(7);
					this.txtNroCuotas.Value = drPromocion.GetInt32(8);
					this.txtCuotaReal.Value = drPromocion.GetDecimal(9);
					this.txtCuotaCalculada.Value = drPromocion.GetDecimal(10);					
					iNumIngreso = drPromocion.GetInt32(11);
					this.lblFinalizado.Text = drPromocion.GetValue(12).ToString();
					this.chkVariable.Checked = drPromocion.GetBoolean(13);
					this.chkSumaDescuentos.Checked = drPromocion.GetBoolean(14);
					this.chkValor.Checked = drPromocion.GetBoolean(15);
					this.chkAdicional.Checked = drPromocion.GetBoolean(16);
					this.chkOpcional.Checked = drPromocion.GetBoolean(17);
					this.chkCupon.Checked = drPromocion.GetBoolean(18);
					this.chkValidaCargaProducto.Checked = drPromocion.GetBoolean(19);
					this.txtMargenTotal.Value = drPromocion.GetDecimal(20);
				}
				drPromocion.Close();
				#endregion Encabezado Campos Individuales

				this.uGridCiudades.DataSource = Funcion.dvProcedimiento(cdsSeteoF, string.Format("Exec ListaCiudadesPromocion {0}", (int)this.txtIdPromocion.Value));			
				this.uGridTiposDePrecio.DataSource = Funcion.dvProcedimiento(cdsSeteoF, string.Format("Exec ListaPreciosPromocion {0}", (int)this.txtIdPromocion.Value));
				this.uGridGrupos.DataSource = Funcion.dvProcedimiento(cdsSeteoF, string.Format("Exec ListaGruposPromocion {0}", (int)this.txtIdPromocion.Value));
				this.uGridSubgrupos.DataSource = Funcion.dvProcedimiento(cdsSeteoF, string.Format("Exec ListaSubGruposPromociones 0, 0, 0"));//, (int)this.txtIdPromocion.Value));
				this.uGridMarcas.DataSource = Funcion.dvProcedimiento(cdsSeteoF, string.Format("Exec ListaMarcasPromociones 0, 0, 0"));//, (int)this.txtIdPromocion.Value));
					
				this.ultraGrid1.DataSource = Funcion.dvProcedimiento(cdsSeteoF, string.Format("Exec BuscaDetallePromocion {0}", (int)this.txtIdPromocion.Value));

				this.btCancelar.Enabled = true;
				this.btEditar.Enabled = true;
				this.btFinalizar.Enabled = true;

				if (this.lblFinalizado.Text.ToString() == "ACTIVO") this.btFinalizar.Text = "Finalizar";
				else if (this.lblFinalizado.Text.ToString() == "FINALIZADO") this.btFinalizar.Text = "Activar";
			}
			catch(Exception Exc)
			{
				MessageBox.Show(string.Format("Error al Consultar: {0}", Exc), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
			}
		}

		private void btBuscar_Click(object sender, System.EventArgs e)
		{
			this.btCancelar_Click(sender, e);

			using (frmBuscaPromociones miBusqueda = new frmBuscaPromociones())
			{
				if (DialogResult.OK == miBusqueda.ShowDialog())
				{
					this.Consulta((int) miBusqueda.uGridPromociones.ActiveRow.Cells["idPromocion"].Value);
				}
			}
		}

		private void btEditar_Click(object sender, System.EventArgs e)
		{		
			if (Funcion.ValidaBloqueo((int)this.txtIdPromocion.Value, 7, cdsSeteoF)) return;

			this.cmbTipo.Enabled = true;
			this.dtInicio.Enabled = true;
			this.dtFin.Enabled = true;
			this.txtDescripcion.Enabled = true;
			this.txtValorMinimo.Enabled = true;
			this.txtValorMaximo.Enabled = true;
			this.txtNroCuotas.Enabled = true;
			this.txtCuotaReal.Enabled = true;
			this.txtCuotaCalculada.Enabled = true;
			this.txtMargenTotal.Enabled = true;

			EstadoGrid(true);

			this.btNuevo.Enabled = false;
			this.btBuscar.Enabled = false;
			this.btEditar.Enabled = false;
			this.btGrabar.Enabled = true;
			this.btFinalizar.Enabled = false;
			this.btCancelar.Enabled = true;

			if (this.cmbTipo.Text.ToString() == "COMBO")
			{
				this.btnActualizaPrecios.Enabled = true;
				this.uGridGrupos.Enabled = false;
				this.uGridSubgrupos.Enabled = false;
				this.uGridMarcas.Enabled = false;				
			}
			else 
			{				
				this.uGridGrupos.Enabled = true;
				this.uGridSubgrupos.Enabled = true;
				this.uGridMarcas.Enabled = true;
				this.chkGrupos.Enabled = true;
				this.chkSubGrupos.Enabled = true;
				this.chkSubGrupos.Enabled = true;
				this.chkVariable.Enabled = true;
				this.chkSumaDescuentos.Enabled = true;
				this.chkValor.Enabled = true;
				this.chkAdicional.Enabled = true;
				this.chkOpcional.Enabled = true;
				this.chkCupon.Enabled = true;
				this.chkValidaCargaProducto.Enabled = true;
			}

			bNuevo = false;
			bModoEdicion = true;
			iNumIngreso += 1;

			string sSQLBloqueo = string.Format("Exec RegistraBloqueoTransacciones 0, 1, {0}, 7", (int)this.txtIdPromocion.Value);
			idBloqueaTransacciones = Funcion.iEscalarSQL(cdsSeteoF, sSQLBloqueo, true);

			this.cmbTipo.Focus();
		}

		private void ultraGrid1_AfterRowsDeleted(object sender, System.EventArgs e)
		{
			CalculaCuota();
		}

		private void btFinalizar_Click(object sender, System.EventArgs e)
		{
			if (Funcion.ValidaBloqueo((int)this.txtIdPromocion.Value, 7, cdsSeteoF)) return;

			if (this.lblFinalizado.Text.ToString() == "ACTIVO") 
			{
				Funcion.EjecutaSQL(cdsSeteoF, string.Format("Update CompraPromocion Set Estado = 'FINALIZADO' Where idPromocion = {0}", (int)this.txtIdPromocion.Value));
				Funcion.EjecutaSQL(cdsSeteoF, string.Format("Update Articulo SET Descontinuado = 1 Where idArticulo = {0}", (int)this.txtidArticulo.Value));
				this.lblFinalizado.Text = "FINALIZADO";
				this.btFinalizar.Text = "Activar";				
				this.btFinalizar.Enabled = true;
			}
			else
			{
        Funcion.EjecutaSQL(cdsSeteoF, string.Format("Update CompraPromocion Set Estado = 'ACTIVO' Where idPromocion = {0}", (int)this.txtIdPromocion.Value));
				Funcion.EjecutaSQL(cdsSeteoF, string.Format("Update Articulo SET Descontinuado = 0 Where idArticulo = {0}", (int)this.txtidArticulo.Value));
				this.lblFinalizado.Text = "ACTIVO";				
				this.btFinalizar.Text = "Finalizar";
				this.btFinalizar.Enabled = true;
			}
		}
		
		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);	
		}

		private void ultraGrid1_BeforeRowsDeleted(object sender, Infragistics.Win.UltraWinGrid.BeforeRowsDeletedEventArgs e)
		{
			if (!bModoEdicion) return;
			if (DialogResult.No ==	MessageBox.Show("¿Esta seguro de borrar la filas seleccionadas?", "Point Technology", MessageBoxButtons.YesNo, MessageBoxIcon.Question)) e.Cancel = true;
			else
			{
				for (int i=0; i<e.Rows.Length; i++)
				{				
				
					if ((int)e.Rows[i].Cells["idDetallePromocion"].Value > 0)
					{
						string sSQL = string.Format("Delete From DetallePromocion Where idDetallePromocion = {0}", (int)e.Rows[i].Cells["idDetallePromocion"].Value);
						Funcion.EjecutaSQL(cdsSeteoF, sSQL, true);
					}
					
					e.DisplayPromptMsg = false;
				}
			}
		}

		private void btnActualizaPrecios_Click(object sender, System.EventArgs e)
		{
			if (this.ultraGrid1.Rows.Count == 0)
			{
				MessageBox.Show("No hay Productos en la lista", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return;
			}

			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in ultraGrid1.Rows.All)
			{
				string sSQL = string.Format("Select CostoUltimo, Precio2, Precio3, ROUND(Precio4, 2) From Articulo Where idArticulo = {0}", (int)dr.Cells["idArticulo"].Value);
				SqlDataReader dReader = Funcion.rEscalarSQL(cdsSeteoF, sSQL, true);
				dReader.Read();
				if (dReader.HasRows)
				{
					dr.Cells["CostoPromedio"].Value = dReader.GetDecimal(0);
					dr.Cells["Precio2"].Value = dReader.GetDecimal(1);
					dr.Cells["Precio3"].Value = dReader.GetDecimal(2);
					dr.Cells["Precio4"].Value = dReader.GetDecimal(3);
					dReader.Close();

					dr.Cells["PromoPrecio2"].Value = Math.Round(Convert.ToDecimal(dr.Cells["Precio2"].Value) - Math.Round(((Convert.ToDecimal(dr.Cells["Precio2"].Value) * Convert.ToDecimal(dr.Cells["Descuento"].Value)) / 100), 2), 2); 
					dr.Cells["PromoPrecio3"].Value = Math.Round(Convert.ToDecimal(dr.Cells["Precio3"].Value) - Math.Round(((Convert.ToDecimal(dr.Cells["Precio3"].Value) * Convert.ToDecimal(dr.Cells["DescuentoTC"].Value)) / 100), 2), 2);
					dr.Cells["PromoPrecio4"].Value = Math.Round(Convert.ToDecimal(dr.Cells["Precio4"].Value) - Math.Round(((Convert.ToDecimal(dr.Cells["Precio4"].Value) * Convert.ToDecimal(dr.Cells["DescuentoCP"].Value)) / 100), 2), 2);

					CalculaCuota();
				}
				dReader.Close();
			}
		}

		private void ultraGrid1_AfterRowInsert(object sender, Infragistics.Win.UltraWinGrid.RowEventArgs e)
		{
			if (bModoEdicion)
			{
				e.Row.Cells["Principal"].Value = false;		
				e.Row.Cells["BloqueaDescuento"].Value = false;
				e.Row.Cells["ValorMinimo"].Value = 0.00m;
				e.Row.Cells["ValorMaximo"].Value = 0.00m;		
				e.Row.Cells["Baratazos"].Value = false;		
				e.Row.Cells["Combos"].Value = false;
				e.Row.Cells["Promociones"].Value = false;
				e.Row.Cells["Cantidad"].Value = 1;
				e.Row.Cells["DescuentoCombo"].Value = 0.00m;
				e.Row.Cells["WEBPRINCIPAL"].Value = false;	
				e.Row.Cells["Aporte"].Value = 0.00m;	
				e.Row.Cells["UnidadMaxima"].Value = 0;
				//if (this.cmbTipo.Text == "COMBO PAYJOY") e.Row.Cells["DescuentoTC"].Value = 0.00m;	
			}
		}

		private void chkLocales_CheckedChanged(object sender, System.EventArgs e)
		{
			if (bSelLocales)
			{
				if (this.chkLocales.Checked)
				{
					foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridBodegas.Rows.All)
					{
						dr.Cells["Sel"].Value = true;
					}
				}
				else
				{
					foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridBodegas.Rows.All)
					{
						dr.Cells["Sel"].Value = false;
					}
				}
			}
		}

		private void chkGrupos_CheckedChanged(object sender, System.EventArgs e)
		{
			if (this.chkGrupos.Checked)
			{
				foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridGrupos.Rows.All)
				{
					dr.Cells["Sel"].Value = true;
				}
			}
			else
			{
				foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridGrupos.Rows.All)
				{
					dr.Cells["Sel"].Value = false;
				}
			}
		}

		private void chkSubGrupos_CheckedChanged(object sender, System.EventArgs e)
		{
			if (bSelSubGrupos)
			{
				if (this.chkSubGrupos.Checked)
				{
					foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridSubgrupos.Rows.All)
					{
						dr.Cells["Sel"].Value = true;
					}
				}
				else
				{
					foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridSubgrupos.Rows.All)
					{
						dr.Cells["Sel"].Value = false;
					}
				}
			}
		}

		private void chkMarcas_CheckedChanged(object sender, System.EventArgs e)
		{
			if (bSelMarcas)
			{
				if (this.chkMarcas.Checked)
				{
					foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridMarcas.Rows.All)
					{
						dr.Cells["Sel"].Value = true;
					}
				}
				else
				{
					foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridMarcas.Rows.All)
					{
						dr.Cells["Sel"].Value = false;
					}
				}
			}
		}

		private void uGridPlanes_AfterCellUpdate(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			if (e.Cell.Column.ToString() == "Sel")
			{
				this.uGridPlanes.Update();
				
				int idPromoPlanes = 0;
				int idPromoBancos = 0;
				int idPromoTarjetas = 0;				

				if (e.Cell.Row.Cells["idPromocionesPlanes"].Value != DBNull.Value) idPromoPlanes = (int)e.Cell.Row.Cells["idPromocionesPlanes"].Value; 
				if (this.uGridBancos.ActiveRow.Cells["idPromocionBanco"].Value != System.DBNull.Value) idPromoBancos = (int)this.uGridBancos.ActiveRow.Cells["idPromocionBanco"].Value;
				if (this.uGridTarjetas.ActiveRow.Cells["idPromocionBancoTarjeta"].Value != System.DBNull.Value) idPromoTarjetas = (int)this.uGridTarjetas.ActiveRow.Cells["idPromocionBancoTarjeta"].Value;
							
				string sSQL = string.Format("Exec ActualizaPromocionesPlanes {0}, {1}, {2}, {3}, {4}, {5}, {6}", 
					idPromoPlanes, idPromoBancos, idPromoTarjetas, (int)this.txtIdPromocion.Value, (int)e.Cell.Row.Cells["idPlan"].Value, iNumIngreso, (bool)e.Cell.Row.Cells["Sel"].Value);				  
				idPromoPlanes = Funcion.iEscalarSQL(cdsSeteoF, sSQL);

				if (e.Cell.Row.Cells["idPromocionesPlanes"].Value == DBNull.Value) e.Cell.Row.Cells["idPromocionesPlanes"].Value = idPromoPlanes; 
				else e.Cell.Row.Cells["idPromocionesPlanes"].Value = DBNull.Value;
			}
		}

		private void uGridMarcas_Click(object sender, System.EventArgs e)
		{
			if (this.uGridMarcas.Rows.Count == 0) return;

			if (this.uGridMarcas.ActiveRow.Cells["idPromocionMarca"].Value == System.DBNull.Value) return;

			if ((bool)this.uGridMarcas.ActiveRow.Cells["Sel"].Value)
			{
				int idSubGrupo = 0;
				if (this.uGridMarcas.ActiveRow != null) idSubGrupo = (int)this.uGridMarcas.ActiveRow.Cells["idPromocionMarca"].Value;
				string sSQL = string.Format("Exec ListaPromocionesModelos {0}, {1}, {2}, {3}, {4}", 
					(int)this.txtIdPromocion.Value, (int)this.uGridMarcas.ActiveRow.Cells["idPromocionMarca"].Value, 
					(int)this.uGridGrupos.ActiveRow.Cells["idGrupoArticulo"].Value, (int)this.uGridSubgrupos.ActiveRow.Cells["idSubGrupo"].Value, 
					(int)this.uGridMarcas.ActiveRow.Cells["idMarca"].Value);
				this.uGridModelos.DataSource = Funcion.dvProcedimiento(cdsSeteoF, sSQL);
				
				this.chkMarcas.Checked = false;
				this.chkMarcas.Enabled = true;
				bSelMarcas = true;
			}
			else if (!(bool)this.uGridMarcas.ActiveRow.Cells["Sel"].Value)
			{				
				this.uGridModelos.DataSource = Funcion.dvProcedimiento(cdsSeteoF, string.Format("Exec [ListaPromocionesModelos] {0}, -1, 0, 0, 0", (int)this.txtIdPromocion.Value));
				this.chkMarcas.Enabled = false;
				bSelMarcas = false;
				this.chkMarcas.Checked = false;
			}
		}

		private void Promociones_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (!bModoEdicion)
			{
				if (e.KeyCode == Keys.N) this.btNuevo_Click(sender, e);
				if (e.KeyCode == Keys.B) this.btBuscar_Click(sender, e);
				if (e.KeyCode == Keys.E) this.btEditar_Click(sender, e);
				if (e.KeyCode == Keys.Escape) this.Close();
				if (e.KeyCode == Keys.F5) this.Consulta((int)this.txtIdPromocion.Value);
			}
			else
			{
				if (e.KeyCode == Keys.Escape) this.btCancelar_Click(sender, e);
				if (e.KeyCode == Keys.F12) this.btGrabar_Click(sender, e);				
			}
		}

		private void uGridModelos_AfterCellUpdate(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			
		}

		private void uGridPlazos_AfterCellUpdate(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			if (e.Cell.Column.ToString() == "Sel")
			{
				this.uGridPlazos.Update();
				
				int idPromoPlazos = 0;
				int idPromoBancos = 0;
				int idPromoTarjetas = 0;
				int idPromoPlanes = 0;
				int MesesDeGracia = 0;

				if (e.Cell.Row.Cells["idPromocionesPlazos"].Value != DBNull.Value) idPromoPlazos = (int)e.Cell.Row.Cells["idPromocionesPlazos"].Value; 
				if (this.uGridBancos.ActiveRow.Cells["idPromocionBanco"].Value != System.DBNull.Value) idPromoBancos = (int)this.uGridBancos.ActiveRow.Cells["idPromocionBanco"].Value;
				if (this.uGridTarjetas.ActiveRow.Cells["idPromocionBancoTarjeta"].Value != System.DBNull.Value) idPromoTarjetas = (int)this.uGridTarjetas.ActiveRow.Cells["idPromocionBancoTarjeta"].Value;
				if (this.uGridPlanes.ActiveRow.Cells["idPromocionesPlanes"].Value != System.DBNull.Value) idPromoPlanes = (int)this.uGridPlanes.ActiveRow.Cells["idPromocionesPlanes"].Value;
				if (e.Cell.Row.Cells["MesesDeGracia"].Value != DBNull.Value) MesesDeGracia = (int)e.Cell.Row.Cells["MesesDeGracia"].Value; 
				
				string sSQL = string.Format("Exec ActualizaPromocionesPlazos {0}, {1}, {2}, {3}, {4}, {5}, {6}, {7}", 
					idPromoPlazos, idPromoBancos, idPromoTarjetas, idPromoPlanes, (int)this.txtIdPromocion.Value, (int)e.Cell.Row.Cells["idPlazo"].Value, iNumIngreso, MesesDeGracia);
				idPromoPlanes = Funcion.iEscalarSQL(cdsSeteoF, sSQL);

				if (e.Cell.Row.Cells["idPromocionesPlazos"].Value == DBNull.Value) e.Cell.Row.Cells["idPromocionesPlazos"].Value = idPromoPlazos; 
				else e.Cell.Row.Cells["idPromocionesPlazos"].Value = DBNull.Value;
			}
		}

		private void uGridPlanes_Click(object sender, System.EventArgs e)
		{
			if (this.uGridPlanes.Rows.Count == 0) return;

			if ((bool)this.uGridPlanes.ActiveRow.Cells["Sel"].Value)
			{
				string sSQL = string.Format("Exec ListaPlazosPromociones {0}, {1}, {2}, {3}", 
					(int)this.txtIdPromocion.Value, (int)this.uGridBancos.ActiveRow.Cells["idPromocionBanco"].Value, 
					(int)this.uGridTarjetas.ActiveRow.Cells["idPromocionBancoTarjeta"].Value, (int)this.uGridPlanes.ActiveRow.Cells["idPromocionesPlanes"].Value);
				this.uGridPlazos.DataSource = Funcion.dvProcedimiento(cdsSeteoF, sSQL);								
			}
			else if (!(bool)this.uGridPlanes.ActiveRow.Cells["Sel"].Value)
			{				
				this.uGridPlazos.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Select Null As idPromocionesPlazos, idPlazo, Null As idPromocion, Descripcion, CONVERT(Bit, 0) As Sel  From Plazos Where idPlazo = 0");
			}
		}

		private void uGridCuotas_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			FuncionesProcedimientos.DesplazamientoGrids(sender, e, this.uGridCuotas);
		}

		private void uGridModelos_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			FuncionesProcedimientos.DesplazamientoGrids(sender, e, this.uGridModelos);
		}

		private void uGridCuotas_AfterCellUpdate(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			if (e.Cell.Column.ToString() == "Sel")
			{
				this.uGridCuotas.Update();

				int idPC = 0;

				if (e.Cell.Row.Cells["idPromocionesCuotas"].Value != DBNull.Value) idPC = (int)e.Cell.Row.Cells["idPromocionesCuotas"].Value; 
							
				string sSQL = string.Format("Exec PromocionesActualizaCuotas {0}, {1}, {2}, {3}, {4}, {5}", 
					idPC, (int)this.txtIdPromocion.Value, (int)e.Cell.Row.Cells["Cuota"].Value, iNumIngreso, (bool)e.Cell.Row.Cells["Sel"].Value, (Decimal)e.Cell.Row.Cells["Porcentaje_Entrada"].Value);				  
				idPC = Funcion.iEscalarSQL(cdsSeteoF, sSQL);

				if (e.Cell.Row.Cells["idPromocionesCuotas"].Value == DBNull.Value) e.Cell.Row.Cells["idPromocionesCuotas"].Value = idPC; 
				else e.Cell.Row.Cells["idPromocionesCuotas"].Value = DBNull.Value;
			}
		}

		private void uGridCiudades_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}

		private void chkCuotas_CheckedChanged(object sender, System.EventArgs e)
		{
			if (bSelCuotas)
			{
				if (this.chkCuotas.Checked)
				{
					foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridCuotas.Rows.All)
					{
						dr.Cells["Sel"].Value = true;
					}
				}
				else
				{
					foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridCuotas.Rows.All)
					{
						dr.Cells["Sel"].Value = false;
					}
				}
			}
		}

		private void uGridMarcas_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}

		private void txtNroCuotas_ValueChanged(object sender, System.EventArgs e)
		{
		

		}

		private void ultraGrid1_InitializeLogicalPrintPage(object sender, Infragistics.Win.UltraWinGrid.CancelableLogicalPrintPageEventArgs e)
		{
		
		}

		private void uGridGrupos_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}

		private void uGridGrupos_CellChange(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			if (e.Cell.Column.ToString() == "Sel")
			{
				this.uGridGrupos.Update();

				int idPG = 0;

				if (e.Cell.Row.Cells["idPromocionGrupo"].Value != DBNull.Value) idPG = (int)e.Cell.Row.Cells["idPromocionGrupo"].Value; 
							
				string sSQL = string.Format("Exec PromocionesActualizaGrupos {0}, {1}, {2}, {3}, {4}", 
					idPG, (int)this.txtIdPromocion.Value, (int)e.Cell.Row.Cells["idGrupoArticulo"].Value, iNumIngreso, (bool)e.Cell.Row.Cells["Sel"].Value);				  
				idPG = Funcion.iEscalarSQL(cdsSeteoF, sSQL);

				if (e.Cell.Row.Cells["idPromocionGrupo"].Value == DBNull.Value) e.Cell.Row.Cells["idPromocionGrupo"].Value = idPG; 
				else e.Cell.Row.Cells["idPromocionGrupo"].Value = DBNull.Value;
			}		
		}

		private void uGridSubgrupos_CellChange(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			if (e.Cell.Column.ToString() == "Sel")
			{
				this.uGridSubgrupos.Update();
				
				int idTS = 0;
				int idPG = 0;

				if (this.uGridGrupos.ActiveRow != null && (bool)this.uGridGrupos.ActiveRow.Cells["Sel"].Value) idPG = (int)this.uGridGrupos.ActiveRow.Cells["idPromocionGrupo"].Value;

				if (e.Cell.Row.Cells["idPromocionSubgrupo"].Value != DBNull.Value) idTS = (int)e.Cell.Row.Cells["idPromocionSubgrupo"].Value; 
							
				string sSQL = string.Format("Exec PromocionesActualizaSubgrupos {0}, {1}, {2}, {3}, {4}, {5}", idTS, idPG, (int)this.txtIdPromocion.Value, (int)e.Cell.Row.Cells["idSubgrupo"].Value, iNumIngreso, (bool)e.Cell.Row.Cells["Sel"].Value);				  
				idTS = Funcion.iEscalarSQL(cdsSeteoF, sSQL);

				if (e.Cell.Row.Cells["idPromocionSubgrupo"].Value == DBNull.Value) e.Cell.Row.Cells["idPromocionSubgrupo"].Value = idTS; 
				else e.Cell.Row.Cells["idPromocionSubGrupo"].Value = DBNull.Value;
			}
		}

		private void uGridMarcas_CellChange(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			if (e.Cell.Column.ToString() == "Sel")
			{
				this.uGridMarcas.Update();
				
				int idPM = 0;
				int idPS = 0;
				int idPG = 0;

				if (this.uGridGrupos.ActiveRow != null && (bool)this.uGridGrupos.ActiveRow.Cells["Sel"].Value) idPG = (int)this.uGridGrupos.ActiveRow.Cells["idPromocionGrupo"].Value;

				if (this.uGridSubgrupos.ActiveRow != null && (bool)this.uGridSubgrupos.ActiveRow.Cells["Sel"].Value) idPS = (int)this.uGridSubgrupos.ActiveRow.Cells["idPromocionSubGrupo"].Value;

				if (e.Cell.Row.Cells["idPromocionMarca"].Value != DBNull.Value) idPM = (int)e.Cell.Row.Cells["idPromocionMarca"].Value; 
							
				string sSQL = string.Format("Exec PromocionesActualizaMarcas {0}, {1}, {2}, {3}, {4}, {5}, {6}", idPM, idPG, idPS, (int)this.txtIdPromocion.Value, (int)e.Cell.Row.Cells["idMarca"].Value, iNumIngreso, (bool)e.Cell.Row.Cells["Sel"].Value);				  
				idPM = Funcion.iEscalarSQL(cdsSeteoF, sSQL);

				if (e.Cell.Row.Cells["idPromocionMarca"].Value == DBNull.Value) e.Cell.Row.Cells["idPromocionMarca"].Value = idPM; 
				else e.Cell.Row.Cells["idPromocionMarca"].Value = DBNull.Value;
			}
		}

		private void uGridModelos_CellChange(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			if (e.Cell.Column.ToString() == "Sel")
			{
				this.uGridModelos.Update();
				
				int idPM = 0;
				int idPS = 0;
				int idPG = 0;
				int idPMod = 0;

				if (this.uGridGrupos.ActiveRow != null && (bool)this.uGridGrupos.ActiveRow.Cells["Sel"].Value) idPG = (int)this.uGridGrupos.ActiveRow.Cells["idPromocionGrupo"].Value;

				if (this.uGridSubgrupos.ActiveRow != null && (bool)this.uGridSubgrupos.ActiveRow.Cells["Sel"].Value) idPS = (int)this.uGridSubgrupos.ActiveRow.Cells["idPromocionSubGrupo"].Value;

				if (this.uGridMarcas.ActiveRow != null && (bool)this.uGridMarcas.ActiveRow.Cells["Sel"].Value) idPM = (int)this.uGridMarcas.ActiveRow.Cells["idPromocionMarca"].Value; 

				if (e.Cell.Row.Cells["idPromocionModelo"].Value != DBNull.Value) idPMod = (int)e.Cell.Row.Cells["idPromocionModelo"].Value; 
							
				string sSQL = string.Format("Exec PromocionesActualizaModelos {0}, {1}, {2}, {3}, {4}, {5}", idPMod, (int)this.txtIdPromocion.Value, idPG, idPS, idPM, 
					(int)e.Cell.Row.Cells["idArticulo"].Value, (bool)e.Cell.Row.Cells["Sel"].Value);
				idPMod = Funcion.iEscalarSQL(cdsSeteoF, sSQL);

				if (e.Cell.Row.Cells["idPromocionModelo"].Value == DBNull.Value) e.Cell.Row.Cells["idPromocionModelo"].Value = idPMod;
				else e.Cell.Row.Cells["idPromocionModelo"].Value = DBNull.Value;
			}		
		}

		private void uGridModelos_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}					
	}
}

