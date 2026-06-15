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
	public class frmCatalogo1 : DevExpress.XtraEditors.XtraForm
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
		private C1.Data.C1DataRow drSeteoF;
		private System.Windows.Forms.Label label15;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtValorMaximo;
		private System.Windows.Forms.Label label3;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtValorMinimo;
		private System.Windows.Forms.Label label2;
		private System.Windows.Forms.Label label16;
		private Infragistics.Win.UltraWinGrid.UltraGrid uGridMarcas;
		private Infragistics.Win.UltraWinGrid.UltraGrid uGridSubgrupos;
		private Infragistics.Win.UltraWinGrid.UltraGrid uGridGrupos;
		private C1.Data.C1DataSet cdsPromocion;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource1;
		private DevExpress.XtraEditors.SimpleButton btNuevo;
		private DevExpress.XtraEditors.SimpleButton btSalir;
		private DevExpress.XtraEditors.SimpleButton btEditar;
		private DevExpress.XtraEditors.SimpleButton btBuscar;
		private DevExpress.XtraEditors.SimpleButton btGrabar;
		private DevExpress.XtraEditors.SimpleButton btCancelar;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource2;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource4;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource5;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource6;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource7;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource8;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource9;
		private DevExpress.XtraEditors.SimpleButton btFinalizar;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtidArticulo;
		private C1.Data.C1DataSet cdsSeteoF;
		private System.Windows.Forms.ImageList imageList1;
		private System.Windows.Forms.GroupBox groupBox1;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkGrupos;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkSubGrupos;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkMarcas;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource10;
		private Infragistics.Win.UltraWinGrid.UltraGrid uGridModelos;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource11;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkSumaDescuentos;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkAdicional;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource12;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource13;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkOpcional;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource14;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkCuotas;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkCupon;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource15;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource3;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource16;
		private Infragistics.Win.UltraWinGrid.UltraGrid uGridPlanes;
		private Infragistics.Win.UltraWinGrid.UltraGrid uGridPlazos;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource17;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtidHoraPromocion;
		private System.Windows.Forms.Label label4;
		private System.Windows.Forms.Label lblnumproducto;
		private DevExpress.XtraEditors.RadioGroup rdgProductoFacturable;
		private DevExpress.XtraEditors.RadioGroup rdgProductoExistencia;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtBuscar;
		private System.Windows.Forms.Label lblContador;
		private Infragistics.Win.UltraWinGrid.UltraGrid uGridCuotas;
				
		decimal dPorcTarjeta = 0.00m;
		decimal dPorcCredito = 0.00m;
		private System.Windows.Forms.Label lbltotalpagcatalogo;
		private System.Windows.Forms.Label label5;
		private System.Windows.Forms.TextBox txtestadocatalogo;
		private DevExpress.XtraEditors.SimpleButton btcrearcatalogo;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtpaginastotalcatalogo;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtidCatalogo;
		private Infragistics.Win.UltraWinEditors.UltraComboEditor cmbTipo;
		private DevExpress.XtraEditors.SimpleButton btnuevocatalogo;
		private Infragistics.Win.UltraWinGrid.UltraGrid uGridResumen;
		private Infragistics.Win.UltraWinGrid.UltraGrid uGridDetalle;
		private System.Windows.Forms.GroupBox groupBox2;
		public DevExpress.XtraEditors.SimpleButton btExportar;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtpaginas;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtnumproductos;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtidDetCatalogo;
		bool bActualiza = true;
		public frmCatalogo1()
		{
			//
			// Required for Windows Form Designer support
			//
			InitializeComponent();

			//
			// TODO: Add any constructor code after InitializeComponent call
			//
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
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmCatalogo1));
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton2 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idPromocion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idArticulo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn3 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Codigo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn4 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Articulo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn5 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Tipo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn6 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Principal");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn7 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Existencia");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn8 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("CostoUltimo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn9 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Precio3");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn10 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("IncrPVP");
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn11 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idPromocionBancoTarjeta");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn12 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idPromocionBanco");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn13 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idTarjeta");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn14 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn15 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Sel");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn16 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idArticulo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn17 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Articulo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn18 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Codigo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn19 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Cantidad");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn20 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("CostoUltimo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn21 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("UtilidadEfectivo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn22 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("UtilidadTC");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn23 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Precio3");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn24 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("cuo6");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn25 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("cuo9");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn26 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("cuo12");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn27 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("cuo24");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn28 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("cuc3");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn29 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("cuc6");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn30 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("cuc9");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn31 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("cuc12");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn32 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("cuc18");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn33 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nuevopvp");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn34 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idPromoPrecio");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn35 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idTipoPrecio");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn36 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idPromocion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn37 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Precio");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn38 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Operador");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn39 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("PorcentajeInicial");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn40 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Porcentaje");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn41 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Sel");
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPromocionMarca");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idMarca");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Marca");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Sel");
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance11 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance12 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance13 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn42 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idPromocionMarca");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn43 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idMarca");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn44 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Marca");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn45 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Sel");
			Infragistics.Win.Appearance appearance14 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPromocionSubGrupo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idSubGrupo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("SubGrupo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Sel");
			Infragistics.Win.Appearance appearance15 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance16 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance17 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance18 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance19 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance20 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance21 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn46 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idPromocionSubGrupo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn47 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idSubGrupo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn48 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("SubGrupo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn49 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Sel");
			Infragistics.Win.Appearance appearance22 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand3 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPromocionGrupo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPromocion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn11 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idGrupoArticulo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn12 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Grupo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn13 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Sel");
			Infragistics.Win.Appearance appearance23 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance24 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance25 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance26 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance27 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance28 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance29 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn50 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idPromocionGrupo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn51 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idPromocion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn52 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idGrupoArticulo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn53 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Grupo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn54 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Sel");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn55 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idPromocionBanco");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn56 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idPromocion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn57 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idBanco");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn58 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn59 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Sel");
			Infragistics.Win.Appearance appearance30 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance31 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance32 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn60 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idPromocionesPlanes");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn61 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idPromocionBanco");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn62 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idPromocionBancoTarjeta");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn63 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idPlan");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn64 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idPromocion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn65 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Descripcion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn66 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Sel");
			Infragistics.Win.Appearance appearance33 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand4 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn14 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPromocionModelo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn15 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idArticulo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn16 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Modelo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn17 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Sel");
			Infragistics.Win.Appearance appearance34 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance35 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance36 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance37 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance38 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance39 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance40 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn67 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idPromocionModelo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn68 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idArticulo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn69 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Modelo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn70 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Sel");
			Infragistics.Win.Appearance appearance41 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance42 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn71 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idPromocionesPlazos");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn72 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idPlazo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn73 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idPromocion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn74 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Descripcion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn75 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Sel");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn76 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("MesesDeGracia");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn77 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idPromocionesCuotas");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn78 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idPromocion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn79 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Numero");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn80 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Sel");
			Infragistics.Win.Appearance appearance43 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance44 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand5 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn18 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPromocionesCuotas");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn19 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPromocion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn20 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cuota");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn21 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Sel");
			Infragistics.Win.Appearance appearance45 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance46 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance47 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance48 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance49 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance50 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance51 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn81 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idPromocionesCuotas");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn82 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idPromocion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn83 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Cuota");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn84 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Sel");
			Infragistics.Win.Appearance appearance52 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance53 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn85 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idHabitacionPromocion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn86 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idPromocion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn87 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idPromCiudadBodega");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn88 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idHabitacion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn89 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("NumeroHabitacion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn90 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn91 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Tematica");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn92 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Sel");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn93 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idBodega");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn94 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn(" Bodega");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn95 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Codigo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn96 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn97 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idPromCiudadBodega");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn98 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idPromocionCiudad");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn99 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Bodega");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn100 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn101 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Sel");
			Infragistics.Win.Appearance appearance54 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand6 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn22 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPromocionesPlanes");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn23 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPromocionBanco");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn24 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPromocionBancoTarjeta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn25 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPlan");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn26 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPromocion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn27 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Descripcion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn28 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Sel");
			Infragistics.Win.Appearance appearance55 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance56 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance57 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance58 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance59 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance60 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance61 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance62 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand7 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn29 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPromocionesPlazos");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn30 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPlazo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn31 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPromocion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn32 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Descripcion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn33 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Sel");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn34 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("MesesDeGracia");
			Infragistics.Win.Appearance appearance63 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance64 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance65 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance66 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance67 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance68 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance69 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance70 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn102 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idDiaPromocion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn103 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idHabitacionPromocion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn104 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idDia");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn105 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn106 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Activa");
			Infragistics.Win.Appearance appearance71 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance72 = new Infragistics.Win.Appearance();
			Infragistics.Win.ValueListItem valueListItem1 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem2 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.Appearance appearance73 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand8 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn35 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPromocion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn36 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idArticulo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn37 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn38 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Articulo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn39 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Tipo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn40 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Principal");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn41 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Existencia");
			Infragistics.Win.Appearance appearance74 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn42 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CostoUltimo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn43 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Precio3");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn44 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("IncrPVP");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn45 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("UnboundColumn1", 0);
			Infragistics.Win.Appearance appearance75 = new Infragistics.Win.Appearance("btSeleccionar");
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings1 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Existencia", 6, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Existencia", 6, true);
			Infragistics.Win.Appearance appearance76 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance77 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance78 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance79 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance80 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance81 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance82 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance83 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand9 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn46 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idArticulo");
			Infragistics.Win.Appearance appearance84 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn47 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Articulo");
			Infragistics.Win.Appearance appearance85 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn48 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.Appearance appearance86 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn49 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cantidad");
			Infragistics.Win.Appearance appearance87 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn50 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CostoUltimo");
			Infragistics.Win.Appearance appearance88 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn51 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("UtilidadEfectivo");
			Infragistics.Win.Appearance appearance89 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn52 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("UtilidadTC");
			Infragistics.Win.Appearance appearance90 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn53 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Precio3");
			Infragistics.Win.Appearance appearance91 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance92 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn54 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("cuo6");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn55 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("cuo9");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn56 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("cuo12");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn57 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("cuo24");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn58 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("cuc3");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn59 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("cuc6");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn60 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("cuc9");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn61 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("cuc12");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn62 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("cuc18");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn63 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nuevopvp");
			Infragistics.Win.Appearance appearance93 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance94 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance95 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance96 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance97 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance98 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance99 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance100 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance101 = new Infragistics.Win.Appearance();
			this.label1 = new System.Windows.Forms.Label();
			this.dtFin = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.label6 = new System.Windows.Forms.Label();
			this.dtInicio = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.lblCliente = new System.Windows.Forms.Label();
			this.txtDescripcion = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtIdPromocion = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.ultraDataSource1 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.label15 = new System.Windows.Forms.Label();
			this.txtValorMaximo = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label3 = new System.Windows.Forms.Label();
			this.txtValorMinimo = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label2 = new System.Windows.Forms.Label();
			this.label16 = new System.Windows.Forms.Label();
			this.ultraDataSource6 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.ultraDataSource2 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.ultraDataSource4 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.uGridMarcas = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource9 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.uGridSubgrupos = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource8 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.uGridGrupos = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource7 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
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
			this.txtidArticulo = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.groupBox1 = new System.Windows.Forms.GroupBox();
			this.chkGrupos = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.chkSubGrupos = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.chkMarcas = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.ultraDataSource10 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.uGridModelos = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource11 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.chkSumaDescuentos = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.chkAdicional = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.ultraDataSource12 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.ultraDataSource13 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.chkOpcional = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.uGridCuotas = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource14 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.chkCuotas = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.chkCupon = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.ultraDataSource15 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.ultraDataSource3 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.ultraDataSource16 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.uGridPlanes = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.uGridPlazos = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource17 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.txtidHoraPromocion = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label4 = new System.Windows.Forms.Label();
			this.lblnumproducto = new System.Windows.Forms.Label();
			this.rdgProductoFacturable = new DevExpress.XtraEditors.RadioGroup();
			this.rdgProductoExistencia = new DevExpress.XtraEditors.RadioGroup();
			this.txtBuscar = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.lblContador = new System.Windows.Forms.Label();
			this.lbltotalpagcatalogo = new System.Windows.Forms.Label();
			this.label5 = new System.Windows.Forms.Label();
			this.txtestadocatalogo = new System.Windows.Forms.TextBox();
			this.btcrearcatalogo = new DevExpress.XtraEditors.SimpleButton();
			this.txtidCatalogo = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtpaginastotalcatalogo = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.cmbTipo = new Infragistics.Win.UltraWinEditors.UltraComboEditor();
			this.btnuevocatalogo = new DevExpress.XtraEditors.SimpleButton();
			this.uGridResumen = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.uGridDetalle = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.groupBox2 = new System.Windows.Forms.GroupBox();
			this.btExportar = new DevExpress.XtraEditors.SimpleButton();
			this.txtpaginas = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtnumproductos = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtidDetCatalogo = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			((System.ComponentModel.ISupportInitialize)(this.dtFin)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtInicio)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDescripcion)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdPromocion)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtValorMaximo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtValorMinimo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource6)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource4)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridMarcas)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource9)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridSubgrupos)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource8)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridGrupos)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource7)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource5)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsPromocion)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtidArticulo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource10)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridModelos)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource11)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource12)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource13)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridCuotas)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource14)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource15)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource3)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource16)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridPlanes)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridPlazos)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource17)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtidHoraPromocion)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.rdgProductoFacturable.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.rdgProductoExistencia.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtBuscar)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtidCatalogo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtpaginastotalcatalogo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbTipo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridResumen)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridDetalle)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtpaginas)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtnumproductos)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtidDetCatalogo)).BeginInit();
			this.SuspendLayout();
			// 
			// label1
			// 
			this.label1.AutoSize = true;
			this.label1.Enabled = false;
			this.label1.Location = new System.Drawing.Point(464, 8);
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
			this.dtFin.Location = new System.Drawing.Point(504, 8);
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
			this.label6.Location = new System.Drawing.Point(232, 8);
			this.label6.Name = "label6";
			this.label6.Size = new System.Drawing.Size(80, 16);
			this.label6.TabIndex = 199;
			this.label6.Text = "Vigencia Desde";
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
			this.dtInicio.Location = new System.Drawing.Point(312, 8);
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
			this.lblCliente.Location = new System.Drawing.Point(656, 8);
			this.lblCliente.Name = "lblCliente";
			this.lblCliente.Size = new System.Drawing.Size(91, 16);
			this.lblCliente.TabIndex = 205;
			this.lblCliente.Text = "Nombre Catálogo";
			this.lblCliente.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtDescripcion
			// 
			appearance3.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtDescripcion.Appearance = appearance3;
			this.txtDescripcion.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtDescripcion.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtDescripcion.Enabled = false;
			this.txtDescripcion.Location = new System.Drawing.Point(752, 8);
			this.txtDescripcion.Name = "txtDescripcion";
			this.txtDescripcion.Size = new System.Drawing.Size(184, 22);
			this.txtDescripcion.TabIndex = 3;
			this.txtDescripcion.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtDescripcion_KeyPress);
			this.txtDescripcion.ValueChanged += new System.EventHandler(this.txtDescripcion_ValueChanged);
			// 
			// txtIdPromocion
			// 
			this.txtIdPromocion.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.txtIdPromocion.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtIdPromocion.Location = new System.Drawing.Point(1048, 8);
			this.txtIdPromocion.Name = "txtIdPromocion";
			this.txtIdPromocion.PromptChar = ' ';
			this.txtIdPromocion.Size = new System.Drawing.Size(24, 22);
			this.txtIdPromocion.TabIndex = 234;
			this.txtIdPromocion.Visible = false;
			// 
			// ultraDataSource1
			// 
			ultraDataColumn1.DataType = typeof(int);
			ultraDataColumn2.DataType = typeof(int);
			ultraDataColumn5.DataType = typeof(int);
			ultraDataColumn6.DataType = typeof(bool);
			ultraDataColumn7.DataType = typeof(int);
			ultraDataColumn8.DataType = typeof(System.Double);
			ultraDataColumn9.DataType = typeof(System.Decimal);
			ultraDataColumn10.DataType = typeof(System.Decimal);
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
																																 ultraDataColumn10});
			// 
			// label15
			// 
			this.label15.AutoSize = true;
			this.label15.BackColor = System.Drawing.Color.Transparent;
			this.label15.Enabled = false;
			this.label15.Location = new System.Drawing.Point(1184, 408);
			this.label15.Name = "label15";
			this.label15.Size = new System.Drawing.Size(38, 16);
			this.label15.TabIndex = 253;
			this.label15.Text = "Cuotas";
			this.label15.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.label15.Visible = false;
			// 
			// txtValorMaximo
			// 
			appearance4.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtValorMaximo.Appearance = appearance4;
			this.txtValorMaximo.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.txtValorMaximo.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtValorMaximo.Enabled = false;
			this.txtValorMaximo.FormatString = "#,##0.00";
			this.txtValorMaximo.Location = new System.Drawing.Point(1272, 456);
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
			this.label3.Location = new System.Drawing.Point(1184, 456);
			this.label3.Name = "label3";
			this.label3.Size = new System.Drawing.Size(77, 16);
			this.label3.TabIndex = 271;
			this.label3.Text = "Precio máximo";
			this.label3.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.label3.Visible = false;
			// 
			// txtValorMinimo
			// 
			appearance5.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtValorMinimo.Appearance = appearance5;
			this.txtValorMinimo.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.txtValorMinimo.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtValorMinimo.Enabled = false;
			this.txtValorMinimo.FormatString = "#,##0.00";
			this.txtValorMinimo.Location = new System.Drawing.Point(1272, 432);
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
			this.label2.Location = new System.Drawing.Point(1184, 432);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(75, 16);
			this.label2.TabIndex = 270;
			this.label2.Text = "Precio mínimo";
			this.label2.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.label2.Visible = false;
			// 
			// label16
			// 
			this.label16.AutoSize = true;
			this.label16.Enabled = false;
			this.label16.Location = new System.Drawing.Point(0, 10);
			this.label16.Name = "label16";
			this.label16.Size = new System.Drawing.Size(89, 16);
			this.label16.TabIndex = 278;
			this.label16.Text = "Tipo de Catalogo";
			this.label16.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// ultraDataSource6
			// 
			ultraDataColumn11.DataType = typeof(int);
			ultraDataColumn12.DataType = typeof(int);
			ultraDataColumn13.DataType = typeof(int);
			ultraDataColumn15.DataType = typeof(bool);
			ultraDataColumn15.DefaultValue = false;
			this.ultraDataSource6.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn11,
																																 ultraDataColumn12,
																																 ultraDataColumn13,
																																 ultraDataColumn14,
																																 ultraDataColumn15});
			// 
			// ultraDataSource2
			// 
			ultraDataColumn16.DataType = typeof(int);
			ultraDataColumn19.DataType = typeof(int);
			ultraDataColumn20.DataType = typeof(System.Decimal);
			ultraDataColumn21.DataType = typeof(System.Decimal);
			ultraDataColumn22.DataType = typeof(System.Decimal);
			ultraDataColumn23.DataType = typeof(System.Decimal);
			ultraDataColumn24.DataType = typeof(System.Decimal);
			ultraDataColumn25.DataType = typeof(System.Decimal);
			ultraDataColumn26.DataType = typeof(System.Decimal);
			ultraDataColumn27.DataType = typeof(System.Decimal);
			ultraDataColumn28.DataType = typeof(System.Decimal);
			ultraDataColumn29.DataType = typeof(System.Decimal);
			ultraDataColumn30.DataType = typeof(System.Decimal);
			ultraDataColumn31.DataType = typeof(System.Decimal);
			ultraDataColumn32.DataType = typeof(System.Decimal);
			ultraDataColumn33.DataType = typeof(System.Decimal);
			this.ultraDataSource2.Band.Columns.AddRange(new object[] {
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
																																 ultraDataColumn33});
			// 
			// ultraDataSource4
			// 
			ultraDataColumn34.DataType = typeof(int);
			ultraDataColumn35.DataType = typeof(int);
			ultraDataColumn36.DataType = typeof(int);
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
			ultraDataColumn41.DataType = typeof(bool);
			ultraDataColumn41.DefaultValue = false;
			this.ultraDataSource4.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn34,
																																 ultraDataColumn35,
																																 ultraDataColumn36,
																																 ultraDataColumn37,
																																 ultraDataColumn38,
																																 ultraDataColumn39,
																																 ultraDataColumn40,
																																 ultraDataColumn41});
			// 
			// uGridMarcas
			// 
			this.uGridMarcas.Cursor = System.Windows.Forms.Cursors.Default;
			this.uGridMarcas.DataSource = this.ultraDataSource9;
			appearance6.BackColor = System.Drawing.Color.White;
			this.uGridMarcas.DisplayLayout.Appearance = appearance6;
			this.uGridMarcas.DisplayLayout.AutoFitColumns = true;
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridColumn2.Hidden = true;
			ultraGridColumn3.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn3.Header.VisiblePosition = 2;
			ultraGridColumn3.Width = 181;
			ultraGridColumn4.Header.Caption = "...";
			ultraGridColumn4.Header.VisiblePosition = 3;
			ultraGridColumn4.Width = 30;
			ultraGridBand1.Columns.AddRange(new object[] {
																										 ultraGridColumn1,
																										 ultraGridColumn2,
																										 ultraGridColumn3,
																										 ultraGridColumn4});
			ultraGridBand1.Header.Caption = "MARCAS DE ARTÍCULOS";
			ultraGridBand1.HeaderVisible = true;
			ultraGridBand1.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			appearance7.FontData.SizeInPoints = 7.25F;
			ultraGridBand1.Override.HeaderAppearance = appearance7;
			this.uGridMarcas.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			appearance8.ForeColor = System.Drawing.Color.Black;
			appearance8.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridMarcas.DisplayLayout.Override.ActiveRowAppearance = appearance8;
			this.uGridMarcas.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.uGridMarcas.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance9.BackColor = System.Drawing.Color.Transparent;
			this.uGridMarcas.DisplayLayout.Override.CardAreaAppearance = appearance9;
			appearance10.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance10.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance10.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance10.FontData.BoldAsString = "True";
			appearance10.FontData.Name = "Arial";
			appearance10.FontData.SizeInPoints = 10F;
			appearance10.ForeColor = System.Drawing.Color.White;
			appearance10.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridMarcas.DisplayLayout.Override.HeaderAppearance = appearance10;
			appearance11.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance11.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance11.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridMarcas.DisplayLayout.Override.RowAlternateAppearance = appearance11;
			appearance12.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance12.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance12.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridMarcas.DisplayLayout.Override.RowSelectorAppearance = appearance12;
			appearance13.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance13.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance13.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridMarcas.DisplayLayout.Override.SelectedRowAppearance = appearance13;
			this.uGridMarcas.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridMarcas.Location = new System.Drawing.Point(472, 112);
			this.uGridMarcas.Name = "uGridMarcas";
			this.uGridMarcas.Size = new System.Drawing.Size(232, 168);
			this.uGridMarcas.TabIndex = 286;
			this.uGridMarcas.KeyDown += new System.Windows.Forms.KeyEventHandler(this.uGridMarcas_KeyDown);
			this.uGridMarcas.AfterCellUpdate += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.uGridMarcas_AfterCellUpdate);
			this.uGridMarcas.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.uGridMarcas_InitializeLayout);
			this.uGridMarcas.Click += new System.EventHandler(this.uGridMarcas_Click);
			// 
			// ultraDataSource9
			// 
			ultraDataColumn42.DataType = typeof(int);
			ultraDataColumn43.DataType = typeof(int);
			ultraDataColumn45.DataType = typeof(bool);
			ultraDataColumn45.DefaultValue = false;
			this.ultraDataSource9.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn42,
																																 ultraDataColumn43,
																																 ultraDataColumn44,
																																 ultraDataColumn45});
			// 
			// uGridSubgrupos
			// 
			this.uGridSubgrupos.Cursor = System.Windows.Forms.Cursors.Default;
			this.uGridSubgrupos.DataSource = this.ultraDataSource8;
			appearance14.BackColor = System.Drawing.Color.White;
			this.uGridSubgrupos.DisplayLayout.Appearance = appearance14;
			this.uGridSubgrupos.DisplayLayout.AutoFitColumns = true;
			ultraGridColumn5.Header.VisiblePosition = 0;
			ultraGridColumn5.Hidden = true;
			ultraGridColumn6.Header.VisiblePosition = 1;
			ultraGridColumn6.Hidden = true;
			ultraGridColumn7.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn7.Header.VisiblePosition = 2;
			ultraGridColumn7.Width = 181;
			ultraGridColumn8.Header.Caption = "...";
			ultraGridColumn8.Header.VisiblePosition = 3;
			ultraGridColumn8.Width = 30;
			ultraGridBand2.Columns.AddRange(new object[] {
																										 ultraGridColumn5,
																										 ultraGridColumn6,
																										 ultraGridColumn7,
																										 ultraGridColumn8});
			ultraGridBand2.Header.Caption = "SUBGRUPOS DE ARTÍCULOS";
			ultraGridBand2.HeaderVisible = true;
			ultraGridBand2.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			appearance15.FontData.SizeInPoints = 7.25F;
			ultraGridBand2.Override.HeaderAppearance = appearance15;
			this.uGridSubgrupos.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			appearance16.ForeColor = System.Drawing.Color.Black;
			appearance16.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridSubgrupos.DisplayLayout.Override.ActiveRowAppearance = appearance16;
			this.uGridSubgrupos.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.uGridSubgrupos.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance17.BackColor = System.Drawing.Color.Transparent;
			this.uGridSubgrupos.DisplayLayout.Override.CardAreaAppearance = appearance17;
			appearance18.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance18.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance18.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance18.FontData.BoldAsString = "True";
			appearance18.FontData.Name = "Arial";
			appearance18.FontData.SizeInPoints = 10F;
			appearance18.ForeColor = System.Drawing.Color.White;
			appearance18.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridSubgrupos.DisplayLayout.Override.HeaderAppearance = appearance18;
			appearance19.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance19.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance19.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridSubgrupos.DisplayLayout.Override.RowAlternateAppearance = appearance19;
			appearance20.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance20.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance20.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridSubgrupos.DisplayLayout.Override.RowSelectorAppearance = appearance20;
			appearance21.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance21.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance21.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridSubgrupos.DisplayLayout.Override.SelectedRowAppearance = appearance21;
			this.uGridSubgrupos.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridSubgrupos.Location = new System.Drawing.Point(240, 112);
			this.uGridSubgrupos.Name = "uGridSubgrupos";
			this.uGridSubgrupos.Size = new System.Drawing.Size(232, 168);
			this.uGridSubgrupos.TabIndex = 285;
			this.uGridSubgrupos.KeyDown += new System.Windows.Forms.KeyEventHandler(this.uGridSubgrupos_KeyDown);
			this.uGridSubgrupos.AfterCellUpdate += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.uGridSubgrupos_AfterCellUpdate);
			this.uGridSubgrupos.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.uGridSubgrupos_InitializeLayout);
			this.uGridSubgrupos.Click += new System.EventHandler(this.uGridSubgrupos_Click);
			// 
			// ultraDataSource8
			// 
			ultraDataColumn46.DataType = typeof(int);
			ultraDataColumn47.DataType = typeof(int);
			ultraDataColumn49.DataType = typeof(bool);
			ultraDataColumn49.DefaultValue = false;
			this.ultraDataSource8.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn46,
																																 ultraDataColumn47,
																																 ultraDataColumn48,
																																 ultraDataColumn49});
			// 
			// uGridGrupos
			// 
			this.uGridGrupos.Cursor = System.Windows.Forms.Cursors.Default;
			this.uGridGrupos.DataSource = this.ultraDataSource7;
			appearance22.BackColor = System.Drawing.Color.White;
			this.uGridGrupos.DisplayLayout.Appearance = appearance22;
			this.uGridGrupos.DisplayLayout.AutoFitColumns = true;
			ultraGridColumn9.Header.VisiblePosition = 0;
			ultraGridColumn9.Hidden = true;
			ultraGridColumn9.Width = 72;
			ultraGridColumn10.Header.VisiblePosition = 1;
			ultraGridColumn10.Hidden = true;
			ultraGridColumn10.Width = 20;
			ultraGridColumn11.Header.VisiblePosition = 2;
			ultraGridColumn11.Hidden = true;
			ultraGridColumn11.Width = 65;
			ultraGridColumn12.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn12.Header.VisiblePosition = 3;
			ultraGridColumn12.Width = 181;
			ultraGridColumn13.Header.Caption = "...";
			ultraGridColumn13.Header.VisiblePosition = 4;
			ultraGridColumn13.Width = 30;
			ultraGridBand3.Columns.AddRange(new object[] {
																										 ultraGridColumn9,
																										 ultraGridColumn10,
																										 ultraGridColumn11,
																										 ultraGridColumn12,
																										 ultraGridColumn13});
			ultraGridBand3.Header.Caption = "GRUPOS DE ARTÍCULOS";
			ultraGridBand3.HeaderVisible = true;
			ultraGridBand3.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			appearance23.FontData.SizeInPoints = 7.25F;
			ultraGridBand3.Override.HeaderAppearance = appearance23;
			this.uGridGrupos.DisplayLayout.BandsSerializer.Add(ultraGridBand3);
			appearance24.ForeColor = System.Drawing.Color.Black;
			appearance24.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridGrupos.DisplayLayout.Override.ActiveRowAppearance = appearance24;
			this.uGridGrupos.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.uGridGrupos.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance25.BackColor = System.Drawing.Color.Transparent;
			this.uGridGrupos.DisplayLayout.Override.CardAreaAppearance = appearance25;
			appearance26.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance26.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance26.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance26.FontData.BoldAsString = "True";
			appearance26.FontData.Name = "Arial";
			appearance26.FontData.SizeInPoints = 10F;
			appearance26.ForeColor = System.Drawing.Color.White;
			appearance26.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridGrupos.DisplayLayout.Override.HeaderAppearance = appearance26;
			appearance27.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance27.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance27.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridGrupos.DisplayLayout.Override.RowAlternateAppearance = appearance27;
			appearance28.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance28.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance28.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridGrupos.DisplayLayout.Override.RowSelectorAppearance = appearance28;
			appearance29.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance29.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance29.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridGrupos.DisplayLayout.Override.SelectedRowAppearance = appearance29;
			this.uGridGrupos.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridGrupos.Location = new System.Drawing.Point(0, 112);
			this.uGridGrupos.Name = "uGridGrupos";
			this.uGridGrupos.Size = new System.Drawing.Size(232, 168);
			this.uGridGrupos.TabIndex = 284;
			this.uGridGrupos.KeyDown += new System.Windows.Forms.KeyEventHandler(this.uGridGrupos_KeyDown);
			this.uGridGrupos.AfterCellUpdate += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.uGridGrupos_AfterCellUpdate);
			this.uGridGrupos.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.uGridGrupos_InitializeLayout);
			this.uGridGrupos.Click += new System.EventHandler(this.uGridGrupos_Click);
			// 
			// ultraDataSource7
			// 
			ultraDataColumn50.DataType = typeof(int);
			ultraDataColumn51.DataType = typeof(int);
			ultraDataColumn52.DataType = typeof(int);
			ultraDataColumn54.DataType = typeof(bool);
			ultraDataColumn54.DefaultValue = false;
			this.ultraDataSource7.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn50,
																																 ultraDataColumn51,
																																 ultraDataColumn52,
																																 ultraDataColumn53,
																																 ultraDataColumn54});
			// 
			// ultraDataSource5
			// 
			ultraDataColumn55.DataType = typeof(int);
			ultraDataColumn56.DataType = typeof(int);
			ultraDataColumn57.DataType = typeof(int);
			ultraDataColumn59.DataType = typeof(bool);
			ultraDataColumn59.DefaultValue = false;
			this.ultraDataSource5.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn55,
																																 ultraDataColumn56,
																																 ultraDataColumn57,
																																 ultraDataColumn58,
																																 ultraDataColumn59});
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
			this.btNuevo.Location = new System.Drawing.Point(576, 80);
			this.btNuevo.Name = "btNuevo";
			this.btNuevo.Size = new System.Drawing.Size(112, 24);
			this.btNuevo.TabIndex = 290;
			this.btNuevo.Text = "&Agregar Pagina";
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
			this.btSalir.Location = new System.Drawing.Point(984, 400);
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
			this.btEditar.Location = new System.Drawing.Point(984, 328);
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
			this.btBuscar.Location = new System.Drawing.Point(984, 296);
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
			this.btGrabar.Location = new System.Drawing.Point(696, 80);
			this.btGrabar.Name = "btGrabar";
			this.btGrabar.Size = new System.Drawing.Size(96, 24);
			this.btGrabar.TabIndex = 291;
			this.btGrabar.Text = "&Grabar Página";
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
			this.btCancelar.Location = new System.Drawing.Point(984, 376);
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
			this.btFinalizar.Location = new System.Drawing.Point(984, 352);
			this.btFinalizar.Name = "btFinalizar";
			this.btFinalizar.Size = new System.Drawing.Size(96, 24);
			this.btFinalizar.TabIndex = 297;
			this.btFinalizar.Text = "&Finalizar";
			// 
			// txtidArticulo
			// 
			this.txtidArticulo.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.txtidArticulo.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtidArticulo.Location = new System.Drawing.Point(1080, 8);
			this.txtidArticulo.Name = "txtidArticulo";
			this.txtidArticulo.PromptChar = ' ';
			this.txtidArticulo.Size = new System.Drawing.Size(16, 22);
			this.txtidArticulo.TabIndex = 299;
			this.txtidArticulo.Visible = false;
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
			// groupBox1
			// 
			this.groupBox1.Location = new System.Drawing.Point(0, 56);
			this.groupBox1.Name = "groupBox1";
			this.groupBox1.Size = new System.Drawing.Size(1167, 8);
			this.groupBox1.TabIndex = 303;
			this.groupBox1.TabStop = false;
			this.groupBox1.Enter += new System.EventHandler(this.groupBox1_Enter);
			// 
			// chkGrupos
			// 
			appearance30.ForeColorDisabled = System.Drawing.Color.Black;
			this.chkGrupos.Appearance = appearance30;
			this.chkGrupos.BackColor = System.Drawing.Color.Transparent;
			this.chkGrupos.CausesValidation = false;
			this.chkGrupos.CheckAlign = System.Drawing.ContentAlignment.MiddleCenter;
			this.chkGrupos.Enabled = false;
			this.chkGrupos.Location = new System.Drawing.Point(216, 248);
			this.chkGrupos.Name = "chkGrupos";
			this.chkGrupos.Size = new System.Drawing.Size(13, 13);
			this.chkGrupos.TabIndex = 305;
			this.chkGrupos.Text = "Precio Manual";
			this.chkGrupos.Visible = false;
			this.chkGrupos.CheckedChanged += new System.EventHandler(this.chkGrupos_CheckedChanged);
			// 
			// chkSubGrupos
			// 
			appearance31.ForeColorDisabled = System.Drawing.Color.Black;
			this.chkSubGrupos.Appearance = appearance31;
			this.chkSubGrupos.BackColor = System.Drawing.Color.Transparent;
			this.chkSubGrupos.CausesValidation = false;
			this.chkSubGrupos.CheckAlign = System.Drawing.ContentAlignment.MiddleCenter;
			this.chkSubGrupos.Enabled = false;
			this.chkSubGrupos.Location = new System.Drawing.Point(464, 240);
			this.chkSubGrupos.Name = "chkSubGrupos";
			this.chkSubGrupos.Size = new System.Drawing.Size(13, 13);
			this.chkSubGrupos.TabIndex = 306;
			this.chkSubGrupos.Text = "Precio Manual";
			this.chkSubGrupos.Visible = false;
			this.chkSubGrupos.CheckedChanged += new System.EventHandler(this.chkSubGrupos_CheckedChanged);
			// 
			// chkMarcas
			// 
			appearance32.ForeColorDisabled = System.Drawing.Color.Black;
			this.chkMarcas.Appearance = appearance32;
			this.chkMarcas.BackColor = System.Drawing.Color.Transparent;
			this.chkMarcas.CausesValidation = false;
			this.chkMarcas.CheckAlign = System.Drawing.ContentAlignment.MiddleCenter;
			this.chkMarcas.Enabled = false;
			this.chkMarcas.Location = new System.Drawing.Point(704, 240);
			this.chkMarcas.Name = "chkMarcas";
			this.chkMarcas.Size = new System.Drawing.Size(13, 13);
			this.chkMarcas.TabIndex = 307;
			this.chkMarcas.Text = "Precio Manual";
			this.chkMarcas.Visible = false;
			this.chkMarcas.CheckedChanged += new System.EventHandler(this.chkMarcas_CheckedChanged);
			// 
			// ultraDataSource10
			// 
			ultraDataColumn60.DataType = typeof(int);
			ultraDataColumn61.DataType = typeof(int);
			ultraDataColumn62.DataType = typeof(int);
			ultraDataColumn63.DataType = typeof(int);
			ultraDataColumn64.DataType = typeof(int);
			ultraDataColumn66.DataType = typeof(bool);
			ultraDataColumn66.DefaultValue = false;
			this.ultraDataSource10.Band.Columns.AddRange(new object[] {
																																	ultraDataColumn60,
																																	ultraDataColumn61,
																																	ultraDataColumn62,
																																	ultraDataColumn63,
																																	ultraDataColumn64,
																																	ultraDataColumn65,
																																	ultraDataColumn66});
			// 
			// uGridModelos
			// 
			this.uGridModelos.Cursor = System.Windows.Forms.Cursors.Default;
			this.uGridModelos.DataSource = this.ultraDataSource11;
			appearance33.BackColor = System.Drawing.Color.White;
			this.uGridModelos.DisplayLayout.Appearance = appearance33;
			this.uGridModelos.DisplayLayout.AutoFitColumns = true;
			ultraGridColumn14.Header.VisiblePosition = 0;
			ultraGridColumn14.Hidden = true;
			ultraGridColumn14.Width = 103;
			ultraGridColumn15.Header.VisiblePosition = 1;
			ultraGridColumn15.Hidden = true;
			ultraGridColumn15.Width = 55;
			ultraGridColumn16.Header.VisiblePosition = 2;
			ultraGridColumn16.Width = 188;
			ultraGridColumn17.Header.Caption = "...";
			ultraGridColumn17.Header.VisiblePosition = 3;
			ultraGridColumn17.Width = 31;
			ultraGridBand4.Columns.AddRange(new object[] {
																										 ultraGridColumn14,
																										 ultraGridColumn15,
																										 ultraGridColumn16,
																										 ultraGridColumn17});
			ultraGridBand4.Header.Caption = "MODELOS DE ARTÍCULOS";
			ultraGridBand4.HeaderVisible = true;
			ultraGridBand4.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			appearance34.FontData.SizeInPoints = 7.25F;
			ultraGridBand4.Override.HeaderAppearance = appearance34;
			this.uGridModelos.DisplayLayout.BandsSerializer.Add(ultraGridBand4);
			appearance35.ForeColor = System.Drawing.Color.Black;
			appearance35.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridModelos.DisplayLayout.Override.ActiveRowAppearance = appearance35;
			this.uGridModelos.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.uGridModelos.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance36.BackColor = System.Drawing.Color.Transparent;
			this.uGridModelos.DisplayLayout.Override.CardAreaAppearance = appearance36;
			appearance37.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance37.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance37.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance37.FontData.BoldAsString = "True";
			appearance37.FontData.Name = "Arial";
			appearance37.FontData.SizeInPoints = 10F;
			appearance37.ForeColor = System.Drawing.Color.White;
			appearance37.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridModelos.DisplayLayout.Override.HeaderAppearance = appearance37;
			appearance38.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance38.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance38.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridModelos.DisplayLayout.Override.RowAlternateAppearance = appearance38;
			appearance39.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance39.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance39.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridModelos.DisplayLayout.Override.RowSelectorAppearance = appearance39;
			appearance40.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance40.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance40.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridModelos.DisplayLayout.Override.SelectedRowAppearance = appearance40;
			this.uGridModelos.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridModelos.Location = new System.Drawing.Point(712, 112);
			this.uGridModelos.Name = "uGridModelos";
			this.uGridModelos.Size = new System.Drawing.Size(240, 168);
			this.uGridModelos.TabIndex = 310;
			this.uGridModelos.KeyDown += new System.Windows.Forms.KeyEventHandler(this.uGridModelos_KeyDown);
			this.uGridModelos.AfterCellUpdate += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.uGridModelos_AfterCellUpdate);
			this.uGridModelos.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.uGridModelos_InitializeLayout);
			// 
			// ultraDataSource11
			// 
			ultraDataColumn67.DataType = typeof(int);
			ultraDataColumn68.DataType = typeof(int);
			ultraDataColumn70.DataType = typeof(bool);
			ultraDataColumn70.DefaultValue = false;
			this.ultraDataSource11.Band.Columns.AddRange(new object[] {
																																	ultraDataColumn67,
																																	ultraDataColumn68,
																																	ultraDataColumn69,
																																	ultraDataColumn70});
			// 
			// chkSumaDescuentos
			// 
			appearance41.FontData.Name = "Tahoma";
			appearance41.ForeColorDisabled = System.Drawing.Color.Black;
			this.chkSumaDescuentos.Appearance = appearance41;
			this.chkSumaDescuentos.BackColor = System.Drawing.Color.Transparent;
			this.chkSumaDescuentos.CausesValidation = false;
			this.chkSumaDescuentos.Enabled = false;
			this.chkSumaDescuentos.Location = new System.Drawing.Point(960, 72);
			this.chkSumaDescuentos.Name = "chkSumaDescuentos";
			this.chkSumaDescuentos.Size = new System.Drawing.Size(136, 21);
			this.chkSumaDescuentos.TabIndex = 311;
			this.chkSumaDescuentos.Text = "Descuentos";
			// 
			// chkAdicional
			// 
			appearance42.FontData.Name = "Tahoma";
			appearance42.ForeColorDisabled = System.Drawing.Color.Black;
			this.chkAdicional.Appearance = appearance42;
			this.chkAdicional.BackColor = System.Drawing.Color.Transparent;
			this.chkAdicional.CausesValidation = false;
			this.chkAdicional.Enabled = false;
			this.chkAdicional.Location = new System.Drawing.Point(960, 120);
			this.chkAdicional.Name = "chkAdicional";
			this.chkAdicional.Size = new System.Drawing.Size(136, 21);
			this.chkAdicional.TabIndex = 313;
			this.chkAdicional.Text = "Adicional";
			// 
			// ultraDataSource12
			// 
			ultraDataColumn71.DataType = typeof(int);
			ultraDataColumn71.DefaultValue = 0;
			ultraDataColumn72.DataType = typeof(int);
			ultraDataColumn72.DefaultValue = 0;
			ultraDataColumn73.DataType = typeof(int);
			ultraDataColumn73.DefaultValue = 0;
			ultraDataColumn75.DataType = typeof(bool);
			ultraDataColumn75.DefaultValue = false;
			ultraDataColumn76.DataType = typeof(int);
			this.ultraDataSource12.Band.Columns.AddRange(new object[] {
																																	ultraDataColumn71,
																																	ultraDataColumn72,
																																	ultraDataColumn73,
																																	ultraDataColumn74,
																																	ultraDataColumn75,
																																	ultraDataColumn76});
			// 
			// ultraDataSource13
			// 
			ultraDataColumn77.DataType = typeof(int);
			ultraDataColumn78.DataType = typeof(int);
			ultraDataColumn79.DataType = typeof(int);
			ultraDataColumn80.DataType = typeof(bool);
			ultraDataColumn80.DefaultValue = false;
			this.ultraDataSource13.Band.Columns.AddRange(new object[] {
																																	ultraDataColumn77,
																																	ultraDataColumn78,
																																	ultraDataColumn79,
																																	ultraDataColumn80});
			// 
			// chkOpcional
			// 
			appearance43.FontData.Name = "Tahoma";
			appearance43.ForeColorDisabled = System.Drawing.Color.Black;
			this.chkOpcional.Appearance = appearance43;
			this.chkOpcional.BackColor = System.Drawing.Color.Transparent;
			this.chkOpcional.CausesValidation = false;
			this.chkOpcional.Enabled = false;
			this.chkOpcional.Location = new System.Drawing.Point(960, 144);
			this.chkOpcional.Name = "chkOpcional";
			this.chkOpcional.Size = new System.Drawing.Size(136, 21);
			this.chkOpcional.TabIndex = 315;
			this.chkOpcional.Text = "Obsequios Opcionales";
			// 
			// uGridCuotas
			// 
			this.uGridCuotas.Cursor = System.Windows.Forms.Cursors.Default;
			this.uGridCuotas.DataSource = this.ultraDataSource14;
			appearance44.BackColor = System.Drawing.Color.White;
			this.uGridCuotas.DisplayLayout.Appearance = appearance44;
			this.uGridCuotas.DisplayLayout.AutoFitColumns = true;
			ultraGridColumn18.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn18.Header.VisiblePosition = 0;
			ultraGridColumn18.Hidden = true;
			ultraGridColumn18.Width = 69;
			ultraGridColumn19.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn19.Header.VisiblePosition = 1;
			ultraGridColumn19.Hidden = true;
			ultraGridColumn19.Width = 44;
			ultraGridColumn20.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn20.Header.VisiblePosition = 2;
			ultraGridColumn20.Width = 73;
			ultraGridColumn21.Header.Caption = "...";
			ultraGridColumn21.Header.VisiblePosition = 3;
			ultraGridColumn21.Width = 17;
			ultraGridBand5.Columns.AddRange(new object[] {
																										 ultraGridColumn18,
																										 ultraGridColumn19,
																										 ultraGridColumn20,
																										 ultraGridColumn21});
			ultraGridBand5.Header.Caption = "CUOTAS CREDIPOINT";
			ultraGridBand5.HeaderVisible = true;
			ultraGridBand5.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			appearance45.FontData.SizeInPoints = 7.25F;
			ultraGridBand5.Override.HeaderAppearance = appearance45;
			this.uGridCuotas.DisplayLayout.BandsSerializer.Add(ultraGridBand5);
			appearance46.ForeColor = System.Drawing.Color.Black;
			appearance46.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridCuotas.DisplayLayout.Override.ActiveRowAppearance = appearance46;
			this.uGridCuotas.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.uGridCuotas.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance47.BackColor = System.Drawing.Color.Transparent;
			this.uGridCuotas.DisplayLayout.Override.CardAreaAppearance = appearance47;
			appearance48.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance48.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance48.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance48.FontData.BoldAsString = "True";
			appearance48.FontData.Name = "Arial";
			appearance48.FontData.SizeInPoints = 10F;
			appearance48.ForeColor = System.Drawing.Color.White;
			appearance48.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridCuotas.DisplayLayout.Override.HeaderAppearance = appearance48;
			appearance49.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance49.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance49.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridCuotas.DisplayLayout.Override.RowAlternateAppearance = appearance49;
			appearance50.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance50.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance50.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridCuotas.DisplayLayout.Override.RowSelectorAppearance = appearance50;
			appearance51.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance51.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance51.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridCuotas.DisplayLayout.Override.SelectedRowAppearance = appearance51;
			this.uGridCuotas.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridCuotas.Location = new System.Drawing.Point(1224, 480);
			this.uGridCuotas.Name = "uGridCuotas";
			this.uGridCuotas.Size = new System.Drawing.Size(128, 56);
			this.uGridCuotas.TabIndex = 316;
			this.uGridCuotas.Visible = false;
			this.uGridCuotas.KeyDown += new System.Windows.Forms.KeyEventHandler(this.uGridCuotas_KeyDown);
			this.uGridCuotas.AfterCellUpdate += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.uGridCuotas_AfterCellUpdate);
			// 
			// ultraDataSource14
			// 
			ultraDataColumn81.DataType = typeof(int);
			ultraDataColumn82.DataType = typeof(int);
			ultraDataColumn83.DataType = typeof(int);
			ultraDataColumn84.DataType = typeof(bool);
			ultraDataColumn84.DefaultValue = false;
			this.ultraDataSource14.Band.Columns.AddRange(new object[] {
																																	ultraDataColumn81,
																																	ultraDataColumn82,
																																	ultraDataColumn83,
																																	ultraDataColumn84});
			// 
			// chkCuotas
			// 
			appearance52.ForeColorDisabled = System.Drawing.Color.Black;
			this.chkCuotas.Appearance = appearance52;
			this.chkCuotas.BackColor = System.Drawing.Color.Transparent;
			this.chkCuotas.CausesValidation = false;
			this.chkCuotas.CheckAlign = System.Drawing.ContentAlignment.MiddleCenter;
			this.chkCuotas.Enabled = false;
			this.chkCuotas.Location = new System.Drawing.Point(1320, 368);
			this.chkCuotas.Name = "chkCuotas";
			this.chkCuotas.Size = new System.Drawing.Size(13, 13);
			this.chkCuotas.TabIndex = 317;
			this.chkCuotas.Text = "Precio Manual";
			this.chkCuotas.Visible = false;
			// 
			// chkCupon
			// 
			appearance53.FontData.Name = "Tahoma";
			appearance53.ForeColorDisabled = System.Drawing.Color.Black;
			this.chkCupon.Appearance = appearance53;
			this.chkCupon.BackColor = System.Drawing.Color.Transparent;
			this.chkCupon.CausesValidation = false;
			this.chkCupon.Enabled = false;
			this.chkCupon.Location = new System.Drawing.Point(960, 96);
			this.chkCupon.Name = "chkCupon";
			this.chkCupon.Size = new System.Drawing.Size(136, 21);
			this.chkCupon.TabIndex = 318;
			this.chkCupon.Text = "Cupon";
			// 
			// ultraDataSource15
			// 
			ultraDataColumn85.DataType = typeof(int);
			ultraDataColumn86.DataType = typeof(int);
			ultraDataColumn87.DataType = typeof(int);
			ultraDataColumn88.DataType = typeof(int);
			ultraDataColumn89.DataType = typeof(int);
			ultraDataColumn92.DataType = typeof(bool);
			ultraDataColumn92.DefaultValue = false;
			this.ultraDataSource15.Band.Columns.AddRange(new object[] {
																																	ultraDataColumn85,
																																	ultraDataColumn86,
																																	ultraDataColumn87,
																																	ultraDataColumn88,
																																	ultraDataColumn89,
																																	ultraDataColumn90,
																																	ultraDataColumn91,
																																	ultraDataColumn92});
			// 
			// ultraDataSource3
			// 
			ultraDataColumn93.DataType = typeof(int);
			ultraDataColumn94.DataType = typeof(int);
			this.ultraDataSource3.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn93,
																																 ultraDataColumn94,
																																 ultraDataColumn95,
																																 ultraDataColumn96});
			// 
			// ultraDataSource16
			// 
			ultraDataColumn97.DataType = typeof(int);
			ultraDataColumn98.DataType = typeof(int);
			ultraDataColumn99.DataType = typeof(int);
			ultraDataColumn101.DataType = typeof(bool);
			ultraDataColumn101.DefaultValue = false;
			this.ultraDataSource16.Band.Columns.AddRange(new object[] {
																																	ultraDataColumn97,
																																	ultraDataColumn98,
																																	ultraDataColumn99,
																																	ultraDataColumn100,
																																	ultraDataColumn101});
			// 
			// uGridPlanes
			// 
			this.uGridPlanes.Cursor = System.Windows.Forms.Cursors.Default;
			this.uGridPlanes.DataSource = this.ultraDataSource10;
			appearance54.BackColor = System.Drawing.Color.White;
			this.uGridPlanes.DisplayLayout.Appearance = appearance54;
			this.uGridPlanes.DisplayLayout.AutoFitColumns = true;
			ultraGridColumn22.Header.VisiblePosition = 0;
			ultraGridColumn22.Hidden = true;
			ultraGridColumn22.Width = 64;
			ultraGridColumn23.Header.VisiblePosition = 1;
			ultraGridColumn23.Hidden = true;
			ultraGridColumn23.Width = 68;
			ultraGridColumn24.Header.VisiblePosition = 2;
			ultraGridColumn24.Hidden = true;
			ultraGridColumn24.Width = 48;
			ultraGridColumn25.Header.VisiblePosition = 3;
			ultraGridColumn25.Hidden = true;
			ultraGridColumn25.Width = 30;
			ultraGridColumn26.Header.VisiblePosition = 4;
			ultraGridColumn26.Hidden = true;
			ultraGridColumn26.Width = 41;
			ultraGridColumn27.Header.VisiblePosition = 5;
			ultraGridColumn27.Width = 50;
			ultraGridColumn28.Header.Caption = "...";
			ultraGridColumn28.Header.VisiblePosition = 6;
			ultraGridColumn28.Width = 17;
			ultraGridBand6.Columns.AddRange(new object[] {
																										 ultraGridColumn22,
																										 ultraGridColumn23,
																										 ultraGridColumn24,
																										 ultraGridColumn25,
																										 ultraGridColumn26,
																										 ultraGridColumn27,
																										 ultraGridColumn28});
			ultraGridBand6.Header.Caption = "PLANES";
			ultraGridBand6.HeaderVisible = true;
			ultraGridBand6.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			appearance55.FontData.SizeInPoints = 7.25F;
			ultraGridBand6.Override.HeaderAppearance = appearance55;
			this.uGridPlanes.DisplayLayout.BandsSerializer.Add(ultraGridBand6);
			appearance56.ForeColor = System.Drawing.Color.Black;
			appearance56.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridPlanes.DisplayLayout.Override.ActiveRowAppearance = appearance56;
			this.uGridPlanes.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.uGridPlanes.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance57.BackColor = System.Drawing.Color.Transparent;
			this.uGridPlanes.DisplayLayout.Override.CardAreaAppearance = appearance57;
			appearance58.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance58.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance58.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance58.FontData.BoldAsString = "True";
			appearance58.FontData.Name = "Arial";
			appearance58.FontData.SizeInPoints = 10F;
			appearance58.ForeColor = System.Drawing.Color.White;
			appearance58.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridPlanes.DisplayLayout.Override.HeaderAppearance = appearance58;
			appearance59.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance59.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance59.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridPlanes.DisplayLayout.Override.RowAlternateAppearance = appearance59;
			appearance60.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance60.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance60.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridPlanes.DisplayLayout.Override.RowSelectorAppearance = appearance60;
			appearance61.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance61.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance61.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridPlanes.DisplayLayout.Override.SelectedRowAppearance = appearance61;
			this.uGridPlanes.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridPlanes.Location = new System.Drawing.Point(1256, 208);
			this.uGridPlanes.Name = "uGridPlanes";
			this.uGridPlanes.Size = new System.Drawing.Size(88, 150);
			this.uGridPlanes.TabIndex = 321;
			this.uGridPlanes.Visible = false;
			// 
			// uGridPlazos
			// 
			this.uGridPlazos.Cursor = System.Windows.Forms.Cursors.Default;
			this.uGridPlazos.DataSource = this.ultraDataSource12;
			appearance62.BackColor = System.Drawing.Color.White;
			this.uGridPlazos.DisplayLayout.Appearance = appearance62;
			this.uGridPlazos.DisplayLayout.AutoFitColumns = true;
			ultraGridColumn29.Header.VisiblePosition = 0;
			ultraGridColumn29.Hidden = true;
			ultraGridColumn29.Width = 60;
			ultraGridColumn30.Header.VisiblePosition = 1;
			ultraGridColumn30.Hidden = true;
			ultraGridColumn30.Width = 28;
			ultraGridColumn31.Header.VisiblePosition = 2;
			ultraGridColumn31.Hidden = true;
			ultraGridColumn31.Width = 41;
			ultraGridColumn32.Header.VisiblePosition = 3;
			ultraGridColumn32.Width = 22;
			ultraGridColumn33.Header.Caption = "...";
			ultraGridColumn33.Header.VisiblePosition = 5;
			ultraGridColumn33.Width = 17;
			appearance63.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn34.CellAppearance = appearance63;
			ultraGridColumn34.Format = "#,##0";
			ultraGridColumn34.Header.Caption = "M Gracia";
			ultraGridColumn34.Header.VisiblePosition = 4;
			ultraGridColumn34.Width = 12;
			ultraGridBand7.Columns.AddRange(new object[] {
																										 ultraGridColumn29,
																										 ultraGridColumn30,
																										 ultraGridColumn31,
																										 ultraGridColumn32,
																										 ultraGridColumn33,
																										 ultraGridColumn34});
			ultraGridBand7.Header.Caption = "PLAZOS";
			ultraGridBand7.HeaderVisible = true;
			ultraGridBand7.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			appearance64.FontData.SizeInPoints = 7.25F;
			ultraGridBand7.Override.HeaderAppearance = appearance64;
			this.uGridPlazos.DisplayLayout.BandsSerializer.Add(ultraGridBand7);
			appearance65.ForeColor = System.Drawing.Color.Black;
			appearance65.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridPlazos.DisplayLayout.Override.ActiveRowAppearance = appearance65;
			this.uGridPlazos.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.uGridPlazos.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance66.BackColor = System.Drawing.Color.Transparent;
			this.uGridPlazos.DisplayLayout.Override.CardAreaAppearance = appearance66;
			appearance67.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance67.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance67.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance67.FontData.BoldAsString = "True";
			appearance67.FontData.Name = "Arial";
			appearance67.FontData.SizeInPoints = 10F;
			appearance67.ForeColor = System.Drawing.Color.White;
			appearance67.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridPlazos.DisplayLayout.Override.HeaderAppearance = appearance67;
			appearance68.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance68.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance68.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridPlazos.DisplayLayout.Override.RowAlternateAppearance = appearance68;
			appearance69.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance69.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance69.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridPlazos.DisplayLayout.Override.RowSelectorAppearance = appearance69;
			appearance70.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance70.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance70.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridPlazos.DisplayLayout.Override.SelectedRowAppearance = appearance70;
			this.uGridPlazos.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridPlazos.Location = new System.Drawing.Point(1176, 208);
			this.uGridPlazos.Name = "uGridPlazos";
			this.uGridPlazos.Size = new System.Drawing.Size(72, 150);
			this.uGridPlazos.TabIndex = 322;
			this.uGridPlazos.Visible = false;
			// 
			// ultraDataSource17
			// 
			ultraDataColumn106.DataType = typeof(bool);
			this.ultraDataSource17.Band.Columns.AddRange(new object[] {
																																	ultraDataColumn102,
																																	ultraDataColumn103,
																																	ultraDataColumn104,
																																	ultraDataColumn105,
																																	ultraDataColumn106});
			// 
			// txtidHoraPromocion
			// 
			this.txtidHoraPromocion.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.txtidHoraPromocion.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtidHoraPromocion.Enabled = false;
			this.txtidHoraPromocion.Location = new System.Drawing.Point(1072, 8);
			this.txtidHoraPromocion.Name = "txtidHoraPromocion";
			this.txtidHoraPromocion.PromptChar = ' ';
			this.txtidHoraPromocion.Size = new System.Drawing.Size(8, 22);
			this.txtidHoraPromocion.TabIndex = 332;
			// 
			// label4
			// 
			this.label4.Location = new System.Drawing.Point(8, 80);
			this.label4.Name = "label4";
			this.label4.TabIndex = 333;
			this.label4.Text = "Paginas:";
			// 
			// lblnumproducto
			// 
			this.lblnumproducto.Location = new System.Drawing.Point(232, 80);
			this.lblnumproducto.Name = "lblnumproducto";
			this.lblnumproducto.Size = new System.Drawing.Size(136, 23);
			this.lblnumproducto.TabIndex = 335;
			this.lblnumproducto.Text = "Número de Productos";
			this.lblnumproducto.Click += new System.EventHandler(this.label5_Click);
			// 
			// rdgProductoFacturable
			// 
			this.rdgProductoFacturable.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.rdgProductoFacturable.EditValue = 0;
			this.rdgProductoFacturable.Location = new System.Drawing.Point(8, 721);
			this.rdgProductoFacturable.Name = "rdgProductoFacturable";
			// 
			// rdgProductoFacturable.Properties
			// 
			this.rdgProductoFacturable.Properties.Appearance.BackColor = System.Drawing.Color.Transparent;
			this.rdgProductoFacturable.Properties.Appearance.Options.UseBackColor = true;
			this.rdgProductoFacturable.Properties.Items.AddRange(new DevExpress.XtraEditors.Controls.RadioGroupItem[] {
																																																									new DevExpress.XtraEditors.Controls.RadioGroupItem(0, "Todos los productos"),
																																																									new DevExpress.XtraEditors.Controls.RadioGroupItem(1, "No Descontinuados"),
																																																									new DevExpress.XtraEditors.Controls.RadioGroupItem(2, "Descontinuados"),
																																																									new DevExpress.XtraEditors.Controls.RadioGroupItem(3, "Combos")});
			this.rdgProductoFacturable.Size = new System.Drawing.Size(464, 24);
			this.rdgProductoFacturable.TabIndex = 339;
			this.rdgProductoFacturable.Visible = false;
			this.rdgProductoFacturable.SelectedIndexChanged += new System.EventHandler(this.rdgProductoFacturable_SelectedIndexChanged);
			// 
			// rdgProductoExistencia
			// 
			this.rdgProductoExistencia.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.rdgProductoExistencia.EditValue = 1;
			this.rdgProductoExistencia.Location = new System.Drawing.Point(8, 697);
			this.rdgProductoExistencia.Name = "rdgProductoExistencia";
			// 
			// rdgProductoExistencia.Properties
			// 
			this.rdgProductoExistencia.Properties.Appearance.BackColor = System.Drawing.Color.Transparent;
			this.rdgProductoExistencia.Properties.Appearance.Options.UseBackColor = true;
			this.rdgProductoExistencia.Properties.Items.AddRange(new DevExpress.XtraEditors.Controls.RadioGroupItem[] {
																																																									new DevExpress.XtraEditors.Controls.RadioGroupItem(0, "Todos los productos"),
																																																									new DevExpress.XtraEditors.Controls.RadioGroupItem(1, "Con existencia"),
																																																									new DevExpress.XtraEditors.Controls.RadioGroupItem(2, "Sin existencia"),
																																																									new DevExpress.XtraEditors.Controls.RadioGroupItem(3, "En remates")});
			this.rdgProductoExistencia.Size = new System.Drawing.Size(464, 24);
			this.rdgProductoExistencia.TabIndex = 338;
			this.rdgProductoExistencia.Visible = false;
			// 
			// txtBuscar
			// 
			this.txtBuscar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			appearance71.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtBuscar.Appearance = appearance71;
			this.txtBuscar.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtBuscar.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtBuscar.Location = new System.Drawing.Point(480, 721);
			this.txtBuscar.Name = "txtBuscar";
			this.txtBuscar.Size = new System.Drawing.Size(250, 22);
			this.txtBuscar.TabIndex = 341;
			// 
			// lblContador
			// 
			this.lblContador.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.lblContador.AutoSize = true;
			this.lblContador.Font = new System.Drawing.Font("Tahoma", 10F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblContador.ForeColor = System.Drawing.Color.Firebrick;
			this.lblContador.Location = new System.Drawing.Point(800, 72);
			this.lblContador.Name = "lblContador";
			this.lblContador.Size = new System.Drawing.Size(0, 20);
			this.lblContador.TabIndex = 632;
			this.lblContador.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// lbltotalpagcatalogo
			// 
			this.lbltotalpagcatalogo.Location = new System.Drawing.Point(0, 40);
			this.lbltotalpagcatalogo.Name = "lbltotalpagcatalogo";
			this.lbltotalpagcatalogo.Size = new System.Drawing.Size(100, 16);
			this.lbltotalpagcatalogo.TabIndex = 633;
			this.lbltotalpagcatalogo.Text = "Páginas Catálogo";
			// 
			// label5
			// 
			this.label5.Location = new System.Drawing.Point(656, 40);
			this.label5.Name = "label5";
			this.label5.Size = new System.Drawing.Size(100, 16);
			this.label5.TabIndex = 635;
			this.label5.Text = "Estado Catálogo";
			// 
			// txtestadocatalogo
			// 
			this.txtestadocatalogo.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtestadocatalogo.Enabled = false;
			this.txtestadocatalogo.Location = new System.Drawing.Point(752, 40);
			this.txtestadocatalogo.Name = "txtestadocatalogo";
			this.txtestadocatalogo.Size = new System.Drawing.Size(184, 20);
			this.txtestadocatalogo.TabIndex = 636;
			this.txtestadocatalogo.Text = "";
			// 
			// btcrearcatalogo
			// 
			this.btcrearcatalogo.Appearance.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.btcrearcatalogo.Appearance.Options.UseFont = true;
			this.btcrearcatalogo.Appearance.Options.UseTextOptions = true;
			this.btcrearcatalogo.Appearance.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Near;
			this.btcrearcatalogo.Enabled = false;
			this.btcrearcatalogo.ImageAlignment = DevExpress.Utils.HorzAlignment.Far;
			this.btcrearcatalogo.ImageIndex = 35;
			this.btcrearcatalogo.ImageList = this.imageList1;
			this.btcrearcatalogo.Location = new System.Drawing.Point(952, 32);
			this.btcrearcatalogo.Name = "btcrearcatalogo";
			this.btcrearcatalogo.Size = new System.Drawing.Size(136, 24);
			this.btcrearcatalogo.TabIndex = 637;
			this.btcrearcatalogo.Text = "&Guardar Catálogo";
			this.btcrearcatalogo.Click += new System.EventHandler(this.btcrearcatalogo_Click);
			// 
			// txtidCatalogo
			// 
			this.txtidCatalogo.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.txtidCatalogo.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtidCatalogo.Enabled = false;
			this.txtidCatalogo.Location = new System.Drawing.Point(240, 137);
			this.txtidCatalogo.Name = "txtidCatalogo";
			this.txtidCatalogo.PromptChar = ' ';
			this.txtidCatalogo.Size = new System.Drawing.Size(16, 22);
			this.txtidCatalogo.SpinButtonAlignment = Infragistics.Win.ButtonAlignment.Left;
			this.txtidCatalogo.TabIndex = 641;
			this.txtidCatalogo.Visible = false;
			// 
			// txtpaginastotalcatalogo
			// 
			this.txtpaginastotalcatalogo.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.txtpaginastotalcatalogo.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtpaginastotalcatalogo.Enabled = false;
			this.txtpaginastotalcatalogo.Location = new System.Drawing.Point(96, 40);
			this.txtpaginastotalcatalogo.Name = "txtpaginastotalcatalogo";
			this.txtpaginastotalcatalogo.PromptChar = ' ';
			this.txtpaginastotalcatalogo.Size = new System.Drawing.Size(128, 22);
			this.txtpaginastotalcatalogo.SpinButtonAlignment = Infragistics.Win.ButtonAlignment.Left;
			this.txtpaginastotalcatalogo.TabIndex = 642;
			this.txtpaginastotalcatalogo.ValueChanged += new System.EventHandler(this.txtpaginastotalcatalogo_ValueChanged);
			// 
			// cmbTipo
			// 
			appearance72.FontData.Name = "Tahoma";
			appearance72.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbTipo.Appearance = appearance72;
			this.cmbTipo.DropDownStyle = Infragistics.Win.DropDownStyle.DropDownList;
			this.cmbTipo.Enabled = false;
			valueListItem1.DataValue = "Credito";
			valueListItem1.DisplayText = "Crédito";
			valueListItem2.DataValue = "Efectivo";
			valueListItem2.DisplayText = "Efectivo";
			this.cmbTipo.Items.Add(valueListItem1);
			this.cmbTipo.Items.Add(valueListItem2);
			this.cmbTipo.Location = new System.Drawing.Point(96, 8);
			this.cmbTipo.Name = "cmbTipo";
			this.cmbTipo.Size = new System.Drawing.Size(128, 22);
			this.cmbTipo.TabIndex = 643;
			this.cmbTipo.ValueChanged += new System.EventHandler(this.cmbTipo_ValueChanged_1);
			// 
			// btnuevocatalogo
			// 
			this.btnuevocatalogo.Appearance.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.btnuevocatalogo.Appearance.Options.UseFont = true;
			this.btnuevocatalogo.Appearance.Options.UseTextOptions = true;
			this.btnuevocatalogo.Appearance.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Near;
			this.btnuevocatalogo.Enabled = false;
			this.btnuevocatalogo.ImageAlignment = DevExpress.Utils.HorzAlignment.Far;
			this.btnuevocatalogo.ImageIndex = 8;
			this.btnuevocatalogo.ImageList = this.imageList1;
			this.btnuevocatalogo.Location = new System.Drawing.Point(952, 8);
			this.btnuevocatalogo.Name = "btnuevocatalogo";
			this.btnuevocatalogo.Size = new System.Drawing.Size(136, 24);
			this.btnuevocatalogo.TabIndex = 644;
			this.btnuevocatalogo.Text = "&Nuevo Catálogo";
			this.btnuevocatalogo.Click += new System.EventHandler(this.btnuevocatalogo_Click);
			// 
			// uGridResumen
			// 
			this.uGridResumen.Cursor = System.Windows.Forms.Cursors.Default;
			this.uGridResumen.DataSource = this.ultraDataSource1;
			appearance73.BackColor = System.Drawing.Color.White;
			appearance73.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridResumen.DisplayLayout.Appearance = appearance73;
			this.uGridResumen.DisplayLayout.AutoFitColumns = true;
			ultraGridColumn35.Header.VisiblePosition = 2;
			ultraGridColumn35.Hidden = true;
			ultraGridColumn35.Width = 168;
			ultraGridColumn36.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn36.Header.VisiblePosition = 3;
			ultraGridColumn36.Hidden = true;
			ultraGridColumn36.Width = 93;
			ultraGridColumn37.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn37.Header.VisiblePosition = 1;
			ultraGridColumn37.Width = 367;
			ultraGridColumn38.Header.VisiblePosition = 4;
			ultraGridColumn38.Width = 132;
			ultraGridColumn39.Header.VisiblePosition = 5;
			ultraGridColumn39.Width = 71;
			ultraGridColumn40.Header.VisiblePosition = 6;
			ultraGridColumn40.Hidden = true;
			ultraGridColumn40.Width = 66;
			appearance74.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(255)), ((System.Byte)(128)), ((System.Byte)(0)));
			ultraGridColumn41.CellAppearance = appearance74;
			ultraGridColumn41.Header.VisiblePosition = 7;
			ultraGridColumn41.Width = 306;
			ultraGridColumn42.Header.VisiblePosition = 8;
			ultraGridColumn42.Hidden = true;
			ultraGridColumn42.Width = 110;
			ultraGridColumn43.Header.VisiblePosition = 9;
			ultraGridColumn43.Hidden = true;
			ultraGridColumn43.Width = 74;
			ultraGridColumn44.Header.VisiblePosition = 10;
			ultraGridColumn44.Hidden = true;
			ultraGridColumn44.Width = 90;
			ultraGridColumn45.ButtonDisplayStyle = Infragistics.Win.UltraWinGrid.ButtonDisplayStyle.Always;
			appearance75.Image = ((object)(resources.GetObject("appearance75.Image")));
			appearance75.ImageHAlign = Infragistics.Win.HAlign.Center;
			appearance75.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn45.CellButtonAppearance = appearance75;
			ultraGridColumn45.Header.Caption = "....";
			ultraGridColumn45.Header.VisiblePosition = 0;
			ultraGridColumn45.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.Button;
			ultraGridColumn45.Width = 55;
			ultraGridBand8.Columns.AddRange(new object[] {
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
																										 ultraGridColumn45});
			ultraGridBand8.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			appearance76.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings1.Appearance = appearance76;
			summarySettings1.DisplayFormat = "{0: #,##0}";
			summarySettings1.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			ultraGridBand8.Summaries.AddRange(new Infragistics.Win.UltraWinGrid.SummarySettings[] {
																																															summarySettings1});
			ultraGridBand8.SummaryFooterCaption = "Totales";
			this.uGridResumen.DisplayLayout.BandsSerializer.Add(ultraGridBand8);
			appearance77.ForeColor = System.Drawing.Color.Black;
			appearance77.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridResumen.DisplayLayout.Override.ActiveRowAppearance = appearance77;
			this.uGridResumen.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.uGridResumen.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance78.BackColor = System.Drawing.Color.Transparent;
			this.uGridResumen.DisplayLayout.Override.CardAreaAppearance = appearance78;
			this.uGridResumen.DisplayLayout.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.Edit;
			appearance79.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance79.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance79.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance79.FontData.BoldAsString = "True";
			appearance79.FontData.Name = "Arial";
			appearance79.FontData.SizeInPoints = 8.25F;
			appearance79.ForeColor = System.Drawing.Color.White;
			appearance79.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridResumen.DisplayLayout.Override.HeaderAppearance = appearance79;
			this.uGridResumen.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortSingle;
			appearance80.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance80.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance80.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridResumen.DisplayLayout.Override.RowAlternateAppearance = appearance80;
			appearance81.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance81.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance81.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridResumen.DisplayLayout.Override.RowSelectorAppearance = appearance81;
			appearance82.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance82.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance82.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridResumen.DisplayLayout.Override.SelectedRowAppearance = appearance82;
			this.uGridResumen.DisplayLayout.ScrollBounds = Infragistics.Win.UltraWinGrid.ScrollBounds.ScrollToFill;
			this.uGridResumen.Font = new System.Drawing.Font("Tahoma", 8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridResumen.Location = new System.Drawing.Point(0, 304);
			this.uGridResumen.Name = "uGridResumen";
			this.uGridResumen.Size = new System.Drawing.Size(952, 160);
			this.uGridResumen.TabIndex = 645;
			this.uGridResumen.ClickCellButton += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.uGridResumen_ClickCellButton);
			this.uGridResumen.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.uGridResumen_InitializeLayout);
			// 
			// uGridDetalle
			// 
			this.uGridDetalle.Cursor = System.Windows.Forms.Cursors.Default;
			this.uGridDetalle.DataSource = this.ultraDataSource2;
			appearance83.BackColor = System.Drawing.Color.White;
			this.uGridDetalle.DisplayLayout.Appearance = appearance83;
			this.uGridDetalle.DisplayLayout.AutoFitColumns = true;
			appearance84.FontData.SizeInPoints = 7F;
			ultraGridColumn46.Header.Appearance = appearance84;
			ultraGridColumn46.Header.VisiblePosition = 0;
			ultraGridColumn46.Width = 127;
			appearance85.FontData.SizeInPoints = 7F;
			ultraGridColumn47.Header.Appearance = appearance85;
			ultraGridColumn47.Header.VisiblePosition = 1;
			ultraGridColumn47.Width = 151;
			appearance86.FontData.SizeInPoints = 7F;
			ultraGridColumn48.Header.Appearance = appearance86;
			ultraGridColumn48.Header.VisiblePosition = 2;
			ultraGridColumn48.Width = 138;
			appearance87.FontData.SizeInPoints = 7F;
			ultraGridColumn49.Header.Appearance = appearance87;
			ultraGridColumn49.Header.VisiblePosition = 3;
			ultraGridColumn49.Width = 77;
			appearance88.FontData.SizeInPoints = 7F;
			ultraGridColumn50.Header.Appearance = appearance88;
			ultraGridColumn50.Header.Caption = "Costo C/IVA";
			ultraGridColumn50.Header.FixedHeaderIndicator = Infragistics.Win.UltraWinGrid.FixedHeaderIndicator.None;
			ultraGridColumn50.Header.VisiblePosition = 4;
			ultraGridColumn50.Width = 97;
			appearance89.FontData.SizeInPoints = 7F;
			ultraGridColumn51.Header.Appearance = appearance89;
			ultraGridColumn51.Header.Caption = "Utilidad Efectivo";
			ultraGridColumn51.Header.VisiblePosition = 5;
			ultraGridColumn51.Width = 119;
			appearance90.FontData.SizeInPoints = 7F;
			ultraGridColumn52.Header.Appearance = appearance90;
			ultraGridColumn52.Header.Caption = "Utilidad TC";
			ultraGridColumn52.Header.VisiblePosition = 16;
			ultraGridColumn52.Width = 111;
			appearance91.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(255)), ((System.Byte)(255)), ((System.Byte)(192)));
			ultraGridColumn53.CellAppearance = appearance91;
			ultraGridColumn53.Format = "#,##0.00";
			appearance92.FontData.SizeInPoints = 7F;
			appearance92.TextVAlign = Infragistics.Win.VAlign.Middle;
			ultraGridColumn53.Header.Appearance = appearance92;
			ultraGridColumn53.Header.Caption = "Nuevo precio con IVA";
			ultraGridColumn53.Header.VisiblePosition = 15;
			ultraGridColumn53.Width = 111;
			ultraGridColumn54.Header.VisiblePosition = 6;
			ultraGridColumn54.Hidden = true;
			ultraGridColumn54.Width = 45;
			ultraGridColumn55.Header.VisiblePosition = 7;
			ultraGridColumn55.Hidden = true;
			ultraGridColumn55.Width = 44;
			ultraGridColumn56.Header.VisiblePosition = 8;
			ultraGridColumn56.Hidden = true;
			ultraGridColumn56.Width = 49;
			ultraGridColumn57.Header.VisiblePosition = 9;
			ultraGridColumn57.Hidden = true;
			ultraGridColumn57.HiddenWhenGroupBy = Infragistics.Win.DefaultableBoolean.True;
			ultraGridColumn57.Width = 48;
			ultraGridColumn58.Header.VisiblePosition = 10;
			ultraGridColumn58.Hidden = true;
			ultraGridColumn58.Width = 53;
			ultraGridColumn59.Header.VisiblePosition = 11;
			ultraGridColumn59.Hidden = true;
			ultraGridColumn59.Width = 59;
			ultraGridColumn60.Header.VisiblePosition = 12;
			ultraGridColumn60.Hidden = true;
			ultraGridColumn60.Width = 55;
			ultraGridColumn61.Header.VisiblePosition = 13;
			ultraGridColumn61.Hidden = true;
			ultraGridColumn61.Width = 55;
			ultraGridColumn62.Header.VisiblePosition = 14;
			ultraGridColumn62.Hidden = true;
			ultraGridColumn62.Width = 81;
			appearance93.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(192)), ((System.Byte)(255)), ((System.Byte)(192)));
			ultraGridColumn63.CellAppearance = appearance93;
			ultraGridColumn63.Format = "#,##0.00";
			appearance94.FontData.SizeInPoints = 7F;
			ultraGridColumn63.Header.Appearance = appearance94;
			ultraGridColumn63.Header.Caption = "Nuevo PVP";
			ultraGridColumn63.Header.VisiblePosition = 17;
			ultraGridColumn63.Hidden = true;
			ultraGridColumn63.Width = 77;
			ultraGridBand9.Columns.AddRange(new object[] {
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
																										 ultraGridColumn63});
			ultraGridBand9.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.Edit;
			this.uGridDetalle.DisplayLayout.BandsSerializer.Add(ultraGridBand9);
			appearance95.ForeColor = System.Drawing.Color.Black;
			appearance95.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridDetalle.DisplayLayout.Override.ActiveRowAppearance = appearance95;
			this.uGridDetalle.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.True;
			this.uGridDetalle.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance96.BackColor = System.Drawing.Color.Transparent;
			this.uGridDetalle.DisplayLayout.Override.CardAreaAppearance = appearance96;
			appearance97.ForeColor = System.Drawing.Color.Black;
			appearance97.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridDetalle.DisplayLayout.Override.CellAppearance = appearance97;
			this.uGridDetalle.DisplayLayout.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.Edit;
			appearance98.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance98.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance98.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance98.FontData.BoldAsString = "True";
			appearance98.FontData.Name = "Arial";
			appearance98.FontData.SizeInPoints = 9F;
			appearance98.ForeColor = System.Drawing.Color.White;
			appearance98.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridDetalle.DisplayLayout.Override.HeaderAppearance = appearance98;
			appearance99.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance99.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance99.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridDetalle.DisplayLayout.Override.RowAlternateAppearance = appearance99;
			appearance100.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance100.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance100.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridDetalle.DisplayLayout.Override.RowSelectorAppearance = appearance100;
			appearance101.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance101.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance101.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridDetalle.DisplayLayout.Override.SelectedRowAppearance = appearance101;
			this.uGridDetalle.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridDetalle.Location = new System.Drawing.Point(0, 472);
			this.uGridDetalle.Name = "uGridDetalle";
			this.uGridDetalle.Size = new System.Drawing.Size(952, 208);
			this.uGridDetalle.TabIndex = 646;
			this.uGridDetalle.KeyDown += new System.Windows.Forms.KeyEventHandler(this.uGridDetalle_KeyDown);
			this.uGridDetalle.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.uGridDetalle_InitializeLayout);
			// 
			// groupBox2
			// 
			this.groupBox2.Location = new System.Drawing.Point(-8, 288);
			this.groupBox2.Name = "groupBox2";
			this.groupBox2.Size = new System.Drawing.Size(1167, 8);
			this.groupBox2.TabIndex = 647;
			this.groupBox2.TabStop = false;
			// 
			// btExportar
			// 
			this.btExportar.Appearance.Options.UseTextOptions = true;
			this.btExportar.Appearance.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Far;
			this.btExportar.Image = ((System.Drawing.Image)(resources.GetObject("btExportar.Image")));
			this.btExportar.ImageAlignment = DevExpress.Utils.HorzAlignment.Far;
			this.btExportar.ImageIndex = 2;
			this.btExportar.Location = new System.Drawing.Point(984, 256);
			this.btExportar.Name = "btExportar";
			this.btExportar.Size = new System.Drawing.Size(96, 32);
			this.btExportar.TabIndex = 648;
			this.btExportar.Text = "&Exportar";
			this.btExportar.Click += new System.EventHandler(this.btExportar_Click);
			// 
			// txtpaginas
			// 
			this.txtpaginas.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.txtpaginas.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtpaginas.Enabled = false;
			this.txtpaginas.Location = new System.Drawing.Point(72, 72);
			this.txtpaginas.Name = "txtpaginas";
			this.txtpaginas.PromptChar = ' ';
			this.txtpaginas.Size = new System.Drawing.Size(128, 22);
			this.txtpaginas.SpinButtonAlignment = Infragistics.Win.ButtonAlignment.Left;
			this.txtpaginas.TabIndex = 649;
			// 
			// txtnumproductos
			// 
			this.txtnumproductos.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.txtnumproductos.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtnumproductos.Enabled = false;
			this.txtnumproductos.Location = new System.Drawing.Point(368, 72);
			this.txtnumproductos.Name = "txtnumproductos";
			this.txtnumproductos.PromptChar = ' ';
			this.txtnumproductos.Size = new System.Drawing.Size(128, 22);
			this.txtnumproductos.SpinButtonAlignment = Infragistics.Win.ButtonAlignment.Left;
			this.txtnumproductos.TabIndex = 650;
			// 
			// txtidDetCatalogo
			// 
			this.txtidDetCatalogo.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.txtidDetCatalogo.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtidDetCatalogo.Enabled = false;
			this.txtidDetCatalogo.Location = new System.Drawing.Point(208, 72);
			this.txtidDetCatalogo.Name = "txtidDetCatalogo";
			this.txtidDetCatalogo.PromptChar = ' ';
			this.txtidDetCatalogo.Size = new System.Drawing.Size(16, 22);
			this.txtidDetCatalogo.SpinButtonAlignment = Infragistics.Win.ButtonAlignment.Left;
			this.txtidDetCatalogo.TabIndex = 651;
			this.txtidDetCatalogo.Visible = false;
			// 
			// frmCatalogo1
			// 
			this.Appearance.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.Appearance.Options.UseFont = true;
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 14);
			this.ClientSize = new System.Drawing.Size(1098, 748);
			this.Controls.Add(this.txtidDetCatalogo);
			this.Controls.Add(this.txtnumproductos);
			this.Controls.Add(this.txtpaginas);
			this.Controls.Add(this.btExportar);
			this.Controls.Add(this.groupBox2);
			this.Controls.Add(this.uGridDetalle);
			this.Controls.Add(this.uGridResumen);
			this.Controls.Add(this.btnuevocatalogo);
			this.Controls.Add(this.cmbTipo);
			this.Controls.Add(this.txtpaginastotalcatalogo);
			this.Controls.Add(this.txtidCatalogo);
			this.Controls.Add(this.btcrearcatalogo);
			this.Controls.Add(this.txtestadocatalogo);
			this.Controls.Add(this.label16);
			this.Controls.Add(this.label3);
			this.Controls.Add(this.label2);
			this.Controls.Add(this.label15);
			this.Controls.Add(this.lblCliente);
			this.Controls.Add(this.label1);
			this.Controls.Add(this.label6);
			this.Controls.Add(this.label5);
			this.Controls.Add(this.lbltotalpagcatalogo);
			this.Controls.Add(this.lblContador);
			this.Controls.Add(this.txtBuscar);
			this.Controls.Add(this.rdgProductoFacturable);
			this.Controls.Add(this.rdgProductoExistencia);
			this.Controls.Add(this.lblnumproducto);
			this.Controls.Add(this.label4);
			this.Controls.Add(this.txtidHoraPromocion);
			this.Controls.Add(this.uGridPlazos);
			this.Controls.Add(this.uGridPlanes);
			this.Controls.Add(this.chkCupon);
			this.Controls.Add(this.chkCuotas);
			this.Controls.Add(this.uGridCuotas);
			this.Controls.Add(this.chkOpcional);
			this.Controls.Add(this.chkAdicional);
			this.Controls.Add(this.chkSumaDescuentos);
			this.Controls.Add(this.uGridModelos);
			this.Controls.Add(this.chkMarcas);
			this.Controls.Add(this.chkSubGrupos);
			this.Controls.Add(this.chkGrupos);
			this.Controls.Add(this.groupBox1);
			this.Controls.Add(this.txtidArticulo);
			this.Controls.Add(this.btFinalizar);
			this.Controls.Add(this.btCancelar);
			this.Controls.Add(this.btBuscar);
			this.Controls.Add(this.btGrabar);
			this.Controls.Add(this.btNuevo);
			this.Controls.Add(this.btSalir);
			this.Controls.Add(this.btEditar);
			this.Controls.Add(this.uGridMarcas);
			this.Controls.Add(this.uGridSubgrupos);
			this.Controls.Add(this.uGridGrupos);
			this.Controls.Add(this.txtValorMaximo);
			this.Controls.Add(this.txtValorMinimo);
			this.Controls.Add(this.txtIdPromocion);
			this.Controls.Add(this.txtDescripcion);
			this.Controls.Add(this.dtFin);
			this.Controls.Add(this.dtInicio);
			this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
			this.KeyPreview = true;
			this.MaximizeBox = false;
			this.Name = "frmCatalogo1";
			this.StartPosition = System.Windows.Forms.FormStartPosition.Manual;
			this.Text = "Catalogo";
			this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.Promociones_KeyDown);
			this.Closing += new System.ComponentModel.CancelEventHandler(this.Promociones_Closing);
			this.Load += new System.EventHandler(this.Promociones_Load);
			((System.ComponentModel.ISupportInitialize)(this.dtFin)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtInicio)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDescripcion)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdPromocion)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtValorMaximo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtValorMinimo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource6)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource4)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridMarcas)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource9)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridSubgrupos)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource8)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridGrupos)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource7)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource5)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsPromocion)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtidArticulo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource10)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridModelos)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource11)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource12)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource13)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridCuotas)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource14)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource15)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource3)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource16)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridPlanes)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridPlazos)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource17)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtidHoraPromocion)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.rdgProductoFacturable.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.rdgProductoExistencia.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtBuscar)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtidCatalogo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtpaginastotalcatalogo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbTipo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridResumen)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridDetalle)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtpaginas)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtnumproductos)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtidDetCatalogo)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		CultureInfo us = new CultureInfo("en-US");
		bool bNuevo = false;		
		bool bModoEdicion = false;
		int iNumIngreso = 0;
		string sArticulo = "";		
		bool bReading = false;
		bool bSelLocales = false;		
		bool bSelSubGrupos = false;
		bool bSelMarcas = false;
		bool bSelCuotas = false;
		bool bSelDias = false;
		bool bSelHabitaciones = false;
		bool bComboArticulo = false;
		int idBloqueaTransacciones = 0;
		int idPromo = 0;
		
		public void EstadoGrid(bool Estado)
		{		
			if (Estado)
			{
							this.uGridCuotas.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;	
					this.uGridPlanes.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
				this.uGridPlazos.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
				this.uGridGrupos.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
				this.uGridSubgrupos.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
				this.uGridMarcas.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;	
				this.uGridModelos.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;					
		
								
			}
			else
			{
							this.uGridCuotas.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.False;	
						this.uGridPlanes.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.False;
				this.uGridPlazos.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.False;
				this.uGridGrupos.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.False;
				this.uGridSubgrupos.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.False;
				this.uGridMarcas.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.False;
				this.uGridModelos.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.False;
									}
		}
		
		private Acceso miAcceso;

		private void UnloadMe()
		{
			this.Close();
		}
		private void Promociones_Load(object sender, System.EventArgs e)
		{	
			miAcceso = new Acceso(cdsSeteoF, "0623");

			if (!Funcion.Permiso("988", cdsSeteoF))
			{				
				MessageBox.Show("No tiene Acceso a CATALOGO", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				BeginInvoke(new MethodInvoker(UnloadMe));
				return;
			}
			this.dtInicio.Value = DateTime.Now.ToShortTimeString();
			this.dtFin.Value = DateTime.Today; 
			Funcion.Resolucion(this, this.Height, this.Width);
//			idSubGrupo =(int)this.uGridSubgrupos.ActiveRow.Cells["idPromocionSubGrupo"].Value;	
			ActualizaMovimientos();
//			ListaArticulosPrecios();
			EstadoGrid(false);
			FuncionesProcedimientos.EstadoGrids(false, this.uGridResumen);
		
			this.btnuevocatalogo.Enabled=true;

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


		
	

		private void CalculaCuota()
		{
//			decimal dValor = 0;
//			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in ultraGrid1.Rows.All)
//			{
//				if (dr.Cells["PromoPrecio4"].Value != System.DBNull.Value) dValor = dValor + Convert.ToDecimal(dr.Cells["PromoPrecio4"].Value);
//			}
//
//			if ((int)this.txtNroCuotas.Value > 0) 
//			{
			//decimal dIVA = Funcion.decEscalarSQL(cdsSeteoF, "Select dbo.RetornaImpuestoIVAVigenteDecimal(GETDATE(), 1)");
//				dValor = dValor * dIVA;
//
//				string sSQL1 = string.Format("Select dbo.CalCuotaPromocion ({0}, {1})", dValor, (int)this.txtNroCuotas.Value);
//				this.txtCuotaCalculada.Value  = Funcion.decEscalarSQL(cdsSeteoF, sSQL1);
//			}
		}

		private void ultraGrid1_InitializeRow(object sender, Infragistics.Win.UltraWinGrid.InitializeRowEventArgs e)
		{
			
		}

		private void txtNroCuotas_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
//			if (e.KeyChar == 13)
//			{
//				if ((int)this.txtNroCuotas.Value == 0) return;
//				if (this.ultraGrid1.Rows.Count == 0) return;
//				decimal dValor = 0;
//				foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in ultraGrid1.Rows.All)
//				{
//					dValor = dValor + (decimal)dr.Cells["Precio4"].Value;
//				}
//
//				if (dValor == 0) return;
//				
//				DateTime dtFecha = DateTime.Today;
//				string sSQL1 = string.Format("Exec CalculaCuotaInicial {0}, {1}, 0, '{2}'", dValor, (int)this.txtNroCuotas.Value, dtFecha.ToString("yyyyMMdd"));
//				SqlDataReader dReader =	Funcion.rEscalarSQL(cdsSeteoF, sSQL1, true);
//				dReader.Read();
//				this.txtCuotaReal.Value = dReader.GetDecimal(2);
//				dReader.Close();
//			}
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
		
			
			this.txtIdPromocion.Value = Funcion.iEscalarSQL(cdsSeteoF, string.Format("Exec GrabaPromociones {0}, 'ACTIVO'", (int)this.txtIdPromocion.Value));
//			string sSQLVentas = string.Format("Exec ResumenCalculoCatalogo {0}", 0);
//			this.uGridResumen.DataSource = Funcion.dvProcedimiento(cdsSeteoF, sSQLVentas);	
//			
			this.uGridGrupos.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Exec ListaGruposPromocion 0");		
			//  this.uGridDetalle.DataSource = Funcion.dvProcedimiento(cdsSeteoF, string.Format("Exec GuardaDetalleCatalogo 0, 0, 0, 0, 0, '', '', 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0,0, 0"));
//			this.uGridDetalle.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Select Null As idDetCatalogo, Null As idCatalogo From DetalleCatalogo Where idDetCatalogo= 0");
			if (this.uGridDetalle.Rows.Count > 0)
				if(this.cmbTipo.Text=="Efectivo")
				{
					this.uGridDetalle.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Select Null As idArticulo,Null As Articulo, Null As Codigo, Null As Cantidad, Null as CostoUltimo, Null as UtilidadEfectivo, Null as UtilidadTC, Null as Precio3, Null as cuo6, Null as cuo9, Null as cuo12, Null as cuo24");// From DetalleCatalogo Where idDetCatalogo= 0
					this.uGridDetalle.DisplayLayout.Bands[0].Columns["cuo6"].Hidden = false;	
					this.uGridDetalle.DisplayLayout.Bands[0].Columns["cuo9"].Hidden = false;
					this.uGridDetalle.DisplayLayout.Bands[0].Columns["cuo12"].Hidden = false;
					this.uGridDetalle.DisplayLayout.Bands[0].Columns["cuo24"].Hidden = false;

				}

				else if(this.cmbTipo.Text=="Crédito")
				{
					this.uGridDetalle.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Select Null As idArticulo,Null As Articulo, Null As Codigo, Null As Cantidad, Null as CostoUltimo, Null as UtilidadEfectivo, Null as UtilidadTC, Null as Precio3, Null as cuc3, Null as cuc6, Null as cuc9, Null as cuc12, Null as cuc18, Null as Nuevopvp");// From DetalleCatalogo Where idDetCatalogo= 0
					this.uGridDetalle.DisplayLayout.Bands[0].Columns["Nuevopvp"].Hidden = false;
					this.uGridDetalle.DisplayLayout.Bands[0].Columns["cuc3"].Hidden = false;
					this.uGridDetalle.DisplayLayout.Bands[0].Columns["cuc6"].Hidden = false;	
					this.uGridDetalle.DisplayLayout.Bands[0].Columns["cuc9"].Hidden = false;	
					this.uGridDetalle.DisplayLayout.Bands[0].Columns["cuc12"].Hidden = false;	
					this.uGridDetalle.DisplayLayout.Bands[0].Columns["cuc18"].Hidden = false;	
				}
			
			//this.uGridDetalle.DataSource = this.ultraDataSource2;//Funcion.dvProcedimiento(cdsSeteoF, "Exec ConsultaDetalleFactura 0");			
			this.uGridResumen.DataSource = this.ultraDataSource1;
			this.txtpaginas.Enabled=true;
			this.txtnumproductos.Enabled= true;
			EstadoGrid(true);
				LimpiaGrids();
   			
			iNumIngreso = 1;

//			FuncionesProcedimientos.EstadoGrids(false, this.uGridDetalle);
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
					this.uGridCuotas.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Select 0 As idPromocionesCuotas, 0 As idPromocion, Cuota, CONVERT(Bit, 0) As Sel From CuotasCrediPOINT Where Cuota = 0");
     
				this.uGridPlanes.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Select Null As idPromocionesPlanes, NULL As idPromocionBancoTarjeta, Null As idPromocionBanco, idPlan, Null As idPromocion, Descripcion, CONVERT(Bit, 0) As Sel From Planes Where idPlan = 0");
			this.uGridPlazos.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Select Null As idPromocionesPlazos, idPlazo, Null As idPromocion, Descripcion, CONVERT(Bit, 0) As Sel  From Plazos Where idPlazo = 0");
			
			//this.uGridGrupos.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Declare @False As Bit = 0 Select Null As idPromocionGrupo, Null As idPromocion, idGrupoArticulo, Grupo, @False As Sel From ArticuloGrupo Where idGrupoArticulo = 0");
			this.uGridSubgrupos.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Declare @False As Bit = 0 Select Null As idPromocionSubGrupo, idSubGrupo, SubGrupo, @False As Sel From ArticuloSubGrupo Where idSubGrupo = 0");
			this.uGridMarcas.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Declare @False As Bit = 0 Select Null As idPromocionMarca, idMarca, Marca, @False As Sel From ArticuloMarca Where idMarca = 0");
			this.uGridModelos.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Select 0 As idPromocionModelo, idArticulo, Modelo, CONVERT(Bit, 0) As Sel From Articulo Where idArticulo = 0");
			
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
			this.txtidHoraPromocion.Value=0;
			this.cmbTipo.Text = "";
			this.dtInicio.Value = DateTime.Today;
			this.dtFin.Value = DateTime.Today;



			this.txtDescripcion.Text = "";
			this.txtValorMinimo.Value = 0.00m;
			this.txtValorMaximo.Value = 0.00m;


			this.txtidArticulo.Value = 0;
		

			this.txtIdPromocion.Enabled = false;
			this.cmbTipo.Enabled = false;
			this.dtInicio.Enabled = false;
			this.dtFin.Enabled = false;
			this.txtDescripcion.Enabled = false;
			this.txtValorMinimo.Enabled = false;
			this.txtValorMaximo.Enabled = false;
		
			
			#endregion Limpiar - Habilitar
//			this.uGridDetalle.DataSource = this.ultraDataSource2;//Funcion.dvProcedimiento(cdsSeteoF, "Exec ConsultaDetalleFactura 0");			
			this.uGridResumen.DataSource = this.ultraDataSource1;
			
				if(this.cmbTipo.Text=="Efectivo")
				{
					this.uGridDetalle.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Select Null As idArticulo,Null As Articulo, Null As Codigo, Null As Cantidad, Null as CostoUltimo, Null as UtilidadEfectivo, Null as UtilidadTC, Null as Precio3, Null as cuo6, Null as cuo9, Null as cuo12, Null as cuo24");// From DetalleCatalogo Where idDetCatalogo= 0
					this.uGridDetalle.DisplayLayout.Bands[0].Columns["cuo6"].Hidden = false;	
					this.uGridDetalle.DisplayLayout.Bands[0].Columns["cuo9"].Hidden = false;
					this.uGridDetalle.DisplayLayout.Bands[0].Columns["cuo12"].Hidden = false;
					this.uGridDetalle.DisplayLayout.Bands[0].Columns["cuo24"].Hidden = false;
					//					this.uGridDetalle.DisplayLayout.Bands[0].Columns["Nuevopvp"].Hidden = true;
					this.uGridDetalle.DisplayLayout.Bands[0].Columns["cuc3"].Hidden = true;
					this.uGridDetalle.DisplayLayout.Bands[0].Columns["cuc6"].Hidden = true;	
					this.uGridDetalle.DisplayLayout.Bands[0].Columns["cuc9"].Hidden = true;	
					this.uGridDetalle.DisplayLayout.Bands[0].Columns["cuc12"].Hidden = true;	
					this.uGridDetalle.DisplayLayout.Bands[0].Columns["cuc18"].Hidden = true;	
				}

				else if(this.cmbTipo.Text=="Crédito")
				{
					this.uGridDetalle.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Select Null As idArticulo,Null As Articulo, Null As Codigo, Null As Cantidad, Null as CostoUltimo, Null as UtilidadEfectivo, Null as UtilidadTC, Null as Precio3, Null as cuc3, Null as cuc6, Null as cuc9, Null as cuc12, Null as cuc18, Null as Nuevopvp");// From DetalleCatalogo Where idDetCatalogo= 0
	
				}
			
			//EstadoGrid(false);
			LimpiaGrids();

			this.btCancelar.Enabled = false;
			this.btGrabar.Enabled = false;
			this.btEditar.Enabled = false;
			this.btFinalizar.Enabled = false;
			this.btNuevo.Enabled = true;
			this.btBuscar.Enabled = true;				
	
			this.chkCuotas.Enabled = false;
			bSelCuotas = false;
			this.chkCuotas.Checked = false;

			

//			this.chkHabitaciones.Enabled = false;
//			bSelHabitaciones = false;
//			this.chkHabitaciones.Checked = false;

			this.chkGrupos.Enabled = false;			
			this.chkGrupos.Checked = false;

			this.chkSubGrupos.Enabled = false;
			bSelSubGrupos = false;
			this.chkSubGrupos.Checked = false;

			this.chkMarcas.Enabled = false;
			bSelMarcas = false;
			this.chkMarcas.Checked = false;	
		
			this.chkSumaDescuentos.Checked = false;
			this.chkSumaDescuentos.Enabled = false;

			this.chkAdicional.Checked = false;
			this.chkAdicional.Enabled = false;
			this.chkOpcional.Checked = false;
			this.chkOpcional.Enabled = false;
			this.chkCupon.Checked = false;
			this.chkCupon.Enabled = false;


			idBloqueaTransacciones = 0;
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
			if (this.txtpaginas.Text.ToString().Length == 0)
			{
				MessageBox.Show("Ingrese la Pagina del Catalogo.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
				this.cmbTipo.Focus();
				return;
			}
			if (this.txtnumproductos.Text.ToString().Length == 0)
			{
				MessageBox.Show("Escriba la cantidad de productos para el catalogo.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
				this.txtDescripcion.Focus();
				return;
			}
		
			#endregion Validación

	
			decimal dContado = 0.00m;
			decimal dTarjeta = 0.00m;
			decimal dCredito = 0.00m;
			DateTime dtInicio = (DateTime)this.dtInicio.Value;
			DateTime dtFin = (DateTime)this.dtFin.Value;

			//			DateTime fecha_ini = Convert.ToDateTime(this.txtHinicio.Text);

using (SqlConnection oConexion = new SqlConnection("Password=Bl45o6$9;Persist Security Info=True;User ID= " + MenuLatinium.stUsuario + ";Initial Catalog=" + MenuLatinium.stDirFacturacion + ";Data Source=" + MenuLatinium.stDirServidor))
					{
						oConexion.Open();

						SqlTransaction oTransaction = null;
						SqlCommand oCmdActualiza = oConexion.CreateCommand();
						oCmdActualiza.Connection = oConexion;
						oCmdActualiza.CommandType = CommandType.Text;
						oCmdActualiza.CommandTimeout = 0;
						
						try
						{
							oTransaction = oConexion.BeginTransaction();
							oCmdActualiza.Transaction = oTransaction;
			#region DetalleHojaCatalogo
							
								foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridDetalle.Rows.All)
								{				
									int idArticulo = 0;
									string Articulo= "";
									string codigo= "";
									int cantidad=0;
									decimal costoUltimo = 0.00m;
									decimal utilidadEfectivo = 0.00m;
									decimal utilidadTC = 0.00m;
									decimal precio3 =0.00m;
									decimal cuo6 = 0.00m;
									decimal cuo9 = 0.00m;
									decimal cuo12 = 0.00m;
									decimal cuo24 = 0.00m;
									decimal cuc3 = 0.00m;
									decimal cuc6 = 0.00m;
									decimal cuc9 = 0.00m;
									decimal cuc12 = 0.00m;
									decimal cuc18 = 0.00m;
									decimal npvp = 0.00m;

									if (dr.Cells["idArticulo"].Value != System.DBNull.Value) idArticulo = (int)dr.Cells["idArticulo"].Value;
									if (dr.Cells["Articulo"].Value != System.DBNull.Value) Articulo = dr.Cells["Articulo"].Value.ToString();
									if (dr.Cells["Codigo"].Value != System.DBNull.Value) codigo =	dr.Cells["Codigo"].Value.ToString();		
									if (dr.Cells["Cantidad"].Value != System.DBNull.Value) cantidad =	(int)dr.Cells["Cantidad"].Value;
									if (dr.Cells["CostoUltimo"].Value != System.DBNull.Value) costoUltimo =(decimal)dr.Cells["CostoUltimo"].Value;
									if (dr.Cells["UtilidadEfectivo"].Value != System.DBNull.Value) utilidadEfectivo =(decimal)dr.Cells["UtilidadEfectivo"].Value;
									if (dr.Cells["UtilidadTC"].Value != System.DBNull.Value) utilidadTC =(decimal)dr.Cells["UtilidadTC"].Value;
									if (dr.Cells["Precio3"].Value != System.DBNull.Value) precio3 =(decimal)dr.Cells["Precio3"].Value;
									if (dr.Cells["cuo6"].Value != System.DBNull.Value) cuo6 =(decimal)dr.Cells["cuo6"].Value;
									if (dr.Cells["cuo9"].Value != System.DBNull.Value) cuo9 =(decimal)dr.Cells["cuo9"].Value;
									if (dr.Cells["cuo12"].Value != System.DBNull.Value) cuo12 =(decimal)dr.Cells["cuo12"].Value;
									if (dr.Cells["cuo24"].Value != System.DBNull.Value) cuo24 =(decimal)dr.Cells["cuo24"].Value;
									if (dr.Cells["cuc3"].Value != System.DBNull.Value) cuc3 =(decimal)dr.Cells["cuc3"].Value;
									if (dr.Cells["cuc6"].Value != System.DBNull.Value) cuc6 =(decimal)dr.Cells["cuc6"].Value;
									if (dr.Cells["cuc9"].Value != System.DBNull.Value) cuc9 =(decimal)dr.Cells["cuc9"].Value;
									if (dr.Cells["cuc12"].Value != System.DBNull.Value) cuc12 =(decimal)dr.Cells["cuc12"].Value;
									if (dr.Cells["cuc18"].Value != System.DBNull.Value) cuc18 =(decimal)dr.Cells["cuc18"].Value;
									if (dr.Cells["Nuevopvp"].Value != System.DBNull.Value) npvp =(decimal)dr.Cells["Nuevopvp"].Value;

									string sSQLDetalle = string.Format("Exec GuardaDetalleCatalogo {0}, {1}, {2}, {3}, {4}, '{5}', '{6}', {7}, {8}, {9}, {10}, {11}, {12}, {13}, {14},{15},{16},{17},{18},{19},{20},{21}",
										(int)this.txtidDetCatalogo.Value,(int)this.txtidCatalogo.Value, (int)this.txtpaginas.Value, (int)this.txtnumproductos.Value, //3
										idArticulo, Articulo, codigo, cantidad, //7
										costoUltimo, utilidadEfectivo, utilidadTC, //10
										precio3, cuo6,//12
										cuo9,cuo12,cuo24,cuc3,cuc6,cuc9,cuc12,cuc18,npvp); //14
									oCmdActualiza.CommandText = sSQLDetalle;
									oCmdActualiza.ExecuteNonQuery();
								}
								MessageBox.Show("Pagina registrada correctamente", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);

							
			#endregion DetalleHojaCatalogo	
							oTransaction.Commit();		

			this.cmbTipo.Enabled = false;
			this.txtDescripcion.Enabled = false;
			this.txtValorMinimo.Enabled = false;
			this.txtValorMaximo.Enabled = false;
	
					
			EstadoGrid(false);
			
			this.btGrabar.Enabled = false;
			this.btEditar.Enabled = true;
			this.btNuevo.Enabled = true;
			this.btBuscar.Enabled = true;
			this.btCancelar.Enabled = true;
			this.btFinalizar.Enabled = true;
	
		

			bModoEdicion = false;
			bNuevo = false;
						}
						catch (Exception ex)
						{
							oTransaction.Rollback();

							MessageBox.Show(string.Format("Commit Exception Type : {0} {1}", ex.GetType(), ex.Message), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);					
						}
						finally
						{
							oConexion.Close();
						}
	}
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
					this.chkGrupos.Enabled = false;

					this.chkOpcional.Enabled = false;
					this.chkCupon.Enabled = false;
				}
				else if (this.cmbTipo.Text.ToString() == "PROMOCIÓN")
				{
					this.txtValorMinimo.Enabled = true;
					this.txtValorMaximo.Enabled = true;
					this.uGridGrupos.Enabled = true;
					this.uGridSubgrupos.Enabled = true;
					this.uGridMarcas.Enabled = true;
					this.chkGrupos.Enabled = true;
		
					this.chkSumaDescuentos.Enabled = true;
			
					this.chkAdicional.Enabled = true;					
					this.chkOpcional.Enabled = true;
					this.chkCupon.Enabled = true;
			
				}
				else if (this.cmbTipo.Text.ToString() == "ESPECÍFICO")
				{
					this.txtValorMinimo.Enabled = true;
					this.txtValorMaximo.Enabled = true;
					this.uGridGrupos.Enabled = true;
					this.uGridSubgrupos.Enabled = true;
					this.uGridMarcas.Enabled = true;
	
					this.chkOpcional.Enabled = false;
					this.chkCupon.Enabled = false;
				}
			}
		}

		private void cmbTipo_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (e.KeyChar == 13) this.dtInicio.Focus();
		}

		private void txtDescripcion_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (e.KeyChar == 13) this.txtpaginastotalcatalogo.Focus();
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
				if (bModoEdicion)
				{
					this.chkSubGrupos.Enabled = true;
					bSelSubGrupos = true;
				}				
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
			if (e.Cell.Column.ToString() == "Sel")
			{
				this.uGridGrupos.Update();

				int idPG = 0;

				if (e.Cell.Row.Cells["idPromocionGrupo"].Value != DBNull.Value) idPG = (int)e.Cell.Row.Cells["idPromocionGrupo"].Value; 
							
				string sSQL = string.Format("Exec PromocionesActualizaGrupos {0}, {1}, {2}, {3}, {4}", idPG, (int)this.txtIdPromocion.Value, (int)e.Cell.Row.Cells["idGrupoArticulo"].Value, iNumIngreso, (bool)e.Cell.Row.Cells["Sel"].Value);				  
				idPG = Funcion.iEscalarSQL(cdsSeteoF, sSQL);

				if (e.Cell.Row.Cells["idPromocionGrupo"].Value == DBNull.Value) e.Cell.Row.Cells["idPromocionGrupo"].Value = idPG; 
				else e.Cell.Row.Cells["idPromocionGrupo"].Value = DBNull.Value;
			}
		}
		
		private void uGridSubgrupos_AfterCellUpdate(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
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

		private void uGridSubgrupos_Click(object sender, System.EventArgs e)
		{
			
			if (this.uGridSubgrupos.Rows.Count == 0) return;

			if (this.uGridSubgrupos.ActiveRow.Cells["idPromocionSubGrupo"].Value == System.DBNull.Value) return;

			if ((bool)this.uGridSubgrupos.ActiveRow.Cells["Sel"].Value)
			{
				int idSubGrupo = 0;
				if (this.uGridSubgrupos.ActiveRow != null) 
				idSubGrupo = (int)this.uGridSubgrupos.ActiveRow.Cells["idPromocionSubGrupo"].Value;

				string sSQL = string.Format("Exec ListaMarcasPromociones {0}, {1}, {2}", idSubGrupo, (int)this.txtIdPromocion.Value, (int)this.uGridSubgrupos.ActiveRow.Cells["idSubGrupo"].Value);
				this.uGridMarcas.DataSource = Funcion.dvProcedimiento(cdsSeteoF, sSQL);
				
				this.chkMarcas.Checked = false;
				if (bModoEdicion)
				{
					this.chkMarcas.Enabled = true;
					bSelMarcas = true;

				}
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

//		private void Consulta(int idPromocion)
//		{
//			try
//			{
//				#region Encabezado Campos Individuales
//				string sSQL = string.Format("Exec ConsultaIndividualPromocionesLotesServicios {0}", idPromocion);
//				SqlDataReader drPromocion = Funcion.rEscalarSQL(cdsSeteoF, sSQL, true);
//				drPromocion.Read();
//				if (drPromocion.HasRows)
//				{
//					this.txtIdPromocion.Value = idPromocion;
//					if (drPromocion.GetValue(1) != System.DBNull.Value)this.txtidArticulo.Value = drPromocion.GetInt32(1);
//					if (drPromocion.GetValue(2) != System.DBNull.Value)this.cmbTipo.Text = drPromocion.GetValue(2).ToString();
//					if (drPromocion.GetValue(3) != System.DBNull.Value)this.dtInicio.Value = drPromocion.GetDateTime(3);
//					if (drPromocion.GetValue(4) != System.DBNull.Value)this.dtFin.Value = drPromocion.GetDateTime(4);
//					if (drPromocion.GetValue(5) != System.DBNull.Value)this.txtDescripcion.Text = drPromocion.GetValue(5).ToString();					
//					if (drPromocion.GetValue(6) != System.DBNull.Value)iNumIngreso = drPromocion.GetInt32(6);
//							if (drPromocion.GetValue(8) != System.DBNull.Value)this.chkCupon.Checked = drPromocion.GetBoolean(8);
//						if (drPromocion.GetValue(10) != System.DBNull.Value)this.chkSumaDescuentos.Checked = drPromocion.GetBoolean(10);
//					if (drPromocion.GetValue(11) != System.DBNull.Value)this.chkAdicional.Checked = drPromocion.GetBoolean(11);
//						if (drPromocion.GetValue(13) != System.DBNull.Value)this.chkOpcional.Checked = drPromocion.GetBoolean(13);								
//							}
//				drPromocion.Close();
//				#endregion Encabezado Campos Individuales
//
//				this.uGridCiudades.DataSource = Funcion.dvProcedimiento(cdsSeteoF, string.Format("Exec ListaCiudadesPromocion {0}", (int)this.txtIdPromocion.Value));	
//				//this.uGridBodegas.DataSource = Funcion.dvProcedimiento(cdsSeteoF, string.Format("Exec ListaCiudadesBodegasPromocion 0, 0"));				
//				this.uDiasSemana.DataSource = Funcion.dvProcedimiento(cdsSeteoF, string.Format("Exec DiasHabitaciones_Promociones {0}", (int)this.txtIdPromocion.Value));
//				this.uGridTiposDePrecio.DataSource = Funcion.dvProcedimiento(cdsSeteoF, string.Format("Exec ListaPreciosPromocion {0}", (int)this.txtIdPromocion.Value));
//				this.uGridGrupos.DataSource = Funcion.dvProcedimiento(cdsSeteoF, string.Format("Exec ListaGruposPromocion {0}", (int)this.txtIdPromocion.Value));
//				this.uGridSubgrupos.DataSource = Funcion.dvProcedimiento(cdsSeteoF, string.Format("Exec ListaSubGruposPromociones 0, 0, 0"));//, (int)this.txtIdPromocion.Value));
//				//this.uGridBodegas.DataSource = Funcion.dvProcedimiento(cdsSeteoF, string.Format("Exec ListaCiudadesBodegasPromocion 0, 0"));
//				//this.ugHabitaciones.DataSource = Funcion.dvProcedimiento(cdsSeteoF, string.Format("Exec ListaHabitacionBodegasPromocion 0,0"));
//				this.uGridMarcas.DataSource = Funcion.dvProcedimiento(cdsSeteoF, string.Format("Exec ListaMarcasPromociones 0, 0, 0"));//, (int)this.txtIdPromocion.Value));
//					
//				this.ultraGrid1.DataSource = Funcion.dvProcedimiento(cdsSeteoF, string.Format("Exec BuscaDetallePromocion {0}", (int)this.txtIdPromocion.Value));
//
//				this.btCancelar.Enabled = true;
//				this.btEditar.Enabled = true;
//				this.btFinalizar.Enabled = true;
//
////				if (this.lblFinalizado.Text.ToString() == "ACTIVO") this.btFinalizar.Text = "Finalizar";
////				else if (this.lblFinalizado.Text.ToString() == "FINALIZADO") this.btFinalizar.Text = "Activar";
////			}
//			catch(Exception Exc)
//			{
//				MessageBox.Show(string.Format("Error al Consultar: {0}", Exc), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
//			}
//		}

		private void btBuscar_Click(object sender, System.EventArgs e)
		{
			this.btCancelar_Click(sender, e);

			using (frmBuscaPromociones miBusqueda = new frmBuscaPromociones())
			{
				if (DialogResult.OK == miBusqueda.ShowDialog())
				{
//					this.Consulta((int) miBusqueda.uGridPromociones.ActiveRow.Cells["idPromocion"].Value);
				}
			}
		}

		private void btEditar_Click(object sender, System.EventArgs e)
		{		
		}

		private void ultraGrid1_AfterRowsDeleted(object sender, System.EventArgs e)
		{
			CalculaCuota();
		}
//
//		private void btFinalizar_Click(object sender, System.EventArgs e)
//		{
//			if (Funcion.ValidaBloqueo((int)this.txtIdPromocion.Value, 7, cdsSeteoF)) return;
//
//			if (this.lblFinalizado.Text.ToString() == "ACTIVO") 
//			{
////				Funcion.EjecutaSQL(cdsSeteoF, string.Format("Update CompraPromocion Set Estado = 'FINALIZADO' Where idPromocion = {0}", (int)this.txtIdPromocion.Value));
////				Funcion.EjecutaSQL(cdsSeteoF, string.Format("Update Articulo SET Descontinuado = 1 Where idArticulo = {0}", (int)this.txtidArticulo.Value));
////				
////				this.btFinalizar.Text = "Activar";				
////				this.btFinalizar.Enabled = true;
////			}
////			else
////			{
//        Funcion.EjecutaSQL(cdsSeteoF, string.Format("Update CompraPromocion Set Estado = 'ACTIVO' Where idPromocion = {0}", (int)this.txtIdPromocion.Value));
//				Funcion.EjecutaSQL(cdsSeteoF, string.Format("Update Articulo SET Descontinuado = 0 Where idArticulo = {0}", (int)this.txtidArticulo.Value));
//				
//				this.btFinalizar.Text = "Finalizar";
//				this.btFinalizar.Enabled = true;
////			}
//		}
		
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



		private void ultraGrid1_AfterRowInsert(object sender, Infragistics.Win.UltraWinGrid.RowEventArgs e)
		{
			if (bModoEdicion)
			{
				e.Row.Cells["Cantidad"].Value = 1;
				e.Row.Cells["Principal"].Value = false;		
				e.Row.Cells["BloqueaDescuento"].Value = false;
				e.Row.Cells["ValorMinimo"].Value = 0.00m;
				e.Row.Cells["ValorMaximo"].Value = 0.00m;		
				e.Row.Cells["Baratazos"].Value = false;		
				e.Row.Cells["Combos"].Value = false;
				e.Row.Cells["Promociones"].Value = false;
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

		private void Resumen(int idArticulo)
		{
			string sSQLVentas = string.Format("Exec ResumenCalculoCatalogo {0}", idArticulo);
			this.uGridResumen.DataSource = Funcion.dvProcedimiento(cdsSeteoF, sSQLVentas);			
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
		private void uGridModelos_Click(object sender, System.EventArgs e)
		{
			if ((bool)this.uGridModelos.ActiveRow.Cells["Sel"].Value)
			{
			Resumen((int)this.uGridModelos.ActiveRow.Cells["idArticulo"].Value);
			}
			else if (!(bool)this.uGridModelos.ActiveRow.Cells["Sel"].Value)
			{				
//				this.uGridModelos.DataSource = Funcion.dvProcedimiento(cdsSeteoF, string.Format("Exec [ListaPromocionesModelos] {0}, -1, 0, 0, 0", (int)this.txtIdPromocion.Value));
			
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
				
			}
			else
			{
				if (e.KeyCode == Keys.Escape) this.btCancelar_Click(sender, e);
				if (e.KeyCode == Keys.F12) this.btGrabar_Click(sender, e);				
			}
		}

		private void uGridModelos_AfterCellUpdate(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			if (e.Cell.Column.ToString() == "Sel")
			{
				this.uGridModelos.Update();
				
				int idPM = 0;
				int idPS = 0;
				int idPG = 0;
				int idPMod = 0;
				int idarticulo=0;
				int idSubGrupo1 = 0;
				int idGrupo=0;
				int idMarca =0;
					
				
				if (this.uGridGrupos.ActiveRow != null && (bool)this.uGridGrupos.ActiveRow.Cells["Sel"].Value) idPG = (int)this.uGridGrupos.ActiveRow.Cells["idPromocionGrupo"].Value;

				if (this.uGridSubgrupos.ActiveRow != null && (bool)this.uGridSubgrupos.ActiveRow.Cells["Sel"].Value) idPS = (int)this.uGridSubgrupos.ActiveRow.Cells["idPromocionSubGrupo"].Value;

				if (this.uGridMarcas.ActiveRow != null && (bool)this.uGridMarcas.ActiveRow.Cells["Sel"].Value) idPM = (int)this.uGridMarcas.ActiveRow.Cells["idPromocionMarca"].Value; 

				if (e.Cell.Row.Cells["idPromocionModelo"].Value != DBNull.Value) idPMod = (int)e.Cell.Row.Cells["idPromocionModelo"].Value; 
							
				string sSQL = string.Format("Exec PromocionesActualizaModelos {0}, {1}, {2}, {3}, {4}, {5}", idPMod, (int)this.txtIdPromocion.Value, idPG, idPS, idPM, 
					(int)e.Cell.Row.Cells["idArticulo"].Value, (bool)e.Cell.Row.Cells["Sel"].Value);
				idPMod = Funcion.iEscalarSQL(cdsSeteoF, sSQL);
				idarticulo=(int)this.uGridModelos.ActiveRow.Cells["idArticulo"].Value;
				Resumen(idarticulo);
				
//				//--
//				this.Cursor = Cursors.WaitCursor;
//			
//				this.ultraGrid1.DataSource = null;
//				this.ultraGrid2.DataSource = null;
//				idSubGrupo1 = (int)this.uGridSubgrupos.ActiveRow.Cells["idSubGrupo"].Value;
//				idGrupo = (int)this.uGridGrupos.ActiveRow.Cells["idGrupoArticulo"].Value;
//				idMarca = (int)this.uGridMarcas.ActiveRow.Cells["idMarca"].Value;
//				
//				this.ultraGrid1.DataSource = Funcion.dvProcedimiento(cdsSeteoF, string.Format("Exec ConsultaActualizacionPreciosArticulos '{0}', {1}, {2}, {3}, {4}, {5}", 
//					idSubGrupo1,idMarca,idGrupo,idarticulo, (int)rdgProductoFacturable.EditValue, false));
//				//					this.ultraGrid1.DataSource = FuncionesProcedimientos.dtGeneral(sSQL);
//			
//				DiseñoGrid();
//
//				this.lblContador.Text = this.ultraGrid1.Rows.Count + " ARTICULOS ENCONTRADOS";
//
//				this.Cursor = Cursors.Default;
//
//
//				//				--

				if (e.Cell.Row.Cells["idPromocionModelo"].Value == DBNull.Value) e.Cell.Row.Cells["idPromocionModelo"].Value = idPMod;
				else e.Cell.Row.Cells["idPromocionModelo"].Value = DBNull.Value;
			}
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
					if (this.uGridPlanes.ActiveRow.Cells["idPromocionesPlanes"].Value != System.DBNull.Value) idPromoPlanes = (int)this.uGridPlanes.ActiveRow.Cells["idPromocionesPlanes"].Value;
				if (e.Cell.Row.Cells["MesesDeGracia"].Value != DBNull.Value) MesesDeGracia = (int)e.Cell.Row.Cells["MesesDeGracia"].Value; 
				
				string sSQL = string.Format("Exec ActualizaPromocionesPlazos {0}, {1}, {2}, {3}, {4}, {5}, {6}, {7}", 
					idPromoPlazos, idPromoBancos, idPromoTarjetas, idPromoPlanes, (int)this.txtIdPromocion.Value, (int)e.Cell.Row.Cells["idPlazo"].Value, iNumIngreso, MesesDeGracia);
				idPromoPlanes = Funcion.iEscalarSQL(cdsSeteoF, sSQL);

				if (e.Cell.Row.Cells["idPromocionesPlazos"].Value == DBNull.Value) e.Cell.Row.Cells["idPromocionesPlazos"].Value = idPromoPlazos; 
				else e.Cell.Row.Cells["idPromocionesPlazos"].Value = DBNull.Value;
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
							
				string sSQL = string.Format("Exec PromocionesActualizaCuotas {0}, {1}, {2}, {3}, {4}", 
					idPC, (int)this.txtIdPromocion.Value, (int)e.Cell.Row.Cells["Cuota"].Value, iNumIngreso, (bool)e.Cell.Row.Cells["Sel"].Value);				  
				idPC = Funcion.iEscalarSQL(cdsSeteoF, sSQL);

				if (e.Cell.Row.Cells["idPromocionesCuotas"].Value == DBNull.Value) e.Cell.Row.Cells["idPromocionesCuotas"].Value = idPC; 
				else e.Cell.Row.Cells["idPromocionesCuotas"].Value = DBNull.Value;
			}
		}

		private void label5_Click(object sender, System.EventArgs e)
		{
		
		}

		
			private void ultraGrid1_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
			{
				#region Decimales
				FuncionesProcedimientos.FormatoGrid(e, "CostoUltimo", 2);
				FuncionesProcedimientos.FormatoGrid(e, "Distribuidor", 2);
				FuncionesProcedimientos.FormatoGrid(e, "MargDist", 2);
				FuncionesProcedimientos.FormatoGrid(e, "Contado", 2);
				FuncionesProcedimientos.FormatoGrid(e, "MargCont", 2);
				FuncionesProcedimientos.FormatoGrid(e, "Tarjeta", 2);
				FuncionesProcedimientos.FormatoGrid(e, "MargTarj", 2);
				FuncionesProcedimientos.FormatoGrid(e, "Credito", 2);
				FuncionesProcedimientos.FormatoGrid(e, "MargCred", 2);	
				FuncionesProcedimientos.FormatoGrid(e, "WEB", 2);
				FuncionesProcedimientos.FormatoGrid(e, "MargWEB", 2);	
				#endregion Decimales
			}

//		private void ListaArticulosPrecios()
//		{
//			this.Cursor = Cursors.WaitCursor;
//			
//		int idSubGrupo1 = 0;		 
//
//						  idSubGrupo1 =(int)this.uGridSubgrupos.ActiveRow.Cells["idSubGrupo"].Value;
//
//			this.ultraGrid1.DataSource = null;
//			this.ultraGrid2.DataSource = null;
//
//			string sSQL = string.Format("Exec ConsultaActualizacionPrecios '{0}', {1}, {2}, {3}, {4}", 
//				this.txtBuscar.Text.ToString().Trim(), idSubGrupo1, (int)rdgProductoExistencia.EditValue, (int)rdgProductoFacturable.EditValue, false);
//			this.ultraGrid1.DataSource = FuncionesProcedimientos.dtGeneral(sSQL);
//			
//			DiseñoGrid();
//
//			this.lblContador.Text = this.ultraGrid1.Rows.Count + " ARTICULOS ENCONTRADOS";
//
//			this.Cursor = Cursors.Default;
//		}
		private void ActualizaMovimientos()
		{
			DateTime dtDesde = DateTime.Now.AddDays(-15);
			DateTime dtHasta = DateTime.Now;

			string sSQL = string.Format("Exec CreaListaArticulosActualizacionDePrecios '{0}', '{1}'", dtDesde.ToString("yyyyMMdd"), dtHasta.ToString("yyyyMMdd"));
			Funcion.EjecutaSQL(cdsSeteoF, sSQL);
			
					
		}
	
		private void Actualiza(Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			string sSQLSel = string.Format("Exec ActualizaActPrecios {0}, {1}, {2}, {3}, {4}, {5}, {6}, {7}, {8}, {9},{10}, {11}", 
				(bool) e.Cell.Row.Cells["Sel"].Value, (int) e.Cell.Row.Cells["idActualizacion"].Value, 
				(decimal) e.Cell.Row.Cells["Distribuidor"].Value, (decimal) e.Cell.Row.Cells["MargDist"].Value,
				(decimal) e.Cell.Row.Cells["Contado"].Value, (decimal) e.Cell.Row.Cells["MargCont"].Value,
				(decimal) e.Cell.Row.Cells["Tarjeta"].Value, (decimal) e.Cell.Row.Cells["MargTarj"].Value,
				(decimal) e.Cell.Row.Cells["Credito"].Value, (decimal) e.Cell.Row.Cells["MargCred"].Value,
				(decimal) e.Cell.Row.Cells["WEB"].Value, (decimal) e.Cell.Row.Cells["MargWEB"].Value);
			Funcion.EjecutaSQL(cdsSeteoF, sSQLSel, true);				
		}
		


		private void uGridSubgrupos_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}

		private void uGridGrupos_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}

		private void rdgProductoFacturable_SelectedIndexChanged(object sender, System.EventArgs e)
		{
		
		}

		private void groupBox1_Enter(object sender, System.EventArgs e)
		{
		
		}

		private void btcrearcatalogo_Click(object sender, System.EventArgs e)
		{
			DateTime dtInicio = (DateTime)this.dtInicio.Value;
			DateTime dtFin = (DateTime)this.dtFin.Value;
			string idTipoCatalogo = "0";
			int aux=0;
			if (this.txtpaginastotalcatalogo.Text.ToString().Trim().Length == 0) 
			{
				MessageBox.Show("Ingrese el numero para crear el Catálogo.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
				this.txtpaginastotalcatalogo.Focus();
				return;
			}
			if (this.txtDescripcion.Text.ToString().Trim().Length == 0)
			{
				MessageBox.Show("Ingrese el nombre del Catálogo", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
				this.txtDescripcion.Focus();
				return;
			}
			if (!Validacion.vbTexto(this.txtDescripcion, 3, 120, "Nombre")) return;
//	if (!Validacion.vbComboVacio(this.cmbTipo, "Seleccione un tipo de Catálogo")) return;
		
			idTipoCatalogo =this.cmbTipo.Text;
			if(idTipoCatalogo=="Credito")
			{
			aux=1;
			}
			else if(idTipoCatalogo=="Efectivo")
			{
			aux=2;
			}
			string sSQL = string.Format("Exec CatalogoGuarda {0}, {1}, '{2}', {3}, '{4}', '{5}', '{6}'", 
				(int)this.txtidCatalogo.Value, 
				aux, 
				this.txtDescripcion.Text.ToString(),
				(int)this.txtpaginastotalcatalogo.Value, 
				this.txtestadocatalogo.Text.ToString(),
				Convert.ToDateTime(this.dtInicio.Value).ToString("yyyyMMdd HH:mm"),
				Convert.ToDateTime(this.dtFin.Value).ToString("yyyyMMdd HH:mm"));
			this.txtidCatalogo.Value = Funcion.iEscalarSQL(cdsSeteoF, sSQL);

			MessageBox.Show("Catálogo registrado correctamente", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);

			DialogResult = DialogResult.OK;
			this.btNuevo.Enabled=true;
		}

		private void txtpaginastotalcatalogo_ValueChanged(object sender, System.EventArgs e)
		{
		
		}
		private void txtpaginastotalcatalogo_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			validar.Solonumeros(e);
//			if ((int)this.cmbTipoRuc.Value == 4 || (int)this.cmbTipoRuc.Value == 5) Funcion.CamposNumericos(sender, e);
		}

		private void txtDescripcion_ValueChanged(object sender, System.EventArgs e)
		{
		
		}

		private void btnuevocatalogo_Click(object sender, System.EventArgs e)
		{
			      DateTime dtHora = DateTime.Now;
						this.dtInicio.Value = DateTime.Now.ToShortTimeString();
						this.dtFin.Value = DateTime.Today;
						this.cmbTipo.Enabled = true;
						this.dtInicio.Enabled = true;
						this.dtFin.Enabled = true;
						this.txtDescripcion.Enabled = true;
						this.txtpaginastotalcatalogo.Enabled = true;
						this.txtestadocatalogo.Enabled = false;
						this.txtestadocatalogo.Text = "INICIO";

			EstadoGrid(false);
			this.btcrearcatalogo.Enabled=true;
			this.btNuevo.Enabled = false;
			this.btBuscar.Enabled = false;
			this.btEditar.Enabled = false;
			this.btGrabar.Enabled = false;
			this.btCancelar.Enabled = false;

			bModoEdicion = true;
			bNuevo = true;
						
			
		}

		private void uGridMarcas_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}

		private void uGridModelos_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}

		private void uGridResumen_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}
		private void CargaArticulos(int idArticulo)
		{
			decimal dDescuento = 0.00m;
			int idPromocionD = 0;
			bool bDescuento = false;
			//bool bAdicional = false;
			bool bOpcional = false;
			bool bCupon = false;
			bool bAplicaPromocion = true;
			bool bPromocionVIP = false;
			bool bValidaNumeroIngresos = true;
			int idarticulo=0;
			int idSubGrupo1 = 0;
			int idGrupo=0;
			int idMarca =0;		
			int idcantidad=0;
		
				#region variables
				decimal dDescuentaPorc = 0.00m;
				int iFila = uGridDetalle.Rows.Count - 1;
				iFila++;
				uGridDetalle.DisplayLayout.Bands[0].AddNew();	
				#endregion variables


				#region Declara Conexion
				DataSet oDS = new DataSet();					
				cdsSeteoF.Schema.Connections[0].Open();
				C1.Data.SchemaObjects.Connection miConeccion = cdsSeteoF.Schema.Connections[0];
				miConeccion = cdsSeteoF.Schema.Connections[0];
				#endregion Declara Conexion

				//--Angélica
				this.Cursor = Cursors.WaitCursor;
//				this.ultraGrid2.DataSource = null;
				idSubGrupo1 = (int)this.uGridSubgrupos.ActiveRow.Cells["idSubGrupo"].Value;
				idGrupo = (int)this.uGridGrupos.ActiveRow.Cells["idGrupoArticulo"].Value;
				idMarca = (int)this.uGridMarcas.ActiveRow.Cells["idMarca"].Value;
				idarticulo=(int)this.uGridModelos.ActiveRow.Cells["idArticulo"].Value;
				
				string sSQL = string.Format ("Exec ConsultaActualizacionPreciosArticulos '{0}', {1}, {2}, {3}, {4}, {5}", 
					idSubGrupo1,idMarca,idGrupo,idarticulo, (int)rdgProductoFacturable.EditValue, false);
			SqlDataAdapter oDA = new SqlDataAdapter(sSQL, (SqlConnection) miConeccion.DbConnection);
			oDA.Fill(oDS);
			
				//					this.ultraGrid1.DataSource = FuncionesProcedimientos.dtGeneral(sSQL);
		
			//**--	DiseñoGrid();
				if (iFila < (int)this.txtnumproductos.Value)
				{	
					double costoU1;
					decimal npiva = 0.00m;	
					decimal npvp = 0.00m;
					decimal cciva = 0.00m;		
					
					#region Informacion Productos sin promocion
//					uGridDetalle.Rows[iFila].Cells["idArticulo"].Value = 0;
						uGridDetalle.Rows[iFila].Cells["idArticulo"].Value = (int)this.uGridResumen.ActiveRow.Cells["idArticulo"].Value;
					uGridDetalle.Rows[iFila].Cells["Articulo"].Value = this.uGridResumen.ActiveRow.Cells["Articulo"].Value.ToString();	
					uGridDetalle.Rows[iFila].Cells["Codigo"].Value = this.uGridResumen.ActiveRow.Cells["Codigo"].Value.ToString();
					uGridDetalle.Rows[iFila].Cells["Cantidad"].Value = (int)this.uGridResumen.ActiveRow.Cells["Existencia"].Value;	
					decimal dIVA = Funcion.decEscalarSQL(cdsSeteoF, "Select dbo.RetornaImpuestoIVAVigenteDecimal(GETDATE(), 1)");
					uGridDetalle.Rows[iFila].Cells["CostoUltimo"].Value = Convert.ToDecimal(this.uGridResumen.ActiveRow.Cells["CostoUltimo"].Value) * dIVA;

						uGridDetalle.Rows[iFila].Cells["Precio3"].Value = Convert.ToDecimal(this.uGridResumen.ActiveRow.Cells["Precio3"].Value) * dIVA;
						npiva = Convert.ToDecimal(this.uGridResumen.ActiveRow.Cells["Precio3"].Value) * dIVA;
							cciva= Convert.ToDecimal(this.uGridResumen.ActiveRow.Cells["CostoUltimo"].Value) * dIVA;
						uGridDetalle.Rows[iFila].Cells["UtilidadEfectivo"].Value = (((npiva-cciva) / npiva)- 0.03m) * 100;
						uGridDetalle.Rows[iFila].Cells["UtilidadTC"].Value = (((npiva-(npiva * 0.05m))-cciva) / (npiva-(npiva * (0.05m)))) * 100;
					if(this.cmbTipo.Text=="Efectivo")
					{
						this.uGridDetalle.DisplayLayout.Bands[0].Columns["cuo6"].Hidden = false;	
						this.uGridDetalle.DisplayLayout.Bands[0].Columns["cuo9"].Hidden = false;
						this.uGridDetalle.DisplayLayout.Bands[0].Columns["cuo12"].Hidden = false;
						this.uGridDetalle.DisplayLayout.Bands[0].Columns["cuo24"].Hidden = false;
						this.uGridDetalle.DisplayLayout.Bands[0].Columns["Nuevopvp"].Hidden = true;
						this.uGridDetalle.DisplayLayout.Bands[0].Columns["cuc3"].Hidden = true;
						this.uGridDetalle.DisplayLayout.Bands[0].Columns["cuc6"].Hidden = true;	
						this.uGridDetalle.DisplayLayout.Bands[0].Columns["cuc9"].Hidden = true;	
						this.uGridDetalle.DisplayLayout.Bands[0].Columns["cuc12"].Hidden = true;	
						this.uGridDetalle.DisplayLayout.Bands[0].Columns["cuc18"].Hidden = true;	
						uGridDetalle.Rows[iFila].Cells["cuo6"].Value = (((npiva * 0.0473m)+ npiva) / 6);
						uGridDetalle.Rows[iFila].Cells["cuo9"].Value = (((npiva * 0.0681m)+ npiva) / 9);
						uGridDetalle.Rows[iFila].Cells["cuo12"].Value = (((npiva * 0.0891m)+ npiva) / 12);
						uGridDetalle.Rows[iFila].Cells["cuo24"].Value = (((npiva * 0.1758m)+ npiva) / 24);
					}	
					else if(this.cmbTipo.Text=="Crédito")
					{ 
						npvp=  npiva * Convert.ToDecimal(this.uGridResumen.ActiveRow.Cells["IncrPVP"].Value);
						this.uGridDetalle.DisplayLayout.Bands[0].Columns["Nuevopvp"].Hidden = false;
						this.uGridDetalle.DisplayLayout.Bands[0].Columns["cuc3"].Hidden = false;
						this.uGridDetalle.DisplayLayout.Bands[0].Columns["cuc6"].Hidden = false;	
						this.uGridDetalle.DisplayLayout.Bands[0].Columns["cuc9"].Hidden = false;	
						this.uGridDetalle.DisplayLayout.Bands[0].Columns["cuc12"].Hidden = false;	
						this.uGridDetalle.DisplayLayout.Bands[0].Columns["cuc18"].Hidden = false;	
						this.uGridDetalle.DisplayLayout.Bands[0].Columns["cuo6"].Hidden = true;	
						this.uGridDetalle.DisplayLayout.Bands[0].Columns["cuo9"].Hidden = true;
						this.uGridDetalle.DisplayLayout.Bands[0].Columns["cuo12"].Hidden = true;
						this.uGridDetalle.DisplayLayout.Bands[0].Columns["cuo24"].Hidden = true;
						uGridDetalle.Rows[iFila].Cells["Nuevopvp"].Value = npvp;
						uGridDetalle.Rows[iFila].Cells["cuc3"].Value = (npvp * 1.0270m)/3;
						uGridDetalle.Rows[iFila].Cells["cuc6"].Value = (npvp * 1.0474m)/6;
						uGridDetalle.Rows[iFila].Cells["cuc9"].Value = (npvp * 1.0681m)/9;
						uGridDetalle.Rows[iFila].Cells["cuc12"].Value = (npvp * 1.0890m)/12;
						uGridDetalle.Rows[iFila].Cells["cuc18"].Value = (npvp * 1.1320m)/18;
					}

										#endregion Informacion Productos sin promocion
				
		}
				else
				{
					MessageBox.Show("Verifique las filas excede el numero de articulos de la pagina", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);

				}
				this.lblContador.Text = this.uGridDetalle.Rows.Count + " Articulos Para la Página";

				this.Cursor = Cursors.Default;

		}

		private void ultraGrid1_InitializeLayout_1(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}

		private void uGridArticulos_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}

		private void uGridResumen_ClickCellButton(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			if (e.Cell.Column.ToString() == "UnboundColumn1") 
			{
				if (this.uGridResumen.ActiveRow != null) 
					uGridResumen.ActiveRow.Selected = false;
//				int articulo=0;
//				
//				articulo=(int)this.uGridResumen.ActiveRow.Cells["idArticulo"].Value;
		
				this.CargaArticulos((int)e.Cell.Row.Cells["idArticulo"].Value);
			}
		}

		private void uGridDetalle_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}

		private void cmbTipo_ValueChanged_1(object sender, System.EventArgs e)
		{
		
		}

		private void btExportar_Click(object sender, System.EventArgs e)
		{
			FuncionesProcedimientos.ExportaExcel(this.uGridDetalle);
		}

		private void uGridDetalle_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
				FuncionesProcedimientos.DesplazamientoGrids(sender, e, this.uGridDetalle);	
		}

		


	

	
		}

	


		
}
