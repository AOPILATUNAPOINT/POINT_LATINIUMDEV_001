using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;
using System.Data.SqlClient;
using DevExpress.XtraEditors;
using CrystalDecisions;
using CrystalDecisions.CrystalReports.Engine;
using CrystalDecisions.ReportSource;
using CrystalDecisions.Shared;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de frmAtributosArticulos.
	/// </summary>
	public class frmAtributosArticulos : System.Windows.Forms.Form
	{
		private C1.Data.C1DataSet cdsSeteoF;
		private System.Windows.Forms.Label label1;
		private DevExpress.XtraEditors.SimpleButton btnPulgadas;
		private System.Windows.Forms.Label label8;
		private System.Windows.Forms.Label label6;
		private System.Windows.Forms.Label label5;
		private System.Windows.Forms.Label label4;
		private DevExpress.XtraEditors.SimpleButton btnDiscoDuro;
		private DevExpress.XtraEditors.SimpleButton btnProcesador;
		private DevExpress.XtraEditors.SimpleButton btnMemoriaRam;
		private System.Windows.Forms.Label label2;
		private System.Windows.Forms.Label label3;
		private DevExpress.XtraEditors.SimpleButton btnMarca;
		private DevExpress.XtraEditors.SimpleButton btnColor;
		private System.Windows.Forms.Label label7;
		private DevExpress.XtraEditors.SimpleButton btnGaming;
		private DevExpress.XtraEditors.SimpleButton btnLumenes;
		private System.Windows.Forms.Label label9;
		private System.Windows.Forms.Label label10;
		private DevExpress.XtraEditors.SimpleButton btnResolucion;
		private System.Windows.Forms.Label label11;
		private DevExpress.XtraEditors.SimpleButton btnCapacidad;
		private System.Windows.Forms.Label label12;
		private System.Windows.Forms.Label label13;
		private System.Windows.Forms.Label label14;
		private DevExpress.XtraEditors.SimpleButton btnAlmacenamiento;
		private DevExpress.XtraEditors.SimpleButton btnTarjetaGrafica;
		private System.Windows.Forms.Label label15;
		private System.Windows.Forms.GroupBox groupBox4;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtIdArticulo;
		private System.Windows.Forms.Button btnGuardar;
		private System.Windows.Forms.Button btnSalir;
		private DevExpress.XtraEditors.SimpleButton btnOfertaElectrodomesticos;
		private System.Windows.Forms.Label label16;
		private System.Windows.Forms.Label label17;
		private DevExpress.XtraEditors.SimpleButton btnOfertaComputo;
		private DevExpress.XtraEditors.SimpleButton btnQuemadores;
		private DevExpress.XtraEditors.SimpleButton btnPixelesCamPost;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbPulgadas;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbColor;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbProcesador;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbMarca;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbMemoriaRam;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbDiscoDuro;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbPrecio;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbGaming;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbLumenes;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbResolucion;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbCapacidad;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbQuemadores;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmdPixelesCamPost;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbAlmacenamiento;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbTarjetaGrafica;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbOfertaElectrodomesticos;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbOfertaComputo;
		private DevExpress.XtraEditors.SimpleButton btnPrecio;
		private System.Windows.Forms.Label label18;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtCostoEnvio;
		private System.Windows.Forms.Label label19;
		private System.Windows.Forms.Label label20;
		private System.Windows.Forms.Label label21;
		private System.Windows.Forms.Label label22;
		private System.Windows.Forms.Label label23;
		private System.Windows.Forms.Label label24;
		private System.Windows.Forms.Label label25;
		private System.Windows.Forms.Label label26;
		private DevExpress.XtraEditors.SimpleButton btnSistemaOperativo;
		private DevExpress.XtraEditors.SimpleButton btnPeso;
		private DevExpress.XtraEditors.SimpleButton btnSistema;
		private DevExpress.XtraEditors.SimpleButton btnConectividad;
		private DevExpress.XtraEditors.SimpleButton btnTipoPantalla;
		private DevExpress.XtraEditors.SimpleButton btnComplementos;
		private DevExpress.XtraEditors.SimpleButton btnTipoCarga;
		private DevExpress.XtraEditors.SimpleButton btnPotencia;
		private System.Windows.Forms.Label label27;
		private DevExpress.XtraEditors.SimpleButton btnBateria;
		private System.Windows.Forms.Label label28;
		private DevExpress.XtraEditors.SimpleButton btnPixelesCamFront;
		private System.Windows.Forms.Label label29;
		private DevExpress.XtraEditors.SimpleButton btnTipo;
		private System.Windows.Forms.Label label30;
		private DevExpress.XtraEditors.SimpleButton btnCarga;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbPeso;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbSistema;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbConectividad;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbComplementos;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbTipoCarga;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbPotencia;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbBateria;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbPixelesCamFront;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbTipo;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbCarga;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource1;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbTipoPantalla;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource2;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbSistemaOperativo;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource3;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource4;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource5;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource6;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource7;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource8;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource9;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource10;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource11;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource12;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public frmAtributosArticulos()
		{
			//
			// Necesario para admitir el Diseñador de Windows Forms
			//
			InitializeComponent();

			//
			// TODO: agregar código de constructor después de llamar a InitializeComponent
			//
		}

		int idArticulo = 0;
		public frmAtributosArticulos(int IdArticulo)
		{
			idArticulo = IdArticulo;

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
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmAtributosArticulos));
			Infragistics.Win.Appearance appearance1 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idArtPulgadas");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Pulgadas");
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idArtColor");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Color");
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand3 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idArtProcesador");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Procesador");
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand4 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idArtMarca");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Marca");
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand5 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idArtMemoriaRAM");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("MemoriaRam");
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand6 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn11 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idArtDiscoDuro");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn12 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("DiscoDuro");
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand7 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn13 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idArtPrecio");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn14 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Precio");
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand8 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn15 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idArtGaming");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn16 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Gaming");
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand9 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn17 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idArtLumenes");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn18 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Lumenes");
			Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand10 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn19 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idArtResolucion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn20 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Resolucion");
			Infragistics.Win.Appearance appearance11 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand11 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn21 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idArtCapacidad");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn22 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Capacidad");
			Infragistics.Win.Appearance appearance12 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand12 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn23 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idArtQuemadores");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn24 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Quemadores");
			Infragistics.Win.Appearance appearance13 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand13 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn25 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idArtPixelesPosterior");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn26 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Pixeles");
			Infragistics.Win.Appearance appearance14 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand14 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn27 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idArtAlmacenamiento");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn28 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Almacenamiento");
			Infragistics.Win.Appearance appearance15 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand15 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn29 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idArtTarjetaGrafica");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn30 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("TarjetaGrafica");
			Infragistics.Win.Appearance appearance16 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand16 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn31 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idArtOfertaElectrodomesticos");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn32 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("OfertaElectrodomestico");
			Infragistics.Win.Appearance appearance17 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand17 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn33 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idArtOfertaComputo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn34 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("OfertaComputo");
			Infragistics.Win.Appearance appearance18 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance19 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand18 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn35 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idArtPeso");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn36 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Peso");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idArtPeso");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Peso");
			Infragistics.Win.Appearance appearance20 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand19 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn37 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idArtSistema");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn38 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Sistema");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn3 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idArtSistema");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn4 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Sistema");
			Infragistics.Win.Appearance appearance21 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand20 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn39 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idArtConectividad");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn40 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Conectividad");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn5 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idArtConectividad");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn6 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Conectividad");
			Infragistics.Win.Appearance appearance22 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand21 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn41 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idArtComplementos");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn42 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Complementos");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn7 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idArtComplementos");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn8 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Complementos");
			Infragistics.Win.Appearance appearance23 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand22 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn43 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idArtTipoCarga");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn44 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("TipoCarga");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn9 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idArtTipoCarga");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn10 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("TipoCarga");
			Infragistics.Win.Appearance appearance24 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand23 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn45 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idArtPotencia");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn46 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Potencia");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn11 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idArtPotencia");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn12 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Potencia");
			Infragistics.Win.Appearance appearance25 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand24 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn47 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idArtBateria");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn48 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bateria");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn13 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idArtBateria");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn14 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Bateria");
			Infragistics.Win.Appearance appearance26 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand25 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn49 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idArtPixelesFrontal");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn50 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("PixelesFrontal");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn15 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idArtPixelesFrontal");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn16 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("PixelesFrontal");
			Infragistics.Win.Appearance appearance27 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand26 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn51 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idArtTipo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn52 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Tipo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn17 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idArtTipo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn18 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Tipo");
			Infragistics.Win.Appearance appearance28 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand27 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn53 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idArtCarga");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn54 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Carga");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn19 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idArtCarga");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn20 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Carga");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn21 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idArtTipoPantalla");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn22 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("TipoPantalla");
			Infragistics.Win.Appearance appearance29 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand28 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn55 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idArtTipoPantalla");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn56 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("TipoPantalla");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn23 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idArtSistemaOperativo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn24 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("SistemaOperativo");
			Infragistics.Win.Appearance appearance30 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand29 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn57 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idArtSistemaOperativo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn58 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("SistemaOperativo");
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.label1 = new System.Windows.Forms.Label();
			this.btnPulgadas = new DevExpress.XtraEditors.SimpleButton();
			this.label8 = new System.Windows.Forms.Label();
			this.label6 = new System.Windows.Forms.Label();
			this.label5 = new System.Windows.Forms.Label();
			this.label4 = new System.Windows.Forms.Label();
			this.btnDiscoDuro = new DevExpress.XtraEditors.SimpleButton();
			this.btnProcesador = new DevExpress.XtraEditors.SimpleButton();
			this.btnMemoriaRam = new DevExpress.XtraEditors.SimpleButton();
			this.label2 = new System.Windows.Forms.Label();
			this.btnColor = new DevExpress.XtraEditors.SimpleButton();
			this.label3 = new System.Windows.Forms.Label();
			this.btnMarca = new DevExpress.XtraEditors.SimpleButton();
			this.label7 = new System.Windows.Forms.Label();
			this.btnGaming = new DevExpress.XtraEditors.SimpleButton();
			this.btnLumenes = new DevExpress.XtraEditors.SimpleButton();
			this.label9 = new System.Windows.Forms.Label();
			this.label10 = new System.Windows.Forms.Label();
			this.btnResolucion = new DevExpress.XtraEditors.SimpleButton();
			this.label11 = new System.Windows.Forms.Label();
			this.btnCapacidad = new DevExpress.XtraEditors.SimpleButton();
			this.label12 = new System.Windows.Forms.Label();
			this.label13 = new System.Windows.Forms.Label();
			this.label14 = new System.Windows.Forms.Label();
			this.btnAlmacenamiento = new DevExpress.XtraEditors.SimpleButton();
			this.btnTarjetaGrafica = new DevExpress.XtraEditors.SimpleButton();
			this.label15 = new System.Windows.Forms.Label();
			this.groupBox4 = new System.Windows.Forms.GroupBox();
			this.txtIdArticulo = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.btnGuardar = new System.Windows.Forms.Button();
			this.btnSalir = new System.Windows.Forms.Button();
			this.btnOfertaElectrodomesticos = new DevExpress.XtraEditors.SimpleButton();
			this.label16 = new System.Windows.Forms.Label();
			this.label17 = new System.Windows.Forms.Label();
			this.btnOfertaComputo = new DevExpress.XtraEditors.SimpleButton();
			this.btnQuemadores = new DevExpress.XtraEditors.SimpleButton();
			this.btnPixelesCamPost = new DevExpress.XtraEditors.SimpleButton();
			this.cmbPulgadas = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.cmbColor = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.cmbProcesador = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.cmbMarca = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.cmbMemoriaRam = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.cmbDiscoDuro = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.cmbPrecio = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.cmbGaming = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.cmbLumenes = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.cmbResolucion = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.cmbCapacidad = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.cmbQuemadores = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.cmdPixelesCamPost = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.cmbAlmacenamiento = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.cmbTarjetaGrafica = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.cmbOfertaElectrodomesticos = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.cmbOfertaComputo = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.btnPrecio = new DevExpress.XtraEditors.SimpleButton();
			this.label18 = new System.Windows.Forms.Label();
			this.txtCostoEnvio = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.btnSistemaOperativo = new DevExpress.XtraEditors.SimpleButton();
			this.cmbPeso = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.ultraDataSource5 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.cmbSistema = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.ultraDataSource4 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.cmbConectividad = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.ultraDataSource3 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.label19 = new System.Windows.Forms.Label();
			this.btnPeso = new DevExpress.XtraEditors.SimpleButton();
			this.label20 = new System.Windows.Forms.Label();
			this.btnSistema = new DevExpress.XtraEditors.SimpleButton();
			this.label21 = new System.Windows.Forms.Label();
			this.btnConectividad = new DevExpress.XtraEditors.SimpleButton();
			this.label22 = new System.Windows.Forms.Label();
			this.label23 = new System.Windows.Forms.Label();
			this.btnTipoPantalla = new DevExpress.XtraEditors.SimpleButton();
			this.cmbComplementos = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.ultraDataSource11 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.cmbTipoCarga = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.ultraDataSource10 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.cmbPotencia = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.ultraDataSource9 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.label24 = new System.Windows.Forms.Label();
			this.btnComplementos = new DevExpress.XtraEditors.SimpleButton();
			this.label25 = new System.Windows.Forms.Label();
			this.btnTipoCarga = new DevExpress.XtraEditors.SimpleButton();
			this.label26 = new System.Windows.Forms.Label();
			this.btnPotencia = new DevExpress.XtraEditors.SimpleButton();
			this.cmbBateria = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.ultraDataSource12 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.label27 = new System.Windows.Forms.Label();
			this.btnBateria = new DevExpress.XtraEditors.SimpleButton();
			this.cmbPixelesCamFront = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.ultraDataSource8 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.cmbTipo = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.ultraDataSource7 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.cmbCarga = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.ultraDataSource6 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.label28 = new System.Windows.Forms.Label();
			this.btnPixelesCamFront = new DevExpress.XtraEditors.SimpleButton();
			this.label29 = new System.Windows.Forms.Label();
			this.btnTipo = new DevExpress.XtraEditors.SimpleButton();
			this.label30 = new System.Windows.Forms.Label();
			this.btnCarga = new DevExpress.XtraEditors.SimpleButton();
			this.ultraDataSource1 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.cmbTipoPantalla = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.ultraDataSource2 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.cmbSistemaOperativo = new Infragistics.Win.UltraWinGrid.UltraCombo();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdArticulo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbPulgadas)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbColor)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbProcesador)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbMarca)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbMemoriaRam)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbDiscoDuro)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbPrecio)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbGaming)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbLumenes)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbResolucion)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCapacidad)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbQuemadores)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmdPixelesCamPost)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbAlmacenamiento)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbTarjetaGrafica)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbOfertaElectrodomesticos)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbOfertaComputo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCostoEnvio)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbPeso)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource5)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbSistema)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource4)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbConectividad)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource3)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbComplementos)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource11)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbTipoCarga)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource10)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbPotencia)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource9)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbBateria)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource12)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbPixelesCamFront)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource8)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbTipo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource7)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCarga)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource6)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbTipoPantalla)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbSistemaOperativo)).BeginInit();
			this.SuspendLayout();
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
			// label1
			// 
			this.label1.AutoSize = true;
			this.label1.BackColor = System.Drawing.Color.Transparent;
			this.label1.Location = new System.Drawing.Point(8, 16);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(51, 16);
			this.label1.TabIndex = 288;
			this.label1.Text = "Pulgadas";
			this.label1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// btnPulgadas
			// 
			this.btnPulgadas.CausesValidation = false;
			this.btnPulgadas.Image = ((System.Drawing.Image)(resources.GetObject("btnPulgadas.Image")));
			this.btnPulgadas.Location = new System.Drawing.Point(88, 13);
			this.btnPulgadas.Name = "btnPulgadas";
			this.btnPulgadas.Size = new System.Drawing.Size(24, 22);
			this.btnPulgadas.TabIndex = 290;
			this.btnPulgadas.ToolTip = "Grupo de artículos";
			this.btnPulgadas.Visible = false;
			this.btnPulgadas.Click += new System.EventHandler(this.btnPulgadas_Click);
			// 
			// label8
			// 
			this.label8.AutoSize = true;
			this.label8.BackColor = System.Drawing.Color.Transparent;
			this.label8.Location = new System.Drawing.Point(8, 210);
			this.label8.Name = "label8";
			this.label8.Size = new System.Drawing.Size(36, 16);
			this.label8.TabIndex = 310;
			this.label8.Text = "Precio";
			this.label8.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label6
			// 
			this.label6.AutoSize = true;
			this.label6.BackColor = System.Drawing.Color.Transparent;
			this.label6.Location = new System.Drawing.Point(8, 178);
			this.label6.Name = "label6";
			this.label6.Size = new System.Drawing.Size(60, 16);
			this.label6.TabIndex = 309;
			this.label6.Text = "Disco Duro";
			this.label6.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label5
			// 
			this.label5.AutoSize = true;
			this.label5.BackColor = System.Drawing.Color.Transparent;
			this.label5.Location = new System.Drawing.Point(8, 146);
			this.label5.Name = "label5";
			this.label5.Size = new System.Drawing.Size(76, 16);
			this.label5.TabIndex = 308;
			this.label5.Text = "Memoria RAM";
			this.label5.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label4
			// 
			this.label4.AutoSize = true;
			this.label4.BackColor = System.Drawing.Color.Transparent;
			this.label4.Location = new System.Drawing.Point(8, 82);
			this.label4.Name = "label4";
			this.label4.Size = new System.Drawing.Size(62, 16);
			this.label4.TabIndex = 307;
			this.label4.Text = "Procesador";
			this.label4.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// btnDiscoDuro
			// 
			this.btnDiscoDuro.CausesValidation = false;
			this.btnDiscoDuro.Image = ((System.Drawing.Image)(resources.GetObject("btnDiscoDuro.Image")));
			this.btnDiscoDuro.Location = new System.Drawing.Point(88, 175);
			this.btnDiscoDuro.Name = "btnDiscoDuro";
			this.btnDiscoDuro.Size = new System.Drawing.Size(24, 22);
			this.btnDiscoDuro.TabIndex = 305;
			this.btnDiscoDuro.ToolTip = "Grupo de artículos";
			this.btnDiscoDuro.Visible = false;
			this.btnDiscoDuro.Click += new System.EventHandler(this.btnDiscoDuro_Click);
			// 
			// btnProcesador
			// 
			this.btnProcesador.CausesValidation = false;
			this.btnProcesador.Image = ((System.Drawing.Image)(resources.GetObject("btnProcesador.Image")));
			this.btnProcesador.Location = new System.Drawing.Point(88, 79);
			this.btnProcesador.Name = "btnProcesador";
			this.btnProcesador.Size = new System.Drawing.Size(24, 22);
			this.btnProcesador.TabIndex = 303;
			this.btnProcesador.ToolTip = "Grupo de artículos";
			this.btnProcesador.Visible = false;
			this.btnProcesador.Click += new System.EventHandler(this.btnProcesador_Click);
			// 
			// btnMemoriaRam
			// 
			this.btnMemoriaRam.CausesValidation = false;
			this.btnMemoriaRam.Image = ((System.Drawing.Image)(resources.GetObject("btnMemoriaRam.Image")));
			this.btnMemoriaRam.Location = new System.Drawing.Point(88, 143);
			this.btnMemoriaRam.Name = "btnMemoriaRam";
			this.btnMemoriaRam.Size = new System.Drawing.Size(24, 22);
			this.btnMemoriaRam.TabIndex = 304;
			this.btnMemoriaRam.ToolTip = "Grupo de artículos";
			this.btnMemoriaRam.Visible = false;
			this.btnMemoriaRam.Click += new System.EventHandler(this.btnMemoriaRam_Click);
			// 
			// label2
			// 
			this.label2.AutoSize = true;
			this.label2.BackColor = System.Drawing.Color.Transparent;
			this.label2.Location = new System.Drawing.Point(8, 50);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(31, 16);
			this.label2.TabIndex = 316;
			this.label2.Text = "Color";
			this.label2.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// btnColor
			// 
			this.btnColor.CausesValidation = false;
			this.btnColor.Image = ((System.Drawing.Image)(resources.GetObject("btnColor.Image")));
			this.btnColor.Location = new System.Drawing.Point(88, 47);
			this.btnColor.Name = "btnColor";
			this.btnColor.Size = new System.Drawing.Size(24, 22);
			this.btnColor.TabIndex = 315;
			this.btnColor.ToolTip = "Grupo de artículos";
			this.btnColor.Visible = false;
			this.btnColor.Click += new System.EventHandler(this.btnColor_Click);
			// 
			// label3
			// 
			this.label3.AutoSize = true;
			this.label3.BackColor = System.Drawing.Color.Transparent;
			this.label3.Location = new System.Drawing.Point(8, 114);
			this.label3.Name = "label3";
			this.label3.Size = new System.Drawing.Size(36, 16);
			this.label3.TabIndex = 319;
			this.label3.Text = "Marca";
			this.label3.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// btnMarca
			// 
			this.btnMarca.CausesValidation = false;
			this.btnMarca.Image = ((System.Drawing.Image)(resources.GetObject("btnMarca.Image")));
			this.btnMarca.Location = new System.Drawing.Point(88, 111);
			this.btnMarca.Name = "btnMarca";
			this.btnMarca.Size = new System.Drawing.Size(24, 22);
			this.btnMarca.TabIndex = 318;
			this.btnMarca.ToolTip = "Grupo de artículos";
			this.btnMarca.Visible = false;
			this.btnMarca.Click += new System.EventHandler(this.btnMarca_Click);
			// 
			// label7
			// 
			this.label7.AutoSize = true;
			this.label7.BackColor = System.Drawing.Color.Transparent;
			this.label7.Location = new System.Drawing.Point(8, 242);
			this.label7.Name = "label7";
			this.label7.Size = new System.Drawing.Size(44, 16);
			this.label7.TabIndex = 322;
			this.label7.Text = "Gaming";
			this.label7.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// btnGaming
			// 
			this.btnGaming.CausesValidation = false;
			this.btnGaming.Image = ((System.Drawing.Image)(resources.GetObject("btnGaming.Image")));
			this.btnGaming.Location = new System.Drawing.Point(88, 239);
			this.btnGaming.Name = "btnGaming";
			this.btnGaming.Size = new System.Drawing.Size(24, 22);
			this.btnGaming.TabIndex = 321;
			this.btnGaming.ToolTip = "Grupo de artículos";
			this.btnGaming.Visible = false;
			this.btnGaming.Click += new System.EventHandler(this.btnGaming_Click);
			// 
			// btnLumenes
			// 
			this.btnLumenes.CausesValidation = false;
			this.btnLumenes.Image = ((System.Drawing.Image)(resources.GetObject("btnLumenes.Image")));
			this.btnLumenes.Location = new System.Drawing.Point(88, 271);
			this.btnLumenes.Name = "btnLumenes";
			this.btnLumenes.Size = new System.Drawing.Size(24, 22);
			this.btnLumenes.TabIndex = 324;
			this.btnLumenes.ToolTip = "Grupo de artículos";
			this.btnLumenes.Visible = false;
			this.btnLumenes.Click += new System.EventHandler(this.btnLumenes_Click);
			// 
			// label9
			// 
			this.label9.AutoSize = true;
			this.label9.BackColor = System.Drawing.Color.Transparent;
			this.label9.Location = new System.Drawing.Point(8, 274);
			this.label9.Name = "label9";
			this.label9.Size = new System.Drawing.Size(51, 16);
			this.label9.TabIndex = 326;
			this.label9.Text = "Lumenes";
			this.label9.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label10
			// 
			this.label10.AutoSize = true;
			this.label10.BackColor = System.Drawing.Color.Transparent;
			this.label10.Location = new System.Drawing.Point(8, 307);
			this.label10.Name = "label10";
			this.label10.Size = new System.Drawing.Size(60, 16);
			this.label10.TabIndex = 329;
			this.label10.Text = "Resolución";
			this.label10.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// btnResolucion
			// 
			this.btnResolucion.CausesValidation = false;
			this.btnResolucion.Image = ((System.Drawing.Image)(resources.GetObject("btnResolucion.Image")));
			this.btnResolucion.Location = new System.Drawing.Point(88, 304);
			this.btnResolucion.Name = "btnResolucion";
			this.btnResolucion.Size = new System.Drawing.Size(24, 22);
			this.btnResolucion.TabIndex = 327;
			this.btnResolucion.ToolTip = "Grupo de artículos";
			this.btnResolucion.Visible = false;
			this.btnResolucion.Click += new System.EventHandler(this.btnResolucion_Click);
			// 
			// label11
			// 
			this.label11.AutoSize = true;
			this.label11.BackColor = System.Drawing.Color.Transparent;
			this.label11.Location = new System.Drawing.Point(328, 16);
			this.label11.Name = "label11";
			this.label11.Size = new System.Drawing.Size(58, 16);
			this.label11.TabIndex = 332;
			this.label11.Text = "Capacidad";
			this.label11.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// btnCapacidad
			// 
			this.btnCapacidad.CausesValidation = false;
			this.btnCapacidad.Image = ((System.Drawing.Image)(resources.GetObject("btnCapacidad.Image")));
			this.btnCapacidad.Location = new System.Drawing.Point(456, 13);
			this.btnCapacidad.Name = "btnCapacidad";
			this.btnCapacidad.Size = new System.Drawing.Size(24, 22);
			this.btnCapacidad.TabIndex = 330;
			this.btnCapacidad.ToolTip = "Grupo de artículos";
			this.btnCapacidad.Visible = false;
			this.btnCapacidad.Click += new System.EventHandler(this.btnCapacidad_Click);
			// 
			// label12
			// 
			this.label12.AutoSize = true;
			this.label12.BackColor = System.Drawing.Color.Transparent;
			this.label12.Location = new System.Drawing.Point(328, 50);
			this.label12.Name = "label12";
			this.label12.Size = new System.Drawing.Size(85, 16);
			this.label12.TabIndex = 334;
			this.label12.Text = "N° Quemadores";
			this.label12.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label13
			// 
			this.label13.AutoSize = true;
			this.label13.BackColor = System.Drawing.Color.Transparent;
			this.label13.Location = new System.Drawing.Point(328, 82);
			this.label13.Name = "label13";
			this.label13.Size = new System.Drawing.Size(112, 16);
			this.label13.TabIndex = 336;
			this.label13.Text = "Px. Camara Posterior";
			this.label13.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label14
			// 
			this.label14.AutoSize = true;
			this.label14.BackColor = System.Drawing.Color.Transparent;
			this.label14.Location = new System.Drawing.Point(328, 114);
			this.label14.Name = "label14";
			this.label14.Size = new System.Drawing.Size(89, 16);
			this.label14.TabIndex = 338;
			this.label14.Text = "Almacenamiento";
			this.label14.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// btnAlmacenamiento
			// 
			this.btnAlmacenamiento.CausesValidation = false;
			this.btnAlmacenamiento.Image = ((System.Drawing.Image)(resources.GetObject("btnAlmacenamiento.Image")));
			this.btnAlmacenamiento.Location = new System.Drawing.Point(456, 111);
			this.btnAlmacenamiento.Name = "btnAlmacenamiento";
			this.btnAlmacenamiento.Size = new System.Drawing.Size(24, 22);
			this.btnAlmacenamiento.TabIndex = 339;
			this.btnAlmacenamiento.ToolTip = "Grupo de artículos";
			this.btnAlmacenamiento.Visible = false;
			this.btnAlmacenamiento.Click += new System.EventHandler(this.btnAlmacenamiento_Click);
			// 
			// btnTarjetaGrafica
			// 
			this.btnTarjetaGrafica.CausesValidation = false;
			this.btnTarjetaGrafica.Image = ((System.Drawing.Image)(resources.GetObject("btnTarjetaGrafica.Image")));
			this.btnTarjetaGrafica.Location = new System.Drawing.Point(456, 143);
			this.btnTarjetaGrafica.Name = "btnTarjetaGrafica";
			this.btnTarjetaGrafica.Size = new System.Drawing.Size(24, 22);
			this.btnTarjetaGrafica.TabIndex = 342;
			this.btnTarjetaGrafica.ToolTip = "Grupo de artículos";
			this.btnTarjetaGrafica.Visible = false;
			this.btnTarjetaGrafica.Click += new System.EventHandler(this.btnTarjetaGrafica_Click);
			// 
			// label15
			// 
			this.label15.AutoSize = true;
			this.label15.BackColor = System.Drawing.Color.Transparent;
			this.label15.Location = new System.Drawing.Point(328, 146);
			this.label15.Name = "label15";
			this.label15.Size = new System.Drawing.Size(79, 16);
			this.label15.TabIndex = 341;
			this.label15.Text = "Tarjeta Gráfica";
			this.label15.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// groupBox4
			// 
			this.groupBox4.Location = new System.Drawing.Point(-184, 464);
			this.groupBox4.Name = "groupBox4";
			this.groupBox4.Size = new System.Drawing.Size(1088, 8);
			this.groupBox4.TabIndex = 344;
			this.groupBox4.TabStop = false;
			this.groupBox4.Enter += new System.EventHandler(this.groupBox4_Enter);
			// 
			// txtIdArticulo
			// 
			this.txtIdArticulo.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.txtIdArticulo.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtIdArticulo.Enabled = false;
			this.txtIdArticulo.Location = new System.Drawing.Point(664, 474);
			this.txtIdArticulo.Name = "txtIdArticulo";
			this.txtIdArticulo.PromptChar = ' ';
			this.txtIdArticulo.Size = new System.Drawing.Size(16, 21);
			this.txtIdArticulo.TabIndex = 345;
			this.txtIdArticulo.Visible = false;
			// 
			// btnGuardar
			// 
			this.btnGuardar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnGuardar.Image = ((System.Drawing.Image)(resources.GetObject("btnGuardar.Image")));
			this.btnGuardar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnGuardar.Location = new System.Drawing.Point(8, 473);
			this.btnGuardar.Name = "btnGuardar";
			this.btnGuardar.Size = new System.Drawing.Size(76, 23);
			this.btnGuardar.TabIndex = 346;
			this.btnGuardar.Text = "&Guardar";
			this.btnGuardar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnGuardar.Click += new System.EventHandler(this.btnGuardar_Click);
			// 
			// btnSalir
			// 
			this.btnSalir.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnSalir.CausesValidation = false;
			this.btnSalir.Image = ((System.Drawing.Image)(resources.GetObject("btnSalir.Image")));
			this.btnSalir.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnSalir.Location = new System.Drawing.Point(96, 473);
			this.btnSalir.Name = "btnSalir";
			this.btnSalir.TabIndex = 347;
			this.btnSalir.Text = "&Salir";
			this.btnSalir.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnSalir.Click += new System.EventHandler(this.btnSalir_Click);
			// 
			// btnOfertaElectrodomesticos
			// 
			this.btnOfertaElectrodomesticos.CausesValidation = false;
			this.btnOfertaElectrodomesticos.Enabled = false;
			this.btnOfertaElectrodomesticos.Image = ((System.Drawing.Image)(resources.GetObject("btnOfertaElectrodomesticos.Image")));
			this.btnOfertaElectrodomesticos.Location = new System.Drawing.Point(456, 472);
			this.btnOfertaElectrodomesticos.Name = "btnOfertaElectrodomesticos";
			this.btnOfertaElectrodomesticos.Size = new System.Drawing.Size(24, 22);
			this.btnOfertaElectrodomesticos.TabIndex = 348;
			this.btnOfertaElectrodomesticos.ToolTip = "Grupo de artículos";
			this.btnOfertaElectrodomesticos.Visible = false;
			this.btnOfertaElectrodomesticos.Click += new System.EventHandler(this.btnOfertaElectrodomesticos_Click);
			// 
			// label16
			// 
			this.label16.AutoSize = true;
			this.label16.BackColor = System.Drawing.Color.Transparent;
			this.label16.Location = new System.Drawing.Point(432, 448);
			this.label16.Name = "label16";
			this.label16.Size = new System.Drawing.Size(131, 16);
			this.label16.TabIndex = 350;
			this.label16.Text = "Oferta Electrodomesticos";
			this.label16.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.label16.Visible = false;
			// 
			// label17
			// 
			this.label17.AutoSize = true;
			this.label17.BackColor = System.Drawing.Color.Transparent;
			this.label17.Location = new System.Drawing.Point(264, 480);
			this.label17.Name = "label17";
			this.label17.Size = new System.Drawing.Size(85, 16);
			this.label17.TabIndex = 353;
			this.label17.Text = "Oferta Computo";
			this.label17.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.label17.Visible = false;
			// 
			// btnOfertaComputo
			// 
			this.btnOfertaComputo.CausesValidation = false;
			this.btnOfertaComputo.Image = ((System.Drawing.Image)(resources.GetObject("btnOfertaComputo.Image")));
			this.btnOfertaComputo.Location = new System.Drawing.Point(584, 440);
			this.btnOfertaComputo.Name = "btnOfertaComputo";
			this.btnOfertaComputo.Size = new System.Drawing.Size(24, 22);
			this.btnOfertaComputo.TabIndex = 351;
			this.btnOfertaComputo.ToolTip = "Grupo de artículos";
			this.btnOfertaComputo.Visible = false;
			this.btnOfertaComputo.Click += new System.EventHandler(this.btnOfertaComputo_Click);
			// 
			// btnQuemadores
			// 
			this.btnQuemadores.CausesValidation = false;
			this.btnQuemadores.Image = ((System.Drawing.Image)(resources.GetObject("btnQuemadores.Image")));
			this.btnQuemadores.Location = new System.Drawing.Point(456, 47);
			this.btnQuemadores.Name = "btnQuemadores";
			this.btnQuemadores.Size = new System.Drawing.Size(24, 22);
			this.btnQuemadores.TabIndex = 354;
			this.btnQuemadores.ToolTip = "Grupo de artículos";
			this.btnQuemadores.Visible = false;
			this.btnQuemadores.Click += new System.EventHandler(this.btnQuemadores_Click);
			// 
			// btnPixelesCamPost
			// 
			this.btnPixelesCamPost.CausesValidation = false;
			this.btnPixelesCamPost.Image = ((System.Drawing.Image)(resources.GetObject("btnPixelesCamPost.Image")));
			this.btnPixelesCamPost.Location = new System.Drawing.Point(456, 79);
			this.btnPixelesCamPost.Name = "btnPixelesCamPost";
			this.btnPixelesCamPost.Size = new System.Drawing.Size(24, 22);
			this.btnPixelesCamPost.TabIndex = 356;
			this.btnPixelesCamPost.ToolTip = "Grupo de artículos";
			this.btnPixelesCamPost.Visible = false;
			this.btnPixelesCamPost.Click += new System.EventHandler(this.btnPixelesCamPost_Click);
			// 
			// cmbPulgadas
			// 
			appearance1.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbPulgadas.Appearance = appearance1;
			this.cmbPulgadas.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.cmbPulgadas.Cursor = System.Windows.Forms.Cursors.Default;
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn1.Width = 8;
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridColumn2.Width = 184;
			ultraGridBand1.Columns.AddRange(new object[] {
																										 ultraGridColumn1,
																										 ultraGridColumn2});
			this.cmbPulgadas.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.cmbPulgadas.DisplayMember = "Pulgadas";
			this.cmbPulgadas.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbPulgadas.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbPulgadas.Enabled = false;
			this.cmbPulgadas.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbPulgadas.Location = new System.Drawing.Point(120, 14);
			this.cmbPulgadas.Name = "cmbPulgadas";
			this.cmbPulgadas.Size = new System.Drawing.Size(184, 21);
			this.cmbPulgadas.TabIndex = 358;
			this.cmbPulgadas.ValueMember = "idArtPulgadas";
			// 
			// cmbColor
			// 
			appearance2.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbColor.Appearance = appearance2;
			this.cmbColor.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.cmbColor.Cursor = System.Windows.Forms.Cursors.Default;
			ultraGridColumn3.Header.VisiblePosition = 0;
			ultraGridColumn3.Hidden = true;
			ultraGridColumn3.Width = 8;
			ultraGridColumn4.Header.VisiblePosition = 1;
			ultraGridColumn4.Width = 184;
			ultraGridBand2.Columns.AddRange(new object[] {
																										 ultraGridColumn3,
																										 ultraGridColumn4});
			this.cmbColor.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			this.cmbColor.DisplayMember = "Color";
			this.cmbColor.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbColor.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbColor.Enabled = false;
			this.cmbColor.Location = new System.Drawing.Point(120, 48);
			this.cmbColor.Name = "cmbColor";
			this.cmbColor.Size = new System.Drawing.Size(184, 21);
			this.cmbColor.TabIndex = 360;
			this.cmbColor.ValueMember = "idArtColor";
			// 
			// cmbProcesador
			// 
			appearance3.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbProcesador.Appearance = appearance3;
			this.cmbProcesador.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.cmbProcesador.Cursor = System.Windows.Forms.Cursors.Default;
			ultraGridColumn5.Header.VisiblePosition = 0;
			ultraGridColumn5.Hidden = true;
			ultraGridColumn5.Width = 8;
			ultraGridColumn6.Header.VisiblePosition = 1;
			ultraGridColumn6.Width = 184;
			ultraGridBand3.Columns.AddRange(new object[] {
																										 ultraGridColumn5,
																										 ultraGridColumn6});
			this.cmbProcesador.DisplayLayout.BandsSerializer.Add(ultraGridBand3);
			this.cmbProcesador.DisplayMember = "Procesador";
			this.cmbProcesador.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbProcesador.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbProcesador.Enabled = false;
			this.cmbProcesador.Location = new System.Drawing.Point(120, 80);
			this.cmbProcesador.Name = "cmbProcesador";
			this.cmbProcesador.Size = new System.Drawing.Size(184, 21);
			this.cmbProcesador.TabIndex = 361;
			this.cmbProcesador.ValueMember = "idArtProcesador";
			// 
			// cmbMarca
			// 
			appearance4.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbMarca.Appearance = appearance4;
			this.cmbMarca.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.cmbMarca.Cursor = System.Windows.Forms.Cursors.Default;
			ultraGridColumn7.Header.VisiblePosition = 0;
			ultraGridColumn7.Hidden = true;
			ultraGridColumn7.Width = 8;
			ultraGridColumn8.Header.VisiblePosition = 1;
			ultraGridColumn8.Width = 184;
			ultraGridBand4.Columns.AddRange(new object[] {
																										 ultraGridColumn7,
																										 ultraGridColumn8});
			this.cmbMarca.DisplayLayout.BandsSerializer.Add(ultraGridBand4);
			this.cmbMarca.DisplayMember = "Marca";
			this.cmbMarca.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbMarca.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbMarca.Enabled = false;
			this.cmbMarca.Location = new System.Drawing.Point(120, 112);
			this.cmbMarca.Name = "cmbMarca";
			this.cmbMarca.Size = new System.Drawing.Size(184, 21);
			this.cmbMarca.TabIndex = 362;
			this.cmbMarca.ValueMember = "idArtMarca";
			// 
			// cmbMemoriaRam
			// 
			appearance5.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbMemoriaRam.Appearance = appearance5;
			this.cmbMemoriaRam.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.cmbMemoriaRam.Cursor = System.Windows.Forms.Cursors.Default;
			ultraGridColumn9.Header.VisiblePosition = 0;
			ultraGridColumn9.Hidden = true;
			ultraGridColumn9.Width = 8;
			ultraGridColumn10.Header.VisiblePosition = 1;
			ultraGridColumn10.Width = 184;
			ultraGridBand5.Columns.AddRange(new object[] {
																										 ultraGridColumn9,
																										 ultraGridColumn10});
			this.cmbMemoriaRam.DisplayLayout.BandsSerializer.Add(ultraGridBand5);
			this.cmbMemoriaRam.DisplayMember = "MemoriaRam";
			this.cmbMemoriaRam.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbMemoriaRam.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbMemoriaRam.Enabled = false;
			this.cmbMemoriaRam.Location = new System.Drawing.Point(120, 144);
			this.cmbMemoriaRam.Name = "cmbMemoriaRam";
			this.cmbMemoriaRam.Size = new System.Drawing.Size(184, 21);
			this.cmbMemoriaRam.TabIndex = 363;
			this.cmbMemoriaRam.ValueMember = "idArtMemoriaRAM";
			// 
			// cmbDiscoDuro
			// 
			appearance6.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbDiscoDuro.Appearance = appearance6;
			this.cmbDiscoDuro.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.cmbDiscoDuro.Cursor = System.Windows.Forms.Cursors.Default;
			ultraGridColumn11.Header.VisiblePosition = 0;
			ultraGridColumn11.Hidden = true;
			ultraGridColumn11.UseEditorMaskSettings = true;
			ultraGridColumn11.Width = 8;
			ultraGridColumn12.Header.VisiblePosition = 1;
			ultraGridColumn12.Width = 184;
			ultraGridBand6.Columns.AddRange(new object[] {
																										 ultraGridColumn11,
																										 ultraGridColumn12});
			this.cmbDiscoDuro.DisplayLayout.BandsSerializer.Add(ultraGridBand6);
			this.cmbDiscoDuro.DisplayMember = "DiscoDuro";
			this.cmbDiscoDuro.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbDiscoDuro.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbDiscoDuro.Enabled = false;
			this.cmbDiscoDuro.Location = new System.Drawing.Point(120, 176);
			this.cmbDiscoDuro.Name = "cmbDiscoDuro";
			this.cmbDiscoDuro.Size = new System.Drawing.Size(184, 21);
			this.cmbDiscoDuro.TabIndex = 364;
			this.cmbDiscoDuro.ValueMember = "idArtDiscoDuro";
			// 
			// cmbPrecio
			// 
			appearance7.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbPrecio.Appearance = appearance7;
			this.cmbPrecio.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.cmbPrecio.Cursor = System.Windows.Forms.Cursors.Default;
			ultraGridColumn13.Header.VisiblePosition = 0;
			ultraGridColumn13.Hidden = true;
			ultraGridColumn14.Header.VisiblePosition = 1;
			ultraGridColumn14.Width = 184;
			ultraGridBand7.Columns.AddRange(new object[] {
																										 ultraGridColumn13,
																										 ultraGridColumn14});
			this.cmbPrecio.DisplayLayout.BandsSerializer.Add(ultraGridBand7);
			this.cmbPrecio.DisplayMember = "Precio";
			this.cmbPrecio.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbPrecio.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbPrecio.Enabled = false;
			this.cmbPrecio.Location = new System.Drawing.Point(120, 208);
			this.cmbPrecio.Name = "cmbPrecio";
			this.cmbPrecio.Size = new System.Drawing.Size(184, 21);
			this.cmbPrecio.TabIndex = 365;
			this.cmbPrecio.ValueMember = "idArtPrecio";
			// 
			// cmbGaming
			// 
			appearance8.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbGaming.Appearance = appearance8;
			this.cmbGaming.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.cmbGaming.Cursor = System.Windows.Forms.Cursors.Default;
			ultraGridColumn15.Header.VisiblePosition = 0;
			ultraGridColumn15.Hidden = true;
			ultraGridColumn16.Header.VisiblePosition = 1;
			ultraGridColumn16.Width = 184;
			ultraGridBand8.Columns.AddRange(new object[] {
																										 ultraGridColumn15,
																										 ultraGridColumn16});
			this.cmbGaming.DisplayLayout.BandsSerializer.Add(ultraGridBand8);
			this.cmbGaming.DisplayMember = "Gaming";
			this.cmbGaming.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbGaming.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbGaming.Enabled = false;
			this.cmbGaming.Location = new System.Drawing.Point(120, 240);
			this.cmbGaming.Name = "cmbGaming";
			this.cmbGaming.Size = new System.Drawing.Size(184, 21);
			this.cmbGaming.TabIndex = 366;
			this.cmbGaming.ValueMember = "idArtGaming";
			// 
			// cmbLumenes
			// 
			appearance9.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbLumenes.Appearance = appearance9;
			this.cmbLumenes.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.cmbLumenes.Cursor = System.Windows.Forms.Cursors.Default;
			ultraGridColumn17.Header.VisiblePosition = 0;
			ultraGridColumn17.Hidden = true;
			ultraGridColumn17.Width = 8;
			ultraGridColumn18.Header.VisiblePosition = 1;
			ultraGridColumn18.Width = 184;
			ultraGridBand9.Columns.AddRange(new object[] {
																										 ultraGridColumn17,
																										 ultraGridColumn18});
			this.cmbLumenes.DisplayLayout.BandsSerializer.Add(ultraGridBand9);
			this.cmbLumenes.DisplayMember = "Lumenes";
			this.cmbLumenes.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbLumenes.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbLumenes.Enabled = false;
			this.cmbLumenes.Location = new System.Drawing.Point(120, 272);
			this.cmbLumenes.Name = "cmbLumenes";
			this.cmbLumenes.Size = new System.Drawing.Size(184, 21);
			this.cmbLumenes.TabIndex = 367;
			this.cmbLumenes.ValueMember = "idArtLumenes";
			// 
			// cmbResolucion
			// 
			appearance10.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbResolucion.Appearance = appearance10;
			this.cmbResolucion.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.cmbResolucion.Cursor = System.Windows.Forms.Cursors.Default;
			ultraGridColumn19.Header.VisiblePosition = 0;
			ultraGridColumn19.Hidden = true;
			ultraGridColumn19.Width = 8;
			ultraGridColumn20.Header.VisiblePosition = 1;
			ultraGridColumn20.Width = 184;
			ultraGridBand10.Columns.AddRange(new object[] {
																											ultraGridColumn19,
																											ultraGridColumn20});
			this.cmbResolucion.DisplayLayout.BandsSerializer.Add(ultraGridBand10);
			this.cmbResolucion.DisplayMember = "Resolucion";
			this.cmbResolucion.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbResolucion.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbResolucion.Enabled = false;
			this.cmbResolucion.Location = new System.Drawing.Point(120, 305);
			this.cmbResolucion.Name = "cmbResolucion";
			this.cmbResolucion.Size = new System.Drawing.Size(184, 21);
			this.cmbResolucion.TabIndex = 368;
			this.cmbResolucion.ValueMember = "idArtResolucion";
			// 
			// cmbCapacidad
			// 
			appearance11.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbCapacidad.Appearance = appearance11;
			this.cmbCapacidad.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.cmbCapacidad.Cursor = System.Windows.Forms.Cursors.Default;
			ultraGridColumn21.Header.VisiblePosition = 0;
			ultraGridColumn21.Hidden = true;
			ultraGridColumn22.Header.VisiblePosition = 1;
			ultraGridColumn22.Width = 184;
			ultraGridBand11.Columns.AddRange(new object[] {
																											ultraGridColumn21,
																											ultraGridColumn22});
			this.cmbCapacidad.DisplayLayout.BandsSerializer.Add(ultraGridBand11);
			this.cmbCapacidad.DisplayMember = "Capacidad";
			this.cmbCapacidad.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbCapacidad.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbCapacidad.Enabled = false;
			this.cmbCapacidad.Location = new System.Drawing.Point(488, 14);
			this.cmbCapacidad.Name = "cmbCapacidad";
			this.cmbCapacidad.Size = new System.Drawing.Size(184, 21);
			this.cmbCapacidad.TabIndex = 369;
			this.cmbCapacidad.ValueMember = "idArtCapacidad";
			// 
			// cmbQuemadores
			// 
			appearance12.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbQuemadores.Appearance = appearance12;
			this.cmbQuemadores.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.cmbQuemadores.Cursor = System.Windows.Forms.Cursors.Default;
			ultraGridColumn23.Header.VisiblePosition = 0;
			ultraGridColumn23.Hidden = true;
			ultraGridColumn23.Width = 8;
			ultraGridColumn24.Header.VisiblePosition = 1;
			ultraGridColumn24.Width = 184;
			ultraGridBand12.Columns.AddRange(new object[] {
																											ultraGridColumn23,
																											ultraGridColumn24});
			this.cmbQuemadores.DisplayLayout.BandsSerializer.Add(ultraGridBand12);
			this.cmbQuemadores.DisplayMember = "Quemadores";
			this.cmbQuemadores.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbQuemadores.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbQuemadores.Enabled = false;
			this.cmbQuemadores.Location = new System.Drawing.Point(488, 48);
			this.cmbQuemadores.Name = "cmbQuemadores";
			this.cmbQuemadores.Size = new System.Drawing.Size(184, 21);
			this.cmbQuemadores.TabIndex = 370;
			this.cmbQuemadores.ValueMember = "idArtQuemadores";
			// 
			// cmdPixelesCamPost
			// 
			appearance13.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmdPixelesCamPost.Appearance = appearance13;
			this.cmdPixelesCamPost.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.cmdPixelesCamPost.Cursor = System.Windows.Forms.Cursors.Default;
			ultraGridColumn25.Header.VisiblePosition = 0;
			ultraGridColumn25.Hidden = true;
			ultraGridColumn26.Header.VisiblePosition = 1;
			ultraGridColumn26.Width = 184;
			ultraGridBand13.Columns.AddRange(new object[] {
																											ultraGridColumn25,
																											ultraGridColumn26});
			this.cmdPixelesCamPost.DisplayLayout.BandsSerializer.Add(ultraGridBand13);
			this.cmdPixelesCamPost.DisplayMember = "Pixeles";
			this.cmdPixelesCamPost.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmdPixelesCamPost.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmdPixelesCamPost.Enabled = false;
			this.cmdPixelesCamPost.Location = new System.Drawing.Point(488, 80);
			this.cmdPixelesCamPost.Name = "cmdPixelesCamPost";
			this.cmdPixelesCamPost.Size = new System.Drawing.Size(184, 21);
			this.cmdPixelesCamPost.TabIndex = 371;
			this.cmdPixelesCamPost.ValueMember = "idArtPixelesPosterior";
			// 
			// cmbAlmacenamiento
			// 
			appearance14.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbAlmacenamiento.Appearance = appearance14;
			this.cmbAlmacenamiento.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.cmbAlmacenamiento.Cursor = System.Windows.Forms.Cursors.Default;
			ultraGridColumn27.Header.VisiblePosition = 0;
			ultraGridColumn27.Hidden = true;
			ultraGridColumn28.Header.VisiblePosition = 1;
			ultraGridColumn28.Width = 184;
			ultraGridBand14.Columns.AddRange(new object[] {
																											ultraGridColumn27,
																											ultraGridColumn28});
			this.cmbAlmacenamiento.DisplayLayout.BandsSerializer.Add(ultraGridBand14);
			this.cmbAlmacenamiento.DisplayMember = "Almacenamiento";
			this.cmbAlmacenamiento.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbAlmacenamiento.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbAlmacenamiento.Enabled = false;
			this.cmbAlmacenamiento.Location = new System.Drawing.Point(488, 112);
			this.cmbAlmacenamiento.Name = "cmbAlmacenamiento";
			this.cmbAlmacenamiento.Size = new System.Drawing.Size(184, 21);
			this.cmbAlmacenamiento.TabIndex = 372;
			this.cmbAlmacenamiento.ValueMember = "idArtAlmacenamiento";
			// 
			// cmbTarjetaGrafica
			// 
			appearance15.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbTarjetaGrafica.Appearance = appearance15;
			this.cmbTarjetaGrafica.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.cmbTarjetaGrafica.Cursor = System.Windows.Forms.Cursors.Default;
			ultraGridColumn29.Header.VisiblePosition = 0;
			ultraGridColumn29.Hidden = true;
			ultraGridColumn30.Header.VisiblePosition = 1;
			ultraGridColumn30.Width = 184;
			ultraGridBand15.Columns.AddRange(new object[] {
																											ultraGridColumn29,
																											ultraGridColumn30});
			this.cmbTarjetaGrafica.DisplayLayout.BandsSerializer.Add(ultraGridBand15);
			this.cmbTarjetaGrafica.DisplayMember = "TarjetaGrafica";
			this.cmbTarjetaGrafica.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbTarjetaGrafica.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbTarjetaGrafica.Enabled = false;
			this.cmbTarjetaGrafica.Location = new System.Drawing.Point(488, 144);
			this.cmbTarjetaGrafica.Name = "cmbTarjetaGrafica";
			this.cmbTarjetaGrafica.Size = new System.Drawing.Size(184, 21);
			this.cmbTarjetaGrafica.TabIndex = 373;
			this.cmbTarjetaGrafica.ValueMember = "idArtTarjetaGrafica";
			// 
			// cmbOfertaElectrodomesticos
			// 
			appearance16.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbOfertaElectrodomesticos.Appearance = appearance16;
			this.cmbOfertaElectrodomesticos.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.cmbOfertaElectrodomesticos.Cursor = System.Windows.Forms.Cursors.Default;
			ultraGridColumn31.Header.VisiblePosition = 0;
			ultraGridColumn31.Hidden = true;
			ultraGridColumn32.Header.VisiblePosition = 1;
			ultraGridColumn32.Width = 184;
			ultraGridBand16.Columns.AddRange(new object[] {
																											ultraGridColumn31,
																											ultraGridColumn32});
			this.cmbOfertaElectrodomesticos.DisplayLayout.BandsSerializer.Add(ultraGridBand16);
			this.cmbOfertaElectrodomesticos.DisplayMember = "OfertaElectrodomestico";
			this.cmbOfertaElectrodomesticos.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbOfertaElectrodomesticos.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbOfertaElectrodomesticos.Enabled = false;
			this.cmbOfertaElectrodomesticos.Location = new System.Drawing.Point(488, 472);
			this.cmbOfertaElectrodomesticos.Name = "cmbOfertaElectrodomesticos";
			this.cmbOfertaElectrodomesticos.Size = new System.Drawing.Size(184, 21);
			this.cmbOfertaElectrodomesticos.TabIndex = 374;
			this.cmbOfertaElectrodomesticos.ValueMember = "idArtOfertaElectrodomesticos";
			this.cmbOfertaElectrodomesticos.Visible = false;
			// 
			// cmbOfertaComputo
			// 
			appearance17.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbOfertaComputo.Appearance = appearance17;
			this.cmbOfertaComputo.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.cmbOfertaComputo.Cursor = System.Windows.Forms.Cursors.Default;
			ultraGridColumn33.Header.VisiblePosition = 0;
			ultraGridColumn33.Hidden = true;
			ultraGridColumn34.Header.VisiblePosition = 1;
			ultraGridColumn34.Width = 184;
			ultraGridBand17.Columns.AddRange(new object[] {
																											ultraGridColumn33,
																											ultraGridColumn34});
			this.cmbOfertaComputo.DisplayLayout.BandsSerializer.Add(ultraGridBand17);
			this.cmbOfertaComputo.DisplayMember = "OfertaComputo";
			this.cmbOfertaComputo.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbOfertaComputo.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbOfertaComputo.Enabled = false;
			this.cmbOfertaComputo.Location = new System.Drawing.Point(344, 440);
			this.cmbOfertaComputo.Name = "cmbOfertaComputo";
			this.cmbOfertaComputo.Size = new System.Drawing.Size(184, 21);
			this.cmbOfertaComputo.TabIndex = 375;
			this.cmbOfertaComputo.ValueMember = "idArtOfertaComputo";
			this.cmbOfertaComputo.Visible = false;
			// 
			// btnPrecio
			// 
			this.btnPrecio.CausesValidation = false;
			this.btnPrecio.Image = ((System.Drawing.Image)(resources.GetObject("btnPrecio.Image")));
			this.btnPrecio.Location = new System.Drawing.Point(88, 208);
			this.btnPrecio.Name = "btnPrecio";
			this.btnPrecio.Size = new System.Drawing.Size(24, 22);
			this.btnPrecio.TabIndex = 376;
			this.btnPrecio.ToolTip = "Grupo de artículos";
			this.btnPrecio.Visible = false;
			this.btnPrecio.Click += new System.EventHandler(this.btnPrecio_Click);
			// 
			// label18
			// 
			this.label18.AutoSize = true;
			this.label18.BackColor = System.Drawing.Color.Transparent;
			this.label18.Location = new System.Drawing.Point(192, 480);
			this.label18.Name = "label18";
			this.label18.Size = new System.Drawing.Size(65, 16);
			this.label18.TabIndex = 377;
			this.label18.Text = "Costo Envio";
			this.label18.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.label18.Visible = false;
			// 
			// txtCostoEnvio
			// 
			appearance18.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtCostoEnvio.Appearance = appearance18;
			this.txtCostoEnvio.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtCostoEnvio.Enabled = false;
			this.txtCostoEnvio.FormatString = "#,##0.00";
			this.txtCostoEnvio.Location = new System.Drawing.Point(336, 440);
			this.txtCostoEnvio.Name = "txtCostoEnvio";
			this.txtCostoEnvio.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtCostoEnvio.PromptChar = ' ';
			this.txtCostoEnvio.Size = new System.Drawing.Size(88, 21);
			this.txtCostoEnvio.TabIndex = 378;
			this.txtCostoEnvio.Visible = false;
			// 
			// btnSistemaOperativo
			// 
			this.btnSistemaOperativo.CausesValidation = false;
			this.btnSistemaOperativo.Image = ((System.Drawing.Image)(resources.GetObject("btnSistemaOperativo.Image")));
			this.btnSistemaOperativo.Location = new System.Drawing.Point(456, 208);
			this.btnSistemaOperativo.Name = "btnSistemaOperativo";
			this.btnSistemaOperativo.Size = new System.Drawing.Size(24, 22);
			this.btnSistemaOperativo.TabIndex = 393;
			this.btnSistemaOperativo.ToolTip = "Grupo de artículos";
			this.btnSistemaOperativo.Visible = false;
			this.btnSistemaOperativo.Click += new System.EventHandler(this.btnSistemaOperativo_Click);
			// 
			// cmbPeso
			// 
			appearance19.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbPeso.Appearance = appearance19;
			this.cmbPeso.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.cmbPeso.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbPeso.DataSource = this.ultraDataSource5;
			ultraGridColumn35.Header.VisiblePosition = 0;
			ultraGridColumn35.Hidden = true;
			ultraGridColumn35.Width = 184;
			ultraGridColumn36.Header.VisiblePosition = 1;
			ultraGridColumn36.Width = 184;
			ultraGridBand18.Columns.AddRange(new object[] {
																											ultraGridColumn35,
																											ultraGridColumn36});
			this.cmbPeso.DisplayLayout.BandsSerializer.Add(ultraGridBand18);
			this.cmbPeso.DisplayMember = "Peso";
			this.cmbPeso.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbPeso.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbPeso.Enabled = false;
			this.cmbPeso.Location = new System.Drawing.Point(488, 304);
			this.cmbPeso.Name = "cmbPeso";
			this.cmbPeso.Size = new System.Drawing.Size(184, 21);
			this.cmbPeso.TabIndex = 392;
			this.cmbPeso.ValueMember = "idArtPeso";
			// 
			// ultraDataSource5
			// 
			ultraDataColumn1.DataType = typeof(int);
			ultraDataColumn1.DefaultValue = 0;
			this.ultraDataSource5.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn1,
																																 ultraDataColumn2});
			// 
			// cmbSistema
			// 
			appearance20.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbSistema.Appearance = appearance20;
			this.cmbSistema.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.cmbSistema.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbSistema.DataSource = this.ultraDataSource4;
			ultraGridColumn37.Header.VisiblePosition = 0;
			ultraGridColumn37.Hidden = true;
			ultraGridColumn37.Width = 184;
			ultraGridColumn38.Header.VisiblePosition = 1;
			ultraGridColumn38.Width = 184;
			ultraGridBand19.Columns.AddRange(new object[] {
																											ultraGridColumn37,
																											ultraGridColumn38});
			this.cmbSistema.DisplayLayout.BandsSerializer.Add(ultraGridBand19);
			this.cmbSistema.DisplayMember = "Sistema";
			this.cmbSistema.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbSistema.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbSistema.Enabled = false;
			this.cmbSistema.Location = new System.Drawing.Point(488, 272);
			this.cmbSistema.Name = "cmbSistema";
			this.cmbSistema.Size = new System.Drawing.Size(184, 21);
			this.cmbSistema.TabIndex = 391;
			this.cmbSistema.ValueMember = "idArtSistema";
			// 
			// ultraDataSource4
			// 
			ultraDataColumn3.DataType = typeof(short);
			ultraDataColumn3.DefaultValue = ((short)(0));
			this.ultraDataSource4.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn3,
																																 ultraDataColumn4});
			// 
			// cmbConectividad
			// 
			appearance21.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbConectividad.Appearance = appearance21;
			this.cmbConectividad.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.cmbConectividad.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbConectividad.DataSource = this.ultraDataSource3;
			ultraGridColumn39.Header.VisiblePosition = 0;
			ultraGridColumn39.Hidden = true;
			ultraGridColumn40.Header.VisiblePosition = 1;
			ultraGridColumn40.Width = 184;
			ultraGridBand20.Columns.AddRange(new object[] {
																											ultraGridColumn39,
																											ultraGridColumn40});
			this.cmbConectividad.DisplayLayout.BandsSerializer.Add(ultraGridBand20);
			this.cmbConectividad.DisplayMember = "Conectividad";
			this.cmbConectividad.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbConectividad.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbConectividad.Enabled = false;
			this.cmbConectividad.Location = new System.Drawing.Point(488, 240);
			this.cmbConectividad.Name = "cmbConectividad";
			this.cmbConectividad.Size = new System.Drawing.Size(184, 21);
			this.cmbConectividad.TabIndex = 390;
			this.cmbConectividad.ValueMember = "idArtConectividad";
			// 
			// ultraDataSource3
			// 
			ultraDataColumn5.DataType = typeof(int);
			ultraDataColumn5.DefaultValue = 0;
			this.ultraDataSource3.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn5,
																																 ultraDataColumn6});
			// 
			// label19
			// 
			this.label19.AutoSize = true;
			this.label19.BackColor = System.Drawing.Color.Transparent;
			this.label19.Location = new System.Drawing.Point(328, 304);
			this.label19.Name = "label19";
			this.label19.Size = new System.Drawing.Size(30, 16);
			this.label19.TabIndex = 387;
			this.label19.Text = "Peso";
			this.label19.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// btnPeso
			// 
			this.btnPeso.CausesValidation = false;
			this.btnPeso.Image = ((System.Drawing.Image)(resources.GetObject("btnPeso.Image")));
			this.btnPeso.Location = new System.Drawing.Point(456, 304);
			this.btnPeso.Name = "btnPeso";
			this.btnPeso.Size = new System.Drawing.Size(24, 22);
			this.btnPeso.TabIndex = 386;
			this.btnPeso.ToolTip = "Grupo de artículos";
			this.btnPeso.Visible = false;
			this.btnPeso.Click += new System.EventHandler(this.btnPeso_Click);
			// 
			// label20
			// 
			this.label20.AutoSize = true;
			this.label20.BackColor = System.Drawing.Color.Transparent;
			this.label20.Location = new System.Drawing.Point(328, 272);
			this.label20.Name = "label20";
			this.label20.Size = new System.Drawing.Size(45, 16);
			this.label20.TabIndex = 385;
			this.label20.Text = "Sistema";
			this.label20.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// btnSistema
			// 
			this.btnSistema.CausesValidation = false;
			this.btnSistema.Image = ((System.Drawing.Image)(resources.GetObject("btnSistema.Image")));
			this.btnSistema.Location = new System.Drawing.Point(456, 272);
			this.btnSistema.Name = "btnSistema";
			this.btnSistema.Size = new System.Drawing.Size(24, 22);
			this.btnSistema.TabIndex = 384;
			this.btnSistema.ToolTip = "Grupo de artículos";
			this.btnSistema.Visible = false;
			this.btnSistema.Click += new System.EventHandler(this.btnSistema_Click);
			// 
			// label21
			// 
			this.label21.AutoSize = true;
			this.label21.BackColor = System.Drawing.Color.Transparent;
			this.label21.Location = new System.Drawing.Point(328, 240);
			this.label21.Name = "label21";
			this.label21.Size = new System.Drawing.Size(70, 16);
			this.label21.TabIndex = 383;
			this.label21.Text = "Conectividad";
			this.label21.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// btnConectividad
			// 
			this.btnConectividad.CausesValidation = false;
			this.btnConectividad.Image = ((System.Drawing.Image)(resources.GetObject("btnConectividad.Image")));
			this.btnConectividad.Location = new System.Drawing.Point(456, 240);
			this.btnConectividad.Name = "btnConectividad";
			this.btnConectividad.Size = new System.Drawing.Size(24, 22);
			this.btnConectividad.TabIndex = 382;
			this.btnConectividad.ToolTip = "Grupo de artículos";
			this.btnConectividad.Visible = false;
			this.btnConectividad.Click += new System.EventHandler(this.btnConectividad_Click);
			// 
			// label22
			// 
			this.label22.AutoSize = true;
			this.label22.BackColor = System.Drawing.Color.Transparent;
			this.label22.Location = new System.Drawing.Point(328, 208);
			this.label22.Name = "label22";
			this.label22.Size = new System.Drawing.Size(97, 16);
			this.label22.TabIndex = 381;
			this.label22.Text = "Sistema Operativo";
			this.label22.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label23
			// 
			this.label23.AutoSize = true;
			this.label23.BackColor = System.Drawing.Color.Transparent;
			this.label23.Location = new System.Drawing.Point(328, 176);
			this.label23.Name = "label23";
			this.label23.Size = new System.Drawing.Size(70, 16);
			this.label23.TabIndex = 380;
			this.label23.Text = "Tipo Pantalla";
			this.label23.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// btnTipoPantalla
			// 
			this.btnTipoPantalla.CausesValidation = false;
			this.btnTipoPantalla.Image = ((System.Drawing.Image)(resources.GetObject("btnTipoPantalla.Image")));
			this.btnTipoPantalla.Location = new System.Drawing.Point(456, 176);
			this.btnTipoPantalla.Name = "btnTipoPantalla";
			this.btnTipoPantalla.Size = new System.Drawing.Size(24, 22);
			this.btnTipoPantalla.TabIndex = 379;
			this.btnTipoPantalla.ToolTip = "Grupo de artículos";
			this.btnTipoPantalla.Visible = false;
			this.btnTipoPantalla.Click += new System.EventHandler(this.btnTipoPantalla_Click);
			// 
			// cmbComplementos
			// 
			appearance22.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbComplementos.Appearance = appearance22;
			this.cmbComplementos.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.cmbComplementos.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbComplementos.DataSource = this.ultraDataSource11;
			ultraGridColumn41.Header.VisiblePosition = 0;
			ultraGridColumn41.Hidden = true;
			ultraGridColumn42.Header.VisiblePosition = 1;
			ultraGridColumn42.UseEditorMaskSettings = true;
			ultraGridColumn42.Width = 184;
			ultraGridBand21.Columns.AddRange(new object[] {
																											ultraGridColumn41,
																											ultraGridColumn42});
			this.cmbComplementos.DisplayLayout.BandsSerializer.Add(ultraGridBand21);
			this.cmbComplementos.DisplayMember = "Complementos";
			this.cmbComplementos.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbComplementos.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbComplementos.Enabled = false;
			this.cmbComplementos.Location = new System.Drawing.Point(120, 400);
			this.cmbComplementos.Name = "cmbComplementos";
			this.cmbComplementos.Size = new System.Drawing.Size(184, 21);
			this.cmbComplementos.TabIndex = 402;
			this.cmbComplementos.ValueMember = "idArtComplementos";
			// 
			// ultraDataSource11
			// 
			ultraDataColumn7.DataType = typeof(int);
			ultraDataColumn7.DefaultValue = 0;
			this.ultraDataSource11.Band.Columns.AddRange(new object[] {
																																	ultraDataColumn7,
																																	ultraDataColumn8});
			// 
			// cmbTipoCarga
			// 
			appearance23.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbTipoCarga.Appearance = appearance23;
			this.cmbTipoCarga.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.cmbTipoCarga.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbTipoCarga.DataSource = this.ultraDataSource10;
			ultraGridColumn43.Header.VisiblePosition = 0;
			ultraGridColumn43.Hidden = true;
			ultraGridColumn43.Width = 184;
			ultraGridColumn44.Header.VisiblePosition = 1;
			ultraGridColumn44.Width = 184;
			ultraGridBand22.Columns.AddRange(new object[] {
																											ultraGridColumn43,
																											ultraGridColumn44});
			this.cmbTipoCarga.DisplayLayout.BandsSerializer.Add(ultraGridBand22);
			this.cmbTipoCarga.DisplayMember = "TipoCarga";
			this.cmbTipoCarga.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbTipoCarga.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbTipoCarga.Enabled = false;
			this.cmbTipoCarga.Location = new System.Drawing.Point(120, 368);
			this.cmbTipoCarga.Name = "cmbTipoCarga";
			this.cmbTipoCarga.Size = new System.Drawing.Size(184, 21);
			this.cmbTipoCarga.TabIndex = 401;
			this.cmbTipoCarga.ValueMember = "idArtTipoCarga";
			this.cmbTipoCarga.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.ultraCombo7_InitializeLayout);
			// 
			// ultraDataSource10
			// 
			ultraDataColumn9.DataType = typeof(int);
			ultraDataColumn9.DefaultValue = 0;
			this.ultraDataSource10.Band.Columns.AddRange(new object[] {
																																	ultraDataColumn9,
																																	ultraDataColumn10});
			// 
			// cmbPotencia
			// 
			appearance24.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbPotencia.Appearance = appearance24;
			this.cmbPotencia.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.cmbPotencia.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbPotencia.DataSource = this.ultraDataSource9;
			ultraGridColumn45.Header.VisiblePosition = 0;
			ultraGridColumn45.Hidden = true;
			ultraGridColumn45.Width = 184;
			ultraGridColumn46.Header.VisiblePosition = 1;
			ultraGridColumn46.Width = 184;
			ultraGridBand23.Columns.AddRange(new object[] {
																											ultraGridColumn45,
																											ultraGridColumn46});
			this.cmbPotencia.DisplayLayout.BandsSerializer.Add(ultraGridBand23);
			this.cmbPotencia.DisplayMember = "Potencia";
			this.cmbPotencia.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbPotencia.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbPotencia.Enabled = false;
			this.cmbPotencia.Location = new System.Drawing.Point(120, 336);
			this.cmbPotencia.Name = "cmbPotencia";
			this.cmbPotencia.Size = new System.Drawing.Size(184, 21);
			this.cmbPotencia.TabIndex = 400;
			this.cmbPotencia.ValueMember = "idArtPotencia";
			// 
			// ultraDataSource9
			// 
			ultraDataColumn11.DataType = typeof(int);
			ultraDataColumn11.DefaultValue = 0;
			this.ultraDataSource9.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn11,
																																 ultraDataColumn12});
			// 
			// label24
			// 
			this.label24.AutoSize = true;
			this.label24.BackColor = System.Drawing.Color.Transparent;
			this.label24.Location = new System.Drawing.Point(8, 400);
			this.label24.Name = "label24";
			this.label24.Size = new System.Drawing.Size(80, 16);
			this.label24.TabIndex = 399;
			this.label24.Text = "Complementos";
			this.label24.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.label24.Click += new System.EventHandler(this.label24_Click);
			// 
			// btnComplementos
			// 
			this.btnComplementos.CausesValidation = false;
			this.btnComplementos.Image = ((System.Drawing.Image)(resources.GetObject("btnComplementos.Image")));
			this.btnComplementos.Location = new System.Drawing.Point(88, 400);
			this.btnComplementos.Name = "btnComplementos";
			this.btnComplementos.Size = new System.Drawing.Size(24, 22);
			this.btnComplementos.TabIndex = 398;
			this.btnComplementos.ToolTip = "Grupo de artículos";
			this.btnComplementos.Visible = false;
			this.btnComplementos.Click += new System.EventHandler(this.btnComplementos_Click);
			// 
			// label25
			// 
			this.label25.AutoSize = true;
			this.label25.BackColor = System.Drawing.Color.Transparent;
			this.label25.Location = new System.Drawing.Point(8, 368);
			this.label25.Name = "label25";
			this.label25.Size = new System.Drawing.Size(60, 16);
			this.label25.TabIndex = 397;
			this.label25.Text = "Tipo Carga";
			this.label25.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// btnTipoCarga
			// 
			this.btnTipoCarga.CausesValidation = false;
			this.btnTipoCarga.Image = ((System.Drawing.Image)(resources.GetObject("btnTipoCarga.Image")));
			this.btnTipoCarga.Location = new System.Drawing.Point(88, 368);
			this.btnTipoCarga.Name = "btnTipoCarga";
			this.btnTipoCarga.Size = new System.Drawing.Size(24, 22);
			this.btnTipoCarga.TabIndex = 396;
			this.btnTipoCarga.ToolTip = "Grupo de artículos";
			this.btnTipoCarga.Visible = false;
			this.btnTipoCarga.Click += new System.EventHandler(this.btnTipoCarga_Click);
			// 
			// label26
			// 
			this.label26.AutoSize = true;
			this.label26.BackColor = System.Drawing.Color.Transparent;
			this.label26.Location = new System.Drawing.Point(8, 336);
			this.label26.Name = "label26";
			this.label26.Size = new System.Drawing.Size(48, 16);
			this.label26.TabIndex = 395;
			this.label26.Text = "Potencia";
			this.label26.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// btnPotencia
			// 
			this.btnPotencia.CausesValidation = false;
			this.btnPotencia.Image = ((System.Drawing.Image)(resources.GetObject("btnPotencia.Image")));
			this.btnPotencia.Location = new System.Drawing.Point(88, 336);
			this.btnPotencia.Name = "btnPotencia";
			this.btnPotencia.Size = new System.Drawing.Size(24, 22);
			this.btnPotencia.TabIndex = 394;
			this.btnPotencia.ToolTip = "Grupo de artículos";
			this.btnPotencia.Visible = false;
			this.btnPotencia.Click += new System.EventHandler(this.btnPotencia_Click);
			// 
			// cmbBateria
			// 
			appearance25.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbBateria.Appearance = appearance25;
			this.cmbBateria.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.cmbBateria.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbBateria.DataSource = this.ultraDataSource12;
			ultraGridColumn47.Header.VisiblePosition = 0;
			ultraGridColumn47.Hidden = true;
			ultraGridColumn48.Header.VisiblePosition = 1;
			ultraGridColumn48.Width = 184;
			ultraGridBand24.Columns.AddRange(new object[] {
																											ultraGridColumn47,
																											ultraGridColumn48});
			this.cmbBateria.DisplayLayout.BandsSerializer.Add(ultraGridBand24);
			this.cmbBateria.DisplayMember = "Bateria";
			this.cmbBateria.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbBateria.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbBateria.Enabled = false;
			this.cmbBateria.Location = new System.Drawing.Point(120, 432);
			this.cmbBateria.Name = "cmbBateria";
			this.cmbBateria.Size = new System.Drawing.Size(184, 21);
			this.cmbBateria.TabIndex = 405;
			this.cmbBateria.ValueMember = "idArtBateria";
			// 
			// ultraDataSource12
			// 
			ultraDataColumn13.DataType = typeof(int);
			ultraDataColumn13.DefaultValue = 0;
			this.ultraDataSource12.Band.Columns.AddRange(new object[] {
																																	ultraDataColumn13,
																																	ultraDataColumn14});
			// 
			// label27
			// 
			this.label27.AutoSize = true;
			this.label27.BackColor = System.Drawing.Color.Transparent;
			this.label27.Location = new System.Drawing.Point(8, 432);
			this.label27.Name = "label27";
			this.label27.Size = new System.Drawing.Size(40, 16);
			this.label27.TabIndex = 404;
			this.label27.Text = "Bateria";
			this.label27.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// btnBateria
			// 
			this.btnBateria.CausesValidation = false;
			this.btnBateria.Image = ((System.Drawing.Image)(resources.GetObject("btnBateria.Image")));
			this.btnBateria.Location = new System.Drawing.Point(88, 432);
			this.btnBateria.Name = "btnBateria";
			this.btnBateria.Size = new System.Drawing.Size(24, 22);
			this.btnBateria.TabIndex = 403;
			this.btnBateria.ToolTip = "Grupo de artículos";
			this.btnBateria.Visible = false;
			this.btnBateria.Click += new System.EventHandler(this.btnBateria_Click);
			// 
			// cmbPixelesCamFront
			// 
			appearance26.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbPixelesCamFront.Appearance = appearance26;
			this.cmbPixelesCamFront.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.cmbPixelesCamFront.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbPixelesCamFront.DataSource = this.ultraDataSource8;
			ultraGridColumn49.Header.VisiblePosition = 0;
			ultraGridColumn49.Hidden = true;
			ultraGridColumn49.Width = 184;
			ultraGridColumn50.Header.VisiblePosition = 1;
			ultraGridColumn50.Width = 184;
			ultraGridBand25.Columns.AddRange(new object[] {
																											ultraGridColumn49,
																											ultraGridColumn50});
			this.cmbPixelesCamFront.DisplayLayout.BandsSerializer.Add(ultraGridBand25);
			this.cmbPixelesCamFront.DisplayMember = "PixelesFrontal";
			this.cmbPixelesCamFront.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbPixelesCamFront.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbPixelesCamFront.Enabled = false;
			this.cmbPixelesCamFront.Location = new System.Drawing.Point(488, 400);
			this.cmbPixelesCamFront.Name = "cmbPixelesCamFront";
			this.cmbPixelesCamFront.Size = new System.Drawing.Size(184, 21);
			this.cmbPixelesCamFront.TabIndex = 414;
			this.cmbPixelesCamFront.ValueMember = "idArtPixelesFrontal";
			// 
			// ultraDataSource8
			// 
			ultraDataColumn15.DataType = typeof(int);
			ultraDataColumn15.DefaultValue = 0;
			this.ultraDataSource8.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn15,
																																 ultraDataColumn16});
			// 
			// cmbTipo
			// 
			appearance27.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbTipo.Appearance = appearance27;
			this.cmbTipo.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.cmbTipo.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbTipo.DataSource = this.ultraDataSource7;
			ultraGridColumn51.Header.VisiblePosition = 0;
			ultraGridColumn51.Hidden = true;
			ultraGridColumn51.Width = 184;
			ultraGridColumn52.Header.VisiblePosition = 1;
			ultraGridColumn52.Width = 184;
			ultraGridBand26.Columns.AddRange(new object[] {
																											ultraGridColumn51,
																											ultraGridColumn52});
			this.cmbTipo.DisplayLayout.BandsSerializer.Add(ultraGridBand26);
			this.cmbTipo.DisplayMember = "Tipo";
			this.cmbTipo.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbTipo.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbTipo.Enabled = false;
			this.cmbTipo.Location = new System.Drawing.Point(488, 368);
			this.cmbTipo.Name = "cmbTipo";
			this.cmbTipo.Size = new System.Drawing.Size(184, 21);
			this.cmbTipo.TabIndex = 413;
			this.cmbTipo.ValueMember = "idArtTipo";
			// 
			// ultraDataSource7
			// 
			ultraDataColumn17.DataType = typeof(int);
			ultraDataColumn17.DefaultValue = 0;
			this.ultraDataSource7.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn17,
																																 ultraDataColumn18});
			// 
			// cmbCarga
			// 
			appearance28.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbCarga.Appearance = appearance28;
			this.cmbCarga.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.cmbCarga.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbCarga.DataSource = this.ultraDataSource6;
			ultraGridColumn53.Header.VisiblePosition = 0;
			ultraGridColumn53.Hidden = true;
			ultraGridColumn54.Header.VisiblePosition = 1;
			ultraGridColumn54.Width = 184;
			ultraGridBand27.Columns.AddRange(new object[] {
																											ultraGridColumn53,
																											ultraGridColumn54});
			this.cmbCarga.DisplayLayout.BandsSerializer.Add(ultraGridBand27);
			this.cmbCarga.DisplayMember = "Carga";
			this.cmbCarga.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbCarga.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbCarga.Enabled = false;
			this.cmbCarga.Location = new System.Drawing.Point(488, 336);
			this.cmbCarga.Name = "cmbCarga";
			this.cmbCarga.Size = new System.Drawing.Size(184, 21);
			this.cmbCarga.TabIndex = 412;
			this.cmbCarga.ValueMember = "idArtCarga";
			// 
			// ultraDataSource6
			// 
			ultraDataColumn19.DataType = typeof(int);
			ultraDataColumn19.DefaultValue = 0;
			this.ultraDataSource6.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn19,
																																 ultraDataColumn20});
			// 
			// label28
			// 
			this.label28.AutoSize = true;
			this.label28.BackColor = System.Drawing.Color.Transparent;
			this.label28.Location = new System.Drawing.Point(328, 400);
			this.label28.Name = "label28";
			this.label28.Size = new System.Drawing.Size(102, 16);
			this.label28.TabIndex = 411;
			this.label28.Text = "Px. Camara Frontal";
			this.label28.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// btnPixelesCamFront
			// 
			this.btnPixelesCamFront.CausesValidation = false;
			this.btnPixelesCamFront.Image = ((System.Drawing.Image)(resources.GetObject("btnPixelesCamFront.Image")));
			this.btnPixelesCamFront.Location = new System.Drawing.Point(456, 400);
			this.btnPixelesCamFront.Name = "btnPixelesCamFront";
			this.btnPixelesCamFront.Size = new System.Drawing.Size(24, 22);
			this.btnPixelesCamFront.TabIndex = 410;
			this.btnPixelesCamFront.ToolTip = "Grupo de artículos";
			this.btnPixelesCamFront.Visible = false;
			this.btnPixelesCamFront.Click += new System.EventHandler(this.btnPixelesCamFront_Click);
			// 
			// label29
			// 
			this.label29.AutoSize = true;
			this.label29.BackColor = System.Drawing.Color.Transparent;
			this.label29.Location = new System.Drawing.Point(328, 368);
			this.label29.Name = "label29";
			this.label29.Size = new System.Drawing.Size(26, 16);
			this.label29.TabIndex = 409;
			this.label29.Text = "Tipo";
			this.label29.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// btnTipo
			// 
			this.btnTipo.CausesValidation = false;
			this.btnTipo.Image = ((System.Drawing.Image)(resources.GetObject("btnTipo.Image")));
			this.btnTipo.Location = new System.Drawing.Point(456, 368);
			this.btnTipo.Name = "btnTipo";
			this.btnTipo.Size = new System.Drawing.Size(24, 22);
			this.btnTipo.TabIndex = 408;
			this.btnTipo.ToolTip = "Grupo de artículos";
			this.btnTipo.Visible = false;
			this.btnTipo.Click += new System.EventHandler(this.btnTipo_Click);
			// 
			// label30
			// 
			this.label30.AutoSize = true;
			this.label30.BackColor = System.Drawing.Color.Transparent;
			this.label30.Location = new System.Drawing.Point(328, 336);
			this.label30.Name = "label30";
			this.label30.Size = new System.Drawing.Size(35, 16);
			this.label30.TabIndex = 407;
			this.label30.Text = "Carga";
			this.label30.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// btnCarga
			// 
			this.btnCarga.CausesValidation = false;
			this.btnCarga.Image = ((System.Drawing.Image)(resources.GetObject("btnCarga.Image")));
			this.btnCarga.Location = new System.Drawing.Point(456, 336);
			this.btnCarga.Name = "btnCarga";
			this.btnCarga.Size = new System.Drawing.Size(24, 22);
			this.btnCarga.TabIndex = 406;
			this.btnCarga.ToolTip = "Grupo de artículos";
			this.btnCarga.Visible = false;
			this.btnCarga.Click += new System.EventHandler(this.btnCarga_Click);
			// 
			// ultraDataSource1
			// 
			ultraDataColumn21.DataType = typeof(int);
			ultraDataColumn21.DefaultValue = 0;
			this.ultraDataSource1.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn21,
																																 ultraDataColumn22});
			// 
			// cmbTipoPantalla
			// 
			appearance29.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbTipoPantalla.Appearance = appearance29;
			this.cmbTipoPantalla.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.cmbTipoPantalla.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbTipoPantalla.DataSource = this.ultraDataSource1;
			ultraGridColumn55.Header.VisiblePosition = 0;
			ultraGridColumn55.Hidden = true;
			ultraGridColumn56.Header.VisiblePosition = 1;
			ultraGridColumn56.Width = 184;
			ultraGridBand28.Columns.AddRange(new object[] {
																											ultraGridColumn55,
																											ultraGridColumn56});
			this.cmbTipoPantalla.DisplayLayout.BandsSerializer.Add(ultraGridBand28);
			this.cmbTipoPantalla.DisplayMember = "TipoPantalla";
			this.cmbTipoPantalla.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbTipoPantalla.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbTipoPantalla.Enabled = false;
			this.cmbTipoPantalla.Location = new System.Drawing.Point(488, 176);
			this.cmbTipoPantalla.Name = "cmbTipoPantalla";
			this.cmbTipoPantalla.Size = new System.Drawing.Size(184, 21);
			this.cmbTipoPantalla.TabIndex = 416;
			this.cmbTipoPantalla.ValueMember = "idArtTipoPantalla";
			// 
			// ultraDataSource2
			// 
			ultraDataColumn23.DataType = typeof(short);
			ultraDataColumn23.DefaultValue = ((short)(0));
			this.ultraDataSource2.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn23,
																																 ultraDataColumn24});
			// 
			// cmbSistemaOperativo
			// 
			appearance30.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbSistemaOperativo.Appearance = appearance30;
			this.cmbSistemaOperativo.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.cmbSistemaOperativo.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbSistemaOperativo.DataSource = this.ultraDataSource2;
			ultraGridColumn57.Header.VisiblePosition = 0;
			ultraGridColumn57.Hidden = true;
			ultraGridColumn58.Header.VisiblePosition = 1;
			ultraGridColumn58.Width = 184;
			ultraGridBand29.Columns.AddRange(new object[] {
																											ultraGridColumn57,
																											ultraGridColumn58});
			this.cmbSistemaOperativo.DisplayLayout.BandsSerializer.Add(ultraGridBand29);
			this.cmbSistemaOperativo.DisplayMember = "SistemaOperativo";
			this.cmbSistemaOperativo.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbSistemaOperativo.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbSistemaOperativo.Enabled = false;
			this.cmbSistemaOperativo.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbSistemaOperativo.Location = new System.Drawing.Point(488, 208);
			this.cmbSistemaOperativo.Name = "cmbSistemaOperativo";
			this.cmbSistemaOperativo.Size = new System.Drawing.Size(184, 21);
			this.cmbSistemaOperativo.TabIndex = 417;
			this.cmbSistemaOperativo.ValueMember = "idArtSistemaOperativo";
			// 
			// frmAtributosArticulos
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(688, 504);
			this.Controls.Add(this.cmbSistemaOperativo);
			this.Controls.Add(this.cmbTipoPantalla);
			this.Controls.Add(this.cmbPixelesCamFront);
			this.Controls.Add(this.cmbTipo);
			this.Controls.Add(this.cmbCarga);
			this.Controls.Add(this.label28);
			this.Controls.Add(this.label29);
			this.Controls.Add(this.label30);
			this.Controls.Add(this.label27);
			this.Controls.Add(this.label24);
			this.Controls.Add(this.label25);
			this.Controls.Add(this.label26);
			this.Controls.Add(this.label19);
			this.Controls.Add(this.label20);
			this.Controls.Add(this.label21);
			this.Controls.Add(this.label22);
			this.Controls.Add(this.label23);
			this.Controls.Add(this.label18);
			this.Controls.Add(this.label17);
			this.Controls.Add(this.label16);
			this.Controls.Add(this.label15);
			this.Controls.Add(this.label14);
			this.Controls.Add(this.label13);
			this.Controls.Add(this.label12);
			this.Controls.Add(this.label11);
			this.Controls.Add(this.label10);
			this.Controls.Add(this.label9);
			this.Controls.Add(this.label7);
			this.Controls.Add(this.label3);
			this.Controls.Add(this.label2);
			this.Controls.Add(this.label8);
			this.Controls.Add(this.label6);
			this.Controls.Add(this.label5);
			this.Controls.Add(this.label4);
			this.Controls.Add(this.label1);
			this.Controls.Add(this.btnPixelesCamFront);
			this.Controls.Add(this.btnTipo);
			this.Controls.Add(this.btnCarga);
			this.Controls.Add(this.cmbBateria);
			this.Controls.Add(this.btnBateria);
			this.Controls.Add(this.cmbComplementos);
			this.Controls.Add(this.cmbTipoCarga);
			this.Controls.Add(this.cmbPotencia);
			this.Controls.Add(this.btnComplementos);
			this.Controls.Add(this.btnTipoCarga);
			this.Controls.Add(this.btnPotencia);
			this.Controls.Add(this.btnSistemaOperativo);
			this.Controls.Add(this.cmbPeso);
			this.Controls.Add(this.cmbSistema);
			this.Controls.Add(this.cmbConectividad);
			this.Controls.Add(this.btnPeso);
			this.Controls.Add(this.btnSistema);
			this.Controls.Add(this.btnConectividad);
			this.Controls.Add(this.btnTipoPantalla);
			this.Controls.Add(this.txtCostoEnvio);
			this.Controls.Add(this.btnPrecio);
			this.Controls.Add(this.cmbOfertaComputo);
			this.Controls.Add(this.cmbOfertaElectrodomesticos);
			this.Controls.Add(this.cmbTarjetaGrafica);
			this.Controls.Add(this.cmbAlmacenamiento);
			this.Controls.Add(this.cmdPixelesCamPost);
			this.Controls.Add(this.cmbQuemadores);
			this.Controls.Add(this.cmbCapacidad);
			this.Controls.Add(this.cmbResolucion);
			this.Controls.Add(this.cmbLumenes);
			this.Controls.Add(this.cmbGaming);
			this.Controls.Add(this.cmbPrecio);
			this.Controls.Add(this.cmbDiscoDuro);
			this.Controls.Add(this.cmbMemoriaRam);
			this.Controls.Add(this.cmbMarca);
			this.Controls.Add(this.cmbProcesador);
			this.Controls.Add(this.cmbColor);
			this.Controls.Add(this.cmbPulgadas);
			this.Controls.Add(this.btnPixelesCamPost);
			this.Controls.Add(this.btnQuemadores);
			this.Controls.Add(this.btnOfertaComputo);
			this.Controls.Add(this.btnOfertaElectrodomesticos);
			this.Controls.Add(this.btnGuardar);
			this.Controls.Add(this.btnSalir);
			this.Controls.Add(this.txtIdArticulo);
			this.Controls.Add(this.groupBox4);
			this.Controls.Add(this.btnTarjetaGrafica);
			this.Controls.Add(this.btnAlmacenamiento);
			this.Controls.Add(this.btnCapacidad);
			this.Controls.Add(this.btnResolucion);
			this.Controls.Add(this.btnLumenes);
			this.Controls.Add(this.btnGaming);
			this.Controls.Add(this.btnMarca);
			this.Controls.Add(this.btnColor);
			this.Controls.Add(this.btnDiscoDuro);
			this.Controls.Add(this.btnProcesador);
			this.Controls.Add(this.btnMemoriaRam);
			this.Controls.Add(this.btnPulgadas);
			this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
			this.MaximizeBox = false;
			this.MinimizeBox = false;
			this.Name = "frmAtributosArticulos";
			this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
			this.Text = "Atributos";
			this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.frmAtributosArticulos_KeyDown);
			this.Load += new System.EventHandler(this.frmAtributosArticulos_Load);
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdArticulo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbPulgadas)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbColor)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbProcesador)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbMarca)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbMemoriaRam)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbDiscoDuro)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbPrecio)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbGaming)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbLumenes)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbResolucion)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCapacidad)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbQuemadores)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmdPixelesCamPost)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbAlmacenamiento)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbTarjetaGrafica)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbOfertaElectrodomesticos)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbOfertaComputo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCostoEnvio)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbPeso)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource5)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbSistema)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource4)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbConectividad)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource3)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbComplementos)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource11)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbTipoCarga)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource10)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbPotencia)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource9)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbBateria)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource12)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbPixelesCamFront)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource8)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbTipo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource7)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCarga)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource6)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbTipoPantalla)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbSistemaOperativo)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);	
		}

		private Acceso miAcceso;

		private void UnloadMe()
		{
			this.Close();
		}

		private void simpleButton4_Click(object sender, System.EventArgs e)
		{
		
		}

		private void frmAtributosArticulos_Load(object sender, System.EventArgs e)
		{
			#region Seguridad
			miAcceso = new Acceso(cdsSeteoF, "060610");

			if (!Funcion.Permiso("1486", cdsSeteoF))
			{				
				MessageBox.Show("No puede ingresar a Articulos Atributos", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				BeginInvoke(new MethodInvoker(UnloadMe));
				return;		
			}

			if (miAcceso.BEditPulgadas) this.btnPulgadas.Visible = true;
			if (miAcceso.BEditColor) this.btnColor.Visible = true;
			if (miAcceso.BEditProcesador) this.btnProcesador.Visible = true;
			if (miAcceso.BEditMarca) this.btnMarca.Visible = true;
			if (miAcceso.BEditMemoria) this.btnMemoriaRam.Visible = true;
			if (miAcceso.BEditDisco) this.btnDiscoDuro.Visible = true;
			if (miAcceso.BEditPrecio) this.btnPrecio.Visible = true;
			if (miAcceso.BEditGaming) this.btnGaming.Visible = true;
			if (miAcceso.BEditLumenes) this.btnLumenes.Visible = true;
			if (miAcceso.BEditResolucion) this.btnResolucion.Visible = true;
			if (miAcceso.BEditCapacidad) this.btnCapacidad.Visible = true;
			if (miAcceso.BEditAlmacenamiento) this.btnAlmacenamiento.Visible = true;
			if (miAcceso.BEditPixeles) this.btnPixelesCamPost.Visible = true;
			if (miAcceso.BEditQuemadores) this.btnQuemadores.Visible = true;
			if (miAcceso.BEditTarjeta) this.btnTarjetaGrafica.Visible = true;
			//if (miAcceso.BEditOfeElectro) this.btnOfertaElectrodomesticos.Visible = true;
			if (miAcceso.BEditOfeComp)
			{
				this.btnTipoPantalla.Visible = true;
				this.btnSistemaOperativo.Visible = true;
				this.btnConectividad.Visible = true;
				this.btnSistema.Visible = true;
				this.btnPeso.Visible = true;
				this.btnPotencia.Visible = true;
				this.btnTipoCarga.Visible = true;
				this.btnCarga.Visible = true;
				this.btnTipo.Visible = true;
				this.btnComplementos.Visible = true;
				this.btnPixelesCamFront.Visible = true;
				this.btnBateria.Visible = true;	
			}
			if (miAcceso.EditarReferencia)
			{
				this.cmbPulgadas.Enabled = true;
				this.cmbColor.Enabled = true;
				this.cmbProcesador.Enabled = true;
				this.cmbMarca.Enabled = true;
				this.cmbMemoriaRam.Enabled = true;
				this.cmbDiscoDuro.Enabled = true;
				this.cmbPrecio.Enabled = true;
				this.cmbGaming.Enabled = true;
				this.cmbLumenes.Enabled = true;
				this.cmbResolucion.Enabled = true;
				this.cmbCapacidad.Enabled = true;
				this.cmbAlmacenamiento.Enabled = true;
				this.cmdPixelesCamPost.Enabled = true;
				this.cmbQuemadores.Enabled = true;
				this.cmbTarjetaGrafica.Enabled = true;
				//this.cmbOfertaElectrodomesticos.Enabled = true;
				//this.cmbOfertaComputo.Enabled = true;
				//this.txtCostoEnvio.Enabled = true;

				this.cmbTipoPantalla.Enabled = true;
				this.cmbSistemaOperativo.Enabled = true;
				this.cmbConectividad.Enabled = true;
				this.cmbSistema.Enabled = true;
				this.cmbPeso.Enabled = true;
				this.cmbPotencia.Enabled = true;
				this.cmbTipoCarga.Enabled = true;
				this.cmbCarga.Enabled = true;
				this.cmbTipo.Enabled = true;
				this.cmbComplementos.Enabled = true;
				this.cmbPixelesCamFront.Enabled = true;
				this.cmbBateria.Enabled = true;
				
			}

			#endregion Seguridad

			this.txtIdArticulo.Value = idArticulo;			
			this.cmbPulgadas.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Select idArtPulgadas, Pulgadas From ArtPulgadas Order By Pulgadas");
			this.cmbColor.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Select idArtColor, Color From ArtColor Order By Color");
			this.cmbProcesador.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Select idArtProcesador, Procesador From ArtProcesador Order By Procesador");
			this.cmbMarca.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Select idArtMarca, Marca From ArtMarca Order By Marca");
			this.cmbMemoriaRam.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Select idArtMemoriaRAM, MemoriaRam From ArtMemoriaRAM Order By MemoriaRam");
			this.cmbDiscoDuro.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Select idArtDiscoDuro, DiscoDuro From ArtDiscoDuro Order By DiscoDuro");
			this.cmbPrecio.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Select idArtPrecio, Precio From ArtPrecio Order By Precio");
			this.cmbGaming.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Select idArtGaming, Gaming From ArtGaming Order By Gaming");
			this.cmbLumenes.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Select idArtLumenes, Lumenes From ArtLumenes Order By Lumenes");
			this.cmbResolucion.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Select idArtResolucion, Resolucion From ArtResolucion Order By Resolucion");
			this.cmbCapacidad.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Select idArtCapacidad, Capacidad From ArtCapacidad Order By Capacidad");
			this.cmbQuemadores.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Select idArtQuemadores, Quemadores From ArtQuemadores Order By Quemadores");
			this.cmdPixelesCamPost.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Select idArtPixelesPosterior, Pixeles From ArtPixelesPosterior Order By Pixeles");
			this.cmbAlmacenamiento.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Select idArtAlmacenamiento, Almacenamiento From ArtAlmacenamiento Order By Almacenamiento");
			this.cmbTarjetaGrafica.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Select idArtTarjetaGrafica, TarjetaGrafica From ArtTarjetaGrafica Order By TarjetaGrafica");
			this.cmbOfertaElectrodomesticos.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Select idArtOfertaElectrodomesticos, OfertaElectrodomestico From ArtOfertaElectrodomesticos Order By OfertaElectrodomestico");
			this.cmbOfertaComputo.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Select idArtOfertaComputo, OfertaComputo From ArtOfertaComputo Order By OfertaComputo");

			this.cmbTipoPantalla.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Select idArtTipoPantalla, TipoPantalla From ArtTipoPantalla Order By TipoPantalla");
			this.cmbSistemaOperativo.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Select idArtSistemaOperativo, SistemaOperativo From ArtSistemaOperativo Order By SistemaOperativo");
			this.cmbConectividad.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Select idArtConectividad, Conectividad From ArtConectividad Order By Conectividad");
			this.cmbSistema.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Select idArtSistema, Sistema From ArtSistema Order By Sistema");
			this.cmbPeso.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Select idArtPeso, Peso From ArtPeso Order By Peso");
			this.cmbPotencia.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Select idArtPotencia, Potencia From ArtPotencia Order By Potencia");
			this.cmbTipoCarga.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Select idArtTipoCarga, TipoCarga From ArtTipoCarga Order By TipoCarga");
			this.cmbCarga.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Select idArtCarga, Carga From ArtCarga Order By Carga");
			this.cmbTipo.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Select idArtTipo, Tipo From ArtTipo Order By Tipo");
			this.cmbComplementos.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Select idArtComplementos, Complementos From ArtComplementos Order By Complementos");
			this.cmbPixelesCamFront.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Select idArtPixelesFrontal, PixelesFrontal From ArtPixelesFrontal Order By PixelesFrontal");
			this.cmbBateria.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Select idArtBateria, Bateria From ArtBateria Order By Bateria");

			string sSQL = string.Format("Exec ConsultaArticulosAtributos {0}", idArticulo);
			SqlDataReader drArticulo = Funcion.rEscalarSQL(cdsSeteoF, sSQL, true);
			drArticulo.Read();
			if (drArticulo.HasRows)
			{	
				this.txtIdArticulo.Value = idArticulo;
				if (drArticulo.GetValue(1) != System.DBNull.Value) this.cmbPulgadas.Value = drArticulo.GetInt32(1);
				if (drArticulo.GetValue(2) != System.DBNull.Value) this.cmbColor.Value = drArticulo.GetInt32(2);
				if (drArticulo.GetValue(3) != System.DBNull.Value) this.cmbProcesador.Value = drArticulo.GetInt32(3);
				if (drArticulo.GetValue(4) != System.DBNull.Value) this.cmbMarca.Value = drArticulo.GetInt32(4);
				if (drArticulo.GetValue(5) != System.DBNull.Value) this.cmbMemoriaRam.Value = drArticulo.GetInt32(5);
				if (drArticulo.GetValue(6) != System.DBNull.Value) this.cmbDiscoDuro.Value = drArticulo.GetInt32(6);
				if (drArticulo.GetValue(7) != System.DBNull.Value) this.cmbPrecio.Value = drArticulo.GetInt32(7);
				if (drArticulo.GetValue(8) != System.DBNull.Value) this.cmbGaming.Value = drArticulo.GetInt32(8);
				if (drArticulo.GetValue(9) != System.DBNull.Value) this.cmbLumenes.Value = drArticulo.GetInt32(9);
				if (drArticulo.GetValue(10) != System.DBNull.Value) this.cmbResolucion.Value = drArticulo.GetInt32(10);
				if (drArticulo.GetValue(11) != System.DBNull.Value) this.cmbCapacidad.Value = drArticulo.GetInt32(11);
				if (drArticulo.GetValue(12) != System.DBNull.Value) this.cmbQuemadores.Value = drArticulo.GetInt32(12);
				if (drArticulo.GetValue(13) != System.DBNull.Value) this.cmdPixelesCamPost.Value = drArticulo.GetInt32(13);
				if (drArticulo.GetValue(14) != System.DBNull.Value) this.cmbAlmacenamiento.Value = drArticulo.GetInt32(14);
				if (drArticulo.GetValue(15) != System.DBNull.Value) this.cmbTarjetaGrafica.Value = drArticulo.GetInt32(15);
				if (drArticulo.GetValue(16) != System.DBNull.Value) this.cmbOfertaElectrodomesticos.Value = drArticulo.GetInt32(16);
				if (drArticulo.GetValue(17) != System.DBNull.Value) this.cmbOfertaComputo.Value = drArticulo.GetInt32(17);
				if (drArticulo.GetValue(18) != System.DBNull.Value) this.txtCostoEnvio.Value = drArticulo.GetDecimal(18);

				if (drArticulo.GetValue(19) != System.DBNull.Value) this.cmbTipoPantalla.Value = drArticulo.GetInt32(19);
				if (drArticulo.GetValue(20) != System.DBNull.Value) this.cmbSistemaOperativo.Value = drArticulo.GetInt32(20);
				if (drArticulo.GetValue(21) != System.DBNull.Value) this.cmbConectividad.Value = drArticulo.GetInt32(21);
				if (drArticulo.GetValue(22) != System.DBNull.Value) this.cmbSistema.Value = drArticulo.GetInt32(22);
				if (drArticulo.GetValue(23) != System.DBNull.Value) this.cmbPeso.Value = drArticulo.GetInt32(23);
				if (drArticulo.GetValue(24) != System.DBNull.Value) this.cmbPotencia.Value = drArticulo.GetInt32(24);
				if (drArticulo.GetValue(25) != System.DBNull.Value) this.cmbTipoCarga.Value = drArticulo.GetInt32(25);
				if (drArticulo.GetValue(26) != System.DBNull.Value) this.cmbCarga.Value = drArticulo.GetInt32(26);
				if (drArticulo.GetValue(27) != System.DBNull.Value) this.cmbTipo.Value = drArticulo.GetInt32(27);
				if (drArticulo.GetValue(28) != System.DBNull.Value) this.cmbComplementos.Value = drArticulo.GetInt32(28);
				if (drArticulo.GetValue(29) != System.DBNull.Value) this.cmbPixelesCamFront.Value = drArticulo.GetInt32(29);
				if (drArticulo.GetValue(30) != System.DBNull.Value) this.cmbBateria.Value = drArticulo.GetInt32(30);
			}
			drArticulo.Close();
		}

		private void cmbDiscoDuro_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		}

		private void btnPulgadas_Click(object sender, System.EventArgs e)
		{
			using (frmArtPulgadas ArtPul = new frmArtPulgadas())
			{
				if (DialogResult.OK == ArtPul.ShowDialog())
				{
					this.cmbPulgadas.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Select idArtPulgadas, Pulgadas From ArtPulgadas Order By Pulgadas");
					this.cmbPulgadas.Focus();
				}
			}
		}

		private void btnColor_Click(object sender, System.EventArgs e)
		{
			using (frmArtColor ArtCol = new frmArtColor())
			{
				if (DialogResult.OK == ArtCol.ShowDialog())
				{
					this.cmbColor.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Select idArtColor, Color From ArtColor Order By Color");
					this.cmbColor.Focus();
				}
			}
		}

		private void btnProcesador_Click(object sender, System.EventArgs e)
		{
			using (frmArtProcesador ArtPro = new frmArtProcesador())
			{
				if (DialogResult.OK == ArtPro.ShowDialog())
				{
					this.cmbProcesador.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Select idArtProcesador, Procesador From ArtProcesador Order By Procesador");				
					this.cmbProcesador.Focus();
				}
			}
		}

		private void btnMarca_Click(object sender, System.EventArgs e)
		{
			using (frmArtMarca ArtMar = new frmArtMarca())
			{
				if (DialogResult.OK == ArtMar.ShowDialog())
				{
					this.cmbMarca.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Select idArtMarca, Marca From ArtMarca Order By Marca");
					this.cmbMarca.Focus();
				}
			}
		}

		private void btnMemoriaRam_Click(object sender, System.EventArgs e)
		{
			using (frmArtMemoriaRAM ArtRAM = new frmArtMemoriaRAM())
			{
				if (DialogResult.OK == ArtRAM.ShowDialog())
				{
					this.cmbMemoriaRam.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Select idArtMemoriaRAM, MemoriaRam From ArtMemoriaRAM Order By MemoriaRam");
					this.cmbMemoriaRam.Focus();
				}
			}
		}

		private void btnDiscoDuro_Click(object sender, System.EventArgs e)
		{
			using (frmArtDiscoDuro ArtDis = new frmArtDiscoDuro())
			{
				if (DialogResult.OK == ArtDis.ShowDialog())
				{
					this.cmbDiscoDuro.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Select idArtDiscoDuro, DiscoDuro From ArtDiscoDuro Order By DiscoDuro");
					this.cmbDiscoDuro.Focus();
				}
			}
		}

		private void btnGaming_Click(object sender, System.EventArgs e)
		{
			using (frmArtGaming ArtGam = new frmArtGaming())
			{
				if (DialogResult.OK == ArtGam.ShowDialog())
				{
					this.cmbGaming.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Select idArtGaming, Gaming From ArtGaming Order By Gaming");
					this.cmbGaming.Focus();
				}
			}
		}

		private void btnPrecio_Click(object sender, System.EventArgs e)
		{
			using (frmArtPrecio ArtPre = new frmArtPrecio())
			{
				if (DialogResult.OK == ArtPre.ShowDialog())
				{
					this.cmbPrecio.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Select idArtPrecio, Precio From ArtPrecio Order By Precio");
					this.cmbPrecio.Focus();
				}
			}
		}

		private void btnLumenes_Click(object sender, System.EventArgs e)
		{
			using (frmArtLumenes ArtPre = new frmArtLumenes())
			{
				if (DialogResult.OK == ArtPre.ShowDialog())
				{
					this.cmbLumenes.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Select idArtLumenes, Lumenes From ArtLumenes Order By Lumenes");
					this.cmbLumenes.Focus();
				}
			}
		}

		private void btnResolucion_Click(object sender, System.EventArgs e)
		{
			using (frmArtResolucion ArtRes = new frmArtResolucion())
			{
				if (DialogResult.OK == ArtRes.ShowDialog())
				{
					this.cmbResolucion.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Select idArtResolucion, Resolucion From ArtResolucion Order By Resolucion");
					this.cmbResolucion.Focus();
				}
			}
		}

		private void btnCapacidad_Click(object sender, System.EventArgs e)
		{
			using (frmArtCapacidad ArtCap = new frmArtCapacidad())
			{
				if (DialogResult.OK == ArtCap.ShowDialog())
				{
					this.cmbCapacidad.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Select idArtCapacidad, Capacidad From ArtCapacidad Order By Capacidad");
					this.cmbCapacidad.Focus();
				}
			}
		}

		private void btnQuemadores_Click(object sender, System.EventArgs e)
		{
			using (frmArtQuemadores ArtQue = new frmArtQuemadores())
			{
				if (DialogResult.OK == ArtQue.ShowDialog())
				{
					this.cmbQuemadores.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Select idArtQuemadores, Quemadores From ArtQuemadores Order By Quemadores");
					this.cmbQuemadores.Focus();
				}
			}
		}

		private void btnPixelesCamPost_Click(object sender, System.EventArgs e)
		{
			using (frmArtPixelesPosterior ArtPix = new frmArtPixelesPosterior())
			{
				if (DialogResult.OK == ArtPix.ShowDialog())
				{
					this.cmdPixelesCamPost.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Select idArtPixelesPosterior, Pixeles From ArtPixelesPosterior Order By Pixeles");
					this.cmdPixelesCamPost.Focus();
				}
			}
		}

		private void btnAlmacenamiento_Click(object sender, System.EventArgs e)
		{
			using (frmArtAlmacenamiento ArtAlm = new frmArtAlmacenamiento())
			{
				if (DialogResult.OK == ArtAlm.ShowDialog())
				{
					this.cmbAlmacenamiento.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Select idArtAlmacenamiento, Almacenamiento From ArtAlmacenamiento Order By Almacenamiento");
					this.cmbAlmacenamiento.Focus();
				}
			}
		}

		private void btnTarjetaGrafica_Click(object sender, System.EventArgs e)
		{
			using (frmArtTarjetaGrafica ArtTar = new frmArtTarjetaGrafica())
			{
				if (DialogResult.OK == ArtTar.ShowDialog())
				{
					this.cmbTarjetaGrafica.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Select idArtTarjetaGrafica, TarjetaGrafica From ArtTarjetaGrafica Order By TarjetaGrafica");
					this.cmbTarjetaGrafica.Focus();
				}
			}
		}

		private void btnOfertaElectrodomesticos_Click(object sender, System.EventArgs e)
		{
			using (frmArtOfertaElectrodomesticos ArtOfe = new frmArtOfertaElectrodomesticos())
			{
				if (DialogResult.OK == ArtOfe.ShowDialog())
				{
					this.cmbOfertaElectrodomesticos.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Select idArtOfertaElectrodomesticos, OfertaElectrodomestico From ArtOfertaElectrodomesticos Order By OfertaElectrodomestico");
					this.cmbOfertaElectrodomesticos.Focus();
				}
			}
		}

		private void btnOfertaComputo_Click(object sender, System.EventArgs e)
		{
			using (frmArtOfertaComputo ArtComp = new frmArtOfertaComputo())
			{
				if (DialogResult.OK == ArtComp.ShowDialog())
				{
					this.cmbOfertaComputo.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Select idArtOfertaComputo, OfertaComputo From ArtOfertaComputo Order By OfertaComputo");
					this.cmbOfertaComputo.Focus();
				}
			}
		}

		private void btnSalir_Click(object sender, System.EventArgs e)
		{
			this.DialogResult = DialogResult.No;
		}

		private void btnGuardar_Click(object sender, System.EventArgs e)
		{
			#region Variables
			int IdArtPulgadas  = 0;
			int IdArtColor = 0;
			int IdArtProcesador = 0;
			int IdArtMarca = 0;
			int IdArtMemoriaRAM = 0;
			int IdArtDiscoDuro = 0;
			int IdArtPrecio = 0;
			int IdArtGaming = 0;
			int IdArtLumenes = 0;
			int IdArtResolucion = 0;
			int IdArtCapacidad = 0;
			int IdArtAlmacenamiento = 0;
			int IdArtPixelesPosterior = 0;
			int IdArtQuemadores = 0;
			int IdArtTarjetaGrafica = 0;
			int IdArtOfertaElectrodomesticos = 0;
			int IdArtOfertaComputo = 0;

			int IdTipoPantalla= 0;
			int IdSistemaOperativo= 0;
			int IdConectividad= 0;
			int IdSistema= 0;
			int IdPeso= 0;
			int IdPotencia= 0;
			int IdTipoCarga= 0;
			int IdCarga= 0;
			int IdTipo= 0;
			int IdComplementos= 0;
			int IdPixelesCamFront= 0;
			int IdBateria= 0;

			if (this.cmbPulgadas.ActiveRow != null) IdArtPulgadas = (int)this.cmbPulgadas.Value;
			if (this.cmbColor.ActiveRow != null) IdArtColor = (int)this.cmbColor.Value;
			if (this.cmbProcesador.ActiveRow != null) IdArtProcesador = (int)this.cmbProcesador.Value;
			if (this.cmbMarca.ActiveRow != null) IdArtMarca = (int)this.cmbMarca.Value;
			if (this.cmbMemoriaRam.ActiveRow != null) IdArtMemoriaRAM = (int)this.cmbMemoriaRam.Value;
			if (this.cmbDiscoDuro.ActiveRow != null) IdArtDiscoDuro = (int)this.cmbDiscoDuro.Value;
			if (this.cmbPrecio.ActiveRow != null) IdArtPrecio = (int)this.cmbPrecio.Value;
			if (this.cmbGaming.ActiveRow != null) IdArtGaming = (int)this.cmbGaming.Value;
			if (this.cmbLumenes.ActiveRow != null) IdArtLumenes = (int)this.cmbLumenes.Value;
			if (this.cmbResolucion.ActiveRow != null) IdArtResolucion = (int)this.cmbResolucion.Value;
			if (this.cmbCapacidad.ActiveRow != null) IdArtCapacidad = (int)this.cmbCapacidad.Value;
			if (this.cmbAlmacenamiento.ActiveRow != null) IdArtAlmacenamiento = (int)this.cmbAlmacenamiento.Value;
			if (this.cmdPixelesCamPost.ActiveRow != null) IdArtPixelesPosterior = (int)this.cmdPixelesCamPost.Value;
			if (this.cmbQuemadores.ActiveRow != null) IdArtQuemadores = (int)this.cmbQuemadores.Value;
			if (this.cmbTarjetaGrafica.ActiveRow != null) IdArtTarjetaGrafica = (int)this.cmbTarjetaGrafica.Value;
			if (this.cmbOfertaElectrodomesticos.ActiveRow != null) IdArtOfertaElectrodomesticos = (int)this.cmbOfertaElectrodomesticos.Value;

			if (this.cmbTipoPantalla.ActiveRow != null)IdTipoPantalla = (int)this.cmbTipoPantalla.Value;
			if (this.cmbSistemaOperativo.ActiveRow != null)IdSistemaOperativo = (int)this.cmbSistemaOperativo.Value;
			if (this.cmbConectividad.ActiveRow != null)IdConectividad = (int)this.cmbConectividad.Value;
			if (this.cmbSistema.ActiveRow != null)IdSistema = (int)this.cmbSistema.Value;
			if (this.cmbPeso.ActiveRow != null)IdPeso = (int)this.cmbPeso.Value;
			if (this.cmbPotencia.ActiveRow != null)IdPotencia = (int)this.cmbPotencia.Value;
			if (this.cmbTipoCarga.ActiveRow != null)IdTipoCarga = (int)this.cmbTipoCarga.Value;
			if (this.cmbCarga.ActiveRow != null)IdCarga = (int)this.cmbCarga.Value;
			if (this.cmbTipo.ActiveRow != null)IdTipo = (int)this.cmbTipo.Value;
			if (this.cmbComplementos.ActiveRow != null)IdComplementos = (int)this.cmbComplementos.Value;
			if (this.cmbPixelesCamFront.ActiveRow != null)IdPixelesCamFront = (int)this.cmbPixelesCamFront.Value;
			if (this.cmbBateria.ActiveRow != null)IdBateria = (int)this.cmbBateria.Value;





			#endregion Variables

			#region Grabar
			string sSQLGuarda = string.Format("Exec GrabaArticulosAtributos {0}, {1}, {2}, {3}, {4}, {5}, {6}, {7}, {8}, {9}, {10}, {11}, {12}, {13}, {14}, {15}, {16}, {17}, {18}, {19}, {20}, {21}, {22}, {23}, {24}, {25}, {26}, {27}, {28}, {29}, {30}" ,
				(int)this.txtIdArticulo.Value, IdArtPulgadas, IdArtColor, IdArtProcesador, IdArtMarca, IdArtMemoriaRAM, IdArtDiscoDuro, IdArtPrecio, IdArtGaming, IdArtLumenes, IdArtResolucion,
				IdArtCapacidad, IdArtAlmacenamiento, IdArtPixelesPosterior, IdArtQuemadores, IdArtTarjetaGrafica, IdArtOfertaElectrodomesticos, IdArtOfertaComputo, this.txtCostoEnvio.Value,
				IdTipoPantalla,IdSistemaOperativo,IdConectividad,IdSistema,IdPeso,IdPotencia,IdTipoCarga,IdCarga,IdTipo,
				IdComplementos,IdPixelesCamFront,IdBateria);

			Funcion.EjecutaSQL(cdsSeteoF, sSQLGuarda, true);

			#endregion Graba
			this.DialogResult = DialogResult.OK;
		}

		private void frmAtributosArticulos_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Escape) this.DialogResult = DialogResult.No;
		}

		private void label24_Click(object sender, System.EventArgs e)
		{
		
		}

		private void groupBox4_Enter(object sender, System.EventArgs e)
		{
		
		}

		private void ultraCombo7_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}

		private void btnTipoPantalla_Click(object sender, System.EventArgs e)
		{
			using (frmArtTipoPantalla ArtTapPan = new frmArtTipoPantalla())
			{
				if (DialogResult.OK == ArtTapPan.ShowDialog())
				{
					this.cmbTipoPantalla.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Select idArtTipoPantalla, TipoPantalla From ArtTipoPantalla Order By TipoPantalla");
					this.cmbTipoPantalla.Focus();
				}
			}
		}

		private void btnSistemaOperativo_Click(object sender, System.EventArgs e)
		{
			using (frmArtSistemaOperativo ArtSO = new frmArtSistemaOperativo())
			{
				if (DialogResult.OK == ArtSO.ShowDialog())
				{
					this.cmbSistemaOperativo.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Select idArtSistemaOperativo, SistemaOperativo From ArtSistemaOperativo Order By SistemaOperativo");
					this.cmbSistemaOperativo.Focus();
				}
			}
		}

		private void btnConectividad_Click(object sender, System.EventArgs e)
		{
			using (frmArtConectividad ArtcCO = new frmArtConectividad())
			{
				if (DialogResult.OK == ArtcCO.ShowDialog())
				{
					this.cmbConectividad.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Select idArtConectividad, Conectividad From ArtConectividad Order By Conectividad");
					this.cmbConectividad.Focus();
				}
			}
		}

		private void btnSistema_Click(object sender, System.EventArgs e)
		{
			
			using (frmArtSistema ArtS = new frmArtSistema())
			{
				if (DialogResult.OK == ArtS.ShowDialog())
				{
					this.cmbSistema.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Select idArtSistema, Sistema From ArtSistema Order By Sistema");
					this.cmbSistema.Focus();
				}
			}
		}

		private void btnPeso_Click(object sender, System.EventArgs e)
		{
			using (frmArtPeso Artps = new frmArtPeso())
			{
				if (DialogResult.OK == Artps.ShowDialog())
				{
					this.cmbPeso.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Select idArtPeso, Peso From ArtPeso Order By Peso");
					this.cmbPeso.Focus();
				}
			}
		}

		private void btnCarga_Click(object sender, System.EventArgs e)
		{
			using (frmArtCarga ArtCG = new frmArtCarga())
			{
				if (DialogResult.OK == ArtCG.ShowDialog())
				{
					this.cmbCarga.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Select idArtCarga, Carga From ArtCarga Order By Carga");
					this.cmbCarga.Focus();
				}
			}
		
		}

		private void btnTipo_Click(object sender, System.EventArgs e)
		{
			using (frmArtTipo Artt = new frmArtTipo())
			{
				if (DialogResult.OK == Artt.ShowDialog())
				{
					this.cmbTipo.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Select idArtTipo, Tipo From ArtTipo Order By Tipo");
					this.cmbTipo.Focus();
				}
			}
		
		}

		private void btnPixelesCamFront_Click(object sender, System.EventArgs e)
		{
			using (frmArtPixelesFronatal ArtPT = new frmArtPixelesFronatal())
			{
				if (DialogResult.OK == ArtPT.ShowDialog())
				{
					this.cmbPixelesCamFront.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Select idArtPixelesFrontal, PixelesFrontal From ArtPixelesFrontal Order By PixelesFrontal");
					this.cmbPixelesCamFront.Focus();
				}
			}
		
		}

		private void btnBateria_Click(object sender, System.EventArgs e)
		{
			using (frmArrBateria ArtB = new frmArrBateria())
			{
				if (DialogResult.OK == ArtB.ShowDialog())
				{
					this.cmbBateria.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Select idArtBateria, Bateria From ArtBateria Order By Bateria");
					this.cmbBateria.Focus();
				}
			}
		
		}

		private void btnComplementos_Click(object sender, System.EventArgs e)
		{
			using (frmArtComplementos ArtCP = new frmArtComplementos())
			{
				if (DialogResult.OK == ArtCP.ShowDialog())
				{
					this.cmbComplementos.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Select idArtComplementos, Complementos From ArtComplementos Order By Complementos");
					this.cmbComplementos.Focus();
				}
			}
		
		}

		private void btnTipoCarga_Click(object sender, System.EventArgs e)
		{
			using (frmArtTipoCarga ArtTC = new frmArtTipoCarga())
			{
				if (DialogResult.OK == ArtTC.ShowDialog())
				{
					this.cmbTipoCarga.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Select idArtTipoCarga, TipoCarga From ArtTipoCarga Order By TipoCarga");
					this.cmbTipoCarga.Focus();
				}
			}
		
		}

		private void btnPotencia_Click(object sender, System.EventArgs e)
		{
			using (frmArtPotencia ArtPP = new frmArtPotencia())
			{
				if (DialogResult.OK == ArtPP.ShowDialog())
				{
					this.cmbPotencia.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Select idArtPotencia, Potencia From ArtPotencia Order By Potencia");
					this.cmbPotencia.Focus();
				}
			}
		}
	}
}
