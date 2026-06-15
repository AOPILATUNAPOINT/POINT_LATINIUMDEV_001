using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using DevExpress.XtraEditors;
using System.Data.SqlClient;

namespace Latinium
{
	/// <summary>
	/// Summary description for frmArticulos.
	/// </summary>
	public class frmArticulos : DevExpress.XtraEditors.XtraForm
	{
		private System.Windows.Forms.Label label58;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkCombo;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtTamaño;
		private System.Windows.Forms.Label label56;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtPeriodoGarantia;
		private System.Windows.Forms.Label label54;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbCategoria;
		private System.Windows.Forms.GroupBox groupBox1;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtPeso;
		private System.Windows.Forms.Label lblPeso;
		private System.Windows.Forms.Label label51;
		private System.Windows.Forms.Label label49;
		private System.Windows.Forms.Label label50;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtModelo;
		private System.Windows.Forms.Label label35;
		private System.Windows.Forms.Label label23;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkSerialAlterna;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkManejaSerial;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkFaturable;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkAplicDescuento;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtNParte;
		private System.Windows.Forms.Label label19;
		private System.Windows.Forms.Label label47;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbColor;
		private System.Windows.Forms.Label label48;
		private DevExpress.XtraEditors.SimpleButton btColor;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtSaldo;
		private System.Windows.Forms.Label label32;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbArticuloSRI;
		private System.Windows.Forms.Label label46;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtCodigo;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtIva;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtCantMaxima;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtCantMinima;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbArancel;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbSubGrupo;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbGrupo;
		private DevExpress.XtraEditors.SimpleButton btArancel;
		private DevExpress.XtraEditors.SimpleButton btGrupo;
		private System.Windows.Forms.Label label14;
		private System.Windows.Forms.Label label13;
		private System.Windows.Forms.Label label12;
		private System.Windows.Forms.Label label9;
		private System.Windows.Forms.Label label6;
		private System.Windows.Forms.Label label7;
		private System.Windows.Forms.Label label3;
		private System.Windows.Forms.Label label2;
		private System.Windows.Forms.Label label1;
		private System.Windows.Forms.Label lblCostoUltim;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtPrecioCostoUltimo;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtPrecioPVP;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtPrecioTarjeta;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtPrecioPA;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtPrecioPI;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtPrecioIncr2;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtPrecioIncr1;
		private System.Windows.Forms.Label label40;
		private System.Windows.Forms.Label label39;
		private System.Windows.Forms.Label label37;
		private System.Windows.Forms.Label label36;
		private System.Windows.Forms.Label label15;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtCodigoReemplazo;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkDescontinuado;
		private System.Windows.Forms.Label label20;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtDepreciacion;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtTiempo;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbSeccion;
		private System.Windows.Forms.Label label31;
		private DevExpress.XtraEditors.SimpleButton btSeccion;
		private System.Windows.Forms.Label label29;
		private System.Windows.Forms.Label lblValorResid;
		private System.Windows.Forms.Label lblDeprecia;
		private DevExpress.XtraEditors.RadioGroup rbgOrigen;
		private System.Windows.Forms.Label label4;
		private System.Windows.Forms.Label label17;
		private System.Windows.Forms.GroupBox groupBox2;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkPrecio;
		private System.Windows.Forms.GroupBox groupBox3;
		private System.Windows.Forms.GroupBox groupBox4;
		private System.Windows.Forms.GroupBox groupBox5;
		private System.Windows.Forms.Button btnNuevo;
		private System.Windows.Forms.Button btnConsultar;
		private System.Windows.Forms.Button btnEditar;
		private System.Windows.Forms.Button btnImprimir;
		private System.Windows.Forms.Button btnGuardar;
		private System.Windows.Forms.Button btnCancelar;
		private System.Windows.Forms.Button btnSalir;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtIdArticulo;
		/// <summary>
		/// Required designer variable.
		/// </summary>
		private System.ComponentModel.Container components = null;

		bool bNuevo = false;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtVidaUtil;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtLargo;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtAlto;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtAncho;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkDescCompras;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkDescVentas;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtFechaUltimo;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbCentroCosto;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource1;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtIdTipoGrupo;
		private C1.Data.C1DataSet cdsSeteoF;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource2;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource3;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource4;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource5;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtResidual;
		bool bEdicion = false;
		private Acceso miAcceso;
		string sCodigo = "";
		decimal dPorcTarjeta = 0.00m;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource6;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource7;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource8;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtCodArtEasyMax;
		private System.Windows.Forms.Label label57;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkTransferible;
		private System.Windows.Forms.Label label5;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbMarca;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource9;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbPreciador;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkInventario;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkImpuestoIVA;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkGarantiaExtendida;
		private System.Windows.Forms.Label label8;
		private System.Windows.Forms.Label label10;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkServicioTecnicoDomicilio;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtTiempoGE;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtPorcentajeGE;
		private System.Windows.Forms.Button btnCaracteristicas;
		private System.Windows.Forms.Label label11;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtLink;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkObsequios;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkcatalogo;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtCSCostoEfectivo;
		private System.Windows.Forms.Label label16;
		private System.Windows.Forms.Label label52;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtCSCostoCredito;
		private System.Windows.Forms.Label label18;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkConsignacion;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtCodigoProveedor;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtPrecioTCsinIntereses;
		private System.Windows.Forms.Label label21;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkProductoWeb;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkPayJoy;
		private System.Windows.Forms.Label label22;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtCodigoPayJoy;
		private System.Windows.Forms.Label label24;
		private System.Windows.Forms.Label label25;
		private System.Windows.Forms.Label label26;
		private System.Windows.Forms.Label label27;
		private System.Windows.Forms.Label label28;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtOrigenMoto;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtTonelaje;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtfechaMoto;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtCilindraje;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtPasajeros;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkMoto;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtNombre;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkMigrar;
		private System.Windows.Forms.Label label30;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkManejaSerie;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtnovasoft;
		private System.Windows.Forms.Label label33;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkPrecioMovil;
		private System.Windows.Forms.Button btnAtributos;
		private System.Windows.Forms.Label label34;
		private System.Windows.Forms.Label label38;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtPrecioAutoconsumo;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtPrecioDescEmpleado;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkDepartamento;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkDepartamento1;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkEfectivo;
		private System.Windows.Forms.Label label41;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtDescripcionWeb;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtTituloWeb;
		private System.Windows.Forms.Label label42;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtCostoPayJoy;
		decimal dPorcCredito = 0.00m;

		public frmArticulos()
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
			Infragistics.Win.Appearance appearance1 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCategoria");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Categoria");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCategoria");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Categoria");
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idArancel");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Porcentaje");
			Infragistics.Win.Appearance appearance11 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Arancel");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn3 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idArancel");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn4 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Codigo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn5 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Porcentaje");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn6 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Arancel");
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmArticulos));
			Infragistics.Win.Appearance appearance12 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance13 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance14 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance15 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance16 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance17 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance18 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand3 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idColor");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Color");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn7 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idColor");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn8 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Color");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn9 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idMarca");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn10 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Marca");
			Infragistics.Win.Appearance appearance19 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance20 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand4 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idArticuloSRI");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn11 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Detalle");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn12 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Porcentaje");
			Infragistics.Win.Appearance appearance21 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn11 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idArticuloSRI");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn12 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Codigo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn13 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Detalle");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn14 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Porcentaje");
			Infragistics.Win.Appearance appearance22 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance23 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance24 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance25 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance26 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance27 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand5 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn13 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idSubGrupo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn14 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("SubGrupo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn15 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idSubGrupo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn16 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("SubGrupo");
			Infragistics.Win.Appearance appearance28 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand6 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn15 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idGrupoArticulo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn16 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Grupo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn17 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idGrupoArticulo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn18 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Grupo");
			Infragistics.Win.Appearance appearance29 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance30 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance31 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance32 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance33 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance34 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance35 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance36 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance37 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance38 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance39 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance40 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton1 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.Appearance appearance41 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance42 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance43 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand7 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn17 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idSeccion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn18 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Seccion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn19 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idSeccion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn20 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Seccion");
			Infragistics.Win.Appearance appearance44 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance45 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance46 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance47 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance48 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance49 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance50 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance51 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance52 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance53 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance54 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance55 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance56 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance57 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance58 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance59 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance60 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance61 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance62 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand8 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn19 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPreciador");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn20 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Preciador");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn21 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idPreciador");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn22 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Preciador");
			Infragistics.Win.Appearance appearance63 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand9 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn21 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idMarca");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn22 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Marca");
			Infragistics.Win.Appearance appearance64 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance65 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance66 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance67 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance68 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance69 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance70 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance71 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance72 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance73 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance74 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance75 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance76 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance77 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance78 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance79 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance80 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance81 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance82 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance83 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance84 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance85 = new Infragistics.Win.Appearance();
			this.txtVidaUtil = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label58 = new System.Windows.Forms.Label();
			this.chkCombo = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.txtTamaño = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label56 = new System.Windows.Forms.Label();
			this.txtPeriodoGarantia = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label54 = new System.Windows.Forms.Label();
			this.cmbCategoria = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.ultraDataSource2 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.groupBox1 = new System.Windows.Forms.GroupBox();
			this.txtLargo = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtPeso = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.lblPeso = new System.Windows.Forms.Label();
			this.label51 = new System.Windows.Forms.Label();
			this.txtAlto = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label49 = new System.Windows.Forms.Label();
			this.label50 = new System.Windows.Forms.Label();
			this.txtAncho = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.rbgOrigen = new DevExpress.XtraEditors.RadioGroup();
			this.label14 = new System.Windows.Forms.Label();
			this.cmbArancel = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.ultraDataSource7 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.btArancel = new DevExpress.XtraEditors.SimpleButton();
			this.txtModelo = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label35 = new System.Windows.Forms.Label();
			this.label23 = new System.Windows.Forms.Label();
			this.chkSerialAlterna = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.chkManejaSerial = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.chkFaturable = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.chkAplicDescuento = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.txtNParte = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label19 = new System.Windows.Forms.Label();
			this.label47 = new System.Windows.Forms.Label();
			this.cmbColor = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.ultraDataSource6 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.label48 = new System.Windows.Forms.Label();
			this.btColor = new DevExpress.XtraEditors.SimpleButton();
			this.ultraDataSource3 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.txtSaldo = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label32 = new System.Windows.Forms.Label();
			this.cmbArticuloSRI = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.ultraDataSource4 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.label46 = new System.Windows.Forms.Label();
			this.txtCodigo = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtNombre = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtIva = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtCantMaxima = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtCantMinima = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.cmbSubGrupo = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.ultraDataSource5 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.cmbGrupo = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.ultraDataSource1 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.btGrupo = new DevExpress.XtraEditors.SimpleButton();
			this.label13 = new System.Windows.Forms.Label();
			this.label12 = new System.Windows.Forms.Label();
			this.label9 = new System.Windows.Forms.Label();
			this.label6 = new System.Windows.Forms.Label();
			this.label7 = new System.Windows.Forms.Label();
			this.label3 = new System.Windows.Forms.Label();
			this.label2 = new System.Windows.Forms.Label();
			this.label1 = new System.Windows.Forms.Label();
			this.lblCostoUltim = new System.Windows.Forms.Label();
			this.txtPrecioCostoUltimo = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtPrecioPVP = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtPrecioTarjeta = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtPrecioPA = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtPrecioPI = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtPrecioIncr2 = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtPrecioIncr1 = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label40 = new System.Windows.Forms.Label();
			this.label39 = new System.Windows.Forms.Label();
			this.label37 = new System.Windows.Forms.Label();
			this.label36 = new System.Windows.Forms.Label();
			this.label15 = new System.Windows.Forms.Label();
			this.txtCodigoReemplazo = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.chkDescCompras = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.chkDescVentas = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.chkDescontinuado = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.dtFechaUltimo = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.label20 = new System.Windows.Forms.Label();
			this.txtDepreciacion = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtTiempo = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.cmbSeccion = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.ultraDataSource8 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.label31 = new System.Windows.Forms.Label();
			this.btSeccion = new DevExpress.XtraEditors.SimpleButton();
			this.label29 = new System.Windows.Forms.Label();
			this.lblValorResid = new System.Windows.Forms.Label();
			this.lblDeprecia = new System.Windows.Forms.Label();
			this.label4 = new System.Windows.Forms.Label();
			this.label17 = new System.Windows.Forms.Label();
			this.groupBox2 = new System.Windows.Forms.GroupBox();
			this.label38 = new System.Windows.Forms.Label();
			this.txtPrecioDescEmpleado = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label34 = new System.Windows.Forms.Label();
			this.txtPrecioAutoconsumo = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label33 = new System.Windows.Forms.Label();
			this.txtCostoPayJoy = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label21 = new System.Windows.Forms.Label();
			this.txtPrecioTCsinIntereses = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.chkPrecio = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.chkDepartamento = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.chkDepartamento1 = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.groupBox3 = new System.Windows.Forms.GroupBox();
			this.txtResidual = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.cmbCentroCosto = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.groupBox4 = new System.Windows.Forms.GroupBox();
			this.groupBox5 = new System.Windows.Forms.GroupBox();
			this.chkEfectivo = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.chkcatalogo = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.chkObsequios = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.chkImpuestoIVA = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.chkInventario = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.chkTransferible = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.chkPrecioMovil = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.chkProductoWeb = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.btnNuevo = new System.Windows.Forms.Button();
			this.btnConsultar = new System.Windows.Forms.Button();
			this.btnEditar = new System.Windows.Forms.Button();
			this.btnImprimir = new System.Windows.Forms.Button();
			this.btnGuardar = new System.Windows.Forms.Button();
			this.btnCancelar = new System.Windows.Forms.Button();
			this.btnSalir = new System.Windows.Forms.Button();
			this.txtIdArticulo = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtIdTipoGrupo = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.txtCodArtEasyMax = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label57 = new System.Windows.Forms.Label();
			this.label5 = new System.Windows.Forms.Label();
			this.cmbPreciador = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.ultraDataSource9 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.cmbMarca = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.chkGarantiaExtendida = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.chkServicioTecnicoDomicilio = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.label8 = new System.Windows.Forms.Label();
			this.txtTiempoGE = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label10 = new System.Windows.Forms.Label();
			this.txtPorcentajeGE = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.btnCaracteristicas = new System.Windows.Forms.Button();
			this.label11 = new System.Windows.Forms.Label();
			this.txtLink = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.chkConsignacion = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.txtCSCostoEfectivo = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label16 = new System.Windows.Forms.Label();
			this.label52 = new System.Windows.Forms.Label();
			this.txtCSCostoCredito = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label18 = new System.Windows.Forms.Label();
			this.txtCodigoProveedor = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.chkPayJoy = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.label22 = new System.Windows.Forms.Label();
			this.txtCodigoPayJoy = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label24 = new System.Windows.Forms.Label();
			this.label25 = new System.Windows.Forms.Label();
			this.label26 = new System.Windows.Forms.Label();
			this.label27 = new System.Windows.Forms.Label();
			this.label28 = new System.Windows.Forms.Label();
			this.txtfechaMoto = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtOrigenMoto = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtTonelaje = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtCilindraje = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtPasajeros = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.chkMoto = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.btnAtributos = new System.Windows.Forms.Button();
			this.chkMigrar = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.label30 = new System.Windows.Forms.Label();
			this.txtnovasoft = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.chkManejaSerie = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.label41 = new System.Windows.Forms.Label();
			this.txtDescripcionWeb = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtTituloWeb = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label42 = new System.Windows.Forms.Label();
			((System.ComponentModel.ISupportInitialize)(this.txtVidaUtil)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtTamaño)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtPeriodoGarantia)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCategoria)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).BeginInit();
			this.groupBox1.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.txtLargo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtPeso)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtAlto)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtAncho)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.rbgOrigen.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbArancel)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource7)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtModelo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNParte)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbColor)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource6)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource3)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtSaldo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbArticuloSRI)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource4)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCodigo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNombre)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIva)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCantMaxima)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCantMinima)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbSubGrupo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource5)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbGrupo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtPrecioCostoUltimo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtPrecioPVP)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtPrecioTarjeta)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtPrecioPA)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtPrecioPI)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtPrecioIncr2)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtPrecioIncr1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCodigoReemplazo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFechaUltimo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDepreciacion)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtTiempo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbSeccion)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource8)).BeginInit();
			this.groupBox2.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.txtPrecioDescEmpleado)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtPrecioAutoconsumo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCostoPayJoy)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtPrecioTCsinIntereses)).BeginInit();
			this.groupBox3.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.txtResidual)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCentroCosto)).BeginInit();
			this.groupBox5.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.txtIdArticulo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdTipoGrupo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCodArtEasyMax)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbPreciador)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource9)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbMarca)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtTiempoGE)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtPorcentajeGE)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtLink)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCSCostoEfectivo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCSCostoCredito)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCodigoProveedor)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCodigoPayJoy)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtfechaMoto)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtOrigenMoto)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtTonelaje)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCilindraje)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtPasajeros)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtnovasoft)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDescripcionWeb)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtTituloWeb)).BeginInit();
			this.SuspendLayout();
			// 
			// txtVidaUtil
			// 
			appearance1.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtVidaUtil.Appearance = appearance1;
			this.txtVidaUtil.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtVidaUtil.Enabled = false;
			this.txtVidaUtil.FormatString = "##0";
			this.txtVidaUtil.Location = new System.Drawing.Point(360, 328);
			this.txtVidaUtil.MaskInput = "nnn";
			this.txtVidaUtil.MaxValue = 365;
			this.txtVidaUtil.MinValue = 0;
			this.txtVidaUtil.Name = "txtVidaUtil";
			this.txtVidaUtil.PromptChar = ' ';
			this.txtVidaUtil.Size = new System.Drawing.Size(72, 22);
			this.txtVidaUtil.TabIndex = 21;
			this.txtVidaUtil.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtVidaUtil_KeyDown);
			this.txtVidaUtil.Click += new System.EventHandler(this.txtVidaUtil_Click);
			this.txtVidaUtil.Enter += new System.EventHandler(this.txtVidaUtil_Enter);
			// 
			// label58
			// 
			this.label58.AutoSize = true;
			this.label58.BackColor = System.Drawing.Color.Transparent;
			this.label58.Location = new System.Drawing.Point(208, 328);
			this.label58.Name = "label58";
			this.label58.Size = new System.Drawing.Size(46, 16);
			this.label58.TabIndex = 158;
			this.label58.Text = "Vida Util";
			this.label58.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// chkCombo
			// 
			appearance2.ForeColorDisabled = System.Drawing.Color.Black;
			this.chkCombo.Appearance = appearance2;
			this.chkCombo.BackColor = System.Drawing.Color.Transparent;
			this.chkCombo.Enabled = false;
			this.chkCombo.Location = new System.Drawing.Point(8, 16);
			this.chkCombo.Name = "chkCombo";
			this.chkCombo.Size = new System.Drawing.Size(88, 22);
			this.chkCombo.TabIndex = 48;
			this.chkCombo.Text = "Combo";
			this.chkCombo.KeyDown += new System.Windows.Forms.KeyEventHandler(this.chkCombo_KeyDown);
			// 
			// txtTamaño
			// 
			appearance3.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtTamaño.Appearance = appearance3;
			this.txtTamaño.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtTamaño.Enabled = false;
			this.txtTamaño.FormatString = "##0";
			this.txtTamaño.Location = new System.Drawing.Point(112, 296);
			this.txtTamaño.MaskInput = "nnn";
			this.txtTamaño.MaxValue = 365;
			this.txtTamaño.MinValue = 0;
			this.txtTamaño.Name = "txtTamaño";
			this.txtTamaño.PromptChar = ' ';
			this.txtTamaño.Size = new System.Drawing.Size(72, 22);
			this.txtTamaño.TabIndex = 19;
			this.txtTamaño.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtTamaño_KeyDown);
			this.txtTamaño.Click += new System.EventHandler(this.txtTamaño_Click);
			this.txtTamaño.Validated += new System.EventHandler(this.txtTamaño_Validated);
			this.txtTamaño.Enter += new System.EventHandler(this.txtTamaño_Enter);
			// 
			// label56
			// 
			this.label56.AutoSize = true;
			this.label56.BackColor = System.Drawing.Color.Transparent;
			this.label56.Location = new System.Drawing.Point(24, 296);
			this.label56.Name = "label56";
			this.label56.Size = new System.Drawing.Size(49, 16);
			this.label56.TabIndex = 155;
			this.label56.Text = "Pulgadas";
			this.label56.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtPeriodoGarantia
			// 
			appearance4.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtPeriodoGarantia.Appearance = appearance4;
			this.txtPeriodoGarantia.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtPeriodoGarantia.Enabled = false;
			this.txtPeriodoGarantia.FormatString = "##0";
			this.txtPeriodoGarantia.Location = new System.Drawing.Point(360, 296);
			this.txtPeriodoGarantia.MaskInput = "nnnn";
			this.txtPeriodoGarantia.MaxValue = 3650;
			this.txtPeriodoGarantia.MinValue = 0;
			this.txtPeriodoGarantia.Name = "txtPeriodoGarantia";
			this.txtPeriodoGarantia.PromptChar = ' ';
			this.txtPeriodoGarantia.Size = new System.Drawing.Size(72, 22);
			this.txtPeriodoGarantia.TabIndex = 20;
			this.txtPeriodoGarantia.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtPeriodoGarantia_KeyDown);
			this.txtPeriodoGarantia.Click += new System.EventHandler(this.txtPeriodoGarantia_Click);
			this.txtPeriodoGarantia.Enter += new System.EventHandler(this.txtPeriodoGarantia_Enter);
			// 
			// label54
			// 
			this.label54.AutoSize = true;
			this.label54.BackColor = System.Drawing.Color.Transparent;
			this.label54.Location = new System.Drawing.Point(208, 304);
			this.label54.Name = "label54";
			this.label54.Size = new System.Drawing.Size(135, 16);
			this.label54.TabIndex = 153;
			this.label54.Text = "Período de garantía (Dias)";
			this.label54.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// cmbCategoria
			// 
			appearance5.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbCategoria.Appearance = appearance5;
			this.cmbCategoria.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbCategoria.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbCategoria.DataSource = this.ultraDataSource2;
			this.cmbCategoria.DisplayLayout.AutoFitColumns = true;
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn1.Width = 101;
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridColumn2.Width = 237;
			ultraGridBand1.Columns.AddRange(new object[] {
																										 ultraGridColumn1,
																										 ultraGridColumn2});
			this.cmbCategoria.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.cmbCategoria.DisplayMember = "Categoria";
			this.cmbCategoria.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbCategoria.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbCategoria.Enabled = false;
			this.cmbCategoria.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbCategoria.Location = new System.Drawing.Point(112, 40);
			this.cmbCategoria.Name = "cmbCategoria";
			this.cmbCategoria.Size = new System.Drawing.Size(256, 21);
			this.cmbCategoria.TabIndex = 5;
			this.cmbCategoria.ValueMember = "idCategoria";
			this.cmbCategoria.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cmbCategoria_KeyDown);
			this.cmbCategoria.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.cmbCategoria_InitializeLayout);
			this.cmbCategoria.Enter += new System.EventHandler(this.cmbCategoria_Enter);
			// 
			// ultraDataSource2
			// 
			ultraDataColumn1.DataType = typeof(int);
			this.ultraDataSource2.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn1,
																																 ultraDataColumn2});
			// 
			// groupBox1
			// 
			this.groupBox1.Controls.Add(this.txtLargo);
			this.groupBox1.Controls.Add(this.txtPeso);
			this.groupBox1.Controls.Add(this.lblPeso);
			this.groupBox1.Controls.Add(this.label51);
			this.groupBox1.Controls.Add(this.txtAlto);
			this.groupBox1.Controls.Add(this.label49);
			this.groupBox1.Controls.Add(this.label50);
			this.groupBox1.Controls.Add(this.txtAncho);
			this.groupBox1.Controls.Add(this.rbgOrigen);
			this.groupBox1.Controls.Add(this.label14);
			this.groupBox1.Controls.Add(this.cmbArancel);
			this.groupBox1.Controls.Add(this.btArancel);
			this.groupBox1.Location = new System.Drawing.Point(8, 400);
			this.groupBox1.Name = "groupBox1";
			this.groupBox1.Size = new System.Drawing.Size(248, 224);
			this.groupBox1.TabIndex = 151;
			this.groupBox1.TabStop = false;
			this.groupBox1.Text = "Origen";
			this.groupBox1.Enter += new System.EventHandler(this.groupBox1_Enter);
			// 
			// txtLargo
			// 
			appearance6.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtLargo.Appearance = appearance6;
			this.txtLargo.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtLargo.Enabled = false;
			this.txtLargo.FormatString = "#,##0.00";
			this.txtLargo.Location = new System.Drawing.Point(80, 152);
			this.txtLargo.Name = "txtLargo";
			this.txtLargo.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtLargo.PromptChar = ' ';
			this.txtLargo.Size = new System.Drawing.Size(96, 22);
			this.txtLargo.TabIndex = 30;
			this.txtLargo.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtLargo_KeyDown);
			this.txtLargo.Click += new System.EventHandler(this.txtLargo_Click);
			this.txtLargo.Enter += new System.EventHandler(this.txtLargo_Enter);
			// 
			// txtPeso
			// 
			appearance7.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtPeso.Appearance = appearance7;
			this.txtPeso.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtPeso.Enabled = false;
			this.txtPeso.FormatString = "#,##0.00";
			this.txtPeso.Location = new System.Drawing.Point(80, 104);
			this.txtPeso.MaskInput = "{LOC}nnnnnnn.nn";
			this.txtPeso.Name = "txtPeso";
			this.txtPeso.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtPeso.PromptChar = ' ';
			this.txtPeso.Size = new System.Drawing.Size(96, 22);
			this.txtPeso.TabIndex = 28;
			this.txtPeso.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtPeso_KeyDown);
			this.txtPeso.Click += new System.EventHandler(this.txtPeso_Click);
			this.txtPeso.Enter += new System.EventHandler(this.txtPeso_Enter);
			// 
			// lblPeso
			// 
			this.lblPeso.AutoSize = true;
			this.lblPeso.Location = new System.Drawing.Point(8, 107);
			this.lblPeso.Name = "lblPeso";
			this.lblPeso.Size = new System.Drawing.Size(63, 16);
			this.lblPeso.TabIndex = 147;
			this.lblPeso.Text = "Peso (Kilos)";
			this.lblPeso.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label51
			// 
			this.label51.AutoSize = true;
			this.label51.Location = new System.Drawing.Point(8, 179);
			this.label51.Name = "label51";
			this.label51.Size = new System.Drawing.Size(23, 16);
			this.label51.TabIndex = 150;
			this.label51.Text = "Alto";
			this.label51.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtAlto
			// 
			appearance8.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtAlto.Appearance = appearance8;
			this.txtAlto.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtAlto.Enabled = false;
			this.txtAlto.FormatString = "#,##0.00";
			this.txtAlto.Location = new System.Drawing.Point(80, 176);
			this.txtAlto.Name = "txtAlto";
			this.txtAlto.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtAlto.PromptChar = ' ';
			this.txtAlto.Size = new System.Drawing.Size(96, 22);
			this.txtAlto.TabIndex = 31;
			this.txtAlto.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtAlto_KeyDown);
			this.txtAlto.Click += new System.EventHandler(this.txtAlto_Click);
			this.txtAlto.Enter += new System.EventHandler(this.txtAlto_Enter);
			// 
			// label49
			// 
			this.label49.AutoSize = true;
			this.label49.Location = new System.Drawing.Point(8, 155);
			this.label49.Name = "label49";
			this.label49.Size = new System.Drawing.Size(32, 16);
			this.label49.TabIndex = 148;
			this.label49.Text = "Largo";
			this.label49.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label50
			// 
			this.label50.AutoSize = true;
			this.label50.Location = new System.Drawing.Point(8, 131);
			this.label50.Name = "label50";
			this.label50.Size = new System.Drawing.Size(35, 16);
			this.label50.TabIndex = 149;
			this.label50.Text = "Ancho";
			this.label50.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtAncho
			// 
			appearance9.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtAncho.Appearance = appearance9;
			this.txtAncho.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtAncho.Enabled = false;
			this.txtAncho.FormatString = "#,##0.00";
			this.txtAncho.Location = new System.Drawing.Point(80, 128);
			this.txtAncho.Name = "txtAncho";
			this.txtAncho.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtAncho.PromptChar = ' ';
			this.txtAncho.Size = new System.Drawing.Size(96, 22);
			this.txtAncho.TabIndex = 29;
			this.txtAncho.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtAncho_KeyDown);
			this.txtAncho.Click += new System.EventHandler(this.txtAncho_Click);
			this.txtAncho.Enter += new System.EventHandler(this.txtAncho_Enter);
			// 
			// rbgOrigen
			// 
			this.rbgOrigen.EditValue = 0;
			this.rbgOrigen.Enabled = false;
			this.rbgOrigen.Location = new System.Drawing.Point(8, 24);
			this.rbgOrigen.Name = "rbgOrigen";
			// 
			// rbgOrigen.Properties
			// 
			this.rbgOrigen.Properties.Appearance.BackColor = System.Drawing.Color.Transparent;
			this.rbgOrigen.Properties.Appearance.BackColor2 = System.Drawing.Color.White;
			this.rbgOrigen.Properties.Appearance.GradientMode = System.Drawing.Drawing2D.LinearGradientMode.Vertical;
			this.rbgOrigen.Properties.Appearance.Options.UseBackColor = true;
			this.rbgOrigen.Properties.BorderStyle = DevExpress.XtraEditors.Controls.BorderStyles.NoBorder;
			this.rbgOrigen.Properties.Items.AddRange(new DevExpress.XtraEditors.Controls.RadioGroupItem[] {
																																																			new DevExpress.XtraEditors.Controls.RadioGroupItem(1, "Nacional"),
																																																			new DevExpress.XtraEditors.Controls.RadioGroupItem(2, "Importado")});
			this.rbgOrigen.Size = new System.Drawing.Size(232, 32);
			this.rbgOrigen.TabIndex = 25;
			this.rbgOrigen.KeyDown += new System.Windows.Forms.KeyEventHandler(this.rbgOrigen_KeyDown);
			this.rbgOrigen.SelectedIndexChanged += new System.EventHandler(this.rbgOrigen_SelectedIndexChanged);
			// 
			// label14
			// 
			this.label14.AutoSize = true;
			this.label14.BackColor = System.Drawing.Color.Transparent;
			this.label14.Location = new System.Drawing.Point(8, 67);
			this.label14.Name = "label14";
			this.label14.Size = new System.Drawing.Size(41, 16);
			this.label14.TabIndex = 102;
			this.label14.Text = "Arancel";
			this.label14.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// cmbArancel
			// 
			appearance10.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbArancel.Appearance = appearance10;
			this.cmbArancel.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbArancel.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbArancel.DataSource = this.ultraDataSource7;
			ultraGridBand2.AddButtonCaption = "ArticuloGrupo";
			ultraGridColumn3.Header.VisiblePosition = 2;
			ultraGridColumn3.Hidden = true;
			ultraGridColumn4.Header.VisiblePosition = 1;
			ultraGridColumn4.Width = 60;
			appearance11.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn5.CellAppearance = appearance11;
			ultraGridColumn5.Format = "#,##0.00";
			ultraGridColumn5.Header.Caption = "%";
			ultraGridColumn5.Header.VisiblePosition = 3;
			ultraGridColumn5.Width = 50;
			ultraGridColumn6.Header.VisiblePosition = 0;
			ultraGridColumn6.Width = 160;
			ultraGridBand2.Columns.AddRange(new object[] {
																										 ultraGridColumn3,
																										 ultraGridColumn4,
																										 ultraGridColumn5,
																										 ultraGridColumn6});
			this.cmbArancel.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			this.cmbArancel.DisplayMember = "Arancel";
			this.cmbArancel.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbArancel.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbArancel.Enabled = false;
			this.cmbArancel.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbArancel.Location = new System.Drawing.Point(56, 64);
			this.cmbArancel.Name = "cmbArancel";
			this.cmbArancel.Size = new System.Drawing.Size(160, 21);
			this.cmbArancel.TabIndex = 26;
			this.cmbArancel.ValueMember = "idArancel";
			this.cmbArancel.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cmbArancel_KeyDown);
			this.cmbArancel.Enter += new System.EventHandler(this.cmbArancel_Enter);
			// 
			// ultraDataSource7
			// 
			ultraDataColumn3.DataType = typeof(int);
			ultraDataColumn5.DataType = typeof(System.Double);
			ultraDataColumn5.DefaultValue = 0;
			this.ultraDataSource7.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn3,
																																 ultraDataColumn4,
																																 ultraDataColumn5,
																																 ultraDataColumn6});
			// 
			// btArancel
			// 
			this.btArancel.CausesValidation = false;
			this.btArancel.Image = ((System.Drawing.Image)(resources.GetObject("btArancel.Image")));
			this.btArancel.Location = new System.Drawing.Point(216, 64);
			this.btArancel.Name = "btArancel";
			this.btArancel.Size = new System.Drawing.Size(24, 22);
			this.btArancel.TabIndex = 27;
			this.btArancel.ToolTip = "Arancel de artículos";
			this.btArancel.Click += new System.EventHandler(this.btArancel_Click);
			// 
			// txtModelo
			// 
			appearance12.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtModelo.Appearance = appearance12;
			this.txtModelo.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtModelo.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtModelo.Enabled = false;
			this.txtModelo.Location = new System.Drawing.Point(112, 72);
			this.txtModelo.Name = "txtModelo";
			this.txtModelo.Size = new System.Drawing.Size(256, 22);
			this.txtModelo.TabIndex = 9;
			this.txtModelo.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtModelo_KeyDown);
			this.txtModelo.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtModelo_KeyPress);
			this.txtModelo.Validating += new System.ComponentModel.CancelEventHandler(this.txtModelo_Validating);
			this.txtModelo.Validated += new System.EventHandler(this.txtModelo_Validated);
			// 
			// label35
			// 
			this.label35.AutoSize = true;
			this.label35.BackColor = System.Drawing.Color.Transparent;
			this.label35.Location = new System.Drawing.Point(24, 75);
			this.label35.Name = "label35";
			this.label35.Size = new System.Drawing.Size(40, 16);
			this.label35.TabIndex = 142;
			this.label35.Text = "Modelo";
			this.label35.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label23
			// 
			this.label23.AutoSize = true;
			this.label23.BackColor = System.Drawing.Color.Transparent;
			this.label23.Location = new System.Drawing.Point(24, 42);
			this.label23.Name = "label23";
			this.label23.Size = new System.Drawing.Size(52, 16);
			this.label23.TabIndex = 140;
			this.label23.Text = "Categoría";
			this.label23.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// chkSerialAlterna
			// 
			appearance13.ForeColorDisabled = System.Drawing.Color.Black;
			this.chkSerialAlterna.Appearance = appearance13;
			this.chkSerialAlterna.Enabled = false;
			this.chkSerialAlterna.Location = new System.Drawing.Point(112, 36);
			this.chkSerialAlterna.Name = "chkSerialAlterna";
			this.chkSerialAlterna.Size = new System.Drawing.Size(88, 22);
			this.chkSerialAlterna.TabIndex = 50;
			this.chkSerialAlterna.Text = "Serial Alterna";
			this.chkSerialAlterna.KeyDown += new System.Windows.Forms.KeyEventHandler(this.chkSerialAlterna_KeyDown);
			this.chkSerialAlterna.CheckedChanged += new System.EventHandler(this.chkSerialAlterna_CheckedChanged);
			// 
			// chkManejaSerial
			// 
			appearance14.ForeColorDisabled = System.Drawing.Color.Black;
			this.chkManejaSerial.Appearance = appearance14;
			this.chkManejaSerial.Enabled = false;
			this.chkManejaSerial.Location = new System.Drawing.Point(8, 36);
			this.chkManejaSerial.Name = "chkManejaSerial";
			this.chkManejaSerial.Size = new System.Drawing.Size(88, 22);
			this.chkManejaSerial.TabIndex = 49;
			this.chkManejaSerial.Text = "Maneja Serial";
			this.chkManejaSerial.KeyDown += new System.Windows.Forms.KeyEventHandler(this.chkManejaSerial_KeyDown);
			// 
			// chkFaturable
			// 
			appearance15.ForeColorDisabled = System.Drawing.Color.Black;
			this.chkFaturable.Appearance = appearance15;
			this.chkFaturable.Enabled = false;
			this.chkFaturable.Location = new System.Drawing.Point(8, 77);
			this.chkFaturable.Name = "chkFaturable";
			this.chkFaturable.Size = new System.Drawing.Size(80, 22);
			this.chkFaturable.TabIndex = 51;
			this.chkFaturable.Text = "Facturable";
			this.chkFaturable.KeyDown += new System.Windows.Forms.KeyEventHandler(this.chkFaturable_KeyDown);
			// 
			// chkAplicDescuento
			// 
			appearance16.ForeColorDisabled = System.Drawing.Color.Black;
			this.chkAplicDescuento.Appearance = appearance16;
			this.chkAplicDescuento.Enabled = false;
			this.chkAplicDescuento.Location = new System.Drawing.Point(88, 99);
			this.chkAplicDescuento.Name = "chkAplicDescuento";
			this.chkAplicDescuento.Size = new System.Drawing.Size(112, 22);
			this.chkAplicDescuento.TabIndex = 52;
			this.chkAplicDescuento.Text = "Aplica Descuento";
			this.chkAplicDescuento.Visible = false;
			this.chkAplicDescuento.KeyDown += new System.Windows.Forms.KeyEventHandler(this.chkAplicDescuento_KeyDown);
			// 
			// txtNParte
			// 
			appearance17.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtNParte.Appearance = appearance17;
			this.txtNParte.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtNParte.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNParte.Enabled = false;
			this.txtNParte.Location = new System.Drawing.Point(472, 72);
			this.txtNParte.Name = "txtNParte";
			this.txtNParte.Size = new System.Drawing.Size(256, 22);
			this.txtNParte.TabIndex = 10;
			this.txtNParte.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtNParte_KeyDown);
			// 
			// label19
			// 
			this.label19.AutoSize = true;
			this.label19.BackColor = System.Drawing.Color.Transparent;
			this.label19.Location = new System.Drawing.Point(384, 75);
			this.label19.Name = "label19";
			this.label19.Size = new System.Drawing.Size(47, 16);
			this.label19.TabIndex = 131;
			this.label19.Text = "Nº Parte";
			this.label19.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label47
			// 
			this.label47.AutoSize = true;
			this.label47.BackColor = System.Drawing.Color.Transparent;
			this.label47.Location = new System.Drawing.Point(384, 42);
			this.label47.Name = "label47";
			this.label47.Size = new System.Drawing.Size(34, 16);
			this.label47.TabIndex = 128;
			this.label47.Text = "Marca";
			this.label47.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// cmbColor
			// 
			appearance18.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbColor.Appearance = appearance18;
			this.cmbColor.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.cmbColor.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbColor.DataSource = this.ultraDataSource6;
			this.cmbColor.DisplayLayout.AutoFitColumns = true;
			ultraGridBand3.AddButtonCaption = "ArticuloGrupo";
			ultraGridColumn7.Header.VisiblePosition = 1;
			ultraGridColumn7.Hidden = true;
			ultraGridColumn7.Width = 64;
			ultraGridColumn8.Header.VisiblePosition = 0;
			ultraGridColumn8.Width = 237;
			ultraGridBand3.Columns.AddRange(new object[] {
																										 ultraGridColumn7,
																										 ultraGridColumn8});
			this.cmbColor.DisplayLayout.BandsSerializer.Add(ultraGridBand3);
			this.cmbColor.DisplayMember = "Color";
			this.cmbColor.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbColor.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbColor.Enabled = false;
			this.cmbColor.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbColor.Location = new System.Drawing.Point(112, 104);
			this.cmbColor.Name = "cmbColor";
			this.cmbColor.Size = new System.Drawing.Size(256, 21);
			this.cmbColor.TabIndex = 11;
			this.cmbColor.ValueMember = "idColor";
			this.cmbColor.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cmbColor_KeyDown);
			this.cmbColor.Validating += new System.ComponentModel.CancelEventHandler(this.cmbColor_Validating);
			this.cmbColor.Validated += new System.EventHandler(this.cmbColor_Validated);
			this.cmbColor.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.cmbColor_InitializeLayout);
			this.cmbColor.Enter += new System.EventHandler(this.cmbColor_Enter);
			// 
			// ultraDataSource6
			// 
			ultraDataColumn7.DataType = typeof(int);
			this.ultraDataSource6.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn7,
																																 ultraDataColumn8});
			// 
			// label48
			// 
			this.label48.AutoSize = true;
			this.label48.BackColor = System.Drawing.Color.Transparent;
			this.label48.Location = new System.Drawing.Point(24, 106);
			this.label48.Name = "label48";
			this.label48.Size = new System.Drawing.Size(30, 16);
			this.label48.TabIndex = 129;
			this.label48.Text = "Color";
			this.label48.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// btColor
			// 
			this.btColor.CausesValidation = false;
			this.btColor.Image = ((System.Drawing.Image)(resources.GetObject("btColor.Image")));
			this.btColor.Location = new System.Drawing.Point(80, 103);
			this.btColor.Name = "btColor";
			this.btColor.Size = new System.Drawing.Size(24, 22);
			this.btColor.TabIndex = 12;
			this.btColor.ToolTip = "Color de artículos";
			this.btColor.Click += new System.EventHandler(this.btColor_Click);
			// 
			// ultraDataSource3
			// 
			ultraDataColumn9.DataType = typeof(int);
			this.ultraDataSource3.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn9,
																																 ultraDataColumn10});
			// 
			// txtSaldo
			// 
			appearance19.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtSaldo.Appearance = appearance19;
			this.txtSaldo.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtSaldo.Enabled = false;
			this.txtSaldo.FormatString = "#,##0.00";
			this.txtSaldo.Location = new System.Drawing.Point(112, 328);
			this.txtSaldo.Name = "txtSaldo";
			this.txtSaldo.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtSaldo.PromptChar = ' ';
			this.txtSaldo.Size = new System.Drawing.Size(72, 22);
			this.txtSaldo.TabIndex = 22;
			this.txtSaldo.Click += new System.EventHandler(this.txtSaldo_Click);
			this.txtSaldo.Enter += new System.EventHandler(this.txtSaldo_Enter);
			// 
			// label32
			// 
			this.label32.AutoSize = true;
			this.label32.BackColor = System.Drawing.Color.Transparent;
			this.label32.Location = new System.Drawing.Point(24, 328);
			this.label32.Name = "label32";
			this.label32.Size = new System.Drawing.Size(31, 16);
			this.label32.TabIndex = 109;
			this.label32.Text = "Saldo";
			this.label32.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// cmbArticuloSRI
			// 
			appearance20.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbArticuloSRI.Appearance = appearance20;
			this.cmbArticuloSRI.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbArticuloSRI.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbArticuloSRI.DataSource = this.ultraDataSource4;
			ultraGridBand4.AddButtonCaption = "ArticuloGrupo";
			ultraGridColumn9.Header.VisiblePosition = 0;
			ultraGridColumn9.Width = 113;
			ultraGridColumn10.Header.VisiblePosition = 2;
			ultraGridColumn10.Width = 100;
			ultraGridColumn11.Header.VisiblePosition = 1;
			ultraGridColumn11.Width = 350;
			appearance21.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn12.CellAppearance = appearance21;
			ultraGridColumn12.Format = "#,##0.00";
			ultraGridColumn12.Header.VisiblePosition = 3;
			ultraGridColumn12.Width = 100;
			ultraGridBand4.Columns.AddRange(new object[] {
																										 ultraGridColumn9,
																										 ultraGridColumn10,
																										 ultraGridColumn11,
																										 ultraGridColumn12});
			this.cmbArticuloSRI.DisplayLayout.BandsSerializer.Add(ultraGridBand4);
			this.cmbArticuloSRI.DisplayMember = "Detalle";
			this.cmbArticuloSRI.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbArticuloSRI.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbArticuloSRI.Enabled = false;
			this.cmbArticuloSRI.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbArticuloSRI.Location = new System.Drawing.Point(112, 264);
			this.cmbArticuloSRI.Name = "cmbArticuloSRI";
			this.cmbArticuloSRI.Size = new System.Drawing.Size(424, 21);
			this.cmbArticuloSRI.TabIndex = 16;
			this.cmbArticuloSRI.ValueMember = "idArticuloSRI";
			this.cmbArticuloSRI.Enter += new System.EventHandler(this.cmbArticuloSRI_Enter);
			// 
			// ultraDataSource4
			// 
			ultraDataColumn11.DataType = typeof(int);
			ultraDataColumn14.DataType = typeof(System.Double);
			ultraDataColumn14.DefaultValue = 0;
			this.ultraDataSource4.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn11,
																																 ultraDataColumn12,
																																 ultraDataColumn13,
																																 ultraDataColumn14});
			// 
			// label46
			// 
			this.label46.AutoSize = true;
			this.label46.BackColor = System.Drawing.Color.Transparent;
			this.label46.Location = new System.Drawing.Point(24, 264);
			this.label46.Name = "label46";
			this.label46.Size = new System.Drawing.Size(32, 16);
			this.label46.TabIndex = 106;
			this.label46.Text = "S.R.I.";
			this.label46.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtCodigo
			// 
			appearance22.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtCodigo.Appearance = appearance22;
			this.txtCodigo.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtCodigo.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtCodigo.Enabled = false;
			this.txtCodigo.Location = new System.Drawing.Point(112, 136);
			this.txtCodigo.Name = "txtCodigo";
			this.txtCodigo.Size = new System.Drawing.Size(256, 22);
			this.txtCodigo.TabIndex = 13;
			this.txtCodigo.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtCodigo_KeyDown);
			this.txtCodigo.ValueChanged += new System.EventHandler(this.txtCodigo_ValueChanged);
			// 
			// txtNombre
			// 
			appearance23.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtNombre.Appearance = appearance23;
			this.txtNombre.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtNombre.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNombre.Enabled = false;
			this.txtNombre.Location = new System.Drawing.Point(112, 168);
			this.txtNombre.MaxLength = 250;
			this.txtNombre.Name = "txtNombre";
			this.txtNombre.Size = new System.Drawing.Size(616, 22);
			this.txtNombre.TabIndex = 15;
			this.txtNombre.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtNombre_KeyDown);
			this.txtNombre.Validating += new System.ComponentModel.CancelEventHandler(this.txtNombre_Validating);
			this.txtNombre.ValueChanged += new System.EventHandler(this.txtNombre_ValueChanged);
			// 
			// txtIva
			// 
			appearance24.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtIva.Appearance = appearance24;
			this.txtIva.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtIva.Enabled = false;
			this.txtIva.FormatString = "#,##0.00";
			this.txtIva.Location = new System.Drawing.Point(648, 264);
			this.txtIva.Name = "txtIva";
			this.txtIva.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtIva.PromptChar = ' ';
			this.txtIva.Size = new System.Drawing.Size(80, 22);
			this.txtIva.TabIndex = 17;
			this.txtIva.Visible = false;
			this.txtIva.Click += new System.EventHandler(this.txtIva_Click);
			this.txtIva.Enter += new System.EventHandler(this.txtIva_Enter);
			// 
			// txtCantMaxima
			// 
			appearance25.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtCantMaxima.Appearance = appearance25;
			this.txtCantMaxima.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtCantMaxima.Enabled = false;
			this.txtCantMaxima.FormatString = "#,##0.00";
			this.txtCantMaxima.Location = new System.Drawing.Point(360, 360);
			this.txtCantMaxima.Name = "txtCantMaxima";
			this.txtCantMaxima.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtCantMaxima.PromptChar = ' ';
			this.txtCantMaxima.Size = new System.Drawing.Size(72, 22);
			this.txtCantMaxima.TabIndex = 24;
			this.txtCantMaxima.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtCantMaxima_KeyDown);
			this.txtCantMaxima.Click += new System.EventHandler(this.txtCantMaxima_Click);
			this.txtCantMaxima.Enter += new System.EventHandler(this.txtCantMaxima_Enter);
			// 
			// txtCantMinima
			// 
			appearance26.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtCantMinima.Appearance = appearance26;
			this.txtCantMinima.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtCantMinima.Enabled = false;
			this.txtCantMinima.FormatString = "#,##0.00";
			this.txtCantMinima.Location = new System.Drawing.Point(112, 360);
			this.txtCantMinima.Name = "txtCantMinima";
			this.txtCantMinima.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtCantMinima.PromptChar = ' ';
			this.txtCantMinima.Size = new System.Drawing.Size(72, 22);
			this.txtCantMinima.TabIndex = 23;
			this.txtCantMinima.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtCantMinima_KeyDown);
			this.txtCantMinima.Click += new System.EventHandler(this.txtCantMinima_Click);
			this.txtCantMinima.ValueChanged += new System.EventHandler(this.txtCantMinima_ValueChanged);
			this.txtCantMinima.Enter += new System.EventHandler(this.txtCantMinima_Enter);
			// 
			// cmbSubGrupo
			// 
			appearance27.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbSubGrupo.Appearance = appearance27;
			this.cmbSubGrupo.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.cmbSubGrupo.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbSubGrupo.DataSource = this.ultraDataSource5;
			this.cmbSubGrupo.DisplayLayout.AutoFitColumns = true;
			ultraGridColumn13.Header.VisiblePosition = 0;
			ultraGridColumn13.Hidden = true;
			ultraGridColumn14.Header.VisiblePosition = 1;
			ultraGridColumn14.Width = 237;
			ultraGridBand5.Columns.AddRange(new object[] {
																										 ultraGridColumn13,
																										 ultraGridColumn14});
			this.cmbSubGrupo.DisplayLayout.BandsSerializer.Add(ultraGridBand5);
			this.cmbSubGrupo.DisplayMember = "Subgrupo";
			this.cmbSubGrupo.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbSubGrupo.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbSubGrupo.Enabled = false;
			this.cmbSubGrupo.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbSubGrupo.Location = new System.Drawing.Point(472, 8);
			this.cmbSubGrupo.Name = "cmbSubGrupo";
			this.cmbSubGrupo.Size = new System.Drawing.Size(256, 21);
			this.cmbSubGrupo.TabIndex = 3;
			this.cmbSubGrupo.ValueMember = "idSubgrupo";
			this.cmbSubGrupo.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cmbSubGrupo_KeyDown);
			this.cmbSubGrupo.Validating += new System.ComponentModel.CancelEventHandler(this.cmbSubGrupo_Validating);
			this.cmbSubGrupo.Validated += new System.EventHandler(this.cmbSubGrupo_Validated);
			this.cmbSubGrupo.ValueChanged += new System.EventHandler(this.cmbSubGrupo_ValueChanged);
			this.cmbSubGrupo.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.cmbSubGrupo_InitializeLayout);
			this.cmbSubGrupo.Enter += new System.EventHandler(this.cmbSubGrupo_Enter);
			// 
			// ultraDataSource5
			// 
			ultraDataColumn15.DataType = typeof(int);
			this.ultraDataSource5.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn15,
																																 ultraDataColumn16});
			// 
			// cmbGrupo
			// 
			appearance28.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbGrupo.Appearance = appearance28;
			this.cmbGrupo.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.cmbGrupo.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbGrupo.DataSource = this.ultraDataSource1;
			ultraGridColumn15.Header.VisiblePosition = 0;
			ultraGridColumn15.Hidden = true;
			ultraGridColumn16.Header.VisiblePosition = 1;
			ultraGridColumn16.Width = 256;
			ultraGridBand6.Columns.AddRange(new object[] {
																										 ultraGridColumn15,
																										 ultraGridColumn16});
			this.cmbGrupo.DisplayLayout.BandsSerializer.Add(ultraGridBand6);
			this.cmbGrupo.DisplayLayout.Override.RowFilterAction = Infragistics.Win.UltraWinGrid.RowFilterAction.HideFilteredOutRows;
			this.cmbGrupo.DisplayLayout.Override.RowFilterMode = Infragistics.Win.UltraWinGrid.RowFilterMode.SiblingRowsOnly;
			this.cmbGrupo.DisplayLayout.Override.RowSizing = Infragistics.Win.UltraWinGrid.RowSizing.Fixed;
			this.cmbGrupo.DisplayLayout.ScrollStyle = Infragistics.Win.UltraWinGrid.ScrollStyle.Immediate;
			this.cmbGrupo.DisplayMember = "Grupo";
			this.cmbGrupo.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbGrupo.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbGrupo.Enabled = false;
			this.cmbGrupo.Location = new System.Drawing.Point(112, 8);
			this.cmbGrupo.Name = "cmbGrupo";
			this.cmbGrupo.Size = new System.Drawing.Size(256, 22);
			this.cmbGrupo.TabIndex = 1;
			this.cmbGrupo.ValueMember = "idGrupoArticulo";
			this.cmbGrupo.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cmbGrupo_KeyDown);
			this.cmbGrupo.Validating += new System.ComponentModel.CancelEventHandler(this.cmbGrupo_Validating);
			this.cmbGrupo.ValueChanged += new System.EventHandler(this.cmbGrupo_ValueChanged);
			this.cmbGrupo.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.cmbGrupo_InitializeLayout);
			this.cmbGrupo.Enter += new System.EventHandler(this.cmbGrupo_Enter);
			// 
			// ultraDataSource1
			// 
			ultraDataColumn17.DataType = typeof(int);
			this.ultraDataSource1.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn17,
																																 ultraDataColumn18});
			// 
			// btGrupo
			// 
			this.btGrupo.CausesValidation = false;
			this.btGrupo.Image = ((System.Drawing.Image)(resources.GetObject("btGrupo.Image")));
			this.btGrupo.Location = new System.Drawing.Point(80, 8);
			this.btGrupo.Name = "btGrupo";
			this.btGrupo.Size = new System.Drawing.Size(24, 22);
			this.btGrupo.TabIndex = 2;
			this.btGrupo.ToolTip = "Grupo de artículos";
			this.btGrupo.Click += new System.EventHandler(this.btGrupo_Click);
			// 
			// label13
			// 
			this.label13.AutoSize = true;
			this.label13.BackColor = System.Drawing.Color.Transparent;
			this.label13.Location = new System.Drawing.Point(24, 72);
			this.label13.Name = "label13";
			this.label13.Size = new System.Drawing.Size(38, 16);
			this.label13.TabIndex = 98;
			this.label13.Text = "Marca:";
			this.label13.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label12
			// 
			this.label12.AutoSize = true;
			this.label12.BackColor = System.Drawing.Color.Transparent;
			this.label12.Location = new System.Drawing.Point(384, 11);
			this.label12.Name = "label12";
			this.label12.Size = new System.Drawing.Size(53, 16);
			this.label12.TabIndex = 96;
			this.label12.Text = "SubGrupo";
			this.label12.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label9
			// 
			this.label9.AutoSize = true;
			this.label9.BackColor = System.Drawing.Color.Transparent;
			this.label9.Location = new System.Drawing.Point(552, 264);
			this.label9.Name = "label9";
			this.label9.Size = new System.Drawing.Size(22, 16);
			this.label9.TabIndex = 87;
			this.label9.Text = "IVA";
			this.label9.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.label9.Visible = false;
			// 
			// label6
			// 
			this.label6.AutoSize = true;
			this.label6.BackColor = System.Drawing.Color.Transparent;
			this.label6.Location = new System.Drawing.Point(208, 360);
			this.label6.Name = "label6";
			this.label6.Size = new System.Drawing.Size(90, 16);
			this.label6.TabIndex = 81;
			this.label6.Text = "Cantidad Máxima";
			this.label6.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label7
			// 
			this.label7.AutoSize = true;
			this.label7.BackColor = System.Drawing.Color.Transparent;
			this.label7.Location = new System.Drawing.Point(24, 360);
			this.label7.Name = "label7";
			this.label7.Size = new System.Drawing.Size(88, 16);
			this.label7.TabIndex = 77;
			this.label7.Text = "Cantidad mínima";
			this.label7.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label3
			// 
			this.label3.AutoSize = true;
			this.label3.BackColor = System.Drawing.Color.Transparent;
			this.label3.Location = new System.Drawing.Point(24, 171);
			this.label3.Name = "label3";
			this.label3.Size = new System.Drawing.Size(61, 16);
			this.label3.TabIndex = 68;
			this.label3.Text = "Descripción";
			this.label3.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label2
			// 
			this.label2.AutoSize = true;
			this.label2.BackColor = System.Drawing.Color.Transparent;
			this.label2.Location = new System.Drawing.Point(24, 11);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(35, 16);
			this.label2.TabIndex = 66;
			this.label2.Text = "Grupo";
			this.label2.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label1
			// 
			this.label1.AutoSize = true;
			this.label1.BackColor = System.Drawing.Color.Transparent;
			this.label1.Location = new System.Drawing.Point(24, 139);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(38, 16);
			this.label1.TabIndex = 64;
			this.label1.Text = "Código";
			this.label1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// lblCostoUltim
			// 
			this.lblCostoUltim.AutoSize = true;
			this.lblCostoUltim.BackColor = System.Drawing.Color.Transparent;
			this.lblCostoUltim.Location = new System.Drawing.Point(8, 144);
			this.lblCostoUltim.Name = "lblCostoUltim";
			this.lblCostoUltim.Size = new System.Drawing.Size(68, 16);
			this.lblCostoUltim.TabIndex = 112;
			this.lblCostoUltim.Text = "Costo Ultimo";
			this.lblCostoUltim.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtPrecioCostoUltimo
			// 
			appearance29.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtPrecioCostoUltimo.Appearance = appearance29;
			this.txtPrecioCostoUltimo.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtPrecioCostoUltimo.Enabled = false;
			this.txtPrecioCostoUltimo.FormatString = "#,##0.00";
			this.txtPrecioCostoUltimo.Location = new System.Drawing.Point(96, 144);
			this.txtPrecioCostoUltimo.Name = "txtPrecioCostoUltimo";
			this.txtPrecioCostoUltimo.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtPrecioCostoUltimo.PromptChar = ' ';
			this.txtPrecioCostoUltimo.Size = new System.Drawing.Size(80, 22);
			this.txtPrecioCostoUltimo.TabIndex = 39;
			this.txtPrecioCostoUltimo.Click += new System.EventHandler(this.txtPrecioCostoUltimo_Click);
			this.txtPrecioCostoUltimo.ValueChanged += new System.EventHandler(this.txtPrecioCostoUltimo_ValueChanged);
			this.txtPrecioCostoUltimo.Enter += new System.EventHandler(this.txtPrecioCostoUltimo_Enter);
			// 
			// txtPrecioPVP
			// 
			appearance30.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtPrecioPVP.Appearance = appearance30;
			this.txtPrecioPVP.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtPrecioPVP.Enabled = false;
			this.txtPrecioPVP.FormatString = "#,##0.00";
			this.txtPrecioPVP.Location = new System.Drawing.Point(96, 120);
			this.txtPrecioPVP.Name = "txtPrecioPVP";
			this.txtPrecioPVP.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtPrecioPVP.PromptChar = ' ';
			this.txtPrecioPVP.Size = new System.Drawing.Size(80, 22);
			this.txtPrecioPVP.TabIndex = 38;
			// 
			// txtPrecioTarjeta
			// 
			appearance31.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtPrecioTarjeta.Appearance = appearance31;
			this.txtPrecioTarjeta.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtPrecioTarjeta.Enabled = false;
			this.txtPrecioTarjeta.FormatString = "#,##0.00";
			this.txtPrecioTarjeta.Location = new System.Drawing.Point(96, 96);
			this.txtPrecioTarjeta.Name = "txtPrecioTarjeta";
			this.txtPrecioTarjeta.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtPrecioTarjeta.PromptChar = ' ';
			this.txtPrecioTarjeta.Size = new System.Drawing.Size(80, 22);
			this.txtPrecioTarjeta.TabIndex = 37;
			this.txtPrecioTarjeta.Validated += new System.EventHandler(this.txtPrecioTarjeta_Validated);
			this.txtPrecioTarjeta.ValueChanged += new System.EventHandler(this.txtPrecioTarjeta_ValueChanged);
			// 
			// txtPrecioPA
			// 
			appearance32.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtPrecioPA.Appearance = appearance32;
			this.txtPrecioPA.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtPrecioPA.Enabled = false;
			this.txtPrecioPA.FormatString = "#,##0.00";
			this.txtPrecioPA.Location = new System.Drawing.Point(96, 72);
			this.txtPrecioPA.MaskInput = "{LOC}nn,nnn,nnn.nn";
			this.txtPrecioPA.Name = "txtPrecioPA";
			this.txtPrecioPA.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtPrecioPA.PromptChar = ' ';
			this.txtPrecioPA.Size = new System.Drawing.Size(80, 22);
			this.txtPrecioPA.TabIndex = 35;
			this.txtPrecioPA.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtPrecioPA_KeyDown);
			this.txtPrecioPA.Click += new System.EventHandler(this.txtPrecioPA_Click);
			this.txtPrecioPA.Validating += new System.ComponentModel.CancelEventHandler(this.txtPrecioPA_Validating);
			this.txtPrecioPA.Validated += new System.EventHandler(this.txtPrecioPA_Validated);
			this.txtPrecioPA.ValueChanged += new System.EventHandler(this.txtPrecioPA_ValueChanged);
			this.txtPrecioPA.Enter += new System.EventHandler(this.txtPrecioPA_Enter);
			// 
			// txtPrecioPI
			// 
			appearance33.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtPrecioPI.Appearance = appearance33;
			this.txtPrecioPI.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtPrecioPI.Enabled = false;
			this.txtPrecioPI.FormatString = "#,##0.00";
			this.txtPrecioPI.Location = new System.Drawing.Point(96, 48);
			this.txtPrecioPI.MaskInput = "{LOC}nn,nnn,nnn.nn";
			this.txtPrecioPI.Name = "txtPrecioPI";
			this.txtPrecioPI.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtPrecioPI.PromptChar = ' ';
			this.txtPrecioPI.Size = new System.Drawing.Size(80, 22);
			this.txtPrecioPI.TabIndex = 33;
			this.txtPrecioPI.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtPrecioPI_KeyDown);
			this.txtPrecioPI.Click += new System.EventHandler(this.txtPrecioPI_Click);
			this.txtPrecioPI.Validating += new System.ComponentModel.CancelEventHandler(this.txtPrecioPI_Validating);
			this.txtPrecioPI.ValueChanged += new System.EventHandler(this.txtPrecioPI_ValueChanged);
			this.txtPrecioPI.Enter += new System.EventHandler(this.txtPrecioPI_Enter);
			// 
			// txtPrecioIncr2
			// 
			appearance34.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtPrecioIncr2.Appearance = appearance34;
			this.txtPrecioIncr2.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtPrecioIncr2.Enabled = false;
			this.txtPrecioIncr2.FormatString = "##0.00";
			this.txtPrecioIncr2.Location = new System.Drawing.Point(192, 88);
			this.txtPrecioIncr2.MaskInput = "nnn.nn";
			this.txtPrecioIncr2.Name = "txtPrecioIncr2";
			this.txtPrecioIncr2.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtPrecioIncr2.PromptChar = ' ';
			this.txtPrecioIncr2.Size = new System.Drawing.Size(72, 22);
			this.txtPrecioIncr2.TabIndex = 36;
			this.txtPrecioIncr2.Visible = false;
			this.txtPrecioIncr2.Click += new System.EventHandler(this.txtPrecioIncr2_Click);
			this.txtPrecioIncr2.Enter += new System.EventHandler(this.txtPrecioIncr2_Enter);
			// 
			// txtPrecioIncr1
			// 
			appearance35.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtPrecioIncr1.Appearance = appearance35;
			this.txtPrecioIncr1.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtPrecioIncr1.Enabled = false;
			this.txtPrecioIncr1.FormatString = "##0.00";
			this.txtPrecioIncr1.Location = new System.Drawing.Point(192, 64);
			this.txtPrecioIncr1.MaskInput = "nnn.nn";
			this.txtPrecioIncr1.Name = "txtPrecioIncr1";
			this.txtPrecioIncr1.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtPrecioIncr1.PromptChar = ' ';
			this.txtPrecioIncr1.Size = new System.Drawing.Size(72, 22);
			this.txtPrecioIncr1.TabIndex = 34;
			this.txtPrecioIncr1.Visible = false;
			this.txtPrecioIncr1.Click += new System.EventHandler(this.txtPrecioIncr1_Click);
			this.txtPrecioIncr1.Enter += new System.EventHandler(this.txtPrecioIncr1_Enter);
			// 
			// label40
			// 
			this.label40.AutoSize = true;
			this.label40.BackColor = System.Drawing.Color.Transparent;
			this.label40.Location = new System.Drawing.Point(8, 96);
			this.label40.Name = "label40";
			this.label40.Size = new System.Drawing.Size(64, 16);
			this.label40.TabIndex = 98;
			this.label40.Text = "Tarjeta - PA";
			this.label40.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label39
			// 
			this.label39.AutoSize = true;
			this.label39.BackColor = System.Drawing.Color.Transparent;
			this.label39.Location = new System.Drawing.Point(8, 120);
			this.label39.Name = "label39";
			this.label39.Size = new System.Drawing.Size(70, 16);
			this.label39.TabIndex = 97;
			this.label39.Text = "Crédito - PVP";
			this.label39.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label37
			// 
			this.label37.AutoSize = true;
			this.label37.BackColor = System.Drawing.Color.Transparent;
			this.label37.Location = new System.Drawing.Point(8, 72);
			this.label37.Name = "label37";
			this.label37.Size = new System.Drawing.Size(70, 16);
			this.label37.TabIndex = 95;
			this.label37.Text = "Contado - PA";
			this.label37.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label36
			// 
			this.label36.AutoSize = true;
			this.label36.BackColor = System.Drawing.Color.Transparent;
			this.label36.Location = new System.Drawing.Point(8, 48);
			this.label36.Name = "label36";
			this.label36.Size = new System.Drawing.Size(83, 16);
			this.label36.TabIndex = 94;
			this.label36.Text = "Distribuidor - PI";
			this.label36.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label15
			// 
			this.label15.AutoSize = true;
			this.label15.BackColor = System.Drawing.Color.Transparent;
			this.label15.Location = new System.Drawing.Point(192, 32);
			this.label15.Name = "label15";
			this.label15.Size = new System.Drawing.Size(76, 16);
			this.label15.TabIndex = 80;
			this.label15.Text = "% Incremento";
			this.label15.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.label15.Visible = false;
			// 
			// txtCodigoReemplazo
			// 
			appearance36.BorderColor = System.Drawing.Color.Black;
			appearance36.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtCodigoReemplazo.Appearance = appearance36;
			this.txtCodigoReemplazo.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtCodigoReemplazo.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtCodigoReemplazo.Enabled = false;
			this.txtCodigoReemplazo.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.txtCodigoReemplazo.Location = new System.Drawing.Point(472, 136);
			this.txtCodigoReemplazo.Name = "txtCodigoReemplazo";
			this.txtCodigoReemplazo.Size = new System.Drawing.Size(256, 22);
			this.txtCodigoReemplazo.TabIndex = 14;
			this.txtCodigoReemplazo.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtCodigoReemplazo_KeyDown);
			// 
			// chkDescCompras
			// 
			appearance37.ForeColorDisabled = System.Drawing.Color.Black;
			this.chkDescCompras.Appearance = appearance37;
			this.chkDescCompras.BackColor = System.Drawing.Color.Transparent;
			this.chkDescCompras.Enabled = false;
			this.chkDescCompras.Location = new System.Drawing.Point(8, 142);
			this.chkDescCompras.Name = "chkDescCompras";
			this.chkDescCompras.Size = new System.Drawing.Size(160, 22);
			this.chkDescCompras.TabIndex = 54;
			this.chkDescCompras.Text = "Descontinuado en Compras";
			this.chkDescCompras.KeyDown += new System.Windows.Forms.KeyEventHandler(this.chkDescCompras_KeyDown);
			this.chkDescCompras.CheckedChanged += new System.EventHandler(this.chkDescCompras_CheckedChanged);
			// 
			// chkDescVentas
			// 
			appearance38.ForeColorDisabled = System.Drawing.Color.Black;
			this.chkDescVentas.Appearance = appearance38;
			this.chkDescVentas.BackColor = System.Drawing.Color.Transparent;
			this.chkDescVentas.Enabled = false;
			this.chkDescVentas.Location = new System.Drawing.Point(8, 164);
			this.chkDescVentas.Name = "chkDescVentas";
			this.chkDescVentas.Size = new System.Drawing.Size(152, 21);
			this.chkDescVentas.TabIndex = 55;
			this.chkDescVentas.Text = "Descontinuado en Ventas";
			this.chkDescVentas.KeyDown += new System.Windows.Forms.KeyEventHandler(this.chkDescVentas_KeyDown);
			// 
			// chkDescontinuado
			// 
			appearance39.ForeColorDisabled = System.Drawing.Color.Black;
			this.chkDescontinuado.Appearance = appearance39;
			this.chkDescontinuado.BackColor = System.Drawing.Color.Transparent;
			this.chkDescontinuado.Enabled = false;
			this.chkDescontinuado.Location = new System.Drawing.Point(8, 121);
			this.chkDescontinuado.Name = "chkDescontinuado";
			this.chkDescontinuado.Size = new System.Drawing.Size(200, 21);
			this.chkDescontinuado.TabIndex = 53;
			this.chkDescontinuado.Text = "Descontinuado en toda transacción";
			this.chkDescontinuado.KeyDown += new System.Windows.Forms.KeyEventHandler(this.chkDescontinuado_KeyDown);
			// 
			// dtFechaUltimo
			// 
			appearance40.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtFechaUltimo.Appearance = appearance40;
			dateButton1.Caption = "Today";
			this.dtFechaUltimo.DateButtons.Add(dateButton1);
			this.dtFechaUltimo.Enabled = false;
			this.dtFechaUltimo.Format = "dd/MM/yyyy";
			this.dtFechaUltimo.Location = new System.Drawing.Point(96, 160);
			this.dtFechaUltimo.Name = "dtFechaUltimo";
			this.dtFechaUltimo.NonAutoSizeHeight = 24;
			this.dtFechaUltimo.Size = new System.Drawing.Size(96, 21);
			this.dtFechaUltimo.TabIndex = 47;
			this.dtFechaUltimo.Value = ((object)(resources.GetObject("dtFechaUltimo.Value")));
			this.dtFechaUltimo.KeyDown += new System.Windows.Forms.KeyEventHandler(this.dtFechaUltimo_KeyDown);
			// 
			// label20
			// 
			this.label20.AutoSize = true;
			this.label20.BackColor = System.Drawing.Color.Transparent;
			this.label20.Location = new System.Drawing.Point(8, 160);
			this.label20.Name = "label20";
			this.label20.Size = new System.Drawing.Size(68, 16);
			this.label20.TabIndex = 82;
			this.label20.Text = "Fecha último";
			this.label20.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtDepreciacion
			// 
			appearance41.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtDepreciacion.Appearance = appearance41;
			this.txtDepreciacion.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtDepreciacion.Enabled = false;
			this.txtDepreciacion.FormatString = "#,##0.00";
			this.txtDepreciacion.Location = new System.Drawing.Point(96, 88);
			this.txtDepreciacion.Name = "txtDepreciacion";
			this.txtDepreciacion.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtDepreciacion.PromptChar = ' ';
			this.txtDepreciacion.Size = new System.Drawing.Size(80, 22);
			this.txtDepreciacion.TabIndex = 44;
			this.txtDepreciacion.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtDepreciacion_KeyDown);
			this.txtDepreciacion.Click += new System.EventHandler(this.txtDepreciacion_Click);
			this.txtDepreciacion.Enter += new System.EventHandler(this.txtDepreciacion_Enter);
			// 
			// txtTiempo
			// 
			appearance42.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtTiempo.Appearance = appearance42;
			this.txtTiempo.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtTiempo.Enabled = false;
			this.txtTiempo.FormatString = "#,##0.00";
			this.txtTiempo.Location = new System.Drawing.Point(96, 136);
			this.txtTiempo.Name = "txtTiempo";
			this.txtTiempo.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtTiempo.PromptChar = ' ';
			this.txtTiempo.Size = new System.Drawing.Size(80, 22);
			this.txtTiempo.TabIndex = 46;
			this.txtTiempo.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtTiempo_KeyDown);
			this.txtTiempo.Click += new System.EventHandler(this.txtTiempo_Click);
			this.txtTiempo.Enter += new System.EventHandler(this.txtTiempo_Enter);
			// 
			// cmbSeccion
			// 
			appearance43.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbSeccion.Appearance = appearance43;
			this.cmbSeccion.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbSeccion.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbSeccion.DataSource = this.ultraDataSource8;
			this.cmbSeccion.DisplayLayout.AutoFitColumns = true;
			ultraGridBand7.AddButtonCaption = "ArticuloGrupo";
			ultraGridColumn17.Header.VisiblePosition = 1;
			ultraGridColumn17.Hidden = true;
			ultraGridColumn17.Width = 226;
			ultraGridColumn18.Header.VisiblePosition = 0;
			ultraGridColumn18.Width = 181;
			ultraGridBand7.Columns.AddRange(new object[] {
																										 ultraGridColumn17,
																										 ultraGridColumn18});
			this.cmbSeccion.DisplayLayout.BandsSerializer.Add(ultraGridBand7);
			this.cmbSeccion.DisplayMember = "Seccion";
			this.cmbSeccion.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbSeccion.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbSeccion.Enabled = false;
			this.cmbSeccion.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbSeccion.Location = new System.Drawing.Point(96, 40);
			this.cmbSeccion.Name = "cmbSeccion";
			this.cmbSeccion.Size = new System.Drawing.Size(200, 21);
			this.cmbSeccion.TabIndex = 42;
			this.cmbSeccion.ValueMember = "idSeccion";
			this.cmbSeccion.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cmbSeccion_KeyDown);
			this.cmbSeccion.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.cmbSeccion_InitializeLayout);
			this.cmbSeccion.Enter += new System.EventHandler(this.cmbSeccion_Enter);
			// 
			// ultraDataSource8
			// 
			ultraDataColumn19.DataType = typeof(int);
			this.ultraDataSource8.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn19,
																																 ultraDataColumn20});
			// 
			// label31
			// 
			this.label31.AutoSize = true;
			this.label31.BackColor = System.Drawing.Color.Transparent;
			this.label31.Location = new System.Drawing.Point(8, 136);
			this.label31.Name = "label31";
			this.label31.Size = new System.Drawing.Size(41, 16);
			this.label31.TabIndex = 73;
			this.label31.Text = "Tiempo";
			this.label31.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// btSeccion
			// 
			this.btSeccion.CausesValidation = false;
			this.btSeccion.Image = ((System.Drawing.Image)(resources.GetObject("btSeccion.Image")));
			this.btSeccion.Location = new System.Drawing.Point(72, 40);
			this.btSeccion.Name = "btSeccion";
			this.btSeccion.Size = new System.Drawing.Size(24, 21);
			this.btSeccion.TabIndex = 41;
			this.btSeccion.ToolTip = "Seccion de artículos";
			this.btSeccion.Click += new System.EventHandler(this.btSeccion_Click);
			// 
			// label29
			// 
			this.label29.AutoSize = true;
			this.label29.BackColor = System.Drawing.Color.Transparent;
			this.label29.Location = new System.Drawing.Point(8, 40);
			this.label29.Name = "label29";
			this.label29.Size = new System.Drawing.Size(42, 16);
			this.label29.TabIndex = 71;
			this.label29.Text = "Sección";
			this.label29.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// lblValorResid
			// 
			this.lblValorResid.AutoSize = true;
			this.lblValorResid.BackColor = System.Drawing.Color.Transparent;
			this.lblValorResid.Location = new System.Drawing.Point(8, 112);
			this.lblValorResid.Name = "lblValorResid";
			this.lblValorResid.Size = new System.Drawing.Size(75, 16);
			this.lblValorResid.TabIndex = 70;
			this.lblValorResid.Text = "Valor Residual";
			this.lblValorResid.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// lblDeprecia
			// 
			this.lblDeprecia.AutoSize = true;
			this.lblDeprecia.BackColor = System.Drawing.Color.Transparent;
			this.lblDeprecia.Location = new System.Drawing.Point(8, 88);
			this.lblDeprecia.Name = "lblDeprecia";
			this.lblDeprecia.Size = new System.Drawing.Size(83, 16);
			this.lblDeprecia.TabIndex = 67;
			this.lblDeprecia.Text = "% Depreciación";
			this.lblDeprecia.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label4
			// 
			this.label4.AutoSize = true;
			this.label4.BackColor = System.Drawing.Color.Transparent;
			this.label4.Location = new System.Drawing.Point(384, 139);
			this.label4.Name = "label4";
			this.label4.Size = new System.Drawing.Size(59, 16);
			this.label4.TabIndex = 159;
			this.label4.Text = "Reemplazo";
			this.label4.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label17
			// 
			this.label17.AutoSize = true;
			this.label17.BackColor = System.Drawing.Color.Transparent;
			this.label17.Location = new System.Drawing.Point(8, 64);
			this.label17.Name = "label17";
			this.label17.Size = new System.Drawing.Size(85, 16);
			this.label17.TabIndex = 112;
			this.label17.Text = "Centro de Costo";
			this.label17.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// groupBox2
			// 
			this.groupBox2.Controls.Add(this.label38);
			this.groupBox2.Controls.Add(this.txtPrecioDescEmpleado);
			this.groupBox2.Controls.Add(this.label34);
			this.groupBox2.Controls.Add(this.txtPrecioAutoconsumo);
			this.groupBox2.Controls.Add(this.label33);
			this.groupBox2.Controls.Add(this.txtCostoPayJoy);
			this.groupBox2.Controls.Add(this.label21);
			this.groupBox2.Controls.Add(this.txtPrecioTCsinIntereses);
			this.groupBox2.Controls.Add(this.txtPrecioIncr1);
			this.groupBox2.Controls.Add(this.label15);
			this.groupBox2.Controls.Add(this.label36);
			this.groupBox2.Controls.Add(this.txtPrecioPI);
			this.groupBox2.Controls.Add(this.txtPrecioIncr2);
			this.groupBox2.Controls.Add(this.label37);
			this.groupBox2.Controls.Add(this.txtPrecioPA);
			this.groupBox2.Controls.Add(this.label40);
			this.groupBox2.Controls.Add(this.label39);
			this.groupBox2.Controls.Add(this.txtPrecioPVP);
			this.groupBox2.Controls.Add(this.txtPrecioTarjeta);
			this.groupBox2.Controls.Add(this.lblCostoUltim);
			this.groupBox2.Controls.Add(this.txtPrecioCostoUltimo);
			this.groupBox2.Controls.Add(this.chkPrecio);
			this.groupBox2.Controls.Add(this.chkDepartamento);
			this.groupBox2.Controls.Add(this.chkDepartamento1);
			this.groupBox2.Location = new System.Drawing.Point(248, 400);
			this.groupBox2.Name = "groupBox2";
			this.groupBox2.Size = new System.Drawing.Size(360, 224);
			this.groupBox2.TabIndex = 160;
			this.groupBox2.TabStop = false;
			this.groupBox2.Text = "Precios";
			this.groupBox2.Enter += new System.EventHandler(this.groupBox2_Enter);
			// 
			// label38
			// 
			this.label38.AutoSize = true;
			this.label38.BackColor = System.Drawing.Color.Transparent;
			this.label38.Location = new System.Drawing.Point(184, 147);
			this.label38.Name = "label38";
			this.label38.Size = new System.Drawing.Size(84, 16);
			this.label38.TabIndex = 120;
			this.label38.Text = "Desc. Empleado";
			this.label38.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.label38.Visible = false;
			// 
			// txtPrecioDescEmpleado
			// 
			appearance44.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtPrecioDescEmpleado.Appearance = appearance44;
			this.txtPrecioDescEmpleado.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtPrecioDescEmpleado.Enabled = false;
			this.txtPrecioDescEmpleado.FormatString = "#,##0.00";
			this.txtPrecioDescEmpleado.Location = new System.Drawing.Point(272, 144);
			this.txtPrecioDescEmpleado.Name = "txtPrecioDescEmpleado";
			this.txtPrecioDescEmpleado.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtPrecioDescEmpleado.PromptChar = ' ';
			this.txtPrecioDescEmpleado.Size = new System.Drawing.Size(80, 22);
			this.txtPrecioDescEmpleado.TabIndex = 119;
			this.txtPrecioDescEmpleado.Visible = false;
			// 
			// label34
			// 
			this.label34.AutoSize = true;
			this.label34.BackColor = System.Drawing.Color.Transparent;
			this.label34.Location = new System.Drawing.Point(184, 171);
			this.label34.Name = "label34";
			this.label34.Size = new System.Drawing.Size(72, 16);
			this.label34.TabIndex = 118;
			this.label34.Text = "Autoconsumo";
			this.label34.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.label34.Visible = false;
			// 
			// txtPrecioAutoconsumo
			// 
			appearance45.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtPrecioAutoconsumo.Appearance = appearance45;
			this.txtPrecioAutoconsumo.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtPrecioAutoconsumo.Enabled = false;
			this.txtPrecioAutoconsumo.FormatString = "#,##0.00";
			this.txtPrecioAutoconsumo.Location = new System.Drawing.Point(272, 168);
			this.txtPrecioAutoconsumo.Name = "txtPrecioAutoconsumo";
			this.txtPrecioAutoconsumo.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtPrecioAutoconsumo.PromptChar = ' ';
			this.txtPrecioAutoconsumo.Size = new System.Drawing.Size(80, 22);
			this.txtPrecioAutoconsumo.TabIndex = 117;
			this.txtPrecioAutoconsumo.Visible = false;
			// 
			// label33
			// 
			this.label33.AutoSize = true;
			this.label33.BackColor = System.Drawing.Color.Transparent;
			this.label33.Location = new System.Drawing.Point(8, 192);
			this.label33.Name = "label33";
			this.label33.Size = new System.Drawing.Size(69, 16);
			this.label33.TabIndex = 116;
			this.label33.Text = "PayJoy - PVP";
			this.label33.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.label33.Click += new System.EventHandler(this.label33_Click);
			// 
			// txtCostoPayJoy
			// 
			appearance46.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtCostoPayJoy.Appearance = appearance46;
			this.txtCostoPayJoy.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtCostoPayJoy.Enabled = false;
			this.txtCostoPayJoy.FormatString = "#,##0.00";
			this.txtCostoPayJoy.Location = new System.Drawing.Point(96, 192);
			this.txtCostoPayJoy.Name = "txtCostoPayJoy";
			this.txtCostoPayJoy.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtCostoPayJoy.PromptChar = ' ';
			this.txtCostoPayJoy.Size = new System.Drawing.Size(80, 22);
			this.txtCostoPayJoy.TabIndex = 115;
			// 
			// label21
			// 
			this.label21.AutoSize = true;
			this.label21.BackColor = System.Drawing.Color.Transparent;
			this.label21.Location = new System.Drawing.Point(8, 168);
			this.label21.Name = "label21";
			this.label21.Size = new System.Drawing.Size(58, 16);
			this.label21.TabIndex = 114;
			this.label21.Text = "Tarjeta- SI";
			this.label21.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtPrecioTCsinIntereses
			// 
			appearance47.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtPrecioTCsinIntereses.Appearance = appearance47;
			this.txtPrecioTCsinIntereses.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtPrecioTCsinIntereses.Enabled = false;
			this.txtPrecioTCsinIntereses.FormatString = "#,##0.00";
			this.txtPrecioTCsinIntereses.Location = new System.Drawing.Point(96, 168);
			this.txtPrecioTCsinIntereses.Name = "txtPrecioTCsinIntereses";
			this.txtPrecioTCsinIntereses.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtPrecioTCsinIntereses.PromptChar = ' ';
			this.txtPrecioTCsinIntereses.Size = new System.Drawing.Size(80, 22);
			this.txtPrecioTCsinIntereses.TabIndex = 113;
			this.txtPrecioTCsinIntereses.Click += new System.EventHandler(this.txtPrecioTCsinIntereses_Click);
			this.txtPrecioTCsinIntereses.Enter += new System.EventHandler(this.txtPrecioTCsinIntereses_Enter);
			// 
			// chkPrecio
			// 
			appearance48.ForeColorDisabled = System.Drawing.Color.Black;
			this.chkPrecio.Appearance = appearance48;
			this.chkPrecio.BackColor = System.Drawing.Color.Transparent;
			this.chkPrecio.CausesValidation = false;
			this.chkPrecio.Enabled = false;
			this.chkPrecio.Location = new System.Drawing.Point(8, 22);
			this.chkPrecio.Name = "chkPrecio";
			this.chkPrecio.Size = new System.Drawing.Size(96, 21);
			this.chkPrecio.TabIndex = 32;
			this.chkPrecio.Text = "Precio Manual";
			this.chkPrecio.KeyDown += new System.Windows.Forms.KeyEventHandler(this.chkPrecio_KeyDown);
			this.chkPrecio.CheckedChanged += new System.EventHandler(this.chkPrecio_CheckedChanged);
			// 
			// chkDepartamento
			// 
			appearance49.ForeColorDisabled = System.Drawing.Color.Black;
			this.chkDepartamento.Appearance = appearance49;
			this.chkDepartamento.BackColor = System.Drawing.Color.Transparent;
			this.chkDepartamento.Enabled = false;
			this.chkDepartamento.Location = new System.Drawing.Point(184, 120);
			this.chkDepartamento.Name = "chkDepartamento";
			this.chkDepartamento.Size = new System.Drawing.Size(48, 22);
			this.chkDepartamento.TabIndex = 311;
			this.chkDepartamento.Text = "Caja";
			this.chkDepartamento.Visible = false;
			// 
			// chkDepartamento1
			// 
			appearance50.ForeColorDisabled = System.Drawing.Color.Black;
			this.chkDepartamento1.Appearance = appearance50;
			this.chkDepartamento1.BackColor = System.Drawing.Color.Transparent;
			this.chkDepartamento1.Enabled = false;
			this.chkDepartamento1.Location = new System.Drawing.Point(240, 120);
			this.chkDepartamento1.Name = "chkDepartamento1";
			this.chkDepartamento1.Size = new System.Drawing.Size(87, 22);
			this.chkDepartamento1.TabIndex = 312;
			this.chkDepartamento1.Text = "Contabilidad";
			this.chkDepartamento1.Visible = false;
			// 
			// groupBox3
			// 
			this.groupBox3.Controls.Add(this.txtResidual);
			this.groupBox3.Controls.Add(this.cmbCentroCosto);
			this.groupBox3.Controls.Add(this.label29);
			this.groupBox3.Controls.Add(this.btSeccion);
			this.groupBox3.Controls.Add(this.cmbSeccion);
			this.groupBox3.Controls.Add(this.label17);
			this.groupBox3.Controls.Add(this.lblDeprecia);
			this.groupBox3.Controls.Add(this.txtDepreciacion);
			this.groupBox3.Controls.Add(this.lblValorResid);
			this.groupBox3.Controls.Add(this.label31);
			this.groupBox3.Controls.Add(this.txtTiempo);
			this.groupBox3.Controls.Add(this.label20);
			this.groupBox3.Controls.Add(this.dtFechaUltimo);
			this.groupBox3.Location = new System.Drawing.Point(608, 400);
			this.groupBox3.Name = "groupBox3";
			this.groupBox3.Size = new System.Drawing.Size(304, 224);
			this.groupBox3.TabIndex = 163;
			this.groupBox3.TabStop = false;
			this.groupBox3.Text = "Activo Fijo";
			this.groupBox3.Enter += new System.EventHandler(this.groupBox3_Enter);
			// 
			// txtResidual
			// 
			appearance51.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtResidual.Appearance = appearance51;
			this.txtResidual.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtResidual.Enabled = false;
			this.txtResidual.FormatString = "#,##0.00";
			this.txtResidual.Location = new System.Drawing.Point(96, 112);
			this.txtResidual.Name = "txtResidual";
			this.txtResidual.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtResidual.PromptChar = ' ';
			this.txtResidual.Size = new System.Drawing.Size(80, 22);
			this.txtResidual.TabIndex = 45;
			this.txtResidual.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtResidual_KeyDown);
			this.txtResidual.Click += new System.EventHandler(this.txtResidual_Click);
			this.txtResidual.Enter += new System.EventHandler(this.txtResidual_Enter);
			// 
			// cmbCentroCosto
			// 
			appearance52.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbCentroCosto.Appearance = appearance52;
			this.cmbCentroCosto.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbCentroCosto.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbCentroCosto.DisplayMember = "Articulo";
			this.cmbCentroCosto.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbCentroCosto.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbCentroCosto.Enabled = false;
			this.cmbCentroCosto.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbCentroCosto.Location = new System.Drawing.Point(96, 64);
			this.cmbCentroCosto.Name = "cmbCentroCosto";
			this.cmbCentroCosto.Size = new System.Drawing.Size(200, 21);
			this.cmbCentroCosto.TabIndex = 43;
			this.cmbCentroCosto.ValueMember = "idArticulo";
			this.cmbCentroCosto.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cmbCentroCosto_KeyDown);
			this.cmbCentroCosto.Enter += new System.EventHandler(this.cmbCentroCosto_Enter);
			// 
			// groupBox4
			// 
			this.groupBox4.Location = new System.Drawing.Point(0, 624);
			this.groupBox4.Name = "groupBox4";
			this.groupBox4.Size = new System.Drawing.Size(1168, 8);
			this.groupBox4.TabIndex = 164;
			this.groupBox4.TabStop = false;
			// 
			// groupBox5
			// 
			this.groupBox5.Controls.Add(this.chkEfectivo);
			this.groupBox5.Controls.Add(this.chkcatalogo);
			this.groupBox5.Controls.Add(this.chkObsequios);
			this.groupBox5.Controls.Add(this.chkImpuestoIVA);
			this.groupBox5.Controls.Add(this.chkInventario);
			this.groupBox5.Controls.Add(this.chkTransferible);
			this.groupBox5.Controls.Add(this.chkCombo);
			this.groupBox5.Controls.Add(this.chkSerialAlterna);
			this.groupBox5.Controls.Add(this.chkManejaSerial);
			this.groupBox5.Controls.Add(this.chkFaturable);
			this.groupBox5.Controls.Add(this.chkAplicDescuento);
			this.groupBox5.Controls.Add(this.chkDescontinuado);
			this.groupBox5.Controls.Add(this.chkDescVentas);
			this.groupBox5.Controls.Add(this.chkDescCompras);
			this.groupBox5.Controls.Add(this.chkPrecioMovil);
			this.groupBox5.Controls.Add(this.chkProductoWeb);
			this.groupBox5.Location = new System.Drawing.Point(912, 400);
			this.groupBox5.Name = "groupBox5";
			this.groupBox5.Size = new System.Drawing.Size(208, 224);
			this.groupBox5.TabIndex = 173;
			this.groupBox5.TabStop = false;
			// 
			// chkEfectivo
			// 
			appearance53.ForeColorDisabled = System.Drawing.Color.Black;
			this.chkEfectivo.Appearance = appearance53;
			this.chkEfectivo.BackColor = System.Drawing.Color.Transparent;
			this.chkEfectivo.Enabled = false;
			this.chkEfectivo.Location = new System.Drawing.Point(88, 184);
			this.chkEfectivo.Name = "chkEfectivo";
			this.chkEfectivo.Size = new System.Drawing.Size(152, 21);
			this.chkEfectivo.TabIndex = 282;
			this.chkEfectivo.Text = "Efectivo-Tarjeta";
			// 
			// chkcatalogo
			// 
			appearance54.ForeColorDisabled = System.Drawing.Color.Black;
			this.chkcatalogo.Appearance = appearance54;
			this.chkcatalogo.BackColor = System.Drawing.Color.Transparent;
			this.chkcatalogo.Enabled = false;
			this.chkcatalogo.Location = new System.Drawing.Point(88, 104);
			this.chkcatalogo.Name = "chkcatalogo";
			this.chkcatalogo.Size = new System.Drawing.Size(120, 22);
			this.chkcatalogo.TabIndex = 281;
			this.chkcatalogo.Text = "Aplica Catálogo";
			this.chkcatalogo.Visible = false;
			// 
			// chkObsequios
			// 
			appearance55.ForeColorDisabled = System.Drawing.Color.Black;
			this.chkObsequios.Appearance = appearance55;
			this.chkObsequios.Enabled = false;
			this.chkObsequios.Location = new System.Drawing.Point(112, 16);
			this.chkObsequios.Name = "chkObsequios";
			this.chkObsequios.Size = new System.Drawing.Size(88, 22);
			this.chkObsequios.TabIndex = 59;
			this.chkObsequios.Text = "Obsequios";
			this.chkObsequios.CheckedChanged += new System.EventHandler(this.chkObsequios_CheckedChanged);
			// 
			// chkImpuestoIVA
			// 
			appearance56.ForeColorDisabled = System.Drawing.Color.Black;
			this.chkImpuestoIVA.Appearance = appearance56;
			this.chkImpuestoIVA.Enabled = false;
			this.chkImpuestoIVA.Location = new System.Drawing.Point(8, 56);
			this.chkImpuestoIVA.Name = "chkImpuestoIVA";
			this.chkImpuestoIVA.Size = new System.Drawing.Size(104, 22);
			this.chkImpuestoIVA.TabIndex = 58;
			this.chkImpuestoIVA.Text = "Impuesto I.V.A.";
			// 
			// chkInventario
			// 
			appearance57.ForeColorDisabled = System.Drawing.Color.Black;
			this.chkInventario.Appearance = appearance57;
			this.chkInventario.BackColor = System.Drawing.Color.Transparent;
			this.chkInventario.Enabled = false;
			this.chkInventario.Location = new System.Drawing.Point(8, 184);
			this.chkInventario.Name = "chkInventario";
			this.chkInventario.Size = new System.Drawing.Size(72, 21);
			this.chkInventario.TabIndex = 57;
			this.chkInventario.Text = "Inventario";
			// 
			// chkTransferible
			// 
			appearance58.ForeColorDisabled = System.Drawing.Color.Black;
			this.chkTransferible.Appearance = appearance58;
			this.chkTransferible.Enabled = false;
			this.chkTransferible.Location = new System.Drawing.Point(8, 99);
			this.chkTransferible.Name = "chkTransferible";
			this.chkTransferible.Size = new System.Drawing.Size(80, 22);
			this.chkTransferible.TabIndex = 56;
			this.chkTransferible.Text = "Transferible";
			// 
			// chkPrecioMovil
			// 
			appearance59.ForeColorDisabled = System.Drawing.Color.Black;
			this.chkPrecioMovil.Appearance = appearance59;
			this.chkPrecioMovil.BackColor = System.Drawing.Color.Transparent;
			this.chkPrecioMovil.Enabled = false;
			this.chkPrecioMovil.Location = new System.Drawing.Point(112, 56);
			this.chkPrecioMovil.Name = "chkPrecioMovil";
			this.chkPrecioMovil.Size = new System.Drawing.Size(88, 22);
			this.chkPrecioMovil.TabIndex = 310;
			this.chkPrecioMovil.Text = "Point  Movil";
			// 
			// chkProductoWeb
			// 
			appearance60.ForeColorDisabled = System.Drawing.Color.Black;
			this.chkProductoWeb.Appearance = appearance60;
			this.chkProductoWeb.BackColor = System.Drawing.Color.Transparent;
			this.chkProductoWeb.Enabled = false;
			this.chkProductoWeb.Location = new System.Drawing.Point(112, 77);
			this.chkProductoWeb.Name = "chkProductoWeb";
			this.chkProductoWeb.Size = new System.Drawing.Size(104, 19);
			this.chkProductoWeb.TabIndex = 289;
			this.chkProductoWeb.Text = "Producto Web";
			this.chkProductoWeb.CheckedChanged += new System.EventHandler(this.chkProductoWeb_CheckedChanged);
			// 
			// btnNuevo
			// 
			this.btnNuevo.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnNuevo.CausesValidation = false;
			this.btnNuevo.Image = ((System.Drawing.Image)(resources.GetObject("btnNuevo.Image")));
			this.btnNuevo.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnNuevo.Location = new System.Drawing.Point(8, 641);
			this.btnNuevo.Name = "btnNuevo";
			this.btnNuevo.Size = new System.Drawing.Size(72, 22);
			this.btnNuevo.TabIndex = 0;
			this.btnNuevo.Text = "&Nuevo";
			this.btnNuevo.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnNuevo.Click += new System.EventHandler(this.btnNuevo_Click);
			// 
			// btnConsultar
			// 
			this.btnConsultar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnConsultar.CausesValidation = false;
			this.btnConsultar.Image = ((System.Drawing.Image)(resources.GetObject("btnConsultar.Image")));
			this.btnConsultar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnConsultar.Location = new System.Drawing.Point(84, 641);
			this.btnConsultar.Name = "btnConsultar";
			this.btnConsultar.Size = new System.Drawing.Size(78, 22);
			this.btnConsultar.TabIndex = 175;
			this.btnConsultar.Text = "&Consultar";
			this.btnConsultar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnConsultar.Click += new System.EventHandler(this.btnConsultar_Click);
			// 
			// btnEditar
			// 
			this.btnEditar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnEditar.CausesValidation = false;
			this.btnEditar.Enabled = false;
			this.btnEditar.Image = ((System.Drawing.Image)(resources.GetObject("btnEditar.Image")));
			this.btnEditar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnEditar.Location = new System.Drawing.Point(166, 641);
			this.btnEditar.Name = "btnEditar";
			this.btnEditar.Size = new System.Drawing.Size(76, 22);
			this.btnEditar.TabIndex = 178;
			this.btnEditar.Text = "&Editar";
			this.btnEditar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnEditar.Click += new System.EventHandler(this.btnEditar_Click);
			// 
			// btnImprimir
			// 
			this.btnImprimir.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnImprimir.CausesValidation = false;
			this.btnImprimir.Enabled = false;
			this.btnImprimir.Image = ((System.Drawing.Image)(resources.GetObject("btnImprimir.Image")));
			this.btnImprimir.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnImprimir.Location = new System.Drawing.Point(246, 641);
			this.btnImprimir.Name = "btnImprimir";
			this.btnImprimir.Size = new System.Drawing.Size(77, 22);
			this.btnImprimir.TabIndex = 177;
			this.btnImprimir.Text = "&Imprimir";
			this.btnImprimir.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnImprimir.Click += new System.EventHandler(this.btnImprimir_Click);
			// 
			// btnGuardar
			// 
			this.btnGuardar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnGuardar.Enabled = false;
			this.btnGuardar.Image = ((System.Drawing.Image)(resources.GetObject("btnGuardar.Image")));
			this.btnGuardar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnGuardar.Location = new System.Drawing.Point(327, 641);
			this.btnGuardar.Name = "btnGuardar";
			this.btnGuardar.Size = new System.Drawing.Size(76, 22);
			this.btnGuardar.TabIndex = 176;
			this.btnGuardar.Text = "&Guardar";
			this.btnGuardar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnGuardar.Click += new System.EventHandler(this.btnGuardar_Click);
			// 
			// btnCancelar
			// 
			this.btnCancelar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnCancelar.CausesValidation = false;
			this.btnCancelar.Enabled = false;
			this.btnCancelar.Image = ((System.Drawing.Image)(resources.GetObject("btnCancelar.Image")));
			this.btnCancelar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnCancelar.Location = new System.Drawing.Point(599, 641);
			this.btnCancelar.Name = "btnCancelar";
			this.btnCancelar.Size = new System.Drawing.Size(72, 22);
			this.btnCancelar.TabIndex = 179;
			this.btnCancelar.Text = "&Cancelar";
			this.btnCancelar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnCancelar.Click += new System.EventHandler(this.btnCancelar_Click);
			// 
			// btnSalir
			// 
			this.btnSalir.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnSalir.CausesValidation = false;
			this.btnSalir.Image = ((System.Drawing.Image)(resources.GetObject("btnSalir.Image")));
			this.btnSalir.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnSalir.Location = new System.Drawing.Point(675, 641);
			this.btnSalir.Name = "btnSalir";
			this.btnSalir.Size = new System.Drawing.Size(75, 22);
			this.btnSalir.TabIndex = 180;
			this.btnSalir.Text = "&Salir";
			this.btnSalir.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnSalir.Click += new System.EventHandler(this.btnSalir_Click);
			// 
			// txtIdArticulo
			// 
			this.txtIdArticulo.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.txtIdArticulo.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtIdArticulo.Enabled = false;
			this.txtIdArticulo.Location = new System.Drawing.Point(1098, 641);
			this.txtIdArticulo.Name = "txtIdArticulo";
			this.txtIdArticulo.PromptChar = ' ';
			this.txtIdArticulo.Size = new System.Drawing.Size(16, 22);
			this.txtIdArticulo.TabIndex = 265;
			this.txtIdArticulo.Visible = false;
			// 
			// txtIdTipoGrupo
			// 
			this.txtIdTipoGrupo.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.txtIdTipoGrupo.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtIdTipoGrupo.Enabled = false;
			this.txtIdTipoGrupo.Location = new System.Drawing.Point(1074, 641);
			this.txtIdTipoGrupo.Name = "txtIdTipoGrupo";
			this.txtIdTipoGrupo.PromptChar = ' ';
			this.txtIdTipoGrupo.Size = new System.Drawing.Size(16, 22);
			this.txtIdTipoGrupo.TabIndex = 266;
			this.txtIdTipoGrupo.Visible = false;
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
			// txtCodArtEasyMax
			// 
			appearance61.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtCodArtEasyMax.Appearance = appearance61;
			this.txtCodArtEasyMax.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtCodArtEasyMax.Enabled = false;
			this.txtCodArtEasyMax.Location = new System.Drawing.Point(928, 640);
			this.txtCodArtEasyMax.Name = "txtCodArtEasyMax";
			this.txtCodArtEasyMax.Size = new System.Drawing.Size(96, 22);
			this.txtCodArtEasyMax.TabIndex = 267;
			this.txtCodArtEasyMax.Visible = false;
			// 
			// label57
			// 
			this.label57.AutoSize = true;
			this.label57.BackColor = System.Drawing.Color.Transparent;
			this.label57.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Bold);
			this.label57.Location = new System.Drawing.Point(824, 640);
			this.label57.Name = "label57";
			this.label57.Size = new System.Drawing.Size(102, 17);
			this.label57.TabIndex = 268;
			this.label57.Text = "Código EASYMAX";
			this.label57.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.label57.Visible = false;
			// 
			// label5
			// 
			this.label5.AutoSize = true;
			this.label5.BackColor = System.Drawing.Color.Transparent;
			this.label5.Location = new System.Drawing.Point(384, 106);
			this.label5.Name = "label5";
			this.label5.Size = new System.Drawing.Size(51, 16);
			this.label5.TabIndex = 270;
			this.label5.Text = "Preciador";
			this.label5.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// cmbPreciador
			// 
			appearance62.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbPreciador.Appearance = appearance62;
			this.cmbPreciador.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbPreciador.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbPreciador.DataSource = this.ultraDataSource9;
			this.cmbPreciador.DisplayLayout.AutoFitColumns = true;
			ultraGridColumn19.Header.VisiblePosition = 0;
			ultraGridColumn19.Hidden = true;
			ultraGridColumn19.Width = 101;
			ultraGridColumn20.Header.VisiblePosition = 1;
			ultraGridColumn20.Width = 237;
			ultraGridBand8.Columns.AddRange(new object[] {
																										 ultraGridColumn19,
																										 ultraGridColumn20});
			this.cmbPreciador.DisplayLayout.BandsSerializer.Add(ultraGridBand8);
			this.cmbPreciador.DisplayMember = "Preciador";
			this.cmbPreciador.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbPreciador.Enabled = false;
			this.cmbPreciador.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbPreciador.Location = new System.Drawing.Point(472, 104);
			this.cmbPreciador.Name = "cmbPreciador";
			this.cmbPreciador.Size = new System.Drawing.Size(256, 21);
			this.cmbPreciador.TabIndex = 269;
			this.cmbPreciador.ValueMember = "idPreciador";
			this.cmbPreciador.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cmbPreciador_KeyDown);
			// 
			// ultraDataSource9
			// 
			ultraDataColumn21.DataType = typeof(int);
			this.ultraDataSource9.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn21,
																																 ultraDataColumn22});
			// 
			// cmbMarca
			// 
			appearance63.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbMarca.Appearance = appearance63;
			this.cmbMarca.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbMarca.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbMarca.DataSource = this.ultraDataSource3;
			this.cmbMarca.DisplayLayout.AutoFitColumns = true;
			ultraGridColumn21.Header.VisiblePosition = 0;
			ultraGridColumn21.Hidden = true;
			ultraGridColumn21.Width = 88;
			ultraGridColumn22.Header.VisiblePosition = 1;
			ultraGridColumn22.Width = 237;
			ultraGridBand9.Columns.AddRange(new object[] {
																										 ultraGridColumn21,
																										 ultraGridColumn22});
			this.cmbMarca.DisplayLayout.BandsSerializer.Add(ultraGridBand9);
			this.cmbMarca.DisplayMember = "Marca";
			this.cmbMarca.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbMarca.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbMarca.Enabled = false;
			this.cmbMarca.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbMarca.Location = new System.Drawing.Point(472, 40);
			this.cmbMarca.Name = "cmbMarca";
			this.cmbMarca.Size = new System.Drawing.Size(256, 21);
			this.cmbMarca.TabIndex = 271;
			this.cmbMarca.ValueMember = "idMarca";
			this.cmbMarca.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cmbMarca_KeyDown);
			this.cmbMarca.ValueChanged += new System.EventHandler(this.cmbMarca_ValueChanged);
			this.cmbMarca.Enter += new System.EventHandler(this.cmbMarca_Enter);
			// 
			// chkGarantiaExtendida
			// 
			appearance64.ForeColorDisabled = System.Drawing.Color.Black;
			this.chkGarantiaExtendida.Appearance = appearance64;
			this.chkGarantiaExtendida.BackColor = System.Drawing.Color.Transparent;
			this.chkGarantiaExtendida.Enabled = false;
			this.chkGarantiaExtendida.Location = new System.Drawing.Point(456, 296);
			this.chkGarantiaExtendida.Name = "chkGarantiaExtendida";
			this.chkGarantiaExtendida.Size = new System.Drawing.Size(120, 22);
			this.chkGarantiaExtendida.TabIndex = 272;
			this.chkGarantiaExtendida.Text = "Garantia Extendida";
			this.chkGarantiaExtendida.CheckedChanged += new System.EventHandler(this.chkGarantiaExtendida_CheckedChanged);
			// 
			// chkServicioTecnicoDomicilio
			// 
			appearance65.ForeColorDisabled = System.Drawing.Color.Black;
			this.chkServicioTecnicoDomicilio.Appearance = appearance65;
			this.chkServicioTecnicoDomicilio.Enabled = false;
			this.chkServicioTecnicoDomicilio.Location = new System.Drawing.Point(576, 296);
			this.chkServicioTecnicoDomicilio.Name = "chkServicioTecnicoDomicilio";
			this.chkServicioTecnicoDomicilio.Size = new System.Drawing.Size(160, 22);
			this.chkServicioTecnicoDomicilio.TabIndex = 273;
			this.chkServicioTecnicoDomicilio.Text = "Servicio Tecnico a Domicilio";
			// 
			// label8
			// 
			this.label8.AutoSize = true;
			this.label8.BackColor = System.Drawing.Color.Transparent;
			this.label8.Location = new System.Drawing.Point(456, 328);
			this.label8.Name = "label8";
			this.label8.Size = new System.Drawing.Size(189, 16);
			this.label8.TabIndex = 275;
			this.label8.Text = "Tiempo de Garantía Extendida (Dias)";
			this.label8.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtTiempoGE
			// 
			appearance66.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtTiempoGE.Appearance = appearance66;
			this.txtTiempoGE.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtTiempoGE.Enabled = false;
			this.txtTiempoGE.FormatString = "##0";
			this.txtTiempoGE.Location = new System.Drawing.Point(648, 328);
			this.txtTiempoGE.MaskInput = "nnn";
			this.txtTiempoGE.MaxValue = 1000;
			this.txtTiempoGE.MinValue = 0;
			this.txtTiempoGE.Name = "txtTiempoGE";
			this.txtTiempoGE.PromptChar = ' ';
			this.txtTiempoGE.Size = new System.Drawing.Size(80, 22);
			this.txtTiempoGE.TabIndex = 274;
			// 
			// label10
			// 
			this.label10.AutoSize = true;
			this.label10.BackColor = System.Drawing.Color.Transparent;
			this.label10.Location = new System.Drawing.Point(456, 360);
			this.label10.Name = "label10";
			this.label10.Size = new System.Drawing.Size(171, 16);
			this.label10.TabIndex = 277;
			this.label10.Text = "Porcentaje de Garantia Extendida";
			this.label10.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtPorcentajeGE
			// 
			appearance67.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtPorcentajeGE.Appearance = appearance67;
			this.txtPorcentajeGE.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtPorcentajeGE.Enabled = false;
			this.txtPorcentajeGE.FormatString = "#,##0.00";
			this.txtPorcentajeGE.Location = new System.Drawing.Point(648, 360);
			this.txtPorcentajeGE.Name = "txtPorcentajeGE";
			this.txtPorcentajeGE.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtPorcentajeGE.PromptChar = ' ';
			this.txtPorcentajeGE.Size = new System.Drawing.Size(80, 22);
			this.txtPorcentajeGE.TabIndex = 276;
			// 
			// btnCaracteristicas
			// 
			this.btnCaracteristicas.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnCaracteristicas.CausesValidation = false;
			this.btnCaracteristicas.Enabled = false;
			this.btnCaracteristicas.Image = ((System.Drawing.Image)(resources.GetObject("btnCaracteristicas.Image")));
			this.btnCaracteristicas.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnCaracteristicas.Location = new System.Drawing.Point(407, 641);
			this.btnCaracteristicas.Name = "btnCaracteristicas";
			this.btnCaracteristicas.Size = new System.Drawing.Size(104, 22);
			this.btnCaracteristicas.TabIndex = 278;
			this.btnCaracteristicas.Text = "&Caracteristicas";
			this.btnCaracteristicas.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnCaracteristicas.Click += new System.EventHandler(this.btnCaracteristicas_Click);
			// 
			// label11
			// 
			this.label11.AutoSize = true;
			this.label11.BackColor = System.Drawing.Color.Transparent;
			this.label11.Location = new System.Drawing.Point(752, 139);
			this.label11.Name = "label11";
			this.label11.Size = new System.Drawing.Size(80, 16);
			this.label11.TabIndex = 280;
			this.label11.Text = "Link Fabricante";
			this.label11.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtLink
			// 
			appearance68.BorderColor = System.Drawing.Color.Black;
			appearance68.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtLink.Appearance = appearance68;
			this.txtLink.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtLink.Enabled = false;
			this.txtLink.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.txtLink.Location = new System.Drawing.Point(864, 136);
			this.txtLink.Name = "txtLink";
			this.txtLink.Size = new System.Drawing.Size(248, 22);
			this.txtLink.TabIndex = 279;
			this.txtLink.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtLink_KeyDown);
			// 
			// chkConsignacion
			// 
			appearance69.ForeColorDisabled = System.Drawing.Color.Black;
			this.chkConsignacion.Appearance = appearance69;
			this.chkConsignacion.BackColor = System.Drawing.Color.Transparent;
			this.chkConsignacion.Enabled = false;
			this.chkConsignacion.Location = new System.Drawing.Point(752, 8);
			this.chkConsignacion.Name = "chkConsignacion";
			this.chkConsignacion.Size = new System.Drawing.Size(120, 22);
			this.chkConsignacion.TabIndex = 282;
			this.chkConsignacion.Text = "En Consignación";
			this.chkConsignacion.CheckedChanged += new System.EventHandler(this.chkConsignacion_CheckedChanged);
			// 
			// txtCSCostoEfectivo
			// 
			appearance70.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtCSCostoEfectivo.Appearance = appearance70;
			this.txtCSCostoEfectivo.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtCSCostoEfectivo.Enabled = false;
			this.txtCSCostoEfectivo.FormatString = "#,##0.00";
			this.txtCSCostoEfectivo.Location = new System.Drawing.Point(848, 39);
			this.txtCSCostoEfectivo.Name = "txtCSCostoEfectivo";
			this.txtCSCostoEfectivo.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtCSCostoEfectivo.PromptChar = ' ';
			this.txtCSCostoEfectivo.Size = new System.Drawing.Size(80, 22);
			this.txtCSCostoEfectivo.TabIndex = 286;
			this.txtCSCostoEfectivo.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtCSCostoEfectivo_KeyDown);
			this.txtCSCostoEfectivo.Click += new System.EventHandler(this.txtCSCostoEfectivo_Click);
			this.txtCSCostoEfectivo.Enter += new System.EventHandler(this.txtCSCostoEfectivo_Enter);
			// 
			// label16
			// 
			this.label16.AutoSize = true;
			this.label16.BackColor = System.Drawing.Color.Transparent;
			this.label16.Location = new System.Drawing.Point(752, 42);
			this.label16.Name = "label16";
			this.label16.Size = new System.Drawing.Size(75, 16);
			this.label16.TabIndex = 285;
			this.label16.Text = "Costo Efectivo";
			this.label16.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label52
			// 
			this.label52.AutoSize = true;
			this.label52.BackColor = System.Drawing.Color.Transparent;
			this.label52.Location = new System.Drawing.Point(944, 42);
			this.label52.Name = "label52";
			this.label52.Size = new System.Drawing.Size(71, 16);
			this.label52.TabIndex = 284;
			this.label52.Text = "Costo Crédito";
			this.label52.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtCSCostoCredito
			// 
			appearance71.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtCSCostoCredito.Appearance = appearance71;
			this.txtCSCostoCredito.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtCSCostoCredito.Enabled = false;
			this.txtCSCostoCredito.FormatString = "#,##0.00";
			this.txtCSCostoCredito.Location = new System.Drawing.Point(1032, 39);
			this.txtCSCostoCredito.Name = "txtCSCostoCredito";
			this.txtCSCostoCredito.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtCSCostoCredito.PromptChar = ' ';
			this.txtCSCostoCredito.Size = new System.Drawing.Size(80, 22);
			this.txtCSCostoCredito.TabIndex = 283;
			this.txtCSCostoCredito.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtCSCostoCredito_KeyDown);
			this.txtCSCostoCredito.Click += new System.EventHandler(this.txtCSCostoCredito_Click);
			this.txtCSCostoCredito.Enter += new System.EventHandler(this.txtCSCostoCredito_Enter);
			// 
			// label18
			// 
			this.label18.AutoSize = true;
			this.label18.BackColor = System.Drawing.Color.Transparent;
			this.label18.Location = new System.Drawing.Point(752, 75);
			this.label18.Name = "label18";
			this.label18.Size = new System.Drawing.Size(92, 16);
			this.label18.TabIndex = 288;
			this.label18.Text = "Codigo Proveedor";
			this.label18.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtCodigoProveedor
			// 
			appearance72.BorderColor = System.Drawing.Color.Black;
			appearance72.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtCodigoProveedor.Appearance = appearance72;
			this.txtCodigoProveedor.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtCodigoProveedor.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtCodigoProveedor.Enabled = false;
			this.txtCodigoProveedor.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.txtCodigoProveedor.Location = new System.Drawing.Point(864, 72);
			this.txtCodigoProveedor.Name = "txtCodigoProveedor";
			this.txtCodigoProveedor.Size = new System.Drawing.Size(248, 22);
			this.txtCodigoProveedor.TabIndex = 287;
			// 
			// chkPayJoy
			// 
			appearance73.ForeColorDisabled = System.Drawing.Color.Black;
			this.chkPayJoy.Appearance = appearance73;
			this.chkPayJoy.BackColor = System.Drawing.Color.Transparent;
			this.chkPayJoy.Enabled = false;
			this.chkPayJoy.Location = new System.Drawing.Point(752, 168);
			this.chkPayJoy.Name = "chkPayJoy";
			this.chkPayJoy.Size = new System.Drawing.Size(104, 22);
			this.chkPayJoy.TabIndex = 290;
			this.chkPayJoy.Text = "Codigo Pay Joy";
			this.chkPayJoy.CheckedChanged += new System.EventHandler(this.chkPayJoy_CheckedChanged);
			// 
			// label22
			// 
			this.label22.AutoSize = true;
			this.label22.BackColor = System.Drawing.Color.Transparent;
			this.label22.Enabled = false;
			this.label22.Location = new System.Drawing.Point(888, 208);
			this.label22.Name = "label22";
			this.label22.Size = new System.Drawing.Size(83, 16);
			this.label22.TabIndex = 292;
			this.label22.Text = "Codigo Pay Joy ";
			this.label22.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.label22.Visible = false;
			this.label22.Click += new System.EventHandler(this.label22_Click);
			// 
			// txtCodigoPayJoy
			// 
			appearance74.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtCodigoPayJoy.Appearance = appearance74;
			this.txtCodigoPayJoy.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtCodigoPayJoy.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtCodigoPayJoy.Enabled = false;
			this.txtCodigoPayJoy.Location = new System.Drawing.Point(864, 168);
			this.txtCodigoPayJoy.Name = "txtCodigoPayJoy";
			this.txtCodigoPayJoy.Size = new System.Drawing.Size(248, 22);
			this.txtCodigoPayJoy.TabIndex = 293;
			this.txtCodigoPayJoy.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtCodigoPayJoy_KeyPress);
			// 
			// label24
			// 
			this.label24.AutoSize = true;
			this.label24.BackColor = System.Drawing.Color.Transparent;
			this.label24.Location = new System.Drawing.Point(752, 296);
			this.label24.Name = "label24";
			this.label24.Size = new System.Drawing.Size(37, 16);
			this.label24.TabIndex = 294;
			this.label24.Text = "Origen";
			this.label24.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label25
			// 
			this.label25.AutoSize = true;
			this.label25.BackColor = System.Drawing.Color.Transparent;
			this.label25.Location = new System.Drawing.Point(928, 296);
			this.label25.Name = "label25";
			this.label25.Size = new System.Drawing.Size(23, 16);
			this.label25.TabIndex = 295;
			this.label25.Text = "Año";
			this.label25.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.label25.Click += new System.EventHandler(this.label25_Click);
			// 
			// label26
			// 
			this.label26.AutoSize = true;
			this.label26.BackColor = System.Drawing.Color.Transparent;
			this.label26.Location = new System.Drawing.Point(752, 328);
			this.label26.Name = "label26";
			this.label26.Size = new System.Drawing.Size(47, 16);
			this.label26.TabIndex = 296;
			this.label26.Text = "Tonelaje";
			this.label26.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label27
			// 
			this.label27.AutoSize = true;
			this.label27.BackColor = System.Drawing.Color.Transparent;
			this.label27.Location = new System.Drawing.Point(928, 328);
			this.label27.Name = "label27";
			this.label27.Size = new System.Drawing.Size(50, 16);
			this.label27.TabIndex = 297;
			this.label27.Text = "Cilindraje";
			this.label27.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label28
			// 
			this.label28.AutoSize = true;
			this.label28.BackColor = System.Drawing.Color.Transparent;
			this.label28.Location = new System.Drawing.Point(752, 360);
			this.label28.Name = "label28";
			this.label28.Size = new System.Drawing.Size(52, 16);
			this.label28.TabIndex = 298;
			this.label28.Text = "Pasajeros";
			this.label28.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtfechaMoto
			// 
			appearance75.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtfechaMoto.Appearance = appearance75;
			this.txtfechaMoto.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtfechaMoto.Enabled = false;
			this.txtfechaMoto.FormatString = "##0";
			this.txtfechaMoto.Location = new System.Drawing.Point(992, 296);
			this.txtfechaMoto.MaskInput = "nnnn";
			this.txtfechaMoto.MaxValue = 3650;
			this.txtfechaMoto.MinValue = 0;
			this.txtfechaMoto.Name = "txtfechaMoto";
			this.txtfechaMoto.PromptChar = ' ';
			this.txtfechaMoto.Size = new System.Drawing.Size(72, 22);
			this.txtfechaMoto.TabIndex = 299;
			this.txtfechaMoto.ValueChanged += new System.EventHandler(this.txtafechaMoto_ValueChanged);
			// 
			// txtOrigenMoto
			// 
			appearance76.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtOrigenMoto.Appearance = appearance76;
			this.txtOrigenMoto.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtOrigenMoto.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtOrigenMoto.Enabled = false;
			this.txtOrigenMoto.Location = new System.Drawing.Point(816, 296);
			this.txtOrigenMoto.Name = "txtOrigenMoto";
			this.txtOrigenMoto.Size = new System.Drawing.Size(72, 22);
			this.txtOrigenMoto.TabIndex = 300;
			// 
			// txtTonelaje
			// 
			appearance77.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtTonelaje.Appearance = appearance77;
			this.txtTonelaje.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtTonelaje.Enabled = false;
			this.txtTonelaje.FormatString = "#,##0.00";
			this.txtTonelaje.Location = new System.Drawing.Point(816, 328);
			this.txtTonelaje.Name = "txtTonelaje";
			this.txtTonelaje.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtTonelaje.PromptChar = ' ';
			this.txtTonelaje.Size = new System.Drawing.Size(72, 22);
			this.txtTonelaje.TabIndex = 301;
			// 
			// txtCilindraje
			// 
			appearance78.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtCilindraje.Appearance = appearance78;
			this.txtCilindraje.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtCilindraje.Enabled = false;
			this.txtCilindraje.FormatString = "##0";
			this.txtCilindraje.Location = new System.Drawing.Point(992, 328);
			this.txtCilindraje.MaskInput = "nnnn";
			this.txtCilindraje.MaxValue = 3650;
			this.txtCilindraje.MinValue = 0;
			this.txtCilindraje.Name = "txtCilindraje";
			this.txtCilindraje.PromptChar = ' ';
			this.txtCilindraje.Size = new System.Drawing.Size(72, 22);
			this.txtCilindraje.TabIndex = 302;
			// 
			// txtPasajeros
			// 
			appearance79.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtPasajeros.Appearance = appearance79;
			this.txtPasajeros.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtPasajeros.Enabled = false;
			this.txtPasajeros.FormatString = "##0";
			this.txtPasajeros.Location = new System.Drawing.Point(816, 360);
			this.txtPasajeros.MaskInput = "nnnn";
			this.txtPasajeros.MaxValue = 3650;
			this.txtPasajeros.MinValue = 0;
			this.txtPasajeros.Name = "txtPasajeros";
			this.txtPasajeros.PromptChar = ' ';
			this.txtPasajeros.Size = new System.Drawing.Size(72, 22);
			this.txtPasajeros.TabIndex = 303;
			// 
			// chkMoto
			// 
			appearance80.ForeColorDisabled = System.Drawing.Color.Black;
			this.chkMoto.Appearance = appearance80;
			this.chkMoto.BackColor = System.Drawing.Color.Transparent;
			this.chkMoto.Enabled = false;
			this.chkMoto.Location = new System.Drawing.Point(752, 264);
			this.chkMoto.Name = "chkMoto";
			this.chkMoto.Size = new System.Drawing.Size(120, 22);
			this.chkMoto.TabIndex = 304;
			this.chkMoto.Text = "Motocicletas";
			this.chkMoto.CheckedChanged += new System.EventHandler(this.chkMoto_CheckedChanged);
			// 
			// btnAtributos
			// 
			this.btnAtributos.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnAtributos.CausesValidation = false;
			this.btnAtributos.Enabled = false;
			this.btnAtributos.Image = ((System.Drawing.Image)(resources.GetObject("btnAtributos.Image")));
			this.btnAtributos.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnAtributos.Location = new System.Drawing.Point(515, 641);
			this.btnAtributos.Name = "btnAtributos";
			this.btnAtributos.Size = new System.Drawing.Size(80, 22);
			this.btnAtributos.TabIndex = 305;
			this.btnAtributos.Text = "&Atributos";
			this.btnAtributos.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnAtributos.Visible = false;
			this.btnAtributos.Click += new System.EventHandler(this.button1_Click);
			// 
			// chkMigrar
			// 
			appearance81.ForeColorDisabled = System.Drawing.Color.Black;
			this.chkMigrar.Appearance = appearance81;
			this.chkMigrar.BackColor = System.Drawing.Color.Transparent;
			this.chkMigrar.Enabled = false;
			this.chkMigrar.Location = new System.Drawing.Point(856, 208);
			this.chkMigrar.Name = "chkMigrar";
			this.chkMigrar.Size = new System.Drawing.Size(128, 22);
			this.chkMigrar.TabIndex = 306;
			this.chkMigrar.Text = "Migrar";
			this.chkMigrar.Visible = false;
			this.chkMigrar.CheckedChanged += new System.EventHandler(this.chkMigrar_CheckedChanged);
			// 
			// label30
			// 
			this.label30.AutoSize = true;
			this.label30.BackColor = System.Drawing.Color.Transparent;
			this.label30.Location = new System.Drawing.Point(856, 208);
			this.label30.Name = "label30";
			this.label30.Size = new System.Drawing.Size(74, 16);
			this.label30.TabIndex = 307;
			this.label30.Text = "Cod Novasoft ";
			this.label30.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.label30.Visible = false;
			// 
			// txtnovasoft
			// 
			appearance82.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtnovasoft.Appearance = appearance82;
			this.txtnovasoft.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtnovasoft.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtnovasoft.Enabled = false;
			this.txtnovasoft.Location = new System.Drawing.Point(856, 208);
			this.txtnovasoft.Name = "txtnovasoft";
			this.txtnovasoft.Size = new System.Drawing.Size(56, 22);
			this.txtnovasoft.TabIndex = 308;
			this.txtnovasoft.Visible = false;
			// 
			// chkManejaSerie
			// 
			appearance83.ForeColorDisabled = System.Drawing.Color.Black;
			this.chkManejaSerie.Appearance = appearance83;
			this.chkManejaSerie.BackColor = System.Drawing.Color.Transparent;
			this.chkManejaSerie.Enabled = false;
			this.chkManejaSerie.Location = new System.Drawing.Point(880, 208);
			this.chkManejaSerie.Name = "chkManejaSerie";
			this.chkManejaSerie.Size = new System.Drawing.Size(128, 22);
			this.chkManejaSerie.TabIndex = 309;
			this.chkManejaSerie.Text = "Maneja Serie";
			this.chkManejaSerie.Visible = false;
			// 
			// label41
			// 
			this.label41.AutoSize = true;
			this.label41.BackColor = System.Drawing.Color.Transparent;
			this.label41.Location = new System.Drawing.Point(24, 200);
			this.label41.Name = "label41";
			this.label41.Size = new System.Drawing.Size(87, 16);
			this.label41.TabIndex = 312;
			this.label41.Text = "Descripción Web";
			this.label41.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtDescripcionWeb
			// 
			appearance84.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtDescripcionWeb.Appearance = appearance84;
			this.txtDescripcionWeb.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtDescripcionWeb.Enabled = false;
			this.txtDescripcionWeb.Location = new System.Drawing.Point(112, 200);
			this.txtDescripcionWeb.Name = "txtDescripcionWeb";
			this.txtDescripcionWeb.Size = new System.Drawing.Size(616, 22);
			this.txtDescripcionWeb.TabIndex = 311;
			// 
			// txtTituloWeb
			// 
			appearance85.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtTituloWeb.Appearance = appearance85;
			this.txtTituloWeb.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtTituloWeb.Enabled = false;
			this.txtTituloWeb.Location = new System.Drawing.Point(112, 232);
			this.txtTituloWeb.Name = "txtTituloWeb";
			this.txtTituloWeb.Size = new System.Drawing.Size(616, 22);
			this.txtTituloWeb.TabIndex = 313;
			// 
			// label42
			// 
			this.label42.AutoSize = true;
			this.label42.BackColor = System.Drawing.Color.Transparent;
			this.label42.Location = new System.Drawing.Point(24, 232);
			this.label42.Name = "label42";
			this.label42.Size = new System.Drawing.Size(58, 16);
			this.label42.TabIndex = 314;
			this.label42.Text = "Titulo Web";
			this.label42.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// frmArticulos
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 14);
			this.ClientSize = new System.Drawing.Size(1121, 666);
			this.Controls.Add(this.label42);
			this.Controls.Add(this.label41);
			this.Controls.Add(this.label30);
			this.Controls.Add(this.label28);
			this.Controls.Add(this.label27);
			this.Controls.Add(this.label26);
			this.Controls.Add(this.label25);
			this.Controls.Add(this.label24);
			this.Controls.Add(this.label22);
			this.Controls.Add(this.label18);
			this.Controls.Add(this.label16);
			this.Controls.Add(this.label52);
			this.Controls.Add(this.label11);
			this.Controls.Add(this.label10);
			this.Controls.Add(this.label8);
			this.Controls.Add(this.label5);
			this.Controls.Add(this.label57);
			this.Controls.Add(this.label4);
			this.Controls.Add(this.label12);
			this.Controls.Add(this.label9);
			this.Controls.Add(this.label58);
			this.Controls.Add(this.label56);
			this.Controls.Add(this.label54);
			this.Controls.Add(this.label3);
			this.Controls.Add(this.label2);
			this.Controls.Add(this.label1);
			this.Controls.Add(this.label35);
			this.Controls.Add(this.label23);
			this.Controls.Add(this.label19);
			this.Controls.Add(this.label47);
			this.Controls.Add(this.label48);
			this.Controls.Add(this.label46);
			this.Controls.Add(this.label13);
			this.Controls.Add(this.label6);
			this.Controls.Add(this.label7);
			this.Controls.Add(this.label32);
			this.Controls.Add(this.txtTituloWeb);
			this.Controls.Add(this.txtDescripcionWeb);
			this.Controls.Add(this.chkManejaSerie);
			this.Controls.Add(this.txtnovasoft);
			this.Controls.Add(this.chkMigrar);
			this.Controls.Add(this.btnAtributos);
			this.Controls.Add(this.chkMoto);
			this.Controls.Add(this.txtPasajeros);
			this.Controls.Add(this.txtCilindraje);
			this.Controls.Add(this.txtTonelaje);
			this.Controls.Add(this.txtOrigenMoto);
			this.Controls.Add(this.txtfechaMoto);
			this.Controls.Add(this.txtCodigoPayJoy);
			this.Controls.Add(this.chkPayJoy);
			this.Controls.Add(this.txtCodigoProveedor);
			this.Controls.Add(this.txtCSCostoEfectivo);
			this.Controls.Add(this.txtCSCostoCredito);
			this.Controls.Add(this.chkConsignacion);
			this.Controls.Add(this.txtLink);
			this.Controls.Add(this.btnCaracteristicas);
			this.Controls.Add(this.txtPorcentajeGE);
			this.Controls.Add(this.txtTiempoGE);
			this.Controls.Add(this.chkGarantiaExtendida);
			this.Controls.Add(this.chkServicioTecnicoDomicilio);
			this.Controls.Add(this.cmbMarca);
			this.Controls.Add(this.cmbPreciador);
			this.Controls.Add(this.txtCodArtEasyMax);
			this.Controls.Add(this.txtIdTipoGrupo);
			this.Controls.Add(this.txtIdArticulo);
			this.Controls.Add(this.btnNuevo);
			this.Controls.Add(this.btnConsultar);
			this.Controls.Add(this.btnEditar);
			this.Controls.Add(this.btnImprimir);
			this.Controls.Add(this.btnGuardar);
			this.Controls.Add(this.btnCancelar);
			this.Controls.Add(this.btnSalir);
			this.Controls.Add(this.groupBox5);
			this.Controls.Add(this.groupBox4);
			this.Controls.Add(this.groupBox3);
			this.Controls.Add(this.groupBox2);
			this.Controls.Add(this.txtIva);
			this.Controls.Add(this.cmbSubGrupo);
			this.Controls.Add(this.txtVidaUtil);
			this.Controls.Add(this.cmbGrupo);
			this.Controls.Add(this.txtTamaño);
			this.Controls.Add(this.txtPeriodoGarantia);
			this.Controls.Add(this.btGrupo);
			this.Controls.Add(this.cmbCategoria);
			this.Controls.Add(this.txtModelo);
			this.Controls.Add(this.txtNParte);
			this.Controls.Add(this.cmbColor);
			this.Controls.Add(this.btColor);
			this.Controls.Add(this.cmbArticuloSRI);
			this.Controls.Add(this.txtCodigo);
			this.Controls.Add(this.txtNombre);
			this.Controls.Add(this.txtCodigoReemplazo);
			this.Controls.Add(this.groupBox1);
			this.Controls.Add(this.txtCantMaxima);
			this.Controls.Add(this.txtCantMinima);
			this.Controls.Add(this.txtSaldo);
			this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
			this.KeyPreview = true;
			this.MaximizeBox = false;
			this.Name = "frmArticulos";
			this.Text = "Articulos";
			this.Closing += new System.ComponentModel.CancelEventHandler(this.frmArticulos_Closing);
			this.Load += new System.EventHandler(this.frmArticulos_Load);
			((System.ComponentModel.ISupportInitialize)(this.txtVidaUtil)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtTamaño)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtPeriodoGarantia)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCategoria)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).EndInit();
			this.groupBox1.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.txtLargo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtPeso)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtAlto)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtAncho)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.rbgOrigen.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbArancel)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource7)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtModelo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNParte)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbColor)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource6)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource3)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtSaldo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbArticuloSRI)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource4)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCodigo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNombre)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIva)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCantMaxima)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCantMinima)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbSubGrupo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource5)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbGrupo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtPrecioCostoUltimo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtPrecioPVP)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtPrecioTarjeta)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtPrecioPA)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtPrecioPI)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtPrecioIncr2)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtPrecioIncr1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCodigoReemplazo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFechaUltimo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDepreciacion)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtTiempo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbSeccion)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource8)).EndInit();
			this.groupBox2.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.txtPrecioDescEmpleado)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtPrecioAutoconsumo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCostoPayJoy)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtPrecioTCsinIntereses)).EndInit();
			this.groupBox3.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.txtResidual)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCentroCosto)).EndInit();
			this.groupBox5.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.txtIdArticulo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdTipoGrupo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCodArtEasyMax)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbPreciador)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource9)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbMarca)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtTiempoGE)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtPorcentajeGE)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtLink)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCSCostoEfectivo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCSCostoCredito)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCodigoProveedor)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCodigoPayJoy)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtfechaMoto)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtOrigenMoto)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtTonelaje)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCilindraje)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtPasajeros)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtnovasoft)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDescripcionWeb)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtTituloWeb)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void UnloadMe()
		{
			this.Close();
		}

		int idSubgrupo = 0;
		int iDescripcion = 0;

		private void frmArticulos_Load(object sender, System.EventArgs e)
		{
			miAcceso = new Acceso(cdsSeteoF, "0606");

			if (!Funcion.Permiso("106",cdsSeteoF))
			{
				MessageBox.Show("No puede ingresar a Articulos", "Informacion");
				BeginInvoke(new MethodInvoker(UnloadMe));
			}

			Funcion.EjecutaSQL(cdsSeteoF, "Exec AuditaCrear 53, 6, 'Articulo'", true);
			
			Funcion.Resolucion(this, this.Height, this.Width);
			
			this.cmbGrupo.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Select idGrupoArticulo, Grupo From ArticuloGrupo Where idTipoGrupo In (1, 5, 6, 13, 16) Order By Grupo");
			this.cmbArticuloSRI.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Select idArticuloSRI, Codigo, Detalle, Porcentaje From ArticuloSRI Order By Codigo");
			this.cmbColor.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Select idColor, Color From ArticuloColor Order By Color");
			this.cmbPreciador.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Select idPreciador, Preciador From Preciador Order By idPreciador");
			this.cmbArancel.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Select idArancel, Codigo, Porcentaje, Arancel From ArticuloArancel Order By Arancel");
			this.cmbSeccion.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Select idSeccion, Seccion From ArticuloSeccion Order By Seccion");

			if (miAcceso.Imprimir) this.btnImprimir.Enabled = true;
			if (miAcceso.BAccesoBoton) this.btnAtributos.Visible = true;
			if (MenuLatinium.stIdDB == 29 || MenuLatinium.stIdDB == 34)
			{
				this.label34.Visible = true;
				this.label38.Visible = true;
				this.txtPrecioDescEmpleado.Visible = true;
				this.txtPrecioAutoconsumo.Visible = true;
				this.chkDepartamento.Visible = true;
				this.chkDepartamento1.Visible = true;
			}
			txtCodigoPayJoy.ContextMenu = new ContextMenu();
		}

		private void btnSalir_Click(object sender, System.EventArgs e)
		{
			this.Close();
		}

		private void btnCancelar_Click(object sender, System.EventArgs e)
		{
			bNuevo = false;
			bEdicion = false;

			idSubgrupo = 0;
			this.txtIdArticulo.Value = 0;
			this.txtIdTipoGrupo.Value = 0;

			this.cmbGrupo.Value = System.DBNull.Value;
			this.cmbSubGrupo.Value = System.DBNull.Value;
			this.cmbCategoria.Value = System.DBNull.Value;
			this.cmbMarca.Value = System.DBNull.Value;
			this.txtModelo.Text = "";
			this.txtNParte.Text = "";
			this.cmbColor.Value = System.DBNull.Value;
			this.cmbPreciador.Value = System.DBNull.Value;
			this.txtCodigo.Text = "";
			this.txtCodigoReemplazo.Text = "";
			this.txtNombre.Text = "";
			this.cmbArticuloSRI.Value = System.DBNull.Value;
			this.txtIva.Value = 0;
			this.txtTamaño.Value = 0;
			this.txtPeriodoGarantia.Value = 0;
			this.txtVidaUtil.Value = 0;
			this.txtSaldo.Value = 0.00m;
			this.txtCantMinima.Value = 0.00m;
			this.txtCantMaxima.Value = 0.00m;
			this.rbgOrigen.EditValue = 0;
			this.cmbArancel.Value = System.DBNull.Value;
			this.txtPeso.Value = 0.00m;
			this.txtAncho.Value = 0.00m;
			this.txtLargo.Value = 0.00m;
			this.txtAlto.Value = 0.00m;
			this.chkPrecio.Checked = false;
			this.txtPrecioPI.Value = 0.00m;
			this.txtPrecioIncr1.Value = 0.00m;
			this.txtPrecioPA.Value = 0.00m;
			this.txtPrecioIncr2.Value = 0.00m;
			this.txtPrecioTarjeta.Value = 0.00m;
			this.txtPrecioPVP.Value = 0.00m;
			this.txtPrecioCostoUltimo.Value = 0.00m;	
			this.txtCostoPayJoy.Value = 0.00m;
			this.txtPrecioTCsinIntereses.Value = 0.00m;		
			this.txtPrecioDescEmpleado.Value = 0.00m;
			this.txtPrecioAutoconsumo.Value = 0.00m;
			this.cmbSeccion.Value = System.DBNull.Value;
			this.cmbCentroCosto.Value = System.DBNull.Value;
			this.txtDepreciacion.Value = 0.00m;
			this.txtResidual.Value = 0.00m;
			this.txtTiempo.Value = 0.00m;
			this.dtFechaUltimo.Value = System.DBNull.Value;
			this.chkConsignacion.Checked = false;
			this.txtCSCostoEfectivo.Value = 0.00m;
			this.txtCSCostoCredito.Value = 0.00m;
			this.txtCodigoProveedor.Text = "";	
			this.txtCodigoPayJoy.Text = "";		
			this.chkMoto.Checked = false;
			this.txtOrigenMoto.Text = "";
			this.txtfechaMoto.Value = 0;
			this.txtTonelaje.Value = 0.00m;
			this.txtCilindraje.Value = 0;
			this.txtPasajeros.Value = 0;
			this.txtDescripcionWeb.Text = "";
			this.txtTituloWeb.Text = "";

			this.txtOrigenMoto.Enabled = false;
			this.txtfechaMoto.Enabled = false;
			this.txtTonelaje.Enabled = false;
			this.txtCilindraje.Enabled = false;
			this.txtPasajeros.Enabled = false;

			this.chkMoto.Enabled = false;
			this.chkConsignacion.Enabled = false;
			this.txtCSCostoEfectivo.Enabled = false;
			this.txtCSCostoCredito.Enabled = false;
			this.txtCodigoProveedor.Enabled = false;
			this.chkCombo.Checked = false;
			this.chkObsequios.Checked = false;
			this.chkManejaSerial.Checked = false;
			this.chkSerialAlterna.Checked = false;
			this.chkImpuestoIVA.Checked = false;
			this.chkFaturable.Checked = false;
			this.chkAplicDescuento.Checked = false;
			this.chkDescontinuado.Checked = false;
			this.chkDescCompras.Checked = false;
			this.chkDescVentas.Checked = false;
			this.chkTransferible.Checked = false;
			this.chkInventario.Checked = false;
			this.chkEfectivo.Checked = false;
			this.chkProductoWeb.Checked = false;
			this.chkPayJoy.Checked = false;
			this.chkPrecioMovil.Checked = false;
			this.chkDepartamento.Enabled = false;
			this.chkDepartamento1.Enabled = false;

			this.chkMigrar.Checked = false;
			this.chkManejaSerie.Checked = false;
			this.txtnovasoft.Text = "";	
			
			this.txtCodArtEasyMax.Enabled = false;
			this.txtCodArtEasyMax.Text = "";

			this.chkGarantiaExtendida.Checked = false;
			this.txtPorcentajeGE.Value = 0.00m;
			this.txtTiempoGE.Value = 0;
			this.chkServicioTecnicoDomicilio.Checked = false;
			this.chkcatalogo.Checked = false;
			this.txtLink.Text = "";

			this.cmbGrupo.Enabled = false;
			this.cmbSubGrupo.Enabled = false;
			this.cmbCategoria.Enabled = false;
			this.cmbMarca.Enabled = false;
			this.txtModelo.Enabled = false;
			this.txtNParte.Enabled = false;
			this.cmbColor.Enabled = false;
			this.cmbPreciador.Enabled = false;
			this.txtCodigo.Enabled = false;
			this.txtCodigoReemplazo.Enabled = false;
			this.txtNombre.Enabled = false;
			this.cmbArticuloSRI.Enabled = false;
			this.txtIva.Enabled = false;
			this.txtTamaño.Enabled = false;
			this.txtPeriodoGarantia.Enabled = false;
			this.txtVidaUtil.Enabled = false;
			this.txtSaldo.Enabled = false;
			this.txtCantMinima.Enabled = false;
			this.txtCantMaxima.Enabled = false;
			this.rbgOrigen.Enabled = false;
			this.cmbArancel.Enabled = false;
			this.txtPeso.Enabled = false;
			this.txtAncho.Enabled = false;
			this.txtLargo.Enabled = false;
			this.txtAlto.Enabled = false;
			this.chkPrecio.Enabled = false;
			this.txtPrecioPI.Enabled = false;
			this.txtPrecioIncr1.Enabled = false;
			this.txtPrecioPA.Enabled = false;
			this.txtPrecioIncr2.Enabled = false;
			this.txtPrecioTarjeta.Enabled = false;
			this.txtPrecioPVP.Enabled = false;
			this.txtPrecioCostoUltimo.Enabled = false;	
			this.txtCostoPayJoy.Enabled = false;
			this.txtPrecioTCsinIntereses.Enabled = false;	
			this.txtPrecioDescEmpleado.Enabled = false;
			this.txtPrecioAutoconsumo.Enabled = false;
			this.cmbSeccion.Enabled = false;
			this.cmbCentroCosto.Enabled = false;
			this.txtDepreciacion.Enabled = false;
			this.txtResidual.Enabled = false;
			this.txtTiempo.Enabled = false;
			this.dtFechaUltimo.Enabled = false;
			this.chkCombo.Enabled = false;
			this.chkObsequios.Enabled = false;
			this.chkManejaSerial.Enabled = false;
			this.chkSerialAlterna.Enabled = false;
			this.chkFaturable.Enabled = false;
			this.chkAplicDescuento.Enabled = false;
			this.chkDescontinuado.Enabled = false;
			this.chkDescCompras.Enabled = false;
			this.chkDescVentas.Enabled = false;
			this.chkTransferible.Enabled = false;
			this.chkInventario.Enabled = false;
			this.chkEfectivo.Enabled = false;
			this.chkDescontinuado.Enabled = false;
			this.chkImpuestoIVA.Enabled = false;
			this.chkProductoWeb.Enabled = false;
			this.chkPayJoy.Enabled = false;
			this.chkDepartamento.Enabled = false;
			this.chkDepartamento1.Enabled = false;
			this.chkPrecioMovil.Enabled = false;
			this.chkMigrar.Enabled = false;
			this.chkManejaSerie.Enabled = false;
			this.txtnovasoft.Enabled= false;	
			this.txtDescripcionWeb.Enabled = false;
			this.txtTituloWeb.Enabled = false;
			
			this.chkGarantiaExtendida.Enabled = false;
			this.txtPorcentajeGE.Enabled = false;
			this.txtTiempoGE.Enabled = false;
			this.chkServicioTecnicoDomicilio.Enabled = false;
			this.chkcatalogo.Enabled = false;
			this.txtLink.Enabled = false;

			this.btnNuevo.Enabled = true;
			this.btnConsultar.Enabled = true;
			this.btnEditar.Enabled = false;
			this.btnImprimir.Enabled = false;
			this.btnGuardar.Enabled = false;
			this.btnCaracteristicas.Enabled = false;
			this.btnCancelar.Enabled = false;
		}

		private void btnNuevo_Click(object sender, System.EventArgs e)
		{
			this.btnCancelar_Click(sender, e);

			this.cmbGrupo.Enabled = true;
			this.cmbSubGrupo.Enabled = true;
			this.cmbCategoria.Enabled = true;
			this.cmbMarca.Enabled = true;
			this.txtModelo.Enabled = true;
			this.txtNParte.Enabled = true;
			this.cmbColor.Enabled = true;	
			this.txtCodigoReemplazo.Enabled = true;
			this.txtNombre.Enabled = true;
			this.txtTamaño.Enabled = true;
			this.txtPeriodoGarantia.Enabled = true;
			this.txtVidaUtil.Enabled = true;
			this.txtCantMinima.Enabled = true;
			this.txtCantMaxima.Enabled = true;
			this.rbgOrigen.Enabled = true;
			this.chkPrecio.Enabled = true;
			this.chkCombo.Enabled = true;
			this.chkObsequios.Enabled = true;
			this.chkManejaSerial.Enabled = true;
			this.chkSerialAlterna.Enabled = true;
			this.chkFaturable.Enabled = true;
			this.chkAplicDescuento.Enabled = true;
			this.chkDescontinuado.Enabled = true;
			this.chkDescCompras.Enabled = true;
			this.chkDescVentas.Enabled = true;
			this.chkTransferible.Enabled = true;
			this.chkImpuestoIVA.Enabled = true;
			this.txtCodArtEasyMax.Enabled = true;
			this.btnNuevo.Enabled = false;
			this.btnConsultar.Enabled = false;
			this.btnGuardar.Enabled = true;
			this.btnCancelar.Enabled = true;
			this.btnImprimir.Enabled = false;
			this.chkGarantiaExtendida.Enabled = true;
			this.chkcatalogo.Enabled = true;
			this.chkProductoWeb.Enabled = true;
			this.chkMigrar.Enabled = true;
			this.chkManejaSerie.Enabled = true;
			this.chkPayJoy.Enabled = true;
			this.chkDepartamento.Enabled = true;
			this.chkDepartamento1.Enabled = true;
			this.chkPrecioMovil.Enabled = true;
			//this.chkMigrar.Checked = true;
			//this.chkManejaSerie.Checked= true;
			//this.txtnovasoft.Enabled = true;
			this.txtPeso.Enabled = true;
			this.txtDescripcionWeb.Enabled = true;
			this.txtTituloWeb.Enabled = true;

			this.txtIva.Value = Funcion.decEscalarSQL(cdsSeteoF, "Select Iva From SeteoF");

			this.chkInventario.Enabled = true;
			this.rbgOrigen.EditValue = 1;
			this.chkEfectivo.Enabled = true;
			this.chkTransferible.Checked = true;
			this.chkFaturable.Checked = true;
			this.chkManejaSerial.Checked = true;
			this.chkInventario.Checked = true;
			this.chkEfectivo.Checked = false;
			this.chkImpuestoIVA.Checked = true;
			this.txtLink.Enabled = true;
			this.chkConsignacion.Enabled = true;			
			this.chkMoto.Enabled = true;			
			this.cmbGrupo.Focus();

			bNuevo = true;
			bEdicion = true;
		}

		private void cmbGrupo_ValueChanged(object sender, System.EventArgs e)
		{
			if (bloquearEventos) return; // Bloquear eventos mientras cargas combos

			if (this.cmbGrupo.ActiveRow != null) 
			{
				string sSQL = string.Format("Exec ListaSubgrupoArticulos {0}", (int)this.cmbGrupo.Value);
				this.cmbSubGrupo.DataSource = Funcion.dvProcedimiento(cdsSeteoF, sSQL);
				
				//this.txtIdTipoGrupo.Value = Funcion.iEscalarSQL(cdsSeteoF, string.Format("Select ISNULL(idTipoGrupo, 0) From ArticuloGrupo Where idGrupoArticulo = {0}", (int)this.cmbGrupo.Value));

				if (bEdicion)
				{
					if (!bNuevo) this.cmbSubGrupo.Value = idSubgrupo;
					this.cmbSubGrupo.Enabled = true;
					this.txtIdTipoGrupo.Value = Funcion.iEscalarSQL(cdsSeteoF, string.Format("Select ISNULL(idTipoGrupo, 0) From ArticuloGrupo Where idGrupoArticulo = {0}", (int)this.cmbGrupo.Value));

					if ((int)this.txtIdTipoGrupo.Value == 1 || (int)this.txtIdTipoGrupo.Value == 6) this.cmbPreciador.Enabled = true;
					else this.cmbPreciador.Enabled = false;				

					if ((int)this.txtIdTipoGrupo.Value == 1 || (int)this.txtIdTipoGrupo.Value == 3 || (int)this.txtIdTipoGrupo.Value == 6) this.cmbArticuloSRI.Value = 80;
					else
					{
						this.cmbArticuloSRI.Enabled = true;
						this.txtCodigo.Enabled = true;
					}

					if ((int)this.txtIdTipoGrupo.Value == 5)
					{
						this.cmbSeccion.Enabled = true;
						this.cmbCentroCosto.Enabled = true;
						this.txtDepreciacion.Enabled = true;
						this.txtResidual.Enabled = true;
						this.txtTiempo.Enabled = true;
						this.dtFechaUltimo.Enabled = true;
					}
					else
					{
						this.cmbSeccion.Value = System.DBNull.Value;
						this.cmbCentroCosto.Value = System.DBNull.Value;
						this.txtDepreciacion.Value = 0;
						this.txtResidual.Value = 0;
						this.txtTiempo.Value = 0;
						this.dtFechaUltimo.Value = System.DBNull.Value;

						this.cmbSeccion.Enabled = false;
						this.cmbCentroCosto.Enabled = false;
						this.txtDepreciacion.Enabled = false;
						this.txtResidual.Enabled = false;
						this.txtTiempo.Enabled = false;
						this.dtFechaUltimo.Enabled = false;
					}
				}
			}
		}

		private void cmbGrupo_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}

		private void cmbGrupo_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.cmbSubGrupo.Focus();
		}

		private void cmbSubGrupo_ValueChanged(object sender, System.EventArgs e)
		{
			if (bloquearEventos) return;  // NO ejecutar nada mientras se cargan datos

			if (this.cmbSubGrupo.ActiveRow != null)
			{
				string sSQL = string.Format("Exec ListaCategoriasMarcas {0}, 1", (int)this.cmbSubGrupo.Value);
				this.cmbCategoria.DataSource = Funcion.dvProcedimiento(cdsSeteoF, sSQL);

				sSQL = string.Format("Exec ListaCategoriasMarcas {0}, 2", (int)this.cmbSubGrupo.Value);
				this.cmbMarca.DataSource = Funcion.dvProcedimiento(cdsSeteoF, sSQL);

				if (bEdicion)
				{
					string sSQLGE = string.Format("Exec ArticuloInformacionGarantiaExtendida {0}", (int)this.cmbSubGrupo.Value);
					SqlDataReader dr = Funcion.rEscalarSQL(cdsSeteoF, sSQLGE, true);
					dr.Read();
					if (dr.HasRows)
					{
						this.chkGarantiaExtendida.Checked = dr.GetBoolean(0);
						this.txtPorcentajeGE.Value = dr.GetDecimal(1);
						this.txtTiempoGE.Value = dr.GetInt32(2);
						this.chkServicioTecnicoDomicilio.Checked = dr.GetBoolean(3);
					}
					dr.Close();
				}
			}

			if (bEdicion && bNuevo)
				if (this.cmbSubGrupo.ActiveRow != null)
					GeneraDescripcion();
		}

		private void cmbSubGrupo_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.cmbCategoria.Focus();
		}

		private void cmbCategoria_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.cmbMarca.Focus();
		}

		private void txtModelo_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.txtNParte.Focus();
		}

		private void txtNParte_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.cmbColor.Focus();
		}

		private void cmbColor_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.cmbPreciador.Focus();			
		}

		private void txtCodigo_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter)
			{
				if (this.txtCodigoReemplazo.Enabled) this.txtCodigoReemplazo.Focus();
				else this.txtNombre.Focus();
			}
		}

		private void txtCodigoReemplazo_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.txtLink.Focus();
		}

		private void txtNombre_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter)
			{
				if (this.cmbArticuloSRI.Enabled) this.cmbArticuloSRI.Focus();
				else this.txtTamaño.Focus();
			}
		}

		private void txtTamaño_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.txtPeriodoGarantia.Focus();
		}

		private void txtPeriodoGarantia_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.txtVidaUtil.Focus();
		}

		private void txtVidaUtil_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.txtCantMinima.Focus();
		}

		private void txtCantMinima_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.txtCantMaxima.Focus();
		}

		private void txtCantMaxima_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.rbgOrigen.Focus();
		}

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);		
		}

		private void txtNombre_ValueChanged(object sender, System.EventArgs e)
		{
		
		}

		private void rbgOrigen_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) 
			{
				if ((int)this.rbgOrigen.EditValue == 2) this.cmbArancel.Focus();
				else if ((int)this.rbgOrigen.EditValue == 1) this.chkPrecio.Focus();
			}
		}

		private void cmbArancel_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.txtPeso.Focus();
		}

		private void txtPeso_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.txtAncho.Focus();
		}

		private void txtAncho_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.txtLargo.Focus();
		}

		private void txtLargo_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.txtAlto.Focus();
		}

		private void txtAlto_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.chkPrecio.Focus();
		}

		private void chkPrecio_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter)
			{
				if (this.chkPrecio.Checked) this.txtPrecioPI.Focus();
				else
				{
					if ((int)this.txtIdTipoGrupo.Value == 5) this.cmbSeccion.Focus();
					else this.chkCombo.Focus();
				}
			}
		}

		private void txtPrecioPI_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.txtPrecioPA.Focus();
		}

		private void txtPrecioPA_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) 
			{
				if ((int)this.txtIdTipoGrupo.Value == 5) this.cmbSeccion.Focus();
				else this.chkCombo.Focus();
			}
		}

		private void cmbSeccion_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.cmbCentroCosto.Focus();
		}

		private void cmbCentroCosto_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.txtDepreciacion.Focus();
		}

		private void txtDepreciacion_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.txtResidual.Focus();
		}

		private void txtTiempo_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.dtFechaUltimo.Focus();
		}

		private void dtFechaUltimo_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.chkCombo.Focus();
		}

		private void chkCombo_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.chkManejaSerial.Focus();
		}

		private void chkManejaSerial_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.chkSerialAlterna.Focus();
		}

		private void chkSerialAlterna_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.chkFaturable.Focus();
		}

		private void chkFaturable_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.chkAplicDescuento.Focus();
		}

		private void chkAplicDescuento_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.chkDescontinuado.Focus();
		}

		private void chkDescontinuado_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.chkDescCompras.Focus();
		}

		private void chkDescCompras_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.chkDescVentas.Focus();
		}

		private void chkDescVentas_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.btnGuardar.Focus();
		}

		private void cmbGrupo_Enter(object sender, System.EventArgs e)
		{
			this.cmbGrupo.ToggleDropdown();
		}

		private void cmbSubGrupo_Enter(object sender, System.EventArgs e)
		{
			this.cmbSubGrupo.ToggleDropdown();
		}

		private void cmbCategoria_Enter(object sender, System.EventArgs e)
		{
			this.cmbCategoria.ToggleDropdown();
		}
		
		private void cmbColor_Enter(object sender, System.EventArgs e)
		{
			this.cmbColor.ToggleDropdown();
		}

		private void cmbArticuloSRI_Enter(object sender, System.EventArgs e)
		{
			this.cmbArticuloSRI.ToggleDropdown();
		}

		private void cmbArancel_Enter(object sender, System.EventArgs e)
		{
			this.cmbArancel.ToggleDropdown();
		}

		private void cmbSeccion_Enter(object sender, System.EventArgs e)
		{
			this.cmbSeccion.ToggleDropdown();
		}

		private void cmbCentroCosto_Enter(object sender, System.EventArgs e)
		{
			this.cmbCentroCosto.ToggleDropdown();
		}

		private void txtModelo_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			//			if (Char.IsLetterOrDigit(e.KeyChar))
			//			{
			//				e.Handled = false;
			//			}
			//			else if (Char.IsNumber(e.KeyChar))
			//			{
			//				e.Handled = false;
			//			}
			//			else if (Char.IsControl(e.KeyChar))
			//			{
			//				e.Handled = false;
			//			}
			//			else
			//			{
			//				e.Handled = true;
			//			}
		}

		private void cmbGrupo_Validating(object sender, System.ComponentModel.CancelEventArgs e)
		{
			if (bEdicion) 
			{
				if (this.cmbGrupo.ActiveRow == null)
				{
					MessageBox.Show("Seleccione un Grupo", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					this.cmbGrupo.Focus();
					e.Cancel = true;
				}
			}
		}

		private void cmbSubGrupo_Validating(object sender, System.ComponentModel.CancelEventArgs e)
		{
			if (bEdicion) 
			{
				if (this.cmbSubGrupo.ActiveRow == null)
				{
					MessageBox.Show("Seleccione un SubGrupo", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					this.cmbSubGrupo.Focus();
					e.Cancel = true;
				}
			}
		}

		private void txtModelo_Validating(object sender, System.ComponentModel.CancelEventArgs e)
		{
			if (bEdicion)
			{
				if (this.txtModelo.Text.ToString().Length == 0)
				{
					MessageBox.Show("Ingrese el modelo", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					this.txtModelo.Focus();
					e.Cancel = true;
				}
			}
		}

		private void txtNombre_Validating(object sender, System.ComponentModel.CancelEventArgs e)
		{
			if (bEdicion)
			{
				if (this.txtNombre.Text.ToString().Length == 0)
				{
					MessageBox.Show("Ingrese la descripcion", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					this.txtNombre.Focus();
					e.Cancel = true;
				}
			}
		}

		private void cmbColor_Validated(object sender, System.EventArgs e)
		{
			//			if (bEdicion) this.txtCodigo.Text = sCodigo;
		}

		private void rbgOrigen_SelectedIndexChanged(object sender, System.EventArgs e)
		{
			if (bEdicion)
			{
				if ((int)this.rbgOrigen.EditValue == 2)
				{
					this.cmbArancel.Enabled = true;
					this.txtPeso.Enabled = true;
					this.txtAncho.Enabled = true;
					this.txtLargo.Enabled = true;
					this.txtAlto.Enabled = true;
				}
				else if ((int)this.rbgOrigen.EditValue == 1)
				{
					this.cmbArancel.Enabled = false;
					//this.txtPeso.Enabled = false;
					this.txtAncho.Enabled = false;
					this.txtLargo.Enabled = false;
					this.txtAlto.Enabled = false;

					this.cmbArancel.Value = System.DBNull.Value;
					this.txtPeso.Value = 0;
					this.txtAncho.Value = 0;
					this.txtLargo.Value = 0;
					this.txtAlto.Value = 0;
				}
			}
		}

		private void chkPrecio_CheckedChanged(object sender, System.EventArgs e)
		{
			if (bEdicion) 
			{
				if (this.chkPrecio.Checked)
				{
					this.txtPrecioPI.Enabled = true;
					this.txtPrecioPA.Enabled = true;
					this.txtPrecioTarjeta.Enabled = true;
					this.txtPrecioCostoUltimo.Enabled = true;
					this.txtCostoPayJoy.Enabled = true;
					this.txtPrecioTCsinIntereses.Enabled = true;
					this.txtPrecioDescEmpleado.Enabled = true;
					this.txtPrecioAutoconsumo.Enabled = true;
					if (this.chkPayJoy.Checked) this.txtCostoPayJoy.Enabled = true;
					if (!this.chkPayJoy.Checked) this.txtCostoPayJoy.Enabled = false;
				}
				else 
				{
					this.txtPrecioPI.Enabled = false;
					this.txtPrecioPA.Enabled = false;
					this.txtPrecioTarjeta.Enabled = false;
					this.txtPrecioCostoUltimo.Enabled = false;
					this.txtCostoPayJoy.Enabled = false;
					this.txtPrecioTCsinIntereses.Enabled = false;
					this.txtPrecioDescEmpleado.Enabled = false;
					this.txtPrecioAutoconsumo.Enabled = false;
					this.txtCostoPayJoy.Enabled = false;


					if (bNuevo)
					{
						this.txtPrecioPI.Value = 0;
						this.txtPrecioPA.Value = 0;
						this.txtPrecioTarjeta.Value = 0;
						this.txtPrecioPVP.Value = 0;
						this.txtPrecioCostoUltimo.Value = 0;
						this.txtCostoPayJoy.Value = 0;
						this.txtPrecioTCsinIntereses.Value = 0;
						this.txtPrecioDescEmpleado.Value = 0;
						this.txtPrecioAutoconsumo.Value = 0;
						this.txtCostoPayJoy.Value = 0;
					}
				}
			}
		}

		private void txtPrecioPA_Validated(object sender, System.EventArgs e)
		{
			if (bEdicion)
			{
				if (this.chkPrecio.Checked)
				{		
					SqlDataReader dr = Funcion.rEscalarSQL(cdsSeteoF, "Select PorcContTarjeta, PorcContCredito From SeteoF", true);
					dr.Read();
					if (dr.HasRows)
					{
						dPorcTarjeta = Convert.ToDecimal(dr.GetValue(0));
						dPorcCredito = Convert.ToDecimal(dr.GetValue(1));
					}
					dr.Close();

					this.txtPrecioTarjeta.Value = Math.Round(Convert.ToDecimal(this.txtPrecioPA.Value) * dPorcTarjeta, 2);
					this.txtPrecioPVP.Value = Math.Round(Convert.ToDecimal(this.txtPrecioTarjeta.Value) * dPorcCredito, 2);
				}
			}
		}

		private void txtPrecioPA_Validating(object sender, System.ComponentModel.CancelEventArgs e)
		{
			if (bEdicion)
			{
				if (this.chkPrecio.Checked)
				{			
					if (this.txtPrecioPA.Value == System.DBNull.Value || Convert.ToDecimal(this.txtPrecioPA.Value) == 0.00m)
					{
						MessageBox.Show("Dijite el Precio Contado", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
						this.txtPrecioPA.Focus();			
					}
				}
			}
		}

		private void cmbColor_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}

		private void txtPrecioPA_ValueChanged(object sender, System.EventArgs e)
		{
		
		}

		private void txtPrecioPI_Validating(object sender, System.ComponentModel.CancelEventArgs e)
		{
			if (bEdicion)
			{
				if (this.chkPrecio.Checked)
				{
					if (MenuLatinium.stIdDB != 34)
					{
						if (this.txtPrecioPI.Value == System.DBNull.Value || Convert.ToDecimal(this.txtPrecioPI.Value) == 0.00m)
						{
							MessageBox.Show("Dijite el Precio Distribuidor", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
							this.txtPrecioPI.Focus();						
						}
					}
				}
			}
		}

		private void txtResidual_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.txtTiempo.Focus();
		}

		private void txtIva_Click(object sender, System.EventArgs e)
		{
			this.txtIva.SelectAll();
		}

		private void txtIva_Enter(object sender, System.EventArgs e)
		{
			this.txtIva.SelectAll();
		}

		private void txtTamaño_Click(object sender, System.EventArgs e)
		{
			this.txtTamaño.SelectAll();
		}

		private void txtTamaño_Enter(object sender, System.EventArgs e)
		{
			this.txtTamaño.SelectAll();
		}

		private void txtPeriodoGarantia_Click(object sender, System.EventArgs e)
		{
			this.txtPeriodoGarantia.SelectAll();
		}

		private void txtPeriodoGarantia_Enter(object sender, System.EventArgs e)
		{
			this.txtPeriodoGarantia.SelectAll();
		}

		private void txtVidaUtil_Click(object sender, System.EventArgs e)
		{
			this.txtVidaUtil.SelectAll();
		}

		private void txtVidaUtil_Enter(object sender, System.EventArgs e)
		{
			this.txtVidaUtil.SelectAll();
		}

		private void txtSaldo_Click(object sender, System.EventArgs e)
		{
			this.txtSaldo.SelectAll();
		}

		private void txtSaldo_Enter(object sender, System.EventArgs e)
		{
			this.txtSaldo.SelectAll();
		}

		private void txtCantMinima_Click(object sender, System.EventArgs e)
		{
			this.txtCantMinima.SelectAll();
		}

		private void txtCantMinima_Enter(object sender, System.EventArgs e)
		{
			this.txtCantMinima.SelectAll();
		}

		private void txtCantMaxima_Click(object sender, System.EventArgs e)
		{
			this.txtCantMaxima.SelectAll();
		}

		private void txtCantMaxima_Enter(object sender, System.EventArgs e)
		{
			this.txtCantMaxima.SelectAll();
		}

		private void txtPeso_Click(object sender, System.EventArgs e)
		{
			this.txtPeso.SelectAll();
		}

		private void txtPeso_Enter(object sender, System.EventArgs e)
		{
			this.txtPeso.SelectAll();
		}

		private void txtAncho_Click(object sender, System.EventArgs e)
		{
			this.txtAncho.SelectAll();
		}

		private void txtAncho_Enter(object sender, System.EventArgs e)
		{
			this.txtAncho.SelectAll();
		}

		private void txtLargo_Click(object sender, System.EventArgs e)
		{
			this.txtLargo.SelectAll();
		}

		private void txtLargo_Enter(object sender, System.EventArgs e)
		{
			this.txtLargo.SelectAll();
		}

		private void txtAlto_Click(object sender, System.EventArgs e)
		{
			this.txtAlto.SelectAll();
		}

		private void txtAlto_Enter(object sender, System.EventArgs e)
		{
			this.txtAlto.SelectAll();
		}

		private void txtPrecioPI_Click(object sender, System.EventArgs e)
		{
			this.txtPrecioPI.SelectAll();
		}

		private void txtPrecioPI_Enter(object sender, System.EventArgs e)
		{
			this.txtPrecioPI.SelectAll();
		}

		private void txtPrecioPA_Click(object sender, System.EventArgs e)
		{
			this.txtPrecioPA.SelectAll();
		}

		private void txtPrecioPA_Enter(object sender, System.EventArgs e)
		{
			this.txtPrecioPA.SelectAll();
		}

		private void txtPrecioIncr1_Click(object sender, System.EventArgs e)
		{
			this.txtPrecioIncr1.SelectAll();
		}

		private void txtPrecioIncr1_Enter(object sender, System.EventArgs e)
		{
			this.txtPrecioIncr1.SelectAll();
		}

		private void txtPrecioIncr2_Click(object sender, System.EventArgs e)
		{
			this.txtPrecioIncr2.SelectAll();
		}

		private void txtPrecioIncr2_Enter(object sender, System.EventArgs e)
		{
			this.txtPrecioIncr2.SelectAll();
		}

		private void txtPrecioCostoUltimo_Click(object sender, System.EventArgs e)
		{
			this.txtPrecioCostoUltimo.SelectAll();
		}

		private void txtPrecioCostoUltimo_Enter(object sender, System.EventArgs e)
		{
			this.txtPrecioCostoUltimo.SelectAll();
		}
		
		private void txtDepreciacion_Click(object sender, System.EventArgs e)
		{
			this.txtDepreciacion.SelectAll();
		}

		private void txtDepreciacion_Enter(object sender, System.EventArgs e)
		{
			this.txtDepreciacion.SelectAll();
		}

		private void txtResidual_Click(object sender, System.EventArgs e)
		{
			this.txtResidual.SelectAll();
		}

		private void txtResidual_Enter(object sender, System.EventArgs e)
		{
			this.txtResidual.SelectAll();
		}

		private void txtTiempo_Click(object sender, System.EventArgs e)
		{
			this.txtTiempo.SelectAll();
		}

		private void txtTiempo_Enter(object sender, System.EventArgs e)
		{
			this.txtTiempo.SelectAll();
		}

		private void Consulta (int idArticulo)
		{
			try
			{	
				string sSQL = string.Format("Exec ConsultaIndividualArticulo {0}", idArticulo);
				SqlDataReader drArticulo = Funcion.rEscalarSQL(cdsSeteoF, sSQL, true);
				drArticulo.Read();
				if (drArticulo.HasRows)
				{	
					this.txtIdArticulo.Value = idArticulo;
					this.txtIdTipoGrupo.Value = drArticulo.GetInt32(1);
					this.cmbGrupo.Value = drArticulo.GetInt32(2);
					this.cmbSubGrupo.Value = drArticulo.GetInt32(3);
					if (drArticulo.GetValue(4) != System.DBNull.Value) this.cmbCategoria.Value = drArticulo.GetInt32(4);
					if (drArticulo.GetValue(5) != System.DBNull.Value) this.cmbMarca.Value = drArticulo.GetInt32(5);
					this.txtModelo.Text = drArticulo.GetValue(6).ToString();
					this.txtNParte.Text = drArticulo.GetValue(7).ToString();
					if (drArticulo.GetValue(8) != System.DBNull.Value) this.cmbColor.Value = (int)drArticulo.GetInt32(8);
					this.txtCodigo.Text = drArticulo.GetValue(9).ToString();
					this.txtCodigoReemplazo.Text = drArticulo.GetValue(10).ToString();
					this.txtNombre.Text = drArticulo.GetValue(11).ToString();
					if (drArticulo.GetValue(12) != System.DBNull.Value) this.cmbArticuloSRI.Value = drArticulo.GetInt32(12);
					this.txtIva.Value = drArticulo.GetDecimal(13);
					this.txtTamaño.Value = drArticulo.GetInt32(14);
					this.txtPeriodoGarantia.Value = drArticulo.GetInt32(15);
					this.txtVidaUtil.Value = drArticulo.GetInt32(16);
					this.txtSaldo.Value = drArticulo.GetDecimal(17);
					this.txtCantMinima.Value = drArticulo.GetDecimal(18);
					this.txtCantMaxima.Value = drArticulo.GetDecimal(19);
					this.rbgOrigen.EditValue = drArticulo.GetInt32(20);
					if (drArticulo.GetValue(21) != System.DBNull.Value) this.cmbArancel.Value = drArticulo.GetInt32(21);
					this.txtPeso.Value = drArticulo.GetDecimal(22);
					this.txtAncho.Value = drArticulo.GetDecimal(23);
					this.txtLargo.Value = drArticulo.GetDecimal(24);
					this.txtAlto.Value = drArticulo.GetDecimal(25);
					this.chkPrecio.Checked = drArticulo.GetBoolean(26);
					this.txtPrecioPI.Value = drArticulo.GetDecimal(27);
					this.txtPrecioIncr1.Value = drArticulo.GetDecimal(28);
					this.txtPrecioPA.Value = drArticulo.GetDecimal(29);
					this.txtPrecioIncr2.Value = drArticulo.GetDecimal(30);
					this.txtPrecioTarjeta.Value = drArticulo.GetDecimal(31);
					this.txtPrecioPVP.Value = drArticulo.GetDecimal(32);
					if (drArticulo.GetValue(33) != System.DBNull.Value) this.txtPrecioCostoUltimo.Value = drArticulo.GetDecimal(33);					
					if (drArticulo.GetValue(35) != System.DBNull.Value) this.cmbSeccion.Value = drArticulo.GetInt32(35);
					if (drArticulo.GetValue(36) != System.DBNull.Value) this.cmbCentroCosto.Value = drArticulo.GetInt32(36);
					this.txtDepreciacion.Value = drArticulo.GetDecimal(37);
					this.txtResidual.Value = drArticulo.GetDecimal(38);
					this.txtTiempo.Value = drArticulo.GetDecimal(39);
					if (drArticulo.GetValue(40) != System.DBNull.Value) this.dtFechaUltimo.Value = drArticulo.GetDateTime(40);
					this.chkCombo.Checked = drArticulo.GetBoolean(41);
					this.chkManejaSerial.Checked = drArticulo.GetBoolean(42);
					this.chkSerialAlterna.Checked = drArticulo.GetBoolean(43);
					this.chkFaturable.Checked = drArticulo.GetBoolean(44);
					this.chkAplicDescuento.Checked = drArticulo.GetBoolean(45);
					this.chkDescontinuado.Checked = drArticulo.GetBoolean(46);
					this.chkDescCompras.Checked = drArticulo.GetBoolean(47);
					this.chkDescVentas.Checked = drArticulo.GetBoolean(48);
					this.chkTransferible.Checked = drArticulo.GetBoolean(49);
					this.txtCodArtEasyMax.Text = drArticulo.GetValue(50).ToString();
					this.cmbPreciador.Value = drArticulo.GetInt32(51);
					this.chkInventario.Checked = drArticulo.GetBoolean(52);
					this.chkImpuestoIVA.Checked = drArticulo.GetBoolean(53);
					this.chkGarantiaExtendida.Checked = drArticulo.GetBoolean(54);
					this.txtPorcentajeGE.Value = drArticulo.GetDecimal(55);
					this.txtTiempoGE.Value = drArticulo.GetInt32(56);
					this.chkServicioTecnicoDomicilio.Checked = drArticulo.GetBoolean(57);
					this.txtLink.Text = drArticulo.GetString(58);
					this.chkObsequios.Checked = drArticulo.GetBoolean(59);
					this.chkcatalogo.Checked = drArticulo.GetBoolean(60);
					this.chkConsignacion.Checked = drArticulo.GetBoolean(62);
					this.txtCSCostoEfectivo.Value = drArticulo.GetDecimal(63);
					this.txtCSCostoCredito.Value = drArticulo.GetDecimal(64);
					this.txtCodigoProveedor.Value = drArticulo.GetString(65);
					if (drArticulo.GetValue(66) != System.DBNull.Value) this.txtPrecioTCsinIntereses.Value = drArticulo.GetDecimal(66);
					this.chkProductoWeb.Checked = drArticulo.GetBoolean(67);
					this.chkPayJoy.Checked = drArticulo.GetBoolean(68);
					this.txtCodigoPayJoy.Value = drArticulo.GetString(69);
					
					this.txtOrigenMoto.Text = drArticulo.GetString(70).ToString();
					
					this.txtfechaMoto.Value = drArticulo.GetInt32(71);
					this.txtTonelaje.Value = drArticulo.GetDecimal(72);
					this.txtCilindraje.Value = drArticulo.GetInt32(73);
					this.txtPasajeros.Value = drArticulo.GetInt32(74);
					this.chkMoto.Checked = drArticulo.GetBoolean(75);
					this.chkMigrar.Checked= drArticulo.GetBoolean(76);
					this.chkManejaSerie.Checked= drArticulo.GetBoolean(77);
					this.txtnovasoft.Value = drArticulo.GetString(78);
					this.txtCostoPayJoy.Value = drArticulo.GetDecimal(79);
					this.chkPrecioMovil.Checked = drArticulo.GetBoolean(80);
					this.txtPrecioDescEmpleado.Value = drArticulo.GetDecimal(81);
					this.txtPrecioAutoconsumo.Value = drArticulo.GetDecimal(82);
					this.chkDepartamento.Checked = drArticulo.GetBoolean(83);
					this.chkDepartamento1.Checked = drArticulo.GetBoolean(84);
					this.chkEfectivo.Checked = drArticulo.GetBoolean(85);
					this.txtDescripcionWeb.Text = drArticulo.GetValue(86).ToString();
					if (drArticulo.GetValue(87) != System.DBNull.Value) this.txtTituloWeb.Text = drArticulo.GetString(87).ToString();
				}
				drArticulo.Close();	

				if (miAcceso.Editar) this.btnEditar.Enabled = true;            						
				if (miAcceso.Imprimir) this.btnImprimir.Enabled = true;						
				this.btnCancelar.Enabled = true;
				this.btnCaracteristicas.Enabled = true;
				if (miAcceso.BAccesoBoton)this.btnAtributos.Enabled = true;
				
				idSubgrupo = (int)this.cmbSubGrupo.Value;
			}
			catch(Exception Exc)
			{
				MessageBox.Show(string.Format("Error al Consultar: {0}", Exc), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
			}			
		}

		private void btnConsultar_Click(object sender, System.EventArgs e)
		{
			btnCancelar_Click(sender, e);

			using (frmBuscaArticulosInventarioActivoFijo Busqueda = new frmBuscaArticulosInventarioActivoFijo())
			{
				if (DialogResult.OK == Busqueda.ShowDialog())
				{
					this.Consulta((int)Busqueda.ultraGrid1.ActiveRow.Cells["idArticulo"].Value);
				}
			}		
		}

		private void cmbSubGrupo_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}

		private void btnEditar_Click(object sender, System.EventArgs e)
		{			
			int iCantUs = Funcion.iEscalarSQL(cdsSeteoF, string.Format("Select COUNT(*) From DetCompra dc Inner Join Compra c On c.idCompra = dc.idCompra Where c.Borrar = 0 And c.Estado Not In (4, 6) And dc.idArticulo = {0}", (int)this.txtIdArticulo.Value));
			if (iCantUs == 0)
			{		
				//this.cmbGrupo.Enabled = true;				
				this.cmbSubGrupo.Enabled = true;				
				this.cmbMarca.Enabled = true;
				this.txtModelo.Enabled = true;
				this.cmbColor.Enabled = true;
				this.txtNombre.Enabled = true;
			}

			this.cmbGrupo.Enabled = true;			
			this.cmbCategoria.Enabled = true;
			this.txtNParte.Enabled = true;			
			this.txtCodigoReemplazo.Enabled = true;			
			this.cmbArticuloSRI.Enabled = true;
			this.txtTamaño.Enabled = true;
			this.txtPeriodoGarantia.Enabled = true;
			this.txtVidaUtil.Enabled = true;
			this.txtCantMinima.Enabled = true;
			this.txtCantMinima.Enabled = true;
			this.rbgOrigen.Enabled = true;		
			this.chkcatalogo.Enabled = true;
			this.txtPeso.Enabled = true;
			this.txtDescripcionWeb.Enabled = true;
			this.txtTituloWeb.Enabled = true;

			if ((int)this.rbgOrigen.EditValue == 2)
			{
				this.cmbArancel.Enabled = true;
				
				this.txtAncho.Enabled = true;
				this.txtLargo.Enabled = true;
				this.txtAlto.Enabled = true;
			}
			
			this.chkGarantiaExtendida.Enabled = true;

			if (this.chkGarantiaExtendida.Checked)
			{
				this.chkServicioTecnicoDomicilio.Enabled = true;
				this.txtPorcentajeGE.Enabled = true;
				this.txtTiempoGE.Enabled = true;				
			}

			this.chkConsignacion.Enabled = true;

			if (this.chkConsignacion.Checked)
			{				
				this.txtCSCostoEfectivo.Enabled = true;
				this.txtCSCostoCredito.Enabled = true;
				this.txtCodigoProveedor.Enabled = true;
			}

			this.chkMoto.Enabled = true;

			if (this.chkMoto.Checked)
			{				
				this.txtOrigenMoto.Enabled = true;
				this.txtfechaMoto.Enabled = true;
				this.txtTonelaje.Enabled = true;
				this.txtCilindraje.Enabled = true;
				//this.txtPasajeros.Enabled = true;
			}

			//			if (this.chkMigrar.Checked)
			//			{				
			//				this.txtnovasoft.Enabled = true;
			//				this.chkManejaSerie.Enabled= true;
			//				}

			this.chkPrecio.Enabled = true;

			if (this.chkPrecio.Checked)
			{
				//				this.txtPrecioPI.Enabled = true;
				//				this.txtPrecioPA.Enabled = true;
				this.txtPrecioTarjeta.Enabled = true;				
				this.txtPrecioCostoUltimo.Enabled = true;
				this.txtCostoPayJoy.Enabled = true;
				this.txtPrecioAutoconsumo.Enabled = true;
				this.txtPrecioDescEmpleado.Enabled = true;
			}

			if ((int)this.txtIdTipoGrupo.Value == 5)
			{
				this.cmbSeccion.Enabled = true;
				this.cmbCentroCosto.Enabled = true;
				this.txtDepreciacion.Enabled = true;
				this.txtResidual.Enabled = true;
				this.txtTiempo.Enabled = true;
				this.dtFechaUltimo.Enabled = true;
			}

			if ((int)this.txtIdTipoGrupo.Value == 1 || (int)this.txtIdTipoGrupo.Value == 6) this.cmbPreciador.Enabled = true;
			
			this.chkCombo.Enabled = true;
			if (miAcceso.BManejaSerial) this.chkManejaSerial.Enabled = true;
			this.chkObsequios.Enabled = true;
			this.chkSerialAlterna.Enabled = true;
			this.chkFaturable.Enabled = true;
			this.chkAplicDescuento.Enabled = true;
			this.chkDescontinuado.Enabled = true;
			this.chkDescCompras.Enabled = true;
			this.chkDescVentas.Enabled = true;
			this.chkTransferible.Enabled = true;
			this.chkInventario.Enabled = true;
			this.chkEfectivo.Enabled = true;
			this.chkImpuestoIVA.Enabled = true;
			this.chkProductoWeb.Enabled = true;
			this.chkPayJoy.Enabled = true;
			this.chkPrecioMovil.Enabled = true;
			this.chkMigrar.Enabled = true;
			this.chkDepartamento.Enabled = true;
			this.chkDepartamento1.Enabled = true;
			this.txtLink.Enabled = true;
			this.chkEfectivo.Enabled = true;

			this.btnNuevo.Enabled = false;
			this.btnConsultar.Enabled = false;
			this.btnEditar.Enabled = false;
			this.btnImprimir.Enabled = false;
			this.btnGuardar.Enabled = true;
			this.btnCaracteristicas.Enabled = false;

			string stMensaje = string.Format("Select descripcionArticulo From SeteoF");
			bool iDescripcion = Funcion.bEscalarSQL(cdsSeteoF, stMensaje, true);
			if (iDescripcion)
			{				
				this.cmbGrupo.Enabled = true;
				this.cmbSubGrupo.Enabled = true;
				this.cmbMarca.Enabled = true;
				this.cmbCategoria.Enabled = true;
				this.txtModelo.Enabled = true;
				this.txtNParte.Enabled = false;
				this.cmbPreciador.Enabled = false;
				this.txtLink.Enabled = false;
				this.cmbArticuloSRI.Enabled = false;
				this.txtTamaño.Enabled = false;				
				this.txtCodigoReemplazo.Enabled = false;
				this.txtPeriodoGarantia.Enabled = false;
				this.txtVidaUtil.Enabled = false;
			}

			bNuevo = false;
			bEdicion = true;
		}

		private void btnImprimir_Click(object sender, System.EventArgs e)
		{
			Cursor = Cursors.WaitCursor;
			Reporte miReporte = new Reporte("Articulo Lista.rpt");
			miReporte.MdiParent = MdiParent;
			miReporte.Show();
			Cursor = Cursors.Default;
		}

		private void cmbColor_Validating(object sender, System.ComponentModel.CancelEventArgs e)
		{
			if (bEdicion) 
			{
				#region Codigo Automatico	
				int IdSubGrupo = 0;
				int IdMarca = 0;
				int IdColor = 0;

				sCodigo = "";

				if (this.cmbSubGrupo.ActiveRow != null) IdSubGrupo = (int)this.cmbSubGrupo.Value;
				if (this.cmbMarca.ActiveRow != null) IdMarca = (int)this.cmbMarca.Value;
				if (this.cmbColor.ActiveRow != null) IdColor = (int)this.cmbColor.Value;			

				sCodigo =	Funcion.sEscalarSQL(cdsSeteoF, string.Format("Select dbo.CodigoArticulo({0}, {1}, '{2}', {3})", IdSubGrupo, IdMarca, this.txtModelo.Value.ToString().Trim(), IdColor));

				if (Funcion.iEscalarSQL(cdsSeteoF, string.Format("Select COUNT(*) From Articulo Where Codigo = '{0}' And idArticulo <> {1}", sCodigo, (int)this.txtIdArticulo.Value)) > 0)
				{
					MessageBox.Show("El Artículo ya existe", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					this.cmbColor.Focus();
					e.Cancel = true;
				}
				#endregion Codigo Automatico
			}
		}
		
		private void GeneraDescripcion()
		{
			if (this.cmbSubGrupo.ActiveRow != null && this.cmbMarca.ActiveRow != null && this.txtModelo.Text.ToString().Length > 0)
			{
				this.txtNombre.Text = Funcion.sEscalarSQL(cdsSeteoF, string.Format("Exec GeneraDescripcionArticulo {0}, {1}, '{2}'",
					(int)this.cmbSubGrupo.Value, (int)this.cmbMarca.Value, this.txtModelo.Text.ToString()));
			}
		}

		private void Codigo()
		{
			int IdSubGrupo = 0;
			int IdMarca = 0;
			int IdColor = 0;

			if (this.cmbSubGrupo.ActiveRow != null && this.cmbMarca.ActiveRow != null && this.txtModelo.Text.ToString().Length > 0)
			{
				if (this.cmbSubGrupo.ActiveRow != null) IdSubGrupo = (int)this.cmbSubGrupo.Value;
				if (this.cmbMarca.ActiveRow != null) IdMarca = (int)this.cmbMarca.Value;
				if (this.cmbColor.ActiveRow != null) IdColor = (int)this.cmbColor.Value;			

				this.txtCodigo.Text =	Funcion.sEscalarSQL(cdsSeteoF, string.Format("Select dbo.CodigoArticulo({0}, {1}, '{2}', {3})", IdSubGrupo, IdMarca, this.txtModelo.Value.ToString().Trim(), IdColor));
			}
		}

		private void cmbSubGrupo_Validated(object sender, System.EventArgs e)
		{
			//if (bEdicion) Codigo();
		}

		private void cmbMarca_Validated(object sender, System.EventArgs e)
		{
			//if (bEdicion) Codigo();
		}

		private void txtModelo_Validated(object sender, System.EventArgs e)
		{
			//if (bEdicion) Codigo();
			if (bEdicion)
				if (this.cmbSubGrupo.ActiveRow != null)
					GeneraDescripcion();
		}

		private void cmbMarca_Validating(object sender, System.ComponentModel.CancelEventArgs e)
		{
			if (bEdicion) 
			{
				if (this.cmbMarca.ActiveRow == null)
				{
					MessageBox.Show("Seleccione una Marca", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					this.cmbMarca.Focus();
					e.Cancel = true;
				}
			}
		}

		private void btnGuardar_Click(object sender, System.EventArgs e)
		{		
			DateTime dtFecha = DateTime.Today;

			#region Validacion
			if (!Validacion.vbComboVacio(this.cmbGrupo, "Seleccione un Grupo")) return;
			if (!Validacion.vbComboVacio(this.cmbSubGrupo, "Seleccione un SubGrupo")) return;
			if (!Validacion.vbComboVacio(this.cmbMarca, "Seleccione una Marca")) return;
			if (!Validacion.vbTexto(this.txtModelo, 1, 100, "Modelo")) return;
			/*QUIITAR VALIDACION PARA CROMATIKA*/
			if (!Validacion.vbComboVacio(this.cmbPreciador, "Seleccione un Preciador")) return;
			if(MenuLatinium.stDirFacturacion == "POINT" || MenuLatinium.stDirFacturacion == "CREDISOLUTIONS")
			{
				if (!Validacion.vbTexto(this.txtNombre, 5, 250, "Descripcion")) return;
			}
			else 
			{
				if (!Validacion.vbTexto(this.txtNombre, 1, 100, "Descripcion")) return;
			}
			if (Funcion.bEscalarSQL(cdsSeteoF, string.Format("Exec RetornaValPulgadas {0}", (int)this.cmbSubGrupo.Value), true) && (int)this.txtTamaño.Value == 0)
			{
				MessageBox.Show("Ingrese el Tamaño del Producto", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.txtTamaño.Focus();				
				return;
			}			
			if (!Validacion.vbCampoEnteroVacio(this.txtPeriodoGarantia, "Ingrese el Periodo de Garantia", 1, 3600)) return;
			
			if (Funcion.iEscalarSQL(cdsSeteoF, string.Format("Select dbo.Existencias('{0}', {1}, 0)", dtFecha.ToString("yyyyMMdd"), (int)this.txtIdArticulo.Value)) > 0 && (this.chkDescontinuado.Checked || this.chkDescCompras.Checked || this.chkDescVentas.Checked))
			{
				MessageBox.Show("No puede Descontinuar un Articulo Con Existencias", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.chkDescontinuado.Focus();
				return;
			}
			//			if ((bool)this.chkMigrar.Checked &&(this.txtnovasoft.Text.Length == 0 ))
			//			{
			//				MessageBox.Show("Ingrese  Codigo Novasoft ", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
			//				this.txtnovasoft.Focus();					
			//				return;
			//			}
			//validacion  novasoft
			//			if (Funcion.iEscalarSQL(cdsSeteoF, string.Format("Select COUNT(*) From Articulo Where Cod_novasoft is not null and Cod_novasoft<>'' AND  Cod_novasoft = '{0}' And idArticulo <> {1}", this.txtnovasoft.Text.ToString(), (int)this.txtIdArticulo.Value)) > 0)
			//			{
			//				string sMensaje = this.txtnovasoft.Text.ToString();
			//				if (DialogResult.Yes ==	MessageBox.Show(string.Format("El codigo Novasoft '{0}' esta registrado en dos o mas productos distintos  ¿Desea  continuar?", sMensaje), "Point Technology", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2))
			//				{
			//					//MessageBox.Show("El Artículo ya existe", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
			//					//this.cmbGrupo.Focus();	
			//					//return;
			//				}
			//				else
			//				{
			//					btnCancelar_Click(sender, e);
			//					return;
			//				}
			//			}	

			if (bNuevo)
				this.txtCodigo.Text =	Funcion.sEscalarSQL(cdsSeteoF, string.Format("Exec GeneraCodigoArticulo {0}", (int)this.cmbSubGrupo.Value));
			//Codigo();

			if (Funcion.iEscalarSQL(cdsSeteoF, string.Format("Select COUNT(*) From Articulo Where Codigo = '{0}' And idArticulo <> {1}", this.txtCodigo.Text.ToString(), (int)this.txtIdArticulo.Value)) > 0)
			{
				MessageBox.Show("El Artículo ya existe", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.cmbGrupo.Focus();	
				return;
			}	
			if ((int)this.rbgOrigen.EditValue == 2 && this.cmbArancel.ActiveRow == null)
			{
				MessageBox.Show("Seleccione el arancel", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
				this.txtPeso.Focus();	
				return;
			}
			if ((int)this.rbgOrigen.EditValue == 2 && (Convert.ToDecimal(this.txtPeso.Value) == 0.00m && Convert.ToDecimal(this.txtAncho.Value) == 0.00m && Convert.ToDecimal(this.txtLargo.Value) == 0.00m && Convert.ToDecimal(this.txtAlto.Value) == 0.00m))
			{
				MessageBox.Show("Ingrese el Peso o Volumén del artículo", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
				this.txtPeso.Focus();	
				return;
			}
			
			if(MenuLatinium.stIdDB != 34)
			{
				if (Convert.ToDecimal(this.txtPeso.Value) == 0.00m)
				{
					MessageBox.Show("Ingrese el Peso ", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
					this.txtPeso.Focus();	
					return;
				}
			}

			if ((int)this.rbgOrigen.EditValue == 2 && (Convert.ToDecimal(this.txtAncho.Value) > 0.00m && (Convert.ToDecimal(this.txtLargo.Value) == 0.00m || Convert.ToDecimal(this.txtAlto.Value) == 0.00m)))
			{
				MessageBox.Show("Debe ingresar Largo - Ancho - Alto", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
				this.txtLargo.Focus();					
				return;
			}			
			if ((int)this.rbgOrigen.EditValue == 2 && (Convert.ToDecimal(this.txtAlto.Value) > 0.00m && (Convert.ToDecimal(this.txtLargo.Value) == 0.00m || Convert.ToDecimal(this.txtAncho.Value) == 0.00m)))
			{
				MessageBox.Show("Debe ingresar Largo - Ancho - Alto", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
				this.txtLargo.Focus();					
				return;
			}
			if ((int)this.rbgOrigen.EditValue == 2 && (Convert.ToDecimal(this.txtLargo.Value) > 0.00m && (Convert.ToDecimal(this.txtAncho.Value) == 0.00m || Convert.ToDecimal(this.txtAlto.Value) == 0.00m)))
			{
				MessageBox.Show("Debe ingresar Largo - Ancho - Alto", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
				this.txtLargo.Focus();					
				return;
			}
			if (MenuLatinium.stIdDB != 34)
			{
				if ((bool)this.chkPrecio.Checked && (this.txtPrecioPI.Value == System.DBNull.Value || Convert.ToDecimal(this.txtPrecioPI.Value) == 0))
				{
					MessageBox.Show("Dijite el precio Distribuidor", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
					this.txtPrecioPI.Focus();					
					return;
				}
			}
			if ((bool)this.chkPrecio.Checked && (this.txtPrecioPA.Value == System.DBNull.Value || Convert.ToDecimal(this.txtPrecioPA.Value) == 0))
			{
				MessageBox.Show("Dijite el precio Contado", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
				this.txtPrecioPA.Focus();					
				return;
			}
			if ((bool)this.chkPayJoy.Checked &&(this.txtCodigoPayJoy.Text.Length == 0 ))
			{
				MessageBox.Show("Ingrese  Codigo Pay Joy ", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
				this.txtPrecioPA.Focus();					
				return;
			}

			if ((bool)this.chkPayJoy.Checked &&(this.txtCostoPayJoy.Value == null || Convert.ToDecimal(this.txtCostoPayJoy.Value) == 0))
			{
				MessageBox.Show("Ingrese  el Precio Pay Joy ", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
				this.txtCostoPayJoy.Focus();					
				return;
			}

			if ((bool)this.chkPayJoy.Checked &&(this.txtCodigoPayJoy.Text.Length == 0 ))
			{
				MessageBox.Show("Ingrese  Codigo Pay Joy ", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
				this.txtPrecioPA.Focus();					
				return;
			}

		
			if ((bool)this.chkMoto.Checked )
			{
				if (this.txtOrigenMoto.Text.Length == 0)
				{
					MessageBox.Show("Ingrese el lugar de origen de la motocicleta", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
					this.txtOrigenMoto.Focus();					
					return;
				}
				else if ((int)this.txtfechaMoto.Value < 2020)
				{
					MessageBox.Show(" El año de origen  no puede ser menor al actual ", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
					this.txtfechaMoto.Focus();					
					return;
				}
			}
			#endregion Validacion							

			#region Variables
			int IdCategoria = 0;
			int IdColor = 0;
			int IdArancel = 0;
			int IdSeccion = 0;
			int IdCentroDeCosto = 0;
			int idPreciador = 0;
			DateTime dtFechaUltimo = DateTime.Today;
			bool bFechaUltimo = false;
			
			if (this.cmbCategoria.ActiveRow != null) IdCategoria = (int)this.cmbCategoria.Value;
			if (this.cmbColor.ActiveRow != null) IdColor = (int)this.cmbColor.Value;
			if (this.cmbPreciador.ActiveRow != null) idPreciador = (int)this.cmbPreciador.Value;
			if (this.cmbArancel.ActiveRow != null) IdArancel = (int)this.cmbArancel.Value;
			if (this.cmbSeccion.ActiveRow != null) IdSeccion = (int)this.cmbSeccion.Value;
			if (this.cmbCentroCosto.ActiveRow != null) IdCentroDeCosto = (int)this.cmbCentroCosto.Value;
			if (this.dtFechaUltimo.Value != System.DBNull.Value)
			{
				dtFechaUltimo = (DateTime) this.dtFechaUltimo.Value;
				bFechaUltimo = true;
			}
			#endregion Variables

			#region Guarda Registro
			/*
			if (bNuevo)
				this.txtCodigo.Text =	Funcion.sEscalarSQL(cdsSeteoF, string.Format("Exec GeneraCodigoArticulo {0}", (int)this.cmbSubGrupo.Value));
			//Codigo();*/


			string sSQLGuarda = string.Format("Exec GuardaArticulos {0}, {1}, {2}, {3}, {4}, {5}, '{6}', '{7}', {8}, '{9}', '{10}', '{11}', {12}, {13}, {14}, {15}, {16}, {17}, {18}, {19}, {20}, {21}, {22}, {23}, {24}, {25}, {26}, {27}, {28}, {29}, {30}, {31}, {32}, {33}, {34}, {35}, {36}, {37}, {38}, {39}, '{40}', {41}, {42}, {43}, {44}, {45}, {46}, {47}, {48}, {49}, {50}, {51}, '{52}', {53}, {54}, {55}, {56}, {57}, {58}, '{59}', {60}, {61}, {62}, {63}, {64}, {65}, '{66}', {67}, {68},{69},'{70}','{71}','{72}','{73}','{74}','{75}','{76}','{77}','{78}','{79}',{80}, '{81}', {82}, {83}, {84}, {85},{86},'{87}', '{88}'", 
				(int)this.txtIdArticulo.Value, (int)this.txtIdTipoGrupo.Value, (int)this.cmbGrupo.Value, (int)this.cmbSubGrupo.Value, IdCategoria, (int)this.cmbMarca.Value, this.txtModelo.Text.ToString(), 
				this.txtNParte.Text.ToString(),
				IdColor, this.txtCodigo.Text.ToString(), this.txtCodigoReemplazo.Text.ToString(), this.txtNombre.Text.ToString(), (int)this.cmbArticuloSRI.Value, Convert.ToDecimal(this.txtIva.Value),
				(int)this.txtTamaño.Value, (int)this.txtPeriodoGarantia.Value, (int)this.txtVidaUtil.Value, Convert.ToDecimal(this.txtSaldo.Value), Convert.ToDecimal(this.txtCantMinima.Value), 
				Convert.ToDecimal(this.txtCantMaxima.Value),
				(int)this.rbgOrigen.EditValue, IdArancel, Convert.ToDecimal(this.txtPeso.Value), Convert.ToDecimal(this.txtAncho.Value), Convert.ToDecimal(this.txtLargo.Value), Convert.ToDecimal(this.txtAlto.Value),
				this.chkPrecio.Checked, Convert.ToDecimal(this.txtPrecioPI.Value), Convert.ToDecimal(this.txtPrecioIncr1.Value), Convert.ToDecimal(this.txtPrecioPA.Value), 
				Convert.ToDecimal(this.txtPrecioIncr2.Value), Convert.ToDecimal(this.txtPrecioTarjeta.Value), Convert.ToDecimal(this.txtPrecioPVP.Value), Convert.ToDecimal(this.txtPrecioCostoUltimo.Value), 0.00m,
				IdSeccion, IdCentroDeCosto, Convert.ToDecimal(this.txtDepreciacion.Value), Convert.ToDecimal(this.txtResidual.Value), Convert.ToDecimal(this.txtTiempo.Value), dtFechaUltimo.ToString("yyyyMMdd"), 
				this.chkCombo.Checked, this.chkManejaSerial.Checked, this.chkSerialAlterna.Checked, this.chkFaturable.Checked, this.chkAplicDescuento.Checked, this.chkDescontinuado.Checked, 
				this.chkDescCompras.Checked, this.chkDescVentas.Checked, bFechaUltimo, this.chkTransferible.Checked, idPreciador, this.txtCodArtEasyMax.Text.ToString(), this.chkInventario.Checked, this.chkImpuestoIVA.Checked,
				this.chkGarantiaExtendida.Checked, Convert.ToDecimal(this.txtPorcentajeGE.Value), (int)this.txtTiempoGE.Value, this.chkServicioTecnicoDomicilio.Checked, this.txtLink.Text.ToString(),
				this.chkObsequios.Checked,this.chkcatalogo.Checked, 0.00m,
				this.chkConsignacion.Checked, Convert.ToDecimal(this.txtCSCostoEfectivo.Value), Convert.ToDecimal(this.txtCSCostoCredito.Value), this.txtCodigoProveedor.Text.ToString(), Convert.ToDecimal(this.txtPrecioTCsinIntereses.Value),
				this.chkProductoWeb.Checked,this.chkPayJoy.Checked, this.txtCodigoPayJoy.Text.ToString(),
				this.txtOrigenMoto.Text.ToString(),(int)this.txtfechaMoto.Value,Convert.ToDecimal(this.txtTonelaje.Value),(int)this.txtCilindraje.Value,(int)this.txtPasajeros.Value,this.chkMoto.Checked,this.chkMigrar.Checked,this.chkManejaSerie.Checked, this.txtnovasoft.Text.ToString(),Convert.ToDecimal(this.txtCostoPayJoy.Value), this.chkPrecioMovil.Checked, 
				Convert.ToDecimal(this.txtPrecioDescEmpleado.Value), Convert.ToDecimal(this.txtPrecioAutoconsumo.Value), this.chkDepartamento.Checked, this.chkDepartamento1.Checked,this.chkEfectivo.Checked,this.txtDescripcionWeb.Text.ToString(), this.txtTituloWeb.Text.ToString());
			this.txtIdArticulo.Value = Funcion.iEscalarSQL(cdsSeteoF, sSQLGuarda, true);
			#endregion Guarda Registro

			#region Estado Controles
			this.cmbGrupo.Enabled = false;
			this.cmbSubGrupo.Enabled = false;
			this.cmbCategoria.Enabled = false;
			this.cmbMarca.Enabled = false;
			this.txtModelo.Enabled = false;
			this.txtNParte.Enabled = false;
			this.cmbColor.Enabled = false;
			this.cmbPreciador.Enabled = false;
			this.txtCodigo.Enabled = false;
			this.txtCodigoReemplazo.Enabled = false;
			this.txtNombre.Enabled = false;
			this.cmbArticuloSRI.Enabled = false;;
			this.txtIva.Enabled = false;
			this.txtTamaño.Enabled = false;
			this.txtPeriodoGarantia.Enabled = false;
			this.txtVidaUtil.Enabled = false;
			this.txtSaldo.Enabled = false;
			this.txtCantMinima.Enabled = false;
			this.txtCantMaxima.Enabled = false;
			this.rbgOrigen.Enabled = false;
			this.cmbArancel.Enabled = false;
			this.txtPeso.Enabled = false;
			this.txtPrecioTCsinIntereses.Enabled = false;
			this.txtAncho.Enabled = false;
			this.txtLargo.Enabled = false;
			this.txtAlto.Enabled = false;
			this.chkPrecio.Enabled = false;
			this.txtPrecioPI.Enabled = false;
			this.txtPrecioIncr1.Enabled = false;
			this.txtPrecioPA.Enabled = false;
			this.txtPrecioIncr2.Enabled = false;
			this.txtPrecioTarjeta.Enabled = false;
			this.txtPrecioPVP.Enabled = false;
			this.txtPrecioCostoUltimo.Enabled = false;	
			this.txtCostoPayJoy.Enabled = false;
			this.txtPrecioAutoconsumo.Enabled = false;
			this.txtPrecioDescEmpleado.Enabled = false;
			this.cmbSeccion.Enabled = false;
			this.cmbCentroCosto.Enabled = false;
			this.txtDepreciacion.Enabled = false;
			this.txtResidual.Enabled = false;
			this.txtTiempo.Enabled = false;
			this.dtFechaUltimo.Enabled = false;
			this.chkCombo.Enabled = false;
			this.chkObsequios.Enabled = false;
			this.chkManejaSerial.Enabled = false;
			this.chkSerialAlterna.Enabled = false;
			this.chkFaturable.Enabled = false;
			this.chkAplicDescuento.Enabled = false;
			this.chkDescontinuado.Enabled = false;
			this.chkDescCompras.Enabled = false;
			this.chkDescVentas.Enabled = false;
			this.chkTransferible.Enabled = false;
			this.chkInventario.Enabled = false;
			this.chkEfectivo.Enabled = false;
			this.chkImpuestoIVA.Enabled = false;
			this.chkConsignacion.Enabled = false;
			this.txtCSCostoEfectivo.Enabled = false;
			this.txtCSCostoCredito.Enabled = false;
			this.txtCodigoProveedor.Enabled = false;
			this.chkProductoWeb.Enabled= false;
			this.chkPayJoy.Enabled= false;
			this.chkDepartamento.Enabled = false;
			this.chkDepartamento1.Enabled = false;
			this.chkPrecioMovil.Enabled= false;
			this.chkMigrar.Enabled= false;
			this.txtCodigoPayJoy.Enabled= false;
			this.txtDescripcionWeb.Enabled = false;
			this.txtTituloWeb.Enabled = false;

			//this.chkMoto.Checked = false;
			this.txtOrigenMoto.Enabled = false;
			this.txtfechaMoto.Enabled = false;
			this.txtTonelaje.Enabled = false;
			this.txtCilindraje.Enabled = false;
			this.txtPasajeros.Enabled = false;
			this.chkMoto.Enabled = false;
			this.chkManejaSerie.Enabled = false;
			this.txtnovasoft.Enabled = false;

			this.chkGarantiaExtendida.Enabled = false;
			this.chkcatalogo.Enabled = false;
			this.chkServicioTecnicoDomicilio.Enabled = false;
			this.txtTiempoGE.Enabled = false;
			this.txtPorcentajeGE.Enabled = false;
			this.txtCodArtEasyMax.Enabled = false;
			this.txtLink.Enabled = false;

			this.btnNuevo.Enabled = true;
			this.btnConsultar.Enabled = true;
			if (miAcceso.Editar) this.btnEditar.Enabled = true;
			if (miAcceso.Imprimir) this.btnImprimir.Enabled = true;
			this.btnGuardar.Enabled = false;
			this.btnCaracteristicas.Enabled = true;
			#endregion Estado Controles					

			bNuevo = false;
			bEdicion = false;

			idSubgrupo = 0;

			if (MenuLatinium.stDirFacturacion == "POINT")
			{
				frmAtributosArticulos ArtAtri = new frmAtributosArticulos((int)this.txtIdArticulo.Value); //((int)this.cmbSubGrupo.Value, (int)this.txtIdArticulo.Value);
				ArtAtri.ShowDialog();
			}
		}

		private void cmbMarca_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}

		private void groupBox1_Enter(object sender, System.EventArgs e)
		{
		
		}

		private void btColor_Click(object sender, System.EventArgs e)
		{
			using (ArticuloColor Color = new ArticuloColor())
			{
				if (DialogResult.OK == Color.ShowDialog())
				{
					this.cmbColor.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Select idColor, Color From ArticuloColor Order By Color");
					this.cmbColor.Focus();
				}
			}
		}

		private bool bloquearEventos = false;
		private void btGrupo_Click(object sender, System.EventArgs e)
		{
			using (frmGruposArticulos miGrupo = new frmGruposArticulos())
			{
				if (DialogResult.OK == miGrupo.ShowDialog())
				{
					bloquearEventos = true; // Bloquear eventos mientras llenamos combos

					this.cmbGrupo.DataSource = Funcion.dvProcedimiento(cdsSeteoF,
						"Select idGrupoArticulo, Grupo From ArticuloGrupo Where idTipoGrupo In (1, 5, 6, 13, 16)");

					if (this.cmbGrupo.Text.ToString() != "")
					{
						string sSQL = string.Format("Exec ListaSubgrupoArticulos {0}", (int)this.cmbGrupo.Value);
						this.cmbSubGrupo.DataSource = Funcion.dvProcedimiento(cdsSeteoF, sSQL);

						if (cmbSubGrupo.ActiveRow != null)
						{
							sSQL = string.Format("Exec ListaCategoriasMarcas {0}, 1", (int)this.cmbSubGrupo.Value);
							this.cmbCategoria.DataSource = Funcion.dvProcedimiento(cdsSeteoF, sSQL);

							sSQL = string.Format("Exec ListaCategoriasMarcas {0}, 2", (int)this.cmbSubGrupo.Value);
							this.cmbMarca.DataSource = Funcion.dvProcedimiento(cdsSeteoF, sSQL);
						}
					}

					bloquearEventos = false; // Habilitar eventos nuevamente
				}
			}
		}

		private void btArancel_Click(object sender, System.EventArgs e)
		{
			using (ArtArancel Arancel = new ArtArancel())
			{
				if (DialogResult.OK == Arancel.ShowDialog())
				{
					this.cmbArancel.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Select idArancel, Codigo, Porcentaje, Arancel From ArticuloArancel Order By Arancel");
					this.cmbArancel.Focus();
				}
			}
		}

		private void btSeccion_Click(object sender, System.EventArgs e)
		{
			using(ArtSeccion Seccion = new ArtSeccion())
			{
				if (DialogResult.OK == Seccion.ShowDialog())
				{
					this.cmbSeccion.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Select idSeccion, Seccion From ArticuloSeccion Order By Seccion");
					this.cmbSeccion.Focus();
				}
			}
		}

		private void cmbSeccion_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}

		private void chkSerialAlterna_CheckedChanged(object sender, System.EventArgs e)
		{
		
		}

		private void chkDescCompras_CheckedChanged(object sender, System.EventArgs e)
		{
		
		}

		private void cmbCategoria_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}

		private void txtPrecioPI_ValueChanged(object sender, System.EventArgs e)
		{
		
		}

		private void txtTamaño_Validated(object sender, System.EventArgs e)
		{
			if (this.txtTamaño.Value == System.DBNull.Value) this.txtTamaño.Value = 0;
		}

		private void cmbPreciador_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter)
			{
				if (this.txtCodigo.Enabled) this.txtCodigo.Focus();
				else if (this.txtCodigoReemplazo.Enabled) this.txtCodigoReemplazo.Focus();
				else this.txtNombre.Focus();
			}
		}

		private void cmbMarca_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.txtModelo.Focus();
		}

		private void cmbMarca_Enter(object sender, System.EventArgs e)
		{
			this.cmbMarca.ToggleDropdown();
		}

		private void frmArticulos_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (!bEdicion)
			{
				if (e.KeyCode == Keys.N) if (miAcceso.Nuevo)	if (this.btnNuevo.Enabled) this.btnNuevo_Click(sender, e);
				if (e.KeyCode == Keys.B) if (miAcceso.Buscar) if (this.btnConsultar.Enabled) this.btnConsultar_Click(sender, e);
				if (e.KeyCode == Keys.E) if (miAcceso.Editar) if (this.btnEditar.Enabled) this.btnEditar_Click(sender, e);
				if (e.KeyCode == Keys.I) this.btnImprimir_Click(sender, e);																 
				if (e.KeyCode == Keys.Escape) this.Close();
				if (e.KeyCode == Keys.F5) this.Consulta((int)this.txtIdArticulo.Value);
			}
			else
			{
				if (e.KeyCode == Keys.Escape) this.btnCancelar_Click(sender, e);
				if (e.KeyCode == Keys.F12) this.btnGuardar_Click(sender, e);				
			}
		}

		private void frmArticulos_Closing(object sender, System.ComponentModel.CancelEventArgs e)
		{
			if (bEdicion)
			{
				MessageBox.Show("No puede salir mientras edita un Articulo", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				e.Cancel = true;
				return;
			}
		}

		private void txtCantMinima_ValueChanged(object sender, System.EventArgs e)
		{
		
		}

		private void txtPrecioTarjeta_Validated(object sender, System.EventArgs e)
		{
			if (bEdicion)
			{
				if (this.chkPrecio.Checked)
				{	
					int idSubGrupo = 0;
					if (this.cmbSubGrupo.ActiveRow != null) idSubGrupo = (int)this.cmbSubGrupo.Value;

					decimal dIncrPVP = Funcion.decEscalarSQL(cdsSeteoF, string.Format("Select IncrPVP From ArticuloSubgrupo Where idSubGrupo = {0}", idSubGrupo, true));
					decimal dTCSI = Funcion.decEscalarSQL(cdsSeteoF, string.Format("Select PorcCredito From ArticuloSubgrupo Where idSubGrupo = {0}", idSubGrupo, true));
										
					this.txtPrecioPVP.Value = Math.Round(Convert.ToDecimal(this.txtPrecioTarjeta.Value) * dIncrPVP, 2);
					this.txtPrecioTCsinIntereses.Value = Math.Round(Convert.ToDecimal(this.txtPrecioTarjeta.Value) * dTCSI, 2);

				}
			}
		}

		private void chkGarantiaExtendida_CheckedChanged(object sender, System.EventArgs e)
		{
			if (bEdicion)
			{
				this.chkServicioTecnicoDomicilio.Enabled = this.chkGarantiaExtendida.Checked;
				this.txtPorcentajeGE.Enabled = this.chkGarantiaExtendida.Checked;
				this.txtTiempoGE.Enabled = this.chkGarantiaExtendida.Checked;
				this.txtPorcentajeGE.Value = 0.00m;
				this.txtTiempoGE.Value = 0;
				this.chkServicioTecnicoDomicilio.Checked = false;
			}
		}

		private void btnCaracteristicas_Click(object sender, System.EventArgs e)
		{
			frmArticulos_Caracteristicas A_C = new frmArticulos_Caracteristicas((int)this.cmbSubGrupo.Value, (int)this.txtIdArticulo.Value);
			A_C.ShowDialog();
		}

		private void txtLink_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.txtNombre.Focus();
		}

		private void groupBox3_Enter(object sender, System.EventArgs e)
		{
		
		}

		private void chkConsignacion_CheckedChanged(object sender, System.EventArgs e)
		{
			if (bEdicion)
			{
				this.txtCSCostoEfectivo.Value = 0.00m;
				this.txtCSCostoCredito.Value = 0.00m;
				this.txtCodigoProveedor.Text = "";

				this.txtCSCostoEfectivo.Enabled = false;
				this.txtCSCostoCredito.Enabled = false;
				this.txtCodigoProveedor.Enabled = false;

				if (this.chkConsignacion.Checked)
				{
					this.txtCSCostoEfectivo.Enabled = true;
					this.txtCSCostoCredito.Enabled = true;
					this.txtCodigoProveedor.Enabled = true;

					this.txtCSCostoEfectivo.Focus();
				}
			}
		}

		private void txtCSCostoEfectivo_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.txtCSCostoCredito.Focus();
		}

		private void txtCSCostoCredito_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.txtCodigoProveedor.Focus();
		}

		private void txtCSCostoEfectivo_Click(object sender, System.EventArgs e)
		{
			this.txtCSCostoEfectivo.SelectAll();
		}

		private void txtCSCostoEfectivo_Enter(object sender, System.EventArgs e)
		{
			this.txtCSCostoEfectivo.SelectAll();
		}

		private void txtCSCostoCredito_Click(object sender, System.EventArgs e)
		{
			this.txtCSCostoCredito.SelectAll();
		}

		private void txtCSCostoCredito_Enter(object sender, System.EventArgs e)
		{
			this.txtCSCostoCredito.SelectAll();
		}

		private void cmbMarca_ValueChanged(object sender, System.EventArgs e)
		{
			if (bloquearEventos) return; // Bloquear eventos mientras cargas combos

			if (bEdicion)
				if (this.cmbMarca.ActiveRow != null)
					GeneraDescripcion();
		}

		private void txtPrecioTarjeta_ValueChanged(object sender, System.EventArgs e)
		{
		
		}

		private void ultraNumericEditor1_ValueChanged(object sender, System.EventArgs e)
		{
		
		}

		private void txtPrecioTCsinIntereses_Enter(object sender, System.EventArgs e)
		{
			this.txtPrecioTCsinIntereses.SelectAll();
		}

		private void txtPrecioTCsinIntereses_Click(object sender, System.EventArgs e)
		{
			this.txtPrecioTCsinIntereses.SelectAll();
		}

		private void chkProductoWeb_CheckedChanged(object sender, System.EventArgs e)
		{
		
		}

		private void chkObsequios_CheckedChanged(object sender, System.EventArgs e)
		{
		
		}

		private void label22_Click(object sender, System.EventArgs e)
		{
		
		}

		private void chkPayJoy_CheckedChanged(object sender, System.EventArgs e)
		{
			if (!bEdicion) return; 

			this.txtCodigoPayJoy.Text = "";
			this.txtCodigoPayJoy.Enabled = false;

			if (this.chkPayJoy.Checked)
			{
				this.txtCodigoPayJoy.Enabled = true;
				this.txtCodigoPayJoy.Focus();


				decimal precioCostoUltimo = 0;
				if (this.txtPrecioCostoUltimo.Value != null)
					precioCostoUltimo = Convert.ToDecimal(this.txtPrecioCostoUltimo.Value);

				// Por defecto, incremento 30%
				decimal incrPrecioC25 = 30M;

				// Solo calcular si el campo está vacío o es 0
				if (this.txtCostoPayJoy.Value == null || Convert.ToDecimal(this.txtCostoPayJoy.Value) == 0)
				{
					// Calcula el costo final usando la fórmula correcta
					decimal costoPayjoy = precioCostoUltimo / (1 - (incrPrecioC25 / 100M));

					// Redondear a 2 decimales
					costoPayjoy = Math.Round(costoPayjoy, 2);

					// Asignar al control
					this.txtCostoPayJoy.Value = costoPayjoy;
				}

				this.txtCostoPayJoy.Enabled = this.chkPrecio.Checked;
			}
			else
			{
	
				this.txtCostoPayJoy.Enabled = false;
			}
		}


		private void label25_Click(object sender, System.EventArgs e)
		{
		
		}

		private void txtCodigo_ValueChanged(object sender, System.EventArgs e)
		{
		
		}

		private void txtafechaMoto_ValueChanged(object sender, System.EventArgs e)
		{
		
		}

		private void chkMoto_CheckedChanged(object sender, System.EventArgs e)
		{
			if (bEdicion)
			{		
				this.txtOrigenMoto.Text = "";
				this.txtfechaMoto.Value = 0;
				this.txtTonelaje.Value = 0.00m;
				this.txtCilindraje.Value = 0;
				this.txtPasajeros.Value = 2;


				this.txtOrigenMoto.Enabled = false;
				this.txtfechaMoto.Enabled = false;
				this.txtTonelaje.Enabled = false;
				this.txtCilindraje.Enabled = false;
				this.txtPasajeros.Enabled = false;

				if (this.chkMoto.Checked)
				{
					this.txtOrigenMoto.Enabled = true;
					this.txtfechaMoto.Enabled = true;
					this.txtTonelaje.Enabled = true;
					this.txtCilindraje.Enabled = true;
					//this.txtPasajeros.Enabled = true;
					this.txtOrigenMoto.Focus();
				}
			}
		}

		private void button1_Click(object sender, System.EventArgs e)
		{
			if (MenuLatinium.stDirFacturacion == "POINT")
			{
				frmAtributosArticulos ArtAtri = new frmAtributosArticulos((int)this.txtIdArticulo.Value); //((int)this.cmbSubGrupo.Value, (int)this.txtIdArticulo.Value);
				ArtAtri.ShowDialog();
			}
		}

		private void chkMigrar_CheckedChanged(object sender, System.EventArgs e)
		{
			if (bEdicion)
			{
				this.txtnovasoft.Text ="";
				this.txtnovasoft.Enabled = false;
				if (this.chkMigrar.Checked)
				{
					this.txtnovasoft.Enabled = true;
					this.txtnovasoft.Focus();
					this.chkManejaSerie.Enabled=true;
				}
				else 
				{
					this.txtnovasoft.Text ="";
					this.chkManejaSerie.Checked=false;
				}
			}
		}

		

		private void label33_Click(object sender, System.EventArgs e)
		{
		
		}

		private void txtPrecioCostoUltimo_ValueChanged(object sender, System.EventArgs e)
		{
		
		}

		private void groupBox2_Enter(object sender, System.EventArgs e)
		{
		
		}

		private void txtCodigoPayJoy_KeyPress(object sender, KeyPressEventArgs e)
		{

			if (e.KeyChar == 22)
			{
				e.Handled = true;
				return;
			}

			//			if (!char.IsLetterOrDigit(e.KeyChar) && e.KeyChar != (char)8)
			//			{
			//				e.Handled = true;
			//			}
		}

		

	}
}

