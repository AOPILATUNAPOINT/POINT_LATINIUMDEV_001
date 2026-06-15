using System;
using System.Globalization;
using System.Data;
using System.Data.SqlClient;
using C1.Data;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using Infragistics.Win.UltraWinGrid;
using System.Web.Mail;
using CrystalDecisions.Shared;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de Compra2.
	/// </summary>
	public class Compra : System.Windows.Forms.Form
	{
		private C1.Data.C1DataSet cdsCompra;
		private C1.Data.C1DataSet cdsCliente;
		private C1.Data.C1DataView cdvCliente;
		private C1.Data.C1DataSet cdsArticulo;
		private C1.Data.C1DataView cdvArtCodigo;
		private C1.Data.C1DataView cdvArticulo;
		private C1.Data.C1DataView cdvComprobante;
		private System.Windows.Forms.ErrorProvider errorProvider;
		private System.Windows.Forms.ImageList imageList;
		private System.Windows.Forms.Label label6;
		private System.Windows.Forms.Label label7;
		private System.Windows.Forms.Label label8;
		private C1.Data.C1DataSet cdsSeguridad;
		private System.Windows.Forms.Label lblAnulado;
		private C1.Data.C1DataSet cdsProyecto;
		private C1.Data.C1DataView cdvProyecto;
		private C1.Data.C1DataView cdvSubProyecto;
		private C1.Data.C1DataView cdvFormaPago;
		private Infragistics.Win.UltraWinGrid.UltraGrid ultraGrid1;
		private C1.Data.C1DataSet cdsSeteoF;
		private C1.Data.C1DataView cdvNumero;
		private C1.Data.C1DataView cdvCompra;
		private C1.Data.C1DataSet cdsCompraTabla;
		public DevExpress.XtraEditors.SimpleButton btTotal;
		private C1.Data.C1DataSet cdsAsiento;
		private C1.Data.C1DataSet cdsRetencion;
		private C1.Data.C1DataSet cdsAsignaCuenta;
		private Infragistics.Win.UltraWinGrid.UltraDropDown CmbDescCod;
		private Infragistics.Win.UltraWinGrid.UltraDropDown CmbCodDesc;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbCliente;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtNumero;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtFecha;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor spnDias;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtVence;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor spnBodega;
		private Infragistics.Win.Misc.UltraButton btCliente;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbComprobante1;
		private C1.Data.C1DataSet cdsPersonal;
		private C1.Data.C1DataView cdvPersonal;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbTipoCompra;
		private Infragistics.Win.Misc.UltraButton btSaldo1;
		private System.Windows.Forms.ToolTip toolTip;
		private Infragistics.Win.Misc.UltraButton btTransforma;
		private Infragistics.Win.Misc.UltraButton btNotaEntrega;
		private Infragistics.Win.Misc.UltraButton btBuscaArticulo;
		private Infragistics.Win.Misc.UltraButton btRetencion;
		private Infragistics.Win.Misc.UltraButton btCreaAsiento;
		private Infragistics.Win.Misc.UltraButton btGuiaRemision;
		private Infragistics.Win.Misc.UltraButton btActualizaPrecio;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtIdAsiento;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtIdCompra;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtBorrar;
		private BarraDatoSQL.BarraDatoSQL barraDato1;
		private Infragistics.Win.UltraWinTabControl.UltraTabSharedControlsPage ultraTabSharedControlsPage1;
		private Infragistics.Win.UltraWinTabControl.UltraTabPageControl ultraTabPageControl1;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtComprobante;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbSubProyecto;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbProyecto;
		private System.Windows.Forms.Label label15;
		private System.Windows.Forms.Label label5;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtNotas;
		private Infragistics.Win.UltraWinTabControl.UltraTabPageControl ultraTabPageControl2;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbFormaPago;
		private Infragistics.Win.UltraWinEditors.UltraOptionSet rdgContadoCred;
		private Infragistics.Win.Misc.UltraButton btPago;
		private Infragistics.Win.Misc.UltraButton cmdFormaPago;
		private Infragistics.Win.UltraWinTabControl.UltraTabPageControl ultraTabPageControl3;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtDescPorcIva;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtDescPorcIva0;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtDescLinea0;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtDescLinea;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtDescIva;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtDescIva0;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtFlete;
		private Infragistics.Win.Misc.UltraButton btRefresaI;
		private Infragistics.Win.Misc.UltraButton btRefrescaD;
		private System.Windows.Forms.Label label17;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkRevisado;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkEntregado;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtFechaRevision;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtFechaEntrega;
		private System.Windows.Forms.Label label20;
		private System.Windows.Forms.Label label19;
		private Infragistics.Win.UltraWinTabControl.UltraTabPageControl ultraTabPageControl5;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtInteresPagar;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtEmpezarEn;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtPagosCada;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtPorcInteres;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtTotalCuotas;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtValorCuotas;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtAnticipo;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtNumCuotas;
		private Infragistics.Win.Misc.UltraButton btTotalCuotas;
		private System.Windows.Forms.Label label25;
		private System.Windows.Forms.Label label26;
		private System.Windows.Forms.Label label27;
		private System.Windows.Forms.Label label28;
		private System.Windows.Forms.Label label23;
		private System.Windows.Forms.Label label22;
		private System.Windows.Forms.Label label21;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtDescIva0Total;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtDescIvaTotal;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtIva0;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtIceTotal;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtTotal;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtIvaTotal;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtIva;
		private System.Windows.Forms.Label label18;
		private System.Windows.Forms.Label label24;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtCheque;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtNumCuenta;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbCredTribut;
		private System.Windows.Forms.Label label29;
		private System.Windows.Forms.Label label30;
		private Infragistics.Win.UltraWinEditors.UltraOptionSet OptBienServicio;
		private Infragistics.Win.UltraWinTabControl.UltraTabPageControl TabPageCOA;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkDevolucion;
		private System.Windows.Forms.Label label31;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtAutFactura;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtAutImprenta;
		private System.Windows.Forms.Label label32;
		private System.Windows.Forms.Label label33;
		private System.Windows.Forms.Label label34;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor spnPrioridad;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkAutPrincipal;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkAutPresupuestaria;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkConvenioInt;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtRefrendo;
		private System.Windows.Forms.Label label35;
		private Infragistics.Win.Misc.UltraButton cmdVuelto;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor spnVuelto;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtDepartamento;
		private System.Windows.Forms.Label label37;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtOtro;
		private System.Windows.Forms.Label label38;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbRecibe;
		private Infragistics.Win.Misc.UltraButton btFlete;
		private Infragistics.Win.Misc.UltraButton btNuevo;
		private C1.Data.C1DataSet cdsSeteo;
		private Infragistics.Win.Misc.UltraButton btCambiarPrecio;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbPrecio;
		private C1.Data.C1DataSet cdsClienteTabla;
		private Infragistics.Win.UltraWinGrid.UltraDropDown cmbSubProyecto1;
		private Infragistics.Win.UltraWinGrid.UltraDropDown cmbProyecto1;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbVendedor;
		private System.Windows.Forms.Label lblCliente;
		private System.Windows.Forms.HelpProvider helpProvider1;
		private Infragistics.Win.UltraWinTabControl.UltraTabPageControl ultraTabPageControl6;
		private Infragistics.Win.UltraWinGrid.UltraGrid ugrBusca;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtBusca;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource udsBusca;
		private Infragistics.Win.Misc.UltraButton btBuscar;
		private Infragistics.Win.Misc.UltraButton btIngresar;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkCodBarras;
		private Infragistics.Win.UltraWinEditors.UltraDateTimeEditor cmbCaducidad;
		private System.Windows.Forms.Label label11;
		private System.Windows.Forms.Label label39;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtUsuario;
		private Infragistics.Win.UltraWinGrid.UltraDropDown cmbUnidad;
		private C1.Data.C1DataSet cdsUnidad;
		private System.Windows.Forms.Label label40;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbIce;
		private Infragistics.Win.Misc.UltraButton btComprobante;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtComprob1;
		private Infragistics.Win.Misc.UltraButton btComprob1;
		private Infragistics.Win.UltraWinTabControl.UltraTabPageControl ultraTabPageControl4;
		private System.Windows.Forms.Label label36;
		private System.Windows.Forms.Label label41;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtBulto;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtPeso;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor ultraTextEditor1;
		private System.Windows.Forms.Label label42;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor ultraTextEditor3;
		private System.Windows.Forms.Label label43;
		private Infragistics.Win.UltraWinTabControl.UltraTabPageControl pgEntrega;
		private Infragistics.Win.Misc.UltraButton btImportacion;
		private Infragistics.Win.Misc.UltraButton btDescuento;
		private C1.Data.C1DataSet cdsTipoImportacion;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtNombre;
		private Infragistics.Win.UltraWinGrid.UltraDropDown cmbImportacion;
		private System.Windows.Forms.Label label44;
		private Infragistics.Win.Misc.UltraButton btBarras;
		private Infragistics.Win.Misc.UltraButton btAuditoria;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtCIF;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtSerieFact;
		private Infragistics.Win.UltraWinTabControl.UltraTabPageControl tbpAutorizacion;
		private Infragistics.Win.UltraWinTabControl.UltraTabControl utcTransaccion;
		private Infragistics.Win.Misc.UltraButton btCentro;
		private Infragistics.Win.Misc.UltraButton btSubProyecto;
		private C1.Data.C1DataView cdvEntrega;
		private System.Windows.Forms.Label label13;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor ultraTextEditor2;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor ultraTextEditor4;
		private System.Windows.Forms.Label label14;
		private System.Windows.Forms.Label label45;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor ultraTextEditor5;
		private System.Windows.Forms.Label label46;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor ultraTextEditor6;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtServicio;
		private System.Windows.Forms.Label lblIva0;
		private System.Windows.Forms.Label lblIvaSubt;
		private System.Windows.Forms.Label lblIva;
		private System.Windows.Forms.Label lblDescuento;
		private System.Windows.Forms.Label lblSubTotal;
		private System.Windows.Forms.Label lblIce;
		private System.Windows.Forms.Label lblServicio;
		private System.Windows.Forms.Label lblPropina;
		private Infragistics.Win.Misc.UltraButton btComision;
		private System.Windows.Forms.Label lblBodega;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbBodega;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource udsBodegas;
		private Infragistics.Win.Misc.UltraButton btCorresponsal;
		private System.Windows.Forms.Label label1;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtComision;
		private System.Windows.Forms.Label lblTiempo;
		private Infragistics.Win.UltraWinGrid.UltraDropDown cmbBodegaLinea;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtPropina;
		private Infragistics.Win.UltraWinGrid.UltraDropDown cmbUnidadB;
		private C1.Data.C1DataSet cdsSeteoC;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor ultraNumericEditor1;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor ultraNumericEditor2;
		private System.Windows.Forms.Label label2;
		private System.Windows.Forms.Label label3;
		private System.ComponentModel.IContainer components;

		public Compra()
		{
			InitializeComponent();
		}

		int IdTipoFactura = 1;
		public Compra(int idTransaccion)
		{
			InitializeComponent();
			IdTipoFactura = idTransaccion;
			drSeteoF = this.cdsSeteoF.TableViews["SeteoF"].Rows[0];
			//			if (drSeteoF["VerDesplegable"] != System.DBNull.Value &&
			//				!(bool) drSeteoF["VerDesplegable"])
			//				cdsArticulo.FillOnRequest = false;
		}

		int IdBusqueda = 0;
		public Compra(int idTransaccion, int idBusca)
		{
			InitializeComponent();
			IdTipoFactura = idTransaccion;
			IdBusqueda = idBusca;
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
			this.components = new System.ComponentModel.Container();
			Infragistics.Win.Appearance appearance1 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Personal", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPersonal");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idAsiDistGrupo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCentroCosto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idDepartamento");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCargo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idRolSucursal");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idRolGrupo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Jornal");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Apellido");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn11 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn12 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Jornada");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn13 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Direccion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn14 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Telefono");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn15 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Celular");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn16 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cedula");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn17 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Afiliacion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn18 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Reserva");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn19 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Renta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn20 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Iess");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn21 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Sueldo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn22 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Comision");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn23 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Tipo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn24 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaIngreso");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn25 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaNacimiento");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn26 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaSalida");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn27 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Notas");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn28 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CargasFamiliares");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn29 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idProvincia");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn30 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCiudad");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn31 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Numero");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn32 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("SistemaSalario");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn33 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cuenta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn34 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CodigoSectorial");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn35 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Condicion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn36 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaNovedad");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn37 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Acceso");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn38 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Vendedor");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn39 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Borrar");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn40 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cobrador");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn41 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Sexo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn42 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Estado_Civil");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn43 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Banco");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn44 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CuentaCorr");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn45 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idTipoCuenta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn46 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CedulaMilitar");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn47 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idFormaPago");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn48 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Quincena");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn49 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CodIngreso");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn50 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CodTelefono");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn51 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("TarjEdificio");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn52 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("TarjetaOficina");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn53 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("email");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn54 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idProyecto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn55 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCuenta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn56 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bono");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn57 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bono1");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn58 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Desc1");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn59 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Desc2");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn60 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("TieneQuincena");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn61 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Fotos");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn62 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Convenio");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn63 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idSubProyecto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn64 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Sindicato");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn65 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Asociacion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn66 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idRolHorario");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn67 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaContrato");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn68 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Educacion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn69 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("NumeroDias");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn70 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idRelacionT");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn71 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("gSalud");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn72 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("gEducacion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn73 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("gVivienda");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn74 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("gVestimenta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn75 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("gAlimentacion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn76 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("gTotal");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn77 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("BaseComision");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn78 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("PorcentajeComision");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn79 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("gDesahucio");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn80 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("gDiscapacidad");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn81 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("IngGravOE");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn82 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("DeduccionOE");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn83 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("OtrasRebOE");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn84 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ImpRentaAsumido");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn85 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("gRebajaIIIEdad");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn86 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("RentaRetenidaOE");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn87 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("RentaRetOE");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn88 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("TipoSangre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn89 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("RegEntrada");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn90 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("RegSalida");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn91 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idTipoContrato");
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("SubProyecto", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn92 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idSubProyecto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn93 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idProyecto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn94 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn95 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Lugar");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn96 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("RespNombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn97 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Notas");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn98 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CodSubproyecto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn99 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Visible");
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand3 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Proyecto", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn100 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idProyecto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn101 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPadre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn102 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn103 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Lugar");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn104 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Responsable");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn105 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Notas");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn106 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CentroCosto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn107 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaIngreso");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn108 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Visible");
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand4 = new Infragistics.Win.UltraWinGrid.UltraGridBand("ClientePrecio", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn109 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idTipoPrecio");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn110 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Precio");
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance11 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand5 = new Infragistics.Win.UltraWinGrid.UltraGridBand("CompraFormaPago", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn111 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idFormaPago");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn112 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPrecio");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn113 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FormaPago");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn114 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCuenta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn115 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Borrar");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn116 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Descuento");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn117 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Ruc");
			Infragistics.Win.Appearance appearance12 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance13 = new Infragistics.Win.Appearance();
			Infragistics.Win.ValueListItem valueListItem1 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem2 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.Appearance appearance14 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance15 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance16 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance17 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance18 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance19 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance20 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand6 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Cliente", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn118 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn119 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCliente");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn120 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn121 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Proveedor");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn122 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Ruc");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn123 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Ambos");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn124 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("DesconProv");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn125 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("DesconCliente");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn126 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Entrega");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn127 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idBanco");
			Infragistics.Win.Appearance appearance21 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance22 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance23 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance24 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance25 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton1 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.Appearance appearance26 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton2 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.Appearance appearance27 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance28 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance29 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance30 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance31 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance32 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance33 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance34 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance35 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance36 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance37 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand7 = new Infragistics.Win.UltraWinGrid.UltraGridBand("PorcIce", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn128 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPorcIce");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn129 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn130 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("porcentaje");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn131 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("concepto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn132 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CodSri");
			Infragistics.Win.Appearance appearance38 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance39 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance40 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance41 = new Infragistics.Win.Appearance();
			Infragistics.Win.ValueListItem valueListItem3 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem4 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.Appearance appearance42 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance43 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand8 = new Infragistics.Win.UltraWinGrid.UltraGridBand("CompraCredTribut", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn133 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCredTributario");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn134 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn135 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CredTributario");
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
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand9 = new Infragistics.Win.UltraWinGrid.UltraGridBand("ImportacionTipo", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn136 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idImportacionTipo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn137 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn138 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Grupo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn139 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Orden");
			Infragistics.Win.Appearance appearance57 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance58 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance59 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance60 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idArticulo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn3 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Codigo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn4 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Alterno");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn5 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Precio1");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn6 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Precio2");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn7 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Precio3");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn8 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Precio4");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn9 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Precio5");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn10 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Precio6");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn11 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Ubicacion");
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(Compra));
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand10 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Articulo", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn140 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Descontinuado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn141 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idArticulo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn142 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("dCompra");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn143 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("dVenta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn144 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idTipoGrupo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn145 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Articulo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn146 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn147 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idGrupoArticulo");
			Infragistics.Win.Appearance appearance61 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand11 = new Infragistics.Win.UltraWinGrid.UltraGridBand("_DetCompra", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn148 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idDetCompra");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn149 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCompra");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn150 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idArticulo", -1, "CmbDescCod");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn151 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idUnidad", -1, "cmbUnidad");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn152 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idSucursal");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn153 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CostoPromedio");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn154 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cantidad");
			Infragistics.Win.Appearance appearance62 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn155 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Precio");
			Infragistics.Win.Appearance appearance63 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn156 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Impuesto");
			Infragistics.Win.Appearance appearance64 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn157 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Notas");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn158 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("RefCodigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn159 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("RefNumero");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn160 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("DescuentoArt");
			Infragistics.Win.Appearance appearance65 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn161 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CantidadPeso");
			Infragistics.Win.Appearance appearance66 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn162 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Flete");
			Infragistics.Win.Appearance appearance67 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn163 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bodega");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn164 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("DescuentoPorc");
			Infragistics.Win.Appearance appearance68 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn165 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Pedido");
			Infragistics.Win.Appearance appearance69 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn166 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Signo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn167 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Borrar");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn168 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Saldo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn169 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaIngreso");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn170 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CentroCosto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn171 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Vencimiento");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn172 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idProyecto", -1, "cmbProyecto1");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn173 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idSubProyecto", -1, "cmbSubProyecto1");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn174 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Unidades");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn175 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn176 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CostoMedio");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn177 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("SubTotal");
			Infragistics.Win.Appearance appearance70 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn178 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idImportacion", -1, "cmbImportacion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn179 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ivaTelco");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn180 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Articulo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn181 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CentroCostoBR");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn182 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ProyectoBR");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn183 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CodArt", 0, "CmbCodDesc");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn184 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("unidadBotrosa", 1, "cmbUnidadB");
			Infragistics.Win.Appearance appearance71 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance72 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance73 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance74 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance75 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand12 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Articulo", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn185 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Descontinuado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn186 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idArticulo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn187 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("dCompra");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn188 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("dVenta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn189 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idTipoGrupo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn190 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Articulo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn191 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn192 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idGrupoArticulo");
			Infragistics.Win.Appearance appearance76 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand13 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Cliente", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn193 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn194 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCliente");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn195 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn196 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Proveedor");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn197 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Ruc");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn198 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Ambos");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn199 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("DesconProv");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn200 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("DesconCliente");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn201 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Entrega");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn202 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idBanco");
			Infragistics.Win.Appearance appearance77 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance78 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton3 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.Appearance appearance79 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance80 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton4 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.Appearance appearance81 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance82 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand14 = new Infragistics.Win.UltraWinGrid.UltraGridBand("CompraComprobante", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn203 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idComprobante");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn204 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn205 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Comprobante");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn206 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCompraNumero");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn207 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Reporte");
			Infragistics.Win.Appearance appearance83 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand15 = new Infragistics.Win.UltraWinGrid.UltraGridBand("CompraNumero", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn208 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idTipoFactura");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn209 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn210 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Numero");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn211 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn212 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Automatico");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn213 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Abrev");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn214 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Name");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn215 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Signo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn216 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Grupo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn217 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CobroPago");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn218 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Transforma");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn219 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Copias");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn220 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idTipoAsientoContado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn221 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idTipoAsientoCredito");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn222 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Proyecto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn223 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Subproyecto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn224 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bodega");
			Infragistics.Win.Appearance appearance84 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance85 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance86 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance87 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance88 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance89 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance90 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance91 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance92 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinTabControl.UltraTab ultraTab1 = new Infragistics.Win.UltraWinTabControl.UltraTab();
			Infragistics.Win.UltraWinTabControl.UltraTab ultraTab2 = new Infragistics.Win.UltraWinTabControl.UltraTab();
			Infragistics.Win.UltraWinTabControl.UltraTab ultraTab3 = new Infragistics.Win.UltraWinTabControl.UltraTab();
			Infragistics.Win.UltraWinTabControl.UltraTab ultraTab4 = new Infragistics.Win.UltraWinTabControl.UltraTab();
			Infragistics.Win.UltraWinTabControl.UltraTab ultraTab5 = new Infragistics.Win.UltraWinTabControl.UltraTab();
			Infragistics.Win.UltraWinTabControl.UltraTab ultraTab6 = new Infragistics.Win.UltraWinTabControl.UltraTab();
			Infragistics.Win.UltraWinTabControl.UltraTab ultraTab7 = new Infragistics.Win.UltraWinTabControl.UltraTab();
			Infragistics.Win.UltraWinTabControl.UltraTab ultraTab8 = new Infragistics.Win.UltraWinTabControl.UltraTab();
			Infragistics.Win.UltraWinTabControl.UltraTab ultraTab9 = new Infragistics.Win.UltraWinTabControl.UltraTab();
			Infragistics.Win.Appearance appearance93 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance94 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance95 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance96 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance97 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance98 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance99 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand16 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Proyecto", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn225 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idProyecto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn226 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPadre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn227 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn228 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Lugar");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn229 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Responsable");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn230 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Notas");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn231 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CentroCosto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn232 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaIngreso");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn233 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Visible");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand17 = new Infragistics.Win.UltraWinGrid.UltraGridBand("SubProyecto", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn234 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idSubProyecto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn235 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idProyecto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn236 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn237 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Lugar");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn238 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("RespNombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn239 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Notas");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn240 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CodSubproyecto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn241 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Visible");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand18 = new Infragistics.Win.UltraWinGrid.UltraGridBand("ArticuloUnidad", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn242 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idUnidad");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn243 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Borrar");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn244 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn245 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cantidad");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn246 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Peso");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn247 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Unidades");
			Infragistics.Win.Appearance appearance100 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance101 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance102 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance103 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand19 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn248 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn249 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bodega");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn12 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn13 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Bodega");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand20 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn250 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn251 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bodega");
			Infragistics.Win.Appearance appearance104 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand21 = new Infragistics.Win.UltraWinGrid.UltraGridBand("ArticuloUnidad", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn252 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idUnidad");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn253 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Borrar");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn254 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn255 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cantidad");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn256 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Peso");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn257 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Unidades");
			this.ultraTabPageControl1 = new Infragistics.Win.UltraWinTabControl.UltraTabPageControl();
			this.btSubProyecto = new Infragistics.Win.Misc.UltraButton();
			this.btCentro = new Infragistics.Win.Misc.UltraButton();
			this.btComprobante = new Infragistics.Win.Misc.UltraButton();
			this.cmbVendedor = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.cdsCompra = new C1.Data.C1DataSet();
			this.cdvPersonal = new C1.Data.C1DataView();
			this.cdsPersonal = new C1.Data.C1DataSet();
			this.txtComprobante = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.cmbSubProyecto = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.cdvSubProyecto = new C1.Data.C1DataView();
			this.cdsProyecto = new C1.Data.C1DataSet();
			this.cmbProyecto = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.cdvProyecto = new C1.Data.C1DataView();
			this.label15 = new System.Windows.Forms.Label();
			this.label5 = new System.Windows.Forms.Label();
			this.txtNotas = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.ultraTabPageControl2 = new Infragistics.Win.UltraWinTabControl.UltraTabPageControl();
			this.txtComision = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label1 = new System.Windows.Forms.Label();
			this.cmbPrecio = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.cdsClienteTabla = new C1.Data.C1DataSet();
			this.btCambiarPrecio = new Infragistics.Win.Misc.UltraButton();
			this.spnVuelto = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.cmdVuelto = new Infragistics.Win.Misc.UltraButton();
			this.label35 = new System.Windows.Forms.Label();
			this.label24 = new System.Windows.Forms.Label();
			this.txtNumCuenta = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label18 = new System.Windows.Forms.Label();
			this.txtCheque = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.cmbFormaPago = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.cdvFormaPago = new C1.Data.C1DataView();
			this.cdsCompraTabla = new C1.Data.C1DataSet();
			this.rdgContadoCred = new Infragistics.Win.UltraWinEditors.UltraOptionSet();
			this.btPago = new Infragistics.Win.Misc.UltraButton();
			this.cmdFormaPago = new Infragistics.Win.Misc.UltraButton();
			this.ultraTabPageControl3 = new Infragistics.Win.UltraWinTabControl.UltraTabPageControl();
			this.btCorresponsal = new Infragistics.Win.Misc.UltraButton();
			this.btComision = new Infragistics.Win.Misc.UltraButton();
			this.btDescuento = new Infragistics.Win.Misc.UltraButton();
			this.btFlete = new Infragistics.Win.Misc.UltraButton();
			this.txtDescPorcIva = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtDescPorcIva0 = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtDescLinea0 = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtDescLinea = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtDescIva = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtDescIva0 = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtFlete = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.btRefresaI = new Infragistics.Win.Misc.UltraButton();
			this.btRefrescaD = new Infragistics.Win.Misc.UltraButton();
			this.label17 = new System.Windows.Forms.Label();
			this.pgEntrega = new Infragistics.Win.UltraWinTabControl.UltraTabPageControl();
			this.btComprob1 = new Infragistics.Win.Misc.UltraButton();
			this.txtComprob1 = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.btImportacion = new Infragistics.Win.Misc.UltraButton();
			this.label38 = new System.Windows.Forms.Label();
			this.cmbRecibe = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.cdvEntrega = new C1.Data.C1DataView();
			this.cdsCliente = new C1.Data.C1DataSet();
			this.label37 = new System.Windows.Forms.Label();
			this.txtOtro = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtDepartamento = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.chkRevisado = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.chkEntregado = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.dtFechaRevision = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.dtFechaEntrega = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.label20 = new System.Windows.Forms.Label();
			this.label19 = new System.Windows.Forms.Label();
			this.ultraTabPageControl5 = new Infragistics.Win.UltraWinTabControl.UltraTabPageControl();
			this.ultraNumericEditor1 = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.ultraNumericEditor2 = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label2 = new System.Windows.Forms.Label();
			this.label3 = new System.Windows.Forms.Label();
			this.txtInteresPagar = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtEmpezarEn = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtPagosCada = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtPorcInteres = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtTotalCuotas = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtValorCuotas = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtAnticipo = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtNumCuotas = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.btTotalCuotas = new Infragistics.Win.Misc.UltraButton();
			this.label25 = new System.Windows.Forms.Label();
			this.label26 = new System.Windows.Forms.Label();
			this.label27 = new System.Windows.Forms.Label();
			this.label28 = new System.Windows.Forms.Label();
			this.label23 = new System.Windows.Forms.Label();
			this.label22 = new System.Windows.Forms.Label();
			this.label21 = new System.Windows.Forms.Label();
			this.TabPageCOA = new Infragistics.Win.UltraWinTabControl.UltraTabPageControl();
			this.label40 = new System.Windows.Forms.Label();
			this.cmbIce = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.chkConvenioInt = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.chkDevolucion = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.OptBienServicio = new Infragistics.Win.UltraWinEditors.UltraOptionSet();
			this.txtRefrendo = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label30 = new System.Windows.Forms.Label();
			this.label29 = new System.Windows.Forms.Label();
			this.cmbCredTribut = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.tbpAutorizacion = new Infragistics.Win.UltraWinTabControl.UltraTabPageControl();
			this.txtUsuario = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label39 = new System.Windows.Forms.Label();
			this.label11 = new System.Windows.Forms.Label();
			this.cmbCaducidad = new Infragistics.Win.UltraWinEditors.UltraDateTimeEditor();
			this.chkAutPresupuestaria = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.chkAutPrincipal = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.spnPrioridad = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label34 = new System.Windows.Forms.Label();
			this.txtSerieFact = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label33 = new System.Windows.Forms.Label();
			this.txtAutImprenta = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label32 = new System.Windows.Forms.Label();
			this.txtAutFactura = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label31 = new System.Windows.Forms.Label();
			this.ultraTabPageControl6 = new Infragistics.Win.UltraWinTabControl.UltraTabPageControl();
			this.chkCodBarras = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.btIngresar = new Infragistics.Win.Misc.UltraButton();
			this.btBuscar = new Infragistics.Win.Misc.UltraButton();
			this.txtBusca = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.ugrBusca = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraTabPageControl4 = new Infragistics.Win.UltraWinTabControl.UltraTabPageControl();
			this.ultraTextEditor6 = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label46 = new System.Windows.Forms.Label();
			this.label45 = new System.Windows.Forms.Label();
			this.ultraTextEditor5 = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label14 = new System.Windows.Forms.Label();
			this.ultraTextEditor4 = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label13 = new System.Windows.Forms.Label();
			this.ultraTextEditor2 = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtCIF = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label44 = new System.Windows.Forms.Label();
			this.cmbImportacion = new Infragistics.Win.UltraWinGrid.UltraDropDown();
			this.cdsTipoImportacion = new C1.Data.C1DataSet();
			this.ultraTextEditor3 = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label43 = new System.Windows.Forms.Label();
			this.ultraTextEditor1 = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label42 = new System.Windows.Forms.Label();
			this.txtPeso = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label36 = new System.Windows.Forms.Label();
			this.label41 = new System.Windows.Forms.Label();
			this.txtBulto = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.udsBusca = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.cdvCliente = new C1.Data.C1DataView();
			this.cdvArticulo = new C1.Data.C1DataView();
			this.cdsArticulo = new C1.Data.C1DataSet();
			this.cdvArtCodigo = new C1.Data.C1DataView();
			this.cdvComprobante = new C1.Data.C1DataView();
			this.cdvNumero = new C1.Data.C1DataView();
			this.lblIva0 = new System.Windows.Forms.Label();
			this.lblIvaSubt = new System.Windows.Forms.Label();
			this.lblIva = new System.Windows.Forms.Label();
			this.lblDescuento = new System.Windows.Forms.Label();
			this.imageList = new System.Windows.Forms.ImageList(this.components);
			this.errorProvider = new System.Windows.Forms.ErrorProvider();
			this.btTotal = new DevExpress.XtraEditors.SimpleButton();
			this.label6 = new System.Windows.Forms.Label();
			this.label7 = new System.Windows.Forms.Label();
			this.label8 = new System.Windows.Forms.Label();
			this.lblBodega = new System.Windows.Forms.Label();
			this.lblCliente = new System.Windows.Forms.Label();
			this.lblSubTotal = new System.Windows.Forms.Label();
			this.lblIce = new System.Windows.Forms.Label();
			this.cdsSeguridad = new C1.Data.C1DataSet();
			this.lblAnulado = new System.Windows.Forms.Label();
			this.CmbDescCod = new Infragistics.Win.UltraWinGrid.UltraDropDown();
			this.ultraGrid1 = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.CmbCodDesc = new Infragistics.Win.UltraWinGrid.UltraDropDown();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.cdvCompra = new C1.Data.C1DataView();
			this.cdsAsiento = new C1.Data.C1DataSet();
			this.cdsRetencion = new C1.Data.C1DataSet();
			this.cdsAsignaCuenta = new C1.Data.C1DataSet();
			this.cmbCliente = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.txtNumero = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.dtFecha = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.spnDias = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.dtVence = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.spnBodega = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.btCliente = new Infragistics.Win.Misc.UltraButton();
			this.cmbComprobante1 = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.cmbTipoCompra = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.btSaldo1 = new Infragistics.Win.Misc.UltraButton();
			this.toolTip = new System.Windows.Forms.ToolTip(this.components);
			this.btBuscaArticulo = new Infragistics.Win.Misc.UltraButton();
			this.btTransforma = new Infragistics.Win.Misc.UltraButton();
			this.btNotaEntrega = new Infragistics.Win.Misc.UltraButton();
			this.btRetencion = new Infragistics.Win.Misc.UltraButton();
			this.btCreaAsiento = new Infragistics.Win.Misc.UltraButton();
			this.btGuiaRemision = new Infragistics.Win.Misc.UltraButton();
			this.btActualizaPrecio = new Infragistics.Win.Misc.UltraButton();
			this.btNuevo = new Infragistics.Win.Misc.UltraButton();
			this.txtIdAsiento = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtIdCompra = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtBorrar = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.barraDato1 = new BarraDatoSQL.BarraDatoSQL();
			this.utcTransaccion = new Infragistics.Win.UltraWinTabControl.UltraTabControl();
			this.ultraTabSharedControlsPage1 = new Infragistics.Win.UltraWinTabControl.UltraTabSharedControlsPage();
			this.txtDescIva0Total = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtDescIvaTotal = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtIva = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtIva0 = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtIceTotal = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtTotal = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtIvaTotal = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.cdsSeteo = new C1.Data.C1DataSet();
			this.cmbProyecto1 = new Infragistics.Win.UltraWinGrid.UltraDropDown();
			this.cmbSubProyecto1 = new Infragistics.Win.UltraWinGrid.UltraDropDown();
			this.helpProvider1 = new System.Windows.Forms.HelpProvider();
			this.cmbUnidad = new Infragistics.Win.UltraWinGrid.UltraDropDown();
			this.cdsUnidad = new C1.Data.C1DataSet();
			this.txtNombre = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.btBarras = new Infragistics.Win.Misc.UltraButton();
			this.btAuditoria = new Infragistics.Win.Misc.UltraButton();
			this.txtServicio = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.lblServicio = new System.Windows.Forms.Label();
			this.lblPropina = new System.Windows.Forms.Label();
			this.cmbBodega = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.udsBodegas = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.lblTiempo = new System.Windows.Forms.Label();
			this.cmbBodegaLinea = new Infragistics.Win.UltraWinGrid.UltraDropDown();
			this.txtPropina = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.cmbUnidadB = new Infragistics.Win.UltraWinGrid.UltraDropDown();
			this.cdsSeteoC = new C1.Data.C1DataSet();
			this.ultraTabPageControl1.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.cmbVendedor)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsCompra)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvPersonal)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsPersonal)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtComprobante)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbSubProyecto)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvSubProyecto)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsProyecto)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbProyecto)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvProyecto)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNotas)).BeginInit();
			this.ultraTabPageControl2.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.txtComision)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbPrecio)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsClienteTabla)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.spnVuelto)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNumCuenta)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCheque)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbFormaPago)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvFormaPago)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsCompraTabla)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.rdgContadoCred)).BeginInit();
			this.ultraTabPageControl3.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.txtDescPorcIva)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDescPorcIva0)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDescLinea0)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDescLinea)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDescIva)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDescIva0)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtFlete)).BeginInit();
			this.pgEntrega.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.txtComprob1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbRecibe)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvEntrega)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsCliente)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtOtro)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDepartamento)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFechaRevision)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFechaEntrega)).BeginInit();
			this.ultraTabPageControl5.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.ultraNumericEditor1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraNumericEditor2)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtInteresPagar)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtEmpezarEn)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtPagosCada)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtPorcInteres)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtTotalCuotas)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtValorCuotas)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtAnticipo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNumCuotas)).BeginInit();
			this.TabPageCOA.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.cmbIce)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.OptBienServicio)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtRefrendo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCredTribut)).BeginInit();
			this.tbpAutorizacion.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.txtUsuario)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCaducidad)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.spnPrioridad)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtSerieFact)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtAutImprenta)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtAutFactura)).BeginInit();
			this.ultraTabPageControl6.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.txtBusca)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ugrBusca)).BeginInit();
			this.ultraTabPageControl4.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.ultraTextEditor6)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraTextEditor5)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraTextEditor4)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraTextEditor2)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCIF)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbImportacion)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsTipoImportacion)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraTextEditor3)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraTextEditor1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtPeso)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtBulto)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.udsBusca)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvCliente)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvArticulo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsArticulo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvArtCodigo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvComprobante)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvNumero)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeguridad)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.CmbDescCod)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.CmbCodDesc)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvCompra)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsAsiento)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsRetencion)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsAsignaCuenta)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCliente)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNumero)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFecha)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.spnDias)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtVence)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.spnBodega)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbComprobante1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbTipoCompra)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdAsiento)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdCompra)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtBorrar)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.utcTransaccion)).BeginInit();
			this.utcTransaccion.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.txtDescIva0Total)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDescIvaTotal)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIva)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIva0)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIceTotal)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtTotal)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIvaTotal)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbProyecto1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbSubProyecto1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbUnidad)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsUnidad)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNombre)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtServicio)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbBodega)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.udsBodegas)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbBodegaLinea)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtPropina)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbUnidadB)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoC)).BeginInit();
			this.SuspendLayout();
			// 
			// ultraTabPageControl1
			// 
			this.ultraTabPageControl1.Controls.Add(this.btSubProyecto);
			this.ultraTabPageControl1.Controls.Add(this.btCentro);
			this.ultraTabPageControl1.Controls.Add(this.btComprobante);
			this.ultraTabPageControl1.Controls.Add(this.cmbVendedor);
			this.ultraTabPageControl1.Controls.Add(this.txtComprobante);
			this.ultraTabPageControl1.Controls.Add(this.cmbSubProyecto);
			this.ultraTabPageControl1.Controls.Add(this.cmbProyecto);
			this.ultraTabPageControl1.Controls.Add(this.label15);
			this.ultraTabPageControl1.Controls.Add(this.label5);
			this.ultraTabPageControl1.Controls.Add(this.txtNotas);
			this.ultraTabPageControl1.Location = new System.Drawing.Point(1, 20);
			this.ultraTabPageControl1.Name = "ultraTabPageControl1";
			this.ultraTabPageControl1.Size = new System.Drawing.Size(486, 115);
			// 
			// btSubProyecto
			// 
			this.btSubProyecto.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.btSubProyecto.Location = new System.Drawing.Point(240, 88);
			this.btSubProyecto.Name = "btSubProyecto";
			this.btSubProyecto.Size = new System.Drawing.Size(72, 23);
			this.btSubProyecto.TabIndex = 93;
			this.btSubProyecto.Text = "Proyecto";
			this.btSubProyecto.Click += new System.EventHandler(this.btSubProyecto_Click);
			this.btSubProyecto.MouseUp += new System.Windows.Forms.MouseEventHandler(this.btSubProyecto_MouseUp);
			// 
			// btCentro
			// 
			this.btCentro.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.btCentro.Location = new System.Drawing.Point(8, 88);
			this.btCentro.Name = "btCentro";
			this.btCentro.Size = new System.Drawing.Size(56, 23);
			this.btCentro.TabIndex = 92;
			this.btCentro.Text = "Centro";
			this.btCentro.Click += new System.EventHandler(this.btCentro_Click);
			this.btCentro.MouseUp += new System.Windows.Forms.MouseEventHandler(this.btCentro_MouseUp);
			// 
			// btComprobante
			// 
			this.btComprobante.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.btComprobante.Location = new System.Drawing.Point(448, 64);
			this.btComprobante.Name = "btComprobante";
			this.btComprobante.Size = new System.Drawing.Size(24, 23);
			this.btComprobante.TabIndex = 91;
			this.btComprobante.Text = "...";
			this.btComprobante.Click += new System.EventHandler(this.btComprobante_Click);
			this.btComprobante.MouseUp += new System.Windows.Forms.MouseEventHandler(this.btComprobante_MouseUp);
			// 
			// cmbVendedor
			// 
			appearance1.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbVendedor.Appearance = appearance1;
			this.cmbVendedor.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbVendedor.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbVendedor.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsCompra, "_Compra.idVendedor"));
			this.cmbVendedor.DataSource = this.cdvPersonal;
			ultraGridBand1.AddButtonCaption = "Cliente";
			ultraGridColumn1.Header.VisiblePosition = 2;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn2.Header.VisiblePosition = 3;
			ultraGridColumn2.Hidden = true;
			ultraGridColumn3.Header.VisiblePosition = 4;
			ultraGridColumn3.Hidden = true;
			ultraGridColumn4.Header.VisiblePosition = 5;
			ultraGridColumn4.Hidden = true;
			ultraGridColumn5.Header.VisiblePosition = 6;
			ultraGridColumn5.Hidden = true;
			ultraGridColumn6.Header.VisiblePosition = 7;
			ultraGridColumn6.Hidden = true;
			ultraGridColumn7.Header.VisiblePosition = 8;
			ultraGridColumn7.Hidden = true;
			ultraGridColumn8.Header.VisiblePosition = 9;
			ultraGridColumn8.Hidden = true;
			ultraGridColumn9.Header.VisiblePosition = 0;
			ultraGridColumn9.Width = 250;
			ultraGridColumn10.Header.VisiblePosition = 10;
			ultraGridColumn10.Hidden = true;
			ultraGridColumn11.Header.VisiblePosition = 1;
			ultraGridColumn12.Header.VisiblePosition = 11;
			ultraGridColumn12.Hidden = true;
			ultraGridColumn13.Header.VisiblePosition = 12;
			ultraGridColumn13.Hidden = true;
			ultraGridColumn14.Header.VisiblePosition = 13;
			ultraGridColumn14.Hidden = true;
			ultraGridColumn15.Header.VisiblePosition = 14;
			ultraGridColumn15.Hidden = true;
			ultraGridColumn16.Header.VisiblePosition = 15;
			ultraGridColumn16.Hidden = true;
			ultraGridColumn17.Header.VisiblePosition = 16;
			ultraGridColumn17.Hidden = true;
			ultraGridColumn18.Header.VisiblePosition = 17;
			ultraGridColumn18.Hidden = true;
			ultraGridColumn19.Header.VisiblePosition = 18;
			ultraGridColumn19.Hidden = true;
			ultraGridColumn20.Header.VisiblePosition = 19;
			ultraGridColumn20.Hidden = true;
			ultraGridColumn21.Header.VisiblePosition = 20;
			ultraGridColumn21.Hidden = true;
			ultraGridColumn22.Header.VisiblePosition = 21;
			ultraGridColumn22.Hidden = true;
			ultraGridColumn23.Header.VisiblePosition = 22;
			ultraGridColumn23.Hidden = true;
			ultraGridColumn24.Header.VisiblePosition = 23;
			ultraGridColumn24.Hidden = true;
			ultraGridColumn25.Header.VisiblePosition = 24;
			ultraGridColumn25.Hidden = true;
			ultraGridColumn26.Header.VisiblePosition = 25;
			ultraGridColumn26.Hidden = true;
			ultraGridColumn27.Header.VisiblePosition = 26;
			ultraGridColumn27.Hidden = true;
			ultraGridColumn28.Header.VisiblePosition = 27;
			ultraGridColumn28.Hidden = true;
			ultraGridColumn29.Header.VisiblePosition = 28;
			ultraGridColumn29.Hidden = true;
			ultraGridColumn30.Header.VisiblePosition = 29;
			ultraGridColumn30.Hidden = true;
			ultraGridColumn31.Header.VisiblePosition = 30;
			ultraGridColumn31.Hidden = true;
			ultraGridColumn32.Header.VisiblePosition = 31;
			ultraGridColumn32.Hidden = true;
			ultraGridColumn33.Header.VisiblePosition = 32;
			ultraGridColumn33.Hidden = true;
			ultraGridColumn34.Header.VisiblePosition = 33;
			ultraGridColumn34.Hidden = true;
			ultraGridColumn35.Header.VisiblePosition = 34;
			ultraGridColumn35.Hidden = true;
			ultraGridColumn36.Header.VisiblePosition = 35;
			ultraGridColumn36.Hidden = true;
			ultraGridColumn37.Header.VisiblePosition = 36;
			ultraGridColumn37.Hidden = true;
			ultraGridColumn38.Header.VisiblePosition = 37;
			ultraGridColumn38.Hidden = true;
			ultraGridColumn39.Header.VisiblePosition = 38;
			ultraGridColumn39.Hidden = true;
			ultraGridColumn40.Header.VisiblePosition = 39;
			ultraGridColumn40.Hidden = true;
			ultraGridColumn41.Header.VisiblePosition = 40;
			ultraGridColumn41.Hidden = true;
			ultraGridColumn42.Header.VisiblePosition = 41;
			ultraGridColumn42.Hidden = true;
			ultraGridColumn43.Header.VisiblePosition = 42;
			ultraGridColumn43.Hidden = true;
			ultraGridColumn44.Header.VisiblePosition = 43;
			ultraGridColumn44.Hidden = true;
			ultraGridColumn45.Header.VisiblePosition = 44;
			ultraGridColumn45.Hidden = true;
			ultraGridColumn46.Header.VisiblePosition = 45;
			ultraGridColumn46.Hidden = true;
			ultraGridColumn47.Header.VisiblePosition = 46;
			ultraGridColumn47.Hidden = true;
			ultraGridColumn48.Header.VisiblePosition = 47;
			ultraGridColumn48.Hidden = true;
			ultraGridColumn49.Header.VisiblePosition = 48;
			ultraGridColumn49.Hidden = true;
			ultraGridColumn50.Header.VisiblePosition = 49;
			ultraGridColumn50.Hidden = true;
			ultraGridColumn51.Header.VisiblePosition = 50;
			ultraGridColumn51.Hidden = true;
			ultraGridColumn52.Header.VisiblePosition = 51;
			ultraGridColumn52.Hidden = true;
			ultraGridColumn53.Header.VisiblePosition = 52;
			ultraGridColumn53.Hidden = true;
			ultraGridColumn54.Header.VisiblePosition = 53;
			ultraGridColumn54.Hidden = true;
			ultraGridColumn55.Header.VisiblePosition = 54;
			ultraGridColumn55.Hidden = true;
			ultraGridColumn56.Header.VisiblePosition = 55;
			ultraGridColumn56.Hidden = true;
			ultraGridColumn57.Header.VisiblePosition = 56;
			ultraGridColumn57.Hidden = true;
			ultraGridColumn58.Header.VisiblePosition = 57;
			ultraGridColumn58.Hidden = true;
			ultraGridColumn59.Header.VisiblePosition = 58;
			ultraGridColumn59.Hidden = true;
			ultraGridColumn60.Header.VisiblePosition = 59;
			ultraGridColumn60.Hidden = true;
			ultraGridColumn61.Header.VisiblePosition = 60;
			ultraGridColumn61.Hidden = true;
			ultraGridColumn62.Header.VisiblePosition = 61;
			ultraGridColumn62.Hidden = true;
			ultraGridColumn63.Header.VisiblePosition = 62;
			ultraGridColumn63.Hidden = true;
			ultraGridColumn64.Header.VisiblePosition = 63;
			ultraGridColumn64.Hidden = true;
			ultraGridColumn65.Header.VisiblePosition = 64;
			ultraGridColumn65.Hidden = true;
			ultraGridColumn66.Header.VisiblePosition = 65;
			ultraGridColumn66.Hidden = true;
			ultraGridColumn67.Header.VisiblePosition = 66;
			ultraGridColumn67.Hidden = true;
			ultraGridColumn68.Header.VisiblePosition = 67;
			ultraGridColumn68.Hidden = true;
			ultraGridColumn69.Header.VisiblePosition = 68;
			ultraGridColumn69.Hidden = true;
			ultraGridColumn70.Header.VisiblePosition = 69;
			ultraGridColumn70.Hidden = true;
			ultraGridColumn71.Header.VisiblePosition = 70;
			ultraGridColumn71.Hidden = true;
			ultraGridColumn72.Header.VisiblePosition = 71;
			ultraGridColumn72.Hidden = true;
			ultraGridColumn73.Header.VisiblePosition = 72;
			ultraGridColumn73.Hidden = true;
			ultraGridColumn74.Header.VisiblePosition = 73;
			ultraGridColumn74.Hidden = true;
			ultraGridColumn75.Header.VisiblePosition = 74;
			ultraGridColumn75.Hidden = true;
			ultraGridColumn76.Header.VisiblePosition = 75;
			ultraGridColumn76.Hidden = true;
			ultraGridColumn77.Header.VisiblePosition = 76;
			ultraGridColumn77.Hidden = true;
			ultraGridColumn78.Header.VisiblePosition = 77;
			ultraGridColumn78.Hidden = true;
			ultraGridColumn79.Header.VisiblePosition = 78;
			ultraGridColumn79.Hidden = true;
			ultraGridColumn80.Header.VisiblePosition = 79;
			ultraGridColumn80.Hidden = true;
			ultraGridColumn81.Header.VisiblePosition = 80;
			ultraGridColumn81.Hidden = true;
			ultraGridColumn82.Header.VisiblePosition = 81;
			ultraGridColumn82.Hidden = true;
			ultraGridColumn83.Header.VisiblePosition = 82;
			ultraGridColumn83.Hidden = true;
			ultraGridColumn84.Header.VisiblePosition = 83;
			ultraGridColumn84.Hidden = true;
			ultraGridColumn85.Header.VisiblePosition = 84;
			ultraGridColumn85.Hidden = true;
			ultraGridColumn86.Header.VisiblePosition = 85;
			ultraGridColumn86.Hidden = true;
			ultraGridColumn87.Header.VisiblePosition = 86;
			ultraGridColumn87.Hidden = true;
			ultraGridColumn88.Header.VisiblePosition = 87;
			ultraGridColumn88.Hidden = true;
			ultraGridColumn89.Header.VisiblePosition = 88;
			ultraGridColumn89.Hidden = true;
			ultraGridColumn90.Header.VisiblePosition = 89;
			ultraGridColumn90.Hidden = true;
			ultraGridColumn91.Header.VisiblePosition = 90;
			ultraGridColumn91.Hidden = true;
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
																										 ultraGridColumn91});
			this.cmbVendedor.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.cmbVendedor.DisplayMember = "Nombre";
			this.cmbVendedor.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbVendedor.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbVendedor.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbVendedor.Location = new System.Drawing.Point(75, 64);
			this.cmbVendedor.Name = "cmbVendedor";
			this.cmbVendedor.Size = new System.Drawing.Size(149, 21);
			this.cmbVendedor.TabIndex = 90;
			this.cmbVendedor.ValueMember = "idPersonal";
			this.cmbVendedor.MouseUp += new System.Windows.Forms.MouseEventHandler(this.cmbVendedor_MouseUp);
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
			this.cdsCompra.PositionChanged += new C1.Data.PositionChangeEventHandler(this.cdsCompra_PositionChanged);
			this.cdsCompra.BeforeFill += new C1.Data.FillEventHandler(this.cdsCompra_BeforeFill);
			// 
			// cdvPersonal
			// 
			this.cdvPersonal.BindingContextCtrl = this;
			this.cdvPersonal.DataSet = this.cdsPersonal;
			this.cdvPersonal.RowFilter = "Vendedor = 1";
			this.cdvPersonal.Sort = "Nombre";
			this.cdvPersonal.TableName = "";
			this.cdvPersonal.TableViewName = "Personal";
			// 
			// cdsPersonal
			// 
			this.cdsPersonal.BindingContextCtrl = this;
			this.cdsPersonal.DataLibrary = "LibPersonal";
			this.cdsPersonal.DataLibraryUrl = "";
			this.cdsPersonal.DataSetDef = "dsPersonal";
			this.cdsPersonal.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsPersonal.Name = "cdsPersonal";
			this.cdsPersonal.SchemaClassName = "LibPersonal.DataClass";
			this.cdsPersonal.SchemaDef = null;
			this.cdsPersonal.BeforeFill += new C1.Data.FillEventHandler(this.cdsPersonal_BeforeFill);
			// 
			// txtComprobante
			// 
			appearance2.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtComprobante.Appearance = appearance2;
			this.txtComprobante.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsCompra, "_Compra.Comprobante"));
			this.txtComprobante.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtComprobante.Location = new System.Drawing.Point(320, 64);
			this.txtComprobante.Name = "txtComprobante";
			this.txtComprobante.Size = new System.Drawing.Size(120, 21);
			this.txtComprobante.TabIndex = 89;
			// 
			// cmbSubProyecto
			// 
			appearance3.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbSubProyecto.Appearance = appearance3;
			this.cmbSubProyecto.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbSubProyecto.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbSubProyecto.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsCompra, "_Compra.idSubProyecto"));
			this.cmbSubProyecto.DataSource = this.cdvSubProyecto;
			ultraGridBand2.AddButtonCaption = "Cliente";
			ultraGridColumn92.Header.VisiblePosition = 3;
			ultraGridColumn92.Hidden = true;
			ultraGridColumn93.Header.VisiblePosition = 2;
			ultraGridColumn93.Hidden = true;
			ultraGridColumn94.Header.VisiblePosition = 0;
			ultraGridColumn94.Width = 341;
			ultraGridColumn95.Header.VisiblePosition = 4;
			ultraGridColumn95.Hidden = true;
			ultraGridColumn96.Header.VisiblePosition = 6;
			ultraGridColumn96.Hidden = true;
			ultraGridColumn97.Header.VisiblePosition = 5;
			ultraGridColumn97.Hidden = true;
			ultraGridColumn98.Header.VisiblePosition = 1;
			ultraGridColumn99.Header.VisiblePosition = 7;
			ultraGridColumn99.Hidden = true;
			ultraGridBand2.Columns.AddRange(new object[] {
																										 ultraGridColumn92,
																										 ultraGridColumn93,
																										 ultraGridColumn94,
																										 ultraGridColumn95,
																										 ultraGridColumn96,
																										 ultraGridColumn97,
																										 ultraGridColumn98,
																										 ultraGridColumn99});
			this.cmbSubProyecto.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			this.cmbSubProyecto.DisplayMember = "Nombre";
			this.cmbSubProyecto.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbSubProyecto.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbSubProyecto.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbSubProyecto.Location = new System.Drawing.Point(320, 88);
			this.cmbSubProyecto.Name = "cmbSubProyecto";
			this.cmbSubProyecto.Size = new System.Drawing.Size(152, 21);
			this.cmbSubProyecto.TabIndex = 88;
			this.cmbSubProyecto.ValueMember = "idSubProyecto";
			this.cmbSubProyecto.MouseUp += new System.Windows.Forms.MouseEventHandler(this.cmbSubProyecto_MouseUp);
			// 
			// cdvSubProyecto
			// 
			this.cdvSubProyecto.BindingContextCtrl = this;
			this.cdvSubProyecto.DataSet = this.cdsProyecto;
			this.cdvSubProyecto.RowFilter = "Visible = 1";
			this.cdvSubProyecto.Sort = "Nombre";
			this.cdvSubProyecto.TableName = "";
			this.cdvSubProyecto.TableViewName = "SubProyecto";
			// 
			// cdsProyecto
			// 
			this.cdsProyecto.BindingContextCtrl = this;
			this.cdsProyecto.DataLibrary = "LibContabilidad";
			this.cdsProyecto.DataLibraryUrl = "";
			this.cdsProyecto.DataSetDef = "dsProyecto";
			this.cdsProyecto.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsProyecto.Name = "cdsProyecto";
			this.cdsProyecto.SchemaClassName = "LibContabilidad.DataClass";
			this.cdsProyecto.SchemaDef = null;
			this.cdsProyecto.BeforeFill += new C1.Data.FillEventHandler(this.cdsProyecto_BeforeFill);
			// 
			// cmbProyecto
			// 
			appearance4.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbProyecto.Appearance = appearance4;
			this.cmbProyecto.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbProyecto.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbProyecto.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsCompra, "_Compra.idProyecto"));
			this.cmbProyecto.DataSource = this.cdvProyecto;
			ultraGridBand3.AddButtonCaption = "Cliente";
			ultraGridColumn100.Header.VisiblePosition = 2;
			ultraGridColumn100.Hidden = true;
			ultraGridColumn101.Header.VisiblePosition = 6;
			ultraGridColumn101.Hidden = true;
			ultraGridColumn102.Header.VisiblePosition = 0;
			ultraGridColumn102.Width = 370;
			ultraGridColumn103.Header.VisiblePosition = 3;
			ultraGridColumn103.Hidden = true;
			ultraGridColumn104.Header.VisiblePosition = 7;
			ultraGridColumn104.Hidden = true;
			ultraGridColumn105.Header.VisiblePosition = 4;
			ultraGridColumn105.Hidden = true;
			ultraGridColumn106.Header.VisiblePosition = 1;
			ultraGridColumn107.Header.VisiblePosition = 5;
			ultraGridColumn107.Hidden = true;
			ultraGridColumn108.Header.VisiblePosition = 8;
			ultraGridColumn108.Hidden = true;
			ultraGridColumn108.Width = 167;
			ultraGridBand3.Columns.AddRange(new object[] {
																										 ultraGridColumn100,
																										 ultraGridColumn101,
																										 ultraGridColumn102,
																										 ultraGridColumn103,
																										 ultraGridColumn104,
																										 ultraGridColumn105,
																										 ultraGridColumn106,
																										 ultraGridColumn107,
																										 ultraGridColumn108});
			this.cmbProyecto.DisplayLayout.BandsSerializer.Add(ultraGridBand3);
			this.cmbProyecto.DisplayMember = "Nombre";
			this.cmbProyecto.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbProyecto.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbProyecto.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbProyecto.Location = new System.Drawing.Point(75, 88);
			this.cmbProyecto.Name = "cmbProyecto";
			this.cmbProyecto.Size = new System.Drawing.Size(149, 21);
			this.cmbProyecto.TabIndex = 87;
			this.cmbProyecto.ValueMember = "idProyecto";
			this.cmbProyecto.Validating += new System.ComponentModel.CancelEventHandler(this.cmbProyecto_Validating);
			this.cmbProyecto.MouseUp += new System.Windows.Forms.MouseEventHandler(this.cmbProyecto_MouseUp);
			// 
			// cdvProyecto
			// 
			this.cdvProyecto.BindingContextCtrl = this;
			this.cdvProyecto.DataSet = this.cdsProyecto;
			this.cdvProyecto.RowFilter = "Visible =1";
			this.cdvProyecto.Sort = "Nombre";
			this.cdvProyecto.TableName = "";
			this.cdvProyecto.TableViewName = "Proyecto";
			// 
			// label15
			// 
			this.label15.AutoSize = true;
			this.label15.BackColor = System.Drawing.Color.Transparent;
			this.label15.Location = new System.Drawing.Point(240, 64);
			this.label15.Name = "label15";
			this.label15.Size = new System.Drawing.Size(73, 16);
			this.label15.TabIndex = 85;
			this.label15.Text = "Comprobante";
			// 
			// label5
			// 
			this.label5.AutoSize = true;
			this.label5.BackColor = System.Drawing.Color.Transparent;
			this.label5.Location = new System.Drawing.Point(11, 64);
			this.label5.Name = "label5";
			this.label5.Size = new System.Drawing.Size(53, 16);
			this.label5.TabIndex = 83;
			this.label5.Text = "Vendedor";
			// 
			// txtNotas
			// 
			this.txtNotas.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			appearance5.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtNotas.Appearance = appearance5;
			this.txtNotas.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsCompra, "_Compra.Notas"));
			this.txtNotas.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNotas.Location = new System.Drawing.Point(8, 8);
			this.txtNotas.Multiline = true;
			this.txtNotas.Name = "txtNotas";
			this.txtNotas.Size = new System.Drawing.Size(466, 48);
			this.txtNotas.TabIndex = 82;
			// 
			// ultraTabPageControl2
			// 
			this.ultraTabPageControl2.Controls.Add(this.txtComision);
			this.ultraTabPageControl2.Controls.Add(this.label1);
			this.ultraTabPageControl2.Controls.Add(this.cmbPrecio);
			this.ultraTabPageControl2.Controls.Add(this.btCambiarPrecio);
			this.ultraTabPageControl2.Controls.Add(this.spnVuelto);
			this.ultraTabPageControl2.Controls.Add(this.cmdVuelto);
			this.ultraTabPageControl2.Controls.Add(this.label35);
			this.ultraTabPageControl2.Controls.Add(this.label24);
			this.ultraTabPageControl2.Controls.Add(this.txtNumCuenta);
			this.ultraTabPageControl2.Controls.Add(this.label18);
			this.ultraTabPageControl2.Controls.Add(this.txtCheque);
			this.ultraTabPageControl2.Controls.Add(this.cmbFormaPago);
			this.ultraTabPageControl2.Controls.Add(this.rdgContadoCred);
			this.ultraTabPageControl2.Controls.Add(this.btPago);
			this.ultraTabPageControl2.Controls.Add(this.cmdFormaPago);
			this.ultraTabPageControl2.Location = new System.Drawing.Point(-10000, -10000);
			this.ultraTabPageControl2.Name = "ultraTabPageControl2";
			this.ultraTabPageControl2.Size = new System.Drawing.Size(486, 115);
			// 
			// txtComision
			// 
			appearance6.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtComision.Appearance = appearance6;
			this.txtComision.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsCompra, "_Compra.Comision"));
			this.txtComision.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtComision.FormatString = "#,##0.00";
			this.txtComision.Location = new System.Drawing.Point(312, 56);
			this.txtComision.MaskInput = "nnnnn.nn";
			this.txtComision.Name = "txtComision";
			this.txtComision.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtComision.PromptChar = ' ';
			this.txtComision.Size = new System.Drawing.Size(104, 21);
			this.txtComision.SpinButtonDisplayStyle = Infragistics.Win.ButtonDisplayStyle.Always;
			this.txtComision.TabIndex = 103;
			// 
			// label1
			// 
			this.label1.AutoSize = true;
			this.label1.BackColor = System.Drawing.Color.Transparent;
			this.label1.Location = new System.Drawing.Point(256, 56);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(52, 16);
			this.label1.TabIndex = 102;
			this.label1.Text = "Comision";
			// 
			// cmbPrecio
			// 
			appearance7.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbPrecio.Appearance = appearance7;
			this.cmbPrecio.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbPrecio.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbPrecio.DataMember = "ClientePrecio";
			this.cmbPrecio.DataSource = this.cdsClienteTabla;
			this.cmbPrecio.DisplayLayout.AutoFitColumns = true;
			ultraGridBand4.AddButtonCaption = "ClienteCiudad";
			ultraGridColumn109.Header.VisiblePosition = 0;
			ultraGridColumn109.Hidden = true;
			ultraGridColumn109.Width = 51;
			ultraGridColumn110.Header.VisiblePosition = 1;
			ultraGridColumn110.Width = 110;
			ultraGridBand4.Columns.AddRange(new object[] {
																										 ultraGridColumn109,
																										 ultraGridColumn110});
			this.cmbPrecio.DisplayLayout.BandsSerializer.Add(ultraGridBand4);
			this.cmbPrecio.DisplayMember = "Precio";
			this.cmbPrecio.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbPrecio.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbPrecio.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbPrecio.Location = new System.Drawing.Point(256, 80);
			this.cmbPrecio.Name = "cmbPrecio";
			this.cmbPrecio.Size = new System.Drawing.Size(128, 21);
			this.cmbPrecio.TabIndex = 100;
			this.toolTip.SetToolTip(this.cmbPrecio, "Escoja Tipo de Precio al que desea modificar");
			this.cmbPrecio.ValueMember = "idTipoPrecio";
			// 
			// cdsClienteTabla
			// 
			this.cdsClienteTabla.BindingContextCtrl = this;
			this.cdsClienteTabla.DataLibrary = "LibFacturacion";
			this.cdsClienteTabla.DataLibraryUrl = "";
			this.cdsClienteTabla.DataSetDef = "dsClienteTabla";
			this.cdsClienteTabla.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsClienteTabla.Name = "cdsClienteTabla";
			this.cdsClienteTabla.SchemaClassName = "LibFacturacion.DataClass";
			this.cdsClienteTabla.SchemaDef = null;
			this.cdsClienteTabla.BeforeFill += new C1.Data.FillEventHandler(this.cdsClienteTabla_BeforeFill);
			// 
			// btCambiarPrecio
			// 
			this.btCambiarPrecio.Location = new System.Drawing.Point(392, 80);
			this.btCambiarPrecio.Name = "btCambiarPrecio";
			this.btCambiarPrecio.Size = new System.Drawing.Size(24, 24);
			this.btCambiarPrecio.TabIndex = 99;
			this.btCambiarPrecio.Text = "$";
			this.btCambiarPrecio.Click += new System.EventHandler(this.btCambiarPrecio_Click);
			this.btCambiarPrecio.MouseUp += new System.Windows.Forms.MouseEventHandler(this.btCambiarPrecio_MouseDown);
			// 
			// spnVuelto
			// 
			appearance8.ForeColorDisabled = System.Drawing.Color.Black;
			this.spnVuelto.Appearance = appearance8;
			this.spnVuelto.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsCompra, "_Compra.Vuelto"));
			this.spnVuelto.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.spnVuelto.FormatString = "#,##0.00";
			this.spnVuelto.Location = new System.Drawing.Point(80, 80);
			this.spnVuelto.MaskInput = "nnnnn.nn";
			this.spnVuelto.Name = "spnVuelto";
			this.spnVuelto.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.spnVuelto.PromptChar = ' ';
			this.spnVuelto.Size = new System.Drawing.Size(88, 21);
			this.spnVuelto.SpinButtonDisplayStyle = Infragistics.Win.ButtonDisplayStyle.Always;
			this.spnVuelto.TabIndex = 98;
			// 
			// cmdVuelto
			// 
			this.cmdVuelto.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.cmdVuelto.Location = new System.Drawing.Point(192, 80);
			this.cmdVuelto.Name = "cmdVuelto";
			this.cmdVuelto.Size = new System.Drawing.Size(48, 23);
			this.cmdVuelto.TabIndex = 57;
			this.cmdVuelto.Text = "Vuelto";
			this.cmdVuelto.Click += new System.EventHandler(this.cmdVuelto_Click);
			// 
			// label35
			// 
			this.label35.AutoSize = true;
			this.label35.BackColor = System.Drawing.Color.Transparent;
			this.label35.Location = new System.Drawing.Point(16, 80);
			this.label35.Name = "label35";
			this.label35.Size = new System.Drawing.Size(52, 16);
			this.label35.TabIndex = 56;
			this.label35.Text = "Recibido:";
			// 
			// label24
			// 
			this.label24.AutoSize = true;
			this.label24.BackColor = System.Drawing.Color.Transparent;
			this.label24.Location = new System.Drawing.Point(256, 32);
			this.label24.Name = "label24";
			this.label24.Size = new System.Drawing.Size(44, 16);
			this.label24.TabIndex = 54;
			this.label24.Text = "Cuenta:";
			// 
			// txtNumCuenta
			// 
			appearance9.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtNumCuenta.Appearance = appearance9;
			this.txtNumCuenta.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsCompra, "_Compra.NumCuenta"));
			this.txtNumCuenta.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNumCuenta.Location = new System.Drawing.Point(312, 32);
			this.txtNumCuenta.Name = "txtNumCuenta";
			this.txtNumCuenta.Size = new System.Drawing.Size(104, 21);
			this.txtNumCuenta.TabIndex = 53;
			// 
			// label18
			// 
			this.label18.AutoSize = true;
			this.label18.BackColor = System.Drawing.Color.Transparent;
			this.label18.Location = new System.Drawing.Point(256, 8);
			this.label18.Name = "label18";
			this.label18.Size = new System.Drawing.Size(47, 16);
			this.label18.TabIndex = 52;
			this.label18.Text = "Cheque:";
			// 
			// txtCheque
			// 
			appearance10.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtCheque.Appearance = appearance10;
			this.txtCheque.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsCompra, "_Compra.NumCheque"));
			this.txtCheque.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtCheque.Location = new System.Drawing.Point(312, 8);
			this.txtCheque.Name = "txtCheque";
			this.txtCheque.Size = new System.Drawing.Size(104, 21);
			this.txtCheque.TabIndex = 51;
			// 
			// cmbFormaPago
			// 
			appearance11.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbFormaPago.Appearance = appearance11;
			this.cmbFormaPago.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbFormaPago.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbFormaPago.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsCompra, "_Compra.idFormaPago"));
			this.cmbFormaPago.DataSource = this.cdvFormaPago;
			this.cmbFormaPago.DisplayLayout.AutoFitColumns = true;
			ultraGridColumn111.Header.VisiblePosition = 3;
			ultraGridColumn111.Hidden = true;
			ultraGridColumn112.Header.VisiblePosition = 4;
			ultraGridColumn112.Hidden = true;
			ultraGridColumn113.Header.VisiblePosition = 0;
			ultraGridColumn113.Width = 134;
			ultraGridColumn114.Header.VisiblePosition = 2;
			ultraGridColumn114.Hidden = true;
			ultraGridColumn115.Header.VisiblePosition = 1;
			ultraGridColumn115.Hidden = true;
			ultraGridColumn116.Header.VisiblePosition = 5;
			ultraGridColumn116.Hidden = true;
			ultraGridColumn116.Width = 29;
			ultraGridColumn117.Header.VisiblePosition = 6;
			ultraGridColumn117.Hidden = true;
			ultraGridColumn117.Width = 58;
			ultraGridBand5.Columns.AddRange(new object[] {
																										 ultraGridColumn111,
																										 ultraGridColumn112,
																										 ultraGridColumn113,
																										 ultraGridColumn114,
																										 ultraGridColumn115,
																										 ultraGridColumn116,
																										 ultraGridColumn117});
			this.cmbFormaPago.DisplayLayout.BandsSerializer.Add(ultraGridBand5);
			this.cmbFormaPago.DisplayMember = "FormaPago";
			this.cmbFormaPago.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbFormaPago.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbFormaPago.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbFormaPago.Location = new System.Drawing.Point(16, 48);
			this.cmbFormaPago.Name = "cmbFormaPago";
			this.cmbFormaPago.Size = new System.Drawing.Size(152, 21);
			this.cmbFormaPago.TabIndex = 50;
			this.cmbFormaPago.ValueMember = "idFormaPago";
			this.cmbFormaPago.Validating += new System.ComponentModel.CancelEventHandler(this.cmbFormaPago_Validating);
			this.cmbFormaPago.Validated += new System.EventHandler(this.cmbFormaPago_Validated);
			// 
			// cdvFormaPago
			// 
			this.cdvFormaPago.BindingContextCtrl = this;
			this.cdvFormaPago.DataSet = this.cdsCompraTabla;
			this.cdvFormaPago.TableName = "";
			this.cdvFormaPago.TableViewName = "CompraFormaPago";
			// 
			// cdsCompraTabla
			// 
			this.cdsCompraTabla.BindingContextCtrl = this;
			this.cdsCompraTabla.DataLibrary = "LibFacturacion";
			this.cdsCompraTabla.DataLibraryUrl = "";
			this.cdsCompraTabla.DataSetDef = "dsCompraTabla";
			this.cdsCompraTabla.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsCompraTabla.Name = "cdsCompraTabla";
			this.cdsCompraTabla.SchemaClassName = "LibFacturacion.DataClass";
			this.cdsCompraTabla.SchemaDef = null;
			this.cdsCompraTabla.BeforeFill += new C1.Data.FillEventHandler(this.cdsComprobante_BeforeFill);
			// 
			// rdgContadoCred
			// 
			appearance12.ForeColorDisabled = System.Drawing.Color.Black;
			this.rdgContadoCred.Appearance = appearance12;
			this.rdgContadoCred.BackColor = System.Drawing.Color.Transparent;
			this.rdgContadoCred.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsCompra, "_Compra.ContadoCredito"));
			this.rdgContadoCred.ItemAppearance = appearance13;
			this.rdgContadoCred.ItemOrigin = new System.Drawing.Point(10, 3);
			valueListItem1.DataValue = 1;
			valueListItem1.DisplayText = "Contado";
			valueListItem2.DataValue = 2;
			valueListItem2.DisplayText = "Crédito";
			this.rdgContadoCred.Items.Add(valueListItem1);
			this.rdgContadoCred.Items.Add(valueListItem2);
			this.rdgContadoCred.Location = new System.Drawing.Point(16, 16);
			this.rdgContadoCred.Name = "rdgContadoCred";
			this.rdgContadoCred.Size = new System.Drawing.Size(152, 24);
			this.rdgContadoCred.TabIndex = 49;
			this.rdgContadoCred.Validated += new System.EventHandler(this.rdgContadoCred_Validated);
			// 
			// btPago
			// 
			this.btPago.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.btPago.Location = new System.Drawing.Point(192, 16);
			this.btPago.Name = "btPago";
			this.btPago.Size = new System.Drawing.Size(48, 23);
			this.btPago.TabIndex = 46;
			this.btPago.Text = "Pagos";
			this.btPago.Click += new System.EventHandler(this.btPagos_Click);
			this.btPago.MouseUp += new System.Windows.Forms.MouseEventHandler(this.btPago_MouseUp);
			// 
			// cmdFormaPago
			// 
			this.cmdFormaPago.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.cmdFormaPago.Location = new System.Drawing.Point(192, 48);
			this.cmdFormaPago.Name = "cmdFormaPago";
			this.cmdFormaPago.Size = new System.Drawing.Size(24, 23);
			this.cmdFormaPago.TabIndex = 45;
			this.cmdFormaPago.Text = "...";
			this.cmdFormaPago.Click += new System.EventHandler(this.cmdFormaPago_Click);
			// 
			// ultraTabPageControl3
			// 
			this.ultraTabPageControl3.Controls.Add(this.btCorresponsal);
			this.ultraTabPageControl3.Controls.Add(this.btComision);
			this.ultraTabPageControl3.Controls.Add(this.btDescuento);
			this.ultraTabPageControl3.Controls.Add(this.btFlete);
			this.ultraTabPageControl3.Controls.Add(this.txtDescPorcIva);
			this.ultraTabPageControl3.Controls.Add(this.txtDescPorcIva0);
			this.ultraTabPageControl3.Controls.Add(this.txtDescLinea0);
			this.ultraTabPageControl3.Controls.Add(this.txtDescLinea);
			this.ultraTabPageControl3.Controls.Add(this.txtDescIva);
			this.ultraTabPageControl3.Controls.Add(this.txtDescIva0);
			this.ultraTabPageControl3.Controls.Add(this.txtFlete);
			this.ultraTabPageControl3.Controls.Add(this.btRefresaI);
			this.ultraTabPageControl3.Controls.Add(this.btRefrescaD);
			this.ultraTabPageControl3.Controls.Add(this.label17);
			this.ultraTabPageControl3.Location = new System.Drawing.Point(-10000, -10000);
			this.ultraTabPageControl3.Name = "ultraTabPageControl3";
			this.ultraTabPageControl3.Size = new System.Drawing.Size(486, 115);
			// 
			// btCorresponsal
			// 
			this.btCorresponsal.ButtonStyle = Infragistics.Win.UIElementButtonStyle.WindowsXPCommandButton;
			this.btCorresponsal.Location = new System.Drawing.Point(328, 72);
			this.btCorresponsal.Name = "btCorresponsal";
			this.btCorresponsal.Size = new System.Drawing.Size(88, 23);
			this.btCorresponsal.TabIndex = 117;
			this.btCorresponsal.Text = "Corresponsal";
			this.btCorresponsal.Click += new System.EventHandler(this.btCorresponsal_Click);
			// 
			// btComision
			// 
			this.btComision.ButtonStyle = Infragistics.Win.UIElementButtonStyle.WindowsXPCommandButton;
			this.btComision.Location = new System.Drawing.Point(232, 72);
			this.btComision.Name = "btComision";
			this.btComision.Size = new System.Drawing.Size(88, 23);
			this.btComision.TabIndex = 116;
			this.btComision.Text = "Comision";
			this.btComision.Click += new System.EventHandler(this.btComision_Click);
			// 
			// btDescuento
			// 
			this.btDescuento.Location = new System.Drawing.Point(16, 16);
			this.btDescuento.Name = "btDescuento";
			this.btDescuento.Size = new System.Drawing.Size(88, 23);
			this.btDescuento.TabIndex = 115;
			this.btDescuento.Text = "Descuento Iva";
			this.btDescuento.Click += new System.EventHandler(this.btDescuento_Click);
			// 
			// btFlete
			// 
			this.btFlete.ButtonStyle = Infragistics.Win.UIElementButtonStyle.WindowsXPCommandButton;
			this.btFlete.Location = new System.Drawing.Point(16, 72);
			this.btFlete.Name = "btFlete";
			this.btFlete.Size = new System.Drawing.Size(80, 23);
			this.btFlete.TabIndex = 114;
			this.btFlete.Text = "Flete";
			this.btFlete.Click += new System.EventHandler(this.btFlete_Click);
			this.btFlete.MouseUp += new System.Windows.Forms.MouseEventHandler(this.btFlete_MouseDown);
			// 
			// txtDescPorcIva
			// 
			appearance14.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtDescPorcIva.Appearance = appearance14;
			this.txtDescPorcIva.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsCompra, "_Compra.DescPorcIva"));
			this.txtDescPorcIva.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtDescPorcIva.FormatString = "#,##0.00";
			this.txtDescPorcIva.Location = new System.Drawing.Point(272, 24);
			this.txtDescPorcIva.Name = "txtDescPorcIva";
			this.txtDescPorcIva.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtDescPorcIva.PromptChar = ' ';
			this.txtDescPorcIva.Size = new System.Drawing.Size(48, 21);
			this.txtDescPorcIva.TabIndex = 113;
			// 
			// txtDescPorcIva0
			// 
			appearance15.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtDescPorcIva0.Appearance = appearance15;
			this.txtDescPorcIva0.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsCompra, "_Compra.DescPorcIva0"));
			this.txtDescPorcIva0.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtDescPorcIva0.FormatString = "#,##0.00";
			this.txtDescPorcIva0.Location = new System.Drawing.Point(272, 48);
			this.txtDescPorcIva0.Name = "txtDescPorcIva0";
			this.txtDescPorcIva0.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtDescPorcIva0.PromptChar = ' ';
			this.txtDescPorcIva0.Size = new System.Drawing.Size(48, 21);
			this.txtDescPorcIva0.TabIndex = 112;
			// 
			// txtDescLinea0
			// 
			this.txtDescLinea0.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsCompra, "_Compra.Desc4"));
			this.txtDescLinea0.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtDescLinea0.FormatString = "#,##0.00";
			this.txtDescLinea0.Location = new System.Drawing.Point(328, 48);
			this.txtDescLinea0.Name = "txtDescLinea0";
			this.txtDescLinea0.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtDescLinea0.PromptChar = ' ';
			this.txtDescLinea0.ReadOnly = true;
			this.txtDescLinea0.Size = new System.Drawing.Size(80, 21);
			this.txtDescLinea0.TabIndex = 111;
			// 
			// txtDescLinea
			// 
			this.txtDescLinea.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsCompra, "_Compra.Desc3"));
			this.txtDescLinea.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtDescLinea.FormatString = "#,##0.00";
			this.txtDescLinea.Location = new System.Drawing.Point(328, 24);
			this.txtDescLinea.Name = "txtDescLinea";
			this.txtDescLinea.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtDescLinea.PromptChar = ' ';
			this.txtDescLinea.ReadOnly = true;
			this.txtDescLinea.Size = new System.Drawing.Size(80, 21);
			this.txtDescLinea.TabIndex = 110;
			// 
			// txtDescIva
			// 
			appearance16.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtDescIva.Appearance = appearance16;
			this.txtDescIva.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsCompra, "_Compra.Desc1"));
			this.txtDescIva.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtDescIva.FormatString = "#,##0.00";
			this.txtDescIva.Location = new System.Drawing.Point(128, 24);
			this.txtDescIva.Name = "txtDescIva";
			this.txtDescIva.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtDescIva.PromptChar = ' ';
			this.txtDescIva.Size = new System.Drawing.Size(96, 21);
			this.txtDescIva.TabIndex = 109;
			this.txtDescIva.Validated += new System.EventHandler(this.btRefrescaD_Click);
			// 
			// txtDescIva0
			// 
			appearance17.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtDescIva0.Appearance = appearance17;
			this.txtDescIva0.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsCompra, "_Compra.Desc2"));
			this.txtDescIva0.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtDescIva0.FormatString = "#,##0.00";
			this.txtDescIva0.Location = new System.Drawing.Point(128, 48);
			this.txtDescIva0.Name = "txtDescIva0";
			this.txtDescIva0.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtDescIva0.PromptChar = ' ';
			this.txtDescIva0.Size = new System.Drawing.Size(96, 21);
			this.txtDescIva0.TabIndex = 108;
			// 
			// txtFlete
			// 
			appearance18.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtFlete.Appearance = appearance18;
			this.txtFlete.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsCompra, "_Compra.Flete"));
			this.txtFlete.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtFlete.FormatString = "#,##0.00";
			this.txtFlete.Location = new System.Drawing.Point(128, 72);
			this.txtFlete.Name = "txtFlete";
			this.txtFlete.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtFlete.PromptChar = ' ';
			this.txtFlete.Size = new System.Drawing.Size(96, 21);
			this.txtFlete.TabIndex = 107;
			// 
			// btRefresaI
			// 
			this.btRefresaI.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.btRefresaI.Location = new System.Drawing.Point(231, 47);
			this.btRefresaI.Name = "btRefresaI";
			this.btRefresaI.Size = new System.Drawing.Size(32, 23);
			this.btRefresaI.TabIndex = 81;
			this.btRefresaI.Text = "<-";
			this.btRefresaI.Click += new System.EventHandler(this.btRefresaI_Click);
			// 
			// btRefrescaD
			// 
			this.btRefrescaD.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.btRefrescaD.Location = new System.Drawing.Point(231, 23);
			this.btRefrescaD.Name = "btRefrescaD";
			this.btRefrescaD.Size = new System.Drawing.Size(32, 23);
			this.btRefrescaD.TabIndex = 80;
			this.btRefrescaD.Text = "->";
			this.btRefrescaD.Click += new System.EventHandler(this.btRefrescaD_Click);
			// 
			// label17
			// 
			this.label17.AutoSize = true;
			this.label17.BackColor = System.Drawing.Color.Transparent;
			this.label17.Location = new System.Drawing.Point(15, 47);
			this.label17.Name = "label17";
			this.label17.Size = new System.Drawing.Size(89, 16);
			this.label17.TabIndex = 79;
			this.label17.Text = "Descuento IVA 0";
			// 
			// pgEntrega
			// 
			this.pgEntrega.Controls.Add(this.btComprob1);
			this.pgEntrega.Controls.Add(this.txtComprob1);
			this.pgEntrega.Controls.Add(this.btImportacion);
			this.pgEntrega.Controls.Add(this.label38);
			this.pgEntrega.Controls.Add(this.cmbRecibe);
			this.pgEntrega.Controls.Add(this.label37);
			this.pgEntrega.Controls.Add(this.txtOtro);
			this.pgEntrega.Controls.Add(this.txtDepartamento);
			this.pgEntrega.Controls.Add(this.chkRevisado);
			this.pgEntrega.Controls.Add(this.chkEntregado);
			this.pgEntrega.Controls.Add(this.dtFechaRevision);
			this.pgEntrega.Controls.Add(this.dtFechaEntrega);
			this.pgEntrega.Controls.Add(this.label20);
			this.pgEntrega.Controls.Add(this.label19);
			this.pgEntrega.Location = new System.Drawing.Point(-10000, -10000);
			this.pgEntrega.Name = "pgEntrega";
			this.pgEntrega.Size = new System.Drawing.Size(486, 115);
			// 
			// btComprob1
			// 
			this.btComprob1.Location = new System.Drawing.Point(344, 8);
			this.btComprob1.Name = "btComprob1";
			this.btComprob1.Size = new System.Drawing.Size(80, 23);
			this.btComprob1.TabIndex = 75;
			this.btComprob1.Text = "Comprob. 1";
			this.btComprob1.Click += new System.EventHandler(this.btComprob1_Click);
			// 
			// txtComprob1
			// 
			appearance19.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtComprob1.Appearance = appearance19;
			this.txtComprob1.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsCompra, "_Compra.Comprob1"));
			this.txtComprob1.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtComprob1.Location = new System.Drawing.Point(328, 32);
			this.txtComprob1.Name = "txtComprob1";
			this.txtComprob1.Size = new System.Drawing.Size(104, 21);
			this.txtComprob1.TabIndex = 74;
			// 
			// btImportacion
			// 
			this.btImportacion.Location = new System.Drawing.Point(16, 64);
			this.btImportacion.Name = "btImportacion";
			this.btImportacion.Size = new System.Drawing.Size(80, 23);
			this.btImportacion.TabIndex = 73;
			this.btImportacion.Text = "Importacion";
			this.btImportacion.Click += new System.EventHandler(this.ultraButton1_Click);
			this.btImportacion.MouseUp += new System.Windows.Forms.MouseEventHandler(this.btImportacion_MouseUp);
			// 
			// label38
			// 
			this.label38.AutoSize = true;
			this.label38.BackColor = System.Drawing.Color.Transparent;
			this.label38.Location = new System.Drawing.Point(24, 88);
			this.label38.Name = "label38";
			this.label38.Size = new System.Drawing.Size(43, 16);
			this.label38.TabIndex = 72;
			this.label38.Text = "Recibe:";
			// 
			// cmbRecibe
			// 
			appearance20.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbRecibe.Appearance = appearance20;
			this.cmbRecibe.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbRecibe.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbRecibe.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsCompra, "_Compra.idRecibe"));
			this.cmbRecibe.DataSource = this.cdvEntrega;
			ultraGridColumn118.Header.VisiblePosition = 1;
			ultraGridColumn119.Header.VisiblePosition = 3;
			ultraGridColumn119.Hidden = true;
			ultraGridColumn120.Header.VisiblePosition = 0;
			ultraGridColumn120.Width = 231;
			ultraGridColumn121.Header.VisiblePosition = 4;
			ultraGridColumn121.Hidden = true;
			ultraGridColumn122.Header.VisiblePosition = 5;
			ultraGridColumn122.Hidden = true;
			ultraGridColumn123.Header.VisiblePosition = 2;
			ultraGridColumn123.Hidden = true;
			ultraGridColumn124.Header.VisiblePosition = 6;
			ultraGridColumn124.Hidden = true;
			ultraGridColumn125.Header.VisiblePosition = 7;
			ultraGridColumn125.Hidden = true;
			ultraGridColumn126.Header.VisiblePosition = 8;
			ultraGridColumn126.Hidden = true;
			ultraGridColumn127.Header.VisiblePosition = 9;
			ultraGridColumn127.Hidden = true;
			ultraGridBand6.Columns.AddRange(new object[] {
																										 ultraGridColumn118,
																										 ultraGridColumn119,
																										 ultraGridColumn120,
																										 ultraGridColumn121,
																										 ultraGridColumn122,
																										 ultraGridColumn123,
																										 ultraGridColumn124,
																										 ultraGridColumn125,
																										 ultraGridColumn126,
																										 ultraGridColumn127});
			this.cmbRecibe.DisplayLayout.BandsSerializer.Add(ultraGridBand6);
			this.cmbRecibe.DisplayMember = "Nombre";
			this.cmbRecibe.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbRecibe.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbRecibe.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbRecibe.Location = new System.Drawing.Point(112, 88);
			this.cmbRecibe.Name = "cmbRecibe";
			this.cmbRecibe.Size = new System.Drawing.Size(304, 21);
			this.cmbRecibe.TabIndex = 71;
			this.cmbRecibe.ValueMember = "idCliente";
			this.cmbRecibe.MouseUp += new System.Windows.Forms.MouseEventHandler(this.cmbRecibe_MouseUp);
			// 
			// cdvEntrega
			// 
			this.cdvEntrega.BindingContextCtrl = this;
			this.cdvEntrega.DataSet = this.cdsCliente;
			this.cdvEntrega.RowFilter = "Entrega = 1";
			this.cdvEntrega.Sort = "Nombre";
			this.cdvEntrega.TableName = "";
			this.cdvEntrega.TableViewName = "Cliente";
			// 
			// cdsCliente
			// 
			this.cdsCliente.BindingContextCtrl = this;
			this.cdsCliente.CaseSensitive = false;
			this.cdsCliente.DataLibrary = "LibFacturacion";
			this.cdsCliente.DataLibraryUrl = "";
			this.cdsCliente.DataSetDef = "dsClientes";
			this.cdsCliente.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsCliente.Name = "cdsCliente";
			this.cdsCliente.SchemaClassName = "LibFacturacion.DataClass";
			this.cdsCliente.SchemaDef = null;
			this.cdsCliente.BeforeFill += new C1.Data.FillEventHandler(this.cdsCliente_BeforeFill);
			// 
			// label37
			// 
			this.label37.AutoSize = true;
			this.label37.BackColor = System.Drawing.Color.Transparent;
			this.label37.Location = new System.Drawing.Point(248, 64);
			this.label37.Name = "label37";
			this.label37.Size = new System.Drawing.Size(26, 16);
			this.label37.TabIndex = 58;
			this.label37.Text = "Otro";
			// 
			// txtOtro
			// 
			appearance21.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtOtro.Appearance = appearance21;
			this.txtOtro.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsCompra, "_Compra.Otro"));
			this.txtOtro.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtOtro.Location = new System.Drawing.Point(288, 64);
			this.txtOtro.Name = "txtOtro";
			this.txtOtro.Size = new System.Drawing.Size(128, 21);
			this.txtOtro.TabIndex = 57;
			// 
			// txtDepartamento
			// 
			appearance22.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtDepartamento.Appearance = appearance22;
			this.txtDepartamento.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsCompra, "_Compra.Departamento"));
			this.txtDepartamento.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtDepartamento.Location = new System.Drawing.Point(112, 64);
			this.txtDepartamento.Name = "txtDepartamento";
			this.txtDepartamento.Size = new System.Drawing.Size(128, 21);
			this.txtDepartamento.TabIndex = 55;
			// 
			// chkRevisado
			// 
			appearance23.ForeColorDisabled = System.Drawing.Color.Black;
			this.chkRevisado.Appearance = appearance23;
			this.chkRevisado.BackColor = System.Drawing.Color.Transparent;
			this.chkRevisado.DataBindings.Add(new System.Windows.Forms.Binding("CheckedValue", this.cdsCompra, "_Compra.Revisado"));
			this.chkRevisado.Location = new System.Drawing.Point(232, 40);
			this.chkRevisado.Name = "chkRevisado";
			this.chkRevisado.Size = new System.Drawing.Size(80, 20);
			this.chkRevisado.TabIndex = 53;
			this.chkRevisado.Text = "Revisado";
			// 
			// chkEntregado
			// 
			appearance24.ForeColorDisabled = System.Drawing.Color.Black;
			this.chkEntregado.Appearance = appearance24;
			this.chkEntregado.BackColor = System.Drawing.Color.Transparent;
			this.chkEntregado.DataBindings.Add(new System.Windows.Forms.Binding("CheckedValue", this.cdsCompra, "_Compra.Entregado"));
			this.chkEntregado.Location = new System.Drawing.Point(232, 16);
			this.chkEntregado.Name = "chkEntregado";
			this.chkEntregado.Size = new System.Drawing.Size(80, 20);
			this.chkEntregado.TabIndex = 52;
			this.chkEntregado.Text = "Entregado";
			// 
			// dtFechaRevision
			// 
			appearance25.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtFechaRevision.Appearance = appearance25;
			this.dtFechaRevision.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.dtFechaRevision.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsCompra, "_Compra.FechaRevision"));
			dateButton1.Caption = "Today";
			this.dtFechaRevision.DateButtons.Add(dateButton1);
			this.dtFechaRevision.Format = "dd/MMM/yyyy";
			this.dtFechaRevision.Location = new System.Drawing.Point(112, 40);
			this.dtFechaRevision.Name = "dtFechaRevision";
			this.dtFechaRevision.NonAutoSizeHeight = 23;
			this.dtFechaRevision.Size = new System.Drawing.Size(112, 21);
			this.dtFechaRevision.SpinButtonsVisible = true;
			this.dtFechaRevision.TabIndex = 51;
			this.dtFechaRevision.Value = new System.DateTime(2005, 7, 1, 0, 0, 0, 0);
			// 
			// dtFechaEntrega
			// 
			appearance26.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtFechaEntrega.Appearance = appearance26;
			this.dtFechaEntrega.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.dtFechaEntrega.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsCompra, "_Compra.FechaEntrega"));
			dateButton2.Caption = "Today";
			this.dtFechaEntrega.DateButtons.Add(dateButton2);
			this.dtFechaEntrega.Format = "dd/MMM/yyyy";
			this.dtFechaEntrega.Location = new System.Drawing.Point(112, 16);
			this.dtFechaEntrega.Name = "dtFechaEntrega";
			this.dtFechaEntrega.NonAutoSizeHeight = 23;
			this.dtFechaEntrega.Size = new System.Drawing.Size(112, 21);
			this.dtFechaEntrega.SpinButtonsVisible = true;
			this.dtFechaEntrega.TabIndex = 50;
			this.dtFechaEntrega.Value = new System.DateTime(2005, 7, 1, 0, 0, 0, 0);
			// 
			// label20
			// 
			this.label20.BackColor = System.Drawing.Color.Transparent;
			this.label20.Location = new System.Drawing.Point(16, 40);
			this.label20.Name = "label20";
			this.label20.Size = new System.Drawing.Size(98, 16);
			this.label20.TabIndex = 45;
			this.label20.Text = "Fecha de Revisión";
			// 
			// label19
			// 
			this.label19.BackColor = System.Drawing.Color.Transparent;
			this.label19.Location = new System.Drawing.Point(16, 16);
			this.label19.Name = "label19";
			this.label19.Size = new System.Drawing.Size(94, 16);
			this.label19.TabIndex = 44;
			this.label19.Text = "Fecha de Entrega";
			// 
			// ultraTabPageControl5
			// 
			this.ultraTabPageControl5.Controls.Add(this.ultraNumericEditor1);
			this.ultraTabPageControl5.Controls.Add(this.ultraNumericEditor2);
			this.ultraTabPageControl5.Controls.Add(this.label2);
			this.ultraTabPageControl5.Controls.Add(this.label3);
			this.ultraTabPageControl5.Controls.Add(this.txtInteresPagar);
			this.ultraTabPageControl5.Controls.Add(this.txtEmpezarEn);
			this.ultraTabPageControl5.Controls.Add(this.txtPagosCada);
			this.ultraTabPageControl5.Controls.Add(this.txtPorcInteres);
			this.ultraTabPageControl5.Controls.Add(this.txtTotalCuotas);
			this.ultraTabPageControl5.Controls.Add(this.txtValorCuotas);
			this.ultraTabPageControl5.Controls.Add(this.txtAnticipo);
			this.ultraTabPageControl5.Controls.Add(this.txtNumCuotas);
			this.ultraTabPageControl5.Controls.Add(this.btTotalCuotas);
			this.ultraTabPageControl5.Controls.Add(this.label25);
			this.ultraTabPageControl5.Controls.Add(this.label26);
			this.ultraTabPageControl5.Controls.Add(this.label27);
			this.ultraTabPageControl5.Controls.Add(this.label28);
			this.ultraTabPageControl5.Controls.Add(this.label23);
			this.ultraTabPageControl5.Controls.Add(this.label22);
			this.ultraTabPageControl5.Controls.Add(this.label21);
			this.ultraTabPageControl5.Location = new System.Drawing.Point(-10000, -10000);
			this.ultraTabPageControl5.Name = "ultraTabPageControl5";
			this.ultraTabPageControl5.Size = new System.Drawing.Size(486, 115);
			// 
			// ultraNumericEditor1
			// 
			appearance27.ForeColorDisabled = System.Drawing.Color.Black;
			this.ultraNumericEditor1.Appearance = appearance27;
			this.ultraNumericEditor1.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsCompra, "_Compra.Interes2"));
			this.ultraNumericEditor1.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.ultraNumericEditor1.FormatString = "#,##0.00";
			this.ultraNumericEditor1.Location = new System.Drawing.Point(432, 56);
			this.ultraNumericEditor1.Name = "ultraNumericEditor1";
			this.ultraNumericEditor1.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.ultraNumericEditor1.PromptChar = ' ';
			this.ultraNumericEditor1.Size = new System.Drawing.Size(48, 21);
			this.ultraNumericEditor1.TabIndex = 121;
			// 
			// ultraNumericEditor2
			// 
			appearance28.ForeColorDisabled = System.Drawing.Color.Black;
			this.ultraNumericEditor2.Appearance = appearance28;
			this.ultraNumericEditor2.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsCompra, "_Compra.Dias2"));
			this.ultraNumericEditor2.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.ultraNumericEditor2.FormatString = "#,##0.00";
			this.ultraNumericEditor2.Location = new System.Drawing.Point(312, 56);
			this.ultraNumericEditor2.Name = "ultraNumericEditor2";
			this.ultraNumericEditor2.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.ultraNumericEditor2.PromptChar = ' ';
			this.ultraNumericEditor2.Size = new System.Drawing.Size(48, 21);
			this.ultraNumericEditor2.TabIndex = 120;
			// 
			// label2
			// 
			this.label2.AutoSize = true;
			this.label2.BackColor = System.Drawing.Color.Transparent;
			this.label2.Location = new System.Drawing.Point(376, 56);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(48, 16);
			this.label2.TabIndex = 119;
			this.label2.Text = "Interes 2";
			// 
			// label3
			// 
			this.label3.AutoSize = true;
			this.label3.BackColor = System.Drawing.Color.Transparent;
			this.label3.Location = new System.Drawing.Point(224, 56);
			this.label3.Name = "label3";
			this.label3.Size = new System.Drawing.Size(39, 16);
			this.label3.TabIndex = 118;
			this.label3.Text = "Dias 2:";
			// 
			// txtInteresPagar
			// 
			appearance29.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtInteresPagar.Appearance = appearance29;
			this.txtInteresPagar.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsCompra, "_Compra.ValorInteres"));
			this.txtInteresPagar.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtInteresPagar.FormatString = "#,##0.00";
			this.txtInteresPagar.Location = new System.Drawing.Point(312, 80);
			this.txtInteresPagar.Name = "txtInteresPagar";
			this.txtInteresPagar.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtInteresPagar.PromptChar = ' ';
			this.txtInteresPagar.Size = new System.Drawing.Size(96, 21);
			this.txtInteresPagar.TabIndex = 117;
			// 
			// txtEmpezarEn
			// 
			appearance30.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtEmpezarEn.Appearance = appearance30;
			this.txtEmpezarEn.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsCompra, "_Compra.DiasInicio"));
			this.txtEmpezarEn.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtEmpezarEn.FormatString = "#,##0.00";
			this.txtEmpezarEn.Location = new System.Drawing.Point(432, 32);
			this.txtEmpezarEn.Name = "txtEmpezarEn";
			this.txtEmpezarEn.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtEmpezarEn.PromptChar = ' ';
			this.txtEmpezarEn.Size = new System.Drawing.Size(48, 21);
			this.txtEmpezarEn.TabIndex = 116;
			// 
			// txtPagosCada
			// 
			appearance31.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtPagosCada.Appearance = appearance31;
			this.txtPagosCada.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsCompra, "_Compra.PagosCada"));
			this.txtPagosCada.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtPagosCada.FormatString = "#,##0.00";
			this.txtPagosCada.Location = new System.Drawing.Point(312, 32);
			this.txtPagosCada.Name = "txtPagosCada";
			this.txtPagosCada.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtPagosCada.PromptChar = ' ';
			this.txtPagosCada.Size = new System.Drawing.Size(48, 21);
			this.txtPagosCada.TabIndex = 115;
			// 
			// txtPorcInteres
			// 
			appearance32.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtPorcInteres.Appearance = appearance32;
			this.txtPorcInteres.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsCompra, "_Compra.Interes"));
			this.txtPorcInteres.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtPorcInteres.FormatString = "#,##0.00";
			this.txtPorcInteres.Location = new System.Drawing.Point(312, 8);
			this.txtPorcInteres.Name = "txtPorcInteres";
			this.txtPorcInteres.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtPorcInteres.PromptChar = ' ';
			this.txtPorcInteres.Size = new System.Drawing.Size(96, 21);
			this.txtPorcInteres.TabIndex = 114;
			// 
			// txtTotalCuotas
			// 
			appearance33.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtTotalCuotas.Appearance = appearance33;
			this.txtTotalCuotas.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsCompra, "_Compra.TotalCuotas"));
			this.txtTotalCuotas.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtTotalCuotas.FormatString = "#,##0.00";
			this.txtTotalCuotas.Location = new System.Drawing.Point(104, 80);
			this.txtTotalCuotas.Name = "txtTotalCuotas";
			this.txtTotalCuotas.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtTotalCuotas.PromptChar = ' ';
			this.txtTotalCuotas.Size = new System.Drawing.Size(96, 21);
			this.txtTotalCuotas.TabIndex = 113;
			// 
			// txtValorCuotas
			// 
			appearance34.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtValorCuotas.Appearance = appearance34;
			this.txtValorCuotas.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsCompra, "_Compra.ValorCuotas"));
			this.txtValorCuotas.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtValorCuotas.FormatString = "#,##0.00";
			this.txtValorCuotas.Location = new System.Drawing.Point(104, 56);
			this.txtValorCuotas.Name = "txtValorCuotas";
			this.txtValorCuotas.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtValorCuotas.PromptChar = ' ';
			this.txtValorCuotas.Size = new System.Drawing.Size(96, 21);
			this.txtValorCuotas.TabIndex = 112;
			// 
			// txtAnticipo
			// 
			appearance35.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtAnticipo.Appearance = appearance35;
			this.txtAnticipo.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsCompra, "_Compra.Contado"));
			this.txtAnticipo.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtAnticipo.FormatString = "#,##0.00";
			this.txtAnticipo.Location = new System.Drawing.Point(104, 32);
			this.txtAnticipo.Name = "txtAnticipo";
			this.txtAnticipo.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtAnticipo.PromptChar = ' ';
			this.txtAnticipo.Size = new System.Drawing.Size(96, 21);
			this.txtAnticipo.TabIndex = 111;
			// 
			// txtNumCuotas
			// 
			appearance36.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtNumCuotas.Appearance = appearance36;
			this.txtNumCuotas.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsCompra, "_Compra.NumCuotas"));
			this.txtNumCuotas.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNumCuotas.FormatString = "#,##0.00";
			this.txtNumCuotas.Location = new System.Drawing.Point(104, 8);
			this.txtNumCuotas.Name = "txtNumCuotas";
			this.txtNumCuotas.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtNumCuotas.PromptChar = ' ';
			this.txtNumCuotas.Size = new System.Drawing.Size(96, 21);
			this.txtNumCuotas.TabIndex = 110;
			// 
			// btTotalCuotas
			// 
			this.btTotalCuotas.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.btTotalCuotas.Location = new System.Drawing.Point(8, 80);
			this.btTotalCuotas.Name = "btTotalCuotas";
			this.btTotalCuotas.Size = new System.Drawing.Size(73, 23);
			this.btTotalCuotas.TabIndex = 106;
			this.btTotalCuotas.Text = "Total Pagar";
			this.btTotalCuotas.Click += new System.EventHandler(this.btTotalCuotas_Click);
			this.btTotalCuotas.MouseUp += new System.Windows.Forms.MouseEventHandler(this.btTotalCuotas_MouseDown);
			// 
			// label25
			// 
			this.label25.AutoSize = true;
			this.label25.BackColor = System.Drawing.Color.Transparent;
			this.label25.Location = new System.Drawing.Point(223, 82);
			this.label25.Name = "label25";
			this.label25.Size = new System.Drawing.Size(81, 16);
			this.label25.TabIndex = 104;
			this.label25.Text = "Interes a Pagar";
			// 
			// label26
			// 
			this.label26.AutoSize = true;
			this.label26.BackColor = System.Drawing.Color.Transparent;
			this.label26.Location = new System.Drawing.Point(376, 32);
			this.label26.Name = "label26";
			this.label26.Size = new System.Drawing.Size(49, 16);
			this.label26.TabIndex = 102;
			this.label26.Text = "Empezar";
			// 
			// label27
			// 
			this.label27.AutoSize = true;
			this.label27.BackColor = System.Drawing.Color.Transparent;
			this.label27.Location = new System.Drawing.Point(223, 34);
			this.label27.Name = "label27";
			this.label27.Size = new System.Drawing.Size(70, 16);
			this.label27.TabIndex = 100;
			this.label27.Text = "Pagos Cada:";
			// 
			// label28
			// 
			this.label28.AutoSize = true;
			this.label28.BackColor = System.Drawing.Color.Transparent;
			this.label28.Location = new System.Drawing.Point(223, 10);
			this.label28.Name = "label28";
			this.label28.Size = new System.Drawing.Size(67, 16);
			this.label28.TabIndex = 98;
			this.label28.Text = "% de Interes";
			// 
			// label23
			// 
			this.label23.AutoSize = true;
			this.label23.BackColor = System.Drawing.Color.Transparent;
			this.label23.Location = new System.Drawing.Point(15, 58);
			this.label23.Name = "label23";
			this.label23.Size = new System.Drawing.Size(70, 16);
			this.label23.TabIndex = 95;
			this.label23.Text = "Valor Cuotas";
			// 
			// label22
			// 
			this.label22.AutoSize = true;
			this.label22.BackColor = System.Drawing.Color.Transparent;
			this.label22.Location = new System.Drawing.Point(15, 34);
			this.label22.Name = "label22";
			this.label22.Size = new System.Drawing.Size(45, 16);
			this.label22.TabIndex = 93;
			this.label22.Text = "Anticipo";
			// 
			// label21
			// 
			this.label21.AutoSize = true;
			this.label21.BackColor = System.Drawing.Color.Transparent;
			this.label21.Location = new System.Drawing.Point(15, 10);
			this.label21.Name = "label21";
			this.label21.Size = new System.Drawing.Size(61, 16);
			this.label21.TabIndex = 91;
			this.label21.Text = "No. Cuotas";
			// 
			// TabPageCOA
			// 
			this.TabPageCOA.Controls.Add(this.label40);
			this.TabPageCOA.Controls.Add(this.cmbIce);
			this.TabPageCOA.Controls.Add(this.chkConvenioInt);
			this.TabPageCOA.Controls.Add(this.chkDevolucion);
			this.TabPageCOA.Controls.Add(this.OptBienServicio);
			this.TabPageCOA.Controls.Add(this.txtRefrendo);
			this.TabPageCOA.Controls.Add(this.label30);
			this.TabPageCOA.Controls.Add(this.label29);
			this.TabPageCOA.Controls.Add(this.cmbCredTribut);
			this.TabPageCOA.Location = new System.Drawing.Point(-10000, -10000);
			this.TabPageCOA.Name = "TabPageCOA";
			this.TabPageCOA.Size = new System.Drawing.Size(486, 115);
			// 
			// label40
			// 
			this.label40.AutoSize = true;
			this.label40.BackColor = System.Drawing.Color.Transparent;
			this.label40.Location = new System.Drawing.Point(232, 8);
			this.label40.Name = "label40";
			this.label40.Size = new System.Drawing.Size(23, 16);
			this.label40.TabIndex = 95;
			this.label40.Text = "ICE";
			// 
			// cmbIce
			// 
			appearance37.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbIce.Appearance = appearance37;
			this.cmbIce.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbIce.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbIce.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsCompra, "_Compra.idPorcIce"));
			this.cmbIce.DataMember = "PorcIce";
			this.cmbIce.DataSource = this.cdsCompraTabla;
			ultraGridBand7.AddButtonCaption = "CompraCredTribut";
			ultraGridColumn128.Header.VisiblePosition = 0;
			ultraGridColumn128.Hidden = true;
			ultraGridColumn129.Header.VisiblePosition = 1;
			ultraGridColumn129.Hidden = true;
			ultraGridColumn130.Header.VisiblePosition = 3;
			ultraGridColumn131.Header.VisiblePosition = 2;
			ultraGridColumn132.Header.VisiblePosition = 4;
			ultraGridColumn132.Hidden = true;
			ultraGridBand7.Columns.AddRange(new object[] {
																										 ultraGridColumn128,
																										 ultraGridColumn129,
																										 ultraGridColumn130,
																										 ultraGridColumn131,
																										 ultraGridColumn132});
			this.cmbIce.DisplayLayout.BandsSerializer.Add(ultraGridBand7);
			this.cmbIce.DisplayMember = "concepto";
			this.cmbIce.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbIce.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbIce.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbIce.Location = new System.Drawing.Point(264, 8);
			this.cmbIce.Name = "cmbIce";
			this.cmbIce.Size = new System.Drawing.Size(144, 21);
			this.cmbIce.TabIndex = 94;
			this.cmbIce.ValueMember = "idPorcIce";
			this.cmbIce.MouseUp += new System.Windows.Forms.MouseEventHandler(this.cmbIce_MouseDown);
			// 
			// chkConvenioInt
			// 
			appearance38.ForeColorDisabled = System.Drawing.Color.Black;
			this.chkConvenioInt.Appearance = appearance38;
			this.chkConvenioInt.BackColor = System.Drawing.Color.Transparent;
			this.chkConvenioInt.DataBindings.Add(new System.Windows.Forms.Binding("CheckedValue", this.cdsCompra, "_Compra.ConvenioInternac"));
			this.chkConvenioInt.Location = new System.Drawing.Point(232, 64);
			this.chkConvenioInt.Name = "chkConvenioInt";
			this.chkConvenioInt.Size = new System.Drawing.Size(160, 20);
			this.chkConvenioInt.TabIndex = 93;
			this.chkConvenioInt.Text = "Convenio Internacional";
			// 
			// chkDevolucion
			// 
			appearance39.ForeColorDisabled = System.Drawing.Color.Black;
			this.chkDevolucion.Appearance = appearance39;
			this.chkDevolucion.BackColor = System.Drawing.Color.Transparent;
			this.chkDevolucion.DataBindings.Add(new System.Windows.Forms.Binding("CheckedValue", this.cdsCompra, "_Compra.DerechoDevolucion"));
			this.chkDevolucion.Location = new System.Drawing.Point(232, 40);
			this.chkDevolucion.Name = "chkDevolucion";
			this.chkDevolucion.Size = new System.Drawing.Size(160, 20);
			this.chkDevolucion.TabIndex = 92;
			this.chkDevolucion.Text = "Derecho a Devolucion";
			// 
			// OptBienServicio
			// 
			appearance40.ForeColorDisabled = System.Drawing.Color.Black;
			this.OptBienServicio.Appearance = appearance40;
			this.OptBienServicio.BackColor = System.Drawing.Color.Transparent;
			this.OptBienServicio.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsCompra, "_Compra.BienServicio"));
			this.OptBienServicio.ItemAppearance = appearance41;
			this.OptBienServicio.ItemOrigin = new System.Drawing.Point(4, 0);
			valueListItem3.DataValue = 1;
			valueListItem3.DisplayText = "Transferencia de Bienes";
			valueListItem4.DataValue = 2;
			valueListItem4.DisplayText = "Prestacion de Servicios";
			this.OptBienServicio.Items.Add(valueListItem3);
			this.OptBienServicio.Items.Add(valueListItem4);
			this.OptBienServicio.ItemSpacingVertical = 6;
			this.OptBienServicio.Location = new System.Drawing.Point(16, 40);
			this.OptBienServicio.Name = "OptBienServicio";
			this.OptBienServicio.Size = new System.Drawing.Size(160, 48);
			this.OptBienServicio.TabIndex = 91;
			// 
			// txtRefrendo
			// 
			appearance42.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtRefrendo.Appearance = appearance42;
			this.txtRefrendo.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsCompra, "_Compra.Refrendo"));
			this.txtRefrendo.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtRefrendo.Location = new System.Drawing.Point(288, 88);
			this.txtRefrendo.Name = "txtRefrendo";
			this.txtRefrendo.Size = new System.Drawing.Size(120, 21);
			this.txtRefrendo.TabIndex = 90;
			// 
			// label30
			// 
			this.label30.AutoSize = true;
			this.label30.BackColor = System.Drawing.Color.Transparent;
			this.label30.Location = new System.Drawing.Point(232, 88);
			this.label30.Name = "label30";
			this.label30.Size = new System.Drawing.Size(51, 16);
			this.label30.TabIndex = 53;
			this.label30.Text = "Refrendo";
			// 
			// label29
			// 
			this.label29.AutoSize = true;
			this.label29.BackColor = System.Drawing.Color.Transparent;
			this.label29.Location = new System.Drawing.Point(8, 8);
			this.label29.Name = "label29";
			this.label29.Size = new System.Drawing.Size(58, 16);
			this.label29.TabIndex = 52;
			this.label29.Text = "Cred. Trib.";
			// 
			// cmbCredTribut
			// 
			appearance43.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbCredTribut.Appearance = appearance43;
			this.cmbCredTribut.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbCredTribut.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbCredTribut.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsCompra, "_Compra.idCredTributario"));
			this.cmbCredTribut.DataMember = "CompraCredTribut";
			this.cmbCredTribut.DataSource = this.cdsCompraTabla;
			ultraGridColumn133.Header.VisiblePosition = 2;
			ultraGridColumn133.Hidden = true;
			ultraGridColumn133.Width = 55;
			ultraGridColumn134.Header.VisiblePosition = 1;
			ultraGridColumn134.Hidden = true;
			ultraGridColumn134.Width = 24;
			ultraGridColumn135.Header.VisiblePosition = 0;
			ultraGridColumn135.Width = 329;
			ultraGridBand8.Columns.AddRange(new object[] {
																										 ultraGridColumn133,
																										 ultraGridColumn134,
																										 ultraGridColumn135});
			this.cmbCredTribut.DisplayLayout.BandsSerializer.Add(ultraGridBand8);
			this.cmbCredTribut.DisplayMember = "CredTributario";
			this.cmbCredTribut.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbCredTribut.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbCredTribut.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbCredTribut.Location = new System.Drawing.Point(72, 8);
			this.cmbCredTribut.Name = "cmbCredTribut";
			this.cmbCredTribut.Size = new System.Drawing.Size(144, 21);
			this.cmbCredTribut.TabIndex = 51;
			this.cmbCredTribut.ValueMember = "idCredTributario";
			this.cmbCredTribut.MouseUp += new System.Windows.Forms.MouseEventHandler(this.cmbCredTribut_MouseDown);
			// 
			// tbpAutorizacion
			// 
			this.tbpAutorizacion.Controls.Add(this.txtUsuario);
			this.tbpAutorizacion.Controls.Add(this.label39);
			this.tbpAutorizacion.Controls.Add(this.label11);
			this.tbpAutorizacion.Controls.Add(this.cmbCaducidad);
			this.tbpAutorizacion.Controls.Add(this.chkAutPresupuestaria);
			this.tbpAutorizacion.Controls.Add(this.chkAutPrincipal);
			this.tbpAutorizacion.Controls.Add(this.spnPrioridad);
			this.tbpAutorizacion.Controls.Add(this.label34);
			this.tbpAutorizacion.Controls.Add(this.txtSerieFact);
			this.tbpAutorizacion.Controls.Add(this.label33);
			this.tbpAutorizacion.Controls.Add(this.txtAutImprenta);
			this.tbpAutorizacion.Controls.Add(this.label32);
			this.tbpAutorizacion.Controls.Add(this.txtAutFactura);
			this.tbpAutorizacion.Controls.Add(this.label31);
			this.tbpAutorizacion.Location = new System.Drawing.Point(-10000, -10000);
			this.tbpAutorizacion.Name = "tbpAutorizacion";
			this.tbpAutorizacion.Size = new System.Drawing.Size(486, 115);
			// 
			// txtUsuario
			// 
			appearance44.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtUsuario.Appearance = appearance44;
			this.txtUsuario.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsCompra, "_Compra.Usuario"));
			this.txtUsuario.Location = new System.Drawing.Point(288, 80);
			this.txtUsuario.Name = "txtUsuario";
			this.txtUsuario.ReadOnly = true;
			this.txtUsuario.Size = new System.Drawing.Size(136, 21);
			this.txtUsuario.TabIndex = 103;
			// 
			// label39
			// 
			this.label39.AutoSize = true;
			this.label39.BackColor = System.Drawing.Color.Transparent;
			this.label39.Location = new System.Drawing.Point(232, 80);
			this.label39.Name = "label39";
			this.label39.Size = new System.Drawing.Size(43, 16);
			this.label39.TabIndex = 102;
			this.label39.Text = "Usuario";
			// 
			// label11
			// 
			this.label11.AutoSize = true;
			this.label11.BackColor = System.Drawing.Color.Transparent;
			this.label11.Location = new System.Drawing.Point(16, 80);
			this.label11.Name = "label11";
			this.label11.Size = new System.Drawing.Size(71, 16);
			this.label11.TabIndex = 101;
			this.label11.Text = "F. Caducidad";
			// 
			// cmbCaducidad
			// 
			appearance45.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbCaducidad.Appearance = appearance45;
			this.cmbCaducidad.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsCompra, "_Compra.FechaCaducidad"));
			this.cmbCaducidad.DateTime = new System.DateTime(2006, 3, 17, 0, 0, 0, 0);
			this.cmbCaducidad.Location = new System.Drawing.Point(96, 80);
			this.cmbCaducidad.Name = "cmbCaducidad";
			this.cmbCaducidad.Size = new System.Drawing.Size(120, 21);
			this.cmbCaducidad.TabIndex = 100;
			this.cmbCaducidad.Value = new System.DateTime(2006, 3, 17, 0, 0, 0, 0);
			// 
			// chkAutPresupuestaria
			// 
			appearance46.ForeColorDisabled = System.Drawing.Color.Black;
			this.chkAutPresupuestaria.Appearance = appearance46;
			this.chkAutPresupuestaria.BackColor = System.Drawing.Color.Transparent;
			this.chkAutPresupuestaria.DataBindings.Add(new System.Windows.Forms.Binding("CheckedValue", this.cdsCompra, "_Compra.AutPresupuesto"));
			this.chkAutPresupuestaria.Location = new System.Drawing.Point(232, 56);
			this.chkAutPresupuestaria.Name = "chkAutPresupuestaria";
			this.chkAutPresupuestaria.Size = new System.Drawing.Size(176, 20);
			this.chkAutPresupuestaria.TabIndex = 99;
			this.chkAutPresupuestaria.Text = "Autorizacion Presupuestaria";
			// 
			// chkAutPrincipal
			// 
			appearance47.ForeColorDisabled = System.Drawing.Color.Black;
			this.chkAutPrincipal.Appearance = appearance47;
			this.chkAutPrincipal.BackColor = System.Drawing.Color.Transparent;
			this.chkAutPrincipal.DataBindings.Add(new System.Windows.Forms.Binding("CheckedValue", this.cdsCompra, "_Compra.AutPrincipal"));
			this.chkAutPrincipal.Location = new System.Drawing.Point(232, 32);
			this.chkAutPrincipal.Name = "chkAutPrincipal";
			this.chkAutPrincipal.Size = new System.Drawing.Size(160, 20);
			this.chkAutPrincipal.TabIndex = 98;
			this.chkAutPrincipal.Text = "Autorizacion Principal";
			this.chkAutPrincipal.Click += new System.EventHandler(this.chkAutPrincipal_Click);
			this.chkAutPrincipal.CheckedChanged += new System.EventHandler(this.chkAutPrincipal_CheckedChanged);
			// 
			// spnPrioridad
			// 
			appearance48.ForeColorDisabled = System.Drawing.Color.Black;
			this.spnPrioridad.Appearance = appearance48;
			this.spnPrioridad.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsCompra, "_Compra.Prioridad"));
			this.spnPrioridad.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.spnPrioridad.Location = new System.Drawing.Point(304, 8);
			this.spnPrioridad.MaskInput = "nn";
			this.spnPrioridad.Name = "spnPrioridad";
			this.spnPrioridad.PromptChar = ' ';
			this.spnPrioridad.Size = new System.Drawing.Size(56, 21);
			this.spnPrioridad.SpinButtonDisplayStyle = Infragistics.Win.ButtonDisplayStyle.Always;
			this.spnPrioridad.TabIndex = 97;
			// 
			// label34
			// 
			this.label34.AutoSize = true;
			this.label34.BackColor = System.Drawing.Color.Transparent;
			this.label34.Location = new System.Drawing.Point(232, 8);
			this.label34.Name = "label34";
			this.label34.Size = new System.Drawing.Size(50, 16);
			this.label34.TabIndex = 96;
			this.label34.Text = "Prioridad";
			// 
			// txtSerieFact
			// 
			appearance49.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtSerieFact.Appearance = appearance49;
			this.txtSerieFact.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsCompra, "_Compra.SerieFactura"));
			this.txtSerieFact.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtSerieFact.Location = new System.Drawing.Point(96, 56);
			this.txtSerieFact.Name = "txtSerieFact";
			this.txtSerieFact.Size = new System.Drawing.Size(120, 21);
			this.txtSerieFact.TabIndex = 95;
			// 
			// label33
			// 
			this.label33.AutoSize = true;
			this.label33.BackColor = System.Drawing.Color.Transparent;
			this.label33.Location = new System.Drawing.Point(16, 56);
			this.label33.Name = "label33";
			this.label33.Size = new System.Drawing.Size(72, 16);
			this.label33.TabIndex = 94;
			this.label33.Text = "Serie Factura";
			// 
			// txtAutImprenta
			// 
			appearance50.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtAutImprenta.Appearance = appearance50;
			this.txtAutImprenta.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsCompra, "_Compra.AutImprenta"));
			this.txtAutImprenta.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtAutImprenta.Location = new System.Drawing.Point(96, 32);
			this.txtAutImprenta.Name = "txtAutImprenta";
			this.txtAutImprenta.Size = new System.Drawing.Size(120, 21);
			this.txtAutImprenta.TabIndex = 93;
			// 
			// label32
			// 
			this.label32.AutoSize = true;
			this.label32.BackColor = System.Drawing.Color.Transparent;
			this.label32.Location = new System.Drawing.Point(16, 32);
			this.label32.Name = "label32";
			this.label32.Size = new System.Drawing.Size(72, 16);
			this.label32.TabIndex = 92;
			this.label32.Text = "Aut. Imprenta";
			// 
			// txtAutFactura
			// 
			appearance51.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtAutFactura.Appearance = appearance51;
			this.txtAutFactura.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsCompra, "_Compra.AutFactura"));
			this.txtAutFactura.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtAutFactura.Location = new System.Drawing.Point(96, 8);
			this.txtAutFactura.Name = "txtAutFactura";
			this.txtAutFactura.Size = new System.Drawing.Size(120, 21);
			this.txtAutFactura.TabIndex = 91;
			this.txtAutFactura.MouseUp += new System.Windows.Forms.MouseEventHandler(this.txtAutFactura_MouseDown);
			// 
			// label31
			// 
			this.label31.AutoSize = true;
			this.label31.BackColor = System.Drawing.Color.Transparent;
			this.label31.Location = new System.Drawing.Point(16, 8);
			this.label31.Name = "label31";
			this.label31.Size = new System.Drawing.Size(66, 16);
			this.label31.TabIndex = 53;
			this.label31.Text = "Aut. Factura";
			// 
			// ultraTabPageControl6
			// 
			this.ultraTabPageControl6.Controls.Add(this.chkCodBarras);
			this.ultraTabPageControl6.Controls.Add(this.btIngresar);
			this.ultraTabPageControl6.Controls.Add(this.btBuscar);
			this.ultraTabPageControl6.Controls.Add(this.txtBusca);
			this.ultraTabPageControl6.Controls.Add(this.ugrBusca);
			this.ultraTabPageControl6.Location = new System.Drawing.Point(-10000, -10000);
			this.ultraTabPageControl6.Name = "ultraTabPageControl6";
			this.ultraTabPageControl6.Size = new System.Drawing.Size(486, 115);
			// 
			// chkCodBarras
			// 
			this.chkCodBarras.Location = new System.Drawing.Point(8, 88);
			this.chkCodBarras.Name = "chkCodBarras";
			this.chkCodBarras.Size = new System.Drawing.Size(88, 20);
			this.chkCodBarras.TabIndex = 4;
			this.chkCodBarras.Text = "Cod. Barras";
			// 
			// btIngresar
			// 
			this.btIngresar.Location = new System.Drawing.Point(16, 64);
			this.btIngresar.Name = "btIngresar";
			this.btIngresar.Size = new System.Drawing.Size(72, 23);
			this.btIngresar.TabIndex = 3;
			this.btIngresar.Text = "Ingresar";
			this.btIngresar.Click += new System.EventHandler(this.btIngresar_Click);
			// 
			// btBuscar
			// 
			this.btBuscar.Location = new System.Drawing.Point(16, 32);
			this.btBuscar.Name = "btBuscar";
			this.btBuscar.Size = new System.Drawing.Size(72, 23);
			this.btBuscar.TabIndex = 2;
			this.btBuscar.Text = "Buscar";
			this.btBuscar.Click += new System.EventHandler(this.btBuscarArtciculos);
			// 
			// txtBusca
			// 
			appearance52.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtBusca.Appearance = appearance52;
			this.txtBusca.Location = new System.Drawing.Point(8, 8);
			this.txtBusca.Name = "txtBusca";
			this.txtBusca.Size = new System.Drawing.Size(88, 21);
			this.txtBusca.TabIndex = 1;
			this.txtBusca.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtBusca_KeyPress);
			// 
			// ugrBusca
			// 
			this.ugrBusca.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.ugrBusca.Cursor = System.Windows.Forms.Cursors.Default;
			this.ugrBusca.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortSingle;
			this.ugrBusca.Location = new System.Drawing.Point(104, 8);
			this.ugrBusca.Name = "ugrBusca";
			this.ugrBusca.Size = new System.Drawing.Size(376, 104);
			this.ugrBusca.TabIndex = 0;
			this.ugrBusca.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.ugrBusca_InitializeLayout);
			// 
			// ultraTabPageControl4
			// 
			this.ultraTabPageControl4.Controls.Add(this.ultraTextEditor6);
			this.ultraTabPageControl4.Controls.Add(this.label46);
			this.ultraTabPageControl4.Controls.Add(this.label45);
			this.ultraTabPageControl4.Controls.Add(this.ultraTextEditor5);
			this.ultraTabPageControl4.Controls.Add(this.label14);
			this.ultraTabPageControl4.Controls.Add(this.ultraTextEditor4);
			this.ultraTabPageControl4.Controls.Add(this.label13);
			this.ultraTabPageControl4.Controls.Add(this.ultraTextEditor2);
			this.ultraTabPageControl4.Controls.Add(this.txtCIF);
			this.ultraTabPageControl4.Controls.Add(this.label44);
			this.ultraTabPageControl4.Controls.Add(this.cmbImportacion);
			this.ultraTabPageControl4.Controls.Add(this.ultraTextEditor3);
			this.ultraTabPageControl4.Controls.Add(this.label43);
			this.ultraTabPageControl4.Controls.Add(this.ultraTextEditor1);
			this.ultraTabPageControl4.Controls.Add(this.label42);
			this.ultraTabPageControl4.Controls.Add(this.txtPeso);
			this.ultraTabPageControl4.Controls.Add(this.label36);
			this.ultraTabPageControl4.Controls.Add(this.label41);
			this.ultraTabPageControl4.Controls.Add(this.txtBulto);
			this.ultraTabPageControl4.Location = new System.Drawing.Point(-10000, -10000);
			this.ultraTabPageControl4.Name = "ultraTabPageControl4";
			this.ultraTabPageControl4.Size = new System.Drawing.Size(486, 115);
			// 
			// ultraTextEditor6
			// 
			appearance53.ForeColorDisabled = System.Drawing.Color.Black;
			this.ultraTextEditor6.Appearance = appearance53;
			this.ultraTextEditor6.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsCompra, "_Compra.Mensaje2"));
			this.ultraTextEditor6.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.ultraTextEditor6.Location = new System.Drawing.Point(360, 16);
			this.ultraTextEditor6.Name = "ultraTextEditor6";
			this.ultraTextEditor6.Size = new System.Drawing.Size(88, 21);
			this.ultraTextEditor6.TabIndex = 7;
			// 
			// label46
			// 
			this.label46.AutoSize = true;
			this.label46.BackColor = System.Drawing.Color.Transparent;
			this.label46.Location = new System.Drawing.Point(328, 16);
			this.label46.Name = "label46";
			this.label46.Size = new System.Drawing.Size(31, 16);
			this.label46.TabIndex = 79;
			this.label46.Text = "Emb.";
			// 
			// label45
			// 
			this.label45.AutoSize = true;
			this.label45.BackColor = System.Drawing.Color.Transparent;
			this.label45.Location = new System.Drawing.Point(184, 64);
			this.label45.Name = "label45";
			this.label45.Size = new System.Drawing.Size(21, 16);
			this.label45.TabIndex = 78;
			this.label45.Text = "Via";
			// 
			// ultraTextEditor5
			// 
			appearance54.ForeColorDisabled = System.Drawing.Color.Black;
			this.ultraTextEditor5.Appearance = appearance54;
			this.ultraTextEditor5.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsCompra, "_Compra.Mensaje3"));
			this.ultraTextEditor5.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.ultraTextEditor5.Location = new System.Drawing.Point(224, 64);
			this.ultraTextEditor5.Name = "ultraTextEditor5";
			this.ultraTextEditor5.Size = new System.Drawing.Size(96, 21);
			this.ultraTextEditor5.TabIndex = 6;
			// 
			// label14
			// 
			this.label14.AutoSize = true;
			this.label14.BackColor = System.Drawing.Color.Transparent;
			this.label14.Location = new System.Drawing.Point(184, 40);
			this.label14.Name = "label14";
			this.label14.Size = new System.Drawing.Size(31, 16);
			this.label14.TabIndex = 76;
			this.label14.Text = "AWG";
			// 
			// ultraTextEditor4
			// 
			appearance55.ForeColorDisabled = System.Drawing.Color.Black;
			this.ultraTextEditor4.Appearance = appearance55;
			this.ultraTextEditor4.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsCompra, "_Compra.Mensaje1"));
			this.ultraTextEditor4.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.ultraTextEditor4.Location = new System.Drawing.Point(224, 40);
			this.ultraTextEditor4.Name = "ultraTextEditor4";
			this.ultraTextEditor4.Size = new System.Drawing.Size(96, 21);
			this.ultraTextEditor4.TabIndex = 5;
			// 
			// label13
			// 
			this.label13.AutoSize = true;
			this.label13.BackColor = System.Drawing.Color.Transparent;
			this.label13.Location = new System.Drawing.Point(184, 16);
			this.label13.Name = "label13";
			this.label13.Size = new System.Drawing.Size(28, 16);
			this.label13.TabIndex = 72;
			this.label13.Text = "Guia";
			// 
			// ultraTextEditor2
			// 
			appearance56.ForeColorDisabled = System.Drawing.Color.Black;
			this.ultraTextEditor2.Appearance = appearance56;
			this.ultraTextEditor2.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsCompra, "_Compra.Otro"));
			this.ultraTextEditor2.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.ultraTextEditor2.Location = new System.Drawing.Point(224, 16);
			this.ultraTextEditor2.Name = "ultraTextEditor2";
			this.ultraTextEditor2.Size = new System.Drawing.Size(96, 21);
			this.ultraTextEditor2.TabIndex = 4;
			// 
			// txtCIF
			// 
			this.txtCIF.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsCompra, "_Compra.Cif"));
			this.txtCIF.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtCIF.FormatString = "#,##0.00";
			this.txtCIF.Location = new System.Drawing.Point(360, 64);
			this.txtCIF.Name = "txtCIF";
			this.txtCIF.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtCIF.PromptChar = ' ';
			this.txtCIF.Size = new System.Drawing.Size(88, 21);
			this.txtCIF.TabIndex = 8;
			// 
			// label44
			// 
			this.label44.AutoSize = true;
			this.label44.BackColor = System.Drawing.Color.Transparent;
			this.label44.Location = new System.Drawing.Point(328, 64);
			this.label44.Name = "label44";
			this.label44.Size = new System.Drawing.Size(22, 16);
			this.label44.TabIndex = 69;
			this.label44.Text = "CIF";
			// 
			// cmbImportacion
			// 
			this.cmbImportacion.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbImportacion.DataMember = "ImportacionTipo";
			this.cmbImportacion.DataSource = this.cdsTipoImportacion;
			this.cmbImportacion.DisplayLayout.AutoFitColumns = true;
			ultraGridColumn136.Header.VisiblePosition = 1;
			ultraGridColumn136.Hidden = true;
			ultraGridColumn137.Header.VisiblePosition = 0;
			ultraGridColumn137.Width = 278;
			ultraGridColumn138.Header.VisiblePosition = 2;
			ultraGridColumn138.Hidden = true;
			ultraGridColumn139.Header.VisiblePosition = 3;
			ultraGridColumn139.Hidden = true;
			ultraGridBand9.Columns.AddRange(new object[] {
																										 ultraGridColumn136,
																										 ultraGridColumn137,
																										 ultraGridColumn138,
																										 ultraGridColumn139});
			this.cmbImportacion.DisplayLayout.BandsSerializer.Add(ultraGridBand9);
			this.cmbImportacion.DisplayMember = "Nombre";
			this.cmbImportacion.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbImportacion.Location = new System.Drawing.Point(168, 16);
			this.cmbImportacion.Name = "cmbImportacion";
			this.cmbImportacion.Size = new System.Drawing.Size(280, 96);
			this.cmbImportacion.TabIndex = 68;
			this.cmbImportacion.ValueMember = "idImportacionTipo";
			this.cmbImportacion.Visible = false;
			// 
			// cdsTipoImportacion
			// 
			this.cdsTipoImportacion.BindingContextCtrl = this;
			this.cdsTipoImportacion.DataLibrary = "LibFacturacion";
			this.cdsTipoImportacion.DataLibraryUrl = "";
			this.cdsTipoImportacion.DataSetDef = "dsTipoImportacion";
			this.cdsTipoImportacion.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsTipoImportacion.Name = "cdsTipoImportacion";
			this.cdsTipoImportacion.SchemaClassName = "LibFacturacion.DataClass";
			this.cdsTipoImportacion.SchemaDef = null;
			this.cdsTipoImportacion.BeforeFill += new C1.Data.FillEventHandler(this.cdsTipoImportacion_BeforeFill);
			// 
			// ultraTextEditor3
			// 
			appearance57.ForeColorDisabled = System.Drawing.Color.Black;
			this.ultraTextEditor3.Appearance = appearance57;
			this.ultraTextEditor3.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsCompra, "_Compra.Depart"));
			this.ultraTextEditor3.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.ultraTextEditor3.Location = new System.Drawing.Point(72, 88);
			this.ultraTextEditor3.Name = "ultraTextEditor3";
			this.ultraTextEditor3.Size = new System.Drawing.Size(100, 21);
			this.ultraTextEditor3.TabIndex = 3;
			// 
			// label43
			// 
			this.label43.AutoSize = true;
			this.label43.BackColor = System.Drawing.Color.Transparent;
			this.label43.Location = new System.Drawing.Point(16, 88);
			this.label43.Name = "label43";
			this.label43.Size = new System.Drawing.Size(41, 16);
			this.label43.TabIndex = 62;
			this.label43.Text = "Depart.";
			// 
			// ultraTextEditor1
			// 
			appearance58.ForeColorDisabled = System.Drawing.Color.Black;
			this.ultraTextEditor1.Appearance = appearance58;
			this.ultraTextEditor1.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsCompra, "_Compra.Pedido"));
			this.ultraTextEditor1.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.ultraTextEditor1.Location = new System.Drawing.Point(72, 64);
			this.ultraTextEditor1.Name = "ultraTextEditor1";
			this.ultraTextEditor1.Size = new System.Drawing.Size(100, 21);
			this.ultraTextEditor1.TabIndex = 2;
			// 
			// label42
			// 
			this.label42.AutoSize = true;
			this.label42.BackColor = System.Drawing.Color.Transparent;
			this.label42.Location = new System.Drawing.Point(16, 64);
			this.label42.Name = "label42";
			this.label42.Size = new System.Drawing.Size(40, 16);
			this.label42.TabIndex = 60;
			this.label42.Text = "Pedido";
			// 
			// txtPeso
			// 
			appearance59.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtPeso.Appearance = appearance59;
			this.txtPeso.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsCompra, "_Compra.Peso"));
			this.txtPeso.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtPeso.Location = new System.Drawing.Point(72, 40);
			this.txtPeso.Name = "txtPeso";
			this.txtPeso.Size = new System.Drawing.Size(100, 21);
			this.txtPeso.TabIndex = 1;
			// 
			// label36
			// 
			this.label36.AutoSize = true;
			this.label36.BackColor = System.Drawing.Color.Transparent;
			this.label36.Location = new System.Drawing.Point(16, 40);
			this.label36.Name = "label36";
			this.label36.Size = new System.Drawing.Size(33, 16);
			this.label36.TabIndex = 58;
			this.label36.Text = "Peso:";
			// 
			// label41
			// 
			this.label41.AutoSize = true;
			this.label41.BackColor = System.Drawing.Color.Transparent;
			this.label41.Location = new System.Drawing.Point(16, 16);
			this.label41.Name = "label41";
			this.label41.Size = new System.Drawing.Size(46, 16);
			this.label41.TabIndex = 56;
			this.label41.Text = "Bulto / #";
			// 
			// txtBulto
			// 
			appearance60.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtBulto.Appearance = appearance60;
			this.txtBulto.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsCompra, "_Compra.Bulto"));
			this.txtBulto.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtBulto.Location = new System.Drawing.Point(72, 16);
			this.txtBulto.Name = "txtBulto";
			this.txtBulto.Size = new System.Drawing.Size(100, 21);
			this.txtBulto.TabIndex = 0;
			// 
			// udsBusca
			// 
			ultraDataColumn1.DataType = typeof(int);
			ultraDataColumn5.DataType = typeof(System.Double);
			ultraDataColumn6.DataType = typeof(System.Double);
			ultraDataColumn7.DataType = typeof(System.Double);
			ultraDataColumn8.DataType = typeof(System.Double);
			ultraDataColumn9.DataType = typeof(System.Double);
			ultraDataColumn10.DataType = typeof(System.Double);
			this.udsBusca.Band.Columns.AddRange(new object[] {
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
																												 ultraDataColumn11});
			// 
			// cdvCliente
			// 
			this.cdvCliente.BindingContextCtrl = this;
			this.cdvCliente.DataSet = this.cdsCliente;
			this.cdvCliente.Sort = "Nombre";
			this.cdvCliente.TableName = "";
			this.cdvCliente.TableViewName = "Cliente";
			// 
			// cdvArticulo
			// 
			this.cdvArticulo.BindingContextCtrl = this;
			this.cdvArticulo.DataSet = this.cdsArticulo;
			this.cdvArticulo.RowFilter = "descontinuado = 0";
			this.cdvArticulo.Sort = "Articulo";
			this.cdvArticulo.TableName = "";
			this.cdvArticulo.TableViewName = "Articulo";
			// 
			// cdsArticulo
			// 
			this.cdsArticulo.BindingContextCtrl = this;
			this.cdsArticulo.CaseSensitive = false;
			this.cdsArticulo.DataLibrary = "LibFacturacion";
			this.cdsArticulo.DataLibraryUrl = "";
			this.cdsArticulo.DataSetDef = "dsArticulos";
			this.cdsArticulo.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsArticulo.Name = "cdsArticulo";
			this.cdsArticulo.SchemaClassName = "LibFacturacion.DataClass";
			this.cdsArticulo.SchemaDef = null;
			this.cdsArticulo.BeforeFill += new C1.Data.FillEventHandler(this.cdsArticulo_BeforeFill);
			// 
			// cdvArtCodigo
			// 
			this.cdvArtCodigo.BindingContextCtrl = this;
			this.cdvArtCodigo.DataSet = this.cdsArticulo;
			this.cdvArtCodigo.RowFilter = "descontinuado = 0";
			this.cdvArtCodigo.Sort = "Codigo";
			this.cdvArtCodigo.TableName = "";
			this.cdvArtCodigo.TableViewName = "Articulo";
			// 
			// cdvComprobante
			// 
			this.cdvComprobante.BindingContextCtrl = this;
			this.cdvComprobante.DataSet = this.cdsCompraTabla;
			this.cdvComprobante.TableName = "";
			this.cdvComprobante.TableViewName = "CompraComprobante";
			// 
			// cdvNumero
			// 
			this.cdvNumero.BindingContextCtrl = this;
			this.cdvNumero.DataSet = this.cdsCompraTabla;
			this.cdvNumero.TableName = "";
			this.cdvNumero.TableViewName = "CompraNumero";
			// 
			// lblIva0
			// 
			this.lblIva0.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.lblIva0.AutoSize = true;
			this.lblIva0.Location = new System.Drawing.Point(602, 248);
			this.lblIva0.Name = "lblIva0";
			this.lblIva0.Size = new System.Drawing.Size(32, 16);
			this.lblIva0.TabIndex = 14;
			this.lblIva0.Text = "IVA 0";
			// 
			// lblIvaSubt
			// 
			this.lblIvaSubt.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.lblIvaSubt.AutoSize = true;
			this.lblIvaSubt.Location = new System.Drawing.Point(696, 248);
			this.lblIvaSubt.Name = "lblIvaSubt";
			this.lblIvaSubt.Size = new System.Drawing.Size(22, 16);
			this.lblIvaSubt.TabIndex = 15;
			this.lblIvaSubt.Text = "IVA";
			// 
			// lblIva
			// 
			this.lblIva.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.lblIva.AutoSize = true;
			this.lblIva.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblIva.Location = new System.Drawing.Point(632, 312);
			this.lblIva.Name = "lblIva";
			this.lblIva.Size = new System.Drawing.Size(22, 16);
			this.lblIva.TabIndex = 17;
			this.lblIva.Text = "IVA";
			// 
			// lblDescuento
			// 
			this.lblDescuento.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.lblDescuento.AutoSize = true;
			this.lblDescuento.Location = new System.Drawing.Point(504, 288);
			this.lblDescuento.Name = "lblDescuento";
			this.lblDescuento.Size = new System.Drawing.Size(58, 16);
			this.lblDescuento.TabIndex = 16;
			this.lblDescuento.Text = "Descuento";
			// 
			// imageList
			// 
			this.imageList.ImageSize = new System.Drawing.Size(16, 16);
			this.imageList.ImageStream = ((System.Windows.Forms.ImageListStreamer)(resources.GetObject("imageList.ImageStream")));
			this.imageList.TransparentColor = System.Drawing.Color.Transparent;
			// 
			// errorProvider
			// 
			this.errorProvider.ContainerControl = this;
			this.errorProvider.DataMember = "";
			// 
			// btTotal
			// 
			this.btTotal.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.btTotal.ButtonStyle = DevExpress.XtraEditors.Controls.BorderStyles.Office2003;
			this.btTotal.Location = new System.Drawing.Point(504, 368);
			this.btTotal.Name = "btTotal";
			this.btTotal.Size = new System.Drawing.Size(40, 23);
			this.btTotal.TabIndex = 20;
			this.btTotal.Text = "Total";
			this.btTotal.Click += new System.EventHandler(this.btTotal_Click);
			// 
			// label6
			// 
			this.label6.AutoSize = true;
			this.label6.Location = new System.Drawing.Point(240, 32);
			this.label6.Name = "label6";
			this.label6.Size = new System.Drawing.Size(36, 16);
			this.label6.TabIndex = 30;
			this.label6.Text = "Fecha";
			// 
			// label7
			// 
			this.label7.AutoSize = true;
			this.label7.Location = new System.Drawing.Point(408, 32);
			this.label7.Name = "label7";
			this.label7.Size = new System.Drawing.Size(27, 16);
			this.label7.TabIndex = 31;
			this.label7.Text = "Días";
			// 
			// label8
			// 
			this.label8.AutoSize = true;
			this.label8.Location = new System.Drawing.Point(512, 32);
			this.label8.Name = "label8";
			this.label8.Size = new System.Drawing.Size(39, 16);
			this.label8.TabIndex = 32;
			this.label8.Text = "Vence:";
			// 
			// lblBodega
			// 
			this.lblBodega.AutoSize = true;
			this.lblBodega.Location = new System.Drawing.Point(568, 8);
			this.lblBodega.Name = "lblBodega";
			this.lblBodega.Size = new System.Drawing.Size(43, 16);
			this.lblBodega.TabIndex = 33;
			this.lblBodega.Text = "Bodega";
			this.lblBodega.TextAlign = System.Drawing.ContentAlignment.TopRight;
			// 
			// lblCliente
			// 
			this.lblCliente.AutoSize = true;
			this.lblCliente.Location = new System.Drawing.Point(240, 8);
			this.lblCliente.Name = "lblCliente";
			this.lblCliente.Size = new System.Drawing.Size(43, 16);
			this.lblCliente.TabIndex = 36;
			this.lblCliente.Text = "Cliente:";
			// 
			// lblSubTotal
			// 
			this.lblSubTotal.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.lblSubTotal.AutoSize = true;
			this.lblSubTotal.Location = new System.Drawing.Point(504, 264);
			this.lblSubTotal.Name = "lblSubTotal";
			this.lblSubTotal.Size = new System.Drawing.Size(50, 16);
			this.lblSubTotal.TabIndex = 37;
			this.lblSubTotal.Text = "SubTotal";
			this.lblSubTotal.Click += new System.EventHandler(this.label10_Click);
			// 
			// lblIce
			// 
			this.lblIce.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.lblIce.AutoSize = true;
			this.lblIce.Location = new System.Drawing.Point(504, 312);
			this.lblIce.Name = "lblIce";
			this.lblIce.Size = new System.Drawing.Size(23, 16);
			this.lblIce.TabIndex = 38;
			this.lblIce.Text = "ICE";
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
			// lblAnulado
			// 
			this.lblAnulado.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.lblAnulado.Font = new System.Drawing.Font("Microsoft Sans Serif", 15F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblAnulado.ForeColor = System.Drawing.Color.Firebrick;
			this.lblAnulado.Location = new System.Drawing.Point(624, 400);
			this.lblAnulado.Name = "lblAnulado";
			this.lblAnulado.Size = new System.Drawing.Size(112, 23);
			this.lblAnulado.TabIndex = 41;
			this.lblAnulado.Text = "ANULADO";
			// 
			// CmbDescCod
			// 
			this.CmbDescCod.Cursor = System.Windows.Forms.Cursors.Default;
			this.CmbDescCod.DataSource = this.cdvArticulo;
			ultraGridColumn140.Header.VisiblePosition = 2;
			ultraGridColumn140.Hidden = true;
			ultraGridColumn141.Header.VisiblePosition = 3;
			ultraGridColumn141.Hidden = true;
			ultraGridColumn142.Header.VisiblePosition = 5;
			ultraGridColumn142.Hidden = true;
			ultraGridColumn143.Header.VisiblePosition = 6;
			ultraGridColumn143.Hidden = true;
			ultraGridColumn144.Header.VisiblePosition = 4;
			ultraGridColumn144.Hidden = true;
			ultraGridColumn145.Header.VisiblePosition = 0;
			ultraGridColumn145.Width = 244;
			ultraGridColumn146.Header.VisiblePosition = 1;
			ultraGridColumn146.Width = 112;
			ultraGridColumn147.Header.VisiblePosition = 7;
			ultraGridColumn147.Hidden = true;
			ultraGridBand10.Columns.AddRange(new object[] {
																											ultraGridColumn140,
																											ultraGridColumn141,
																											ultraGridColumn142,
																											ultraGridColumn143,
																											ultraGridColumn144,
																											ultraGridColumn145,
																											ultraGridColumn146,
																											ultraGridColumn147});
			this.CmbDescCod.DisplayLayout.BandsSerializer.Add(ultraGridBand10);
			this.CmbDescCod.DisplayLayout.Override.RowFilterAction = Infragistics.Win.UltraWinGrid.RowFilterAction.HideFilteredOutRows;
			this.CmbDescCod.DisplayLayout.Override.RowSelectors = Infragistics.Win.DefaultableBoolean.False;
			this.CmbDescCod.DisplayMember = "Articulo";
			this.CmbDescCod.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.CmbDescCod.Location = new System.Drawing.Point(184, 112);
			this.CmbDescCod.Name = "CmbDescCod";
			this.CmbDescCod.Size = new System.Drawing.Size(368, 64);
			this.CmbDescCod.TabIndex = 53;
			this.CmbDescCod.ValueMember = "idArticulo";
			this.CmbDescCod.Visible = false;
			// 
			// ultraGrid1
			// 
			this.ultraGrid1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.ultraGrid1.Cursor = System.Windows.Forms.Cursors.Default;
			this.ultraGrid1.DataMember = "_Compra.Compra - DetCompra";
			this.ultraGrid1.DataSource = this.cdsCompra;
			appearance61.BackColor = System.Drawing.Color.White;
			this.ultraGrid1.DisplayLayout.Appearance = appearance61;
			ultraGridColumn148.Header.VisiblePosition = 24;
			ultraGridColumn148.Hidden = true;
			ultraGridColumn149.Header.VisiblePosition = 23;
			ultraGridColumn149.Hidden = true;
			ultraGridColumn150.Header.Caption = "Artículo";
			ultraGridColumn150.Header.VisiblePosition = 1;
			ultraGridColumn150.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownValidate;
			ultraGridColumn150.Width = 216;
			ultraGridColumn151.Header.Caption = "Unidad";
			ultraGridColumn151.Header.VisiblePosition = 6;
			ultraGridColumn151.Hidden = true;
			ultraGridColumn151.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownValidate;
			ultraGridColumn151.Width = 53;
			ultraGridColumn152.Header.VisiblePosition = 31;
			ultraGridColumn152.Hidden = true;
			ultraGridColumn153.Header.VisiblePosition = 27;
			ultraGridColumn153.Hidden = true;
			appearance62.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn154.CellAppearance = appearance62;
			ultraGridColumn154.Header.VisiblePosition = 0;
			ultraGridColumn154.MaskInput = "{LOC}nn,nnn,nnn.nn";
			ultraGridColumn154.NullText = "0";
			ultraGridColumn154.PromptChar = ' ';
			appearance63.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn155.CellAppearance = appearance63;
			ultraGridColumn155.Format = "";
			ultraGridColumn155.Header.VisiblePosition = 8;
			ultraGridColumn155.MaskInput = "{LOC}nn,nnn,nnn.nn";
			ultraGridColumn155.PromptChar = ' ';
			ultraGridColumn155.Width = 89;
			appearance64.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn156.CellAppearance = appearance64;
			ultraGridColumn156.Header.Caption = "Iva";
			ultraGridColumn156.Header.VisiblePosition = 9;
			ultraGridColumn156.PromptChar = ' ';
			ultraGridColumn156.Width = 34;
			ultraGridColumn157.CellMultiLine = Infragistics.Win.DefaultableBoolean.True;
			ultraGridColumn157.Header.VisiblePosition = 21;
			ultraGridColumn158.Header.Caption = "Ref. Codigo";
			ultraGridColumn158.Header.VisiblePosition = 19;
			ultraGridColumn159.Header.Caption = "Ref. Numero";
			ultraGridColumn159.Header.VisiblePosition = 20;
			appearance65.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn160.CellAppearance = appearance65;
			ultraGridColumn160.Header.Caption = "Descuento";
			ultraGridColumn160.Header.VisiblePosition = 11;
			ultraGridColumn160.MaskInput = "{LOC}nn,nnn,nnn.nn";
			ultraGridColumn160.PromptChar = ' ';
			appearance66.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn161.CellAppearance = appearance66;
			ultraGridColumn161.Header.Caption = "Peso";
			ultraGridColumn161.Header.VisiblePosition = 18;
			appearance67.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn162.CellAppearance = appearance67;
			ultraGridColumn162.Header.VisiblePosition = 26;
			ultraGridColumn162.Hidden = true;
			ultraGridColumn163.Header.VisiblePosition = 17;
			appearance68.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn164.CellAppearance = appearance68;
			ultraGridColumn164.Header.Caption = "% Desc.";
			ultraGridColumn164.Header.VisiblePosition = 12;
			ultraGridColumn164.MaskInput = "{LOC}nnnnnnn.nn";
			ultraGridColumn164.PromptChar = ' ';
			appearance69.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn165.CellAppearance = appearance69;
			ultraGridColumn165.Header.VisiblePosition = 22;
			ultraGridColumn165.Hidden = true;
			ultraGridColumn166.Header.VisiblePosition = 25;
			ultraGridColumn166.Hidden = true;
			ultraGridColumn167.Header.VisiblePosition = 36;
			ultraGridColumn167.Hidden = true;
			ultraGridColumn168.Header.VisiblePosition = 28;
			ultraGridColumn168.Hidden = true;
			ultraGridColumn169.Header.VisiblePosition = 29;
			ultraGridColumn169.Hidden = true;
			ultraGridColumn170.Header.VisiblePosition = 14;
			ultraGridColumn170.Hidden = true;
			ultraGridColumn171.Header.VisiblePosition = 30;
			ultraGridColumn171.Hidden = true;
			ultraGridColumn172.Header.Caption = "Centro Costo";
			ultraGridColumn172.Header.VisiblePosition = 32;
			ultraGridColumn172.Hidden = true;
			ultraGridColumn172.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownValidate;
			ultraGridColumn173.Header.Caption = "Proyecto";
			ultraGridColumn173.Header.VisiblePosition = 16;
			ultraGridColumn173.Hidden = true;
			ultraGridColumn173.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownValidate;
			ultraGridColumn174.Header.VisiblePosition = 5;
			ultraGridColumn174.Hidden = true;
			ultraGridColumn175.Header.VisiblePosition = 2;
			ultraGridColumn175.Hidden = true;
			ultraGridColumn175.Width = 118;
			ultraGridColumn176.Header.VisiblePosition = 33;
			ultraGridColumn176.Hidden = true;
			ultraGridColumn177.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance70.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn177.CellAppearance = appearance70;
			ultraGridColumn177.Format = "#,##0.00";
			ultraGridColumn177.Header.VisiblePosition = 10;
			ultraGridColumn178.Header.Caption = "Importacion";
			ultraGridColumn178.Header.VisiblePosition = 34;
			ultraGridColumn178.Hidden = true;
			ultraGridColumn178.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownValidate;
			ultraGridColumn179.Header.VisiblePosition = 35;
			ultraGridColumn179.Hidden = true;
			ultraGridColumn180.Header.VisiblePosition = 4;
			ultraGridColumn180.Hidden = true;
			ultraGridColumn181.Header.Caption = "CentroCosto";
			ultraGridColumn181.Header.VisiblePosition = 13;
			ultraGridColumn181.Hidden = true;
			ultraGridColumn182.Header.Caption = "Proyecto";
			ultraGridColumn182.Header.VisiblePosition = 15;
			ultraGridColumn182.Hidden = true;
			ultraGridColumn183.Header.Caption = "Código";
			ultraGridColumn183.Header.VisiblePosition = 3;
			ultraGridColumn183.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownValidate;
			ultraGridColumn184.Header.Caption = "Unidades";
			ultraGridColumn184.Header.VisiblePosition = 7;
			ultraGridColumn184.Hidden = true;
			ultraGridColumn184.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.Edit;
			ultraGridBand11.Columns.AddRange(new object[] {
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
																											ultraGridColumn184});
			this.ultraGrid1.DisplayLayout.BandsSerializer.Add(ultraGridBand11);
			this.ultraGrid1.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.TemplateOnBottom;
			this.ultraGrid1.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.True;
			this.ultraGrid1.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance71.BackColor = System.Drawing.Color.Transparent;
			this.ultraGrid1.DisplayLayout.Override.CardAreaAppearance = appearance71;
			appearance72.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance72.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance72.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance72.FontData.BoldAsString = "True";
			appearance72.FontData.Name = "Arial";
			appearance72.FontData.SizeInPoints = 10F;
			appearance72.ForeColor = System.Drawing.Color.White;
			appearance72.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.ultraGrid1.DisplayLayout.Override.HeaderAppearance = appearance72;
			this.ultraGrid1.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortSingle;
			appearance73.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance73.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance73.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.RowAlternateAppearance = appearance73;
			appearance74.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance74.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance74.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.RowSelectorAppearance = appearance74;
			appearance75.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance75.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance75.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.SelectedRowAppearance = appearance75;
			this.ultraGrid1.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.ultraGrid1.Location = new System.Drawing.Point(8, 64);
			this.ultraGrid1.Name = "ultraGrid1";
			this.ultraGrid1.Size = new System.Drawing.Size(720, 184);
			this.ultraGrid1.TabIndex = 8;
			this.ultraGrid1.InitializeRow += new Infragistics.Win.UltraWinGrid.InitializeRowEventHandler(this.ultraGrid1_InitializeRow);
			this.ultraGrid1.Error += new Infragistics.Win.UltraWinGrid.ErrorEventHandler(this.ultraGrid1_Error);
			this.ultraGrid1.AfterRowInsert += new Infragistics.Win.UltraWinGrid.RowEventHandler(this.ultraGrid1_AfterRowInsert);
			this.ultraGrid1.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.ultraGrid1_KeyPress);
			this.ultraGrid1.BeforeRowsDeleted += new Infragistics.Win.UltraWinGrid.BeforeRowsDeletedEventHandler(this.ultraGrid1_BeforeRowsDeleted);
			this.ultraGrid1.BeforeRowInsert += new Infragistics.Win.UltraWinGrid.BeforeRowInsertEventHandler(this.ultraGrid1_BeforeRowInsert);
			this.ultraGrid1.KeyDown += new System.Windows.Forms.KeyEventHandler(this.ultraGrid1_KeyDown);
			this.ultraGrid1.AfterCellUpdate += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.ultraGrid1_AfterCellUpdate);
			this.ultraGrid1.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.ultraGrid1_InitializeLayout);
			// 
			// CmbCodDesc
			// 
			this.CmbCodDesc.Cursor = System.Windows.Forms.Cursors.Default;
			this.CmbCodDesc.DataSource = this.cdvArtCodigo;
			ultraGridColumn185.Header.VisiblePosition = 2;
			ultraGridColumn185.Hidden = true;
			ultraGridColumn186.Header.VisiblePosition = 3;
			ultraGridColumn186.Hidden = true;
			ultraGridColumn187.Header.VisiblePosition = 5;
			ultraGridColumn187.Hidden = true;
			ultraGridColumn188.Header.VisiblePosition = 6;
			ultraGridColumn188.Hidden = true;
			ultraGridColumn189.Header.VisiblePosition = 4;
			ultraGridColumn189.Hidden = true;
			ultraGridColumn190.Header.VisiblePosition = 1;
			ultraGridColumn190.Width = 270;
			ultraGridColumn191.Header.VisiblePosition = 0;
			ultraGridColumn191.Width = 134;
			ultraGridColumn192.Header.VisiblePosition = 7;
			ultraGridColumn192.Hidden = true;
			ultraGridBand12.Columns.AddRange(new object[] {
																											ultraGridColumn185,
																											ultraGridColumn186,
																											ultraGridColumn187,
																											ultraGridColumn188,
																											ultraGridColumn189,
																											ultraGridColumn190,
																											ultraGridColumn191,
																											ultraGridColumn192});
			this.CmbCodDesc.DisplayLayout.BandsSerializer.Add(ultraGridBand12);
			this.CmbCodDesc.DisplayLayout.Override.RowFilterMode = Infragistics.Win.UltraWinGrid.RowFilterMode.SiblingRowsOnly;
			this.CmbCodDesc.DisplayLayout.Override.RowSizing = Infragistics.Win.UltraWinGrid.RowSizing.AutoFree;
			this.CmbCodDesc.DisplayMember = "Codigo";
			this.CmbCodDesc.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.CmbCodDesc.Location = new System.Drawing.Point(80, 152);
			this.CmbCodDesc.Name = "CmbCodDesc";
			this.CmbCodDesc.Size = new System.Drawing.Size(392, 80);
			this.CmbCodDesc.TabIndex = 54;
			this.CmbCodDesc.ValueMember = "idArticulo";
			this.CmbCodDesc.Visible = false;
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
			// cdvCompra
			// 
			this.cdvCompra.BindingContextCtrl = this;
			this.cdvCompra.DataSet = this.cdsCompra;
			this.cdvCompra.Sort = "idCompra";
			this.cdvCompra.TableName = "";
			this.cdvCompra.TableViewName = "Compra";
			// 
			// cdsAsiento
			// 
			this.cdsAsiento.BindingContextCtrl = this;
			this.cdsAsiento.DataLibrary = "LibContabilidad";
			this.cdsAsiento.DataLibraryUrl = "";
			this.cdsAsiento.DataSetDef = "dsAsiento";
			this.cdsAsiento.FillOnRequest = false;
			this.cdsAsiento.Locale = new System.Globalization.CultureInfo("es-MX");
			this.cdsAsiento.Name = "cdsAsiento";
			this.cdsAsiento.SchemaClassName = "LibContabilidad.DataClass";
			this.cdsAsiento.SchemaDef = null;
			this.cdsAsiento.BeforeFill += new C1.Data.FillEventHandler(this.cdsAsiento_BeforeFill);
			// 
			// cdsRetencion
			// 
			this.cdsRetencion.BindingContextCtrl = this;
			this.cdsRetencion.DataLibrary = "LibContabilidad";
			this.cdsRetencion.DataLibraryUrl = "";
			this.cdsRetencion.DataSetDef = "dsRetencion";
			this.cdsRetencion.FillOnRequest = false;
			this.cdsRetencion.Locale = new System.Globalization.CultureInfo("es-MX");
			this.cdsRetencion.Name = "cdsRetencion";
			this.cdsRetencion.SchemaClassName = "LibContabilidad.DataClass";
			this.cdsRetencion.SchemaDef = null;
			this.cdsRetencion.BeforeFill += new C1.Data.FillEventHandler(this.cdsRetencion_BeforeFill);
			// 
			// cdsAsignaCuenta
			// 
			this.cdsAsignaCuenta.BindingContextCtrl = this;
			this.cdsAsignaCuenta.DataLibrary = "LibContabilidad";
			this.cdsAsignaCuenta.DataLibraryUrl = "";
			this.cdsAsignaCuenta.DataSetDef = "dsAsignaCuenta";
			this.cdsAsignaCuenta.Locale = new System.Globalization.CultureInfo("es-MX");
			this.cdsAsignaCuenta.Name = "cdsAsignaCuenta";
			this.cdsAsignaCuenta.SchemaClassName = "LibContabilidad.DataClass";
			this.cdsAsignaCuenta.SchemaDef = null;
			this.cdsAsignaCuenta.BeforeFill += new C1.Data.FillEventHandler(this.cdsAsignaCuenta_BeforeFill);
			// 
			// cmbCliente
			// 
			appearance76.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbCliente.Appearance = appearance76;
			this.cmbCliente.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbCliente.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbCliente.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsCompra, "_Compra.idCliente"));
			this.cmbCliente.DataSource = this.cdvCliente;
			ultraGridColumn193.Header.VisiblePosition = 1;
			ultraGridColumn194.Header.VisiblePosition = 3;
			ultraGridColumn194.Hidden = true;
			ultraGridColumn195.Header.VisiblePosition = 0;
			ultraGridColumn195.Width = 231;
			ultraGridColumn196.Header.VisiblePosition = 4;
			ultraGridColumn196.Hidden = true;
			ultraGridColumn197.Header.VisiblePosition = 5;
			ultraGridColumn197.Hidden = true;
			ultraGridColumn198.Header.VisiblePosition = 2;
			ultraGridColumn198.Hidden = true;
			ultraGridColumn199.Header.VisiblePosition = 6;
			ultraGridColumn199.Hidden = true;
			ultraGridColumn200.Header.VisiblePosition = 7;
			ultraGridColumn200.Hidden = true;
			ultraGridColumn201.Header.VisiblePosition = 8;
			ultraGridColumn201.Hidden = true;
			ultraGridColumn202.Header.VisiblePosition = 9;
			ultraGridColumn202.Hidden = true;
			ultraGridBand13.Columns.AddRange(new object[] {
																											ultraGridColumn193,
																											ultraGridColumn194,
																											ultraGridColumn195,
																											ultraGridColumn196,
																											ultraGridColumn197,
																											ultraGridColumn198,
																											ultraGridColumn199,
																											ultraGridColumn200,
																											ultraGridColumn201,
																											ultraGridColumn202});
			this.cmbCliente.DisplayLayout.BandsSerializer.Add(ultraGridBand13);
			this.cmbCliente.DisplayMember = "Nombre";
			this.cmbCliente.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbCliente.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbCliente.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbCliente.Location = new System.Drawing.Point(288, 8);
			this.cmbCliente.Name = "cmbCliente";
			this.cmbCliente.Size = new System.Drawing.Size(248, 21);
			this.cmbCliente.TabIndex = 2;
			this.cmbCliente.ValueMember = "idCliente";
			this.cmbCliente.Validated += new System.EventHandler(this.cmbCliente_Validated);
			this.cmbCliente.MouseUp += new System.Windows.Forms.MouseEventHandler(this.cmbCliente_MouseUp);
			// 
			// txtNumero
			// 
			appearance77.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtNumero.Appearance = appearance77;
			this.txtNumero.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsCompra, "_Compra.Numero"));
			this.txtNumero.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNumero.Location = new System.Drawing.Point(120, 8);
			this.txtNumero.Name = "txtNumero";
			this.txtNumero.Size = new System.Drawing.Size(112, 21);
			this.txtNumero.TabIndex = 0;
			// 
			// dtFecha
			// 
			appearance78.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtFecha.Appearance = appearance78;
			this.dtFecha.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.dtFecha.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsCompra, "_Compra.Fecha"));
			dateButton3.Caption = "Today";
			this.dtFecha.DateButtons.Add(dateButton3);
			this.dtFecha.Format = "dd/MMM/yyyy";
			this.dtFecha.Location = new System.Drawing.Point(288, 32);
			this.dtFecha.Name = "dtFecha";
			this.dtFecha.NonAutoSizeHeight = 23;
			this.dtFecha.Size = new System.Drawing.Size(112, 21);
			this.dtFecha.SpinButtonsVisible = true;
			this.dtFecha.TabIndex = 5;
			this.dtFecha.Value = new System.DateTime(2005, 6, 8, 0, 0, 0, 0);
			// 
			// spnDias
			// 
			appearance79.ForeColorDisabled = System.Drawing.Color.Black;
			this.spnDias.Appearance = appearance79;
			this.spnDias.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.spnDias.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsCompra, "_Compra.DiasCredito"));
			this.spnDias.Location = new System.Drawing.Point(440, 32);
			this.spnDias.MaskInput = "nnnn";
			this.spnDias.Name = "spnDias";
			this.spnDias.PromptChar = ' ';
			this.spnDias.Size = new System.Drawing.Size(64, 21);
			this.spnDias.SpinButtonDisplayStyle = Infragistics.Win.ButtonDisplayStyle.Always;
			this.spnDias.TabIndex = 6;
			// 
			// dtVence
			// 
			appearance80.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtVence.Appearance = appearance80;
			this.dtVence.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.dtVence.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsCompra, "_Compra.FechaVencimiento"));
			dateButton4.Caption = "Today";
			this.dtVence.DateButtons.Add(dateButton4);
			this.dtVence.Format = "dd/MMM/yyyy";
			this.dtVence.Location = new System.Drawing.Point(560, 32);
			this.dtVence.Name = "dtVence";
			this.dtVence.NonAutoSizeHeight = 23;
			this.dtVence.Size = new System.Drawing.Size(112, 21);
			this.dtVence.SpinButtonsVisible = true;
			this.dtVence.TabIndex = 7;
			this.dtVence.Value = new System.DateTime(2005, 6, 8, 0, 0, 0, 0);
			// 
			// spnBodega
			// 
			appearance81.ForeColorDisabled = System.Drawing.Color.Black;
			this.spnBodega.Appearance = appearance81;
			this.spnBodega.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.spnBodega.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsCompra, "_Compra.Bodega"));
			this.spnBodega.Location = new System.Drawing.Point(616, 8);
			this.spnBodega.MaskInput = "nnn";
			this.spnBodega.MinValue = 0;
			this.spnBodega.Name = "spnBodega";
			this.spnBodega.PromptChar = ' ';
			this.spnBodega.Size = new System.Drawing.Size(72, 21);
			this.spnBodega.SpinButtonDisplayStyle = Infragistics.Win.ButtonDisplayStyle.Always;
			this.spnBodega.TabIndex = 3;
			// 
			// btCliente
			// 
			this.btCliente.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.btCliente.Location = new System.Drawing.Point(536, 8);
			this.btCliente.Name = "btCliente";
			this.btCliente.Size = new System.Drawing.Size(24, 23);
			this.btCliente.TabIndex = 76;
			this.btCliente.Text = "...";
			this.btCliente.Click += new System.EventHandler(this.btCliente_Click);
			this.btCliente.MouseUp += new System.Windows.Forms.MouseEventHandler(this.btCliente_MouseDown);
			// 
			// cmbComprobante1
			// 
			appearance82.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbComprobante1.Appearance = appearance82;
			this.cmbComprobante1.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbComprobante1.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbComprobante1.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsCompra, "_Compra.idComprobante"));
			this.cmbComprobante1.DataSource = this.cdvComprobante;
			ultraGridBand14.AddButtonCaption = "Cliente";
			ultraGridColumn203.Header.VisiblePosition = 2;
			ultraGridColumn203.Hidden = true;
			ultraGridColumn204.Header.VisiblePosition = 1;
			ultraGridColumn204.Width = 62;
			ultraGridColumn205.Header.VisiblePosition = 0;
			ultraGridColumn205.Width = 299;
			ultraGridColumn206.Header.VisiblePosition = 3;
			ultraGridColumn206.Hidden = true;
			ultraGridColumn207.Header.VisiblePosition = 4;
			ultraGridColumn207.Hidden = true;
			ultraGridBand14.Columns.AddRange(new object[] {
																											ultraGridColumn203,
																											ultraGridColumn204,
																											ultraGridColumn205,
																											ultraGridColumn206,
																											ultraGridColumn207});
			this.cmbComprobante1.DisplayLayout.BandsSerializer.Add(ultraGridBand14);
			this.cmbComprobante1.DisplayMember = "Comprobante";
			this.cmbComprobante1.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbComprobante1.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbComprobante1.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbComprobante1.Location = new System.Drawing.Point(8, 32);
			this.cmbComprobante1.Name = "cmbComprobante1";
			this.cmbComprobante1.Size = new System.Drawing.Size(224, 21);
			this.cmbComprobante1.TabIndex = 4;
			this.cmbComprobante1.ValueMember = "idComprobante";
			this.cmbComprobante1.Validated += new System.EventHandler(this.cmbComprobante1_Validated);
			this.cmbComprobante1.MouseUp += new System.Windows.Forms.MouseEventHandler(this.cmbComprobante1_MouseUp);
			// 
			// cmbTipoCompra
			// 
			appearance83.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbTipoCompra.Appearance = appearance83;
			this.cmbTipoCompra.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbTipoCompra.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbTipoCompra.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsCompra, "_Compra.idTipoFactura"));
			this.cmbTipoCompra.DataSource = this.cdvNumero;
			ultraGridBand15.AddButtonCaption = "Cliente";
			ultraGridColumn208.Header.VisiblePosition = 7;
			ultraGridColumn208.Hidden = true;
			ultraGridColumn209.Header.VisiblePosition = 0;
			ultraGridColumn210.Header.VisiblePosition = 8;
			ultraGridColumn210.Hidden = true;
			ultraGridColumn211.Header.VisiblePosition = 3;
			ultraGridColumn211.Hidden = true;
			ultraGridColumn212.Header.VisiblePosition = 2;
			ultraGridColumn212.Hidden = true;
			ultraGridColumn213.Header.VisiblePosition = 1;
			ultraGridColumn213.Hidden = true;
			ultraGridColumn214.Header.VisiblePosition = 5;
			ultraGridColumn214.Hidden = true;
			ultraGridColumn215.Header.VisiblePosition = 6;
			ultraGridColumn215.Hidden = true;
			ultraGridColumn216.Header.VisiblePosition = 4;
			ultraGridColumn216.Hidden = true;
			ultraGridColumn217.Header.VisiblePosition = 9;
			ultraGridColumn218.Header.VisiblePosition = 10;
			ultraGridColumn219.Header.VisiblePosition = 11;
			ultraGridColumn220.Header.VisiblePosition = 12;
			ultraGridColumn221.Header.VisiblePosition = 13;
			ultraGridColumn222.Header.VisiblePosition = 14;
			ultraGridColumn223.Header.VisiblePosition = 15;
			ultraGridColumn224.Header.VisiblePosition = 16;
			ultraGridBand15.Columns.AddRange(new object[] {
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
																											ultraGridColumn224});
			this.cmbTipoCompra.DisplayLayout.BandsSerializer.Add(ultraGridBand15);
			this.cmbTipoCompra.DisplayMember = "Nombre";
			this.cmbTipoCompra.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbTipoCompra.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbTipoCompra.Enabled = false;
			this.cmbTipoCompra.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbTipoCompra.Location = new System.Drawing.Point(8, 8);
			this.cmbTipoCompra.Name = "cmbTipoCompra";
			this.cmbTipoCompra.ReadOnly = true;
			this.cmbTipoCompra.Size = new System.Drawing.Size(112, 21);
			this.cmbTipoCompra.TabIndex = 78;
			this.cmbTipoCompra.ValueMember = "idTipoFactura";
			// 
			// btSaldo1
			// 
			this.btSaldo1.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			appearance84.Image = 5;
			this.btSaldo1.Appearance = appearance84;
			this.btSaldo1.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.btSaldo1.ImageList = this.imageList;
			this.btSaldo1.Location = new System.Drawing.Point(328, 400);
			this.btSaldo1.Name = "btSaldo1";
			this.btSaldo1.Size = new System.Drawing.Size(24, 23);
			this.btSaldo1.TabIndex = 11;
			this.toolTip.SetToolTip(this.btSaldo1, "Saldo de Artículo / Kardex de Articulo");
			this.btSaldo1.Click += new System.EventHandler(this.btSaldo_Click);
			this.btSaldo1.MouseUp += new System.Windows.Forms.MouseEventHandler(this.btSaldo1_MouseUp);
			// 
			// btBuscaArticulo
			// 
			this.btBuscaArticulo.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			appearance85.Image = 6;
			this.btBuscaArticulo.Appearance = appearance85;
			this.btBuscaArticulo.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.btBuscaArticulo.ImageList = this.imageList;
			this.btBuscaArticulo.Location = new System.Drawing.Point(352, 400);
			this.btBuscaArticulo.Name = "btBuscaArticulo";
			this.btBuscaArticulo.Size = new System.Drawing.Size(24, 23);
			this.btBuscaArticulo.TabIndex = 12;
			this.toolTip.SetToolTip(this.btBuscaArticulo, "Busqueda de articulo seleccionado / Ultimos Movimientos");
			this.btBuscaArticulo.Click += new System.EventHandler(this.btBuscaArticulo_Click);
			this.btBuscaArticulo.MouseUp += new System.Windows.Forms.MouseEventHandler(this.btBuscaArticulo_MouseUp);
			// 
			// btTransforma
			// 
			this.btTransforma.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			appearance86.Image = 3;
			this.btTransforma.Appearance = appearance86;
			this.btTransforma.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.btTransforma.ImageList = this.imageList;
			this.btTransforma.Location = new System.Drawing.Point(376, 400);
			this.btTransforma.Name = "btTransforma";
			this.btTransforma.Size = new System.Drawing.Size(24, 23);
			this.btTransforma.TabIndex = 13;
			this.toolTip.SetToolTip(this.btTransforma, "Transformacion / Distribuye Porcentajes");
			this.btTransforma.Click += new System.EventHandler(this.btTransforma_Click);
			this.btTransforma.MouseUp += new System.Windows.Forms.MouseEventHandler(this.btTransforma_MouseUp);
			// 
			// btNotaEntrega
			// 
			this.btNotaEntrega.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			appearance87.Image = 7;
			this.btNotaEntrega.Appearance = appearance87;
			this.btNotaEntrega.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.btNotaEntrega.ImageList = this.imageList;
			this.btNotaEntrega.Location = new System.Drawing.Point(400, 400);
			this.btNotaEntrega.Name = "btNotaEntrega";
			this.btNotaEntrega.Size = new System.Drawing.Size(24, 23);
			this.btNotaEntrega.TabIndex = 14;
			this.toolTip.SetToolTip(this.btNotaEntrega, "Impresión de Nota de Entrega / Envio de Correo Electronico");
			this.btNotaEntrega.Click += new System.EventHandler(this.btNotaEntrega_Click);
			this.btNotaEntrega.MouseDown += new System.Windows.Forms.MouseEventHandler(this.btNotaEntrega_MouseDown);
			// 
			// btRetencion
			// 
			this.btRetencion.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			appearance88.Image = 0;
			this.btRetencion.Appearance = appearance88;
			this.btRetencion.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.btRetencion.ImageList = this.imageList;
			this.btRetencion.Location = new System.Drawing.Point(424, 400);
			this.btRetencion.Name = "btRetencion";
			this.btRetencion.Size = new System.Drawing.Size(24, 23);
			this.btRetencion.TabIndex = 15;
			this.toolTip.SetToolTip(this.btRetencion, "Retención en la fuente");
			this.btRetencion.Click += new System.EventHandler(this.btRetencion_Click);
			this.btRetencion.MouseUp += new System.Windows.Forms.MouseEventHandler(this.btRetencion_MouseDown);
			// 
			// btCreaAsiento
			// 
			this.btCreaAsiento.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			appearance89.Image = 9;
			this.btCreaAsiento.Appearance = appearance89;
			this.btCreaAsiento.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.btCreaAsiento.Location = new System.Drawing.Point(448, 400);
			this.btCreaAsiento.Name = "btCreaAsiento";
			this.btCreaAsiento.Size = new System.Drawing.Size(24, 23);
			this.btCreaAsiento.TabIndex = 16;
			this.btCreaAsiento.Text = "A";
			this.toolTip.SetToolTip(this.btCreaAsiento, "Crea Asiento");
			this.btCreaAsiento.Click += new System.EventHandler(this.btCreaAsiento_Click);
			this.btCreaAsiento.MouseUp += new System.Windows.Forms.MouseEventHandler(this.btCreaAsiento_MouseDown);
			// 
			// btGuiaRemision
			// 
			this.btGuiaRemision.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			appearance90.Image = 8;
			this.btGuiaRemision.Appearance = appearance90;
			this.btGuiaRemision.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.btGuiaRemision.ImageList = this.imageList;
			this.btGuiaRemision.Location = new System.Drawing.Point(472, 400);
			this.btGuiaRemision.Name = "btGuiaRemision";
			this.btGuiaRemision.Size = new System.Drawing.Size(24, 23);
			this.btGuiaRemision.TabIndex = 17;
			this.toolTip.SetToolTip(this.btGuiaRemision, "Guia de Remisión");
			this.btGuiaRemision.Visible = false;
			this.btGuiaRemision.Click += new System.EventHandler(this.btGuiaRemision_Click);
			this.btGuiaRemision.MouseUp += new System.Windows.Forms.MouseEventHandler(this.btGuiaRemision_MouseDown);
			// 
			// btActualizaPrecio
			// 
			this.btActualizaPrecio.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			appearance91.Image = 11;
			this.btActualizaPrecio.Appearance = appearance91;
			this.btActualizaPrecio.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.btActualizaPrecio.ImageList = this.imageList;
			this.btActualizaPrecio.Location = new System.Drawing.Point(496, 400);
			this.btActualizaPrecio.Name = "btActualizaPrecio";
			this.btActualizaPrecio.Size = new System.Drawing.Size(24, 23);
			this.btActualizaPrecio.TabIndex = 18;
			this.toolTip.SetToolTip(this.btActualizaPrecio, "Actualiza Precios");
			this.btActualizaPrecio.Visible = false;
			this.btActualizaPrecio.Click += new System.EventHandler(this.btActualizaPrecio_Click);
			this.btActualizaPrecio.MouseDown += new System.Windows.Forms.MouseEventHandler(this.btActualizaPrecio_MouseDown);
			// 
			// btNuevo
			// 
			this.btNuevo.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btNuevo.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.btNuevo.ImageList = this.imageList;
			this.btNuevo.Location = new System.Drawing.Point(520, 400);
			this.btNuevo.Name = "btNuevo";
			this.btNuevo.Size = new System.Drawing.Size(24, 23);
			this.btNuevo.TabIndex = 19;
			this.btNuevo.Text = "&N";
			this.toolTip.SetToolTip(this.btNuevo, "Nueva Linea");
			this.btNuevo.Click += new System.EventHandler(this.btNuevo_Click);
			// 
			// txtIdAsiento
			// 
			this.txtIdAsiento.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.txtIdAsiento.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsCompra, "_Compra.idAsiento"));
			this.txtIdAsiento.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtIdAsiento.Location = new System.Drawing.Point(512, 384);
			this.txtIdAsiento.Name = "txtIdAsiento";
			this.txtIdAsiento.PromptChar = ' ';
			this.txtIdAsiento.Size = new System.Drawing.Size(48, 21);
			this.txtIdAsiento.TabIndex = 88;
			// 
			// txtIdCompra
			// 
			this.txtIdCompra.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.txtIdCompra.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsCompra, "_Compra.idCompra"));
			this.txtIdCompra.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtIdCompra.Location = new System.Drawing.Point(568, 376);
			this.txtIdCompra.Name = "txtIdCompra";
			this.txtIdCompra.PromptChar = ' ';
			this.txtIdCompra.Size = new System.Drawing.Size(56, 21);
			this.txtIdCompra.TabIndex = 89;
			// 
			// txtBorrar
			// 
			this.txtBorrar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.txtBorrar.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsCompra, "_Compra.Borrar"));
			this.txtBorrar.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtBorrar.Location = new System.Drawing.Point(624, 376);
			this.txtBorrar.Name = "txtBorrar";
			this.txtBorrar.Size = new System.Drawing.Size(48, 21);
			this.txtBorrar.TabIndex = 90;
			// 
			// barraDato1
			// 
			this.barraDato1.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.barraDato1.BarraMovimiento = true;
			this.barraDato1.DataMember = "_Compra";
			this.barraDato1.DataNombreId = "idCompra";
			this.barraDato1.DataOrden = "";
			this.barraDato1.DataSource = this.cdsCompra;
			this.barraDato1.DataTabla = "Compra";
			this.barraDato1.DataTablaHija = "DetCompra";
			this.barraDato1.Location = new System.Drawing.Point(8, 400);
			this.barraDato1.Name = "barraDato1";
			this.barraDato1.Size = new System.Drawing.Size(304, 24);
			this.barraDato1.TabIndex = 10;
			this.barraDato1.VisibleBorrar = true;
			this.barraDato1.VisibleBuscar = true;
			this.barraDato1.VisibleEditar = true;
			this.barraDato1.VisibleImprimir = true;
			this.barraDato1.VisibleNuevo = true;
			this.barraDato1.Buscar += new System.EventHandler(this.barraDato1_Buscar);
			this.barraDato1.Borrar += new System.EventHandler(this.barraDato1_Borrar);
			this.barraDato1.Imprimir += new System.EventHandler(this.barraDato1_Imprimir);
			this.barraDato1.Restaurar += new System.EventHandler(this.barraDato1_Restaurar);
			this.barraDato1.Refresca += new System.EventHandler(this.barraDato1_Refresca);
			this.barraDato1.Crear += new System.EventHandler(this.barraDato1_Crear);
			this.barraDato1.Deshacer += new System.EventHandler(this.barraDato1_Deshacer);
			this.barraDato1.Copiar += new System.EventHandler(this.barraDato1_Copiar);
			this.barraDato1.Impresora += new System.EventHandler(this.barraDato1_Impresora);
			this.barraDato1.Editar += new System.EventHandler(this.barraDato1_Editar);
			this.barraDato1.Grabar += new System.EventHandler(this.barraDato1_Grabar);
			// 
			// utcTransaccion
			// 
			this.utcTransaccion.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			appearance92.ForeColorDisabled = System.Drawing.Color.Black;
			this.utcTransaccion.Appearance = appearance92;
			this.utcTransaccion.Controls.Add(this.ultraTabSharedControlsPage1);
			this.utcTransaccion.Controls.Add(this.ultraTabPageControl1);
			this.utcTransaccion.Controls.Add(this.ultraTabPageControl2);
			this.utcTransaccion.Controls.Add(this.ultraTabPageControl3);
			this.utcTransaccion.Controls.Add(this.pgEntrega);
			this.utcTransaccion.Controls.Add(this.ultraTabPageControl5);
			this.utcTransaccion.Controls.Add(this.TabPageCOA);
			this.utcTransaccion.Controls.Add(this.tbpAutorizacion);
			this.utcTransaccion.Controls.Add(this.ultraTabPageControl6);
			this.utcTransaccion.Controls.Add(this.ultraTabPageControl4);
			this.utcTransaccion.Location = new System.Drawing.Point(8, 256);
			this.utcTransaccion.Name = "utcTransaccion";
			this.utcTransaccion.SharedControlsPage = this.ultraTabSharedControlsPage1;
			this.utcTransaccion.Size = new System.Drawing.Size(488, 136);
			this.utcTransaccion.Style = Infragistics.Win.UltraWinTabControl.UltraTabControlStyle.Excel;
			this.utcTransaccion.TabIndex = 9;
			ultraTab1.TabPage = this.ultraTabPageControl1;
			ultraTab1.Text = "Notas";
			ultraTab2.TabPage = this.ultraTabPageControl2;
			ultraTab2.Text = "Pago";
			ultraTab3.TabPage = this.ultraTabPageControl3;
			ultraTab3.Text = "Desc.";
			ultraTab4.TabPage = this.pgEntrega;
			ultraTab4.Text = "Entrega";
			ultraTab5.TabPage = this.ultraTabPageControl5;
			ultraTab5.Text = "Cuotas";
			ultraTab6.TabPage = this.TabPageCOA;
			ultraTab6.Text = "A.T.";
			ultraTab7.TabPage = this.tbpAutorizacion;
			ultraTab7.Text = "Autoriz.";
			ultraTab8.TabPage = this.ultraTabPageControl6;
			ultraTab8.Text = "Busca";
			ultraTab9.TabPage = this.ultraTabPageControl4;
			ultraTab9.Text = "Otro";
			this.utcTransaccion.Tabs.AddRange(new Infragistics.Win.UltraWinTabControl.UltraTab[] {
																																														 ultraTab1,
																																														 ultraTab2,
																																														 ultraTab3,
																																														 ultraTab4,
																																														 ultraTab5,
																																														 ultraTab6,
																																														 ultraTab7,
																																														 ultraTab8,
																																														 ultraTab9});
			this.utcTransaccion.ViewStyle = Infragistics.Win.UltraWinTabControl.ViewStyle.Office2003;
			// 
			// ultraTabSharedControlsPage1
			// 
			this.ultraTabSharedControlsPage1.Location = new System.Drawing.Point(-10000, -10000);
			this.ultraTabSharedControlsPage1.Name = "ultraTabSharedControlsPage1";
			this.ultraTabSharedControlsPage1.Size = new System.Drawing.Size(486, 115);
			// 
			// txtDescIva0Total
			// 
			this.txtDescIva0Total.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			appearance93.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtDescIva0Total.Appearance = appearance93;
			this.txtDescIva0Total.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsCompra, "_Compra.Descuento0"));
			this.txtDescIva0Total.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtDescIva0Total.FormatString = "#,##0.00";
			this.txtDescIva0Total.Location = new System.Drawing.Point(560, 288);
			this.txtDescIva0Total.Name = "txtDescIva0Total";
			this.txtDescIva0Total.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtDescIva0Total.PromptChar = ' ';
			this.txtDescIva0Total.ReadOnly = true;
			this.txtDescIva0Total.Size = new System.Drawing.Size(80, 21);
			this.txtDescIva0Total.TabIndex = 112;
			// 
			// txtDescIvaTotal
			// 
			this.txtDescIvaTotal.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			appearance94.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtDescIvaTotal.Appearance = appearance94;
			this.txtDescIvaTotal.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsCompra, "_Compra.Descuento"));
			this.txtDescIvaTotal.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtDescIvaTotal.FormatString = "#,##0.00";
			this.txtDescIvaTotal.Location = new System.Drawing.Point(640, 288);
			this.txtDescIvaTotal.Name = "txtDescIvaTotal";
			this.txtDescIvaTotal.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtDescIvaTotal.PromptChar = ' ';
			this.txtDescIvaTotal.ReadOnly = true;
			this.txtDescIvaTotal.Size = new System.Drawing.Size(88, 21);
			this.txtDescIvaTotal.TabIndex = 111;
			// 
			// txtIva
			// 
			this.txtIva.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			appearance95.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtIva.Appearance = appearance95;
			this.txtIva.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsCompra, "_Compra.SubtotalIva"));
			this.txtIva.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtIva.FormatString = "#,##0.00";
			this.txtIva.Location = new System.Drawing.Point(640, 264);
			this.txtIva.Name = "txtIva";
			this.txtIva.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtIva.PromptChar = ' ';
			this.txtIva.ReadOnly = true;
			this.txtIva.Size = new System.Drawing.Size(88, 21);
			this.txtIva.TabIndex = 110;
			// 
			// txtIva0
			// 
			this.txtIva0.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			appearance96.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtIva0.Appearance = appearance96;
			this.txtIva0.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsCompra, "_Compra.SubtotalExcento"));
			this.txtIva0.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtIva0.FormatString = "#,##0.00";
			this.txtIva0.Location = new System.Drawing.Point(560, 264);
			this.txtIva0.Name = "txtIva0";
			this.txtIva0.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtIva0.PromptChar = ' ';
			this.txtIva0.ReadOnly = true;
			this.txtIva0.Size = new System.Drawing.Size(80, 21);
			this.txtIva0.TabIndex = 109;
			// 
			// txtIceTotal
			// 
			this.txtIceTotal.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			appearance97.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtIceTotal.Appearance = appearance97;
			this.txtIceTotal.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsCompra, "_Compra.Ice"));
			this.txtIceTotal.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtIceTotal.FormatString = "#,##0.00";
			this.txtIceTotal.Location = new System.Drawing.Point(552, 312);
			this.txtIceTotal.Name = "txtIceTotal";
			this.txtIceTotal.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtIceTotal.PromptChar = ' ';
			this.txtIceTotal.Size = new System.Drawing.Size(64, 21);
			this.txtIceTotal.TabIndex = 115;
			// 
			// txtTotal
			// 
			this.txtTotal.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			appearance98.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtTotal.Appearance = appearance98;
			this.txtTotal.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsCompra, "_Compra.Total"));
			this.txtTotal.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtTotal.Font = new System.Drawing.Font("Microsoft Sans Serif", 15.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtTotal.FormatString = "#,##0.00";
			this.txtTotal.Location = new System.Drawing.Point(552, 360);
			this.txtTotal.Name = "txtTotal";
			this.txtTotal.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtTotal.PromptChar = ' ';
			this.txtTotal.ReadOnly = true;
			this.txtTotal.Size = new System.Drawing.Size(176, 34);
			this.txtTotal.TabIndex = 114;
			// 
			// txtIvaTotal
			// 
			this.txtIvaTotal.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			appearance99.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtIvaTotal.Appearance = appearance99;
			this.txtIvaTotal.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsCompra, "_Compra.Iva"));
			this.txtIvaTotal.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtIvaTotal.FormatString = "#,##0.00";
			this.txtIvaTotal.Location = new System.Drawing.Point(656, 312);
			this.txtIvaTotal.Name = "txtIvaTotal";
			this.txtIvaTotal.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtIvaTotal.PromptChar = ' ';
			this.txtIvaTotal.ReadOnly = true;
			this.txtIvaTotal.Size = new System.Drawing.Size(72, 21);
			this.txtIvaTotal.TabIndex = 113;
			// 
			// cdsSeteo
			// 
			this.cdsSeteo.BindingContextCtrl = this;
			this.cdsSeteo.DataLibrary = "LibContabilidad";
			this.cdsSeteo.DataLibraryUrl = "";
			this.cdsSeteo.DataSetDef = "dsSeteo";
			this.cdsSeteo.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsSeteo.Name = "cdsSeteo";
			this.cdsSeteo.SchemaClassName = "LibContabilidad.DataClass";
			this.cdsSeteo.SchemaDef = null;
			this.cdsSeteo.BeforeFill += new C1.Data.FillEventHandler(this.cdsSeteo_BeforeFill);
			// 
			// cmbProyecto1
			// 
			this.cmbProyecto1.Cursor = System.Windows.Forms.Cursors.Hand;
			this.cmbProyecto1.DataSource = this.cdvProyecto;
			ultraGridColumn225.Header.VisiblePosition = 2;
			ultraGridColumn225.Hidden = true;
			ultraGridColumn226.Header.VisiblePosition = 6;
			ultraGridColumn226.Hidden = true;
			ultraGridColumn227.Header.VisiblePosition = 0;
			ultraGridColumn227.Width = 144;
			ultraGridColumn228.Header.VisiblePosition = 3;
			ultraGridColumn228.Hidden = true;
			ultraGridColumn229.Header.VisiblePosition = 7;
			ultraGridColumn229.Hidden = true;
			ultraGridColumn230.Header.VisiblePosition = 4;
			ultraGridColumn230.Hidden = true;
			ultraGridColumn231.Header.VisiblePosition = 1;
			ultraGridColumn232.Header.VisiblePosition = 5;
			ultraGridColumn232.Hidden = true;
			ultraGridColumn233.Header.VisiblePosition = 8;
			ultraGridColumn233.Hidden = true;
			ultraGridBand16.Columns.AddRange(new object[] {
																											ultraGridColumn225,
																											ultraGridColumn226,
																											ultraGridColumn227,
																											ultraGridColumn228,
																											ultraGridColumn229,
																											ultraGridColumn230,
																											ultraGridColumn231,
																											ultraGridColumn232,
																											ultraGridColumn233});
			this.cmbProyecto1.DisplayLayout.BandsSerializer.Add(ultraGridBand16);
			this.cmbProyecto1.DisplayMember = "Nombre";
			this.cmbProyecto1.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbProyecto1.Location = new System.Drawing.Point(440, 128);
			this.cmbProyecto1.Name = "cmbProyecto1";
			this.cmbProyecto1.Size = new System.Drawing.Size(200, 80);
			this.cmbProyecto1.TabIndex = 116;
			this.cmbProyecto1.ValueMember = "idProyecto";
			this.cmbProyecto1.Visible = false;
			// 
			// cmbSubProyecto1
			// 
			this.cmbSubProyecto1.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbSubProyecto1.DataSource = this.cdvSubProyecto;
			ultraGridBand17.AddButtonCaption = "Proyecto";
			ultraGridColumn234.Header.VisiblePosition = 3;
			ultraGridColumn234.Hidden = true;
			ultraGridColumn235.Header.VisiblePosition = 2;
			ultraGridColumn235.Hidden = true;
			ultraGridColumn236.Header.VisiblePosition = 0;
			ultraGridColumn236.Width = 165;
			ultraGridColumn237.Header.VisiblePosition = 4;
			ultraGridColumn237.Hidden = true;
			ultraGridColumn238.Header.VisiblePosition = 6;
			ultraGridColumn238.Hidden = true;
			ultraGridColumn239.Header.VisiblePosition = 5;
			ultraGridColumn239.Hidden = true;
			ultraGridColumn240.Header.VisiblePosition = 1;
			ultraGridColumn241.Header.VisiblePosition = 7;
			ultraGridColumn241.Hidden = true;
			ultraGridBand17.Columns.AddRange(new object[] {
																											ultraGridColumn234,
																											ultraGridColumn235,
																											ultraGridColumn236,
																											ultraGridColumn237,
																											ultraGridColumn238,
																											ultraGridColumn239,
																											ultraGridColumn240,
																											ultraGridColumn241});
			this.cmbSubProyecto1.DisplayLayout.BandsSerializer.Add(ultraGridBand17);
			this.cmbSubProyecto1.DisplayMember = "Nombre";
			this.cmbSubProyecto1.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbSubProyecto1.Location = new System.Drawing.Point(256, 144);
			this.cmbSubProyecto1.Name = "cmbSubProyecto1";
			this.cmbSubProyecto1.Size = new System.Drawing.Size(232, 80);
			this.cmbSubProyecto1.TabIndex = 117;
			this.cmbSubProyecto1.ValueMember = "idSubProyecto";
			this.cmbSubProyecto1.Visible = false;
			// 
			// helpProvider1
			// 
			this.helpProvider1.HelpNamespace = "C:\\Latinium\\Latinium.chm";
			// 
			// cmbUnidad
			// 
			this.cmbUnidad.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbUnidad.DataMember = "ArticuloUnidad";
			this.cmbUnidad.DataSource = this.cdsUnidad;
			ultraGridBand18.AddButtonCaption = "CompraComprobante";
			ultraGridColumn242.Header.VisiblePosition = 3;
			ultraGridColumn242.Hidden = true;
			ultraGridColumn243.Header.VisiblePosition = 2;
			ultraGridColumn243.Hidden = true;
			ultraGridColumn244.Header.VisiblePosition = 1;
			ultraGridColumn244.Hidden = true;
			ultraGridColumn245.Header.VisiblePosition = 4;
			ultraGridColumn245.Hidden = true;
			ultraGridColumn246.Header.VisiblePosition = 5;
			ultraGridColumn246.Hidden = true;
			ultraGridColumn247.Header.VisiblePosition = 0;
			ultraGridBand18.Columns.AddRange(new object[] {
																											ultraGridColumn242,
																											ultraGridColumn243,
																											ultraGridColumn244,
																											ultraGridColumn245,
																											ultraGridColumn246,
																											ultraGridColumn247});
			this.cmbUnidad.DisplayLayout.BandsSerializer.Add(ultraGridBand18);
			this.cmbUnidad.DisplayMember = "Unidades";
			this.cmbUnidad.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbUnidad.Location = new System.Drawing.Point(528, 160);
			this.cmbUnidad.Name = "cmbUnidad";
			this.cmbUnidad.Size = new System.Drawing.Size(136, 72);
			this.cmbUnidad.TabIndex = 118;
			this.cmbUnidad.ValueMember = "idUnidad";
			this.cmbUnidad.Visible = false;
			// 
			// cdsUnidad
			// 
			this.cdsUnidad.BindingContextCtrl = this;
			this.cdsUnidad.DataLibrary = "LibFacturacion";
			this.cdsUnidad.DataLibraryUrl = "";
			this.cdsUnidad.DataSetDef = "dsArticuloTabla";
			this.cdsUnidad.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsUnidad.Name = "cdsUnidad";
			this.cdsUnidad.SchemaClassName = "LibFacturacion.DataClass";
			this.cdsUnidad.SchemaDef = null;
			this.cdsUnidad.BeforeFill += new C1.Data.FillEventHandler(this.cdsUnidad_BeforeFill);
			// 
			// txtNombre
			// 
			appearance100.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtNombre.Appearance = appearance100;
			this.txtNombre.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsCompra, "_Compra.Nombre"));
			this.txtNombre.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNombre.Location = new System.Drawing.Point(288, 8);
			this.txtNombre.Name = "txtNombre";
			this.txtNombre.Size = new System.Drawing.Size(232, 21);
			this.txtNombre.TabIndex = 1;
			this.txtNombre.Validating += new System.ComponentModel.CancelEventHandler(this.txtNombre_Validating);
			// 
			// btBarras
			// 
			this.btBarras.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			appearance101.Image = ((object)(resources.GetObject("appearance101.Image")));
			this.btBarras.Appearance = appearance101;
			this.btBarras.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.btBarras.Location = new System.Drawing.Point(544, 400);
			this.btBarras.Name = "btBarras";
			this.btBarras.Size = new System.Drawing.Size(24, 23);
			this.btBarras.TabIndex = 119;
			this.btBarras.Click += new System.EventHandler(this.btBarras_Click);
			// 
			// btAuditoria
			// 
			this.btAuditoria.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			appearance102.Image = ((object)(resources.GetObject("appearance102.Image")));
			this.btAuditoria.Appearance = appearance102;
			this.btAuditoria.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.btAuditoria.Location = new System.Drawing.Point(568, 400);
			this.btAuditoria.Name = "btAuditoria";
			this.btAuditoria.Size = new System.Drawing.Size(24, 24);
			this.btAuditoria.TabIndex = 121;
			this.btAuditoria.Click += new System.EventHandler(this.btAuditoria_Click);
			this.btAuditoria.MouseUp += new System.Windows.Forms.MouseEventHandler(this.btAuditoria_MouseUp);
			// 
			// txtServicio
			// 
			this.txtServicio.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			appearance103.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtServicio.Appearance = appearance103;
			this.txtServicio.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsCompra, "_Compra.Servicio"));
			this.txtServicio.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtServicio.FormatString = "#,##0.00";
			this.txtServicio.Location = new System.Drawing.Point(552, 336);
			this.txtServicio.Name = "txtServicio";
			this.txtServicio.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtServicio.PromptChar = ' ';
			this.txtServicio.Size = new System.Drawing.Size(64, 21);
			this.txtServicio.TabIndex = 123;
			// 
			// lblServicio
			// 
			this.lblServicio.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.lblServicio.AutoSize = true;
			this.lblServicio.Location = new System.Drawing.Point(504, 336);
			this.lblServicio.Name = "lblServicio";
			this.lblServicio.Size = new System.Drawing.Size(45, 16);
			this.lblServicio.TabIndex = 122;
			this.lblServicio.Text = "Servicio";
			// 
			// lblPropina
			// 
			this.lblPropina.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.lblPropina.AutoSize = true;
			this.lblPropina.Location = new System.Drawing.Point(616, 336);
			this.lblPropina.Name = "lblPropina";
			this.lblPropina.Size = new System.Drawing.Size(43, 16);
			this.lblPropina.TabIndex = 124;
			this.lblPropina.Text = "Propina";
			// 
			// cmbBodega
			// 
			this.cmbBodega.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbBodega.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbBodega.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsCompra, "_Compra.Bodega"));
			this.cmbBodega.DataSource = this.udsBodegas;
			ultraGridColumn248.Header.VisiblePosition = 0;
			ultraGridColumn248.Width = 250;
			ultraGridColumn249.Header.VisiblePosition = 1;
			ultraGridColumn249.Width = 50;
			ultraGridBand19.Columns.AddRange(new object[] {
																											ultraGridColumn248,
																											ultraGridColumn249});
			this.cmbBodega.DisplayLayout.BandsSerializer.Add(ultraGridBand19);
			this.cmbBodega.DisplayMember = "Nombre";
			this.cmbBodega.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbBodega.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbBodega.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbBodega.Location = new System.Drawing.Point(576, 8);
			this.cmbBodega.MaxDropDownItems = 30;
			this.cmbBodega.Name = "cmbBodega";
			this.cmbBodega.Size = new System.Drawing.Size(112, 21);
			this.cmbBodega.TabIndex = 126;
			this.cmbBodega.ValueMember = "Bodega";
			// 
			// udsBodegas
			// 
			ultraDataColumn13.DataType = typeof(int);
			this.udsBodegas.Band.Columns.AddRange(new object[] {
																													 ultraDataColumn12,
																													 ultraDataColumn13});
			// 
			// lblTiempo
			// 
			this.lblTiempo.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.lblTiempo.AutoSize = true;
			this.lblTiempo.Location = new System.Drawing.Point(504, 250);
			this.lblTiempo.Name = "lblTiempo";
			this.lblTiempo.Size = new System.Drawing.Size(17, 16);
			this.lblTiempo.TabIndex = 127;
			this.lblTiempo.Text = "T: ";
			// 
			// cmbBodegaLinea
			// 
			this.cmbBodegaLinea.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbBodegaLinea.DataSource = this.udsBodegas;
			ultraGridBand20.AddButtonCaption = "CompraComprobante";
			ultraGridColumn250.Header.VisiblePosition = 0;
			ultraGridColumn250.Width = 250;
			ultraGridColumn251.Header.VisiblePosition = 1;
			ultraGridColumn251.Width = 50;
			ultraGridBand20.Columns.AddRange(new object[] {
																											ultraGridColumn250,
																											ultraGridColumn251});
			this.cmbBodegaLinea.DisplayLayout.BandsSerializer.Add(ultraGridBand20);
			this.cmbBodegaLinea.DisplayMember = "Nombre";
			this.cmbBodegaLinea.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbBodegaLinea.Location = new System.Drawing.Point(40, 80);
			this.cmbBodegaLinea.Name = "cmbBodegaLinea";
			this.cmbBodegaLinea.Size = new System.Drawing.Size(176, 120);
			this.cmbBodegaLinea.TabIndex = 128;
			this.cmbBodegaLinea.ValueMember = "Bodega";
			this.cmbBodegaLinea.Visible = false;
			// 
			// txtPropina
			// 
			this.txtPropina.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			appearance104.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtPropina.Appearance = appearance104;
			this.txtPropina.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsCompra, "_Compra.Propina"));
			this.txtPropina.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtPropina.FormatString = "#,##0.00";
			this.txtPropina.Location = new System.Drawing.Point(656, 336);
			this.txtPropina.Name = "txtPropina";
			this.txtPropina.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtPropina.PromptChar = ' ';
			this.txtPropina.Size = new System.Drawing.Size(72, 21);
			this.txtPropina.TabIndex = 129;
			// 
			// cmbUnidadB
			// 
			this.cmbUnidadB.Cursor = System.Windows.Forms.Cursors.Hand;
			this.cmbUnidadB.DataMember = "ArticuloUnidad";
			this.cmbUnidadB.DataSource = this.cdsUnidad;
			ultraGridBand21.AddButtonCaption = "CompraComprobante";
			ultraGridColumn252.Header.VisiblePosition = 3;
			ultraGridColumn252.Hidden = true;
			ultraGridColumn253.Header.VisiblePosition = 2;
			ultraGridColumn253.Hidden = true;
			ultraGridColumn254.Header.VisiblePosition = 1;
			ultraGridColumn254.Hidden = true;
			ultraGridColumn255.Header.VisiblePosition = 4;
			ultraGridColumn255.Hidden = true;
			ultraGridColumn256.Header.VisiblePosition = 5;
			ultraGridColumn256.Hidden = true;
			ultraGridColumn257.Header.VisiblePosition = 0;
			ultraGridBand21.Columns.AddRange(new object[] {
																											ultraGridColumn252,
																											ultraGridColumn253,
																											ultraGridColumn254,
																											ultraGridColumn255,
																											ultraGridColumn256,
																											ultraGridColumn257});
			this.cmbUnidadB.DisplayLayout.BandsSerializer.Add(ultraGridBand21);
			this.cmbUnidadB.DisplayMember = "Unidades";
			this.cmbUnidadB.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbUnidadB.Location = new System.Drawing.Point(288, 178);
			this.cmbUnidadB.Name = "cmbUnidadB";
			this.cmbUnidadB.Size = new System.Drawing.Size(136, 72);
			this.cmbUnidadB.TabIndex = 130;
			this.cmbUnidadB.ValueMember = "idUnidad";
			this.cmbUnidadB.Visible = false;
			// 
			// cdsSeteoC
			// 
			this.cdsSeteoC.BindingContextCtrl = this;
			this.cdsSeteoC.DataLibrary = "LibContabilidad";
			this.cdsSeteoC.DataLibraryUrl = "";
			this.cdsSeteoC.DataSetDef = "dsSeteoC";
			this.cdsSeteoC.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsSeteoC.Name = "cdsSeteoC";
			this.cdsSeteoC.SchemaClassName = "LibContabilidad.DataClass";
			this.cdsSeteoC.SchemaDef = null;
			this.cdsSeteoC.BeforeFill += new C1.Data.FillEventHandler(this.cdsSeteoC_BeforeFill);
			// 
			// Compra
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.BackColor = System.Drawing.SystemColors.Control;
			this.ClientSize = new System.Drawing.Size(744, 429);
			this.Controls.Add(this.cmbUnidadB);
			this.Controls.Add(this.txtPropina);
			this.Controls.Add(this.cmbBodegaLinea);
			this.Controls.Add(this.lblTiempo);
			this.Controls.Add(this.lblPropina);
			this.Controls.Add(this.lblServicio);
			this.Controls.Add(this.lblIce);
			this.Controls.Add(this.lblSubTotal);
			this.Controls.Add(this.lblCliente);
			this.Controls.Add(this.label8);
			this.Controls.Add(this.label7);
			this.Controls.Add(this.label6);
			this.Controls.Add(this.lblIva);
			this.Controls.Add(this.lblDescuento);
			this.Controls.Add(this.lblIvaSubt);
			this.Controls.Add(this.lblIva0);
			this.Controls.Add(this.lblBodega);
			this.Controls.Add(this.cmbBodega);
			this.Controls.Add(this.txtServicio);
			this.Controls.Add(this.btAuditoria);
			this.Controls.Add(this.btBarras);
			this.Controls.Add(this.txtNombre);
			this.Controls.Add(this.cmbUnidad);
			this.Controls.Add(this.cmbSubProyecto1);
			this.Controls.Add(this.cmbProyecto1);
			this.Controls.Add(this.btNuevo);
			this.Controls.Add(this.txtIceTotal);
			this.Controls.Add(this.txtTotal);
			this.Controls.Add(this.txtIvaTotal);
			this.Controls.Add(this.txtDescIva0Total);
			this.Controls.Add(this.txtDescIvaTotal);
			this.Controls.Add(this.txtIva);
			this.Controls.Add(this.txtIva0);
			this.Controls.Add(this.utcTransaccion);
			this.Controls.Add(this.barraDato1);
			this.Controls.Add(this.txtBorrar);
			this.Controls.Add(this.txtIdCompra);
			this.Controls.Add(this.txtIdAsiento);
			this.Controls.Add(this.btActualizaPrecio);
			this.Controls.Add(this.btGuiaRemision);
			this.Controls.Add(this.btCreaAsiento);
			this.Controls.Add(this.btRetencion);
			this.Controls.Add(this.btNotaEntrega);
			this.Controls.Add(this.btTransforma);
			this.Controls.Add(this.btBuscaArticulo);
			this.Controls.Add(this.btSaldo1);
			this.Controls.Add(this.cmbTipoCompra);
			this.Controls.Add(this.cmbComprobante1);
			this.Controls.Add(this.btCliente);
			this.Controls.Add(this.spnBodega);
			this.Controls.Add(this.dtVence);
			this.Controls.Add(this.spnDias);
			this.Controls.Add(this.dtFecha);
			this.Controls.Add(this.txtNumero);
			this.Controls.Add(this.cmbCliente);
			this.Controls.Add(this.CmbCodDesc);
			this.Controls.Add(this.CmbDescCod);
			this.Controls.Add(this.ultraGrid1);
			this.Controls.Add(this.lblAnulado);
			this.Controls.Add(this.btTotal);
			this.helpProvider1.SetHelpKeyword(this, "Facturación.htm");
			this.helpProvider1.SetHelpNavigator(this, System.Windows.Forms.HelpNavigator.Topic);
			this.Name = "Compra";
			this.helpProvider1.SetShowHelp(this, true);
			this.Text = "Compra";
			this.Resize += new System.EventHandler(this.Compra_Resize);
			this.Closing += new System.ComponentModel.CancelEventHandler(this.Compra_Closing);
			this.Load += new System.EventHandler(this.Compra2_Load);
			this.ultraTabPageControl1.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.cmbVendedor)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsCompra)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvPersonal)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsPersonal)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtComprobante)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbSubProyecto)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvSubProyecto)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsProyecto)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbProyecto)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvProyecto)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNotas)).EndInit();
			this.ultraTabPageControl2.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.txtComision)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbPrecio)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsClienteTabla)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.spnVuelto)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNumCuenta)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCheque)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbFormaPago)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvFormaPago)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsCompraTabla)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.rdgContadoCred)).EndInit();
			this.ultraTabPageControl3.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.txtDescPorcIva)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDescPorcIva0)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDescLinea0)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDescLinea)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDescIva)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDescIva0)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtFlete)).EndInit();
			this.pgEntrega.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.txtComprob1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbRecibe)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvEntrega)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsCliente)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtOtro)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDepartamento)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFechaRevision)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFechaEntrega)).EndInit();
			this.ultraTabPageControl5.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.ultraNumericEditor1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraNumericEditor2)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtInteresPagar)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtEmpezarEn)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtPagosCada)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtPorcInteres)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtTotalCuotas)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtValorCuotas)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtAnticipo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNumCuotas)).EndInit();
			this.TabPageCOA.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.cmbIce)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.OptBienServicio)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtRefrendo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCredTribut)).EndInit();
			this.tbpAutorizacion.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.txtUsuario)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCaducidad)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.spnPrioridad)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtSerieFact)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtAutImprenta)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtAutFactura)).EndInit();
			this.ultraTabPageControl6.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.txtBusca)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ugrBusca)).EndInit();
			this.ultraTabPageControl4.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.ultraTextEditor6)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraTextEditor5)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraTextEditor4)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraTextEditor2)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCIF)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbImportacion)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsTipoImportacion)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraTextEditor3)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraTextEditor1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtPeso)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtBulto)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.udsBusca)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvCliente)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvArticulo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsArticulo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvArtCodigo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvComprobante)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvNumero)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeguridad)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.CmbDescCod)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.CmbCodDesc)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvCompra)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsAsiento)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsRetencion)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsAsignaCuenta)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCliente)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNumero)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFecha)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.spnDias)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtVence)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.spnBodega)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbComprobante1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbTipoCompra)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdAsiento)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdCompra)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtBorrar)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.utcTransaccion)).EndInit();
			this.utcTransaccion.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.txtDescIva0Total)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDescIvaTotal)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIva)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIva0)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIceTotal)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtTotal)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIvaTotal)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbProyecto1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbSubProyecto1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbUnidad)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsUnidad)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNombre)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtServicio)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbBodega)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.udsBodegas)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbBodegaLinea)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtPropina)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbUnidadB)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoC)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion


		private void ValoresIniciales()
		{
			this.barraDato1_Refresca(this, new System.EventArgs());
			if (this.IdTipoFactura == 1 || this.IdTipoFactura == 3) // Factura de venta
			{
				if (drSeteoF["NotaVentaP"] != System.DBNull.Value && (bool) drSeteoF["NotaVentaP"])
					drCompra["idComprobante"] = 2;
				else drCompra["idComprobante"] = 15;
				drCompra["idCliente"] = (int) drSeteoF["idCliente"];
				if (drSeteo["SriaAutoriz"] != System.DBNull.Value)
					drCompra["AutFactura"] = drSeteo["SriaAutoriz"];
				if (drSeteo["SriSerie"] != System.DBNull.Value)
					drCompra["SerieFactura"] = drSeteo["SriSerie"];
				if (drSeteo["SriCaduca"] != System.DBNull.Value)
					drCompra["FechaCaducidad"] = drSeteo["SriCaduca"];
				if (bAutoImpresor)
					drCompra["FechaCaducidad"] = FechaCaducidadSri;

				if (!bListaCliente)
				{
					string stNombre = "Select Count(*) From Cliente Where idCliente = "
						+ drCompra["idCliente"].ToString();
					int iCuenta = Funcion.iEscalarSQL(cdsCliente, stNombre);
					if (iCuenta > 0)
					{
						stNombre = "Select Nombre From Cliente Where idCliente = "
							+ drCompra["idCliente"].ToString();
						drCompra["Nombre"] = Funcion.sEscalarSQL(cdsCliente, stNombre, false);
					}
				}
			}
			else
				drCompra["idComprobante"] = 1;
			if ((bool) drSeteoF["VentaCredito"] == true)
				drCompra["ContadoCredito"] = 2;
			else
				drCompra["ContadoCredito"] = 1;

			//Asignacion de Proyectos, sub y bodega predefinidos o por maquina
			if (IProyecto > 0) drCompra["idProyecto"] = IProyecto;
			if (ISubProyecto > 0) drCompra["idSubProyecto"] = ISubProyecto;
			if (IBodega > 0) drCompra["Bodega"] = IBodega;

			drCompra["idFormaPago"] = 1;
			if (Funcion.bEjecutaSQL(cdsArticulo, "Exec SeteoGExiste 'FPFC'"))
				drCompra["Fecha"] = DateTime.Today;
			else
				drCompra["Fecha"] = Funcion.dtEscalarSQL(cdsArticulo, "Select Convert(DateTime, Convert(Char(10), getdate(),112), 112)");

			drCompra["idTipoFactura"] = IdTipoFactura;
			drCompra["idSucursal"] = (int) drSeteo["NumEmpresa"];
			if (IdTipoFactura == 5 || IdTipoFactura == 6 ||
				IdTipoFactura == 27 || IdTipoFactura == 28) // Devoluciones y notas de Credito
			{
				drCompra["ContadoCredito"] = 2;
				if (IdTipoFactura == 5 || IdTipoFactura == 6)
					drCompra["idComprobante"] = 4;
				else
					drCompra["idComprobante"] = 5;
			}
		} 

		private void barraDato1_Crear(object sender, System.EventArgs e)
		{
			if (MenuLatinium.bAcademica && iCuentaTotal > 600)
			{
				MessageBox.Show("No tiene acceso a creación de Facturas en version Academica", "Confirmación", 
					MessageBoxButtons.OK, MessageBoxIcon.Information);
				return;
			}
			if (!miAcceso.Nuevo)
			{
				MessageBox.Show("No tiene Acceso a Crear " + this.Text.Trim(),
					"Seguridad", MessageBoxButtons.OK, MessageBoxIcon.Information);
				return;
			}
			if (bAutoImpresor)
			{
				string stExec = "Exec AutoImpresorCaducado " + IdTipoFactura 
					+ ", " + cmbComprobante1.Value.ToString();
				bool bCaducado = false;
				//Funcion.bEjecutaSQL(cdsArticulo, stExec);
				if (bCaducado)
				{
					MessageBox.Show("Caducidad de Autorizacion excede la fecha actual ",
						"Seguridad", MessageBoxButtons.OK, MessageBoxIcon.Information);
					return;
				}
			}

			this.barraDato1.CrearRegistro();
			ValoresIniciales();
			this.HabilitaTab();
			// Precio por Forma de Pago
			if ((int) this.drSeteoF["TipoPrecio"] == 2)
			{
				switch(this.IdTipoFactura)
				{
					case 1: // Factura
					case 3: // Cotizacion
					case 8: // Transferencia Enviada
					case 9: // Transferencia Recibida
					case 10: // Orden de Produccion
					case 13: // Reservacion Cliente
					case 15: // Pedido Cliente
						using (CompraFormaPagoPrecio miFormaPago = new CompraFormaPagoPrecio())
						{
							miFormaPago.ShowDialog();
							drCompra["idFormaPago"] = int.Parse(miFormaPago.IdFormaPago.Value.ToString());
						}
						break;
				}
			}
			if (bListaCliente) cmbCliente.Select();
			else txtNombre.Select();
			ModificaGrilla(true);
		}

		private void ActualizaSaldo(int idSigno)
		{
			if (!(idSigno == 1 || idSigno == -1)) return;

			double fCantOriginal = 0;
			double fSaldo = 0;
			double fCantidad = 0;
			int IdArticulo = 0;
			int IdDetCompra = 0;

			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.ultraGrid1.Rows.All)
			{
				if (dr.Cells["Cantidad"].Value != System.DBNull.Value)
					fCantidad = (double) dr.Cells["Cantidad"].Value;
				if (dr.Cells["idArticulo"].Value != System.DBNull.Value)
					IdArticulo = (int) dr.Cells["idArticulo"].Value;
				if (dr.Cells["idDetCompra"].Value != System.DBNull.Value)
					IdDetCompra = (int) dr.Cells["idDetCompra"].Value;
				this.cdsArticulo.Schema.Connections[0].Open();
				C1.Data.SchemaObjects.Connection miConeccion = this.cdsArticulo.Schema.Connections[0];
				SqlCommand cmdSaldo = new SqlCommand (
					"SELECT IsNull(Saldo, 0) FROM Articulo WHERE idArticulo = " +
					IdArticulo.ToString(), (SqlConnection) miConeccion.DbConnection);

				// Para evitar el error cuando no existe articulo IdArt = 0
				try
				{
					fSaldo = (double) cmdSaldo.ExecuteScalar();
				}
				catch {}

				// Busqueda de Valor Original
				if (!this.barraDato1.bNuevo)
				{
					SqlCommand cmdCant = new SqlCommand (
						"SELECT Cantidad FROM DetCompra WHERE idDetCompra = " +
						IdDetCompra.ToString(), (SqlConnection)miConeccion.DbConnection);

					try
					{
						fCantOriginal = (double) cmdCant.ExecuteScalar();
					}
					catch 
					{
						fCantOriginal = 0;
					}
				}

				if (fCantidad != fCantOriginal)
				{
					fCantidad = fSaldo + (fCantidad- fCantOriginal)* idSigno;
					SqlCommand cmdActualiza = new SqlCommand (
						"UPDATE Articulo Set Saldo = " +
						fCantidad.ToString() + " WHERE idArticulo = " +
						IdArticulo.ToString()
						, (SqlConnection)miConeccion.DbConnection);
					cmdActualiza.Transaction = (SqlTransaction)
						miConeccion.DbTransaction;
					cmdActualiza.ExecuteNonQuery();
				}
				this.cdsCompra.Schema.Connections[0].Close();
			}
		}

		private void SaldoCxC()
		{
			double dSaldo = 0;
			this.cdsCompra.Schema.Connections[0].Open();
			C1.Data.SchemaObjects.Connection miConeccion = this.cdsCompra.Schema.Connections[0];
			SqlCommand cmdSaldo = new SqlCommand (
				"SELECT SUM(Pago) FROM Pago WHERE idCompra = " +
				this.txtIdCompra.Value.ToString(), (SqlConnection) miConeccion.DbConnection);

			try
			{
				dSaldo = (double) cmdSaldo.ExecuteScalar();
			}
			catch(System.Exception ex)
			{
				Console.WriteLine(ex.Message);
			}

			//Calculo de retencion *******************************************************************
			double dRetencion = 0;
			cmdSaldo = new SqlCommand (
				"SELECT RentaRetenida FROM Retencion WHERE idCompra = " +
				this.txtIdCompra.Value.ToString(), (SqlConnection) miConeccion.DbConnection);

			try
			{
				dRetencion += (double) cmdSaldo.ExecuteScalar();
			}
			catch(System.Exception ex)
			{
				Console.WriteLine(ex.Message);
			}
			cmdSaldo = new SqlCommand (
				"SELECT IvaRetenido FROM Retencion WHERE idCompra = " +
				this.txtIdCompra.Value.ToString(), (SqlConnection) miConeccion.DbConnection);

			try
			{
				dRetencion += (double) cmdSaldo.ExecuteScalar();
			}
			catch(System.Exception ex)
			{
				Console.WriteLine(ex.Message);
			}
			cmdSaldo = new SqlCommand (
				"SELECT RentaRetenida1 FROM Retencion WHERE idCompra = " +
				this.txtIdCompra.Value.ToString(), (SqlConnection) miConeccion.DbConnection);

			try
			{
				dRetencion += (double) cmdSaldo.ExecuteScalar();
			}
			catch(System.Exception ex)
			{
				Console.WriteLine(ex.Message);
			}
			cmdSaldo = new SqlCommand (
				"SELECT IvaRetenido1 FROM Retencion WHERE idCompra = " +
				this.txtIdCompra.Value.ToString(), (SqlConnection) miConeccion.DbConnection);

			try
			{
				dRetencion += (double) cmdSaldo.ExecuteScalar();
			}
			catch(System.Exception ex)
			{
				Console.WriteLine(ex.Message);
			}

			this.cdsCompra.Schema.Connections[0].Close();
			// Fin de calculo de retenciones ***********************************************

			this.drCompra["Saldo"] = double.Parse(this.txtTotal.Value.ToString()) 
				- dSaldo - dRetencion;
			this.drCompra["Retenido"] = dRetencion;
		}

		private void BorraErrores()
		{
			errorProvider.SetError(cmbCliente, "");
			errorProvider.SetError(txtNumero, "");
			errorProvider.SetError(dtFecha, "");
			errorProvider.SetError(rdgContadoCred, "");
			errorProvider.SetError(txtTotal, "");
			errorProvider.SetError(cmbVendedor, "");
			errorProvider.SetError(txtAutFactura, "");
			errorProvider.SetError(utcTransaccion, "");
			errorProvider.SetError(txtSerieFact, "");
			errorProvider.SetError(utcTransaccion, "");
			errorProvider.SetError(cmbCaducidad, "");
			errorProvider.SetError(cmbProyecto, "");
			errorProvider.SetError(cmbSubProyecto, "");
		}

		private bool VerificaNumero()
		{
			bool bOk = true;

			if (txtNumero.Text.Trim().Length == 0 && bNumeraGrabar == false)
			{
				errorProvider.SetError(txtNumero, "Ingrese Numero de Comprobante");
				return false;
			}
			int iCuenta = 0;
			string stSelect = "SELECT COUNT(*) FROM Compra WHERE Compra.Numero = '" +
				this.txtNumero.Text.ToString().Trim() + "' And Compra.idTipoFactura = "
				+ this.IdTipoFactura.ToString().Trim() + " And Compra.idComprobante = "
				+ cmbComprobante1.Value.ToString() ;
			if (!barraDato1.bNuevo)
			{
				stSelect += " AND Compra.idCompra <> " + txtIdCompra.Text.Trim();
			}
			if (this.IdTipoFactura == 4) // Caso de Compras se verifica que no tenga el cliente factura repetida
			{
				stSelect += " And Compra.idCliente = " + this.cmbCliente.Value.ToString().Trim();
			}
			iCuenta = Funcion.iEscalarSQL(cdsCompra, stSelect);
			if (iCuenta > 0) 
			{
				bOk = false;
				errorProvider.SetError(txtNumero, "Numero de Comprobante ya existe");
			}
			return bOk;
		}

		private double SelectSaldo()
		{
			string stIdCompra = "0";
			if (barraDato1.bEditar || IdTipoFactura == 3) stIdCompra = txtIdCompra.Value.ToString();
			string stExec = "Exec SaldoClienteMenosId " + this.cmbCliente.Value.ToString().Trim()
				+ ", " + IdTipoFactura.ToString() + ", " + stIdCompra;
//			string stIgnoraActual = "";
//			if (barraDato1.bEditar) stIgnoraActual = " And Compra.idCompra <> " + txtIdCompra.Value.ToString();
//
//			string stSelect = "SELECT Sum(ISNULL(Compra.Saldo* CompraNumero.CobroPago, 0)) "
//				+ "FROM Compra Inner Join " +
//				"CompraNumero On Compra.idTipoFactura = CompraNumero.idTipoFactura "
//				+ "	WHERE Compra.idCliente = "
//				+ this.cmbCliente.Value.ToString().Trim() 
//				+ " And CompraNumero.CobroPago IN (1, -1) "
//				+ " And Compra.ContadoCredito = 2"
//				+ " And Compra.Borrar = 0" + stIgnoraActual;
			double dSaldo = Funcion.dEscalarSQL(cdsCliente, stExec, false);
			if (barraDato1.bEditar && (int) drCompra["ContadoCredito"] == 2) 
				dSaldo += (double) txtTotal.Value;
			return dSaldo;
		}

		private string LimiteCredito(int idTipoFactura)
		{
			string stMensaje = "";
			if (idTipoFactura == 1 && (int) rdgContadoCred.Value == 2 
				&& Funcion.Permiso("511", cdsSeguridad) && !(bool) chkAutPrincipal.Checked)
			{
				try
				{
					double dSaldoCxC = SelectSaldo();
					double dMaximoCred = miProveedor.MaximoCredito;
					string stSaldo = "Select IsNull(Sum(Pago), 0) From Pago Where idCompra = "
						+ txtIdCompra.Value.ToString();
					double dPago = Funcion.dEscalarSQL(cdsCompra, stSaldo, false);
					dSaldoCxC += - dPago;
					if (dSaldoCxC > dMaximoCred)
					{
						stMensaje = "Infringe Cupo de Credito. Saldo: " + dSaldoCxC.ToString("#,##0.00", us) + 
							" Cupo Maximo: " + dMaximoCred.ToString("#,##0.00", us);
						errorProvider.SetError(cmbCliente, stMensaje);
					}
				}
				catch
				{
					stMensaje = "Cliente No valido";
					errorProvider.SetError(cmbCliente, stMensaje);
				}
			}
			return stMensaje;
		}
		private bool VerificaCentroCosto()
		{
			bool bOk = true;
			BorraErrores();
			
			// Verificacion de Centro de Costo o proyecto  Venta y Compra
			if (IdTipoFactura == 1 || IdTipoFactura == 4)
			{
				if (Funcion.bEjecutaSQL(cdsArticulo, "Exec SeteoGExiste 'FVCP'"))
				{
					if ((int)cmbProyecto.Value == 0)
					{
						errorProvider.SetError(cmbProyecto, "Ingrese Centro de Costo");
						errorProvider.SetError(utcTransaccion, "Error en Ficha Notas");
						bOk = false;
					}
					if((int) cmbSubProyecto.Value == 0)
					{
						errorProvider.SetError(cmbSubProyecto, "Ingrese Proyecto");
						errorProvider.SetError(utcTransaccion, "Error en Ficha Notas");
						bOk = false;
					}
				}
			}
			return bOk;
		}
		private bool Verifica()
		{
			bool bOk = true;
			BorraErrores();
			//Verificacion de Autorizaciones
			string CodComprob = "";
			if (cmbComprobante1.ActiveRow != null &&
				cmbComprobante1.ActiveRow.Cells["Codigo"].Value != null
				&& cmbComprobante1.ActiveRow.Cells["Codigo"].Value != DBNull.Value)
					CodComprob = cmbComprobante1.ActiveRow.Cells["Codigo"].Value.ToString().Trim();
			if (IdTipoFactura == 4 && !(CodComprob == "99") &&
				drSeteoF["ExigeAutorizacion"] != System.DBNull.Value &&
				(bool) drSeteoF["ExigeAutorizacion"] == true)
			{
				if (drCompra["AutFactura"] == System.DBNull.Value
					|| drCompra["AutFactura"].ToString().Trim().Length != 10)
				{
					errorProvider.SetError(txtAutFactura, "Autorizacion de Factura debe ser de 10 digitos");
					errorProvider.SetError(utcTransaccion, "Error en autorizacion");
					bOk = false;
				}
				if (drCompra["SerieFactura"] == System.DBNull.Value
					|| drCompra["SerieFactura"].ToString().Trim().Length != 6)
				{
					errorProvider.SetError(txtSerieFact, "Serie de Factura debe ser de 6 digitos");
					errorProvider.SetError(utcTransaccion, "Error en autorizacion");
					bOk = false;
				}
				if (drCompra["FechaCaducidad"] == System.DBNull.Value
					|| (DateTime) drCompra["FechaCaducidad"] < (DateTime) drCompra["Fecha"])
				{
					errorProvider.SetError(utcTransaccion, "Error en autorizacion");
					errorProvider.SetError(cmbCaducidad, "Fecha de Caducidad caducada");
					bOk = false;
				}
				int IdCliente = (int) this.cmbCliente.Value;
				string stVerificaRango = "Exec ClienteVerificaRango " + 
					IdCliente.ToString() + ", '" + txtNumero.Text.Trim() + "',"+ cmbComprobante1.Value.ToString();
				if (!Funcion.bEjecutaSQL(cdsArticulo, stVerificaRango))
				{
					errorProvider.SetError(cmbCliente, "Error en Rango de Autorizacion / Inconsistencia (Cedula con Factura) o (Ruc con Liquidacion)");
					bOk = false;
				}
			}

			if (this.rdgContadoCred.CheckedIndex < 0)
			{
				MessageBox.Show("Ingrese Forma de Pago", "Informacion");
				this.errorProvider.SetError(this.rdgContadoCred, 
					"Ingrese forma de pago.");
				errorProvider.SetError(utcTransaccion, "Error en Ficha Pago");
				return false;
			}

			if (this.cmbCliente.Value == System.DBNull.Value || (int) this.cmbCliente.Value < 1)
			{
				int idTipoF = (int) cmbTipoCompra.Value;
				if ((idTipoF <= 6 || idTipoF >= 23 || idTipoF == 13 || idTipoF == 14)
					&& idTipoF != 36) // no verifica 36
				{
					this.errorProvider.SetError(this.cmbCliente, "Ingrese Nombre");
					bOk = false;
				}
			}
			// Limitar Credito en Ventas
			if (LimiteCredito(IdTipoFactura).Length > 0)
				bOk = false;

			// Exige vendedor en venta y transferencia enviada
			if ((bool) drSeteoF["ExigeVendedor"] == true && 
				(IdTipoFactura == 1 || IdTipoFactura == 8) 
				&& (int) cmbVendedor.Value == 0)
			{
				errorProvider.SetError(cmbVendedor, "Ingrese Vendedor");
				errorProvider.SetError(utcTransaccion, "Error en Ficha Nota Vendedor");
				bOk = false;
			}

			double dPagos = 0;
			double dRetenido = 0;

			double dTotal = (double) this.drCompra["Total"];
			if (drCompra["Retenido"] != System.DBNull.Value)
				dRetenido = (double) this.drCompra["Retenido"];
			if (!barraDato1.bNuevo)
			{
				string stPagos = "Select Sum(IsNull(Pago, 0)) From Pago "
					+ "Where idCompra = " + txtIdCompra.Value.ToString();
				dPagos = Funcion.dEscalarSQL(cdsCompra, stPagos, false);
			}

			if ((int) this.rdgContadoCred.Value == 1 && Math.Round(dPagos, 2) > 0)
			{
				MessageBox.Show("Forma de Pago no puede cambiar a contado", "Informacion");
				this.errorProvider.SetError(this.rdgContadoCred, 
					"No puede modificar forma de pago. "
					+ "Pagos de esta Factura por: " 
					+ dPagos.ToString("#,##0.00"));
				bOk = false;
			}
			//Calculo de Pagos reales sin pendientes
			if (!barraDato1.bNuevo)
			{
				string stPagos = "Select Sum(IsNull(Pago, 0)) From Pago "
					+ "Where idCompra = " + txtIdCompra.Value.ToString()
					+ " And Pendiente = 0";
				dPagos = Funcion.dEscalarSQL(cdsCompra, stPagos, false);
			}
			dPagos += dRetenido;
			if ((int) this.rdgContadoCred.Value == 2 && 
				Math.Round(dPagos, 2) > Math.Round(dTotal, 2) + .001) // Pago Excede factura
			{
				this.errorProvider.SetError(this.txtTotal, 
					"Valor Total de Factura " + dTotal.ToString("#,##0.00", us)
					+ " es menor que lo cancelado " + dPagos.ToString("#,##0.00", us));
				bOk = false;
			}

			// Cuando esta habilitado cambiar el periodo no realiza verificacion de fecha
			if (miAcceso.CambiarPeriodo) return bOk;

			// Verificacion de Fecha
			DateTime dtFechaFact;
			dtFechaFact = DateTime.Today.AddYears(-1000);
			DateTime dtFechaIni;
			DateTime dtFechaFin = DateTime.Today;
			try // En caso de ser nulos no verifica
			{
				dtFechaIni = (DateTime) drSeteoF["FechaInicio"];
				dtFechaFin = (DateTime) drSeteoF["FechaFinal"];
			}
			catch
			{
				return bOk;
			}

			if (this.dtFecha.Value.ToString().Trim().Length < 3)
			{
				this.errorProvider.SetError(this.dtFecha, "Ingrese Fecha");
				bOk = false;
			}
			else
			{
				dtFechaFact = DateTime.Parse(this.dtFecha.Value.ToString());
			}

			if (drSeteoF["FechaInicio"] == System.DBNull.Value 
				||drSeteoF["FechaFinal"] == System.DBNull.Value
				|| (dtFechaFact < DateTime.Today.AddYears(-100)))
			{
				this.errorProvider.SetError(this.dtFecha, "");
			}
			else // Existen las tres fechas
			{
				if (dtFechaFact > dtFechaFin || dtFechaFact < dtFechaIni)
				{
					bOk = false;
					this.errorProvider.SetError(this.dtFecha, "Fecha debe estar entre el " 
						+ dtFechaIni.ToString("dd/MMM/yyyy")+ " y el " 
						+ dtFechaFin.ToString("dd/MMM/yyyy") + ".\n"
						+ "Fechas definidas en Archivo - Propiedades - Facturación");
				}
			}
			// Fin de verificacion de fecha

			return bOk;
		}

		private bool VerificaBodega()
		{
			bool bOk = true;
			if (Funcion.bEjecutaSQL(cdsArticulo, "Exec SeteoGExiste 'FPVCNSP'")) return true;

			if (drSeteoF["SaldoPorBodega"] != DBNull.Value && (bool) drSeteoF["SaldoPorBodega"] == true)
			{
				// Verificacion de Signo Negativo
				string stVerificaSigno = "Select Signo From CompraNumero Where idTipoFactura = "
					+ IdTipoFactura.ToString();
				bool bVerificaCotiza = false;
				if (IdTipoFactura == 3 && (drSeteoF["SaldoEnCotiza"] != DBNull.Value 
					&& (bool) drSeteoF["SaldoEnCotiza"] == true)) bVerificaCotiza = true;

				if (Funcion.iEscalarSQL(cdsArticulo, stVerificaSigno, true) != -1 && !bVerificaCotiza)
				{
					return bOk;
				}
				string stArticulos = "";
				bool bBodegaGlobal = true;
				int iBodega = (int)	drCompra["Bodega"];
				if (drSeteoF["VerBodega"] != DBNull.Value && (bool) drSeteoF["VerBodega"] == true)
					bBodegaGlobal = false;
				foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in ultraGrid1.Rows.All)
				{
					if (dr.Cells["idArticulo"].Value != DBNull.Value)
						if (dr.Cells["Cantidad"].Value != DBNull.Value)
						{
							if (!bBodegaGlobal) 
							{
								if (dr.Cells["Bodega"].Value != DBNull.Value)
									iBodega = iBodega;
								//bOk = false;
								//string stMensaje = "Bodega en linea no valida";
							}
							double dCantidad = (double) dr.Cells["Cantidad"].Value;
							stArticulos = dr.Cells["idArticulo"].Value.ToString() + ", "
								+ dCantidad.ToString("0.00", us) + ", " + iBodega.ToString("0", us);
							string stExec = "Exec CompraBodegaNegativo1 " + stArticulos
								+ ", " + txtIdCompra.Value.ToString();

							double dSaldo = Funcion.dEscalarSQL(cdsArticulo, stExec, true) + dCantidad;
							if (dSaldo - dCantidad < -0.002)
							{
								int IdArt = (int) dr.Cells["idArticulo"].Value;
								string stArticulo = "Select Articulo From Articulo Where idArticulo = "
									+ IdArt.ToString();
								stArticulo = Funcion.sEscalarSQL(cdsArticulo, stArticulo);
								MessageBox.Show("No puede vender: " + dCantidad.ToString("0.00", us)
									+ " Saldo: " + dSaldo.ToString("0.00", us) 
									+ " Articulo: " + stArticulo, "Informacion");
								return false;
							}
							//							stArticulos += dr.Cells["idArticulo"].Value.ToString() + "; "
							//								+ dCantidad.ToString("0.00", us) + "; ";
						}
				}
				if (stArticulos.EndsWith("; ")) stArticulos = stArticulos.Substring(0, stArticulos.Length -2);
				//string stExec = "Exec CompraBodegaNegativo '" + stArticulos + "'";
			}
			return bOk;
		}

		bool bAutoImpresor = false;
		bool bNumeraGrabar = false;
		private void barraDato1_Grabar(object sender, System.EventArgs e)
		{
			if (!CambiosBotrosa())
			{
				MessageBox.Show("Ingrese Centro de Costo..... \n Campo no pude estar vacio","Informacion.....",MessageBoxButtons.OK,MessageBoxIcon.Information);
				return;
			}
			
			this.cmdTotal();
			if (!Verifica()) return;
			if (!VerificaCentroCosto()) return;
			if (!VerificaBodega()) return;
			if (!miAcceso.GrabaCero && (double) txtTotal.Value == 0) 
			{
				MessageBox.Show("No puede grabar la transaccion en Cero");
				return;
			}
			this.barraDato1.Focus();
			bool bNuevoReg = this.barraDato1.bNuevo;
			int iTipoNumera = IdTipoFactura;
			if (bAutoImpresor && IdTipoFactura == 1)
			{
				if (!((int)cmbComprobante1.Value == 15 || (int) cmbComprobante1.Value == 2))
				{
					MessageBox.Show("Solo se permiten Documentos Autorizados (18) o Nota de Venta", "Informacion");
					return;
				}
			}
			if (bAutoImpresor && IdTipoFactura == 5)
			{
				if (txtComprobante.Text.Trim().Length==0)
				{
					MessageBox.Show("Notas de Credito Deben tener comprobante que afectan", "Informacion");
					return;
				}
			}

			if (!Funcion.bEjecutaSQL(cdsArticulo, "Exec SeteoGExiste 'FTRNUM'"))
			{
				if (bNuevoReg && !bAutoImpresor && !bNumeraGrabar) // Numeración
				{
					if (cmbComprobante1.ActiveRow != null && 
						cmbComprobante1.ActiveRow.Cells["idCompraNumero"].Value 
						!= System.DBNull.Value)
						iTipoNumera = (int) cmbComprobante1.ActiveRow.Cells["idCompraNumero"].Value;
					if (iTipoNumera == 0 || IdTipoFactura == 3) // Diferente de Cotizacion
						iTipoNumera = IdTipoFactura;
					string stNumero = Funcion.NumeraFactura(cdsCompraTabla, iTipoNumera);
					if (stNumero.Trim().Length > 0) drCompra["Numero"] = stNumero;
					else
					{
						if (txtNumero.ToString().Trim().Length == 0)
						{
							this.errorProvider.SetError(txtNumero, "Ingrese Codigo");
							return;
						}
					}
				}
			}

			if (!bAutoImpresor && !VerificaNumero()) return;

			if (!barraDato1.GrabaRegistro())
				if (bNuevoReg)
				{
					MessageBox.Show("No se puede grabar la factura sistema ocupado", "Informacion");
					return;
//					Funcion.RegresaFactura(cdsCompra, iTipoNumera);
				}
		
			this.HabilitaTab();
			ModificaGrilla(false);
			if (bNuevoReg) // Actualiza Posicion de nuevo Registro
			{
				if (Funcion.bEjecutaSQL(cdsArticulo, "Exec SeteoGExiste 'FTRNUM'"))
				{
					if (bNuevoReg && !bAutoImpresor && !bNumeraGrabar) // Numeración
					{
						if (cmbComprobante1.ActiveRow != null && 
							cmbComprobante1.ActiveRow.Cells["idCompraNumero"].Value 
							!= System.DBNull.Value)
							iTipoNumera = (int) cmbComprobante1.ActiveRow.Cells["idCompraNumero"].Value;
						if (iTipoNumera == 0) iTipoNumera = IdTipoFactura;
						string stNumero = Funcion.NumeraFactura(cdsCompraTabla, iTipoNumera);
						if (stNumero.Trim().Length > 0) drCompra["Numero"] = stNumero;
						else
						{
							if (txtNumero.ToString().Trim().Length == 0)
							{
								this.errorProvider.SetError(txtNumero, "Ingrese Codigo");
								return;
							}
						}
					}
				}

				if ((bool) drSeteoF["VerBodega"] == false)
				{
					string ActualizaBodega = "Exec DetCompraBodegaAct " + txtIdCompra.Value.ToString();
					Funcion.EjecutaSQL(cdsArticulo, ActualizaBodega, true);
				}

				barraDato1.IdRegistro = (int) this.txtIdCompra.Value;
				if (IdTipoFactura == 1 && drSeteoF["Promocion"] != DBNull.Value && (bool) drSeteoF["Promocion"])
				{
					string stExec = "Exec ArticuloPromocion " + txtIdCompra.Value.ToString();
					Funcion.EjecutaSQL(cdsArticulo, stExec, true);
					this.cmdTotal();
					barraDato1.GrabaRegistro();
				}
			}
			this.barraDato1.ProximoId(5);
			///////////////// SANTELMO
			if (Funcion.bEjecutaSQL(cdsArticulo, "Exec SeteoGExiste 'FTRSTELMO'"))
			{
				Funcion.EjecutaSQL(cdsCompra, "EXEC CambiaBodega " 
					+ txtIdCompra.Value.ToString() + ", " + IdTipoFactura.ToString());
				this.barraDato1.ProximoId(5);
			}
				
			/////////////////////

			ActualizaCostoUltimo();
			if ((bool) drSeteoF["ActualizarCosto"]) ActualizaPrecio();
			if ((bool) drSeteoF["ConsumoFact"]) 
				Funcion.EjecutaSQL(cdsCompra, "EXEC ConsumoFactura " 
					+ txtIdCompra.Value.ToString() + ", " + IdTipoFactura.ToString());

			if (Funcion.bEjecutaSQL(cdsArticulo, "Exec SeteoGExiste 'FPVCNSP'") && !miAcceso.ReImprimir)
				barraDato1_Impresora(this, e);

		}

		private void txtIdCompra_EditValueChanged(object sender, System.EventArgs e)
		{
			return;
		}

		private void barraDato1_Editar(object sender, System.EventArgs e)
		{
			if (!miAcceso.Autorizacion && chkAutPresupuestaria.Checked)
			{
				MessageBox.Show("No puede Editar despues de Autorizar ",
					"Seguridad", MessageBoxButtons.OK, MessageBoxIcon.Information);
				return;
			}

			if (bAutoImpresor && (bool) drCompra["impreso"] == true)
			{
				MessageBox.Show("No puede Editar en AutoImpresores despues de imprimir ",
					"Seguridad", MessageBoxButtons.OK, MessageBoxIcon.Information);
				return;
			}

			if (drCompra["Bloquear"] != DBNull.Value && (bool) drCompra["Bloquear"] == true)
			{
				MessageBox.Show("No puede editar registro por estar bloqueado", "Informacion",
					MessageBoxButtons.OK, MessageBoxIcon.Information);
				return;
			}

			if (!miAcceso.Editar)
			{
				MessageBox.Show("No tiene Acceso a Editar " + this.Text.Trim(),
					"Seguridad", MessageBoxButtons.OK, MessageBoxIcon.Information);
				return;
			}
			if (this.txtIdCompra.Value == System.DBNull.Value)
			{
				MessageBox.Show("Tiene que existir un registro para poder editar",
					"Edicion", MessageBoxButtons.OK, MessageBoxIcon.Information);
				return;
			}
			if (!miAcceso.EditaImpresa && (bool) drCompra["impreso"] == true)
			{
				MessageBox.Show("No puede Editar Facturas Impresas ",
					"Seguridad", MessageBoxButtons.OK, MessageBoxIcon.Information);
				return;
			}
			if (miAcceso.EditarSoloDia)
			{
				DateTime dtFechaHoy = Funcion.dtEscalarSQL(cdsCliente, "Select GetDate()");
				DateTime dtFechaFac = (DateTime) this.dtFecha.Value;
				if (dtFechaFac.ToShortDateString() != dtFechaHoy.ToShortDateString())
				{
					MessageBox.Show("Solo puede editar Registros de fecha " +
						dtFechaHoy.ToString("dd/MMM/yyyy"), 
						"Información", MessageBoxButtons.OK, MessageBoxIcon.Information);
					return;
				}
			}
			if (miAcceso.EditarDiaIngreso)
			{
				DateTime dtFechaHoy = Funcion.dtEscalarSQL(cdsCliente, "Select GetDate()");
				DateTime dtFechaIng = (DateTime) drCompra["FechaIngreso"];
				if (dtFechaHoy.ToShortDateString() != dtFechaIng.ToShortDateString())
				{	
					MessageBox.Show("Registro Ingresado el " +
						dtFechaIng.ToString("dd/MMM/yyyy") + 
						".\nSolo puede editar Asientos Ingresados el día de Hoy",
						"Información", MessageBoxButtons.OK, MessageBoxIcon.Information);
					return;
				}
			}
			this.barraDato1.EditarRegistro();
			this.HabilitaTab();
			ModificaGrilla(true);
			if (drCompra["idCliente"] != System.DBNull.Value)
				miProveedor = new Proveedor(cdsCliente, (int) drCompra["idCliente"]);
			if (bAutoImpresor) 
			{
				txtNumero.ReadOnly = true;
				dtFecha.ReadOnly = true;
			}
			
		}

		private void barraDato1_Deshacer(object sender, System.EventArgs e)
		{
			this.BorraErrores();
			this.barraDato1.DeshacerRegistro();
			this.HabilitaTab();
			ModificaGrilla(false);
		}

		private void barraDato1_Borrar(object sender, System.EventArgs e)
		{
			//			drCompra = cdsCompra.TableViews["Compra"].Rows[0]; // .Find((int) this.txtIdCompra.Value);
			if ((bool) this.lblAnulado.Visible == false)
			{
//				if (bAutoImpresor)
//				{
//					MessageBox.Show("No puede Anular en AutoImpresores.",
//						"Seguridad", MessageBoxButtons.OK, MessageBoxIcon.Information);
//					return;
//				}
				if (!miAcceso.Anular)
				{
					MessageBox.Show("No tiene acceso a Anular " + this.Text.Trim(),
						"Seguridad", MessageBoxButtons.OK, MessageBoxIcon.Information);
					return;
				}
				if (DialogResult.No == MessageBox.Show(
					"¿Está seguro que desea Anular el Registro?", 
					"Confirmación", MessageBoxButtons.YesNo, MessageBoxIcon.Warning, 
					MessageBoxDefaultButton.Button2)) return;
		
				string stBuscaPago = "Select IsNull(Sum(Pago), 0) From Pago Where idcompra = "
					+ txtIdCompra.Value.ToString();
				double dPagos = Funcion.dEscalarSQL(cdsCompra, stBuscaPago, true);
				if (dPagos > 0)
				{
					MessageBox.Show("No puede Anular Factura porque Tiene Pagos realizados" ,
						"Seguridad", MessageBoxButtons.OK, MessageBoxIcon.Information);
					return;
				}
				this.lblAnulado.Visible = true;
//				string stBorraSQL = "UPDATE Asiento SET Borrar = 1 WHERE idAsiento = "
//					+ txtIdAsiento.Value.ToString();
//				Contabilidad.EjecutaSQL(cdsAsiento, stBorraSQL);
//				stBorraSQL = "UPDATE Retencion SET Borrar = 1 WHERE idCompra = "
//					+ txtIdCompra.Value.ToString();
//				Contabilidad.EjecutaSQL(cdsAsiento, stBorraSQL);
//				this.drCompra["Borrar"] = true;
//				this.cdsCompra.Update();
				string stExecAnula = "Exec CompraBorrar " + txtIdCompra.Value.ToString() + ", 0";
				Funcion.EjecutaSQL(cdsCompra, stExecAnula, true);
				barraDato1.ProximoId(5);
			}
			else // Si esta anulado se elimina
			{
				if (bAutoImpresor)
				{
					MessageBox.Show("No puede Eliminar en AutoImpresores.",
						"Seguridad", MessageBoxButtons.OK, MessageBoxIcon.Information);
					return;
				}
				if (!miAcceso.Eliminar)
				{
					MessageBox.Show("No tiene acceso a Eliminar " + this.Text.Trim(),
						"Seguridad", MessageBoxButtons.OK, MessageBoxIcon.Information);
					return;
				}
				if (DialogResult.No == MessageBox.Show(
					"¿Está seguro que desea borrar el Registro?"
					+ ".\nEste proceso no puede ser revertido", 
					"Confirmación", MessageBoxButtons.YesNo, MessageBoxIcon.Stop, 
					MessageBoxDefaultButton.Button2)) 
				{
					return;
				}

//				Contabilidad.BorrarAsiento(cdsAsiento, (int) txtIdAsiento.Value);
//				Contabilidad.BorrarRetencion(cdsAsiento, (int) txtIdCompra.Value);
//				barraDato1.BorraRegistro();
				string stExecBorrar = "Exec CompraBorrar " + txtIdCompra.Value.ToString() + ", 1";
				Funcion.EjecutaSQL(cdsCompra, stExecBorrar, true);
				barraDato1.ProximoId(2);
			}
		}

		private void BuscarCompra(int idCompra)
		{
			if (idCompra > 0) 
			{
				this.barraDato1.IdRegistro = idCompra;
				this.barraDato1.ProximoId(5);
			}
		}

		private void barraDato1_Buscar(object sender, System.EventArgs e)
		{
			if (!miAcceso.Buscar)
			{
				MessageBox.Show("No tiene acceso a Buscar " + this.Text.Trim(),
					"Seguridad", MessageBoxButtons.OK, MessageBoxIcon.Information);
				return;
			}
			using (BuscaCompra miBusqueda = new BuscaCompra(IdTipoFactura))
			{
				if (DialogResult.OK == miBusqueda.ShowDialog())
				{
					try
					{
						if (miBusqueda.chkOrden.Checked) barraDato1.DataOrden = miBusqueda.stOrden;
						else barraDato1.DataOrden = "";
						int id = int.Parse(miBusqueda.idBusca.Text);
						BuscarCompra(id);
					}
					catch{}
				}
			}
		}

		private void Imprimir(int iTipo)
		{
			if (bNumeraGrabar && (bool)drCompra["impreso"] == false)
			{
				int iTipoNumera = 1;
				if (cmbComprobante1.ActiveRow != null && 
					cmbComprobante1.ActiveRow.Cells["idCompraNumero"].Value 
					!= System.DBNull.Value)
					iTipoNumera = (int) cmbComprobante1.ActiveRow.Cells["idCompraNumero"].Value;
				if (iTipoNumera == 0 || IdTipoFactura == 3) // Diferente de Cotizacion
					iTipoNumera = IdTipoFactura;

        string stNumero = Funcion.sEscalarSQL(cdsCompra, "EXEC FacturaNumeroEstacion " + iTipoNumera.ToString());
				string stSelect = string.Format("Update Compra Set Numero = '{0}' Where idCompra = {1}", 
					stNumero, (int) drCompra["idCompra"]);
				Funcion.EjecutaSQL(cdsArticulo, stSelect);
				drCompra["Numero"] = stNumero;
			}

			DateTime dtIni = DateTime.Now;
			if (!miAcceso.Imprimir)
			{
				MessageBox.Show("No tiene acceso a Imprimir " + this.Text.Trim(),
					"Seguridad", MessageBoxButtons.OK, MessageBoxIcon.Information);
				return;
			}
			if (bAutoImpresor && (bool) drCompra["impreso"] == true)
			{
				MessageBox.Show("No puede ReImprimir en AutoImpresores ",
					"Seguridad", MessageBoxButtons.OK, MessageBoxIcon.Information);
				return;
			}
			
			if (!miAcceso.ReImprimir && (bool) drCompra["impreso"] == true)
			{
				MessageBox.Show("No tiene acceso a ReImprimir " + this.Text.Trim(),
					"Seguridad", MessageBoxButtons.OK, MessageBoxIcon.Information);
				return;
			}
			if (cmbTipoCompra.Value == System.DBNull.Value)
			{
				MessageBox.Show("No hay informacion para imprimir", "Informacion");
				return;
			}

			this.Cursor = Cursors.WaitCursor;
			string stReporte = "Compra.rpt";
			string stFiltro = "{Compra.idCompra} = " + this.txtIdCompra.Value.ToString();
			stFiltro += " And {Compra.idCliente} = " + cmbCliente.Value.ToString();
			switch((int) this.cmbTipoCompra.Value)
			{
				case 1: // Venta
					stReporte = "Venta.Rpt";
					if (cmbComprobante1.ActiveRow.Cells["Reporte"].Value != System.DBNull.Value
						&& cmbComprobante1.ActiveRow.Cells["Reporte"].Value.ToString().Length > 0)
					{
						stReporte = cmbComprobante1.ActiveRow.Cells["Reporte"].Value.ToString().Trim();
					}
					if (bAutoImpresor)
					{
						if ((int) cmbComprobante1.Value == 15) stReporte = "VentaAutoImpresor.Rpt";
						else stReporte = "NotaVentaAutoImpresor.Rpt";
					}
					break;
				case 2: // Orden de Compra
					stReporte = "OrdenCompra.Rpt";
					break;
				case 3: // Cotizacion
					stReporte = "Cotizacion.Rpt";
					break;
				case 4: // Compra
					stReporte = "Compra.Rpt";
					if (bAutoImpresor && (int) cmbComprobante1.Value == 3)  stReporte = "LiquidaCompraAutoImpresor.Rpt";
					break;
				case 5: // Devolucion en venta
					stReporte = "DevolucionVenta.Rpt";
					if (cmbComprobante1.ActiveRow.Cells["Reporte"].Value != System.DBNull.Value)
						stReporte = cmbComprobante1.ActiveRow.Cells["Reporte"].Value.ToString().Trim();
					break;
				case 6: // Devolucion en compra
					stReporte = "DevolucionCompra.Rpt";
					break;
				case 7: // Ingreso de Bodega
					stReporte = "IngresoBodega.Rpt";
					break;
				case 8: // Egreso de Bodega
					stReporte = "EgresoBodega.Rpt";
					break;
				case 9: // Orden de Produccion
					stReporte = "IngresoBodega.Rpt";
					break;
				case 10: // Producto Terminado
					stReporte = "ProductoTerminado.Rpt";
					break;
				case 11: // Reservacion Cliente
					stReporte = "ReservaCliente.Rpt";
					break;
				case 12: // Reservacion Proveedor
					stReporte = "ReservaProveedor.Rpt";
					break;
				case 13: // Pédido Cliente
					stReporte = "PedidoCliente.Rpt";
					break;
				case 14: // Pédido Proveedor
					stReporte = "PedidoProveedor.Rpt";
					break;
				case 26: // Reservacion de Proveedor
					stReporte = "ReservaProveedor.Rpt";
					break;
				case 27: // Nota Debito Venta
					stReporte = "NotaDebitoVenta.Rpt";
					if (bAutoImpresor && (int) cmbComprobante1.Value == 5)  stReporte = "NotaDebitoAutoImpresor.Rpt";
					break;
				case 28: // Nota de Debito Compra
					stReporte = "NotaDebitoCompra.Rpt";
					break;
				case 36: // Requisicion
					stReporte = "Requisicion.Rpt";
					break;
				default:
					stReporte = "TransaccionPredefinida.Rpt";
//					MessageBox.Show("Tipo de Factura Desconocida");
					break;
			}
			if (drSeteoF["SerieLinea"] != DBNull.Value && (bool) drSeteoF["SerieLinea"] == true)
			{
				string stExec = "Exec SerieLinea " + txtIdCompra.Value.ToString();
				Funcion.EjecutaSQL(cdsCompra, stExec, true);
			}

			string stProc = "Exec SerieExisteTipo " + txtIdCompra.Value.ToString() + ", 'FPVTISP', "
				+ cmbComprobante1.Value.ToString();
			bool bProc = Funcion.bEjecutaSQL(cdsArticulo, stProc);
			if (bProc == true) 
			{
				stFiltro = "";
				stReporte = stReporte.Substring(0, stReporte.Length - 4) + "SP.rpt";
			}
			using (Reporte miReporte = new Reporte(stReporte, stFiltro))
			{
				if (bProc == true)
				{
					// Definicion de Parametros
					ParameterFields paramFields = new ParameterFields ();
					ParameterField paramField = new ParameterField ();

					ParameterDiscreteValue discreteVal = new ParameterDiscreteValue ();
					paramField.ParameterFieldName = "@idCompra";
					discreteVal.Value = txtIdCompra.Value.ToString();
					paramField.CurrentValues.Add (discreteVal);
					paramFields.Add (paramField);

					miReporte.crVista.ParameterFieldInfo = paramFields;

					// Fin de parametros
				}

				if (iTipo == 0)
					miReporte.ShowDialog();
				if (iTipo == 1)
				{
					string stImpresora = MenuLatinium.ImpresoraNombre;
					if (stReporte == "NotaVenta.rpt" && 
						Funcion.iEscalarSQL(cdsArticulo, "Select Count(*) From SeteoG Where Codigo = 'FPVI' + Host_Name() And Activado = 1") > 0)
					{
						MenuLatinium.ImpresoraNombre = Funcion.sEscalarSQL(cdsArticulo,
              "Select Nombre From SeteoG Where Codigo = 'FPVI' + Host_Name() And Activado = 1");
//						MessageBox.Show(stImpresora, "NotaVenta y FPVI");
						miReporte.ImprimeRep(iTipo);
						MenuLatinium.ImpresoraNombre = stImpresora;
					}
					else
					{
//						MessageBox.Show(stImpresora, "Predefinida");
						miReporte.ImprimeRep(iTipo);
					}
				}
				if (iTipo > 1)
				{
					miReporte.iNumeroCopias = iTipo;
					miReporte.ImprimeRep(1);
				}
			}
			string stGrabaImprime = "Update Compra Set Impreso = 1 "
				+ "Where idCompra = " + txtIdCompra.Value.ToString();
			Funcion.EjecutaSQL(cdsCompra, stGrabaImprime);
			drCompra["impreso"] = true;
			DateTime dtFin = DateTime.Now;
			TimeSpan ts = dtFin.Subtract(dtIni);
			lblTiempo.Text = ts.Seconds.ToString() + ":" + ts.Milliseconds.ToString();
			this.Cursor = Cursors.Default;
		}

		private void barraDato1_Imprimir(object sender, System.EventArgs e)
		{
			this.Imprimir(0);
		}

		C1.Data.C1DataRow drCompra;
		private void cdsCompra_PositionChanged(object sender, C1.Data.PositionChangeEventArgs e)
		{
			if (e.DataMember.ToString() == "_Compra")
			{
				drCompra = e.Row;
				this.lblAnulado.Visible = (bool) drCompra["Borrar"];
				if ((int) e.Row["idTipoFactura"] < 1)	
				{
					e.Row["idTipoFactura"] = this.IdTipoFactura;
				}
			}
		}

		private int IdSigno = 2;
		private Acceso miAcceso;
		private C1DataRow drSeteoF;
		private C1DataRow drSeteoC;
		private C1DataRow drSeteo;

		private void VisibleBotones()
		{
			if (IdTipoFactura == 4 || IdTipoFactura == 8 || IdTipoFactura == 9) // No es compra
			{
				if ((bool)drSeteoF["CompraUltimaC"] && !(bool) drSeteoF["ActualizarCosto"])
					this.btActualizaPrecio.Visible = true;
			}
			switch (this.IdTipoFactura)
			{
				case 1: // Venta 
				case 6: // Devolucion en compra
				case 4: // Compra
				case 5: // Devolucion en venta
				case 8: // Ingreso de Bodega
				case 9: // Egreso de Bodega
				case 10: // Consumo Real
				case 12: // Producto Terminado
				case 27: // Nota debito venta
				case 11: // Reservacion Cliente
				case 26: // Reservacion Proveedor
				case 28: // Nota debito compra
					break;
				default:
					this.btCreaAsiento.Visible = false;
					break;
			}
			switch (this.IdTipoFactura)
			{
				case 1: // Venta 
				case 4: // Compra
					this.btRetencion.Visible = true;
					break;
			}

			switch (this.IdTipoFactura)
			{
				case 1: // Venta 
				case 6: // Devolucion en compra
				case 4: // Compra
				case 5: // Devolucion en venta
				case 8: // Egreso de Bodega
					this.btGuiaRemision.Visible = true;
					break;
			}

			if (this.IdTipoFactura == 6 || this.IdTipoFactura == 5) // Devolucion en compra o venta
				this.btTransforma.Visible = false;

			this.txtIdCompra.Width=0;
			this.txtBorrar.Width = 0;
			//			this.txtIdAsiento.Width = 0;

			// Posicion de Botones
			int iPosDer = this.btBuscaArticulo.Left + this.btBuscaArticulo.Width;
			if (this.btTransforma.Visible)
			{
				this.btTransforma.Left = iPosDer;
				iPosDer += this.btTransforma.Width;
			}
			if (this.btNotaEntrega.Visible)
			{
				this.btNotaEntrega.Left = iPosDer;
				iPosDer += this.btNotaEntrega.Width;
			}
			if (this.btRetencion.Visible)
			{
				this.btRetencion.Left = iPosDer;
				iPosDer += this.btRetencion.Width;
			}
			if (this.btCreaAsiento.Visible)
			{
				this.btCreaAsiento.Left = iPosDer;
				iPosDer += this.btCreaAsiento.Width;
			}
			if (this.btGuiaRemision.Visible)
			{
				this.btGuiaRemision.Left = iPosDer;
				iPosDer += this.btGuiaRemision.Width;
			}
			if (this.btActualizaPrecio.Visible)
			{
				this.btActualizaPrecio.Left = iPosDer;
				iPosDer += this.btActualizaPrecio.Width;
			}
			if (this.btNuevo.Visible)
			{
				this.btNuevo.Left = iPosDer;
				iPosDer += this.btNuevo.Width;
			}
			if (this.btBarras.Visible)
			{
				this.btBarras.Left = iPosDer;
				iPosDer += this.btBarras.Width;
			}
			if (this.btAuditoria.Visible)
			{
				this.btAuditoria.Left = iPosDer;
				iPosDer += this.btAuditoria.Width;
			}
		}
	
		private void HabilitaTab()
		{
			foreach(Infragistics.Win.UltraWinTabControl.UltraTab ctr in utcTransaccion.Tabs)
				foreach(Control cn in ctr.TabPage.Controls)
					this.barraDato1.HabilitaControl(cn, this.barraDato1.bEditar);
			this.rdgContadoCred.Enabled = this.barraDato1.bEditar;
		}

		private bool bIncluyeIva = false;
		struct Proveedor
		{
			public DateTime FechaCaducidad;
			public string SerieFactura;
			public string AutorizaImprenta;
			public string AutorizaFactura;
			public int DiasCredito;
			public int idVendedor;
			public int idCredTribut;
			public double MaximoCredito;
			public double DescuentoCredito;
			public double DescuentoContado;
			public int idTipoPrecio;
			public bool InstPublica;

			public Proveedor(C1.Data.C1DataSet ds, int idCliente)
			{
				string stSelect = "Select IsNull(FechaCaducidad, '19000101'), IsNull(SerieFactura, ''), "
					+ "IsNull(AutorizaImprenta, ''), IsNull(AutorizaFactura, ''), "
					+ "IsNull(DiasCredito, 0), IsNull(idVendedor, 0), "
					+ "IsNull(idCredTribut, 0), IsNull(MaximoCredito, 0), "
					+ "IsNull(DescuentoCredito, 0), IsNull(DescuentoContado, 0), "
					+ "IsNull(idTipoPrecio, 0), IsNull(instPublica, 0) "
					+ "From Cliente Where idCliente = " + idCliente.ToString();
				SqlDataReader miReader = null;
				this.FechaCaducidad = DateTime.Today;
				this.SerieFactura = "001001";
				this.AutorizaImprenta = "";
				this.AutorizaFactura = "";
				this.DiasCredito = 0;
				this.idVendedor = 0;
				this.idCredTribut = 1;
				this.MaximoCredito = 0;
				this.DescuentoCredito = 0;
				this.DescuentoContado = 0;
				this.idTipoPrecio = 1;
				this.InstPublica = false;
				try
				{
					miReader = Funcion.rEscalarSQL(ds, stSelect, true);
					miReader.Read();
					this.FechaCaducidad = miReader.GetDateTime(0);
					this.SerieFactura = miReader.GetString(1);
					this.AutorizaImprenta = miReader.GetString(2);
					this.AutorizaFactura = miReader.GetString(3);
					this.DiasCredito = miReader.GetInt32(4);
					this.idVendedor = miReader.GetInt32(5);
					this.idCredTribut = miReader.GetInt32(6);
					this.MaximoCredito = miReader.GetDouble(7);
					this.DescuentoCredito = miReader.GetDouble(8);
					this.DescuentoContado = miReader.GetDouble(9);
					this.idTipoPrecio = miReader.GetInt32(10);
					this.InstPublica = miReader.GetBoolean(11);
				}
				catch{}
				finally
				{
					miReader.Close();
				}
			}
		}

		private Proveedor miProveedor;
		
		private Items miArt;
		struct Items
		{
			public double Iva;
			public double CostoPromedio;
			public double CostoUltimo;
			public double[] Precio;
			public double[] Cantidad;

			public Items(C1.Data.C1DataSet ds, int IdArticulo)
			{
				string stBuscaArt = "Select IsNull(Iva, 0), IsNull(CostoPromedio, 0), "
					+ "IsNull(Precio1, 0), "
					+ "IsNull(Precio2, 0), IsNull(Precio3, 0), IsNUll(Precio4, 0), IsNull(Precio5, 0), "
					+ "IsNull(Cantidad1, 0), IsNull(Cantidad2, 0), IsNUll(Cantidad3, 0), "
					+ "IsNull(Cantidad4, 0), IsNull(Cantidad5, 0), IsNUll(CostoUltimo, 0) "
					+ "From Articulo Where idArticulo = " + IdArticulo.ToString();
				SqlDataReader miReader = Funcion.rEscalarSQL(ds, stBuscaArt, true);
				miReader.Read();
				this.Precio = new double[6];
				this.Cantidad = new double[6];
				this.Iva = miReader.GetDouble(0);
				this.CostoPromedio = miReader.GetDouble(1);
				this.Precio[1] = miReader.GetDouble(2);
				this.Precio[2] = miReader.GetDouble(3);
				this.Precio[3] = miReader.GetDouble(4);
				this.Precio[4] = miReader.GetDouble(5);
				this.Precio[5] = miReader.GetDouble(6);
				this.Cantidad[1]= miReader.GetInt32(7);
				this.Cantidad[2]= miReader.GetInt32(8);
				this.Cantidad[3]= miReader.GetInt32(9);
				this.Cantidad[4]= miReader.GetInt32(10);
				this.Cantidad[5]= miReader.GetInt32(11);
				this.CostoUltimo = miReader.GetDouble(12);
				miReader.Close();
			}
		}
		int IProyecto = 0;
		int ISubProyecto = 0;
		int IBodega = 0;
		private void CargaProyecto()
		{
			string stExec = "Exec ProyectoPredefinido " + IdTipoFactura.ToString();
			IProyecto = Funcion.iEscalarSQL(cdsCompra, stExec + ", 1", true);
			ISubProyecto = Funcion.iEscalarSQL(cdsCompra, stExec + ", 2", true);
			IBodega = Funcion.iEscalarSQL(cdsCompra, stExec + ", 3", true);

			if (IProyecto > 0) cmbProyecto.ReadOnly = true;
			if (ISubProyecto > 0) cmbSubProyecto.ReadOnly = true;
			if (IBodega > 0) spnBodega.ReadOnly = true;
		}

		private int iCuentaTotal = 0;
		private void MoverTotales()
		{
			lblIce.Top = lblIce.Top + 16;
			lblIva.Top = lblIva.Top + 16;
			lblIva0.Top = lblIva0.Top + 16;
			lblIvaSubt.Top = lblIvaSubt.Top + 16;
			lblSubTotal.Top = lblSubTotal.Top + 16;
			lblDescuento.Top = lblDescuento.Top + 16;
			txtIceTotal.Top = txtIceTotal.Top + 16;
			txtIvaTotal.Top = txtIvaTotal.Top + 16;
			txtIva.Top = txtIva.Top + 16;
			txtIva0.Top = txtIva0.Top + 16;
			txtDescIva0Total.Top = txtDescIva0Total.Top + 16;
			txtDescIvaTotal.Top = txtDescIvaTotal.Top + 16;

			lblServicio.Visible  = false;
			lblPropina.Visible = false;
			txtServicio.Width = 0;
			txtPropina.Width = 0;
		}
		public DateTime FechaCaducidadSri;
		private void Compra2_Load(object sender, System.EventArgs e)
		{			
			///////////////////////////////////////////////////////////
			//Cambios para Botrosa 
			//agrandar el campo de bodegas para ver el nombre			
			// agrandar campo de vendedor
			//Filtrar en articulos que se escoja solo los de Inventarios
			//en Ingresos y egresos de BOdega
			//									if (Funcion.bEjecutaSQL(cdsArticulo, "Exec SeteoGExiste 'FPAI'")) 
			//									{if (IdTipoFactura == 8 || IdTipoFactura == 9)
			//										{ this.cdvArtCodigo.RowFilter = "idTipoGrupo In (1)";
			//											this.cdvArticulo.RowFilter = "idTipoGrupo In (1)";									}									}
			//poner como obligatorio ingresar Centro de costo: en la funcion grabar llamar a funcion CambiosBotrosa
			////funcoion para proyectos por centro de costo
			if (Funcion.bEjecutaSQL(cdsArticulo, "Exec SeteoGExiste 'FTRPBP'")) 				
			{
				this.ultraGrid1.DisplayLayout.Bands[0].Columns["idProyecto"].Hidden = true;										
				this.ultraGrid1.DisplayLayout.Bands[0].Columns["idSubProyecto"].Hidden = true;				
				this.ultraGrid1.DisplayLayout.Bands[0].Columns["ProyectoBR"].Hidden = false;										
				this.ultraGrid1.DisplayLayout.Bands[0].Columns["CentroCostoBr"].Hidden = false;
			}
			if (Funcion.bEjecutaSQL(cdsArticulo, "Exec SeteoGExiste 'FVPSA'")) 				
				bServicioAuto = true;

			//////////////////////////////////////////////////////////
			if (Funcion.bEjecutaSQL(cdsArticulo, "Exec SeteoGExiste 'FTRPB'")) 				
			{				
				//crear en el grid UnidadBotrosa y un crear un cmbUnidadB y enlazar 
				this.ultraGrid1.DisplayLayout.Bands[0].Columns["UnidadBotrosa"].Hidden = false;						
				//craer un cmbBodegaLinea se asigna y se hace el enlace 			
				this.ultraGrid1.DisplayLayout.Bands[0].Columns["Bodega"].ValueList = cmbBodegaLinea;
			}
			//////////////////////////////////////////////////////////
      /////cambio STI
      if (Funcion.bEjecutaSQL(cdsArticulo, "Exec SeteoGExiste 'FTRSTI'"))
      {
          this.ultraGrid1.DisplayLayout.Bands[0].Columns["RefNumero"].Header.Caption = "Numero OP";
      }


			if (Funcion.bEjecutaSQL(cdsArticulo, "Exec SeteoGExiste 'FTRNBG'"))
				if (IdTipoFactura == 1)
					bNumeraGrabar = true;



			cmbBodega.Width = 0;
			if (drSeteoF["VerBodegaNombre"] != DBNull.Value && (bool) drSeteoF["VerBodegaNombre"] == true)
			{
				udsBodegas.Rows.Clear();
				string stExecBodega = "Exec BodegaUsuario "	+ IdTipoFactura.ToString();
				SqlDataReader drBod = Funcion.rEscalarSQL(cdsArticulo, stExecBodega, true);
				int iNumFilas = 0;

				while(drBod.Read())
				{				
					udsBodegas.Rows.Add();
					udsBodegas.Rows[iNumFilas]["Nombre"]=drBod.GetString(0);
					udsBodegas.Rows[iNumFilas]["Bodega"]=drBod.GetInt32(1);
					iNumFilas++;
				}
				drBod.Close();

				if (iNumFilas > 0)
				{
					cmbBodega.Width = 110;
					cmbBodega.Left = lblBodega.Left-6;
					spnBodega.Width = 0;
					//lblBodega.Width = 0;
					lblBodega.Visible=false;
				}
			}
			Cursor = Cursors.Default;
			if (MenuLatinium.bAcademica)
			{
				string stCuenta = "Select Count(*) From Compra";
				iCuentaTotal = Funcion.iEscalarSQL(cdsAsiento, stCuenta, true);
			}

			this.txtIdAsiento.Width = 0;
			string stTipo = "04";
			#region Tipo de Factura
			switch( this.IdTipoFactura)
			{
				case 1:
				case 11:
				case 27:
					stTipo += "03";
					break;
				case 3:
					stTipo += "04";
					break;
				case 4:
				case 28:
					stTipo += "01";
					break;
				case 2:
				case 5:
				case 6:
					stTipo += "0" + this.IdTipoFactura.ToString().Trim();
					break;
				case 8:
				case 9:
					int idT = IdTipoFactura - 1;
					stTipo += "0" + idT.ToString().Trim();
					break;
				case 10:
					stTipo = "MP04";
					break;
				case 12:
					stTipo = "MP03";
					break;
				case 31:
					stTipo = "0408";
					break;
				case 99:
					stTipo = "04090L9";
					break;
				default:
					stTipo = "0401";
					break;
			}
			#endregion Tipo de Factura

			miAcceso =new Acceso(stTipo, cdsSeguridad.StorageDataSet.Tables["Seguridad"]);
			if (!miAcceso.Autorizacion) 
				chkAutPresupuestaria.Visible = false;

			if (!miAcceso.Buscar) 
			{
				barraDato1.BarraMovimiento = false;
			}
			if (miAcceso.ListaPrecios) 
			{
				cmbPrecio.Visible = true;
				btCambiarPrecio.Visible = true;
			}

			CargaProyecto();
			this.VisibleBotones();

			this.barraDato1.stFiltroExt = " And idTipoFactura = " + this.IdTipoFactura.ToString();
			this.barraDato1.PosUltima();
			int iLong = cmbTipoCompra.Text.Trim().Length;
			if (iLong > 20) iLong = 20;
			string stAudita = "Exec AuditaCrear 53, 6, '" + cmbTipoCompra.Text.Trim().Substring(0, iLong) + "'";
			Funcion.EjecutaSQL(cdsArticulo, stAudita, true);

			// Ubicacion inicial en el registro
			if (this.IdBusqueda > 0)
				this.BuscarCompra(IdBusqueda);
			else
				this.barraDato1.PosUltima();
			this.barraDato1.HabilitaControles(false);
			this.HabilitaTab();
			ModificaGrilla(false);

			string stSelect = "SELECT Signo FROM CompraNumero WHERE idTipoFactura = " 
				+ IdTipoFactura.ToString();
			IdSigno = Funcion.iEscalarSQL(cdsCompra, stSelect);
			if (IdTipoFactura == 2 || IdTipoFactura == 4 || IdTipoFactura == 25
				|| IdTipoFactura == 26 || IdTipoFactura == 6 || IdTipoFactura == 9
				|| IdTipoFactura == 14)
			{
				lblCliente.Text = "Proveed.";
			}
			this.Text = this.cmbTipoCompra.Text;

			if (!miAcceso.CambiarFecha) dtFecha.ReadOnly = true;
			if (!miAcceso.CambiarNumeracion) txtNumero.ReadOnly = true;

			drSeteoF = this.cdsSeteoF.TableViews["SeteoF"].Rows[0];
			drSeteoC = this.cdsSeteoC.TableViews["SeteoC"].Rows[0];
			drSeteo = this.cdsSeteo.TableViews["Seteo"].Rows[0];
			// Filtro de Articulos
			if (IdTipoFactura == 19 || IdTipoFactura == 20 || IdTipoFactura == 12)
			{
				this.cdvArtCodigo.RowFilter = "idTipoGrupo In (6, 8)";
				this.cdvArticulo.RowFilter = "idTipoGrupo In (6, 8)";
			}

			//Filtrar en articulos que se escoja solo los de Inventarios
			//en Ingresos y egresos de BOdega
			if (Funcion.bEjecutaSQL(cdsArticulo, "Exec SeteoGExiste 'FPAI'")) 
			{
				if (IdTipoFactura == 8 || IdTipoFactura == 9)
				{
					this.cdvArtCodigo.RowFilter = "idTipoGrupo In (1)";
					this.cdvArticulo.RowFilter = "idTipoGrupo In (1)";
				}
			}
			///fin
			if (true)
			{
				string stExecFiltro = "Exec ArticuloFiltrado " + IdTipoFactura.ToString();
				string stFiltro = Funcion.sEscalarSQL(cdsArticulo, stExecFiltro, true);
				if (stFiltro.Length > 0)
				{
					this.cdvArtCodigo.RowFilter = stFiltro;
					this.cdvArticulo.RowFilter = stFiltro;
				}
			}


			if (drSeteoF["VerServicio"] == DBNull.Value || (bool) drSeteoF["VerServicio"] == false)
				MoverTotales();

			if ((bool) drSeteoF["IncluyeIva"] && (IdTipoFactura == 1 
				|| IdTipoFactura == 3 || IdTipoFactura == 5 
				|| IdTipoFactura == 8 || IdTipoFactura == 9)) bIncluyeIva = true;

			stSelect = "SELECT Count(*) FROM FACTURANUMERO WHERE ESTACION = Host_Name() "
				+ "AND IDTIPOFACTURA = " +  IdTipoFactura.ToString();
			int iCuentaAutoImp = Funcion.iEscalarSQL(cdsArticulo, stSelect, true);
			if (iCuentaAutoImp == 1)
			{
				stSelect = "SELECT IsNull(AutoImpresor, 0) FROM FACTURANUMERO WHERE ESTACION = Host_Name() "
					+ "AND IDTIPOFACTURA = " +  IdTipoFactura.ToString();
				bAutoImpresor = Funcion.bEjecutaSQL(cdsArticulo, stSelect);
			}
			// En ventas hay nota de venta y factura
			if (iCuentaAutoImp == 2 && IdTipoFactura == 1)
			{
				stSelect = "SELECT IsNull(AutoImpresor, 0) FROM FACTURANUMERO WHERE ESTACION = Host_Name() "
					+ "AND IDTIPOFACTURA = " +  IdTipoFactura.ToString() + " And idDocumento In (1, 2)";
				bAutoImpresor = Funcion.bEjecutaSQL(cdsArticulo, stSelect);
				iCuentaAutoImp = 1;
			}
			if (iCuentaAutoImp > 1)
			{
				MessageBox.Show("Numeracion de autoimpresor esta duplicada", "Informacion");
			}

			if (bAutoImpresor)
			{
				dtFecha.ReadOnly = true;
				txtNumero.ReadOnly = true;
				cmbCaducidad.ReadOnly = true;
				txtAutFactura.ReadOnly = true;
				string stExecCaducado = "Exec SriAutorizacionCaducada " + IdTipoFactura.ToString();

				int iDiasCaducidad = Funcion.iEscalarSQL(cdsArticulo, stExecCaducado, true);
				FechaCaducidadSri = DateTime.Today.AddDays(iDiasCaducidad);
				if (iDiasCaducidad <= 0)
				{
					MessageBox.Show("Autorizacion Caducada", "Informacion");
					barraDato1.VisibleNuevo = false;
					Close();
				}
				if (iDiasCaducidad < 30)
				{
					MessageBox.Show("Faltan " + iDiasCaducidad.ToString()
						+ " dias para que caduque autorizacion");
				}
			}

//			if (drSeteo["AutoImpresor"] != DBNull.Value && (bool) drSeteo["AutoImpresor"] == true)
			// Valores por omision de Grilla al crear nueva linea
			string stActivado = "Select IsNull(Activado, 0) From SeteoG Where Codigo = 'I'";
			if (Funcion.bEjecutaSQL(cdsArticulo, stActivado))// ultraTextEditor4.ReadOnly = true;
		{
			ultraTextEditor4.ReadOnly = true;
			label14.Text= "BIO";
		}
			

			if (!miAcceso.CambiarPrecio && (IdTipoFactura == 1 || IdTipoFactura == 3))
				this.ultraGrid1.DisplayLayout.Bands[0].Columns["Precio"].CellActivation = 
					Infragistics.Win.UltraWinGrid.Activation.Disabled;
			if (!miAcceso.CambiarDescuento && (IdTipoFactura == 1 || IdTipoFactura == 3))
			{
				ultraGrid1.DisplayLayout.Bands[0].Columns["DescuentoArt"].CellActivation = 
					Infragistics.Win.UltraWinGrid.Activation.NoEdit;
				ultraGrid1.DisplayLayout.Bands[0].Columns["DescuentoPorc"].CellActivation = 
					Infragistics.Win.UltraWinGrid.Activation.NoEdit;
				txtDescIva.ReadOnly = true;
				txtDescIva0.ReadOnly = true;
				txtDescPorcIva.ReadOnly = true;
				txtDescPorcIva0.ReadOnly = true;
			}			
		}

		private void ModificaGrilla(bool bModifica)
		{
			ultraGrid1.Enabled = true;
			if (bModifica)
			{
				ultraGrid1.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
				ultraGrid1.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.True;
			}
			else
			{
				ultraGrid1.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.False;
				ultraGrid1.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			}
		}

		private string stbProveedor = "0";
		bool bListaCliente = true;
		private void cdsCliente_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			this.cdsCliente.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);		

			if (Funcion.EsProveedor(IdTipoFactura))
				stbProveedor = "True";
			else
				stbProveedor = "False";
			C1.Data.SchemaObjects.DataSetDef dataSetDef = e.DataSet.Schema.DataSetDefs["dsClientes"];

			if (stbProveedor == "True") stbProveedor = " <> 0";
			else stbProveedor = " = 0";

			string stFiltro = "(Ambos <> 0 OR Proveedor " + stbProveedor;
			if (stbProveedor != " = 0") stFiltro += ") And DesconProv = 0 ";
			else stFiltro += ") And DesconCliente = 0 ";

			string stSelect = "Select IsNull(VerNombreLista, 1) From SeteoF ";
			bListaCliente = Funcion.bEscalarSQL(cdsCliente, stSelect, true);
			if (!bListaCliente)
				stFiltro = "idCliente < 2";

			e.Filter.Add(new C1.Data.FilterCondition(dataSetDef.TableViews["Cliente"],
				stFiltro));
		}

		private void cdsCompra_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			int iSeg = (int) cdsCompra.Schema.Connections[0].CommandTimeout;
			cdsCompra.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);
			iSeg = (int) cdsCompra.Schema.Connections[0].CommandTimeout;
		}

		private void ultraGrid1_InitializeRow(object sender, Infragistics.Win.UltraWinGrid.InitializeRowEventArgs e)
		{
			e.Row.Cells["CodArt"].Value = e.Row.Cells["IdArticulo"].Value;
			e.Row.Cells["UnidadBotrosa"].Value = e.Row.Cells["IdUnidad"].Value;
		}

		private bool bActualiza = true;

		string controlW="";
		string controlW1="";
		string NuevoCodigo="";
		private void ultraGrid1_AfterCellUpdate(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			////el siguiente codigo es para filtrar cc por proy
			if (Funcion.bEjecutaSQL(cdsArticulo, "Exec SeteoGExiste 'FTRPBP'")) 				
			{
				if (e.Cell.IsInEditMode)
				{
					if (e.Cell.Column.ToString() == "CentroCostoBR")
					{
						if(controlW != e.Cell.Row.Cells["CentroCostoBR"].Value.ToString())
						{
							controlCCP();
							controlW = e.Cell.Row.Cells["CentroCostoBR"].Value.ToString();							
							NuevoCodigo = codCC(controlW);
							NuevoCodigo=NuevoCodigo.Substring(0,5);		
							string stSelect = "Select idProyecto From Proyecto Where CentroCosto like '"
								+ NuevoCodigo + "%'";
							int IdProy = Funcion.iEscalarSQL(cdsArticulo, stSelect, true);
							e.Cell.Row.Cells["idProyecto"].Value = IdProy;
						}
						else return;
					}
				}
				if (e.Cell.IsInEditMode)
				{
					if (e.Cell.Column.ToString() == "ProyectoBR")
					{
						if(ultraGrid1.Rows.Count < 1 )return;
						string CodCentroCosto = ultraGrid1.ActiveRow.Cells["CentroCostoBR"].Value.ToString();							
						string NuevoCodigo1 = codCC(CodCentroCosto);
						NuevoCodigo1=NuevoCodigo1.Substring(0,5);	

						if(controlW1 != e.Cell.Row.Cells["ProyectoBR"].Value.ToString())
						{
							controlCPP(NuevoCodigo1);							
							controlW1 = e.Cell.Row.Cells["ProyectoBR"].Value.ToString();

							string stSelect1 = "Select idSubProyecto From subProyecto Where nombre like '"
								+ controlW1 + "%'";
							int IdProy1 = Funcion.iEscalarSQL(cdsArticulo, stSelect1, true);
							e.Cell.Row.Cells["idSubProyecto"].Value = IdProy1;
						}
						else return;
					}
				}
			}
			/////


			if (!bActualiza) return;

			if (e.Cell.IsInEditMode)
				if (e.Cell.Column.ToString() == "CodArt")
					if (e.Cell.OriginalValue != e.Cell.Value)
					{
						e.Cell.Row.Cells["idArticulo"].Value = e.Cell.Row.Cells["CodArt"].Value;
					}

			bActualiza = false;

			string stFiltro = "";
			if (IdTipoFactura == 1) stFiltro = " And IsNull(dVenta, 0) = 0";
			if (IdTipoFactura == 4) stFiltro = " And IsNull(dCompra, 0) = 0";
			if (e.Cell.Column.ToString() == "Codigo")
			{
				string stCod = e.Cell.Row.Cells["Codigo"].Value.ToString();
				if (Funcion.bEjecutaSQL(cdsArticulo, "Exec SeteoGExiste 'FTRCxG'")) // cambia la comilla por guion				
				{						
					string cadena = revisaCadena(stCod);
					stCod = cadena;												
				}
			
				string stSelect = "Select Count(*) From Articulo Where Codigo like '"
					+ stCod + "%'" + stFiltro;
				int iCuenta = Funcion.iEscalarSQL(cdsArticulo, stSelect, true);
				int IdArt = 0;
				if (iCuenta != 1)
				{
					using (BuscaArticulos miBusqueda = new BuscaArticulos(stCod, IdTipoFactura))
					{
						if (DialogResult.OK == miBusqueda.ShowDialog())
						{
							IdArt = (int) miBusqueda.txtIdBusca.Value;
							e.Cell.Row.Cells["Articulo"].Value = miBusqueda.txtDescripcion.Text;
							e.Cell.Row.Cells["Codigo"].Value = miBusqueda.txtCodigo.Text;
							if ((bool) drSeteoF["VerRefNumero"] == true)
								e.Cell.Row.Cells["RefNumero"].Value = miBusqueda.txtUbicacion.Text;
							e.Cell.Row.Cells["idUnidad"].Value = int.Parse(miBusqueda.txtIdUnidad.Text);
						}
						else
						{
							bActualiza = true;
							return;						
						}
					}
				}
				else
				{
					stSelect = "Select idArticulo From Articulo Where Codigo like '"
						+ stCod + "%'" + stFiltro;
					IdArt = Funcion.iEscalarSQL(cdsArticulo, stSelect, true);
					stSelect = "Select Codigo From Articulo Where Codigo like '" + stCod + "%'";
					e.Cell.Row.Cells["Codigo"].Value = Funcion.sEscalarSQL(cdsArticulo, stSelect, true);
					stSelect = "Select Articulo From Articulo Where Codigo like '" + stCod + "%'";
					e.Cell.Row.Cells["Articulo"].Value = Funcion.sEscalarSQL(cdsArticulo, stSelect, true);
					stSelect = "Select idUnidad From Articulo Where Codigo like '" + stCod + "%'";
					e.Cell.Row.Cells["idUnidad"].Value = Funcion.iEscalarSQL(cdsArticulo, stSelect, true);
				}
				e.Cell.Row.Cells["idArticulo"].Value = IdArt;
				if (!Limites(e)) 
				{
					bActualiza = true;
					return;
				}
				CambiaCantidad(e);
				Precio(e);
				DescuentoPorcentaje(e);
				Promocion(e);
				cmdTotal();
			}
			if (e.Cell.Column.ToString() == "Articulo")
			{
				string stCod = e.Cell.Row.Cells["Articulo"].Value.ToString();
				string stSelect = "Select Count(*) From Articulo Where Articulo like '"
					+ stCod + "%'";
				int iCuenta = Funcion.iEscalarSQL(cdsArticulo, stSelect, true);
				int IdArt = 0;
				if (iCuenta != 1)
				{
					using (BuscaArticulos miBusqueda = new BuscaArticulos(stCod, IdTipoFactura))
					{
						if (DialogResult.OK == miBusqueda.ShowDialog())
						{
							IdArt = (int) miBusqueda.txtIdBusca.Value;
							e.Cell.Row.Cells["Articulo"].Value = miBusqueda.txtDescripcion.Text;
							e.Cell.Row.Cells["Codigo"].Value = miBusqueda.txtCodigo.Text;
							e.Cell.Row.Cells["idUnidad"].Value = int.Parse(miBusqueda.txtIdUnidad.Text);
						}
						else
						{
							bActualiza = true;
							return;						
						}
					}
				}
				else
				{
					stSelect = "Select idArticulo From Articulo Where Articulo like '"
						+ stCod + "%'";
					IdArt = Funcion.iEscalarSQL(cdsArticulo, stSelect, true);
					stSelect = "Select Codigo From Articulo Where Articulo like '" + stCod + "%'";
					e.Cell.Row.Cells["Codigo"].Value = Funcion.sEscalarSQL(cdsArticulo, stSelect, true);
					stSelect = "Select Articulo From Articulo Where Articulo like '" + stCod + "%'";
					e.Cell.Row.Cells["Articulo"].Value = Funcion.sEscalarSQL(cdsArticulo, stSelect, true);
					stSelect = "Select idUnidad From Articulo Where Articulo like '" + stCod + "%'";
					e.Cell.Row.Cells["idUnidad"].Value = Funcion.iEscalarSQL(cdsArticulo, stSelect, true);
				}
				e.Cell.Row.Cells["idArticulo"].Value = IdArt;
				if (!Limites(e)) 
				{
					bActualiza = true;
					return;
				}
				CambiaCantidad(e);
				Precio(e);
				DescuentoPorcentaje(e);
				Promocion(e);
				cmdTotal();
			}

			if (e.Cell.Column.ToString() == "idArticulo")
			{
				if (!Limites(e)) 
				{
					bActualiza = true;
					return;
				}
				Precio(e);
				DescuentoPorcentaje(e);
				Promocion(e);
				int IdArtU = (int) e.Cell.Row.Cells["idArticulo"].Value;
				string stSelectU = "Select idUnidad From Articulo Where idArticulo = " + IdArtU.ToString();
				e.Cell.Row.Cells["idUnidad"].Value = Funcion.iEscalarSQL(cdsArticulo, stSelectU, true);
				cmdTotal();
			}

			if (e.Cell.Column.ToString() == "idUnidad" || e.Cell.Column.ToString() == "Unidades")
			{
				if (e.Cell.OriginalValue == e.Cell.Value) return;
				CambiaCantidad(e);
				Precio(e);
				if (!Limites(e)) 
				{
					bActualiza = true;
					return;
				}
				DescuentoPorcentaje(e);
				Promocion(e);
				cmdTotal();
			}


			if (e.Cell.Column.ToString() == "Cantidad")
			{
				if (!Limites(e)) 
				{
					bActualiza = true;
					return;
				}
				// Cambio de precio por cantidad
				if ((int) this.drSeteoF["TipoPrecio"] == 1 || (int) this.drSeteoF["TipoPrecio"] == 3) Precio(e);
				Promocion(e);
				DescuentoPorcentaje(e);

				cmdTotal();
			}

			if (e.Cell.Column.ToString() == "Unidad")
			{
				CambiaCantidad(e);
				if (!Limites(e)) 
				{
					bActualiza = true;
					return;
				}
				// Cambio de precio por cantidad
				if ((int) this.drSeteoF["TipoPrecio"] == 1) Precio(e);
				Promocion(e);
				DescuentoPorcentaje(e);

				cmdTotal();
			}
			
			if (e.Cell.Column.ToString() == "DescuentoPorc")
			{
				bActualiza = false;
				DescuentoPorcentaje(e);
				cmdTotal();
				bActualiza = true;
			}
			if (e.Cell.Column.ToString() == "DescuentoArt")
			{
				double dCantidad = (double) e.Cell.Row.Cells["Cantidad"].Value;
				double dPrecio = 0;
				if (e.Cell.Row.Cells["Precio"].Value != System.DBNull.Value)
					dPrecio = (double) e.Cell.Row.Cells["Precio"].Value;
				if (dCantidad == 0 || dPrecio == 0) 
				{
					bActualiza = true;
					return;
				}

				double dDesc = dCantidad * dPrecio * (double) 
					e.Cell.Row.Cells["DescuentoPorc"].Value / 100.0;
				double dDesc1 = Math.Round((double) e.Cell.Row.Cells["DescuentoArt"].Value, 2);
				double dPorc = dDesc1 / dCantidad / dPrecio *100.0;
				dPorc = Math.Round(dPorc, 2);
				if ((double) drSeteoF["LimiteDescuento"] >= 0.01 && dPorc > (double) drSeteoF["LimiteDescuento"])
				{
					MessageBox.Show("Descuento no Permitido", "Informacion");
					e.Cell.Row.Cells["DescuentoArt"].Value = 0;
					e.Cell.Row.Cells["DescuentoPorc"].Value = 0;
				}
				else
				{
					if (Math.Abs(dDesc - dDesc1) > .009)
						e.Cell.Row.Cells["DescuentoPorc"].Value = dPorc;
				}
				cmdTotal();
			}

			if (e.Cell.Column.ToString() == "Precio")
			{
				DescuentoPorcentaje(e);
				cmdTotal();
			}
			bActualiza = true;
		}

		private void CambiaCantidad(Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			if ((bool) drSeteoF["VerUnidad"] == false) return;
			int IdUnidad = (int) e.Cell.Row.Cells["idUnidad"].Value;
			int IdArticulo = (int) e.Cell.Row.Cells["idArticulo"].Value;
			string stSelect = "Exec UnidadCambio "	+ IdArticulo.ToString();
			SqlDataReader miReader = Funcion.rEscalarSQL(cdsArticulo, stSelect, false);
			miReader.Read();

			int IdUnidadOr = miReader.GetInt32(0);
			//Poner solo cuando hay unidades visible
			if (IdUnidad == IdUnidadOr)
			{
				double dCant = miReader.GetDouble(4);
				e.Cell.Row.Cells["Cantidad"].Value = 
					(double) e.Cell.Row.Cells["Unidades"].Value * dCant;
			}
			else
			{
				double dCantidad = miReader.GetDouble(1);
				if (dCantidad > 0)
				{
					e.Cell.Row.Cells["Cantidad"].Value = 
						(double) e.Cell.Row.Cells["Unidades"].Value * dCantidad;
					bool bPeso = miReader.GetBoolean(5);
					if (bPeso)
					{
						e.Cell.Row.Cells["Precio"].Value = 
							(double) e.Cell.Row.Cells["Precio"].Value / dCantidad;
						e.Cell.Row.Cells["Precio"].Refresh();
					}
				}
			}
			miReader.Close();
		}

		/// <summary>
		/// Cambio de Descuento en porcentaje
		/// </summary>
		/// <returns></returns>
		private void DescuentoPorcentaje(Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			if ((double) e.Cell.Row.Cells["DescuentoPorc"].Value < 0.001) return;
			if ((IdTipoFactura == 1 || IdTipoFactura == 3 
				|| IdTipoFactura == 13 || IdTipoFactura == 11)
				&& (double) drSeteoF["LimiteDescuento"] < 
				(double) e.Cell.Row.Cells["DescuentoPorc"].Value && 
				(double) drSeteoF["LimiteDescuento"] > 0)
			{
				e.Cell.Row.Cells["DescuentoPorc"].Value = 0;
				MessageBox.Show("Descuento Supera el Limite de Descuento Permitido", 
					"Informacion", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
			}
			double dCantidad = (double) e.Cell.Row.Cells["Cantidad"].Value;
			double dPrecio = 0;
			if (e.Cell.Row.Cells["Precio"].Value != System.DBNull.Value)
				dPrecio = (double) e.Cell.Row.Cells["Precio"].Value;
			double dSubTLinea = dCantidad * dPrecio;
			if (bIncluyeIva)
			{
				double dIva = (double) e.Cell.Row.Cells["Impuesto"].Value;
				dSubTLinea = dSubTLinea / (1.0 + dIva / 100.0);
			}
			double dDesc = dSubTLinea * (double) e.Cell.Row.Cells["DescuentoPorc"].Value;
			try
			{
				e.Cell.Row.Cells["DescuentoArt"].Value = dDesc / 100.0;
			}
			catch(System.Exception ex)
			{
				MessageBox.Show(ex.Message, "Descuento Porcentaje");
			}
		}

		/// <summary>
		/// Verifica si existe Promociones de Articulo
		/// </summary>
		/// <returns></returns>
		private void Promocion(Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			if (!(IdTipoFactura == 1 || IdTipoFactura == 3)) return;
			if (e.Cell.Row.Cells["idArticulo"].Value == System.DBNull.Value) return;
			if (e.Cell.Row.Cells["Cantidad"].Value == System.DBNull.Value) return;
			double fCantidad = (double) e.Cell.Row.Cells["Cantidad"].Value;

			int IdArt = (int) e.Cell.Row.Cells["idArticulo"].Value;
			string stSelect = "Select TipoPromocion FROM Articulo Where "
				+ " idArticulo = " + IdArt.ToString();
			int ITipoProm = Funcion.iEscalarSQL(cdsArticulo, stSelect);
			if (ITipoProm == 2) // Descuento por Valor
			{
				stSelect = "Select CantReqDesc FROM Articulo Where "
					+ " idArticulo = " + IdArt.ToString();
				double dCantReq = Funcion.dEscalarSQL(cdsArticulo, stSelect);
				if (fCantidad >= dCantReq)
				{
					if (dCantReq == 0) dCantReq = 1;
					int iFraccion = (int) (fCantidad / dCantReq);
					string stDescuento = "DescuentoContado";
					if ((int) rdgContadoCred.Value == 1) stDescuento = "DescuentoCredito";
					stSelect = "Select " + stDescuento + " FROM Articulo Where "
						+ " idArticulo = " + IdArt.ToString();
					double dDescuento = Funcion.dEscalarSQL(cdsArticulo, stSelect);
					e.Cell.Row.Cells["DescuentoArt"].Value = dDescuento * iFraccion;
				} 
			} // Fin de Descuento por valor
			if (ITipoProm == 3) // Descuento por Porcentaje
			{
				stSelect = "Select CantReqPorc FROM Articulo Where "
					+ " idArticulo = " + IdArt.ToString();
				double dCantReq = Funcion.dEscalarSQL(cdsArticulo, stSelect);
				if (fCantidad >= dCantReq)
				{
					string stDescuento = "DescPorcContado";
					if ((int) rdgContadoCred.Value == 1) stDescuento = "DescPorcCre";
					stSelect = "Select " + stDescuento + " FROM Articulo Where "
						+ " idArticulo = " + IdArt.ToString();
					double dDescuento = Funcion.dEscalarSQL(cdsArticulo, stSelect);
					e.Cell.Row.Cells["DescuentoPorc"].Value = dDescuento;
					double fPrecio = (double) e.Cell.Row.Cells["Precio"].Value;
					e.Cell.Row.Cells["DescuentoArt"].Value = dDescuento * fPrecio * fCantidad / 100.0;
				} // Fin de cantidad
			}// Fin de Descuento por valor
		}

		/// <summary>
		/// Verifica si existe saldo para el articulos seleccionado
		/// </summary>
		/// <returns></returns>
		private bool Limites(Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			// Retorno en valores nulos de cantidad o de idArt
			if (e.Cell.Row.Cells["idArticulo"].Value == System.DBNull.Value) return true;
			if (e.Cell.Row.Cells["Cantidad"].Value == System.DBNull.Value) return true;

			if (Funcion.bEjecutaSQL(cdsArticulo, "Exec SeteoGExiste 'FPVCNSP'"))
			{
				string stIdArt = e.Cell.Row.Cells["idArticulo"].Value.ToString();
				string stCant = e.Cell.Row.Cells["Cantidad"].Value.ToString();
				string stTipo = cmbTipoCompra.Value.ToString();
				string stVerif = "Exec DetCompraVerifNegValida " + stIdArt + ", "
					+ stCant + ", " + stTipo;
				if(!Funcion.bEjecutaSQL(cdsArticulo, stVerif)) return true;

				string stBodega = cmbBodega.Value.ToString();
				if ((bool) drSeteoF["VerBodega"])
					stBodega = e.Cell.Row.Cells["Bodega"].Value.ToString();
				string stidDetC = e.Cell.Row.Cells["idDetCompra"].Value.ToString();
				stVerif = "Exec DetCompraVerifNegativo " + stIdArt + ", "
					+ stCant + ", " + stBodega + ", " + stidDetC + ", " + stTipo;
				string stMensaje = Funcion.sEscalarSQL(cdsArticulo, stVerif, true);
				if (stMensaje.Length > 0)
				{
					if (stMensaje.StartsWith("-"))
						e.Cell.Row.Cells["Cantidad"].Value = 0;
					MessageBox.Show(stMensaje.Substring(1), "Informacion", 
						MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
					return false;
				}
				return true;
			}

			if ((double) e.Cell.Row.Cells["Cantidad"].Value == 0)  return true;
			int iTipo = (int) this.cmbTipoCompra.Value;
			if ((iTipo == 3 || iTipo == 13)  && drSeteoF["SaldoEnCotiza"] != DBNull.Value && 
				(bool) drSeteoF["SaldoEnCotiza"] == true)
			{
				// Cotizacion que revisa inventario
			}
			else
			{
				if (!(iTipo == 1 || iTipo == 6 || iTipo == 8)) return true;
			}

			double fSaldo = 0;
			double fSaldoMin = 0;
			double fSaldoBodega = 0;
			double fCantidad = (double) e.Cell.Row.Cells["Cantidad"].Value;
			if ((bool) drSeteoF["AlertaSaldo"] == true || 
				(bool) drSeteoF["SaldoNegativo"] == true)
			{
				double fCantOriginal = 0;
				int iBodegaAntes = 0;
				int IdDet = (int) e.Cell.Row.Cells["idDetCompra"].Value;
				if (this.barraDato1.bNuevo == false && iTipo != 13 && IdDet > 0) // Se encuentra editando
				{
					cdsCompra.Schema.Connections[0].Open();
					C1.Data.SchemaObjects.Connection connection = this.cdsCompra.Schema.Connections[0];
					string stSelect = "SELECT IsNull(Cantidad, 0) FROM DetCompra WHERE idDetCompra = " +
						IdDet.ToString();
					SqlCommand cmdCant = new SqlCommand (stSelect
						, (SqlConnection)connection.DbConnection);

					fCantOriginal = (double) cmdCant.ExecuteScalar();

					cmdCant = new SqlCommand (
						"SELECT IsNull(Bodega, 0) FROM DetCompra WHERE idDetCompra = " +
						IdDet.ToString(), (SqlConnection)connection.DbConnection);

					iBodegaAntes = (int) cmdCant.ExecuteScalar();
					cdsCompra.Schema.Connections[0].Close();
					// Fin de Lectura de valor
				}

//				C1DataRow drArt = cdsArticulo.TableViews["Articulo"].Rows.Find(e.Cell.Row.Cells["idArticulo"].Value);
				string stSaldoArt = "SELECT idTipoGrupo FROM Articulo WHERE idArticulo = " +
					e.Cell.Row.Cells["idArticulo"].Value.ToString();
				int iTipoGrupo = Funcion.iEscalarSQL(cdsArticulo, stSaldoArt, true);
				if (!(iTipoGrupo == 1 || iTipoGrupo == 5)) return true;

				stSaldoArt = "SELECT IsNull(Saldo, 0) FROM Articulo WHERE idArticulo = " +
					e.Cell.Row.Cells["idArticulo"].Value.ToString();
				fSaldo = Funcion.dEscalarSQL(cdsArticulo, stSaldoArt, true) + fCantOriginal;

				stSaldoArt = "SELECT CantidadMinima FROM Articulo WHERE idArticulo = " +
					e.Cell.Row.Cells["idArticulo"].Value.ToString();
				fSaldoMin = Funcion.dEscalarSQL(cdsArticulo, stSaldoArt, true);

				// Calculo de saldo en una bodega especifica
				if (false && (bool) this.drSeteoF["SaldoPorBodega"])
				{
					int iBodegaActual = (int) e.Cell.Row.Cells["Bodega"].Value;
					string stBodega = "Bodega" + iBodegaActual.ToString().Trim();
					stSaldoArt = "SELECT " + stBodega + " FROM Articulo WHERE idArticulo = " +
						e.Cell.Row.Cells["idArticulo"].Value.ToString();
					try
					{
						fSaldoBodega = Funcion.dEscalarSQL(cdsArticulo, stSaldoArt, true);
						if (iBodegaAntes == iBodegaActual)
							fSaldoBodega += fCantOriginal;
					}
					catch(System.Exception ex)
					{
						MessageBox.Show(ex.Message, "Saldo de Bodega");
					}
				}
				// Fin de calculo de saldo por bodega
			}

			// Alerta de saldos minimos
			if ((bool) drSeteoF["AlertaSaldo"] == true)
				if (fSaldo - fSaldoMin  < fCantidad)
					MessageBox.Show("Saldo Antes: " 
						+ fSaldo.ToString("#,##0.00") + "\nSaldo Minimo: "
						+ fSaldoMin.ToString("#,##0.00") + 
						"\nCantidad: " + fCantidad.ToString("#,##0.00"));
			// Fin de alerta de saldos minimos

			// Impide Saldos Negativos
			if ((bool) drSeteoF["SaldoNegativo"] == false) return true;
			string strBodega = "";
			if ((bool) drSeteoF["SaldoPorBodega"])
				strBodega = "\nSaldo de Bodega: " + fSaldoBodega.ToString("#,##0.00");
			else
				fSaldoBodega = 0;
			double fDiferencia = fSaldo + fSaldoBodega - (double) e.Cell.Row.Cells["Cantidad"].Value;
			if (fDiferencia < -0.00001)
			{
				e.Cell.Row.Cells["Cantidad"].Value = 0;
				e.Cell.Row.Cells["Unidades"].Value = 0;
				MessageBox.Show("Cantidad excede Saldo " + fSaldo.ToString() + strBodega);
				return false;
			}
			// Fin de Bloque de saldos negativos
			return true;
		}
	

		/// <summary>
		/// Calculo del precio e Iva de acuerdo a tipo de cliente
		/// </summary>
		/// <param name="e"></param>
		private void Precio(Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			if (e.Cell.Row.Cells["idArticulo"].Value == System.DBNull.Value) return;

			int IdArt = (int) e.Cell.Row.Cells["idArticulo"].Value;
			miArt = new Items(cdsArticulo, IdArt);

			if (miProveedor.InstPublica)// && ((int) this.cmbTipoCompra.Value == 3 || (int) this.cmbTipoCompra.Value == 1))
				e.Cell.Row.Cells["Impuesto"].Value = 0;
			else
				e.Cell.Row.Cells["Impuesto"].Value = miArt.Iva;
			
			if (!((int) this.cmbTipoCompra.Value == 8 || (int) this.cmbTipoCompra.Value == 9))
				if ((int) this.drSeteoF["TipoPrecio"] == 4) return;

			double fPrecio = 0;
			// Calculo de precio
			switch ((int) this.cmbTipoCompra.Value)
			{
				case 9: // Ingreso de Bodega
					e.Cell.Row.Cells["Impuesto"].Value = 0;
					return;
				case 1: // Factura
				case 3: // Cotizacion
				case 8: // Transferencia Enviada
					//				case 10: // Orden de Produccion
				case 13: // Reservacion Cliente
				case 15: // Pedido Cliente

					// Egreso de Bodega a Precio de Venta o Costo
					if (drSeteoF["PrecioTE"] != System.DBNull.Value && !(bool) drSeteoF["PrecioTE"])
					{
						if ((int) this.cmbTipoCompra.Value == 8) 
						{
							fPrecio = miArt.CostoPromedio;
							e.Cell.Row.Cells["Precio"].Value = Math.Round(fPrecio, 5);
							e.Cell.Row.Cells["Impuesto"].Value = 0;
							return;
						}
						if ((int) this.cmbTipoCompra.Value == 9) 
						{
							fPrecio = miArt.CostoPromedio;
							e.Cell.Row.Cells["Precio"].Value = Math.Round(fPrecio, 5);
							e.Cell.Row.Cells["Impuesto"].Value = 0;
							return;
						}
					}
				switch ((int) this.drSeteoF["TipoPrecio"])
				{
					case 0: // Precio definido por cliente
						int stPrecio = 1;
							stPrecio = (int) miProveedor.idTipoPrecio;
							fPrecio = (double) miArt.Precio[stPrecio];
						break;
					case 1: // Precio definido por Cantidad Vendida
						double dCantidad = (double) e.Cell.Row.Cells["Cantidad"].Value;
						if (dCantidad >= miArt.Cantidad[5])
							fPrecio = (double) miArt.Precio[5];
						else if (dCantidad >= miArt.Cantidad[4])
							fPrecio = (double) miArt.Precio[4];
						else if (dCantidad >= miArt.Cantidad[3])
							fPrecio = (double) miArt.Precio[3];
						else if (dCantidad >= miArt.Cantidad[2])
							fPrecio = (double) miArt.Precio[2];
						else //(dCantidad >= (int) miArt.Cantidad[1])
							fPrecio = (double) miArt.Precio[1];
						break;
					case 2: // Precio definido por Forma de Pago
						C1DataRow drFormaPago = 
							cdsCompraTabla.TableViews["CompraFormaPago"].Rows.Find((int) this.cmbFormaPago.Value);
						stPrecio = 1;
						if (drFormaPago != null)
							stPrecio = (int) drFormaPago["idPrecio"];
							fPrecio = (double) miArt.Precio[stPrecio];
						break;
					case 3:
						double dCantidadV = (double) e.Cell.Row.Cells["Cantidad"].Value;
						int IdClienteV = (int) cmbCliente.Value;
						string stExec = "Exec ArticuloPrecioLista " + IdClienteV.ToString() + ", "
							+ IdArt + ", " + dCantidadV.ToString("0.00", us);
						fPrecio = Funcion.dEscalarSQL(cdsArticulo, stExec, true);
						break;
				}
					break;
				case 4: // Compra
				case 2: // Orden de Compra
				case 6: // Dev Compra
				case 10: // Consumo Real
				case 12: // Producto Terminado
				case 14: // Reservacion de Proveedor
				case 16: // Pedido Proveedor
				case 36: // Requisicion
					if ((int) this.cmbTipoCompra.Value == 12)
						e.Cell.Row.Cells["Impuesto"].Value = 0;
					if ((bool)this.drSeteoF["c_u_c_p"])
            fPrecio = miArt.CostoUltimo;
					else
						fPrecio = 0;
					break;
			}
			e.Cell.Row.Cells["Precio"].Value = fPrecio;
		}

		private void ultraGrid1_AfterRowInsert(object sender, Infragistics.Win.UltraWinGrid.RowEventArgs e)
		{
			e.Row.Cells["Bodega"].Value = (int) this.spnBodega.Value;
			e.Row.Cells["Signo"].Value = IdSigno;
		}

		private void ultraGrid1_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
			drSeteoF = this.cdsSeteoF.TableViews["SeteoF"].Rows[0];
			e.Layout.Bands[0].Columns["Bodega"].Hidden = !(bool) drSeteoF["VerBodega"];
			e.Layout.Bands[0].Columns["CodArt"].Hidden = !(bool) drSeteoF["EscogerCodigo"];
			e.Layout.Bands[0].Columns["RefCodigo"].Hidden = !(bool) drSeteoF["VerRefCodigo"];
			e.Layout.Bands[0].Columns["RefNumero"].Hidden = !(bool) drSeteoF["VerRefNumero"];

			if (drSeteoF["LimiteTransforma"] != DBNull.Value && (bool) drSeteoF["LimiteTransforma"] == true)
			{
				string stRefC = "Ref. Codigo";
				string stRefN = "Ref. Numero";
				if (IdTipoFactura == 4)
				{
					stRefC = "Orden Compra";
					stRefN = "Requisicion";
				}
				if (IdTipoFactura == 2)
				{
					stRefC = "Compra";
					stRefN = "Requisicion";
				}
				if (IdTipoFactura == 36)
				{
					stRefC = "Compra";
					stRefN = "Orden Compra";
					chkRevisado.Text = "Centralizado";
				}

				e.Layout.Bands[0].Columns["RefCodigo"].Header.Caption = stRefC;
				e.Layout.Bands[0].Columns["RefNumero"].Header.Caption = stRefN;
			}

			e.Layout.Bands[0].Columns["CantidadPeso"].Hidden = !(bool) drSeteoF["VerPeso"];
			e.Layout.Bands[0].Columns["Pedido"].Hidden = !(bool) drSeteoF["VerPedido"];
			e.Layout.Bands[0].Columns["DescuentoArt"].Hidden = !(bool) drSeteoF["VerDescuento"];
			e.Layout.Bands[0].Columns["DescuentoPorc"].Hidden = !(bool) drSeteoF["VerDescPorc"];
			e.Layout.Bands[0].Columns["Flete"].Hidden = !(bool) drSeteoF["VerFlete"];
			e.Layout.Bands[0].Columns["Notas"].Hidden = !(bool) drSeteoF["VerNotas"];
			e.Layout.Bands[0].Columns["Vencimiento"].Hidden = !(bool) drSeteoF["VerVencimiento"];
			e.Layout.Bands[0].Columns["idProyecto"].Hidden = !(bool) drSeteoF["VerProyecto"];
			e.Layout.Bands[0].Columns["idSubProyecto"].Hidden = !(bool) drSeteoF["VerSubProyecto"];
			e.Layout.Bands[0].Columns["idUnidad"].Hidden = !(bool) drSeteoF["VerUnidad"];
			e.Layout.Bands[0].Columns["Unidades"].Hidden = !(bool) drSeteoF["VerUnidad"];
			if (drSeteoF["Importacion"] != System.DBNull.Value && (bool) drSeteoF["Importacion"])
				e.Layout.Bands[0].Columns["idImportacion"].Hidden = false;
			if ((bool) drSeteoF["VerUnidad"])
				e.Layout.Bands[0].Columns["Cantidad"].CellActivation = 
					Infragistics.Win.UltraWinGrid.Activation.Disabled;
			if (drSeteoF["VerNombreLista"] == System.DBNull.Value || (bool) drSeteoF["VerNombreLista"])
				txtNombre.Width = 0;
			else
				cmbCliente.Width = 0;
			
			if (!(IdTipoFactura == 1 || IdTipoFactura == 3 || IdTipoFactura == 8))
			{
				cmbPrecio.Visible = false;
				btCambiarPrecio.Visible = false;
			}

			if (drSeteoF["VerDesplegable"] != System.DBNull.Value && !(bool) drSeteoF["VerDesplegable"])
			{
				e.Layout.Bands[0].Columns["Articulo"].Hidden = false;
				e.Layout.Bands[0].Columns["Codigo"].Hidden = false;
				e.Layout.Bands[0].Columns["CodArt"].Hidden = true;
				e.Layout.Bands[0].Columns["idArticulo"].Hidden = true;
			}

			if (!(bool) drSeteoF["CambiarPrecio"] && (IdTipoFactura == 1 || IdTipoFactura == 3))
			{
				e.Layout.Bands[0].Columns["Precio"].CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
				cmbPrecio.Visible = false;
				btCambiarPrecio.Visible = false;
			}
			if (!(bool)drSeteoF["CambiarIva"])
				e.Layout.Bands[0].Columns["Impuesto"].CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;


			// Compra y Transferencia Enviado y Recibida
			if (IdTipoFactura == 4 || IdTipoFactura == 8 || IdTipoFactura == 9)
			{
				if (drSeteoF["CambiarIvaC"] == System.DBNull.Value || !(bool)drSeteoF["CambiarIvaC"])
					e.Layout.Bands[0].Columns["Impuesto"].CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
				else
					e.Layout.Bands[0].Columns["Impuesto"].CellActivation = Infragistics.Win.UltraWinGrid.Activation.AllowEdit;
			}
			//			if (!miAcceso.CambiarPrecio)
			//				e.Layout.Bands[0].Columns["Precio"].CellActivation = 
			//					Infragistics.Win.UltraWinGrid.Activation.Disabled;

			// Formato numerico
			CultureInfo cultureENUS = CultureInfo.CreateSpecificCulture("ja-jp");
			string stMoneda = "#,##0.";
			string stMascara = "nnn,nnn,nnn.";
			string stCeros = "0000000000";
			string stNumero = "nnnnnnnnnn";
			string stFormato = "";
			string stInput = "";

			// Decimales para el total
			stFormato = stMoneda + stCeros.Substring(0, (int) drSeteoF["decTotal"]);
			stInput = stMascara + stNumero.Substring(0, (int) drSeteoF["decTotal"]);
			e.Layout.Bands[0].Columns["SubTotal"].Format = stFormato; 
			e.Layout.Bands[0].Columns["DescuentoArt"].Format = stFormato;

			this.txtDescIvaTotal.FormatString = stFormato;
			this.txtIceTotal.FormatString = stFormato;
			this.txtIvaTotal.FormatString = stFormato;
			this.txtIva.FormatString = stFormato;
			this.txtIva0.FormatString = stFormato;
			this.txtTotal.FormatString = stFormato;
			this.txtDescIva.FormatString = stFormato;
			this.txtDescIva0.FormatString = stFormato;
			this.txtDescIva0Total.FormatString = stFormato;
			this.txtDescLinea.FormatString = stFormato;
			this.txtDescLinea0.FormatString = stFormato;
			this.txtFlete.FormatString = stFormato;

			this.txtDescIvaTotal.MaskInput = stInput;
			this.txtIceTotal.MaskInput = stInput;
			this.txtIvaTotal.MaskInput = stInput;
			this.txtIva.MaskInput = stInput;
			this.txtIva0.MaskInput = stInput;
			this.txtTotal.MaskInput = stInput;
			this.txtDescIva.MaskInput = stInput;
			this.txtDescIva0.MaskInput = stInput;
			this.txtDescIva0Total.MaskInput = stInput;
			this.txtDescLinea.MaskInput = stInput;
			this.txtDescLinea0.MaskInput = stInput;
			this.txtFlete.MaskInput = stInput;

			// Decimales para la cantidad, peso y Pedido
			stFormato = stMoneda + stCeros.Substring(0, (int) drSeteoF["decCantidad"]);
			e.Layout.Bands[0].Columns["Cantidad"].Format = stFormato;
			e.Layout.Bands[0].Columns["CantidadPeso"].Format = stFormato;
			e.Layout.Bands[0].Columns["Pedido"].Format = stFormato;
			stInput = stMascara + stNumero.Substring(0, (int) drSeteoF["decCantidad"]);
			e.Layout.Bands[0].Columns["Cantidad"].MaskInput = stInput;
			e.Layout.Bands[0].Columns["CantidadPeso"].MaskInput = stInput;
			e.Layout.Bands[0].Columns["Pedido"].MaskInput = stInput;

			// Decimales para el precio
			stFormato = stMoneda + stCeros.Substring(0, (int) drSeteoF["decPrecio"]);
			e.Layout.Bands[0].Columns["Precio"].Format = stFormato;
			stInput = stMascara + stNumero.Substring(0, (int) drSeteoF["decPrecio"]);
			e.Layout.Bands[0].Columns["Precio"].MaskInput = stInput;

			// Decimales para el Impuesto
			stFormato = stMoneda + stCeros.Substring(0, (int) drSeteoF["decImpuesto"]);
			e.Layout.Bands[0].Columns["Impuesto"].Format = stFormato;
			stInput = stMascara + stNumero.Substring(0, (int) drSeteoF["decImpuesto"]);
			e.Layout.Bands[0].Columns["Impuesto"].MaskInput = stInput;

			// Decimales para porcetaje de descuento
			stFormato = stMoneda + stCeros.Substring(0, (int) drSeteoF["decPorcDesc"]);
			e.Layout.Bands[0].Columns["DescuentoPorc"].Format = stFormato;
			this.txtDescPorcIva.FormatString = stFormato;
			this.txtDescPorcIva0.FormatString = stFormato;
			int iDecPorD = 0;
			if (drSeteoF["decPorcDesc"] != null) iDecPorD = (int) drSeteoF["decPorcDesc"];

			stInput = stMascara + stNumero.Substring(0, iDecPorD);
			e.Layout.Bands[0].Columns["DescuentoPorc"].MaskInput = stInput;
			this.txtDescPorcIva.MaskInput = stInput;
			this.txtDescPorcIva0.MaskInput = stInput;
		}

		private int CopiarFactura(int idDestino)
		{
			C1.Data.C1DataRow drC = this.cdsCompra.TableViews["Compra"].AddNew();
			// Calculo de numeracion
			C1DataRow drNumero = this.cdsCompraTabla.TableViews["CompraNumero"].Rows.Find(idDestino);
			int IdSigno1 = (int) drNumero["Signo"];
			if ((bool) drNumero["Automatico"] == true || MenuLatinium.bFox==true)
			{
				int iTipoNumCopia = 0;
				if (cmbComprobante1.ActiveRow != null && 
					cmbComprobante1.ActiveRow.Cells["idCompraNumero"].Value 
					!= System.DBNull.Value)
					iTipoNumCopia = (int) cmbComprobante1.ActiveRow.Cells["idCompraNumero"].Value;
				if (iTipoNumCopia == 0) iTipoNumCopia = idDestino;

//				string stNumero = Funcion.NumeraFactura(this.cdsCompraTabla, iTipoNumCopia);
//				if (stNumero.Trim().Length > 0) drC["Numero"] = stNumero;
//				this.cdsCompraTabla.Update();
			}// fin de numeracion
			drC["idCliente"] = drCompra["idCliente"];
			drC["Bulto"] = drCompra["Bulto"];
			drC["Peso"] = drCompra["Peso"];
			drC["Pedido"] = drCompra["Pedido"];
			drC["Depart"] = drCompra["Depart"];
			drC["Otro"] = drCompra["Otro"];
			drC["Mensaje1"] = drCompra["Mensaje1"];
			drC["Mensaje2"] = drCompra["Mensaje2"];
			drC["Mensaje3"] = drCompra["Mensaje3"];
			drC["Cif"] = drCompra["Cif"];

			drC["idComprobante"] = drCompra["idComprobante"];
			drC["idRecibe"] = drCompra["idRecibe"];
			drC["idProyecto"] = drCompra["idProyecto"];
			drC["idSubProyecto"] = drCompra["idSubProyecto"];
			drC["idFormaPago"] = drCompra["idFormaPago"];
			drC["idTipoMoneda"] = drCompra["idTipoMoneda"];
			if (idDestino == 6)	drC["idComprobante"] = 4;
			else if (idDestino == 5) drC["idComprobante"] = 4;
			else drC["idComprobante"] = drCompra["idComprobante"];
			drC["idCredTributario"] = drCompra["idCredTributario"];
			drC["ContadoCredito"] = drCompra["ContadoCredito"];
			drC["Otro"] = drCompra["Otro"];
			drC["Fecha"] = DateTime.Today;
			drC["FechaIngreso"] = DateTime.Now;
			drC["FechaVencimiento"] = DateTime.Today;
			drC["idTipoFactura"] = idDestino;
			drC["Comprobante"] = drCompra["Numero"];
			drC["Total"] = drCompra["Total"];
			drC["Desc1"] = drCompra["Desc1"];
			drC["Desc2"] = drCompra["Desc2"];
			drC["Desc3"] = drCompra["Desc3"];
			drC["Desc4"] = drCompra["Desc4"];
			drC["Descuento"] = drCompra["Descuento"];
			drC["Descuento0"] = drCompra["Descuento0"];
			drC["Iva"] = drCompra["Iva"];
			drC["SubTotalIva"] = drCompra["SubTotalIva"];
			drC["SubTotalExcento"] = drCompra["SubTotalExcento"];
			drC["Retencion"] = drCompra["Retencion"];
			drC["Ice"] = drCompra["Ice"];
			drC["Flete"] = drCompra["Flete"];
			drC["Interes"] = drCompra["Interes"];
			drC["ValorDolar"] = drCompra["ValorDolar"];
			drC["DiasCredito"] = drCompra["DiasCredito"];
			drC["FechaVencimiento"] = DateTime.Today.AddDays((int) drCompra["DiasCredito"]);
			drC["idVendedor"] = drCompra["idVendedor"];
			drC["ValorDolar"] = drCompra["ValorDolar"];
			drC["Notas"] = drCompra["Notas"];
			drC["Bodega"] = drCompra["Bodega"];
			drC["AutFactura"] = drCompra["AutFactura"];
			drC["FechaCaducidad"] = drCompra["FechaCaducidad"];
			drC["AutImprenta"] = drCompra["AutImprenta"];
			drC["SerieFactura"] = drCompra["SerieFactura"];
			// drC["Editar"] = true;
			if (!bAutoImpresor && !bNumeraGrabar)
			{
				int iTipoNumera = 0;
				if (idDestino == 1 || idDestino == 3) // Para las notas de venta
				{
					if (cmbComprobante1.ActiveRow != null && 
						cmbComprobante1.ActiveRow.Cells["idCompraNumero"].Value 
						!= System.DBNull.Value)
						iTipoNumera = (int) cmbComprobante1.ActiveRow.Cells["idCompraNumero"].Value;
				}
				if (iTipoNumera == 0) iTipoNumera = idDestino;

				string stNumeroFact = Funcion.NumeraFactura(cdsCompraTabla, iTipoNumera);
				drC["Numero"] = stNumeroFact;
			}
			this.cdsCompra.Update();

			// Copia de lineas
			foreach(C1.Data.C1DataRow dr in drCompra.GetChildRows())
			{
				C1.Data.C1DataRow drDC = this.cdsCompra.TableViews["DetCompra"].AddNew();
				drDC["idCompra"] = drC["idCompra"];
				drDC["idArticulo"] = dr["idArticulo"];
				drDC["idUnidad"] = dr["idUnidad"];
				drDC["Cantidad"] = dr["Cantidad"];
				drDC["Unidades"] = dr["Unidades"];
				drDC["Precio"] = dr["Precio"];
				drDC["Impuesto"] = dr["Impuesto"];
				drDC["Notas"] = dr["Notas"];
				drDC["RefCodigo"] = dr["RefCodigo"];
				drDC["RefNumero"] = dr["RefNumero"];
				drDC["DescuentoArt"] = dr["DescuentoArt"];
				drDC["CantidadPeso"] = dr["CantidadPeso"];
				drDC["Flete"] = dr["Flete"];
				drDC["Bodega"] = dr["Bodega"];
				drDC["DescuentoPorc"] = dr["DescuentoPorc"];
				drDC["Pedido"] = dr["Pedido"];
				drDC["Signo"] = IdSigno1;
				drDC["Codigo"] = dr["Codigo"];
				drDC["Articulo"] = dr["Articulo"];
				drDC["idProyecto"] = dr["idProyecto"];
				drDC["idSubProyecto"] = dr["idSubProyecto"];
			}
			this.cdsCompra.Update();

			cmdTotal();
			return (int) drC["idCompra"];
		}

		private void btTransforma_Click(object sender, System.EventArgs e)
		{
			if (barraDato1.bEditar) return;
			if(!miAcceso.Transformacion)
			{
				MessageBox.Show("No tiene acceso a Transformar " + this.Text.Trim(),
					"Confirmación", MessageBoxButtons.OK, MessageBoxIcon.Information);
				return;
			}
			if (Funcion.bEjecutaSQL(cdsArticulo, "Exec SeteoGExiste 'FTT'"))
			{
				string stMensaje = Funcion.sEscalarSQL(cdsArticulo, 
					"Exec CompraTransformaV " + txtIdCompra.Value.ToString(), true);
				if (stMensaje.Length > 0)
				{
					MessageBox.Show(stMensaje, "Información", MessageBoxButtons.OK, MessageBoxIcon.Warning);
					return;
				}

				stMensaje = Funcion.sEscalarSQL(cdsArticulo, 
					"Exec CompraTransformaM " + txtIdCompra.Value.ToString() + ", 1", true);
				if (DialogResult.Yes == MessageBox.Show(stMensaje, 
					"Confirmación", MessageBoxButtons.YesNo, MessageBoxIcon.Question))
				{
					this.Cursor = Cursors.WaitCursor;
					int iDestino = Funcion.iEscalarSQL(cdsArticulo, 
						"Exec CompraTransformaM " + txtIdCompra.Value.ToString() + ", 2", true);
					int idCompraDest = Funcion.iEscalarSQL(cdsArticulo, 
						"Exec CompraTransformaCopia " + txtIdCompra.Value.ToString(), true);
					Compra miCompra1 = new Compra(iDestino, idCompraDest);
					miCompra1.MdiParent = this.MdiParent;
					miCompra1.Show();
					this.Cursor = Cursors.Default;
				}
				return;
			}

			if ((bool) this.lblAnulado.Visible == true)
			{
				MessageBox.Show("No puede transformar registros Anulados ",
					"Información", MessageBoxButtons.OK, MessageBoxIcon.Warning);
				return;
			}
			string stTransforma = "";
			int idDestino = 0;
			if (drCompra["idCliente"] != System.DBNull.Value)
				miProveedor = new Proveedor(cdsCliente, (int) drCompra["idCliente"]);

			switch(this.IdTipoFactura)
			{
				case 1:
					stTransforma = "Desea crear una devolución en venta de esta Venta";
					idDestino = 5;
					break;
				case 2:
					string stUsuarioT = "Exec VerificaUsuarioT " + txtIdCompra.Value.ToString();
					int iVerifica = Funcion.iEscalarSQL(cdsArticulo, stUsuarioT, true);
					if (iVerifica == 0) 
					{
						MessageBox.Show("No tiene el cupo de credito para transformar", "Informacion");
						return;
					}

					if (drSeteoF["OrdenCIngreso"] != System.DBNull.Value &&
						(bool) drSeteoF["OrdenCIngreso"] == true)
					{
						string stSelect = "Select Count(*) From CompraNumero where IsNull(Transforma, 0) = 2";
						int iCuenta1 = Funcion.iEscalarSQL(cdsCliente, stSelect, true);
						if (iCuenta1 > 0)
						{
							ListaTransformacion miLista = new ListaTransformacion();
							if (DialogResult.Cancel == miLista.ShowDialog())
							{
								miLista.Dispose();
								return;
							}
						
							stTransforma = "Desea crear un " + miLista.txtNombre.Text + " con esta orden de compra";
							idDestino = (int) miLista.txtIdFactura.Value;
							miLista.Dispose();
						}
						else
						{
							MessageBox.Show("Defina las transformaciones en Archivo - Propiedades - Facturacion - Numeracion "
								+ "\n ingresando en la Columna Transforma el valor 2 en las filas a transformar", "Informacion"
								, MessageBoxButtons.OK, MessageBoxIcon.Information);
							return;
						}
					}
					else if (drSeteoF["OrdenCPedido"] != System.DBNull.Value &&
						(bool) drSeteoF["OrdenCPedido"] == true)
					{
						stTransforma = "Desea crear un Pedido de Proveedor con esta Orden de Compra";
						idDestino = 14;
					}
					else
					{
						stTransforma = "Desea crear una Factura de Compra con esta orden de compra";
						idDestino = 4;
					}
					break;
				case 3:
					if (drSeteoF["CotizaReserva"] == System.DBNull.Value ||
						(bool) drSeteoF["CotizaReserva"] == false)
					{
						stTransforma = "Desea crear una Factura de Venta con esta Cotización";
						idDestino = 1;
					}
					else
					{
						stTransforma = "Desea crear un Pedido de Cliente con esta Cotización";
						idDestino = 13;
					}
					if (drSeteoF["CotizaEgreso"] != System.DBNull.Value &&
						(bool) drSeteoF["CotizaEgreso"] == true)
					{
						if (!chkAutPrincipal.Checked)
						{
							MessageBox.Show("No puede transformar una cotizacion a Egreso sin Autorizacion Principal", 
								"Informacion");
							return;
						}
						stTransforma = "Desea crear un Egreso con esta Cotización";
						idDestino = 8;
					}
          break;
				case 4:
					stTransforma = "Desea crear una Devolución de compra con esta Compra";
					idDestino = 6;
					break;
				case 5:
					stTransforma = "Desea Abonar esta devolucion a su cuenta por Cobrar";
					return;
				case 8:
					stTransforma = "Desea crear un Ingreso a Bodega con este Egreso de Bodega";
					idDestino = 9;
					break;
				case 9:
					stTransforma = "Desea crear un Egreso de Bodega con este Ingreso a bodega";
					idDestino = 8;
					break;
				case 11:
					stTransforma = "Desea crear una Factura de Venta con esta Reservacion del Cliente";
					idDestino = 1;
					break;
				case 12:
					stTransforma = "Desea crear un Egreso de Bodega de los componentes del Producto Terminado";
					break;
				case 26:
					stTransforma = "Desea crear una Factura de Compra con esta Reservacion del Proveedor";
					idDestino = 4;
					break;
				case 13:
					stTransforma = "Desea crear una Factura con este Pedido";
					idDestino = 1;
					break;
				case 14:
					stTransforma = "Desea crear un Ingreso de Bodega con este Pedido";
					idDestino = 9;
					break;
				case 31:
					stTransforma = "Desea crear un Ingreso de Bodega con esta Importacion en Transito";
					idDestino = 9;
					break;
				default:
					MessageBox.Show("Transformación no implementada");
					return;
			}
			if (DialogResult.Yes == MessageBox.Show(stTransforma, 
				"Confirmación", 
				MessageBoxButtons.YesNo, MessageBoxIcon.Question))
			{
				if (IdTipoFactura == 12) // Transforma Producto Terminado
				{
					string stExecPT = "Exec ConsumoFactura " + txtIdCompra.Value.ToString()
						+ ", 12";
					string stNegativoReceta = Funcion.sEscalarSQL(cdsArticulo, stExecPT, true);
					if (stNegativoReceta.Length > 0) 
					{
						MessageBox.Show(stNegativoReceta, "Informacion");
						return;
					}
					MessageBox.Show("Egreso de Bodega Generado", "Informacion");
					return;
				}
				if (IdTipoFactura == 3 && idDestino != 8) // Transforma Cotizacion
					{
					string stLimiteCredito = LimiteCredito(1);
					if (stLimiteCredito.Length > 0)
					{
						MessageBox.Show("Se infringe Cupo de Credito", "Informacion");
						return;
					}

					if (!(bool) drSeteoF["SaldoNegativo"] || ListaNegativos() == 0)
					{
						Cursor = Cursors.WaitCursor;
						string stExec = "EXEC TransformaCotizacion "
							+ txtIdCompra.Value.ToString();
						Funcion.EjecutaSQL(cdsArticulo, stExec, true);
						Compra miCompra = new Compra(idDestino);
						miCompra.MdiParent = this.MdiParent;
						miCompra.Show();
						Cursor = Cursors.Default;
						return;
					}
					return;
				}
				if (IdTipoFactura == 11 || IdTipoFactura == 26)
				{
					string stNum = Funcion.NumeraFactura(cdsCompra, idDestino);
					if (stNum.Trim().Length > 0) drCompra["Numero"] = stNum;
					drCompra["idTipoFactura"] = idDestino;
					cdsCompra.Update();
					IdTipoFactura = idDestino;
					int iSigno = 1;
					if (idDestino == 1) iSigno = -1;
					stNum = "Update detCompra Set Signo = " + iSigno.ToString()
						+ " Where idCompra = " + txtIdCompra.Value.ToString();
					Funcion.EjecutaSQL(cdsCompra, stNum, true);
					MessageBox.Show("Reservacion Transformada en factura", "Informacion");
					Compra2_Load(this, e);
					return;
				}

				this.Cursor = Cursors.WaitCursor;
				int IdOrigen = (int) txtIdCompra.Value;
				int idC = CopiarFactura(idDestino);
				if (idDestino == 9) // Ingreso de Bodega Copia serie de Egreso
				{
					string stCopiar = "Exec SerieCopiar " + IdOrigen.ToString()
						+ ", " + idC.ToString();
					Funcion.EjecutaSQL(cdsArticulo, stCopiar, true);
				}
			///////////////// SANTELMO
			if (Funcion.bEjecutaSQL(cdsArticulo, "Exec SeteoGExiste 'FTRSTELMO'"))
			    {
				Funcion.EjecutaSQL(cdsCompra, "EXEC CambiaBodega " 
					+ idC.ToString()  + ", " + idDestino.ToString());

			    }
				
			/////////////////////

				Compra miCompra1 = new Compra(idDestino, idC);
				miCompra1.MdiParent = this.MdiParent;
				miCompra1.Show();
				this.Cursor = Cursors.Default;
			}
		}

		private int ListaNegativos()
		{
			int num3 = 0;
			string text1 = "";
			foreach (UltraGridRow row1 in this.ultraGrid1.Rows.All)
			{
				int num1 = 0;
				double num2 = 0;
				if (row1.Cells["idArticulo"].Value != DBNull.Value)
				{
					num1 = (int) row1.Cells["idArticulo"].Value;
				}
				if (row1.Cells["Cantidad"].Value != DBNull.Value)
				{
					num2 = (double) row1.Cells["Cantidad"].Value;
				}
				if (num1 > 0)
				{
					string text2 = "Select IsNull(Saldo, 0) From Articulo Where idArticulo = " + num1.ToString();
					double num4 = Funcion.dEscalarSQL(this.cdsArticulo, text2, true);
					if (num4 < num2)
					{
						num3++;
						text2 = "Select Articulo From Articulo Where idArticulo = " + num1.ToString();
						text1 = text1 + Funcion.sEscalarSQL(this.cdsArticulo, text2, true);
						string text3 = text1;
						text1 = text3 + " Saldo " + num4.ToString("0.00", this.us) + " Pedido " + num2.ToString("0.00", this.us) + "\n";
					}
				}
			}
			if (num3 > 0)
			{
				MessageBox.Show(text1, "Informacion");
			}
			return num3;
		}

		private void btBuscaArticulo_Click(object sender, System.EventArgs e)
		{
			bool bExiste = true;
			try
			{
				if (ultraGrid1.ActiveRow.Cells["idArticulo"].Value == null 
					|| ultraGrid1.ActiveRow.Cells["idArticulo"].Value == DBNull.Value)
					bExiste = false;
			}
			catch
			{
				bExiste = false;
			}

			Articulo miArticulo;
			if (bExiste)

				if (Funcion.bEjecutaSQL(cdsArticulo, "Exec SeteoGExiste 'ARTMPREC'")) 				
				{
					miArticulo = new Articulo((int)  
						this.ultraGrid1.ActiveRow.Cells["idArticulo"].Value,
						(double) this.ultraGrid1.ActiveRow.Cells["Precio"].Value);
				}
				else				
				miArticulo = new Articulo((int)  
					this.ultraGrid1.ActiveRow.Cells["idArticulo"].Value);

			else
				miArticulo = new Articulo();

			miArticulo.ShowDialog();
			miArticulo.Dispose();
			miArticulo = null;
			Cursor = Cursors.WaitCursor;
			cdsArticulo.Clear();
			cdsArticulo.Fill();
			Cursor = Cursors.Default;
		}

		private void btSaldo_Click(object sender, System.EventArgs e)
		{
			if (ultraGrid1.ActiveRow == null ||
				ultraGrid1.ActiveRow.Cells["idArticulo"].Value == System.DBNull.Value)
			{
				MessageBox.Show("Seleccione el articulo que desea ver el saldo");
				return;
			}
			this.Cursor = Cursors.WaitCursor;
			int IdArticulo = (int) this.ultraGrid1.ActiveRow.Cells["idArticulo"].Value;
			string stSelect = "Select Articulo + ' ' + Codigo From Articulo Where IdArticulo = " + IdArticulo.ToString();
			string stArticulo = Funcion.sEscalarSQL(cdsArticulo, stSelect, true);
			if (Funcion.bEjecutaSQL(cdsArticulo, "Exec SeteoGExiste 'FPVCNSP'"))
			{
				CompraSaldoArt miSaldo = new CompraSaldoArt(IdArticulo, stArticulo);
				this.Cursor = Cursors.Default;
				miSaldo.ShowDialog();
				return;
			}

			// Para evitar el error cuando no existe articulo IdArt = 0
			double fSaldo = 0;
			stSelect = "exec CalculoSaldo " + IdArticulo.ToString();
			fSaldo = Funcion.dEscalarSQL(cdsArticulo, stSelect, true);

			this.Cursor = Cursors.Default;
			MessageBox.Show("Saldo de " + stArticulo.Trim() + ": "+ fSaldo.ToString("#,##0.00"));
		}

		private void Compra_Resize(object sender, System.EventArgs e)
		{
			this.ultraGrid1.DisplayLayout.AutoFitColumns = true;
		}

		private void txtNumero_DoubleClick(object sender, System.EventArgs e)
		{
			Asiento miAsiento = new Asiento((int) this.txtIdAsiento.Value);
			miAsiento.MdiParent = this.MdiParent;
			miAsiento.Show();		
		}

		private bool bCargaArticulo = true;
		private void cdsArticulo_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			this.cdsArticulo.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);

			string stFiltro = "[Descontinuado] = 0";
			if (IdTipoFactura == 1) stFiltro += " And dVenta <> 1";
			if (IdTipoFactura == 4) stFiltro += " And dCompra <> 1";
			string stSelect = "Select IsNull(CargaArticulo, 1) From SeteoF";
			bCargaArticulo = Funcion.bEscalarSQL(cdsArticulo, stSelect, true);
			if (!bCargaArticulo)
        stFiltro = "idArticulo < 2";
			C1.Data.SchemaObjects.DataSetDef dataSetDef = 
				e.DataSet.Schema.DataSetDefs["dsArticulos"];
			e.Filter.Add(new C1.Data.FilterCondition(dataSetDef.TableViews["Articulo"],
				stFiltro));
		}

		private void cdsComprobante_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsCompraTabla.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirNumero);
		}

		private void cdsProyecto_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsProyecto.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);
		}

		private void cdsSeguridad_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeguridad.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);		
		}

		private void btRetencion_Click(object sender, System.EventArgs e)
		{
			if (barraDato1.bEditar) 
			{
				return;
			}
			if (!miAcceso.Retencion)
			{
				MessageBox.Show("No tiene acceso a creación de retenciones",
					"Seguridad",MessageBoxButtons.OK, MessageBoxIcon.Information);
				return;
			}
//			if ((int)rdgContadoCred.Value == 2)
//			{
//				if ((double) drCompra["Saldo"] < 0.01)
//				{
//					if (DialogResult.No == MessageBox.Show("Factura cancelada totalmente. Desea Crear Retencion?",
//						"Seguridad",MessageBoxButtons.YesNo, MessageBoxIcon.Question)) return;
//				}
//			}

			this.Cursor = Cursors.WaitCursor;
			if ((int) this.cmbTipoCompra.Value == 4) // Compra
			{
				using (EgresoAT miEgreso = new EgresoAT())
				{
					miEgreso.IdFactura = (int) this.txtIdCompra.Value;
					miEgreso.drFactura = this.drCompra;
					miEgreso.ShowDialog();
				}
				barraDato1.ProximoId(5);
			}

			if ((int) this.cmbTipoCompra.Value == 1) // Venta
			{
				using(IngresoAT miIngreso = new IngresoAT())
				{
					miIngreso.IdFactura = (int) txtIdCompra.Value;
					miIngreso.drFactura = drCompra;
					miIngreso.ShowDialog();
				}
				
				barraDato1.ProximoId(5);
			}
			this.Cursor = Cursors.Default;
		}

		private double SaldoMora()
		{
			DateTime dtFecha = DateTime.Today;
			string stSelect = "Select IsNull(Sum(Compra.Saldo), 0) From Compra "
				+ "Where idCliente = " + cmbCliente.Value.ToString() 
				+ " And Borrar = 0 And idTipoFactura = 1 And Compra.FechaVencimiento < '"
				+ dtFecha.ToString("yyyyMMdd") + "'";
			double dSaldo = Funcion.dEscalarSQL(cdsCompra, stSelect, true);
			return dSaldo;
		}

		private bool bServicioAuto = false;
		private void btCliente_Click(object sender, System.EventArgs e)
		{
			bool bCliente = true;
			int IdCliente = 0;
			if (this.stbProveedor.Trim() != "= 0") bCliente = false;
			Cliente miCliente;
			if (this.cmbCliente.Value == System.DBNull.Value)
				miCliente = new Cliente(bCliente);
			else
			{
				IdCliente = (int) this.cmbCliente.Value;
				miCliente = new Cliente(bCliente, IdCliente);
			}

			miCliente.ShowDialog();
			miCliente.Dispose();
			miCliente = null;
			this.cdsCliente.Clear();
			this.cdsCliente.Fill();
			this.cmbCliente.Refresh();
			if (barraDato1.bNuevo)
			{
				try
				{
					cmbCliente.Value = int.Parse(miCliente.txtIdCliente.Value.ToString());
				}
				catch{};
			}
			cmbCliente.Select();
		}

		private void cmdTotal()
		{
			#region Creacion de Variables
			double dSubtIva = 0;
			double dSubtIva0 = 0;
			double dSubTotal = 0;
			double dImpuesto = 0;
			double dDescIva = 0;
			double dDescIva0 = 0;
			double dDescuento = 0;
			double dCantidad = 0;
			double dPrecio = 0;
			double dIva = 0;
			int dRedondeo = (int) drSeteoF["DecRedondeo"];
			#endregion Creacion de Variables
			#region Desplazamiento de grilla
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in ultraGrid1.Rows.All)
			{
				dCantidad = 0;
				dPrecio = 0;
				dIva = 0;
				dDescuento = 0;
				if (dr.Cells["Cantidad"].Value != System.DBNull.Value)
					dCantidad = (double) dr.Cells["Cantidad"].Value;
				if (dr.Cells["Precio"].Value != System.DBNull.Value)
					dPrecio = (double) dr.Cells["Precio"].Value;
				if (dr.Cells["Impuesto"].Value != System.DBNull.Value)
					dIva = (double) dr.Cells["Impuesto"].Value;
				if (dr.Cells["DescuentoArt"].Value != System.DBNull.Value)
					dDescuento = (double) dr.Cells["DescuentoArt"].Value;
				dSubTotal = dCantidad * dPrecio + .000001;
				dSubTotal = Math.Round(dSubTotal, dRedondeo);
				if (dIva > 0)
				{
					if (bIncluyeIva)
					{
						double dIvaLinea = Math.Round(dSubTotal - dSubTotal /(1+dIva/100.0), dRedondeo);
						double dSubTotalL = dSubTotal - dIvaLinea;
						dSubtIva += dSubTotal - dIvaLinea;
						dImpuesto += (dSubTotalL - dDescuento) * dIva / 100.0;
					}
					else
					{
						dSubtIva += Math.Round(dSubTotal, dRedondeo);
						dImpuesto += (dSubTotal - dDescuento) * dIva / 100.0;
					}

					dDescIva += dDescuento;
				}
				else
				{
					dSubtIva0 += dSubTotal;
					dDescIva0 += dDescuento;
				}
			}
			#endregion Desplazamiento de grilla
			#region Calculo de Descuentos
			if (txtDescPorcIva.Value != System.DBNull.Value && (double) txtDescPorcIva.Value > 0)
				drCompra["Desc1"] = (double) txtDescPorcIva.Value * (dSubtIva - dDescIva) / 100;
			if (txtDescPorcIva0.Value != System.DBNull.Value && (double) txtDescPorcIva0.Value > 0)
				drCompra["Desc2"] = (double) txtDescPorcIva0.Value * (dSubtIva0 - dDescIva0) / 100;
			drCompra["Desc3"] = dDescIva;
			drCompra["Desc4"] = dDescIva0;

			dDescuento = dDescIva0 + dDescIva;
			dDescIva0 = 0;
			dDescIva = 0;
			dDescIva = (double) drCompra["Desc1"];
			dDescIva0 = (double) drCompra["Desc2"];
			dDescuento += dDescIva0 + dDescIva;
			drCompra["Descuento"] = dDescIva + (double) drCompra["Desc3"];
			drCompra["Descuento0"] = dDescIva0 + (double) drCompra["Desc4"];
			#endregion Calculo de Descuentos

			// Calculo de impuesto con descuento
			//			dImpuesto = (dSubtIva - (double) drCompra["Descuento"]) 
			//				* (double) drSeteoF["IVA"]/100.0;
			dDescuento = Math.Round(dDescuento, dRedondeo);
			dSubtIva0 += (double) this.txtFlete.Value;
			dSubtIva0 = Math.Round(dSubtIva0, dRedondeo);
			dSubtIva = Math.Round(dSubtIva, dRedondeo);
		
			//			this.drCompra["Descuento"] = dDescuento;
			this.drCompra["SubTotalExcento"] = dSubtIva0;
			this.drCompra["SubTotalIva"] = dSubtIva;
			if ((bool) drSeteoF["IncluyeIva"])
				drCompra["SubTotalIva"] = dSubtIva;
			else
				drCompra["SubTotalIva"] = dSubtIva;

			double dSubtIvaAntes = dSubtIva - (double) drCompra["Desc3"];
			if ((double) drCompra["Desc1"] > 0 && dSubtIvaAntes > 0)
				dImpuesto -= dImpuesto * (1 - (dSubtIvaAntes - (double) drCompra["Desc1"])/ dSubtIvaAntes);
			dImpuesto = Math.Round(dImpuesto, dRedondeo);
			if (txtServicio.Width > 0 && IdTipoFactura == 1 && bServicioAuto) 
			{
				txtServicio.Value = (dSubtIva0 + dSubtIva- dDescuento) * 0.1 ;
				drCompra["Servicio"] = (dSubtIva0 + dSubtIva- dDescuento) * 0.1 ;
			}
			this.drCompra["Iva"] = dImpuesto;
			double dTotal = dSubtIva0 + dSubtIva + dImpuesto - dDescuento 
				+ (double) this.txtIceTotal.Value 
				+ (double) this.txtServicio.Value
				+ (double) this.txtPropina.Value;

			dTotal = Math.Round(dTotal,dRedondeo);
			if(dTotal <0)	
			{
				this.drCompra["Total"] = 0;
				this.errorProvider.SetError(this.txtTotal, "Error en el calculo del total: " 
					+ dTotal.ToString("#,##0.00"));
			}
			else
			{
				this.drCompra["Total"] = dTotal;
				this.errorProvider.SetError(this.txtTotal, ""); 
			}
			//			this.Refresh();
		}


		private void btTotal_Click(object sender, System.EventArgs e)
		{
			this.cmdTotal();
		}

		private void CreaDetalle(int iCuenta, double dDebe, Double dHaber, int iCuentaP)
		{
			C1.Data.C1DataRow drDetAsi = this.cdsAsiento.TableViews["AsientoDetalle"].AddNew();
			drDetAsi["idCuenta"] = iCuenta;
			drDetAsi["idAsiento"] = iCuentaP;
			dDebe = Math.Round(dDebe, 2);
			dHaber = Math.Round(dHaber, 2);
			drDetAsi["Debe"] = dDebe;
			drDetAsi["Haber"] = dHaber;
		}

		private void CreaDetalle(int iCuenta, double dDebe, Double dHaber, int iCuentaP, string stDetalle)
		{
			C1.Data.C1DataRow drDetAsi = this.cdsAsiento.TableViews["AsientoDetalle"].AddNew();
			drDetAsi["idCuenta"] = iCuenta;
			drDetAsi["idAsiento"] = iCuentaP;
			dDebe = Math.Round(dDebe, 2);
			dHaber = Math.Round(dHaber, 2);
			drDetAsi["Debe"] = dDebe;
			drDetAsi["Haber"] = dHaber;
			drDetAsi["Detalle"] = stDetalle;
		}

		private void AsientoVenta(int idPrinc)
		{
			double dValor = 0;
			int iCuentaDeudora = 0;

			// Cuentas por Cobrar o caja o forma de pago o grupo de cliente
			if ((int) this.rdgContadoCred.Value == 1)
			{
				iCuentaDeudora = Contabilidad.AsignaCuenta(this.cdsAsignaCuenta,1); // Caja o Bancos
				// Cambio de cuenta si hay forma de pago
				if (this.cmbFormaPago.Value != System.DBNull.Value)
				{
					int iCuenta = Funcion.CuentaFormaPago(cdsCompra, (int) cmbFormaPago.Value);
					if (iCuenta > 0) iCuentaDeudora = iCuenta;
				}
			}
			else // Credito
			{
				iCuentaDeudora = Contabilidad.AsignaCuenta(cdsAsignaCuenta,4); // Cuenta por Cobrar
				// cambio de cuenta si hay grupo de clientes
				if (this.cmbCliente.Value != System.DBNull.Value)
				{
					int iCuenta = Funcion.CuentaGrupoCliente(cdsCompra, (int) cmbCliente.Value);
					if (iCuenta > 0) iCuentaDeudora = iCuenta;
				}
			} // Fin de credito

			dValor =	(double) txtTotal.Value + (double) txtInteresPagar.Value;
			dValor += - Funcion.RetAsiento(cdsCompra, int.Parse(txtIdCompra.Value.ToString()));

			// Manejo de Anticipos
			if ((double) drCompra["Contado"] > 0)
				dValor += - (double) drCompra["Contado"];

			CreaDetalle(iCuentaDeudora, dValor, 0, idPrinc);

			// Manejo de Anticipos
			if ((double) drCompra["Contado"] > 0)
			{
				dValor = (double) drCompra["Contado"];
				iCuentaDeudora = Contabilidad.AsignaCuenta(this.cdsAsignaCuenta,1); // Caja o Bancos
				CreaDetalle(iCuentaDeudora, dValor, 0, idPrinc);
			}

			// IVA en Ventas
			dValor = (double) this.txtIvaTotal.Value;
			if (dValor > 0)
			{
				iCuentaDeudora = Contabilidad.AsignaCuenta(this.cdsAsignaCuenta, 2); // Iva
				CreaDetalle(iCuentaDeudora, 0, dValor, idPrinc);
			}

			// Descuento en Ventas
			try
			{
				dValor = 0;
				dValor = (double) this.drCompra["Descuento"];
				dValor+= (double) this.drCompra["Descuento0"];
			}
			catch{}
			if (dValor > 0)
			{
				iCuentaDeudora = Contabilidad.AsignaCuenta(this.cdsAsignaCuenta,3);
				CreaDetalle(iCuentaDeudora, dValor, 0, idPrinc);
			}

			// Flete en Ventas
			dValor = (double) this.txtFlete.Value;
			if (dValor > 0)
			{
				iCuentaDeudora = Contabilidad.AsignaCuenta(this.cdsAsignaCuenta, 7);
				CreaDetalle(iCuentaDeudora, 0, dValor, idPrinc);
			}

			// ICE
			dValor = (double) this.txtIceTotal.Value;
			if (dValor > 0)
			{
				iCuentaDeudora = Contabilidad.AsignaCuenta(this.cdsAsignaCuenta, 9);
				CreaDetalle(iCuentaDeudora, 0, dValor, idPrinc);
			}
			// Interes
			dValor = (double) this.txtInteresPagar.Value;
			if (dValor > 0)
			{
				iCuentaDeudora = Contabilidad.AsignaCuenta(this.cdsAsignaCuenta, 6);
				CreaDetalle(iCuentaDeudora, 0, dValor, idPrinc);
			}

			// Crea Detalles de Asiento
			double dCantidad = 0;
			double dPrecio = 0;
			double dImpuesto = 0;
			double dDescArt = 0;
			int iIdArticulo = 0;

			if ((bool) drSeteoF["Inv_Unico"])
			{
				this.cdsAsiento.Update();
				Contabilidad.AsientoFactura(cdsCompra, (int) txtIdCompra.Value, idPrinc, 1, drSeteoF);
			}
			else
			{
				foreach(C1.Data.C1DataRow drDet in drCompra.GetChildRows())
				{
					dCantidad = (double) drDet["Cantidad"];
					dPrecio = (double) drDet["Precio"];
					dImpuesto = (double) drDet["Impuesto"];
					dDescArt = (double) drDet["DescuentoArt"];
					iIdArticulo = (int) drDet["idArticulo"];
					dValor = dCantidad * dPrecio - dDescArt;
					if (dValor > 0)
					{
						iCuentaDeudora = Funcion.AsignaCuenta(cdsArticulo, iIdArticulo, 3);
						string stSel = "Select Codigo From Articulo Where idArticulo = "
							+ iIdArticulo.ToString();
						stSel = Funcion.sEscalarSQL(cdsArticulo, stSel, true);
						CreaDetalle(iCuentaDeudora, 0, dValor, idPrinc, stSel);
					}
				} // Fin de for para crear linea de asiento por cada linea
			}

			if (!(bool) drSeteoF["CostoManual"]) // 
			{
				foreach(C1.Data.C1DataRow drDet in drCompra.GetChildRows())
				{
					dCantidad = (double) drDet["Cantidad"];
					iIdArticulo = (int) drDet["idArticulo"];
					dPrecio = Funcion.CostoArticulo(cdsArticulo, iIdArticulo);
					dValor = dCantidad * dPrecio;
					if (dValor > 0)
					{
						iCuentaDeudora = Funcion.AsignaCuenta(cdsArticulo, iIdArticulo, 2);
						CreaDetalle(iCuentaDeudora, dValor, 0, idPrinc);
						iCuentaDeudora = Funcion.AsignaCuenta(cdsArticulo, iIdArticulo, 4);
						CreaDetalle(iCuentaDeudora, 0, dValor, idPrinc);
					}
				}
			}
			//*************** Fin de asiento de venta	************

		}

		private void AsientoCompra(int idPrinc)
		{
			// Crea Detalles de Asiento
			double dCantidad = 0;
			double dPrecio = 0;
			double dImpuesto = 0;
			int iIdArticulo = 0;
			double dValor = 0;
			int iCuentaDeudora = 0;

			if ((bool) drSeteoF["Inv_Unico"])
			{
				Contabilidad.AsientoFactura(cdsCompra, (int) txtIdCompra.Value, idPrinc, 2, drSeteoF);
			}
			else
			{
				foreach(C1.Data.C1DataRow drDet in drCompra.GetChildRows())
				{
					dCantidad = (double) drDet["Cantidad"];
					dPrecio = (double) drDet["Precio"];
					dImpuesto = (double) drDet["Impuesto"];
					//				dDescArt = (double) drDet["DescuentoArt"];
					iIdArticulo = (int) drDet["idArticulo"];
					dValor = dCantidad * dPrecio;
					if (dValor > 0)
					{
						iCuentaDeudora = Funcion.AsignaCuenta(cdsArticulo, iIdArticulo, 2);
						CreaDetalle(iCuentaDeudora, dValor, 0, idPrinc);
					}
				}
			}

			// IVA en Compras
			dValor = (double) txtIvaTotal.Value;
			if (dValor > 0)
			{
				iCuentaDeudora = Contabilidad.AsignaCuenta(cdsAsignaCuenta, 11); // Iva
				CreaDetalle(iCuentaDeudora, dValor, 0, idPrinc);
			}

			// Intereses en Compras
			dValor = (double) txtInteresPagar.Value;
			if (dValor > 0)
			{
				iCuentaDeudora = Contabilidad.AsignaCuenta(cdsAsignaCuenta, 15); // Interes
				CreaDetalle(iCuentaDeudora, dValor, 0, idPrinc);
			}
			
			// Flete en Compras
			dValor = (double) txtFlete.Value;
			if (dValor > 0)
			{
				iCuentaDeudora = Contabilidad.AsignaCuenta(cdsAsignaCuenta, 16);
				CreaDetalle(iCuentaDeudora, dValor, 0, idPrinc);
			}

			// ICE en compras
			dValor = (double) txtIceTotal.Value;
			if (dValor > 0)
			{
				iCuentaDeudora = Contabilidad.AsignaCuenta(cdsAsignaCuenta, 18);
				CreaDetalle(iCuentaDeudora, dValor, 0, idPrinc);
			}

			// Descuento en Compras
			dValor = (double) drCompra["Descuento"] + (double) drCompra["Descuento0"];
			if (dValor > 0)
			{
				iCuentaDeudora = Contabilidad.AsignaCuenta(cdsAsignaCuenta, 12);
				CreaDetalle(iCuentaDeudora, 0, dValor, idPrinc);
			}

			// Cuentas por Pagar o caja o forma de pago o grupo de cliente
			if ((int) rdgContadoCred.Value == 1)
			{
				iCuentaDeudora = Contabilidad.AsignaCuenta(cdsAsignaCuenta, 10); // Caja o Bancos
				// Cambio de cuenta si hay forma de pago
				if (cmbFormaPago.Value != System.DBNull.Value)
				{
					int iCuenta = Funcion.CuentaFormaPago(cdsCompra, (int) cmbFormaPago.Value);
					if (iCuenta > 0) iCuentaDeudora = iCuenta;
				}
			}
			else // Credito
			{
				iCuentaDeudora = Contabilidad.AsignaCuenta(cdsAsignaCuenta, 13); // Cuenta por Pagar
				// cambio de cuenta si hay grupo de clientes
				if (cmbCliente.Value != System.DBNull.Value)
				{
					int iCuenta = Funcion.CuentaGrupoCliente(cdsCompra, (int) cmbCliente.Value);
					if (iCuenta > 0) iCuentaDeudora = iCuenta;
				}
			}

			dValor =	(double) txtTotal.Value + (double) txtInteresPagar.Value;
			//			if ((bool) drCompra["Retencion"])
			//				dValor += - (double) drCompra["Retenido"];
			dValor += - Funcion.RetAsiento(cdsCompra, int.Parse(txtIdCompra.Value.ToString()));

			CreaDetalle(iCuentaDeudora, 0, dValor, idPrinc);

			//*************** Fin de asiento de Compra	************
		}


		private void AsientoDevCompra(int idPrinc)
		{
			double dValor = 0;
			int iCuentaDeudora = 0;

			// Cuentas por Cobrar o caja o forma de pago o grupo de cliente
			if ((int) this.rdgContadoCred.Value == 1)
			{
				iCuentaDeudora = Contabilidad.AsignaCuenta(this.cdsAsignaCuenta,10); // Caja o Bancos
				// Cambio de cuenta si hay forma de pago
				if (this.cmbFormaPago.Value != System.DBNull.Value)
				{
					int iCuenta = Funcion.CuentaFormaPago(this.cdsCompra, (int) this.cmbFormaPago.Value);
					if (iCuenta > 0) iCuentaDeudora = iCuenta;
				}
			}
			else // Credito
			{
				iCuentaDeudora = Contabilidad.AsignaCuenta(this.cdsAsignaCuenta,13); // Cuenta por Cobrar
				// cambio de cuenta si hay grupo de clientes
				if (this.cmbCliente.Value != System.DBNull.Value)
				{
					int iCuenta = Funcion.CuentaGrupoCliente(this.cdsCompra, (int) this.cmbCliente.Value);
					if (iCuenta > 0) iCuentaDeudora = iCuenta;
				}
			} // Fin de credito

			dValor =	(double) this.txtTotal.Value + (double) txtInteresPagar.Value;
			if ((bool) this.drCompra["Retencion"])
				dValor += - (double) this.drCompra["Retenido"];

			// Manejo de Anticipos
			if ((double) drCompra["Contado"] > 0)
				dValor += - (double) drCompra["Contado"];


			CreaDetalle(iCuentaDeudora, dValor, 0, idPrinc);

			// Manejo de Anticipos
			if ((double) drCompra["Contado"] > 0)
			{
				dValor = (double) drCompra["Contado"];
				iCuentaDeudora = Contabilidad.AsignaCuenta(this.cdsAsignaCuenta,10); // Caja o Bancos
				CreaDetalle(iCuentaDeudora, dValor, 0, idPrinc);
			}

			// IVA en Compras
			dValor = (double) this.txtIvaTotal.Value;
			if (dValor > 0)
			{
				iCuentaDeudora = Contabilidad.AsignaCuenta(this.cdsAsignaCuenta, 11); // Iva
				CreaDetalle(iCuentaDeudora, 0, dValor, idPrinc);
			}

			// Descuento en Compras
			try
			{
				dValor = 0;
				dValor = (double) this.drCompra["Descuento"];
				dValor+= (double) this.drCompra["Descuento0"];
			}
			catch{}
			if (dValor > 0)
			{
				iCuentaDeudora = Contabilidad.AsignaCuenta(this.cdsAsignaCuenta,12);
				CreaDetalle(iCuentaDeudora, dValor, 0, idPrinc);
			}

			// Flete en Compras
			dValor = (double) this.txtFlete.Value;
			if (dValor > 0)
			{
				iCuentaDeudora = Contabilidad.AsignaCuenta(this.cdsAsignaCuenta, 16);
				CreaDetalle(iCuentaDeudora, 0, dValor, idPrinc);
			}

			// ICE en Compras
			dValor = (double) this.txtIceTotal.Value;
			if (dValor > 0)
			{
				iCuentaDeudora = Contabilidad.AsignaCuenta(this.cdsAsignaCuenta, 18);
				CreaDetalle(iCuentaDeudora, 0, dValor, idPrinc);
			}

			// Interes en Compras
			dValor = (double) this.txtInteresPagar.Value;
			if (dValor > 0)
			{
				iCuentaDeudora = Contabilidad.AsignaCuenta(this.cdsAsignaCuenta, 15);
				CreaDetalle(iCuentaDeudora, 0, dValor, idPrinc);
			}

			// Crea Detalles de Asiento
			double dCantidad = 0;
			double dPrecio = 0;
			double dImpuesto = 0;
			double dDescArt = 0;
			int iIdArticulo = 0;

			if ((bool) drSeteoF["Inv_Unico"])
			{
				this.cdsAsiento.Update();
				Contabilidad.AsientoFactura(cdsCompra, (int) txtIdCompra.Value, idPrinc, 4, drSeteoF);
			}
			else
			{
				foreach(C1.Data.C1DataRow drDet in drCompra.GetChildRows())
				{
					dCantidad = (double) drDet["Cantidad"];
					dPrecio = (double) drDet["Precio"];
					dImpuesto = (double) drDet["Impuesto"];
					dDescArt = (double) drDet["DescuentoArt"];
					iIdArticulo = (int) drDet["idArticulo"];
					dValor = dCantidad * dPrecio - dDescArt;
					if (dValor > 0)
					{
						iCuentaDeudora = Funcion.AsignaCuenta(cdsArticulo, iIdArticulo, 2);
						CreaDetalle(iCuentaDeudora, 0, dValor, idPrinc);
					}
				} // Fin de for para crear linea de asiento por cada linea
			}


			if (!(bool) drSeteoF["CostoManual"]) // 
			{
				foreach(C1.Data.C1DataRow drDet in drCompra.GetChildRows())
				{
					dCantidad = (double) drDet["Cantidad"];
					iIdArticulo = (int) drDet["idArticulo"];
					dPrecio = Funcion.CostoArticulo(this.cdsArticulo, iIdArticulo);
					dValor = dCantidad * dPrecio;
					if (dValor > 0)
					{
						iCuentaDeudora = Funcion.AsignaCuenta(cdsArticulo, iIdArticulo, 2);
						CreaDetalle(iCuentaDeudora, dValor, 0, idPrinc);
						iCuentaDeudora = Funcion.AsignaCuenta(cdsArticulo, iIdArticulo, 4);
						CreaDetalle(iCuentaDeudora, 0, dValor, idPrinc);
					}
				}
			}
			//*************** Fin de asiento de Devolucion en Compra	************
		}

		private void AsientoDevVenta(int idPrinc)
		{
			// Crea Detalles de Asiento
			double dCantidad = 0;
			double dPrecio = 0;
			double dImpuesto = 0;
			int iIdArticulo = 0;
			double dValor = 0;
			int iCuentaDeudora = 0;

			if ((bool) drSeteoF["Inv_Unico"])
			{
				Contabilidad.AsientoFactura(cdsCompra, (int) txtIdCompra.Value, idPrinc, 5, drSeteoF);
			}
			else
			{
				foreach(C1.Data.C1DataRow drDet in drCompra.GetChildRows())
				{
					dCantidad = (double) drDet["Cantidad"];
					dPrecio = (double) drDet["Precio"];
					dImpuesto = (double) drDet["Impuesto"];
					//				dDescArt = (double) drDet["DescuentoArt"];
					iIdArticulo = (int) drDet["idArticulo"];
					dValor = dCantidad * dPrecio;
					if (dValor > 0)
					{
						iCuentaDeudora = Funcion.AsignaCuenta(cdsArticulo, iIdArticulo, 6);
						CreaDetalle(iCuentaDeudora, dValor, 0, idPrinc);
					}
				}
			}

			// IVA en Ventas
			dValor = (double) txtIvaTotal.Value;
			if (dValor > 0)
			{
				iCuentaDeudora = Contabilidad.AsignaCuenta(cdsAsignaCuenta, 2); // Iva
				CreaDetalle(iCuentaDeudora, dValor, 0, idPrinc);
			}

			// Intereses en Ventas
			dValor = (double) txtInteresPagar.Value;
			if (dValor > 0)
			{
				iCuentaDeudora = Contabilidad.AsignaCuenta(cdsAsignaCuenta, 6); // Interes
				CreaDetalle(iCuentaDeudora, dValor, 0, idPrinc);
			}
			
			// Flete en Ventas
			dValor = (double) txtFlete.Value;
			if (dValor > 0)
			{
				iCuentaDeudora = Contabilidad.AsignaCuenta(cdsAsignaCuenta, 7);
				CreaDetalle(iCuentaDeudora, dValor, 0, idPrinc);
			}

			// ICE en Ventas
			dValor = (double) txtIceTotal.Value;
			if (dValor > 0)
			{
				iCuentaDeudora = Contabilidad.AsignaCuenta(cdsAsignaCuenta, 9);
				CreaDetalle(iCuentaDeudora, dValor, 0, idPrinc);
			}

			// Descuento en Ventas
			dValor = (double) drCompra["Descuento"] + (double) drCompra["Descuento0"];
			if (dValor > 0)
			{
				//				iCuentaDeudora = Contabilidad.AsignaCuenta(cdsAsignaCuenta, 3);
				iCuentaDeudora = Funcion.AsignaCuenta(cdsArticulo, iIdArticulo, 5);
				CreaDetalle(iCuentaDeudora, 0, dValor, idPrinc);
			}

			// Cuentas por Cobrar o caja o forma de pago o grupo de cliente
			if ((int) rdgContadoCred.Value == 1)
			{
				iCuentaDeudora = Contabilidad.AsignaCuenta(cdsAsignaCuenta, 1); // Caja o Bancos
				// Cambio de cuenta si hay forma de pago
				if (cmbFormaPago.Value != System.DBNull.Value)
				{
					int iCuenta = Funcion.CuentaFormaPago(cdsCompra, (int) cmbFormaPago.Value);
					if (iCuenta > 0) iCuentaDeudora = iCuenta;
				}
			}
			else // Credito
			{
				iCuentaDeudora = Contabilidad.AsignaCuenta(cdsAsignaCuenta, 4); // Cuenta por Pagar
				// cambio de cuenta si hay grupo de clientes
				if (cmbCliente.Value != System.DBNull.Value)
				{
					int iCuenta = Funcion.CuentaGrupoCliente(cdsCompra, (int) cmbCliente.Value);
					if (iCuenta > 0) iCuentaDeudora = iCuenta;
				}
			}

			dValor =	(double) txtTotal.Value + (double) txtInteresPagar.Value ;
			if ((bool) drCompra["Retencion"])
				dValor += - (double) drCompra["Retenido"];

			CreaDetalle(iCuentaDeudora, 0, dValor, idPrinc);

			//*************** Fin de asiento de devolucon de venta	************
		}


		private void btCreaAsiento_Click(object sender, System.EventArgs e)
		{
			if (barraDato1.bEditar) 
			{
				MessageBox.Show("No puede crear un asiento mientras edita la factura");
				return;
			}
			if ((bool) drCompra["Borrar"] == true) 
			{
				MessageBox.Show("No puede crear un asiento de una transaccion Anulada");
				return;
			}

			if (!(drSeteoC["FechaFin"] == DBNull.Value || drSeteoC["FechaIni"] == DBNull.Value))
			{
				DateTime dtFechaAsi = (DateTime)dtFecha.Value;
				DateTime dtFechaAFin = (DateTime)drSeteoC["FechaFin"];
				DateTime dtFechaAIni = (DateTime)drSeteoC["FechaIni"];
				if (dtFechaAsi > dtFechaAFin || dtFechaAsi < dtFechaAIni)
				{
					MessageBox.Show("Fecha debe estar entre el " 
						+ dtFechaAIni.ToString("dd/MMM/yyyy")+ " y el " 
						+ dtFechaAFin.ToString("dd/MMM/yyyy") + ".\n"
						+ "Fechas definidas en Archivo - Propiedades - Contabilidad", "Informacion");
					return;
				}
			}

			string stProcedimiento = Funcion.sEscalarSQL(cdsArticulo, 
				"Exec CompraAsientoTipo " + txtIdCompra.Value.ToString(), true);

			if (stProcedimiento.Length == 0)
			{
				MessageBox.Show("Tipo de Transaccion no genera Asiento", "Informacion", 
					MessageBoxButtons.OK, MessageBoxIcon.Information);
				return;
			}
			// Crea asiento de Factura
			if (!miAcceso.CrearAsiento)
			{
				MessageBox.Show("No tiene Acceso a Crear Asientos", "Seguridad", 
					MessageBoxButtons.OK, MessageBoxIcon.Information);
				return;
			}

			if (btCreaAsiento.Visible)
				if (DialogResult.No == MessageBox.Show("¿Desea crear Asiento?", "Confirmación", 
					MessageBoxButtons.YesNo, MessageBoxIcon.Question))
					return;

			this.Cursor = Cursors.WaitCursor;
			#region verificacion de Existencia
			string stExisteAsiento = "Exec AsientoMaestroExiste "	+ txtIdCompra.Value.ToString();
			string stMensaje = Funcion.sEscalarSQL(cdsAsiento, stExisteAsiento, true);
			if (stMensaje.Length > 0)
			{
				MessageBox.Show(stMensaje, "Informacion", 
					MessageBoxButtons.OK, MessageBoxIcon.Warning);
				this.Cursor = Cursors.Default;
				return;
			}
			stMensaje = Funcion.sEscalarSQL(cdsAsiento, stExisteAsiento + ", 2", true);
			if (stMensaje.Length > 0)
			{
				if (DialogResult.No == MessageBox.Show(stMensaje, "Confirmación", 
					MessageBoxButtons.YesNo, MessageBoxIcon.Asterisk))
				{
					this.Cursor = Cursors.Default;
					return;
				}
				stMensaje = Funcion.sEscalarSQL(cdsAsiento, stExisteAsiento + ", 3", true);
			}
			#endregion verificacion

			string stGeneraCompra = "EXEC " + stProcedimiento + txtIdCompra.Value.ToString();
			Funcion.EjecutaSQL(cdsAsiento, stGeneraCompra, true);
			MessageBox.Show("Asiento Generado con éxito", "Informacion",
				MessageBoxButtons.OK, MessageBoxIcon.Information);
			barraDato1.ProximoId(5);
			Cursor = Cursors.Default;
		}

		private void cdsAsiento_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			this.cdsAsiento.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirConta);
		}

		private void cdsRetencion_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			this.cdsRetencion.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirRetenc);
		}

		private void cdsAsignaCuenta_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			this.cdsAsignaCuenta.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);
		}

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			this.cdsSeteoF.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);		
		}

		private void btCreaAsiento_MouseDown(object sender, System.Windows.Forms.MouseEventArgs e)
		{
			if (e.Button != MouseButtons.Right) return;
			if (txtIdAsiento.Value == System.DBNull.Value || (int) txtIdAsiento.Value == 0)
			{
				MessageBox.Show("No existe un asiento creado de esta factura");
				return;
			}
			int IdAsiento = (int) txtIdAsiento.Value;
			if (drCompra["idAsientoPer"] != DBNull.Value && (int) drCompra["idAsientoPer"] > 0)
			{
				MessageBox.Show("Origen es un Asiento de Comprobante", "Informacion");
				IdAsiento = (int) drCompra["idAsientoPer"];
			}
			
			Cursor = Cursors.WaitCursor;
			Asiento miAsiento = new Asiento(IdAsiento);
			miAsiento.MdiParent = MdiParent;
			miAsiento.Show();
			Cursor = Cursors.Default;
		}

		private void cmdFormaPago_Click(object sender, System.EventArgs e)
		{
			CompraFormaPago miFormaPago = new CompraFormaPago();
//			miFormaPago.MdiParent = this.MdiParent;
			miFormaPago.ShowDialog();
			cdsCompraTabla.Clear();
			cdsCompraTabla.Fill();
		}

		/// <summary>
		/// Funcion para calcular el porcentaje de descuento
		/// </summary>
		/// <param name="sender"></param>
		/// <param name="e"></param>
		private void btRefrescaD_Click(object sender, System.EventArgs e)
		{
			double dTotalIva = (double) this.txtIva.Value; // Valor que paga Iva de la factura
			dTotalIva -= (double) txtDescLinea.Value;
			double dDescIva = (double) this.txtDescIva.Value;
			double dPorcIva = 0;
			if (dTotalIva != 0) dPorcIva = dDescIva / (dTotalIva) * 100.0;
			this.drCompra["DescPorcIva"] = dPorcIva;

			double dTotalIva0 = (double) this.txtIva0.Value; // Valor que paga Iva 0 de la factura
			double dDescIva0 = (double) this.txtDescIva0.Value;
			double dDescIvaLinea0 = (double) this.txtDescLinea0.Value;
			double dPorcIva0 = 0;
			if (dTotalIva0 != 0) dPorcIva0 = dDescIva0 / (dTotalIva0 - dDescIvaLinea0) * 100.0;
			this.drCompra["DescPorcIva0"] = dPorcIva0;
		}

		private void btRefresaI_Click(object sender, System.EventArgs e)
		{
			double dTotalIva = (double) this.txtIva.Value; // Valor que paga Iva de la factura
			dTotalIva -= (double) txtDescLinea.Value;
			double dDescPIva = (double) this.txtDescPorcIva.Value;
			double dDescIva = dDescPIva *  dTotalIva / 100.0;
			this.drCompra["Desc1"] = dDescIva;

			double dTotalIva0 = (double) this.txtIva0.Value; // Valor que no paga Iva de la factura
			dTotalIva0 -= (double) txtDescLinea0.Value;
			double dDescPIva0 = (double) this.txtDescPorcIva0.Value;
			double dDescIva0 = dDescPIva0 * dTotalIva0 / 100.0;
			this.drCompra["Desc2"] = dDescIva0;
		}

		private void barraDato1_Impresora(object sender, System.EventArgs e)
		{
			string stCopia = "Select IsNull(Copias, 0) From CompraNumero Where idTipoFactura = " 
				+ IdTipoFactura.ToString();
			int iCopias = Funcion.iEscalarSQL(cdsCompraTabla, stCopia);
			if (iCopias < 1) iCopias = 1;
			if (IdTipoFactura == 1 && drSeteoF["TiempoVenta"] != System.DBNull.Value
				&& (int) drSeteoF["TiempoVenta"] > 0)
			{
				double dTotal = (double) txtTotal.Value;
				double dRecibido = (double) spnVuelto.Value;
				CompraGracias miGracias = new CompraGracias((int) drSeteoF["TiempoVenta"], dTotal, dRecibido);
				miGracias.MdiParent = this.MdiParent;
				miGracias.Show();
			}
//			for (int i=0; i<iCopias; i++)
			
				Imprimir(iCopias);		
			//miGracias.Dispose();
		}

		private void btPagos_Click(object sender, System.EventArgs e)
		{
			// Pagos solo existen en ciertos tipos de comprobantes
			if (IdTipoFactura != 1 && IdTipoFactura != 4 &&
				IdTipoFactura != 27 && IdTipoFactura != 28 &&
				IdTipoFactura != 5 && IdTipoFactura != 6 &&
				IdTipoFactura != 11 && IdTipoFactura != 26 && IdTipoFactura != 99) return;

			if (barraDato1.bEditar)
			{
				MessageBox.Show("No puede Crear pagos mientras edita la factura");
				return;
			}

			if ((int) this.rdgContadoCred.Value != 2) 
			{
				MessageBox.Show("Pagos solo se puede realizar en Creditos",
					"Informacion", MessageBoxButtons.OK, MessageBoxIcon.Information);
				return;
			}
			Pagos miPago = new Pagos(this.IdTipoFactura, (int) this.txtIdCompra.Value);
			miPago.MdiParent = this.MdiParent;
			miPago.Show();
		}

		private void barraDato1_Restaurar(object sender, System.EventArgs e)
		{
			if (!miAcceso.Recuperar)
			{
				MessageBox.Show("No tiene Acceso a Restaurar " + this.Text.Trim(),
					"Seguridad", MessageBoxButtons.OK, MessageBoxIcon.Information);
				return;
			}
			drCompra = cdsCompra.TableViews["Compra"].Rows[0]; //.Find((int) this.txtIdCompra.Value);

			if (DialogResult.No == MessageBox.Show(
				"¿Está seguro que desea Restaurar el Registro?", 
				"Confirmación", MessageBoxButtons.YesNo, MessageBoxIcon.Information, 
				MessageBoxDefaultButton.Button2)) return;

			string stBorraSQL = "UPDATE Asiento SET Borrar = 0 WHERE idAsiento = "
				+ txtIdAsiento.Value.ToString();
			Contabilidad.EjecutaSQL(cdsAsiento, stBorraSQL);
			stBorraSQL = "UPDATE Retencion SET Borrar = 0 WHERE idCompra = "
				+ txtIdCompra.Value.ToString();
			Contabilidad.EjecutaSQL(cdsAsiento, stBorraSQL);
			this.lblAnulado.Visible = false;
			this.drCompra["Borrar"] = false;
			this.cdsCompra.Update();
		}


		private void ImprimeNota(int iTipo)
		{
			if (!miAcceso.Imprimir)
			{
				MessageBox.Show("No tiene acceso a Imprimir Nota de Entrega",
					"Seguridad", MessageBoxButtons.OK, MessageBoxIcon.Information);
				return;
			}

			string stFiltro = "{Compra.idCompra} = " + this.txtIdCompra.Value.ToString();
			string stReporte = "NotaEntrega.rpt";
			if (IdTipoFactura == 4) 
				stReporte = "IngresoBodega.rpt";
			Reporte miReporte = new Reporte(stReporte, stFiltro);
			if (iTipo == 0)
			{
				miReporte.MdiParent = this.MdiParent;
				miReporte.Show();
			}
			else
				miReporte.ImprimeRep();

			this.Cursor = Cursors.Default;
		}

		private void btNotaEntrega_Click(object sender, System.EventArgs e)
		{
			if (barraDato1.bEditar) return;

			this.Cursor = Cursors.WaitCursor;
			this.ImprimeNota(0);
			this.Cursor = Cursors.Default;
		}

		private void btGuiaRemision_Click(object sender, System.EventArgs e)
		{
			if (barraDato1.bEditar) return;
			CompraGuia miGuia = new CompraGuia((int) this.txtIdCompra.Value);

			miGuia.dtFecha = (DateTime) this.dtFecha.Value;
			miGuia.IdCliente = (int) this.cmbCliente.Value;
			miGuia.MdiParent = this.MdiParent;
			miGuia.Show();
		}

		private void ActualizaCostoUltimo()
		{
			if (IdTipoFactura != 4) return;
			int  iIdArt = 0;
			double dPrecio = 0;
			string stSelect = "";
			DateTime dtFechaC = (DateTime) this.dtFecha.Value;
			cdsArticulo.Schema.Connections[0].Open();
			foreach(C1.Data.C1DataRow drDet in drCompra.GetChildRows())
			{
				iIdArt = (int) drDet["idArticulo"];
				dPrecio = (double) drDet["Precio"];
				stSelect = "UPDATE Articulo SET ";
				stSelect += "idCompra3 = idCompra2";
				stSelect += ", idCompra2 = idCompra1";
				stSelect += ", idCompra1 = " + this.txtIdCompra.Value.ToString();
				stSelect += ", CostoUltimo = " + dPrecio.ToString("0.00", us);
				stSelect += ", FechaUltimo = '" + dtFechaC.ToString("yyyyMMdd") +"'";
				stSelect += " WHERE idArticulo = " + iIdArt.ToString();
				stSelect += " And ISNULL(FechaUltimo, '19000101') <= '" + dtFechaC.ToString("yyyyMMdd") +"'";

				C1.Data.SchemaObjects.Connection miConeccion = this.cdsArticulo.Schema.Connections[0];
				SqlCommand cmdActualiza = new SqlCommand (
					stSelect, (SqlConnection) miConeccion.DbConnection);

				cmdActualiza.Transaction = (SqlTransaction)
					miConeccion.DbTransaction;
				try
				{
					cmdActualiza.ExecuteNonQuery();
				}
				catch(System.Exception ex)
				{
					MessageBox.Show(ex.Message, "ActualizaCostoUltimo");
				}
			}
			cdsCompra.Schema.Connections[0].Close();
		}

		private void ActualizaPrecio()
		{
			if (IdTipoFactura != 4) return;
			if (barraDato1.bEditar) return;
			if (this.btActualizaPrecio.Visible)
			{
				if (DialogResult.No == MessageBox.Show(
					"Desea Actualizar precios en base a costo de Compra + % de Incremento?",
					"Confirmacion", MessageBoxButtons.YesNo, MessageBoxIcon.Question))
					return;
			}
			string stExec = "Exec ArticuloActualizaPrecio " + txtIdCompra.Value.ToString();
			Funcion.EjecutaSQL(cdsArticulo, stExec, true);

			if (this.btActualizaPrecio.Visible)
			{
				MessageBox.Show("Precios Actualizados",
					"Informacion", MessageBoxButtons.OK, MessageBoxIcon.Information);
			}
		}

		CultureInfo us = new CultureInfo("en-US");
		private void btActualizaPrecio_Click(object sender, System.EventArgs e)
		{
			if (barraDato1.bEditar) return;
			if (IdTipoFactura != 4)
			{
				MessageBox.Show("Actualización de precios solo en compras");
				return;
			}
			if (!miAcceso.ActualizaPrecio)
			{
				MessageBox.Show("No tiene acceso a Actualización de precios");
				return;
			}
			if (DialogResult.No == MessageBox.Show(
				"¿Está seguro que desea Actualizar los precios de venta en base a esta compra?", 
				"Confirmación", MessageBoxButtons.YesNo, MessageBoxIcon.Warning, 
				MessageBoxDefaultButton.Button2)) return;
			ActualizaCostoUltimo();
//			if ((bool) drSeteoF["ActualizarCosto"])
      ActualizaPrecio();
		}

		private void barraDato1_Copiar(object sender, System.EventArgs e)
		{
			if (MessageBox.Show("¿Desea copiar este registro?",
				"Confirmación", MessageBoxButtons.YesNo, 
				MessageBoxIcon.Question) == DialogResult.No) return;

			CopiarFactura((int) this.IdTipoFactura);
			string stOrden = barraDato1.DataOrden.ToString();
			barraDato1.DataOrden = "";
			barraDato1.PosUltima();
			barraDato1.DataOrden = stOrden;
			MessageBox.Show("Copia de Factura Terminada",
				"Información", MessageBoxButtons.OK, 
				MessageBoxIcon.Information);
		}

		private void ultraGrid1_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (ultraGrid1.ActiveCell == null) return;
			try
			{
				if (ultraGrid1.ActiveCell.Column.Key.ToString().ToUpper() == "NOTAS") return;
				switch(e.KeyValue)
				{
					case (int) Keys.F3:
						if (ultraGrid1.ActiveCell.Column.Key.ToString().ToUpper() == "REFNUMERO")
						{
							int IdDetCompra = (int) ultraGrid1.ActiveRow.Cells["idDetCompra"].Value;
							if (barraDato1.bNuevo || IdDetCompra < 0)
							{
								MessageBox.Show("Antes de ingresar la serie debe grabar el registro",
									"Informacion", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
								return;
							}
							int IdArt = (int) ultraGrid1.ActiveRow.Cells["idArticulo"].Value;
							if (IdArt < 1)
							{
								MessageBox.Show("Ingrese Articulo",
									"Informacion", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
								return;
							}
							int IdCompra = (int) txtIdCompra.Value;
							double dCantidad = (double) ultraGrid1.ActiveRow.Cells["Cantidad"].Value;
							int iCantidad = int.Parse(dCantidad.ToString("0", us));
							int iDisponible = 0;
							// Compra o Ingreso de Bodega
							if (IdTipoFactura == 4 || IdTipoFactura == 9)
							{
								if (!e.Shift)
								{
									string stExec = "Exec SerieCrear " + IdArt.ToString() + ", "
										+ IdCompra.ToString() + ", " + iCantidad.ToString()
										+ ", 1, " + IdDetCompra.ToString();

									Funcion.EjecutaSQL(cdsArticulo, stExec, true);
								}
							}
							// Venta o Egreso de Bodega
							if (IdTipoFactura == 1 || IdTipoFactura == 8)
							{
								if (!e.Shift)	iDisponible = iCantidad;
							}
							// Devolucion de Venta
							int idTipoFactura = 0;
							if (IdTipoFactura == 5)
							{
								if (!e.Shift)	iDisponible = -iCantidad;
								idTipoFactura = 1;
							}
							// Devolucion de Compra
							if (IdTipoFactura == 6)
							{
								if (!e.Shift)	iDisponible = iCantidad;
								idTipoFactura = 4;
							}
							bool bUsado = e.Shift;
							CompraSerie miSerie = new CompraSerie(IdArt, IdCompra, iDisponible, 
								idTipoFactura, IdDetCompra, bUsado);
							miSerie.ShowDialog();
						}
						if (ultraGrid1.ActiveCell.Column.Key.ToString().ToUpper() == "PRECIO")
						{
							int IdArt = (int) ultraGrid1.ActiveRow.Cells["idArticulo"].Value;
							if (IdArt < 1)
							{
								MessageBox.Show("Ingrese Articulo",
									"Informacion", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
								return;
							}
							CompraArt miCompraArt = new CompraArt(IdArt);
							miCompraArt.ShowDialog();
						}
						break;

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
				//Console.WriteLine(ex.Message);
				MessageBox.Show(ex.Message, "Tecla");
			}
		}

		private void btCliente_MouseDown(object sender, System.Windows.Forms.MouseEventArgs e)
		{
			if (e.Button != MouseButtons.Right) return;
			// Tipo de Factura genera CxC o CxP
			string stSelect = "SELECT CobroPago FROM CompraNumero WHERE idTipoFactura = "
				+ this.cmbTipoCompra.Value.ToString().Trim();
			int iCobroPago = Funcion.iEscalarSQL(cdsCompra, stSelect, true);

			string stTipoSaldo = "";

			// Saldo de Cuentas por Cobrar o Pagar de Clientes
			double dSaldoCxC = 0;
			if (iCobroPago == 1 || iCobroPago == -1 || iCobroPago == 0)
			{
				dSaldoCxC = SelectSaldo();

				if (dSaldoCxC > 0 || iCobroPago != 0)
					stTipoSaldo = "Saldo por Cobrar: " + dSaldoCxC.ToString("#,##.00", us) + "\n";
			}
			double dSaldoCxP = 0;
			if (iCobroPago == 2 || iCobroPago == -2 )
			{
				dSaldoCxP = SelectSaldo();
				if (dSaldoCxP > 0 || iCobroPago != 0)
					stTipoSaldo += "Saldo por Pagar: " + dSaldoCxP.ToString("#,##.00", us) + "\n";
			}

			double dSaldoCompra = 0;

			if (iCobroPago != 0) 
			{
				stSelect = "SELECT ISNULL(Saldo, 0) FROM Compra WHERE idCompra = "
					+ this.txtIdCompra.Value.ToString().Trim();

				dSaldoCompra = Funcion.dEscalarSQL(cdsCompra, stSelect, true);
				stTipoSaldo += "Saldo de Factura: " + dSaldoCompra.ToString("#,##.00", us);
			}
			double dVencido = SaldoMora();
			if (dVencido > 0) 
			{
				stTipoSaldo += "\nVencido: " + dVencido.ToString("0.00", us);
			}
			if (stTipoSaldo.Length == 0) stTipoSaldo = "No existen Saldos";
			MessageBox.Show(stTipoSaldo, "Saldos");	
		}

		private void btGuiaRemision_MouseDown(object sender, System.Windows.Forms.MouseEventArgs e)
		{
			if (e.Button != MouseButtons.Right) return;
			Cursor = Cursors.WaitCursor;
			string stFiltro = "{Compra.idCompra} = " + 
				this.txtIdCompra.Value.ToString();
			using (Reporte miReporte = new Reporte("GuiaRemision.Rpt", stFiltro))
			{
				miReporte.ShowDialog();
			}
			Cursor = Cursors.Default;
		}

		private void cdsPersonal_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsPersonal.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirRoles);		
		}

		private void btTotalCuotas_Click(object sender, System.EventArgs e)
		{
			this.errorProvider.SetError(this.txtNumCuotas, "");
			if (this.txtNumCuotas.Value.ToString().Length == 0)
			{
				this.errorProvider.SetError(this.txtNumCuotas, "Ingrese Numero de Cuotas");
				return;
			}
			int iPlazo = int.Parse(this.txtNumCuotas.Value.ToString());
			if (iPlazo == 0)
			{
				this.txtValorCuotas.Value = 0;
				this.txtTotalCuotas.Value = 0;
				this.txtInteresPagar.Value = 0;
				return;
			}

			double dMonto = (double) this.txtTotal.Value - (double) this.txtAnticipo.Value
				- (double) drCompra["Retenido"];
			double dInteres = (double) this.txtPorcInteres.Value / 12 / 100;
			double dCuota = 0;
			if ((bool) this.drSeteoF["InteresLineal"] == true)
				dCuota = (dMonto * (1 + dInteres * iPlazo))/iPlazo;
			else
			{
				if (dInteres == 0)
					dCuota = dMonto / iPlazo;
				else
					dCuota = dMonto * dInteres / (1-Math.Pow(1 + dInteres, -iPlazo));
			}
			//Caso de interes no calculado
			if ((double) this.txtPorcInteres.Value == 0 && iPlazo == 1)
				dCuota = dMonto + (double) this.txtInteresPagar.Value;

			this.drCompra["ValorCuotas"] = Math.Round(dCuota, 2);
			this.drCompra["TotalCuotas"] = Math.Round(dCuota, 2) * iPlazo + 
				(double) this.txtAnticipo.Value;
			if ((double) this.txtPorcInteres.Value ==0)
			{
				this.drCompra["TotalCuotas"] = (double) this.txtTotal.Value;
				this.drCompra["ValorInteres"] = 0;
			}
			else
			{
				if ((double) this.txtTotalCuotas.Value < (double) this.txtTotal.Value)
					this.drCompra["TotalCuotas"] = (double) this.txtTotal.Value;
				this.drCompra["ValorInteres"] = (double) this.txtTotalCuotas.Value
					- (double) this.txtTotal.Value;
			}
		}

		private void cmdVuelto_Click(object sender, System.EventArgs e)
		{
			float fTotal, fVuelto, fRecibido;
			if (spnVuelto.Value == System.DBNull.Value) return;
			fRecibido = float.Parse(this.spnVuelto.Value.ToString());
			fTotal = float.Parse(this.txtTotal.Value.ToString());
			if (fRecibido < fTotal)
			{
				MessageBox.Show("Valor Recibido no Puede ser menor que total a cobrar",
					"Error en Valor Recibido", MessageBoxButtons.OK, 
					MessageBoxIcon.Error);
				return;
			}
			fVuelto = fRecibido - fTotal;
			MessageBox.Show("Su Cambio es de: " + fVuelto.ToString("#,##0.00"),
				"Informacion", MessageBoxButtons.OK, MessageBoxIcon.Information);
		}

		private void barraDato1_Refresca(object sender, System.EventArgs e)
		{
			try
			{
				drCompra = cdsCompra.TableViews["Compra"].Rows[0];
				if ((bool)drCompra["Borrar"] == true) this.lblAnulado.Visible = true;
				else this.lblAnulado.Visible = false;
			}
			catch{};
		}

		private void ultraGrid1_BeforeRowInsert(object sender, Infragistics.Win.UltraWinGrid.BeforeRowInsertEventArgs e)
		{
			if (this.ultraGrid1.Rows.Count >= (int) this.drSeteoF["NumeroLineas"] && (IdTipoFactura == 1 || IdTipoFactura == 3)) 
			{
				e.Cancel = true;
				MessageBox.Show("No puede crear mas lineas");
			}
		}

		private void btFlete_Click(object sender, System.EventArgs e)
		{
			if (DialogResult.No == MessageBox.Show("Desea calcular el total del flete?", 
				"Confirmacion",	MessageBoxButtons.YesNo, 
				MessageBoxIcon.Information)) return;
			double dFlete = 0;
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.ultraGrid1.Rows.All)
			{
				if (dr.Cells["Flete"].Value != System.DBNull.Value)
					dFlete += (double) dr.Cells["Flete"].Value;
			}
			drCompra["Flete"] = dFlete;
			MessageBox.Show("Flete Calculado", "Informacion",
				MessageBoxButtons.OK, MessageBoxIcon.Information);
 
//			CompraFlete miFlete = new CompraFlete();
//			miFlete.MdiParent = this.MdiParent;
//			miFlete.Show();
		}

		private void ultraGrid1_Error(object sender, Infragistics.Win.UltraWinGrid.ErrorEventArgs e)
		{
			if (e.MaskErrorInfo == null)
			{
				e.Cancel = true; 
				MessageBox.Show(e.DataErrorInfo.ErrorText, "Informacion de Error");
				//				e.DataErrorInfo.Cell.Value = 1;
			}
		}

		private void btNuevo_Click(object sender, System.EventArgs e)
		{
			if (!barraDato1.bEditar) return;
			this.ultraGrid1.Rows.Band.AddNew();
			if (!(bool) drSeteoF["VerDesplegable"])
			{
				ultraGrid1.ActiveCell = ultraGrid1.ActiveRow.Cells["Codigo"];
				ultraGrid1.PerformAction(UltraGridAction.EnterEditMode, false, false);
			}
			else
			{
				ultraGrid1.ActiveCell = ultraGrid1.ActiveRow.Cells["CodArt"];
				ultraGrid1.PerformAction(UltraGridAction.EnterEditMode, false, false);
			}
		}

		private void btCambiarPrecio_Click(object sender, System.EventArgs e)
		{
			if (!this.barraDato1.bEditar) return;
			if (miAcceso.CambiarPrecio == false) return;
			if (this.cmbPrecio.Value == null) return;
			string stPrecio = "Precio" + this.cmbPrecio.Value.ToString().Trim();

			if (ultraGrid1.ActiveRow.Cells["idArticulo"].Value == System.DBNull.Value) return;
			int IdArt = (int) ultraGrid1.ActiveRow.Cells["idArticulo"].Value;
			if (IdArt < 1) return;
			string stSelectPrecio = "Select "  +stPrecio + " From Articulo Where idArticulo = "
				+ IdArt.ToString();
			ultraGrid1.ActiveRow.Cells["Precio"].Value = Funcion.dEscalarSQL(cdsArticulo, stSelectPrecio, true);
			ultraGrid1.ActiveRow.Cells["Precio"].Activate();
			ultraGrid1.ActiveRow.Cells["Precio"].Refresh();
	
			cmdTotal();
			MessageBox.Show("Precio Modificado",
				"Informacion", MessageBoxButtons.OK,
				MessageBoxIcon.Information);
		}

		private void cdsClienteTabla_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			this.cdsCliente.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);		
		}

		private void cdsSeteo_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteo.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);		
		}

		private void btFlete_MouseDown(object sender, System.Windows.Forms.MouseEventArgs e)
		{
			if (e.Button != MouseButtons.Right) return;
			if (!barraDato1.bEditar) return;
//			if (!(IdTipoFactura == 4 || IdTipoFactura == 8 || IdTipoFactura == 3 || IdTipoFactura == 9)) 
//				return; // Aplica solo a compras, Egreso, Cotizacion
			if (this.txtFlete.Value == System.DBNull.Value) return;
			if ((double) txtFlete.Value == 0) return;
			cmdTotal();
			double dFlete = double.Parse(txtFlete.Value.ToString());
			double dSubTotalIva = 0;
			double dSubTotalIva0 = 0;
			if (txtIva.Value != System.DBNull.Value) dSubTotalIva = double.Parse(txtIva.Value.ToString());
			if (txtIva0.Value != System.DBNull.Value) 
				dSubTotalIva0 = double.Parse(txtIva0.Value.ToString()) - dFlete;
			if (dSubTotalIva + dSubTotalIva0 == 0) return;
			if (DialogResult.No == MessageBox.Show("Desea Distribuir flete proporcionalmente?", 
				"Confirmacion",	MessageBoxButtons.YesNo, 
				MessageBoxIcon.Information)) return;
			double dPorc = dFlete / (dSubTotalIva + dSubTotalIva0);
			double dCantidad, dPrecio;
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.ultraGrid1.Rows.All)
			{
				dCantidad = 0;
				dPrecio = 0;
				if (dr.Cells["Cantidad"].Value != System.DBNull.Value)
					dCantidad = (double) dr.Cells["Cantidad"].Value;
				if (dr.Cells["Precio"].Value != System.DBNull.Value)
					dPrecio = (double) dr.Cells["Precio"].Value;
				dr.Cells["Flete"].Activate();
				dr.Cells["Flete"].Value = dCantidad * dPrecio * dPorc;
			}
//			drCompra["Flete"] = 0;
			MessageBox.Show("Flete Repartido", "Informacion",
				MessageBoxButtons.OK, MessageBoxIcon.Information);
		}

		private void cmbCliente_MouseUp(object sender, System.Windows.Forms.MouseEventArgs e)
		{
			if (e.Button == MouseButtons.Right)
			{
				string stPrimero = "Nombre";
				if (cmbCliente.DisplayLayout.Bands[0].Columns[stPrimero].Header.VisiblePosition == 0)
					stPrimero = "Codigo";
				cdvCliente.Sort = stPrimero;
				cmbCliente.DisplayLayout.Bands[0].Columns[stPrimero].Header.VisiblePosition = 0;
				cmbCliente.DisplayMember = stPrimero;
			}
		}

		private void ultraGrid1_BeforeRowsDeleted(object sender, Infragistics.Win.UltraWinGrid.BeforeRowsDeletedEventArgs e)
		{
			if (DialogResult.No ==	MessageBox.Show("Desea Borrar Lineas Seleccionadas?",
				"Confirmacion", MessageBoxButtons.YesNo, MessageBoxIcon.Question))
			{
				e.Cancel = true;
			}
			else
			{
				for (int i=0; i<e.Rows.Length; i++)
				{
					string stIdDetCompra = e.Rows[i].Cells["idDetCompra"].Value.ToString();
					string stExec = "Exec DetCompraBorrarL " + stIdDetCompra;
					string stMensaje = Funcion.sEscalarSQL(cdsArticulo, stExec, true);
					if (stMensaje.Length > 0)
					{
						MessageBox.Show(stMensaje, "Informacion");
						e.Cancel = true;
						return;
					}
				}

				e.DisplayPromptMsg = false;
			}
		}

		private void btBuscaArticulo_MouseUp(object sender, System.Windows.Forms.MouseEventArgs e)
		{
			if (e.Button == MouseButtons.Right)
			{
				Cursor = Cursors.WaitCursor;
				string stFiltro = "";
				string stTitulo = "";
				if (this.ultraGrid1.ActiveRow.Cells["idArticulo"].Value != null
					&& this.ultraGrid1.ActiveRow.Cells["idArticulo"].Value != System.DBNull.Value)
				{
					stFiltro = "{Articulo.idArticulo} = " 
						+ ultraGrid1.ActiveRow.Cells["idArticulo"].Value.ToString();
				}
				if (drCompra["idCliente"] != System.DBNull.Value)
				{
					if (stFiltro.Length > 0) stFiltro += " And ";
					stFiltro += "{Cliente.idCliente} = " 
						+ drCompra["idCliente"].ToString();
					if (txtNombre.Width > 0)
						stTitulo = "Cliente " + txtNombre.Text;
					else
            stTitulo = "Cliente " + cmbCliente.ActiveRow.Cells["Nombre"].Value.ToString();
				}
				Reporte miReporte = new Reporte("Articulo por Fecha.rpt", stFiltro);
				miReporte.Titulo(stTitulo);
				miReporte.Show();
				Cursor = Cursors.Default;
			}
		}

		private void btPago_MouseUp(object sender, System.Windows.Forms.MouseEventArgs e)
		{
			if (e.Button != MouseButtons.Right) return;
				
			if (barraDato1.bEditar)
			{
				MessageBox.Show("Para generar un pago debe grabar previamente el registro", 
					"Informacion", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
				return;
			}

			string stExec = "Exec AbonoCruzeVerifica " + txtIdCompra.Value.ToString();
			string stMensaje = Funcion.sEscalarSQL(cdsArticulo, stExec, true);
			if (stMensaje.Length > 0)
			{
				MessageBox.Show(stMensaje, 
					"Informacion", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
				return;
			}

			if (DialogResult.No == MessageBox.Show("Desea Generar un abono?",
				"Confirmacion", MessageBoxButtons.YesNo, MessageBoxIcon.Question))
				return;

			stExec = "Exec AbonoCruzePago " + txtIdCompra.Value.ToString();
			Funcion.sEscalarSQL(cdsArticulo, stExec, true);

			MessageBox.Show("Abono Generado", "Informacion",
				MessageBoxButtons.OK, MessageBoxIcon.Information);	
		}

		private void btTotalCuotas_MouseDown(object sender, System.Windows.Forms.MouseEventArgs e)
		{
			if (e.Button != MouseButtons.Right) return;

			int IdCompra = 0;
			try
			{
				IdCompra = (int) txtIdCompra.Value;
			}
			catch
			{
				MessageBox.Show("Debe grabar la factura antes de generar pagos", "informacion");
				return;
			}

			string stVerifica = "SELECT COUNT(*) From Pago Where IdCompra = "
				+ IdCompra.ToString();
			if (Funcion.iEscalarSQL(cdsCompra, stVerifica) > 0)
			{
				MessageBox.Show("No puede generar tabla de amortizacion cuando existen pagos");
				return;
			}
			
			this.errorProvider.SetError(this.txtNumCuotas, "");
			if (this.txtNumCuotas.Value.ToString().Length == 0)
			{
				this.errorProvider.SetError(this.txtNumCuotas, "Ingrese Numero de Cuotas");
				return;
			}
			if (IdTipoFactura != 1 && IdTipoFactura != 4 && IdTipoFactura != 11) return;
			if ((int) this.rdgContadoCred.Value != 2) 
			{
				MessageBox.Show("Pagos solo se puede realizar en Creditos",
					"Informacion", MessageBoxButtons.OK, MessageBoxIcon.Information);
				return;
			}
			int iPlazo = int.Parse(this.txtNumCuotas.Value.ToString());
			if (iPlazo == 0)
			{
				return;
			}
			if (drSeteoF["PagoInteres"] != DBNull.Value && (bool) drSeteoF["PagoInteres"] == true)
			{
				string stExec = "Exec AmortizacionTabla " + txtIdCompra.Value.ToString();
				Funcion.EjecutaSQL(cdsArticulo, stExec, true);
				MessageBox.Show("Tabla Amortizacion Generada", "Informacion");
				return;
			}			

			string stInsert = "Insert Into Pago (Fecha, Interes, Pago, Pendiente, idCompra, Notas) "
				+ "	VALUES (";
			DateTime dtFechaIni = (DateTime) dtFecha.Value;
			double dInteres = (double) txtPorcInteres.Value / 12.0/ 100.0;
			double dCapital = (double) txtTotal.Value - (double) txtAnticipo.Value 
				- (double) drCompra["Retenido"];
			double dCuota = (double) txtValorCuotas.Value;
			for (int i=0 ; i<iPlazo; i++)
			{
				double dInt = dCapital* dInteres;
				double dPagado = dCuota - dInt;
				if (i == iPlazo - 1)
				{
					dPagado = dCapital;
				}
				dtFechaIni = dtFechaIni.AddMonths(1);
				string stValor = stInsert + "'" + dtFechaIni.ToString("yyyyMMdd") + "'";
        string stNota = "Cuota " + (i + 1).ToString() + "/" + iPlazo.ToString();
        stValor += ", " + dInt.ToString("0.00", us) + ", " + dPagado.ToString("0.00", us) + ", 1, "
          + IdCompra.ToString() + ",'" + stNota + "')";
				dCapital -= dPagado;
				Funcion.EjecutaSQL(cdsCompra, stValor);
			}
			MessageBox.Show("Tabla de Amortizacion Generada", "Informacion");
		}

		private void cmbCliente_Validated(object sender, System.EventArgs e)
		{
			if (!barraDato1.bEditar) return;
			if (drCompra["idCliente"] != null && (int) drCompra["idCliente"] > 0)
				miProveedor = new Proveedor(cdsCliente, (int) drCompra["idCliente"]);
			#region Nuevo Registro
			if (barraDato1.bNuevo) 
			{
				if ((IdTipoFactura == 4 || IdTipoFactura == 2) && drCompra["idCliente"] != null && (int) drCompra["idCliente"] > 0)
				{
					drCompra["AutImprenta"] = miProveedor.AutorizaImprenta;
					drCompra["SerieFactura"] = miProveedor.SerieFactura;
					drCompra["AutFactura"] = miProveedor.AutorizaFactura;
					if (miProveedor.FechaCaducidad > DateTime.Today.AddYears(-10))
						drCompra["FechaCaducidad"] = miProveedor.FechaCaducidad;
					if (miProveedor.idCredTribut > 0)
						drCompra["idCredTributario"] = miProveedor.idCredTribut;
				}
				if ((bool) drSeteoF["Copia_Dist"] == false) return;
				if ((int) rdgContadoCred.Value == 2)
				{
					drCompra["DescPorcIva"] = miProveedor.DescuentoCredito;
					drCompra["DescPorcIva0"] = miProveedor.DescuentoCredito;
				}
				else
				{
					drCompra["DescPorcIva"] = miProveedor.DescuentoContado;
					drCompra["DescPorcIva0"] = miProveedor.DescuentoContado;
				}

				drCompra["idVendedor"] = miProveedor.idVendedor;
				drCompra["DiasCredito"] = miProveedor.DiasCredito;
			}
			#endregion Nuevo Registro
		}

		private void rdgContadoCred_Validated(object sender, System.EventArgs e)
		{
			if (!barraDato1.bNuevo) return;
			if ((bool) drSeteoF["Copia_Dist"] == false) return;
			if ((int) rdgContadoCred.Value == 2)
			{
				drCompra["DescPorcIva"] = miProveedor.DescuentoCredito;
				drCompra["DescPorcIva0"] = miProveedor.DescuentoCredito;
			}
			else
			{
				drCompra["DescPorcIva"] = miProveedor.DescuentoContado;
				drCompra["DescPorcIva0"] = miProveedor.DescuentoContado;
			}
		}

		private void btRetencion_MouseDown(object sender, System.Windows.Forms.MouseEventArgs e)
		{
			if (barraDato1.bEditar)
			{
				MessageBox.Show("No puede crear una retencion mientras esta en edicion"
					, "Confirmación", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
				return;
			}
			if (e.Button != MouseButtons.Right) return;
			if ((double) drCompra["Retenido"] < .001)
			{
				MessageBox.Show("No existe retencion", "Informacion");
				return;
			}
			// Crea asiento de Retencion
			if (!miAcceso.CrearAsiento)
			{
				MessageBox.Show("No tiene Acceso a Crear Asientos", "Seguridad", 
					MessageBoxButtons.OK, MessageBoxIcon.Information);
				return;
			}

			if (DialogResult.No == MessageBox.Show("¿Desea crear Asiento de Retencion?", "Confirmación", 
				MessageBoxButtons.YesNo, MessageBoxIcon.Question))
				return;

//			bool bRepetido = false;
			int IdAsiRet = 0;
			if (drCompra["idAsiRet"] != System.DBNull.Value && (int) drCompra["idAsiRet"] > 0)
			{
				IdAsiRet = (int) drCompra["idAsiRet"];
				string stSelect = "SELECT COUNT(*) FROM Asiento WHERE Borrar = 0 And idAsiento = "+
					IdAsiRet.ToString().Trim();
				int iCuenta = Funcion.iEscalarSQL(cdsAsiento, stSelect);

				if (iCuenta > 0)
				{
					if (DialogResult.No == MessageBox.Show("Asiento de Retencion ya existe desea sobreescribirlo?", 
						"Confirmación", MessageBoxButtons.YesNo, MessageBoxIcon.Asterisk))
					{
						Cursor = Cursors.Default;
						return;
					}
					// Si desea sobrrescribir
					string stBorrar = "DELETE FROM AsientoDetalle WHERE idAsiento = "+
						IdAsiRet.ToString().Trim();
//					bRepetido = true;
					Funcion.EjecutaSQL(cdsAsiento, stBorrar);
				}
			}

			Cursor = Cursors.WaitCursor;
			int IdCompra = int.Parse(txtIdCompra.Value.ToString());
//			if (!bRepetido)
//			{
//				Contabilidad.AsientoRetencion(cdsAsiento, IdCompra, IdTipoFactura);
//				// Grabacion de Asiento de Retencion
//				string stIdAsi = "Select IdAsiento FROM Asiento Order by IdAsiento DESC ";
//				IdAsiRet = Funcion.iEscalarSQL(cdsAsiento, stIdAsi, true);
//				drCompra["idAsiRet"] = IdAsiRet;
//				string stSelect = "Update Compra SET idAsiRet = " + IdAsiRet.ToString()
//					+ " Where idCompra = " + IdCompra.ToString();
//				Funcion.EjecutaSQL(cdsCompraTabla, stSelect, false);
//				// Grabacion de Codigo de Asiento de Retencion
//				stIdAsi = "Select CodAsiento FROM Asiento Where IdAsiento = " 
//					+ IdAsiRet.ToString();
//				stIdAsi = Funcion.sEscalarSQL(cdsAsiento, stIdAsi, true);
//				if (drCompra["Notas"] != System.DBNull.Value)
//					stIdAsi = drCompra["Notas"] + " " + stIdAsi;
//				stSelect = "Update Compra SET Notas = Notas + ' ' + '" + stIdAsi
//					+ "' Where idCompra = " + IdCompra.ToString();
//				Funcion.EjecutaSQL(cdsCompraTabla, stSelect, false);
//			}
			string stExec = "EXEC AsientoRetencionC " + IdAsiRet.ToString() + ", -"
				+ IdCompra.ToString() + ", " + IdTipoFactura.ToString();
			Funcion.EjecutaSQL(cdsAsiento, stExec, true);
			MessageBox.Show("Asiento de Retencion Generado", "Informacion");
			Cursor = Cursors.Default;
		}

		private void btIngresar_Click(object sender, System.EventArgs e)
		{
			if (!barraDato1.bEditar) return;
			if (ugrBusca.ActiveRow == null) return;
			int idArt = (int) ugrBusca.ActiveRow.Cells["idArticulo"].Value;

			ultraGrid1.Rows.Band.AddNew();
			ultraGrid1.ActiveCell = ultraGrid1.ActiveRow.Cells["idArticulo"];
			ultraGrid1.ActiveCell.Value = idArt;
			if (drSeteoF["VerDesplegable"] != System.DBNull.Value && !(bool) drSeteoF["VerDesplegable"])
			{
				string Codigo = ugrBusca.ActiveRow.Cells["Codigo"].Value.ToString();
				string Articulo = ugrBusca.ActiveRow.Cells["Nombre"].Value.ToString();
				ultraGrid1.ActiveCell = ultraGrid1.ActiveRow.Cells["Codigo"];
				ultraGrid1.ActiveCell.Value = Codigo;
				ultraGrid1.ActiveCell = ultraGrid1.ActiveRow.Cells["Articulo"];
				ultraGrid1.ActiveCell.Value = Articulo;
			}
			else
			{
				ultraGrid1.ActiveCell = ultraGrid1.ActiveRow.Cells["CodArt"];
				ultraGrid1.ActiveCell.Value = idArt;
			}

			ultraGrid1.PerformAction(UltraGridAction.EnterEditMode, false, false);
		}

		private void btBuscarArtciculos(object sender, System.EventArgs e)
		{
			string stBusca = txtBusca.Text;
			string stExec = string.Format("Exec VentaArticuloLista '{0}'", stBusca);
			ugrBusca.DataSource = Funcion.dvProcedimiento(cdsSeteo, stExec);
			return;
			string stSelect = "Select Top 50 Codigo, Articulo, CodAlterno, idArticulo, " +
				" Precio1, Precio2, Precio3, Precio4, Precio5, Precio6, Ubicacion FROM " +
				"Articulo Where Codigo like '%" + stBusca + "%'" +
				" Or CodAlterno like '%" + stBusca + "%'" +
				" Or Articulo like '%" + stBusca + "%'";
			SqlDataReader miReader=	Funcion.rEscalarSQL(cdsArticulo, stSelect, true);
			udsBusca.Rows.Clear();
			int iNumFilas = 0;
			double Precio1, Precio2, Precio3, Precio4, Precio5, Precio6;
			while (miReader.Read())
			{
				Precio1 = 0; Precio2=0; Precio3=0; Precio4=0; Precio5=0; Precio6=0;
				string stCodigo = "";
				string stCodAlterno = "";
				string stArticulo = "";
				string stUbicacion = "";
				if (miReader.GetValue(0) != System.DBNull.Value)
					stCodigo = miReader.GetString(0);
				if (miReader.GetValue(1) != System.DBNull.Value)
					stArticulo = miReader.GetString(1);
				if (miReader.GetValue(2) != System.DBNull.Value)
					stCodAlterno = miReader.GetString(2);
				int IdArt = miReader.GetInt32(3);
				if (miReader.GetValue(4) != System.DBNull.Value)
					Precio1 = miReader.GetDouble(4);
				if (miReader.GetValue(5) != System.DBNull.Value)
					Precio2 = miReader.GetDouble(5);
				if (miReader.GetValue(6) != System.DBNull.Value)
					Precio3 = miReader.GetDouble(6);
				if (miReader.GetValue(7) != System.DBNull.Value)
					Precio4 = miReader.GetDouble(7);
				if (miReader.GetValue(8) != System.DBNull.Value)
					Precio5 = miReader.GetDouble(8);
				if (miReader.GetValue(9) != System.DBNull.Value)
					Precio6 = miReader.GetDouble(9);
				if (miReader.GetValue(10) != System.DBNull.Value)
					stUbicacion = miReader.GetString(10);
				udsBusca.Rows.Add();
				udsBusca.Rows[iNumFilas]["Codigo"]=stCodigo;
				udsBusca.Rows[iNumFilas]["Nombre"]=stArticulo;
				udsBusca.Rows[iNumFilas]["Alterno"]=stCodAlterno;
				udsBusca.Rows[iNumFilas]["idArticulo"]=IdArt;
				udsBusca.Rows[iNumFilas]["Precio1"]=Precio1;
				udsBusca.Rows[iNumFilas]["Precio2"]=Precio2;
				udsBusca.Rows[iNumFilas]["Precio3"]=Precio3;
				udsBusca.Rows[iNumFilas]["Precio4"]=Precio4;
				udsBusca.Rows[iNumFilas]["Precio5"]=Precio5;
				udsBusca.Rows[iNumFilas]["Precio6"]=Precio6;
				udsBusca.Rows[iNumFilas]["Ubicacion"]=stUbicacion;
				iNumFilas ++;
			}		
			miReader.Close();
		}

		private void btSaldo1_MouseUp(object sender, System.Windows.Forms.MouseEventArgs e)
		{
			if (e.Button != MouseButtons.Right) return;
			// Muestra la forma de Entrega de un pedido
			if (IdTipoFactura == 13)
			{
				string stFiltro = "{DetCompra.RefNumero} = '" + txtNumero.Text + "'";
				Reporte miRep = new Reporte("PedidoEntrega", stFiltro);
				miRep.MdiParent = this.MdiParent;
				miRep.Show();
			}

		}

		private void chkAutPrincipal_Click(object sender, System.EventArgs e)
		{
			if (!barraDato1.bEditar) return;
			if (chkAutPrincipal.Checked) return;
			using (CompraAutoriza miAutoriza = new CompraAutoriza())
			{
				if (DialogResult.OK == miAutoriza.ShowDialog())
				{
					drCompra["AutPrincipal"] = true;
				}
				else
				{
					drCompra["AutPrincipal"] = false;
					chkAutPrincipal.Checked = false;
				}
			}
		}

		private void txtBusca_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (e.KeyChar != 13) return;
			string stBusca = txtBusca.Text.Trim();
			string stSelect = "Select Count(*) FROM " +
				"Articulo Where Codigo like '%" + stBusca + "%'" +
				" Or CodAlterno like '%" + stBusca + "%'" +
				" Or Articulo like '%" + stBusca + "%'";
			int iCuenta = Funcion.iEscalarSQL(cdsArticulo, stSelect, true);
			if (iCuenta == 1)
			{
				stSelect = "Select idArticulo FROM " +
					"Articulo Where Codigo like '%" + stBusca + "%'" +
					" Or CodAlterno like '%" + stBusca + "%'" +
					" Or Articulo like '%" + stBusca + "%'";
				int idArt = Funcion.iEscalarSQL(cdsArticulo, stSelect, true);

				bool bExisteArt = false;
				// Recorre grilla para buscar si ya fue ingresado
				foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in ultraGrid1.Rows.All)
				{
					if (dr.Cells["idArticulo"].Value != System.DBNull.Value)
					{
						if ((int) dr.Cells["idArticulo"].Value == idArt)
						{
							dr.Cells["Cantidad"].Value = (double) dr.Cells["Cantidad"].Value + 1;
							if (!chkCodBarras.Checked) dr.Activate();
							bExisteArt = true;
							break;
						}
					}
				}
				// Si no existe crea linea nueva en caso contrario solo aumenta la cantidad
				if (!bExisteArt)
				{
					ultraGrid1.Rows.Band.AddNew();
					ultraGrid1.ActiveCell = ultraGrid1.ActiveRow.Cells["CodArt"];
					ultraGrid1.ActiveCell.Value = idArt;
					ultraGrid1.ActiveCell = ultraGrid1.ActiveRow.Cells["idArticulo"];
					ultraGrid1.ActiveCell.Value = idArt;
				}
				if (chkCodBarras.Checked)	txtBusca.SelectAll();
				else 
				{
					ultraGrid1.Select();
					ultraGrid1.ActiveCell = ultraGrid1.ActiveRow.Cells["idArticulo"];
					ultraGrid1.PerformAction(UltraGridAction.EnterEditMode, false, false);
				}
			}
			else if (iCuenta > 1)
			{
				btBuscaArticulo_Click(this, e);
			}
			else
				MessageBox.Show("No existen articulos que cumplan condicion", "Informacion");
		}

		private void cmbCredTribut_MouseDown(object sender, System.Windows.Forms.MouseEventArgs e)
		{
			if (e.Button != MouseButtons.Right) return;
			using(CompraCredTribut miCred = new CompraCredTribut())
			{
				miCred.ShowDialog();
			}
		}

		private void cmbIce_MouseDown(object sender, System.Windows.Forms.MouseEventArgs e)
		{
			if (e.Button != MouseButtons.Right) return;
			using (CompraPorcIce miPorcIce = new CompraPorcIce())
			{
				miPorcIce.ShowDialog();
			}
		}

		private void btComprobante_Click(object sender, System.EventArgs e)
		{
			if (this.txtComprobante.Text.Trim().Length == 0)
			{
				MessageBox.Show("Ingrese Comprobante", "Informacion");
			}
			else
			{
				int idCliente = (int) cmbCliente.Value;
				int iTipoDest = 4;
				if (IdTipoFactura == 5 || IdTipoFactura == 11) iTipoDest = 1;
				string stSelect = "Select idCompra From Compra Where Numero = '"
					+ this.txtComprobante.Text.Trim() + "' And ContadoCredito = 2 And idTipoFactura = "
					+ iTipoDest.ToString() + " And idCliente = " + idCliente.ToString();

				int num1 = Funcion.iEscalarSQL(this.cdsCompra, stSelect, false);
				if (num1 < 1) // No es devolucion en cualquier parte
				{
					stSelect = "Select idCompra From Compra Where Numero = '" 
						+ this.txtComprobante.Text.Trim() + "' And idCliente = " + idCliente.ToString();
					num1 = Funcion.iEscalarSQL(this.cdsCompra, stSelect, false);
				}

				if (num1 < 1)
				{
					MessageBox.Show("No existe una factura que coincida con este comprobante", "Informacion");
				}
				else
				{
					stSelect = "Select idTipoFactura From Compra Where idCompra = " + num1.ToString();
					int num2 = Funcion.iEscalarSQL(this.cdsCompra, stSelect, false);
					Compra compra1 = new Compra(num2, num1);
					compra1.MdiParent = base.MdiParent;
					compra1.Show();
				}
			}
		}

		private void ultraButton1_Click(object sender, System.EventArgs e)
		{
			string text1 = this.txtDepartamento.Text.Trim();
			if (text1.Length == 0)
			{
				MessageBox.Show("Ingrese Importacion", "Informacion");
			}
			else
			{
				this.Cursor = Cursors.WaitCursor;
				string text2 = "{Compra.Departamento} = '" + text1 
					+ "' And {Compra.idTipoFactura} = " + this.IdTipoFactura.ToString();
				if (IdTipoFactura == 14)
					new Reporte("ImportacionArancel", text2).Show();
				else
					new Reporte("NotaEntregaImp", text2).Show();
				this.Cursor = Cursors.Default;
			}
		}

		private void cmbProyecto_MouseUp(object sender, System.Windows.Forms.MouseEventArgs e)
		{
			if (e.Button == MouseButtons.Middle)
			{
				using (Proyecto miProyecto = new Proyecto())
				{
					miProyecto.ShowDialog();
				}
				this.cdsProyecto.Clear();
				this.cdsProyecto.Fill();
				return;
			}
			if (e.Button == MouseButtons.Right)
			{
				if (ultraGrid1.ActiveRow != null)
					ultraGrid1.ActiveRow.Cells["idProyecto"].Value = 0;
				cmbProyecto.Value = 0;
				cmbProyecto.Select();
				cmbProyecto.Refresh();
			}
		}

		private void cmbVendedor_MouseUp(object sender, System.Windows.Forms.MouseEventArgs e)
		{
			if (e.Button != MouseButtons.Right) return;
			cmbVendedor.Value = 0;
			cmbVendedor.Select();
			cmbVendedor.Refresh();
		}

		private void cmbSubProyecto_MouseUp(object sender, System.Windows.Forms.MouseEventArgs e)
		{
			if (e.Button == MouseButtons.Middle)
			{
				using (SubProyecto miSubProyecto = new SubProyecto())
				{
					miSubProyecto.ShowDialog();
				}
				this.cdsProyecto.Clear();
				this.cdsProyecto.Fill();
				return;
			}
			if (e.Button == MouseButtons.Right)
			{
				if (ultraGrid1.ActiveRow != null)
					ultraGrid1.ActiveRow.Cells["idSubProyecto"].Value = 0;

				cmbSubProyecto.Value = 0;
				cmbSubProyecto.Select();
				cmbSubProyecto.Refresh();
			}
		}

		private void cmbRecibe_MouseUp(object sender, System.Windows.Forms.MouseEventArgs e)
		{
			if (e.Button != MouseButtons.Right) return;
			cmbRecibe.Value = 0;
			cmbRecibe.Select();
			cmbRecibe.Refresh();
		}

		private void txtAutFactura_MouseDown(object sender, System.Windows.Forms.MouseEventArgs e)
		{
			if (e.Button != MouseButtons.Right || e.Button != MouseButtons.Middle)
			{
				if (IdTipoFactura != 4) return;
				if (drCompra["idCliente"] == null) 
				{
					MessageBox.Show("Selecione el Proveedor", "Informacion");
					return;
				}
				drCompra["AutFactura"] = miProveedor.AutorizaFactura;
				drCompra["AutImprenta"] = miProveedor.AutorizaImprenta;
				drCompra["SerieFactura"] =miProveedor.SerieFactura;
				if (miProveedor.FechaCaducidad > DateTime.Today.AddYears(-10))
					drCompra["FechaCaducidad"] = miProveedor.FechaCaducidad;
			}
		}

		private void btComprob1_Click(object sender, System.EventArgs e)
		{
			if (txtComprob1.Text.Trim().Length == 0)
			{
				MessageBox.Show("Ingrese Comprobante", "Informacion");
			}
			else
			{
				string text1 = "Select idCompra From Compra Where Numero = '" + txtComprob1.Text.Trim() + "'";
				int num1 = Funcion.iEscalarSQL(this.cdsCompra, text1, false);
				if (num1 < 1)
				{
					MessageBox.Show("No existe una factura que coincida con este comprobante", "Informacion");
				}
				else
				{
					text1 = "Select idTipoFactura From Compra Where idCompra = " + num1.ToString();
					int num2 = Funcion.iEscalarSQL(this.cdsCompra, text1, false);
					Compra compra1 = new Compra(num2, num1);
					compra1.MdiParent = base.MdiParent;
					compra1.Show();
				}
			}
		}

		private void btDescuento_Click(object sender, System.EventArgs e)
		{
			if (IdTipoFactura != 4) return; // Aplica solo a compras
			if (txtDescIva.Value == System.DBNull.Value || (double) txtDescIva.Value == 0) return;
			if (txtDescIva0.Value == System.DBNull.Value || (double) txtDescIva0.Value == 0) return;
			double dDescuento = double.Parse(txtDescIva.Value.ToString());
			double dDescuento0 = double.Parse(txtDescIva0.Value.ToString());
			double dSubTotalIva = 0;
			double dSubTotalIva0 = 0;
			if (txtIva.Value != System.DBNull.Value) dSubTotalIva = double.Parse(txtIva.Value.ToString());
			if (txtIva0.Value != System.DBNull.Value) dSubTotalIva0 = double.Parse(txtIva0.Value.ToString());
			if (dSubTotalIva + dSubTotalIva0 == 0) return;
			if (DialogResult.No == MessageBox.Show("Desea Distribuir Descuento proporcionalmente?", 
				"Confirmacion",	MessageBoxButtons.YesNo, 
				MessageBoxIcon.Information)) return;
			double dPorc = dDescuento / (dSubTotalIva);
			double dPorc0 = dDescuento0 / (dSubTotalIva0);
			double dCantidad, dPrecio, dImpuesto;
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.ultraGrid1.Rows.All)
			{
				dCantidad = 0;
				dPrecio = 0;
				dImpuesto = 0;
				if (dr.Cells["Cantidad"].Value != System.DBNull.Value)
					dCantidad = (double) dr.Cells["Cantidad"].Value;
				if (dr.Cells["Precio"].Value != System.DBNull.Value)
					dPrecio = (double) dr.Cells["Precio"].Value;
				if (dr.Cells["Impuesto"].Value != System.DBNull.Value)
					dImpuesto = (double) dr.Cells["Impuesto"].Value;
				if (dImpuesto > 0)
					dr.Cells["DescuentoArt"].Value = dCantidad * dPrecio * dPorc;
				else
					dr.Cells["DescuentoArt"].Value = dCantidad * dPrecio * dPorc0;
			}
			txtDescIva.Value = 0;
			txtDescIva0.Value = 0;
			MessageBox.Show("Descuento Repartido", "Informacion",
				MessageBoxButtons.OK, MessageBoxIcon.Information);
		}

		private void btCambiarPrecio_MouseDown(object sender, System.Windows.Forms.MouseEventArgs e)
		{
			if (e.Button != MouseButtons.Right) return;
			if (!this.barraDato1.bEditar) return;

			// Solo se permite cambiar precios de articulo seleccionado si esta habilitado
			// la opcion de cambiar precio en la ficha del articulo
			if (miAcceso.CambiarPrecio == false)
			{
				int IdArt = int.Parse(ultraGrid1.ActiveRow.Cells["idArticulo"].Value.ToString());
				string stSelect = "SELECT CambiarPrecio FROM Articulo " + 
					"WHERE idArticulo = " + IdArt.ToString();
				bool bCambia = Funcion.bEscalarSQL(cdsArticulo, stSelect, false);
				if (bCambia) // Permite cambiar el precio de un articulo
				{
					stSelect = "SELECT Articulo FROM Articulo " + 
						"WHERE idArticulo = " + IdArt.ToString();
					string stCodigo = Funcion.sEscalarSQL(cdsArticulo, stSelect);
					stSelect = "SELECT Codigo FROM Articulo " + 
						"WHERE idArticulo = " + IdArt.ToString();
					stCodigo = stCodigo.Trim() + "\n";
					stCodigo += Funcion.sEscalarSQL(cdsArticulo, stSelect);
					using (CambiarPrecio miCambio = new CambiarPrecio(stCodigo))
					{
						if (DialogResult.OK == miCambio.ShowDialog())
						{
							ultraGrid1.ActiveRow.Cells["Precio"].Value = miCambio.txtPrecio.Value;
							MessageBox.Show("Cambio de precio Aceptado",
								"Informacion", MessageBoxButtons.OK, MessageBoxIcon.Information);
						}
					}
				}
				else
				{
					MessageBox.Show("No tiene acceso a cambiar precios de Articulo Seleccionado",
						"Informacion", MessageBoxButtons.OK, MessageBoxIcon.Information);
				}
				return;
			}

			// Cambia precio de toda la lista
			if (this.cmbPrecio.Value == null) 
			{
				MessageBox.Show("Ingrese Precio al que desea Actualizar la lista",
					"Informacion", MessageBoxButtons.OK, MessageBoxIcon.Information);
				this.cmbPrecio.Select();
				return;
			}
			string stPrecio = "Precio" + this.cmbPrecio.Value.ToString().Trim();
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.ultraGrid1.Rows.All)
			{
				
				if (dr.Cells["idArticulo"].Value == System.DBNull.Value) continue;
				int IdArt = (int) dr.Cells["idArticulo"].Value;
				if (IdArt < 1) return;
				string stSelectPrecio = "Select "  +stPrecio + " From Articulo Where idArticulo = "
					+ IdArt.ToString();
				dr.Cells["Precio"].Value = Funcion.dEscalarSQL(cdsArticulo, stSelectPrecio, true);
				dr.Cells["Precio"].Activate();
			}
			cmdTotal();
			MessageBox.Show("Precios Modificados",
				"Informacion", MessageBoxButtons.OK,
				MessageBoxIcon.Information);
		}

		private void btComprobante_MouseUp(object sender, System.Windows.Forms.MouseEventArgs e)
		{
			if ((e.Button == MouseButtons.Right) && 
				(IdTipoFactura == 4 || IdTipoFactura == 9 || IdTipoFactura == 14 || IdTipoFactura == 31))
			{
				this.Cursor = Cursors.WaitCursor;
				string stExec = "Exec OrdenCompraConsumo '";
				if (drSeteoF["OrdenC_ConsumoL"] != System.DBNull.Value 
					&& (bool) drSeteoF["OrdenC_ConsumoL"] == true)
					stExec = "Exec OrdenCompraConsumoPedProv '";
				stExec += this.txtComprobante.Text.Trim() + "'";
				Funcion.EjecutaSQL(cdsCompra, stExec, true);
				new Reporte("OrdenCompraConsumo.rpt").Show();
				this.Cursor = Cursors.Default;
			}
		}

		private void btImportacion_MouseUp(object sender, System.Windows.Forms.MouseEventArgs e)
		{
			if (MouseButtons.Right != e.Button) return;
			this.Cursor = Cursors.WaitCursor;
			string stIdCompra = "{Compra.idCompra} = " + txtIdCompra.Value.ToString();
			new Reporte("ImportacionArancel.Rpt", stIdCompra).Show();
			this.Cursor = Cursors.Default;
		}

		private void cdsTipoImportacion_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsTipoImportacion.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);
		}

		private void cmbComprobante1_MouseUp(object sender, System.Windows.Forms.MouseEventArgs e)
		{
			if (e.Button != MouseButtons.Right) return;
			using (CompraComprobante miComprob = new CompraComprobante())
			{
				miComprob.ShowDialog();
			}
			cdsCompraTabla.Clear();
			cdsCompraTabla.Fill();
			cmbComprobante1.Refresh();
			cmbComprobante1.Select();
		}

		private void txtNombre_Validating(object sender, System.ComponentModel.CancelEventArgs e)
		{
			string stFiltro = "0";
			if (Funcion.EsProveedor(IdTipoFactura))
				stFiltro = "1";

			string stNombre = txtNombre.Text.Trim();
			if (stNombre.Length == 0) return;
			string stSelect = "Select Count(*) From Cliente Where Nombre like '"
				+ stNombre + "%' And (Ambos = 1 Or Proveedor = " + stFiltro + ")";
			int iCuenta = Funcion.iEscalarSQL(cdsArticulo, stSelect, true);
			int IdCliente = 0;
			if (iCuenta != 1)
			{
				using (BuscaClientes miBusqueda = new BuscaClientes(stNombre, IdTipoFactura))
				{
					if (DialogResult.OK == miBusqueda.ShowDialog())
					{
						IdCliente = (int) miBusqueda.txtIdBusca.Value;
						drCompra["Nombre"] = miBusqueda.txtDescripcion.Text;
					}
					else e.Cancel=true;
				}
			}
			else
			{
				stSelect = "Select idCliente From Cliente Where Nombre like '"
					+ stNombre + "%' And (Ambos = 1 Or Proveedor = " + stFiltro + ")";
				IdCliente = Funcion.iEscalarSQL(cdsCliente, stSelect, true);
				stSelect = "Select Nombre From Cliente Where idCliente = " + IdCliente.ToString();
				drCompra["Nombre"] = Funcion.sEscalarSQL(cdsCliente, stSelect, true);
			}
			drCompra["idCliente"] = IdCliente;

			if (!barraDato1.bEditar) return;
			if (drCompra["idCliente"] != null && (int) drCompra["idCliente"] > 0)
				miProveedor = new Proveedor(cdsCliente, (int) drCompra["idCliente"]);
			#region Nuevo Registro
			if (barraDato1.bNuevo) 
			{
				if ((IdTipoFactura == 4 || IdTipoFactura == 2) && drCompra["idCliente"] != null && (int) drCompra["idCliente"] > 0)
				{
					drCompra["AutImprenta"] = miProveedor.AutorizaImprenta;
					drCompra["SerieFactura"] = miProveedor.SerieFactura;
					drCompra["AutFactura"] = miProveedor.AutorizaFactura;
					if (miProveedor.FechaCaducidad > DateTime.Today.AddYears(-10))
						drCompra["FechaCaducidad"] = miProveedor.FechaCaducidad;
					if (miProveedor.idCredTribut > 0)
						drCompra["idCredTributario"] = miProveedor.idCredTribut;
				}
				if ((bool) drSeteoF["Copia_Dist"] == false) return;
				if ((int) rdgContadoCred.Value == 2)
				{
					drCompra["DescPorcIva"] = miProveedor.DescuentoCredito;
					drCompra["DescPorcIva0"] = miProveedor.DescuentoCredito;
				}
				else
				{
					drCompra["DescPorcIva"] = miProveedor.DescuentoContado;
					drCompra["DescPorcIva0"] = miProveedor.DescuentoContado;
				}

				drCompra["idVendedor"] = miProveedor.idVendedor;
				drCompra["DiasCredito"] = miProveedor.DiasCredito;
			}
			#endregion Nuevo Registro
		}

		private void Compra_Closing(object sender, System.ComponentModel.CancelEventArgs e)
		{
			if (barraDato1.bEditar == true)
			{
				if (DialogResult.No == MessageBox.Show("Desea Cerrar esta pantalla", "Confirmacion",
					MessageBoxButtons.YesNo, MessageBoxIcon.Question)) 
				{
					e.Cancel = true;
					return;
				}
			}
			int iLong = cmbTipoCompra.Text.Trim().Length;
			if (iLong > 20) iLong = 20;

			string stAudita = "Exec AuditaCrear 53, 7, '" + cmbTipoCompra.Text.Trim().Substring(0, iLong) + "'";
			Funcion.EjecutaSQL(cdsArticulo, stAudita, true);
		}

		private void cdsUnidad_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsUnidad.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);		
		}

		private void btBarras_Click(object sender, System.EventArgs e)
		{
			Cursor = Cursors.WaitCursor;
			string stExec = "Exec BarraFactura " + txtIdCompra.Value.ToString();
			Funcion.EjecutaSQL(cdsArticulo, stExec, true);
			Reporte miRep = new Reporte("BarrasArticulo.rpt", "{BarrasCodigos.idArticulo} > 0");
			miRep.MdiParent = this.MdiParent;
			miRep.Show();
			Cursor = Cursors.Default;		
		}

		private void btAuditoria_Click(object sender, System.EventArgs e)
		{
			Cursor = Cursors.WaitCursor;
			Funcion.AuditoriaImprime(IdTipoFactura, (int) txtIdCompra.Value, true);
			Cursor = Cursors.Default;
		}

		private void btAuditoria_MouseUp(object sender, System.Windows.Forms.MouseEventArgs e)
		{
			if (e.Button != MouseButtons.Right) return;
			Cursor = Cursors.WaitCursor;
			Funcion.AuditoriaImprime(IdTipoFactura, (int) txtIdCompra.Value, false);
			Cursor = Cursors.Default;
		}

		private void btNotaEntrega_MouseDown(object sender, System.Windows.Forms.MouseEventArgs e)
		{
			if (e.Button != MouseButtons.Right)
				return;

			if (DialogResult.No == MessageBox.Show("Desea enviar estos datos por Email?", "Confirmacion",
					MessageBoxButtons.YesNo, MessageBoxIcon.Question)) return;

			string stCorreo = drSeteo["EMail"].ToString();
			if (stCorreo.Trim().Length == 0)
			{
				MessageBox.Show("Ingrese Correo Electronico en Archivo Propiedades Generales", "Informacion",
					MessageBoxButtons.OK, MessageBoxIcon.Information);
				return;
			}
			string stClaveCorreo = drSeteo["ClaveCorreo"].ToString();
			if (stClaveCorreo.Trim().Length == 0)
			{
				MessageBox.Show("Ingrese Clave de Correo, en Archivo Propiedades Generales Pestana con Correo Electronico", "Informacion",
					MessageBoxButtons.OK, MessageBoxIcon.Information);
				return;
			}

			string stUsuarioCorreo = drSeteo["UsuarioCorreo"].ToString();
			if (stUsuarioCorreo.Trim().Length == 0)
			{
				MessageBox.Show("Ingrese Usuario de Correo, en Archivo Propiedades Generales en la Pestana con Correo Electronico", "Informacion",
					MessageBoxButtons.OK, MessageBoxIcon.Information);
				return;
			}

			string stPuertoCorreo = drSeteo["PuertoCorreo"].ToString();
			if (stPuertoCorreo.Trim().Length == 0)
			{
				MessageBox.Show("Ingrese Puerto de Correo, en Archivo Propiedades Generales en la Pestana con Correo Electronico", "Informacion",
					MessageBoxButtons.OK, MessageBoxIcon.Information);
				return;
			}

			if (drSeteo["SslCorreo"] == DBNull.Value)
			{
				MessageBox.Show("Ingrese habilitacion ssl del Correo, en Archivo Propiedades Generales en la Pestana con Correo Electronico", "Informacion",
					MessageBoxButtons.OK, MessageBoxIcon.Information);
				return;
			}
			bool bEnabledSSL = bool.Parse(drSeteo["SslCorreo"].ToString());

			string stServidorCorreo = drSeteo["ServidorCorreo"].ToString();
			if (stServidorCorreo.Trim().Length == 0)
			{
				MessageBox.Show("Ingrese Servidor de Correo ej. smtp.correo.yahoo.es, en Archivo Propiedades Generales Complemento", "Informacion",
					MessageBoxButtons.OK, MessageBoxIcon.Information);
				return;
			}

			
			if (cmbCliente.Value == DBNull.Value || (int) cmbCliente.Value == 0)
			{
				MessageBox.Show("Ingrese Cliente", "Informacion",
					MessageBoxButtons.OK, MessageBoxIcon.Information);
				return;
			}


			string stDestino = "Select IsNull(EMail, '') From Cliente Where idCliente = "
				+ cmbCliente.Value.ToString();
			stDestino = Funcion.sEscalarSQL(cdsCliente, stDestino, true);
			if (stDestino.Trim().Length == 0)
			{
				MessageBox.Show("Cliente no tiene correo electronico", "Informacion",
					MessageBoxButtons.OK, MessageBoxIcon.Information);
				return;
			}
	
			string stCorreoOrigen = drSeteo["Email"].ToString();

			if (stCorreoOrigen.Length == 0)
			{
				MessageBox.Show("Empresa no tiene correo electronico. Ingresele en Archivo - Propiedades - Empresa", "Informacion",
					MessageBoxButtons.OK, MessageBoxIcon.Information);
				return;
			}

			Imprimir(2);

			MailMessage msg = new MailMessage();
			msg.To = stDestino;
			msg.From = stCorreoOrigen;
			msg.Subject = cmbTipoCompra.Text.Trim() + " " + txtNumero.Text.Trim();
			msg.Body = "Revisar informacion de archivo Adjunto";
			msg.Priority = MailPriority.Normal;
			msg.BodyFormat = MailFormat.Text;			//o MailFormat.Html

			// Activar Autenticación

			msg.Fields.Add("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate", "1"); 
			msg.Fields.Add("http://schemas.microsoft.com/cdo/configuration/sendusername", stUsuarioCorreo);// Nombre de usuario
			msg.Fields.Add("http://schemas.microsoft.com/cdo/configuration/sendpassword", stClaveCorreo);  // Contraseña
			msg.Fields.Add("http://schemas.microsoft.com/cdo/configuration/smtpserverport", stPuertoCorreo);
			msg.Fields.Add("http://schemas.microsoft.com/cdo/configuration/smtpusessl", bEnabledSSL);
			
			SmtpMail.SmtpServer = stServidorCorreo; 	// El servidor de correo
			try
			{
				MailAttachment Attachment = new MailAttachment(MenuLatinium.stDirInicio + @"\Correo.pdf");
				msg.Attachments.Add(Attachment);
				SmtpMail.Send(msg); 
				MessageBox.Show("Correo Electronico Enviado", "Informacion",
					MessageBoxButtons.OK, MessageBoxIcon.Information);
			}
			catch (Exception ex)
			{
				MessageBox.Show(ex.Message, "error en envio de mail",
					MessageBoxButtons.OK, MessageBoxIcon.Error);
			}
		}

		private void label10_Click(object sender, System.EventArgs e)
		{
		
		}

		private void btCentro_Click(object sender, System.EventArgs e)
		{
			using (Proyecto miProyecto = new Proyecto())
			{
				miProyecto.ShowDialog();
			}
			this.cdsProyecto.Clear();
			this.cdsProyecto.Fill();
		}

		private void btCentro_MouseUp(object sender, System.Windows.Forms.MouseEventArgs e)
		{
			if (e.Button != MouseButtons.Right) return;
			if (barraDato1.bEditar)
			{
				MessageBox.Show("Tiene que estar grabada la factura para repartir centros de costo",
					"Informacion");
				return;
			}
			if (DialogResult.No == MessageBox.Show("Desea pasar los Centros de Costo a las lineas?", 
				"Informacion", MessageBoxButtons.YesNo, MessageBoxIcon.Question)) return;
			if (drCompra["idProyecto"] == System.DBNull.Value || (int) drCompra["idProyecto"] == 0)
			{
				MessageBox.Show("No puede pasar Centros de Costo vacios a las lineas", "Informacion");
				return;
			}
			int IdProyecto = (int) drCompra["idProyecto"];
			string stExec = "Exec CentroCostoCompra " + txtIdCompra.Value.ToString();
			Funcion.EjecutaSQL(cdsArticulo, stExec, true);
			barraDato1.ProximoId(5);
		}

		private void btSubProyecto_Click(object sender, System.EventArgs e)
		{
			using (SubProyecto miProyecto = new SubProyecto())
			{
				miProyecto.ShowDialog();
			}
			this.cdsProyecto.Clear();
			this.cdsProyecto.Fill();
		}

		private void btSubProyecto_MouseUp(object sender, System.Windows.Forms.MouseEventArgs e)
		{
			if (e.Button != MouseButtons.Right) return;
			if (barraDato1.bEditar)
			{
				MessageBox.Show("Tiene que estar grabada la factura para repartir Proyectos",
					"Informacion");
				return;
			}
			if (DialogResult.No == MessageBox.Show("Desea pasar los Proyectos a las lineas?", 
				"Informacion", MessageBoxButtons.YesNo, MessageBoxIcon.Question)) return;
			if (drCompra["idSubProyecto"] == System.DBNull.Value || (int) drCompra["idSubProyecto"] == 0)
			{
				MessageBox.Show("No puede pasar Proyectos vacios a las lineas", "Informacion");
				return;
			}
			string stExec = "Exec ProyectoCompra " + txtIdCompra.Value.ToString();
			Funcion.EjecutaSQL(cdsArticulo, stExec, true);
			barraDato1.ProximoId(5);
		}

		private void cmbFormaPago_Validating(object sender, System.ComponentModel.CancelEventArgs e)
		{
		}

		private void cmbFormaPago_Validated(object sender, System.EventArgs e)
		{
			if (IdTipoFactura != 1 || !barraDato1.bEditar) return;
			if (cmbFormaPago.ActiveRow.Cells["Descuento"].Value == DBNull.Value) return;
			double dDescuento = (double) cmbFormaPago.ActiveRow.Cells["Descuento"].Value;
			if (dDescuento == 0) drCompra["Desc1"] = 0;
			drCompra["DescPorcIva"] = dDescuento;
		}

		private void chkAutPrincipal_CheckedChanged(object sender, System.EventArgs e)
		{
		
		}

		private void btTransforma_MouseUp(object sender, System.Windows.Forms.MouseEventArgs e)
		{
			if (e.Button != MouseButtons.Right) return;
			if (barraDato1.bEditar) return;
			if (ultraGrid1.ActiveRow == null) return;
			int IdDetCompra = (int) ultraGrid1.ActiveRow.Cells["idDetCompra"].Value;
			using (AsientoDist miDist = new AsientoDist(IdDetCompra, true))
			{
				try
				{
					miDist.ShowDialog();
				}
				catch(System.Exception ex)
				{
					MessageBox.Show(ex.Message);
				}
				barraDato1.ProximoId(5);
			}
		
		}

		private void btComision_Click(object sender, System.EventArgs e)
		{
			if (barraDato1.bEditar == true)
			{
				MessageBox.Show("No puede generar comisiones mientras edita", "Informacion");
				return;
			}
			CompraComision miComision = new CompraComision((int) txtIdCompra.Value);
			miComision.MdiParent=this.MdiParent;
			miComision.Show();
		}

		private void btCorresponsal_Click(object sender, System.EventArgs e)
		{
			if (barraDato1.bEditar == true)
			{
				MessageBox.Show("No puede generar corresponsales mientras edita", "Informacion");
				return;
			}
			CompraCorresponsal miCorresponsal = new CompraCorresponsal((int) txtIdCompra.Value);
			miCorresponsal.MdiParent=this.MdiParent;
			miCorresponsal.Show();
		}

		private void btActualizaPrecio_MouseDown(object sender, System.Windows.Forms.MouseEventArgs e)
		{
			if (e.Button != MouseButtons.Right) return;
			if (!(IdTipoFactura == 8 || IdTipoFactura == 9)) return;
			if (DialogResult.No == MessageBox.Show("Desea Actualizar precio a costo promedio",
				"Confirmacion", MessageBoxButtons.YesNo, MessageBoxIcon.Question)) return;

			string stExec = "Exec CompraPrecioCosto " + txtIdCompra.Value.ToString();
			Funcion.EjecutaSQL(cdsArticulo, stExec, true);
			barraDato1.ProximoId(5);
			MessageBox.Show("Precios actualizados al costo promedio", "Informacion");
		}
		public bool CambiosBotrosa()
		{
			bool chek = true;
			Int32 iC = 0 ;
			if (Funcion.bEjecutaSQL(cdsArticulo, "Exec SeteoGExiste 'FTRPB'")) 				
			{				
				#region Desplazamiento de grilla
				foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in ultraGrid1.Rows.All)
				{
					//dCentroCosto = 0;					
					if (dr.Cells["idProyecto"].Value != System.DBNull.Value)
					{						
						iC = (Int32) dr.Cells["idProyecto"].Value;						
					}
					if (iC <= 0 )
					{
						chek=false;
					}
				}
				#endregion Desplazamiento de grilla	
			}
			 return chek;
		}

		private void cmbComprobante1_Validated(object sender, System.EventArgs e)
		{
			if ((int) cmbComprobante1.Value == 2 || (int) cmbComprobante1.Value == 12 || (int) cmbComprobante1.Value == 18) //Reembolso Gastos
				drCompra["idCredTributario"] = 2;
			else
				drCompra["idCredTributario"] = 1;
		}

		private void cdsSeteoC_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoC.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);		
		}

		private void controlProyecto()
		{
			try
			{
				string va;
				va = cmbProyecto.Value.ToString();				
				if (va.Length <= 0)
				{					
					return;
				}
				else
				{			
					string stExec = "SELECT CentroCosto FROM PROYECTO WHERE IdProyecto = " + va;
					string cod = Funcion.sEscalarSQL(cdsArticulo, stExec, true);
					cod.Substring(0,5);			
					cdvSubProyecto.RowFilter="CodSubProyecto like '"+ cod+"%'";
				}			
			}
			catch(Exception ex){Console.Write(ex);}
		}


		private void cmbProyecto_Validating(object sender, System.ComponentModel.CancelEventArgs e)
		{
			if (Funcion.bEjecutaSQL(cdsArticulo, "Exec SeteoGExiste 'FTRPBP'")) 				
			{				
				controlProyecto();
			}
		}
		private void controlCCP()
		{
			int iCuenta =0;
			string txtE = ultraGrid1.ActiveRow.Cells["CentroCostoBR"].Value.ToString();
			if (txtE.Length > 0)
			{
				try
				{
					string stSelect = "EXEC BuscaRCentroC '" + txtE + "',1";
					iCuenta = Funcion.iEscalarSQL(cdsArticulo, stSelect, true);
					if (iCuenta == 0)
					{
						MessageBox.Show("No Existen Registros Coincidentes");
						ultraGrid1.ActiveRow.Cells["CentroCostoBR"].Value = "";												                  
						return;                  
					}
					if (iCuenta > 0)
					{						
						string stEmpresa = "Select Nombre from Proyecto where idProyecto = "+iCuenta;
						string CentroC = Funcion.sEscalarSQL(cdsArticulo, stEmpresa, true);							
						controlW = CentroC;							
						ultraGrid1.ActiveRow.Cells["CentroCostoBR"].Value = CentroC;					
					}            
					if (iCuenta < 0)
					{
						using (BuscaCentroCosto miBusca = new BuscaCentroCosto(txtE))
						{
							if (DialogResult.OK == miBusca.ShowDialog())
							{
								try
								{
									int id = miBusca.idBusca;
									string stEmpresa1 = "Select Nombre from Proyecto where idProyecto = "+id;
									string CentroC1 = Funcion.sEscalarSQL(cdsArticulo, stEmpresa1, true);									
									controlW = CentroC1;									
									ultraGrid1.ActiveRow.Cells["CentroCostoBR"].Value = CentroC1;										
								}
								catch { }
							}
						}
					}			
				}
				catch (Exception ex) {Console.Write(ex); }			
			}
		}
		public string codCC(string cod)
		{
			string stCodCC = "Select CentroCosto from Proyecto where nombre like '%"+cod+"%'";
			string IdGlobal = Funcion.sEscalarSQL(cdsArticulo, stCodCC, true);	
			return IdGlobal;
		}
		private void controlCPP(string cod1)
		{
			int iCuenta =0;
			string txtE = ultraGrid1.ActiveRow.Cells["ProyectoBR"].Value.ToString();
			if (txtE.Length > 0)
			{
				try
				{
					string stSelect = "EXEC BuscaRProyecto '" + txtE + "',1,"+cod1;
					iCuenta = Funcion.iEscalarSQL(cdsArticulo, stSelect, true);
					if (iCuenta == 0)
					{
						MessageBox.Show("No Existen Registros Coincidentes");
						ultraGrid1.ActiveRow.Cells["ProyectoBR"].Value = "";												                  
						return;                  
					}
					if (iCuenta > 0)
					{						
						string stEmpresa = "Select Nombre from SubProyecto where idSubProyecto = "+iCuenta;
						string CentroC = Funcion.sEscalarSQL(cdsArticulo, stEmpresa, true);							
						controlW1 = CentroC;							
						ultraGrid1.ActiveRow.Cells["ProyectoBR"].Value = CentroC;					
					}            
					if (iCuenta < 0)
					{
					
						using (BuscaCentroCosto miBusca1 = new BuscaCentroCosto(txtE,cod1))
						{
							if (DialogResult.OK == miBusca1.ShowDialog())
							{
								try
								{
									int id = miBusca1.idBusca;
									string stEmpresa1 = "Select Nombre from SubProyecto where idSubProyecto = "+id;
									string CentroC1 = Funcion.sEscalarSQL(cdsArticulo, stEmpresa1, true);									
									controlW1 = CentroC1;									
									ultraGrid1.ActiveRow.Cells["ProyectoBR"].Value = CentroC1;										
								}
								catch { }
							}
						}
					}			
				}
				catch (Exception ex) {Console.Write(ex); }			
			}
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
					MessageBox.Show(ex.Message, "Error");
				}
			}

		}
		public string revisaCadena(string cadenaOrg)
		{
			string letra = "";
			string stL = "'";
			string cadena = cadenaOrg;
			string cadena1 = "";
			for (int i=0;i<cadena.Length;i++)
			{
				letra = cadena.Substring(i,1);				
				if(letra == stL)
				{					
					cadena1 += "-";
				}
				else
				{
					cadena1 += letra;
				}
			}
			//				MessageBox.Show(" "+cadena1);
			return cadena1;
		}

		private void ugrBusca_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
			e.Layout.Bands[0].Columns["idArticulo"].Hidden = true;
		}

	}

}
