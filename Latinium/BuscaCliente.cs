using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de BuscaAsiento.
	/// </summary>
	public class BuscaCliente : DevExpress.XtraEditors.XtraForm
	{
		public System.Windows.Forms.TextBox idBusca;
		private System.Windows.Forms.ErrorProvider errorProvider;
		private C1.Data.C1DataView cdvCliente;
		private C1.Data.C1DataSet cdsCliente;
		private System.ComponentModel.IContainer components=null;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtNombre;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtCodigo;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtContacto;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtNumero;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtRepresentante;
		private DevExpress.XtraEditors.SimpleButton btBuscar;
		private DevExpress.XtraEditors.SimpleButton btAceptar;
		private DevExpress.XtraEditors.SimpleButton btCancelar;
		private Infragistics.Win.UltraWinGrid.UltraGrid ultraGrid1;
		private Infragistics.Win.UltraWinEditors.UltraOptionSet rdgTipoBusqueda;
		private Infragistics.Win.UltraWinEditors.UltraOptionSet rdgOpcion;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtFecha;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtEmail;
		public Infragistics.Win.UltraWinEditors.UltraCheckEditor chkOrden;


		private bool TipoCliente = true;
		public BuscaCliente(bool bTipo)
		{
			InitializeComponent();
			TipoCliente = bTipo;
		}

		public BuscaCliente()
		{
			InitializeComponent();
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
			Infragistics.Win.Appearance appearance1 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Cliente", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCliente");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idGrupoCliente");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCobrador");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idVendedor");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idTipoPrecio");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCiudad");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idSector");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idTipoRuc");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idSucursal");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPadre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn11 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idFormaPago");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn12 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCredTribut");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn13 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn14 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn15 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Direccion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn16 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Telefono");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn17 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Fax");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn18 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Ruc");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn19 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Observacion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn20 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Representante");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn21 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Email");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn22 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Celular");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn23 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Contacto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn24 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Pais");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn25 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Notas");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn26 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Ambos");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn27 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Proveedor");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn28 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("AutorizaFactura");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn29 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("AutorizaImprenta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn30 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("MaximoCredito");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn31 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("DiasCredito");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn32 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("SerieFactura");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn33 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Fecha");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn34 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaIngreso");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn35 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("DescuentoContado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn36 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("DescuentoCredito");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn37 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Borrar");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn38 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("SaldoCxC");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn39 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("SaldoCxP");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn40 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Editar");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn41 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nuevo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn42 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaEditado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn43 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idRetRenta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn44 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idRetIva");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn45 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaCaducidad");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn46 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("AutorizaRet");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn47 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("SerieRet");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn48 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ContribEspecial");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn49 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CuentaCorr");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn50 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CuentaAh");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn51 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Conyuge");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn52 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ConyDireccion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn53 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ConyTelefono");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn54 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ConyCelular");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn55 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Referencia");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn56 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("RefTelefono");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn57 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("RefDireccion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn58 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Poliza");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn59 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Parroquia");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn60 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Sector");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn61 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("DesconProv");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn62 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("DesconCliente");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn63 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idClientePublicidad");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn64 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("InstPublica");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn65 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("PorComT");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn66 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("NetoValorT");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn67 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ProveMaster");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn68 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Entrega");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn69 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("RangoHasta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn70 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("RangoDesde");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn71 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idBanco");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn72 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Prospecto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn73 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Foto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn74 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Interes");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn75 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaCaducaCredito");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn76 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Comercial");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn77 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Vencimiento");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn78 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("RefLocal");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn79 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPais");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn80 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Especial");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn81 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ContadoCred");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn82 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCuentaAnticipo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn83 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCuentaFondo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn84 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Chofer");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn85 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("SujetoRetencion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn86 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Aprobado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn87 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ClteCredito");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn88 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idEstado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn89 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idInstitucion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn90 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("PointAutorizado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn91 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("PointCuota");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn92 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("PointIdInstitucion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn93 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("PointIdInstitucionFinanaciera");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn94 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idObservaciones");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn95 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idSolicitud");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn96 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("HaceRetencion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn97 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("AplicaConsignacion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn98 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idTipoCredito");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn99 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("DiasCredito1");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn100 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("DiasCredito2");
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance11 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance12 = new Infragistics.Win.Appearance();
			Infragistics.Win.ValueListItem valueListItem1 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem2 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem3 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.Appearance appearance13 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance14 = new Infragistics.Win.Appearance();
			Infragistics.Win.ValueListItem valueListItem4 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem5 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem6 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem7 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem8 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem9 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem10 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.Appearance appearance15 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton1 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.Appearance appearance16 = new Infragistics.Win.Appearance();
			this.cdvCliente = new C1.Data.C1DataView();
			this.cdsCliente = new C1.Data.C1DataSet();
			this.idBusca = new System.Windows.Forms.TextBox();
			this.errorProvider = new System.Windows.Forms.ErrorProvider();
			this.txtNombre = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtCodigo = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtContacto = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtNumero = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtRepresentante = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.btBuscar = new DevExpress.XtraEditors.SimpleButton();
			this.btAceptar = new DevExpress.XtraEditors.SimpleButton();
			this.btCancelar = new DevExpress.XtraEditors.SimpleButton();
			this.ultraGrid1 = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.rdgTipoBusqueda = new Infragistics.Win.UltraWinEditors.UltraOptionSet();
			this.rdgOpcion = new Infragistics.Win.UltraWinEditors.UltraOptionSet();
			this.dtFecha = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.txtEmail = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.chkOrden = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			((System.ComponentModel.ISupportInitialize)(this.cdvCliente)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsCliente)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNombre)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCodigo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtContacto)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNumero)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtRepresentante)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.rdgTipoBusqueda)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.rdgOpcion)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFecha)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtEmail)).BeginInit();
			this.SuspendLayout();
			// 
			// cdvCliente
			// 
			this.cdvCliente.BindingContextCtrl = this;
			this.cdvCliente.DataSet = this.cdsCliente;
			this.cdvCliente.Sort = "idCliente";
			this.cdvCliente.TableName = "";
			this.cdvCliente.TableViewName = "Cliente";
			// 
			// cdsCliente
			// 
			this.cdsCliente.BindingContextCtrl = this;
			this.cdsCliente.CaseSensitive = false;
			this.cdsCliente.DataLibrary = "LibFacturacion";
			this.cdsCliente.DataLibraryUrl = "";
			this.cdsCliente.DataSetDef = "dsCliente";
			this.cdsCliente.FillOnRequest = false;
			this.cdsCliente.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsCliente.Name = "cdsCliente";
			this.cdsCliente.SchemaClassName = "LibFacturacion.DataClass";
			this.cdsCliente.SchemaDef = null;
			this.cdsCliente.BeforeFill += new C1.Data.FillEventHandler(this.cdsCliente_BeforeFill);
			// 
			// idBusca
			// 
			this.idBusca.DataBindings.Add(new System.Windows.Forms.Binding("Text", this.cdsCliente, "Cliente.idCliente"));
			this.idBusca.Location = new System.Drawing.Point(480, 17);
			this.idBusca.Name = "idBusca";
			this.idBusca.Size = new System.Drawing.Size(48, 20);
			this.idBusca.TabIndex = 23;
			this.idBusca.Text = "";
			// 
			// errorProvider
			// 
			this.errorProvider.ContainerControl = this;
			this.errorProvider.DataMember = "";
			// 
			// txtNombre
			// 
			appearance1.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtNombre.Appearance = appearance1;
			this.txtNombre.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNombre.Location = new System.Drawing.Point(136, 17);
			this.txtNombre.Name = "txtNombre";
			this.txtNombre.Size = new System.Drawing.Size(328, 22);
			this.txtNombre.TabIndex = 29;
			this.txtNombre.ValueChanged += new System.EventHandler(this.txtNombre_ValueChanged);
			// 
			// txtCodigo
			// 
			appearance2.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtCodigo.Appearance = appearance2;
			this.txtCodigo.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtCodigo.Location = new System.Drawing.Point(136, 43);
			this.txtCodigo.Name = "txtCodigo";
			this.txtCodigo.Size = new System.Drawing.Size(112, 22);
			this.txtCodigo.TabIndex = 30;
			this.txtCodigo.ValueChanged += new System.EventHandler(this.txtCodigo_ValueChanged);
			// 
			// txtContacto
			// 
			appearance3.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtContacto.Appearance = appearance3;
			this.txtContacto.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtContacto.Location = new System.Drawing.Point(136, 69);
			this.txtContacto.Name = "txtContacto";
			this.txtContacto.Size = new System.Drawing.Size(112, 22);
			this.txtContacto.TabIndex = 31;
			// 
			// txtNumero
			// 
			appearance4.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtNumero.Appearance = appearance4;
			this.txtNumero.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNumero.Location = new System.Drawing.Point(136, 95);
			this.txtNumero.Name = "txtNumero";
			this.txtNumero.Size = new System.Drawing.Size(112, 22);
			this.txtNumero.TabIndex = 32;
			// 
			// txtRepresentante
			// 
			appearance5.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtRepresentante.Appearance = appearance5;
			this.txtRepresentante.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtRepresentante.Location = new System.Drawing.Point(136, 121);
			this.txtRepresentante.Name = "txtRepresentante";
			this.txtRepresentante.Size = new System.Drawing.Size(112, 22);
			this.txtRepresentante.TabIndex = 33;
			// 
			// btBuscar
			// 
			this.btBuscar.Location = new System.Drawing.Point(264, 172);
			this.btBuscar.Name = "btBuscar";
			this.btBuscar.Size = new System.Drawing.Size(75, 25);
			this.btBuscar.TabIndex = 34;
			this.btBuscar.Text = "&Buscar";
			this.btBuscar.Click += new System.EventHandler(this.btBuscar_Click);
			// 
			// btAceptar
			// 
			this.btAceptar.Location = new System.Drawing.Point(352, 172);
			this.btAceptar.Name = "btAceptar";
			this.btAceptar.Size = new System.Drawing.Size(75, 25);
			this.btAceptar.TabIndex = 35;
			this.btAceptar.Text = "&Aceptar";
			this.btAceptar.Click += new System.EventHandler(this.btAceptar_Click);
			this.btAceptar.MouseUp += new System.Windows.Forms.MouseEventHandler(this.btAceptar_MouseUp);
			// 
			// btCancelar
			// 
			this.btCancelar.DialogResult = System.Windows.Forms.DialogResult.Cancel;
			this.btCancelar.Location = new System.Drawing.Point(440, 172);
			this.btCancelar.Name = "btCancelar";
			this.btCancelar.Size = new System.Drawing.Size(75, 25);
			this.btCancelar.TabIndex = 36;
			this.btCancelar.Text = "&Cancelar";
			this.btCancelar.Click += new System.EventHandler(this.btCancelar_Click);
			// 
			// ultraGrid1
			// 
			this.ultraGrid1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.ultraGrid1.Cursor = System.Windows.Forms.Cursors.Default;
			this.ultraGrid1.DataMember = "Cliente";
			this.ultraGrid1.DataSource = this.cdsCliente;
			appearance6.BackColor = System.Drawing.Color.White;
			this.ultraGrid1.DisplayLayout.Appearance = appearance6;
			ultraGridColumn1.Header.VisiblePosition = 21;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn2.Header.VisiblePosition = 23;
			ultraGridColumn2.Hidden = true;
			ultraGridColumn3.Header.VisiblePosition = 22;
			ultraGridColumn3.Hidden = true;
			ultraGridColumn4.Header.VisiblePosition = 27;
			ultraGridColumn4.Hidden = true;
			ultraGridColumn5.Header.VisiblePosition = 25;
			ultraGridColumn5.Hidden = true;
			ultraGridColumn6.Header.VisiblePosition = 20;
			ultraGridColumn6.Hidden = true;
			ultraGridColumn7.Header.VisiblePosition = 24;
			ultraGridColumn7.Hidden = true;
			ultraGridColumn8.Header.VisiblePosition = 26;
			ultraGridColumn8.Hidden = true;
			ultraGridColumn9.Header.VisiblePosition = 35;
			ultraGridColumn9.Hidden = true;
			ultraGridColumn10.Header.VisiblePosition = 39;
			ultraGridColumn10.Hidden = true;
			ultraGridColumn11.Header.VisiblePosition = 45;
			ultraGridColumn11.Hidden = true;
			ultraGridColumn12.Header.VisiblePosition = 48;
			ultraGridColumn12.Hidden = true;
			ultraGridColumn13.Header.VisiblePosition = 0;
			ultraGridColumn13.Width = 199;
			ultraGridColumn14.Header.VisiblePosition = 1;
			ultraGridColumn14.Width = 86;
			ultraGridColumn15.Header.VisiblePosition = 17;
			ultraGridColumn15.Hidden = true;
			ultraGridColumn16.Header.VisiblePosition = 34;
			ultraGridColumn16.Hidden = true;
			ultraGridColumn17.Header.VisiblePosition = 18;
			ultraGridColumn17.Hidden = true;
			ultraGridColumn18.Header.VisiblePosition = 3;
			ultraGridColumn18.Width = 101;
			ultraGridColumn19.Header.VisiblePosition = 30;
			ultraGridColumn19.Hidden = true;
			ultraGridColumn20.Header.VisiblePosition = 4;
			ultraGridColumn20.Width = 97;
			ultraGridColumn21.Header.VisiblePosition = 6;
			ultraGridColumn22.Header.VisiblePosition = 13;
			ultraGridColumn22.Hidden = true;
			ultraGridColumn23.Header.VisiblePosition = 2;
			ultraGridColumn23.Width = 95;
			ultraGridColumn24.Header.VisiblePosition = 31;
			ultraGridColumn24.Hidden = true;
			ultraGridColumn25.Header.VisiblePosition = 29;
			ultraGridColumn25.Hidden = true;
			ultraGridColumn26.Header.VisiblePosition = 9;
			ultraGridColumn26.Hidden = true;
			ultraGridColumn27.Header.VisiblePosition = 32;
			ultraGridColumn27.Hidden = true;
			ultraGridColumn28.Header.VisiblePosition = 10;
			ultraGridColumn28.Hidden = true;
			ultraGridColumn29.Header.VisiblePosition = 11;
			ultraGridColumn29.Hidden = true;
			ultraGridColumn30.Header.VisiblePosition = 28;
			ultraGridColumn30.Hidden = true;
			ultraGridColumn31.Header.VisiblePosition = 16;
			ultraGridColumn31.Hidden = true;
			ultraGridColumn32.Header.VisiblePosition = 33;
			ultraGridColumn32.Hidden = true;
			ultraGridColumn33.Header.VisiblePosition = 5;
			ultraGridColumn33.Width = 92;
			ultraGridColumn34.Header.VisiblePosition = 19;
			ultraGridColumn34.Hidden = true;
			ultraGridColumn35.Header.VisiblePosition = 14;
			ultraGridColumn35.Hidden = true;
			ultraGridColumn36.Header.VisiblePosition = 15;
			ultraGridColumn36.Hidden = true;
			ultraGridColumn37.Header.VisiblePosition = 12;
			ultraGridColumn37.Hidden = true;
			ultraGridColumn38.Header.VisiblePosition = 36;
			ultraGridColumn38.Hidden = true;
			ultraGridColumn39.Header.VisiblePosition = 37;
			ultraGridColumn39.Hidden = true;
			ultraGridColumn40.Header.VisiblePosition = 38;
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
			ultraGridColumn46.Header.VisiblePosition = 46;
			ultraGridColumn46.Hidden = true;
			ultraGridColumn47.Header.VisiblePosition = 47;
			ultraGridColumn47.Hidden = true;
			ultraGridColumn48.Header.VisiblePosition = 49;
			ultraGridColumn48.Hidden = true;
			ultraGridColumn49.Header.VisiblePosition = 50;
			ultraGridColumn49.Hidden = true;
			ultraGridColumn50.Header.VisiblePosition = 51;
			ultraGridColumn50.Hidden = true;
			ultraGridColumn51.Header.VisiblePosition = 52;
			ultraGridColumn51.Hidden = true;
			ultraGridColumn52.Header.VisiblePosition = 53;
			ultraGridColumn52.Hidden = true;
			ultraGridColumn53.Header.VisiblePosition = 54;
			ultraGridColumn53.Hidden = true;
			ultraGridColumn54.Header.VisiblePosition = 55;
			ultraGridColumn54.Hidden = true;
			ultraGridColumn55.Header.VisiblePosition = 56;
			ultraGridColumn55.Hidden = true;
			ultraGridColumn56.Header.VisiblePosition = 57;
			ultraGridColumn56.Hidden = true;
			ultraGridColumn57.Header.VisiblePosition = 58;
			ultraGridColumn57.Hidden = true;
			ultraGridColumn58.Header.VisiblePosition = 59;
			ultraGridColumn58.Hidden = true;
			ultraGridColumn59.Header.VisiblePosition = 60;
			ultraGridColumn59.Hidden = true;
			ultraGridColumn60.Header.VisiblePosition = 61;
			ultraGridColumn60.Hidden = true;
			ultraGridColumn61.Header.VisiblePosition = 63;
			ultraGridColumn61.Hidden = true;
			ultraGridColumn62.Header.VisiblePosition = 64;
			ultraGridColumn62.Hidden = true;
			ultraGridColumn63.Header.VisiblePosition = 65;
			ultraGridColumn63.Hidden = true;
			ultraGridColumn64.Header.VisiblePosition = 67;
			ultraGridColumn64.Hidden = true;
			ultraGridColumn65.Header.VisiblePosition = 68;
			ultraGridColumn65.Hidden = true;
			ultraGridColumn66.Header.VisiblePosition = 69;
			ultraGridColumn66.Hidden = true;
			ultraGridColumn67.Header.VisiblePosition = 70;
			ultraGridColumn67.Hidden = true;
			ultraGridColumn68.Header.VisiblePosition = 66;
			ultraGridColumn68.Hidden = true;
			ultraGridColumn69.Header.VisiblePosition = 71;
			ultraGridColumn69.Hidden = true;
			ultraGridColumn70.Header.VisiblePosition = 72;
			ultraGridColumn70.Hidden = true;
			ultraGridColumn71.Header.VisiblePosition = 73;
			ultraGridColumn71.Hidden = true;
			ultraGridColumn72.Header.VisiblePosition = 8;
			ultraGridColumn73.Header.VisiblePosition = 74;
			ultraGridColumn73.Hidden = true;
			ultraGridColumn74.Header.VisiblePosition = 75;
			ultraGridColumn74.Hidden = true;
			ultraGridColumn75.Header.VisiblePosition = 76;
			ultraGridColumn75.Hidden = true;
			ultraGridColumn76.Header.VisiblePosition = 7;
			ultraGridColumn77.Header.VisiblePosition = 77;
			ultraGridColumn77.Hidden = true;
			ultraGridColumn78.Header.VisiblePosition = 62;
			ultraGridColumn78.Hidden = true;
			ultraGridColumn79.Header.VisiblePosition = 78;
			ultraGridColumn79.Hidden = true;
			ultraGridColumn80.Header.VisiblePosition = 79;
			ultraGridColumn80.Hidden = true;
			ultraGridColumn81.Header.VisiblePosition = 80;
			ultraGridColumn81.Hidden = true;
			ultraGridColumn82.Header.VisiblePosition = 81;
			ultraGridColumn83.Header.VisiblePosition = 82;
			ultraGridColumn84.Header.VisiblePosition = 83;
			ultraGridColumn85.Header.VisiblePosition = 84;
			ultraGridColumn86.Header.VisiblePosition = 85;
			ultraGridColumn87.Header.VisiblePosition = 86;
			ultraGridColumn88.Header.VisiblePosition = 87;
			ultraGridColumn89.Header.VisiblePosition = 88;
			ultraGridColumn90.Header.VisiblePosition = 89;
			ultraGridColumn91.Header.VisiblePosition = 90;
			ultraGridColumn92.Header.VisiblePosition = 91;
			ultraGridColumn93.Header.VisiblePosition = 92;
			ultraGridColumn94.Header.VisiblePosition = 93;
			ultraGridColumn95.Header.VisiblePosition = 94;
			ultraGridColumn96.Header.VisiblePosition = 95;
			ultraGridColumn97.Header.VisiblePosition = 96;
			ultraGridColumn98.Header.VisiblePosition = 97;
			ultraGridColumn99.Header.VisiblePosition = 98;
			ultraGridColumn100.Header.VisiblePosition = 99;
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
																										 ultraGridColumn100});
			this.ultraGrid1.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.ultraGrid1.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
			this.ultraGrid1.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.ultraGrid1.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.False;
			appearance7.BackColor = System.Drawing.Color.Transparent;
			this.ultraGrid1.DisplayLayout.Override.CardAreaAppearance = appearance7;
			this.ultraGrid1.DisplayLayout.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.RowSelect;
			this.ultraGrid1.DisplayLayout.Override.ColumnAutoSizeMode = Infragistics.Win.UltraWinGrid.ColumnAutoSizeMode.AllRowsInBand;
			appearance8.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance8.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance8.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance8.FontData.BoldAsString = "True";
			appearance8.FontData.Name = "Arial";
			appearance8.FontData.SizeInPoints = 10F;
			appearance8.ForeColor = System.Drawing.Color.White;
			appearance8.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.ultraGrid1.DisplayLayout.Override.HeaderAppearance = appearance8;
			this.ultraGrid1.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti;
			appearance9.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance9.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance9.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.RowSelectorAppearance = appearance9;
			appearance10.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance10.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance10.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.SelectedRowAppearance = appearance10;
			this.ultraGrid1.DisplayLayout.Override.SelectTypeCell = Infragistics.Win.UltraWinGrid.SelectType.None;
			this.ultraGrid1.DisplayLayout.Override.SelectTypeCol = Infragistics.Win.UltraWinGrid.SelectType.None;
			this.ultraGrid1.DisplayLayout.Override.SelectTypeRow = Infragistics.Win.UltraWinGrid.SelectType.Single;
			this.ultraGrid1.DisplayLayout.ScrollBounds = Infragistics.Win.UltraWinGrid.ScrollBounds.ScrollToFill;
			this.ultraGrid1.DisplayLayout.ScrollStyle = Infragistics.Win.UltraWinGrid.ScrollStyle.Immediate;
			this.ultraGrid1.DisplayLayout.TabNavigation = Infragistics.Win.UltraWinGrid.TabNavigation.NextControl;
			this.ultraGrid1.DisplayLayout.ViewStyle = Infragistics.Win.UltraWinGrid.ViewStyle.SingleBand;
			this.ultraGrid1.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.ultraGrid1.Location = new System.Drawing.Point(16, 207);
			this.ultraGrid1.Name = "ultraGrid1";
			this.ultraGrid1.Size = new System.Drawing.Size(504, 232);
			this.ultraGrid1.TabIndex = 37;
			// 
			// rdgTipoBusqueda
			// 
			appearance11.ForeColorDisabled = System.Drawing.Color.Black;
			this.rdgTipoBusqueda.Appearance = appearance11;
			this.rdgTipoBusqueda.CheckedIndex = 2;
			this.rdgTipoBusqueda.ItemAppearance = appearance12;
			this.rdgTipoBusqueda.ItemOrigin = new System.Drawing.Point(10, 0);
			valueListItem1.DataValue = 0;
			valueListItem1.DisplayText = "Que Empiece";
			valueListItem2.DataValue = 1;
			valueListItem2.DisplayText = "Exacta";
			valueListItem3.DataValue = 2;
			valueListItem3.DisplayText = "En cualquier parte";
			this.rdgTipoBusqueda.Items.Add(valueListItem1);
			this.rdgTipoBusqueda.Items.Add(valueListItem2);
			this.rdgTipoBusqueda.Items.Add(valueListItem3);
			this.rdgTipoBusqueda.ItemSpacingVertical = 10;
			this.rdgTipoBusqueda.Location = new System.Drawing.Point(312, 52);
			this.rdgTipoBusqueda.Name = "rdgTipoBusqueda";
			this.rdgTipoBusqueda.Size = new System.Drawing.Size(136, 86);
			this.rdgTipoBusqueda.TabIndex = 42;
			this.rdgTipoBusqueda.Text = "En cualquier parte";
			// 
			// rdgOpcion
			// 
			appearance13.ForeColorDisabled = System.Drawing.Color.Black;
			this.rdgOpcion.Appearance = appearance13;
			this.rdgOpcion.ItemAppearance = appearance14;
			this.rdgOpcion.ItemOrigin = new System.Drawing.Point(6, 0);
			valueListItem4.DataValue = 0;
			valueListItem4.DisplayText = "Nombre";
			valueListItem5.DataValue = 1;
			valueListItem5.DisplayText = "Código";
			valueListItem6.DataValue = 2;
			valueListItem6.DisplayText = "Contacto";
			valueListItem7.DataValue = 3;
			valueListItem7.DisplayText = "Ruc";
			valueListItem8.DataValue = 4;
			valueListItem8.DisplayText = "Representante";
			valueListItem9.DataValue = 5;
			valueListItem9.DisplayText = "Fecha";
			valueListItem10.DataValue = 6;
			valueListItem10.DisplayText = "EMail";
			this.rdgOpcion.Items.Add(valueListItem4);
			this.rdgOpcion.Items.Add(valueListItem5);
			this.rdgOpcion.Items.Add(valueListItem6);
			this.rdgOpcion.Items.Add(valueListItem7);
			this.rdgOpcion.Items.Add(valueListItem8);
			this.rdgOpcion.Items.Add(valueListItem9);
			this.rdgOpcion.Items.Add(valueListItem10);
			this.rdgOpcion.ItemSpacingVertical = 10;
			this.rdgOpcion.Location = new System.Drawing.Point(16, 17);
			this.rdgOpcion.Name = "rdgOpcion";
			this.rdgOpcion.Size = new System.Drawing.Size(112, 181);
			this.rdgOpcion.TabIndex = 43;
			this.rdgOpcion.ValueChanged += new System.EventHandler(this.rdgOpcion_SelectedIndexChanged);
			// 
			// dtFecha
			// 
			appearance15.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtFecha.Appearance = appearance15;
			this.dtFecha.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton1.Caption = "Today";
			this.dtFecha.DateButtons.Add(dateButton1);
			this.dtFecha.Format = "dd/MMM/yyyy";
			this.dtFecha.Location = new System.Drawing.Point(136, 146);
			this.dtFecha.Name = "dtFecha";
			this.dtFecha.NonAutoSizeHeight = 23;
			this.dtFecha.Size = new System.Drawing.Size(112, 21);
			this.dtFecha.SpinButtonsVisible = true;
			this.dtFecha.TabIndex = 44;
			this.dtFecha.Value = new System.DateTime(2005, 5, 17, 0, 0, 0, 0);
			// 
			// txtEmail
			// 
			appearance16.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtEmail.Appearance = appearance16;
			this.txtEmail.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtEmail.Location = new System.Drawing.Point(136, 172);
			this.txtEmail.Name = "txtEmail";
			this.txtEmail.Size = new System.Drawing.Size(112, 22);
			this.txtEmail.TabIndex = 45;
			// 
			// chkOrden
			// 
			this.chkOrden.Location = new System.Drawing.Point(312, 138);
			this.chkOrden.Name = "chkOrden";
			this.chkOrden.Size = new System.Drawing.Size(120, 21);
			this.chkOrden.TabIndex = 46;
			this.chkOrden.Text = "Ordenar al salir";
			// 
			// BuscaCliente
			// 
			this.AcceptButton = this.btBuscar;
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 14);
			this.CancelButton = this.btCancelar;
			this.ClientSize = new System.Drawing.Size(536, 445);
			this.Controls.Add(this.chkOrden);
			this.Controls.Add(this.txtEmail);
			this.Controls.Add(this.dtFecha);
			this.Controls.Add(this.rdgOpcion);
			this.Controls.Add(this.rdgTipoBusqueda);
			this.Controls.Add(this.ultraGrid1);
			this.Controls.Add(this.btCancelar);
			this.Controls.Add(this.btAceptar);
			this.Controls.Add(this.btBuscar);
			this.Controls.Add(this.txtRepresentante);
			this.Controls.Add(this.txtNumero);
			this.Controls.Add(this.txtContacto);
			this.Controls.Add(this.txtCodigo);
			this.Controls.Add(this.txtNombre);
			this.Controls.Add(this.idBusca);
			this.MinimizeBox = false;
			this.Name = "BuscaCliente";
			this.Text = "Búsqueda de Clientes";
			this.Load += new System.EventHandler(this.BuscaCliente_Load);
			((System.ComponentModel.ISupportInitialize)(this.cdvCliente)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsCliente)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNombre)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCodigo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtContacto)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNumero)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtRepresentante)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.rdgTipoBusqueda)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.rdgOpcion)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFecha)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtEmail)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void BuscaCliente_Load(object sender, System.EventArgs e)
		{
			Cursor = Cursors.Default;
			this.idBusca.Width =0;
			this.rdgOpcion.Value = 0;
			this.dtFecha.Value = DateTime.Today;
		}

		private bool Verifica()
		{
			bool bOk = true;
			this.errorProvider.SetError(this.txtNombre, "");
			this.errorProvider.SetError(this.txtCodigo, "");
			this.errorProvider.SetError(this.txtContacto, "");
			this.errorProvider.SetError(this.txtNumero, "");
			this.errorProvider.SetError(this.txtRepresentante, "");

			switch((int) this.rdgOpcion.Value)
			{
				case 0:
					if (this.txtNombre.Value == null)
					{
						this.errorProvider.SetError(this.txtNombre, "Ingrese Nombre");
						this.txtNombre.Select();
						bOk = false;
					}
					break;
				case 1:
					if (this.txtCodigo.Value == null)
					{
						this.errorProvider.SetError(this.txtCodigo, "Ingrese Codigo");
						this.txtCodigo.Select();
						bOk = false;
					}
					break;
				case 2:
					if (this.txtContacto.Value == null)
					{
						this.errorProvider.SetError(this.txtContacto, "Ingrese Contacto");
						this.txtContacto.Select();
						bOk = false;
					}
					break;
				case 3:
					if (this.txtNumero.Value == null)
					{
						this.errorProvider.SetError(this.txtNumero, "Ingrese RUC");
						this.txtNumero.Select();
						bOk = false;
					}
					break;
				case 4:
					if (this.txtRepresentante.Value == null)
					{
						this.errorProvider.SetError(this.txtRepresentante, "Ingrese Representante");
						this.txtRepresentante.Select();
						bOk = false;
					}
					break;
			}
			return bOk;
		}

		public string stOrden = "";
		private void btBuscar_Click(object sender, System.EventArgs e)
		{
			if (!Verifica()) return;
			string strFiltro;
			string strAntes = "";
			string strDespues="'";

//		Opciones de comodines
			if ((int) this.rdgTipoBusqueda.Value == 0) strDespues = "%'";
			if ((int) this.rdgTipoBusqueda.Value == 2) strAntes = "%";
			if ((int) this.rdgTipoBusqueda.Value == 2) strDespues = "%'";

			strFiltro = "";
			string stBusca = "";
			switch((int) this.rdgOpcion.Value)
			{
				case 0:
					stBusca = this.txtNombre.Value.ToString().Trim();
					stBusca = stBusca.Replace("'", "''");
					strFiltro = "Nombre LIKE '" + strAntes + stBusca + strDespues;
					if (Funcion.bEjecutaSQL(cdsCliente, "Exec SeteoGExiste 'CPBCOM'"))
						strFiltro = "(" + strFiltro + " Or Comercial LIKE '" + strAntes + stBusca + strDespues + ") ";
					break;
				case 1:
					stBusca = this.txtCodigo.Value.ToString().Trim();
					stBusca = stBusca.Replace("'", "''");
					strFiltro = "Codigo LIKE '" + strAntes + stBusca + strDespues;
					break;
				case 2:
					stBusca = this.txtContacto.Value.ToString().Trim();
					stBusca = stBusca.Replace("'", "''");
					strFiltro = "Contacto Like '" + strAntes + stBusca + strDespues;
					break;
				case 3:
					stBusca = this.txtNumero.Value.ToString().Trim();
					stBusca = stBusca.Replace("'", "''");
					strFiltro = "Ruc LIKE '" + strAntes + stBusca + strDespues;
					break;
				case 4:
					stBusca = this.txtRepresentante.Value.ToString().Trim();
					stBusca = stBusca.Replace("'", "''");
					strFiltro = "Representante LIKE '" + strAntes + stBusca + strDespues;
					break;
				case 5:
					DateTime txtFecha = (DateTime) this.dtFecha.Value;
					strFiltro = "Fecha = '" + txtFecha.ToString("yyyyMMdd") + "'";
					break;
				case 6:
					stBusca = this.txtEmail.Value.ToString().Trim();
					stBusca = stBusca.Replace("'", "''");
					strFiltro = "EMail LIKE '" + strAntes + stBusca + strDespues;
					break;
			}

//			MessageBox.Show(strFiltro);
			string stFiltro = strFiltro;
			if (this.TipoCliente)
					stFiltro += " And (Proveedor = 0 Or Ambos <> 0)";
			else stFiltro += " And (Proveedor <> 0 Or Ambos <> 0)";

			C1.Data.FilterCondition fcPadre = new 
				C1.Data.FilterCondition(this.cdsCliente, "Cliente", stFiltro);
			C1.Data.FilterConditions fcTotal = new C1.Data.FilterConditions();
			fcTotal.Add(fcPadre);

			this.cdsCliente.Fill(fcTotal, false);

			if (this.cdvCliente.Count == 0)
			{
				MessageBox.Show("No existen registros que cumplan la condicion de busqueda",
					"Información", MessageBoxButtons.OK, MessageBoxIcon.Information);
				return;
			}

			if (this.cdvCliente.Count == 1)
			{
				this.ultraGrid1.Select();
				this.DialogResult = DialogResult.OK;
			}
			else
			{
				this.AcceptButton=this.btAceptar;
				this.ultraGrid1.Select();
			}
		}

//		public int IdBuscado()
//		{
//			int iFila = this.gridView1.FocusedRowHandle;
//			return int.Parse(this.cdvSocio[iFila]["idSocio"].ToString());
//		}

		private void btAceptar_Click(object sender, System.EventArgs e)
		{
			if (chkOrden.Checked)
			{
				switch ((int) rdgOpcion.Value)
				{
					case 0:
						stOrden = "Nombre";
						break;
					case 1:
						stOrden = "Codigo";
						break;
					case 2:
						stOrden = "Contacto";
						break;
					case 3:
						stOrden = "Ruc";
						break;
					case 4:
						stOrden = "Representante";
						break;
					case 5:
						stOrden = "Fecha";
						break;
					case 6:
						stOrden = "EMail";
						break;
				}
			}
			this.DialogResult = DialogResult.OK;
		}

		private void cdsCliente_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			this.cdsCliente.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);
//			return;
//
//			C1.Data.SchemaObjects.DataSetDef dataSetDef = e.DataSet.Schema.DataSetDefs["dsCliente"];
//
//			if (this.TipoCliente)
//				e.Filter.Add(new C1.Data.FilterCondition(dataSetDef.TableViews["Cliente"],
//					"[Proveedor] = 0 Or Ambos <> 0"));
//			else
//				e.Filter.Add(new C1.Data.FilterCondition(dataSetDef.TableViews["Cliente"],
//					"[Proveedor] <> 0 Or Ambos <> 0"));
		}

		private void cdsGrupoArt_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			this.cdsCliente.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);		
		}

		private void rdgOpcion_SelectedIndexChanged(object sender, System.EventArgs e)
		{
			this.txtNombre.Enabled = false;
			this.txtCodigo.Enabled = false;
			this.txtContacto.Enabled = false;
			this.txtNumero.Enabled = false;
			this.txtRepresentante.Enabled = false;
			this.dtFecha.Enabled = false;
			this.txtEmail.Enabled = false;

			switch((int) this.rdgOpcion.Value)
			{
				case 0:
					this.txtNombre.Enabled = true;
					this.txtNombre.Select();
					break;
				case 1:
					this.txtCodigo.Enabled = true;
					this.txtCodigo.Select();
					break;
				case 2:
					this.txtContacto.Enabled = true;
					this.txtContacto.Select();
					break;
				case 3:
					this.txtNumero.Enabled = true;
					this.txtNumero.Select();
					break;
				case 4:
					this.txtRepresentante.Enabled = true;
					this.txtRepresentante.Select();
					break;
				case 5:
					this.dtFecha.Enabled = true;
					this.dtFecha.Select();
					break;
				case 6:
					this.txtEmail.Enabled = true;
					this.txtEmail.Select();
					break;
			}
		}

		private void btCancelar_Click(object sender, System.EventArgs e)
		{
			this.Close();
		}

		private void btAceptar_MouseUp(object sender, System.Windows.Forms.MouseEventArgs e)
		{
			if (e.Button != MouseButtons.Right) return;
			Cliente miCliente = new Cliente(TipoCliente, int.Parse(idBusca.Text));
			miCliente.MdiParent = this.MdiParent;
			miCliente.Show();
		}

		private void txtNombre_ValueChanged(object sender, System.EventArgs e)
		{
		
		}

		private void txtCodigo_ValueChanged(object sender, System.EventArgs e)
		{
		
		}
	}
}
