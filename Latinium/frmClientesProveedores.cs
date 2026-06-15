
using System.Data;
using System.Data.SqlClient;
using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
namespace Latinium
{
	/// <summary>
	/// Descripción breve de frmClientesProveedores.
	/// </summary>
	public class frmClientesProveedores : System.Windows.Forms.Form
	{
		#region descripcion del formulaio
		private Infragistics.Win.UltraWinTabControl.UltraTabControl tbBasico;
		private Infragistics.Win.UltraWinTabControl.UltraTabSharedControlsPage ultraTabSharedControlsPage1;
		private Infragistics.Win.UltraWinTabControl.UltraTabPageControl ultraTabPageControl1;
		private System.Windows.Forms.Label lblTipoProveedor;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbSujetoRetencion;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbAgenteDeRetencion;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbTipoProveedor;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbTipoRuc;
		public Infragistics.Win.UltraWinEditors.UltraTextEditor txtRuc;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtNombreComercial;
		private System.Windows.Forms.Label lblNotas;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtNotas;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtCelular;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtEMail;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtRepresentante;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbCiudad;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbGrupo;
		public Infragistics.Win.UltraWinEditors.UltraTextEditor txtNombre;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtDireccion;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtTelefono;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtFax;
		private Infragistics.Win.UltraWinTabControl.UltraTabPageControl tpComplemento;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtDiasCredito3;
		private System.Windows.Forms.Label label55;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtDiasCredito2;
		private System.Windows.Forms.Label label54;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtValorCredito;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtDiasCredito1;
		private System.Windows.Forms.Label label14;
		private System.Windows.Forms.Label label12;
		private Infragistics.Win.UltraWinTabControl.UltraTabPageControl ultraTabPageControl5;
		private Infragistics.Win.UltraWinTabControl.UltraTabPageControl ultraTabPageControl7;
		private System.Windows.Forms.Label lblAGRetencion;
		private System.Windows.Forms.Label lblRucCed;
		private System.Windows.Forms.Label lblNombreComercial;
		private System.Windows.Forms.Label lblNombre;
		private System.Windows.Forms.Label lblRepresentante;
		private System.Windows.Forms.Label lblDireccion;
		private System.Windows.Forms.Label lblGrupo;
		private System.Windows.Forms.Label lblCelular;
		private System.Windows.Forms.Label lblCiudad;
		private System.Windows.Forms.Label lblTelefono;
		private System.Windows.Forms.Label lblFax;
		private System.Windows.Forms.Label lblEmail;
		private System.Windows.Forms.Label lblPais;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbSector;
		private System.Windows.Forms.Label label19;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtCodigo;
		private System.Windows.Forms.Label label18;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor spnDescCredito;
		private System.Windows.Forms.Label label24;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor spnDescuento;
		private System.Windows.Forms.Label label23;
		private C1.Data.C1DataSet cdsSeteoF;
		private System.Windows.Forms.Label lblSujetoDeRetencion;
		public System.Windows.Forms.Button btnEditar;
		public System.Windows.Forms.Button btnCancelar;
		public System.Windows.Forms.Button btnSalir;
		public System.Windows.Forms.Button btnNuevo;
		public System.Windows.Forms.Button btnConsultar;
		public System.Windows.Forms.Button btnGuardarComoCliente;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbPais;
		public Infragistics.Win.UltraWinEditors.UltraNumericEditor txtidCliente;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtAgenteRetencion;
		private Infragistics.Win.UltraWinEditors.UltraOptionSet optContadoCredito;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource1;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtverificar;
		public Infragistics.Win.UltraWinGrid.UltraGrid ultgContactos;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource2;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkRise;
		private Infragistics.Win.Misc.UltraButton btGrupoCliente;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtDiasConsignacion;
		private System.Windows.Forms.Label lblDiasdeConsignacion;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkCodigoBarras;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkAplicaConsignacion;
		#endregion descripcion del formulaio

		private System.Windows.Forms.Label label1;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtDiasCredito4;
		private System.Windows.Forms.Label lblproductos;
		private Infragistics.Win.UltraWinGrid.UltraGrid ultDirecciones;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource3;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkCourier;
		
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public frmClientesProveedores()
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

		bool bNuevo = false;
		bool bEdicion = false;
		bool bClienteProveedor;
		
		int iOrigen = 0;
		int idClienteC = 0;

		public frmClientesProveedores(bool BClienteProveedor)
		{
		
			InitializeComponent();
		
			bClienteProveedor = BClienteProveedor;
			
		}

		public frmClientesProveedores(bool BClienteProveedor,int IOrigen, int IdClienteC)
		{
			//
			// Necesario para admitir el Diseñador de Windows Forms
			//
			InitializeComponent();

			bClienteProveedor = BClienteProveedor;
			iOrigen = IOrigen;
			idClienteC = IdClienteC;
			//
			// TODO: agregar código de constructor después de llamar a InitializeComponent
			//
		}

	

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
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idClientePais");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Pais");
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idSujetoRetencion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand3 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idAgenteRetencion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand4 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idTipoProveedor");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Tipo");
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand5 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idTipoRuc");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn11 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("TipoRuc");
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance11 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance12 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance13 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand6 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn12 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCiudad");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn13 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Ciudad");
			Infragistics.Win.Appearance appearance14 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand7 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn14 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idGrupoCliente");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn15 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Grupo");
			Infragistics.Win.Appearance appearance15 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance16 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance17 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance18 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance19 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance20 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance21 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance22 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance23 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance24 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance25 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance26 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance27 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance28 = new Infragistics.Win.Appearance();
			Infragistics.Win.ValueListItem valueListItem1 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem2 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.Appearance appearance29 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand8 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn16 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idSector");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn17 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Sector");
			Infragistics.Win.Appearance appearance30 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance31 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance32 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance33 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance34 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance35 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance36 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance37 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand9 = new Infragistics.Win.UltraWinGrid.UltraGridBand("ClienteContacto", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn18 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idClienteContacto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn19 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Contacto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn20 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Telefono");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn21 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Detalle");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn22 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Email");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn23 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Celular");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn24 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cargo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn25 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("EnvioMail");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn26 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("EFE");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn27 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCliente");
			Infragistics.Win.Appearance appearance38 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance39 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance40 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance41 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idClienteContacto");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Contacto");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn3 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Telefono");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn4 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Detalle");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn5 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Email");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn6 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Celular");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn7 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Cargo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn8 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("EnvioMail");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn9 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("EFE");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn10 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCliente");
			Infragistics.Win.Appearance appearance42 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand10 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn28 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idDreccion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn29 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCliente");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn30 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Direccion");
			Infragistics.Win.Appearance appearance43 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance44 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance45 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance46 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn11 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idDreccion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn12 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCliente");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn13 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Direccion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn14 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idClienteContacto");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn15 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCliente");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn16 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Contacto");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn17 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Cargo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn18 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Telefono");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn19 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Celular");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn20 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Detalle");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn21 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("EnvioMail");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn22 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Mail");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn23 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("EFE");
			Infragistics.Win.Appearance appearance47 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance48 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinTabControl.UltraTab ultraTab1 = new Infragistics.Win.UltraWinTabControl.UltraTab();
			Infragistics.Win.Appearance appearance49 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance50 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance51 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance52 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinTabControl.UltraTab ultraTab2 = new Infragistics.Win.UltraWinTabControl.UltraTab();
			Infragistics.Win.UltraWinTabControl.UltraTab ultraTab3 = new Infragistics.Win.UltraWinTabControl.UltraTab();
			Infragistics.Win.UltraWinTabControl.UltraTab ultraTab4 = new Infragistics.Win.UltraWinTabControl.UltraTab();
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmClientesProveedores));
			this.ultraTabPageControl1 = new Infragistics.Win.UltraWinTabControl.UltraTabPageControl();
			this.lblproductos = new System.Windows.Forms.Label();
			this.btGrupoCliente = new Infragistics.Win.Misc.UltraButton();
			this.chkRise = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.txtAgenteRetencion = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.cmbPais = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.lblSujetoDeRetencion = new System.Windows.Forms.Label();
			this.lblAGRetencion = new System.Windows.Forms.Label();
			this.lblTipoProveedor = new System.Windows.Forms.Label();
			this.cmbSujetoRetencion = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.cmbAgenteDeRetencion = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.cmbTipoProveedor = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.cmbTipoRuc = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.lblRucCed = new System.Windows.Forms.Label();
			this.txtRuc = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtNombreComercial = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.lblNombreComercial = new System.Windows.Forms.Label();
			this.lblNotas = new System.Windows.Forms.Label();
			this.txtCelular = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.lblCelular = new System.Windows.Forms.Label();
			this.lblCiudad = new System.Windows.Forms.Label();
			this.txtEMail = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtRepresentante = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.lblTelefono = new System.Windows.Forms.Label();
			this.lblFax = new System.Windows.Forms.Label();
			this.cmbCiudad = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.lblNombre = new System.Windows.Forms.Label();
			this.cmbGrupo = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.lblEmail = new System.Windows.Forms.Label();
			this.lblRepresentante = new System.Windows.Forms.Label();
			this.lblPais = new System.Windows.Forms.Label();
			this.lblDireccion = new System.Windows.Forms.Label();
			this.lblGrupo = new System.Windows.Forms.Label();
			this.txtNombre = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtDireccion = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtTelefono = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtFax = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtNotas = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtidCliente = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtverificar = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.tpComplemento = new Infragistics.Win.UltraWinTabControl.UltraTabPageControl();
			this.chkCourier = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.label1 = new System.Windows.Forms.Label();
			this.txtDiasCredito4 = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtDiasConsignacion = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.lblDiasdeConsignacion = new System.Windows.Forms.Label();
			this.chkCodigoBarras = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.chkAplicaConsignacion = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.optContadoCredito = new Infragistics.Win.UltraWinEditors.UltraOptionSet();
			this.cmbSector = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.label19 = new System.Windows.Forms.Label();
			this.txtCodigo = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label18 = new System.Windows.Forms.Label();
			this.spnDescCredito = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label24 = new System.Windows.Forms.Label();
			this.spnDescuento = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label23 = new System.Windows.Forms.Label();
			this.txtDiasCredito3 = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label55 = new System.Windows.Forms.Label();
			this.txtDiasCredito2 = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label54 = new System.Windows.Forms.Label();
			this.txtValorCredito = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtDiasCredito1 = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label14 = new System.Windows.Forms.Label();
			this.label12 = new System.Windows.Forms.Label();
			this.ultraTabPageControl5 = new Infragistics.Win.UltraWinTabControl.UltraTabPageControl();
			this.ultgContactos = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource2 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.ultraTabPageControl7 = new Infragistics.Win.UltraWinTabControl.UltraTabPageControl();
			this.ultDirecciones = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource3 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.ultraDataSource1 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.tbBasico = new Infragistics.Win.UltraWinTabControl.UltraTabControl();
			this.ultraTabSharedControlsPage1 = new Infragistics.Win.UltraWinTabControl.UltraTabSharedControlsPage();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.btnEditar = new System.Windows.Forms.Button();
			this.btnCancelar = new System.Windows.Forms.Button();
			this.btnSalir = new System.Windows.Forms.Button();
			this.btnNuevo = new System.Windows.Forms.Button();
			this.btnConsultar = new System.Windows.Forms.Button();
			this.btnGuardarComoCliente = new System.Windows.Forms.Button();
			this.ultraTabPageControl1.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.txtAgenteRetencion)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbPais)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbSujetoRetencion)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbAgenteDeRetencion)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbTipoProveedor)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbTipoRuc)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtRuc)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNombreComercial)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCelular)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtEMail)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtRepresentante)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCiudad)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbGrupo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNombre)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDireccion)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtTelefono)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtFax)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNotas)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtidCliente)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtverificar)).BeginInit();
			this.tpComplemento.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.txtDiasCredito4)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDiasConsignacion)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.optContadoCredito)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbSector)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCodigo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.spnDescCredito)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.spnDescuento)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDiasCredito3)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDiasCredito2)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtValorCredito)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDiasCredito1)).BeginInit();
			this.ultraTabPageControl5.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.ultgContactos)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).BeginInit();
			this.ultraTabPageControl7.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.ultDirecciones)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource3)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.tbBasico)).BeginInit();
			this.tbBasico.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			this.SuspendLayout();
			// 
			// ultraTabPageControl1
			// 
			this.ultraTabPageControl1.Controls.Add(this.lblproductos);
			this.ultraTabPageControl1.Controls.Add(this.btGrupoCliente);
			this.ultraTabPageControl1.Controls.Add(this.chkRise);
			this.ultraTabPageControl1.Controls.Add(this.txtAgenteRetencion);
			this.ultraTabPageControl1.Controls.Add(this.cmbPais);
			this.ultraTabPageControl1.Controls.Add(this.lblSujetoDeRetencion);
			this.ultraTabPageControl1.Controls.Add(this.lblAGRetencion);
			this.ultraTabPageControl1.Controls.Add(this.lblTipoProveedor);
			this.ultraTabPageControl1.Controls.Add(this.cmbSujetoRetencion);
			this.ultraTabPageControl1.Controls.Add(this.cmbAgenteDeRetencion);
			this.ultraTabPageControl1.Controls.Add(this.cmbTipoProveedor);
			this.ultraTabPageControl1.Controls.Add(this.cmbTipoRuc);
			this.ultraTabPageControl1.Controls.Add(this.lblRucCed);
			this.ultraTabPageControl1.Controls.Add(this.txtRuc);
			this.ultraTabPageControl1.Controls.Add(this.txtNombreComercial);
			this.ultraTabPageControl1.Controls.Add(this.lblNombreComercial);
			this.ultraTabPageControl1.Controls.Add(this.lblNotas);
			this.ultraTabPageControl1.Controls.Add(this.txtCelular);
			this.ultraTabPageControl1.Controls.Add(this.lblCelular);
			this.ultraTabPageControl1.Controls.Add(this.lblCiudad);
			this.ultraTabPageControl1.Controls.Add(this.txtEMail);
			this.ultraTabPageControl1.Controls.Add(this.txtRepresentante);
			this.ultraTabPageControl1.Controls.Add(this.lblTelefono);
			this.ultraTabPageControl1.Controls.Add(this.lblFax);
			this.ultraTabPageControl1.Controls.Add(this.cmbCiudad);
			this.ultraTabPageControl1.Controls.Add(this.lblNombre);
			this.ultraTabPageControl1.Controls.Add(this.cmbGrupo);
			this.ultraTabPageControl1.Controls.Add(this.lblEmail);
			this.ultraTabPageControl1.Controls.Add(this.lblRepresentante);
			this.ultraTabPageControl1.Controls.Add(this.lblPais);
			this.ultraTabPageControl1.Controls.Add(this.lblDireccion);
			this.ultraTabPageControl1.Controls.Add(this.lblGrupo);
			this.ultraTabPageControl1.Controls.Add(this.txtNombre);
			this.ultraTabPageControl1.Controls.Add(this.txtDireccion);
			this.ultraTabPageControl1.Controls.Add(this.txtTelefono);
			this.ultraTabPageControl1.Controls.Add(this.txtFax);
			this.ultraTabPageControl1.Controls.Add(this.txtNotas);
			this.ultraTabPageControl1.Controls.Add(this.txtidCliente);
			this.ultraTabPageControl1.Controls.Add(this.txtverificar);
			this.ultraTabPageControl1.Location = new System.Drawing.Point(1, 23);
			this.ultraTabPageControl1.Name = "ultraTabPageControl1";
			this.ultraTabPageControl1.Size = new System.Drawing.Size(892, 350);
			this.ultraTabPageControl1.Paint += new System.Windows.Forms.PaintEventHandler(this.ultraTabPageControl1_Paint);
			// 
			// lblproductos
			// 
			this.lblproductos.AutoSize = true;
			this.lblproductos.BackColor = System.Drawing.Color.Transparent;
			this.lblproductos.Location = new System.Drawing.Point(8, 264);
			this.lblproductos.Name = "lblproductos";
			this.lblproductos.Size = new System.Drawing.Size(144, 17);
			this.lblproductos.TabIndex = 292;
			this.lblproductos.Text = "Productos que comercializa:";
			this.lblproductos.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.lblproductos.Visible = false;
			// 
			// btGrupoCliente
			// 
			this.btGrupoCliente.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.btGrupoCliente.Location = new System.Drawing.Point(472, 112);
			this.btGrupoCliente.Name = "btGrupoCliente";
			this.btGrupoCliente.Size = new System.Drawing.Size(24, 21);
			this.btGrupoCliente.TabIndex = 291;
			this.btGrupoCliente.Text = "...";
			this.btGrupoCliente.Click += new System.EventHandler(this.btGrupoCliente_Click_1);
			// 
			// chkRise
			// 
			appearance1.ForeColorDisabled = System.Drawing.Color.Black;
			this.chkRise.Appearance = appearance1;
			this.chkRise.BackColor = System.Drawing.Color.Transparent;
			this.chkRise.Enabled = false;
			this.chkRise.Location = new System.Drawing.Point(448, 16);
			this.chkRise.Name = "chkRise";
			this.chkRise.Size = new System.Drawing.Size(56, 21);
			this.chkRise.TabIndex = 289;
			this.chkRise.Text = "RISE";
			this.chkRise.KeyDown += new System.Windows.Forms.KeyEventHandler(this.chkRise_KeyDown);
			// 
			// txtAgenteRetencion
			// 
			appearance2.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtAgenteRetencion.Appearance = appearance2;
			this.txtAgenteRetencion.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtAgenteRetencion.Enabled = false;
			this.txtAgenteRetencion.FormatString = "0";
			this.txtAgenteRetencion.Location = new System.Drawing.Point(400, 208);
			this.txtAgenteRetencion.MinValue = 0;
			this.txtAgenteRetencion.Name = "txtAgenteRetencion";
			this.txtAgenteRetencion.PromptChar = ' ';
			this.txtAgenteRetencion.Size = new System.Drawing.Size(96, 22);
			this.txtAgenteRetencion.TabIndex = 288;
			this.txtAgenteRetencion.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtAgenteRetencion_KeyDown);
			this.txtAgenteRetencion.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtAgenteRetencion_KeyPress);
			this.txtAgenteRetencion.Validating += new System.ComponentModel.CancelEventHandler(this.txtAgenteRetencion_Validating);
			this.txtAgenteRetencion.Validated += new System.EventHandler(this.txtAgenteRetencion_Validated);
			// 
			// cmbPais
			// 
			appearance3.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbPais.Appearance = appearance3;
			this.cmbPais.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.cmbPais.Cursor = System.Windows.Forms.Cursors.Default;
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn1.Width = 8;
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridColumn2.Width = 152;
			ultraGridBand1.Columns.AddRange(new object[] {
																										 ultraGridColumn1,
																										 ultraGridColumn2});
			this.cmbPais.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.cmbPais.DisplayMember = "Pais";
			this.cmbPais.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbPais.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbPais.Enabled = false;
			this.cmbPais.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbPais.Location = new System.Drawing.Point(698, 160);
			this.cmbPais.Name = "cmbPais";
			this.cmbPais.Size = new System.Drawing.Size(167, 21);
			this.cmbPais.TabIndex = 247;
			this.cmbPais.ValueMember = "idClientePais";
			this.cmbPais.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cmbPais_KeyDown_1);
			this.cmbPais.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.cmbPais_InitializeLayout_1);
			// 
			// lblSujetoDeRetencion
			// 
			this.lblSujetoDeRetencion.AutoSize = true;
			this.lblSujetoDeRetencion.BackColor = System.Drawing.Color.Transparent;
			this.lblSujetoDeRetencion.Location = new System.Drawing.Point(16, 208);
			this.lblSujetoDeRetencion.Name = "lblSujetoDeRetencion";
			this.lblSujetoDeRetencion.Size = new System.Drawing.Size(106, 17);
			this.lblSujetoDeRetencion.TabIndex = 246;
			this.lblSujetoDeRetencion.Text = "Sujeto de retención:";
			this.lblSujetoDeRetencion.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.lblSujetoDeRetencion.Click += new System.EventHandler(this.lblSujetoDeRetencion_Click);
			// 
			// lblAGRetencion
			// 
			this.lblAGRetencion.AutoSize = true;
			this.lblAGRetencion.BackColor = System.Drawing.Color.Transparent;
			this.lblAGRetencion.Location = new System.Drawing.Point(16, 208);
			this.lblAGRetencion.Name = "lblAGRetencion";
			this.lblAGRetencion.Size = new System.Drawing.Size(109, 17);
			this.lblAGRetencion.TabIndex = 241;
			this.lblAGRetencion.Text = "Agente de retención:";
			this.lblAGRetencion.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.lblAGRetencion.Click += new System.EventHandler(this.lblAGRetencion_Click);
			// 
			// lblTipoProveedor
			// 
			this.lblTipoProveedor.AutoSize = true;
			this.lblTipoProveedor.BackColor = System.Drawing.Color.Transparent;
			this.lblTipoProveedor.Location = new System.Drawing.Point(554, 224);
			this.lblTipoProveedor.Name = "lblTipoProveedor";
			this.lblTipoProveedor.TabIndex = 159;
			this.lblTipoProveedor.Text = "Tipo de Proveedor:";
			this.lblTipoProveedor.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.lblTipoProveedor.Visible = false;
			// 
			// cmbSujetoRetencion
			// 
			appearance4.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbSujetoRetencion.Appearance = appearance4;
			this.cmbSujetoRetencion.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.cmbSujetoRetencion.Cursor = System.Windows.Forms.Cursors.Default;
			ultraGridColumn3.Header.VisiblePosition = 0;
			ultraGridColumn3.Hidden = true;
			ultraGridColumn4.Header.VisiblePosition = 1;
			ultraGridColumn4.Width = 328;
			ultraGridBand2.Columns.AddRange(new object[] {
																										 ultraGridColumn3,
																										 ultraGridColumn4});
			this.cmbSujetoRetencion.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			this.cmbSujetoRetencion.DisplayMember = "Nombre";
			this.cmbSujetoRetencion.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbSujetoRetencion.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbSujetoRetencion.Enabled = false;
			this.cmbSujetoRetencion.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbSujetoRetencion.Location = new System.Drawing.Point(168, 208);
			this.cmbSujetoRetencion.Name = "cmbSujetoRetencion";
			this.cmbSujetoRetencion.Size = new System.Drawing.Size(216, 21);
			this.cmbSujetoRetencion.TabIndex = 244;
			this.cmbSujetoRetencion.ValueMember = "idSujetoRetencion";
			this.cmbSujetoRetencion.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cmbSujetoRetencion_KeyDown);
			this.cmbSujetoRetencion.Validating += new System.ComponentModel.CancelEventHandler(this.cmbSujetoRetencion_Validating);
			this.cmbSujetoRetencion.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.cmbSujetoRetencion_InitializeLayout);
			// 
			// cmbAgenteDeRetencion
			// 
			appearance5.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbAgenteDeRetencion.Appearance = appearance5;
			this.cmbAgenteDeRetencion.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.cmbAgenteDeRetencion.Cursor = System.Windows.Forms.Cursors.Default;
			ultraGridColumn5.Header.VisiblePosition = 0;
			ultraGridColumn5.Hidden = true;
			ultraGridColumn6.Header.VisiblePosition = 1;
			ultraGridColumn6.Width = 350;
			ultraGridBand3.Columns.AddRange(new object[] {
																										 ultraGridColumn5,
																										 ultraGridColumn6});
			this.cmbAgenteDeRetencion.DisplayLayout.BandsSerializer.Add(ultraGridBand3);
			this.cmbAgenteDeRetencion.DisplayMember = "Nombre";
			this.cmbAgenteDeRetencion.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbAgenteDeRetencion.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbAgenteDeRetencion.Enabled = false;
			this.cmbAgenteDeRetencion.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbAgenteDeRetencion.Location = new System.Drawing.Point(168, 208);
			this.cmbAgenteDeRetencion.Name = "cmbAgenteDeRetencion";
			this.cmbAgenteDeRetencion.Size = new System.Drawing.Size(216, 21);
			this.cmbAgenteDeRetencion.TabIndex = 243;
			this.cmbAgenteDeRetencion.ValueMember = "idAgenteRetencion";
			this.cmbAgenteDeRetencion.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cmbAgenteDeRetencion_KeyDown);
			this.cmbAgenteDeRetencion.Validating += new System.ComponentModel.CancelEventHandler(this.cmbAgenteDeRetencion_Validating);
			this.cmbAgenteDeRetencion.ValueChanged += new System.EventHandler(this.cmbAgenteDeRetencion_ValueChanged);
			// 
			// cmbTipoProveedor
			// 
			appearance6.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbTipoProveedor.Appearance = appearance6;
			this.cmbTipoProveedor.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbTipoProveedor.Cursor = System.Windows.Forms.Cursors.Default;
			ultraGridColumn7.Header.VisiblePosition = 0;
			ultraGridColumn7.Hidden = true;
			ultraGridColumn8.Header.VisiblePosition = 1;
			ultraGridColumn8.Hidden = true;
			ultraGridColumn8.Width = 60;
			ultraGridColumn9.Header.VisiblePosition = 2;
			ultraGridColumn9.Width = 200;
			ultraGridBand4.Columns.AddRange(new object[] {
																										 ultraGridColumn7,
																										 ultraGridColumn8,
																										 ultraGridColumn9});
			this.cmbTipoProveedor.DisplayLayout.BandsSerializer.Add(ultraGridBand4);
			this.cmbTipoProveedor.DisplayMember = "Tipo";
			this.cmbTipoProveedor.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbTipoProveedor.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbTipoProveedor.Enabled = false;
			this.cmbTipoProveedor.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbTipoProveedor.Location = new System.Drawing.Point(696, 224);
			this.cmbTipoProveedor.Name = "cmbTipoProveedor";
			this.cmbTipoProveedor.Size = new System.Drawing.Size(168, 21);
			this.cmbTipoProveedor.TabIndex = 238;
			this.cmbTipoProveedor.ValueMember = "idTipoProveedor";
			this.cmbTipoProveedor.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cmbTipoProveedor_KeyDown);
			// 
			// cmbTipoRuc
			// 
			appearance7.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbTipoRuc.Appearance = appearance7;
			this.cmbTipoRuc.CausesValidation = false;
			this.cmbTipoRuc.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbTipoRuc.Cursor = System.Windows.Forms.Cursors.Default;
			ultraGridColumn10.Header.VisiblePosition = 0;
			ultraGridColumn10.Hidden = true;
			ultraGridColumn11.Header.VisiblePosition = 1;
			ultraGridColumn11.Width = 233;
			ultraGridBand5.Columns.AddRange(new object[] {
																										 ultraGridColumn10,
																										 ultraGridColumn11});
			this.cmbTipoRuc.DisplayLayout.BandsSerializer.Add(ultraGridBand5);
			this.cmbTipoRuc.DisplayMember = "TipoRuc";
			this.cmbTipoRuc.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbTipoRuc.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbTipoRuc.Enabled = false;
			this.cmbTipoRuc.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbTipoRuc.Location = new System.Drawing.Point(168, 16);
			this.cmbTipoRuc.Name = "cmbTipoRuc";
			this.cmbTipoRuc.Size = new System.Drawing.Size(144, 21);
			this.cmbTipoRuc.TabIndex = 0;
			this.cmbTipoRuc.ValueMember = "idTipoRuc";
			this.cmbTipoRuc.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cmbTipoRuc_KeyDown);
			this.cmbTipoRuc.ValueChanged += new System.EventHandler(this.cmbTipoRuc_ValueChanged);
			// 
			// lblRucCed
			// 
			this.lblRucCed.AutoSize = true;
			this.lblRucCed.BackColor = System.Drawing.Color.Transparent;
			this.lblRucCed.Location = new System.Drawing.Point(16, 16);
			this.lblRucCed.Name = "lblRucCed";
			this.lblRucCed.Size = new System.Drawing.Size(49, 17);
			this.lblRucCed.TabIndex = 156;
			this.lblRucCed.Text = "RUC/Ced";
			this.lblRucCed.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtRuc
			// 
			appearance8.BackColorDisabled = System.Drawing.Color.Silver;
			appearance8.BackColorDisabled2 = System.Drawing.Color.Silver;
			appearance8.BorderColor = System.Drawing.Color.Black;
			appearance8.BorderColor3DBase = System.Drawing.Color.Black;
			appearance8.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtRuc.Appearance = appearance8;
			this.txtRuc.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtRuc.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtRuc.Enabled = false;
			this.txtRuc.Location = new System.Drawing.Point(320, 16);
			this.txtRuc.Name = "txtRuc";
			this.txtRuc.Size = new System.Drawing.Size(112, 22);
			this.txtRuc.TabIndex = 1;
			this.txtRuc.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtRuc_KeyDown);
			this.txtRuc.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtRuc_KeyPress);
			this.txtRuc.Validating += new System.ComponentModel.CancelEventHandler(this.txtRuc_Validating);
			this.txtRuc.Validated += new System.EventHandler(this.txtRuc_Validated);
			// 
			// txtNombreComercial
			// 
			appearance9.BackColorDisabled = System.Drawing.Color.Silver;
			appearance9.BackColorDisabled2 = System.Drawing.Color.Silver;
			appearance9.BorderColor = System.Drawing.Color.Black;
			appearance9.BorderColor3DBase = System.Drawing.Color.Black;
			appearance9.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtNombreComercial.Appearance = appearance9;
			this.txtNombreComercial.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtNombreComercial.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNombreComercial.Enabled = false;
			this.txtNombreComercial.Location = new System.Drawing.Point(168, 176);
			this.txtNombreComercial.Name = "txtNombreComercial";
			this.txtNombreComercial.Size = new System.Drawing.Size(328, 22);
			this.txtNombreComercial.TabIndex = 8;
			this.txtNombreComercial.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtNombreComercial_KeyDown);
			this.txtNombreComercial.ValueChanged += new System.EventHandler(this.txtNombreComercial_ValueChanged);
			// 
			// lblNombreComercial
			// 
			this.lblNombreComercial.AutoSize = true;
			this.lblNombreComercial.BackColor = System.Drawing.Color.Transparent;
			this.lblNombreComercial.Location = new System.Drawing.Point(16, 176);
			this.lblNombreComercial.Name = "lblNombreComercial";
			this.lblNombreComercial.TabIndex = 148;
			this.lblNombreComercial.Text = "Nombre Comercial:";
			this.lblNombreComercial.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.lblNombreComercial.Click += new System.EventHandler(this.lblNombreComercial_Click);
			// 
			// lblNotas
			// 
			this.lblNotas.AutoSize = true;
			this.lblNotas.BackColor = System.Drawing.Color.Transparent;
			this.lblNotas.Location = new System.Drawing.Point(16, 264);
			this.lblNotas.Name = "lblNotas";
			this.lblNotas.Size = new System.Drawing.Size(33, 17);
			this.lblNotas.TabIndex = 105;
			this.lblNotas.Text = "Notas";
			this.lblNotas.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.lblNotas.Visible = false;
			this.lblNotas.Click += new System.EventHandler(this.lblNotas_Click);
			// 
			// txtCelular
			// 
			appearance10.BackColorDisabled = System.Drawing.Color.Silver;
			appearance10.BackColorDisabled2 = System.Drawing.Color.Silver;
			appearance10.BorderColor = System.Drawing.Color.Black;
			appearance10.BorderColor3DBase = System.Drawing.Color.Black;
			appearance10.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtCelular.Appearance = appearance10;
			this.txtCelular.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtCelular.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtCelular.Enabled = false;
			this.txtCelular.Location = new System.Drawing.Point(697, 88);
			this.txtCelular.MaxLength = 10;
			this.txtCelular.Name = "txtCelular";
			this.txtCelular.Size = new System.Drawing.Size(168, 22);
			this.txtCelular.TabIndex = 11;
			this.txtCelular.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtCelular_KeyDown);
			this.txtCelular.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtCelular_KeyPress);
			// 
			// lblCelular
			// 
			this.lblCelular.AutoSize = true;
			this.lblCelular.BackColor = System.Drawing.Color.Transparent;
			this.lblCelular.Location = new System.Drawing.Point(584, 96);
			this.lblCelular.Name = "lblCelular";
			this.lblCelular.Size = new System.Drawing.Size(38, 17);
			this.lblCelular.TabIndex = 87;
			this.lblCelular.Text = "Celular";
			this.lblCelular.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// lblCiudad
			// 
			this.lblCiudad.AutoSize = true;
			this.lblCiudad.BackColor = System.Drawing.Color.Transparent;
			this.lblCiudad.Location = new System.Drawing.Point(581, 192);
			this.lblCiudad.Name = "lblCiudad";
			this.lblCiudad.Size = new System.Drawing.Size(42, 17);
			this.lblCiudad.TabIndex = 10;
			this.lblCiudad.Text = "Ciudad:";
			this.lblCiudad.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtEMail
			// 
			appearance11.BackColorDisabled = System.Drawing.Color.Silver;
			appearance11.BackColorDisabled2 = System.Drawing.Color.Silver;
			appearance11.BorderColor = System.Drawing.Color.Black;
			appearance11.BorderColor3DBase = System.Drawing.Color.Black;
			appearance11.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtEMail.Appearance = appearance11;
			this.txtEMail.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtEMail.Enabled = false;
			this.txtEMail.Location = new System.Drawing.Point(696, 128);
			this.txtEMail.Name = "txtEMail";
			this.txtEMail.Size = new System.Drawing.Size(168, 22);
			this.txtEMail.TabIndex = 12;
			this.txtEMail.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtEMail_KeyDown);
			// 
			// txtRepresentante
			// 
			appearance12.BackColorDisabled = System.Drawing.Color.Silver;
			appearance12.BackColorDisabled2 = System.Drawing.Color.Silver;
			appearance12.BorderColor = System.Drawing.Color.Black;
			appearance12.BorderColor3DBase = System.Drawing.Color.Black;
			appearance12.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtRepresentante.Appearance = appearance12;
			this.txtRepresentante.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtRepresentante.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtRepresentante.Enabled = false;
			this.txtRepresentante.Location = new System.Drawing.Point(168, 144);
			this.txtRepresentante.Name = "txtRepresentante";
			this.txtRepresentante.Size = new System.Drawing.Size(328, 22);
			this.txtRepresentante.TabIndex = 7;
			this.txtRepresentante.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtRepresentante_KeyDown);
			// 
			// lblTelefono
			// 
			this.lblTelefono.AutoSize = true;
			this.lblTelefono.BackColor = System.Drawing.Color.Transparent;
			this.lblTelefono.Location = new System.Drawing.Point(577, 24);
			this.lblTelefono.Name = "lblTelefono";
			this.lblTelefono.Size = new System.Drawing.Size(52, 17);
			this.lblTelefono.TabIndex = 14;
			this.lblTelefono.Text = "Teléfono:";
			this.lblTelefono.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// lblFax
			// 
			this.lblFax.AutoSize = true;
			this.lblFax.BackColor = System.Drawing.Color.Transparent;
			this.lblFax.Location = new System.Drawing.Point(589, 64);
			this.lblFax.Name = "lblFax";
			this.lblFax.Size = new System.Drawing.Size(26, 17);
			this.lblFax.TabIndex = 16;
			this.lblFax.Text = "Fax:";
			this.lblFax.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// cmbCiudad
			// 
			appearance13.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbCiudad.Appearance = appearance13;
			this.cmbCiudad.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.cmbCiudad.Cursor = System.Windows.Forms.Cursors.Default;
			ultraGridColumn12.Header.VisiblePosition = 0;
			ultraGridColumn12.Hidden = true;
			ultraGridColumn13.Header.VisiblePosition = 1;
			ultraGridColumn13.Width = 276;
			ultraGridBand6.Columns.AddRange(new object[] {
																										 ultraGridColumn12,
																										 ultraGridColumn13});
			this.cmbCiudad.DisplayLayout.BandsSerializer.Add(ultraGridBand6);
			this.cmbCiudad.DisplayMember = "Ciudad";
			this.cmbCiudad.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbCiudad.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbCiudad.Enabled = false;
			this.cmbCiudad.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbCiudad.Location = new System.Drawing.Point(696, 192);
			this.cmbCiudad.Name = "cmbCiudad";
			this.cmbCiudad.Size = new System.Drawing.Size(167, 21);
			this.cmbCiudad.TabIndex = 16;
			this.cmbCiudad.ValueMember = "idCiudad";
			this.cmbCiudad.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cmbCiudad_KeyDown);
			// 
			// lblNombre
			// 
			this.lblNombre.AutoSize = true;
			this.lblNombre.BackColor = System.Drawing.Color.Transparent;
			this.lblNombre.Location = new System.Drawing.Point(16, 48);
			this.lblNombre.Name = "lblNombre";
			this.lblNombre.Size = new System.Drawing.Size(48, 17);
			this.lblNombre.TabIndex = 1;
			this.lblNombre.Text = "Nombre:";
			this.lblNombre.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// cmbGrupo
			// 
			appearance14.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbGrupo.Appearance = appearance14;
			this.cmbGrupo.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.cmbGrupo.Cursor = System.Windows.Forms.Cursors.Default;
			ultraGridColumn14.Header.VisiblePosition = 0;
			ultraGridColumn14.Hidden = true;
			ultraGridColumn15.Header.VisiblePosition = 1;
			ultraGridColumn15.Width = 385;
			ultraGridBand7.Columns.AddRange(new object[] {
																										 ultraGridColumn14,
																										 ultraGridColumn15});
			this.cmbGrupo.DisplayLayout.BandsSerializer.Add(ultraGridBand7);
			this.cmbGrupo.DisplayMember = "Grupo";
			this.cmbGrupo.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbGrupo.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbGrupo.Enabled = false;
			this.cmbGrupo.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbGrupo.Location = new System.Drawing.Point(168, 112);
			this.cmbGrupo.Name = "cmbGrupo";
			this.cmbGrupo.Size = new System.Drawing.Size(296, 21);
			this.cmbGrupo.TabIndex = 5;
			this.cmbGrupo.ValueMember = "idGrupoCliente";
			this.cmbGrupo.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cmbGrupo_KeyDown);
			// 
			// lblEmail
			// 
			this.lblEmail.AutoSize = true;
			this.lblEmail.BackColor = System.Drawing.Color.Transparent;
			this.lblEmail.Location = new System.Drawing.Point(584, 136);
			this.lblEmail.Name = "lblEmail";
			this.lblEmail.Size = new System.Drawing.Size(35, 17);
			this.lblEmail.TabIndex = 24;
			this.lblEmail.Text = "Email:";
			this.lblEmail.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// lblRepresentante
			// 
			this.lblRepresentante.AutoSize = true;
			this.lblRepresentante.BackColor = System.Drawing.Color.Transparent;
			this.lblRepresentante.Location = new System.Drawing.Point(16, 144);
			this.lblRepresentante.Name = "lblRepresentante";
			this.lblRepresentante.Size = new System.Drawing.Size(81, 17);
			this.lblRepresentante.TabIndex = 28;
			this.lblRepresentante.Text = "Representante:";
			this.lblRepresentante.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// lblPais
			// 
			this.lblPais.AutoSize = true;
			this.lblPais.BackColor = System.Drawing.Color.Transparent;
			this.lblPais.Location = new System.Drawing.Point(588, 168);
			this.lblPais.Name = "lblPais";
			this.lblPais.Size = new System.Drawing.Size(28, 17);
			this.lblPais.TabIndex = 8;
			this.lblPais.Text = "País:";
			this.lblPais.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// lblDireccion
			// 
			this.lblDireccion.AutoSize = true;
			this.lblDireccion.BackColor = System.Drawing.Color.Transparent;
			this.lblDireccion.Location = new System.Drawing.Point(16, 80);
			this.lblDireccion.Name = "lblDireccion";
			this.lblDireccion.Size = new System.Drawing.Size(54, 17);
			this.lblDireccion.TabIndex = 5;
			this.lblDireccion.Text = "Dirección:";
			this.lblDireccion.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// lblGrupo
			// 
			this.lblGrupo.AutoSize = true;
			this.lblGrupo.BackColor = System.Drawing.Color.Transparent;
			this.lblGrupo.Location = new System.Drawing.Point(16, 112);
			this.lblGrupo.Name = "lblGrupo";
			this.lblGrupo.Size = new System.Drawing.Size(39, 17);
			this.lblGrupo.TabIndex = 36;
			this.lblGrupo.Text = "Grupo:";
			this.lblGrupo.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtNombre
			// 
			appearance15.BackColorDisabled = System.Drawing.Color.Silver;
			appearance15.BackColorDisabled2 = System.Drawing.Color.Silver;
			appearance15.BorderColor = System.Drawing.Color.Black;
			appearance15.BorderColor3DBase = System.Drawing.Color.Black;
			appearance15.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtNombre.Appearance = appearance15;
			this.txtNombre.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtNombre.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNombre.Enabled = false;
			this.txtNombre.Location = new System.Drawing.Point(168, 48);
			this.txtNombre.Name = "txtNombre";
			this.txtNombre.Size = new System.Drawing.Size(328, 22);
			this.txtNombre.TabIndex = 3;
			this.txtNombre.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtNombre_KeyDown);
			this.txtNombre.Validated += new System.EventHandler(this.txtNombre_Validated);
			this.txtNombre.InitializeEditorButtonCheckState += new Infragistics.Win.UltraWinEditors.InitializeCheckStateEventHandler(this.txtNombre_InitializeEditorButtonCheckState);
			// 
			// txtDireccion
			// 
			appearance16.BackColorDisabled = System.Drawing.Color.Silver;
			appearance16.BackColorDisabled2 = System.Drawing.Color.Silver;
			appearance16.BorderColor = System.Drawing.Color.Black;
			appearance16.BorderColor3DBase = System.Drawing.Color.Black;
			appearance16.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtDireccion.Appearance = appearance16;
			this.txtDireccion.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtDireccion.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtDireccion.Enabled = false;
			this.txtDireccion.Location = new System.Drawing.Point(168, 80);
			this.txtDireccion.Name = "txtDireccion";
			this.txtDireccion.Size = new System.Drawing.Size(328, 22);
			this.txtDireccion.TabIndex = 4;
			this.txtDireccion.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtDireccion_KeyDown);
			this.txtDireccion.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtDireccion_KeyPress);
			// 
			// txtTelefono
			// 
			appearance17.BackColorDisabled = System.Drawing.Color.Silver;
			appearance17.BackColorDisabled2 = System.Drawing.Color.Silver;
			appearance17.BorderColor = System.Drawing.Color.Black;
			appearance17.BorderColor3DBase = System.Drawing.Color.Black;
			appearance17.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtTelefono.Appearance = appearance17;
			this.txtTelefono.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtTelefono.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtTelefono.Enabled = false;
			this.txtTelefono.Location = new System.Drawing.Point(697, 16);
			this.txtTelefono.MaxLength = 9;
			this.txtTelefono.Name = "txtTelefono";
			this.txtTelefono.Size = new System.Drawing.Size(168, 22);
			this.txtTelefono.TabIndex = 9;
			this.txtTelefono.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtTelefono_KeyDown);
			this.txtTelefono.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtTelefono_KeyPress);
			// 
			// txtFax
			// 
			appearance18.BackColorDisabled = System.Drawing.Color.Silver;
			appearance18.BackColorDisabled2 = System.Drawing.Color.Silver;
			appearance18.BorderColor = System.Drawing.Color.Black;
			appearance18.BorderColor3DBase = System.Drawing.Color.Black;
			appearance18.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtFax.Appearance = appearance18;
			this.txtFax.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtFax.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtFax.Enabled = false;
			this.txtFax.Location = new System.Drawing.Point(697, 56);
			this.txtFax.MaxLength = 9;
			this.txtFax.Name = "txtFax";
			this.txtFax.Size = new System.Drawing.Size(168, 22);
			this.txtFax.TabIndex = 10;
			this.txtFax.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtFax_KeyDown);
			this.txtFax.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtFax_KeyPress);
			// 
			// txtNotas
			// 
			appearance19.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtNotas.Appearance = appearance19;
			this.txtNotas.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtNotas.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNotas.Enabled = false;
			this.txtNotas.Location = new System.Drawing.Point(168, 256);
			this.txtNotas.Multiline = true;
			this.txtNotas.Name = "txtNotas";
			this.txtNotas.Size = new System.Drawing.Size(664, 72);
			this.txtNotas.TabIndex = 19;
			this.txtNotas.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtNotas_KeyDown);
			// 
			// txtidCliente
			// 
			appearance20.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtidCliente.Appearance = appearance20;
			this.txtidCliente.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtidCliente.Enabled = false;
			this.txtidCliente.FormatString = "0";
			this.txtidCliente.Location = new System.Drawing.Point(16, 312);
			this.txtidCliente.MinValue = 0;
			this.txtidCliente.Name = "txtidCliente";
			this.txtidCliente.PromptChar = ' ';
			this.txtidCliente.Size = new System.Drawing.Size(40, 22);
			this.txtidCliente.SpinButtonDisplayStyle = Infragistics.Win.ButtonDisplayStyle.Always;
			this.txtidCliente.TabIndex = 287;
			this.txtidCliente.Visible = false;
			// 
			// txtverificar
			// 
			appearance21.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtverificar.Appearance = appearance21;
			this.txtverificar.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtverificar.Enabled = false;
			this.txtverificar.FormatString = "0";
			this.txtverificar.Location = new System.Drawing.Point(520, 16);
			this.txtverificar.MinValue = 0;
			this.txtverificar.Name = "txtverificar";
			this.txtverificar.PromptChar = ' ';
			this.txtverificar.Size = new System.Drawing.Size(48, 22);
			this.txtverificar.SpinButtonDisplayStyle = Infragistics.Win.ButtonDisplayStyle.Always;
			this.txtverificar.TabIndex = 288;
			this.txtverificar.Visible = false;
			this.txtverificar.ValueChanged += new System.EventHandler(this.txtverificar_ValueChanged);
			// 
			// tpComplemento
			// 
			this.tpComplemento.Controls.Add(this.chkCourier);
			this.tpComplemento.Controls.Add(this.label1);
			this.tpComplemento.Controls.Add(this.txtDiasCredito4);
			this.tpComplemento.Controls.Add(this.txtDiasConsignacion);
			this.tpComplemento.Controls.Add(this.lblDiasdeConsignacion);
			this.tpComplemento.Controls.Add(this.chkCodigoBarras);
			this.tpComplemento.Controls.Add(this.chkAplicaConsignacion);
			this.tpComplemento.Controls.Add(this.optContadoCredito);
			this.tpComplemento.Controls.Add(this.cmbSector);
			this.tpComplemento.Controls.Add(this.label19);
			this.tpComplemento.Controls.Add(this.txtCodigo);
			this.tpComplemento.Controls.Add(this.label18);
			this.tpComplemento.Controls.Add(this.spnDescCredito);
			this.tpComplemento.Controls.Add(this.label24);
			this.tpComplemento.Controls.Add(this.spnDescuento);
			this.tpComplemento.Controls.Add(this.label23);
			this.tpComplemento.Controls.Add(this.txtDiasCredito3);
			this.tpComplemento.Controls.Add(this.label55);
			this.tpComplemento.Controls.Add(this.txtDiasCredito2);
			this.tpComplemento.Controls.Add(this.label54);
			this.tpComplemento.Controls.Add(this.txtValorCredito);
			this.tpComplemento.Controls.Add(this.txtDiasCredito1);
			this.tpComplemento.Controls.Add(this.label14);
			this.tpComplemento.Controls.Add(this.label12);
			this.tpComplemento.Location = new System.Drawing.Point(-10000, -10000);
			this.tpComplemento.Name = "tpComplemento";
			this.tpComplemento.Size = new System.Drawing.Size(892, 350);
			this.tpComplemento.Paint += new System.Windows.Forms.PaintEventHandler(this.tpComplemento_Paint);
			// 
			// chkCourier
			// 
			appearance22.ForeColorDisabled = System.Drawing.Color.Black;
			this.chkCourier.Appearance = appearance22;
			this.chkCourier.BackColor = System.Drawing.Color.Transparent;
			this.chkCourier.Enabled = false;
			this.chkCourier.Location = new System.Drawing.Point(688, 280);
			this.chkCourier.Name = "chkCourier";
			this.chkCourier.Size = new System.Drawing.Size(64, 21);
			this.chkCourier.TabIndex = 618;
			this.chkCourier.Text = "Courier";
			this.chkCourier.Visible = false;
			this.chkCourier.KeyDown += new System.Windows.Forms.KeyEventHandler(this.chkCourier_KeyDown);
			// 
			// label1
			// 
			this.label1.AutoSize = true;
			this.label1.BackColor = System.Drawing.Color.Transparent;
			this.label1.Location = new System.Drawing.Point(112, 216);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(78, 17);
			this.label1.TabIndex = 617;
			this.label1.Text = "Días Crédito 4:";
			// 
			// txtDiasCredito4
			// 
			appearance23.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtDiasCredito4.Appearance = appearance23;
			this.txtDiasCredito4.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtDiasCredito4.Enabled = false;
			this.txtDiasCredito4.Location = new System.Drawing.Point(224, 214);
			this.txtDiasCredito4.MinValue = 0;
			this.txtDiasCredito4.Name = "txtDiasCredito4";
			this.txtDiasCredito4.PromptChar = ' ';
			this.txtDiasCredito4.Size = new System.Drawing.Size(136, 22);
			this.txtDiasCredito4.SpinButtonDisplayStyle = Infragistics.Win.ButtonDisplayStyle.Always;
			this.txtDiasCredito4.TabIndex = 616;
			this.txtDiasCredito4.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtDiasCredito4_KeyDown);
			// 
			// txtDiasConsignacion
			// 
			appearance24.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtDiasConsignacion.Appearance = appearance24;
			this.txtDiasConsignacion.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtDiasConsignacion.Enabled = false;
			this.txtDiasConsignacion.FormatString = "0";
			this.txtDiasConsignacion.Location = new System.Drawing.Point(528, 279);
			this.txtDiasConsignacion.MinValue = 0;
			this.txtDiasConsignacion.Name = "txtDiasConsignacion";
			this.txtDiasConsignacion.PromptChar = ' ';
			this.txtDiasConsignacion.Size = new System.Drawing.Size(136, 22);
			this.txtDiasConsignacion.TabIndex = 615;
			this.txtDiasConsignacion.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtDiasConsignacion_KeyDown_1);
			this.txtDiasConsignacion.ValueChanged += new System.EventHandler(this.txtDiasConsignacion_ValueChanged);
			// 
			// lblDiasdeConsignacion
			// 
			this.lblDiasdeConsignacion.AutoSize = true;
			this.lblDiasdeConsignacion.BackColor = System.Drawing.Color.Transparent;
			this.lblDiasdeConsignacion.Location = new System.Drawing.Point(400, 282);
			this.lblDiasdeConsignacion.Name = "lblDiasdeConsignacion";
			this.lblDiasdeConsignacion.Size = new System.Drawing.Size(114, 17);
			this.lblDiasdeConsignacion.TabIndex = 613;
			this.lblDiasdeConsignacion.Text = "Dias de Consignación:";
			this.lblDiasdeConsignacion.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// chkCodigoBarras
			// 
			appearance25.ForeColorDisabled = System.Drawing.Color.Black;
			this.chkCodigoBarras.Appearance = appearance25;
			this.chkCodigoBarras.BackColor = System.Drawing.Color.Transparent;
			this.chkCodigoBarras.Enabled = false;
			this.chkCodigoBarras.Location = new System.Drawing.Point(112, 280);
			this.chkCodigoBarras.Name = "chkCodigoBarras";
			this.chkCodigoBarras.Size = new System.Drawing.Size(127, 21);
			this.chkCodigoBarras.TabIndex = 614;
			this.chkCodigoBarras.Text = "Codigo de Barras";
			this.chkCodigoBarras.Visible = false;
			this.chkCodigoBarras.KeyDown += new System.Windows.Forms.KeyEventHandler(this.chkCodigoBarras_KeyDown_1);
			// 
			// chkAplicaConsignacion
			// 
			appearance26.ForeColorDisabled = System.Drawing.Color.Black;
			this.chkAplicaConsignacion.Appearance = appearance26;
			this.chkAplicaConsignacion.BackColor = System.Drawing.Color.Transparent;
			this.chkAplicaConsignacion.Enabled = false;
			this.chkAplicaConsignacion.Location = new System.Drawing.Point(248, 280);
			this.chkAplicaConsignacion.Name = "chkAplicaConsignacion";
			this.chkAplicaConsignacion.Size = new System.Drawing.Size(127, 21);
			this.chkAplicaConsignacion.TabIndex = 612;
			this.chkAplicaConsignacion.Text = "Aplica Consignación";
			this.chkAplicaConsignacion.KeyDown += new System.Windows.Forms.KeyEventHandler(this.chkAplicaConsignacion_KeyDown_1);
			this.chkAplicaConsignacion.Validating += new System.ComponentModel.CancelEventHandler(this.chkAplicaConsignacion_Validating);
			this.chkAplicaConsignacion.CheckedChanged += new System.EventHandler(this.chkAplicaConsignacion_CheckedChanged_1);
			// 
			// optContadoCredito
			// 
			appearance27.ForeColorDisabled = System.Drawing.Color.Black;
			this.optContadoCredito.Appearance = appearance27;
			this.optContadoCredito.Enabled = false;
			this.optContadoCredito.ItemAppearance = appearance28;
			this.optContadoCredito.ItemOrigin = new System.Drawing.Point(10, 0);
			valueListItem1.DataValue = 1;
			valueListItem1.DisplayText = "Contado";
			valueListItem2.DataValue = 2;
			valueListItem2.DisplayText = "Credito";
			this.optContadoCredito.Items.Add(valueListItem1);
			this.optContadoCredito.Items.Add(valueListItem2);
			this.optContadoCredito.ItemSpacingVertical = 10;
			this.optContadoCredito.Location = new System.Drawing.Point(104, 8);
			this.optContadoCredito.Name = "optContadoCredito";
			this.optContadoCredito.Size = new System.Drawing.Size(200, 26);
			this.optContadoCredito.TabIndex = 611;
			this.optContadoCredito.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.optContadoCredito_KeyPress);
			this.optContadoCredito.KeyDown += new System.Windows.Forms.KeyEventHandler(this.optContadoCredito_KeyDown);
			this.optContadoCredito.ValueChanged += new System.EventHandler(this.optContadoCredito_ValueChanged);
			// 
			// cmbSector
			// 
			appearance29.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbSector.Appearance = appearance29;
			this.cmbSector.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbSector.Cursor = System.Windows.Forms.Cursors.Default;
			ultraGridColumn16.Header.VisiblePosition = 0;
			ultraGridColumn16.Hidden = true;
			ultraGridColumn17.Header.VisiblePosition = 1;
			ultraGridColumn17.Width = 274;
			ultraGridBand8.Columns.AddRange(new object[] {
																										 ultraGridColumn16,
																										 ultraGridColumn17});
			this.cmbSector.DisplayLayout.BandsSerializer.Add(ultraGridBand8);
			this.cmbSector.DisplayMember = "Sector";
			this.cmbSector.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbSector.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbSector.Enabled = false;
			this.cmbSector.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbSector.Location = new System.Drawing.Point(616, 136);
			this.cmbSector.Name = "cmbSector";
			this.cmbSector.Size = new System.Drawing.Size(136, 21);
			this.cmbSector.TabIndex = 177;
			this.cmbSector.ValueMember = "idSector";
			this.cmbSector.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cmbSector_KeyDown);
			// 
			// label19
			// 
			this.label19.AutoSize = true;
			this.label19.BackColor = System.Drawing.Color.Transparent;
			this.label19.Location = new System.Drawing.Point(520, 138);
			this.label19.Name = "label19";
			this.label19.Size = new System.Drawing.Size(40, 17);
			this.label19.TabIndex = 183;
			this.label19.Text = "Sector:";
			// 
			// txtCodigo
			// 
			appearance30.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtCodigo.Appearance = appearance30;
			this.txtCodigo.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtCodigo.Location = new System.Drawing.Point(616, 176);
			this.txtCodigo.Name = "txtCodigo";
			this.txtCodigo.Size = new System.Drawing.Size(136, 22);
			this.txtCodigo.TabIndex = 179;
			this.txtCodigo.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtCodigo_KeyDown);
			this.txtCodigo.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtCodigo_KeyPress);
			// 
			// label18
			// 
			this.label18.AutoSize = true;
			this.label18.BackColor = System.Drawing.Color.Transparent;
			this.label18.Location = new System.Drawing.Point(520, 178);
			this.label18.Name = "label18";
			this.label18.Size = new System.Drawing.Size(42, 17);
			this.label18.TabIndex = 182;
			this.label18.Text = "Código:";
			// 
			// spnDescCredito
			// 
			appearance31.ForeColorDisabled = System.Drawing.Color.Black;
			this.spnDescCredito.Appearance = appearance31;
			this.spnDescCredito.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.spnDescCredito.Enabled = false;
			this.spnDescCredito.Location = new System.Drawing.Point(616, 94);
			this.spnDescCredito.MinValue = 0;
			this.spnDescCredito.Name = "spnDescCredito";
			this.spnDescCredito.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.spnDescCredito.PromptChar = ' ';
			this.spnDescCredito.Size = new System.Drawing.Size(136, 22);
			this.spnDescCredito.SpinButtonDisplayStyle = Infragistics.Win.ButtonDisplayStyle.Always;
			this.spnDescCredito.TabIndex = 176;
			this.spnDescCredito.KeyDown += new System.Windows.Forms.KeyEventHandler(this.spnDescCredito_KeyDown);
			// 
			// label24
			// 
			this.label24.AutoSize = true;
			this.label24.BackColor = System.Drawing.Color.Transparent;
			this.label24.Location = new System.Drawing.Point(504, 96);
			this.label24.Name = "label24";
			this.label24.Size = new System.Drawing.Size(75, 17);
			this.label24.TabIndex = 181;
			this.label24.Text = "Desc. Crédito:";
			// 
			// spnDescuento
			// 
			appearance32.ForeColorDisabled = System.Drawing.Color.Black;
			this.spnDescuento.Appearance = appearance32;
			this.spnDescuento.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.spnDescuento.Enabled = false;
			this.spnDescuento.Location = new System.Drawing.Point(616, 54);
			this.spnDescuento.MinValue = 0;
			this.spnDescuento.Name = "spnDescuento";
			this.spnDescuento.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.spnDescuento.PromptChar = ' ';
			this.spnDescuento.Size = new System.Drawing.Size(136, 22);
			this.spnDescuento.SpinButtonDisplayStyle = Infragistics.Win.ButtonDisplayStyle.Always;
			this.spnDescuento.TabIndex = 175;
			this.spnDescuento.KeyDown += new System.Windows.Forms.KeyEventHandler(this.spnDescuento_KeyDown);
			// 
			// label23
			// 
			this.label23.AutoSize = true;
			this.label23.BackColor = System.Drawing.Color.Transparent;
			this.label23.Location = new System.Drawing.Point(504, 56);
			this.label23.Name = "label23";
			this.label23.Size = new System.Drawing.Size(80, 17);
			this.label23.TabIndex = 180;
			this.label23.Text = "Desc. Contado:";
			// 
			// txtDiasCredito3
			// 
			appearance33.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtDiasCredito3.Appearance = appearance33;
			this.txtDiasCredito3.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtDiasCredito3.Enabled = false;
			this.txtDiasCredito3.Location = new System.Drawing.Point(224, 176);
			this.txtDiasCredito3.MinValue = 0;
			this.txtDiasCredito3.Name = "txtDiasCredito3";
			this.txtDiasCredito3.PromptChar = ' ';
			this.txtDiasCredito3.Size = new System.Drawing.Size(136, 22);
			this.txtDiasCredito3.SpinButtonDisplayStyle = Infragistics.Win.ButtonDisplayStyle.Always;
			this.txtDiasCredito3.TabIndex = 26;
			this.txtDiasCredito3.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtDiasCredito3_KeyDown);
			// 
			// label55
			// 
			this.label55.AutoSize = true;
			this.label55.BackColor = System.Drawing.Color.Transparent;
			this.label55.Location = new System.Drawing.Point(112, 178);
			this.label55.Name = "label55";
			this.label55.Size = new System.Drawing.Size(78, 17);
			this.label55.TabIndex = 164;
			this.label55.Text = "Días Crédito 3:";
			// 
			// txtDiasCredito2
			// 
			appearance34.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtDiasCredito2.Appearance = appearance34;
			this.txtDiasCredito2.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtDiasCredito2.Enabled = false;
			this.txtDiasCredito2.Location = new System.Drawing.Point(224, 136);
			this.txtDiasCredito2.MinValue = 0;
			this.txtDiasCredito2.Name = "txtDiasCredito2";
			this.txtDiasCredito2.PromptChar = ' ';
			this.txtDiasCredito2.Size = new System.Drawing.Size(136, 22);
			this.txtDiasCredito2.SpinButtonDisplayStyle = Infragistics.Win.ButtonDisplayStyle.Always;
			this.txtDiasCredito2.TabIndex = 25;
			this.txtDiasCredito2.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtDiasCredito2_KeyDown);
			// 
			// label54
			// 
			this.label54.AutoSize = true;
			this.label54.BackColor = System.Drawing.Color.Transparent;
			this.label54.Location = new System.Drawing.Point(112, 138);
			this.label54.Name = "label54";
			this.label54.Size = new System.Drawing.Size(78, 17);
			this.label54.TabIndex = 162;
			this.label54.Text = "Días Crédito 2:";
			// 
			// txtValorCredito
			// 
			appearance35.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtValorCredito.Appearance = appearance35;
			this.txtValorCredito.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtValorCredito.Enabled = false;
			this.txtValorCredito.FormatString = "#,##0.00";
			this.txtValorCredito.Location = new System.Drawing.Point(224, 54);
			this.txtValorCredito.MinValue = 0;
			this.txtValorCredito.Name = "txtValorCredito";
			this.txtValorCredito.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtValorCredito.PromptChar = ' ';
			this.txtValorCredito.Size = new System.Drawing.Size(136, 22);
			this.txtValorCredito.TabIndex = 23;
			this.txtValorCredito.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtValorCredito_KeyDown);
			this.txtValorCredito.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtValorCredito_KeyPress);
			// 
			// txtDiasCredito1
			// 
			appearance36.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtDiasCredito1.Appearance = appearance36;
			this.txtDiasCredito1.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtDiasCredito1.Enabled = false;
			this.txtDiasCredito1.Location = new System.Drawing.Point(224, 94);
			this.txtDiasCredito1.MinValue = 0;
			this.txtDiasCredito1.Name = "txtDiasCredito1";
			this.txtDiasCredito1.PromptChar = ' ';
			this.txtDiasCredito1.Size = new System.Drawing.Size(136, 22);
			this.txtDiasCredito1.SpinButtonDisplayStyle = Infragistics.Win.ButtonDisplayStyle.Always;
			this.txtDiasCredito1.TabIndex = 24;
			this.txtDiasCredito1.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtDiasCredito1_KeyDown);
			// 
			// label14
			// 
			this.label14.AutoSize = true;
			this.label14.BackColor = System.Drawing.Color.Transparent;
			this.label14.Location = new System.Drawing.Point(106, 56);
			this.label14.Name = "label14";
			this.label14.Size = new System.Drawing.Size(89, 17);
			this.label14.TabIndex = 109;
			this.label14.Text = "Cupo de Crédito:";
			// 
			// label12
			// 
			this.label12.AutoSize = true;
			this.label12.BackColor = System.Drawing.Color.Transparent;
			this.label12.Location = new System.Drawing.Point(112, 96);
			this.label12.Name = "label12";
			this.label12.Size = new System.Drawing.Size(78, 17);
			this.label12.TabIndex = 108;
			this.label12.Text = "Días Crédito 1:";
			// 
			// ultraTabPageControl5
			// 
			this.ultraTabPageControl5.Controls.Add(this.ultgContactos);
			this.ultraTabPageControl5.Location = new System.Drawing.Point(-10000, -10000);
			this.ultraTabPageControl5.Name = "ultraTabPageControl5";
			this.ultraTabPageControl5.Size = new System.Drawing.Size(892, 350);
			this.ultraTabPageControl5.Paint += new System.Windows.Forms.PaintEventHandler(this.ultraTabPageControl5_Paint);
			// 
			// ultgContactos
			// 
			this.ultgContactos.Cursor = System.Windows.Forms.Cursors.Default;
			this.ultgContactos.DataSource = this.ultraDataSource2;
			appearance37.BackColor = System.Drawing.Color.White;
			this.ultgContactos.DisplayLayout.Appearance = appearance37;
			this.ultgContactos.DisplayLayout.AutoFitColumns = true;
			ultraGridBand9.AddButtonCaption = "ClienteTipoRuc";
			ultraGridColumn18.Header.VisiblePosition = 9;
			ultraGridColumn18.Hidden = true;
			ultraGridColumn18.Width = 96;
			ultraGridColumn19.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn19.Header.VisiblePosition = 1;
			ultraGridColumn19.Width = 170;
			ultraGridColumn20.FieldLen = 9;
			ultraGridColumn20.Format = "";
			ultraGridColumn20.Header.VisiblePosition = 3;
			ultraGridColumn20.Width = 82;
			ultraGridColumn21.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn21.Header.Caption = "Observación";
			ultraGridColumn21.Header.VisiblePosition = 5;
			ultraGridColumn21.Width = 145;
			ultraGridColumn22.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn22.Header.Caption = "E-Mail";
			ultraGridColumn22.Header.VisiblePosition = 7;
			ultraGridColumn22.Width = 71;
			ultraGridColumn23.FieldLen = 10;
			ultraGridColumn23.Format = "";
			ultraGridColumn23.Header.VisiblePosition = 4;
			ultraGridColumn23.MaskInput = "";
			ultraGridColumn23.Width = 92;
			ultraGridColumn24.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn24.Header.VisiblePosition = 2;
			ultraGridColumn24.Width = 69;
			ultraGridColumn25.Header.Caption = "Envía Mail";
			ultraGridColumn25.Header.VisiblePosition = 6;
			ultraGridColumn25.Width = 60;
			ultraGridColumn26.Header.VisiblePosition = 8;
			ultraGridColumn26.Width = 122;
			ultraGridColumn27.Header.VisiblePosition = 0;
			ultraGridColumn27.Hidden = true;
			ultraGridColumn27.Width = 94;
			ultraGridBand9.Columns.AddRange(new object[] {
																										 ultraGridColumn18,
																										 ultraGridColumn19,
																										 ultraGridColumn20,
																										 ultraGridColumn21,
																										 ultraGridColumn22,
																										 ultraGridColumn23,
																										 ultraGridColumn24,
																										 ultraGridColumn25,
																										 ultraGridColumn26,
																										 ultraGridColumn27});
			this.ultgContactos.DisplayLayout.BandsSerializer.Add(ultraGridBand9);
			this.ultgContactos.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.TemplateOnBottom;
			this.ultgContactos.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.True;
			this.ultgContactos.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance38.BackColor = System.Drawing.Color.Transparent;
			this.ultgContactos.DisplayLayout.Override.CardAreaAppearance = appearance38;
			appearance39.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance39.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance39.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance39.FontData.BoldAsString = "True";
			appearance39.FontData.Name = "Arial";
			appearance39.FontData.SizeInPoints = 10F;
			appearance39.ForeColor = System.Drawing.Color.White;
			appearance39.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.ultgContactos.DisplayLayout.Override.HeaderAppearance = appearance39;
			appearance40.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance40.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance40.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultgContactos.DisplayLayout.Override.RowSelectorAppearance = appearance40;
			appearance41.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance41.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance41.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultgContactos.DisplayLayout.Override.SelectedRowAppearance = appearance41;
			this.ultgContactos.Enabled = false;
			this.ultgContactos.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.ultgContactos.Location = new System.Drawing.Point(24, 16);
			this.ultgContactos.Name = "ultgContactos";
			this.ultgContactos.Size = new System.Drawing.Size(832, 296);
			this.ultgContactos.TabIndex = 3;
			this.ultgContactos.InitializeRow += new Infragistics.Win.UltraWinGrid.InitializeRowEventHandler(this.ultgContactos_InitializeRow);
			this.ultgContactos.CellChange += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.ultgContactos_CellChange_1);
			this.ultgContactos.AfterRowInsert += new Infragistics.Win.UltraWinGrid.RowEventHandler(this.ultgContactos_AfterRowInsert);
			this.ultgContactos.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.ultgContactos_KeyPress);
			this.ultgContactos.BeforeRowInsert += new Infragistics.Win.UltraWinGrid.BeforeRowInsertEventHandler(this.ultgContactos_BeforeRowInsert);
			this.ultgContactos.KeyDown += new System.Windows.Forms.KeyEventHandler(this.ultgContactos_KeyDown);
			this.ultgContactos.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.ultraGrid2_InitializeLayout_1);
			// 
			// ultraDataSource2
			// 
			ultraDataColumn1.DataType = typeof(int);
			ultraDataColumn8.DataType = typeof(bool);
			ultraDataColumn8.DefaultValue = false;
			ultraDataColumn9.DataType = typeof(bool);
			ultraDataColumn9.DefaultValue = false;
			ultraDataColumn10.DataType = typeof(int);
			this.ultraDataSource2.Band.Columns.AddRange(new object[] {
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
			this.ultraDataSource2.Band.Key = "ClienteContacto";
			// 
			// ultraTabPageControl7
			// 
			this.ultraTabPageControl7.Controls.Add(this.ultDirecciones);
			this.ultraTabPageControl7.Location = new System.Drawing.Point(-10000, -10000);
			this.ultraTabPageControl7.Name = "ultraTabPageControl7";
			this.ultraTabPageControl7.Size = new System.Drawing.Size(892, 350);
			// 
			// ultDirecciones
			// 
			this.ultDirecciones.Cursor = System.Windows.Forms.Cursors.Default;
			this.ultDirecciones.DataSource = this.ultraDataSource3;
			appearance42.BackColor = System.Drawing.Color.White;
			this.ultDirecciones.DisplayLayout.Appearance = appearance42;
			this.ultDirecciones.DisplayLayout.AutoFitColumns = true;
			ultraGridBand10.AddButtonCaption = "ClientesDirecciones";
			ultraGridColumn28.Header.VisiblePosition = 0;
			ultraGridColumn28.Hidden = true;
			ultraGridColumn28.Width = 139;
			ultraGridColumn29.Header.VisiblePosition = 1;
			ultraGridColumn29.Hidden = true;
			ultraGridColumn29.Width = 33;
			ultraGridColumn30.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn30.Header.VisiblePosition = 2;
			ultraGridColumn30.Width = 474;
			ultraGridBand10.Columns.AddRange(new object[] {
																											ultraGridColumn28,
																											ultraGridColumn29,
																											ultraGridColumn30});
			this.ultDirecciones.DisplayLayout.BandsSerializer.Add(ultraGridBand10);
			this.ultDirecciones.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.TemplateOnBottom;
			this.ultDirecciones.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.True;
			this.ultDirecciones.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance43.BackColor = System.Drawing.Color.Transparent;
			this.ultDirecciones.DisplayLayout.Override.CardAreaAppearance = appearance43;
			appearance44.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance44.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance44.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance44.FontData.BoldAsString = "True";
			appearance44.FontData.Name = "Arial";
			appearance44.FontData.SizeInPoints = 10F;
			appearance44.ForeColor = System.Drawing.Color.White;
			appearance44.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.ultDirecciones.DisplayLayout.Override.HeaderAppearance = appearance44;
			appearance45.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance45.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance45.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultDirecciones.DisplayLayout.Override.RowSelectorAppearance = appearance45;
			appearance46.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance46.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance46.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultDirecciones.DisplayLayout.Override.SelectedRowAppearance = appearance46;
			this.ultDirecciones.Enabled = false;
			this.ultDirecciones.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.ultDirecciones.Location = new System.Drawing.Point(9, 6);
			this.ultDirecciones.Name = "ultDirecciones";
			this.ultDirecciones.Size = new System.Drawing.Size(495, 218);
			this.ultDirecciones.TabIndex = 2;
			this.ultDirecciones.AfterRowInsert += new Infragistics.Win.UltraWinGrid.RowEventHandler(this.ultDirecciones_AfterRowInsert);
			this.ultDirecciones.KeyDown += new System.Windows.Forms.KeyEventHandler(this.ultDirecciones_KeyDown);
			// 
			// ultraDataSource3
			// 
			ultraDataColumn11.DataType = typeof(int);
			ultraDataColumn12.DataType = typeof(int);
			this.ultraDataSource3.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn11,
																																 ultraDataColumn12,
																																 ultraDataColumn13});
			// 
			// ultraDataSource1
			// 
			ultraDataColumn14.DataType = typeof(int);
			ultraDataColumn15.DataType = typeof(int);
			ultraDataColumn21.DataType = typeof(bool);
			ultraDataColumn23.DataType = typeof(bool);
			this.ultraDataSource1.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn14,
																																 ultraDataColumn15,
																																 ultraDataColumn16,
																																 ultraDataColumn17,
																																 ultraDataColumn18,
																																 ultraDataColumn19,
																																 ultraDataColumn20,
																																 ultraDataColumn21,
																																 ultraDataColumn22,
																																 ultraDataColumn23});
			// 
			// tbBasico
			// 
			appearance47.BackColor = System.Drawing.Color.LightSteelBlue;
			appearance47.BackColor2 = System.Drawing.Color.LightSteelBlue;
			this.tbBasico.Appearance = appearance47;
			this.tbBasico.BackColor = System.Drawing.Color.LightSteelBlue;
			appearance48.BackColor = System.Drawing.Color.LightSteelBlue;
			appearance48.BackColor2 = System.Drawing.Color.LightSteelBlue;
			this.tbBasico.ClientAreaAppearance = appearance48;
			this.tbBasico.Controls.Add(this.ultraTabSharedControlsPage1);
			this.tbBasico.Controls.Add(this.ultraTabPageControl1);
			this.tbBasico.Controls.Add(this.tpComplemento);
			this.tbBasico.Controls.Add(this.ultraTabPageControl5);
			this.tbBasico.Controls.Add(this.ultraTabPageControl7);
			this.tbBasico.Location = new System.Drawing.Point(0, 0);
			this.tbBasico.Name = "tbBasico";
			this.tbBasico.SharedControlsPage = this.ultraTabSharedControlsPage1;
			this.tbBasico.Size = new System.Drawing.Size(896, 376);
			this.tbBasico.TabIndex = 48;
			appearance49.BackColor = System.Drawing.Color.LightSteelBlue;
			appearance49.BackColor2 = System.Drawing.Color.LightSteelBlue;
			ultraTab1.ActiveAppearance = appearance49;
			appearance50.BackColor = System.Drawing.Color.LightSteelBlue;
			appearance50.BackColor2 = System.Drawing.Color.LightSteelBlue;
			ultraTab1.Appearance = appearance50;
			appearance51.BackColor = System.Drawing.Color.LightSteelBlue;
			appearance51.BackColor2 = System.Drawing.Color.LightSteelBlue;
			ultraTab1.ClientAreaAppearance = appearance51;
			appearance52.BackColor = System.Drawing.Color.LightSteelBlue;
			appearance52.BackColor2 = System.Drawing.Color.LightSteelBlue;
			ultraTab1.SelectedAppearance = appearance52;
			ultraTab1.TabPage = this.ultraTabPageControl1;
			ultraTab1.Text = "Básico";
			ultraTab2.TabPage = this.tpComplemento;
			ultraTab2.Text = "Complemento";
			ultraTab3.TabPage = this.ultraTabPageControl5;
			ultraTab3.Text = "Contactos";
			ultraTab4.TabPage = this.ultraTabPageControl7;
			ultraTab4.Text = "Direcciones de entrega";
			this.tbBasico.Tabs.AddRange(new Infragistics.Win.UltraWinTabControl.UltraTab[] {
																																											 ultraTab1,
																																											 ultraTab2,
																																											 ultraTab3,
																																											 ultraTab4});
			this.tbBasico.SelectedTabChanged += new Infragistics.Win.UltraWinTabControl.SelectedTabChangedEventHandler(this.tbBasico_SelectedTabChanged);
			// 
			// ultraTabSharedControlsPage1
			// 
			this.ultraTabSharedControlsPage1.Location = new System.Drawing.Point(-10000, -10000);
			this.ultraTabSharedControlsPage1.Name = "ultraTabSharedControlsPage1";
			this.ultraTabSharedControlsPage1.Size = new System.Drawing.Size(892, 350);
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
			// btnEditar
			// 
			this.btnEditar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnEditar.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(230)), ((System.Byte)(230)), ((System.Byte)(230)));
			this.btnEditar.CausesValidation = false;
			this.btnEditar.Enabled = false;
			this.btnEditar.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.btnEditar.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(50)), ((System.Byte)(50)), ((System.Byte)(50)));
			this.btnEditar.Image = ((System.Drawing.Image)(resources.GetObject("btnEditar.Image")));
			this.btnEditar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnEditar.Location = new System.Drawing.Point(232, 378);
			this.btnEditar.Name = "btnEditar";
			this.btnEditar.Size = new System.Drawing.Size(72, 23);
			this.btnEditar.TabIndex = 282;
			this.btnEditar.Text = "Editar";
			this.btnEditar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnEditar.Click += new System.EventHandler(this.btnEditar_Click);
			// 
			// btnCancelar
			// 
			this.btnCancelar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnCancelar.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(230)), ((System.Byte)(230)), ((System.Byte)(230)));
			this.btnCancelar.CausesValidation = false;
			this.btnCancelar.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.btnCancelar.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(50)), ((System.Byte)(50)), ((System.Byte)(50)));
			this.btnCancelar.Image = ((System.Drawing.Image)(resources.GetObject("btnCancelar.Image")));
			this.btnCancelar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnCancelar.Location = new System.Drawing.Point(416, 378);
			this.btnCancelar.Name = "btnCancelar";
			this.btnCancelar.Size = new System.Drawing.Size(72, 23);
			this.btnCancelar.TabIndex = 283;
			this.btnCancelar.Text = "Cancelar";
			this.btnCancelar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnCancelar.Click += new System.EventHandler(this.btnCancelar_Click);
			// 
			// btnSalir
			// 
			this.btnSalir.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnSalir.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(230)), ((System.Byte)(230)), ((System.Byte)(230)));
			this.btnSalir.CausesValidation = false;
			this.btnSalir.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.btnSalir.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(50)), ((System.Byte)(50)), ((System.Byte)(50)));
			this.btnSalir.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnSalir.ImageIndex = 62;
			this.btnSalir.Location = new System.Drawing.Point(512, 378);
			this.btnSalir.Name = "btnSalir";
			this.btnSalir.Size = new System.Drawing.Size(72, 23);
			this.btnSalir.TabIndex = 284;
			this.btnSalir.Text = "Salir";
			this.btnSalir.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnSalir.Click += new System.EventHandler(this.btnSalir_Click);
			// 
			// btnNuevo
			// 
			this.btnNuevo.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnNuevo.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(230)), ((System.Byte)(230)), ((System.Byte)(230)));
			this.btnNuevo.CausesValidation = false;
			this.btnNuevo.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.btnNuevo.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(50)), ((System.Byte)(50)), ((System.Byte)(50)));
			this.btnNuevo.Image = ((System.Drawing.Image)(resources.GetObject("btnNuevo.Image")));
			this.btnNuevo.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnNuevo.Location = new System.Drawing.Point(64, 378);
			this.btnNuevo.Name = "btnNuevo";
			this.btnNuevo.Size = new System.Drawing.Size(72, 23);
			this.btnNuevo.TabIndex = 279;
			this.btnNuevo.Text = "Nuevo";
			this.btnNuevo.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnNuevo.Click += new System.EventHandler(this.btnNuevo_Click);
			// 
			// btnConsultar
			// 
			this.btnConsultar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnConsultar.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(230)), ((System.Byte)(230)), ((System.Byte)(230)));
			this.btnConsultar.CausesValidation = false;
			this.btnConsultar.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.btnConsultar.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(50)), ((System.Byte)(50)), ((System.Byte)(50)));
			this.btnConsultar.Image = ((System.Drawing.Image)(resources.GetObject("btnConsultar.Image")));
			this.btnConsultar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnConsultar.Location = new System.Drawing.Point(144, 378);
			this.btnConsultar.Name = "btnConsultar";
			this.btnConsultar.Size = new System.Drawing.Size(72, 23);
			this.btnConsultar.TabIndex = 280;
			this.btnConsultar.Text = "Consultar";
			this.btnConsultar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnConsultar.Click += new System.EventHandler(this.btnConsultar_Click);
			// 
			// btnGuardarComoCliente
			// 
			this.btnGuardarComoCliente.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnGuardarComoCliente.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(230)), ((System.Byte)(230)), ((System.Byte)(230)));
			this.btnGuardarComoCliente.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.btnGuardarComoCliente.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(50)), ((System.Byte)(50)), ((System.Byte)(50)));
			this.btnGuardarComoCliente.Image = ((System.Drawing.Image)(resources.GetObject("btnGuardarComoCliente.Image")));
			this.btnGuardarComoCliente.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnGuardarComoCliente.Location = new System.Drawing.Point(320, 378);
			this.btnGuardarComoCliente.Name = "btnGuardarComoCliente";
			this.btnGuardarComoCliente.Size = new System.Drawing.Size(80, 23);
			this.btnGuardarComoCliente.TabIndex = 285;
			this.btnGuardarComoCliente.Text = "Guardar";
			this.btnGuardarComoCliente.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnGuardarComoCliente.Click += new System.EventHandler(this.button1_Click_1);
			// 
			// frmClientesProveedores
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 14);
			this.ClientSize = new System.Drawing.Size(898, 408);
			this.Controls.Add(this.btnGuardarComoCliente);
			this.Controls.Add(this.btnEditar);
			this.Controls.Add(this.btnCancelar);
			this.Controls.Add(this.btnSalir);
			this.Controls.Add(this.btnNuevo);
			this.Controls.Add(this.btnConsultar);
			this.Controls.Add(this.tbBasico);
			this.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
			this.KeyPreview = true;
			this.MaximizeBox = false;
			this.Name = "frmClientesProveedores";
			this.ShowInTaskbar = false;
			this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
			this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.frmClientesProveedores_KeyDown);
			this.Closing += new System.ComponentModel.CancelEventHandler(this.frmClientesProveedores_Closing);
			this.Load += new System.EventHandler(this.frmClientesProveedores_Load);
			this.ultraTabPageControl1.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.txtAgenteRetencion)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbPais)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbSujetoRetencion)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbAgenteDeRetencion)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbTipoProveedor)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbTipoRuc)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtRuc)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNombreComercial)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCelular)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtEMail)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtRepresentante)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCiudad)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbGrupo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNombre)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDireccion)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtTelefono)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtFax)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNotas)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtidCliente)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtverificar)).EndInit();
			this.tpComplemento.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.txtDiasCredito4)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDiasConsignacion)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.optContadoCredito)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbSector)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCodigo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.spnDescCredito)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.spnDescuento)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDiasCredito3)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDiasCredito2)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtValorCredito)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDiasCredito1)).EndInit();
			this.ultraTabPageControl5.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.ultgContactos)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).EndInit();
			this.ultraTabPageControl7.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.ultDirecciones)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource3)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.tbBasico)).EndInit();
			this.tbBasico.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		#region Variables Globales

		private bool TipoCliente = true;
		private int IdCliente = 0;
		private string stClientes = "Clientes";
		private string stCliente = "Cliente: ";
		string stFiltro = "";
		private int idTipoF = 34;
		private bool bCodAuto = false;
		private C1.Data.C1DataRow drCliente;
		private C1.Data.C1DataRow drSeteoF;
		private Acceso miAcceso;
		bool bcarga;
		//private Compra miCompra;
		#endregion Variables Globales

		private void frmClientesProveedores_Load(object sender, System.EventArgs e)
		{
			bcarga=true;
			int iCliente = 0;
			if (!bClienteProveedor) iCliente = 1; 

			cmbTipoRuc.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Exec ClienteTipoRucCarga " + iCliente);

			cmbGrupo.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Select idGrupoCliente, Grupo From ClienteGrupo Where Proveedor = '" + iCliente + "' Order by Grupo");

			//			int  bClienetProveedor = 0;
			//			this.txtverificar.Value = Convert.ToInt32(bClienteProveedor);
			//			int iCliente = 0;
			//			if (!bClienteProveedor) iCliente = 1; 
			//				cmbTipoRuc.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Exec ClienteTipoRucCarga " + iCliente);
			//				cmbGrupo.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Select idGrupoCliente, Grupo From ClienteGrupo Where Proveedor = '" + iCliente + "' Order by Grupo");
			#region Elementos de la pantalla cliente
			if(!bClienteProveedor)
			{			  
				this.lblNotas.Visible= true;
				this.lblproductos.Visible=false;
				this.lblTipoProveedor.Visible = false;
				this.cmbTipoProveedor.Visible = false;
				this.chkCodigoBarras.Visible=false;
				this.chkRise.Visible=false;
				this.lblSujetoDeRetencion.Visible=false;
				this.cmbSujetoRetencion.Visible=false;
				this.cmbAgenteDeRetencion.Visible=true;
				this.lblAGRetencion.Visible=true;
				this.chkAplicaConsignacion.Visible= true;
				this.lblDiasdeConsignacion.Visible=true;
				this.txtDiasConsignacion.Visible= true;
				this.chkCourier.Visible = false;
				this.txtverificar.Value=0;				
				//				this.chkCodigoBarras.Visible=false;
				this.Text = "CLIENTE";
			}
			#endregion Elementos de la pantalla cliente
			#region elementos de la pantalla proveedor
			if(bClienteProveedor)
			{
				//				int  bClienetProveedor = 1;
				//				
				//				int iCliente = 1;
				//				cmbTipoRuc.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Exec ClienteTipoRucCarga " + iCliente);
				//				this.txtverificar.Value = bClienetProveedor;
				this.lblNotas.Visible= false;
				this.lblproductos.Visible=true;
		
				this.lblTipoProveedor.Visible = true;
				this.cmbTipoProveedor.Visible = true;
				this.chkCodigoBarras.Visible=true;
				this.lblSujetoDeRetencion.Visible=true;
				this.cmbSujetoRetencion.Visible=true;
				this.cmbAgenteDeRetencion.Visible=false;
				this.lblAGRetencion.Visible=false;
				this.chkRise.Visible=true;
				this.chkAplicaConsignacion.Visible= true;
				this.lblDiasdeConsignacion.Visible=true;
				this.txtDiasConsignacion.Visible= true;
				this.chkCourier.Visible = true;
				this.txtverificar.Value=1;
				
				//				this.chkCodigoBarras.Visible=true;
				this.Text = "PROVEEDOR";

			}
			#endregion elementos de la pantalla proveedor
			#region Cargar Datos a los ComboBox
			
			this.cmbAgenteDeRetencion.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Select idAgenteRetencion, Nombre From AgenteRetencion");
			cmbPais.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Select idClientePais, Pais From ClientePais Order by Pais");
			this.cmbTipoProveedor.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Select idTipoProveedor, Codigo, Tipo From TipoDeProveedor Order By idTipoProveedor Asc");
			cmbSector.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Select idSector, Sector From ClienteSector Order by Sector");
			this.cmbSujetoRetencion.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Select idSujetoRetencion, Nombre From SujetoRetencion");
			cmbPais.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Select idClientePais, Pais From ClientePais Order by Pais");
			cmbCiudad.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Select idCiudad, Ciudad From ClienteCiudad Order By Ciudad");
			#endregion Cargar Datos a los ComboBox

			#region Pantalla Deposito
		  
			if (idClienteC > 0) this.Consultar(idClienteC);

			if (iOrigen == 1)
			{
				if (idClienteC > 0) this.btnEditar_Click(sender, e);			
				else this.btnNuevo_Click(sender, e);
			}
			#endregion Pantalla Deposito

bcarga=false;
		}

		private void ultraTabPageControl1_Paint(object sender, System.Windows.Forms.PaintEventArgs e)
		{
		}

		private void lblNotas_Click(object sender, System.EventArgs e)
		{	}

		private void cmbTipoRuc_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter) this.txtRuc.Focus();
		}

		private void txtRuc_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{

			if(!bClienteProveedor ){if(e.KeyCode == Keys.Enter) {this.txtNombre.Focus();}}

			if(bClienteProveedor ){if(e.KeyCode == Keys.Enter) {this.chkRise.Focus();}}

		
		}

		private void txtNombre_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter) this.txtDireccion.Focus();
		}

		private void txtDireccion_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter) this.cmbGrupo.Focus();
		}

		private void cmbGrupo_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter) this.txtRepresentante.Focus();
		}

		private void txtRepresentante_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter) this.txtNombreComercial.Focus();
		}

		private void txtNombreComercial_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{

			if(!bClienteProveedor )if(e.KeyCode == Keys.Enter) {this.cmbAgenteDeRetencion.Focus();}
			if(bClienteProveedor )if(e.KeyCode == Keys.Enter) {this.cmbSujetoRetencion.Focus();}
			
		}
		private void cmbSujetoRetencion_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter) this.txtAgenteRetencion.Focus();
		}

		private void txtTelefono_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter) this.txtFax.Focus();
		}

		private void txtFax_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter) this.txtCelular.Focus();
		}

		private void txtCelular_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter) this.txtEMail.Focus();
		}

		private void txtEMail_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter) this.cmbPais.Focus();
		}

		private void cmbPais_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter) this.cmbCiudad.Focus();
		}

		private void cmbCiudad_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(!bClienteProveedor )if(e.KeyCode == Keys.Enter) {this.txtNotas.Focus();}
			if(bClienteProveedor )if(e.KeyCode == Keys.Enter) {this.cmbTipoProveedor.Focus();}
		}

		private void cmbTipoProveedor_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter) this.txtNotas.Focus();
		}

		private void txtNotas_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			// ahi es cliente
			if(e.KeyCode == Keys.Enter)
			{
				this.tbBasico.Tabs[1].Selected = true;
			  this.optContadoCredito.Focus ();
			} 
		
		}

		private void chkAplicaConsignacion_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter) this.txtDiasConsignacion.Focus();
		}

		private void txtDiasConsignacion_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			this.tbBasico.Tabs[1].Selected = true;
			if(e.KeyCode == Keys.Enter) this.optContadoCredito.Focus();	
		}

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);		
		
		}


		private void lblAGRetencion_Click(object sender, System.EventArgs e)
		{}

		private void txtCelular_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			validar.Solonumeros(e);
		}

		private void txtTelefono_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			validar.Solonumeros(e);
		}


		private void chkAplicaConsignacion_CheckedChanged(object sender, System.EventArgs e)
		{
		
		}

		private void txtRuc_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			validar.Solonumeros(e);
			if ((int)this.cmbTipoRuc.Value == 4 || (int)this.cmbTipoRuc.Value == 5) Funcion.CamposNumericos(sender, e);
		}

		private void txtFax_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			validar.Solonumeros(e);
		}

		private void cmbAgenteDeRetencion_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter) this.txtAgenteRetencion.Focus();
		}
		private void btnNuevo_Click(object sender, System.EventArgs e)
		{
			this.btnCancelar_Click(sender, e);

			this.btnNuevo.Enabled = false;
			this.txtidCliente.Value=0;
			this.btnConsultar.Enabled = false;
			this.btnEditar.Enabled = false;
			this.btnGuardarComoCliente.Enabled = true;
			this.btnCancelar.Enabled = true;

			#region Activa controles Cliente
			this.txtAgenteRetencion.Value= System.DBNull.Value;
			this.tbBasico.Tabs[0].Selected = true;
			this.cmbTipoRuc.Focus();
			this.cmbTipoRuc.Enabled=true;
			this.txtRuc.Enabled=true;
			this.txtRuc.Text="";
			this.txtNombre.Enabled=true;
			this.txtNombre.Text="";
			this.txtDireccion.Enabled=true;
			this.txtDireccion.Text="";
			this.cmbGrupo.Enabled=true;
			this.cmbGrupo.Text="";
			this.txtidCliente.Value=0;
			this.txtRepresentante.Enabled=true;
			this.txtRepresentante.Text="";
			this.txtNombreComercial.Enabled=true;
			this.txtNombreComercial.Text="";
			this.cmbSujetoRetencion.Enabled=true;
			this.cmbSujetoRetencion.Text="";
			this.cmbAgenteDeRetencion.Enabled=true;
			this.cmbAgenteDeRetencion.Text="";
			this.txtAgenteRetencion.Enabled=true;
			this.txtAgenteRetencion.Value=0;
			this.txtTelefono.Enabled=true;
			this.txtTelefono.Text="";
			this.txtFax.Enabled=true;
			this.txtFax.Text="";
			this.txtCelular.Enabled=true;
			this.txtCelular.Text="";
			this.txtEMail.Text="";
			this.txtEMail.Enabled=true;
			this.cmbPais.Enabled=true;
			this.cmbPais.Text="";
			this.cmbCiudad.Enabled=true;
			this.cmbCiudad.Text="";
			this.cmbTipoProveedor.Enabled=true;
			this.cmbTipoProveedor.Text="";
			this.txtNotas.Enabled=true;
			this.txtNotas.Text="";
			this.txtValorCredito.Value=0;
			this.txtDiasCredito1.Value=0;
			this.txtDiasCredito2.Value=0;
			this.txtDiasCredito3.Value=0;
      this.txtDiasCredito4.Value=0;
		  
			this.chkCourier.Enabled = true;
	    this.chkAplicaConsignacion.Enabled = true;
			this.spnDescuento.Enabled=true;
			this.spnDescuento.Value=0;
			this.spnDescCredito.Enabled = true;
			this.spnDescCredito.Value=0;
			this.cmbSector.Enabled = true;
			this.cmbSector.Text="";
			this.txtCodigo.Text="";
			this.txtCodigo.Enabled= true;
			this.chkAplicaConsignacion.Checked=false;
			this.txtDiasConsignacion.Value=0;
			this.chkCodigoBarras.Checked=false;
			this.chkCodigoBarras.Enabled=true;
			this.chkRise.Checked=false;
			this.chkRise.Enabled=true;
			this.optContadoCredito.Enabled=true;
			this.ultgContactos.Enabled=true;
			this.ultDirecciones.Enabled=true;
			if(!bClienteProveedor)this.cmbTipoRuc.Value = 4;
	    if(bClienteProveedor)this.cmbTipoRuc.Value = 1;			
			#endregion Activa controles 

			bNuevo = true;
			bEdicion = true;

			if(!bClienteProveedor)this.optContadoCredito.Value= 1;
			if(bClienteProveedor)this.optContadoCredito.Value= 2;

			this.cmbTipoRuc.Focus();
		}

		private void button1_Click_1(object sender, System.EventArgs e)
		{																																													 
			#region validar campos generales
			this.tbBasico.Tabs[0].Selected = true;

			if (!Validacion.vbComboVacio(this.cmbTipoRuc, "Seleccione Ruc o Ced o Pasaporte")) return;	
			if (!Validacion.vbTexto(this.txtRuc, 10, 13, "Por favor ingrese su identidad ")) return;
			if (!Validacion.vbTexto(this.txtNombre, 5, 60, "Ingrese un nombre")) return;
			if (!Validacion.vbTexto(this.txtDireccion, 9, 100, "Ingrese la direccion")) return;
			if (!Validacion.vbComboVacio(this.cmbGrupo, "Seleccione al Grupo que pertenece")) return;
			if(!bClienteProveedor) {if (!Validacion.vbComboVacio(this.cmbAgenteDeRetencion, "Seleccione El Agente de Retención")) return;}
			else {if (!Validacion.vbComboVacio(this.cmbSujetoRetencion, "Seleccione El sujeto de Retención")) return;}
			if (this.txtEMail.Text.ToString().Trim().Length == 0   )
			{
				MessageBox.Show(string.Format("El E-Mail esta vacio por favor ingrese uno "), "Point Technology", MessageBoxButtons.OK,MessageBoxIcon.Stop);
				this.tbBasico.Tabs[0].Selected = true;
				return;
				
			}
			if (!Validacion.bEMail(this.txtEMail.Text.ToString())) return;
			if(bClienteProveedor)if (!Validacion.vbComboVacio(this.cmbPais, "Seleccione el Pais"))return;
			if (!Validacion.vbComboVacio(this.cmbCiudad, "Seleccione la ciudad")) return;
			if(bClienteProveedor){if (!Validacion.vbComboVacio(this.cmbTipoProveedor, "Seleccione el Tipo de proveedor")) return;}
			if (this.txtTelefono.Text.ToString().Trim().Length == 0 &&  this.txtCelular.Text.ToString().Trim().Length == 0  )
			{
				MessageBox.Show(string.Format("Debe Ingresar el Telefono Convencional o el Telefono Celular"), "Point Technology", MessageBoxButtons.OK,MessageBoxIcon.Stop);
				this.txtTelefono.Focus();
				this.tbBasico.Tabs[0].Selected = true;
				return;
				
			}
			
			if (!Validacion.vbValidaTelefono(this.txtTelefono, 1, false)) return;
			if (!Validacion.vbValidaTelefono(this.txtFax, 1, false)) return;
      if (!Validacion.vbValidaTelefono(this.txtCelular, 2, false)) return;
			#region validar sujeto retencion
			if(bClienteProveedor)
			{
				if ((int)this.cmbSujetoRetencion.Value == 8 ) 
				{
					if( this.txtAgenteRetencion.Value == System.DBNull.Value || (int)this.txtAgenteRetencion.Value == 0) 
					{
						MessageBox.Show(string.Format("El campo  no puede estar vacio o Igual a 0, mientras sea contribuyente especial"), "Información", MessageBoxButtons.OK,MessageBoxIcon.Stop);
						this.tbBasico.Tabs[0].Selected = true;
						this.txtAgenteRetencion.Focus();
						return;
					}
				}
			}
			#endregion validar sujeto retencion
		  #region validar agente de retencion
			if(!bClienteProveedor)
			{
				if ((int)this.cmbAgenteDeRetencion.Value == 11) 
															{
																if( this.txtAgenteRetencion.Value == System.DBNull.Value || (int)this.txtAgenteRetencion.Value == 0) 
																{
																	MessageBox.Show(string.Format("El campo  no puede estar vacio o Igual a 0"), "Información", MessageBoxButtons.OK,MessageBoxIcon.Stop);
																	this.tbBasico.Tabs[0].Selected = true;
																	this.txtAgenteRetencion.Focus();
																	return;
																}
															}
			}
			#endregion validar agente de retencion
        this.tbBasico.Tabs[0].Selected = true;
			#region valida Campos dias de Credito
			if( (int)this.optContadoCredito.Value == 2)
			{
				this.tbBasico.Tabs[1].Selected = true;
				if ((int)this.txtDiasCredito1.Value == 0 && (int)this.txtDiasCredito2.Value == 0 && (int)this.txtDiasCredito3.Value == 0 && (int)this.txtDiasCredito4.Value == 0)
				{
					MessageBox.Show("Ingrese los dias de credito para el pago del documento", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					this.txtDiasCredito1.Focus();
					return;
				}
				if ((int)this.txtDiasCredito1.Value == 0 && ((int)this.txtDiasCredito2.Value > 0 || (int)this.txtDiasCredito3.Value > 0 || (int)this.txtDiasCredito4.Value > 0))
				{
					MessageBox.Show("Dias Credito 1 no puede estar en Cero si Dias Credito 2, 3, 4 son mayores a Cero", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					this.txtDiasCredito1.Focus();
					return;
				}
				if ((int)this.txtDiasCredito2.Value == 0 && ((int)this.txtDiasCredito3.Value > 0 || (int)this.txtDiasCredito4.Value > 0))
				{
					MessageBox.Show("Dias Credito 2 no puede estar en Cero si Dias Credito 3, 4 son mayores a Cero", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					this.txtDiasCredito2.Focus();
					return;
				}
				if ((int)this.txtDiasCredito3.Value == 0 && (int)this.txtDiasCredito4.Value > 0)
				{
					MessageBox.Show("Dias Credito 3 no puede estar en Cero si Dias Credito 4 es mayor a Cero", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					this.txtDiasCredito3.Focus();
					return;
				}
				if ((int)this.txtDiasCredito2.Value > 0 && (int)this.txtDiasCredito1.Value > (int)this.txtDiasCredito2.Value)
				{
					MessageBox.Show("Dias Credito 1 no puede ser mayor a Dias de Credito 2", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					this.txtDiasCredito2.Focus();
					return;
				}
				if ((int)this.txtDiasCredito3.Value > 0 && (int)this.txtDiasCredito2.Value > (int)this.txtDiasCredito3.Value)
				{
					MessageBox.Show("Dias Credito 2 no puede ser mayor a Dias de Credito 3", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					this.txtDiasCredito3.Focus();
					return;
				}
				if ((int)this.txtDiasCredito4.Value > 0 && (int)this.txtDiasCredito3.Value > (int)this.txtDiasCredito4.Value)
				{
					MessageBox.Show("Dias Credito 3 no puede ser mayor a Dias de Credito 4", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					this.txtDiasCredito4.Focus();
					return;
				}
				 
			}
			#endregion valida Campos dias de Credito
			
			#endregion validar campos generales
			#region validar combo box
			int idSujetoRetencion = 0;
			int idAgenteRetencion = 0;
      int idTipoProveedor = 0;
			int idCiudad = 0;
			int idClientePais = 0;
			int idGrupocliente = 0;
      int idTipoRuc = 0;
			int idSector = 0;
      if (this.cmbSujetoRetencion.ActiveRow != null) idSujetoRetencion = (int) this.cmbSujetoRetencion.Value;
			if(this.cmbAgenteDeRetencion.ActiveRow != null) idAgenteRetencion = (int) this.cmbAgenteDeRetencion.Value;
      if(this.cmbTipoProveedor.ActiveRow != null) idTipoProveedor  = (int) this.cmbTipoProveedor.Value;
			 if(this.cmbCiudad.ActiveRow != null) idCiudad   = (int) this.cmbCiudad.Value;
      if(this.cmbPais.ActiveRow != null) idClientePais  = (int) this.cmbPais.Value;
			if(this.cmbGrupo.ActiveRow != null) idGrupocliente  = (int) this.cmbGrupo.Value;
      if(this.cmbTipoRuc.ActiveRow != null) idTipoRuc   = (int) this.cmbTipoRuc.Value;
			if(this.cmbSector.ActiveRow != null) idSector    = (int) this.cmbSector.Value;
			if (this.ultgContactos.Rows.Count == 0)
			{
				MessageBox.Show("Para guardar debe ingresar un contacto ", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);					
				this.tbBasico.Tabs[2].Selected = true;
				return;
			}
			#endregion validar combo box
			
        int a = (int) this.txtidCliente.Value;
			#region validar campos datagrid 
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.ultgContactos.Rows.All)
			{
				
				#region Telefonos Vacios
				if (dr.Cells["Telefono"].Value.ToString().Trim().Length == 0 && dr.Cells["Celular"].Value.ToString().Trim().Length == 0)
				{
					MessageBox.Show("Debe Ingresar el Telefono Convencional o el Telefono Celular del Contacto Seleccionado ", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					this.tbBasico.Tabs[2].Selected = true;
					this.ultgContactos.ActiveRow = dr;
					this.ultgContactos.ActiveRow.Selected = true;
					return;
				}
				#endregion Telefonos Vacios

				if((bool)dr.Cells["EnvioMail"].Value)
				{
					if(dr.Cells["Email"].Value.ToString().Length==0)
					{
						
						MessageBox.Show("Ingrese el email del contacto", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
						this.tbBasico.Tabs[2].Selected = true;
						this.ultgContactos.ActiveRow = dr;
						this.ultgContactos.ActiveRow.Selected = true;
						return;
					}
					if (!Validacion.bEMail(dr.Cells["Email"].Value.ToString())) {
						this.tbBasico.Tabs[2].Selected = true; 
						this.ultgContactos.ActiveRow = dr;
						this.ultgContactos.ActiveRow.Selected = true;
						return;
					}
					
				}
				else{dr.Cells["Email"].Value = "";}
			}
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow drValidaContactos in this.ultgContactos.Rows.All)
			{
				#region Telefono
				if (drValidaContactos.Cells["Telefono"].Value.ToString().Trim().Length > 0)
				{
					if (drValidaContactos.Cells["Telefono"].Value.ToString().Trim().Length != 9)
					{
						MessageBox.Show("El Telefono Convencional debe tener 9 Digitos.\n\nLos dos primeros corresponden al Codigo de Provincia", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
						this.tbBasico.Tabs[2].Selected = true;
						this.ultgContactos.ActiveRow = drValidaContactos;
						this.ultgContactos.ActiveRow.Selected = true;
						return;
					}
				}
				#endregion Telefono
				#region Celular
				if (drValidaContactos.Cells["Celular"].Value.ToString().Trim().Length > 0)
				{
					if (drValidaContactos.Cells["Celular"].Value.ToString().Trim().Length != 10)
					{
						MessageBox.Show("El Telefono Celular debe tener 10 digitos", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
						this.tbBasico.Tabs[2].Selected = true;
						this.ultgContactos.ActiveRow = drValidaContactos;
						this.ultgContactos.ActiveRow.Selected = true;
						return;
					}
				}
				#endregion Celular
			}
			#endregion validar campos datagrid 
       
			if ( this.chkAplicaConsignacion.Checked == true)
			{
				if((int)this.txtDiasConsignacion.Value == 0){
					MessageBox.Show("Los días de consignación no pueden estar en cero", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				  	this.tbBasico.Tabs[1].Selected = true;
					this.txtDiasConsignacion.Focus();
					return;
				
				}
			}

			#region Guardar Cliente
			string sSQL = string.Format("Exec GuardaClientesProveedor '{0}','{1}',{2},'{3}',{4},{5},'{6}','{7}',{8},{9},'{10}','{11}','{12}','{13}',{14},{15},'{16}',{17},{18},{19},{20},{21},{22},{23},{24},{25},'{26}',{27},{28},{29},{30},{31},{32},{33},{34}",
				this.txtNombre.Text.ToString(),
				this.txtDireccion.Text.ToString(),
				(int) this.txtidCliente.Value,
				this.txtRuc.Text.ToString(), 
				idTipoRuc ,
				idGrupocliente,
				this.txtRepresentante.Text.ToString(), 
				this.txtNombreComercial.Text.ToString(), 
				idAgenteRetencion ,
				idSujetoRetencion, 
				this.txtTelefono.Text.ToString(), 
				this.txtFax.Text.ToString(), 
				this.txtCelular.Text.ToString(),
				this.txtEMail.Text.ToString(), 
				idCiudad , 
				idTipoProveedor ,
				this.txtNotas.Text.ToString(),
				(bool) this.chkAplicaConsignacion.Checked,
				(int) this.txtDiasConsignacion.Value, 
				(bool) this.chkCodigoBarras.Checked, 
				(int) this.txtDiasCredito1.Value,
				(int) this.txtDiasCredito2.Value,
				(int) this.txtDiasCredito3.Value,
				Convert.ToDecimal (this.spnDescuento.Value), 
				Convert.ToDecimal ( this.spnDescCredito.Value), 
				idSector , 
				this.txtCodigo.Text.ToString(), 
				idClientePais , 
				(int) this.txtAgenteRetencion.Value ,
				Convert.ToDecimal ( this.txtValorCredito.Value ),
				(int) this.txtverificar.Value,
				(int) this.optContadoCredito.Value,
				(bool) this.chkRise.Checked,
				(int) this.txtDiasCredito4.Value,
				(bool) this.chkCourier.Checked);
			this.txtidCliente.Value=Funcion.iEscalarSQL(cdsSeteoF,sSQL);
			this.btnNuevo.Enabled = true;
			this.btnConsultar.Enabled = true;
			this.btnEditar.Enabled = true;
			this.btnGuardarComoCliente.Enabled = true;
			this.btnCancelar.Enabled = true;
			#endregion Guardar Cliente

			#region Guarda Registros contacto
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.ultgContactos.Rows.All)
			{
				string sSQLDetalle = string.Format("Exec GuardaClientesProveedorContacto {0},'{1}','{2}','{3}',{4},'{5}','{6}','{7}',{8},{9}", 
					(int)this.txtidCliente.Value,
					dr.Cells["Contacto"].Value.ToString(),
					dr.Cells["Cargo"].Value.ToString(),
					dr.Cells["Telefono"].Value.ToString(),
          (bool)dr.Cells["EFE"].Value,
					dr.Cells["Email"].Value.ToString(),
					dr.Cells["Celular"].Value.ToString(),
					dr.Cells["Detalle"].Value.ToString(),
					(bool)dr.Cells["EnvioMail"].Value,
					(int)dr.Cells["idClienteContacto"].Value);
				Funcion.EjecutaSQL(cdsSeteoF, sSQLDetalle);
			}
			#endregion Guarda Registros contacto
			#region Guarda Registro Direcciones
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow di in this.ultDirecciones.Rows.All)
			{
				string sSQLDirecciones = string.Format("Exec GuardaClientesProveedorDirecciones {0},{1},'{2}'", 
	
				(int)di.Cells["idDreccion"].Value,
				(int)this.txtidCliente.Value,
				di.Cells["Direccion"].Value.ToString());
				Funcion.EjecutaSQL(cdsSeteoF, sSQLDirecciones);
			}
			#endregion Guarda Registro Direcciones
			#region  Controles
			this.btnCancelar.Enabled = true;
			this.tbBasico.Tabs[0].Selected = true;
			this.cmbTipoRuc.Enabled=false;
			this.txtRuc.Enabled=false;
			this.txtNombre.Enabled=false;
			this.txtDireccion.Enabled=false;
			this.cmbGrupo.Enabled=false;
			this.txtRepresentante.Enabled=false;
			this.txtNombreComercial.Enabled=false;
			this.cmbSujetoRetencion.Enabled=false;
			this.cmbAgenteDeRetencion.Enabled=false;
			this.txtAgenteRetencion.Enabled=false;	
			this.txtTelefono.Enabled=false;
			this.txtFax.Enabled=false;	
			this.txtCelular.Enabled=false;		
			this.txtEMail.Enabled=false;
			this.cmbPais.Enabled=false;
			this.cmbCiudad.Enabled=false;
			this.cmbTipoProveedor.Enabled=false;
			this.txtNotas.Enabled=false;
			this.spnDescuento.Enabled=false;
			this.spnDescCredito.Enabled =false;
			this.cmbSector.Enabled =false;
			this.txtCodigo.Enabled=false;
			this.chkAplicaConsignacion.Enabled=false;
			this.txtDiasConsignacion.Enabled=false;
			this.chkCodigoBarras.Enabled=false;
			this.optContadoCredito.Enabled=false;
			#endregion Controles
			#region bloquear controles
			this.tbBasico.Tabs[0].Selected = true;
			this.cmbTipoRuc.Focus();
			this.cmbTipoRuc.Enabled=false;
			this.txtRuc.Enabled=false;
			this.txtNombre.Enabled=false;
			this.txtDireccion.Enabled=false;
			this.cmbGrupo.Enabled=false;
			this.txtRepresentante.Enabled=false;
			this.txtNombreComercial.Enabled=false;
			this.cmbSujetoRetencion.Enabled=false;
			this.cmbAgenteDeRetencion.Enabled=false;
			this.txtAgenteRetencion.Enabled=false;
			this.txtTelefono.Enabled=false;
			this.txtFax.Enabled=false;
			this.txtCelular.Enabled=false;
			this.txtEMail.Enabled=false;
			this.cmbPais.Enabled=false;
			this.cmbCiudad.Enabled=false;
			this.cmbTipoProveedor.Enabled=false;
			this.txtNotas.Enabled=false;
			this.spnDescuento.Enabled=false;
			this.spnDescCredito.Enabled = false;
			this.cmbSector.Enabled = false;
			this.txtCodigo.Enabled= false;
			this.chkAplicaConsignacion.Enabled=false;
			this.txtDiasConsignacion.Enabled=false;
			this.chkCodigoBarras.Enabled=false;
			this.optContadoCredito.Enabled=false;
			this.ultgContactos.Enabled=false;
			this.ultDirecciones.Enabled=false;
			this.chkRise.Enabled=false;
			this.chkCourier.Enabled = false;
			this.btnNuevo.Enabled = true;
			this.btnConsultar.Enabled = true;
			this.btnEditar.Enabled = true;
			this.btnGuardarComoCliente.Enabled = false;
			this.btnCancelar.Enabled = true;
			this.txtValorCredito.Enabled=false;
			this.txtDiasCredito1.Enabled=false;
			this.txtDiasCredito2.Enabled=false;
			this.txtDiasCredito3.Enabled =false;
			this.txtDiasCredito4.Enabled =false;
			
			bNuevo = false;
			bEdicion = false;
			#endregion bloquear controles

			this.tbBasico.Tabs[0].Selected = true;

			if (iOrigen  == 1){
			this.DialogResult = DialogResult.OK;
			}
			if (iOrigen  == 2)
			{
				this.DialogResult = DialogResult.OK;
			}
		}

		private void btnSalir_Click(object sender, System.EventArgs e)
		{
			this.Close();	
		}

		private void btnCancelar_Click(object sender, System.EventArgs e)
		{
			#region Bloquea Controles
			bNuevo = false;
			bEdicion = false;

			this.txtidCliente.Value=0;
			this.btnCancelar.Enabled = true;
			this.tbBasico.Tabs[0].Selected = true;
			this.cmbTipoRuc.Focus();
			this.cmbTipoRuc.Enabled=false;
			this.cmbTipoRuc.Text="";
			this.txtRuc.Enabled=false;
			this.txtRuc.Text="";
			this.txtNombre.Enabled=false;
			this.txtNombre.Text="";
			this.txtDireccion.Enabled=false;
			this.txtDireccion.Text="";
			this.cmbGrupo.Enabled=false;
			this.cmbGrupo.Text="";
			this.txtRepresentante.Enabled=false;
			this.txtRepresentante.Text="";
			this.txtNombreComercial.Enabled=false;
			this.txtNombreComercial.Text="";
			this.cmbSujetoRetencion.Enabled=false;
			this.cmbSujetoRetencion.Text="";
			this.cmbAgenteDeRetencion.Enabled=false;
			this.cmbAgenteDeRetencion.Text="";
			this.txtAgenteRetencion.Enabled=false;
			this.txtAgenteRetencion.Value=0;
			this.txtTelefono.Enabled=false;
			this.txtTelefono.Text="";
			this.txtFax.Enabled=false;
			this.txtFax.Text="";
			this.txtCelular.Enabled=false;
			this.txtCelular.Text="";
			this.txtEMail.Text="";
			this.txtEMail.Enabled=false;
			this.cmbPais.Enabled=false;
			this.cmbPais.Text="";
			this.cmbCiudad.Enabled=false;
			this.cmbCiudad.Text="";
			this.cmbTipoProveedor.Enabled=false;
			this.cmbTipoProveedor.Text="";
			this.txtNotas.Enabled=false;
			this.txtNotas.Text="";
			this.txtValorCredito.Value=0;
			this.txtDiasCredito1.Value=0;
			this.txtDiasCredito2.Value=0;
			this.txtDiasCredito3.Value=0;
			this.txtDiasCredito4.Value=0;
			this.spnDescuento.Enabled=false;
			this.spnDescuento.Value=0;
			this.spnDescCredito.Enabled =false;
			this.spnDescCredito.Value=0;
			this.cmbSector.Enabled =false;
			this.cmbSector.Text="";
			this.txtCodigo.Text="";
			this.txtCodigo.Enabled=false;
			this.chkAplicaConsignacion.Checked=false;
			this.chkAplicaConsignacion.Enabled=false;
			this.chkRise.Checked=false;
			this.chkRise.Enabled=false;
			this.txtDiasConsignacion.Enabled=false;
			this.txtDiasConsignacion.Value=0;
			this.chkCodigoBarras.Checked=false;
			this.chkCodigoBarras.Enabled=false;
			this.txtValorCredito.Enabled=false;
			this.txtDiasCredito1.Enabled=false;
			this.txtDiasCredito2.Enabled=false;
			this.txtDiasCredito3.Enabled=false;
      this.txtDiasCredito4.Enabled=false;
			this.optContadoCredito.Enabled=false;
			this.optContadoCredito.Value= 0;
			this.ultgContactos.Enabled=false;
			this.ultDirecciones.Enabled=false;
			this.chkCourier.Enabled=false;
			this.chkCourier.Checked= false;
		
     #endregion Bloquea Controles
			string sSQL = string.Format("Exec ConsultaGeneralClienteProveedorContacto  {0}, '{1}'", 
				this.txtidCliente.Value	,this.txtRuc.Text.ToString());
			this.ultgContactos.DataSource = Funcion.dvProcedimiento(cdsSeteoF, sSQL);
      
			string sSQLL = string.Format("Exec ConsultaGeneralClienteProveedorDirecciones  {0}", 
				this.txtidCliente.Value);
			this.ultDirecciones.DataSource = Funcion.dvProcedimiento(cdsSeteoF, sSQLL);


			this.btnNuevo.Enabled = true;
			this.btnConsultar.Enabled = true;
			this.btnEditar.Enabled = false;
		  this.btnGuardarComoCliente.Enabled = false;
			this.btnCancelar.Enabled = false;

			if (iOrigen  == 1 && bcarga == false)
			{
				this.DialogResult = DialogResult.No;
			}
			
		}

		private void btnGuardar_Click(object sender, System.EventArgs e)
		{
			this.btnNuevo.Enabled = true;
			this.btnConsultar.Enabled = true;
			this.btnEditar.Enabled = true;
			this.btnGuardarComoCliente.Enabled = false;
			this.btnCancelar.Enabled = true;
		}

		private void btnEditar_Click(object sender, System.EventArgs e)
		{
			#region Activa controles Boton editar
			this.btnNuevo.Enabled = false;
			this.btnConsultar.Enabled = false;
			this.btnEditar.Enabled = false;
			this.btnGuardarComoCliente.Enabled = true;
			this.btnCancelar.Enabled = true;
			this.tbBasico.Tabs[0].Selected = true;
			this.cmbTipoRuc.Focus();
			this.cmbTipoRuc.Enabled=true;
			this.txtRuc.Enabled=true;
	    this.txtNombre.Enabled=true;
	    this.txtDireccion.Enabled=true;
			this.cmbGrupo.Enabled=true;
			this.txtRepresentante.Enabled=true;
			this.txtNombreComercial.Enabled=true;
			this.cmbSujetoRetencion.Enabled=true;
			this.cmbAgenteDeRetencion.Enabled=true;
			this.txtAgenteRetencion.Enabled=true;
			this.txtTelefono.Enabled=true;
			this.txtFax.Enabled=true;
			this.txtCelular.Enabled=true;
			this.txtEMail.Enabled=true;
			this.cmbPais.Enabled=true;
			this.cmbCiudad.Enabled=true;
			this.cmbTipoProveedor.Enabled=true;
			this.txtNotas.Enabled=true;
		  this.spnDescuento.Enabled=true;
      this.spnDescCredito.Enabled = true;
      this.cmbSector.Enabled = true;
      this.txtCodigo.Enabled= true;
			this.chkAplicaConsignacion.Enabled=true;
			this.chkRise.Enabled=true;
			this.txtDiasConsignacion.Enabled=false;
			this.chkCodigoBarras.Enabled=true;
			this.optContadoCredito.Enabled=true;
			this.ultgContactos.Enabled=true;
			this.ultDirecciones.Enabled=true;
			this.chkCourier.Enabled = true;
			#endregion Activa controles Boton editar
			#region Valida la existencia del ruc o cedula
			int IdCliente = (int)this.txtidCliente.Value;
			if (IdCliente > 0)
			{
				this.cmbTipoRuc.Enabled = false;
				this.txtRuc.Enabled = false; 
			}
			this.btnGuardarComoCliente.Enabled = true ;
			this.btnNuevo.Enabled = false;
			this.btnConsultar.Enabled = false;
			this.btnEditar.Enabled = false;
			this.btnCancelar.Enabled = true;
//			this.btnGuardarComoCliente.Enabled = true;
      #endregion Valida la existencia del ruc o cedula

	    if( (int)this.optContadoCredito.Value == 1)
			{
				bool vcredito = false;
				valor_credio( vcredito );
			} 
			else 
			{
				bool vcredito  = true;
				valor_credio( vcredito );
			}
			
			bEdicion=true;
			
		}
		private void Consultar (int IdCliente )
		{
			
			#region Consulta Individual
			string ssql= string.Format ("Exec ConsultaIndividualClienteProveedor {0}",IdCliente); 
			SqlDataReader dr = Funcion.rEscalarSQL(cdsSeteoF,ssql,true);
			dr.Read();
			if (dr.HasRows)
			{
				if (dr.GetValue(0) != System.DBNull.Value)this.txtNombre.Text= dr.GetString  (0);
				if (dr.GetValue(1) != System.DBNull.Value)this.txtDireccion.Text = dr.GetString (1);
				if (dr.GetValue(2) != System.DBNull.Value)this.txtidCliente.Value = dr.GetInt32(2);
				if (dr.GetValue(3) != System.DBNull.Value)this.txtRuc.Text = dr.GetString (3);
				this.cmbTipoRuc.Value = dr.GetInt32(4);
				if (dr.GetValue(5) != System.DBNull.Value)this.cmbGrupo.Value = dr.GetInt32(5);
				if (dr.GetValue(6) != System.DBNull.Value)this.txtRepresentante.Text= dr.GetString  (6);
				if (dr.GetValue(7) != System.DBNull.Value)this.txtNombreComercial.Text= dr.GetString  (7);
				if (dr.GetValue(8) != System.DBNull.Value)this.cmbAgenteDeRetencion.Value = dr.GetInt32(8);
				if (dr.GetValue(9) != System.DBNull.Value)this.cmbSujetoRetencion.Value = dr.GetInt32(9);
				if (dr.GetValue(10) != System.DBNull.Value)this.txtTelefono.Text = dr.GetString  (10);
				if (dr.GetValue(11) != System.DBNull.Value)this.txtFax.Text = dr.GetString  (11);
				if (dr.GetValue(12) != System.DBNull.Value)this.txtCelular.Text = dr.GetString  (12);
				if (dr.GetValue(13) != System.DBNull.Value)this.txtEMail.Text = dr.GetString  (13);
				if (dr.GetValue(14) != System.DBNull.Value)this.cmbCiudad.Value = dr.GetInt32(14);
				if (dr.GetValue(15) != System.DBNull.Value)this.cmbTipoProveedor.Value = dr.GetInt32(15);
				if (dr.GetValue(16) != System.DBNull.Value)this.txtNotas.Text = dr.GetString  (16);
				this.chkAplicaConsignacion.Checked = dr.GetBoolean(17);
				if (dr.GetValue(18) != System.DBNull.Value)this.txtDiasConsignacion.Value = dr.GetInt32(18);
				this.chkCodigoBarras.Checked = dr.GetBoolean(19);
				if (dr.GetValue(20) != System.DBNull.Value)this.txtDiasCredito1.Value = dr.GetInt32(20);
				if (dr.GetValue(21) != System.DBNull.Value)this.txtDiasCredito2.Value = dr.GetInt32(21);
				if (dr.GetValue(22) != System.DBNull.Value)this.txtDiasCredito3.Value = dr.GetInt32(22);
				if (dr.GetValue(23) != System.DBNull.Value)this.spnDescuento.Value = dr.GetDecimal(23);
				if (dr.GetValue(24) != System.DBNull.Value)this.spnDescCredito.Value = dr.GetDecimal(24);
				if (dr.GetValue(25) != System.DBNull.Value)this.cmbSector.Value = dr.GetInt32(25);
				if (dr.GetValue(26) != System.DBNull.Value)this.txtCodigo.Text = dr.GetString(26);
				if (dr.GetValue(27) != System.DBNull.Value)this.cmbPais.Value = dr.GetInt32(27);
				if (dr.GetValue(28) != System.DBNull.Value)this.txtAgenteRetencion.Value = dr.GetInt32(28);
				if (dr.GetValue(29) != System.DBNull.Value)this.txtValorCredito.Value = dr.GetDecimal(29);
				if (dr.GetValue(30) != System.DBNull.Value){this.optContadoCredito.Value = dr.GetInt32(30);}
//				else
//				{
//					if(bClienteProveedor== true )
//					{
//					this.optContadoCredito.Value = 1;
//					}
//					else{
//					this.optContadoCredito.Value = 0;
//					} 
//				
//				}
				if (dr.GetValue(31) != System.DBNull.Value)this.chkRise.Checked = dr.GetBoolean(31);
			  if (dr.GetValue(32) != System.DBNull.Value)this.txtDiasCredito4.Value = dr.GetInt32(32);
				if (dr.GetValue(33) != System.DBNull.Value)this.chkCourier.Checked = dr.GetBoolean(33);

			}
			
			
			dr.Close();
			#endregion Consulta Individual

			string sSQL = string.Format("Exec ConsultaGeneralClienteProveedorContacto  {0}, '{1}'", 
				this.txtidCliente.Value	,this.txtRuc.Text.ToString());
			this.ultgContactos.DataSource = Funcion.dvProcedimiento(cdsSeteoF, sSQL);
     
			string sSQLL = string.Format("Exec ConsultaGeneralClienteProveedorDirecciones  {0}", 
				this.txtidCliente.Value);
			this.ultDirecciones.DataSource = Funcion.dvProcedimiento(cdsSeteoF, sSQLL);


      this.btnNuevo.Enabled = true;
			this.btnConsultar.Enabled = true;
			this.btnEditar.Enabled = true;
		  this.btnGuardarComoCliente.Enabled = false;
		   this.btnCancelar.Enabled = true;
		}
		
		private void btnConsultar_Click(object sender, System.EventArgs e)
		{
			this.btnCancelar_Click(sender, e);

			using (frmConsultaClienteProveedor consultab = new frmConsultaClienteProveedor(bClienteProveedor))
			{
				if (DialogResult.OK == consultab.ShowDialog())
				{
					this.Consultar((int)consultab.udtG.ActiveRow.Cells["idCliente"].Value);

					string sSQL = string.Format("Exec ConsultaGeneralClienteProveedorContacto  {0}, '{1}'", 
					this.txtidCliente.Value	,this.txtRuc.Text.ToString());
					this.ultgContactos.DataSource = Funcion.dvProcedimiento(cdsSeteoF, sSQL);

					if (this.ultgContactos.Rows.Count > 0) this.ultgContactos.Focus(); else this.txtRuc.Focus();
				}						
			}
			

		  
		}

		private void cmbAgenteDeRetencion_ValueChanged(object sender, System.EventArgs e)
		{
			
		}

		private void cmbPais_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}

		private void cmbPais_Click(object sender, System.EventArgs e)
		{
		
		}

		private void bPais_Click(object sender, System.EventArgs e)
		{
			using (ClientePais miPais = new ClientePais())
			{
				miPais.ShowDialog();
			}
			cmbPais.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Select idClientePais, Pais From ClientePais");
		}

		private void btGrupoCliente_Click(object sender, System.EventArgs e)
		{
		
		}

		private void tpComplemento_Paint(object sender, System.Windows.Forms.PaintEventArgs e)
		{
		
		}

		private void txtCodigo_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			validar.Solonumeros(e);
		}

		private void txtAgenteRetencion_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter) this.txtTelefono.Focus();
		}

		private void optContadoCredito_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			
		}
	
		private void optContadoCredito_ValueChanged(object sender, System.EventArgs e)
		{
			
			if( !bEdicion)return;

			if( (int)this.optContadoCredito.Value == 1)
			{
				bool vcredito = false;
				valor_credio( vcredito );
				} 
			else 
			{
				bool vcredito  =true;
				valor_credio( vcredito );
			}
		}

		private void cmbPais_KeyDown_1(object sender, System.Windows.Forms.KeyEventArgs e)
		{

			if(e.KeyCode == Keys.Enter) this.cmbCiudad.Focus();
		}

		private void cmbPais_InitializeLayout_1(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}

		private void chkCodigoBarras_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			this.tbBasico.Tabs[1].Selected = true;
			if(e.KeyCode == Keys.Enter) this.optContadoCredito.Focus();
		}

		private void optContadoCredito_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if( (int)this.optContadoCredito.Value == 2){if(e.KeyCode == Keys.Enter) this.txtValorCredito.Focus();}
			if( (int)this.optContadoCredito.Value == 1){if(e.KeyCode == Keys.Enter) this.spnDescuento.Focus();}
			}

		private void txtValorCredito_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter) this.txtDiasCredito1.Focus();
		}

		private void txtDiasCredito1_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter) this.txtDiasCredito2.Focus();
		}

		private void txtDiasCredito2_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter) this.txtDiasCredito3.Focus();
		}

		private void txtDiasCredito3_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter) this.txtDiasCredito4.Focus();
		}

		private void spnDescuento_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter) this.spnDescCredito.Focus();
		}

		private void spnDescCredito_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter) this.cmbSector.Focus();
		}

		private void cmbSector_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter) this.txtCodigo.Focus();
		}

		private void ultraGrid2_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}

		private void frmClientesProveedores_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (!bEdicion)
			{
				if (e.KeyCode == Keys.N) this.btnNuevo_Click(sender, e);
				if (e.KeyCode == Keys.B) this.btnConsultar_Click(sender, e);
				if ((int)this.txtidCliente.Value > 0)
				{
					if (e.KeyCode == Keys.E) if (this.btnEditar.Enabled) this.btnEditar_Click(sender, e);					
					if (e.KeyCode == Keys.F5) this.Consultar((int)this.txtidCliente.Value);					
				}
				if (e.KeyCode == Keys.Escape) this.Close();				
			}
			else
			{
				if (e.KeyCode == Keys.Escape) 
					if (DialogResult.Yes == MessageBox.Show("Esta seguro de Cancelar la Edición", "Point Technology", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2))
						this.btnCancelar_Click(sender, e);

				if (e.KeyCode == Keys.F12) this.button1_Click_1(sender, e);				
			}
		}

		private void frmClientesProveedores_Closing(object sender, System.ComponentModel.CancelEventArgs e)
		{
			
		}
		private void ultraGrid2_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			
		}
		private void ultgContactos_BeforeSelectChange(object sender, Infragistics.Win.UltraWinGrid.BeforeSelectChangeEventArgs e)
		{
			
		}

		private void ultgContactos_CellChange(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			
		}

		private void ultraGrid2_InitializeLayout_1(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
			
		}

		private void ultgContactos_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			FuncionesProcedimientos.DesplazamientoGrids(sender,e,this.ultgContactos);
		}

		private void ultgContactos_CellChange_1(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			
			if(e.Cell.Column.ToString()=="EnvioMail")
			{
				this.ultgContactos.UpdateData();
				if((bool)e.Cell.Row.Cells["EnvioMail"].Value== true)
				{
					e.Cell.Row.Cells["Email"].Hidden=false;
				}
				else
				{
					e.Cell.Row.Cells["Email"].Hidden=true;
				
				}
			}
		}

		private void ultgContactos_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
//			if (	(bool)this.ultgContactos.ActiveRow.Cells["EnvioMail"].Value)
//				e.Handled=true;}
			if (ultgContactos.ActiveCell.Column.Key.ToString().ToUpper() == "TELEFONO" || ultgContactos.ActiveCell.Column.Key.ToString().ToUpper() == "CELULAR"){
				Funcion.CamposNumericos(sender, e);}

		
           
		}

		private void txtverificar_ValueChanged(object sender, System.EventArgs e)
		{
		
		}

		private void txtNombreComercial_ValueChanged(object sender, System.EventArgs e)
		{
		
		}

		private void ultgContactos_AfterRowInsert(object sender, Infragistics.Win.UltraWinGrid.RowEventArgs e)
		{
			e.Row.Cells["idClienteContacto"].Value = 0;
			e.Row.Cells["EFE"].Value = false;
			e.Row.Cells["EnvioMail"].Value = false;
			
			

		}

		private void txtRuc_Validating(object sender, System.ComponentModel.CancelEventArgs e)
		{
			
			if (bEdicion)
				if (!Validacion.vbIdentificacion(this.txtRuc, (int)this.cmbTipoRuc.Value, cdsSeteoF)) e.Cancel = true;
		}

		private void cmbTipoRuc_ValueChanged(object sender, System.EventArgs e)
		{
			if (this.cmbTipoRuc.ActiveRow != null)
			{
				if ((int)this.cmbTipoRuc.Value == 4 || (int)this.cmbTipoRuc.Value == 6) this.txtRuc.MaxLength = 13;
				if ((int)this.cmbTipoRuc.Value == 5) this.txtRuc.MaxLength = 10;

			}
			if(!bEdicion) return;
			if (!TipoCliente)
			{
				if ((int)this.cmbTipoRuc.Value == 3) 
				{
					cmbGrupo.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Select idGrupoCliente, Grupo From ClienteGrupo Where Proveedor = '" + TipoCliente + "' And idGrupoCliente In (2, 3) Order by Grupo ");
					drCliente["idGrupoCliente"] = 0;
					drCliente["idTipoProveedor"] = 2;

					this.ultgContactos.DisplayLayout.Bands[0].Columns["Telefono"].FieldLen = 15;
					this.ultgContactos.DisplayLayout.Bands[0].Columns["Celular"].FieldLen = 15;
					
				}
				else
				{
					cmbGrupo.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Select idGrupoCliente, Grupo From ClienteGrupo Where Proveedor = '" + TipoCliente + "' And idGrupoCliente Not In (2, 3) Order by Grupo ");
					drCliente["idGrupoCliente"] = 0;
					drCliente["idTipoProveedor"] = System.DBNull.Value;

					this.ultgContactos.DisplayLayout.Bands[0].Columns["Telefono"].FieldLen = 9;
					this.ultgContactos.DisplayLayout.Bands[0].Columns["Celular"].FieldLen = 10;
				}
			}	
	
			if((int)this.txtverificar.Value==0)
			{

				if ((int)this.cmbTipoRuc.Value ==5) 
				{
				   this.cmbAgenteDeRetencion.Value = 5;
					this.cmbAgenteDeRetencion.Enabled=false;
				}
				else {
					  this.cmbAgenteDeRetencion.Enabled=true;this.cmbAgenteDeRetencion.Text="";
				 }
			
			}
		}

		private void txtNombre_Validated(object sender, System.EventArgs e)
		{
		
		}

		private void txtRuc_Validated(object sender, System.EventArgs e)
		{		
//	    if (bEdicion && bNuevo)
//	    this.Consultar(Funcion.iEscalarSQL(cdsSeteoF, string.Format("if Exists(select idCliente from Cliente where Ruc= '{0}' and Proveedor = {1})select idCliente from Cliente where Ruc= '{0}' and Proveedor = {1}  Else Select CONVERT(Int, 0)",this.txtRuc.Text.ToString(),(int) this.txtverificar.Value)));
//		
			if (bEdicion)// && bNuevo)
			{
				int IdCliente= Funcion.iEscalarSQL(cdsSeteoF,string.Format("if Exists(select idCliente from Cliente where Ruc= '{0}' and Proveedor = {1})select idCliente from Cliente where Ruc= '{0}' and Proveedor = {1}  Else Select CONVERT(Int, 0)",this.txtRuc.Text.ToString(),(int) this.txtverificar.Value));
 	       if (IdCliente > 0)
				{
					this.Consultar(IdCliente);

					 this.cmbTipoRuc.Enabled = false;
					 this.txtRuc.Enabled = false;
					 
					this.bNuevo = false;
					this.bEdicion = true;
					 this.btnNuevo.Enabled = false;
					 this.btnConsultar.Enabled = false;
					 this.btnEditar.Enabled = false;
					 this.btnCancelar.Enabled = true;
					 this.btnGuardarComoCliente.Enabled = true ;

				}
			}
		}

		private void btGrupoCliente_Click_1(object sender, System.EventArgs e)
		{
			if(!bClienteProveedor){
				TipoCliente = true;
				using (ClienteGrupo miGrupoC = new ClienteGrupo(TipoCliente))
				{				
					if (DialogResult.OK == miGrupoC.ShowDialog())
					{
						string sSQL = string.Format("Select idGrupoCliente, Grupo From ClienteGrupo Where Proveedor = {0}", TipoCliente);
					
						cmbGrupo.DataSource = Funcion.dvProcedimiento(cdsSeteoF, sSQL);
					}
				}
				}
			else
			{
					TipoCliente = false;
				using (ClienteGrupo miGrupoC = new ClienteGrupo(TipoCliente))
				{				
					if (DialogResult.OK == miGrupoC.ShowDialog())
					{
						string sSQL = string.Format("Select idGrupoCliente, Grupo From ClienteGrupo Where Proveedor = {0}", 	TipoCliente);
					
						cmbGrupo.DataSource = Funcion.dvProcedimiento(cdsSeteoF, sSQL);
					}
				}
			}

		}

		private void txtAgenteRetencion_Validating(object sender, System.ComponentModel.CancelEventArgs e)
		{	
		}

		private void txtAgenteRetencion_Validated(object sender, System.EventArgs e)
		{
			 
		}

		private void button1_Click(object sender, System.EventArgs e)
		{
//			bNuevo =false;
//			Consultar2();
		}

		private void ultgContactos_BeforeRowInsert(object sender, Infragistics.Win.UltraWinGrid.BeforeRowInsertEventArgs e)
		{
		}
		private void txtAgenteRetencion_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			
		}
		private void Consultar2 ()
		{

		
		}

		private void txtValorCredito_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			
		}

		private void button1_Click_2(object sender, System.EventArgs e)
		{

		}

		private void chkRise_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter) {this.txtNombre.Focus();}
		}

		private void txtDiasConsignacion_ValueChanged(object sender, System.EventArgs e)
		{
		
		}

		private void txtCodigo_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{

			if(!bClienteProveedor )if(e.KeyCode == Keys.Enter) {this.chkAplicaConsignacion.Focus();}
			if(bClienteProveedor )if(e.KeyCode == Keys.Enter) {this.chkCodigoBarras.Focus();}
			
			
		}

		private void chkCodigoBarras_KeyDown_1(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter) this.chkAplicaConsignacion.Focus();
			
			
		}

		private void txtDiasConsignacion_KeyDown_1(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(!bClienteProveedor )if(e.KeyCode == Keys.Enter) 
														 {
															 this.tbBasico.Tabs[2].Selected = true; this.ultgContactos.DisplayLayout.Bands[0].AddNew();
															 this.ultgContactos.ActiveCell = this.ultgContactos.Rows[0].Cells["Contacto"];}
			if(bClienteProveedor )if(e.KeyCode == Keys.Enter) {this.chkCourier.Focus();}

						
		
		}

		private void chkAplicaConsignacion_KeyDown_1(object sender, System.Windows.Forms.KeyEventArgs e)
		{

			
			if(e.KeyCode == Keys.Enter) 
				
			{
				if( (bool)this.chkAplicaConsignacion.Checked == true)
				{
					this.txtDiasConsignacion.Focus();
				}
				else
				{
				  this.chkCourier.Focus();
				}
			}
			

		}

		private void cmbSujetoRetencion_Validating(object sender, System.ComponentModel.CancelEventArgs e)
		{

		}

		private void cmbAgenteDeRetencion_Validating(object sender, System.ComponentModel.CancelEventArgs e)
		{
		
		}

		private void cmbSujetoRetencion_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}

		private void txtDiasCredito4_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter) this.spnDescuento.Focus();
		}

		private void tbBasico_SelectedTabChanged(object sender, Infragistics.Win.UltraWinTabControl.SelectedTabChangedEventArgs e)
		{
		
		}
		private void  valor_credio (bool vcredito )
		{
			if(vcredito )
			{
				this.txtValorCredito.Enabled=true;
				this.txtDiasCredito1.Enabled=true;
				this.txtDiasCredito2.Enabled=true;
				this.txtDiasCredito3.Enabled =true;
				this.txtDiasCredito4.Enabled =true;
				this.txtValorCredito.Focus();
			}
			else 
			{
				this.spnDescuento.Focus();
				this.txtValorCredito.Enabled=false;
				this.txtDiasCredito1.Enabled=false;
				this.txtDiasCredito2.Enabled=false;
				this.txtDiasCredito3.Enabled =false;
				this.txtDiasCredito4.Enabled =false;
				this.txtValorCredito.Value=0;
				this.txtDiasCredito1.Value=0;
				this.txtDiasCredito2.Value=0;
				this.txtDiasCredito3.Value=0;
				this.txtDiasCredito4.Value=0;
			}
		}

		private void ultgContactos_InitializeRow(object sender, Infragistics.Win.UltraWinGrid.InitializeRowEventArgs e)
		{
		
		}

		private void lblNombreComercial_Click(object sender, System.EventArgs e)
		{
		
		}

		private void lblSujetoDeRetencion_Click(object sender, System.EventArgs e)
		{
		
		}

		private void ultDirecciones_AfterRowInsert(object sender, Infragistics.Win.UltraWinGrid.RowEventArgs e)
		{
			e.Row.Cells["idDreccion"].Value = 0;
		}

		private void ultDirecciones_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			FuncionesProcedimientos.DesplazamientoGrids(sender,e,this.ultDirecciones);
		}

		private void ultraTabPageControl5_Paint(object sender, System.Windows.Forms.PaintEventArgs e)
		{
		
		}

		private void chkCourier_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
						if(e.KeyCode == Keys.Enter) 
						{
							this.tbBasico.Tabs[2].Selected = true; this.ultgContactos.DisplayLayout.Bands[0].AddNew();
							this.ultgContactos.ActiveCell = this.ultgContactos.Rows[0].Cells["Contacto"];}
		}

		private void chkAplicaConsignacion_Validating(object sender, System.ComponentModel.CancelEventArgs e)
		{
			if( !bEdicion)return;

			if( (bool)this.chkAplicaConsignacion.Checked == true)
			{
				this.txtDiasConsignacion.Enabled = true;
			} 
			else 
			{
				this.txtDiasConsignacion.Enabled = false;
        this.txtDiasConsignacion.Value = 0;
			}
		}

		private void chkAplicaConsignacion_CheckedChanged_1(object sender, System.EventArgs e)
		{
			if (!bEdicion) return;
			if( this.chkAplicaConsignacion.Checked)
			{
				this.txtDiasConsignacion.Enabled = true;
				
			}
			else
			{
				this.txtDiasConsignacion.Enabled = false;
				this.txtDiasConsignacion.Value = 0;
			}
		}

		private void txtNombre_InitializeEditorButtonCheckState(object sender, Infragistics.Win.UltraWinEditors.InitializeCheckStateEventArgs e)
		{
		
		}

		private void txtDireccion_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			
					validar.SoloCaracter(e);
		}
	
		class validar 
		{
			public static void Solonumeros (KeyPressEventArgs Nu)
			{
				if(char.IsDigit(Nu.KeyChar) )
				{
			
					Nu.Handled = false;
			
				}
				else if (char.IsControl(Nu.KeyChar))
				{
			
					Nu.Handled = false;		
				} 
				else
				{
					Nu.Handled= true;
				}
			}


			public static void SoloCaracter (KeyPressEventArgs Nu)
			{
				if(char.IsLetterOrDigit(Nu.KeyChar) || char.IsSeparator(Nu.KeyChar)  || char.IsControl(Nu.KeyChar)  )
				{
			
					Nu.Handled = false;
			
				}
				else if ( char.IsPunctuation (Nu.KeyChar)|| char.IsSurrogate (Nu.KeyChar) || char.IsSymbol (Nu.KeyChar)    )
				{
			
					Nu.Handled = true;		
				} 
				else
				{
					Nu.Handled= true;
				}
			}
			


			
	
		}
	


	}
}
