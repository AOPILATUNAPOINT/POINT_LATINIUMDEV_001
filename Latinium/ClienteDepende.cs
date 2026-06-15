using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de ClienteDepende.
	/// </summary>
	public class ClienteDepende : DevExpress.XtraEditors.XtraForm
	{
		private Infragistics.Win.UltraWinGrid.UltraGrid ultraGrid1;
		private C1.Data.C1DataSet cdsCliente;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbPadre;
		private System.Windows.Forms.Label label1;
		private C1.Data.C1DataSet cdsPadre;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtIdCliente;
		public Infragistics.Win.UltraWinEditors.UltraTextEditor IdBuscado;
		private Infragistics.Win.Misc.UltraButton btBuscar;
		private Infragistics.Win.Misc.UltraButton btCerrar;

		int IdPadre = 0;
		public ClienteDepende(int idPadre)
		{
			InitializeComponent();
			IdPadre = idPadre;
		}

		public ClienteDepende()
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

		#region Código generado por el Diseñador de Windows Forms
		/// <summary>
		/// Método necesario para admitir el Diseñador. No se puede modificar
		/// el contenido del método con el editor de código.
		/// </summary>
		private void InitializeComponent()
		{
			Infragistics.Win.Appearance appearance1 = new Infragistics.Win.Appearance();
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
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn24 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Foto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn25 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Pais");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn26 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Notas");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn27 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Ambos");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn28 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Proveedor");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn29 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("AutorizaFactura");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn30 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("AutorizaImprenta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn31 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("MaximoCredito");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn32 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("DiasCredito");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn33 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("SerieFactura");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn34 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Fecha");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn35 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaIngreso");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn36 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("DescuentoContado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn37 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("DescuentoCredito");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn38 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Borrar");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn39 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("SaldoCxC");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn40 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("SaldoCxP");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn41 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Editar");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn42 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nuevo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn43 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaEditado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn44 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idRetRenta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn45 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idRetIva");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn46 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaCaducidad");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn47 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("AutorizaRet");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn48 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("SerieRet");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn49 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ContribEspecial");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn50 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CuentaCorr");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn51 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CuentaAh");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn52 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Conyuge");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn53 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ConyDireccion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn54 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ConyTelefono");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn55 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ConyCelular");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn56 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Referencia");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn57 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("RefTelefono");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn58 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("RefDireccion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn59 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Poliza");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn60 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Parroquia");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn61 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Sector");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn62 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("RefLocal");
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Cliente", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn63 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCliente");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn64 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idGrupoCliente");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn65 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCobrador");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn66 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idVendedor");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn67 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idTipoPrecio");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn68 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCiudad");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn69 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idSector");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn70 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idTipoRuc");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn71 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idSucursal");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn72 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPadre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn73 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idFormaPago");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn74 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCredTribut");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn75 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn76 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn77 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Direccion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn78 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Telefono");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn79 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Fax");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn80 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Ruc");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn81 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Observacion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn82 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Representante");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn83 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Email");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn84 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Celular");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn85 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Contacto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn86 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Foto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn87 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Pais");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn88 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Notas");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn89 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Ambos");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn90 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Proveedor");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn91 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("AutorizaFactura");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn92 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("AutorizaImprenta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn93 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("MaximoCredito");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn94 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("DiasCredito");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn95 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("SerieFactura");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn96 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Fecha");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn97 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaIngreso");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn98 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("DescuentoContado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn99 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("DescuentoCredito");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn100 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Borrar");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn101 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("SaldoCxC");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn102 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("SaldoCxP");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn103 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Editar");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn104 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nuevo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn105 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaEditado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn106 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idRetRenta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn107 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idRetIva");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn108 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaCaducidad");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn109 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("AutorizaRet");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn110 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("SerieRet");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn111 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ContribEspecial");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn112 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CuentaCorr");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn113 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CuentaAh");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn114 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Conyuge");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn115 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ConyDireccion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn116 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ConyTelefono");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn117 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ConyCelular");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn118 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Referencia");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn119 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("RefTelefono");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn120 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("RefDireccion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn121 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Poliza");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn122 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Parroquia");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn123 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Sector");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn124 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("RefLocal");
			this.ultraGrid1 = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.cdsCliente = new C1.Data.C1DataSet();
			this.btBuscar = new Infragistics.Win.Misc.UltraButton();
			this.cmbPadre = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.cdsPadre = new C1.Data.C1DataSet();
			this.label1 = new System.Windows.Forms.Label();
			this.btCerrar = new Infragistics.Win.Misc.UltraButton();
			this.txtIdCliente = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.IdBuscado = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsCliente)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbPadre)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsPadre)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdCliente)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.IdBuscado)).BeginInit();
			this.SuspendLayout();
			// 
			// ultraGrid1
			// 
			this.ultraGrid1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.ultraGrid1.Cursor = System.Windows.Forms.Cursors.Default;
			this.ultraGrid1.DataMember = "Cliente";
			this.ultraGrid1.DataSource = this.cdsCliente;
			appearance1.BackColor = System.Drawing.Color.White;
			this.ultraGrid1.DisplayLayout.Appearance = appearance1;
			ultraGridColumn1.Header.VisiblePosition = 18;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn2.Header.VisiblePosition = 20;
			ultraGridColumn2.Hidden = true;
			ultraGridColumn3.Header.VisiblePosition = 19;
			ultraGridColumn3.Hidden = true;
			ultraGridColumn4.Header.VisiblePosition = 24;
			ultraGridColumn4.Hidden = true;
			ultraGridColumn5.Header.VisiblePosition = 22;
			ultraGridColumn5.Hidden = true;
			ultraGridColumn6.Header.VisiblePosition = 17;
			ultraGridColumn6.Hidden = true;
			ultraGridColumn7.Header.VisiblePosition = 21;
			ultraGridColumn7.Hidden = true;
			ultraGridColumn8.Header.VisiblePosition = 23;
			ultraGridColumn8.Hidden = true;
			ultraGridColumn9.Header.VisiblePosition = 34;
			ultraGridColumn9.Hidden = true;
			ultraGridColumn10.Header.VisiblePosition = 40;
			ultraGridColumn10.Hidden = true;
			ultraGridColumn11.Header.VisiblePosition = 44;
			ultraGridColumn12.Header.VisiblePosition = 47;
			ultraGridColumn13.Header.VisiblePosition = 0;
			ultraGridColumn13.Width = 427;
			ultraGridColumn14.Header.VisiblePosition = 6;
			ultraGridColumn14.Hidden = true;
			ultraGridColumn15.Header.VisiblePosition = 11;
			ultraGridColumn15.Hidden = true;
			ultraGridColumn16.Header.VisiblePosition = 33;
			ultraGridColumn16.Hidden = true;
			ultraGridColumn17.Header.VisiblePosition = 13;
			ultraGridColumn17.Hidden = true;
			ultraGridColumn18.Header.VisiblePosition = 31;
			ultraGridColumn18.Hidden = true;
			ultraGridColumn19.Header.VisiblePosition = 27;
			ultraGridColumn19.Hidden = true;
			ultraGridColumn20.Header.VisiblePosition = 30;
			ultraGridColumn20.Hidden = true;
			ultraGridColumn21.Header.VisiblePosition = 12;
			ultraGridColumn21.Hidden = true;
			ultraGridColumn22.Header.VisiblePosition = 5;
			ultraGridColumn22.Hidden = true;
			ultraGridColumn23.Header.VisiblePosition = 7;
			ultraGridColumn23.Hidden = true;
			ultraGridColumn24.Header.VisiblePosition = 16;
			ultraGridColumn24.Hidden = true;
			ultraGridColumn25.Header.VisiblePosition = 28;
			ultraGridColumn25.Hidden = true;
			ultraGridColumn26.Header.VisiblePosition = 26;
			ultraGridColumn26.Hidden = true;
			ultraGridColumn27.Header.VisiblePosition = 1;
			ultraGridColumn27.Hidden = true;
			ultraGridColumn28.Header.VisiblePosition = 29;
			ultraGridColumn28.Hidden = true;
			ultraGridColumn29.Header.VisiblePosition = 2;
			ultraGridColumn29.Hidden = true;
			ultraGridColumn30.Header.VisiblePosition = 3;
			ultraGridColumn30.Hidden = true;
			ultraGridColumn31.Header.VisiblePosition = 25;
			ultraGridColumn31.Hidden = true;
			ultraGridColumn32.Header.VisiblePosition = 10;
			ultraGridColumn32.Hidden = true;
			ultraGridColumn33.Header.VisiblePosition = 32;
			ultraGridColumn33.Hidden = true;
			ultraGridColumn34.Header.VisiblePosition = 14;
			ultraGridColumn34.Hidden = true;
			ultraGridColumn35.Header.VisiblePosition = 15;
			ultraGridColumn35.Hidden = true;
			ultraGridColumn36.Header.VisiblePosition = 8;
			ultraGridColumn36.Hidden = true;
			ultraGridColumn37.Header.VisiblePosition = 9;
			ultraGridColumn37.Hidden = true;
			ultraGridColumn38.Header.VisiblePosition = 4;
			ultraGridColumn38.Hidden = true;
			ultraGridColumn39.Header.VisiblePosition = 35;
			ultraGridColumn39.Hidden = true;
			ultraGridColumn40.Header.VisiblePosition = 36;
			ultraGridColumn40.Hidden = true;
			ultraGridColumn41.Header.VisiblePosition = 37;
			ultraGridColumn41.Hidden = true;
			ultraGridColumn42.Header.VisiblePosition = 38;
			ultraGridColumn42.Hidden = true;
			ultraGridColumn43.Header.VisiblePosition = 39;
			ultraGridColumn43.Hidden = true;
			ultraGridColumn44.Header.VisiblePosition = 41;
			ultraGridColumn45.Header.VisiblePosition = 42;
			ultraGridColumn46.Header.VisiblePosition = 43;
			ultraGridColumn47.Header.VisiblePosition = 45;
			ultraGridColumn48.Header.VisiblePosition = 46;
			ultraGridColumn49.Header.VisiblePosition = 48;
			ultraGridColumn50.Header.VisiblePosition = 49;
			ultraGridColumn51.Header.VisiblePosition = 50;
			ultraGridColumn52.Header.VisiblePosition = 51;
			ultraGridColumn53.Header.VisiblePosition = 52;
			ultraGridColumn54.Header.VisiblePosition = 53;
			ultraGridColumn55.Header.VisiblePosition = 54;
			ultraGridColumn56.Header.VisiblePosition = 55;
			ultraGridColumn57.Header.VisiblePosition = 56;
			ultraGridColumn58.Header.VisiblePosition = 57;
			ultraGridColumn59.Header.VisiblePosition = 58;
			ultraGridColumn60.Header.VisiblePosition = 59;
			ultraGridColumn61.Header.VisiblePosition = 60;
			ultraGridColumn62.Header.VisiblePosition = 61;
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
																										 ultraGridColumn62});
			this.ultraGrid1.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.ultraGrid1.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
			this.ultraGrid1.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.ultraGrid1.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.False;
			appearance2.BackColor = System.Drawing.Color.Transparent;
			this.ultraGrid1.DisplayLayout.Override.CardAreaAppearance = appearance2;
			this.ultraGrid1.DisplayLayout.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.RowSelect;
			this.ultraGrid1.DisplayLayout.Override.ColumnAutoSizeMode = Infragistics.Win.UltraWinGrid.ColumnAutoSizeMode.AllRowsInBand;
			appearance3.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance3.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance3.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance3.FontData.BoldAsString = "True";
			appearance3.FontData.Name = "Arial";
			appearance3.FontData.SizeInPoints = 10F;
			appearance3.ForeColor = System.Drawing.Color.White;
			appearance3.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.ultraGrid1.DisplayLayout.Override.HeaderAppearance = appearance3;
			this.ultraGrid1.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti;
			appearance4.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance4.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance4.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.RowSelectorAppearance = appearance4;
			appearance5.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance5.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance5.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.SelectedRowAppearance = appearance5;
			this.ultraGrid1.DisplayLayout.Override.SelectTypeCell = Infragistics.Win.UltraWinGrid.SelectType.None;
			this.ultraGrid1.DisplayLayout.Override.SelectTypeCol = Infragistics.Win.UltraWinGrid.SelectType.None;
			this.ultraGrid1.DisplayLayout.Override.SelectTypeRow = Infragistics.Win.UltraWinGrid.SelectType.Single;
			this.ultraGrid1.DisplayLayout.ScrollBounds = Infragistics.Win.UltraWinGrid.ScrollBounds.ScrollToFill;
			this.ultraGrid1.DisplayLayout.ScrollStyle = Infragistics.Win.UltraWinGrid.ScrollStyle.Immediate;
			this.ultraGrid1.DisplayLayout.TabNavigation = Infragistics.Win.UltraWinGrid.TabNavigation.NextControl;
			this.ultraGrid1.DisplayLayout.ViewStyle = Infragistics.Win.UltraWinGrid.ViewStyle.SingleBand;
			this.ultraGrid1.Location = new System.Drawing.Point(16, 40);
			this.ultraGrid1.Name = "ultraGrid1";
			this.ultraGrid1.Size = new System.Drawing.Size(448, 184);
			this.ultraGrid1.TabIndex = 0;
			// 
			// cdsCliente
			// 
			this.cdsCliente.BindingContextCtrl = this;
			this.cdsCliente.DataLibrary = "LibFacturacion";
			this.cdsCliente.DataLibraryUrl = "";
			this.cdsCliente.DataSetDef = "dsCliente";
			this.cdsCliente.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsCliente.Name = "cdsCliente";
			this.cdsCliente.SchemaClassName = "LibFacturacion.DataClass";
			this.cdsCliente.SchemaDef = null;
			this.cdsCliente.BeforeFill += new C1.Data.FillEventHandler(this.cdsCliente_BeforeFill);
			// 
			// btBuscar
			// 
			this.btBuscar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btBuscar.DialogResult = System.Windows.Forms.DialogResult.OK;
			this.btBuscar.Location = new System.Drawing.Point(48, 240);
			this.btBuscar.Name = "btBuscar";
			this.btBuscar.TabIndex = 1;
			this.btBuscar.Text = "&Buscar";
			this.btBuscar.Click += new System.EventHandler(this.btBuscar_Click);
			this.btBuscar.MouseMove += new System.Windows.Forms.MouseEventHandler(this.btBuscar_MouseDown);
			// 
			// cmbPadre
			// 
			appearance6.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbPadre.Appearance = appearance6;
			this.cmbPadre.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbPadre.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbPadre.DataMember = "Cliente";
			this.cmbPadre.DataSource = this.cdsPadre;
			this.cmbPadre.DisplayLayout.AutoFitColumns = true;
			ultraGridColumn63.Header.VisiblePosition = 18;
			ultraGridColumn63.Hidden = true;
			ultraGridColumn64.Header.VisiblePosition = 20;
			ultraGridColumn64.Hidden = true;
			ultraGridColumn65.Header.VisiblePosition = 19;
			ultraGridColumn65.Hidden = true;
			ultraGridColumn66.Header.VisiblePosition = 24;
			ultraGridColumn66.Hidden = true;
			ultraGridColumn67.Header.VisiblePosition = 22;
			ultraGridColumn67.Hidden = true;
			ultraGridColumn68.Header.VisiblePosition = 17;
			ultraGridColumn68.Hidden = true;
			ultraGridColumn69.Header.VisiblePosition = 21;
			ultraGridColumn69.Hidden = true;
			ultraGridColumn70.Header.VisiblePosition = 23;
			ultraGridColumn70.Hidden = true;
			ultraGridColumn71.Header.VisiblePosition = 34;
			ultraGridColumn71.Hidden = true;
			ultraGridColumn72.Header.VisiblePosition = 40;
			ultraGridColumn72.Hidden = true;
			ultraGridColumn73.Header.VisiblePosition = 44;
			ultraGridColumn73.Width = 8;
			ultraGridColumn74.Header.VisiblePosition = 47;
			ultraGridColumn74.Width = 8;
			ultraGridColumn75.Header.VisiblePosition = 0;
			ultraGridColumn75.Width = 8;
			ultraGridColumn76.Header.VisiblePosition = 6;
			ultraGridColumn76.Hidden = true;
			ultraGridColumn77.Header.VisiblePosition = 11;
			ultraGridColumn77.Hidden = true;
			ultraGridColumn78.Header.VisiblePosition = 33;
			ultraGridColumn78.Hidden = true;
			ultraGridColumn79.Header.VisiblePosition = 13;
			ultraGridColumn79.Hidden = true;
			ultraGridColumn80.Header.VisiblePosition = 31;
			ultraGridColumn80.Hidden = true;
			ultraGridColumn81.Header.VisiblePosition = 27;
			ultraGridColumn81.Hidden = true;
			ultraGridColumn82.Header.VisiblePosition = 30;
			ultraGridColumn82.Hidden = true;
			ultraGridColumn83.Header.VisiblePosition = 12;
			ultraGridColumn83.Hidden = true;
			ultraGridColumn84.Header.VisiblePosition = 5;
			ultraGridColumn84.Hidden = true;
			ultraGridColumn85.Header.VisiblePosition = 7;
			ultraGridColumn85.Hidden = true;
			ultraGridColumn86.Header.VisiblePosition = 16;
			ultraGridColumn86.Hidden = true;
			ultraGridColumn87.Header.VisiblePosition = 28;
			ultraGridColumn87.Hidden = true;
			ultraGridColumn88.Header.VisiblePosition = 26;
			ultraGridColumn88.Hidden = true;
			ultraGridColumn89.Header.VisiblePosition = 1;
			ultraGridColumn89.Hidden = true;
			ultraGridColumn90.Header.VisiblePosition = 29;
			ultraGridColumn90.Hidden = true;
			ultraGridColumn91.Header.VisiblePosition = 2;
			ultraGridColumn91.Hidden = true;
			ultraGridColumn92.Header.VisiblePosition = 3;
			ultraGridColumn92.Hidden = true;
			ultraGridColumn93.Header.VisiblePosition = 25;
			ultraGridColumn93.Hidden = true;
			ultraGridColumn94.Header.VisiblePosition = 10;
			ultraGridColumn94.Hidden = true;
			ultraGridColumn95.Header.VisiblePosition = 32;
			ultraGridColumn95.Hidden = true;
			ultraGridColumn96.Header.VisiblePosition = 14;
			ultraGridColumn96.Hidden = true;
			ultraGridColumn97.Header.VisiblePosition = 15;
			ultraGridColumn97.Hidden = true;
			ultraGridColumn98.Header.VisiblePosition = 8;
			ultraGridColumn98.Hidden = true;
			ultraGridColumn99.Header.VisiblePosition = 9;
			ultraGridColumn99.Hidden = true;
			ultraGridColumn100.Header.VisiblePosition = 4;
			ultraGridColumn100.Hidden = true;
			ultraGridColumn101.Header.VisiblePosition = 35;
			ultraGridColumn101.Hidden = true;
			ultraGridColumn102.Header.VisiblePosition = 36;
			ultraGridColumn102.Hidden = true;
			ultraGridColumn103.Header.VisiblePosition = 37;
			ultraGridColumn103.Hidden = true;
			ultraGridColumn104.Header.VisiblePosition = 38;
			ultraGridColumn104.Hidden = true;
			ultraGridColumn105.Header.VisiblePosition = 39;
			ultraGridColumn105.Hidden = true;
			ultraGridColumn106.Header.VisiblePosition = 41;
			ultraGridColumn106.Width = 8;
			ultraGridColumn107.Header.VisiblePosition = 42;
			ultraGridColumn107.Width = 8;
			ultraGridColumn108.Header.VisiblePosition = 43;
			ultraGridColumn108.Width = 16;
			ultraGridColumn109.Header.VisiblePosition = 45;
			ultraGridColumn109.Width = 8;
			ultraGridColumn110.Header.VisiblePosition = 46;
			ultraGridColumn110.Width = 8;
			ultraGridColumn111.Header.VisiblePosition = 48;
			ultraGridColumn111.Width = 16;
			ultraGridColumn112.Header.VisiblePosition = 49;
			ultraGridColumn112.Width = 10;
			ultraGridColumn113.Header.VisiblePosition = 50;
			ultraGridColumn113.Width = 21;
			ultraGridColumn114.Header.VisiblePosition = 51;
			ultraGridColumn114.Width = 21;
			ultraGridColumn115.Header.VisiblePosition = 52;
			ultraGridColumn115.Width = 21;
			ultraGridColumn116.Header.VisiblePosition = 53;
			ultraGridColumn116.Width = 21;
			ultraGridColumn117.Header.VisiblePosition = 54;
			ultraGridColumn117.Width = 21;
			ultraGridColumn118.Header.VisiblePosition = 55;
			ultraGridColumn118.Width = 21;
			ultraGridColumn119.Header.VisiblePosition = 56;
			ultraGridColumn119.Width = 21;
			ultraGridColumn120.Header.VisiblePosition = 57;
			ultraGridColumn120.Width = 21;
			ultraGridColumn121.Header.VisiblePosition = 58;
			ultraGridColumn121.Width = 21;
			ultraGridColumn122.Header.VisiblePosition = 59;
			ultraGridColumn122.Width = 21;
			ultraGridColumn123.Header.VisiblePosition = 60;
			ultraGridColumn123.Width = 21;
			ultraGridColumn124.Header.VisiblePosition = 61;
			ultraGridColumn124.Width = 21;
			ultraGridBand2.Columns.AddRange(new object[] {
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
																										 ultraGridColumn124});
			this.cmbPadre.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			this.cmbPadre.DisplayMember = "Nombre";
			this.cmbPadre.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbPadre.Location = new System.Drawing.Point(88, 8);
			this.cmbPadre.Name = "cmbPadre";
			this.cmbPadre.Size = new System.Drawing.Size(368, 21);
			this.cmbPadre.TabIndex = 2;
			this.cmbPadre.ValueMember = "idCliente";
			// 
			// cdsPadre
			// 
			this.cdsPadre.BindingContextCtrl = this;
			this.cdsPadre.DataLibrary = "LibFacturacion";
			this.cdsPadre.DataLibraryUrl = "";
			this.cdsPadre.DataSetDef = "dsCliente";
			this.cdsPadre.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsPadre.Name = "cdsPadre";
			this.cdsPadre.SchemaClassName = "LibFacturacion.DataClass";
			this.cdsPadre.SchemaDef = null;
			this.cdsPadre.BeforeFill += new C1.Data.FillEventHandler(this.cdsPadre_BeforeFill);
			// 
			// label1
			// 
			this.label1.AutoSize = true;
			this.label1.Location = new System.Drawing.Point(16, 8);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(48, 16);
			this.label1.TabIndex = 3;
			this.label1.Text = "Principal";
			// 
			// btCerrar
			// 
			this.btCerrar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btCerrar.DialogResult = System.Windows.Forms.DialogResult.Cancel;
			this.btCerrar.Location = new System.Drawing.Point(352, 240);
			this.btCerrar.Name = "btCerrar";
			this.btCerrar.TabIndex = 4;
			this.btCerrar.Text = "&Cerrar";
			// 
			// txtIdCliente
			// 
			this.txtIdCliente.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsCliente, "Cliente.idCliente"));
			this.txtIdCliente.Location = new System.Drawing.Point(168, 240);
			this.txtIdCliente.Name = "txtIdCliente";
			this.txtIdCliente.Size = new System.Drawing.Size(40, 21);
			this.txtIdCliente.TabIndex = 5;
			// 
			// IdBuscado
			// 
			this.IdBuscado.Location = new System.Drawing.Point(232, 240);
			this.IdBuscado.Name = "IdBuscado";
			this.IdBuscado.Size = new System.Drawing.Size(40, 21);
			this.IdBuscado.TabIndex = 6;
			// 
			// ClienteDepende
			// 
			this.AcceptButton = this.btBuscar;
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(480, 278);
			this.Controls.Add(this.IdBuscado);
			this.Controls.Add(this.txtIdCliente);
			this.Controls.Add(this.btCerrar);
			this.Controls.Add(this.label1);
			this.Controls.Add(this.cmbPadre);
			this.Controls.Add(this.btBuscar);
			this.Controls.Add(this.ultraGrid1);
			this.Name = "ClienteDepende";
			this.Text = "Dependientes";
			this.Closing += new System.ComponentModel.CancelEventHandler(this.ClienteDepende_Closing);
			this.Load += new System.EventHandler(this.ClienteDepende_Load);
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsCliente)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbPadre)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsPadre)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdCliente)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.IdBuscado)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void ClienteDepende_Load(object sender, System.EventArgs e)
		{
			string stAudita = "Exec AuditaCrear 53, 6, 'Depende Cl.'";
			Funcion.EjecutaSQL(cdsCliente, stAudita, true);
			IdBuscado.Width = 0;
			txtIdCliente.Width = 0;
			string stFiltro = "idCliente = " + IdPadre.ToString();
			C1.Data.FilterCondition fcPadre = new 
				C1.Data.FilterCondition(cdsPadre, "Cliente", stFiltro);
			C1.Data.FilterConditions fcTotal = new C1.Data.FilterConditions();
			fcTotal.Add(fcPadre);

			cdsPadre.Fill(fcTotal, false);
			cmbPadre.Value = IdPadre;

			stFiltro = "idPadre = " + IdPadre.ToString();
			C1.Data.FilterCondition fcPadre1 = new 
				C1.Data.FilterCondition(cdsCliente, "Cliente", stFiltro);
			C1.Data.FilterConditions fcTotal1 = new C1.Data.FilterConditions();
			fcTotal1.Add(fcPadre1);

			cdsCliente.Fill(fcTotal1, false);
		}

		private void btBuscar_Click(object sender, System.EventArgs e)
		{
			IdBuscado.Value = txtIdCliente.Value;
		}

		private void btBuscar_MouseDown(object sender, System.Windows.Forms.MouseEventArgs e)
		{
			if (e.Button != MouseButtons.Right) return;
			IdBuscado.Value = cmbPadre.Value;
			DialogResult = DialogResult.OK;
		}

		private void cdsCliente_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			this.cdsCliente.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);

		}

		private void cdsPadre_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			this.cdsPadre.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);

		}

		private void ClienteDepende_Closing(object sender, System.ComponentModel.CancelEventArgs e)
		{
			string stAudita = "Exec AuditaCrear 53, 7, 'Depende Cl.'";
			Funcion.EjecutaSQL(cdsCliente, stAudita, true);
		}
	}
}
