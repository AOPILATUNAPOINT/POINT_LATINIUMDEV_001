using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using Microsoft.VisualBasic;
using System.IO;
using System.Diagnostics;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de frmCob_InformeNotCompro.
	/// </summary>
	public class frmCob_InformeNotCompro : System.Windows.Forms.Form
	{
		
		private System.Windows.Forms.Label label3;
		private System.Windows.Forms.Label label11;
		private System.Windows.Forms.Label label6;
		private System.Windows.Forms.Button btnGuardar;
		private System.Windows.Forms.Button btnSalir;
		private System.Windows.Forms.GroupBox groupBox1;
		public Infragistics.Win.UltraWinEditors.UltraTextEditor txtObservaciones;
		public Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtFecha;
		public Infragistics.Win.UltraWinGrid.UltraCombo cmbGestionCredito;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource1;
		private C1.Data.C1DataSet cdsSeteoF;
		private System.Windows.Forms.Label label1;
		public Infragistics.Win.UltraWinGrid.UltraCombo cmbGestion;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource2;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtValor;
		private System.Windows.Forms.Label lblFechaPago;
		private System.Windows.Forms.Label lblValor;
		private Infragistics.Win.UltraWinEditors.UltraOptionSet optAsignados;
		public Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtFechaPago;
		private System.Windows.Forms.Label lblParentesco;
		public Infragistics.Win.UltraWinGrid.UltraCombo cmbParentesco;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource3;
		private System.Windows.Forms.Button btnTelefonia;
		private System.Windows.Forms.Button btnGestion;
		public Infragistics.Win.UltraWinEditors.UltraNumericEditor txtidPersonal;
		public Infragistics.Win.UltraWinEditors.UltraNumericEditor idCompra;
		public Infragistics.Win.UltraWinEditors.UltraNumericEditor txtidCliente;
		private System.ComponentModel.IContainer components;
		public Infragistics.Win.UltraWinEditors.UltraNumericEditor txtidCompra;
		private System.Windows.Forms.Timer timer1;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource4;
		private System.Windows.Forms.Label lblContador;
		private System.Windows.Forms.GroupBox groupBox2;
		private System.Windows.Forms.Timer timer2;
		private System.Windows.Forms.Label lblDireccion;
		public Infragistics.Win.UltraWinEditors.UltraTextEditor txtDireccion;
		private System.Windows.Forms.Button btnCobranza;
		public Infragistics.Win.UltraWinEditors.UltraTextEditor txtEmail;
		public Infragistics.Win.UltraWinEditors.UltraTextEditor txttelefono;
		private System.Windows.Forms.Label label2;
		public Infragistics.Win.UltraWinEditors.UltraTextEditor txtUbicacion;
		private System.Windows.Forms.GroupBox groupBox3;

		int idPersonal ;
		int idCliente;
		int idCompras;
		int idEntidadFinanciera;
		private System.Windows.Forms.Button btnWhatssap;
		public Infragistics.Win.UltraWinEditors.UltraTextEditor txtURL;
		public Infragistics.Win.UltraWinEditors.UltraTextEditor txtDirectorio;
		public Infragistics.Win.UltraWinEditors.UltraTextEditor txtDirectorioFinal;
		public Infragistics.Win.UltraWinEditors.UltraTextEditor txtDirectorioarchivo;
		private System.Windows.Forms.Label label4;
		string sRuc ;
		string sNumeroFactura;
		private System.Windows.Forms.Label lblMail;
		private Infragistics.Win.UltraWinGrid.UltraGrid uGridDocumentos;
		string sCliente;
		string sOperadorCobrador ;
		public Infragistics.Win.UltraWinEditors.UltraTextEditor txtEmailOpeSup;
		private System.Windows.Forms.Label label5;
		private System.Windows.Forms.Label lblVendedor;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbVendedor;
		int iOperadorCobradorVendedor = 0;
		private System.Windows.Forms.Button btnTelefoniacli;
		private Infragistics.Win.UltraWinGrid.UltraGrid uWeb;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource5;
		private System.Windows.Forms.Label label7;
		public Infragistics.Win.UltraWinGrid.UltraCombo cmbTipocontacto;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource6;
		int iBodega = 0;

		public frmCob_InformeNotCompro(int IdCompra, int IdPersonal, string SRuc , int IDEntidadFinanciera, string SNumeroFactura, string SCliente, string SOperadorCobrador, int IOperadorCobradorVendedor, int IdCbo_GestorDeCobranzas)
		{
			//
			// Necesario para admitir el Diseñador de Windows Forms
			//
			InitializeComponent();
			idCompras = IdCompra  ;
			idPersonal = IdPersonal ;
			sRuc = SRuc ;
			idEntidadFinanciera = IDEntidadFinanciera;
			sNumeroFactura = SNumeroFactura;
			sCliente = SCliente;
			sOperadorCobrador = SOperadorCobrador ; 
			iOperadorCobradorVendedor = IOperadorCobradorVendedor ;
			idCbo_GestorDeCobranzas = IdCbo_GestorDeCobranzas;
			//
			// TODO: agregar código de constructor después de llamar a InitializeComponent
			//
		}

		public frmCob_InformeNotCompro(int IdCompra, string SRuc , int IDEntidadFinanciera, string SNumeroFactura, string SCliente, int IOperadorCobradorVendedor, int IBodega, int IdCbo_GestorDeCobranzas)
		{
			//
			// Necesario para admitir el Diseñador de Windows Forms
			//
			InitializeComponent();
			idCompras = IdCompra  ;
			sRuc = SRuc ;
			idEntidadFinanciera = IDEntidadFinanciera;
			sNumeroFactura = SNumeroFactura;
			sCliente = SCliente;
			iOperadorCobradorVendedor = IOperadorCobradorVendedor ;
			iBodega = IBodega;
			idCbo_GestorDeCobranzas = IdCbo_GestorDeCobranzas;

			//
			// TODO: agregar código de constructor después de llamar a InitializeComponent
			//
		}

		string sDireccionEnvio ="";
		int idUsuario = 0;
		string sDireccion = "";
		string  sTelefono ="";
		string sEmail ="";
		bool bWeb = true;
		string sClienteW ="";
		string sCedulaW ="";
		int iPedido;
		int idCbo_GestorDeCobranzas = 0;
		public frmCob_InformeNotCompro(string SDireccionEnvio, int IdUsuario, string  STelefono, string SEmail, bool BWeb , string SClienteW, string sCedulaW , int IPedido, int IdCbo_GestorDeCobranzas )
		{
			//
			// Necesario para admitir el Diseñador de Windows Forms
			//

			sDireccionEnvio =SDireccionEnvio;
			idUsuario = IdUsuario;
			sTelefono =STelefono;
			sEmail =SEmail;
			bWeb = BWeb;
			sClienteW =SClienteW;
			sCedulaW =sCedulaW;
			iPedido =  IPedido;
			idCbo_GestorDeCobranzas = IdCbo_GestorDeCobranzas;
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
			this.components = new System.ComponentModel.Container();
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmCob_InformeNotCompro));
			Infragistics.Win.Appearance appearance1 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton1 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCbo_EstadoGestion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Estado");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCbo_EstadoGestion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Estado");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCbo_ResultadoGestion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Resultado");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn3 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCbo_ResultadoGestion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn4 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Resultado");
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton2 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.ValueListItem valueListItem1 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem2 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem3 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand3 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idParentesco");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn5 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idParentesco");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn6 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn7 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCompra");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn8 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCbo_GestionesDeCobranzas");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn9 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Fecha");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn10 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Operador_Cobrador");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn11 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Estado");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn12 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Resultado");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn13 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("FechaPago");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn14 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Valor");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn15 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("ESTADO_PAGO");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn16 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn17 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Usuario");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn18 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Notas");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn19 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Cliente");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn20 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Direccion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn21 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Telefono");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn22 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("DireccionUrlWSP");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn23 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Tipo");
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance11 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance12 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance13 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance14 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance15 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand4 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCompra");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCbo_GestionesDeCobranzas");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Fecha");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Operador_Cobrador");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn11 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Estado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn12 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Resultado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn13 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaPago");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn14 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Valor");
			Infragistics.Win.Appearance appearance16 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn15 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ESTADO_PAGO");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn16 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn17 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Usuario");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn18 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Notas");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn19 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cliente", -1, null, 0, Infragistics.Win.UltraWinGrid.SortIndicator.Ascending, false);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn20 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Direccion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn21 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Telefono");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn22 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("DireccionUrlWSP");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn23 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Tipo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn24 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Ir", 0);
			Infragistics.Win.Appearance appearance17 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance18 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance19 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance20 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance21 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance22 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance23 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance24 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance25 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand5 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn25 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPersonal");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn26 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn27 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.Appearance appearance26 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand6 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn28 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Fecha");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn29 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Estado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn30 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cliente");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn31 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Operador");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn32 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Notas");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn33 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Tipo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn34 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Ir", 0);
			Infragistics.Win.Appearance appearance27 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance28 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance29 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance30 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance31 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance32 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance33 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn24 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Fecha");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn25 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Estado");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn26 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Cliente");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn27 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Operador");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn28 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Notas");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn29 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Tipo");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand7 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn35 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCbo_EstadosTipocontacto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn36 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Estado");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn30 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCbo_EstadosTipocontacto");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn31 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Estado");
			this.label3 = new System.Windows.Forms.Label();
			this.label11 = new System.Windows.Forms.Label();
			this.label6 = new System.Windows.Forms.Label();
			this.btnGuardar = new System.Windows.Forms.Button();
			this.btnSalir = new System.Windows.Forms.Button();
			this.groupBox1 = new System.Windows.Forms.GroupBox();
			this.txtObservaciones = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtidPersonal = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.dtFecha = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.cmbGestionCredito = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.ultraDataSource1 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.label1 = new System.Windows.Forms.Label();
			this.cmbGestion = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.ultraDataSource2 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.dtFechaPago = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.lblFechaPago = new System.Windows.Forms.Label();
			this.lblValor = new System.Windows.Forms.Label();
			this.txtValor = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.optAsignados = new Infragistics.Win.UltraWinEditors.UltraOptionSet();
			this.lblParentesco = new System.Windows.Forms.Label();
			this.cmbParentesco = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.ultraDataSource3 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.btnTelefonia = new System.Windows.Forms.Button();
			this.btnGestion = new System.Windows.Forms.Button();
			this.txtidCliente = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtidCompra = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.timer1 = new System.Windows.Forms.Timer(this.components);
			this.ultraDataSource4 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.lblContador = new System.Windows.Forms.Label();
			this.groupBox2 = new System.Windows.Forms.GroupBox();
			this.timer2 = new System.Windows.Forms.Timer(this.components);
			this.lblDireccion = new System.Windows.Forms.Label();
			this.txtDireccion = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.btnCobranza = new System.Windows.Forms.Button();
			this.txtEmail = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txttelefono = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label2 = new System.Windows.Forms.Label();
			this.txtUbicacion = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.groupBox3 = new System.Windows.Forms.GroupBox();
			this.btnWhatssap = new System.Windows.Forms.Button();
			this.txtURL = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtDirectorio = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtDirectorioarchivo = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtDirectorioFinal = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label4 = new System.Windows.Forms.Label();
			this.lblMail = new System.Windows.Forms.Label();
			this.uGridDocumentos = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.txtEmailOpeSup = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label5 = new System.Windows.Forms.Label();
			this.lblVendedor = new System.Windows.Forms.Label();
			this.cmbVendedor = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.btnTelefoniacli = new System.Windows.Forms.Button();
			this.uWeb = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource5 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.cmbTipocontacto = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.ultraDataSource6 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.label7 = new System.Windows.Forms.Label();
			((System.ComponentModel.ISupportInitialize)(this.txtObservaciones)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtidPersonal)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFecha)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbGestionCredito)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbGestion)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFechaPago)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtValor)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.optAsignados)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbParentesco)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource3)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtidCliente)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtidCompra)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource4)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDireccion)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtEmail)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txttelefono)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtUbicacion)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtURL)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDirectorio)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDirectorioarchivo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDirectorioFinal)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridDocumentos)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtEmailOpeSup)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbVendedor)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.uWeb)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource5)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbTipocontacto)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource6)).BeginInit();
			this.SuspendLayout();
			// 
			// label3
			// 
			this.label3.AutoSize = true;
			this.label3.Location = new System.Drawing.Point(8, 184);
			this.label3.Name = "label3";
			this.label3.Size = new System.Drawing.Size(63, 16);
			this.label3.TabIndex = 449;
			this.label3.Text = "Descripción";
			this.label3.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label11
			// 
			this.label11.AutoSize = true;
			this.label11.Location = new System.Drawing.Point(8, 256);
			this.label11.Name = "label11";
			this.label11.Size = new System.Drawing.Size(68, 16);
			this.label11.TabIndex = 447;
			this.label11.Text = "Observación";
			// 
			// label6
			// 
			this.label6.AutoSize = true;
			this.label6.Location = new System.Drawing.Point(8, 17);
			this.label6.Name = "label6";
			this.label6.Size = new System.Drawing.Size(36, 16);
			this.label6.TabIndex = 444;
			this.label6.Text = "Fecha";
			this.label6.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// btnGuardar
			// 
			this.btnGuardar.Image = ((System.Drawing.Image)(resources.GetObject("btnGuardar.Image")));
			this.btnGuardar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnGuardar.Location = new System.Drawing.Point(792, 336);
			this.btnGuardar.Name = "btnGuardar";
			this.btnGuardar.Size = new System.Drawing.Size(76, 23);
			this.btnGuardar.TabIndex = 441;
			this.btnGuardar.Text = "&Guardar";
			this.btnGuardar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnGuardar.Click += new System.EventHandler(this.btnGuardar_Click);
			// 
			// btnSalir
			// 
			this.btnSalir.CausesValidation = false;
			this.btnSalir.Image = ((System.Drawing.Image)(resources.GetObject("btnSalir.Image")));
			this.btnSalir.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnSalir.Location = new System.Drawing.Point(880, 336);
			this.btnSalir.Name = "btnSalir";
			this.btnSalir.Size = new System.Drawing.Size(72, 23);
			this.btnSalir.TabIndex = 442;
			this.btnSalir.Text = "&Salir";
			this.btnSalir.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnSalir.Click += new System.EventHandler(this.btnSalir_Click);
			// 
			// groupBox1
			// 
			this.groupBox1.Location = new System.Drawing.Point(0, 368);
			this.groupBox1.Name = "groupBox1";
			this.groupBox1.Size = new System.Drawing.Size(1032, 8);
			this.groupBox1.TabIndex = 448;
			this.groupBox1.TabStop = false;
			// 
			// txtObservaciones
			// 
			this.txtObservaciones.AcceptsReturn = true;
			appearance1.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtObservaciones.Appearance = appearance1;
			this.txtObservaciones.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtObservaciones.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtObservaciones.Location = new System.Drawing.Point(72, 280);
			this.txtObservaciones.MaxLength = 300;
			this.txtObservaciones.Multiline = true;
			this.txtObservaciones.Name = "txtObservaciones";
			this.txtObservaciones.Size = new System.Drawing.Size(344, 88);
			this.txtObservaciones.TabIndex = 440;
			// 
			// txtidPersonal
			// 
			this.txtidPersonal.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.txtidPersonal.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtidPersonal.Enabled = false;
			this.txtidPersonal.Location = new System.Drawing.Point(48, 352);
			this.txtidPersonal.Name = "txtidPersonal";
			this.txtidPersonal.PromptChar = ' ';
			this.txtidPersonal.Size = new System.Drawing.Size(16, 21);
			this.txtidPersonal.TabIndex = 446;
			this.txtidPersonal.Visible = false;
			// 
			// dtFecha
			// 
			appearance2.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtFecha.Appearance = appearance2;
			this.dtFecha.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton1.Caption = "Today";
			this.dtFecha.DateButtons.Add(dateButton1);
			this.dtFecha.Enabled = false;
			this.dtFecha.Format = "dd/MM/yyyy HH:mm:ss";
			this.dtFecha.Location = new System.Drawing.Point(72, 15);
			this.dtFecha.Name = "dtFecha";
			this.dtFecha.NonAutoSizeHeight = 23;
			this.dtFecha.Size = new System.Drawing.Size(152, 21);
			this.dtFecha.SpinButtonsVisible = true;
			this.dtFecha.TabIndex = 438;
			this.dtFecha.Value = ((object)(resources.GetObject("dtFecha.Value")));
			// 
			// cmbGestionCredito
			// 
			this.cmbGestionCredito.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbGestionCredito.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbGestionCredito.DataSource = this.ultraDataSource1;
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn1.Width = 150;
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridColumn2.Width = 150;
			ultraGridBand1.Columns.AddRange(new object[] {
																										 ultraGridColumn1,
																										 ultraGridColumn2});
			this.cmbGestionCredito.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.cmbGestionCredito.DisplayMember = "Estado";
			this.cmbGestionCredito.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbGestionCredito.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbGestionCredito.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbGestionCredito.Location = new System.Drawing.Point(96, 120);
			this.cmbGestionCredito.MaxDropDownItems = 30;
			this.cmbGestionCredito.Name = "cmbGestionCredito";
			this.cmbGestionCredito.Size = new System.Drawing.Size(152, 21);
			this.cmbGestionCredito.TabIndex = 450;
			this.cmbGestionCredito.ValueMember = "idCbo_EstadoGestion";
			this.cmbGestionCredito.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cmbGestionCredito_KeyDown);
			this.cmbGestionCredito.ValueChanged += new System.EventHandler(this.cmbGestionCredito_ValueChanged);
			// 
			// ultraDataSource1
			// 
			ultraDataColumn1.DataType = typeof(int);
			this.ultraDataSource1.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn1,
																																 ultraDataColumn2});
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
			this.label1.Location = new System.Drawing.Point(8, 120);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(28, 16);
			this.label1.TabIndex = 451;
			this.label1.Text = "Dato";
			this.label1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// cmbGestion
			// 
			this.cmbGestion.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbGestion.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbGestion.DataSource = this.ultraDataSource2;
			ultraGridColumn3.Header.VisiblePosition = 0;
			ultraGridColumn3.Hidden = true;
			ultraGridColumn4.Header.Caption = "Gestión";
			ultraGridColumn4.Header.VisiblePosition = 1;
			ultraGridColumn4.Width = 180;
			ultraGridBand2.Columns.AddRange(new object[] {
																										 ultraGridColumn3,
																										 ultraGridColumn4});
			this.cmbGestion.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			this.cmbGestion.DisplayMember = "Resultado";
			this.cmbGestion.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbGestion.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbGestion.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbGestion.Location = new System.Drawing.Point(72, 184);
			this.cmbGestion.MaxDropDownItems = 30;
			this.cmbGestion.Name = "cmbGestion";
			this.cmbGestion.Size = new System.Drawing.Size(180, 21);
			this.cmbGestion.TabIndex = 452;
			this.cmbGestion.ValueMember = "idCbo_ResultadoGestion";
			this.cmbGestion.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cmbGestion_KeyDown);
			this.cmbGestion.VisibleChanged += new System.EventHandler(this.cmbGestion_VisibleChanged);
			this.cmbGestion.ValueChanged += new System.EventHandler(this.cmbGestion_ValueChanged);
			// 
			// ultraDataSource2
			// 
			ultraDataColumn3.DataType = typeof(int);
			this.ultraDataSource2.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn3,
																																 ultraDataColumn4});
			// 
			// dtFechaPago
			// 
			appearance3.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtFechaPago.Appearance = appearance3;
			this.dtFechaPago.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton2.Caption = "Today";
			this.dtFechaPago.DateButtons.Add(dateButton2);
			this.dtFechaPago.Enabled = false;
			this.dtFechaPago.Format = "dd/MM/yyyy HH:mm";
			this.dtFechaPago.Location = new System.Drawing.Point(336, 184);
			this.dtFechaPago.Name = "dtFechaPago";
			this.dtFechaPago.NonAutoSizeHeight = 23;
			this.dtFechaPago.Size = new System.Drawing.Size(104, 21);
			this.dtFechaPago.SpinButtonsVisible = true;
			this.dtFechaPago.TabIndex = 453;
			this.dtFechaPago.Value = ((object)(resources.GetObject("dtFechaPago.Value")));
			this.dtFechaPago.KeyDown += new System.Windows.Forms.KeyEventHandler(this.dtFechaPago_KeyDown);
			// 
			// lblFechaPago
			// 
			this.lblFechaPago.AutoSize = true;
			this.lblFechaPago.Enabled = false;
			this.lblFechaPago.Location = new System.Drawing.Point(264, 184);
			this.lblFechaPago.Name = "lblFechaPago";
			this.lblFechaPago.Size = new System.Drawing.Size(65, 16);
			this.lblFechaPago.TabIndex = 454;
			this.lblFechaPago.Text = "Fecha Pago";
			this.lblFechaPago.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// lblValor
			// 
			this.lblValor.AutoSize = true;
			this.lblValor.Enabled = false;
			this.lblValor.Location = new System.Drawing.Point(448, 184);
			this.lblValor.Name = "lblValor";
			this.lblValor.Size = new System.Drawing.Size(31, 16);
			this.lblValor.TabIndex = 840;
			this.lblValor.Text = "Valor";
			this.lblValor.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtValor
			// 
			appearance4.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtValor.Appearance = appearance4;
			this.txtValor.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtValor.Enabled = false;
			this.txtValor.FormatString = "#,##0.00";
			this.txtValor.Location = new System.Drawing.Point(488, 184);
			this.txtValor.MaskInput = "{LOC}nn,nnn,nnn.nn";
			this.txtValor.Name = "txtValor";
			this.txtValor.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtValor.PromptChar = ' ';
			this.txtValor.Size = new System.Drawing.Size(80, 21);
			this.txtValor.TabIndex = 841;
			this.txtValor.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtValor_KeyDown);
			// 
			// optAsignados
			// 
			appearance5.ForeColorDisabled = System.Drawing.Color.Black;
			this.optAsignados.Appearance = appearance5;
			this.optAsignados.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.optAsignados.CheckedIndex = 0;
			this.optAsignados.Enabled = false;
			this.optAsignados.ItemAppearance = appearance6;
			this.optAsignados.ItemOrigin = new System.Drawing.Point(10, 0);
			valueListItem1.DataValue = 1;
			valueListItem1.DisplayText = "Espera";
			valueListItem2.DataValue = 2;
			valueListItem2.DisplayText = "Cumplido";
			valueListItem3.DataValue = 3;
			valueListItem3.DisplayText = "Incumplido";
			this.optAsignados.Items.Add(valueListItem1);
			this.optAsignados.Items.Add(valueListItem2);
			this.optAsignados.Items.Add(valueListItem3);
			this.optAsignados.ItemSpacingVertical = 10;
			this.optAsignados.Location = new System.Drawing.Point(688, 264);
			this.optAsignados.Name = "optAsignados";
			this.optAsignados.Size = new System.Drawing.Size(240, 26);
			this.optAsignados.TabIndex = 842;
			this.optAsignados.Text = "Espera";
			this.optAsignados.Visible = false;
			// 
			// lblParentesco
			// 
			this.lblParentesco.AutoSize = true;
			this.lblParentesco.Location = new System.Drawing.Point(8, 216);
			this.lblParentesco.Name = "lblParentesco";
			this.lblParentesco.Size = new System.Drawing.Size(61, 16);
			this.lblParentesco.TabIndex = 843;
			this.lblParentesco.Text = "Parentesco";
			this.lblParentesco.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.lblParentesco.Visible = false;
			// 
			// cmbParentesco
			// 
			this.cmbParentesco.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbParentesco.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbParentesco.DataSource = this.ultraDataSource3;
			ultraGridColumn5.Header.VisiblePosition = 0;
			ultraGridColumn5.Hidden = true;
			ultraGridColumn6.Header.Caption = "Parentesco";
			ultraGridColumn6.Header.VisiblePosition = 1;
			ultraGridColumn6.Width = 180;
			ultraGridBand3.Columns.AddRange(new object[] {
																										 ultraGridColumn5,
																										 ultraGridColumn6});
			this.cmbParentesco.DisplayLayout.BandsSerializer.Add(ultraGridBand3);
			this.cmbParentesco.DisplayMember = "Nombre";
			this.cmbParentesco.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbParentesco.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbParentesco.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbParentesco.Location = new System.Drawing.Point(72, 216);
			this.cmbParentesco.MaxDropDownItems = 30;
			this.cmbParentesco.Name = "cmbParentesco";
			this.cmbParentesco.Size = new System.Drawing.Size(180, 21);
			this.cmbParentesco.TabIndex = 844;
			this.cmbParentesco.ValueMember = "idParentesco";
			this.cmbParentesco.Visible = false;
			this.cmbParentesco.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cmbParentesco_KeyDown);
			// 
			// ultraDataSource3
			// 
			this.ultraDataSource3.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn5,
																																 ultraDataColumn6});
			// 
			// btnTelefonia
			// 
			this.btnTelefonia.CausesValidation = false;
			this.btnTelefonia.Image = ((System.Drawing.Image)(resources.GetObject("btnTelefonia.Image")));
			this.btnTelefonia.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnTelefonia.Location = new System.Drawing.Point(424, 272);
			this.btnTelefonia.Name = "btnTelefonia";
			this.btnTelefonia.Size = new System.Drawing.Size(72, 23);
			this.btnTelefonia.TabIndex = 845;
			this.btnTelefonia.Text = "&Contacto";
			this.btnTelefonia.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnTelefonia.Click += new System.EventHandler(this.btnTelefonia_Click);
			// 
			// btnGestion
			// 
			this.btnGestion.CausesValidation = false;
			this.btnGestion.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.btnGestion.Image = ((System.Drawing.Image)(resources.GetObject("btnGestion.Image")));
			this.btnGestion.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnGestion.Location = new System.Drawing.Point(424, 304);
			this.btnGestion.Name = "btnGestion";
			this.btnGestion.Size = new System.Drawing.Size(72, 23);
			this.btnGestion.TabIndex = 846;
			this.btnGestion.Text = "&Gestión";
			this.btnGestion.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnGestion.Click += new System.EventHandler(this.btnGestion_Click);
			// 
			// txtidCliente
			// 
			this.txtidCliente.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.txtidCliente.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtidCliente.Enabled = false;
			this.txtidCliente.Location = new System.Drawing.Point(48, 328);
			this.txtidCliente.Name = "txtidCliente";
			this.txtidCliente.PromptChar = ' ';
			this.txtidCliente.Size = new System.Drawing.Size(16, 21);
			this.txtidCliente.TabIndex = 847;
			this.txtidCliente.Visible = false;
			// 
			// txtidCompra
			// 
			this.txtidCompra.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.txtidCompra.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtidCompra.Enabled = false;
			this.txtidCompra.Location = new System.Drawing.Point(48, 368);
			this.txtidCompra.Name = "txtidCompra";
			this.txtidCompra.PromptChar = ' ';
			this.txtidCompra.Size = new System.Drawing.Size(16, 21);
			this.txtidCompra.TabIndex = 848;
			this.txtidCompra.Visible = false;
			// 
			// timer1
			// 
			this.timer1.Interval = 1000;
			this.timer1.Tick += new System.EventHandler(this.timer1_Tick);
			// 
			// ultraDataSource4
			// 
			ultraDataColumn7.DataType = typeof(int);
			ultraDataColumn8.DataType = typeof(int);
			ultraDataColumn9.DataType = typeof(System.DateTime);
			ultraDataColumn14.DataType = typeof(System.Decimal);
			this.ultraDataSource4.Band.Columns.AddRange(new object[] {
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
																																 ultraDataColumn23});
			// 
			// lblContador
			// 
			this.lblContador.AutoSize = true;
			this.lblContador.Font = new System.Drawing.Font("Tahoma", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblContador.ForeColor = System.Drawing.Color.Firebrick;
			this.lblContador.Location = new System.Drawing.Point(16, 296);
			this.lblContador.Name = "lblContador";
			this.lblContador.Size = new System.Drawing.Size(0, 19);
			this.lblContador.TabIndex = 850;
			this.lblContador.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			this.lblContador.Visible = false;
			// 
			// groupBox2
			// 
			this.groupBox2.Location = new System.Drawing.Point(-16, 248);
			this.groupBox2.Name = "groupBox2";
			this.groupBox2.Size = new System.Drawing.Size(1032, 8);
			this.groupBox2.TabIndex = 851;
			this.groupBox2.TabStop = false;
			// 
			// timer2
			// 
			this.timer2.Enabled = true;
			this.timer2.Interval = 1000;
			this.timer2.Tick += new System.EventHandler(this.timer2_Tick);
			// 
			// lblDireccion
			// 
			this.lblDireccion.AutoSize = true;
			this.lblDireccion.Location = new System.Drawing.Point(264, 112);
			this.lblDireccion.Name = "lblDireccion";
			this.lblDireccion.Size = new System.Drawing.Size(52, 16);
			this.lblDireccion.TabIndex = 853;
			this.lblDireccion.Text = "Dirección";
			this.lblDireccion.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtDireccion
			// 
			this.txtDireccion.AcceptsReturn = true;
			appearance7.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtDireccion.Appearance = appearance7;
			this.txtDireccion.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtDireccion.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtDireccion.Enabled = false;
			this.txtDireccion.Location = new System.Drawing.Point(328, 88);
			this.txtDireccion.MaxLength = 100;
			this.txtDireccion.Multiline = true;
			this.txtDireccion.Name = "txtDireccion";
			this.txtDireccion.Size = new System.Drawing.Size(624, 48);
			this.txtDireccion.TabIndex = 854;
			// 
			// btnCobranza
			// 
			this.btnCobranza.CausesValidation = false;
			this.btnCobranza.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.btnCobranza.Image = ((System.Drawing.Image)(resources.GetObject("btnCobranza.Image")));
			this.btnCobranza.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnCobranza.Location = new System.Drawing.Point(424, 336);
			this.btnCobranza.Name = "btnCobranza";
			this.btnCobranza.Size = new System.Drawing.Size(72, 23);
			this.btnCobranza.TabIndex = 855;
			this.btnCobranza.Text = "&Cobranza";
			this.btnCobranza.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnCobranza.Click += new System.EventHandler(this.btnCobranza_Click);
			// 
			// txtEmail
			// 
			this.txtEmail.AcceptsReturn = true;
			appearance8.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtEmail.Appearance = appearance8;
			this.txtEmail.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtEmail.Enabled = false;
			this.txtEmail.Location = new System.Drawing.Point(336, 216);
			this.txtEmail.MaxLength = 100;
			this.txtEmail.Multiline = true;
			this.txtEmail.Name = "txtEmail";
			this.txtEmail.Size = new System.Drawing.Size(232, 21);
			this.txtEmail.TabIndex = 856;
			// 
			// txttelefono
			// 
			this.txttelefono.AcceptsReturn = true;
			appearance9.ForeColorDisabled = System.Drawing.Color.Black;
			this.txttelefono.Appearance = appearance9;
			this.txttelefono.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txttelefono.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txttelefono.Location = new System.Drawing.Point(72, 88);
			this.txttelefono.MaxLength = 11;
			this.txttelefono.Name = "txttelefono";
			this.txttelefono.Size = new System.Drawing.Size(152, 21);
			this.txttelefono.TabIndex = 858;
			this.txttelefono.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txttelefono_KeyDown);
			this.txttelefono.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txttelefono_KeyPress);
			// 
			// label2
			// 
			this.label2.AutoSize = true;
			this.label2.Location = new System.Drawing.Point(8, 88);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(48, 16);
			this.label2.TabIndex = 857;
			this.label2.Text = "Teléfono";
			this.label2.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtUbicacion
			// 
			this.txtUbicacion.AcceptsReturn = true;
			appearance10.BackColor = System.Drawing.SystemColors.Control;
			appearance10.BorderColor = System.Drawing.SystemColors.Control;
			appearance10.BorderColor3DBase = System.Drawing.SystemColors.Control;
			appearance10.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtUbicacion.Appearance = appearance10;
			this.txtUbicacion.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtUbicacion.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtUbicacion.Enabled = false;
			this.txtUbicacion.Location = new System.Drawing.Point(240, 6);
			this.txtUbicacion.MaxLength = 100;
			this.txtUbicacion.Multiline = true;
			this.txtUbicacion.Name = "txtUbicacion";
			this.txtUbicacion.Size = new System.Drawing.Size(710, 42);
			this.txtUbicacion.TabIndex = 860;
			// 
			// groupBox3
			// 
			this.groupBox3.Location = new System.Drawing.Point(-64, 48);
			this.groupBox3.Name = "groupBox3";
			this.groupBox3.Size = new System.Drawing.Size(1032, 8);
			this.groupBox3.TabIndex = 861;
			this.groupBox3.TabStop = false;
			// 
			// btnWhatssap
			// 
			this.btnWhatssap.CausesValidation = false;
			this.btnWhatssap.Enabled = false;
			this.btnWhatssap.Image = ((System.Drawing.Image)(resources.GetObject("btnWhatssap.Image")));
			this.btnWhatssap.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnWhatssap.Location = new System.Drawing.Point(880, 216);
			this.btnWhatssap.Name = "btnWhatssap";
			this.btnWhatssap.Size = new System.Drawing.Size(72, 23);
			this.btnWhatssap.TabIndex = 862;
			this.btnWhatssap.Text = "&Cargar";
			this.btnWhatssap.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnWhatssap.Click += new System.EventHandler(this.btnWhatssap_Click);
			// 
			// txtURL
			// 
			this.txtURL.AcceptsReturn = true;
			appearance11.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtURL.Appearance = appearance11;
			this.txtURL.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtURL.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtURL.Enabled = false;
			this.txtURL.Location = new System.Drawing.Point(720, 216);
			this.txtURL.MaxLength = 300;
			this.txtURL.Name = "txtURL";
			this.txtURL.Size = new System.Drawing.Size(152, 21);
			this.txtURL.TabIndex = 863;
			// 
			// txtDirectorio
			// 
			this.txtDirectorio.AcceptsReturn = true;
			appearance12.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtDirectorio.Appearance = appearance12;
			this.txtDirectorio.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtDirectorio.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtDirectorio.Location = new System.Drawing.Point(600, 264);
			this.txtDirectorio.MaxLength = 5000;
			this.txtDirectorio.Name = "txtDirectorio";
			this.txtDirectorio.Size = new System.Drawing.Size(16, 21);
			this.txtDirectorio.TabIndex = 864;
			this.txtDirectorio.Visible = false;
			// 
			// txtDirectorioarchivo
			// 
			this.txtDirectorioarchivo.AcceptsReturn = true;
			appearance13.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtDirectorioarchivo.Appearance = appearance13;
			this.txtDirectorioarchivo.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtDirectorioarchivo.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtDirectorioarchivo.Location = new System.Drawing.Point(648, 264);
			this.txtDirectorioarchivo.MaxLength = 5000;
			this.txtDirectorioarchivo.Name = "txtDirectorioarchivo";
			this.txtDirectorioarchivo.Size = new System.Drawing.Size(16, 21);
			this.txtDirectorioarchivo.TabIndex = 865;
			this.txtDirectorioarchivo.Visible = false;
			// 
			// txtDirectorioFinal
			// 
			this.txtDirectorioFinal.AcceptsReturn = true;
			appearance14.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtDirectorioFinal.Appearance = appearance14;
			this.txtDirectorioFinal.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtDirectorioFinal.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtDirectorioFinal.Location = new System.Drawing.Point(624, 264);
			this.txtDirectorioFinal.MaxLength = 5000;
			this.txtDirectorioFinal.Name = "txtDirectorioFinal";
			this.txtDirectorioFinal.Size = new System.Drawing.Size(16, 21);
			this.txtDirectorioFinal.TabIndex = 866;
			this.txtDirectorioFinal.Visible = false;
			// 
			// label4
			// 
			this.label4.AutoSize = true;
			this.label4.Enabled = false;
			this.label4.Location = new System.Drawing.Point(664, 216);
			this.label4.Name = "label4";
			this.label4.Size = new System.Drawing.Size(42, 16);
			this.label4.TabIndex = 867;
			this.label4.Text = "Archivo";
			this.label4.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// lblMail
			// 
			this.lblMail.AutoSize = true;
			this.lblMail.Location = new System.Drawing.Point(264, 216);
			this.lblMail.Name = "lblMail";
			this.lblMail.Size = new System.Drawing.Size(71, 16);
			this.lblMail.TabIndex = 869;
			this.lblMail.Text = "Email Cliente";
			this.lblMail.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// uGridDocumentos
			// 
			this.uGridDocumentos.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.uGridDocumentos.Cursor = System.Windows.Forms.Cursors.Default;
			this.uGridDocumentos.DataSource = this.ultraDataSource4;
			appearance15.BackColor = System.Drawing.Color.White;
			appearance15.BackColor2 = System.Drawing.Color.White;
			this.uGridDocumentos.DisplayLayout.Appearance = appearance15;
			ultraGridColumn7.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn7.Header.VisiblePosition = 0;
			ultraGridColumn7.Hidden = true;
			ultraGridColumn8.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn8.Header.VisiblePosition = 1;
			ultraGridColumn8.Hidden = true;
			ultraGridColumn9.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn9.Format = "yyy/MM/dd HH:mm:ss";
			ultraGridColumn9.Header.VisiblePosition = 2;
			ultraGridColumn9.Width = 100;
			ultraGridColumn10.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn10.Header.Caption = "Operador/Cobrador/Vendedor";
			ultraGridColumn10.Header.VisiblePosition = 4;
			ultraGridColumn10.Width = 220;
			ultraGridColumn11.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn11.Header.VisiblePosition = 6;
			ultraGridColumn11.Width = 100;
			ultraGridColumn12.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn12.Header.Caption = "Gestión Estado";
			ultraGridColumn12.Header.VisiblePosition = 7;
			ultraGridColumn12.Width = 100;
			ultraGridColumn13.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn13.Header.Caption = "Fecha Pago";
			ultraGridColumn13.Header.VisiblePosition = 8;
			ultraGridColumn13.Width = 90;
			ultraGridColumn14.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance16.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn14.CellAppearance = appearance16;
			ultraGridColumn14.Formula = "#,##0.00";
			ultraGridColumn14.Header.VisiblePosition = 9;
			ultraGridColumn14.Width = 80;
			ultraGridColumn15.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn15.Header.Caption = "Estado Compromiso";
			ultraGridColumn15.Header.VisiblePosition = 10;
			ultraGridColumn15.Width = 120;
			ultraGridColumn16.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn16.Header.Caption = "Parentesco";
			ultraGridColumn16.Header.VisiblePosition = 11;
			ultraGridColumn16.Width = 160;
			ultraGridColumn17.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn17.Header.VisiblePosition = 12;
			ultraGridColumn17.Width = 100;
			ultraGridColumn18.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn18.Header.VisiblePosition = 13;
			ultraGridColumn18.Width = 250;
			ultraGridColumn19.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn19.Header.VisiblePosition = 5;
			ultraGridColumn19.Width = 250;
			ultraGridColumn20.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn20.Header.VisiblePosition = 14;
			ultraGridColumn20.Width = 250;
			ultraGridColumn21.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn21.Header.VisiblePosition = 15;
			ultraGridColumn22.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn22.Header.VisiblePosition = 16;
			ultraGridColumn22.Width = 250;
			ultraGridColumn23.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn23.Header.VisiblePosition = 3;
			ultraGridColumn24.ButtonDisplayStyle = Infragistics.Win.UltraWinGrid.ButtonDisplayStyle.Always;
			appearance17.Image = ((object)(resources.GetObject("appearance17.Image")));
			appearance17.ImageHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn24.CellButtonAppearance = appearance17;
			ultraGridColumn24.Header.Caption = "...";
			ultraGridColumn24.Header.VisiblePosition = 17;
			ultraGridColumn24.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.Button;
			ultraGridColumn24.Width = 30;
			ultraGridBand4.Columns.AddRange(new object[] {
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
																										 ultraGridColumn24});
			ultraGridBand4.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
			ultraGridBand4.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			ultraGridBand4.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			ultraGridBand4.SummaryFooterCaption = "";
			this.uGridDocumentos.DisplayLayout.BandsSerializer.Add(ultraGridBand4);
			this.uGridDocumentos.DisplayLayout.GroupByBox.Hidden = true;
			appearance18.ForeColor = System.Drawing.Color.Black;
			appearance18.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridDocumentos.DisplayLayout.Override.ActiveRowAppearance = appearance18;
			appearance19.BackColor = System.Drawing.Color.Transparent;
			this.uGridDocumentos.DisplayLayout.Override.CardAreaAppearance = appearance19;
			appearance20.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance20.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance20.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance20.FontData.BoldAsString = "True";
			appearance20.FontData.Name = "Arial";
			appearance20.FontData.SizeInPoints = 8F;
			appearance20.ForeColor = System.Drawing.Color.White;
			appearance20.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridDocumentos.DisplayLayout.Override.HeaderAppearance = appearance20;
			this.uGridDocumentos.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti;
			appearance21.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance21.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance21.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridDocumentos.DisplayLayout.Override.RowAlternateAppearance = appearance21;
			appearance22.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance22.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance22.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridDocumentos.DisplayLayout.Override.RowSelectorAppearance = appearance22;
			appearance23.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance23.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance23.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridDocumentos.DisplayLayout.Override.SelectedRowAppearance = appearance23;
			this.uGridDocumentos.DisplayLayout.ViewStyleBand = Infragistics.Win.UltraWinGrid.ViewStyleBand.OutlookGroupBy;
			this.uGridDocumentos.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridDocumentos.Location = new System.Drawing.Point(8, 384);
			this.uGridDocumentos.Name = "uGridDocumentos";
			this.uGridDocumentos.Size = new System.Drawing.Size(952, 176);
			this.uGridDocumentos.TabIndex = 849;
			this.uGridDocumentos.ClickCellButton += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.uGridDocumentos_ClickCellButton);
			// 
			// txtEmailOpeSup
			// 
			this.txtEmailOpeSup.AcceptsReturn = true;
			appearance24.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtEmailOpeSup.Appearance = appearance24;
			this.txtEmailOpeSup.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtEmailOpeSup.Enabled = false;
			this.txtEmailOpeSup.Location = new System.Drawing.Point(720, 184);
			this.txtEmailOpeSup.MaxLength = 100;
			this.txtEmailOpeSup.Multiline = true;
			this.txtEmailOpeSup.Name = "txtEmailOpeSup";
			this.txtEmailOpeSup.Size = new System.Drawing.Size(152, 21);
			this.txtEmailOpeSup.TabIndex = 870;
			// 
			// label5
			// 
			this.label5.AutoSize = true;
			this.label5.Enabled = false;
			this.label5.Location = new System.Drawing.Point(576, 184);
			this.label5.Name = "label5";
			this.label5.Size = new System.Drawing.Size(134, 16);
			this.label5.TabIndex = 871;
			this.label5.Text = "Email Operador/Cobrador";
			this.label5.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// lblVendedor
			// 
			this.lblVendedor.AutoSize = true;
			this.lblVendedor.Location = new System.Drawing.Point(8, 64);
			this.lblVendedor.Name = "lblVendedor";
			this.lblVendedor.Size = new System.Drawing.Size(53, 16);
			this.lblVendedor.TabIndex = 873;
			this.lblVendedor.Text = "Vendedor";
			this.lblVendedor.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// cmbVendedor
			// 
			appearance25.FontData.Name = "Tahoma";
			appearance25.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbVendedor.Appearance = appearance25;
			this.cmbVendedor.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbVendedor.Cursor = System.Windows.Forms.Cursors.Default;
			ultraGridColumn25.Header.VisiblePosition = 0;
			ultraGridColumn25.Hidden = true;
			ultraGridColumn26.Header.VisiblePosition = 1;
			ultraGridColumn26.Width = 335;
			ultraGridColumn27.Header.VisiblePosition = 2;
			ultraGridBand5.Columns.AddRange(new object[] {
																										 ultraGridColumn25,
																										 ultraGridColumn26,
																										 ultraGridColumn27});
			this.cmbVendedor.DisplayLayout.BandsSerializer.Add(ultraGridBand5);
			this.cmbVendedor.DisplayMember = "Nombre";
			this.cmbVendedor.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbVendedor.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbVendedor.Enabled = false;
			this.cmbVendedor.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbVendedor.Location = new System.Drawing.Point(72, 58);
			this.cmbVendedor.Name = "cmbVendedor";
			this.cmbVendedor.Size = new System.Drawing.Size(250, 22);
			this.cmbVendedor.TabIndex = 872;
			this.cmbVendedor.ValueMember = "idPersonal";
			this.cmbVendedor.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cmbVendedor_KeyDown);
			// 
			// btnTelefoniacli
			// 
			this.btnTelefoniacli.CausesValidation = false;
			this.btnTelefoniacli.Image = ((System.Drawing.Image)(resources.GetObject("btnTelefoniacli.Image")));
			this.btnTelefoniacli.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnTelefoniacli.Location = new System.Drawing.Point(504, 272);
			this.btnTelefoniacli.Name = "btnTelefoniacli";
			this.btnTelefoniacli.Size = new System.Drawing.Size(72, 23);
			this.btnTelefoniacli.TabIndex = 874;
			this.btnTelefoniacli.Text = "&Historico";
			this.btnTelefoniacli.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnTelefoniacli.Visible = false;
			this.btnTelefoniacli.Click += new System.EventHandler(this.btnTelefoniacli_Click);
			// 
			// uWeb
			// 
			this.uWeb.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.uWeb.Cursor = System.Windows.Forms.Cursors.Default;
			this.uWeb.DataSource = this.ultraDataSource5;
			appearance26.BackColor = System.Drawing.Color.White;
			appearance26.BackColor2 = System.Drawing.Color.White;
			this.uWeb.DisplayLayout.Appearance = appearance26;
			ultraGridColumn28.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn28.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn28.Format = "yyy/MM/dd HH:mm:ss";
			ultraGridColumn28.Header.VisiblePosition = 0;
			ultraGridColumn28.Width = 100;
			ultraGridColumn29.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn29.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn29.Header.VisiblePosition = 1;
			ultraGridColumn29.Width = 100;
			ultraGridColumn30.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn30.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn30.Header.VisiblePosition = 3;
			ultraGridColumn30.Width = 250;
			ultraGridColumn31.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn31.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn31.Header.VisiblePosition = 4;
			ultraGridColumn31.Width = 180;
			ultraGridColumn32.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn32.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn32.Header.VisiblePosition = 5;
			ultraGridColumn32.Width = 250;
			ultraGridColumn33.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn33.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn33.Header.VisiblePosition = 2;
			ultraGridColumn33.Hidden = true;
			ultraGridColumn34.ButtonDisplayStyle = Infragistics.Win.UltraWinGrid.ButtonDisplayStyle.Always;
			appearance27.Image = ((object)(resources.GetObject("appearance27.Image")));
			appearance27.ImageHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn34.CellButtonAppearance = appearance27;
			ultraGridColumn34.Header.Caption = "...";
			ultraGridColumn34.Header.VisiblePosition = 6;
			ultraGridColumn34.Hidden = true;
			ultraGridColumn34.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.Button;
			ultraGridColumn34.Width = 30;
			ultraGridBand6.Columns.AddRange(new object[] {
																										 ultraGridColumn28,
																										 ultraGridColumn29,
																										 ultraGridColumn30,
																										 ultraGridColumn31,
																										 ultraGridColumn32,
																										 ultraGridColumn33,
																										 ultraGridColumn34});
			ultraGridBand6.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
			ultraGridBand6.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			ultraGridBand6.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			ultraGridBand6.SummaryFooterCaption = "";
			this.uWeb.DisplayLayout.BandsSerializer.Add(ultraGridBand6);
			this.uWeb.DisplayLayout.GroupByBox.Hidden = true;
			appearance28.ForeColor = System.Drawing.Color.Black;
			appearance28.ForeColorDisabled = System.Drawing.Color.Black;
			this.uWeb.DisplayLayout.Override.ActiveRowAppearance = appearance28;
			appearance29.BackColor = System.Drawing.Color.Transparent;
			this.uWeb.DisplayLayout.Override.CardAreaAppearance = appearance29;
			appearance30.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance30.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance30.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance30.FontData.BoldAsString = "True";
			appearance30.FontData.Name = "Arial";
			appearance30.FontData.SizeInPoints = 8F;
			appearance30.ForeColor = System.Drawing.Color.White;
			appearance30.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uWeb.DisplayLayout.Override.HeaderAppearance = appearance30;
			this.uWeb.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti;
			appearance31.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance31.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance31.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uWeb.DisplayLayout.Override.RowAlternateAppearance = appearance31;
			appearance32.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance32.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance32.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uWeb.DisplayLayout.Override.RowSelectorAppearance = appearance32;
			appearance33.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance33.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance33.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uWeb.DisplayLayout.Override.SelectedRowAppearance = appearance33;
			this.uWeb.DisplayLayout.ViewStyleBand = Infragistics.Win.UltraWinGrid.ViewStyleBand.OutlookGroupBy;
			this.uWeb.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uWeb.Location = new System.Drawing.Point(8, 384);
			this.uWeb.Name = "uWeb";
			this.uWeb.Size = new System.Drawing.Size(952, 176);
			this.uWeb.TabIndex = 875;
			// 
			// ultraDataSource5
			// 
			ultraDataColumn24.DataType = typeof(System.DateTime);
			this.ultraDataSource5.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn24,
																																 ultraDataColumn25,
																																 ultraDataColumn26,
																																 ultraDataColumn27,
																																 ultraDataColumn28,
																																 ultraDataColumn29});
			// 
			// cmbTipocontacto
			// 
			this.cmbTipocontacto.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbTipocontacto.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbTipocontacto.DataSource = this.ultraDataSource6;
			ultraGridColumn35.Header.VisiblePosition = 0;
			ultraGridColumn35.Hidden = true;
			ultraGridColumn36.Header.VisiblePosition = 1;
			ultraGridColumn36.Width = 150;
			ultraGridBand7.Columns.AddRange(new object[] {
																										 ultraGridColumn35,
																										 ultraGridColumn36});
			this.cmbTipocontacto.DisplayLayout.BandsSerializer.Add(ultraGridBand7);
			this.cmbTipocontacto.DisplayMember = "Estado";
			this.cmbTipocontacto.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbTipocontacto.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbTipocontacto.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbTipocontacto.Location = new System.Drawing.Point(88, 152);
			this.cmbTipocontacto.MaxDropDownItems = 30;
			this.cmbTipocontacto.Name = "cmbTipocontacto";
			this.cmbTipocontacto.Size = new System.Drawing.Size(160, 21);
			this.cmbTipocontacto.TabIndex = 876;
			this.cmbTipocontacto.ValueMember = "idCbo_EstadosTipocontacto";
			this.cmbTipocontacto.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cmbTipocontacto_KeyDown);
			this.cmbTipocontacto.ValueChanged += new System.EventHandler(this.cmbTipocontacto_ValueChanged);
			// 
			// ultraDataSource6
			// 
			ultraDataColumn30.DataType = typeof(int);
			this.ultraDataSource6.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn30,
																																 ultraDataColumn31});
			// 
			// label7
			// 
			this.label7.AutoSize = true;
			this.label7.Location = new System.Drawing.Point(8, 152);
			this.label7.Name = "label7";
			this.label7.Size = new System.Drawing.Size(72, 16);
			this.label7.TabIndex = 877;
			this.label7.Text = "Tipo contacto";
			this.label7.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// frmCob_InformeNotCompro
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.BackColor = System.Drawing.SystemColors.Control;
			this.ClientSize = new System.Drawing.Size(968, 574);
			this.Controls.Add(this.label7);
			this.Controls.Add(this.lblVendedor);
			this.Controls.Add(this.label5);
			this.Controls.Add(this.lblMail);
			this.Controls.Add(this.label4);
			this.Controls.Add(this.label2);
			this.Controls.Add(this.lblDireccion);
			this.Controls.Add(this.lblParentesco);
			this.Controls.Add(this.lblValor);
			this.Controls.Add(this.lblFechaPago);
			this.Controls.Add(this.label1);
			this.Controls.Add(this.label3);
			this.Controls.Add(this.label11);
			this.Controls.Add(this.label6);
			this.Controls.Add(this.cmbTipocontacto);
			this.Controls.Add(this.uWeb);
			this.Controls.Add(this.btnTelefoniacli);
			this.Controls.Add(this.cmbVendedor);
			this.Controls.Add(this.txtEmailOpeSup);
			this.Controls.Add(this.txtDirectorioFinal);
			this.Controls.Add(this.txtDirectorioarchivo);
			this.Controls.Add(this.txtDirectorio);
			this.Controls.Add(this.txtURL);
			this.Controls.Add(this.btnWhatssap);
			this.Controls.Add(this.groupBox3);
			this.Controls.Add(this.txtUbicacion);
			this.Controls.Add(this.txttelefono);
			this.Controls.Add(this.txtEmail);
			this.Controls.Add(this.btnCobranza);
			this.Controls.Add(this.txtDireccion);
			this.Controls.Add(this.groupBox2);
			this.Controls.Add(this.lblContador);
			this.Controls.Add(this.uGridDocumentos);
			this.Controls.Add(this.txtidCompra);
			this.Controls.Add(this.txtidCliente);
			this.Controls.Add(this.btnGestion);
			this.Controls.Add(this.btnTelefonia);
			this.Controls.Add(this.cmbParentesco);
			this.Controls.Add(this.optAsignados);
			this.Controls.Add(this.txtValor);
			this.Controls.Add(this.dtFechaPago);
			this.Controls.Add(this.cmbGestion);
			this.Controls.Add(this.cmbGestionCredito);
			this.Controls.Add(this.btnGuardar);
			this.Controls.Add(this.btnSalir);
			this.Controls.Add(this.groupBox1);
			this.Controls.Add(this.txtObservaciones);
			this.Controls.Add(this.txtidPersonal);
			this.Controls.Add(this.dtFecha);
			this.KeyPreview = true;
			this.MaximizeBox = false;
			this.MinimizeBox = false;
			this.Name = "frmCob_InformeNotCompro";
			this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
			this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.frmCob_InformeNotCompro_KeyDown);
			this.Load += new System.EventHandler(this.frmCob_InformeNotCompro_Load);
			((System.ComponentModel.ISupportInitialize)(this.txtObservaciones)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtidPersonal)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFecha)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbGestionCredito)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbGestion)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFechaPago)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtValor)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.optAsignados)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbParentesco)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource3)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtidCliente)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtidCompra)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource4)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDireccion)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtEmail)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txttelefono)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtUbicacion)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtURL)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDirectorio)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDirectorioarchivo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDirectorioFinal)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridDocumentos)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtEmailOpeSup)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbVendedor)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.uWeb)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource5)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbTipocontacto)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource6)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void frmCob_InformeNotCompro_Load(object sender, System.EventArgs e)
		{
			if (bWeb)
			{
				//			frmGestionCobranza idcompra , false , false

				if(iOperadorCobradorVendedor == 2)
				{
					this.cmbVendedor.DataSource = Funcion.dvProcedimiento(cdsSeteoF, string.Format("Exec [PersonalTransferenciasRango] '{0}', {1}, {2}, {3}", 
						Convert.ToDateTime(DateTime.Today).ToString("yyyyMMdd"), iBodega, MenuLatinium.IdUsuario, iBodega));
					//				this.cmbVendedor.DataSource = FuncionesProcedimientos.Lista_Vendedores(Convert.ToDateTime(DateTime.Today), iBodega, cdsSeteoF);
					this.cmbVendedor.Enabled =  true;
					this.btnTelefoniacli.Visible = true;
					this.cmbVendedor.Focus ();
				}
				this.txttelefono.Focus();
				this.txtidCompra.Value = idCompras;
				this.txtidPersonal.Value = idPersonal;

				idCliente = Funcion.iEscalarSQL(cdsSeteoF, string.Format("Select idCliente from Compra where idcompra = {0}",(int)this.txtidCompra.Value));	

				if(iOperadorCobradorVendedor != 10)
				{
					this.txtUbicacion.Text  = Funcion.sEscalarSQL(cdsSeteoF, string.Format("select dbo.CobranzasUbicacion ( '{0}')",sRuc));	
				
				}
				else
				{
					this.txtUbicacion.Text  = Funcion.sEscalarSQL(cdsSeteoF, string.Format("EXEC  ConsultaProveedorCXC  '{0}'",sRuc));	
					idPersonal  = Funcion.iEscalarSQL(cdsSeteoF, string.Format("EXEC  ConsultaUsuarioCXC  {0}",idPersonal));	
					this.btnGestion.Visible = false;
					this.btnCobranza.Visible = false;
					this.btnTelefonia.Visible = false;
					this.txtidPersonal.Value = idPersonal;
				
				}
				this.dtFecha.Value = DateTime.Today;
				this.cmbGestionCredito.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Select idCbo_EstadoGestion, Estado from Cbo_EstadosGestion where activo = 1");

				#region Correos Operadores y cobradores

				this.txtEmailOpeSup.Text =  Funcion.sEscalarSQL(cdsSeteoF, string.Format(" exec Cbo_RetornaEmailOperadorCobrador {0}, {1}, {2}", idPersonal,iOperadorCobradorVendedor, iBodega));
     
				#endregion Correos Operadores y cobradores
				string sSQLg = string.Format("Exec ConsultaCbo_GestionesDeCobranzas {0}", (int) this.txtidCompra.Value);
				FuncionesProcedimientos.dsGeneral(sSQLg, this.uGridDocumentos);
				if(idEntidadFinanciera == 2) this.btnCobranza.Enabled = false;
				this.Text = " * Cliente: " + sCliente + "  * Cedula : " + sRuc + "*"  ;
				this.uGridDocumentos.Visible = true;
				this.uWeb.Visible = false;
			}
			else
			{
				#region Desabilita Web
				this.btnCobranza.Visible = false;
				this.btnGestion.Visible = false;
				this.btnTelefonia.Visible = false;
				this.btnTelefoniacli.Visible = false;
				this.cmbParentesco.Visible = false;
				this.cmbGestion.Visible = false;
				this.txtURL.Visible = false;
				this.btnWhatssap.Visible = false;
				this.txtEmailOpeSup.Visible = false;
				this.txtValor.Visible = false;
				this.dtFechaPago.Visible = false;
				this.txtDireccion.Visible = false;
				this.label4.Visible = false;
				this.label5.Visible = false;
				this.lblValor.Visible = false;
				this.lblFechaPago.Visible = false;
				this.lblDireccion.Text ="Producto";
				//								this.txtDireccion.Enabled = true;
				this.txtDireccion.Visible = true;
				this.lblParentesco.Visible = false;
				this.label3.Visible = false;
				#endregion Desabilita Web
				this.txtUbicacion.Text  = Funcion.sEscalarSQL(cdsSeteoF, string.Format(" select ' '+observacion + '   '+ 'Codigo Postal :' + codigo_postal + '    Telefono : ' + telefono+ '      '+ 'Dirección Entrega :' + direccion_entrega+ ''  from web_cabecera_pedido where numero_pedido = {0} ", iPedido));	
				this.txtDireccion.Text  = Funcion.sEscalarSQL(cdsSeteoF, string.Format("   select replace(replace (replace ((   select replace(replace (replace ((select isnull ( Articulo, cod_arcticulo) as _ from  web_detalle_pedido w left outer join Articulo a on a.Codigo = w.cod_arcticulo where numero_pedido = {0} for XML path ('')),'<',''),'>',''),'_','')),'<',''),'>',''),'_','') ", iPedido));	
				this.cmbGestionCredito.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "select  Estado as idCbo_EstadoGestion ,Nombre as Estado from Estado_WEB order by nombre ");
				this.txtEmail.Text = sEmail;

				string sSQLg = string.Format("Exec [ConsultaGestionVentasWeb] {0}", iPedido );
				FuncionesProcedimientos.dsGeneral(sSQLg, this.uWeb);

				this.uGridDocumentos.Visible = false;
				this.uWeb.Visible = true;
				this.Text = " * Cliente: " + sClienteW + "  * Cedula : " + sCedulaW  ;

			}
		   
			this.txttelefono.Focus();
		}

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);		
		}

		private void cmbGestionCredito_ValueChanged(object sender, System.EventArgs e)
		{
			if (this.cmbGestionCredito.ActiveRow != null)
			{
				if( bWeb)
				{
					Vaciar ();
					this.cmbTipocontacto.Value = System.DBNull.Value;
					this.cmbGestion.Value = System.DBNull.Value;
					this.cmbTipocontacto.DataSource = Funcion.dvProcedimiento(cdsSeteoF, string.Format("select idCbo_EstadosTipocontacto, Estado from Cbo_EstadosTipocontacto where idCbo_EstadoGestion = {0}", (int) this.cmbGestionCredito.Value));			
				}
			}
		}
		private void Vaciar ()
		{
			carpeta ="";
			this.dtFechaPago.Value = System.DBNull.Value;
			this.txtValor.Value = 0.00m;
			this.optAsignados.Value = 1;
			this.cmbParentesco.Value = System.DBNull.Value;

	
			this.dtFechaPago.Enabled = false;
			this.txtValor.Enabled = false;
			this.optAsignados.Enabled = false;
			this.cmbParentesco.Visible = false;
			this.btnWhatssap.Enabled = false;
			this.txtDireccion.Enabled = false;
			this.lblParentesco.Visible = false;

			this.txtURL.Text ="";
			this.txtURL.Enabled = false;
			this.txtDirectorioarchivo.Text = "";
			this.txtDirectorioFinal.Text = "";
			this.txtDirectorio.Text ="";
			this.txtEmail.Text = "";
			this.txtEmail.Enabled = false;

		}
		private void cmbGestion_ValueChanged(object sender, System.EventArgs e)
		{
			int iContWha = this.uGridDocumentos.Rows.Count ;
			if (this.cmbGestion.ActiveRow != null)
			{
				if (bWeb)
				{
					Vaciar ();
					#region Compromiso De Pago

					if((int)this.cmbGestion.Value == 4 || (int)this.cmbGestion.Value == 39 || (int)this.cmbGestion.Value == 54 || (int)this.cmbGestion.Value == 69 )
					{
						this.dtFechaPago.Value = DateTime.Today;
						this.dtFechaPago.CalendarInfo.MinDate = DateTime.Today;
						this.lblFechaPago.Enabled = true;
						this.txtValor.Enabled = true;
						this.dtFechaPago.Enabled = true;
						this.optAsignados.Enabled = true;
					}
					else
					{
						this.lblFechaPago.Enabled = false;
						this.lblValor.Enabled = false;
						this.txtValor.Enabled = false;
						this.dtFechaPago.Enabled = false;
						this.optAsignados.Enabled = false;
				  
					}

					#endregion Compromiso De Pago

					#region Pariente Domicilio

					if((int)this.cmbGestion.Value == 5 || (int)this.cmbGestion.Value == 6 )
					{
						this.cmbParentesco.Visible = true;
						this.lblParentesco.Visible = true;
						if((int)this.cmbGestion.Value == 5)
						{
							this.cmbParentesco.DataSource = Funcion.dvProcedimiento(cdsSeteoF, string.Format("SELECT idParentesco, Nombre FROM Cre_Parentesco where idParentesco not in (20,21)"));
							this.txtDireccion.Enabled = true;
						}
						else
						{
							this.cmbParentesco.DataSource = Funcion.dvProcedimiento(cdsSeteoF, string.Format("SELECT idParentesco, Nombre FROM Cre_Parentesco where idParentesco  in (20,21)"));
							this.txtDireccion.Enabled = true;
						}
					}
					else
					{
						this.cmbParentesco.Visible = false;
						this.lblParentesco.Visible = false;
					}
					#endregion Pariente Domicilio

					if( (int) this.cmbGestion.Value ==12 )
					{
						if (iContWha >= 3)
						{
							this.txtURL.Enabled = true;
							this.btnWhatssap.Enabled = true;
							this.txtEmail.Enabled = false;
							this.txtEmail.Text = "";
						}
						else
						{
							MessageBox.Show(string.Format("Deben existir minimo 3 gestiones antes del WHATSAPP"), "Point Technology", MessageBoxButtons.OK,MessageBoxIcon.Stop);
							return;

						}
					}
				
					if( (int) this.cmbGestion.Value ==13 )
					{
						this.txtURL.Enabled = true;
						this.btnWhatssap.Enabled = true;
						this.txtEmail.Enabled = false;
						this.txtEmail.Text = "";
					}
				
					if( (int) this.cmbGestion.Value ==14 )
					{
						this.btnWhatssap.Enabled = false;
						this.txtURL.Text = "";
						this.txtEmail.Enabled = true;

					}
				
				}
			}
		}

		private void cmbGestion_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter)
			{
				if (this.cmbGestion.ActiveRow != null)
				{
					if ((int)this.cmbGestion.Value == 4)
					{
						this.dtFechaPago.Focus();
					}
					else
					{
						if ((int)this.cmbGestion.Value == 5 || (int)this.cmbGestion.Value == 6)
						{
							this.cmbParentesco.Focus();
						}
						else
						{
							this.txtObservaciones.Focus();
						}
		
					}
				}
			}
		}

		private void dtFechaPago_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter)this.txtValor.Focus ();
		}

		private void txtValor_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter)this.txtObservaciones.Focus ();
		}

		private void cmbParentesco_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter)this.txtObservaciones.Focus ();
		}

		private void cmbGestionCredito_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter)this.cmbTipocontacto.Focus ();
		}

		private void btnGuardar_Click(object sender, System.EventArgs e)
		{
			if (bWeb)
			{
				#region validacion

				DateTime dtFechaPago = DateTime.Parse("01/01/2000");;
				decimal dValor = 0.00m ;
				int iParentesco = 0;
				int EstadoG = 0;

	
				if (idCliente < 1) return;
				if (iOperadorCobradorVendedor == 0 || iOperadorCobradorVendedor == 2 )
				{
					if (!Validacion.vbTexto(this.txttelefono, 9, 10, "Telefono")) return;
				}
				if (this.txttelefono.Text.ToString().Trim ().Length == 9)
				{
					if (!Validacion.vbValidaTelefono(this.txttelefono, 1, false)) return;
				
				}
				if (this.txttelefono.Text.ToString().Trim ().Length == 10)
				{
					if (!Validacion.vbValidaTelefono(this.txttelefono, 2, false)) return;
				}


				if ( this.cmbParentesco.ActiveRow != null) iParentesco  = ( int) this.cmbParentesco.Value;

				if (!Validacion.vbFechaEnDocumentos(this.dtFecha, "Seleccione la fecha ", false, "", cdsSeteoF)) return;

				if (!Validacion.vbComboVacio(this.cmbGestionCredito, "Seleccione el estado de la Gestión")) return;

				if (!Validacion.vbComboVacio(this.cmbGestion, "Seleccione la Gestión")) return;
				if (!Validacion.vbComboVacio(this.cmbTipocontacto, "Seleccione la tipo de contacto")) return;
      
		
				if ((int)this.cmbGestion.Value == 4 || (int)this.cmbGestion.Value == 5 || (int)this.cmbGestion.Value == 6  || (int)this.cmbGestion.Value == 39   || (int)this.cmbGestion.Value == 54   || (int)this.cmbGestion.Value == 55  || (int)this.cmbGestion.Value == 69  )
				{
					//if ((int)this.cmbGestion.Value == 4)
					//{
						if (!Validacion.vbFechaEnDocumentos(this.dtFechaPago, "Seleccione la Fecha de Pago ", false, "", cdsSeteoF)) return;
						if (!Validacion.vbCampoDecimalVacio(this.txtValor, "Ingrese el Valor de Cobro ",1,2000)) return;
						if((int)this.optAsignados.Value < 0)
						{
							MessageBox.Show("Seleccione el Estado de laGestion", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
							this.optAsignados.Focus();
							return;
						}

						dtFechaPago = (DateTime) this.dtFechaPago.Value;
						dValor = Convert.ToDecimal( this.txtValor.Value);
						EstadoG = (int)this.optAsignados.Value;

					//}
					if ((int)this.cmbGestion.Value == 5 || (int)this.cmbGestion.Value == 6)
					{
						if (!Validacion.vbComboVacio(this.cmbParentesco, "Seleccione el Parentesco ")) return;
					
						if (!Validacion.vbTexto(this.txtDireccion, 1, 100, "Dirección")) return;
					}
				}
				if (!Validacion.vbTexto(this.txtObservaciones, 20, 300, "Observaciones")) return;
				if( (int) this.cmbGestion.Value == 12)
				{
					if (!Validacion.vbTexto(this.txtURL, 1, 300, " Archivo")) return;
				
				}
				if( (int) this.cmbGestion.Value == 13)
				{
					if (!Validacion.vbTexto(this.txtURL, 1, 300, " Archivo")) return;
				
				}
				#endregion Validacion

				if ((int) this.cmbGestion.Value == 14)
				{		
	
					if (this.txtEmail.Text.ToString().Trim().Length == 0   )
					{
						MessageBox.Show(string.Format("Ingrese El Email al cual se va enviar la notificación "), "Point Technology", MessageBoxButtons.OK,MessageBoxIcon.Stop);
						return;
					}
					if (!Validacion.bEMail(this.txtEmail.Text.ToString())) return;

					if (this.txtEmailOpeSup.Text.ToString().Trim().Length == 0   )
					{
						MessageBox.Show(string.Format("Por favor ingrese un Email de respaldo. Comuniquese con el Administrador"), "Point Technology", MessageBoxButtons.OK,MessageBoxIcon.Stop);
						return;
					}
					if (!Validacion.bEMail(this.txtEmailOpeSup.Text.ToString())) return;
				}
				if( (int) this.cmbGestion.Value == 12)
				{
					Directorio (carpeta, false);
				}
				if( (int) this.cmbGestion.Value == 13)
				{
					Directorio (carpeta, false);
				}
				if(iOperadorCobradorVendedor == 2)
				{
					if (!Validacion.vbComboVacio(this.cmbVendedor, "Seleccione el Vendedor ")) return;

					this.txtidPersonal.Value =  (int)this.cmbVendedor.Value;
				}

				string sSQL = string.Format("Exec GuardaCbo_GestionesDeCobranzas 0, {0},{1},'{2}', {3}, {4}, '{5}', {6}, {7}, {8}, '{9}', {10}, '{11}', '{12}','{13}', '{14}', {15}, {16}, {17}",
					(int)this.txtidCompra.Value, (int)this.txtidPersonal.Value , Convert.ToDateTime(this.dtFecha.Value).ToString("yyyyMMdd HH:mm:ss"), (int)this.cmbGestionCredito.Value, 
					(int)this.cmbGestion.Value, Convert.ToDateTime(dtFechaPago ).ToString("yyyyMMdd"), dValor, EstadoG, iParentesco, this.txtObservaciones.Text.ToString (),
					idCliente,this.txtDireccion.Text.ToString(), this.txttelefono.Text.ToString (), this.txtDirectorioFinal.Text, this.txtEmail.Text.ToString (), iOperadorCobradorVendedor, idCbo_GestorDeCobranzas,
					(int)this.cmbTipocontacto.Value);
				int iRetGestionesDeCobranzas =Funcion.iEscalarSQL(cdsSeteoF,sSQL);
			
				string sSQLg = string.Format("Exec ConsultaCbo_GestionesDeCobranzas {0}", (int) this.txtidCompra.Value);
				FuncionesProcedimientos.dsGeneral(sSQLg, this.uGridDocumentos);
     
				if ((int) this.cmbGestion.Value == 14)
				{
					string sSQLMail = string.Format(@"Exec MailCbo_GestionesDeCobranzas {0}, {1}, '{2}','{3}',{4},'{5}'", 
						idCliente, (int) this.txtidCompra.Value, this.txtEmail.Text.ToString (),Convert.ToDateTime(this.dtFecha.Value).ToString("yyyyMMdd HH:mm:ss"),
						(int)this.txtidPersonal.Value,  this.txtEmailOpeSup.Text );
					string sMensaje = Funcion.sEscalarSQL(cdsSeteoF, sSQLMail, true);
				}

			}
			else
			{
				if (!Validacion.vbComboVacio(this.cmbGestionCredito, "Seleccione el estado de la Gestión")) return;
				if (!Validacion.vbTexto(this.txtObservaciones, 20, 300, " Observaciones")) return;
				/* 1 es vendedor web */
				string sSQL = string.Format("Exec [GuardaGestionesWeb] 0, {0},'{1}','{2}', {3}, '{4}', {5}, {6}",
					                                
					iPedido, 
					Convert.ToDateTime(this.dtFecha.Value).ToString("yyyyMMdd HH:mm:ss"), 
					this.txttelefono.Text, 
					1,
					this.txtObservaciones.Text.ToString (),
					MenuLatinium.IdUsuario, 
					(int)this.cmbGestionCredito.Value );
				int web = Funcion.iEscalarSQL(cdsSeteoF,sSQL);


				string sSQLg = string.Format("Exec [ConsultaGestionVentasWeb] {0}", iPedido );
				FuncionesProcedimientos.dsGeneral(sSQLg, this.uWeb);
			}
			MessageBox.Show("Registro Guardado Correctamente ", "Informacion", MessageBoxButtons.OK, MessageBoxIcon.Information);  

			Limpia_Campos ();
		}

		private void Limpia_Campos ()
		{
			this.cmbGestionCredito.Value = System.DBNull.Value;
			this.cmbParentesco.Value = System.DBNull.Value;
			this.txtDireccion.Enabled = false;
			this.txtDireccion.Text = "";
			this.txtObservaciones.Text = "";
			this.dtFechaPago.Enabled = false;
			this.txtValor.Enabled = false;
			this.optAsignados.Enabled = false;
			this.txtEmail.Text = "";
			this.txttelefono.Text ="";
			this.lblParentesco.Visible = false;
			this.cmbGestion.Value = System.DBNull.Value;
			this.dtFechaPago.Value = System.DBNull.Value;
			this.txtValor.Value = 0.00m;
			this.optAsignados.Value = 1;
			this.cmbParentesco.Visible = false;
			this.cmbParentesco.Value =  System.DBNull.Value;
			this.txtDirectorio.Text ="";
			this.txtDirectorioarchivo.Text ="";
			this.txtDirectorioFinal.Text ="";
			this.txtURL.Text ="";
			this.txtURL.Enabled = false;
			this.btnWhatssap.Enabled = false;
			this.txtEmail.Enabled = false;
		}

		private void btnGestion_Click(object sender, System.EventArgs e)
		{
			using (frmCred_ReporteVerificacionTelefonica Gestion = new frmCred_ReporteVerificacionTelefonica(sRuc, true ))
			{
				if (DialogResult.OK == Gestion.ShowDialog( ))
				{
	
				}
			}
		}

		private void timer1_Tick(object sender, System.EventArgs e)
		{
			this.dtFecha.Value = DateTime.Now;
		}

		private void btnTelefonia_Click(object sender, System.EventArgs e)
		{
			this.txtidCliente.Value = idCliente;
			using (frmCbo_GCTelefonosNuevos Gestion = new frmCbo_GCTelefonosNuevos((int)this.txtidCliente.Value))
			{
				if (DialogResult.OK == Gestion.ShowDialog())
				{
								
				}
			}
		}

		private void btnSalir_Click(object sender, System.EventArgs e)
		{
			this.DialogResult = DialogResult.No;
		}

		private void timer2_Tick(object sender, System.EventArgs e)
		{
			this.dtFecha.Value = DateTime.Now.ToLongTimeString ();
		}

		private void frmCob_InformeNotCompro_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Escape) this.DialogResult = DialogResult.No;
		}

		

		private void txttelefono_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter)this.cmbGestionCredito.Focus ();
		}

		private void btnCobranza_Click(object sender, System.EventArgs e)
		{
			using (frmCobranzas Cobranza = new frmCobranzas(sRuc, true ))
			{
				if (DialogResult.OK == Cobranza.ShowDialog( ))
				{
					
				}

			}
		}
		string carpeta ;
		private void btnWhatssap_Click(object sender, System.EventArgs e)
		{
			//		 carpeta = @"C:\Documents and Settings\Infoelec\Escritorio\"+ sRuc+sNumeroFactura;  /* aqui creamos el nombre de la carpeta con el #cedual el cliente*/
			carpeta = @"\\192.168.1.10\Proyectos\Latinium\ed\"+ sRuc+sNumeroFactura; 
			try
			{
				if (Directory.Exists(carpeta))
				{
					Directorio (carpeta , true);
				}
				else
				{
					Directory.CreateDirectory(carpeta);
					Directorio (carpeta , true);
				}
			}
			catch (Exception ex)
			{
				MessageBox.Show ("Error: "+ ex.Message);

			}
		}
		private void Directorio (string sCarpeta, bool bArchivoGuarda)
		{ 
			string URL ="" ;
			string sArchivo="";

			if (bArchivoGuarda)
			{
				try 
				{
					OpenFileDialog buscar = new OpenFileDialog ();
					if (buscar.ShowDialog () == DialogResult.OK)
					{
						this.txtDirectorioarchivo.Text = buscar.FileName;
						this.txtURL.Text =  Funcion.sEscalarSQL(cdsSeteoF, string.Format(" SELECT RIGHT('{0}', CHARINDEX('{1}', REVERSE('{0}')) - 1) AS Archivo",this.txtDirectorioarchivo.Text.ToString (),@"\"));
						//						this.txtURL.Text = sArchivo;
					}
				}
				catch (Exception ex)
				{
					MessageBox.Show ("Error: "+ ex.Message);
					return;

				}
			}
			else
			{
				try 
				{
					sCarpeta  = (sCarpeta + @"\" + this.txtURL.Text);
					System.IO.File.Copy (  this.txtDirectorioarchivo.Text ,sCarpeta, false);
					MessageBox.Show(string.Format("El archivo '{0}' se a Guardado en la carpeta '{1}'  ", this.txtURL.Text,sCarpeta), "Validacion", MessageBoxButtons.OK, MessageBoxIcon.Information);
					this.txtDirectorioFinal.Text = sCarpeta;
				}
				catch (Exception ex)
				{
					MessageBox.Show ("Error: "+ ex.Message);
					return;
				}
			}
		}

		private void uGridDocumentos_ClickCellButton(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			try 
			{
				
				if (e.Cell.Column.ToString() == "Ir")
				{
					if (e.Cell.Row.Cells["DireccionUrlWSP"].Value.ToString().Length == 0)
					{
						MessageBox.Show(string.Format("No un existe Archivo Guardado"), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);

						return;
					}
					else
					{
						Process.Start ("explorer.exe",e.Cell.Row.Cells["DireccionUrlWSP"].Value.ToString() );
						//FileStream archivo = new FileStream ( e.Cell.Row.Cells["DireccionUrlWSP"].Value.ToString() , FileMode.Open ); 
					}
				}
			}
			catch (Exception ex)
			{
				MessageBox.Show ("Error: "+ ex.Message);
				return;
			}
		}

		private void btnTelefoniacli_Click(object sender, System.EventArgs e)
		{
			using (frmGestionCobranza Gestion = new frmGestionCobranza(idCompras, false, false))
			{
				if (DialogResult.OK == Gestion.ShowDialog())
				{
								
				}
			}
		}

		private void cmbVendedor_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter)this.txttelefono.Focus ();
		}

		private void cmbGestion_VisibleChanged(object sender, System.EventArgs e)
		{
		
		}

		private void txttelefono_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			validar.Solonumeros(e);
		}

		
		private void cmbTipocontacto_ValueChanged(object sender, System.EventArgs e)
		{
			if (this.cmbTipocontacto.ActiveRow != null)
			{
				if( bWeb)
				{
					Vaciar ();
					this.cmbGestion.Value = System.DBNull.Value;
					this.cmbGestion.DataSource = Funcion.dvProcedimiento(cdsSeteoF, string.Format("select idCbo_ResultadoGestion, Resultado from Cbo_ResultadoGestion where idCbo_EstadosTipocontacto = {0}", (int) this.cmbTipocontacto.Value));	
				}
			}
		}

		private void cmbTipocontacto_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter)this.cmbGestion.Focus ();
		}
	
		class validar 
		{
			public static void Solonumeros (KeyPressEventArgs Nu)
			{
				if(char.IsDigit(Nu.KeyChar))
				{
			
					Nu.Handled = false;
			
				}
				else if (char.IsControl(Nu.KeyChar))
				{
			
					Nu.Handled = false;			} 
				else
				{
					Nu.Handled= true;
				}
			}
	
	
		}
	}
}
