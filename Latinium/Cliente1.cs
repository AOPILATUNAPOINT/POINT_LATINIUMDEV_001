using System;
using System.Data.SqlClient;
using C1.Data;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de Cliente.
	/// </summary>
	public class Cliente1 : DevExpress.XtraEditors.XtraForm
	{
		private System.Windows.Forms.Label label1;
		private System.Windows.Forms.Label label3;
		private System.Windows.Forms.Label label4;
		private System.Windows.Forms.Label label5;
		private System.Windows.Forms.Label label6;
		private System.Windows.Forms.Label label7;
		private System.Windows.Forms.Label label8;
		private System.Windows.Forms.Label label11;
		private System.Windows.Forms.Label label13;
		private System.Windows.Forms.Label label17;
		private System.Windows.Forms.Label label18;
		private C1.Data.C1DataSet cdsCliente;
		private System.Windows.Forms.ErrorProvider errorProvider;
		private Infragistics.Win.Misc.UltraButton bCiudad;
		private Infragistics.Win.Misc.UltraButton btGrupoCliente;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbCiudad;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbGrupo;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbTipoRuc;
		private System.Windows.Forms.Label label22;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo cmbFecha;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtNombre;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtDireccion;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtTelefono;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor ultraTextEditor3;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtRuc;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtEMail;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor ultraTextEditor10;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtCodigo;
		public Infragistics.Win.UltraWinEditors.UltraTextEditor txtIdCliente;
		private C1.Data.C1DataSet cdsSeteoF;
		private Infragistics.Win.Misc.UltraButton btVerifica;
		private Infragistics.Win.UltraWinTabControl.UltraTabControl tbBasico;
		private Infragistics.Win.UltraWinTabControl.UltraTabSharedControlsPage ultraTabSharedControlsPage1;
		private Infragistics.Win.UltraWinTabControl.UltraTabPageControl ultraTabPageControl1;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor ultraCheckEditor1;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor ultraTextEditor5;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtCelular;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbPrecio;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbSector;
		private Infragistics.Win.Misc.UltraButton btSector;
		private Infragistics.Win.Misc.UltraButton btPrecio;
		private System.Windows.Forms.Label label10;
		private System.Windows.Forms.Label label9;
		private System.Windows.Forms.Label label2;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor spnDescCredito;
		private System.Windows.Forms.Label label24;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor spnDescuento;
		private System.Windows.Forms.Label label23;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtValorCredito;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor ultraNumericEditor1;
		private System.Windows.Forms.Label label14;
		private System.Windows.Forms.Label label12;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtAutFactura;
		private System.Windows.Forms.Label label15;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtAutImprenta;
		private System.Windows.Forms.Label label16;
		private System.Windows.Forms.Label label20;
		private System.Windows.Forms.Label label21;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbCobrador;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbVendedor;
		private Infragistics.Win.Misc.UltraButton btVendedor;
		private Infragistics.Win.UltraWinTabControl.UltraTabPageControl tpComplemento;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbDepende;
		private System.Windows.Forms.Label label25;
		private Infragistics.Win.Misc.UltraButton btDepende;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor ultraTextEditor1;
		private System.Windows.Forms.Label label26;
		private System.Windows.Forms.Label label27;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor ultraTextEditor2;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbPorcIvaRet;
		private System.Windows.Forms.Label label29;
		private Infragistics.Win.UltraWinEditors.UltraDateTimeEditor cmbFechaCaduca;
		private System.Windows.Forms.Label label31;
		private System.Windows.Forms.Label label32;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbFormaPago;
		private System.Windows.Forms.Label label33;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbCredTribut;
		private System.Windows.Forms.Label label34;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkContribuyenteEsp;
		private Infragistics.Win.UltraWinTabControl.UltraTabPageControl ultraTabPageControl2;
		private System.Windows.Forms.Label label35;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor ultraTextEditor7;
		private System.Windows.Forms.Label label36;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor ultraTextEditor8;
		private System.Windows.Forms.Label label37;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor ultraTextEditor9;
		private System.Windows.Forms.Label label38;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor ultraTextEditor11;
		private System.Windows.Forms.Label label39;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor ultraTextEditor12;
		private System.Windows.Forms.Label label40;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor ultraTextEditor13;
		private System.Windows.Forms.Label label41;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor ultraTextEditor14;
		private System.Windows.Forms.Label label42;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor ultraTextEditor15;
		private System.Windows.Forms.Label label43;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor ultraTextEditor16;
		private System.Windows.Forms.Label label44;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor ultraTextEditor17;
		private System.Windows.Forms.Label label45;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor ultraTextEditor18;
		private System.Windows.Forms.Label label46;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor ultraTextEditor19;
		private System.Windows.Forms.Label label47;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor ultraTextEditor20;
		private System.Windows.Forms.GroupBox groupBox1;
		private System.Windows.Forms.GroupBox groupBox2;
		private Infragistics.Win.Misc.UltraButton btCRMClientes;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkDesCompras;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor ultraCheckEditor2;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtAutRetencion;
		private System.Windows.Forms.Label lblAutRetencion;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbClientePublicidad;
		private System.Windows.Forms.Label label30;
		private Infragistics.Win.Misc.UltraButton btPublicidad;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkEntrega;
		private Infragistics.Win.Misc.UltraButton btAuditoria;
		private System.Windows.Forms.Label label48;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor ultraCheckEditor3;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor ultraNumericEditor2;
		private System.Windows.Forms.Label label49;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor ultraNumericEditor3;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor ultraTextEditor4;
		private System.Windows.Forms.Label label50;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor ultraTextEditor21;
		private System.Windows.Forms.Label label51;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkPublica;
		private Infragistics.Win.Misc.UltraButton btBanco;
		private System.Windows.Forms.Label label53;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtSerie;
		private Infragistics.Win.UltraWinTabControl.UltraTabPageControl ultraTabPageControl3;
		private Infragistics.Win.UltraWinGrid.UltraGrid ultraGrid1;
		private Infragistics.Win.Misc.UltraButton btListadoPendiente;
		private Infragistics.Win.Misc.UltraButton btVencido;
		private Infragistics.Win.Misc.UltraButton btFactura;
		private Infragistics.Win.Misc.UltraButton btTotales;
		private Infragistics.Win.UltraWinTabControl.UltraTabPageControl ultraTabPageControl4;
		private Infragistics.Win.UltraWinGrid.UltraGrid grdItems;
		private Infragistics.Win.Misc.UltraButton btListadoItems;
		private System.Windows.Forms.Label label19;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbBanco;
		private System.Windows.Forms.ImageList imageList;
		private System.Windows.Forms.ToolTip toolTip1;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkProspecto;
		private Infragistics.Win.UltraWinTabControl.UltraTabPageControl ultraTabPageControl5;
		private Infragistics.Win.UltraWinGrid.UltraGrid ultraGrid2;
		private C1.Data.C1DataSet cdsContacto;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor ultraTextEditor6;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo ultraCalendarCombo1;
		private System.Windows.Forms.Label label28;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkEspecial;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtPais;
		private Infragistics.Win.Misc.UltraButton bPais;
		private C1.Data.C1DataSet cdsPais;
		private Infragistics.Win.Misc.UltraButton btExcel;
		private Infragistics.Win.UltraWinGrid.ExcelExport.UltraGridExcelExporter ultraGridExcelExporter1;
		private Infragistics.Win.Misc.UltraButton ultraButton1;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbPais;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource1;
		private DevExpress.XtraEditors.RadioGroup radioGroup1;
		private BarraDatoSQL.BarraDatoSQL barraDato1;
		private System.ComponentModel.IContainer components;

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
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idClientePais");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Pais");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idClientePais");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Pais");
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton1 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idClientePublicidad");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Medio");
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance11 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand3 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idTipoPrecio");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Precio");
			Infragistics.Win.Appearance appearance12 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand4 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idSector");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Sector");
			Infragistics.Win.Appearance appearance13 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance14 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance15 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand5 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCiudad");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Ciudad");
			Infragistics.Win.Appearance appearance16 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand6 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn11 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idGrupoCliente");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn12 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Grupo");
			Infragistics.Win.Appearance appearance17 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand7 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn13 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idTipoRuc");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn14 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("TipoRuc");
			Infragistics.Win.Appearance appearance18 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance19 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance20 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance21 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance22 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance23 = new Infragistics.Win.Appearance();
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(Cliente1));
			Infragistics.Win.Appearance appearance24 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance25 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton2 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.Appearance appearance26 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance27 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance28 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance29 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance30 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance31 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand8 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn15 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCredTributario");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn16 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CredTributario");
			Infragistics.Win.Appearance appearance32 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand9 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn17 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idFormaPago");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn18 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FormaPago");
			Infragistics.Win.Appearance appearance33 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance34 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance35 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance36 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand10 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn19 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idRetIva");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn20 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Porcentaje");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn21 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Detalle");
			Infragistics.Win.Appearance appearance37 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand11 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn22 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCliente");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn23 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.Appearance appearance38 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand12 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn24 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPersonal");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn25 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.Appearance appearance39 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand13 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn26 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPersonal");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn27 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.Appearance appearance40 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance41 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance42 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance43 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance44 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance45 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance46 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance47 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand14 = new Infragistics.Win.UltraWinGrid.UltraGridBand("ClienteContacto", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn28 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idClienteContacto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn29 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCliente");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn30 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Contacto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn31 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Telefono");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn32 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Detalle");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn33 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Email");
			Infragistics.Win.Appearance appearance48 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance49 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance50 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance51 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance52 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand15 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn34 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idBanco");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn35 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
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
			Infragistics.Win.Appearance appearance63 = new Infragistics.Win.Appearance();
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
			Infragistics.Win.UltraWinTabControl.UltraTab ultraTab1 = new Infragistics.Win.UltraWinTabControl.UltraTab();
			Infragistics.Win.UltraWinTabControl.UltraTab ultraTab2 = new Infragistics.Win.UltraWinTabControl.UltraTab();
			Infragistics.Win.UltraWinTabControl.UltraTab ultraTab3 = new Infragistics.Win.UltraWinTabControl.UltraTab();
			Infragistics.Win.UltraWinTabControl.UltraTab ultraTab4 = new Infragistics.Win.UltraWinTabControl.UltraTab();
			Infragistics.Win.UltraWinTabControl.UltraTab ultraTab5 = new Infragistics.Win.UltraWinTabControl.UltraTab();
			Infragistics.Win.UltraWinTabControl.UltraTab ultraTab6 = new Infragistics.Win.UltraWinTabControl.UltraTab();
			Infragistics.Win.Appearance appearance79 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance80 = new Infragistics.Win.Appearance();
			this.ultraTabPageControl1 = new Infragistics.Win.UltraWinTabControl.UltraTabPageControl();
			this.txtPais = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.cdsCliente = new C1.Data.C1DataSet();
			this.cmbPais = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.ultraDataSource1 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.bPais = new Infragistics.Win.Misc.UltraButton();
			this.ultraTextEditor6 = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.ultraCalendarCombo1 = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.label28 = new System.Windows.Forms.Label();
			this.btPublicidad = new Infragistics.Win.Misc.UltraButton();
			this.cmbClientePublicidad = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.label30 = new System.Windows.Forms.Label();
			this.label27 = new System.Windows.Forms.Label();
			this.ultraTextEditor2 = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label26 = new System.Windows.Forms.Label();
			this.ultraTextEditor1 = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.ultraCheckEditor1 = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.ultraTextEditor5 = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtCelular = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.cmbPrecio = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.cmbSector = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.btSector = new Infragistics.Win.Misc.UltraButton();
			this.btPrecio = new Infragistics.Win.Misc.UltraButton();
			this.label19 = new System.Windows.Forms.Label();
			this.label10 = new System.Windows.Forms.Label();
			this.label9 = new System.Windows.Forms.Label();
			this.label2 = new System.Windows.Forms.Label();
			this.label5 = new System.Windows.Forms.Label();
			this.txtEMail = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.ultraTextEditor10 = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label7 = new System.Windows.Forms.Label();
			this.label8 = new System.Windows.Forms.Label();
			this.btGrupoCliente = new Infragistics.Win.Misc.UltraButton();
			this.cmbCiudad = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.label1 = new System.Windows.Forms.Label();
			this.cmbGrupo = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.label11 = new System.Windows.Forms.Label();
			this.cmbTipoRuc = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.bCiudad = new Infragistics.Win.Misc.UltraButton();
			this.label13 = new System.Windows.Forms.Label();
			this.label4 = new System.Windows.Forms.Label();
			this.label22 = new System.Windows.Forms.Label();
			this.label3 = new System.Windows.Forms.Label();
			this.label17 = new System.Windows.Forms.Label();
			this.txtNombre = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtDireccion = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtTelefono = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.ultraTextEditor3 = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtRuc = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.btVerifica = new Infragistics.Win.Misc.UltraButton();
			this.txtCodigo = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.cmbFecha = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.label18 = new System.Windows.Forms.Label();
			this.label6 = new System.Windows.Forms.Label();
			this.chkProspecto = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.tpComplemento = new Infragistics.Win.UltraWinTabControl.UltraTabPageControl();
			this.radioGroup1 = new DevExpress.XtraEditors.RadioGroup();
			this.chkEspecial = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.ultraTextEditor21 = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label51 = new System.Windows.Forms.Label();
			this.ultraTextEditor4 = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label50 = new System.Windows.Forms.Label();
			this.chkContribuyenteEsp = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.label34 = new System.Windows.Forms.Label();
			this.cmbCredTribut = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.label33 = new System.Windows.Forms.Label();
			this.cmbFormaPago = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.txtSerie = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label32 = new System.Windows.Forms.Label();
			this.txtAutRetencion = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label31 = new System.Windows.Forms.Label();
			this.cmbFechaCaduca = new Infragistics.Win.UltraWinEditors.UltraDateTimeEditor();
			this.lblAutRetencion = new System.Windows.Forms.Label();
			this.label29 = new System.Windows.Forms.Label();
			this.cmbPorcIvaRet = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.btDepende = new Infragistics.Win.Misc.UltraButton();
			this.label25 = new System.Windows.Forms.Label();
			this.cmbDepende = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.label20 = new System.Windows.Forms.Label();
			this.label21 = new System.Windows.Forms.Label();
			this.cmbCobrador = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.cmbVendedor = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.btVendedor = new Infragistics.Win.Misc.UltraButton();
			this.txtAutImprenta = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label16 = new System.Windows.Forms.Label();
			this.txtAutFactura = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label15 = new System.Windows.Forms.Label();
			this.spnDescCredito = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label24 = new System.Windows.Forms.Label();
			this.spnDescuento = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label23 = new System.Windows.Forms.Label();
			this.txtValorCredito = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.ultraNumericEditor1 = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label14 = new System.Windows.Forms.Label();
			this.label12 = new System.Windows.Forms.Label();
			this.chkPublica = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.ultraTabPageControl5 = new Infragistics.Win.UltraWinTabControl.UltraTabPageControl();
			this.ultraGrid2 = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.cdsContacto = new C1.Data.C1DataSet();
			this.ultraTabPageControl2 = new Infragistics.Win.UltraWinTabControl.UltraTabPageControl();
			this.btBanco = new Infragistics.Win.Misc.UltraButton();
			this.cmbBanco = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.label53 = new System.Windows.Forms.Label();
			this.ultraNumericEditor3 = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label49 = new System.Windows.Forms.Label();
			this.ultraNumericEditor2 = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.ultraCheckEditor3 = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.label48 = new System.Windows.Forms.Label();
			this.chkEntrega = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.ultraCheckEditor2 = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.chkDesCompras = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.groupBox2 = new System.Windows.Forms.GroupBox();
			this.ultraTextEditor17 = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label41 = new System.Windows.Forms.Label();
			this.label40 = new System.Windows.Forms.Label();
			this.label44 = new System.Windows.Forms.Label();
			this.ultraTextEditor14 = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.ultraTextEditor13 = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label36 = new System.Windows.Forms.Label();
			this.ultraTextEditor8 = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label47 = new System.Windows.Forms.Label();
			this.ultraTextEditor20 = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label45 = new System.Windows.Forms.Label();
			this.ultraTextEditor18 = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label46 = new System.Windows.Forms.Label();
			this.ultraTextEditor19 = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label42 = new System.Windows.Forms.Label();
			this.ultraTextEditor15 = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label35 = new System.Windows.Forms.Label();
			this.ultraTextEditor7 = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.groupBox1 = new System.Windows.Forms.GroupBox();
			this.label37 = new System.Windows.Forms.Label();
			this.ultraTextEditor9 = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.ultraTextEditor12 = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label39 = new System.Windows.Forms.Label();
			this.label43 = new System.Windows.Forms.Label();
			this.label38 = new System.Windows.Forms.Label();
			this.ultraTextEditor16 = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.ultraTextEditor11 = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.ultraTabPageControl3 = new Infragistics.Win.UltraWinTabControl.UltraTabPageControl();
			this.btExcel = new Infragistics.Win.Misc.UltraButton();
			this.btTotales = new Infragistics.Win.Misc.UltraButton();
			this.btFactura = new Infragistics.Win.Misc.UltraButton();
			this.btVencido = new Infragistics.Win.Misc.UltraButton();
			this.btListadoPendiente = new Infragistics.Win.Misc.UltraButton();
			this.ultraGrid1 = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraTabPageControl4 = new Infragistics.Win.UltraWinTabControl.UltraTabPageControl();
			this.ultraButton1 = new Infragistics.Win.Misc.UltraButton();
			this.btListadoItems = new Infragistics.Win.Misc.UltraButton();
			this.grdItems = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.errorProvider = new System.Windows.Forms.ErrorProvider();
			this.txtIdCliente = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.tbBasico = new Infragistics.Win.UltraWinTabControl.UltraTabControl();
			this.ultraTabSharedControlsPage1 = new Infragistics.Win.UltraWinTabControl.UltraTabSharedControlsPage();
			this.btCRMClientes = new Infragistics.Win.Misc.UltraButton();
			this.imageList = new System.Windows.Forms.ImageList(this.components);
			this.btAuditoria = new Infragistics.Win.Misc.UltraButton();
			this.toolTip1 = new System.Windows.Forms.ToolTip(this.components);
			this.cdsPais = new C1.Data.C1DataSet();
			this.ultraGridExcelExporter1 = new Infragistics.Win.UltraWinGrid.ExcelExport.UltraGridExcelExporter();
			this.barraDato1 = new BarraDatoSQL.BarraDatoSQL();
			this.ultraTabPageControl1.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.txtPais)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsCliente)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbPais)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraTextEditor6)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraCalendarCombo1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbClientePublicidad)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraTextEditor2)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraTextEditor1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraTextEditor5)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCelular)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbPrecio)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbSector)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtEMail)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraTextEditor10)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCiudad)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbGrupo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbTipoRuc)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNombre)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDireccion)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtTelefono)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraTextEditor3)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtRuc)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCodigo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbFecha)).BeginInit();
			this.tpComplemento.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.radioGroup1.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraTextEditor21)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraTextEditor4)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCredTribut)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbFormaPago)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtSerie)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtAutRetencion)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbFechaCaduca)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbPorcIvaRet)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbDepende)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCobrador)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbVendedor)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtAutImprenta)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtAutFactura)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.spnDescCredito)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.spnDescuento)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtValorCredito)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraNumericEditor1)).BeginInit();
			this.ultraTabPageControl5.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid2)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsContacto)).BeginInit();
			this.ultraTabPageControl2.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.cmbBanco)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraNumericEditor3)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraNumericEditor2)).BeginInit();
			this.groupBox2.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.ultraTextEditor17)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraTextEditor14)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraTextEditor13)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraTextEditor8)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraTextEditor20)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraTextEditor18)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraTextEditor19)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraTextEditor15)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraTextEditor7)).BeginInit();
			this.groupBox1.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.ultraTextEditor9)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraTextEditor12)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraTextEditor16)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraTextEditor11)).BeginInit();
			this.ultraTabPageControl3.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).BeginInit();
			this.ultraTabPageControl4.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.grdItems)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdCliente)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.tbBasico)).BeginInit();
			this.tbBasico.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.cdsPais)).BeginInit();
			this.SuspendLayout();
			// 
			// ultraTabPageControl1
			// 
			this.ultraTabPageControl1.Controls.Add(this.txtPais);
			this.ultraTabPageControl1.Controls.Add(this.cmbPais);
			this.ultraTabPageControl1.Controls.Add(this.bPais);
			this.ultraTabPageControl1.Controls.Add(this.ultraTextEditor6);
			this.ultraTabPageControl1.Controls.Add(this.ultraCalendarCombo1);
			this.ultraTabPageControl1.Controls.Add(this.label28);
			this.ultraTabPageControl1.Controls.Add(this.btPublicidad);
			this.ultraTabPageControl1.Controls.Add(this.cmbClientePublicidad);
			this.ultraTabPageControl1.Controls.Add(this.label30);
			this.ultraTabPageControl1.Controls.Add(this.label27);
			this.ultraTabPageControl1.Controls.Add(this.ultraTextEditor2);
			this.ultraTabPageControl1.Controls.Add(this.label26);
			this.ultraTabPageControl1.Controls.Add(this.ultraTextEditor1);
			this.ultraTabPageControl1.Controls.Add(this.ultraCheckEditor1);
			this.ultraTabPageControl1.Controls.Add(this.ultraTextEditor5);
			this.ultraTabPageControl1.Controls.Add(this.txtCelular);
			this.ultraTabPageControl1.Controls.Add(this.cmbPrecio);
			this.ultraTabPageControl1.Controls.Add(this.cmbSector);
			this.ultraTabPageControl1.Controls.Add(this.btSector);
			this.ultraTabPageControl1.Controls.Add(this.btPrecio);
			this.ultraTabPageControl1.Controls.Add(this.label19);
			this.ultraTabPageControl1.Controls.Add(this.label10);
			this.ultraTabPageControl1.Controls.Add(this.label9);
			this.ultraTabPageControl1.Controls.Add(this.label2);
			this.ultraTabPageControl1.Controls.Add(this.label5);
			this.ultraTabPageControl1.Controls.Add(this.txtEMail);
			this.ultraTabPageControl1.Controls.Add(this.ultraTextEditor10);
			this.ultraTabPageControl1.Controls.Add(this.label7);
			this.ultraTabPageControl1.Controls.Add(this.label8);
			this.ultraTabPageControl1.Controls.Add(this.btGrupoCliente);
			this.ultraTabPageControl1.Controls.Add(this.cmbCiudad);
			this.ultraTabPageControl1.Controls.Add(this.label1);
			this.ultraTabPageControl1.Controls.Add(this.cmbGrupo);
			this.ultraTabPageControl1.Controls.Add(this.label11);
			this.ultraTabPageControl1.Controls.Add(this.cmbTipoRuc);
			this.ultraTabPageControl1.Controls.Add(this.bCiudad);
			this.ultraTabPageControl1.Controls.Add(this.label13);
			this.ultraTabPageControl1.Controls.Add(this.label4);
			this.ultraTabPageControl1.Controls.Add(this.label22);
			this.ultraTabPageControl1.Controls.Add(this.label3);
			this.ultraTabPageControl1.Controls.Add(this.label17);
			this.ultraTabPageControl1.Controls.Add(this.txtNombre);
			this.ultraTabPageControl1.Controls.Add(this.txtDireccion);
			this.ultraTabPageControl1.Controls.Add(this.txtTelefono);
			this.ultraTabPageControl1.Controls.Add(this.ultraTextEditor3);
			this.ultraTabPageControl1.Controls.Add(this.txtRuc);
			this.ultraTabPageControl1.Controls.Add(this.btVerifica);
			this.ultraTabPageControl1.Controls.Add(this.txtCodigo);
			this.ultraTabPageControl1.Controls.Add(this.cmbFecha);
			this.ultraTabPageControl1.Controls.Add(this.label18);
			this.ultraTabPageControl1.Controls.Add(this.label6);
			this.ultraTabPageControl1.Controls.Add(this.chkProspecto);
			this.ultraTabPageControl1.Location = new System.Drawing.Point(1, 23);
			this.ultraTabPageControl1.Name = "ultraTabPageControl1";
			this.ultraTabPageControl1.Size = new System.Drawing.Size(668, 371);
			// 
			// txtPais
			// 
			appearance1.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtPais.Appearance = appearance1;
			this.txtPais.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsCliente, "Cliente.Pais"));
			this.txtPais.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtPais.Location = new System.Drawing.Point(16, 104);
			this.txtPais.Name = "txtPais";
			this.txtPais.Size = new System.Drawing.Size(104, 22);
			this.txtPais.TabIndex = 151;
			// 
			// cdsCliente
			// 
			this.cdsCliente.BindingContextCtrl = this;
			this.cdsCliente.DataLibrary = "LibFacturacion";
			this.cdsCliente.DataLibraryUrl = "";
			this.cdsCliente.DataSetDef = "dsCliente";
			this.cdsCliente.FillOnRequest = false;
			this.cdsCliente.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsCliente.Name = "cdsCliente";
			this.cdsCliente.SchemaClassName = "LibFacturacion.DataClass";
			this.cdsCliente.SchemaDef = null;
			this.cdsCliente.PositionChanged += new C1.Data.PositionChangeEventHandler(this.cdsCliente_PositionChanged);
			this.cdsCliente.BeforeFill += new C1.Data.FillEventHandler(this.cdsCliente_BeforeFill);
			// 
			// cmbPais
			// 
			appearance2.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbPais.Appearance = appearance2;
			this.cmbPais.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbPais.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbPais.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsCliente, "Cliente.idPais"));
			this.cmbPais.DataSource = this.ultraDataSource1;
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
			this.cmbPais.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbPais.Location = new System.Drawing.Point(80, 88);
			this.cmbPais.Name = "cmbPais";
			this.cmbPais.Size = new System.Drawing.Size(80, 21);
			this.cmbPais.TabIndex = 149;
			this.cmbPais.ValueMember = "idClientePais";
			// 
			// ultraDataSource1
			// 
			ultraDataColumn1.DataType = typeof(int);
			this.ultraDataSource1.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn1,
																																 ultraDataColumn2});
			// 
			// bPais
			// 
			this.bPais.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.bPais.Location = new System.Drawing.Point(160, 86);
			this.bPais.Name = "bPais";
			this.bPais.Size = new System.Drawing.Size(24, 26);
			this.bPais.TabIndex = 150;
			this.bPais.Text = "...";
			this.bPais.Click += new System.EventHandler(this.bPais_Click);
			// 
			// ultraTextEditor6
			// 
			appearance3.ForeColorDisabled = System.Drawing.Color.Black;
			this.ultraTextEditor6.Appearance = appearance3;
			this.ultraTextEditor6.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsCliente, "Cliente.Comercial"));
			this.ultraTextEditor6.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.ultraTextEditor6.Location = new System.Drawing.Point(80, 327);
			this.ultraTextEditor6.Name = "ultraTextEditor6";
			this.ultraTextEditor6.Size = new System.Drawing.Size(288, 22);
			this.ultraTextEditor6.TabIndex = 12;
			this.toolTip1.SetToolTip(this.ultraTextEditor6, "Nombre comercial");
			// 
			// ultraCalendarCombo1
			// 
			appearance4.ForeColorDisabled = System.Drawing.Color.Black;
			this.ultraCalendarCombo1.Appearance = appearance4;
			this.ultraCalendarCombo1.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.ultraCalendarCombo1.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsCliente, "Cliente.Vencimiento"));
			dateButton1.Caption = "Today";
			this.ultraCalendarCombo1.DateButtons.Add(dateButton1);
			this.ultraCalendarCombo1.Format = "dd/MMM/yyyy";
			this.ultraCalendarCombo1.Location = new System.Drawing.Point(240, 190);
			this.ultraCalendarCombo1.Name = "ultraCalendarCombo1";
			this.ultraCalendarCombo1.NonAutoSizeHeight = 23;
			this.ultraCalendarCombo1.Size = new System.Drawing.Size(128, 21);
			this.ultraCalendarCombo1.SpinButtonsVisible = true;
			this.ultraCalendarCombo1.TabIndex = 8;
			this.ultraCalendarCombo1.Value = new System.DateTime(2006, 1, 14, 0, 0, 0, 0);
			// 
			// label28
			// 
			this.label28.AutoSize = true;
			this.label28.BackColor = System.Drawing.Color.Transparent;
			this.label28.Location = new System.Drawing.Point(16, 327);
			this.label28.Name = "label28";
			this.label28.Size = new System.Drawing.Size(53, 16);
			this.label28.TabIndex = 148;
			this.label28.Text = "Comercial";
			// 
			// btPublicidad
			// 
			this.btPublicidad.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.btPublicidad.Location = new System.Drawing.Point(624, 293);
			this.btPublicidad.Name = "btPublicidad";
			this.btPublicidad.Size = new System.Drawing.Size(24, 26);
			this.btPublicidad.TabIndex = 28;
			this.btPublicidad.Text = "...";
			this.btPublicidad.Click += new System.EventHandler(this.btPublicidad_Click);
			// 
			// cmbClientePublicidad
			// 
			appearance5.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbClientePublicidad.Appearance = appearance5;
			this.cmbClientePublicidad.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbClientePublicidad.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbClientePublicidad.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsCliente, "Cliente.idClientePublicidad"));
			ultraGridColumn3.Header.VisiblePosition = 0;
			ultraGridColumn3.Hidden = true;
			ultraGridColumn4.Header.VisiblePosition = 1;
			ultraGridColumn4.Width = 310;
			ultraGridBand2.Columns.AddRange(new object[] {
																										 ultraGridColumn3,
																										 ultraGridColumn4});
			this.cmbClientePublicidad.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			this.cmbClientePublicidad.DisplayMember = "Medio";
			this.cmbClientePublicidad.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbClientePublicidad.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbClientePublicidad.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbClientePublicidad.Location = new System.Drawing.Point(480, 293);
			this.cmbClientePublicidad.Name = "cmbClientePublicidad";
			this.cmbClientePublicidad.Size = new System.Drawing.Size(136, 21);
			this.cmbClientePublicidad.TabIndex = 21;
			this.cmbClientePublicidad.ValueMember = "idClientePublicidad";
			// 
			// label30
			// 
			this.label30.AutoSize = true;
			this.label30.BackColor = System.Drawing.Color.Transparent;
			this.label30.Location = new System.Drawing.Point(384, 293);
			this.label30.Name = "label30";
			this.label30.Size = new System.Drawing.Size(54, 16);
			this.label30.TabIndex = 107;
			this.label30.Text = "Publicidad";
			// 
			// label27
			// 
			this.label27.AutoSize = true;
			this.label27.BackColor = System.Drawing.Color.Transparent;
			this.label27.Location = new System.Drawing.Point(384, 224);
			this.label27.Name = "label27";
			this.label27.Size = new System.Drawing.Size(33, 16);
			this.label27.TabIndex = 105;
			this.label27.Text = "Notas";
			// 
			// ultraTextEditor2
			// 
			appearance6.ForeColorDisabled = System.Drawing.Color.Black;
			this.ultraTextEditor2.Appearance = appearance6;
			this.ultraTextEditor2.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsCliente, "Cliente.Notas"));
			this.ultraTextEditor2.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.ultraTextEditor2.Location = new System.Drawing.Point(480, 224);
			this.ultraTextEditor2.Multiline = true;
			this.ultraTextEditor2.Name = "ultraTextEditor2";
			this.ultraTextEditor2.Size = new System.Drawing.Size(168, 60);
			this.ultraTextEditor2.TabIndex = 20;
			// 
			// label26
			// 
			this.label26.AutoSize = true;
			this.label26.BackColor = System.Drawing.Color.Transparent;
			this.label26.Location = new System.Drawing.Point(384, 190);
			this.label26.Name = "label26";
			this.label26.Size = new System.Drawing.Size(65, 16);
			this.label26.TabIndex = 103;
			this.label26.Text = "Observación";
			// 
			// ultraTextEditor1
			// 
			appearance7.ForeColorDisabled = System.Drawing.Color.Black;
			this.ultraTextEditor1.Appearance = appearance7;
			this.ultraTextEditor1.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsCliente, "Cliente.Observacion"));
			this.ultraTextEditor1.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.ultraTextEditor1.Location = new System.Drawing.Point(480, 190);
			this.ultraTextEditor1.Name = "ultraTextEditor1";
			this.ultraTextEditor1.Size = new System.Drawing.Size(168, 22);
			this.ultraTextEditor1.TabIndex = 19;
			// 
			// ultraCheckEditor1
			// 
			appearance8.ForeColorDisabled = System.Drawing.Color.Black;
			this.ultraCheckEditor1.Appearance = appearance8;
			this.ultraCheckEditor1.BackColor = System.Drawing.Color.Transparent;
			this.ultraCheckEditor1.DataBindings.Add(new System.Windows.Forms.Binding("CheckedValue", this.cdsCliente, "Cliente.Ambos"));
			this.ultraCheckEditor1.Location = new System.Drawing.Point(384, 52);
			this.ultraCheckEditor1.Name = "ultraCheckEditor1";
			this.ultraCheckEditor1.Size = new System.Drawing.Size(128, 21);
			this.ultraCheckEditor1.TabIndex = 14;
			this.ultraCheckEditor1.Text = "Proveedor y Cliente";
			// 
			// ultraTextEditor5
			// 
			appearance9.ForeColorDisabled = System.Drawing.Color.Black;
			this.ultraTextEditor5.Appearance = appearance9;
			this.ultraTextEditor5.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsCliente, "Cliente.Contacto"));
			this.ultraTextEditor5.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.ultraTextEditor5.Location = new System.Drawing.Point(480, 155);
			this.ultraTextEditor5.Name = "ultraTextEditor5";
			this.ultraTextEditor5.Size = new System.Drawing.Size(168, 22);
			this.ultraTextEditor5.TabIndex = 18;
			// 
			// txtCelular
			// 
			appearance10.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtCelular.Appearance = appearance10;
			this.txtCelular.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsCliente, "Cliente.Celular"));
			this.txtCelular.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtCelular.Location = new System.Drawing.Point(480, 121);
			this.txtCelular.Name = "txtCelular";
			this.txtCelular.Size = new System.Drawing.Size(168, 22);
			this.txtCelular.TabIndex = 17;
			// 
			// cmbPrecio
			// 
			appearance11.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbPrecio.Appearance = appearance11;
			this.cmbPrecio.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbPrecio.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbPrecio.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsCliente, "Cliente.idTipoPrecio"));
			ultraGridColumn5.Header.VisiblePosition = 0;
			ultraGridColumn5.Hidden = true;
			ultraGridColumn6.Header.VisiblePosition = 1;
			ultraGridColumn6.Width = 291;
			ultraGridBand3.Columns.AddRange(new object[] {
																										 ultraGridColumn5,
																										 ultraGridColumn6});
			this.cmbPrecio.DisplayLayout.BandsSerializer.Add(ultraGridBand3);
			this.cmbPrecio.DisplayMember = "Precio";
			this.cmbPrecio.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbPrecio.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbPrecio.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbPrecio.Location = new System.Drawing.Point(480, 17);
			this.cmbPrecio.Name = "cmbPrecio";
			this.cmbPrecio.Size = new System.Drawing.Size(136, 21);
			this.cmbPrecio.TabIndex = 13;
			this.cmbPrecio.ValueMember = "idTipoPrecio";
			// 
			// cmbSector
			// 
			appearance12.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbSector.Appearance = appearance12;
			this.cmbSector.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbSector.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbSector.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsCliente, "Cliente.idSector"));
			ultraGridColumn7.Header.VisiblePosition = 0;
			ultraGridColumn7.Hidden = true;
			ultraGridColumn8.Header.VisiblePosition = 1;
			ultraGridColumn8.Width = 274;
			ultraGridBand4.Columns.AddRange(new object[] {
																										 ultraGridColumn7,
																										 ultraGridColumn8});
			this.cmbSector.DisplayLayout.BandsSerializer.Add(ultraGridBand4);
			this.cmbSector.DisplayMember = "Sector";
			this.cmbSector.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbSector.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbSector.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbSector.Location = new System.Drawing.Point(480, 86);
			this.cmbSector.Name = "cmbSector";
			this.cmbSector.Size = new System.Drawing.Size(136, 21);
			this.cmbSector.TabIndex = 16;
			this.cmbSector.ValueMember = "idSector";
			// 
			// btSector
			// 
			this.btSector.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.btSector.Location = new System.Drawing.Point(624, 86);
			this.btSector.Name = "btSector";
			this.btSector.Size = new System.Drawing.Size(24, 26);
			this.btSector.TabIndex = 27;
			this.btSector.Text = "...";
			this.btSector.Click += new System.EventHandler(this.btSector_Click);
			// 
			// btPrecio
			// 
			this.btPrecio.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.btPrecio.Location = new System.Drawing.Point(624, 17);
			this.btPrecio.Name = "btPrecio";
			this.btPrecio.Size = new System.Drawing.Size(24, 26);
			this.btPrecio.TabIndex = 26;
			this.btPrecio.Text = "...";
			this.btPrecio.Click += new System.EventHandler(this.btPrecio_Click);
			// 
			// label19
			// 
			this.label19.AutoSize = true;
			this.label19.BackColor = System.Drawing.Color.Transparent;
			this.label19.Location = new System.Drawing.Point(384, 86);
			this.label19.Name = "label19";
			this.label19.Size = new System.Drawing.Size(36, 16);
			this.label19.TabIndex = 101;
			this.label19.Text = "Sector";
			// 
			// label10
			// 
			this.label10.AutoSize = true;
			this.label10.BackColor = System.Drawing.Color.Transparent;
			this.label10.Location = new System.Drawing.Point(384, 155);
			this.label10.Name = "label10";
			this.label10.Size = new System.Drawing.Size(52, 16);
			this.label10.TabIndex = 90;
			this.label10.Text = "Contacto:";
			// 
			// label9
			// 
			this.label9.AutoSize = true;
			this.label9.BackColor = System.Drawing.Color.Transparent;
			this.label9.Location = new System.Drawing.Point(384, 121);
			this.label9.Name = "label9";
			this.label9.Size = new System.Drawing.Size(38, 16);
			this.label9.TabIndex = 87;
			this.label9.Text = "Celular";
			// 
			// label2
			// 
			this.label2.AutoSize = true;
			this.label2.BackColor = System.Drawing.Color.Transparent;
			this.label2.Location = new System.Drawing.Point(384, 17);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(38, 16);
			this.label2.TabIndex = 83;
			this.label2.Text = "Precio:";
			// 
			// label5
			// 
			this.label5.AutoSize = true;
			this.label5.BackColor = System.Drawing.Color.Transparent;
			this.label5.Location = new System.Drawing.Point(192, 86);
			this.label5.Name = "label5";
			this.label5.Size = new System.Drawing.Size(42, 16);
			this.label5.TabIndex = 10;
			this.label5.Text = "Ciudad:";
			// 
			// txtEMail
			// 
			appearance13.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtEMail.Appearance = appearance13;
			this.txtEMail.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsCliente, "Cliente.Email"));
			this.txtEMail.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtEMail.Location = new System.Drawing.Point(80, 224);
			this.txtEMail.Name = "txtEMail";
			this.txtEMail.Size = new System.Drawing.Size(288, 22);
			this.txtEMail.TabIndex = 9;
			// 
			// ultraTextEditor10
			// 
			appearance14.ForeColorDisabled = System.Drawing.Color.Black;
			this.ultraTextEditor10.Appearance = appearance14;
			this.ultraTextEditor10.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsCliente, "Cliente.Representante"));
			this.ultraTextEditor10.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.ultraTextEditor10.Location = new System.Drawing.Point(80, 293);
			this.ultraTextEditor10.Name = "ultraTextEditor10";
			this.ultraTextEditor10.Size = new System.Drawing.Size(288, 22);
			this.ultraTextEditor10.TabIndex = 11;
			// 
			// label7
			// 
			this.label7.AutoSize = true;
			this.label7.BackColor = System.Drawing.Color.Transparent;
			this.label7.Location = new System.Drawing.Point(16, 121);
			this.label7.Name = "label7";
			this.label7.Size = new System.Drawing.Size(52, 16);
			this.label7.TabIndex = 14;
			this.label7.Text = "Teléfono:";
			// 
			// label8
			// 
			this.label8.AutoSize = true;
			this.label8.BackColor = System.Drawing.Color.Transparent;
			this.label8.Location = new System.Drawing.Point(192, 121);
			this.label8.Name = "label8";
			this.label8.Size = new System.Drawing.Size(26, 16);
			this.label8.TabIndex = 16;
			this.label8.Text = "Fax:";
			// 
			// btGrupoCliente
			// 
			this.btGrupoCliente.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.btGrupoCliente.Location = new System.Drawing.Point(344, 258);
			this.btGrupoCliente.Name = "btGrupoCliente";
			this.btGrupoCliente.Size = new System.Drawing.Size(24, 26);
			this.btGrupoCliente.TabIndex = 25;
			this.btGrupoCliente.Text = "...";
			this.btGrupoCliente.Click += new System.EventHandler(this.btGrupoCliente_Click);
			// 
			// cmbCiudad
			// 
			appearance15.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbCiudad.Appearance = appearance15;
			this.cmbCiudad.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbCiudad.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbCiudad.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsCliente, "Cliente.idCiudad"));
			ultraGridColumn9.Header.VisiblePosition = 0;
			ultraGridColumn9.Hidden = true;
			ultraGridColumn10.Header.VisiblePosition = 1;
			ultraGridColumn10.Width = 276;
			ultraGridBand5.Columns.AddRange(new object[] {
																										 ultraGridColumn9,
																										 ultraGridColumn10});
			this.cmbCiudad.DisplayLayout.BandsSerializer.Add(ultraGridBand5);
			this.cmbCiudad.DisplayMember = "Ciudad";
			this.cmbCiudad.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbCiudad.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbCiudad.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbCiudad.Location = new System.Drawing.Point(240, 86);
			this.cmbCiudad.Name = "cmbCiudad";
			this.cmbCiudad.Size = new System.Drawing.Size(104, 21);
			this.cmbCiudad.TabIndex = 2;
			this.cmbCiudad.ValueMember = "idCiudad";
			// 
			// label1
			// 
			this.label1.AutoSize = true;
			this.label1.BackColor = System.Drawing.Color.Transparent;
			this.label1.Location = new System.Drawing.Point(16, 17);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(48, 16);
			this.label1.TabIndex = 1;
			this.label1.Text = "Nombre:";
			// 
			// cmbGrupo
			// 
			appearance16.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbGrupo.Appearance = appearance16;
			this.cmbGrupo.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbGrupo.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbGrupo.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsCliente, "Cliente.idGrupoCliente"));
			ultraGridColumn11.Header.VisiblePosition = 0;
			ultraGridColumn11.Hidden = true;
			ultraGridColumn12.Header.VisiblePosition = 1;
			ultraGridColumn12.Width = 385;
			ultraGridBand6.Columns.AddRange(new object[] {
																										 ultraGridColumn11,
																										 ultraGridColumn12});
			this.cmbGrupo.DisplayLayout.BandsSerializer.Add(ultraGridBand6);
			this.cmbGrupo.DisplayMember = "Grupo";
			this.cmbGrupo.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbGrupo.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbGrupo.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbGrupo.Location = new System.Drawing.Point(80, 258);
			this.cmbGrupo.Name = "cmbGrupo";
			this.cmbGrupo.Size = new System.Drawing.Size(256, 21);
			this.cmbGrupo.TabIndex = 10;
			this.cmbGrupo.ValueMember = "idGrupoCliente";
			this.cmbGrupo.MouseDown += new System.Windows.Forms.MouseEventHandler(this.cmbGrupo_MouseDown);
			// 
			// label11
			// 
			this.label11.AutoSize = true;
			this.label11.BackColor = System.Drawing.Color.Transparent;
			this.label11.Location = new System.Drawing.Point(16, 224);
			this.label11.Name = "label11";
			this.label11.Size = new System.Drawing.Size(35, 16);
			this.label11.TabIndex = 24;
			this.label11.Text = "Email:";
			// 
			// cmbTipoRuc
			// 
			appearance17.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbTipoRuc.Appearance = appearance17;
			this.cmbTipoRuc.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbTipoRuc.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbTipoRuc.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsCliente, "Cliente.idTipoRuc"));
			ultraGridColumn13.Header.VisiblePosition = 0;
			ultraGridColumn13.Hidden = true;
			ultraGridColumn14.Header.VisiblePosition = 1;
			ultraGridColumn14.Width = 233;
			ultraGridBand7.Columns.AddRange(new object[] {
																										 ultraGridColumn13,
																										 ultraGridColumn14});
			this.cmbTipoRuc.DisplayLayout.BandsSerializer.Add(ultraGridBand7);
			this.cmbTipoRuc.DisplayMember = "TipoRuc";
			this.cmbTipoRuc.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbTipoRuc.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbTipoRuc.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbTipoRuc.Location = new System.Drawing.Point(80, 155);
			this.cmbTipoRuc.Name = "cmbTipoRuc";
			this.cmbTipoRuc.Size = new System.Drawing.Size(144, 21);
			this.cmbTipoRuc.TabIndex = 5;
			this.cmbTipoRuc.ValueMember = "idTipoRuc";
			// 
			// bCiudad
			// 
			this.bCiudad.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.bCiudad.Location = new System.Drawing.Point(344, 86);
			this.bCiudad.Name = "bCiudad";
			this.bCiudad.Size = new System.Drawing.Size(24, 26);
			this.bCiudad.TabIndex = 23;
			this.bCiudad.Text = "...";
			this.bCiudad.Click += new System.EventHandler(this.bCiudad_Click);
			// 
			// label13
			// 
			this.label13.AutoSize = true;
			this.label13.BackColor = System.Drawing.Color.Transparent;
			this.label13.Location = new System.Drawing.Point(16, 293);
			this.label13.Name = "label13";
			this.label13.Size = new System.Drawing.Size(58, 16);
			this.label13.TabIndex = 28;
			this.label13.Text = "Represent.";
			// 
			// label4
			// 
			this.label4.AutoSize = true;
			this.label4.BackColor = System.Drawing.Color.Transparent;
			this.label4.Location = new System.Drawing.Point(16, 86);
			this.label4.Name = "label4";
			this.label4.Size = new System.Drawing.Size(28, 16);
			this.label4.TabIndex = 8;
			this.label4.Text = "País:";
			// 
			// label22
			// 
			this.label22.AutoSize = true;
			this.label22.BackColor = System.Drawing.Color.Transparent;
			this.label22.Location = new System.Drawing.Point(16, 155);
			this.label22.Name = "label22";
			this.label22.Size = new System.Drawing.Size(49, 16);
			this.label22.TabIndex = 62;
			this.label22.Text = "RUC/Ced";
			// 
			// label3
			// 
			this.label3.AutoSize = true;
			this.label3.BackColor = System.Drawing.Color.Transparent;
			this.label3.Location = new System.Drawing.Point(16, 52);
			this.label3.Name = "label3";
			this.label3.Size = new System.Drawing.Size(54, 16);
			this.label3.TabIndex = 5;
			this.label3.Text = "Dirección:";
			// 
			// label17
			// 
			this.label17.AutoSize = true;
			this.label17.BackColor = System.Drawing.Color.Transparent;
			this.label17.Location = new System.Drawing.Point(16, 258);
			this.label17.Name = "label17";
			this.label17.Size = new System.Drawing.Size(39, 16);
			this.label17.TabIndex = 36;
			this.label17.Text = "Grupo:";
			// 
			// txtNombre
			// 
			appearance18.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtNombre.Appearance = appearance18;
			this.txtNombre.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsCliente, "Cliente.Nombre"));
			this.txtNombre.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNombre.Location = new System.Drawing.Point(80, 17);
			this.txtNombre.Name = "txtNombre";
			this.txtNombre.Size = new System.Drawing.Size(288, 22);
			this.txtNombre.TabIndex = 0;
			// 
			// txtDireccion
			// 
			appearance19.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtDireccion.Appearance = appearance19;
			this.txtDireccion.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsCliente, "Cliente.Direccion"));
			this.txtDireccion.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtDireccion.Location = new System.Drawing.Point(80, 52);
			this.txtDireccion.Name = "txtDireccion";
			this.txtDireccion.Size = new System.Drawing.Size(288, 22);
			this.txtDireccion.TabIndex = 0;
			// 
			// txtTelefono
			// 
			appearance20.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtTelefono.Appearance = appearance20;
			this.txtTelefono.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsCliente, "Cliente.Telefono"));
			this.txtTelefono.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtTelefono.Location = new System.Drawing.Point(80, 121);
			this.txtTelefono.Name = "txtTelefono";
			this.txtTelefono.Size = new System.Drawing.Size(104, 22);
			this.txtTelefono.TabIndex = 3;
			// 
			// ultraTextEditor3
			// 
			appearance21.ForeColorDisabled = System.Drawing.Color.Black;
			this.ultraTextEditor3.Appearance = appearance21;
			this.ultraTextEditor3.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsCliente, "Cliente.Fax"));
			this.ultraTextEditor3.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.ultraTextEditor3.Location = new System.Drawing.Point(240, 121);
			this.ultraTextEditor3.Name = "ultraTextEditor3";
			this.ultraTextEditor3.Size = new System.Drawing.Size(128, 22);
			this.ultraTextEditor3.TabIndex = 4;
			// 
			// txtRuc
			// 
			appearance22.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtRuc.Appearance = appearance22;
			this.txtRuc.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsCliente, "Cliente.Ruc"));
			this.txtRuc.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtRuc.Location = new System.Drawing.Point(240, 155);
			this.txtRuc.Name = "txtRuc";
			this.txtRuc.Size = new System.Drawing.Size(104, 22);
			this.txtRuc.TabIndex = 6;
			// 
			// btVerifica
			// 
			appearance23.Image = ((object)(resources.GetObject("appearance23.Image")));
			this.btVerifica.Appearance = appearance23;
			this.btVerifica.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.btVerifica.Location = new System.Drawing.Point(344, 155);
			this.btVerifica.Name = "btVerifica";
			this.btVerifica.Size = new System.Drawing.Size(24, 26);
			this.btVerifica.TabIndex = 24;
			this.btVerifica.Click += new System.EventHandler(this.btVerifica_Click);
			// 
			// txtCodigo
			// 
			appearance24.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtCodigo.Appearance = appearance24;
			this.txtCodigo.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsCliente, "Cliente.Codigo"));
			this.txtCodigo.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtCodigo.Location = new System.Drawing.Point(480, 327);
			this.txtCodigo.Name = "txtCodigo";
			this.txtCodigo.Size = new System.Drawing.Size(136, 22);
			this.txtCodigo.TabIndex = 22;
			// 
			// cmbFecha
			// 
			appearance25.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbFecha.Appearance = appearance25;
			this.cmbFecha.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.cmbFecha.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsCliente, "Cliente.Fecha"));
			dateButton2.Caption = "Today";
			this.cmbFecha.DateButtons.Add(dateButton2);
			this.cmbFecha.Format = "dd/MMM/yyyy";
			this.cmbFecha.Location = new System.Drawing.Point(80, 190);
			this.cmbFecha.Name = "cmbFecha";
			this.cmbFecha.NonAutoSizeHeight = 23;
			this.cmbFecha.Size = new System.Drawing.Size(144, 21);
			this.cmbFecha.SpinButtonsVisible = true;
			this.cmbFecha.TabIndex = 7;
			this.cmbFecha.Value = new System.DateTime(2006, 1, 14, 0, 0, 0, 0);
			// 
			// label18
			// 
			this.label18.AutoSize = true;
			this.label18.BackColor = System.Drawing.Color.Transparent;
			this.label18.Location = new System.Drawing.Point(384, 327);
			this.label18.Name = "label18";
			this.label18.Size = new System.Drawing.Size(42, 16);
			this.label18.TabIndex = 39;
			this.label18.Text = "Código:";
			// 
			// label6
			// 
			this.label6.AutoSize = true;
			this.label6.BackColor = System.Drawing.Color.Transparent;
			this.label6.Location = new System.Drawing.Point(16, 190);
			this.label6.Name = "label6";
			this.label6.Size = new System.Drawing.Size(46, 16);
			this.label6.TabIndex = 11;
			this.label6.Text = "Periodo:";
			// 
			// chkProspecto
			// 
			appearance26.ForeColorDisabled = System.Drawing.Color.Black;
			this.chkProspecto.Appearance = appearance26;
			this.chkProspecto.BackColor = System.Drawing.Color.Transparent;
			this.chkProspecto.DataBindings.Add(new System.Windows.Forms.Binding("CheckedValue", this.cdsCliente, "Cliente.Prospecto"));
			this.chkProspecto.Location = new System.Drawing.Point(520, 52);
			this.chkProspecto.Name = "chkProspecto";
			this.chkProspecto.Size = new System.Drawing.Size(96, 21);
			this.chkProspecto.TabIndex = 15;
			this.chkProspecto.Text = "Prospecto";
			this.chkProspecto.CheckedChanged += new System.EventHandler(this.chkProspecto_CheckedChanged);
			// 
			// tpComplemento
			// 
			this.tpComplemento.Controls.Add(this.radioGroup1);
			this.tpComplemento.Controls.Add(this.chkEspecial);
			this.tpComplemento.Controls.Add(this.ultraTextEditor21);
			this.tpComplemento.Controls.Add(this.label51);
			this.tpComplemento.Controls.Add(this.ultraTextEditor4);
			this.tpComplemento.Controls.Add(this.label50);
			this.tpComplemento.Controls.Add(this.chkContribuyenteEsp);
			this.tpComplemento.Controls.Add(this.label34);
			this.tpComplemento.Controls.Add(this.cmbCredTribut);
			this.tpComplemento.Controls.Add(this.label33);
			this.tpComplemento.Controls.Add(this.cmbFormaPago);
			this.tpComplemento.Controls.Add(this.txtSerie);
			this.tpComplemento.Controls.Add(this.label32);
			this.tpComplemento.Controls.Add(this.txtAutRetencion);
			this.tpComplemento.Controls.Add(this.label31);
			this.tpComplemento.Controls.Add(this.cmbFechaCaduca);
			this.tpComplemento.Controls.Add(this.lblAutRetencion);
			this.tpComplemento.Controls.Add(this.label29);
			this.tpComplemento.Controls.Add(this.cmbPorcIvaRet);
			this.tpComplemento.Controls.Add(this.btDepende);
			this.tpComplemento.Controls.Add(this.label25);
			this.tpComplemento.Controls.Add(this.cmbDepende);
			this.tpComplemento.Controls.Add(this.label20);
			this.tpComplemento.Controls.Add(this.label21);
			this.tpComplemento.Controls.Add(this.cmbCobrador);
			this.tpComplemento.Controls.Add(this.cmbVendedor);
			this.tpComplemento.Controls.Add(this.btVendedor);
			this.tpComplemento.Controls.Add(this.txtAutImprenta);
			this.tpComplemento.Controls.Add(this.label16);
			this.tpComplemento.Controls.Add(this.txtAutFactura);
			this.tpComplemento.Controls.Add(this.label15);
			this.tpComplemento.Controls.Add(this.spnDescCredito);
			this.tpComplemento.Controls.Add(this.label24);
			this.tpComplemento.Controls.Add(this.spnDescuento);
			this.tpComplemento.Controls.Add(this.label23);
			this.tpComplemento.Controls.Add(this.txtValorCredito);
			this.tpComplemento.Controls.Add(this.ultraNumericEditor1);
			this.tpComplemento.Controls.Add(this.label14);
			this.tpComplemento.Controls.Add(this.label12);
			this.tpComplemento.Controls.Add(this.chkPublica);
			this.tpComplemento.Location = new System.Drawing.Point(-10000, -10000);
			this.tpComplemento.Name = "tpComplemento";
			this.tpComplemento.Size = new System.Drawing.Size(668, 371);
			// 
			// radioGroup1
			// 
			this.radioGroup1.Location = new System.Drawing.Point(456, 80);
			this.radioGroup1.Name = "radioGroup1";
			// 
			// radioGroup1.Properties
			// 
			this.radioGroup1.Properties.Items.AddRange(new DevExpress.XtraEditors.Controls.RadioGroupItem[] {
																																																				new DevExpress.XtraEditors.Controls.RadioGroupItem(1, "Contado"),
																																																				new DevExpress.XtraEditors.Controls.RadioGroupItem(2, "Crédito")});
			this.radioGroup1.Size = new System.Drawing.Size(160, 32);
			this.radioGroup1.TabIndex = 145;
			// 
			// chkEspecial
			// 
			appearance27.ForeColorDisabled = System.Drawing.Color.Black;
			this.chkEspecial.Appearance = appearance27;
			this.chkEspecial.BackColor = System.Drawing.Color.Transparent;
			this.chkEspecial.DataBindings.Add(new System.Windows.Forms.Binding("CheckedValue", this.cdsCliente, "Cliente.Especial"));
			this.chkEspecial.Location = new System.Drawing.Point(456, 120);
			this.chkEspecial.Name = "chkEspecial";
			this.chkEspecial.Size = new System.Drawing.Size(128, 21);
			this.chkEspecial.TabIndex = 144;
			this.chkEspecial.Text = "Especial";
			// 
			// ultraTextEditor21
			// 
			appearance28.ForeColorDisabled = System.Drawing.Color.Black;
			this.ultraTextEditor21.Appearance = appearance28;
			this.ultraTextEditor21.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsCliente, "Cliente.RangoHasta"));
			this.ultraTextEditor21.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.ultraTextEditor21.Location = new System.Drawing.Point(128, 155);
			this.ultraTextEditor21.Name = "ultraTextEditor21";
			this.ultraTextEditor21.Size = new System.Drawing.Size(136, 22);
			this.ultraTextEditor21.TabIndex = 142;
			// 
			// label51
			// 
			this.label51.AutoSize = true;
			this.label51.BackColor = System.Drawing.Color.Transparent;
			this.label51.Location = new System.Drawing.Point(32, 155);
			this.label51.Name = "label51";
			this.label51.Size = new System.Drawing.Size(67, 16);
			this.label51.TabIndex = 143;
			this.label51.Text = "Rango Final:";
			// 
			// ultraTextEditor4
			// 
			appearance29.ForeColorDisabled = System.Drawing.Color.Black;
			this.ultraTextEditor4.Appearance = appearance29;
			this.ultraTextEditor4.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsCliente, "Cliente.RangoDesde"));
			this.ultraTextEditor4.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.ultraTextEditor4.Location = new System.Drawing.Point(128, 121);
			this.ultraTextEditor4.Name = "ultraTextEditor4";
			this.ultraTextEditor4.Size = new System.Drawing.Size(136, 22);
			this.ultraTextEditor4.TabIndex = 140;
			// 
			// label50
			// 
			this.label50.AutoSize = true;
			this.label50.BackColor = System.Drawing.Color.Transparent;
			this.label50.Location = new System.Drawing.Point(32, 121);
			this.label50.Name = "label50";
			this.label50.Size = new System.Drawing.Size(73, 16);
			this.label50.TabIndex = 141;
			this.label50.Text = "Rango Inicial:";
			// 
			// chkContribuyenteEsp
			// 
			appearance30.ForeColorDisabled = System.Drawing.Color.Black;
			this.chkContribuyenteEsp.Appearance = appearance30;
			this.chkContribuyenteEsp.BackColor = System.Drawing.Color.Transparent;
			this.chkContribuyenteEsp.DataBindings.Add(new System.Windows.Forms.Binding("CheckedValue", this.cdsCliente, "Cliente.ContribEspecial"));
			this.chkContribuyenteEsp.Location = new System.Drawing.Point(320, 86);
			this.chkContribuyenteEsp.Name = "chkContribuyenteEsp";
			this.chkContribuyenteEsp.Size = new System.Drawing.Size(136, 22);
			this.chkContribuyenteEsp.TabIndex = 11;
			this.chkContribuyenteEsp.Text = "Contribuyente Especial";
			// 
			// label34
			// 
			this.label34.AutoSize = true;
			this.label34.BackColor = System.Drawing.Color.Transparent;
			this.label34.Location = new System.Drawing.Point(320, 52);
			this.label34.Name = "label34";
			this.label34.Size = new System.Drawing.Size(73, 16);
			this.label34.TabIndex = 139;
			this.label34.Text = "Crédito Tribut";
			// 
			// cmbCredTribut
			// 
			appearance31.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbCredTribut.Appearance = appearance31;
			this.cmbCredTribut.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbCredTribut.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbCredTribut.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsCliente, "Cliente.idCredTribut"));
			ultraGridColumn15.Header.VisiblePosition = 0;
			ultraGridColumn15.Hidden = true;
			ultraGridColumn16.Header.VisiblePosition = 1;
			ultraGridColumn16.Width = 295;
			ultraGridBand8.Columns.AddRange(new object[] {
																										 ultraGridColumn15,
																										 ultraGridColumn16});
			this.cmbCredTribut.DisplayLayout.BandsSerializer.Add(ultraGridBand8);
			this.cmbCredTribut.DisplayMember = "CredTributario";
			this.cmbCredTribut.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbCredTribut.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbCredTribut.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbCredTribut.Location = new System.Drawing.Point(416, 52);
			this.cmbCredTribut.Name = "cmbCredTribut";
			this.cmbCredTribut.Size = new System.Drawing.Size(160, 21);
			this.cmbCredTribut.TabIndex = 10;
			this.cmbCredTribut.ValueMember = "idCredTributario";
			// 
			// label33
			// 
			this.label33.AutoSize = true;
			this.label33.BackColor = System.Drawing.Color.Transparent;
			this.label33.Location = new System.Drawing.Point(320, 224);
			this.label33.Name = "label33";
			this.label33.Size = new System.Drawing.Size(64, 16);
			this.label33.TabIndex = 137;
			this.label33.Text = "Forma Pago";
			// 
			// cmbFormaPago
			// 
			appearance32.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbFormaPago.Appearance = appearance32;
			this.cmbFormaPago.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbFormaPago.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbFormaPago.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsCliente, "Cliente.idFormaPago"));
			ultraGridColumn17.Header.VisiblePosition = 0;
			ultraGridColumn17.Hidden = true;
			ultraGridColumn18.Header.VisiblePosition = 1;
			ultraGridColumn18.Width = 203;
			ultraGridBand9.Columns.AddRange(new object[] {
																										 ultraGridColumn17,
																										 ultraGridColumn18});
			this.cmbFormaPago.DisplayLayout.BandsSerializer.Add(ultraGridBand9);
			this.cmbFormaPago.DisplayMember = "FormaPago";
			this.cmbFormaPago.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbFormaPago.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbFormaPago.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbFormaPago.Location = new System.Drawing.Point(424, 224);
			this.cmbFormaPago.Name = "cmbFormaPago";
			this.cmbFormaPago.Size = new System.Drawing.Size(160, 21);
			this.cmbFormaPago.TabIndex = 13;
			this.cmbFormaPago.ValueMember = "idFormaPago";
			// 
			// txtSerie
			// 
			appearance33.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtSerie.Appearance = appearance33;
			this.txtSerie.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsCliente, "Cliente.SerieFactura"));
			this.txtSerie.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtSerie.Location = new System.Drawing.Point(128, 52);
			this.txtSerie.Name = "txtSerie";
			this.txtSerie.Size = new System.Drawing.Size(136, 22);
			this.txtSerie.TabIndex = 1;
			// 
			// label32
			// 
			this.label32.AutoSize = true;
			this.label32.BackColor = System.Drawing.Color.Transparent;
			this.label32.Location = new System.Drawing.Point(32, 52);
			this.label32.Name = "label32";
			this.label32.Size = new System.Drawing.Size(86, 16);
			this.label32.TabIndex = 135;
			this.label32.Text = "Serie de Factura";
			// 
			// txtAutRetencion
			// 
			appearance34.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtAutRetencion.Appearance = appearance34;
			this.txtAutRetencion.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsCliente, "Cliente.AutorizaRet"));
			this.txtAutRetencion.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtAutRetencion.Location = new System.Drawing.Point(448, 155);
			this.txtAutRetencion.Name = "txtAutRetencion";
			this.txtAutRetencion.Size = new System.Drawing.Size(136, 22);
			this.txtAutRetencion.TabIndex = 9;
			// 
			// label31
			// 
			this.label31.AutoSize = true;
			this.label31.BackColor = System.Drawing.Color.Transparent;
			this.label31.Location = new System.Drawing.Point(32, 86);
			this.label31.Name = "label31";
			this.label31.Size = new System.Drawing.Size(89, 16);
			this.label31.TabIndex = 132;
			this.label31.Text = "Fecha Caducidad";
			// 
			// cmbFechaCaduca
			// 
			appearance35.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbFechaCaduca.Appearance = appearance35;
			this.cmbFechaCaduca.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsCliente, "Cliente.FechaCaducidad"));
			this.cmbFechaCaduca.DateTime = new System.DateTime(2010, 8, 27, 0, 0, 0, 0);
			this.cmbFechaCaduca.FormatString = "MMM/yyyy";
			this.cmbFechaCaduca.Location = new System.Drawing.Point(128, 86);
			this.cmbFechaCaduca.Name = "cmbFechaCaduca";
			this.cmbFechaCaduca.Size = new System.Drawing.Size(136, 22);
			this.cmbFechaCaduca.SpinButtonDisplayStyle = Infragistics.Win.ButtonDisplayStyle.Always;
			this.cmbFechaCaduca.TabIndex = 2;
			this.cmbFechaCaduca.Value = new System.DateTime(2010, 8, 27, 0, 0, 0, 0);
			// 
			// lblAutRetencion
			// 
			this.lblAutRetencion.BackColor = System.Drawing.Color.Transparent;
			this.lblAutRetencion.Location = new System.Drawing.Point(320, 155);
			this.lblAutRetencion.Name = "lblAutRetencion";
			this.lblAutRetencion.Size = new System.Drawing.Size(121, 17);
			this.lblAutRetencion.TabIndex = 130;
			this.lblAutRetencion.Text = "Autorización Retención";
			// 
			// label29
			// 
			this.label29.AutoSize = true;
			this.label29.BackColor = System.Drawing.Color.Transparent;
			this.label29.Location = new System.Drawing.Point(320, 17);
			this.label29.Name = "label29";
			this.label29.Size = new System.Drawing.Size(118, 16);
			this.label29.TabIndex = 129;
			this.label29.Text = "% Ret. Iva Predefinido";
			// 
			// cmbPorcIvaRet
			// 
			appearance36.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbPorcIvaRet.Appearance = appearance36;
			this.cmbPorcIvaRet.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbPorcIvaRet.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbPorcIvaRet.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsCliente, "Cliente.idRetIva"));
			ultraGridColumn19.Header.VisiblePosition = 0;
			ultraGridColumn19.Hidden = true;
			ultraGridColumn20.Header.VisiblePosition = 1;
			ultraGridColumn21.Header.VisiblePosition = 2;
			ultraGridBand10.Columns.AddRange(new object[] {
																											ultraGridColumn19,
																											ultraGridColumn20,
																											ultraGridColumn21});
			this.cmbPorcIvaRet.DisplayLayout.BandsSerializer.Add(ultraGridBand10);
			this.cmbPorcIvaRet.DisplayMember = "Porcentaje";
			this.cmbPorcIvaRet.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbPorcIvaRet.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbPorcIvaRet.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbPorcIvaRet.Location = new System.Drawing.Point(520, 17);
			this.cmbPorcIvaRet.Name = "cmbPorcIvaRet";
			this.cmbPorcIvaRet.Size = new System.Drawing.Size(56, 21);
			this.cmbPorcIvaRet.TabIndex = 8;
			this.cmbPorcIvaRet.ValueMember = "idRetIva";
			// 
			// btDepende
			// 
			this.btDepende.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.btDepende.Location = new System.Drawing.Point(592, 327);
			this.btDepende.Name = "btDepende";
			this.btDepende.Size = new System.Drawing.Size(24, 26);
			this.btDepende.TabIndex = 18;
			this.btDepende.Text = "...";
			this.btDepende.Click += new System.EventHandler(this.btDepende_Click);
			this.btDepende.MouseUp += new System.Windows.Forms.MouseEventHandler(this.btDepende_MouseDown);
			// 
			// label25
			// 
			this.label25.AutoSize = true;
			this.label25.BackColor = System.Drawing.Color.Transparent;
			this.label25.Location = new System.Drawing.Point(32, 327);
			this.label25.Name = "label25";
			this.label25.Size = new System.Drawing.Size(68, 16);
			this.label25.TabIndex = 124;
			this.label25.Text = "Depende de:";
			// 
			// cmbDepende
			// 
			appearance37.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbDepende.Appearance = appearance37;
			this.cmbDepende.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbDepende.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbDepende.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsCliente, "Cliente.idPadre"));
			this.cmbDepende.DisplayLayout.AutoFitColumns = true;
			ultraGridColumn22.Header.VisiblePosition = 0;
			ultraGridColumn22.Hidden = true;
			ultraGridColumn23.Header.VisiblePosition = 1;
			ultraGridColumn23.Width = 437;
			ultraGridBand11.Columns.AddRange(new object[] {
																											ultraGridColumn22,
																											ultraGridColumn23});
			this.cmbDepende.DisplayLayout.BandsSerializer.Add(ultraGridBand11);
			this.cmbDepende.DisplayMember = "Nombre";
			this.cmbDepende.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbDepende.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbDepende.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbDepende.Location = new System.Drawing.Point(128, 327);
			this.cmbDepende.Name = "cmbDepende";
			this.cmbDepende.Size = new System.Drawing.Size(456, 21);
			this.cmbDepende.TabIndex = 17;
			this.cmbDepende.ValueMember = "idCliente";
			// 
			// label20
			// 
			this.label20.AutoSize = true;
			this.label20.BackColor = System.Drawing.Color.Transparent;
			this.label20.Location = new System.Drawing.Point(320, 258);
			this.label20.Name = "label20";
			this.label20.Size = new System.Drawing.Size(54, 16);
			this.label20.TabIndex = 121;
			this.label20.Text = "Cobrador:";
			// 
			// label21
			// 
			this.label21.AutoSize = true;
			this.label21.BackColor = System.Drawing.Color.Transparent;
			this.label21.Location = new System.Drawing.Point(320, 293);
			this.label21.Name = "label21";
			this.label21.Size = new System.Drawing.Size(52, 16);
			this.label21.TabIndex = 122;
			this.label21.Text = "Vendedor";
			// 
			// cmbCobrador
			// 
			appearance38.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbCobrador.Appearance = appearance38;
			this.cmbCobrador.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbCobrador.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbCobrador.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsCliente, "Cliente.idCobrador"));
			ultraGridColumn24.Header.VisiblePosition = 0;
			ultraGridColumn24.Hidden = true;
			ultraGridColumn25.Header.VisiblePosition = 1;
			ultraGridColumn25.Width = 304;
			ultraGridBand12.Columns.AddRange(new object[] {
																											ultraGridColumn24,
																											ultraGridColumn25});
			this.cmbCobrador.DisplayLayout.BandsSerializer.Add(ultraGridBand12);
			this.cmbCobrador.DisplayMember = "Nombre";
			this.cmbCobrador.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbCobrador.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbCobrador.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbCobrador.Location = new System.Drawing.Point(424, 258);
			this.cmbCobrador.Name = "cmbCobrador";
			this.cmbCobrador.Size = new System.Drawing.Size(160, 21);
			this.cmbCobrador.TabIndex = 14;
			this.cmbCobrador.ValueMember = "idPersonal";
			// 
			// cmbVendedor
			// 
			appearance39.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbVendedor.Appearance = appearance39;
			this.cmbVendedor.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbVendedor.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbVendedor.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsCliente, "Cliente.idVendedor"));
			ultraGridColumn26.Header.VisiblePosition = 0;
			ultraGridColumn26.Hidden = true;
			ultraGridColumn27.Header.VisiblePosition = 1;
			ultraGridColumn27.Width = 350;
			ultraGridBand13.Columns.AddRange(new object[] {
																											ultraGridColumn26,
																											ultraGridColumn27});
			this.cmbVendedor.DisplayLayout.BandsSerializer.Add(ultraGridBand13);
			this.cmbVendedor.DisplayMember = "Nombre";
			this.cmbVendedor.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbVendedor.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbVendedor.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbVendedor.Location = new System.Drawing.Point(424, 293);
			this.cmbVendedor.Name = "cmbVendedor";
			this.cmbVendedor.Size = new System.Drawing.Size(160, 21);
			this.cmbVendedor.TabIndex = 15;
			this.cmbVendedor.ValueMember = "idPersonal";
			// 
			// btVendedor
			// 
			this.btVendedor.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.btVendedor.Location = new System.Drawing.Point(592, 293);
			this.btVendedor.Name = "btVendedor";
			this.btVendedor.Size = new System.Drawing.Size(24, 26);
			this.btVendedor.TabIndex = 16;
			this.btVendedor.Text = "...";
			this.btVendedor.Click += new System.EventHandler(this.btVendedor_Click);
			// 
			// txtAutImprenta
			// 
			appearance40.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtAutImprenta.Appearance = appearance40;
			this.txtAutImprenta.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsCliente, "Cliente.AutorizaImprenta"));
			this.txtAutImprenta.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtAutImprenta.Location = new System.Drawing.Point(448, 190);
			this.txtAutImprenta.Name = "txtAutImprenta";
			this.txtAutImprenta.Size = new System.Drawing.Size(136, 22);
			this.txtAutImprenta.TabIndex = 3;
			// 
			// label16
			// 
			this.label16.AutoSize = true;
			this.label16.BackColor = System.Drawing.Color.Transparent;
			this.label16.Location = new System.Drawing.Point(320, 190);
			this.label16.Name = "label16";
			this.label16.Size = new System.Drawing.Size(96, 16);
			this.label16.TabIndex = 117;
			this.label16.Text = "Autoriz. Imprenta:";
			// 
			// txtAutFactura
			// 
			appearance41.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtAutFactura.Appearance = appearance41;
			this.txtAutFactura.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsCliente, "Cliente.AutorizaFactura"));
			this.txtAutFactura.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtAutFactura.Location = new System.Drawing.Point(128, 17);
			this.txtAutFactura.Name = "txtAutFactura";
			this.txtAutFactura.Size = new System.Drawing.Size(136, 22);
			this.txtAutFactura.TabIndex = 0;
			// 
			// label15
			// 
			this.label15.AutoSize = true;
			this.label15.BackColor = System.Drawing.Color.Transparent;
			this.label15.Location = new System.Drawing.Point(32, 17);
			this.label15.Name = "label15";
			this.label15.Size = new System.Drawing.Size(87, 16);
			this.label15.TabIndex = 115;
			this.label15.Text = "Autoriz. Factura:";
			// 
			// spnDescCredito
			// 
			appearance42.ForeColorDisabled = System.Drawing.Color.Black;
			this.spnDescCredito.Appearance = appearance42;
			this.spnDescCredito.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsCliente, "Cliente.DescuentoCredito"));
			this.spnDescCredito.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.spnDescCredito.Location = new System.Drawing.Point(128, 293);
			this.spnDescCredito.MinValue = 0;
			this.spnDescCredito.Name = "spnDescCredito";
			this.spnDescCredito.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.spnDescCredito.PromptChar = ' ';
			this.spnDescCredito.Size = new System.Drawing.Size(136, 22);
			this.spnDescCredito.SpinButtonDisplayStyle = Infragistics.Win.ButtonDisplayStyle.Always;
			this.spnDescCredito.TabIndex = 7;
			// 
			// label24
			// 
			this.label24.AutoSize = true;
			this.label24.BackColor = System.Drawing.Color.Transparent;
			this.label24.Location = new System.Drawing.Point(32, 293);
			this.label24.Name = "label24";
			this.label24.Size = new System.Drawing.Size(71, 16);
			this.label24.TabIndex = 113;
			this.label24.Text = "Desc. Crédito";
			// 
			// spnDescuento
			// 
			appearance43.ForeColorDisabled = System.Drawing.Color.Black;
			this.spnDescuento.Appearance = appearance43;
			this.spnDescuento.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsCliente, "Cliente.DescuentoContado"));
			this.spnDescuento.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.spnDescuento.Location = new System.Drawing.Point(128, 258);
			this.spnDescuento.MinValue = 0;
			this.spnDescuento.Name = "spnDescuento";
			this.spnDescuento.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.spnDescuento.PromptChar = ' ';
			this.spnDescuento.Size = new System.Drawing.Size(136, 22);
			this.spnDescuento.SpinButtonDisplayStyle = Infragistics.Win.ButtonDisplayStyle.Always;
			this.spnDescuento.TabIndex = 6;
			// 
			// label23
			// 
			this.label23.AutoSize = true;
			this.label23.BackColor = System.Drawing.Color.Transparent;
			this.label23.Location = new System.Drawing.Point(32, 258);
			this.label23.Name = "label23";
			this.label23.Size = new System.Drawing.Size(76, 16);
			this.label23.TabIndex = 111;
			this.label23.Text = "Desc. Contado";
			// 
			// txtValorCredito
			// 
			appearance44.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtValorCredito.Appearance = appearance44;
			this.txtValorCredito.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsCliente, "Cliente.MaximoCredito"));
			this.txtValorCredito.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtValorCredito.Location = new System.Drawing.Point(128, 224);
			this.txtValorCredito.MinValue = 0;
			this.txtValorCredito.Name = "txtValorCredito";
			this.txtValorCredito.PromptChar = ' ';
			this.txtValorCredito.Size = new System.Drawing.Size(136, 22);
			this.txtValorCredito.SpinButtonDisplayStyle = Infragistics.Win.ButtonDisplayStyle.Always;
			this.txtValorCredito.TabIndex = 5;
			// 
			// ultraNumericEditor1
			// 
			appearance45.ForeColorDisabled = System.Drawing.Color.Black;
			this.ultraNumericEditor1.Appearance = appearance45;
			this.ultraNumericEditor1.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsCliente, "Cliente.DiasCredito"));
			this.ultraNumericEditor1.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.ultraNumericEditor1.Location = new System.Drawing.Point(128, 190);
			this.ultraNumericEditor1.MinValue = 0;
			this.ultraNumericEditor1.Name = "ultraNumericEditor1";
			this.ultraNumericEditor1.PromptChar = ' ';
			this.ultraNumericEditor1.Size = new System.Drawing.Size(136, 22);
			this.ultraNumericEditor1.SpinButtonDisplayStyle = Infragistics.Win.ButtonDisplayStyle.Always;
			this.ultraNumericEditor1.TabIndex = 4;
			// 
			// label14
			// 
			this.label14.AutoSize = true;
			this.label14.BackColor = System.Drawing.Color.Transparent;
			this.label14.Location = new System.Drawing.Point(32, 224);
			this.label14.Name = "label14";
			this.label14.Size = new System.Drawing.Size(73, 16);
			this.label14.TabIndex = 109;
			this.label14.Text = "Valor Crédito:";
			// 
			// label12
			// 
			this.label12.AutoSize = true;
			this.label12.BackColor = System.Drawing.Color.Transparent;
			this.label12.Location = new System.Drawing.Point(32, 190);
			this.label12.Name = "label12";
			this.label12.Size = new System.Drawing.Size(69, 16);
			this.label12.TabIndex = 108;
			this.label12.Text = "Días Crédito:";
			// 
			// chkPublica
			// 
			appearance46.ForeColorDisabled = System.Drawing.Color.Black;
			this.chkPublica.Appearance = appearance46;
			this.chkPublica.BackColor = System.Drawing.Color.Transparent;
			this.chkPublica.DataBindings.Add(new System.Windows.Forms.Binding("CheckedValue", this.cdsCliente, "Cliente.InstPublica"));
			this.chkPublica.Location = new System.Drawing.Point(320, 121);
			this.chkPublica.Name = "chkPublica";
			this.chkPublica.Size = new System.Drawing.Size(128, 21);
			this.chkPublica.TabIndex = 109;
			this.chkPublica.Text = "Institución Pública";
			// 
			// ultraTabPageControl5
			// 
			this.ultraTabPageControl5.Controls.Add(this.ultraGrid2);
			this.ultraTabPageControl5.Location = new System.Drawing.Point(-10000, -10000);
			this.ultraTabPageControl5.Name = "ultraTabPageControl5";
			this.ultraTabPageControl5.Size = new System.Drawing.Size(668, 371);
			// 
			// ultraGrid2
			// 
			this.ultraGrid2.Cursor = System.Windows.Forms.Cursors.Default;
			this.ultraGrid2.DataMember = "ClienteContacto";
			this.ultraGrid2.DataSource = this.cdsContacto;
			appearance47.BackColor = System.Drawing.Color.White;
			this.ultraGrid2.DisplayLayout.Appearance = appearance47;
			this.ultraGrid2.DisplayLayout.AutoFitColumns = true;
			ultraGridBand14.AddButtonCaption = "ClienteTipoRuc";
			ultraGridColumn28.Header.VisiblePosition = 0;
			ultraGridColumn28.Hidden = true;
			ultraGridColumn29.Header.VisiblePosition = 1;
			ultraGridColumn29.Hidden = true;
			ultraGridColumn30.Header.VisiblePosition = 2;
			ultraGridColumn30.Width = 177;
			ultraGridColumn31.Header.VisiblePosition = 3;
			ultraGridColumn31.Width = 173;
			ultraGridColumn32.Header.VisiblePosition = 4;
			ultraGridColumn32.Width = 173;
			ultraGridColumn33.Header.VisiblePosition = 5;
			ultraGridColumn33.Width = 88;
			ultraGridBand14.Columns.AddRange(new object[] {
																											ultraGridColumn28,
																											ultraGridColumn29,
																											ultraGridColumn30,
																											ultraGridColumn31,
																											ultraGridColumn32,
																											ultraGridColumn33});
			this.ultraGrid2.DisplayLayout.BandsSerializer.Add(ultraGridBand14);
			this.ultraGrid2.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.TemplateOnBottom;
			this.ultraGrid2.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.True;
			this.ultraGrid2.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance48.BackColor = System.Drawing.Color.Transparent;
			this.ultraGrid2.DisplayLayout.Override.CardAreaAppearance = appearance48;
			appearance49.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance49.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance49.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance49.FontData.BoldAsString = "True";
			appearance49.FontData.Name = "Arial";
			appearance49.FontData.SizeInPoints = 10F;
			appearance49.ForeColor = System.Drawing.Color.White;
			appearance49.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.ultraGrid2.DisplayLayout.Override.HeaderAppearance = appearance49;
			appearance50.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance50.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance50.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid2.DisplayLayout.Override.RowSelectorAppearance = appearance50;
			appearance51.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance51.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance51.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid2.DisplayLayout.Override.SelectedRowAppearance = appearance51;
			this.ultraGrid2.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.ultraGrid2.Location = new System.Drawing.Point(16, 17);
			this.ultraGrid2.Name = "ultraGrid2";
			this.ultraGrid2.Size = new System.Drawing.Size(632, 328);
			this.ultraGrid2.TabIndex = 1;
			this.ultraGrid2.AfterRowInsert += new Infragistics.Win.UltraWinGrid.RowEventHandler(this.ultraGrid2_AfterRowInsert);
			// 
			// cdsContacto
			// 
			this.cdsContacto.BindingContextCtrl = this;
			this.cdsContacto.DataLibrary = "LibFacturacion";
			this.cdsContacto.DataLibraryUrl = "";
			this.cdsContacto.DataSetDef = "dsClienteTabla";
			this.cdsContacto.FillOnRequest = false;
			this.cdsContacto.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsContacto.Name = "cdsContacto";
			this.cdsContacto.SchemaClassName = "LibFacturacion.DataClass";
			this.cdsContacto.SchemaDef = null;
			// 
			// ultraTabPageControl2
			// 
			this.ultraTabPageControl2.Controls.Add(this.btBanco);
			this.ultraTabPageControl2.Controls.Add(this.cmbBanco);
			this.ultraTabPageControl2.Controls.Add(this.label53);
			this.ultraTabPageControl2.Controls.Add(this.ultraNumericEditor3);
			this.ultraTabPageControl2.Controls.Add(this.label49);
			this.ultraTabPageControl2.Controls.Add(this.ultraNumericEditor2);
			this.ultraTabPageControl2.Controls.Add(this.ultraCheckEditor3);
			this.ultraTabPageControl2.Controls.Add(this.label48);
			this.ultraTabPageControl2.Controls.Add(this.chkEntrega);
			this.ultraTabPageControl2.Controls.Add(this.ultraCheckEditor2);
			this.ultraTabPageControl2.Controls.Add(this.chkDesCompras);
			this.ultraTabPageControl2.Controls.Add(this.groupBox2);
			this.ultraTabPageControl2.Controls.Add(this.label47);
			this.ultraTabPageControl2.Controls.Add(this.ultraTextEditor20);
			this.ultraTabPageControl2.Controls.Add(this.label45);
			this.ultraTabPageControl2.Controls.Add(this.ultraTextEditor18);
			this.ultraTabPageControl2.Controls.Add(this.label46);
			this.ultraTabPageControl2.Controls.Add(this.ultraTextEditor19);
			this.ultraTabPageControl2.Controls.Add(this.label42);
			this.ultraTabPageControl2.Controls.Add(this.ultraTextEditor15);
			this.ultraTabPageControl2.Controls.Add(this.label35);
			this.ultraTabPageControl2.Controls.Add(this.ultraTextEditor7);
			this.ultraTabPageControl2.Controls.Add(this.groupBox1);
			this.ultraTabPageControl2.Location = new System.Drawing.Point(-10000, -10000);
			this.ultraTabPageControl2.Name = "ultraTabPageControl2";
			this.ultraTabPageControl2.Size = new System.Drawing.Size(668, 371);
			// 
			// btBanco
			// 
			this.btBanco.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.btBanco.Location = new System.Drawing.Point(616, 327);
			this.btBanco.Name = "btBanco";
			this.btBanco.Size = new System.Drawing.Size(24, 26);
			this.btBanco.TabIndex = 152;
			this.btBanco.Text = "...";
			this.btBanco.Click += new System.EventHandler(this.btBanco_Click);
			// 
			// cmbBanco
			// 
			appearance52.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbBanco.Appearance = appearance52;
			this.cmbBanco.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbBanco.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbBanco.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsCliente, "Cliente.idBanco"));
			ultraGridColumn34.Header.VisiblePosition = 0;
			ultraGridColumn35.Header.VisiblePosition = 1;
			ultraGridBand15.Columns.AddRange(new object[] {
																											ultraGridColumn34,
																											ultraGridColumn35});
			this.cmbBanco.DisplayLayout.BandsSerializer.Add(ultraGridBand15);
			this.cmbBanco.DisplayMember = "Nombre";
			this.cmbBanco.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbBanco.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbBanco.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbBanco.Location = new System.Drawing.Point(448, 327);
			this.cmbBanco.Name = "cmbBanco";
			this.cmbBanco.Size = new System.Drawing.Size(168, 21);
			this.cmbBanco.TabIndex = 151;
			this.cmbBanco.ValueMember = "idBanco";
			// 
			// label53
			// 
			this.label53.AutoSize = true;
			this.label53.BackColor = System.Drawing.Color.Transparent;
			this.label53.Location = new System.Drawing.Point(384, 327);
			this.label53.Name = "label53";
			this.label53.Size = new System.Drawing.Size(34, 16);
			this.label53.TabIndex = 150;
			this.label53.Text = "Banco";
			// 
			// ultraNumericEditor3
			// 
			this.ultraNumericEditor3.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsCliente, "Cliente.ProveMaster"));
			this.ultraNumericEditor3.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.ultraNumericEditor3.FormatString = "";
			this.ultraNumericEditor3.Location = new System.Drawing.Point(272, 327);
			this.ultraNumericEditor3.Name = "ultraNumericEditor3";
			this.ultraNumericEditor3.PromptChar = ' ';
			this.ultraNumericEditor3.Size = new System.Drawing.Size(80, 22);
			this.ultraNumericEditor3.TabIndex = 149;
			// 
			// label49
			// 
			this.label49.AutoSize = true;
			this.label49.BackColor = System.Drawing.Color.Transparent;
			this.label49.Location = new System.Drawing.Point(200, 327);
			this.label49.Name = "label49";
			this.label49.Size = new System.Drawing.Size(63, 16);
			this.label49.TabIndex = 148;
			this.label49.Text = "Prov Master";
			// 
			// ultraNumericEditor2
			// 
			this.ultraNumericEditor2.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsCliente, "Cliente.PorComT"));
			this.ultraNumericEditor2.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.ultraNumericEditor2.FormatString = "#,##0.00";
			this.ultraNumericEditor2.Location = new System.Drawing.Point(96, 327);
			this.ultraNumericEditor2.Name = "ultraNumericEditor2";
			this.ultraNumericEditor2.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.ultraNumericEditor2.PromptChar = ' ';
			this.ultraNumericEditor2.Size = new System.Drawing.Size(80, 22);
			this.ultraNumericEditor2.TabIndex = 147;
			// 
			// ultraCheckEditor3
			// 
			this.ultraCheckEditor3.DataBindings.Add(new System.Windows.Forms.Binding("CheckedValue", this.cdsCliente, "Cliente.NetoValorT"));
			this.ultraCheckEditor3.Location = new System.Drawing.Point(560, 302);
			this.ultraCheckEditor3.Name = "ultraCheckEditor3";
			this.ultraCheckEditor3.Size = new System.Drawing.Size(80, 25);
			this.ultraCheckEditor3.TabIndex = 146;
			this.ultraCheckEditor3.Text = "Valor Neto";
			// 
			// label48
			// 
			this.label48.AutoSize = true;
			this.label48.BackColor = System.Drawing.Color.Transparent;
			this.label48.Location = new System.Drawing.Point(24, 327);
			this.label48.Name = "label48";
			this.label48.Size = new System.Drawing.Size(47, 16);
			this.label48.TabIndex = 145;
			this.label48.Text = "Por Com";
			// 
			// chkEntrega
			// 
			appearance53.ForeColorDisabled = System.Drawing.Color.Black;
			this.chkEntrega.Appearance = appearance53;
			this.chkEntrega.BackColor = System.Drawing.Color.Transparent;
			this.chkEntrega.DataBindings.Add(new System.Windows.Forms.Binding("CheckedValue", this.cdsCliente, "Cliente.Entrega"));
			this.chkEntrega.Location = new System.Drawing.Point(384, 302);
			this.chkEntrega.Name = "chkEntrega";
			this.chkEntrega.Size = new System.Drawing.Size(168, 21);
			this.chkEntrega.TabIndex = 143;
			this.chkEntrega.Text = "Recibe Mercaderia";
			// 
			// ultraCheckEditor2
			// 
			appearance54.ForeColorDisabled = System.Drawing.Color.Black;
			this.ultraCheckEditor2.Appearance = appearance54;
			this.ultraCheckEditor2.BackColor = System.Drawing.Color.Transparent;
			this.ultraCheckEditor2.DataBindings.Add(new System.Windows.Forms.Binding("CheckedValue", this.cdsCliente, "Cliente.DesconCliente"));
			this.ultraCheckEditor2.Location = new System.Drawing.Point(200, 302);
			this.ultraCheckEditor2.Name = "ultraCheckEditor2";
			this.ultraCheckEditor2.Size = new System.Drawing.Size(168, 21);
			this.ultraCheckEditor2.TabIndex = 142;
			this.ultraCheckEditor2.Text = "Descontinuado en Ventas";
			// 
			// chkDesCompras
			// 
			appearance55.ForeColorDisabled = System.Drawing.Color.Black;
			this.chkDesCompras.Appearance = appearance55;
			this.chkDesCompras.BackColor = System.Drawing.Color.Transparent;
			this.chkDesCompras.DataBindings.Add(new System.Windows.Forms.Binding("CheckedValue", this.cdsCliente, "Cliente.DesconProv"));
			this.chkDesCompras.Location = new System.Drawing.Point(24, 302);
			this.chkDesCompras.Name = "chkDesCompras";
			this.chkDesCompras.Size = new System.Drawing.Size(168, 21);
			this.chkDesCompras.TabIndex = 141;
			this.chkDesCompras.Text = "Descontinuado en Compras";
			// 
			// groupBox2
			// 
			this.groupBox2.BackColor = System.Drawing.Color.Transparent;
			this.groupBox2.Controls.Add(this.ultraTextEditor17);
			this.groupBox2.Controls.Add(this.label41);
			this.groupBox2.Controls.Add(this.label40);
			this.groupBox2.Controls.Add(this.label44);
			this.groupBox2.Controls.Add(this.ultraTextEditor14);
			this.groupBox2.Controls.Add(this.ultraTextEditor13);
			this.groupBox2.Controls.Add(this.label36);
			this.groupBox2.Controls.Add(this.ultraTextEditor8);
			this.groupBox2.Location = new System.Drawing.Point(16, 86);
			this.groupBox2.Name = "groupBox2";
			this.groupBox2.Size = new System.Drawing.Size(640, 86);
			this.groupBox2.TabIndex = 2;
			this.groupBox2.TabStop = false;
			this.groupBox2.Text = " Conyuge ";
			// 
			// ultraTextEditor17
			// 
			appearance56.ForeColorDisabled = System.Drawing.Color.Black;
			this.ultraTextEditor17.Appearance = appearance56;
			this.ultraTextEditor17.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsCliente, "Cliente.ConyTelefono"));
			this.ultraTextEditor17.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.ultraTextEditor17.Location = new System.Drawing.Point(112, 52);
			this.ultraTextEditor17.Name = "ultraTextEditor17";
			this.ultraTextEditor17.Size = new System.Drawing.Size(192, 22);
			this.ultraTextEditor17.TabIndex = 2;
			// 
			// label41
			// 
			this.label41.AutoSize = true;
			this.label41.BackColor = System.Drawing.Color.Transparent;
			this.label41.Location = new System.Drawing.Point(344, 26);
			this.label41.Name = "label41";
			this.label41.Size = new System.Drawing.Size(50, 16);
			this.label41.TabIndex = 15;
			this.label41.Text = "Dirección";
			// 
			// label40
			// 
			this.label40.AutoSize = true;
			this.label40.BackColor = System.Drawing.Color.Transparent;
			this.label40.Location = new System.Drawing.Point(344, 52);
			this.label40.Name = "label40";
			this.label40.Size = new System.Drawing.Size(38, 16);
			this.label40.TabIndex = 13;
			this.label40.Text = "Celular";
			// 
			// label44
			// 
			this.label44.AutoSize = true;
			this.label44.BackColor = System.Drawing.Color.Transparent;
			this.label44.Location = new System.Drawing.Point(8, 52);
			this.label44.Name = "label44";
			this.label44.Size = new System.Drawing.Size(48, 16);
			this.label44.TabIndex = 21;
			this.label44.Text = "Teléfono";
			// 
			// ultraTextEditor14
			// 
			appearance57.ForeColorDisabled = System.Drawing.Color.Black;
			this.ultraTextEditor14.Appearance = appearance57;
			this.ultraTextEditor14.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsCliente, "Cliente.ConyDireccion"));
			this.ultraTextEditor14.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.ultraTextEditor14.Location = new System.Drawing.Point(432, 26);
			this.ultraTextEditor14.Name = "ultraTextEditor14";
			this.ultraTextEditor14.Size = new System.Drawing.Size(192, 22);
			this.ultraTextEditor14.TabIndex = 1;
			// 
			// ultraTextEditor13
			// 
			appearance58.ForeColorDisabled = System.Drawing.Color.Black;
			this.ultraTextEditor13.Appearance = appearance58;
			this.ultraTextEditor13.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsCliente, "Cliente.ConyCelular"));
			this.ultraTextEditor13.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.ultraTextEditor13.Location = new System.Drawing.Point(432, 52);
			this.ultraTextEditor13.Name = "ultraTextEditor13";
			this.ultraTextEditor13.Size = new System.Drawing.Size(192, 22);
			this.ultraTextEditor13.TabIndex = 3;
			// 
			// label36
			// 
			this.label36.AutoSize = true;
			this.label36.BackColor = System.Drawing.Color.Transparent;
			this.label36.Location = new System.Drawing.Point(8, 26);
			this.label36.Name = "label36";
			this.label36.Size = new System.Drawing.Size(48, 16);
			this.label36.TabIndex = 5;
			this.label36.Text = "Conyuge";
			// 
			// ultraTextEditor8
			// 
			appearance59.ForeColorDisabled = System.Drawing.Color.Black;
			this.ultraTextEditor8.Appearance = appearance59;
			this.ultraTextEditor8.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsCliente, "Cliente.Conyuge"));
			this.ultraTextEditor8.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.ultraTextEditor8.Location = new System.Drawing.Point(112, 26);
			this.ultraTextEditor8.Name = "ultraTextEditor8";
			this.ultraTextEditor8.Size = new System.Drawing.Size(192, 22);
			this.ultraTextEditor8.TabIndex = 0;
			// 
			// label47
			// 
			this.label47.AutoSize = true;
			this.label47.BackColor = System.Drawing.Color.Transparent;
			this.label47.Location = new System.Drawing.Point(24, 17);
			this.label47.Name = "label47";
			this.label47.Size = new System.Drawing.Size(53, 16);
			this.label47.TabIndex = 27;
			this.label47.Text = "Ref. Local";
			// 
			// ultraTextEditor20
			// 
			appearance60.ForeColorDisabled = System.Drawing.Color.Black;
			this.ultraTextEditor20.Appearance = appearance60;
			this.ultraTextEditor20.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsCliente, "Cliente.RefLocal"));
			this.ultraTextEditor20.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.ultraTextEditor20.Location = new System.Drawing.Point(128, 17);
			this.ultraTextEditor20.Name = "ultraTextEditor20";
			this.ultraTextEditor20.Size = new System.Drawing.Size(512, 22);
			this.ultraTextEditor20.TabIndex = 0;
			// 
			// label45
			// 
			this.label45.AutoSize = true;
			this.label45.BackColor = System.Drawing.Color.Transparent;
			this.label45.Location = new System.Drawing.Point(360, 267);
			this.label45.Name = "label45";
			this.label45.Size = new System.Drawing.Size(36, 16);
			this.label45.TabIndex = 25;
			this.label45.Text = "Sector";
			// 
			// ultraTextEditor18
			// 
			appearance61.ForeColorDisabled = System.Drawing.Color.Black;
			this.ultraTextEditor18.Appearance = appearance61;
			this.ultraTextEditor18.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsCliente, "Cliente.Sector"));
			this.ultraTextEditor18.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.ultraTextEditor18.Location = new System.Drawing.Point(448, 267);
			this.ultraTextEditor18.Name = "ultraTextEditor18";
			this.ultraTextEditor18.Size = new System.Drawing.Size(192, 22);
			this.ultraTextEditor18.TabIndex = 4;
			// 
			// label46
			// 
			this.label46.AutoSize = true;
			this.label46.BackColor = System.Drawing.Color.Transparent;
			this.label46.Location = new System.Drawing.Point(24, 267);
			this.label46.Name = "label46";
			this.label46.Size = new System.Drawing.Size(52, 16);
			this.label46.TabIndex = 23;
			this.label46.Text = "Parroquia";
			// 
			// ultraTextEditor19
			// 
			appearance62.ForeColorDisabled = System.Drawing.Color.Black;
			this.ultraTextEditor19.Appearance = appearance62;
			this.ultraTextEditor19.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsCliente, "Cliente.Parroquia"));
			this.ultraTextEditor19.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.ultraTextEditor19.Location = new System.Drawing.Point(128, 267);
			this.ultraTextEditor19.Name = "ultraTextEditor19";
			this.ultraTextEditor19.Size = new System.Drawing.Size(192, 22);
			this.ultraTextEditor19.TabIndex = 3;
			// 
			// label42
			// 
			this.label42.AutoSize = true;
			this.label42.BackColor = System.Drawing.Color.Transparent;
			this.label42.Location = new System.Drawing.Point(360, 52);
			this.label42.Name = "label42";
			this.label42.Size = new System.Drawing.Size(76, 16);
			this.label42.TabIndex = 17;
			this.label42.Text = "Cuenta Ahorro";
			// 
			// ultraTextEditor15
			// 
			appearance63.ForeColorDisabled = System.Drawing.Color.Black;
			this.ultraTextEditor15.Appearance = appearance63;
			this.ultraTextEditor15.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsCliente, "Cliente.CuentaAh"));
			this.ultraTextEditor15.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.ultraTextEditor15.Location = new System.Drawing.Point(448, 52);
			this.ultraTextEditor15.Name = "ultraTextEditor15";
			this.ultraTextEditor15.Size = new System.Drawing.Size(192, 22);
			this.ultraTextEditor15.TabIndex = 2;
			// 
			// label35
			// 
			this.label35.AutoSize = true;
			this.label35.BackColor = System.Drawing.Color.Transparent;
			this.label35.Location = new System.Drawing.Point(24, 52);
			this.label35.Name = "label35";
			this.label35.Size = new System.Drawing.Size(89, 16);
			this.label35.TabIndex = 3;
			this.label35.Text = "Cuenta Corriente";
			// 
			// ultraTextEditor7
			// 
			appearance64.ForeColorDisabled = System.Drawing.Color.Black;
			this.ultraTextEditor7.Appearance = appearance64;
			this.ultraTextEditor7.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsCliente, "Cliente.CuentaCorr"));
			this.ultraTextEditor7.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.ultraTextEditor7.Location = new System.Drawing.Point(128, 52);
			this.ultraTextEditor7.Name = "ultraTextEditor7";
			this.ultraTextEditor7.Size = new System.Drawing.Size(192, 22);
			this.ultraTextEditor7.TabIndex = 1;
			// 
			// groupBox1
			// 
			this.groupBox1.BackColor = System.Drawing.Color.Transparent;
			this.groupBox1.Controls.Add(this.label37);
			this.groupBox1.Controls.Add(this.ultraTextEditor9);
			this.groupBox1.Controls.Add(this.ultraTextEditor12);
			this.groupBox1.Controls.Add(this.label39);
			this.groupBox1.Controls.Add(this.label43);
			this.groupBox1.Controls.Add(this.label38);
			this.groupBox1.Controls.Add(this.ultraTextEditor16);
			this.groupBox1.Controls.Add(this.ultraTextEditor11);
			this.groupBox1.Location = new System.Drawing.Point(16, 172);
			this.groupBox1.Name = "groupBox1";
			this.groupBox1.Size = new System.Drawing.Size(640, 86);
			this.groupBox1.TabIndex = 3;
			this.groupBox1.TabStop = false;
			this.groupBox1.Text = " Referencia ";
			// 
			// label37
			// 
			this.label37.AutoSize = true;
			this.label37.BackColor = System.Drawing.Color.Transparent;
			this.label37.Location = new System.Drawing.Point(8, 52);
			this.label37.Name = "label37";
			this.label37.Size = new System.Drawing.Size(50, 16);
			this.label37.TabIndex = 7;
			this.label37.Text = "Dirección";
			// 
			// ultraTextEditor9
			// 
			appearance65.ForeColorDisabled = System.Drawing.Color.Black;
			this.ultraTextEditor9.Appearance = appearance65;
			this.ultraTextEditor9.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsCliente, "Cliente.RefDireccion"));
			this.ultraTextEditor9.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.ultraTextEditor9.Location = new System.Drawing.Point(112, 52);
			this.ultraTextEditor9.Name = "ultraTextEditor9";
			this.ultraTextEditor9.Size = new System.Drawing.Size(192, 22);
			this.ultraTextEditor9.TabIndex = 2;
			// 
			// ultraTextEditor12
			// 
			appearance66.ForeColorDisabled = System.Drawing.Color.Black;
			this.ultraTextEditor12.Appearance = appearance66;
			this.ultraTextEditor12.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsCliente, "Cliente.RefTelefono"));
			this.ultraTextEditor12.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.ultraTextEditor12.Location = new System.Drawing.Point(432, 26);
			this.ultraTextEditor12.Name = "ultraTextEditor12";
			this.ultraTextEditor12.Size = new System.Drawing.Size(192, 22);
			this.ultraTextEditor12.TabIndex = 1;
			// 
			// label39
			// 
			this.label39.AutoSize = true;
			this.label39.BackColor = System.Drawing.Color.Transparent;
			this.label39.Location = new System.Drawing.Point(344, 26);
			this.label39.Name = "label39";
			this.label39.Size = new System.Drawing.Size(48, 16);
			this.label39.TabIndex = 11;
			this.label39.Text = "Teléfono";
			// 
			// label43
			// 
			this.label43.AutoSize = true;
			this.label43.BackColor = System.Drawing.Color.Transparent;
			this.label43.Location = new System.Drawing.Point(8, 26);
			this.label43.Name = "label43";
			this.label43.Size = new System.Drawing.Size(44, 16);
			this.label43.TabIndex = 19;
			this.label43.Text = "Nombre";
			// 
			// label38
			// 
			this.label38.AutoSize = true;
			this.label38.BackColor = System.Drawing.Color.Transparent;
			this.label38.Location = new System.Drawing.Point(344, 52);
			this.label38.Name = "label38";
			this.label38.Size = new System.Drawing.Size(33, 16);
			this.label38.TabIndex = 9;
			this.label38.Text = "Póliza";
			// 
			// ultraTextEditor16
			// 
			appearance67.ForeColorDisabled = System.Drawing.Color.Black;
			this.ultraTextEditor16.Appearance = appearance67;
			this.ultraTextEditor16.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsCliente, "Cliente.Referencia"));
			this.ultraTextEditor16.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.ultraTextEditor16.Location = new System.Drawing.Point(112, 26);
			this.ultraTextEditor16.Name = "ultraTextEditor16";
			this.ultraTextEditor16.Size = new System.Drawing.Size(192, 22);
			this.ultraTextEditor16.TabIndex = 0;
			// 
			// ultraTextEditor11
			// 
			appearance68.ForeColorDisabled = System.Drawing.Color.Black;
			this.ultraTextEditor11.Appearance = appearance68;
			this.ultraTextEditor11.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsCliente, "Cliente.Poliza"));
			this.ultraTextEditor11.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.ultraTextEditor11.Location = new System.Drawing.Point(432, 52);
			this.ultraTextEditor11.Name = "ultraTextEditor11";
			this.ultraTextEditor11.Size = new System.Drawing.Size(192, 22);
			this.ultraTextEditor11.TabIndex = 3;
			// 
			// ultraTabPageControl3
			// 
			this.ultraTabPageControl3.Controls.Add(this.btExcel);
			this.ultraTabPageControl3.Controls.Add(this.btTotales);
			this.ultraTabPageControl3.Controls.Add(this.btFactura);
			this.ultraTabPageControl3.Controls.Add(this.btVencido);
			this.ultraTabPageControl3.Controls.Add(this.btListadoPendiente);
			this.ultraTabPageControl3.Controls.Add(this.ultraGrid1);
			this.ultraTabPageControl3.Location = new System.Drawing.Point(-10000, -10000);
			this.ultraTabPageControl3.Name = "ultraTabPageControl3";
			this.ultraTabPageControl3.Size = new System.Drawing.Size(668, 371);
			// 
			// btExcel
			// 
			this.btExcel.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btExcel.Location = new System.Drawing.Point(567, 9);
			this.btExcel.Name = "btExcel";
			this.btExcel.Size = new System.Drawing.Size(88, 24);
			this.btExcel.TabIndex = 13;
			this.btExcel.Text = "Excel";
			this.btExcel.Click += new System.EventHandler(this.btExcel_Click);
			// 
			// btTotales
			// 
			this.btTotales.Location = new System.Drawing.Point(264, 9);
			this.btTotales.Name = "btTotales";
			this.btTotales.Size = new System.Drawing.Size(120, 24);
			this.btTotales.TabIndex = 4;
			this.btTotales.Text = "Totales";
			this.btTotales.Click += new System.EventHandler(this.btTotales_Click);
			// 
			// btFactura
			// 
			this.btFactura.Location = new System.Drawing.Point(392, 9);
			this.btFactura.Name = "btFactura";
			this.btFactura.Size = new System.Drawing.Size(120, 24);
			this.btFactura.TabIndex = 3;
			this.btFactura.Text = "Factura";
			this.btFactura.Click += new System.EventHandler(this.btFactura_Click);
			this.btFactura.MouseUp += new System.Windows.Forms.MouseEventHandler(this.btFactura_MouseUp);
			// 
			// btVencido
			// 
			this.btVencido.Location = new System.Drawing.Point(136, 9);
			this.btVencido.Name = "btVencido";
			this.btVencido.Size = new System.Drawing.Size(120, 24);
			this.btVencido.TabIndex = 2;
			this.btVencido.Text = "Vencido";
			this.btVencido.Click += new System.EventHandler(this.btVencido_Click);
			// 
			// btListadoPendiente
			// 
			this.btListadoPendiente.Location = new System.Drawing.Point(8, 9);
			this.btListadoPendiente.Name = "btListadoPendiente";
			this.btListadoPendiente.Size = new System.Drawing.Size(120, 24);
			this.btListadoPendiente.TabIndex = 1;
			this.btListadoPendiente.Text = "Listado de Pendiente";
			this.btListadoPendiente.Click += new System.EventHandler(this.btListadoPendiente_Click);
			// 
			// ultraGrid1
			// 
			this.ultraGrid1.Cursor = System.Windows.Forms.Cursors.Default;
			appearance69.BackColor = System.Drawing.Color.White;
			this.ultraGrid1.DisplayLayout.Appearance = appearance69;
			appearance70.BackColor = System.Drawing.Color.Transparent;
			this.ultraGrid1.DisplayLayout.Override.CardAreaAppearance = appearance70;
			appearance71.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance71.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance71.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance71.FontData.BoldAsString = "True";
			appearance71.FontData.Name = "Arial";
			appearance71.FontData.SizeInPoints = 10F;
			appearance71.ForeColor = System.Drawing.Color.White;
			appearance71.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.ultraGrid1.DisplayLayout.Override.HeaderAppearance = appearance71;
			this.ultraGrid1.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti;
			appearance72.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance72.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance72.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.RowSelectorAppearance = appearance72;
			appearance73.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance73.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance73.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.SelectedRowAppearance = appearance73;
			this.ultraGrid1.DisplayLayout.ViewStyleBand = Infragistics.Win.UltraWinGrid.ViewStyleBand.OutlookGroupBy;
			this.ultraGrid1.Location = new System.Drawing.Point(8, 43);
			this.ultraGrid1.Name = "ultraGrid1";
			this.ultraGrid1.Size = new System.Drawing.Size(648, 310);
			this.ultraGrid1.TabIndex = 0;
			// 
			// ultraTabPageControl4
			// 
			this.ultraTabPageControl4.Controls.Add(this.ultraButton1);
			this.ultraTabPageControl4.Controls.Add(this.btListadoItems);
			this.ultraTabPageControl4.Controls.Add(this.grdItems);
			this.ultraTabPageControl4.Location = new System.Drawing.Point(-10000, -10000);
			this.ultraTabPageControl4.Name = "ultraTabPageControl4";
			this.ultraTabPageControl4.Size = new System.Drawing.Size(668, 371);
			// 
			// ultraButton1
			// 
			this.ultraButton1.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.ultraButton1.Location = new System.Drawing.Point(559, 9);
			this.ultraButton1.Name = "ultraButton1";
			this.ultraButton1.Size = new System.Drawing.Size(88, 24);
			this.ultraButton1.TabIndex = 14;
			this.ultraButton1.Text = "Excel";
			this.ultraButton1.Click += new System.EventHandler(this.ultraButton1_Click);
			// 
			// btListadoItems
			// 
			this.btListadoItems.Location = new System.Drawing.Point(24, 9);
			this.btListadoItems.Name = "btListadoItems";
			this.btListadoItems.Size = new System.Drawing.Size(136, 24);
			this.btListadoItems.TabIndex = 2;
			this.btListadoItems.Text = "Listado de Items";
			this.btListadoItems.Click += new System.EventHandler(this.btListadoItems_Click);
			// 
			// grdItems
			// 
			this.grdItems.Cursor = System.Windows.Forms.Cursors.Default;
			appearance74.BackColor = System.Drawing.Color.White;
			this.grdItems.DisplayLayout.Appearance = appearance74;
			appearance75.BackColor = System.Drawing.Color.Transparent;
			this.grdItems.DisplayLayout.Override.CardAreaAppearance = appearance75;
			appearance76.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance76.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance76.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance76.FontData.BoldAsString = "True";
			appearance76.FontData.Name = "Arial";
			appearance76.FontData.SizeInPoints = 10F;
			appearance76.ForeColor = System.Drawing.Color.White;
			appearance76.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.grdItems.DisplayLayout.Override.HeaderAppearance = appearance76;
			this.grdItems.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti;
			appearance77.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance77.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance77.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.grdItems.DisplayLayout.Override.RowSelectorAppearance = appearance77;
			appearance78.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance78.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance78.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.grdItems.DisplayLayout.Override.SelectedRowAppearance = appearance78;
			this.grdItems.DisplayLayout.ViewStyleBand = Infragistics.Win.UltraWinGrid.ViewStyleBand.OutlookGroupBy;
			this.grdItems.Location = new System.Drawing.Point(16, 43);
			this.grdItems.Name = "grdItems";
			this.grdItems.Size = new System.Drawing.Size(640, 302);
			this.grdItems.TabIndex = 1;
			// 
			// errorProvider
			// 
			this.errorProvider.ContainerControl = this;
			// 
			// txtIdCliente
			// 
			this.txtIdCliente.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsCliente, "Cliente.idCliente"));
			this.txtIdCliente.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtIdCliente.Location = new System.Drawing.Point(592, 431);
			this.txtIdCliente.Name = "txtIdCliente";
			this.txtIdCliente.Size = new System.Drawing.Size(48, 22);
			this.txtIdCliente.TabIndex = 76;
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
			// tbBasico
			// 
			this.tbBasico.Controls.Add(this.ultraTabSharedControlsPage1);
			this.tbBasico.Controls.Add(this.ultraTabPageControl1);
			this.tbBasico.Controls.Add(this.tpComplemento);
			this.tbBasico.Controls.Add(this.ultraTabPageControl2);
			this.tbBasico.Controls.Add(this.ultraTabPageControl3);
			this.tbBasico.Controls.Add(this.ultraTabPageControl4);
			this.tbBasico.Controls.Add(this.ultraTabPageControl5);
			this.tbBasico.Location = new System.Drawing.Point(16, 17);
			this.tbBasico.Name = "tbBasico";
			this.tbBasico.SharedControlsPage = this.ultraTabSharedControlsPage1;
			this.tbBasico.Size = new System.Drawing.Size(672, 397);
			this.tbBasico.TabIndex = 83;
			ultraTab1.TabPage = this.ultraTabPageControl1;
			ultraTab1.Text = "Básico";
			ultraTab2.TabPage = this.tpComplemento;
			ultraTab2.Text = "Complemento";
			ultraTab3.TabPage = this.ultraTabPageControl5;
			ultraTab3.Text = "Contactos";
			ultraTab4.TabPage = this.ultraTabPageControl2;
			ultraTab4.Text = "Referencias";
			ultraTab5.TabPage = this.ultraTabPageControl3;
			ultraTab5.Text = "Análisis de cartera";
			ultraTab6.TabPage = this.ultraTabPageControl4;
			ultraTab6.Text = "Análisis de Items";
			this.tbBasico.Tabs.AddRange(new Infragistics.Win.UltraWinTabControl.UltraTab[] {
																																											 ultraTab1,
																																											 ultraTab2,
																																											 ultraTab3,
																																											 ultraTab4,
																																											 ultraTab5,
																																											 ultraTab6});
			// 
			// ultraTabSharedControlsPage1
			// 
			this.ultraTabSharedControlsPage1.Location = new System.Drawing.Point(-10000, -10000);
			this.ultraTabSharedControlsPage1.Name = "ultraTabSharedControlsPage1";
			this.ultraTabSharedControlsPage1.Size = new System.Drawing.Size(668, 371);
			// 
			// btCRMClientes
			// 
			appearance79.Image = 5;
			this.btCRMClientes.Appearance = appearance79;
			this.btCRMClientes.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.btCRMClientes.ImageList = this.imageList;
			this.btCRMClientes.Location = new System.Drawing.Point(352, 431);
			this.btCRMClientes.Name = "btCRMClientes";
			this.btCRMClientes.Size = new System.Drawing.Size(24, 25);
			this.btCRMClientes.TabIndex = 84;
			this.btCRMClientes.Text = ">";
			this.toolTip1.SetToolTip(this.btCRMClientes, "Análisis de Cliente");
			this.btCRMClientes.Click += new System.EventHandler(this.btCRMClientes_Click);
			// 
			// imageList
			// 
			this.imageList.ImageSize = new System.Drawing.Size(16, 16);
			this.imageList.ImageStream = ((System.Windows.Forms.ImageListStreamer)(resources.GetObject("imageList.ImageStream")));
			this.imageList.TransparentColor = System.Drawing.Color.Transparent;
			// 
			// btAuditoria
			// 
			appearance80.Image = ((object)(resources.GetObject("appearance80.Image")));
			this.btAuditoria.Appearance = appearance80;
			this.btAuditoria.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.btAuditoria.Location = new System.Drawing.Point(328, 431);
			this.btAuditoria.Name = "btAuditoria";
			this.btAuditoria.Size = new System.Drawing.Size(24, 25);
			this.btAuditoria.TabIndex = 85;
			this.toolTip1.SetToolTip(this.btAuditoria, "Auditoria de Cliente");
			this.btAuditoria.Click += new System.EventHandler(this.btAuditoria_Click);
			this.btAuditoria.MouseUp += new System.Windows.Forms.MouseEventHandler(this.btAuditoria_MouseUp);
			// 
			// cdsPais
			// 
			this.cdsPais.BindingContextCtrl = this;
			this.cdsPais.DataLibrary = "LibFacturacion";
			this.cdsPais.DataLibraryUrl = "";
			this.cdsPais.DataSetDef = "dsCliente";
			this.cdsPais.FillOnRequest = false;
			this.cdsPais.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsPais.Name = "cdsPais";
			this.cdsPais.SchemaClassName = "LibFacturacion.DataClass";
			this.cdsPais.SchemaDef = null;
			// 
			// barraDato1
			// 
			this.barraDato1.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.barraDato1.BarraMovimiento = true;
			this.barraDato1.DataMember = "Cliente";
			this.barraDato1.DataNombreId = "idCliente";
			this.barraDato1.DataOrden = "";
			this.barraDato1.DataSource = this.cdsCliente;
			this.barraDato1.DataTabla = "";
			this.barraDato1.DataTablaHija = "";
			this.barraDato1.Location = new System.Drawing.Point(16, 432);
			this.barraDato1.Name = "barraDato1";
			this.barraDato1.Size = new System.Drawing.Size(304, 26);
			this.barraDato1.TabIndex = 86;
			this.barraDato1.VisibleBorrar = true;
			this.barraDato1.VisibleBuscar = true;
			this.barraDato1.VisibleEditar = true;
			this.barraDato1.VisibleImprimir = true;
			this.barraDato1.VisibleNuevo = true;
			// 
			// Cliente1
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 14);
			this.ClientSize = new System.Drawing.Size(704, 470);
			this.Controls.Add(this.barraDato1);
			this.Controls.Add(this.btAuditoria);
			this.Controls.Add(this.btCRMClientes);
			this.Controls.Add(this.tbBasico);
			this.Controls.Add(this.txtIdCliente);
			this.MaximizeBox = false;
			this.Name = "Cliente1";
			this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
			this.Text = "Clientes";
			this.Closing += new System.ComponentModel.CancelEventHandler(this.Cliente_Closing);
			this.Load += new System.EventHandler(this.Cliente_Load);
			this.ultraTabPageControl1.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.txtPais)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsCliente)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbPais)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraTextEditor6)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraCalendarCombo1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbClientePublicidad)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraTextEditor2)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraTextEditor1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraTextEditor5)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCelular)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbPrecio)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbSector)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtEMail)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraTextEditor10)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCiudad)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbGrupo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbTipoRuc)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNombre)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDireccion)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtTelefono)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraTextEditor3)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtRuc)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCodigo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbFecha)).EndInit();
			this.tpComplemento.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.radioGroup1.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraTextEditor21)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraTextEditor4)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCredTribut)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbFormaPago)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtSerie)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtAutRetencion)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbFechaCaduca)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbPorcIvaRet)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbDepende)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCobrador)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbVendedor)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtAutImprenta)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtAutFactura)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.spnDescCredito)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.spnDescuento)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtValorCredito)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraNumericEditor1)).EndInit();
			this.ultraTabPageControl5.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid2)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsContacto)).EndInit();
			this.ultraTabPageControl2.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.cmbBanco)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraNumericEditor3)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraNumericEditor2)).EndInit();
			this.groupBox2.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.ultraTextEditor17)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraTextEditor14)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraTextEditor13)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraTextEditor8)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraTextEditor20)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraTextEditor18)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraTextEditor19)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraTextEditor15)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraTextEditor7)).EndInit();
			this.groupBox1.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.ultraTextEditor9)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraTextEditor12)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraTextEditor16)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraTextEditor11)).EndInit();
			this.ultraTabPageControl3.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).EndInit();
			this.ultraTabPageControl4.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.grdItems)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdCliente)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.tbBasico)).EndInit();
			this.tbBasico.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.cdsPais)).EndInit();
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
		private Compra miCompra;
		#endregion Variables Globales
		#region Pantalla Cliente
		public Cliente1(bool bTipo)
		{
			InitializeComponent();
			TipoCliente = bTipo;
		}
		public Cliente1(bool bTipo, int id)
		{
			InitializeComponent();
			TipoCliente = bTipo;
			IdCliente = id;
		}
		private void UnloadMe()
		{
			this.Close();
		}
		private void Cliente_Load(object sender, System.EventArgs e)
		{
			string stCodCliente = "71";
			if (TipoCliente) stCodCliente = "65";
		
			if (!Funcion.Permiso(stCodCliente,cdsSeteoF))
			{
				stCodCliente = "Proveedores";
				if (TipoCliente) stCodCliente = "Clientes";

				MessageBox.Show("No puede ingresar a " + stCodCliente, "Informacion");
				BeginInvoke(new MethodInvoker(UnloadMe));
				return;
			}

			Cursor = Cursors.Default;
			#region Auditoria
			string stAudita = "Exec AuditaCrear 53, 6, '" + stCliente + "'";
			Funcion.EjecutaSQL(cdsCliente, stAudita, true);
			#endregion Auditoria
			ConfiguracionInicial();
			Seguridad();
			CargarCombos();
			#region Busqueda de Registro
			barraDato1.HabilitaControles(false);
			if (this.IdCliente > 0) // Busca un cliente
				ClienteBuscado(IdCliente);
			else
				barraDato1.PosUltima();
			HabilitaTab();
			#endregion Busqueda de Registro
		}

		private void Seguridad()
		{
			if (!TipoCliente) 
				miAcceso =new Acceso(cdsSeteoF,  "0412");
			else
				miAcceso =new Acceso(cdsSeteoF,  "0411");

			if (!miAcceso.Eliminar) barraDato1.sbtBorrar.Enabled = false;
			if (!miAcceso.Editar) barraDato1.sbtEditar.Enabled = false;
			if (!miAcceso.Nuevo) barraDato1.sbtNuevo.Enabled = false;
			if (!miAcceso.Imprimir) barraDato1.sbtImprimir.Enabled = false;
			if (!miAcceso.Auditoria) btAuditoria.Visible = false;
		}
		private void ConfiguracionInicial()
		{
			#region Codigo Automatico
			if (TipoCliente) idTipoF = 33;
			bCodAuto = Funcion.bEjecutaSQL(cdsCliente, 
				"SELECT IsNull(AUTOMATICO, 0) FROM COMPRANUMERO WHERE IDTIPOFACTURA = " + idTipoF.ToString());
			if (bCodAuto) txtCodigo.ReadOnly = true;
			#endregion Codigo Automatico
			if (!TipoCliente) 
			{
				stCliente = "Proveedor ";
				txtAutRetencion.Width = 0;
				lblAutRetencion.Width = 0;
				stFiltro = "Proveedor = 1 Or Ambos = 1";
				stClientes = this.Text;
				chkPublica.Text = "RISE";
			}
			else
			{
				stFiltro = "Proveedor = 0 Or Ambos = 1";
			}
			barraDato1.stFiltroExt = " And (" + stFiltro + ")";
			txtIdCliente.Width = 0;
			drSeteoF = cdsSeteoF.TableViews["SeteoF"].Rows[0];
		}
		private void CargarCombos()
		{
			int iCliente = 1;
			if (!TipoCliente) iCliente = 0;
			cmbClientePublicidad.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Select idClientePublicidad, Medio From ClientePublicidad");
			cmbPrecio.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Select idTipoPrecio, Precio From ClientePrecio");
			cmbCiudad.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Select idCiudad, Ciudad From ClienteCiudad");
			cmbSector.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Select idSector, Sector From ClienteSector");
			cmbGrupo.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Select idGrupoCliente, Grupo From ClienteGrupo");
			cmbFormaPago.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Exec CompraFormaPagoCarga");
			cmbTipoRuc.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Exec ClienteTipoRucCarga " + iCliente);
			cmbBanco.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Select idBanco, Nombre From Banco");
			cmbPorcIvaRet.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Select idRetIva, Porcentaje, Detalle From RetencionIva");
			cmbPais.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Select idClientePais, Pais From ClientePais");
			cmbPais.Visible = false;
			bPais.Visible = false;
			if(Funcion.bEjecutaSQL(cdsPais,"Exec SeteoGExiste 'FClPais'"))
			{
				txtPais.Visible = false;
				cmbPais.Visible = true;
				bPais.Visible = true;
			}

			cmbCredTribut.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Select idCredTributario, CredTributario from CompraCredTribut");
			cmbDepende.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Select idCliente, Nombre From Cliente Where Entrega = 1 And " + stFiltro);
			cmbVendedor.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Select idPersonal, Nombre From Personal Where Vendedor = 1");
			cmbCobrador.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Select idPersonal, Nombre From Personal Where Cobrador = 1");
		}

		private void Cliente_Closing(object sender, System.ComponentModel.CancelEventArgs e)
		{
			string stAudita = "Exec AuditaCrear 53, 7, '" + stCliente + "'";
			Funcion.EjecutaSQL(cdsCliente, stAudita, true);
		}

		#endregion Pantalla Cliente
		#region DataSet
		private void cdsCliente_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsCliente.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);
		}

		private void cdsCliente_PositionChanged(object sender, C1.Data.PositionChangeEventArgs e)
		{
			drCliente = e.Row;
		}

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);
		}

		#endregion DataSet
		#region Barra Standard
		private void barraDato1_Borrar(object sender, System.EventArgs e)
		{
			if (drCliente == null)
			{
				MessageBox.Show("No existe un registro para borrar", "Información",
					MessageBoxButtons.OK, MessageBoxIcon.Information);
				return;
			}
			if (MessageBox.Show(string.Format("¿Desea Borrar este {0}?", stCliente), 
				"Confirmación", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) == DialogResult.No) 
				return;

			#region Verificacion para Borrar
			string stSelect = "Exec ClienteVerificaBorrar " + txtIdCliente.Text;
			string stMensaje = Funcion.sEscalarSQL(cdsSeteoF, stSelect);
			if (stMensaje.Length > 0) 
			{
				MessageBox.Show(stMensaje, "Información", MessageBoxButtons.OK, MessageBoxIcon.Information);
				return;
			}
			#endregion Fin de verificacion para borrar
			barraDato1.BorraRegistro();
		}

		private void barraDato1_Buscar(object sender, System.EventArgs e)
		{
			using (BuscaCliente miBusqueda = new BuscaCliente(this.TipoCliente))
			{
				if (miBusqueda.ShowDialog() == DialogResult.OK)
				{
					if (miBusqueda.chkOrden.Checked) barraDato1.DataOrden = miBusqueda.stOrden;
					else barraDato1.DataOrden = "";

					if (miBusqueda.idBusca.Text.Length > 0)
						ClienteBuscado(int.Parse(miBusqueda.idBusca.Text));
				}
			}
		}

		private void ClienteBuscado(int id)
		{
			this.barraDato1.IdRegistro = id;
			this.barraDato1.ProximoId(5);
		}

		private void barraDato1_Crear(object sender, System.EventArgs e)
		{
			barraDato1.CrearRegistro();
			#region Predefinidos
			if (TipoCliente) drCliente["idTipoRuc"] = 4;
			else drCliente["idTipoRuc"] = 1;
			drCliente["Pais"] = "Ecuador";
			drCliente["Fecha"] = DateTime.Today;
			drCliente["Proveedor"] = !TipoCliente;
//			drCliente["idSucursal"] = (int) drSeteo["NumEmpresa"];
			#endregion Predefinidos
			HabilitaTab();
			ultraTabPageControl1.Select();
			ultraTabPageControl1.Show();
			txtNombre.Select();
		}

		private void barraDato1_Copiar(object sender, System.EventArgs e)
		{
			if (DialogResult.No == MessageBox.Show("¿Desea Copiar " + stClientes + "?",
				"Confirmación", MessageBoxButtons.YesNo, MessageBoxIcon.Question,
				MessageBoxDefaultButton.Button2)) return;

			string stNombre = txtNombre.Text.Trim() + ".";
			// Copia Cliente Cuyo id = txtIdCliente y cambia el Nombre a stNombre
			string stSelect = "Exec CopiaRegistros 'Cliente', "
				+ "'idCliente = "	+ txtIdCliente.Text.ToString()
				+ "', 'Nombre,', '" + stNombre 
				+ "', 'idCliente,'";
			Funcion.EjecutaSQL(cdsCliente, stSelect);
			barraDato1.PosUltima();
			barraDato1_Editar(this, e);
			ultraTabPageControl1.Select();
			ultraTabPageControl1.Show();
			txtNombre.Select();
			MessageBox.Show("Copia de Clientes Terminada", "Información");
		}

		private void barraDato1_Grabar(object sender, System.EventArgs e)
		{
			if (!Verifica()) 
			{
				ultraTabPageControl1.Select();
				ultraTabPageControl1.Show();
				return;
			}
			bool bNuevo = barraDato1.bNuevo;
			barraDato1.GrabaRegistro();
			if (bNuevo)
			{
				int IdCliente = int.Parse(txtIdCliente.Value.ToString());
				foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in ultraGrid2.Rows.All)
				{
					if (dr.Cells["idCliente"].Value != System.DBNull.Value)
						dr.Cells["idCliente"].Value = IdCliente;
				}
			}
			cdsContacto.Update();
			if (bNuevo && bCodAuto)
			{
				string stExec = "Exec FacturaNumeroEstacion " + idTipoF;
				string stCodigo = Funcion.sEscalarSQL(cdsCliente, stExec, true);
				if (stCodigo.Trim().Length > 0)
				{
					stExec = "Update Cliente Set Codigo = '" + stCodigo 
						+ "' Where idCliente = " + txtIdCliente.Text;
					Funcion.EjecutaSQL(cdsCliente, stExec, true);
					txtCodigo.Value = stCodigo;
				}
			}
			HabilitaTab();
		}
		private bool Verifica()
		{
			bool bOk = true;
			BorrarErrores();

			#region Verifica Existencia de Tipo de Ruc
			if (cmbTipoRuc.Value == DBNull.Value || (int) cmbTipoRuc.Value == 0)
			{
				errorProvider.SetError(cmbTipoRuc, "Ingrese Tipo de RUC");
				bOk = false;
			}
			#endregion Verifica Existencia de Tipo de Ruc
			string stProv = " And (Proveedor = 1 Or Ambos = 1) ";
			if (TipoCliente) stProv = " And (Proveedor = 0 Or Ambos = 1) ";
			if (!TipoCliente)  // Caso de proveedor
			{
				#region Verificacion que no sea consumidor Final
				if ((int) cmbTipoRuc.Value == 7)
				{
					errorProvider.SetError(cmbTipoRuc, "No puede existir consumidor final en Proveedores");
					bOk = false;
				}
				#endregion Verificacion de Tipo de Ruc
				#region Verificacion de Autorizacion, Serie, Caducidad
				if (Funcion.bEjecutaSQL(cdsCliente, "Exec SeteoGExiste 'FCVA'"))
				{
					if (txtAutFactura.Value == null || (int) txtAutFactura.Value.ToString().Length != 10)
					{
						errorProvider.SetError(txtAutFactura, "La autorizacion debe tener 10 digitos");
						bOk = false;
					}
					if (txtSerie.Value == null || (int) txtSerie.Value.ToString().Length != 6)
					{
						errorProvider.SetError(txtSerie, "La Serie debe tener debe tener 6 digitos");
						bOk = false;
					}
					if (cmbFechaCaduca.Value == null || ((DateTime) cmbFechaCaduca.Value).Year < 2000)
					{
						errorProvider.SetError(cmbFechaCaduca, "Ingrese Fecha de Caducidad valida");
						bOk = false;
					}
				}
				#endregion Verificacion de Autorizacion, Serie, Caducidad
			}

			#region Verificacion de Campos obligatorios definidos en Configuracion
			string stCamposObligatorios = "";
			string stCodigoVerificador = "'FCLV'";
			if (!TipoCliente) stCodigoVerificador = "'FPRV'";
			SqlDataReader dr = Funcion.rEscalarSQL(cdsCliente, "Exec VerificaCliente " + stCodigoVerificador, true);
			while (dr.Read())
			{
				string stCampo = dr.GetString(0);
				string stDisplay = "";
				int iPos_ = stCampo.IndexOf("_");
				if (iPos_ > 0)
				{
					stDisplay = stCampo.Substring(iPos_ + 1);
					stCampo = stCampo.Substring(0, iPos_);
				}
				try
				{
					bool bCampoValido = true;
					if (stCampo.StartsWith("id"))
					{
						if (drCliente[stCampo] == DBNull.Value)	bCampoValido = false;
						else if (drCliente[stCampo].ToString().Length == 0) bCampoValido = false;
						else if (int.Parse(drCliente[stCampo].ToString()) == 0) bCampoValido = false;
					}
					else
						if (drCliente[stCampo] == DBNull.Value || drCliente[stCampo].ToString().Length == 0)
						bCampoValido = false;
				
					if (!bCampoValido)
					{
						if (stDisplay.Length > 0) stCampo = stDisplay;
						stCamposObligatorios += "\n" + stCampo;
					}
				}
				catch(System.Exception ex)
				{
					MessageBox.Show(ex.Message, "Error en campo: " + stCampo);
				}
			}
			dr.Close();
			if (stCamposObligatorios.Length > 0)
			{
				MessageBox.Show("Ingrese estos campos obligatorios:" + stCamposObligatorios, "Información",
					MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
				return false;
			}
			#endregion Verificacion de Campos obligatorios definidos en Configuracion

			#region Verifica Nombre
			if (txtNombre.Text.ToString().Trim().Length == 0)
			{
				errorProvider.SetError(txtNombre, "Ingrese el nombre");
				bOk = false;
			}
			else
			{
				#region Verificacion de Nombre Repetido
				if (drSeteoF["ClienteRepetido"] == DBNull.Value || ! (bool) drSeteoF["ClienteRepetido"])
				{
					string stSelect = "SELECT COUNT(*) FROM Cliente WHERE Nombre = '" +
						txtNombre.Text.ToString().Trim() + "'" + stProv + " And idCliente <> " + txtIdCliente.Text;

					int iCuenta = Funcion.iEscalarSQL(cdsSeteoF, stSelect);
					if (iCuenta > 0) 
					{
						errorProvider.SetError(txtNombre,"Nombre ya existe");
						bOk = false;
					}
				}
				#endregion Verificacion de Nombre Repetido
			}
			#endregion Verifica Nombre

			#region Verificacion de RUC
			if ((bool) drSeteoF["VerificaRUC"] && (int) cmbTipoRuc.Value !=7 && !txtRuc.Text.ToUpper().StartsWith("SN"))
			{
				if (txtRuc.Text.ToString().Trim().Length == 0)
				{
					errorProvider.SetError(txtRuc, "Ingrese el RUC");
					bOk = false;
				}
				else
				{
					#region Ruc Repetido
					string stSelect = "SELECT COUNT(*) FROM Cliente WHERE Ruc = '" +
						txtRuc.Text.ToString().Trim() + "'" + stProv + " And idCliente <> " + txtIdCliente.Text;

					int iCuenta = Funcion.iEscalarSQL(cdsSeteoF, stSelect);
					if (iCuenta > 0) 
					{
						errorProvider.SetError(txtRuc,"RUC ya existe");
						bOk = false;
					}					
					#endregion Fin de verificacion de RUC repetido
					#region 13 digitos en ruc
					if ((int)cmbTipoRuc.Value == 1 || (int)cmbTipoRuc.Value == 4)
					{
						if (txtRuc.Text.Trim().Length != 13)
						{
							errorProvider.SetError(txtRuc, "RUC debe tener 13 digitos");
							bOk = false;
						}
					}
					#endregion 13 digitos en ruc
					#region 10 digitos en cedula
					if ((int)cmbTipoRuc.Value == 2 || (int)cmbTipoRuc.Value == 5)
					{
						if (txtRuc.Text.Trim().Length != 10)
						{
							errorProvider.SetError(txtRuc, "Cedula debe tener 10 digitos");
							bOk = false;
						}
					}
					#endregion 10 digitos en cedula
					#region Verificacion de cedula
					// Compra con Ruc o Cedula, Venta con Ruc o Cedula
					if ((int)cmbTipoRuc.Value == 1 || (int) cmbTipoRuc.Value == 2 || (int) cmbTipoRuc.Value == 4 || (int) cmbTipoRuc.Value == 5)
					{
						string stError = Funcion.VerificaCedula(txtRuc.Text, false);
						if (stError.Length > 0)
						{
							errorProvider.SetError(txtRuc, stError);
							bOk = false;
						}
					}
					#endregion Verificacion de cedula
				}
			} // Fin de drseteoF[""]
			#endregion Verificacion de RUC
			return bOk;
		}

		private void barraDato1_Editar(object sender, System.EventArgs e)
		{
			if (drCliente == null)
			{
				MessageBox.Show("No existe un registro para editar", "Información",
					MessageBoxButtons.OK, MessageBoxIcon.Information);
				return;
			}
			barraDato1.EditarRegistro();
			HabilitaTab();
		}

		private void barraDato1_Deshacer(object sender, System.EventArgs e)
		{
			BorrarErrores();
			barraDato1.DeshacerRegistro();
			HabilitaTab();
		}

		private void BorrarErrores()
		{
			errorProvider.SetError(txtNombre, "");
			errorProvider.SetError(txtRuc, "");
			errorProvider.SetError(cmbTipoRuc, "");
			errorProvider.SetError(txtAutFactura, "");
			errorProvider.SetError(txtSerie, "");
		}

		private void HabilitaTab()
		{
			txtAutFactura.Enabled = true;
			foreach(Infragistics.Win.UltraWinTabControl.UltraTab ctr in tbBasico.Tabs)
				foreach(Control cn in ctr.TabPage.Controls)
					barraDato1.HabilitaControl(cn, barraDato1.bEditar);
		}

		private void barraDato1_Imprimir(object sender, System.EventArgs e)
		{
			RepCliente miReporte = new RepCliente();
			miReporte.MdiParent = MdiParent;
			if (!TipoCliente) miReporte.optTipo.Value = 1;
			miReporte.Show();
		}

		private void barraDato1_Impresora(object sender, System.EventArgs e)
		{
		
		}

		private void barraDato1_Refresca(object sender, System.EventArgs e)
		{
			this.Text = stCliente + txtNombre.Text;
			if (txtIdCliente.Value == null) return;
			C1.Data.FilterCondition fcContactos = new 
				C1.Data.FilterCondition(cdsContacto, "ClienteContacto", "idCliente = " 
				+ txtIdCliente.Value.ToString());
			C1.Data.FilterConditions fcTotal = new C1.Data.FilterConditions();

			fcTotal.Add(fcContactos);
			cdsContacto.Fill(fcTotal, false);
		}

		private void barraDato1_Restaurar(object sender, System.EventArgs e)
		{
		
		}

		#endregion Barra Standard
		#region Barra Extendida
		private void btAuditoria_Click(object sender, System.EventArgs e)
		{
			Cursor = Cursors.WaitCursor;
			int iProveedor = 1;
			if (TipoCliente) iProveedor = 0;
			Funcion.AuditoriaImprime(25 + iProveedor, int.Parse(txtIdCliente.Value.ToString()), true);
			Cursor = Cursors.Default;
		}

		private void btAuditoria_MouseUp(object sender, System.Windows.Forms.MouseEventArgs e)
		{
			if (e.Button != MouseButtons.Right) return;
			Cursor = Cursors.WaitCursor;
			int iProveedor = 1;
			if (TipoCliente) iProveedor = 0;
			Funcion.AuditoriaImprime(25 + iProveedor, int.Parse(txtIdCliente.Value.ToString()), false);
			Cursor = Cursors.Default;		
		}

		private void btCRMClientes_Click(object sender, System.EventArgs e)
		{
			CRMClientes miCrm = new CRMClientes();
			miCrm.MdiParent = this.MdiParent;
			miCrm.Show();
		}

		#endregion Barra Extendida
		#region Pestaña
		#region Basico
		private void btPrecio_Click(object sender, System.EventArgs e)
		{
			bool bPrecioCliente = Funcion.bEscalarSQL(cdsCliente, "Exec SeteoGExiste 'FCLPA'", true);
			if (bPrecioCliente)
			{
				using (ClienteListaPrecio miPrecio = new ClienteListaPrecio())
				{
					miPrecio.ShowDialog();
					return;
				}
			}

			using (ClientePrecio miPrecio = new ClientePrecio())
			{
				miPrecio.ShowDialog();
			}
			cmbPrecio.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Select idTipoPrecio, Precio From ClientePrecio");
		}
		private void bCiudad_Click(object sender, System.EventArgs e)
		{
			using (ClienteCiudad miCiudad = new ClienteCiudad())
			{
				miCiudad.ShowDialog();
			}
			cmbCiudad.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Select idCiudad, Ciudad From ClienteCiudad");
		}

		private void btSector_Click(object sender, System.EventArgs e)
		{
			using(ClienteSector miSector = new ClienteSector())
			{
				miSector.ShowDialog();
			}
			cmbSector.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Select idSector, Sector From ClienteSector");
		}

		private void btGrupoCliente_Click(object sender, System.EventArgs e)
		{
			using(ClienteGrupo miGrupoC = new ClienteGrupo())
			{
				miGrupoC.ShowDialog();
			}
			cmbGrupo.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Select idGrupoCliente, Grupo From ClienteGrupo");
		}

		private void cmbGrupo_MouseDown(object sender, System.Windows.Forms.MouseEventArgs e)
		{
			if (e.Button != MouseButtons.Right) return;
			cmbGrupo.Value = 0;
		}

		private void btVerifica_Click(object sender, System.EventArgs e)
		{
			string stMensaje = Funcion.VerificaCedula(this.txtRuc.Text.Trim(), true);
			MessageBox.Show(stMensaje, "Informacion",
				MessageBoxButtons.OK, MessageBoxIcon.Information);
		}

		private void btPublicidad_Click(object sender, System.EventArgs e)
		{
			using(ClientePublicidad miPublicidad = new ClientePublicidad())
			{
				miPublicidad.ShowDialog();
			}
			cmbClientePublicidad.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Select idClientePublicidad, Medio From ClientePublicidad");
		}

		#endregion Basico
		#region Complemento
		private void btDepende_Click(object sender, System.EventArgs e)
		{
			int IdPadre = 0;
			if (cmbDepende.Value != DBNull.Value && (int) cmbDepende.Value != 0)
				IdPadre = (int) cmbDepende.Value;
			else
				IdPadre = int.Parse(txtIdCliente.Text);
			if (IdPadre == 0)
			{
				MessageBox.Show("No existen dependientes", "Información");
				return;
			}

			using (ClienteDepende miDepende = new ClienteDepende(IdPadre))
			{
				if (DialogResult.OK == miDepende.ShowDialog())
				{
					int IdB = int.Parse(miDepende.IdBuscado.Text);
					if (IdB > 0) ClienteBuscado(IdB);
				}
			}
		}

		private void btDepende_MouseDown(object sender, System.Windows.Forms.MouseEventArgs e)
		{
			if (e.Button != MouseButtons.Right) return;
			if (cmbDepende.Value == System.DBNull.Value || (int) cmbDepende.Value == 0)
			{
				MessageBox.Show("No tiene dependencias", "Información");
				return;
			}
			int IdPadre = (int) cmbDepende.Value;
			if (DialogResult.No == MessageBox.Show("Desea ir a la ficha de "
				+ cmbDepende.Text.Trim() + "?", "Confirmación", MessageBoxButtons.YesNo,
				MessageBoxIcon.Question)) return;
			ClienteBuscado(IdPadre);
		}

		private void btVendedor_Click(object sender, System.EventArgs e)
		{
			using(RolPersonal miPersonal = new RolPersonal())
			{
				miPersonal.ShowDialog();
			}
			cmbVendedor.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Select idPersonal, Nombre From Personal Where Vendedor = 1");
		}

		#endregion Complemento
		#region Referencias
		private void btBanco_Click(object sender, System.EventArgs e)
		{
			using (Bancos miBanco = new Bancos())
			{
				miBanco.ShowDialog();
			}
			cmbBanco.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Select idBanco, Nombre From Banco");
		}

		#endregion Referencias
		#region Analisis Cartera
		private void btListadoPendiente_Click(object sender, System.EventArgs e)
		{
			LlenarGrilla(1);
		}

		private void btVencido_Click(object sender, System.EventArgs e)
		{
			LlenarGrilla(2);		
		}

		private void btTotales_Click(object sender, System.EventArgs e)
		{
			LlenarGrilla(3);		
		}		
		private void btFactura_Click(object sender, System.EventArgs e)
		{
			if (ultraGrid1.ActiveRow == null ||	ultraGrid1.ActiveRow.Cells["idCompra"].Value == DBNull.Value)
			{
				MessageBox.Show("Seleccione la factura que desea abrir", "Información",
					MessageBoxButtons.OK, MessageBoxIcon.Information);
				return;
			}
			Cursor = Cursors.WaitCursor;
			int IdCompra = (int) ultraGrid1.ActiveRow.Cells["idCompra"].Value;
			miCompra = new Compra(-IdCompra);
			miCompra.MdiParent = MdiParent;	
			miCompra.Show();
		}

		private void btFactura_MouseUp(object sender, System.Windows.Forms.MouseEventArgs e)
		{
			if (e.Button != MouseButtons.Right) return;
			if (ultraGrid1.ActiveRow == null ||	ultraGrid1.ActiveRow.Cells["idCompra"].Value == DBNull.Value)
			{
				MessageBox.Show("Seleccione la factura que desea abrir", "Información",
					MessageBoxButtons.OK, MessageBoxIcon.Information);
				return;
			}
			Cursor = Cursors.WaitCursor;
			int IdCompra = (int) ultraGrid1.ActiveRow.Cells["idCompra"].Value;
			miCompra.BuscarCompra(IdCompra);
			miCompra.Show();
		}

		private void LlenarGrilla(int Tipo)
		{
			string stExec = string.Format("Exec ClienteAnalisisCartera {0}, {1}",
				txtIdCliente.Text, Tipo);
			ultraGrid1.DataSource = Funcion.dvProcedimiento(cdsSeteoF, stExec);
			ultraGrid1.Text = "Listado de Pendientes";
		}

		#endregion Analisis Cartera
		#region Listado Items
		private void btListadoItems_Click(object sender, System.EventArgs e)
		{
			LlenarItems(1);		
		}
		private void LlenarItems(int Tipo)
		{
			string stExec = string.Format("Exec ClienteAnalisisItem {0}, {1}",
				txtIdCliente.Text, Tipo);
			grdItems.DataSource = Funcion.dvProcedimiento(cdsSeteoF, stExec);
			grdItems.Text = "Listado de Items";
		}

		#endregion Listado Items

		private void chkProspecto_CheckedChanged(object sender, System.EventArgs e)
		{
		
		}
		#endregion Pestaña

		private void ultraGrid2_AfterRowInsert(object sender, Infragistics.Win.UltraWinGrid.RowEventArgs e)
		{
			e.Row.Cells["idCliente"].Value = int.Parse(txtIdCliente.Value.ToString());
		}

		private void btExcel_Click(object sender, System.EventArgs e)
		{
			string stDir = MenuLatinium.stDirInicio + "\\"; 
			ultraGridExcelExporter1.Export(ultraGrid1, "ClienteRep.xls");
			if (DialogResult.Yes == MessageBox.Show(
				"¿Desea Ver el Archivo Generado?", 
				"Información", MessageBoxButtons.YesNo, MessageBoxIcon.Question))
				System.Diagnostics.Process.Start(stDir + "ClienteRep.xls");
		}

		private void ultraButton1_Click(object sender, System.EventArgs e)
		{
			string stDir = MenuLatinium.stDirInicio + "\\"; 
			ultraGridExcelExporter1.Export(ultraGrid2, "ClienteItem.xls");
			if (DialogResult.Yes == MessageBox.Show(
				"¿Desea Ver el Archivo Generado?", 
				"Información", MessageBoxButtons.YesNo, MessageBoxIcon.Question))
				System.Diagnostics.Process.Start(stDir + "ClienteItem.xls");		
		}

		private void bPais_Click(object sender, System.EventArgs e)
		{

		}
	}
}
