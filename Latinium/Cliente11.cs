using System;
using System.Data.SqlClient;
using C1.Data;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using Infragistics.Win.UltraWinGrid;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de Cliente.
	/// </summary>
	public class Cliente : DevExpress.XtraEditors.XtraForm
	{
		private System.Windows.Forms.Label label1;
		private System.Windows.Forms.Label label3;
		private System.Windows.Forms.Label label4;
		private System.Windows.Forms.Label label5;
		private System.Windows.Forms.Label label7;
		private System.Windows.Forms.Label label8;
		private System.Windows.Forms.Label label11;
		private System.Windows.Forms.Label label13;
		private System.Windows.Forms.Label label17;
		private C1.Data.C1DataSet cdsCliente;
		private System.Windows.Forms.ErrorProvider errorProvider;
		private Infragistics.Win.Misc.UltraButton bCiudad;
		private Infragistics.Win.Misc.UltraButton btGrupoCliente;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbCiudad;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbGrupo;
		public Infragistics.Win.UltraWinEditors.UltraTextEditor txtNombre;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtDireccion;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtTelefono;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtEMail;
		public Infragistics.Win.UltraWinEditors.UltraTextEditor txtIdCliente;
		private BarraDatoSQL.BarraDatoSQL barraDato1;
		private C1.Data.C1DataSet cdsSeteoF;
		private Infragistics.Win.UltraWinTabControl.UltraTabControl tbBasico;
		private Infragistics.Win.UltraWinTabControl.UltraTabSharedControlsPage ultraTabSharedControlsPage1;
		private Infragistics.Win.UltraWinTabControl.UltraTabPageControl ultraTabPageControl1;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtCelular;
		private System.Windows.Forms.Label label9;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor spnDescCredito;
		private System.Windows.Forms.Label label24;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor spnDescuento;
		private System.Windows.Forms.Label label23;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtValorCredito;
		private System.Windows.Forms.Label label14;
		private System.Windows.Forms.Label label12;
		private System.Windows.Forms.Label label21;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbVendedor;
		private Infragistics.Win.Misc.UltraButton btVendedor;
		private Infragistics.Win.UltraWinTabControl.UltraTabPageControl tpComplemento;
		private System.Windows.Forms.Label label26;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbPorcIvaRet;
		private System.Windows.Forms.Label label29;
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
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkEntrega;
		private Infragistics.Win.Misc.UltraButton btAuditoria;
		private System.Windows.Forms.Label label48;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor ultraCheckEditor3;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor ultraNumericEditor2;
		private System.Windows.Forms.Label label49;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor ultraNumericEditor3;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkPublica;
		private System.Windows.Forms.Label label53;
		private Infragistics.Win.UltraWinTabControl.UltraTabPageControl ultraTabPageControl3;
		private Infragistics.Win.UltraWinGrid.UltraGrid ultraGrid1;
		private Infragistics.Win.Misc.UltraButton btListadoPendiente;
		private Infragistics.Win.Misc.UltraButton btVencido;
		private Infragistics.Win.Misc.UltraButton btFactura;
		private Infragistics.Win.Misc.UltraButton btTotales;
		private Infragistics.Win.UltraWinTabControl.UltraTabPageControl ultraTabPageControl4;
		private Infragistics.Win.UltraWinGrid.UltraGrid grdItems;
		private Infragistics.Win.Misc.UltraButton btListadoItems;
		private System.Windows.Forms.ImageList imageList;
		private System.Windows.Forms.ToolTip toolTip1;
		private Infragistics.Win.UltraWinTabControl.UltraTabPageControl ultraTabPageControl5;
		private Infragistics.Win.UltraWinGrid.UltraGrid ultraGrid2;
		private C1.Data.C1DataSet cdsContacto;
		private System.Windows.Forms.Label label28;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtPais;
		private Infragistics.Win.Misc.UltraButton bPais;
		private C1.Data.C1DataSet cdsPais;
		private Infragistics.Win.Misc.UltraButton btExcel;
		private Infragistics.Win.UltraWinGrid.ExcelExport.UltraGridExcelExporter ultraGridExcelExporter1;
		private Infragistics.Win.Misc.UltraButton ultraButton1;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbPais;
		private Infragistics.Win.Misc.UltraButton btCedRep;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtCedRep;
		private System.Windows.Forms.Label label52;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbTipoCredito;
		private Infragistics.Win.Misc.UltraButton ultraButton2;
		private Infragistics.Win.Misc.UltraButton btProveedorCliente;
		private System.Windows.Forms.Label label54;
		private System.Windows.Forms.Label label55;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbTipoRuc;
		private System.Windows.Forms.Label label22;
		public Infragistics.Win.UltraWinEditors.UltraTextEditor txtRuc;
		private Infragistics.Win.Misc.UltraButton btVerifica;
		private Infragistics.Win.UltraWinTabControl.UltraTabPageControl ultraTabPageControl6;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo ultraCalendarCombo1;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor ultraCheckEditor1;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor ultraTextEditor5;
		private System.Windows.Forms.Label label10;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo cmbFecha;
		private System.Windows.Forms.Label label6;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkProspecto;
		private System.Windows.Forms.GroupBox groupBox3;
		private System.Windows.Forms.GroupBox groupBox4;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor ultraTextEditor21;
		private System.Windows.Forms.Label label51;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor ultraTextEditor4;
		private System.Windows.Forms.Label label50;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtSerie;
		private System.Windows.Forms.Label label32;
		private System.Windows.Forms.Label label31;
		private Infragistics.Win.UltraWinEditors.UltraDateTimeEditor cmbFechaCaduca;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtAutFactura;
		private System.Windows.Forms.Label label15;
		private System.Windows.Forms.Label lblCuentaFondo;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbCuentaFondo;
		private System.Windows.Forms.Label lblCuentaAnticipo;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbCuentaAnticipo;
		private Infragistics.Win.Misc.UltraButton btDepende;
		private System.Windows.Forms.Label label25;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbDepende;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkSujetoRetencion;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkChofer;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkEspecial;
		private System.Windows.Forms.Label label33;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbFormaPago;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtAutRetencion;
		private System.Windows.Forms.Label lblAutRetencion;
		private System.Windows.Forms.Label label20;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbCobrador;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtAutImprenta;
		private System.Windows.Forms.Label label16;
		private Infragistics.Win.Misc.UltraButton btPublicidad;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbClientePublicidad;
		private System.Windows.Forms.Label label30;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbPrecio;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbSector;
		private Infragistics.Win.Misc.UltraButton btSector;
		private Infragistics.Win.Misc.UltraButton btPrecio;
		private System.Windows.Forms.Label label19;
		private System.Windows.Forms.Label label2;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtCodigo;
		private System.Windows.Forms.Label label18;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtFax;
		private System.Windows.Forms.RadioButton rdbContado;
		private System.Windows.Forms.RadioButton rdbCredito;
		private DevExpress.XtraEditors.RadioGroup radioGroup1;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkAplicaConsignacion;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkHaceRetencion;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbDias;
		private System.Windows.Forms.Label lblDias;
		private System.Windows.Forms.Label lblHoraHasta;
		private System.Windows.Forms.Label lblHoraDesde;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtDiasCredito3;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtDiasCredito2;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtDiasCredito1;
		private System.Windows.Forms.DateTimePicker dtHoraDesde;
		private System.Windows.Forms.DateTimePicker dtHoraHasta;
		private System.Windows.Forms.Label lblNotas;
		private System.Windows.Forms.Label lblCuentaSwift;
		private System.Windows.Forms.Label lblBanco;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtCuentaSwift;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtBeneficiario;
		private System.Windows.Forms.Label lblBeneficiario;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbBanco;
		private Infragistics.Win.Misc.UltraButton btBanco;
		private Infragistics.Win.UltraWinTabControl.UltraTabPageControl ultraTabPageControl7;
		private Infragistics.Win.UltraWinGrid.UltraGrid ultraGrid3;
		private C1.Data.C1DataSet cdsDirecciones;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtDiasConsignacion;
		private System.Windows.Forms.Label label27;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtRepresentante;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtNombreComercial;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtNotas;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtObservacion;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbArticuloSRI;
		private System.Windows.Forms.Label lblArticuloSRI;
		private C1.Data.C1DataSet cdsArticuloTabla;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource1;
		private System.Windows.Forms.Label lblTipoProveedor;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbTipoProveedor;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkCodigoBarras;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbAgenteDeRetencion;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource2;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbInstitucion;
		private System.Windows.Forms.Label label57;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource3;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbSujetoRetencion;
		private System.Windows.Forms.Label lblASRetencion;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource4;
		public Infragistics.Win.UltraWinEditors.UltraTextEditor ultraTextEditor1;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkCourier;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtDiasCredito4;
		private System.Windows.Forms.Label label56;
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
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idSujetoRetencion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idSujetoRetencion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idAgenteRetencion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn3 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idAgenteRetencion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn4 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand3 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idTipoProveedor");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Tipo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn5 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idTipoProveedor");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn6 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Codigo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn7 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Tipo");
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand4 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idTipoRuc");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("TipoRuc");
			Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance11 = new Infragistics.Win.Appearance();
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(Cliente));
			Infragistics.Win.Appearance appearance12 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand5 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idClientePais");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn11 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Pais");
			Infragistics.Win.Appearance appearance13 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance14 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance15 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance16 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance17 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance18 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance19 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance20 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand6 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn12 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCiudad");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn13 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Ciudad");
			Infragistics.Win.Appearance appearance21 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand7 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn14 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idGrupoCliente");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn15 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Grupo");
			Infragistics.Win.Appearance appearance22 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance23 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance24 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance25 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance26 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance27 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand8 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn16 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idInstitucion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn17 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Descripcion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn8 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idInstitucion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn9 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Descripcion");
			Infragistics.Win.Appearance appearance28 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand9 = new Infragistics.Win.UltraWinGrid.UltraGridBand("ArticuloSRI", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn18 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idArticuloSRI");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn19 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Detalle");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn20 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Porcentaje");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn21 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCuentaI");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn22 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCuentaE");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn23 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn24 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Activo");
			Infragistics.Win.Appearance appearance29 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand10 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn25 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idBanco");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn26 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.Appearance appearance30 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance31 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance32 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance33 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance34 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand11 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn27 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idClientePublicidad");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn28 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Medio");
			Infragistics.Win.Appearance appearance35 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand12 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn29 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idTipoPrecio");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn30 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Precio");
			Infragistics.Win.Appearance appearance36 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand13 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn31 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idSector");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn32 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Sector");
			Infragistics.Win.Appearance appearance37 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance38 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance39 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance40 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand14 = new Infragistics.Win.UltraWinGrid.UltraGridBand("TIPOCRED", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn33 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idTipoCredito");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn34 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn35 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn36 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Dias");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn37 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Giros");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn38 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("PctIneteres");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn39 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Tipo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn40 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("AfectaA");
			Infragistics.Win.Appearance appearance41 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance42 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand15 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn41 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCredTributario");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn42 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CredTributario");
			Infragistics.Win.Appearance appearance43 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand16 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn43 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idRetIva");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn44 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Porcentaje");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn45 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Detalle");
			Infragistics.Win.Appearance appearance44 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand17 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn46 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPersonal");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn47 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.Appearance appearance45 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance46 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance47 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance48 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance49 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance50 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand18 = new Infragistics.Win.UltraWinGrid.UltraGridBand("ClienteContacto", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn48 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idClienteContacto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn49 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCliente");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn50 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Contacto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn51 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Telefono");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn52 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Detalle");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn53 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Email");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn54 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Regalo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn55 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaCumple");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn56 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Celular");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn57 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cargo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn58 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("EnvioMail");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn59 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("EnviaRegalo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn60 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("EFE");
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
			Infragistics.Win.Appearance appearance79 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance80 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance81 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance82 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance83 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance84 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance85 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand19 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn61 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idFormaPago");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn62 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FormaPago");
			Infragistics.Win.Appearance appearance86 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance87 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand20 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn63 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPersonal");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn64 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.Appearance appearance88 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance89 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand21 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn65 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPersonal");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn66 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.Appearance appearance90 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand22 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn67 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPersonal");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn68 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.Appearance appearance91 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand23 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn69 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCliente");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn70 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.Appearance appearance92 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance93 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance94 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance95 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance96 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance97 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance98 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance99 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance100 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton1 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.Appearance appearance101 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton2 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.Appearance appearance102 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand24 = new Infragistics.Win.UltraWinGrid.UltraGridBand("ClientesDirecciones", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn71 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idDreccion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn72 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCliente");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn73 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Direccion");
			Infragistics.Win.Appearance appearance103 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance104 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance105 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance106 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance107 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance108 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinTabControl.UltraTab ultraTab1 = new Infragistics.Win.UltraWinTabControl.UltraTab();
			Infragistics.Win.Appearance appearance109 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance110 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance111 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance112 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinTabControl.UltraTab ultraTab2 = new Infragistics.Win.UltraWinTabControl.UltraTab();
			Infragistics.Win.UltraWinTabControl.UltraTab ultraTab3 = new Infragistics.Win.UltraWinTabControl.UltraTab();
			Infragistics.Win.UltraWinTabControl.UltraTab ultraTab4 = new Infragistics.Win.UltraWinTabControl.UltraTab();
			Infragistics.Win.UltraWinTabControl.UltraTab ultraTab5 = new Infragistics.Win.UltraWinTabControl.UltraTab();
			Infragistics.Win.UltraWinTabControl.UltraTab ultraTab6 = new Infragistics.Win.UltraWinTabControl.UltraTab();
			Infragistics.Win.UltraWinTabControl.UltraTab ultraTab7 = new Infragistics.Win.UltraWinTabControl.UltraTab();
			Infragistics.Win.UltraWinTabControl.UltraTab ultraTab8 = new Infragistics.Win.UltraWinTabControl.UltraTab();
			Infragistics.Win.Appearance appearance113 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance114 = new Infragistics.Win.Appearance();
			this.ultraTabPageControl1 = new Infragistics.Win.UltraWinTabControl.UltraTabPageControl();
			this.chkCourier = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.cdsCliente = new C1.Data.C1DataSet();
			this.ultraTextEditor1 = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.cmbSujetoRetencion = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.ultraDataSource4 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.cmbAgenteDeRetencion = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.ultraDataSource2 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.lblASRetencion = new System.Windows.Forms.Label();
			this.chkCodigoBarras = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.cmbTipoProveedor = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.ultraDataSource1 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.lblTipoProveedor = new System.Windows.Forms.Label();
			this.label27 = new System.Windows.Forms.Label();
			this.txtDiasConsignacion = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.chkAplicaConsignacion = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.cmbTipoRuc = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.label22 = new System.Windows.Forms.Label();
			this.txtRuc = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.btVerifica = new Infragistics.Win.Misc.UltraButton();
			this.cmbPais = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.txtPais = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.bPais = new Infragistics.Win.Misc.UltraButton();
			this.txtNombreComercial = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label28 = new System.Windows.Forms.Label();
			this.lblNotas = new System.Windows.Forms.Label();
			this.txtNotas = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label26 = new System.Windows.Forms.Label();
			this.txtObservacion = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtCelular = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label9 = new System.Windows.Forms.Label();
			this.label5 = new System.Windows.Forms.Label();
			this.txtEMail = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtRepresentante = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label7 = new System.Windows.Forms.Label();
			this.label8 = new System.Windows.Forms.Label();
			this.btGrupoCliente = new Infragistics.Win.Misc.UltraButton();
			this.cmbCiudad = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.label1 = new System.Windows.Forms.Label();
			this.cmbGrupo = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.label11 = new System.Windows.Forms.Label();
			this.bCiudad = new Infragistics.Win.Misc.UltraButton();
			this.label13 = new System.Windows.Forms.Label();
			this.label4 = new System.Windows.Forms.Label();
			this.label3 = new System.Windows.Forms.Label();
			this.label17 = new System.Windows.Forms.Label();
			this.txtNombre = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtDireccion = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtTelefono = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtFax = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.tpComplemento = new Infragistics.Win.UltraWinTabControl.UltraTabPageControl();
			this.txtDiasCredito4 = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label56 = new System.Windows.Forms.Label();
			this.cmbInstitucion = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.ultraDataSource3 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.label57 = new System.Windows.Forms.Label();
			this.cmbArticuloSRI = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.cdsArticuloTabla = new C1.Data.C1DataSet();
			this.lblArticuloSRI = new System.Windows.Forms.Label();
			this.btBanco = new Infragistics.Win.Misc.UltraButton();
			this.cmbBanco = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.txtBeneficiario = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.lblBeneficiario = new System.Windows.Forms.Label();
			this.txtCuentaSwift = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.lblCuentaSwift = new System.Windows.Forms.Label();
			this.lblBanco = new System.Windows.Forms.Label();
			this.dtHoraHasta = new System.Windows.Forms.DateTimePicker();
			this.dtHoraDesde = new System.Windows.Forms.DateTimePicker();
			this.cmbDias = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.lblHoraHasta = new System.Windows.Forms.Label();
			this.lblHoraDesde = new System.Windows.Forms.Label();
			this.lblDias = new System.Windows.Forms.Label();
			this.chkHaceRetencion = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.rdbCredito = new System.Windows.Forms.RadioButton();
			this.rdbContado = new System.Windows.Forms.RadioButton();
			this.btPublicidad = new Infragistics.Win.Misc.UltraButton();
			this.cmbClientePublicidad = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.label30 = new System.Windows.Forms.Label();
			this.cmbPrecio = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.cmbSector = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.btSector = new Infragistics.Win.Misc.UltraButton();
			this.btPrecio = new Infragistics.Win.Misc.UltraButton();
			this.label19 = new System.Windows.Forms.Label();
			this.label2 = new System.Windows.Forms.Label();
			this.txtCodigo = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label18 = new System.Windows.Forms.Label();
			this.txtDiasCredito3 = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label55 = new System.Windows.Forms.Label();
			this.txtDiasCredito2 = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label54 = new System.Windows.Forms.Label();
			this.ultraButton2 = new Infragistics.Win.Misc.UltraButton();
			this.cmbTipoCredito = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.cdsContacto = new C1.Data.C1DataSet();
			this.label52 = new System.Windows.Forms.Label();
			this.chkContribuyenteEsp = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.label34 = new System.Windows.Forms.Label();
			this.cmbCredTribut = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.label29 = new System.Windows.Forms.Label();
			this.cmbPorcIvaRet = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.label21 = new System.Windows.Forms.Label();
			this.cmbVendedor = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.btVendedor = new Infragistics.Win.Misc.UltraButton();
			this.spnDescCredito = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label24 = new System.Windows.Forms.Label();
			this.spnDescuento = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label23 = new System.Windows.Forms.Label();
			this.txtValorCredito = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtDiasCredito1 = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label14 = new System.Windows.Forms.Label();
			this.label12 = new System.Windows.Forms.Label();
			this.chkPublica = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.ultraTabPageControl5 = new Infragistics.Win.UltraWinTabControl.UltraTabPageControl();
			this.ultraGrid2 = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraTabPageControl2 = new Infragistics.Win.UltraWinTabControl.UltraTabPageControl();
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
			this.btCedRep = new Infragistics.Win.Misc.UltraButton();
			this.label37 = new System.Windows.Forms.Label();
			this.ultraTextEditor9 = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.ultraTextEditor12 = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label39 = new System.Windows.Forms.Label();
			this.label43 = new System.Windows.Forms.Label();
			this.label38 = new System.Windows.Forms.Label();
			this.ultraTextEditor16 = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtCedRep = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
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
			this.ultraTabPageControl6 = new Infragistics.Win.UltraWinTabControl.UltraTabPageControl();
			this.radioGroup1 = new DevExpress.XtraEditors.RadioGroup();
			this.groupBox4 = new System.Windows.Forms.GroupBox();
			this.chkSujetoRetencion = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.chkChofer = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.chkEspecial = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.label33 = new System.Windows.Forms.Label();
			this.cmbFormaPago = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.txtAutRetencion = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.lblAutRetencion = new System.Windows.Forms.Label();
			this.label20 = new System.Windows.Forms.Label();
			this.cmbCobrador = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.txtAutImprenta = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label16 = new System.Windows.Forms.Label();
			this.lblCuentaFondo = new System.Windows.Forms.Label();
			this.cmbCuentaFondo = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.lblCuentaAnticipo = new System.Windows.Forms.Label();
			this.cmbCuentaAnticipo = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.btDepende = new Infragistics.Win.Misc.UltraButton();
			this.label25 = new System.Windows.Forms.Label();
			this.cmbDepende = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.ultraTextEditor21 = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label51 = new System.Windows.Forms.Label();
			this.ultraTextEditor4 = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label50 = new System.Windows.Forms.Label();
			this.txtSerie = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label32 = new System.Windows.Forms.Label();
			this.label31 = new System.Windows.Forms.Label();
			this.cmbFechaCaduca = new Infragistics.Win.UltraWinEditors.UltraDateTimeEditor();
			this.txtAutFactura = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label15 = new System.Windows.Forms.Label();
			this.groupBox3 = new System.Windows.Forms.GroupBox();
			this.ultraCheckEditor1 = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.chkProspecto = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.label10 = new System.Windows.Forms.Label();
			this.ultraTextEditor5 = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label6 = new System.Windows.Forms.Label();
			this.ultraCalendarCombo1 = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.cmbFecha = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.ultraTabPageControl7 = new Infragistics.Win.UltraWinTabControl.UltraTabPageControl();
			this.ultraGrid3 = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.cdsDirecciones = new C1.Data.C1DataSet();
			this.errorProvider = new System.Windows.Forms.ErrorProvider();
			this.txtIdCliente = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.barraDato1 = new BarraDatoSQL.BarraDatoSQL();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.tbBasico = new Infragistics.Win.UltraWinTabControl.UltraTabControl();
			this.ultraTabSharedControlsPage1 = new Infragistics.Win.UltraWinTabControl.UltraTabSharedControlsPage();
			this.btCRMClientes = new Infragistics.Win.Misc.UltraButton();
			this.imageList = new System.Windows.Forms.ImageList(this.components);
			this.btAuditoria = new Infragistics.Win.Misc.UltraButton();
			this.toolTip1 = new System.Windows.Forms.ToolTip(this.components);
			this.btProveedorCliente = new Infragistics.Win.Misc.UltraButton();
			this.cdsPais = new C1.Data.C1DataSet();
			this.ultraGridExcelExporter1 = new Infragistics.Win.UltraWinGrid.ExcelExport.UltraGridExcelExporter();
			this.ultraTabPageControl1.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.cdsCliente)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraTextEditor1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbSujetoRetencion)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource4)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbAgenteDeRetencion)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbTipoProveedor)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDiasConsignacion)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbTipoRuc)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtRuc)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbPais)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtPais)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNombreComercial)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNotas)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtObservacion)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCelular)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtEMail)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtRepresentante)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCiudad)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbGrupo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNombre)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDireccion)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtTelefono)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtFax)).BeginInit();
			this.tpComplemento.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.txtDiasCredito4)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbInstitucion)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource3)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbArticuloSRI)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsArticuloTabla)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbBanco)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtBeneficiario)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCuentaSwift)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbDias)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbClientePublicidad)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbPrecio)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbSector)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCodigo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDiasCredito3)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDiasCredito2)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbTipoCredito)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsContacto)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCredTribut)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbPorcIvaRet)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbVendedor)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.spnDescCredito)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.spnDescuento)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtValorCredito)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDiasCredito1)).BeginInit();
			this.ultraTabPageControl5.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid2)).BeginInit();
			this.ultraTabPageControl2.SuspendLayout();
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
			((System.ComponentModel.ISupportInitialize)(this.txtCedRep)).BeginInit();
			this.ultraTabPageControl3.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).BeginInit();
			this.ultraTabPageControl4.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.grdItems)).BeginInit();
			this.ultraTabPageControl6.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.radioGroup1.Properties)).BeginInit();
			this.groupBox4.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.cmbFormaPago)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtAutRetencion)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCobrador)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtAutImprenta)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCuentaFondo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCuentaAnticipo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbDepende)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraTextEditor21)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraTextEditor4)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtSerie)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbFechaCaduca)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtAutFactura)).BeginInit();
			this.groupBox3.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.ultraTextEditor5)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraCalendarCombo1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbFecha)).BeginInit();
			this.ultraTabPageControl7.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid3)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsDirecciones)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdCliente)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.tbBasico)).BeginInit();
			this.tbBasico.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.cdsPais)).BeginInit();
			this.SuspendLayout();
			// 
			// ultraTabPageControl1
			// 
			this.ultraTabPageControl1.Controls.Add(this.chkCourier);
			this.ultraTabPageControl1.Controls.Add(this.ultraTextEditor1);
			this.ultraTabPageControl1.Controls.Add(this.cmbSujetoRetencion);
			this.ultraTabPageControl1.Controls.Add(this.cmbAgenteDeRetencion);
			this.ultraTabPageControl1.Controls.Add(this.lblASRetencion);
			this.ultraTabPageControl1.Controls.Add(this.chkCodigoBarras);
			this.ultraTabPageControl1.Controls.Add(this.cmbTipoProveedor);
			this.ultraTabPageControl1.Controls.Add(this.lblTipoProveedor);
			this.ultraTabPageControl1.Controls.Add(this.label27);
			this.ultraTabPageControl1.Controls.Add(this.txtDiasConsignacion);
			this.ultraTabPageControl1.Controls.Add(this.chkAplicaConsignacion);
			this.ultraTabPageControl1.Controls.Add(this.cmbTipoRuc);
			this.ultraTabPageControl1.Controls.Add(this.label22);
			this.ultraTabPageControl1.Controls.Add(this.txtRuc);
			this.ultraTabPageControl1.Controls.Add(this.btVerifica);
			this.ultraTabPageControl1.Controls.Add(this.cmbPais);
			this.ultraTabPageControl1.Controls.Add(this.txtPais);
			this.ultraTabPageControl1.Controls.Add(this.bPais);
			this.ultraTabPageControl1.Controls.Add(this.txtNombreComercial);
			this.ultraTabPageControl1.Controls.Add(this.label28);
			this.ultraTabPageControl1.Controls.Add(this.lblNotas);
			this.ultraTabPageControl1.Controls.Add(this.txtNotas);
			this.ultraTabPageControl1.Controls.Add(this.label26);
			this.ultraTabPageControl1.Controls.Add(this.txtObservacion);
			this.ultraTabPageControl1.Controls.Add(this.txtCelular);
			this.ultraTabPageControl1.Controls.Add(this.label9);
			this.ultraTabPageControl1.Controls.Add(this.label5);
			this.ultraTabPageControl1.Controls.Add(this.txtEMail);
			this.ultraTabPageControl1.Controls.Add(this.txtRepresentante);
			this.ultraTabPageControl1.Controls.Add(this.label7);
			this.ultraTabPageControl1.Controls.Add(this.label8);
			this.ultraTabPageControl1.Controls.Add(this.btGrupoCliente);
			this.ultraTabPageControl1.Controls.Add(this.cmbCiudad);
			this.ultraTabPageControl1.Controls.Add(this.label1);
			this.ultraTabPageControl1.Controls.Add(this.cmbGrupo);
			this.ultraTabPageControl1.Controls.Add(this.label11);
			this.ultraTabPageControl1.Controls.Add(this.bCiudad);
			this.ultraTabPageControl1.Controls.Add(this.label13);
			this.ultraTabPageControl1.Controls.Add(this.label4);
			this.ultraTabPageControl1.Controls.Add(this.label3);
			this.ultraTabPageControl1.Controls.Add(this.label17);
			this.ultraTabPageControl1.Controls.Add(this.txtNombre);
			this.ultraTabPageControl1.Controls.Add(this.txtDireccion);
			this.ultraTabPageControl1.Controls.Add(this.txtTelefono);
			this.ultraTabPageControl1.Controls.Add(this.txtFax);
			this.ultraTabPageControl1.Location = new System.Drawing.Point(1, 23);
			this.ultraTabPageControl1.Name = "ultraTabPageControl1";
			this.ultraTabPageControl1.Size = new System.Drawing.Size(938, 422);
			this.ultraTabPageControl1.Paint += new System.Windows.Forms.PaintEventHandler(this.ultraTabPageControl1_Paint);
			// 
			// chkCourier
			// 
			appearance1.ForeColorDisabled = System.Drawing.Color.Black;
			this.chkCourier.Appearance = appearance1;
			this.chkCourier.BackColor = System.Drawing.Color.Transparent;
			this.chkCourier.DataBindings.Add(new System.Windows.Forms.Binding("CheckedValue", this.cdsCliente, "Cliente.Courier"));
			this.chkCourier.Location = new System.Drawing.Point(672, 392);
			this.chkCourier.Name = "chkCourier";
			this.chkCourier.Size = new System.Drawing.Size(127, 21);
			this.chkCourier.TabIndex = 246;
			this.chkCourier.Text = "Courier";
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
			// ultraTextEditor1
			// 
			appearance2.BackColorDisabled = System.Drawing.Color.Silver;
			appearance2.BackColorDisabled2 = System.Drawing.Color.Silver;
			appearance2.BorderColor = System.Drawing.Color.Black;
			appearance2.BorderColor3DBase = System.Drawing.Color.Black;
			appearance2.ForeColorDisabled = System.Drawing.Color.Black;
			this.ultraTextEditor1.Appearance = appearance2;
			this.ultraTextEditor1.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.ultraTextEditor1.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsCliente, "Cliente.NResolucion"));
			this.ultraTextEditor1.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.ultraTextEditor1.Location = new System.Drawing.Point(504, 208);
			this.ultraTextEditor1.Name = "ultraTextEditor1";
			this.ultraTextEditor1.Size = new System.Drawing.Size(80, 22);
			this.ultraTextEditor1.TabIndex = 245;
			// 
			// cmbSujetoRetencion
			// 
			appearance3.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbSujetoRetencion.Appearance = appearance3;
			this.cmbSujetoRetencion.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.cmbSujetoRetencion.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbSujetoRetencion.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsCliente, "Cliente.idSujetoRetencion"));
			this.cmbSujetoRetencion.DataSource = this.ultraDataSource4;
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridColumn2.Width = 328;
			ultraGridBand1.Columns.AddRange(new object[] {
																										 ultraGridColumn1,
																										 ultraGridColumn2});
			this.cmbSujetoRetencion.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.cmbSujetoRetencion.DisplayMember = "Nombre";
			this.cmbSujetoRetencion.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbSujetoRetencion.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbSujetoRetencion.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbSujetoRetencion.Location = new System.Drawing.Point(200, 208);
			this.cmbSujetoRetencion.Name = "cmbSujetoRetencion";
			this.cmbSujetoRetencion.Size = new System.Drawing.Size(296, 21);
			this.cmbSujetoRetencion.TabIndex = 244;
			this.cmbSujetoRetencion.ValueMember = "idSujetoRetencion";
			this.cmbSujetoRetencion.Visible = false;
			// 
			// ultraDataSource4
			// 
			ultraDataColumn1.DataType = typeof(int);
			this.ultraDataSource4.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn1,
																																 ultraDataColumn2});
			// 
			// cmbAgenteDeRetencion
			// 
			appearance4.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbAgenteDeRetencion.Appearance = appearance4;
			this.cmbAgenteDeRetencion.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.cmbAgenteDeRetencion.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbAgenteDeRetencion.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsCliente, "Cliente.idAgenteRetencion"));
			this.cmbAgenteDeRetencion.DataSource = this.ultraDataSource2;
			ultraGridColumn3.Header.VisiblePosition = 0;
			ultraGridColumn3.Hidden = true;
			ultraGridColumn4.Header.VisiblePosition = 1;
			ultraGridColumn4.Width = 350;
			ultraGridBand2.Columns.AddRange(new object[] {
																										 ultraGridColumn3,
																										 ultraGridColumn4});
			this.cmbAgenteDeRetencion.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			this.cmbAgenteDeRetencion.DisplayMember = "Nombre";
			this.cmbAgenteDeRetencion.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbAgenteDeRetencion.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbAgenteDeRetencion.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbAgenteDeRetencion.Location = new System.Drawing.Point(200, 208);
			this.cmbAgenteDeRetencion.Name = "cmbAgenteDeRetencion";
			this.cmbAgenteDeRetencion.Size = new System.Drawing.Size(296, 21);
			this.cmbAgenteDeRetencion.TabIndex = 243;
			this.cmbAgenteDeRetencion.ValueMember = "idAgenteRetencion";
			this.cmbAgenteDeRetencion.Visible = false;
			// 
			// ultraDataSource2
			// 
			ultraDataColumn3.DataType = typeof(int);
			this.ultraDataSource2.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn3,
																																 ultraDataColumn4});
			// 
			// lblASRetencion
			// 
			this.lblASRetencion.AutoSize = true;
			this.lblASRetencion.BackColor = System.Drawing.Color.Transparent;
			this.lblASRetencion.Location = new System.Drawing.Point(48, 208);
			this.lblASRetencion.Name = "lblASRetencion";
			this.lblASRetencion.Size = new System.Drawing.Size(105, 16);
			this.lblASRetencion.TabIndex = 241;
			this.lblASRetencion.Text = "Agente de retención";
			this.lblASRetencion.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// chkCodigoBarras
			// 
			appearance5.ForeColorDisabled = System.Drawing.Color.Black;
			this.chkCodigoBarras.Appearance = appearance5;
			this.chkCodigoBarras.BackColor = System.Drawing.Color.Transparent;
			this.chkCodigoBarras.DataBindings.Add(new System.Windows.Forms.Binding("CheckedValue", this.cdsCliente, "Cliente.CodigoBarras"));
			this.chkCodigoBarras.Location = new System.Drawing.Point(536, 392);
			this.chkCodigoBarras.Name = "chkCodigoBarras";
			this.chkCodigoBarras.Size = new System.Drawing.Size(127, 21);
			this.chkCodigoBarras.TabIndex = 239;
			this.chkCodigoBarras.Text = "Codigo de Barras";
			this.chkCodigoBarras.Visible = false;
			// 
			// cmbTipoProveedor
			// 
			appearance6.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbTipoProveedor.Appearance = appearance6;
			this.cmbTipoProveedor.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbTipoProveedor.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbTipoProveedor.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsCliente, "Cliente.idTipoProveedor"));
			this.cmbTipoProveedor.DataSource = this.ultraDataSource1;
			ultraGridColumn5.Header.VisiblePosition = 0;
			ultraGridColumn5.Hidden = true;
			ultraGridColumn6.Header.VisiblePosition = 1;
			ultraGridColumn6.Width = 60;
			ultraGridColumn7.Header.VisiblePosition = 2;
			ultraGridColumn7.Width = 200;
			ultraGridBand3.Columns.AddRange(new object[] {
																										 ultraGridColumn5,
																										 ultraGridColumn6,
																										 ultraGridColumn7});
			this.cmbTipoProveedor.DisplayLayout.BandsSerializer.Add(ultraGridBand3);
			this.cmbTipoProveedor.DisplayMember = "Tipo";
			this.cmbTipoProveedor.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbTipoProveedor.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbTipoProveedor.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbTipoProveedor.Location = new System.Drawing.Point(696, 208);
			this.cmbTipoProveedor.Name = "cmbTipoProveedor";
			this.cmbTipoProveedor.Size = new System.Drawing.Size(168, 21);
			this.cmbTipoProveedor.TabIndex = 238;
			this.cmbTipoProveedor.ValueMember = "idTipoProveedor";
			this.cmbTipoProveedor.Visible = false;
			// 
			// ultraDataSource1
			// 
			ultraDataColumn5.DataType = typeof(int);
			this.ultraDataSource1.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn5,
																																 ultraDataColumn6,
																																 ultraDataColumn7});
			// 
			// lblTipoProveedor
			// 
			this.lblTipoProveedor.AutoSize = true;
			this.lblTipoProveedor.BackColor = System.Drawing.Color.Transparent;
			this.lblTipoProveedor.Location = new System.Drawing.Point(592, 208);
			this.lblTipoProveedor.Name = "lblTipoProveedor";
			this.lblTipoProveedor.Size = new System.Drawing.Size(96, 16);
			this.lblTipoProveedor.TabIndex = 159;
			this.lblTipoProveedor.Text = "Tipo de Proveedor";
			this.lblTipoProveedor.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.lblTipoProveedor.Visible = false;
			// 
			// label27
			// 
			this.label27.AutoSize = true;
			this.label27.BackColor = System.Drawing.Color.Transparent;
			this.label27.Location = new System.Drawing.Point(336, 394);
			this.label27.Name = "label27";
			this.label27.Size = new System.Drawing.Size(110, 16);
			this.label27.TabIndex = 157;
			this.label27.Text = "Dias de Consignación";
			this.label27.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtDiasConsignacion
			// 
			appearance7.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtDiasConsignacion.Appearance = appearance7;
			this.txtDiasConsignacion.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsCliente, "Cliente.DiasConsignacion"));
			this.txtDiasConsignacion.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtDiasConsignacion.FormatString = "0";
			this.txtDiasConsignacion.Location = new System.Drawing.Point(456, 391);
			this.txtDiasConsignacion.MinValue = 0;
			this.txtDiasConsignacion.Name = "txtDiasConsignacion";
			this.txtDiasConsignacion.PromptChar = ' ';
			this.txtDiasConsignacion.Size = new System.Drawing.Size(64, 22);
			this.txtDiasConsignacion.SpinButtonDisplayStyle = Infragistics.Win.ButtonDisplayStyle.Always;
			this.txtDiasConsignacion.TabIndex = 21;
			this.txtDiasConsignacion.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtDiasConsignacion_KeyPress);
			this.txtDiasConsignacion.ValueChanged += new System.EventHandler(this.txtDiasConsignacion_ValueChanged);
			// 
			// chkAplicaConsignacion
			// 
			appearance8.ForeColorDisabled = System.Drawing.Color.Black;
			this.chkAplicaConsignacion.Appearance = appearance8;
			this.chkAplicaConsignacion.BackColor = System.Drawing.Color.Transparent;
			this.chkAplicaConsignacion.DataBindings.Add(new System.Windows.Forms.Binding("CheckedValue", this.cdsCliente, "Cliente.AplicaConsignacion"));
			this.chkAplicaConsignacion.Location = new System.Drawing.Point(200, 392);
			this.chkAplicaConsignacion.Name = "chkAplicaConsignacion";
			this.chkAplicaConsignacion.Size = new System.Drawing.Size(127, 21);
			this.chkAplicaConsignacion.TabIndex = 20;
			this.chkAplicaConsignacion.Text = "Aplica Consignación";
			this.chkAplicaConsignacion.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.chkAplicaConsignacion_KeyPress);
			// 
			// cmbTipoRuc
			// 
			appearance9.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbTipoRuc.Appearance = appearance9;
			this.cmbTipoRuc.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbTipoRuc.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbTipoRuc.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsCliente, "Cliente.idTipoRuc"));
			ultraGridColumn8.Header.VisiblePosition = 0;
			ultraGridColumn8.Hidden = true;
			ultraGridColumn9.Header.VisiblePosition = 1;
			ultraGridColumn9.Width = 233;
			ultraGridBand4.Columns.AddRange(new object[] {
																										 ultraGridColumn8,
																										 ultraGridColumn9});
			this.cmbTipoRuc.DisplayLayout.BandsSerializer.Add(ultraGridBand4);
			this.cmbTipoRuc.DisplayMember = "TipoRuc";
			this.cmbTipoRuc.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbTipoRuc.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbTipoRuc.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbTipoRuc.Location = new System.Drawing.Point(201, 16);
			this.cmbTipoRuc.Name = "cmbTipoRuc";
			this.cmbTipoRuc.Size = new System.Drawing.Size(144, 21);
			this.cmbTipoRuc.TabIndex = 0;
			this.cmbTipoRuc.ValueMember = "idTipoRuc";
			this.cmbTipoRuc.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.cmbTipoRuc_KeyPress);
			this.cmbTipoRuc.Validated += new System.EventHandler(this.cmbTipoRuc_Validated);
			this.cmbTipoRuc.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.cmbTipoRuc_InitializeLayout);
			// 
			// label22
			// 
			this.label22.AutoSize = true;
			this.label22.BackColor = System.Drawing.Color.Transparent;
			this.label22.Location = new System.Drawing.Point(48, 16);
			this.label22.Name = "label22";
			this.label22.Size = new System.Drawing.Size(49, 16);
			this.label22.TabIndex = 156;
			this.label22.Text = "RUC/Ced";
			this.label22.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtRuc
			// 
			appearance10.BackColorDisabled = System.Drawing.Color.Silver;
			appearance10.BackColorDisabled2 = System.Drawing.Color.Silver;
			appearance10.BorderColor = System.Drawing.Color.Black;
			appearance10.BorderColor3DBase = System.Drawing.Color.Black;
			appearance10.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtRuc.Appearance = appearance10;
			this.txtRuc.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtRuc.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsCliente, "Cliente.Ruc"));
			this.txtRuc.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtRuc.Location = new System.Drawing.Point(361, 16);
			this.txtRuc.Name = "txtRuc";
			this.txtRuc.Size = new System.Drawing.Size(136, 22);
			this.txtRuc.TabIndex = 1;
			this.txtRuc.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtRuc_KeyDown);
			this.txtRuc.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtRuc_KeyPress);
			this.txtRuc.Validating += new System.ComponentModel.CancelEventHandler(this.txtRuc_Validating);
			this.txtRuc.Validated += new System.EventHandler(this.txtRuc_Validated);
			this.txtRuc.ValueChanged += new System.EventHandler(this.txtRuc_ValueChanged);
			// 
			// btVerifica
			// 
			appearance11.Image = ((object)(resources.GetObject("appearance11.Image")));
			this.btVerifica.Appearance = appearance11;
			this.btVerifica.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.btVerifica.Location = new System.Drawing.Point(505, 16);
			this.btVerifica.Name = "btVerifica";
			this.btVerifica.Size = new System.Drawing.Size(24, 22);
			this.btVerifica.TabIndex = 2;
			this.btVerifica.Visible = false;
			this.btVerifica.Click += new System.EventHandler(this.btVerifica_Click);
			// 
			// cmbPais
			// 
			appearance12.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbPais.Appearance = appearance12;
			this.cmbPais.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.cmbPais.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbPais.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsCliente, "Cliente.idPais"));
			ultraGridColumn10.Header.VisiblePosition = 0;
			ultraGridColumn10.Hidden = true;
			ultraGridColumn10.Width = 8;
			ultraGridColumn11.Header.VisiblePosition = 1;
			ultraGridColumn11.Width = 152;
			ultraGridBand5.Columns.AddRange(new object[] {
																										 ultraGridColumn10,
																										 ultraGridColumn11});
			this.cmbPais.DisplayLayout.BandsSerializer.Add(ultraGridBand5);
			this.cmbPais.DisplayMember = "Pais";
			this.cmbPais.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbPais.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbPais.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbPais.Location = new System.Drawing.Point(697, 144);
			this.cmbPais.Name = "cmbPais";
			this.cmbPais.Size = new System.Drawing.Size(136, 21);
			this.cmbPais.TabIndex = 13;
			this.cmbPais.ValueMember = "idClientePais";
			this.cmbPais.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.cmbPais_KeyPress);
			// 
			// txtPais
			// 
			appearance13.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtPais.Appearance = appearance13;
			this.txtPais.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtPais.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsCliente, "Cliente.Pais"));
			this.txtPais.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtPais.Location = new System.Drawing.Point(697, 144);
			this.txtPais.Name = "txtPais";
			this.txtPais.Size = new System.Drawing.Size(168, 22);
			this.txtPais.TabIndex = 15;
			this.txtPais.Visible = false;
			// 
			// bPais
			// 
			this.bPais.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.bPais.Location = new System.Drawing.Point(841, 144);
			this.bPais.Name = "bPais";
			this.bPais.Size = new System.Drawing.Size(24, 21);
			this.bPais.TabIndex = 14;
			this.bPais.Text = "...";
			this.bPais.Click += new System.EventHandler(this.bPais_Click);
			// 
			// txtNombreComercial
			// 
			appearance14.BackColorDisabled = System.Drawing.Color.Silver;
			appearance14.BackColorDisabled2 = System.Drawing.Color.Silver;
			appearance14.BorderColor = System.Drawing.Color.Black;
			appearance14.BorderColor3DBase = System.Drawing.Color.Black;
			appearance14.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtNombreComercial.Appearance = appearance14;
			this.txtNombreComercial.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtNombreComercial.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsCliente, "Cliente.Comercial"));
			this.txtNombreComercial.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNombreComercial.Location = new System.Drawing.Point(201, 176);
			this.txtNombreComercial.Name = "txtNombreComercial";
			this.txtNombreComercial.Size = new System.Drawing.Size(328, 22);
			this.txtNombreComercial.TabIndex = 8;
			this.toolTip1.SetToolTip(this.txtNombreComercial, "Nombre comercial");
			this.txtNombreComercial.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtNombreComercial_KeyPress);
			// 
			// label28
			// 
			this.label28.AutoSize = true;
			this.label28.BackColor = System.Drawing.Color.Transparent;
			this.label28.Location = new System.Drawing.Point(48, 176);
			this.label28.Name = "label28";
			this.label28.Size = new System.Drawing.Size(96, 16);
			this.label28.TabIndex = 148;
			this.label28.Text = "Nombre Comercial";
			this.label28.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// lblNotas
			// 
			this.lblNotas.AutoSize = true;
			this.lblNotas.BackColor = System.Drawing.Color.Transparent;
			this.lblNotas.Location = new System.Drawing.Point(48, 312);
			this.lblNotas.Name = "lblNotas";
			this.lblNotas.Size = new System.Drawing.Size(33, 16);
			this.lblNotas.TabIndex = 105;
			this.lblNotas.Text = "Notas";
			this.lblNotas.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtNotas
			// 
			appearance15.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtNotas.Appearance = appearance15;
			this.txtNotas.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtNotas.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsCliente, "Cliente.Notas"));
			this.txtNotas.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNotas.Location = new System.Drawing.Point(201, 312);
			this.txtNotas.Multiline = true;
			this.txtNotas.Name = "txtNotas";
			this.txtNotas.Size = new System.Drawing.Size(664, 72);
			this.txtNotas.TabIndex = 19;
			this.txtNotas.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtNotas_KeyPress);
			this.txtNotas.ValueChanged += new System.EventHandler(this.txtNotas_ValueChanged);
			// 
			// label26
			// 
			this.label26.AutoSize = true;
			this.label26.BackColor = System.Drawing.Color.Transparent;
			this.label26.Location = new System.Drawing.Point(48, 240);
			this.label26.Name = "label26";
			this.label26.Size = new System.Drawing.Size(65, 16);
			this.label26.TabIndex = 103;
			this.label26.Text = "Observación";
			this.label26.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtObservacion
			// 
			appearance16.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtObservacion.Appearance = appearance16;
			this.txtObservacion.BorderStyle = Infragistics.Win.UIElementBorderStyle.Raised;
			this.txtObservacion.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtObservacion.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsCliente, "Cliente.Observacion"));
			this.txtObservacion.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtObservacion.Location = new System.Drawing.Point(201, 240);
			this.txtObservacion.Multiline = true;
			this.txtObservacion.Name = "txtObservacion";
			this.txtObservacion.Size = new System.Drawing.Size(664, 64);
			this.txtObservacion.TabIndex = 18;
			this.txtObservacion.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtObservacion_KeyPress);
			// 
			// txtCelular
			// 
			appearance17.BackColorDisabled = System.Drawing.Color.Silver;
			appearance17.BackColorDisabled2 = System.Drawing.Color.Silver;
			appearance17.BorderColor = System.Drawing.Color.Black;
			appearance17.BorderColor3DBase = System.Drawing.Color.Black;
			appearance17.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtCelular.Appearance = appearance17;
			this.txtCelular.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtCelular.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsCliente, "Cliente.Celular"));
			this.txtCelular.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtCelular.Location = new System.Drawing.Point(697, 80);
			this.txtCelular.MaxLength = 10;
			this.txtCelular.Name = "txtCelular";
			this.txtCelular.Size = new System.Drawing.Size(168, 22);
			this.txtCelular.TabIndex = 11;
			this.txtCelular.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtCelular_KeyPress);
			this.txtCelular.ValueChanged += new System.EventHandler(this.txtCelular_ValueChanged);
			// 
			// label9
			// 
			this.label9.AutoSize = true;
			this.label9.BackColor = System.Drawing.Color.Transparent;
			this.label9.Location = new System.Drawing.Point(592, 80);
			this.label9.Name = "label9";
			this.label9.Size = new System.Drawing.Size(38, 16);
			this.label9.TabIndex = 87;
			this.label9.Text = "Celular";
			this.label9.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label5
			// 
			this.label5.AutoSize = true;
			this.label5.BackColor = System.Drawing.Color.Transparent;
			this.label5.Location = new System.Drawing.Point(592, 176);
			this.label5.Name = "label5";
			this.label5.Size = new System.Drawing.Size(42, 16);
			this.label5.TabIndex = 10;
			this.label5.Text = "Ciudad:";
			this.label5.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtEMail
			// 
			appearance18.BackColorDisabled = System.Drawing.Color.Silver;
			appearance18.BackColorDisabled2 = System.Drawing.Color.Silver;
			appearance18.BorderColor = System.Drawing.Color.Black;
			appearance18.BorderColor3DBase = System.Drawing.Color.Black;
			appearance18.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtEMail.Appearance = appearance18;
			this.txtEMail.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsCliente, "Cliente.Email"));
			this.txtEMail.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtEMail.Location = new System.Drawing.Point(697, 112);
			this.txtEMail.Name = "txtEMail";
			this.txtEMail.Size = new System.Drawing.Size(168, 22);
			this.txtEMail.TabIndex = 12;
			this.txtEMail.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtEMail_KeyPress);
			// 
			// txtRepresentante
			// 
			appearance19.BackColorDisabled = System.Drawing.Color.Silver;
			appearance19.BackColorDisabled2 = System.Drawing.Color.Silver;
			appearance19.BorderColor = System.Drawing.Color.Black;
			appearance19.BorderColor3DBase = System.Drawing.Color.Black;
			appearance19.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtRepresentante.Appearance = appearance19;
			this.txtRepresentante.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtRepresentante.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsCliente, "Cliente.Representante"));
			this.txtRepresentante.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtRepresentante.Location = new System.Drawing.Point(201, 144);
			this.txtRepresentante.Name = "txtRepresentante";
			this.txtRepresentante.Size = new System.Drawing.Size(328, 22);
			this.txtRepresentante.TabIndex = 7;
			this.txtRepresentante.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtRepresentante_KeyPress);
			// 
			// label7
			// 
			this.label7.AutoSize = true;
			this.label7.BackColor = System.Drawing.Color.Transparent;
			this.label7.Location = new System.Drawing.Point(592, 16);
			this.label7.Name = "label7";
			this.label7.Size = new System.Drawing.Size(52, 16);
			this.label7.TabIndex = 14;
			this.label7.Text = "Teléfono:";
			this.label7.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label8
			// 
			this.label8.AutoSize = true;
			this.label8.BackColor = System.Drawing.Color.Transparent;
			this.label8.Location = new System.Drawing.Point(592, 48);
			this.label8.Name = "label8";
			this.label8.Size = new System.Drawing.Size(26, 16);
			this.label8.TabIndex = 16;
			this.label8.Text = "Fax:";
			this.label8.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// btGrupoCliente
			// 
			this.btGrupoCliente.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.btGrupoCliente.Location = new System.Drawing.Point(505, 112);
			this.btGrupoCliente.Name = "btGrupoCliente";
			this.btGrupoCliente.Size = new System.Drawing.Size(24, 21);
			this.btGrupoCliente.TabIndex = 6;
			this.btGrupoCliente.Text = "...";
			this.btGrupoCliente.Click += new System.EventHandler(this.btGrupoCliente_Click);
			// 
			// cmbCiudad
			// 
			appearance20.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbCiudad.Appearance = appearance20;
			this.cmbCiudad.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.cmbCiudad.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbCiudad.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsCliente, "Cliente.idCiudad"));
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
			this.cmbCiudad.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbCiudad.Location = new System.Drawing.Point(697, 176);
			this.cmbCiudad.Name = "cmbCiudad";
			this.cmbCiudad.Size = new System.Drawing.Size(136, 21);
			this.cmbCiudad.TabIndex = 16;
			this.cmbCiudad.ValueMember = "idCiudad";
			this.cmbCiudad.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.cmbCiudad_KeyPress);
			this.cmbCiudad.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.cmbCiudad_InitializeLayout);
			// 
			// label1
			// 
			this.label1.AutoSize = true;
			this.label1.BackColor = System.Drawing.Color.Transparent;
			this.label1.Location = new System.Drawing.Point(48, 48);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(48, 16);
			this.label1.TabIndex = 1;
			this.label1.Text = "Nombre:";
			this.label1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// cmbGrupo
			// 
			appearance21.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbGrupo.Appearance = appearance21;
			this.cmbGrupo.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.cmbGrupo.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbGrupo.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsCliente, "Cliente.idGrupoCliente"));
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
			this.cmbGrupo.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbGrupo.Location = new System.Drawing.Point(201, 112);
			this.cmbGrupo.Name = "cmbGrupo";
			this.cmbGrupo.Size = new System.Drawing.Size(296, 21);
			this.cmbGrupo.TabIndex = 5;
			this.cmbGrupo.ValueMember = "idGrupoCliente";
			this.cmbGrupo.MouseDown += new System.Windows.Forms.MouseEventHandler(this.cmbGrupo_MouseDown);
			this.cmbGrupo.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.cmbGrupo_KeyPress);
			this.cmbGrupo.Validated += new System.EventHandler(this.cmbGrupo_Validated);
			this.cmbGrupo.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.cmbGrupo_InitializeLayout);
			// 
			// label11
			// 
			this.label11.AutoSize = true;
			this.label11.BackColor = System.Drawing.Color.Transparent;
			this.label11.Location = new System.Drawing.Point(592, 112);
			this.label11.Name = "label11";
			this.label11.Size = new System.Drawing.Size(35, 16);
			this.label11.TabIndex = 24;
			this.label11.Text = "Email:";
			this.label11.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// bCiudad
			// 
			this.bCiudad.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.bCiudad.Location = new System.Drawing.Point(841, 176);
			this.bCiudad.Name = "bCiudad";
			this.bCiudad.Size = new System.Drawing.Size(24, 21);
			this.bCiudad.TabIndex = 17;
			this.bCiudad.Text = "...";
			this.bCiudad.Click += new System.EventHandler(this.bCiudad_Click);
			// 
			// label13
			// 
			this.label13.AutoSize = true;
			this.label13.BackColor = System.Drawing.Color.Transparent;
			this.label13.Location = new System.Drawing.Point(48, 144);
			this.label13.Name = "label13";
			this.label13.Size = new System.Drawing.Size(77, 16);
			this.label13.TabIndex = 28;
			this.label13.Text = "Representante";
			this.label13.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label4
			// 
			this.label4.AutoSize = true;
			this.label4.BackColor = System.Drawing.Color.Transparent;
			this.label4.Location = new System.Drawing.Point(592, 144);
			this.label4.Name = "label4";
			this.label4.Size = new System.Drawing.Size(28, 16);
			this.label4.TabIndex = 8;
			this.label4.Text = "País:";
			this.label4.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label3
			// 
			this.label3.AutoSize = true;
			this.label3.BackColor = System.Drawing.Color.Transparent;
			this.label3.Location = new System.Drawing.Point(48, 80);
			this.label3.Name = "label3";
			this.label3.Size = new System.Drawing.Size(54, 16);
			this.label3.TabIndex = 5;
			this.label3.Text = "Dirección:";
			this.label3.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label17
			// 
			this.label17.AutoSize = true;
			this.label17.BackColor = System.Drawing.Color.Transparent;
			this.label17.Location = new System.Drawing.Point(48, 112);
			this.label17.Name = "label17";
			this.label17.Size = new System.Drawing.Size(39, 16);
			this.label17.TabIndex = 36;
			this.label17.Text = "Grupo:";
			this.label17.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtNombre
			// 
			appearance22.BackColorDisabled = System.Drawing.Color.Silver;
			appearance22.BackColorDisabled2 = System.Drawing.Color.Silver;
			appearance22.BorderColor = System.Drawing.Color.Black;
			appearance22.BorderColor3DBase = System.Drawing.Color.Black;
			appearance22.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtNombre.Appearance = appearance22;
			this.txtNombre.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtNombre.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsCliente, "Cliente.Nombre"));
			this.txtNombre.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNombre.Location = new System.Drawing.Point(201, 48);
			this.txtNombre.Name = "txtNombre";
			this.txtNombre.Size = new System.Drawing.Size(328, 22);
			this.txtNombre.TabIndex = 3;
			this.txtNombre.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtNombre_KeyPress);
			// 
			// txtDireccion
			// 
			appearance23.BackColorDisabled = System.Drawing.Color.Silver;
			appearance23.BackColorDisabled2 = System.Drawing.Color.Silver;
			appearance23.BorderColor = System.Drawing.Color.Black;
			appearance23.BorderColor3DBase = System.Drawing.Color.Black;
			appearance23.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtDireccion.Appearance = appearance23;
			this.txtDireccion.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtDireccion.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsCliente, "Cliente.Direccion"));
			this.txtDireccion.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtDireccion.Location = new System.Drawing.Point(201, 80);
			this.txtDireccion.Name = "txtDireccion";
			this.txtDireccion.Size = new System.Drawing.Size(328, 22);
			this.txtDireccion.TabIndex = 4;
			this.txtDireccion.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtDireccion_KeyPress);
			// 
			// txtTelefono
			// 
			appearance24.BackColorDisabled = System.Drawing.Color.Silver;
			appearance24.BackColorDisabled2 = System.Drawing.Color.Silver;
			appearance24.BorderColor = System.Drawing.Color.Black;
			appearance24.BorderColor3DBase = System.Drawing.Color.Black;
			appearance24.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtTelefono.Appearance = appearance24;
			this.txtTelefono.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtTelefono.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsCliente, "Cliente.Telefono"));
			this.txtTelefono.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtTelefono.Location = new System.Drawing.Point(697, 16);
			this.txtTelefono.MaxLength = 9;
			this.txtTelefono.Name = "txtTelefono";
			this.txtTelefono.Size = new System.Drawing.Size(168, 22);
			this.txtTelefono.TabIndex = 9;
			this.txtTelefono.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtTelefono_KeyPress);
			this.txtTelefono.ValueChanged += new System.EventHandler(this.txtTelefono_ValueChanged);
			// 
			// txtFax
			// 
			appearance25.BackColorDisabled = System.Drawing.Color.Silver;
			appearance25.BackColorDisabled2 = System.Drawing.Color.Silver;
			appearance25.BorderColor = System.Drawing.Color.Black;
			appearance25.BorderColor3DBase = System.Drawing.Color.Black;
			appearance25.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtFax.Appearance = appearance25;
			this.txtFax.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtFax.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsCliente, "Cliente.Fax"));
			this.txtFax.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtFax.Location = new System.Drawing.Point(697, 48);
			this.txtFax.MaxLength = 9;
			this.txtFax.Name = "txtFax";
			this.txtFax.Size = new System.Drawing.Size(168, 22);
			this.txtFax.TabIndex = 10;
			this.txtFax.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtFax_KeyPress);
			// 
			// tpComplemento
			// 
			this.tpComplemento.Controls.Add(this.txtDiasCredito4);
			this.tpComplemento.Controls.Add(this.label56);
			this.tpComplemento.Controls.Add(this.cmbInstitucion);
			this.tpComplemento.Controls.Add(this.label57);
			this.tpComplemento.Controls.Add(this.cmbArticuloSRI);
			this.tpComplemento.Controls.Add(this.lblArticuloSRI);
			this.tpComplemento.Controls.Add(this.btBanco);
			this.tpComplemento.Controls.Add(this.cmbBanco);
			this.tpComplemento.Controls.Add(this.txtBeneficiario);
			this.tpComplemento.Controls.Add(this.lblBeneficiario);
			this.tpComplemento.Controls.Add(this.txtCuentaSwift);
			this.tpComplemento.Controls.Add(this.lblCuentaSwift);
			this.tpComplemento.Controls.Add(this.lblBanco);
			this.tpComplemento.Controls.Add(this.dtHoraHasta);
			this.tpComplemento.Controls.Add(this.dtHoraDesde);
			this.tpComplemento.Controls.Add(this.cmbDias);
			this.tpComplemento.Controls.Add(this.lblHoraHasta);
			this.tpComplemento.Controls.Add(this.lblHoraDesde);
			this.tpComplemento.Controls.Add(this.lblDias);
			this.tpComplemento.Controls.Add(this.chkHaceRetencion);
			this.tpComplemento.Controls.Add(this.rdbCredito);
			this.tpComplemento.Controls.Add(this.rdbContado);
			this.tpComplemento.Controls.Add(this.btPublicidad);
			this.tpComplemento.Controls.Add(this.cmbClientePublicidad);
			this.tpComplemento.Controls.Add(this.label30);
			this.tpComplemento.Controls.Add(this.cmbPrecio);
			this.tpComplemento.Controls.Add(this.cmbSector);
			this.tpComplemento.Controls.Add(this.btSector);
			this.tpComplemento.Controls.Add(this.btPrecio);
			this.tpComplemento.Controls.Add(this.label19);
			this.tpComplemento.Controls.Add(this.label2);
			this.tpComplemento.Controls.Add(this.txtCodigo);
			this.tpComplemento.Controls.Add(this.label18);
			this.tpComplemento.Controls.Add(this.txtDiasCredito3);
			this.tpComplemento.Controls.Add(this.label55);
			this.tpComplemento.Controls.Add(this.txtDiasCredito2);
			this.tpComplemento.Controls.Add(this.label54);
			this.tpComplemento.Controls.Add(this.ultraButton2);
			this.tpComplemento.Controls.Add(this.cmbTipoCredito);
			this.tpComplemento.Controls.Add(this.label52);
			this.tpComplemento.Controls.Add(this.chkContribuyenteEsp);
			this.tpComplemento.Controls.Add(this.label34);
			this.tpComplemento.Controls.Add(this.cmbCredTribut);
			this.tpComplemento.Controls.Add(this.label29);
			this.tpComplemento.Controls.Add(this.cmbPorcIvaRet);
			this.tpComplemento.Controls.Add(this.label21);
			this.tpComplemento.Controls.Add(this.cmbVendedor);
			this.tpComplemento.Controls.Add(this.btVendedor);
			this.tpComplemento.Controls.Add(this.spnDescCredito);
			this.tpComplemento.Controls.Add(this.label24);
			this.tpComplemento.Controls.Add(this.spnDescuento);
			this.tpComplemento.Controls.Add(this.label23);
			this.tpComplemento.Controls.Add(this.txtValorCredito);
			this.tpComplemento.Controls.Add(this.txtDiasCredito1);
			this.tpComplemento.Controls.Add(this.label14);
			this.tpComplemento.Controls.Add(this.label12);
			this.tpComplemento.Controls.Add(this.chkPublica);
			this.tpComplemento.Location = new System.Drawing.Point(-10000, -10000);
			this.tpComplemento.Name = "tpComplemento";
			this.tpComplemento.Size = new System.Drawing.Size(938, 422);
			this.tpComplemento.Paint += new System.Windows.Forms.PaintEventHandler(this.tpComplemento_Paint);
			// 
			// txtDiasCredito4
			// 
			appearance26.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtDiasCredito4.Appearance = appearance26;
			this.txtDiasCredito4.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsCliente, "Cliente.DiasCredito3"));
			this.txtDiasCredito4.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtDiasCredito4.Location = new System.Drawing.Point(232, 168);
			this.txtDiasCredito4.MinValue = 0;
			this.txtDiasCredito4.Name = "txtDiasCredito4";
			this.txtDiasCredito4.PromptChar = ' ';
			this.txtDiasCredito4.Size = new System.Drawing.Size(136, 22);
			this.txtDiasCredito4.SpinButtonDisplayStyle = Infragistics.Win.ButtonDisplayStyle.Always;
			this.txtDiasCredito4.TabIndex = 202;
			// 
			// label56
			// 
			this.label56.AutoSize = true;
			this.label56.BackColor = System.Drawing.Color.Transparent;
			this.label56.Location = new System.Drawing.Point(136, 168);
			this.label56.Name = "label56";
			this.label56.Size = new System.Drawing.Size(78, 16);
			this.label56.TabIndex = 203;
			this.label56.Text = "Días Crédito 4:";
			// 
			// cmbInstitucion
			// 
			appearance27.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbInstitucion.Appearance = appearance27;
			this.cmbInstitucion.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbInstitucion.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbInstitucion.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsCliente, "Cliente.idInstitucion"));
			this.cmbInstitucion.DataSource = this.ultraDataSource3;
			ultraGridBand8.AddButtonCaption = "ArticuloGrupo";
			ultraGridColumn16.Header.VisiblePosition = 1;
			ultraGridColumn16.Hidden = true;
			ultraGridColumn17.Header.VisiblePosition = 0;
			ultraGridColumn17.Width = 272;
			ultraGridBand8.Columns.AddRange(new object[] {
																										 ultraGridColumn16,
																										 ultraGridColumn17});
			this.cmbInstitucion.DisplayLayout.BandsSerializer.Add(ultraGridBand8);
			this.cmbInstitucion.DisplayMember = "Descripcion";
			this.cmbInstitucion.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbInstitucion.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbInstitucion.Location = new System.Drawing.Point(232, 360);
			this.cmbInstitucion.Name = "cmbInstitucion";
			this.cmbInstitucion.Size = new System.Drawing.Size(272, 21);
			this.cmbInstitucion.TabIndex = 200;
			this.cmbInstitucion.ValueMember = "idInstitucion";
			// 
			// ultraDataSource3
			// 
			ultraDataColumn8.DataType = typeof(int);
			this.ultraDataSource3.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn8,
																																 ultraDataColumn9});
			// 
			// label57
			// 
			this.label57.AutoSize = true;
			this.label57.BackColor = System.Drawing.Color.Transparent;
			this.label57.Location = new System.Drawing.Point(136, 360);
			this.label57.Name = "label57";
			this.label57.Size = new System.Drawing.Size(57, 16);
			this.label57.TabIndex = 201;
			this.label57.Text = "Institución";
			this.label57.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// cmbArticuloSRI
			// 
			appearance28.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbArticuloSRI.Appearance = appearance28;
			this.cmbArticuloSRI.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbArticuloSRI.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbArticuloSRI.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsCliente, "Cliente.idArticuloSRI"));
			this.cmbArticuloSRI.DataMember = "ArticuloSRI";
			this.cmbArticuloSRI.DataSource = this.cdsArticuloTabla;
			ultraGridBand9.AddButtonCaption = "ArticuloGrupo";
			ultraGridColumn18.Header.VisiblePosition = 0;
			ultraGridColumn18.Hidden = true;
			ultraGridColumn18.Width = 113;
			ultraGridColumn19.Header.VisiblePosition = 1;
			ultraGridColumn19.Width = 248;
			ultraGridColumn20.Header.VisiblePosition = 3;
			ultraGridColumn20.Width = 57;
			ultraGridColumn21.Header.VisiblePosition = 4;
			ultraGridColumn21.Hidden = true;
			ultraGridColumn22.Header.VisiblePosition = 5;
			ultraGridColumn22.Hidden = true;
			ultraGridColumn23.Header.VisiblePosition = 2;
			ultraGridColumn23.Width = 73;
			ultraGridColumn24.Header.VisiblePosition = 6;
			ultraGridBand9.Columns.AddRange(new object[] {
																										 ultraGridColumn18,
																										 ultraGridColumn19,
																										 ultraGridColumn20,
																										 ultraGridColumn21,
																										 ultraGridColumn22,
																										 ultraGridColumn23,
																										 ultraGridColumn24});
			this.cmbArticuloSRI.DisplayLayout.BandsSerializer.Add(ultraGridBand9);
			this.cmbArticuloSRI.DisplayMember = "Detalle";
			this.cmbArticuloSRI.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbArticuloSRI.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbArticuloSRI.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbArticuloSRI.Location = new System.Drawing.Point(232, 392);
			this.cmbArticuloSRI.Name = "cmbArticuloSRI";
			this.cmbArticuloSRI.Size = new System.Drawing.Size(272, 21);
			this.cmbArticuloSRI.TabIndex = 198;
			this.toolTip1.SetToolTip(this.cmbArticuloSRI, "Retenciones de Renta");
			this.cmbArticuloSRI.ValueMember = "idArticuloSRI";
			this.cmbArticuloSRI.Visible = false;
			// 
			// cdsArticuloTabla
			// 
			this.cdsArticuloTabla.BindingContextCtrl = this;
			this.cdsArticuloTabla.CaseSensitive = false;
			this.cdsArticuloTabla.DataLibrary = "LibFacturacion";
			this.cdsArticuloTabla.DataLibraryUrl = "";
			this.cdsArticuloTabla.DataSetDef = "dsArticuloTabla";
			this.cdsArticuloTabla.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsArticuloTabla.Name = "cdsArticuloTabla";
			this.cdsArticuloTabla.SchemaClassName = "LibFacturacion.DataClass";
			this.cdsArticuloTabla.SchemaDef = null;
			// 
			// lblArticuloSRI
			// 
			this.lblArticuloSRI.AutoSize = true;
			this.lblArticuloSRI.BackColor = System.Drawing.Color.Transparent;
			this.lblArticuloSRI.Location = new System.Drawing.Point(136, 392);
			this.lblArticuloSRI.Name = "lblArticuloSRI";
			this.lblArticuloSRI.Size = new System.Drawing.Size(32, 16);
			this.lblArticuloSRI.TabIndex = 199;
			this.lblArticuloSRI.Text = "S.R.I.";
			this.lblArticuloSRI.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.lblArticuloSRI.Visible = false;
			// 
			// btBanco
			// 
			this.btBanco.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.btBanco.Location = new System.Drawing.Point(408, 264);
			this.btBanco.Name = "btBanco";
			this.btBanco.Size = new System.Drawing.Size(24, 24);
			this.btBanco.TabIndex = 197;
			this.btBanco.Text = "...";
			this.btBanco.Visible = false;
			this.btBanco.Click += new System.EventHandler(this.btBanco_Click);
			// 
			// cmbBanco
			// 
			appearance29.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbBanco.Appearance = appearance29;
			this.cmbBanco.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.cmbBanco.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbBanco.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsCliente, "Cliente.idBanco"));
			ultraGridColumn25.Header.VisiblePosition = 0;
			ultraGridColumn26.Header.VisiblePosition = 1;
			ultraGridBand10.Columns.AddRange(new object[] {
																											ultraGridColumn25,
																											ultraGridColumn26});
			this.cmbBanco.DisplayLayout.BandsSerializer.Add(ultraGridBand10);
			this.cmbBanco.DisplayMember = "Nombre";
			this.cmbBanco.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbBanco.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbBanco.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbBanco.Location = new System.Drawing.Point(232, 264);
			this.cmbBanco.Name = "cmbBanco";
			this.cmbBanco.Size = new System.Drawing.Size(168, 21);
			this.cmbBanco.TabIndex = 196;
			this.cmbBanco.ValueMember = "idBanco";
			this.cmbBanco.Visible = false;
			// 
			// txtBeneficiario
			// 
			appearance30.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtBeneficiario.Appearance = appearance30;
			this.txtBeneficiario.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtBeneficiario.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsCliente, "Cliente.Beneficiario"));
			this.txtBeneficiario.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtBeneficiario.Location = new System.Drawing.Point(232, 328);
			this.txtBeneficiario.Name = "txtBeneficiario";
			this.txtBeneficiario.Size = new System.Drawing.Size(304, 22);
			this.txtBeneficiario.TabIndex = 192;
			this.txtBeneficiario.Visible = false;
			// 
			// lblBeneficiario
			// 
			this.lblBeneficiario.AutoSize = true;
			this.lblBeneficiario.BackColor = System.Drawing.Color.Transparent;
			this.lblBeneficiario.Location = new System.Drawing.Point(133, 328);
			this.lblBeneficiario.Name = "lblBeneficiario";
			this.lblBeneficiario.Size = new System.Drawing.Size(62, 16);
			this.lblBeneficiario.TabIndex = 193;
			this.lblBeneficiario.Text = "Beneficiario";
			this.lblBeneficiario.Visible = false;
			// 
			// txtCuentaSwift
			// 
			appearance31.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtCuentaSwift.Appearance = appearance31;
			this.txtCuentaSwift.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsCliente, "Cliente.CuentaSwift"));
			this.txtCuentaSwift.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtCuentaSwift.Location = new System.Drawing.Point(232, 296);
			this.txtCuentaSwift.MaxLength = 20;
			this.txtCuentaSwift.Name = "txtCuentaSwift";
			this.txtCuentaSwift.Size = new System.Drawing.Size(168, 22);
			this.txtCuentaSwift.TabIndex = 191;
			this.txtCuentaSwift.Visible = false;
			this.txtCuentaSwift.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtCuentaSwift_KeyPress);
			this.txtCuentaSwift.ValueChanged += new System.EventHandler(this.txtCuentaSwift_ValueChanged);
			// 
			// lblCuentaSwift
			// 
			this.lblCuentaSwift.AutoSize = true;
			this.lblCuentaSwift.BackColor = System.Drawing.Color.Transparent;
			this.lblCuentaSwift.Location = new System.Drawing.Point(133, 296);
			this.lblCuentaSwift.Name = "lblCuentaSwift";
			this.lblCuentaSwift.Size = new System.Drawing.Size(68, 16);
			this.lblCuentaSwift.TabIndex = 190;
			this.lblCuentaSwift.Text = "Cuenta Swift";
			this.lblCuentaSwift.Visible = false;
			// 
			// lblBanco
			// 
			this.lblBanco.AutoSize = true;
			this.lblBanco.BackColor = System.Drawing.Color.Transparent;
			this.lblBanco.Location = new System.Drawing.Point(133, 264);
			this.lblBanco.Name = "lblBanco";
			this.lblBanco.Size = new System.Drawing.Size(34, 16);
			this.lblBanco.TabIndex = 189;
			this.lblBanco.Text = "Banco";
			this.lblBanco.Visible = false;
			// 
			// dtHoraHasta
			// 
			this.dtHoraHasta.CustomFormat = "HH:mm";
			this.dtHoraHasta.DataBindings.Add(new System.Windows.Forms.Binding("Text", this.cdsCliente, "Cliente.HoraHasta"));
			this.dtHoraHasta.Format = System.Windows.Forms.DateTimePickerFormat.Custom;
			this.dtHoraHasta.Location = new System.Drawing.Point(648, 328);
			this.dtHoraHasta.Name = "dtHoraHasta";
			this.dtHoraHasta.ShowUpDown = true;
			this.dtHoraHasta.Size = new System.Drawing.Size(72, 20);
			this.dtHoraHasta.TabIndex = 186;
			this.dtHoraHasta.Visible = false;
			// 
			// dtHoraDesde
			// 
			this.dtHoraDesde.CustomFormat = "HH:mm";
			this.dtHoraDesde.DataBindings.Add(new System.Windows.Forms.Binding("Text", this.cdsCliente, "Cliente.HoraDesde"));
			this.dtHoraDesde.Format = System.Windows.Forms.DateTimePickerFormat.Custom;
			this.dtHoraDesde.Location = new System.Drawing.Point(456, 328);
			this.dtHoraDesde.Name = "dtHoraDesde";
			this.dtHoraDesde.ShowUpDown = true;
			this.dtHoraDesde.Size = new System.Drawing.Size(72, 20);
			this.dtHoraDesde.TabIndex = 185;
			this.dtHoraDesde.Visible = false;
			// 
			// cmbDias
			// 
			appearance32.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbDias.Appearance = appearance32;
			this.cmbDias.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.cmbDias.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbDias.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsCliente, "Cliente.idDia"));
			this.cmbDias.DisplayMember = "Dia";
			this.cmbDias.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbDias.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbDias.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbDias.Location = new System.Drawing.Point(232, 328);
			this.cmbDias.Name = "cmbDias";
			this.cmbDias.Size = new System.Drawing.Size(160, 21);
			this.cmbDias.TabIndex = 45;
			this.cmbDias.ValueMember = "idDia";
			this.cmbDias.Visible = false;
			// 
			// lblHoraHasta
			// 
			this.lblHoraHasta.AutoSize = true;
			this.lblHoraHasta.BackColor = System.Drawing.Color.Transparent;
			this.lblHoraHasta.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.lblHoraHasta.ForeColor = System.Drawing.Color.Black;
			this.lblHoraHasta.Location = new System.Drawing.Point(592, 328);
			this.lblHoraHasta.Name = "lblHoraHasta";
			this.lblHoraHasta.Size = new System.Drawing.Size(33, 17);
			this.lblHoraHasta.TabIndex = 180;
			this.lblHoraHasta.Text = "Hasta";
			this.lblHoraHasta.Visible = false;
			// 
			// lblHoraDesde
			// 
			this.lblHoraDesde.AutoSize = true;
			this.lblHoraDesde.BackColor = System.Drawing.Color.Transparent;
			this.lblHoraDesde.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.lblHoraDesde.ForeColor = System.Drawing.Color.Black;
			this.lblHoraDesde.Location = new System.Drawing.Point(424, 328);
			this.lblHoraDesde.Name = "lblHoraDesde";
			this.lblHoraDesde.Size = new System.Drawing.Size(35, 17);
			this.lblHoraDesde.TabIndex = 179;
			this.lblHoraDesde.Text = "Desde";
			this.lblHoraDesde.Visible = false;
			// 
			// lblDias
			// 
			this.lblDias.AutoSize = true;
			this.lblDias.BackColor = System.Drawing.Color.Transparent;
			this.lblDias.Location = new System.Drawing.Point(133, 328);
			this.lblDias.Name = "lblDias";
			this.lblDias.Size = new System.Drawing.Size(20, 16);
			this.lblDias.TabIndex = 178;
			this.lblDias.Text = "Día";
			this.lblDias.Visible = false;
			// 
			// chkHaceRetencion
			// 
			appearance33.ForeColorDisabled = System.Drawing.Color.Black;
			this.chkHaceRetencion.Appearance = appearance33;
			this.chkHaceRetencion.BackColor = System.Drawing.Color.Transparent;
			this.chkHaceRetencion.DataBindings.Add(new System.Windows.Forms.Binding("CheckedValue", this.cdsCliente, "Cliente.HaceRetencion"));
			this.chkHaceRetencion.Location = new System.Drawing.Point(619, 200);
			this.chkHaceRetencion.Name = "chkHaceRetencion";
			this.chkHaceRetencion.Size = new System.Drawing.Size(128, 21);
			this.chkHaceRetencion.TabIndex = 34;
			this.chkHaceRetencion.Text = "Hace Retención";
			// 
			// rdbCredito
			// 
			this.rdbCredito.Enabled = false;
			this.rdbCredito.Location = new System.Drawing.Point(261, 8);
			this.rdbCredito.Name = "rdbCredito";
			this.rdbCredito.TabIndex = 22;
			this.rdbCredito.Text = "Crédito";
			this.rdbCredito.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.rdbCredito_KeyPress);
			this.rdbCredito.CheckedChanged += new System.EventHandler(this.rdbCredito_CheckedChanged);
			// 
			// rdbContado
			// 
			this.rdbContado.Enabled = false;
			this.rdbContado.Location = new System.Drawing.Point(133, 8);
			this.rdbContado.Name = "rdbContado";
			this.rdbContado.TabIndex = 21;
			this.rdbContado.Text = "Contado";
			this.rdbContado.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.rdbContado_KeyPress);
			this.rdbContado.CheckedChanged += new System.EventHandler(this.rdbContado_CheckedChanged);
			// 
			// btPublicidad
			// 
			this.btPublicidad.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.btPublicidad.Location = new System.Drawing.Point(757, 264);
			this.btPublicidad.Name = "btPublicidad";
			this.btPublicidad.Size = new System.Drawing.Size(24, 21);
			this.btPublicidad.TabIndex = 42;
			this.btPublicidad.Text = "...";
			// 
			// cmbClientePublicidad
			// 
			appearance34.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbClientePublicidad.Appearance = appearance34;
			this.cmbClientePublicidad.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.cmbClientePublicidad.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbClientePublicidad.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsCliente, "Cliente.idClientePublicidad"));
			ultraGridColumn27.Header.VisiblePosition = 0;
			ultraGridColumn27.Hidden = true;
			ultraGridColumn28.Header.VisiblePosition = 1;
			ultraGridColumn28.Width = 310;
			ultraGridBand11.Columns.AddRange(new object[] {
																											ultraGridColumn27,
																											ultraGridColumn28});
			this.cmbClientePublicidad.DisplayLayout.BandsSerializer.Add(ultraGridBand11);
			this.cmbClientePublicidad.DisplayMember = "Medio";
			this.cmbClientePublicidad.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbClientePublicidad.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbClientePublicidad.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbClientePublicidad.Location = new System.Drawing.Point(613, 264);
			this.cmbClientePublicidad.Name = "cmbClientePublicidad";
			this.cmbClientePublicidad.Size = new System.Drawing.Size(136, 21);
			this.cmbClientePublicidad.TabIndex = 41;
			this.cmbClientePublicidad.ValueMember = "idClientePublicidad";
			// 
			// label30
			// 
			this.label30.AutoSize = true;
			this.label30.BackColor = System.Drawing.Color.Transparent;
			this.label30.Location = new System.Drawing.Point(517, 264);
			this.label30.Name = "label30";
			this.label30.Size = new System.Drawing.Size(54, 16);
			this.label30.TabIndex = 175;
			this.label30.Text = "Publicidad";
			// 
			// cmbPrecio
			// 
			appearance35.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbPrecio.Appearance = appearance35;
			this.cmbPrecio.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.cmbPrecio.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbPrecio.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsCliente, "Cliente.idTipoPrecio"));
			ultraGridColumn29.Header.VisiblePosition = 0;
			ultraGridColumn29.Hidden = true;
			ultraGridColumn30.Header.VisiblePosition = 1;
			ultraGridColumn30.Width = 291;
			ultraGridBand12.Columns.AddRange(new object[] {
																											ultraGridColumn29,
																											ultraGridColumn30});
			this.cmbPrecio.DisplayLayout.BandsSerializer.Add(ultraGridBand12);
			this.cmbPrecio.DisplayMember = "Precio";
			this.cmbPrecio.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbPrecio.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbPrecio.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbPrecio.Location = new System.Drawing.Point(613, 296);
			this.cmbPrecio.Name = "cmbPrecio";
			this.cmbPrecio.Size = new System.Drawing.Size(136, 21);
			this.cmbPrecio.TabIndex = 43;
			this.cmbPrecio.ValueMember = "idTipoPrecio";
			// 
			// cmbSector
			// 
			appearance36.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbSector.Appearance = appearance36;
			this.cmbSector.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbSector.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbSector.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsCliente, "Cliente.idSector"));
			ultraGridColumn31.Header.VisiblePosition = 0;
			ultraGridColumn31.Hidden = true;
			ultraGridColumn32.Header.VisiblePosition = 1;
			ultraGridColumn32.Width = 274;
			ultraGridBand13.Columns.AddRange(new object[] {
																											ultraGridColumn31,
																											ultraGridColumn32});
			this.cmbSector.DisplayLayout.BandsSerializer.Add(ultraGridBand13);
			this.cmbSector.DisplayMember = "Sector";
			this.cmbSector.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbSector.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbSector.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbSector.Location = new System.Drawing.Point(613, 104);
			this.cmbSector.Name = "cmbSector";
			this.cmbSector.Size = new System.Drawing.Size(160, 21);
			this.cmbSector.TabIndex = 29;
			this.cmbSector.ValueMember = "idSector";
			// 
			// btSector
			// 
			this.btSector.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.btSector.Location = new System.Drawing.Point(781, 104);
			this.btSector.Name = "btSector";
			this.btSector.Size = new System.Drawing.Size(24, 21);
			this.btSector.TabIndex = 30;
			this.btSector.Text = "...";
			// 
			// btPrecio
			// 
			this.btPrecio.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.btPrecio.Location = new System.Drawing.Point(757, 296);
			this.btPrecio.Name = "btPrecio";
			this.btPrecio.Size = new System.Drawing.Size(24, 21);
			this.btPrecio.TabIndex = 44;
			this.btPrecio.Text = "...";
			// 
			// label19
			// 
			this.label19.AutoSize = true;
			this.label19.BackColor = System.Drawing.Color.Transparent;
			this.label19.Location = new System.Drawing.Point(517, 104);
			this.label19.Name = "label19";
			this.label19.Size = new System.Drawing.Size(36, 16);
			this.label19.TabIndex = 174;
			this.label19.Text = "Sector";
			// 
			// label2
			// 
			this.label2.AutoSize = true;
			this.label2.BackColor = System.Drawing.Color.Transparent;
			this.label2.Location = new System.Drawing.Point(517, 296);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(38, 16);
			this.label2.TabIndex = 173;
			this.label2.Text = "Precio:";
			// 
			// txtCodigo
			// 
			appearance37.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtCodigo.Appearance = appearance37;
			this.txtCodigo.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsCliente, "Cliente.Codigo"));
			this.txtCodigo.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtCodigo.Enabled = false;
			this.txtCodigo.Location = new System.Drawing.Point(613, 136);
			this.txtCodigo.Name = "txtCodigo";
			this.txtCodigo.Size = new System.Drawing.Size(160, 22);
			this.txtCodigo.TabIndex = 31;
			// 
			// label18
			// 
			this.label18.AutoSize = true;
			this.label18.BackColor = System.Drawing.Color.Transparent;
			this.label18.Location = new System.Drawing.Point(517, 136);
			this.label18.Name = "label18";
			this.label18.Size = new System.Drawing.Size(42, 16);
			this.label18.TabIndex = 172;
			this.label18.Text = "Código:";
			// 
			// txtDiasCredito3
			// 
			appearance38.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtDiasCredito3.Appearance = appearance38;
			this.txtDiasCredito3.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsCliente, "Cliente.DiasCredito2"));
			this.txtDiasCredito3.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtDiasCredito3.Location = new System.Drawing.Point(232, 136);
			this.txtDiasCredito3.MinValue = 0;
			this.txtDiasCredito3.Name = "txtDiasCredito3";
			this.txtDiasCredito3.PromptChar = ' ';
			this.txtDiasCredito3.Size = new System.Drawing.Size(136, 22);
			this.txtDiasCredito3.SpinButtonDisplayStyle = Infragistics.Win.ButtonDisplayStyle.Always;
			this.txtDiasCredito3.TabIndex = 26;
			this.txtDiasCredito3.Click += new System.EventHandler(this.txtDiasCredito3_Click);
			this.txtDiasCredito3.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtDiasCredito3_KeyPress);
			this.txtDiasCredito3.Enter += new System.EventHandler(this.txtDiasCredito3_Enter);
			// 
			// label55
			// 
			this.label55.AutoSize = true;
			this.label55.BackColor = System.Drawing.Color.Transparent;
			this.label55.Location = new System.Drawing.Point(133, 136);
			this.label55.Name = "label55";
			this.label55.Size = new System.Drawing.Size(78, 16);
			this.label55.TabIndex = 164;
			this.label55.Text = "Días Crédito 3:";
			// 
			// txtDiasCredito2
			// 
			appearance39.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtDiasCredito2.Appearance = appearance39;
			this.txtDiasCredito2.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsCliente, "Cliente.DiasCredito1"));
			this.txtDiasCredito2.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtDiasCredito2.Location = new System.Drawing.Point(232, 104);
			this.txtDiasCredito2.MinValue = 0;
			this.txtDiasCredito2.Name = "txtDiasCredito2";
			this.txtDiasCredito2.PromptChar = ' ';
			this.txtDiasCredito2.Size = new System.Drawing.Size(136, 22);
			this.txtDiasCredito2.SpinButtonDisplayStyle = Infragistics.Win.ButtonDisplayStyle.Always;
			this.txtDiasCredito2.TabIndex = 25;
			this.txtDiasCredito2.Click += new System.EventHandler(this.txtDiasCredito2_Click);
			this.txtDiasCredito2.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtDiasCredito2_KeyPress);
			this.txtDiasCredito2.Enter += new System.EventHandler(this.txtDiasCredito2_Enter);
			// 
			// label54
			// 
			this.label54.AutoSize = true;
			this.label54.BackColor = System.Drawing.Color.Transparent;
			this.label54.Location = new System.Drawing.Point(133, 104);
			this.label54.Name = "label54";
			this.label54.Size = new System.Drawing.Size(78, 16);
			this.label54.TabIndex = 162;
			this.label54.Text = "Días Crédito 2:";
			// 
			// ultraButton2
			// 
			this.ultraButton2.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.ultraButton2.Location = new System.Drawing.Point(408, 296);
			this.ultraButton2.Name = "ultraButton2";
			this.ultraButton2.Size = new System.Drawing.Size(24, 21);
			this.ultraButton2.TabIndex = 39;
			this.ultraButton2.Text = "...";
			this.ultraButton2.Click += new System.EventHandler(this.ultraButton2_Click);
			// 
			// cmbTipoCredito
			// 
			appearance40.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbTipoCredito.Appearance = appearance40;
			this.cmbTipoCredito.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.cmbTipoCredito.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbTipoCredito.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsCliente, "Cliente.idTipoCredito"));
			this.cmbTipoCredito.DataMember = "TIPOCRED";
			this.cmbTipoCredito.DataSource = this.cdsContacto;
			ultraGridColumn33.Header.VisiblePosition = 0;
			ultraGridColumn33.Hidden = true;
			ultraGridColumn34.Header.VisiblePosition = 1;
			ultraGridColumn34.Width = 50;
			ultraGridColumn35.Header.VisiblePosition = 2;
			ultraGridColumn35.Width = 150;
			ultraGridColumn36.Header.VisiblePosition = 3;
			ultraGridColumn36.Hidden = true;
			ultraGridColumn37.Header.VisiblePosition = 4;
			ultraGridColumn37.Hidden = true;
			ultraGridColumn38.Header.VisiblePosition = 5;
			ultraGridColumn38.Hidden = true;
			ultraGridColumn39.Header.VisiblePosition = 6;
			ultraGridColumn39.Hidden = true;
			ultraGridColumn40.Header.VisiblePosition = 7;
			ultraGridColumn40.Hidden = true;
			ultraGridBand14.Columns.AddRange(new object[] {
																											ultraGridColumn33,
																											ultraGridColumn34,
																											ultraGridColumn35,
																											ultraGridColumn36,
																											ultraGridColumn37,
																											ultraGridColumn38,
																											ultraGridColumn39,
																											ultraGridColumn40});
			this.cmbTipoCredito.DisplayLayout.BandsSerializer.Add(ultraGridBand14);
			this.cmbTipoCredito.DisplayMember = "Nombre";
			this.cmbTipoCredito.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbTipoCredito.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbTipoCredito.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbTipoCredito.Location = new System.Drawing.Point(232, 296);
			this.cmbTipoCredito.Name = "cmbTipoCredito";
			this.cmbTipoCredito.Size = new System.Drawing.Size(160, 21);
			this.cmbTipoCredito.TabIndex = 38;
			this.cmbTipoCredito.ValueMember = "idTipoCredito";
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
			// label52
			// 
			this.label52.AutoSize = true;
			this.label52.BackColor = System.Drawing.Color.Transparent;
			this.label52.Location = new System.Drawing.Point(133, 296);
			this.label52.Name = "label52";
			this.label52.Size = new System.Drawing.Size(64, 16);
			this.label52.TabIndex = 155;
			this.label52.Text = "Forma Pago";
			// 
			// chkContribuyenteEsp
			// 
			appearance41.ForeColorDisabled = System.Drawing.Color.Black;
			this.chkContribuyenteEsp.Appearance = appearance41;
			this.chkContribuyenteEsp.BackColor = System.Drawing.Color.Transparent;
			this.chkContribuyenteEsp.DataBindings.Add(new System.Windows.Forms.Binding("CheckedValue", this.cdsCliente, "Cliente.ContribEspecial"));
			this.chkContribuyenteEsp.Location = new System.Drawing.Point(371, 200);
			this.chkContribuyenteEsp.Name = "chkContribuyenteEsp";
			this.chkContribuyenteEsp.Size = new System.Drawing.Size(152, 22);
			this.chkContribuyenteEsp.TabIndex = 33;
			this.chkContribuyenteEsp.Text = "Contribuyente Especial";
			this.chkContribuyenteEsp.Validated += new System.EventHandler(this.chkContribuyenteEsp_Validated);
			this.chkContribuyenteEsp.CheckedChanged += new System.EventHandler(this.chkContribuyenteEsp_CheckedChanged);
			// 
			// label34
			// 
			this.label34.AutoSize = true;
			this.label34.BackColor = System.Drawing.Color.Transparent;
			this.label34.Location = new System.Drawing.Point(133, 232);
			this.label34.Name = "label34";
			this.label34.Size = new System.Drawing.Size(92, 16);
			this.label34.TabIndex = 139;
			this.label34.Text = "Crédito Tributario";
			// 
			// cmbCredTribut
			// 
			appearance42.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbCredTribut.Appearance = appearance42;
			this.cmbCredTribut.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.cmbCredTribut.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbCredTribut.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsCliente, "Cliente.idCredTribut"));
			ultraGridColumn41.Header.VisiblePosition = 0;
			ultraGridColumn41.Hidden = true;
			ultraGridColumn42.Header.VisiblePosition = 1;
			ultraGridColumn42.Width = 295;
			ultraGridBand15.Columns.AddRange(new object[] {
																											ultraGridColumn41,
																											ultraGridColumn42});
			this.cmbCredTribut.DisplayLayout.BandsSerializer.Add(ultraGridBand15);
			this.cmbCredTribut.DisplayMember = "CredTributario";
			this.cmbCredTribut.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbCredTribut.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbCredTribut.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbCredTribut.Location = new System.Drawing.Point(232, 232);
			this.cmbCredTribut.Name = "cmbCredTribut";
			this.cmbCredTribut.Size = new System.Drawing.Size(160, 21);
			this.cmbCredTribut.TabIndex = 35;
			this.cmbCredTribut.ValueMember = "idCredTributario";
			// 
			// label29
			// 
			this.label29.AutoSize = true;
			this.label29.BackColor = System.Drawing.Color.Transparent;
			this.label29.Location = new System.Drawing.Point(517, 232);
			this.label29.Name = "label29";
			this.label29.Size = new System.Drawing.Size(87, 16);
			this.label29.TabIndex = 129;
			this.label29.Text = "% Retención Iva";
			this.label29.Visible = false;
			// 
			// cmbPorcIvaRet
			// 
			appearance43.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbPorcIvaRet.Appearance = appearance43;
			this.cmbPorcIvaRet.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.cmbPorcIvaRet.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbPorcIvaRet.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsCliente, "Cliente.idRetIva"));
			ultraGridColumn43.Header.VisiblePosition = 0;
			ultraGridColumn43.Hidden = true;
			ultraGridColumn44.Header.VisiblePosition = 1;
			ultraGridColumn45.Header.VisiblePosition = 2;
			ultraGridBand16.Columns.AddRange(new object[] {
																											ultraGridColumn43,
																											ultraGridColumn44,
																											ultraGridColumn45});
			this.cmbPorcIvaRet.DisplayLayout.BandsSerializer.Add(ultraGridBand16);
			this.cmbPorcIvaRet.DisplayMember = "Porcentaje";
			this.cmbPorcIvaRet.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbPorcIvaRet.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbPorcIvaRet.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbPorcIvaRet.Location = new System.Drawing.Point(613, 232);
			this.cmbPorcIvaRet.Name = "cmbPorcIvaRet";
			this.cmbPorcIvaRet.Size = new System.Drawing.Size(171, 21);
			this.cmbPorcIvaRet.TabIndex = 40;
			this.cmbPorcIvaRet.ValueMember = "idRetIva";
			this.cmbPorcIvaRet.Visible = false;
			// 
			// label21
			// 
			this.label21.AutoSize = true;
			this.label21.BackColor = System.Drawing.Color.Transparent;
			this.label21.Location = new System.Drawing.Point(133, 264);
			this.label21.Name = "label21";
			this.label21.Size = new System.Drawing.Size(52, 16);
			this.label21.TabIndex = 122;
			this.label21.Text = "Vendedor";
			// 
			// cmbVendedor
			// 
			appearance44.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbVendedor.Appearance = appearance44;
			this.cmbVendedor.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.cmbVendedor.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbVendedor.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsCliente, "Cliente.idVendedor"));
			ultraGridColumn46.Header.VisiblePosition = 0;
			ultraGridColumn46.Hidden = true;
			ultraGridColumn47.Header.VisiblePosition = 1;
			ultraGridColumn47.Width = 350;
			ultraGridBand17.Columns.AddRange(new object[] {
																											ultraGridColumn46,
																											ultraGridColumn47});
			this.cmbVendedor.DisplayLayout.BandsSerializer.Add(ultraGridBand17);
			this.cmbVendedor.DisplayMember = "Nombre";
			this.cmbVendedor.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbVendedor.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbVendedor.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbVendedor.Location = new System.Drawing.Point(232, 264);
			this.cmbVendedor.Name = "cmbVendedor";
			this.cmbVendedor.Size = new System.Drawing.Size(160, 21);
			this.cmbVendedor.TabIndex = 36;
			this.cmbVendedor.ValueMember = "idPersonal";
			// 
			// btVendedor
			// 
			this.btVendedor.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.btVendedor.Location = new System.Drawing.Point(408, 264);
			this.btVendedor.Name = "btVendedor";
			this.btVendedor.Size = new System.Drawing.Size(24, 21);
			this.btVendedor.TabIndex = 37;
			this.btVendedor.Text = "...";
			this.btVendedor.Click += new System.EventHandler(this.btVendedor_Click);
			// 
			// spnDescCredito
			// 
			appearance45.ForeColorDisabled = System.Drawing.Color.Black;
			this.spnDescCredito.Appearance = appearance45;
			this.spnDescCredito.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsCliente, "Cliente.DescuentoCredito"));
			this.spnDescCredito.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.spnDescCredito.Location = new System.Drawing.Point(613, 72);
			this.spnDescCredito.MinValue = 0;
			this.spnDescCredito.Name = "spnDescCredito";
			this.spnDescCredito.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.spnDescCredito.PromptChar = ' ';
			this.spnDescCredito.Size = new System.Drawing.Size(136, 22);
			this.spnDescCredito.SpinButtonDisplayStyle = Infragistics.Win.ButtonDisplayStyle.Always;
			this.spnDescCredito.TabIndex = 28;
			this.spnDescCredito.Click += new System.EventHandler(this.spnDescCredito_Click);
			this.spnDescCredito.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.spnDescCredito_KeyPress);
			this.spnDescCredito.Enter += new System.EventHandler(this.spnDescCredito_Enter);
			// 
			// label24
			// 
			this.label24.AutoSize = true;
			this.label24.BackColor = System.Drawing.Color.Transparent;
			this.label24.Location = new System.Drawing.Point(517, 72);
			this.label24.Name = "label24";
			this.label24.Size = new System.Drawing.Size(71, 16);
			this.label24.TabIndex = 113;
			this.label24.Text = "Desc. Crédito";
			// 
			// spnDescuento
			// 
			appearance46.ForeColorDisabled = System.Drawing.Color.Black;
			this.spnDescuento.Appearance = appearance46;
			this.spnDescuento.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsCliente, "Cliente.DescuentoContado"));
			this.spnDescuento.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.spnDescuento.Location = new System.Drawing.Point(613, 40);
			this.spnDescuento.MinValue = 0;
			this.spnDescuento.Name = "spnDescuento";
			this.spnDescuento.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.spnDescuento.PromptChar = ' ';
			this.spnDescuento.Size = new System.Drawing.Size(136, 22);
			this.spnDescuento.SpinButtonDisplayStyle = Infragistics.Win.ButtonDisplayStyle.Always;
			this.spnDescuento.TabIndex = 27;
			this.spnDescuento.Click += new System.EventHandler(this.spnDescuento_Click);
			this.spnDescuento.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.spnDescuento_KeyPress);
			this.spnDescuento.Enter += new System.EventHandler(this.spnDescuento_Enter);
			// 
			// label23
			// 
			this.label23.AutoSize = true;
			this.label23.BackColor = System.Drawing.Color.Transparent;
			this.label23.Location = new System.Drawing.Point(517, 40);
			this.label23.Name = "label23";
			this.label23.Size = new System.Drawing.Size(76, 16);
			this.label23.TabIndex = 111;
			this.label23.Text = "Desc. Contado";
			// 
			// txtValorCredito
			// 
			appearance47.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtValorCredito.Appearance = appearance47;
			this.txtValorCredito.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsCliente, "Cliente.MaximoCredito"));
			this.txtValorCredito.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtValorCredito.FormatString = "#,##0.00";
			this.txtValorCredito.Location = new System.Drawing.Point(232, 40);
			this.txtValorCredito.MinValue = 0;
			this.txtValorCredito.Name = "txtValorCredito";
			this.txtValorCredito.PromptChar = ' ';
			this.txtValorCredito.Size = new System.Drawing.Size(136, 22);
			this.txtValorCredito.SpinButtonDisplayStyle = Infragistics.Win.ButtonDisplayStyle.Always;
			this.txtValorCredito.TabIndex = 23;
			this.txtValorCredito.Click += new System.EventHandler(this.txtValorCredito_Click);
			this.txtValorCredito.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtValorCredito_KeyPress);
			this.txtValorCredito.ValueChanged += new System.EventHandler(this.txtValorCredito_ValueChanged);
			this.txtValorCredito.Enter += new System.EventHandler(this.txtValorCredito_Enter);
			// 
			// txtDiasCredito1
			// 
			appearance48.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtDiasCredito1.Appearance = appearance48;
			this.txtDiasCredito1.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsCliente, "Cliente.DiasCredito"));
			this.txtDiasCredito1.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtDiasCredito1.Location = new System.Drawing.Point(232, 72);
			this.txtDiasCredito1.MinValue = 0;
			this.txtDiasCredito1.Name = "txtDiasCredito1";
			this.txtDiasCredito1.PromptChar = ' ';
			this.txtDiasCredito1.Size = new System.Drawing.Size(136, 22);
			this.txtDiasCredito1.SpinButtonDisplayStyle = Infragistics.Win.ButtonDisplayStyle.Always;
			this.txtDiasCredito1.TabIndex = 24;
			this.txtDiasCredito1.Click += new System.EventHandler(this.txtDiasCredito1_Click);
			this.txtDiasCredito1.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtDiasCredito1_KeyPress);
			this.txtDiasCredito1.Validated += new System.EventHandler(this.ultraNumericEditor1_Validated);
			this.txtDiasCredito1.ValueChanged += new System.EventHandler(this.ultraNumericEditor1_ValueChanged);
			this.txtDiasCredito1.Enter += new System.EventHandler(this.txtDiasCredito1_Enter);
			// 
			// label14
			// 
			this.label14.AutoSize = true;
			this.label14.BackColor = System.Drawing.Color.Transparent;
			this.label14.Location = new System.Drawing.Point(133, 40);
			this.label14.Name = "label14";
			this.label14.Size = new System.Drawing.Size(89, 16);
			this.label14.TabIndex = 109;
			this.label14.Text = "Cupo de Crédito:";
			// 
			// label12
			// 
			this.label12.AutoSize = true;
			this.label12.BackColor = System.Drawing.Color.Transparent;
			this.label12.Location = new System.Drawing.Point(133, 72);
			this.label12.Name = "label12";
			this.label12.Size = new System.Drawing.Size(78, 16);
			this.label12.TabIndex = 108;
			this.label12.Text = "Días Crédito 1:";
			// 
			// chkPublica
			// 
			appearance49.ForeColorDisabled = System.Drawing.Color.Black;
			this.chkPublica.Appearance = appearance49;
			this.chkPublica.BackColor = System.Drawing.Color.Transparent;
			this.chkPublica.DataBindings.Add(new System.Windows.Forms.Binding("CheckedValue", this.cdsCliente, "Cliente.InstPublica"));
			this.chkPublica.Location = new System.Drawing.Point(147, 200);
			this.chkPublica.Name = "chkPublica";
			this.chkPublica.Size = new System.Drawing.Size(128, 21);
			this.chkPublica.TabIndex = 32;
			this.chkPublica.Text = "Institución Pública";
			this.chkPublica.Validated += new System.EventHandler(this.chkPublica_Validated);
			// 
			// ultraTabPageControl5
			// 
			this.ultraTabPageControl5.Controls.Add(this.ultraGrid2);
			this.ultraTabPageControl5.Location = new System.Drawing.Point(-10000, -10000);
			this.ultraTabPageControl5.Name = "ultraTabPageControl5";
			this.ultraTabPageControl5.Size = new System.Drawing.Size(938, 422);
			// 
			// ultraGrid2
			// 
			this.ultraGrid2.Cursor = System.Windows.Forms.Cursors.Default;
			this.ultraGrid2.DataMember = "ClienteContacto";
			this.ultraGrid2.DataSource = this.cdsContacto;
			appearance50.BackColor = System.Drawing.Color.White;
			this.ultraGrid2.DisplayLayout.Appearance = appearance50;
			this.ultraGrid2.DisplayLayout.AutoFitColumns = true;
			ultraGridBand18.AddButtonCaption = "ClienteTipoRuc";
			ultraGridColumn48.Header.VisiblePosition = 0;
			ultraGridColumn48.Hidden = true;
			ultraGridColumn49.Header.VisiblePosition = 1;
			ultraGridColumn49.Hidden = true;
			ultraGridColumn50.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn50.Header.VisiblePosition = 2;
			ultraGridColumn50.Width = 154;
			ultraGridColumn51.FieldLen = 9;
			ultraGridColumn51.Format = "";
			ultraGridColumn51.Header.VisiblePosition = 4;
			ultraGridColumn51.Width = 118;
			ultraGridColumn52.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn52.Header.VisiblePosition = 6;
			ultraGridColumn52.Width = 154;
			ultraGridColumn53.Header.Caption = "E-Mail";
			ultraGridColumn53.Header.VisiblePosition = 8;
			ultraGridColumn53.Width = 89;
			ultraGridColumn54.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn54.Header.VisiblePosition = 9;
			ultraGridColumn54.Hidden = true;
			ultraGridColumn54.Width = 94;
			ultraGridColumn55.Header.VisiblePosition = 11;
			ultraGridColumn55.Hidden = true;
			ultraGridColumn55.Width = 72;
			ultraGridColumn56.FieldLen = 10;
			ultraGridColumn56.Format = "";
			ultraGridColumn56.Header.VisiblePosition = 5;
			ultraGridColumn56.Width = 130;
			ultraGridColumn57.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn57.Header.VisiblePosition = 3;
			ultraGridColumn57.Width = 141;
			ultraGridColumn58.Header.Caption = "Envía Mail";
			ultraGridColumn58.Header.VisiblePosition = 7;
			ultraGridColumn58.Width = 75;
			ultraGridColumn59.Header.VisiblePosition = 10;
			ultraGridColumn59.Hidden = true;
			ultraGridColumn59.Width = 78;
			ultraGridColumn60.Header.VisiblePosition = 12;
			ultraGridBand18.Columns.AddRange(new object[] {
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
																											ultraGridColumn60});
			this.ultraGrid2.DisplayLayout.BandsSerializer.Add(ultraGridBand18);
			this.ultraGrid2.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.TemplateOnBottom;
			this.ultraGrid2.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.True;
			this.ultraGrid2.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance51.BackColor = System.Drawing.Color.Transparent;
			this.ultraGrid2.DisplayLayout.Override.CardAreaAppearance = appearance51;
			appearance52.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance52.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance52.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance52.FontData.BoldAsString = "True";
			appearance52.FontData.Name = "Arial";
			appearance52.FontData.SizeInPoints = 10F;
			appearance52.ForeColor = System.Drawing.Color.White;
			appearance52.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.ultraGrid2.DisplayLayout.Override.HeaderAppearance = appearance52;
			appearance53.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance53.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance53.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid2.DisplayLayout.Override.RowSelectorAppearance = appearance53;
			appearance54.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance54.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance54.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid2.DisplayLayout.Override.SelectedRowAppearance = appearance54;
			this.ultraGrid2.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.ultraGrid2.Location = new System.Drawing.Point(9, 6);
			this.ultraGrid2.Name = "ultraGrid2";
			this.ultraGrid2.Size = new System.Drawing.Size(920, 410);
			this.ultraGrid2.TabIndex = 1;
			this.ultraGrid2.AfterRowInsert += new Infragistics.Win.UltraWinGrid.RowEventHandler(this.ultraGrid2_AfterRowInsert);
			this.ultraGrid2.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.ultraGrid2_KeyPress);
			this.ultraGrid2.KeyDown += new System.Windows.Forms.KeyEventHandler(this.ultraGrid2_KeyDown);
			this.ultraGrid2.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.ultraGrid2_InitializeLayout);
			// 
			// ultraTabPageControl2
			// 
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
			this.ultraTabPageControl2.Size = new System.Drawing.Size(938, 422);
			// 
			// label53
			// 
			this.label53.AutoSize = true;
			this.label53.BackColor = System.Drawing.Color.Transparent;
			this.label53.Location = new System.Drawing.Point(517, 353);
			this.label53.Name = "label53";
			this.label53.Size = new System.Drawing.Size(34, 16);
			this.label53.TabIndex = 150;
			this.label53.Text = "Banco";
			this.label53.Visible = false;
			// 
			// ultraNumericEditor3
			// 
			this.ultraNumericEditor3.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsCliente, "Cliente.ProveMaster"));
			this.ultraNumericEditor3.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.ultraNumericEditor3.FormatString = "";
			this.ultraNumericEditor3.Location = new System.Drawing.Point(405, 353);
			this.ultraNumericEditor3.Name = "ultraNumericEditor3";
			this.ultraNumericEditor3.PromptChar = ' ';
			this.ultraNumericEditor3.Size = new System.Drawing.Size(80, 22);
			this.ultraNumericEditor3.TabIndex = 149;
			this.ultraNumericEditor3.Visible = false;
			// 
			// label49
			// 
			this.label49.AutoSize = true;
			this.label49.BackColor = System.Drawing.Color.Transparent;
			this.label49.Location = new System.Drawing.Point(333, 353);
			this.label49.Name = "label49";
			this.label49.Size = new System.Drawing.Size(63, 16);
			this.label49.TabIndex = 148;
			this.label49.Text = "Prov Master";
			this.label49.Visible = false;
			// 
			// ultraNumericEditor2
			// 
			this.ultraNumericEditor2.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsCliente, "Cliente.PorComT"));
			this.ultraNumericEditor2.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.ultraNumericEditor2.FormatString = "#,##0.00";
			this.ultraNumericEditor2.Location = new System.Drawing.Point(229, 353);
			this.ultraNumericEditor2.Name = "ultraNumericEditor2";
			this.ultraNumericEditor2.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.ultraNumericEditor2.PromptChar = ' ';
			this.ultraNumericEditor2.Size = new System.Drawing.Size(80, 22);
			this.ultraNumericEditor2.TabIndex = 147;
			this.ultraNumericEditor2.Visible = false;
			// 
			// ultraCheckEditor3
			// 
			this.ultraCheckEditor3.DataBindings.Add(new System.Windows.Forms.Binding("CheckedValue", this.cdsCliente, "Cliente.NetoValorT"));
			this.ultraCheckEditor3.Location = new System.Drawing.Point(693, 328);
			this.ultraCheckEditor3.Name = "ultraCheckEditor3";
			this.ultraCheckEditor3.Size = new System.Drawing.Size(80, 25);
			this.ultraCheckEditor3.TabIndex = 146;
			this.ultraCheckEditor3.Text = "Valor Neto";
			this.ultraCheckEditor3.Visible = false;
			// 
			// label48
			// 
			this.label48.AutoSize = true;
			this.label48.BackColor = System.Drawing.Color.Transparent;
			this.label48.Location = new System.Drawing.Point(157, 353);
			this.label48.Name = "label48";
			this.label48.Size = new System.Drawing.Size(47, 16);
			this.label48.TabIndex = 145;
			this.label48.Text = "Por Com";
			this.label48.Visible = false;
			// 
			// chkEntrega
			// 
			appearance55.ForeColorDisabled = System.Drawing.Color.Black;
			this.chkEntrega.Appearance = appearance55;
			this.chkEntrega.BackColor = System.Drawing.Color.Transparent;
			this.chkEntrega.DataBindings.Add(new System.Windows.Forms.Binding("CheckedValue", this.cdsCliente, "Cliente.Entrega"));
			this.chkEntrega.Location = new System.Drawing.Point(517, 328);
			this.chkEntrega.Name = "chkEntrega";
			this.chkEntrega.Size = new System.Drawing.Size(120, 21);
			this.chkEntrega.TabIndex = 143;
			this.chkEntrega.Text = "Recibe Mercaderia";
			this.chkEntrega.Visible = false;
			// 
			// ultraCheckEditor2
			// 
			appearance56.ForeColorDisabled = System.Drawing.Color.Black;
			this.ultraCheckEditor2.Appearance = appearance56;
			this.ultraCheckEditor2.BackColor = System.Drawing.Color.Transparent;
			this.ultraCheckEditor2.DataBindings.Add(new System.Windows.Forms.Binding("CheckedValue", this.cdsCliente, "Cliente.DesconCliente"));
			this.ultraCheckEditor2.Location = new System.Drawing.Point(333, 328);
			this.ultraCheckEditor2.Name = "ultraCheckEditor2";
			this.ultraCheckEditor2.Size = new System.Drawing.Size(168, 21);
			this.ultraCheckEditor2.TabIndex = 142;
			this.ultraCheckEditor2.Text = "Descontinuado en Ventas";
			this.ultraCheckEditor2.Visible = false;
			// 
			// chkDesCompras
			// 
			appearance57.ForeColorDisabled = System.Drawing.Color.Black;
			this.chkDesCompras.Appearance = appearance57;
			this.chkDesCompras.BackColor = System.Drawing.Color.Transparent;
			this.chkDesCompras.DataBindings.Add(new System.Windows.Forms.Binding("CheckedValue", this.cdsCliente, "Cliente.DesconProv"));
			this.chkDesCompras.Location = new System.Drawing.Point(157, 328);
			this.chkDesCompras.Name = "chkDesCompras";
			this.chkDesCompras.Size = new System.Drawing.Size(168, 21);
			this.chkDesCompras.TabIndex = 141;
			this.chkDesCompras.Text = "Descontinuado en Compras";
			this.chkDesCompras.Visible = false;
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
			this.groupBox2.Location = new System.Drawing.Point(149, 112);
			this.groupBox2.Name = "groupBox2";
			this.groupBox2.Size = new System.Drawing.Size(640, 86);
			this.groupBox2.TabIndex = 2;
			this.groupBox2.TabStop = false;
			this.groupBox2.Text = " Conyuge ";
			// 
			// ultraTextEditor17
			// 
			appearance58.ForeColorDisabled = System.Drawing.Color.Black;
			this.ultraTextEditor17.Appearance = appearance58;
			this.ultraTextEditor17.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.ultraTextEditor17.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsCliente, "Cliente.ConyTelefono"));
			this.ultraTextEditor17.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.ultraTextEditor17.Location = new System.Drawing.Point(112, 52);
			this.ultraTextEditor17.Name = "ultraTextEditor17";
			this.ultraTextEditor17.Size = new System.Drawing.Size(192, 22);
			this.ultraTextEditor17.TabIndex = 2;
			this.ultraTextEditor17.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.ultraTextEditor17_KeyPress);
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
			appearance59.ForeColorDisabled = System.Drawing.Color.Black;
			this.ultraTextEditor14.Appearance = appearance59;
			this.ultraTextEditor14.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.ultraTextEditor14.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsCliente, "Cliente.ConyDireccion"));
			this.ultraTextEditor14.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.ultraTextEditor14.Location = new System.Drawing.Point(432, 26);
			this.ultraTextEditor14.Name = "ultraTextEditor14";
			this.ultraTextEditor14.Size = new System.Drawing.Size(192, 22);
			this.ultraTextEditor14.TabIndex = 1;
			this.ultraTextEditor14.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.ultraTextEditor14_KeyPress);
			// 
			// ultraTextEditor13
			// 
			appearance60.ForeColorDisabled = System.Drawing.Color.Black;
			this.ultraTextEditor13.Appearance = appearance60;
			this.ultraTextEditor13.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.ultraTextEditor13.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsCliente, "Cliente.ConyCelular"));
			this.ultraTextEditor13.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.ultraTextEditor13.Location = new System.Drawing.Point(432, 52);
			this.ultraTextEditor13.Name = "ultraTextEditor13";
			this.ultraTextEditor13.Size = new System.Drawing.Size(192, 22);
			this.ultraTextEditor13.TabIndex = 3;
			this.ultraTextEditor13.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.ultraTextEditor13_KeyPress);
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
			appearance61.ForeColorDisabled = System.Drawing.Color.Black;
			this.ultraTextEditor8.Appearance = appearance61;
			this.ultraTextEditor8.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.ultraTextEditor8.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsCliente, "Cliente.Conyuge"));
			this.ultraTextEditor8.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.ultraTextEditor8.Location = new System.Drawing.Point(112, 26);
			this.ultraTextEditor8.Name = "ultraTextEditor8";
			this.ultraTextEditor8.Size = new System.Drawing.Size(192, 22);
			this.ultraTextEditor8.TabIndex = 0;
			this.ultraTextEditor8.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.ultraTextEditor8_KeyPress);
			// 
			// label47
			// 
			this.label47.AutoSize = true;
			this.label47.BackColor = System.Drawing.Color.Transparent;
			this.label47.Location = new System.Drawing.Point(157, 43);
			this.label47.Name = "label47";
			this.label47.Size = new System.Drawing.Size(53, 16);
			this.label47.TabIndex = 27;
			this.label47.Text = "Ref. Local";
			// 
			// ultraTextEditor20
			// 
			appearance62.ForeColorDisabled = System.Drawing.Color.Black;
			this.ultraTextEditor20.Appearance = appearance62;
			this.ultraTextEditor20.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.ultraTextEditor20.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsCliente, "Cliente.RefLocal"));
			this.ultraTextEditor20.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.ultraTextEditor20.Location = new System.Drawing.Point(261, 43);
			this.ultraTextEditor20.Name = "ultraTextEditor20";
			this.ultraTextEditor20.Size = new System.Drawing.Size(512, 22);
			this.ultraTextEditor20.TabIndex = 0;
			this.ultraTextEditor20.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.ultraTextEditor20_KeyPress);
			// 
			// label45
			// 
			this.label45.AutoSize = true;
			this.label45.BackColor = System.Drawing.Color.Transparent;
			this.label45.Location = new System.Drawing.Point(493, 293);
			this.label45.Name = "label45";
			this.label45.Size = new System.Drawing.Size(36, 16);
			this.label45.TabIndex = 25;
			this.label45.Text = "Sector";
			// 
			// ultraTextEditor18
			// 
			appearance63.ForeColorDisabled = System.Drawing.Color.Black;
			this.ultraTextEditor18.Appearance = appearance63;
			this.ultraTextEditor18.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.ultraTextEditor18.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsCliente, "Cliente.Sector"));
			this.ultraTextEditor18.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.ultraTextEditor18.Location = new System.Drawing.Point(581, 293);
			this.ultraTextEditor18.Name = "ultraTextEditor18";
			this.ultraTextEditor18.Size = new System.Drawing.Size(192, 22);
			this.ultraTextEditor18.TabIndex = 4;
			// 
			// label46
			// 
			this.label46.AutoSize = true;
			this.label46.BackColor = System.Drawing.Color.Transparent;
			this.label46.Location = new System.Drawing.Point(157, 293);
			this.label46.Name = "label46";
			this.label46.Size = new System.Drawing.Size(52, 16);
			this.label46.TabIndex = 23;
			this.label46.Text = "Parroquia";
			// 
			// ultraTextEditor19
			// 
			appearance64.ForeColorDisabled = System.Drawing.Color.Black;
			this.ultraTextEditor19.Appearance = appearance64;
			this.ultraTextEditor19.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.ultraTextEditor19.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsCliente, "Cliente.Parroquia"));
			this.ultraTextEditor19.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.ultraTextEditor19.Location = new System.Drawing.Point(261, 293);
			this.ultraTextEditor19.Name = "ultraTextEditor19";
			this.ultraTextEditor19.Size = new System.Drawing.Size(192, 22);
			this.ultraTextEditor19.TabIndex = 3;
			this.ultraTextEditor19.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.ultraTextEditor19_KeyPress);
			// 
			// label42
			// 
			this.label42.AutoSize = true;
			this.label42.BackColor = System.Drawing.Color.Transparent;
			this.label42.Location = new System.Drawing.Point(493, 78);
			this.label42.Name = "label42";
			this.label42.Size = new System.Drawing.Size(76, 16);
			this.label42.TabIndex = 17;
			this.label42.Text = "Cuenta Ahorro";
			// 
			// ultraTextEditor15
			// 
			appearance65.ForeColorDisabled = System.Drawing.Color.Black;
			this.ultraTextEditor15.Appearance = appearance65;
			this.ultraTextEditor15.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.ultraTextEditor15.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsCliente, "Cliente.CuentaAh"));
			this.ultraTextEditor15.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.ultraTextEditor15.Location = new System.Drawing.Point(581, 78);
			this.ultraTextEditor15.Name = "ultraTextEditor15";
			this.ultraTextEditor15.Size = new System.Drawing.Size(192, 22);
			this.ultraTextEditor15.TabIndex = 2;
			this.ultraTextEditor15.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.ultraTextEditor15_KeyPress);
			// 
			// label35
			// 
			this.label35.AutoSize = true;
			this.label35.BackColor = System.Drawing.Color.Transparent;
			this.label35.Location = new System.Drawing.Point(157, 78);
			this.label35.Name = "label35";
			this.label35.Size = new System.Drawing.Size(89, 16);
			this.label35.TabIndex = 3;
			this.label35.Text = "Cuenta Corriente";
			// 
			// ultraTextEditor7
			// 
			appearance66.ForeColorDisabled = System.Drawing.Color.Black;
			this.ultraTextEditor7.Appearance = appearance66;
			this.ultraTextEditor7.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.ultraTextEditor7.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsCliente, "Cliente.CuentaCorr"));
			this.ultraTextEditor7.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.ultraTextEditor7.Location = new System.Drawing.Point(261, 78);
			this.ultraTextEditor7.Name = "ultraTextEditor7";
			this.ultraTextEditor7.Size = new System.Drawing.Size(192, 22);
			this.ultraTextEditor7.TabIndex = 1;
			this.ultraTextEditor7.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.ultraTextEditor7_KeyPress);
			// 
			// groupBox1
			// 
			this.groupBox1.BackColor = System.Drawing.Color.Transparent;
			this.groupBox1.Controls.Add(this.btCedRep);
			this.groupBox1.Controls.Add(this.label37);
			this.groupBox1.Controls.Add(this.ultraTextEditor9);
			this.groupBox1.Controls.Add(this.ultraTextEditor12);
			this.groupBox1.Controls.Add(this.label39);
			this.groupBox1.Controls.Add(this.label43);
			this.groupBox1.Controls.Add(this.label38);
			this.groupBox1.Controls.Add(this.ultraTextEditor16);
			this.groupBox1.Controls.Add(this.txtCedRep);
			this.groupBox1.Location = new System.Drawing.Point(149, 198);
			this.groupBox1.Name = "groupBox1";
			this.groupBox1.Size = new System.Drawing.Size(640, 86);
			this.groupBox1.TabIndex = 3;
			this.groupBox1.TabStop = false;
			this.groupBox1.Text = " Referencia ";
			// 
			// btCedRep
			// 
			appearance67.Image = ((object)(resources.GetObject("appearance67.Image")));
			this.btCedRep.Appearance = appearance67;
			this.btCedRep.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.btCedRep.Location = new System.Drawing.Point(600, 48);
			this.btCedRep.Name = "btCedRep";
			this.btCedRep.Size = new System.Drawing.Size(24, 26);
			this.btCedRep.TabIndex = 25;
			this.btCedRep.Click += new System.EventHandler(this.btCedRep_Click);
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
			appearance68.ForeColorDisabled = System.Drawing.Color.Black;
			this.ultraTextEditor9.Appearance = appearance68;
			this.ultraTextEditor9.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.ultraTextEditor9.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsCliente, "Cliente.RefDireccion"));
			this.ultraTextEditor9.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.ultraTextEditor9.Location = new System.Drawing.Point(112, 52);
			this.ultraTextEditor9.Name = "ultraTextEditor9";
			this.ultraTextEditor9.Size = new System.Drawing.Size(192, 22);
			this.ultraTextEditor9.TabIndex = 2;
			this.ultraTextEditor9.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.ultraTextEditor9_KeyPress);
			// 
			// ultraTextEditor12
			// 
			appearance69.ForeColorDisabled = System.Drawing.Color.Black;
			this.ultraTextEditor12.Appearance = appearance69;
			this.ultraTextEditor12.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.ultraTextEditor12.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsCliente, "Cliente.RefTelefono"));
			this.ultraTextEditor12.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.ultraTextEditor12.Location = new System.Drawing.Point(432, 26);
			this.ultraTextEditor12.Name = "ultraTextEditor12";
			this.ultraTextEditor12.Size = new System.Drawing.Size(192, 22);
			this.ultraTextEditor12.TabIndex = 1;
			this.ultraTextEditor12.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.ultraTextEditor12_KeyPress);
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
			this.label38.Size = new System.Drawing.Size(64, 16);
			this.label38.TabIndex = 9;
			this.label38.Text = "Cedula Rep.";
			// 
			// ultraTextEditor16
			// 
			appearance70.ForeColorDisabled = System.Drawing.Color.Black;
			this.ultraTextEditor16.Appearance = appearance70;
			this.ultraTextEditor16.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.ultraTextEditor16.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsCliente, "Cliente.Referencia"));
			this.ultraTextEditor16.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.ultraTextEditor16.Location = new System.Drawing.Point(112, 26);
			this.ultraTextEditor16.Name = "ultraTextEditor16";
			this.ultraTextEditor16.Size = new System.Drawing.Size(192, 22);
			this.ultraTextEditor16.TabIndex = 0;
			this.ultraTextEditor16.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.ultraTextEditor16_KeyPress);
			// 
			// txtCedRep
			// 
			appearance71.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtCedRep.Appearance = appearance71;
			this.txtCedRep.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtCedRep.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsCliente, "Cliente.Poliza"));
			this.txtCedRep.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtCedRep.Location = new System.Drawing.Point(432, 52);
			this.txtCedRep.Name = "txtCedRep";
			this.txtCedRep.Size = new System.Drawing.Size(160, 22);
			this.txtCedRep.TabIndex = 3;
			this.txtCedRep.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtCedRep_KeyPress);
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
			this.ultraTabPageControl3.Size = new System.Drawing.Size(938, 422);
			// 
			// btExcel
			// 
			this.btExcel.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btExcel.Location = new System.Drawing.Point(837, 9);
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
			appearance72.BackColor = System.Drawing.Color.White;
			this.ultraGrid1.DisplayLayout.Appearance = appearance72;
			appearance73.BackColor = System.Drawing.Color.Transparent;
			this.ultraGrid1.DisplayLayout.Override.CardAreaAppearance = appearance73;
			appearance74.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance74.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance74.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance74.FontData.BoldAsString = "True";
			appearance74.FontData.Name = "Arial";
			appearance74.FontData.SizeInPoints = 10F;
			appearance74.ForeColor = System.Drawing.Color.White;
			appearance74.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.ultraGrid1.DisplayLayout.Override.HeaderAppearance = appearance74;
			this.ultraGrid1.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti;
			appearance75.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance75.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance75.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.RowSelectorAppearance = appearance75;
			appearance76.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance76.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance76.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.SelectedRowAppearance = appearance76;
			this.ultraGrid1.DisplayLayout.ViewStyleBand = Infragistics.Win.UltraWinGrid.ViewStyleBand.OutlookGroupBy;
			this.ultraGrid1.Location = new System.Drawing.Point(8, 40);
			this.ultraGrid1.Name = "ultraGrid1";
			this.ultraGrid1.Size = new System.Drawing.Size(923, 373);
			this.ultraGrid1.TabIndex = 0;
			// 
			// ultraTabPageControl4
			// 
			this.ultraTabPageControl4.Controls.Add(this.ultraButton1);
			this.ultraTabPageControl4.Controls.Add(this.btListadoItems);
			this.ultraTabPageControl4.Controls.Add(this.grdItems);
			this.ultraTabPageControl4.Location = new System.Drawing.Point(-10000, -10000);
			this.ultraTabPageControl4.Name = "ultraTabPageControl4";
			this.ultraTabPageControl4.Size = new System.Drawing.Size(938, 422);
			// 
			// ultraButton1
			// 
			this.ultraButton1.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.ultraButton1.Location = new System.Drawing.Point(829, 9);
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
			appearance77.BackColor = System.Drawing.Color.White;
			this.grdItems.DisplayLayout.Appearance = appearance77;
			appearance78.BackColor = System.Drawing.Color.Transparent;
			this.grdItems.DisplayLayout.Override.CardAreaAppearance = appearance78;
			appearance79.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance79.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance79.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance79.FontData.BoldAsString = "True";
			appearance79.FontData.Name = "Arial";
			appearance79.FontData.SizeInPoints = 10F;
			appearance79.ForeColor = System.Drawing.Color.White;
			appearance79.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.grdItems.DisplayLayout.Override.HeaderAppearance = appearance79;
			this.grdItems.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti;
			appearance80.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance80.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance80.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.grdItems.DisplayLayout.Override.RowSelectorAppearance = appearance80;
			appearance81.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance81.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance81.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.grdItems.DisplayLayout.Override.SelectedRowAppearance = appearance81;
			this.grdItems.DisplayLayout.ViewStyleBand = Infragistics.Win.UltraWinGrid.ViewStyleBand.OutlookGroupBy;
			this.grdItems.Location = new System.Drawing.Point(8, 43);
			this.grdItems.Name = "grdItems";
			this.grdItems.Size = new System.Drawing.Size(923, 373);
			this.grdItems.TabIndex = 1;
			// 
			// ultraTabPageControl6
			// 
			this.ultraTabPageControl6.Controls.Add(this.radioGroup1);
			this.ultraTabPageControl6.Controls.Add(this.groupBox4);
			this.ultraTabPageControl6.Controls.Add(this.groupBox3);
			this.ultraTabPageControl6.Location = new System.Drawing.Point(-10000, -10000);
			this.ultraTabPageControl6.Name = "ultraTabPageControl6";
			this.ultraTabPageControl6.Size = new System.Drawing.Size(938, 422);
			// 
			// radioGroup1
			// 
			this.radioGroup1.EditValue = 1;
			this.radioGroup1.Location = new System.Drawing.Point(312, 32);
			this.radioGroup1.Name = "radioGroup1";
			// 
			// radioGroup1.Properties
			// 
			this.radioGroup1.Properties.Items.AddRange(new DevExpress.XtraEditors.Controls.RadioGroupItem[] {
																																																				new DevExpress.XtraEditors.Controls.RadioGroupItem(1, "Contado"),
																																																				new DevExpress.XtraEditors.Controls.RadioGroupItem(2, "Crédto")});
			this.radioGroup1.Size = new System.Drawing.Size(232, 32);
			this.radioGroup1.TabIndex = 146;
			// 
			// groupBox4
			// 
			this.groupBox4.Controls.Add(this.chkSujetoRetencion);
			this.groupBox4.Controls.Add(this.chkChofer);
			this.groupBox4.Controls.Add(this.chkEspecial);
			this.groupBox4.Controls.Add(this.label33);
			this.groupBox4.Controls.Add(this.cmbFormaPago);
			this.groupBox4.Controls.Add(this.txtAutRetencion);
			this.groupBox4.Controls.Add(this.lblAutRetencion);
			this.groupBox4.Controls.Add(this.label20);
			this.groupBox4.Controls.Add(this.cmbCobrador);
			this.groupBox4.Controls.Add(this.txtAutImprenta);
			this.groupBox4.Controls.Add(this.label16);
			this.groupBox4.Controls.Add(this.lblCuentaFondo);
			this.groupBox4.Controls.Add(this.cmbCuentaFondo);
			this.groupBox4.Controls.Add(this.lblCuentaAnticipo);
			this.groupBox4.Controls.Add(this.cmbCuentaAnticipo);
			this.groupBox4.Controls.Add(this.btDepende);
			this.groupBox4.Controls.Add(this.label25);
			this.groupBox4.Controls.Add(this.cmbDepende);
			this.groupBox4.Controls.Add(this.ultraTextEditor21);
			this.groupBox4.Controls.Add(this.label51);
			this.groupBox4.Controls.Add(this.ultraTextEditor4);
			this.groupBox4.Controls.Add(this.label50);
			this.groupBox4.Controls.Add(this.txtSerie);
			this.groupBox4.Controls.Add(this.label32);
			this.groupBox4.Controls.Add(this.label31);
			this.groupBox4.Controls.Add(this.cmbFechaCaduca);
			this.groupBox4.Controls.Add(this.txtAutFactura);
			this.groupBox4.Controls.Add(this.label15);
			this.groupBox4.Location = new System.Drawing.Point(16, 128);
			this.groupBox4.Name = "groupBox4";
			this.groupBox4.Size = new System.Drawing.Size(536, 224);
			this.groupBox4.TabIndex = 99;
			this.groupBox4.TabStop = false;
			this.groupBox4.Text = "Complemento";
			// 
			// chkSujetoRetencion
			// 
			appearance82.ForeColorDisabled = System.Drawing.Color.Black;
			this.chkSujetoRetencion.Appearance = appearance82;
			this.chkSujetoRetencion.BackColor = System.Drawing.Color.Transparent;
			this.chkSujetoRetencion.DataBindings.Add(new System.Windows.Forms.Binding("CheckedValue", this.cdsCliente, "Cliente.SujetoRetencion"));
			this.chkSujetoRetencion.Location = new System.Drawing.Point(344, 24);
			this.chkSujetoRetencion.Name = "chkSujetoRetencion";
			this.chkSujetoRetencion.Size = new System.Drawing.Size(112, 21);
			this.chkSujetoRetencion.TabIndex = 173;
			this.chkSujetoRetencion.Text = "Sujeto Retencion";
			this.chkSujetoRetencion.Visible = false;
			// 
			// chkChofer
			// 
			appearance83.ForeColorDisabled = System.Drawing.Color.Black;
			this.chkChofer.Appearance = appearance83;
			this.chkChofer.BackColor = System.Drawing.Color.Transparent;
			this.chkChofer.DataBindings.Add(new System.Windows.Forms.Binding("CheckedValue", this.cdsCliente, "Cliente.Chofer"));
			this.chkChofer.Location = new System.Drawing.Point(272, 24);
			this.chkChofer.Name = "chkChofer";
			this.chkChofer.Size = new System.Drawing.Size(64, 21);
			this.chkChofer.TabIndex = 172;
			this.chkChofer.Text = "Chofer";
			this.chkChofer.Visible = false;
			// 
			// chkEspecial
			// 
			appearance84.ForeColorDisabled = System.Drawing.Color.Black;
			this.chkEspecial.Appearance = appearance84;
			this.chkEspecial.BackColor = System.Drawing.Color.Transparent;
			this.chkEspecial.DataBindings.Add(new System.Windows.Forms.Binding("CheckedValue", this.cdsCliente, "Cliente.Especial"));
			this.chkEspecial.Location = new System.Drawing.Point(464, 24);
			this.chkEspecial.Name = "chkEspecial";
			this.chkEspecial.Size = new System.Drawing.Size(64, 21);
			this.chkEspecial.TabIndex = 171;
			this.chkEspecial.Text = "Especial";
			this.chkEspecial.Visible = false;
			// 
			// label33
			// 
			this.label33.AutoSize = true;
			this.label33.BackColor = System.Drawing.Color.Transparent;
			this.label33.Location = new System.Drawing.Point(264, 96);
			this.label33.Name = "label33";
			this.label33.Size = new System.Drawing.Size(64, 16);
			this.label33.TabIndex = 170;
			this.label33.Text = "Forma Pago";
			this.label33.Visible = false;
			// 
			// cmbFormaPago
			// 
			appearance85.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbFormaPago.Appearance = appearance85;
			this.cmbFormaPago.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbFormaPago.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbFormaPago.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsCliente, "Cliente.idFormaPago"));
			ultraGridColumn61.Header.VisiblePosition = 0;
			ultraGridColumn61.Hidden = true;
			ultraGridColumn62.Header.VisiblePosition = 1;
			ultraGridColumn62.Width = 203;
			ultraGridBand19.Columns.AddRange(new object[] {
																											ultraGridColumn61,
																											ultraGridColumn62});
			this.cmbFormaPago.DisplayLayout.BandsSerializer.Add(ultraGridBand19);
			this.cmbFormaPago.DisplayMember = "FormaPago";
			this.cmbFormaPago.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbFormaPago.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbFormaPago.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbFormaPago.Location = new System.Drawing.Point(360, 96);
			this.cmbFormaPago.Name = "cmbFormaPago";
			this.cmbFormaPago.Size = new System.Drawing.Size(160, 21);
			this.cmbFormaPago.TabIndex = 165;
			this.cmbFormaPago.ValueMember = "idFormaPago";
			this.cmbFormaPago.Visible = false;
			// 
			// txtAutRetencion
			// 
			appearance86.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtAutRetencion.Appearance = appearance86;
			this.txtAutRetencion.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsCliente, "Cliente.AutorizaRet"));
			this.txtAutRetencion.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtAutRetencion.Location = new System.Drawing.Point(392, 48);
			this.txtAutRetencion.Name = "txtAutRetencion";
			this.txtAutRetencion.Size = new System.Drawing.Size(128, 22);
			this.txtAutRetencion.TabIndex = 164;
			this.txtAutRetencion.Visible = false;
			// 
			// lblAutRetencion
			// 
			this.lblAutRetencion.BackColor = System.Drawing.Color.Transparent;
			this.lblAutRetencion.Location = new System.Drawing.Point(264, 48);
			this.lblAutRetencion.Name = "lblAutRetencion";
			this.lblAutRetencion.Size = new System.Drawing.Size(121, 17);
			this.lblAutRetencion.TabIndex = 169;
			this.lblAutRetencion.Text = "Autorización Retención";
			this.lblAutRetencion.Visible = false;
			// 
			// label20
			// 
			this.label20.AutoSize = true;
			this.label20.BackColor = System.Drawing.Color.Transparent;
			this.label20.Location = new System.Drawing.Point(264, 120);
			this.label20.Name = "label20";
			this.label20.Size = new System.Drawing.Size(54, 16);
			this.label20.TabIndex = 168;
			this.label20.Text = "Cobrador:";
			this.label20.Visible = false;
			// 
			// cmbCobrador
			// 
			appearance87.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbCobrador.Appearance = appearance87;
			this.cmbCobrador.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbCobrador.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbCobrador.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsCliente, "Cliente.idCobrador"));
			ultraGridColumn63.Header.VisiblePosition = 0;
			ultraGridColumn63.Hidden = true;
			ultraGridColumn64.Header.VisiblePosition = 1;
			ultraGridColumn64.Width = 304;
			ultraGridBand20.Columns.AddRange(new object[] {
																											ultraGridColumn63,
																											ultraGridColumn64});
			this.cmbCobrador.DisplayLayout.BandsSerializer.Add(ultraGridBand20);
			this.cmbCobrador.DisplayMember = "Nombre";
			this.cmbCobrador.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbCobrador.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbCobrador.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbCobrador.Location = new System.Drawing.Point(360, 120);
			this.cmbCobrador.Name = "cmbCobrador";
			this.cmbCobrador.Size = new System.Drawing.Size(160, 21);
			this.cmbCobrador.TabIndex = 166;
			this.cmbCobrador.ValueMember = "idPersonal";
			this.cmbCobrador.Visible = false;
			// 
			// txtAutImprenta
			// 
			appearance88.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtAutImprenta.Appearance = appearance88;
			this.txtAutImprenta.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsCliente, "Cliente.AutorizaImprenta"));
			this.txtAutImprenta.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtAutImprenta.Location = new System.Drawing.Point(392, 72);
			this.txtAutImprenta.Name = "txtAutImprenta";
			this.txtAutImprenta.Size = new System.Drawing.Size(128, 22);
			this.txtAutImprenta.TabIndex = 163;
			this.txtAutImprenta.Visible = false;
			// 
			// label16
			// 
			this.label16.AutoSize = true;
			this.label16.BackColor = System.Drawing.Color.Transparent;
			this.label16.Location = new System.Drawing.Point(264, 72);
			this.label16.Name = "label16";
			this.label16.Size = new System.Drawing.Size(96, 16);
			this.label16.TabIndex = 167;
			this.label16.Text = "Autoriz. Imprenta:";
			this.label16.Visible = false;
			// 
			// lblCuentaFondo
			// 
			this.lblCuentaFondo.AutoSize = true;
			this.lblCuentaFondo.BackColor = System.Drawing.Color.Transparent;
			this.lblCuentaFondo.Location = new System.Drawing.Point(16, 192);
			this.lblCuentaFondo.Name = "lblCuentaFondo";
			this.lblCuentaFondo.Size = new System.Drawing.Size(70, 16);
			this.lblCuentaFondo.TabIndex = 160;
			this.lblCuentaFondo.Text = "Cta Fondo G:";
			this.lblCuentaFondo.Visible = false;
			// 
			// cmbCuentaFondo
			// 
			appearance89.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbCuentaFondo.Appearance = appearance89;
			this.cmbCuentaFondo.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbCuentaFondo.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbCuentaFondo.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsCliente, "Cliente.idCuentaFondo"));
			ultraGridColumn65.Header.VisiblePosition = 0;
			ultraGridColumn65.Hidden = true;
			ultraGridColumn66.Header.VisiblePosition = 1;
			ultraGridColumn66.Width = 350;
			ultraGridBand21.Columns.AddRange(new object[] {
																											ultraGridColumn65,
																											ultraGridColumn66});
			this.cmbCuentaFondo.DisplayLayout.BandsSerializer.Add(ultraGridBand21);
			this.cmbCuentaFondo.DisplayMember = "";
			this.cmbCuentaFondo.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbCuentaFondo.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbCuentaFondo.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbCuentaFondo.Location = new System.Drawing.Point(96, 192);
			this.cmbCuentaFondo.Name = "cmbCuentaFondo";
			this.cmbCuentaFondo.Size = new System.Drawing.Size(144, 21);
			this.cmbCuentaFondo.TabIndex = 159;
			this.cmbCuentaFondo.ValueMember = "";
			this.cmbCuentaFondo.Visible = false;
			// 
			// lblCuentaAnticipo
			// 
			this.lblCuentaAnticipo.AutoSize = true;
			this.lblCuentaAnticipo.BackColor = System.Drawing.Color.Transparent;
			this.lblCuentaAnticipo.Location = new System.Drawing.Point(16, 168);
			this.lblCuentaAnticipo.Name = "lblCuentaAnticipo";
			this.lblCuentaAnticipo.Size = new System.Drawing.Size(68, 16);
			this.lblCuentaAnticipo.TabIndex = 158;
			this.lblCuentaAnticipo.Text = "Cta Anticipo:";
			this.lblCuentaAnticipo.Visible = false;
			// 
			// cmbCuentaAnticipo
			// 
			appearance90.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbCuentaAnticipo.Appearance = appearance90;
			this.cmbCuentaAnticipo.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbCuentaAnticipo.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbCuentaAnticipo.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsCliente, "Cliente.idCuentaAnticipo"));
			ultraGridColumn67.Header.VisiblePosition = 0;
			ultraGridColumn67.Hidden = true;
			ultraGridColumn68.Header.VisiblePosition = 1;
			ultraGridColumn68.Width = 350;
			ultraGridBand22.Columns.AddRange(new object[] {
																											ultraGridColumn67,
																											ultraGridColumn68});
			this.cmbCuentaAnticipo.DisplayLayout.BandsSerializer.Add(ultraGridBand22);
			this.cmbCuentaAnticipo.DisplayMember = "";
			this.cmbCuentaAnticipo.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbCuentaAnticipo.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbCuentaAnticipo.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbCuentaAnticipo.Location = new System.Drawing.Point(112, 168);
			this.cmbCuentaAnticipo.Name = "cmbCuentaAnticipo";
			this.cmbCuentaAnticipo.Size = new System.Drawing.Size(104, 21);
			this.cmbCuentaAnticipo.TabIndex = 157;
			this.cmbCuentaAnticipo.ValueMember = "";
			this.cmbCuentaAnticipo.Visible = false;
			// 
			// btDepende
			// 
			this.btDepende.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.btDepende.Location = new System.Drawing.Point(216, 168);
			this.btDepende.Name = "btDepende";
			this.btDepende.Size = new System.Drawing.Size(24, 26);
			this.btDepende.TabIndex = 155;
			this.btDepende.Text = "...";
			this.btDepende.Visible = false;
			// 
			// label25
			// 
			this.label25.AutoSize = true;
			this.label25.BackColor = System.Drawing.Color.Transparent;
			this.label25.Location = new System.Drawing.Point(16, 152);
			this.label25.Name = "label25";
			this.label25.Size = new System.Drawing.Size(68, 16);
			this.label25.TabIndex = 156;
			this.label25.Text = "Depende de:";
			this.label25.Visible = false;
			// 
			// cmbDepende
			// 
			appearance91.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbDepende.Appearance = appearance91;
			this.cmbDepende.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbDepende.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbDepende.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsCliente, "Cliente.idPadre"));
			this.cmbDepende.DisplayLayout.AutoFitColumns = true;
			ultraGridColumn69.Header.VisiblePosition = 0;
			ultraGridColumn69.Hidden = true;
			ultraGridColumn70.Header.VisiblePosition = 1;
			ultraGridColumn70.Width = 53;
			ultraGridBand23.Columns.AddRange(new object[] {
																											ultraGridColumn69,
																											ultraGridColumn70});
			this.cmbDepende.DisplayLayout.BandsSerializer.Add(ultraGridBand23);
			this.cmbDepende.DisplayMember = "Nombre";
			this.cmbDepende.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbDepende.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbDepende.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbDepende.Location = new System.Drawing.Point(112, 144);
			this.cmbDepende.Name = "cmbDepende";
			this.cmbDepende.Size = new System.Drawing.Size(72, 21);
			this.cmbDepende.TabIndex = 154;
			this.cmbDepende.ValueMember = "idCliente";
			this.cmbDepende.Visible = false;
			// 
			// ultraTextEditor21
			// 
			appearance92.ForeColorDisabled = System.Drawing.Color.Black;
			this.ultraTextEditor21.Appearance = appearance92;
			this.ultraTextEditor21.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsCliente, "Cliente.RangoHasta"));
			this.ultraTextEditor21.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.ultraTextEditor21.Location = new System.Drawing.Point(116, 120);
			this.ultraTextEditor21.Name = "ultraTextEditor21";
			this.ultraTextEditor21.Size = new System.Drawing.Size(136, 22);
			this.ultraTextEditor21.TabIndex = 152;
			// 
			// label51
			// 
			this.label51.AutoSize = true;
			this.label51.BackColor = System.Drawing.Color.Transparent;
			this.label51.Location = new System.Drawing.Point(20, 120);
			this.label51.Name = "label51";
			this.label51.Size = new System.Drawing.Size(67, 16);
			this.label51.TabIndex = 153;
			this.label51.Text = "Rango Final:";
			// 
			// ultraTextEditor4
			// 
			appearance93.ForeColorDisabled = System.Drawing.Color.Black;
			this.ultraTextEditor4.Appearance = appearance93;
			this.ultraTextEditor4.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsCliente, "Cliente.RangoDesde"));
			this.ultraTextEditor4.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.ultraTextEditor4.Location = new System.Drawing.Point(116, 96);
			this.ultraTextEditor4.Name = "ultraTextEditor4";
			this.ultraTextEditor4.Size = new System.Drawing.Size(136, 22);
			this.ultraTextEditor4.TabIndex = 150;
			// 
			// label50
			// 
			this.label50.AutoSize = true;
			this.label50.BackColor = System.Drawing.Color.Transparent;
			this.label50.Location = new System.Drawing.Point(20, 96);
			this.label50.Name = "label50";
			this.label50.Size = new System.Drawing.Size(73, 16);
			this.label50.TabIndex = 151;
			this.label50.Text = "Rango Inicial:";
			// 
			// txtSerie
			// 
			appearance94.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtSerie.Appearance = appearance94;
			this.txtSerie.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsCliente, "Cliente.SerieFactura"));
			this.txtSerie.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtSerie.Location = new System.Drawing.Point(116, 48);
			this.txtSerie.Name = "txtSerie";
			this.txtSerie.Size = new System.Drawing.Size(136, 22);
			this.txtSerie.TabIndex = 145;
			// 
			// label32
			// 
			this.label32.AutoSize = true;
			this.label32.BackColor = System.Drawing.Color.Transparent;
			this.label32.Location = new System.Drawing.Point(20, 48);
			this.label32.Name = "label32";
			this.label32.Size = new System.Drawing.Size(86, 16);
			this.label32.TabIndex = 149;
			this.label32.Text = "Serie de Factura";
			// 
			// label31
			// 
			this.label31.AutoSize = true;
			this.label31.BackColor = System.Drawing.Color.Transparent;
			this.label31.Location = new System.Drawing.Point(20, 72);
			this.label31.Name = "label31";
			this.label31.Size = new System.Drawing.Size(89, 16);
			this.label31.TabIndex = 148;
			this.label31.Text = "Fecha Caducidad";
			// 
			// cmbFechaCaduca
			// 
			appearance95.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbFechaCaduca.Appearance = appearance95;
			this.cmbFechaCaduca.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsCliente, "Cliente.FechaCaducidad"));
			this.cmbFechaCaduca.DateTime = new System.DateTime(2010, 8, 27, 0, 0, 0, 0);
			this.cmbFechaCaduca.FormatString = "dd/MMM/yyyy";
			this.cmbFechaCaduca.Location = new System.Drawing.Point(116, 72);
			this.cmbFechaCaduca.Name = "cmbFechaCaduca";
			this.cmbFechaCaduca.Size = new System.Drawing.Size(136, 22);
			this.cmbFechaCaduca.SpinButtonDisplayStyle = Infragistics.Win.ButtonDisplayStyle.Always;
			this.cmbFechaCaduca.TabIndex = 146;
			this.cmbFechaCaduca.Value = new System.DateTime(2010, 8, 27, 0, 0, 0, 0);
			// 
			// txtAutFactura
			// 
			appearance96.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtAutFactura.Appearance = appearance96;
			this.txtAutFactura.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsCliente, "Cliente.AutorizaFactura"));
			this.txtAutFactura.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtAutFactura.Location = new System.Drawing.Point(116, 24);
			this.txtAutFactura.Name = "txtAutFactura";
			this.txtAutFactura.Size = new System.Drawing.Size(136, 22);
			this.txtAutFactura.TabIndex = 144;
			// 
			// label15
			// 
			this.label15.AutoSize = true;
			this.label15.BackColor = System.Drawing.Color.Transparent;
			this.label15.Location = new System.Drawing.Point(20, 24);
			this.label15.Name = "label15";
			this.label15.Size = new System.Drawing.Size(87, 16);
			this.label15.TabIndex = 147;
			this.label15.Text = "Autoriz. Factura:";
			// 
			// groupBox3
			// 
			this.groupBox3.Controls.Add(this.ultraCheckEditor1);
			this.groupBox3.Controls.Add(this.chkProspecto);
			this.groupBox3.Controls.Add(this.label10);
			this.groupBox3.Controls.Add(this.ultraTextEditor5);
			this.groupBox3.Controls.Add(this.label6);
			this.groupBox3.Controls.Add(this.ultraCalendarCombo1);
			this.groupBox3.Controls.Add(this.cmbFecha);
			this.groupBox3.Location = new System.Drawing.Point(16, 16);
			this.groupBox3.Name = "groupBox3";
			this.groupBox3.Size = new System.Drawing.Size(272, 104);
			this.groupBox3.TabIndex = 98;
			this.groupBox3.TabStop = false;
			this.groupBox3.Text = "Basico";
			// 
			// ultraCheckEditor1
			// 
			appearance97.ForeColorDisabled = System.Drawing.Color.Black;
			this.ultraCheckEditor1.Appearance = appearance97;
			this.ultraCheckEditor1.BackColor = System.Drawing.Color.Transparent;
			this.ultraCheckEditor1.DataBindings.Add(new System.Windows.Forms.Binding("CheckedValue", this.cdsCliente, "Cliente.Ambos"));
			this.ultraCheckEditor1.Location = new System.Drawing.Point(8, 16);
			this.ultraCheckEditor1.Name = "ultraCheckEditor1";
			this.ultraCheckEditor1.Size = new System.Drawing.Size(128, 21);
			this.ultraCheckEditor1.TabIndex = 94;
			this.ultraCheckEditor1.Text = "Proveedor y Cliente";
			this.ultraCheckEditor1.Visible = false;
			// 
			// chkProspecto
			// 
			appearance98.ForeColorDisabled = System.Drawing.Color.Black;
			this.chkProspecto.Appearance = appearance98;
			this.chkProspecto.BackColor = System.Drawing.Color.Transparent;
			this.chkProspecto.DataBindings.Add(new System.Windows.Forms.Binding("CheckedValue", this.cdsCliente, "Cliente.Prospecto"));
			this.chkProspecto.Location = new System.Drawing.Point(136, 16);
			this.chkProspecto.Name = "chkProspecto";
			this.chkProspecto.Size = new System.Drawing.Size(96, 21);
			this.chkProspecto.TabIndex = 95;
			this.chkProspecto.Text = "Prospecto";
			this.chkProspecto.Visible = false;
			// 
			// label10
			// 
			this.label10.AutoSize = true;
			this.label10.BackColor = System.Drawing.Color.Transparent;
			this.label10.Location = new System.Drawing.Point(8, 40);
			this.label10.Name = "label10";
			this.label10.Size = new System.Drawing.Size(52, 16);
			this.label10.TabIndex = 97;
			this.label10.Text = "Contacto:";
			this.label10.Visible = false;
			// 
			// ultraTextEditor5
			// 
			appearance99.ForeColorDisabled = System.Drawing.Color.Black;
			this.ultraTextEditor5.Appearance = appearance99;
			this.ultraTextEditor5.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.ultraTextEditor5.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsCliente, "Cliente.Contacto"));
			this.ultraTextEditor5.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.ultraTextEditor5.Location = new System.Drawing.Point(72, 40);
			this.ultraTextEditor5.Name = "ultraTextEditor5";
			this.ultraTextEditor5.Size = new System.Drawing.Size(72, 22);
			this.ultraTextEditor5.TabIndex = 96;
			this.ultraTextEditor5.Visible = false;
			// 
			// label6
			// 
			this.label6.AutoSize = true;
			this.label6.BackColor = System.Drawing.Color.Transparent;
			this.label6.Location = new System.Drawing.Point(8, 72);
			this.label6.Name = "label6";
			this.label6.Size = new System.Drawing.Size(46, 16);
			this.label6.TabIndex = 93;
			this.label6.Text = "Periodo:";
			this.label6.Visible = false;
			// 
			// ultraCalendarCombo1
			// 
			appearance100.ForeColorDisabled = System.Drawing.Color.Black;
			this.ultraCalendarCombo1.Appearance = appearance100;
			this.ultraCalendarCombo1.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.ultraCalendarCombo1.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsCliente, "Cliente.Vencimiento"));
			dateButton1.Caption = "Today";
			this.ultraCalendarCombo1.DateButtons.Add(dateButton1);
			this.ultraCalendarCombo1.Format = "dd/MMM/yyyy";
			this.ultraCalendarCombo1.Location = new System.Drawing.Point(56, 72);
			this.ultraCalendarCombo1.Name = "ultraCalendarCombo1";
			this.ultraCalendarCombo1.NonAutoSizeHeight = 23;
			this.ultraCalendarCombo1.Size = new System.Drawing.Size(104, 21);
			this.ultraCalendarCombo1.SpinButtonsVisible = true;
			this.ultraCalendarCombo1.TabIndex = 92;
			this.ultraCalendarCombo1.Value = new System.DateTime(2006, 1, 14, 0, 0, 0, 0);
			this.ultraCalendarCombo1.Visible = false;
			// 
			// cmbFecha
			// 
			appearance101.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbFecha.Appearance = appearance101;
			this.cmbFecha.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.cmbFecha.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsCliente, "Cliente.Fecha"));
			dateButton2.Caption = "Today";
			this.cmbFecha.DateButtons.Add(dateButton2);
			this.cmbFecha.Format = "dd/MMM/yyyy";
			this.cmbFecha.Location = new System.Drawing.Point(160, 72);
			this.cmbFecha.Name = "cmbFecha";
			this.cmbFecha.NonAutoSizeHeight = 23;
			this.cmbFecha.ReadOnly = true;
			this.cmbFecha.Size = new System.Drawing.Size(104, 21);
			this.cmbFecha.SpinButtonsVisible = true;
			this.cmbFecha.TabIndex = 91;
			this.cmbFecha.Value = new System.DateTime(2013, 3, 14, 0, 0, 0, 0);
			this.cmbFecha.Visible = false;
			// 
			// ultraTabPageControl7
			// 
			this.ultraTabPageControl7.Controls.Add(this.ultraGrid3);
			this.ultraTabPageControl7.Location = new System.Drawing.Point(-10000, -10000);
			this.ultraTabPageControl7.Name = "ultraTabPageControl7";
			this.ultraTabPageControl7.Size = new System.Drawing.Size(938, 422);
			// 
			// ultraGrid3
			// 
			this.ultraGrid3.Cursor = System.Windows.Forms.Cursors.Default;
			this.ultraGrid3.DataMember = "ClientesDirecciones";
			this.ultraGrid3.DataSource = this.cdsDirecciones;
			appearance102.BackColor = System.Drawing.Color.White;
			this.ultraGrid3.DisplayLayout.Appearance = appearance102;
			this.ultraGrid3.DisplayLayout.AutoFitColumns = true;
			ultraGridColumn71.Header.VisiblePosition = 0;
			ultraGridColumn71.Hidden = true;
			ultraGridColumn71.Width = 273;
			ultraGridColumn72.Header.VisiblePosition = 1;
			ultraGridColumn72.Hidden = true;
			ultraGridColumn72.Width = 338;
			ultraGridColumn73.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn73.Header.VisiblePosition = 2;
			ultraGridColumn73.Width = 474;
			ultraGridBand24.Columns.AddRange(new object[] {
																											ultraGridColumn71,
																											ultraGridColumn72,
																											ultraGridColumn73});
			this.ultraGrid3.DisplayLayout.BandsSerializer.Add(ultraGridBand24);
			this.ultraGrid3.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.TemplateOnBottom;
			this.ultraGrid3.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.True;
			this.ultraGrid3.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance103.BackColor = System.Drawing.Color.Transparent;
			this.ultraGrid3.DisplayLayout.Override.CardAreaAppearance = appearance103;
			appearance104.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance104.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance104.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance104.FontData.BoldAsString = "True";
			appearance104.FontData.Name = "Arial";
			appearance104.FontData.SizeInPoints = 10F;
			appearance104.ForeColor = System.Drawing.Color.White;
			appearance104.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.ultraGrid3.DisplayLayout.Override.HeaderAppearance = appearance104;
			appearance105.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance105.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance105.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid3.DisplayLayout.Override.RowSelectorAppearance = appearance105;
			appearance106.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance106.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance106.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid3.DisplayLayout.Override.SelectedRowAppearance = appearance106;
			this.ultraGrid3.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.ultraGrid3.Location = new System.Drawing.Point(9, 6);
			this.ultraGrid3.Name = "ultraGrid3";
			this.ultraGrid3.Size = new System.Drawing.Size(495, 218);
			this.ultraGrid3.TabIndex = 2;
			this.ultraGrid3.AfterRowInsert += new Infragistics.Win.UltraWinGrid.RowEventHandler(this.ultraGrid3_AfterRowInsert);
			this.ultraGrid3.KeyDown += new System.Windows.Forms.KeyEventHandler(this.ultraGrid3_KeyDown);
			// 
			// cdsDirecciones
			// 
			this.cdsDirecciones.BindingContextCtrl = this;
			this.cdsDirecciones.DataLibrary = "LibFacturacion";
			this.cdsDirecciones.DataLibraryUrl = "";
			this.cdsDirecciones.DataSetDef = "dsClienteDirecciones";
			this.cdsDirecciones.FillOnRequest = false;
			this.cdsDirecciones.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsDirecciones.Name = "cdsDirecciones";
			this.cdsDirecciones.SchemaClassName = "LibFacturacion.DataClass";
			this.cdsDirecciones.SchemaDef = null;
			this.cdsDirecciones.BeforeFill += new C1.Data.FillEventHandler(this.cdsDirecciones_BeforeFill);
			// 
			// errorProvider
			// 
			this.errorProvider.ContainerControl = this;
			// 
			// txtIdCliente
			// 
			this.txtIdCliente.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsCliente, "Cliente.idCliente"));
			this.txtIdCliente.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtIdCliente.Location = new System.Drawing.Point(592, 464);
			this.txtIdCliente.Name = "txtIdCliente";
			this.txtIdCliente.Size = new System.Drawing.Size(48, 22);
			this.txtIdCliente.TabIndex = 76;
			// 
			// barraDato1
			// 
			this.barraDato1.BarraMovimiento = true;
			this.barraDato1.CausesValidation = false;
			this.barraDato1.DataMember = "Cliente";
			this.barraDato1.DataNombreId = "idCliente";
			this.barraDato1.DataOrden = "";
			this.barraDato1.DataSource = this.cdsCliente;
			this.barraDato1.DataTabla = "Cliente";
			this.barraDato1.DataTablaHija = "";
			this.barraDato1.Location = new System.Drawing.Point(8, 464);
			this.barraDato1.Name = "barraDato1";
			this.barraDato1.Size = new System.Drawing.Size(296, 26);
			this.barraDato1.TabIndex = 46;
			this.barraDato1.VisibleBorrar = true;
			this.barraDato1.VisibleBuscar = true;
			this.barraDato1.VisibleEditar = true;
			this.barraDato1.VisibleImprimir = true;
			this.barraDato1.VisibleNuevo = true;
			this.barraDato1.ContextMenuChanged += new System.EventHandler(this.barraDato1_ContextMenuChanged);
			this.barraDato1.Buscar += new System.EventHandler(this.barraDato1_Buscar);
			this.barraDato1.Load += new System.EventHandler(this.barraDato1_Load);
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
			this.tbBasico.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			appearance107.BackColor = System.Drawing.Color.LightSteelBlue;
			appearance107.BackColor2 = System.Drawing.Color.LightSteelBlue;
			this.tbBasico.Appearance = appearance107;
			this.tbBasico.BackColor = System.Drawing.Color.LightSteelBlue;
			appearance108.BackColor = System.Drawing.Color.LightSteelBlue;
			appearance108.BackColor2 = System.Drawing.Color.LightSteelBlue;
			this.tbBasico.ClientAreaAppearance = appearance108;
			this.tbBasico.Controls.Add(this.ultraTabSharedControlsPage1);
			this.tbBasico.Controls.Add(this.ultraTabPageControl1);
			this.tbBasico.Controls.Add(this.tpComplemento);
			this.tbBasico.Controls.Add(this.ultraTabPageControl2);
			this.tbBasico.Controls.Add(this.ultraTabPageControl3);
			this.tbBasico.Controls.Add(this.ultraTabPageControl4);
			this.tbBasico.Controls.Add(this.ultraTabPageControl5);
			this.tbBasico.Controls.Add(this.ultraTabPageControl6);
			this.tbBasico.Controls.Add(this.ultraTabPageControl7);
			this.tbBasico.Location = new System.Drawing.Point(5, 8);
			this.tbBasico.Name = "tbBasico";
			this.tbBasico.SharedControlsPage = this.ultraTabSharedControlsPage1;
			this.tbBasico.Size = new System.Drawing.Size(942, 448);
			this.tbBasico.TabIndex = 47;
			appearance109.BackColor = System.Drawing.Color.LightSteelBlue;
			appearance109.BackColor2 = System.Drawing.Color.LightSteelBlue;
			ultraTab1.ActiveAppearance = appearance109;
			appearance110.BackColor = System.Drawing.Color.LightSteelBlue;
			appearance110.BackColor2 = System.Drawing.Color.LightSteelBlue;
			ultraTab1.Appearance = appearance110;
			appearance111.BackColor = System.Drawing.Color.LightSteelBlue;
			appearance111.BackColor2 = System.Drawing.Color.LightSteelBlue;
			ultraTab1.ClientAreaAppearance = appearance111;
			appearance112.BackColor = System.Drawing.Color.LightSteelBlue;
			appearance112.BackColor2 = System.Drawing.Color.LightSteelBlue;
			ultraTab1.SelectedAppearance = appearance112;
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
			ultraTab7.TabPage = this.ultraTabPageControl6;
			ultraTab7.Text = "tab1";
			ultraTab7.Visible = false;
			ultraTab8.TabPage = this.ultraTabPageControl7;
			ultraTab8.Text = "Direcciones de entrega";
			ultraTab8.Visible = false;
			this.tbBasico.Tabs.AddRange(new Infragistics.Win.UltraWinTabControl.UltraTab[] {
																																											 ultraTab1,
																																											 ultraTab2,
																																											 ultraTab3,
																																											 ultraTab4,
																																											 ultraTab5,
																																											 ultraTab6,
																																											 ultraTab7,
																																											 ultraTab8});
			this.tbBasico.SelectedTabChanged += new Infragistics.Win.UltraWinTabControl.SelectedTabChangedEventHandler(this.tbBasico_SelectedTabChanged);
			// 
			// ultraTabSharedControlsPage1
			// 
			this.ultraTabSharedControlsPage1.Location = new System.Drawing.Point(-10000, -10000);
			this.ultraTabSharedControlsPage1.Name = "ultraTabSharedControlsPage1";
			this.ultraTabSharedControlsPage1.Size = new System.Drawing.Size(938, 422);
			// 
			// btCRMClientes
			// 
			appearance113.Image = 5;
			this.btCRMClientes.Appearance = appearance113;
			this.btCRMClientes.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.btCRMClientes.ImageList = this.imageList;
			this.btCRMClientes.Location = new System.Drawing.Point(352, 464);
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
			appearance114.Image = ((object)(resources.GetObject("appearance114.Image")));
			this.btAuditoria.Appearance = appearance114;
			this.btAuditoria.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.btAuditoria.Location = new System.Drawing.Point(328, 464);
			this.btAuditoria.Name = "btAuditoria";
			this.btAuditoria.Size = new System.Drawing.Size(24, 25);
			this.btAuditoria.TabIndex = 85;
			this.toolTip1.SetToolTip(this.btAuditoria, "Auditoria de Cliente");
			this.btAuditoria.Click += new System.EventHandler(this.btAuditoria_Click);
			this.btAuditoria.MouseUp += new System.Windows.Forms.MouseEventHandler(this.btAuditoria_MouseUp);
			// 
			// btProveedorCliente
			// 
			this.btProveedorCliente.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.btProveedorCliente.ImageList = this.imageList;
			this.btProveedorCliente.Location = new System.Drawing.Point(392, 464);
			this.btProveedorCliente.Name = "btProveedorCliente";
			this.btProveedorCliente.Size = new System.Drawing.Size(115, 25);
			this.btProveedorCliente.TabIndex = 45;
			this.btProveedorCliente.Text = "&Grabar como cliente";
			this.toolTip1.SetToolTip(this.btProveedorCliente, "Grabar como cliente");
			this.btProveedorCliente.Visible = false;
			this.btProveedorCliente.Click += new System.EventHandler(this.btProveedorCliente_Click);
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
			// Cliente
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 14);
			this.CausesValidation = false;
			this.ClientSize = new System.Drawing.Size(952, 496);
			this.Controls.Add(this.btProveedorCliente);
			this.Controls.Add(this.btAuditoria);
			this.Controls.Add(this.btCRMClientes);
			this.Controls.Add(this.tbBasico);
			this.Controls.Add(this.txtIdCliente);
			this.Controls.Add(this.barraDato1);
			this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
			this.KeyPreview = true;
			this.MaximizeBox = false;
			this.Name = "Cliente";
			this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
			this.Text = "Cliente";
			this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.Cliente_KeyDown);
			this.Closing += new System.ComponentModel.CancelEventHandler(this.Cliente_Closing);
			this.Load += new System.EventHandler(this.Cliente_Load);
			this.ultraTabPageControl1.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.cdsCliente)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraTextEditor1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbSujetoRetencion)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource4)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbAgenteDeRetencion)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbTipoProveedor)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDiasConsignacion)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbTipoRuc)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtRuc)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbPais)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtPais)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNombreComercial)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNotas)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtObservacion)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCelular)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtEMail)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtRepresentante)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCiudad)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbGrupo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNombre)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDireccion)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtTelefono)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtFax)).EndInit();
			this.tpComplemento.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.txtDiasCredito4)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbInstitucion)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource3)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbArticuloSRI)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsArticuloTabla)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbBanco)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtBeneficiario)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCuentaSwift)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbDias)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbClientePublicidad)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbPrecio)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbSector)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCodigo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDiasCredito3)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDiasCredito2)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbTipoCredito)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsContacto)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCredTribut)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbPorcIvaRet)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbVendedor)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.spnDescCredito)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.spnDescuento)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtValorCredito)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDiasCredito1)).EndInit();
			this.ultraTabPageControl5.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid2)).EndInit();
			this.ultraTabPageControl2.ResumeLayout(false);
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
			((System.ComponentModel.ISupportInitialize)(this.txtCedRep)).EndInit();
			this.ultraTabPageControl3.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).EndInit();
			this.ultraTabPageControl4.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.grdItems)).EndInit();
			this.ultraTabPageControl6.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.radioGroup1.Properties)).EndInit();
			this.groupBox4.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.cmbFormaPago)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtAutRetencion)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCobrador)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtAutImprenta)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCuentaFondo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCuentaAnticipo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbDepende)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraTextEditor21)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraTextEditor4)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtSerie)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbFechaCaduca)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtAutFactura)).EndInit();
			this.groupBox3.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.ultraTextEditor5)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraCalendarCombo1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbFecha)).EndInit();
			this.ultraTabPageControl7.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid3)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsDirecciones)).EndInit();
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
		//private Compra miCompra;
		#endregion Variables Globales
		#region Pantalla Cliente
		public Cliente(bool bTipo)
		{
			InitializeComponent();
			TipoCliente = bTipo;
		}
		public Cliente(bool bTipo, int id)
		{
			InitializeComponent();
			TipoCliente = bTipo;
			IdCliente = id;
		}
		private void UnloadMe()
		{
			this.Close();
		}

		private int WM_CLOSE = 0x0010;
		protected override void WndProc(ref System.Windows.Forms.Message msg)
		{
			if (msg.Msg == WM_CLOSE) 
			{
				ActiveControl.CausesValidation = false;
			}			
			base.WndProc(ref msg);			
		}

		private void Cliente_Load(object sender, System.EventArgs e)
		{
			Funcion.Resolucion(this, this.Height, this.Width);

			#region Acceso Global
			string stCodCliente = "101";
			if (TipoCliente) stCodCliente = "116";
		
			if (!Funcion.Permiso(stCodCliente,cdsSeteoF))
			{
				stCodCliente = "Proveedores";
				if (TipoCliente) stCodCliente = "Clientes";

				MessageBox.Show("No puede ingresar a " + stCodCliente, "Informacion");
				BeginInvoke(new MethodInvoker(UnloadMe));
				return;
			}
			#endregion Acceso Global
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
			#region Personalizaciones
			#region Mutualista
			if (Funcion.bEjecutaSQL(cdsCliente, "Exec SeteoGExiste 'FVCPV'")) 				
			{
				if (!TipoCliente)
				{
					cmbCuentaAnticipo.Visible = true;
					lblCuentaAnticipo.Visible = true;
					cmbCuentaFondo.Visible = true;
					lblCuentaFondo.Visible= true;
					llenaComboCuentaAnticipo();
					llenaComboCuentaFondo();
				}
			}
			#endregion Mutualista
			#region Guias Natranscom
			if (Funcion.bEjecutaSQL(cdsCliente, "Exec SeteoGExiste 'GNAT'")) 				
			{
				chkChofer.Visible =true;			
			}
			#endregion
			#endregion Personalizaciones
		}

		private void Seguridad()
		{
			if (!TipoCliente) 
				miAcceso =new Acceso(cdsSeteoF,  "0601");
			else
				miAcceso =new Acceso(cdsSeteoF,  "0701");

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
			bCodAuto = Funcion.bEjecutaSQL(cdsCliente, "SELECT IsNull(AUTOMATICO, 0) FROM COMPRANUMERO WHERE IDTIPOFACTURA = " + idTipoF.ToString());
			if (bCodAuto) txtCodigo.ReadOnly = true;
			#endregion Codigo Automatico
			#region Campos Visibles
			if (!TipoCliente) 
			{
				stCliente = "Proveedor: ";
				txtAutRetencion.Width = 0;
				lblAutRetencion.Width = 0;

				this.label2.Visible = false;
				this.cmbPrecio.Visible = false;
				this.btPrecio.Visible = false;
				this.ultraCheckEditor1.Visible = false;
				this.chkProspecto.Visible = false;
				this.label30.Visible = false;
				this.btPublicidad.Visible = false;
				this.cmbClientePublicidad.Visible = false;
				this.label25.Visible = false;
				this.cmbDepende.Visible = false;
				this.btDepende.Visible = false;
				this.label33.Visible = false;
				this.cmbFormaPago.Visible = false;
				this.label20.Visible = false;
				this.cmbCobrador.Visible = false;
				this.label21.Visible = false;
				this.cmbVendedor.Visible = false;
				this.btVendedor.Visible = false;
				this.chkSujetoRetencion.Visible = false;
				this.chkEspecial.Visible = false;
				this.ultraCalendarCombo1.Visible = false;
				this.cmbFecha.Value = false; 
				this.btProveedorCliente.Visible = true;
				//				this.chkAplicaConsignacion.Visible = false;	
				this.txtAutFactura.Visible = true;
				this.label15.Visible = true;
				this.txtSerie.Visible = true;
				this.label32.Visible = true;
				this.cmbFechaCaduca.Visible = true;
				this.label31.Visible = true;
				this.ultraTextEditor4.Visible = true;
				this.label50.Visible = true;
				this.ultraTextEditor21.Visible = true;
				this.label51.Visible = true;
				this.chkHaceRetencion.Visible = false;
				this.tbBasico.Tabs[3].Visible = false;
//				this.lblArticuloSRI.Visible = true;
//				this.cmbArticuloSRI.Visible = true; 
				this.lblTipoProveedor.Visible = true;
				this.cmbTipoProveedor.Visible = true;
				this.cmbSujetoRetencion.Visible = true;
				this.lblASRetencion.Text = "Sujeto De Retención";
				
				#region FormaPago
				this.label52.Visible = false;					
				this.cmbTipoCredito.Visible = false;
				this.ultraButton2.Visible = false;
				#endregion FormaPago

				stFiltro = "Proveedor = 1 Or Ambos = 1";
				stClientes = this.Text;
				chkPublica.Text = "RISE";	
				this.lblNotas.Text = "Productos que comercializa";
			}
			else
			{
				stFiltro = "Proveedor = 0 Or Ambos = 1";
				#region Oculta Campo para Pantalla de Cliente
				this.txtAutFactura.Visible = false;
				this.label15.Visible = false;
				this.txtSerie.Visible = false;
				this.label32.Visible = false;
				this.cmbFechaCaduca.Visible = false;
				this.label31.Visible = false;
				this.ultraTextEditor4.Visible = false;
				this.label50.Visible = false;
				this.ultraTextEditor21.Visible = false;
				this.label51.Visible = false;
				this.chkHaceRetencion.Visible = true;
				this.label29.Visible = false;
				this.cmbPorcIvaRet.Visible = false;
				this.lblDias.Visible = true;
				this.cmbDias.Visible = true;
				this.lblHoraDesde.Visible = true;
				this.dtHoraDesde.Visible = true;
				this.lblHoraHasta.Visible = true;
				this.dtHoraHasta.Visible = true;
				this.chkCodigoBarras.Visible = true;				
				this.cmbAgenteDeRetencion.Visible = true;
				this.ultraGrid2.DisplayLayout.Bands[0].Columns["EnvioMail"].Hidden = true;
				//this.ultraGrid2.DisplayLayout.Bands[0].Columns["EMail"].Hidden = true;
				this.tbBasico.Tabs[7].Visible = true;
				#endregion Oculta Campo para Pantalla de Cliente
			}
			this.ultraCheckEditor1.Visible = false;
			#endregion Campos Visibles

			barraDato1.stFiltroExt = " And (" + stFiltro + ")";
			txtIdCliente.Width = 0;
			drSeteoF = cdsSeteoF.TableViews["SeteoF"].Rows[0];		
		}

		private void ObjetosChina(bool bEstado)
		{
			this.lblBanco.Visible = bEstado;
			this.cmbBanco.Visible = bEstado;
			this.btBanco.Visible = bEstado;
			this.lblCuentaSwift.Visible = bEstado;
			this.txtCuentaSwift.Visible = bEstado;
			this.lblBeneficiario.Visible = bEstado;
			this.txtBeneficiario.Visible = bEstado;
		}

		private void CargarCombos()
		{
			int iCliente = 1;
			if (!TipoCliente) iCliente = 0;
			cmbClientePublicidad.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Select idClientePublicidad, Medio From ClientePublicidad Order By Medio");
			cmbPrecio.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Select idTipoPrecio, Precio From ClientePrecio");
			cmbCiudad.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Select idCiudad, Ciudad From ClienteCiudad Order By Ciudad");
			cmbSector.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Select idSector, Sector From ClienteSector Order by Sector");
			cmbGrupo.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Select idGrupoCliente, Grupo From ClienteGrupo Where Proveedor = '" + TipoCliente + "' Order by Grupo");
			cmbFormaPago.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Exec CompraFormaPagoCarga");
			cmbTipoRuc.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Exec ClienteTipoRucCarga " + iCliente);
			this.cmbTipoProveedor.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Select idTipoProveedor, Codigo, Tipo From TipoDeProveedor Order By idTipoProveedor Asc");
			cmbBanco.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Select idBanco, Nombre From Banco Order By Nombre");
			this.cmbBanco.DisplayLayout.Bands[0].Columns["idBanco"].Hidden = true;
			this.cmbBanco.DisplayLayout.Bands[0].Columns["Nombre"].Width = 168;
			
			this.cmbAgenteDeRetencion.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Select idAgenteRetencion, Nombre From AgenteRetencion");

			this.cmbSujetoRetencion.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Select idSujetoRetencion, Nombre From SujetoRetencion");

			this.cmbInstitucion.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Select idInstitucion, Descripcion From Cre_Instituciones Order By Descripcion");
			cmbPorcIvaRet.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Select idRetIva, Porcentaje, Detalle From RetencionIva Order by Porcentaje");
			cmbPais.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Select idClientePais, Pais From ClientePais Order by Pais");
			cmbPais.Visible = false;
			bPais.Visible = false;
			if(Funcion.bEjecutaSQL(cdsPais,"Exec SeteoGExiste 'FClPais'"))
			{
				txtPais.Visible = false;
				cmbPais.Visible = true;
				bPais.Visible = true;
			}

			cmbCredTribut.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Select idCredTributario, CredTributario from CompraCredTribut");
			cmbDepende.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Select idCliente, Nombre From Cliente Where Entrega = 1 And " + stFiltro + " Order by Nombre");
			cmbVendedor.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Select idPersonal, Nombre From Personal Where Vendedor = 1 Order by Nombre");
			cmbCobrador.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Select idPersonal, Nombre From Personal Where Cobrador = 1 Order by Nombre");
			this.cmbDias.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Select idDia, Dia From ClienteDias Order by idDia");
			cmbDias.DisplayLayout.Bands[0].Columns["idDia"].Hidden = true;
			cmbDias.DisplayLayout.Bands[0].Columns["Dia"].Width = 160;
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
			HabilitaTab();
			#region Predefinidos
			if (TipoCliente) 
			{
				drCliente["idTipoRuc"] = 4;
				drCliente["idGrupoCliente"] = 6;
				drCliente["idPais"] = 1;
				drCliente["idCiudad"] = 1;
				if ((bool)drSeteoF["VentaCredito"] == true)
					drCliente["ContadoCred"] = 1;
				else
					drCliente["ContadoCred"] = 1;

				this.rdbContado.Checked = true;

				this.txtValorCredito.Enabled = false;
				this.txtDiasCredito1.Enabled = false;
				this.txtDiasCredito2.Enabled = false;
				this.txtDiasCredito3.Enabled = false;

				this.cmbTipoCredito.Enabled = false;
			}
			else 
			{
				drCliente["idTipoRuc"] = 1;
				drCliente["ContadoCred"] = 2;
				
				this.rdbCredito.Checked = true;
				cmbGrupo.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Select idGrupoCliente, Grupo From ClienteGrupo Where Proveedor = '" + TipoCliente + "' And idGrupoCliente Not In (2, 3) Order by Grupo ");
				ObjetosChina(false);
			}

			drCliente["Especial"] = false;
			drCliente["Pais"] = "Ecuador";
			drCliente["Fecha"] = DateTime.Today;
			drCliente["Proveedor"] = !TipoCliente;			
			//			drCliente["idSucursal"] = (int) drSeteo["NumEmpresa"];
			#endregion Predefinidos
			//HabilitaTab();
			if (!TipoCliente) this.btProveedorCliente.Enabled = false;
			this.radioGroup1.Enabled = false;
			this.cmbTipoCredito.Enabled = false;
			this.rdbContado.Enabled = true;
			this.rdbCredito.Enabled = true;
			//			ultraTabPageControl1.Focus();
			this.tbBasico.Tabs[0].Selected = true;

			//ultraTabPageControl1.Show();
			//this.cmbTipoRuc.Focus();

			this.txtCelular.MaxLength = 10;
			this.txtTelefono.MaxLength = 9;
			this.txtFax.MaxLength = 9;
			
			//if (!TipoCliente) this.cmbGrupo.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Select idGrupoCliente, Grupo From ClienteGrupo Where Proveedor = '" + TipoCliente + "' And idGrupoCliente Not In (2, 3) Order by Grupo ");
			this.txtCodigo.Enabled = false;
			txtRuc.Focus();
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
			ultraTabPageControl1.Focus();
			ultraTabPageControl1.Show();
			txtNombre.Focus();
			MessageBox.Show("Copia de Clientes Terminada", "Información");
		}

		private void barraDato1_Grabar(object sender, System.EventArgs e)
		{
			if (!Verifica()) 
			{
				ultraTabPageControl1.Focus();
				ultraTabPageControl1.Show();
				return;
			}

			#region Crédito
//			if ((int)drCliente["ContadoCred"] == 2)
//			{
//				int iGrupoCliente = (int)drCliente["idGrupoCliente"];
//				if (iGrupoCliente != 2 && iGrupoCliente != 3)
//				{
//					//					if (iGrupoCliente != 3)
//					//					{
//					if (!TipoCliente)
//					{
//						//							double MontoCredito = (double)drCliente["MaximoCredito"];
//						//							if (MontoCredito == 0)
//						//							{
//						////								errorProvider.SetError(this.tbBasico, "Error en la pestaña Complemento");
//						////								errorProvider.SetError(this.txtValorCredito, "Ingrese el Valor de crédito");
//						//								MessageBox.Show("Escriba el cupo de crédito", "Point Technology - Error en la pestaña Complemento", MessageBoxButtons.OK, MessageBoxIcon.Error);
//						//								this.tbBasico.Tabs[1].Selected = true;
//						//								this.txtValorCredito.Focus();
//						//								return;
//						//							}
//					}
//		
//					int Dias1 = (int)drCliente["DiasCredito"];
//					int Dias2 = (int)drCliente["DiasCredito1"];
//					int Dias3 = (int)drCliente["DiasCredito2"];
//
//					if (Dias1 < 1)
//					{
//						//							errorProvider.SetError(this.tbBasico, "Error en la pestaña Complemento");
//						//							errorProvider.SetError(this.ultraNumericEditor1, "Ingrese los dias de crédito");
//						MessageBox.Show("Ingrese los dias de crédito", "Point Technology - Error en la pestaña Complemento", MessageBoxButtons.OK, MessageBoxIcon.Error);
//						this.tbBasico.Tabs[1].Selected = true;
//						this.txtDiasCredito1.Focus();
//						return;
//					}
//
//					if ((Dias2 != 0 && Dias2 < Dias1))
//					{
//						//							errorProvider.SetError(this.tbBasico, "Error en la pestaña Complemento");
//						//							errorProvider.SetError(this.ultraNumericEditor4, "El campo dias de crédito 2 no puede ser menor a dias de crédito 1");
//						MessageBox.Show("El campo dias de crédito 2 no puede ser menor a dias de crédito 1", "Point Technology - Error en la pestaña Complemento", MessageBoxButtons.OK, MessageBoxIcon.Error);
//						this.tbBasico.Tabs[1].Selected = true;
//						this.txtDiasCredito1.Focus();
//						return;
//					}
//
//					if ((Dias3 != 0 && Dias3 < Dias2))
//					{
//						//							errorProvider.SetError(this.tbBasico, "Error en la pestaña Complemento");
//						//							errorProvider.SetError(this.ultraNumericEditor4, "El campo dias de crédito 3 no puede ser menor a dias de crédito 2");
//						MessageBox.Show("El campo dias de crédito 3 no puede ser menor a dias de crédito 2", "Point Technology - Error en la pestaña Complemento", MessageBoxButtons.OK, MessageBoxIcon.Error);
//						this.tbBasico.Tabs[1].Selected = true;
//						this.txtDiasCredito1.Focus();
//						return;
//					}
//				}
//				//				}
//			}
			#endregion Crédito

			#region Contactos
			//if (!TipoCliente)  // Caso de proveedor
			//{ 
			int iContContactos = this.ultraGrid2.Rows.Count;
			if (iContContactos < 1) 
			{
				MessageBox.Show("Debe ingresar un contacto en la Ficha Contactos", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
				this.tbBasico.Tabs[2].Selected = true;
				return;
			}
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in ultraGrid2.Rows.All)
			{
				if (dr.Cells["Contacto"].Value == System.DBNull.Value || dr.Cells["Contacto"].Value.ToString() == "") 
				{
					MessageBox.Show("Ingrese un nombre de contacto en la Ficha Contactos", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
					this.tbBasico.Tabs[2].Selected = true;
					this.ultraGrid2.ActiveRow = dr;
					this.ultraGrid2.ActiveRow.Selected = true;
					return;
				}

				if ((dr.Cells["Telefono"].Value == System.DBNull.Value || dr.Cells["Telefono"].Value.ToString() == "") && (dr.Cells["Celular"].Value == System.DBNull.Value || dr.Cells["Celular"].Value.ToString() == ""))
				{
					MessageBox.Show("Ingrese un número de contacto en la Ficha Contactos", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
					this.tbBasico.Tabs[2].Selected = true;
					this.ultraGrid2.ActiveRow = dr;
					this.ultraGrid2.ActiveRow.Selected = true;
					return;
				}

				if ((bool)dr.Cells["EnvioMail"].Value)
				{
					if (dr.Cells["Email"].Value == System.DBNull.Value || dr.Cells["Email"].Value.ToString() == "") 
					{
						MessageBox.Show("Ingrese el E-Mail del contacto " + dr.Cells["Contacto"].Value.ToString() + ".", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
						this.tbBasico.Tabs[2].Selected = true;
						this.ultraGrid2.ActiveRow = dr;
						this.ultraGrid2.ActiveRow.Selected = true;
						return;
					}
				}
			}
			//}
			#endregion Contactos
			
			//drCliente["Codigo"] = this.txtRuc.Value.ToString();// drCliente["RUC"].ToString();

			bool bNuevo = barraDato1.bNuevo;
			
			barraDato1.GrabaRegistro();

			Funcion.EjecutaSQL(cdsSeteoF, string.Format("Update Cliente Set Codigo = '{0}' Where idCliente = {1}", (string)drCliente["Ruc"], (int)drCliente["idCliente"]), true);

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
			cdsDirecciones.Update();

			if (bNuevo && bCodAuto)
			{
				string stExec = "Exec FacturaNumeroEstacion " + idTipoF;
				string stCodigo = Funcion.sEscalarSQL(cdsCliente, stExec, true);
				if (stCodigo.Trim().Length > 0)
				{
					stExec = "Update Cliente Set Codigo = '" + stCodigo + "' Where idCliente = " + txtIdCliente.Text;
					Funcion.EjecutaSQL(cdsCliente, stExec, true);
					txtCodigo.Value = stCodigo;
				}
			}

			if (!TipoCliente) this.btProveedorCliente.Enabled = true;

			HabilitaTab();

			ProveedorInternacional();

			this.rdbContado.Enabled = false;
			this.rdbCredito.Enabled = false;
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

			#region Representante y Nombre Comercial
			int idTipoRuc = (int) this.cmbTipoRuc.Value;
			if (idTipoRuc == 1 || idTipoRuc == 4)
			{
				#region Representante
				if (txtRepresentante.Text.ToString().Trim().Length == 0)
				{
					errorProvider.SetError(this.txtRepresentante, "Ingrese el nombre del representante");
					bOk = false;
				}
				#endregion Representante

				#region Nombre Comercial
				if (this.txtNombreComercial.Text.ToString().Trim().Length == 0)
				{
					errorProvider.SetError(this.txtNombreComercial, "Ingrese el nombre comercial");
					bOk = false;
				}
				#endregion Nombre Comercial
			}
			#endregion Representante y Nombre Comercial

			#region Verifica Nombre
			if (txtNombre.Text.ToString().Trim().Length == 0)
			{
				errorProvider.SetError(txtNombre, "Ingrese el Nombre comercial");
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
			
			#region Direccion
			if (this.txtDireccion.Text.ToString().Trim().Length == 0)
			{
				errorProvider.SetError(this.txtDireccion, "Ingrese la dirección");
				bOk = false;
			}
			#endregion Direccion 

			#region Teléfono
			if (txtTelefono.Text.ToString().Trim() == "")
			{
				errorProvider.SetError(this.txtTelefono, "Ingrese el número de teléfono");
				bOk = false;
			}
			if ((int)drCliente["idTipoRuc"] != 3) 
			{
				if ((int)drCliente["idTipoRuc"] != 6)
				{
					if (txtTelefono.TextLength < 7 || this.txtTelefono.TextLength > 9)
					{
						errorProvider.SetError(this.txtTelefono, "Teléfono no válido");
						bOk = false;
					}
				}
			}
			#endregion Teléfono

			#region Fax
			if ((string)this.txtFax.Text.ToString().Trim() != "")
			{
				if ((int)drCliente["idTipoRuc"] != 3)
				{
					if ((int)drCliente["idTipoRuc"] != 6)
					{
						if (txtFax.TextLength < 7 || this.txtFax.TextLength > 9)
						{
							errorProvider.SetError(this.txtFax, "FAX no válido");
							bOk = false;
						}
					}
				}
			}
			#endregion Fax

			#region Celular
			if ((int) this.txtCelular.TextLength > 0)
			{
				if ((int)this.cmbTipoRuc.Value != 3) // || (int)this.cmbTipoRuc.Value != 6)
				{
					if (this.txtCelular.TextLength != 10)
					{
						errorProvider.SetError(this.txtCelular, "Celular no válido");
						bOk = false;
					}
				}
			}
			#endregion Celular

			#region Grupo
			if (cmbGrupo.Value == DBNull.Value || (int) cmbGrupo.Value == 0)
			{
				errorProvider.SetError(cmbGrupo, "Grupo de cliente no puede estar vacio");
				bOk = false;
			}
			#endregion Grupo

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
					if (this.barraDato1.bNuevo)
					{
						string stSelect = "SELECT COUNT(*) FROM Cliente WHERE Ruc = '" + txtRuc.Text.ToString().Trim() + "'" + stProv + " And idCliente <> " + txtIdCliente.Text;

						int iCuenta = Funcion.iEscalarSQL(cdsSeteoF, stSelect);
						if (iCuenta > 0) 
						{
							errorProvider.SetError(txtRuc,"RUC ya existe");
							bOk = false;
						}
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

			#region Tipo de proveedor 
			if (!TipoCliente && (int)this.cmbTipoRuc.Value == 3 && this.cmbTipoProveedor.Text.ToString().Length == 0)
			{
				errorProvider.SetError(this.cmbTipoProveedor, "Para proveedor con pasaporte debe seleccionar el tipo de proveedor.");
				bOk = false;
			}
			#endregion Tipo de proveedor 
			return bOk;
		}

		private void barraDato1_Editar(object sender, System.EventArgs e)
		{
			if (drCliente == null)
			{
				MessageBox.Show("No existe un registro para editar", "Información", MessageBoxButtons.OK, MessageBoxIcon.Information);
				return;
			}

			barraDato1.EditarRegistro();

			HabilitaTab();

			if (this.cmbGrupo.Value != System.DBNull.Value)
			{
				if ((int)this.cmbGrupo.Value == 2 || (int)this.cmbGrupo.Value == 3)
				{
					this.chkPublica.Enabled = false;
					this.chkContribuyenteEsp.Enabled = false;
					this.txtCelular.MaxLength = 30;
					this.txtTelefono.MaxLength = 30;
					this.txtFax.MaxLength = 30;
				}
				else
				{
					this.chkPublica.Enabled = true;
					this.chkContribuyenteEsp.Enabled = true;
					this.txtCelular.MaxLength = 10;
					this.txtTelefono.MaxLength = 9;
					this.txtFax.MaxLength = 9;
				}
			}

			if (TipoCliente)
			{
				this.btProveedorCliente.Enabled = false;
				this.radioGroup1.Enabled = false;
				//if ((int) this.radioGroup1.EditValue == 2) this.cmbTipoCredito.Enabled = true;
				//else this.cmbTipoCredito.Enabled = false;
			}
			else
			{
				this.radioGroup1.Enabled = false;
				if (this.chkContribuyenteEsp.Checked) this.cmbPorcIvaRet.Enabled = false;
				else this.cmbPorcIvaRet.Enabled = true;
			}

			if (Funcion.iEscalarSQL(cdsSeteoF, string.Format("Exec ClienteValidaFacturas {0}, {1}", TipoCliente, (int)drCliente["idCliente"])) > 0)
			{
				this.cmbTipoRuc.Enabled = false;
				this.txtRuc.Enabled = false;
			}		
	
			if ((int)drCliente["ContadoCred"] == 1)
			{
				this.txtValorCredito.Enabled = false;
				this.txtDiasCredito1.Enabled = false;
				this.txtDiasCredito2.Enabled = false;
				this.txtDiasCredito3.Enabled = false;
			}
			else if ((int)drCliente["ContadoCred"] == 2)
			{
				this.txtValorCredito.Enabled = true;
				this.txtDiasCredito1.Enabled = true;
				this.txtDiasCredito2.Enabled = true;
				this.txtDiasCredito3.Enabled = true;
			}

			this.rdbContado.Enabled = true;
			this.rdbCredito.Enabled = true;
			this.txtCodigo.Enabled = false;
		}

		private void barraDato1_Deshacer(object sender, System.EventArgs e)
		{
			this.barraDato1.CausesValidation = false;
			BorrarErrores();
			barraDato1.DeshacerRegistro();
			if (!TipoCliente) this.btProveedorCliente.Enabled = true;

			if (!TipoCliente)
			{
				if (this.cmbTipoRuc.Value != System.DBNull.Value)
				{
					if ((int)this.cmbTipoRuc.Value == 3) 
					{
						cmbGrupo.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Select idGrupoCliente, Grupo From ClienteGrupo Where Proveedor = '" + TipoCliente + "' And idGrupoCliente In (2, 3) Order by Grupo ");

						ObjetosChina(true);
					}
					else
					{
						cmbGrupo.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Select idGrupoCliente, Grupo From ClienteGrupo Where Proveedor = '" + TipoCliente + "' And idGrupoCliente Not In (2, 3) Order by Grupo ");

						ObjetosChina(false);
					}
				}
			}		

			HabilitaTab();
			
			this.rdbContado.Enabled = false;
			this.rdbCredito.Enabled = false;
		}

		private void BorrarErrores()
		{
			errorProvider.SetError(txtNombre, "");
			errorProvider.SetError(txtRuc, "");
			errorProvider.SetError(cmbTipoRuc, "");
			errorProvider.SetError(txtAutFactura, "");
			errorProvider.SetError(txtSerie, "");
			errorProvider.SetError(this.txtTelefono, "");
			errorProvider.SetError(this.txtDireccion, "");
			errorProvider.SetError(this.txtTelefono, "");
			errorProvider.SetError(this.txtFax, "");
			errorProvider.SetError(this.txtCelular, "");
			errorProvider.SetError(this.txtNombreComercial, "");
			errorProvider.SetError(this.txtRepresentante, "");
			errorProvider.SetError(this.cmbGrupo, "");
			errorProvider.SetError(this.tbBasico, "");
			errorProvider.SetError(this.txtValorCredito, "");
			errorProvider.SetError(this.txtDiasCredito1, "");
			errorProvider.SetError(this.cmbTipoProveedor, "");
		}

		private void HabilitaTab()
		{
			txtAutFactura.Enabled = true;
			foreach(Infragistics.Win.UltraWinTabControl.UltraTab ctr in tbBasico.Tabs)
				foreach(Control cn in ctr.TabPage.Controls)
				{
					barraDato1.HabilitaControl(cn, barraDato1.bEditar);
					this.dtHoraDesde.Enabled = barraDato1.bEditar;
					this.dtHoraHasta.Enabled = barraDato1.bEditar;
				}
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
			C1.Data.FilterCondition fcContactos = new C1.Data.FilterCondition(cdsContacto, "ClienteContacto", "idCliente = " + txtIdCliente.Value.ToString());
			C1.Data.FilterConditions fcTotal = new C1.Data.FilterConditions();
			fcTotal.Add(fcContactos);
			cdsContacto.Fill(fcTotal, false);

			C1.Data.FilterCondition fcDirecciones = new C1.Data.FilterCondition(cdsDirecciones, "ClientesDirecciones", "idCliente = " + txtIdCliente.Value.ToString());
			C1.Data.FilterConditions fcTot = new C1.Data.FilterConditions();
			fcTot.Add(fcDirecciones);
			cdsDirecciones.Fill(fcTot, false);

			if (!this.barraDato1.bEditar)
			{
				if ((int) drCliente["ContadoCred"] == 1) this.rdbContado.Checked = true;
				else if ((int) drCliente["ContadoCred"] == 2) this.rdbCredito.Checked = true;
			}	

			if (!this.barraDato1.bNuevo)
			{
				if (!TipoCliente)
				{
					if ((int)this.cmbTipoRuc.Value == 3) 
					{
						cmbGrupo.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Select idGrupoCliente, Grupo From ClienteGrupo Where Proveedor = '" + TipoCliente + "' And idGrupoCliente In (2, 3) Order by Grupo ");

						ObjetosChina(true);
					}
					else
					{
						cmbGrupo.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Select idGrupoCliente, Grupo From ClienteGrupo Where Proveedor = '" + TipoCliente + "' And idGrupoCliente Not In (2, 3) Order by Grupo ");

						ObjetosChina(false);
					}
				}
			}
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
			ClienteGrupo miGrupoC = new ClienteGrupo(TipoCliente);
			miGrupoC.ShowDialog();
			string sSQL = ""; 
			if (TipoCliente) sSQL = string.Format("Select idGrupoCliente, Grupo From ClienteGrupo Where Proveedor = 1");
			else sSQL = string.Format("Select idGrupoCliente, Grupo From ClienteGrupo Where Proveedor = 0");
			cmbGrupo.DataSource = Funcion.dvProcedimiento(cdsSeteoF, sSQL);
		}

		private void cmbGrupo_MouseDown(object sender, System.Windows.Forms.MouseEventArgs e)
		{
			if (e.Button != MouseButtons.Right) return;
			cmbGrupo.Value = 0;
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
//			miCompra = new Compra(-IdCompra);
//			miCompra.MdiParent = MdiParent;	
//			miCompra.Show();
		}

		private void btFactura_MouseUp(object sender, System.Windows.Forms.MouseEventArgs e)
		{
//			if (e.Button != MouseButtons.Right) return;
//			if (ultraGrid1.ActiveRow == null ||	ultraGrid1.ActiveRow.Cells["idCompra"].Value == DBNull.Value)
//			{
//				MessageBox.Show("Seleccione la factura que desea abrir", "Información",
//					MessageBoxButtons.OK, MessageBoxIcon.Information);
//				return;
//			}
//			Cursor = Cursors.WaitCursor;
//			int IdCompra = (int) ultraGrid1.ActiveRow.Cells["idCompra"].Value;
//			miCompra.BuscarCompra(IdCompra);
//			miCompra.Show();
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
			using (ClientePais miPais = new ClientePais())
			{
				miPais.ShowDialog();
			}
			cmbPais.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Select idClientePais, Pais From ClientePais");
		}

		private void btCedRep_Click(object sender, System.EventArgs e)
		{
			string stMensaje = Funcion.VerificaCedula(txtCedRep.Text.Trim(), true);
			MessageBox.Show(stMensaje, "Informacion", MessageBoxButtons.OK, MessageBoxIcon.Information);
		}
		#region Personalizacion Presupuesto
		private void llenaComboCuentaAnticipo()
		{			 
			//combo Cuenta
			string stSelect ="select idCuenta,CodRapido,Descripcion from cuenta where grupo = 0 order by codigo";
			cmbCuentaAnticipo.DataSource = Funcion.dvProcedimiento(cdsCliente,stSelect);			
			cmbCuentaAnticipo.DisplayMember = "CodRapido";
			cmbCuentaAnticipo.ValueMember = "idCuenta";
			cmbCuentaAnticipo.Rows.Band.Columns["idCuenta"].Hidden = true;
			//llena el valor predefinido en la tabla cuenta
			stSelect = "Select count(idCuentaAnticipo) from cliente where idCliente = " + IdCliente;
			if (Funcion.iEscalarSQL(cdsCliente, stSelect) > 0)
			{
				stSelect = "Select isnull(idCuentaAnticipo,0) from cliente where idCliente = " + IdCliente;
				int idCuenta = Funcion.iEscalarSQL(cdsCliente, stSelect);
				cmbCuentaAnticipo.Value = idCuenta;
			}
		}
		private void llenaComboCuentaFondo()
		{			 
			//combo Cuenta
			string stSelect ="select idCuenta,CodRapido,Descripcion from cuenta where grupo = 0 order by codigo";
			cmbCuentaFondo.DataSource = Funcion.dvProcedimiento(cdsCliente,stSelect);			
			cmbCuentaFondo.DisplayMember = "CodRapido";
			cmbCuentaFondo.ValueMember = "idCuenta";
			cmbCuentaFondo.Rows.Band.Columns["idCuenta"].Hidden = true;
			//llena el valor predefinido en la tabla cuenta
			stSelect = "Select count(idCuentaFondo) from cliente where idCliente = " + IdCliente;
			if (Funcion.iEscalarSQL(cdsCliente, stSelect) > 0)
			{
				stSelect = "Select isnull(idCuentaFondo,0) from cliente where idCliente = " + IdCliente;
				int idCuenta = Funcion.iEscalarSQL(cdsCliente, stSelect);
				cmbCuentaFondo.Value = idCuenta;
			}
		}

		private void cmbCuentaAnticipo_MouseDown(object sender, System.Windows.Forms.MouseEventArgs e)
		{
			if (e.Button != MouseButtons.Right)return;			 
			string stPrimero = "Descripcion";
			if (cmbCuentaAnticipo.DisplayLayout.Bands[0].Columns[stPrimero].Header.VisiblePosition == 0)
				stPrimero = "CodRapido";			
			cmbCuentaAnticipo.DisplayLayout.Bands[0].Columns[stPrimero].Header.VisiblePosition = 0;
			cmbCuentaAnticipo.DisplayMember = stPrimero;
		}

		private void cmbCuentaFondo_MouseDown(object sender, System.Windows.Forms.MouseEventArgs e)
		{
			if (e.Button != MouseButtons.Right)return;			 
			string stPrimero = "Descripcion";
			if (cmbCuentaFondo.DisplayLayout.Bands[0].Columns[stPrimero].Header.VisiblePosition == 0)
				stPrimero = "CodRapido";			
			cmbCuentaFondo.DisplayLayout.Bands[0].Columns[stPrimero].Header.VisiblePosition = 0;
			cmbCuentaFondo.DisplayMember = stPrimero;
		}
		#endregion Personalizacion Presupuesto

		private void barraDato1_Load(object sender, System.EventArgs e)
		{
		
		}

		private void ultraTabPageControl1_Paint(object sender, System.Windows.Forms.PaintEventArgs e)
		{
		
		}

		private void txtRuc_Leave(object sender, System.EventArgs e)
		{

		}

		private void ultraButton2_Click(object sender, System.EventArgs e)
		{
			using(POINT_TipoCredito miTPCred = new POINT_TipoCredito())
			{
				miTPCred.ShowDialog();
			}
			//cmbSector.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Select idSector, Sector From ClienteSector");
		}

		private void btProveedorCliente_Click(object sender, System.EventArgs e)
		{
			if (this.barraDato1.bEditar)
			{
				MessageBox.Show("No puede crear como cliente en Edición.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
				return;
			}
			string sSQL = string.Format("Exec ClienteProveedor {0}, '{1}', {2}", this.txtIdCliente.Value, this.txtRuc.Text, (int)drCliente["idTipoRuc"]);
			int iRes = Funcion.iEscalarSQL(cdsSeteoF, sSQL);

			if (iRes > 0)
				MessageBox.Show("Creo al proveedor " + this.txtNombre.Text.Trim() + " como cliente.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
			else
				MessageBox.Show("El proveedor " + this.txtNombre.Text.Trim() + " ya existe como cliente.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
		}

		private void txtRuc_Validating(object sender, System.ComponentModel.CancelEventArgs e)
		{
			if (txtRuc.Text.ToString().Trim() != "")
			{
				#region 13 digitos en ruc cmbTipoRuc.Value 
				//				if ((int)drCliente["idTipoRuc"] == 1 || (int)drCliente["idTipoRuc"] == 4)
				//				{
				//					if (txtRuc.Text.Trim().Length != 13)
				//					{
				//						MessageBox.Show("RUC debe tener 13 digitos", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
				////						this.txtRuc.Focus();
				//						//errorProvider.SetError(txtRuc, "RUC debe tener 13 digitos");
				//						e.Cancel = true;
				//						return;
				//					}
				//				}
				#endregion 13 digitos en ruc
				
				#region Valida Documento de identidad
				if ((int)this.cmbTipoRuc.Value != 3 || (int)this.cmbTipoRuc.Value != 6)
				{
					string stExec = string.Format("Select dbo.ValidaIdentificacion ('{0}', {1})", txtRuc.Text, (int)this.cmbTipoRuc.Value);
					string VlCed = Funcion.sEscalarSQL(cdsCliente, stExec);
					if (VlCed.StartsWith("Error"))
					{
						MessageBox.Show(VlCed, "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
						e.Cancel=true;
						return;
					}
				}
				#endregion Valida Documento de identidad
			}
			
			if (this.barraDato1.bNuevo)
			{
				if (txtRuc.Text.ToString().Trim() != "")
				{
					string stProv = "";
					if (TipoCliente) stProv = " And (Proveedor = 0) ";
					else stProv = " And (Proveedor = 1) ";

					string stSelect = string.Format("SELECT COUNT(*) FROM Cliente WHERE Ruc = '{0}' {1} And idCliente <> {2} And idTipoRuc = {3}", 
						this.txtRuc.Text.ToString().Trim(), stProv, txtIdCliente.Text, (int)this.cmbTipoRuc.Value);//"SELECT COUNT(*) FROM Cliente WHERE (Ruc = '" + this.txtRuc.Text.ToString().Trim() + "' Or Ruc = '" + this.txtRuc.Text.ToString().Trim().Substring(1, 10) + "')" + stProv + " And idCliente <> " + txtIdCliente.Text;
					int iCuenta = Funcion.iEscalarSQL(cdsSeteoF, stSelect);
					if (iCuenta > 0) 
					{
						MessageBox.Show("Número de Identificación ya existe", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
						string sSQL = string.Format("SELECT idCliente FROM Cliente WHERE Ruc = '{0}' {1} And idCliente <> {2} And idTipoRuc = {3}", 
							this.txtRuc.Text.ToString().Trim(), stProv, txtIdCliente.Text, (int)this.cmbTipoRuc.Value);
						//"SELECT idCliente FROM Cliente WHERE Ruc = '" + this.txtRuc.Text.ToString().Trim() + "'" + stProv + " And idCliente <> " + txtIdCliente.Text;
						int IdCliente = Funcion.iEscalarSQL(cdsSeteoF, sSQL);
						this.barraDato1_Deshacer(sender, e);
						ClienteBuscado(IdCliente);
						this.barraDato1_Editar(sender, e);

						this.txtRuc.Focus();
						e.Cancel = true;
						return;
					}
				}
			}
		}

		private void txtRuc_ValueChanged(object sender, System.EventArgs e)
		{
		
		}

		private void ultraNumericEditor1_ValueChanged(object sender, System.EventArgs e)
		{
			
		}

		private void chkContribuyenteEsp_CheckedChanged(object sender, System.EventArgs e)
		{
			if (!this.barraDato1.bEditar) return;
			if (!TipoCliente)
			{
				if (this.chkContribuyenteEsp.Checked)
				{
					this.cmbPorcIvaRet.Value = 1;
					this.cmbPorcIvaRet.Enabled = false;
				}
				else
					this.cmbPorcIvaRet.Enabled = true;
			}
		}

		private void radioGroup1_SelectedIndexChanged(object sender, System.EventArgs e)
		{


		}

		private void ultraNumericEditor1_Validated(object sender, System.EventArgs e)
		{
			
		}

		private void txtRuc_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (e.KeyChar == 13) this.txtNombre.Focus();
		}

		private void txtRuc_Validated(object sender, System.EventArgs e)
		{
		
		}

		private void txtRuc_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Return || e.KeyCode == Keys.Tab)
			{
				
			}
		}

		private void txtValorCredito_ValueChanged(object sender, System.EventArgs e)
		{
		
		}

		private void txtTelefono_ValueChanged(object sender, System.EventArgs e)
		{
		
		}

		private void btVerifica_Click(object sender, System.EventArgs e)
		{
			string stMensaje = Funcion.VerificaCedula(this.txtRuc.Text.Trim(), true);
			MessageBox.Show(stMensaje, "Información", MessageBoxButtons.OK, MessageBoxIcon.Information);
		}

		private void rdbContado_CheckedChanged(object sender, System.EventArgs e)
		{
			if (this.rdbContado.Checked)
			{
				drCliente["ContadoCred"] = 1;

				drCliente["MaximoCredito"] = 0;
				drCliente["DiasCredito"] = 0;
				drCliente["DiasCredito1"] = 0;	
				drCliente["DiasCredito2"] = 0;

				this.txtValorCredito.Enabled = false;
				this.txtDiasCredito1.Enabled = false;
				this.txtDiasCredito2.Enabled = false;
				this.txtDiasCredito3.Enabled = false;
			}
		}

		private void rdbCredito_CheckedChanged(object sender, System.EventArgs e)
		{
			if (this.rdbCredito.Checked)
			{
				drCliente["ContadoCred"] = 2;
				
				if (this.barraDato1.bEditar)
				{
					this.txtValorCredito.Enabled = true;
					this.txtDiasCredito1.Enabled = true;
					this.txtDiasCredito2.Enabled = true;
					this.txtDiasCredito3.Enabled = true;
				}
			}
		}

		private void cmbGrupo_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}

		private void cmbGrupo_Validated(object sender, System.EventArgs e)
		{
			if (this.cmbGrupo.Value == System.DBNull.Value) return;
			if ((int)this.cmbGrupo.Value == 2 || (int)this.cmbGrupo.Value == 3)
			{
				this.chkPublica.Enabled = false;
				this.chkContribuyenteEsp.Enabled = false;
				drCliente["InstPublica"] = false;
				drCliente["ContribEspecial"] = false;
				this.txtTelefono.MaxLength = 15;
				this.txtFax.MaxLength = 15;
				this.txtCelular.MaxLength = 15;
				this.cmbCredTribut.Enabled = false;
				this.cmbPorcIvaRet.Enabled = false;				
			}
			else if ((int) this.cmbGrupo.Value == 1)
			{
				this.chkPublica.Enabled = true;
				this.chkContribuyenteEsp.Enabled = true;
				drCliente["idPais"] = 1;
				drCliente["idCiudad"] = 1;
				this.txtTelefono.MaxLength = 10;
				this.txtFax.MaxLength = 10;
				this.txtCelular.MaxLength = 10;
				this.cmbCredTribut.Enabled = true;
				this.cmbPorcIvaRet.Enabled = true;
			}
			else
			{
				drCliente["idPais"] = 1;
				drCliente["idCiudad"] = 1;
				this.chkPublica.Enabled = true;
				this.chkContribuyenteEsp.Enabled = true;
				this.chkPublica.Enabled = true;
				this.chkContribuyenteEsp.Enabled = true;
				this.txtCelular.MaxLength = 10;
				this.txtTelefono.MaxLength = 9;
				this.txtFax.MaxLength = 9;
				this.cmbCredTribut.Enabled = true;
				this.cmbPorcIvaRet.Enabled = true;
			}
		}

		private void txtValorCredito_Enter(object sender, System.EventArgs e)
		{
			txtValorCredito.SelectAll();
		}

		private void Cliente_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Escape) 
			{
				if (this.barraDato1.bEditar)
					this.barraDato1_Deshacer(sender, e);
				else
					this.Close();
			}
			
			if (e.KeyCode == Keys.N)
			{
				if (!this.barraDato1.bEditar)
				{
					this.barraDato1_Crear(sender, e);
				}
			}
			if (e.KeyCode == Keys.E)
			{
				if (!this.barraDato1.bEditar)
				{
					this.barraDato1_Editar(sender, e);
				}
			}
			if (e.KeyCode == Keys.B)
			{
				if (!this.barraDato1.bEditar)
				{
					this.barraDato1_Buscar(sender, e);
				}
			}
			if (e.KeyCode == Keys.I)
			{
				if (!this.barraDato1.bEditar)
				{
					this.barraDato1_Impresora(sender, e);
				}
			}
			if (e.KeyCode == Keys.Delete)
			{
				if (!this.barraDato1.bEditar)
				{
					this.barraDato1_Borrar(sender, e);
				}
			}
			if (e.KeyCode == Keys.F12)
			{
				if (this.barraDato1.bEditar)
				{
					this.barraDato1_Grabar(sender, e);
				}
			}
			if (e.KeyCode == Keys.F5)
			{
				if (!this.barraDato1.bEditar)
				{
					barraDato1.IdRegistro = (int)drCliente["idCliente"];
					barraDato1.ProximoId(5);					
				}
			}			
		}

		private void tbBasico_SelectedTabChanged(object sender, Infragistics.Win.UltraWinTabControl.SelectedTabChangedEventArgs e)
		{
		
		}

		private void txtTelefono_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			Funcion.CamposNumericos(sender, e);
			if (e.KeyChar == 13) this.txtFax.Focus();
		}

		private void txtFax_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			Funcion.CamposNumericos(sender, e);
			if (e.KeyChar == 13) this.txtCelular.Focus();
		}

		private void txtCelular_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			Funcion.CamposNumericos(sender, e);
			if (e.KeyChar == 13) this.txtEMail.Focus();
		}

		private void tpComplemento_Paint(object sender, System.Windows.Forms.PaintEventArgs e)
		{
		
		}

		private void txtValorCredito_Click(object sender, System.EventArgs e)
		{
			txtValorCredito.SelectAll();
		}

		private void txtDiasCredito1_Click(object sender, System.EventArgs e)
		{
			this.txtDiasCredito1.SelectAll();
		}

		private void txtDiasCredito1_Enter(object sender, System.EventArgs e)
		{
			this.txtDiasCredito1.SelectAll();
		}

		private void txtDiasCredito2_Click(object sender, System.EventArgs e)
		{
			this.txtDiasCredito2.SelectAll();
		}

		private void txtDiasCredito2_Enter(object sender, System.EventArgs e)
		{
			this.txtDiasCredito2.SelectAll();
		}

		private void txtDiasCredito3_Click(object sender, System.EventArgs e)
		{
			this.txtDiasCredito3.SelectAll();
		}

		private void txtDiasCredito3_Enter(object sender, System.EventArgs e)
		{
			this.txtDiasCredito3.SelectAll();
		}

		private void spnDescuento_Click(object sender, System.EventArgs e)
		{
			this.spnDescuento.SelectAll();
		}

		private void spnDescuento_Enter(object sender, System.EventArgs e)
		{
			this.spnDescuento.SelectAll();
		}

		private void spnDescCredito_Click(object sender, System.EventArgs e)
		{
			this.spnDescCredito.SelectAll();
		}

		private void spnDescCredito_Enter(object sender, System.EventArgs e)
		{
			this.spnDescCredito.SelectAll();
		}

		private void ultraGrid2_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (ultraGrid2.ActiveCell.Column.Key.ToString().ToUpper() == "TELEFONO" || ultraGrid2.ActiveCell.Column.Key.ToString().ToUpper() == "CELULAR")
				Funcion.CamposNumericos(sender, e);
		}

		private void ultraGrid2_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			switch(e.KeyValue)
			{
					#region Flechas
				case (int) Keys.Up:
					if (ultraGrid2.ActiveCell.DroppedDown) return;

					ultraGrid2.PerformAction(UltraGridAction.ExitEditMode, false, false);
					ultraGrid2.PerformAction(UltraGridAction.AboveCell, false, false);
					e.Handled = true;
					ultraGrid2.PerformAction(UltraGridAction.EnterEditMode, false, false);
					break;
				case (int) Keys.Left:
					ultraGrid2.PerformAction(UltraGridAction.ExitEditMode, false, false);
					ultraGrid2.PerformAction(UltraGridAction.PrevCellByTab, false, false);
					e.Handled = true;
					ultraGrid2.PerformAction(UltraGridAction.EnterEditMode, false, false);
					break;
				case (int) Keys.Down:
					if (ultraGrid2.ActiveCell.DroppedDown) return;
			
					ultraGrid2.PerformAction(UltraGridAction.ExitEditMode, false, false);
					ultraGrid2.PerformAction(UltraGridAction.BelowCell, false, false);
					e.Handled = true;
					ultraGrid2.PerformAction(UltraGridAction.EnterEditMode, false, false);
					break;
					#endregion Flechas
					#region Enter
				case (int) Keys.Enter:
					ultraGrid2.PerformAction(UltraGridAction.ExitEditMode, false, false);
					ultraGrid2.PerformAction(UltraGridAction.NextCellByTab, false, false);
					e.Handled = true;
					ultraGrid2.PerformAction(UltraGridAction.EnterEditMode, false, false);
					if(ultraGrid2.ActiveCell.DroppedDown == false)
						ultraGrid2.ActiveCell.DroppedDown = true;
					//						ultraGrid2.PerformAction(UltraGridAction.EnterEditModeAndDropdown, false, false);
					break;
					#endregion Enter
			}
		}

		private void txtCelular_ValueChanged(object sender, System.EventArgs e)
		{
		
		}

		private void cmbTipoRuc_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}

		private void cmbTipoRuc_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (e.KeyChar == 13)
			{
				this.txtRuc.Focus();
			}
		}

		private void ProveedorInternacional()
		{
			
		}

		private void cmbTipoRuc_Validated(object sender, System.EventArgs e)
		{
			if (!TipoCliente)
			{
				if ((int)this.cmbTipoRuc.Value == 3) 
				{
					cmbGrupo.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Select idGrupoCliente, Grupo From ClienteGrupo Where Proveedor = '" + TipoCliente + "' And idGrupoCliente In (2, 3) Order by Grupo ");
					drCliente["idGrupoCliente"] = 0;
					drCliente["idTipoProveedor"] = 2;

					this.ultraGrid2.DisplayLayout.Bands[0].Columns["Telefono"].FieldLen = 15;
					this.ultraGrid2.DisplayLayout.Bands[0].Columns["Celular"].FieldLen = 15;
					ObjetosChina(true);
				}
				else
				{
					cmbGrupo.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Select idGrupoCliente, Grupo From ClienteGrupo Where Proveedor = '" + TipoCliente + "' And idGrupoCliente Not In (2, 3) Order by Grupo ");
					drCliente["idGrupoCliente"] = 0;
					drCliente["idTipoProveedor"] = System.DBNull.Value;

					this.ultraGrid2.DisplayLayout.Bands[0].Columns["Telefono"].FieldLen = 9;
					this.ultraGrid2.DisplayLayout.Bands[0].Columns["Celular"].FieldLen = 10;

					ObjetosChina(false);
				}
			}		
		}

		private void btBanco_Click(object sender, System.EventArgs e)
		{
			using (Bancos miBanco = new Bancos())
			{
				miBanco.ShowDialog();
			}
			this.cmbBanco.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Select idBanco, Nombre From Banco");
			this.cmbBanco.DisplayLayout.Bands[0].Columns["idBanco"].Hidden = true;
			this.cmbBanco.DisplayLayout.Bands[0].Columns["Nombre"].Width = 168;
		}

		private void txtCuentaSwift_ValueChanged(object sender, System.EventArgs e)
		{
		
		}

		private void txtCuentaSwift_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			Funcion.CamposNumericos(sender, e);
		}

		private void ultraGrid2_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}

		private void ultraGrid3_AfterRowInsert(object sender, Infragistics.Win.UltraWinGrid.RowEventArgs e)
		{
			e.Row.Cells["idCliente"].Value = int.Parse(txtIdCliente.Value.ToString());
		}

		private void ultraGrid3_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			switch(e.KeyValue)
			{
					#region Flechas
				case (int) Keys.Up:
					if (ultraGrid3.ActiveCell.DroppedDown) return;

					ultraGrid3.PerformAction(UltraGridAction.ExitEditMode, false, false);
					ultraGrid3.PerformAction(UltraGridAction.AboveCell, false, false);
					e.Handled = true;
					ultraGrid3.PerformAction(UltraGridAction.EnterEditMode, false, false);
					break;
				case (int) Keys.Left:
					ultraGrid3.PerformAction(UltraGridAction.ExitEditMode, false, false);
					ultraGrid3.PerformAction(UltraGridAction.PrevCellByTab, false, false);
					e.Handled = true;
					ultraGrid3.PerformAction(UltraGridAction.EnterEditMode, false, false);
					break;
				case (int) Keys.Down:
					if (ultraGrid3.ActiveCell.DroppedDown) return;
			
					ultraGrid3.PerformAction(UltraGridAction.ExitEditMode, false, false);
					ultraGrid3.PerformAction(UltraGridAction.BelowCell, false, false);
					e.Handled = true;
					ultraGrid3.PerformAction(UltraGridAction.EnterEditMode, false, false);
					break;
					#endregion Flechas
					#region Enter
				case (int) Keys.Enter:
					ultraGrid3.PerformAction(UltraGridAction.ExitEditMode, false, false);
					ultraGrid3.PerformAction(UltraGridAction.NextCellByTab, false, false);
					e.Handled = true;
					ultraGrid3.PerformAction(UltraGridAction.EnterEditMode, false, false);
					if(ultraGrid3.ActiveCell.DroppedDown == false)
						ultraGrid3.ActiveCell.DroppedDown = true;
					//						ultraGrid3.PerformAction(UltraGridAction.EnterEditModeAndDropdown, false, false);
					break;
					#endregion Enter
			}
		}

		private void cdsDirecciones_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsDirecciones.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);
		}

		private void txtNombre_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (e.KeyChar == 13) this.txtDireccion.Focus();
		}

		private void txtDireccion_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (e.KeyChar == 13) this.cmbGrupo.Focus();
		}

		private void cmbGrupo_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (e.KeyChar == 13) this.txtRepresentante.Focus();
		}

		private void txtNombreComercial_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (e.KeyChar == 13) this.txtTelefono.Focus();
		}

		private void txtRepresentante_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (e.KeyChar == 13) this.txtNombreComercial.Focus();
		}

		private void txtEMail_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (e.KeyChar == 13) this.cmbPais.Focus();
		}

		private void cmbPais_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (e.KeyChar == 13) this.cmbCiudad.Focus();
		}

		private void cmbCiudad_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (e.KeyChar == 13) this.txtObservacion.Focus();
		}

		private void txtObservacion_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (e.KeyChar == 13) this.txtNotas.Focus();
		}

		private void txtNotas_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (e.KeyChar == 13) this.chkAplicaConsignacion.Focus();
		}

		private void chkAplicaConsignacion_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (e.KeyChar == 13) this.txtDiasConsignacion.Focus();
		}

		private void txtDiasConsignacion_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (e.KeyChar == 13) 
			{
				this.tbBasico.Tabs[1].Selected = true;
				this.rdbContado.Focus();				
			}
		}

		private void txtDiasConsignacion_ValueChanged(object sender, System.EventArgs e)
		{
		
		}

		private void rdbContado_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (e.KeyChar == 13) this.rdbCredito.Focus();
		}

		private void rdbCredito_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (e.KeyChar == 13)
			{
				if ((int)drCliente["ContadoCred"] == 2)	this.txtValorCredito.Focus();
				else this.spnDescuento.Focus();
			}
		}
	
		private void txtValorCredito_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (e.KeyChar == 13) this.txtDiasCredito1.Focus();
		}

		private void txtDiasCredito1_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (e.KeyChar == 13) this.txtDiasCredito2.Focus();
		}

		private void txtDiasCredito2_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (e.KeyChar == 13) this.txtDiasCredito3.Focus();
		}

		private void txtDiasCredito3_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (e.KeyChar == 13) this.spnDescuento.Focus();
		}

		private void spnDescCredito_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (e.KeyChar == 13) this.cmbSector.Focus();
		}

		private void spnDescuento_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (e.KeyChar == 13) this.spnDescCredito.Focus();
		}

		private void ultraTextEditor20_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (e.KeyChar == 13) this.ultraTextEditor7.Focus();
		}

		private void ultraTextEditor7_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (e.KeyChar == 13) this.ultraTextEditor15.Focus();
		}

		private void ultraTextEditor15_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (e.KeyChar == 13) this.ultraTextEditor8.Focus();
		}

		private void ultraTextEditor8_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (e.KeyChar == 13) this.ultraTextEditor14.Focus();
		}

		private void ultraTextEditor14_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (e.KeyChar == 13) this.ultraTextEditor17.Focus();
		}

		private void ultraTextEditor17_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (e.KeyChar == 13) this.ultraTextEditor13.Focus();
		}

		private void ultraTextEditor13_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (e.KeyChar == 13) this.ultraTextEditor16.Focus();
		}

		private void ultraTextEditor16_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (e.KeyChar == 13) this.ultraTextEditor12.Focus();
		}

		private void ultraTextEditor12_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (e.KeyChar == 13) this.ultraTextEditor9.Focus();
		}

		private void ultraTextEditor9_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (e.KeyChar == 13) this.txtCedRep.Focus();
		}

		private void txtCedRep_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (e.KeyChar == 13) this.ultraTextEditor19.Focus();
		}

		private void ultraTextEditor19_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (e.KeyChar == 13) this.ultraTextEditor18.Focus();
		}

		private void txtNotas_ValueChanged(object sender, System.EventArgs e)
		{
		
		}

		private void cmbCiudad_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}

		private void chkPublica_Validated(object sender, System.EventArgs e)
		{
			if ((bool)this.chkPublica.Checked) drCliente["ContribEspecial"] = false;			
		}

		private void chkContribuyenteEsp_Validated(object sender, System.EventArgs e)
		{
			if ((bool)this.chkContribuyenteEsp.Checked) drCliente["InstPublica"] = false;
		}

		private void barraDato1_ContextMenuChanged(object sender, System.EventArgs e)
		{
		
		}
	}
}
