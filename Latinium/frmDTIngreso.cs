using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using DevExpress.XtraEditors;
using System.Data.SqlClient;
using System.IO;
using System.Data;
using System.Text;
using CrystalDecisions;
using CrystalDecisions.CrystalReports.Engine;
using CrystalDecisions.ReportSource;
using CrystalDecisions.Shared;

using System.Diagnostics;
using System.Threading;
using System.Net;
namespace Latinium
{
	/// <summary>
	/// Summary description for frmDTIngreso.
	/// </summary>
	public class frmDTIngreso : DevExpress.XtraEditors.XtraForm
	{
		private System.Windows.Forms.Button btnCancelar;
		private System.Windows.Forms.Button btnAnular;
		private System.Windows.Forms.Button btnSalir;
		private System.Windows.Forms.GroupBox gbProducto;
		private System.Windows.Forms.Label label38;
		private System.Windows.Forms.Label label30;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtSerial;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtIdSerial;
		private System.Windows.Forms.Label label13;
		private System.Windows.Forms.Label label9;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtidArticulo;
		private System.Windows.Forms.Label label7;
		private System.Windows.Forms.Label label8;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtCodArticulo;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtArticulo;
		private System.Windows.Forms.GroupBox gbCliente;
		private System.Windows.Forms.Label label39;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtCelular;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtidCliente;
		private System.Windows.Forms.Label label1;
		private System.Windows.Forms.Label label2;
		private System.Windows.Forms.Label label6;
		private System.Windows.Forms.Label label5;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtTelefono;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtDireccion;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtCodigo;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtNombre;
		private System.Windows.Forms.Label label14;
		private System.Windows.Forms.Button btnNuevo;
		private System.Windows.Forms.Button btnConsultar;
		private System.Windows.Forms.Button btnEditar;
		private System.Windows.Forms.Button btnImprimir;
		private System.Windows.Forms.Button btnGuardar;
		private System.Windows.Forms.Label lblUsuario;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtBodegaPredef;
		private System.Windows.Forms.Label label36;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbBodega;
		private System.Windows.Forms.Label lblEstado;
		private System.Windows.Forms.Label label33;
		private System.Windows.Forms.Label label32;
		private System.Windows.Forms.GroupBox gbProveedor;
		private System.Windows.Forms.Label label37;
		private System.Windows.Forms.Label label12;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo cmbFechaFactura;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtProveedor;
		private System.Windows.Forms.Label label10;
		private System.Windows.Forms.Label label11;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtIdIngreso;
		private System.Windows.Forms.Label label4;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtFecha;
		private C1.Data.C1DataSet cdsSeteoF;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtTiempoDeGarantia;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtDaño;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtGrupoProveedor;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtNumeroDeFactura;
		private System.ComponentModel.IContainer components;

		#region Variables
		bool bNuevo = false;
		bool bEdicion = false;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtEstado;
		bool bAnulado = false;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtIdProveedor;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtOrigen;
		private System.Windows.Forms.Label lblEgreso;
		private System.Windows.Forms.Label lblIngreso;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtIngresos;
		private System.Windows.Forms.ContextMenu mnuImprimir;
		private System.Windows.Forms.MenuItem mnuImprimirIngreso;
		private System.Windows.Forms.MenuItem mnuImprimirEgreso;
		private System.Windows.Forms.ImageList imageList1;
		private System.Windows.Forms.Label label16;
		private System.Windows.Forms.Label label17;
		private System.Windows.Forms.Label label18;
		private System.Windows.Forms.Label label19;
		private System.Windows.Forms.GroupBox groupBox1;
		private System.Windows.Forms.GroupBox groupBox3;
		private System.Windows.Forms.GroupBox groupBox4;
		private System.Windows.Forms.GroupBox groupBox5;
		private System.Windows.Forms.Label label3;
		private System.Windows.Forms.Label label23;
		private System.Windows.Forms.Label label24;
		private System.Windows.Forms.ToolTip toolTip1;
		private System.Windows.Forms.CheckBox chkEmpqs;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtEmpqsObservaciones;
		private System.Windows.Forms.CheckBox chkBordes;
		private System.Windows.Forms.CheckBox chkSupDelantera;
		private System.Windows.Forms.CheckBox chkSupPosterior;
		private System.Windows.Forms.CheckBox chkPines;
		private System.Windows.Forms.CheckBox chkBotones;
		private System.Windows.Forms.CheckBox chkSerial;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtAccesorios;
		private System.Windows.Forms.CheckBox chkAccesorios;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtImagen1;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtImagen2;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtImagen3;
		private System.Windows.Forms.Button btnImg1;
		private System.Windows.Forms.Button btnImg2;
		private System.Windows.Forms.Button btnImg3;
		private System.Windows.Forms.Label label25;
		private System.Windows.Forms.Label label26;
		private System.Windows.Forms.Label label27;
		private System.Windows.Forms.Label label28;
		private System.Windows.Forms.Label label29;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtTiempGarantiaProveedr;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo cmbFechaVenta;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtNumFCVenta;
		private Infragistics.Win.UltraWinEditors.UltraOptionSet rbBordes;
		private Infragistics.Win.UltraWinEditors.UltraOptionSet rbSupDelantera;
		private Infragistics.Win.UltraWinEditors.UltraOptionSet rbSupPosterior;
		private Infragistics.Win.UltraWinEditors.UltraOptionSet rbPines;
		private Infragistics.Win.UltraWinEditors.UltraOptionSet rbBotones;
		private Infragistics.Win.UltraWinEditors.UltraOptionSet rbSerial;
		private System.Windows.Forms.Button button1;
		private System.Windows.Forms.ToolTip toolTip2;
		private System.Windows.Forms.Button button2;
		private System.Windows.Forms.Button button3;
		private System.Windows.Forms.Label label31;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbUsuario;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource1;
		private System.Windows.Forms.Label label34;
		private System.Windows.Forms.Label label35;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtVendedor;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtLocal;
		private System.Windows.Forms.Label label40;
		private System.Windows.Forms.GroupBox groupBox6;
		private System.Windows.Forms.Label label42;
		private System.Windows.Forms.Label label41;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtFonoAdicional;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtNombreEntrega;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtCIEntrega;
		private Acceso miAcceso;
		private System.Windows.Forms.Label label43;
		private System.Windows.Forms.Label label44;
		private System.Windows.Forms.Label label45;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbAsignaTecnico;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtCorreo;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtCelAdicional;
		private System.Windows.Forms.ComboBox cmbGarantia;
		private System.Windows.Forms.ComboBox cmbGarantiaProveedor;
		private Infragistics.Win.UltraWinEditors.UltraOptionSet rbEmpaques;
		private Infragistics.Win.UltraWinEditors.UltraOptionSet rbAccesorios;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtEstadoFisico;
		private System.Windows.Forms.GroupBox groupBox2;
		private System.Windows.Forms.Button btnCliente;
		private System.Windows.Forms.CheckBox chkGarantiaExtendida;
		int validarIngresoArt = 0;
		#endregion Variables

		public frmDTIngreso()
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
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmDTIngreso));
			Infragistics.Win.Appearance appearance1 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton1 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance11 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance12 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance13 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance14 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance15 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance16 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance17 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance18 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance19 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance20 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance21 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance22 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bodega");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo", 0);
			Infragistics.Win.Appearance appearance23 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance24 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance25 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton2 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.Appearance appearance26 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance27 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance28 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton3 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.Appearance appearance29 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance30 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance31 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance32 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance33 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance34 = new Infragistics.Win.Appearance();
			Infragistics.Win.ValueListItem valueListItem1 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem2 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem3 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.Appearance appearance35 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance36 = new Infragistics.Win.Appearance();
			Infragistics.Win.ValueListItem valueListItem4 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem5 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem6 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem7 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem8 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.Appearance appearance37 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance38 = new Infragistics.Win.Appearance();
			Infragistics.Win.ValueListItem valueListItem9 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem10 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem11 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem12 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem13 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.Appearance appearance39 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance40 = new Infragistics.Win.Appearance();
			Infragistics.Win.ValueListItem valueListItem14 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem15 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem16 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem17 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem18 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.Appearance appearance41 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance42 = new Infragistics.Win.Appearance();
			Infragistics.Win.ValueListItem valueListItem19 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem20 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem21 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem22 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem23 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.Appearance appearance43 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance44 = new Infragistics.Win.Appearance();
			Infragistics.Win.ValueListItem valueListItem24 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem25 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem26 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem27 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem28 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.Appearance appearance45 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance46 = new Infragistics.Win.Appearance();
			Infragistics.Win.ValueListItem valueListItem29 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem30 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem31 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem32 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem33 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.Appearance appearance47 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance48 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance49 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance50 = new Infragistics.Win.Appearance();
			Infragistics.Win.ValueListItem valueListItem34 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem35 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem36 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.Appearance appearance51 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance52 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance53 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPersonal");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idNomina");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idPersonal");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idNomina");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn3 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Codigo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn4 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.Appearance appearance54 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance55 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand3 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPersonal");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idNomina");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn11 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			this.btnCancelar = new System.Windows.Forms.Button();
			this.btnAnular = new System.Windows.Forms.Button();
			this.btnSalir = new System.Windows.Forms.Button();
			this.imageList1 = new System.Windows.Forms.ImageList(this.components);
			this.gbProducto = new System.Windows.Forms.GroupBox();
			this.cmbGarantia = new System.Windows.Forms.ComboBox();
			this.txtLocal = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label34 = new System.Windows.Forms.Label();
			this.txtNumFCVenta = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label19 = new System.Windows.Forms.Label();
			this.label18 = new System.Windows.Forms.Label();
			this.cmbFechaVenta = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.label30 = new System.Windows.Forms.Label();
			this.txtSerial = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label13 = new System.Windows.Forms.Label();
			this.txtTiempoDeGarantia = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label9 = new System.Windows.Forms.Label();
			this.label7 = new System.Windows.Forms.Label();
			this.label8 = new System.Windows.Forms.Label();
			this.txtCodArticulo = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtArticulo = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtVendedor = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label35 = new System.Windows.Forms.Label();
			this.txtIngresos = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label38 = new System.Windows.Forms.Label();
			this.txtIdSerial = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtidArticulo = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.gbCliente = new System.Windows.Forms.GroupBox();
			this.btnCliente = new System.Windows.Forms.Button();
			this.label44 = new System.Windows.Forms.Label();
			this.txtCelAdicional = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label43 = new System.Windows.Forms.Label();
			this.txtCorreo = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label40 = new System.Windows.Forms.Label();
			this.txtFonoAdicional = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label39 = new System.Windows.Forms.Label();
			this.txtCelular = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label1 = new System.Windows.Forms.Label();
			this.label2 = new System.Windows.Forms.Label();
			this.label6 = new System.Windows.Forms.Label();
			this.label5 = new System.Windows.Forms.Label();
			this.txtTelefono = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtDireccion = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtCodigo = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtNombre = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtidCliente = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label14 = new System.Windows.Forms.Label();
			this.txtDaño = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.btnNuevo = new System.Windows.Forms.Button();
			this.btnConsultar = new System.Windows.Forms.Button();
			this.btnEditar = new System.Windows.Forms.Button();
			this.btnImprimir = new System.Windows.Forms.Button();
			this.btnGuardar = new System.Windows.Forms.Button();
			this.lblUsuario = new System.Windows.Forms.Label();
			this.txtBodegaPredef = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label36 = new System.Windows.Forms.Label();
			this.cmbBodega = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.lblEstado = new System.Windows.Forms.Label();
			this.label33 = new System.Windows.Forms.Label();
			this.label32 = new System.Windows.Forms.Label();
			this.lblEgreso = new System.Windows.Forms.Label();
			this.gbProveedor = new System.Windows.Forms.GroupBox();
			this.chkGarantiaExtendida = new System.Windows.Forms.CheckBox();
			this.cmbGarantiaProveedor = new System.Windows.Forms.ComboBox();
			this.label16 = new System.Windows.Forms.Label();
			this.txtTiempGarantiaProveedr = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label17 = new System.Windows.Forms.Label();
			this.txtGrupoProveedor = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label37 = new System.Windows.Forms.Label();
			this.label12 = new System.Windows.Forms.Label();
			this.cmbFechaFactura = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.txtProveedor = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtNumeroDeFactura = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label10 = new System.Windows.Forms.Label();
			this.label11 = new System.Windows.Forms.Label();
			this.txtIdIngreso = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.lblIngreso = new System.Windows.Forms.Label();
			this.label4 = new System.Windows.Forms.Label();
			this.dtFecha = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.txtEstado = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtIdProveedor = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtOrigen = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.mnuImprimir = new System.Windows.Forms.ContextMenu();
			this.mnuImprimirIngreso = new System.Windows.Forms.MenuItem();
			this.mnuImprimirEgreso = new System.Windows.Forms.MenuItem();
			this.groupBox1 = new System.Windows.Forms.GroupBox();
			this.txtEmpqsObservaciones = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.chkEmpqs = new System.Windows.Forms.CheckBox();
			this.rbEmpaques = new Infragistics.Win.UltraWinEditors.UltraOptionSet();
			this.groupBox3 = new System.Windows.Forms.GroupBox();
			this.rbSerial = new Infragistics.Win.UltraWinEditors.UltraOptionSet();
			this.rbBotones = new Infragistics.Win.UltraWinEditors.UltraOptionSet();
			this.rbPines = new Infragistics.Win.UltraWinEditors.UltraOptionSet();
			this.rbSupPosterior = new Infragistics.Win.UltraWinEditors.UltraOptionSet();
			this.rbSupDelantera = new Infragistics.Win.UltraWinEditors.UltraOptionSet();
			this.rbBordes = new Infragistics.Win.UltraWinEditors.UltraOptionSet();
			this.chkSerial = new System.Windows.Forms.CheckBox();
			this.chkBotones = new System.Windows.Forms.CheckBox();
			this.chkPines = new System.Windows.Forms.CheckBox();
			this.chkSupPosterior = new System.Windows.Forms.CheckBox();
			this.chkSupDelantera = new System.Windows.Forms.CheckBox();
			this.txtEstadoFisico = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.chkBordes = new System.Windows.Forms.CheckBox();
			this.groupBox4 = new System.Windows.Forms.GroupBox();
			this.txtAccesorios = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.chkAccesorios = new System.Windows.Forms.CheckBox();
			this.rbAccesorios = new Infragistics.Win.UltraWinEditors.UltraOptionSet();
			this.groupBox5 = new System.Windows.Forms.GroupBox();
			this.button3 = new System.Windows.Forms.Button();
			this.button2 = new System.Windows.Forms.Button();
			this.button1 = new System.Windows.Forms.Button();
			this.btnImg3 = new System.Windows.Forms.Button();
			this.btnImg2 = new System.Windows.Forms.Button();
			this.btnImg1 = new System.Windows.Forms.Button();
			this.label24 = new System.Windows.Forms.Label();
			this.txtImagen3 = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label23 = new System.Windows.Forms.Label();
			this.txtImagen2 = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label3 = new System.Windows.Forms.Label();
			this.txtImagen1 = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.toolTip1 = new System.Windows.Forms.ToolTip(this.components);
			this.label25 = new System.Windows.Forms.Label();
			this.label26 = new System.Windows.Forms.Label();
			this.label27 = new System.Windows.Forms.Label();
			this.label28 = new System.Windows.Forms.Label();
			this.label29 = new System.Windows.Forms.Label();
			this.toolTip2 = new System.Windows.Forms.ToolTip(this.components);
			this.label31 = new System.Windows.Forms.Label();
			this.cmbUsuario = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.ultraDataSource1 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.groupBox6 = new System.Windows.Forms.GroupBox();
			this.label42 = new System.Windows.Forms.Label();
			this.txtNombreEntrega = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label41 = new System.Windows.Forms.Label();
			this.txtCIEntrega = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label45 = new System.Windows.Forms.Label();
			this.cmbAsignaTecnico = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.groupBox2 = new System.Windows.Forms.GroupBox();
			this.gbProducto.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.txtLocal)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNumFCVenta)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbFechaVenta)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtSerial)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtTiempoDeGarantia)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCodArticulo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtArticulo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtVendedor)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIngresos)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdSerial)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtidArticulo)).BeginInit();
			this.gbCliente.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.txtCelAdicional)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCorreo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtFonoAdicional)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCelular)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtTelefono)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDireccion)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCodigo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNombre)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtidCliente)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDaño)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtBodegaPredef)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbBodega)).BeginInit();
			this.gbProveedor.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.txtTiempGarantiaProveedr)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtGrupoProveedor)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbFechaFactura)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtProveedor)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNumeroDeFactura)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdIngreso)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFecha)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtEstado)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdProveedor)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtOrigen)).BeginInit();
			this.groupBox1.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.txtEmpqsObservaciones)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.rbEmpaques)).BeginInit();
			this.groupBox3.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.rbSerial)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.rbBotones)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.rbPines)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.rbSupPosterior)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.rbSupDelantera)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.rbBordes)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtEstadoFisico)).BeginInit();
			this.groupBox4.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.txtAccesorios)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.rbAccesorios)).BeginInit();
			this.groupBox5.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.txtImagen3)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtImagen2)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtImagen1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbUsuario)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).BeginInit();
			this.groupBox6.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.txtNombreEntrega)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCIEntrega)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbAsignaTecnico)).BeginInit();
			this.SuspendLayout();
			// 
			// btnCancelar
			// 
			this.btnCancelar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnCancelar.CausesValidation = false;
			this.btnCancelar.Enabled = false;
			this.btnCancelar.Image = ((System.Drawing.Image)(resources.GetObject("btnCancelar.Image")));
			this.btnCancelar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnCancelar.Location = new System.Drawing.Point(469, 492);
			this.btnCancelar.Name = "btnCancelar";
			this.btnCancelar.Size = new System.Drawing.Size(72, 23);
			this.btnCancelar.TabIndex = 361;
			this.btnCancelar.Text = "&Cancelar";
			this.btnCancelar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnCancelar.Click += new System.EventHandler(this.btnCancelar_Click);
			// 
			// btnAnular
			// 
			this.btnAnular.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnAnular.CausesValidation = false;
			this.btnAnular.Enabled = false;
			this.btnAnular.Image = ((System.Drawing.Image)(resources.GetObject("btnAnular.Image")));
			this.btnAnular.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnAnular.Location = new System.Drawing.Point(392, 492);
			this.btnAnular.Name = "btnAnular";
			this.btnAnular.Size = new System.Drawing.Size(76, 23);
			this.btnAnular.TabIndex = 360;
			this.btnAnular.Text = "&Anular";
			this.btnAnular.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnAnular.Click += new System.EventHandler(this.btnAnular_Click);
			// 
			// btnSalir
			// 
			this.btnSalir.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnSalir.CausesValidation = false;
			this.btnSalir.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnSalir.ImageIndex = 62;
			this.btnSalir.ImageList = this.imageList1;
			this.btnSalir.Location = new System.Drawing.Point(542, 492);
			this.btnSalir.Name = "btnSalir";
			this.btnSalir.TabIndex = 362;
			this.btnSalir.Text = "&Salir";
			this.btnSalir.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnSalir.Click += new System.EventHandler(this.btnSalir_Click);
			// 
			// imageList1
			// 
			this.imageList1.ImageSize = new System.Drawing.Size(16, 16);
			this.imageList1.ImageStream = ((System.Windows.Forms.ImageListStreamer)(resources.GetObject("imageList1.ImageStream")));
			this.imageList1.TransparentColor = System.Drawing.Color.Transparent;
			// 
			// gbProducto
			// 
			this.gbProducto.Controls.Add(this.cmbGarantia);
			this.gbProducto.Controls.Add(this.txtLocal);
			this.gbProducto.Controls.Add(this.label34);
			this.gbProducto.Controls.Add(this.txtNumFCVenta);
			this.gbProducto.Controls.Add(this.label19);
			this.gbProducto.Controls.Add(this.label18);
			this.gbProducto.Controls.Add(this.cmbFechaVenta);
			this.gbProducto.Controls.Add(this.label30);
			this.gbProducto.Controls.Add(this.txtSerial);
			this.gbProducto.Controls.Add(this.label13);
			this.gbProducto.Controls.Add(this.txtTiempoDeGarantia);
			this.gbProducto.Controls.Add(this.label9);
			this.gbProducto.Controls.Add(this.label7);
			this.gbProducto.Controls.Add(this.label8);
			this.gbProducto.Controls.Add(this.txtCodArticulo);
			this.gbProducto.Controls.Add(this.txtArticulo);
			this.gbProducto.Controls.Add(this.txtVendedor);
			this.gbProducto.Controls.Add(this.label35);
			this.gbProducto.Controls.Add(this.txtIngresos);
			this.gbProducto.Controls.Add(this.label38);
			this.gbProducto.Location = new System.Drawing.Point(388, 40);
			this.gbProducto.Name = "gbProducto";
			this.gbProducto.Size = new System.Drawing.Size(388, 160);
			this.gbProducto.TabIndex = 342;
			this.gbProducto.TabStop = false;
			this.gbProducto.Text = "Articulo";
			// 
			// cmbGarantia
			// 
			this.cmbGarantia.Enabled = false;
			this.cmbGarantia.Location = new System.Drawing.Point(72, 84);
			this.cmbGarantia.Name = "cmbGarantia";
			this.cmbGarantia.Size = new System.Drawing.Size(120, 21);
			this.cmbGarantia.TabIndex = 296;
			// 
			// txtLocal
			// 
			appearance1.ForeColorDisabled = System.Drawing.Color.Black;
			appearance1.TextHAlign = Infragistics.Win.HAlign.Center;
			this.txtLocal.Appearance = appearance1;
			this.txtLocal.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtLocal.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtLocal.Enabled = false;
			this.txtLocal.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtLocal.Location = new System.Drawing.Point(328, 18);
			this.txtLocal.Name = "txtLocal";
			this.txtLocal.Size = new System.Drawing.Size(56, 20);
			this.txtLocal.TabIndex = 295;
			// 
			// label34
			// 
			this.label34.AutoSize = true;
			this.label34.BackColor = System.Drawing.Color.Transparent;
			this.label34.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label34.Location = new System.Drawing.Point(272, 21);
			this.label34.Name = "label34";
			this.label34.Size = new System.Drawing.Size(29, 15);
			this.label34.TabIndex = 294;
			this.label34.Text = "Local";
			this.label34.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtNumFCVenta
			// 
			appearance2.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtNumFCVenta.Appearance = appearance2;
			this.txtNumFCVenta.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtNumFCVenta.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNumFCVenta.Enabled = false;
			this.txtNumFCVenta.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtNumFCVenta.Location = new System.Drawing.Point(72, 106);
			this.txtNumFCVenta.Name = "txtNumFCVenta";
			this.txtNumFCVenta.Size = new System.Drawing.Size(120, 20);
			this.txtNumFCVenta.TabIndex = 293;
			// 
			// label19
			// 
			this.label19.AutoSize = true;
			this.label19.BackColor = System.Drawing.Color.Transparent;
			this.label19.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label19.Location = new System.Drawing.Point(8, 109);
			this.label19.Name = "label19";
			this.label19.Size = new System.Drawing.Size(52, 15);
			this.label19.TabIndex = 292;
			this.label19.Text = "N. Factura";
			this.label19.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label18
			// 
			this.label18.AutoSize = true;
			this.label18.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label18.Location = new System.Drawing.Point(200, 109);
			this.label18.Name = "label18";
			this.label18.Size = new System.Drawing.Size(70, 15);
			this.label18.TabIndex = 291;
			this.label18.Text = "Fecha Factura";
			this.label18.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// cmbFechaVenta
			// 
			appearance3.ForeColor = System.Drawing.Color.Black;
			appearance3.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbFechaVenta.Appearance = appearance3;
			this.cmbFechaVenta.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton1.Caption = "Today";
			this.cmbFechaVenta.DateButtons.Add(dateButton1);
			this.cmbFechaVenta.Enabled = false;
			this.cmbFechaVenta.Format = "dd/MM/yyyy";
			this.cmbFechaVenta.Location = new System.Drawing.Point(280, 106);
			this.cmbFechaVenta.Name = "cmbFechaVenta";
			this.cmbFechaVenta.NonAutoSizeHeight = 23;
			this.cmbFechaVenta.Size = new System.Drawing.Size(104, 21);
			this.cmbFechaVenta.SpinButtonsVisible = true;
			this.cmbFechaVenta.TabIndex = 290;
			this.cmbFechaVenta.Value = ((object)(resources.GetObject("cmbFechaVenta.Value")));
			// 
			// label30
			// 
			this.label30.AutoSize = true;
			this.label30.BackColor = System.Drawing.Color.Transparent;
			this.label30.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label30.Location = new System.Drawing.Point(8, 21);
			this.label30.Name = "label30";
			this.label30.Size = new System.Drawing.Size(30, 15);
			this.label30.TabIndex = 278;
			this.label30.Text = "Serial";
			this.label30.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtSerial
			// 
			appearance4.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtSerial.Appearance = appearance4;
			this.txtSerial.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtSerial.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtSerial.Enabled = false;
			this.txtSerial.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtSerial.Location = new System.Drawing.Point(72, 18);
			this.txtSerial.Name = "txtSerial";
			this.txtSerial.Size = new System.Drawing.Size(192, 20);
			this.txtSerial.TabIndex = 277;
			this.txtSerial.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtSerial_KeyDown);
			this.txtSerial.Validated += new System.EventHandler(this.txtSerial_Validated);
			// 
			// label13
			// 
			this.label13.AutoSize = true;
			this.label13.BackColor = System.Drawing.Color.Transparent;
			this.label13.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label13.Location = new System.Drawing.Point(200, 87);
			this.label13.Name = "label13";
			this.label13.Size = new System.Drawing.Size(95, 15);
			this.label13.TabIndex = 275;
			this.label13.Text = "Tiempo de Garantía";
			this.label13.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtTiempoDeGarantia
			// 
			appearance5.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtTiempoDeGarantia.Appearance = appearance5;
			this.txtTiempoDeGarantia.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtTiempoDeGarantia.Enabled = false;
			this.txtTiempoDeGarantia.FormatString = "";
			this.txtTiempoDeGarantia.Location = new System.Drawing.Point(328, 83);
			this.txtTiempoDeGarantia.Name = "txtTiempoDeGarantia";
			this.txtTiempoDeGarantia.PromptChar = ' ';
			this.txtTiempoDeGarantia.Size = new System.Drawing.Size(56, 22);
			this.txtTiempoDeGarantia.TabIndex = 274;
			// 
			// label9
			// 
			this.label9.AutoSize = true;
			this.label9.BackColor = System.Drawing.Color.Transparent;
			this.label9.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label9.Location = new System.Drawing.Point(8, 87);
			this.label9.Name = "label9";
			this.label9.Size = new System.Drawing.Size(43, 15);
			this.label9.TabIndex = 272;
			this.label9.Text = "Garantía";
			this.label9.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label7
			// 
			this.label7.AutoSize = true;
			this.label7.BackColor = System.Drawing.Color.Transparent;
			this.label7.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label7.Location = new System.Drawing.Point(8, 43);
			this.label7.Name = "label7";
			this.label7.Size = new System.Drawing.Size(37, 15);
			this.label7.TabIndex = 270;
			this.label7.Text = "Código";
			this.label7.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label8
			// 
			this.label8.AutoSize = true;
			this.label8.BackColor = System.Drawing.Color.Transparent;
			this.label8.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label8.Location = new System.Drawing.Point(8, 65);
			this.label8.Name = "label8";
			this.label8.Size = new System.Drawing.Size(38, 15);
			this.label8.TabIndex = 268;
			this.label8.Text = "Artículo";
			this.label8.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtCodArticulo
			// 
			appearance6.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtCodArticulo.Appearance = appearance6;
			this.txtCodArticulo.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtCodArticulo.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtCodArticulo.Enabled = false;
			this.txtCodArticulo.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtCodArticulo.Location = new System.Drawing.Point(72, 40);
			this.txtCodArticulo.Name = "txtCodArticulo";
			this.txtCodArticulo.Size = new System.Drawing.Size(192, 20);
			this.txtCodArticulo.TabIndex = 269;
			this.txtCodArticulo.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtCodArticulo_KeyDown);
			this.txtCodArticulo.Validated += new System.EventHandler(this.txtCodArticulo_Validated);
			this.txtCodArticulo.ValueChanged += new System.EventHandler(this.txtCodArticulo_ValueChanged);
			// 
			// txtArticulo
			// 
			appearance7.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtArticulo.Appearance = appearance7;
			this.txtArticulo.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtArticulo.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtArticulo.Enabled = false;
			this.txtArticulo.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtArticulo.Location = new System.Drawing.Point(72, 62);
			this.txtArticulo.Name = "txtArticulo";
			this.txtArticulo.Size = new System.Drawing.Size(312, 20);
			this.txtArticulo.TabIndex = 267;
			this.txtArticulo.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtArticulo_KeyDown);
			// 
			// txtVendedor
			// 
			appearance8.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtVendedor.Appearance = appearance8;
			this.txtVendedor.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtVendedor.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtVendedor.Enabled = false;
			this.txtVendedor.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtVendedor.Location = new System.Drawing.Point(72, 128);
			this.txtVendedor.MaxLength = 13;
			this.txtVendedor.Name = "txtVendedor";
			this.txtVendedor.Size = new System.Drawing.Size(312, 20);
			this.txtVendedor.TabIndex = 276;
			// 
			// label35
			// 
			this.label35.AutoSize = true;
			this.label35.BackColor = System.Drawing.Color.Transparent;
			this.label35.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label35.Location = new System.Drawing.Point(8, 131);
			this.label35.Name = "label35";
			this.label35.Size = new System.Drawing.Size(48, 15);
			this.label35.TabIndex = 277;
			this.label35.Text = "Vendedor";
			this.label35.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtIngresos
			// 
			appearance9.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtIngresos.Appearance = appearance9;
			this.txtIngresos.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtIngresos.Enabled = false;
			this.txtIngresos.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtIngresos.FormatString = "";
			this.txtIngresos.Location = new System.Drawing.Point(328, 40);
			this.txtIngresos.Name = "txtIngresos";
			this.txtIngresos.PromptChar = ' ';
			this.txtIngresos.Size = new System.Drawing.Size(56, 20);
			this.txtIngresos.TabIndex = 289;
			// 
			// label38
			// 
			this.label38.AutoSize = true;
			this.label38.BackColor = System.Drawing.Color.Transparent;
			this.label38.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label38.Location = new System.Drawing.Point(272, 43);
			this.label38.Name = "label38";
			this.label38.Size = new System.Drawing.Size(43, 15);
			this.label38.TabIndex = 288;
			this.label38.Text = "Ingresos";
			this.label38.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtIdSerial
			// 
			this.txtIdSerial.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			appearance10.ForeColorDisabled = System.Drawing.Color.Black;
			appearance10.TextHAlign = Infragistics.Win.HAlign.Right;
			appearance10.TextVAlign = Infragistics.Win.VAlign.Middle;
			this.txtIdSerial.Appearance = appearance10;
			this.txtIdSerial.Enabled = false;
			this.txtIdSerial.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtIdSerial.FormatString = "";
			this.txtIdSerial.Location = new System.Drawing.Point(964, 498);
			this.txtIdSerial.Name = "txtIdSerial";
			this.txtIdSerial.PromptChar = ' ';
			this.txtIdSerial.Size = new System.Drawing.Size(25, 21);
			this.txtIdSerial.TabIndex = 276;
			this.txtIdSerial.Visible = false;
			// 
			// txtidArticulo
			// 
			this.txtidArticulo.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			appearance11.ForeColorDisabled = System.Drawing.Color.Black;
			appearance11.TextHAlign = Infragistics.Win.HAlign.Right;
			appearance11.TextVAlign = Infragistics.Win.VAlign.Middle;
			this.txtidArticulo.Appearance = appearance11;
			this.txtidArticulo.Enabled = false;
			this.txtidArticulo.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtidArticulo.FormatString = "";
			this.txtidArticulo.Location = new System.Drawing.Point(940, 498);
			this.txtidArticulo.Name = "txtidArticulo";
			this.txtidArticulo.PromptChar = ' ';
			this.txtidArticulo.Size = new System.Drawing.Size(25, 21);
			this.txtidArticulo.TabIndex = 271;
			this.txtidArticulo.Visible = false;
			// 
			// gbCliente
			// 
			this.gbCliente.Controls.Add(this.btnCliente);
			this.gbCliente.Controls.Add(this.label44);
			this.gbCliente.Controls.Add(this.txtCelAdicional);
			this.gbCliente.Controls.Add(this.label43);
			this.gbCliente.Controls.Add(this.txtCorreo);
			this.gbCliente.Controls.Add(this.label40);
			this.gbCliente.Controls.Add(this.txtFonoAdicional);
			this.gbCliente.Controls.Add(this.label39);
			this.gbCliente.Controls.Add(this.txtCelular);
			this.gbCliente.Controls.Add(this.label1);
			this.gbCliente.Controls.Add(this.label2);
			this.gbCliente.Controls.Add(this.label6);
			this.gbCliente.Controls.Add(this.label5);
			this.gbCliente.Controls.Add(this.txtTelefono);
			this.gbCliente.Controls.Add(this.txtDireccion);
			this.gbCliente.Controls.Add(this.txtCodigo);
			this.gbCliente.Controls.Add(this.txtNombre);
			this.gbCliente.Location = new System.Drawing.Point(5, 40);
			this.gbCliente.Name = "gbCliente";
			this.gbCliente.Size = new System.Drawing.Size(383, 160);
			this.gbCliente.TabIndex = 341;
			this.gbCliente.TabStop = false;
			this.gbCliente.Text = "Cliente";
			// 
			// btnCliente
			// 
			this.btnCliente.Enabled = false;
			this.btnCliente.Image = ((System.Drawing.Image)(resources.GetObject("btnCliente.Image")));
			this.btnCliente.Location = new System.Drawing.Point(168, 16);
			this.btnCliente.Name = "btnCliente";
			this.btnCliente.Size = new System.Drawing.Size(24, 22);
			this.btnCliente.TabIndex = 348;
			this.toolTip1.SetToolTip(this.btnCliente, "Cliente");
			this.btnCliente.Click += new System.EventHandler(this.btnCliente_Click);
			// 
			// label44
			// 
			this.label44.AutoSize = true;
			this.label44.BackColor = System.Drawing.Color.Transparent;
			this.label44.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label44.Location = new System.Drawing.Point(160, 131);
			this.label44.Name = "label44";
			this.label44.Size = new System.Drawing.Size(37, 15);
			this.label44.TabIndex = 347;
			this.label44.Text = "Celular";
			this.label44.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtCelAdicional
			// 
			appearance12.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtCelAdicional.Appearance = appearance12;
			this.txtCelAdicional.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtCelAdicional.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtCelAdicional.Enabled = false;
			this.txtCelAdicional.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtCelAdicional.Location = new System.Drawing.Point(208, 128);
			this.txtCelAdicional.MaxLength = 10;
			this.txtCelAdicional.Name = "txtCelAdicional";
			this.txtCelAdicional.Size = new System.Drawing.Size(80, 20);
			this.txtCelAdicional.TabIndex = 346;
			this.toolTip1.SetToolTip(this.txtCelAdicional, "Celular adicional");
			this.txtCelAdicional.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtCelAdicional_KeyDown);
			this.txtCelAdicional.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtCelAdicional_KeyPress);
			// 
			// label43
			// 
			this.label43.AutoSize = true;
			this.label43.BackColor = System.Drawing.Color.Transparent;
			this.label43.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label43.Location = new System.Drawing.Point(8, 87);
			this.label43.Name = "label43";
			this.label43.Size = new System.Drawing.Size(35, 15);
			this.label43.TabIndex = 345;
			this.label43.Text = "Correo";
			this.label43.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtCorreo
			// 
			appearance13.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtCorreo.Appearance = appearance13;
			this.txtCorreo.CharacterCasing = System.Windows.Forms.CharacterCasing.Lower;
			this.txtCorreo.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtCorreo.Enabled = false;
			this.txtCorreo.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtCorreo.Location = new System.Drawing.Point(64, 84);
			this.txtCorreo.MaxLength = 100;
			this.txtCorreo.Name = "txtCorreo";
			this.txtCorreo.Size = new System.Drawing.Size(224, 20);
			this.txtCorreo.TabIndex = 344;
			// 
			// label40
			// 
			this.label40.AutoSize = true;
			this.label40.BackColor = System.Drawing.Color.Transparent;
			this.label40.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label40.Location = new System.Drawing.Point(8, 131);
			this.label40.Name = "label40";
			this.label40.Size = new System.Drawing.Size(46, 15);
			this.label40.TabIndex = 343;
			this.label40.Text = "Adicional";
			this.label40.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtFonoAdicional
			// 
			appearance14.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtFonoAdicional.Appearance = appearance14;
			this.txtFonoAdicional.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtFonoAdicional.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtFonoAdicional.Enabled = false;
			this.txtFonoAdicional.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtFonoAdicional.Location = new System.Drawing.Point(64, 128);
			this.txtFonoAdicional.MaxLength = 9;
			this.txtFonoAdicional.Name = "txtFonoAdicional";
			this.txtFonoAdicional.Size = new System.Drawing.Size(80, 20);
			this.txtFonoAdicional.TabIndex = 342;
			this.toolTip1.SetToolTip(this.txtFonoAdicional, "Teléfono adicional");
			this.txtFonoAdicional.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtFonoAdicional_KeyDown);
			this.txtFonoAdicional.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtFonoAdicional_KeyPress);
			this.txtFonoAdicional.Click += new System.EventHandler(this.txtFonoAdicional_Click);
			this.txtFonoAdicional.Enter += new System.EventHandler(this.txtFonoAdicional_Enter);
			// 
			// label39
			// 
			this.label39.AutoSize = true;
			this.label39.BackColor = System.Drawing.Color.Transparent;
			this.label39.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label39.Location = new System.Drawing.Point(160, 109);
			this.label39.Name = "label39";
			this.label39.Size = new System.Drawing.Size(37, 15);
			this.label39.TabIndex = 275;
			this.label39.Text = "Celular";
			this.label39.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtCelular
			// 
			appearance15.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtCelular.Appearance = appearance15;
			this.txtCelular.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtCelular.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtCelular.Enabled = false;
			this.txtCelular.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtCelular.Location = new System.Drawing.Point(208, 106);
			this.txtCelular.MaxLength = 10;
			this.txtCelular.Name = "txtCelular";
			this.txtCelular.Size = new System.Drawing.Size(80, 20);
			this.txtCelular.TabIndex = 274;
			this.txtCelular.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtCelular_KeyDown);
			this.txtCelular.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtCelular_KeyPress);
			this.txtCelular.ValueChanged += new System.EventHandler(this.txtCelular_ValueChanged);
			this.txtCelular.Click += new System.EventHandler(this.txtCelular_Click);
			this.txtCelular.Enter += new System.EventHandler(this.txtCelular_Enter);
			// 
			// label1
			// 
			this.label1.AutoSize = true;
			this.label1.BackColor = System.Drawing.Color.Transparent;
			this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label1.Location = new System.Drawing.Point(8, 109);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(44, 15);
			this.label1.TabIndex = 272;
			this.label1.Text = "Teléfono";
			this.label1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label2
			// 
			this.label2.AutoSize = true;
			this.label2.BackColor = System.Drawing.Color.Transparent;
			this.label2.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label2.Location = new System.Drawing.Point(8, 65);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(47, 15);
			this.label2.TabIndex = 270;
			this.label2.Text = "Dirección";
			this.label2.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label6
			// 
			this.label6.AutoSize = true;
			this.label6.BackColor = System.Drawing.Color.Transparent;
			this.label6.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label6.Location = new System.Drawing.Point(8, 21);
			this.label6.Name = "label6";
			this.label6.Size = new System.Drawing.Size(37, 15);
			this.label6.TabIndex = 268;
			this.label6.Text = "Cédula";
			this.label6.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label5
			// 
			this.label5.AutoSize = true;
			this.label5.BackColor = System.Drawing.Color.Transparent;
			this.label5.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label5.Location = new System.Drawing.Point(8, 43);
			this.label5.Name = "label5";
			this.label5.Size = new System.Drawing.Size(40, 15);
			this.label5.TabIndex = 266;
			this.label5.Text = "Nombre";
			this.label5.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtTelefono
			// 
			appearance16.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtTelefono.Appearance = appearance16;
			this.txtTelefono.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtTelefono.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtTelefono.Enabled = false;
			this.txtTelefono.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtTelefono.Location = new System.Drawing.Point(64, 106);
			this.txtTelefono.MaxLength = 9;
			this.txtTelefono.Name = "txtTelefono";
			this.txtTelefono.Size = new System.Drawing.Size(80, 20);
			this.txtTelefono.TabIndex = 271;
			this.txtTelefono.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtTelefono_KeyDown);
			this.txtTelefono.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtTelefono_KeyPress);
			this.txtTelefono.Click += new System.EventHandler(this.txtTelefono_Click);
			this.txtTelefono.Enter += new System.EventHandler(this.txtTelefono_Enter);
			// 
			// txtDireccion
			// 
			appearance17.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtDireccion.Appearance = appearance17;
			this.txtDireccion.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtDireccion.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtDireccion.Enabled = false;
			this.txtDireccion.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtDireccion.Location = new System.Drawing.Point(64, 62);
			this.txtDireccion.Name = "txtDireccion";
			this.txtDireccion.Size = new System.Drawing.Size(312, 20);
			this.txtDireccion.TabIndex = 269;
			this.txtDireccion.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtDireccion_KeyDown);
			this.txtDireccion.Click += new System.EventHandler(this.txtDireccion_Click);
			this.txtDireccion.Enter += new System.EventHandler(this.txtDireccion_Enter);
			// 
			// txtCodigo
			// 
			appearance18.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtCodigo.Appearance = appearance18;
			this.txtCodigo.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtCodigo.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtCodigo.Enabled = false;
			this.txtCodigo.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtCodigo.Location = new System.Drawing.Point(64, 18);
			this.txtCodigo.MaxLength = 13;
			this.txtCodigo.Name = "txtCodigo";
			this.txtCodigo.Size = new System.Drawing.Size(96, 20);
			this.txtCodigo.TabIndex = 267;
			this.txtCodigo.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtCodigo_KeyDown);
			this.txtCodigo.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtCodigo_KeyPress);
			this.txtCodigo.Validating += new System.ComponentModel.CancelEventHandler(this.txtCodigo_Validating);
			this.txtCodigo.Validated += new System.EventHandler(this.txtCodigo_Validated);
			this.txtCodigo.Click += new System.EventHandler(this.txtCodigo_Click);
			this.txtCodigo.Enter += new System.EventHandler(this.txtCodigo_Enter);
			// 
			// txtNombre
			// 
			appearance19.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtNombre.Appearance = appearance19;
			this.txtNombre.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtNombre.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNombre.Enabled = false;
			this.txtNombre.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtNombre.Location = new System.Drawing.Point(64, 40);
			this.txtNombre.Name = "txtNombre";
			this.txtNombre.Size = new System.Drawing.Size(312, 20);
			this.txtNombre.TabIndex = 265;
			this.txtNombre.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtNombre_KeyDown);
			this.txtNombre.Click += new System.EventHandler(this.txtNombre_Click);
			this.txtNombre.Enter += new System.EventHandler(this.txtNombre_Enter);
			// 
			// txtidCliente
			// 
			this.txtidCliente.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			appearance20.ForeColorDisabled = System.Drawing.Color.Black;
			appearance20.TextHAlign = Infragistics.Win.HAlign.Right;
			appearance20.TextVAlign = Infragistics.Win.VAlign.Middle;
			this.txtidCliente.Appearance = appearance20;
			this.txtidCliente.Enabled = false;
			this.txtidCliente.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtidCliente.FormatString = "";
			this.txtidCliente.Location = new System.Drawing.Point(916, 498);
			this.txtidCliente.Name = "txtidCliente";
			this.txtidCliente.PromptChar = ' ';
			this.txtidCliente.Size = new System.Drawing.Size(25, 21);
			this.txtidCliente.TabIndex = 273;
			this.txtidCliente.Visible = false;
			// 
			// label14
			// 
			this.label14.BackColor = System.Drawing.Color.Transparent;
			this.label14.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label14.Location = new System.Drawing.Point(8, 368);
			this.label14.Name = "label14";
			this.label14.Size = new System.Drawing.Size(40, 16);
			this.label14.TabIndex = 338;
			this.label14.Text = "Daño";
			this.label14.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtDaño
			// 
			appearance21.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtDaño.Appearance = appearance21;
			this.txtDaño.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtDaño.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtDaño.Enabled = false;
			this.txtDaño.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtDaño.Location = new System.Drawing.Point(8, 392);
			this.txtDaño.Multiline = true;
			this.txtDaño.Name = "txtDaño";
			this.txtDaño.Scrollbars = System.Windows.Forms.ScrollBars.Vertical;
			this.txtDaño.Size = new System.Drawing.Size(404, 80);
			this.txtDaño.TabIndex = 337;
			this.txtDaño.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtDaño_KeyDown);
			// 
			// btnNuevo
			// 
			this.btnNuevo.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnNuevo.CausesValidation = false;
			this.btnNuevo.Enabled = false;
			this.btnNuevo.Image = ((System.Drawing.Image)(resources.GetObject("btnNuevo.Image")));
			this.btnNuevo.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnNuevo.Location = new System.Drawing.Point(8, 492);
			this.btnNuevo.Name = "btnNuevo";
			this.btnNuevo.Size = new System.Drawing.Size(72, 23);
			this.btnNuevo.TabIndex = 355;
			this.btnNuevo.Text = "&Nuevo";
			this.btnNuevo.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnNuevo.Click += new System.EventHandler(this.btnNuevo_Click);
			// 
			// btnConsultar
			// 
			this.btnConsultar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnConsultar.CausesValidation = false;
			this.btnConsultar.Enabled = false;
			this.btnConsultar.Image = ((System.Drawing.Image)(resources.GetObject("btnConsultar.Image")));
			this.btnConsultar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnConsultar.Location = new System.Drawing.Point(81, 492);
			this.btnConsultar.Name = "btnConsultar";
			this.btnConsultar.Size = new System.Drawing.Size(78, 23);
			this.btnConsultar.TabIndex = 356;
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
			this.btnEditar.Location = new System.Drawing.Point(160, 492);
			this.btnEditar.Name = "btnEditar";
			this.btnEditar.Size = new System.Drawing.Size(76, 23);
			this.btnEditar.TabIndex = 359;
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
			this.btnImprimir.Location = new System.Drawing.Point(237, 492);
			this.btnImprimir.Name = "btnImprimir";
			this.btnImprimir.Size = new System.Drawing.Size(77, 23);
			this.btnImprimir.TabIndex = 358;
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
			this.btnGuardar.Location = new System.Drawing.Point(315, 492);
			this.btnGuardar.Name = "btnGuardar";
			this.btnGuardar.Size = new System.Drawing.Size(76, 23);
			this.btnGuardar.TabIndex = 357;
			this.btnGuardar.Text = "&Guardar";
			this.btnGuardar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnGuardar.Click += new System.EventHandler(this.btnGuardar_Click);
			// 
			// lblUsuario
			// 
			this.lblUsuario.AutoSize = true;
			this.lblUsuario.Font = new System.Drawing.Font("Tahoma", 12F, System.Drawing.FontStyle.Bold);
			this.lblUsuario.ForeColor = System.Drawing.Color.Firebrick;
			this.lblUsuario.Location = new System.Drawing.Point(864, 506);
			this.lblUsuario.Name = "lblUsuario";
			this.lblUsuario.Size = new System.Drawing.Size(0, 23);
			this.lblUsuario.TabIndex = 354;
			this.lblUsuario.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtBodegaPredef
			// 
			this.txtBodegaPredef.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			appearance22.ForeColorDisabled = System.Drawing.Color.Black;
			appearance22.TextHAlign = Infragistics.Win.HAlign.Right;
			appearance22.TextVAlign = Infragistics.Win.VAlign.Middle;
			this.txtBodegaPredef.Appearance = appearance22;
			this.txtBodegaPredef.Enabled = false;
			this.txtBodegaPredef.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtBodegaPredef.FormatString = "";
			this.txtBodegaPredef.Location = new System.Drawing.Point(988, 498);
			this.txtBodegaPredef.Name = "txtBodegaPredef";
			this.txtBodegaPredef.PromptChar = ' ';
			this.txtBodegaPredef.Size = new System.Drawing.Size(25, 21);
			this.txtBodegaPredef.TabIndex = 353;
			this.txtBodegaPredef.Visible = false;
			// 
			// label36
			// 
			this.label36.AutoSize = true;
			this.label36.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label36.Location = new System.Drawing.Point(184, 11);
			this.label36.Name = "label36";
			this.label36.Size = new System.Drawing.Size(29, 15);
			this.label36.TabIndex = 349;
			this.label36.Text = "Local";
			this.label36.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// cmbBodega
			// 
			this.cmbBodega.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbBodega.Cursor = System.Windows.Forms.Cursors.Default;
			ultraGridColumn1.Header.VisiblePosition = 2;
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridColumn2.Width = 180;
			ultraGridColumn3.Header.VisiblePosition = 0;
			ultraGridBand1.Columns.AddRange(new object[] {
																										 ultraGridColumn1,
																										 ultraGridColumn2,
																										 ultraGridColumn3});
			this.cmbBodega.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.cmbBodega.DisplayMember = "Nombre";
			this.cmbBodega.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbBodega.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbBodega.Enabled = false;
			this.cmbBodega.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbBodega.Location = new System.Drawing.Point(224, 8);
			this.cmbBodega.MaxDropDownItems = 30;
			this.cmbBodega.Name = "cmbBodega";
			this.cmbBodega.Size = new System.Drawing.Size(200, 20);
			this.cmbBodega.TabIndex = 350;
			this.cmbBodega.ValueMember = "Bodega";
			this.cmbBodega.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cmbBodega_KeyDown);
			this.cmbBodega.ValueChanged += new System.EventHandler(this.cmbBodega_ValueChanged);
			this.cmbBodega.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.cmbBodega_InitializeLayout);
			// 
			// lblEstado
			// 
			this.lblEstado.AutoSize = true;
			this.lblEstado.Font = new System.Drawing.Font("Tahoma", 12F, System.Drawing.FontStyle.Bold);
			this.lblEstado.ForeColor = System.Drawing.Color.Firebrick;
			this.lblEstado.Location = new System.Drawing.Point(632, 492);
			this.lblEstado.Name = "lblEstado";
			this.lblEstado.Size = new System.Drawing.Size(0, 23);
			this.lblEstado.TabIndex = 347;
			this.lblEstado.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label33
			// 
			this.label33.AutoSize = true;
			this.label33.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label33.Location = new System.Drawing.Point(736, 491);
			this.label33.Name = "label33";
			this.label33.Size = new System.Drawing.Size(36, 15);
			this.label33.TabIndex = 346;
			this.label33.Text = "Egreso";
			this.label33.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.label33.Visible = false;
			// 
			// label32
			// 
			this.label32.AutoSize = true;
			this.label32.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label32.Location = new System.Drawing.Point(944, 11);
			this.label32.Name = "label32";
			this.label32.Size = new System.Drawing.Size(38, 15);
			this.label32.TabIndex = 345;
			this.label32.Text = "Ingreso";
			this.label32.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// lblEgreso
			// 
			this.lblEgreso.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
			this.lblEgreso.Font = new System.Drawing.Font("Microsoft Sans Serif", 6.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblEgreso.Location = new System.Drawing.Point(784, 488);
			this.lblEgreso.Name = "lblEgreso";
			this.lblEgreso.Size = new System.Drawing.Size(96, 21);
			this.lblEgreso.TabIndex = 344;
			this.lblEgreso.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.lblEgreso.Visible = false;
			// 
			// gbProveedor
			// 
			this.gbProveedor.Controls.Add(this.chkGarantiaExtendida);
			this.gbProveedor.Controls.Add(this.cmbGarantiaProveedor);
			this.gbProveedor.Controls.Add(this.label16);
			this.gbProveedor.Controls.Add(this.txtTiempGarantiaProveedr);
			this.gbProveedor.Controls.Add(this.label17);
			this.gbProveedor.Controls.Add(this.txtGrupoProveedor);
			this.gbProveedor.Controls.Add(this.label37);
			this.gbProveedor.Controls.Add(this.label12);
			this.gbProveedor.Controls.Add(this.cmbFechaFactura);
			this.gbProveedor.Controls.Add(this.txtProveedor);
			this.gbProveedor.Controls.Add(this.txtNumeroDeFactura);
			this.gbProveedor.Controls.Add(this.label10);
			this.gbProveedor.Controls.Add(this.label11);
			this.gbProveedor.Location = new System.Drawing.Point(776, 40);
			this.gbProveedor.Name = "gbProveedor";
			this.gbProveedor.Size = new System.Drawing.Size(312, 160);
			this.gbProveedor.TabIndex = 343;
			this.gbProveedor.TabStop = false;
			this.gbProveedor.Text = "Proveedor";
			this.gbProveedor.Visible = false;
			// 
			// chkGarantiaExtendida
			// 
			this.chkGarantiaExtendida.CheckAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.chkGarantiaExtendida.Enabled = false;
			this.chkGarantiaExtendida.Font = new System.Drawing.Font("Tahoma", 7.5F);
			this.chkGarantiaExtendida.Location = new System.Drawing.Point(8, 136);
			this.chkGarantiaExtendida.Name = "chkGarantiaExtendida";
			this.chkGarantiaExtendida.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
			this.chkGarantiaExtendida.Size = new System.Drawing.Size(112, 16);
			this.chkGarantiaExtendida.TabIndex = 291;
			this.chkGarantiaExtendida.Text = "Garantía Extendida";
			this.chkGarantiaExtendida.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
			// 
			// cmbGarantiaProveedor
			// 
			this.cmbGarantiaProveedor.Enabled = false;
			this.cmbGarantiaProveedor.Items.AddRange(new object[] {
																															"SI",
																															"NO"});
			this.cmbGarantiaProveedor.Location = new System.Drawing.Point(72, 106);
			this.cmbGarantiaProveedor.Name = "cmbGarantiaProveedor";
			this.cmbGarantiaProveedor.Size = new System.Drawing.Size(72, 21);
			this.cmbGarantiaProveedor.TabIndex = 290;
			// 
			// label16
			// 
			this.label16.AutoSize = true;
			this.label16.BackColor = System.Drawing.Color.Transparent;
			this.label16.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label16.Location = new System.Drawing.Point(164, 109);
			this.label16.Name = "label16";
			this.label16.Size = new System.Drawing.Size(81, 15);
			this.label16.TabIndex = 289;
			this.label16.Text = "Tiempo Garantia";
			this.label16.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtTiempGarantiaProveedr
			// 
			appearance23.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtTiempGarantiaProveedr.Appearance = appearance23;
			this.txtTiempGarantiaProveedr.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtTiempGarantiaProveedr.Enabled = false;
			this.txtTiempGarantiaProveedr.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtTiempGarantiaProveedr.FormatString = "";
			this.txtTiempGarantiaProveedr.Location = new System.Drawing.Point(260, 106);
			this.txtTiempGarantiaProveedr.Name = "txtTiempGarantiaProveedr";
			this.txtTiempGarantiaProveedr.PromptChar = ' ';
			this.txtTiempGarantiaProveedr.Size = new System.Drawing.Size(44, 20);
			this.txtTiempGarantiaProveedr.TabIndex = 288;
			// 
			// label17
			// 
			this.label17.AutoSize = true;
			this.label17.BackColor = System.Drawing.Color.Transparent;
			this.label17.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label17.Location = new System.Drawing.Point(8, 109);
			this.label17.Name = "label17";
			this.label17.Size = new System.Drawing.Size(43, 15);
			this.label17.TabIndex = 287;
			this.label17.Text = "Garantía";
			this.label17.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtGrupoProveedor
			// 
			appearance24.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtGrupoProveedor.Appearance = appearance24;
			this.txtGrupoProveedor.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtGrupoProveedor.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtGrupoProveedor.Enabled = false;
			this.txtGrupoProveedor.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtGrupoProveedor.Location = new System.Drawing.Point(72, 84);
			this.txtGrupoProveedor.Name = "txtGrupoProveedor";
			this.txtGrupoProveedor.Size = new System.Drawing.Size(232, 20);
			this.txtGrupoProveedor.TabIndex = 259;
			// 
			// label37
			// 
			this.label37.AutoSize = true;
			this.label37.BackColor = System.Drawing.Color.Transparent;
			this.label37.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label37.Location = new System.Drawing.Point(8, 87);
			this.label37.Name = "label37";
			this.label37.Size = new System.Drawing.Size(32, 15);
			this.label37.TabIndex = 258;
			this.label37.Text = "Grupo";
			this.label37.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label12
			// 
			this.label12.AutoSize = true;
			this.label12.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label12.Location = new System.Drawing.Point(8, 65);
			this.label12.Name = "label12";
			this.label12.Size = new System.Drawing.Size(70, 15);
			this.label12.TabIndex = 257;
			this.label12.Text = "Fecha Factura";
			this.label12.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// cmbFechaFactura
			// 
			appearance25.ForeColor = System.Drawing.Color.Black;
			appearance25.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbFechaFactura.Appearance = appearance25;
			this.cmbFechaFactura.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton2.Caption = "Today";
			this.cmbFechaFactura.DateButtons.Add(dateButton2);
			this.cmbFechaFactura.Enabled = false;
			this.cmbFechaFactura.Format = "dd/MM/yyyy";
			this.cmbFechaFactura.Location = new System.Drawing.Point(88, 62);
			this.cmbFechaFactura.Name = "cmbFechaFactura";
			this.cmbFechaFactura.NonAutoSizeHeight = 23;
			this.cmbFechaFactura.Size = new System.Drawing.Size(128, 21);
			this.cmbFechaFactura.SpinButtonsVisible = true;
			this.cmbFechaFactura.TabIndex = 256;
			this.cmbFechaFactura.Value = ((object)(resources.GetObject("cmbFechaFactura.Value")));
			// 
			// txtProveedor
			// 
			appearance26.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtProveedor.Appearance = appearance26;
			this.txtProveedor.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtProveedor.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtProveedor.Enabled = false;
			this.txtProveedor.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtProveedor.Location = new System.Drawing.Point(72, 18);
			this.txtProveedor.Name = "txtProveedor";
			this.txtProveedor.Size = new System.Drawing.Size(232, 20);
			this.txtProveedor.TabIndex = 255;
			// 
			// txtNumeroDeFactura
			// 
			appearance27.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtNumeroDeFactura.Appearance = appearance27;
			this.txtNumeroDeFactura.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtNumeroDeFactura.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNumeroDeFactura.Enabled = false;
			this.txtNumeroDeFactura.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtNumeroDeFactura.Location = new System.Drawing.Point(72, 40);
			this.txtNumeroDeFactura.Name = "txtNumeroDeFactura";
			this.txtNumeroDeFactura.Size = new System.Drawing.Size(128, 20);
			this.txtNumeroDeFactura.TabIndex = 254;
			// 
			// label10
			// 
			this.label10.AutoSize = true;
			this.label10.BackColor = System.Drawing.Color.Transparent;
			this.label10.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label10.Location = new System.Drawing.Point(8, 21);
			this.label10.Name = "label10";
			this.label10.Size = new System.Drawing.Size(51, 15);
			this.label10.TabIndex = 253;
			this.label10.Text = "Proveedor";
			this.label10.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label11
			// 
			this.label11.AutoSize = true;
			this.label11.BackColor = System.Drawing.Color.Transparent;
			this.label11.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label11.Location = new System.Drawing.Point(8, 43);
			this.label11.Name = "label11";
			this.label11.Size = new System.Drawing.Size(52, 15);
			this.label11.TabIndex = 252;
			this.label11.Text = "N. Factura";
			this.label11.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtIdIngreso
			// 
			this.txtIdIngreso.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			appearance28.ForeColorDisabled = System.Drawing.Color.Black;
			appearance28.TextHAlign = Infragistics.Win.HAlign.Right;
			appearance28.TextVAlign = Infragistics.Win.VAlign.Middle;
			this.txtIdIngreso.Appearance = appearance28;
			this.txtIdIngreso.Enabled = false;
			this.txtIdIngreso.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtIdIngreso.FormatString = "";
			this.txtIdIngreso.Location = new System.Drawing.Point(892, 498);
			this.txtIdIngreso.Name = "txtIdIngreso";
			this.txtIdIngreso.PromptChar = ' ';
			this.txtIdIngreso.Size = new System.Drawing.Size(25, 21);
			this.txtIdIngreso.TabIndex = 335;
			this.txtIdIngreso.Visible = false;
			// 
			// lblIngreso
			// 
			this.lblIngreso.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
			this.lblIngreso.Font = new System.Drawing.Font("Microsoft Sans Serif", 6.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblIngreso.Location = new System.Drawing.Point(992, 8);
			this.lblIngreso.Name = "lblIngreso";
			this.lblIngreso.Size = new System.Drawing.Size(96, 21);
			this.lblIngreso.TabIndex = 336;
			this.lblIngreso.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label4
			// 
			this.label4.AutoSize = true;
			this.label4.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label4.Location = new System.Drawing.Point(8, 11);
			this.label4.Name = "label4";
			this.label4.Size = new System.Drawing.Size(32, 15);
			this.label4.TabIndex = 332;
			this.label4.Text = "Fecha";
			this.label4.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// dtFecha
			// 
			this.dtFecha.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton3.Caption = "Today";
			this.dtFecha.DateButtons.Add(dateButton3);
			this.dtFecha.Enabled = false;
			this.dtFecha.Format = "dd/MM/yyyy";
			this.dtFecha.Location = new System.Drawing.Point(56, 8);
			this.dtFecha.Name = "dtFecha";
			this.dtFecha.NonAutoSizeHeight = 23;
			this.dtFecha.Size = new System.Drawing.Size(112, 21);
			this.dtFecha.SpinButtonsVisible = true;
			this.dtFecha.TabIndex = 331;
			this.dtFecha.Value = ((object)(resources.GetObject("dtFecha.Value")));
			this.dtFecha.KeyDown += new System.Windows.Forms.KeyEventHandler(this.dtFecha_KeyDown);
			this.dtFecha.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.dtFecha_KeyPress);
			this.dtFecha.BeforeDropDown += new System.ComponentModel.CancelEventHandler(this.dtFecha_BeforeDropDown);
			this.dtFecha.ValueChanged += new System.EventHandler(this.dtFecha_ValueChanged);
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
			// txtEstado
			// 
			this.txtEstado.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			appearance29.ForeColorDisabled = System.Drawing.Color.Black;
			appearance29.TextHAlign = Infragistics.Win.HAlign.Right;
			appearance29.TextVAlign = Infragistics.Win.VAlign.Middle;
			this.txtEstado.Appearance = appearance29;
			this.txtEstado.Enabled = false;
			this.txtEstado.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtEstado.FormatString = "";
			this.txtEstado.Location = new System.Drawing.Point(1012, 498);
			this.txtEstado.Name = "txtEstado";
			this.txtEstado.PromptChar = ' ';
			this.txtEstado.Size = new System.Drawing.Size(25, 21);
			this.txtEstado.TabIndex = 364;
			this.txtEstado.Visible = false;
			// 
			// txtIdProveedor
			// 
			this.txtIdProveedor.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			appearance30.ForeColorDisabled = System.Drawing.Color.Black;
			appearance30.TextHAlign = Infragistics.Win.HAlign.Right;
			appearance30.TextVAlign = Infragistics.Win.VAlign.Middle;
			this.txtIdProveedor.Appearance = appearance30;
			this.txtIdProveedor.Enabled = false;
			this.txtIdProveedor.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtIdProveedor.FormatString = "";
			this.txtIdProveedor.Location = new System.Drawing.Point(1036, 498);
			this.txtIdProveedor.Name = "txtIdProveedor";
			this.txtIdProveedor.PromptChar = ' ';
			this.txtIdProveedor.Size = new System.Drawing.Size(25, 21);
			this.txtIdProveedor.TabIndex = 366;
			this.txtIdProveedor.Visible = false;
			// 
			// txtOrigen
			// 
			this.txtOrigen.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			appearance31.ForeColorDisabled = System.Drawing.Color.Black;
			appearance31.TextHAlign = Infragistics.Win.HAlign.Right;
			appearance31.TextVAlign = Infragistics.Win.VAlign.Middle;
			this.txtOrigen.Appearance = appearance31;
			this.txtOrigen.Enabled = false;
			this.txtOrigen.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtOrigen.FormatString = "";
			this.txtOrigen.Location = new System.Drawing.Point(1060, 498);
			this.txtOrigen.Name = "txtOrigen";
			this.txtOrigen.PromptChar = ' ';
			this.txtOrigen.Size = new System.Drawing.Size(25, 21);
			this.txtOrigen.TabIndex = 367;
			this.txtOrigen.Visible = false;
			// 
			// mnuImprimir
			// 
			this.mnuImprimir.MenuItems.AddRange(new System.Windows.Forms.MenuItem[] {
																																								this.mnuImprimirIngreso,
																																								this.mnuImprimirEgreso});
			this.mnuImprimir.Popup += new System.EventHandler(this.mnuImprimir_Popup);
			// 
			// mnuImprimirIngreso
			// 
			this.mnuImprimirIngreso.Index = 0;
			this.mnuImprimirIngreso.Text = "Ingreso";
			this.mnuImprimirIngreso.Click += new System.EventHandler(this.mnuImprimirIngreso_Click);
			// 
			// mnuImprimirEgreso
			// 
			this.mnuImprimirEgreso.Index = 1;
			this.mnuImprimirEgreso.Text = "Salida";
			this.mnuImprimirEgreso.Click += new System.EventHandler(this.mnuImprimirEgreso_Click);
			// 
			// groupBox1
			// 
			this.groupBox1.Controls.Add(this.txtEmpqsObservaciones);
			this.groupBox1.Controls.Add(this.chkEmpqs);
			this.groupBox1.Controls.Add(this.rbEmpaques);
			this.groupBox1.Font = new System.Drawing.Font("Tahoma", 7F);
			this.groupBox1.Location = new System.Drawing.Point(624, 200);
			this.groupBox1.Name = "groupBox1";
			this.groupBox1.Size = new System.Drawing.Size(232, 168);
			this.groupBox1.TabIndex = 368;
			this.groupBox1.TabStop = false;
			// 
			// txtEmpqsObservaciones
			// 
			appearance32.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtEmpqsObservaciones.Appearance = appearance32;
			this.txtEmpqsObservaciones.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtEmpqsObservaciones.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtEmpqsObservaciones.Enabled = false;
			this.txtEmpqsObservaciones.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtEmpqsObservaciones.Location = new System.Drawing.Point(8, 64);
			this.txtEmpqsObservaciones.Multiline = true;
			this.txtEmpqsObservaciones.Name = "txtEmpqsObservaciones";
			this.txtEmpqsObservaciones.Scrollbars = System.Windows.Forms.ScrollBars.Vertical;
			this.txtEmpqsObservaciones.Size = new System.Drawing.Size(216, 100);
			this.txtEmpqsObservaciones.TabIndex = 339;
			this.txtEmpqsObservaciones.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtEmpqsObservaciones_KeyDown);
			// 
			// chkEmpqs
			// 
			this.chkEmpqs.Enabled = false;
			this.chkEmpqs.Font = new System.Drawing.Font("Tahoma", 7.5F);
			this.chkEmpqs.Location = new System.Drawing.Point(0, 2);
			this.chkEmpqs.Name = "chkEmpqs";
			this.chkEmpqs.Size = new System.Drawing.Size(112, 24);
			this.chkEmpqs.TabIndex = 0;
			this.chkEmpqs.Text = "Empaques de Caja";
			this.chkEmpqs.KeyDown += new System.Windows.Forms.KeyEventHandler(this.chkEmpqs_KeyDown);
			this.chkEmpqs.CheckedChanged += new System.EventHandler(this.chkEmpqs_CheckedChanged);
			// 
			// rbEmpaques
			// 
			appearance33.FontData.SizeInPoints = 7.5F;
			appearance33.ForeColorDisabled = System.Drawing.Color.Black;
			this.rbEmpaques.Appearance = appearance33;
			this.rbEmpaques.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.rbEmpaques.Enabled = false;
			this.rbEmpaques.ItemAppearance = appearance34;
			this.rbEmpaques.ItemOrigin = new System.Drawing.Point(10, 0);
			valueListItem1.DataValue = 1;
			valueListItem1.DisplayText = "Bueno";
			valueListItem2.DataValue = 2;
			valueListItem2.DisplayText = "Malo";
			valueListItem3.DataValue = 3;
			valueListItem3.DisplayText = "Novedades";
			this.rbEmpaques.Items.Add(valueListItem1);
			this.rbEmpaques.Items.Add(valueListItem2);
			this.rbEmpaques.Items.Add(valueListItem3);
			this.rbEmpaques.ItemSpacingVertical = 10;
			this.rbEmpaques.Location = new System.Drawing.Point(8, 32);
			this.rbEmpaques.Name = "rbEmpaques";
			this.rbEmpaques.Size = new System.Drawing.Size(216, 24);
			this.rbEmpaques.TabIndex = 383;
			// 
			// groupBox3
			// 
			this.groupBox3.Controls.Add(this.rbSerial);
			this.groupBox3.Controls.Add(this.rbBotones);
			this.groupBox3.Controls.Add(this.rbPines);
			this.groupBox3.Controls.Add(this.rbSupPosterior);
			this.groupBox3.Controls.Add(this.rbSupDelantera);
			this.groupBox3.Controls.Add(this.rbBordes);
			this.groupBox3.Controls.Add(this.chkSerial);
			this.groupBox3.Controls.Add(this.chkBotones);
			this.groupBox3.Controls.Add(this.chkPines);
			this.groupBox3.Controls.Add(this.chkSupPosterior);
			this.groupBox3.Controls.Add(this.chkSupDelantera);
			this.groupBox3.Controls.Add(this.txtEstadoFisico);
			this.groupBox3.Controls.Add(this.chkBordes);
			this.groupBox3.Font = new System.Drawing.Font("Tahoma", 7F);
			this.groupBox3.Location = new System.Drawing.Point(5, 200);
			this.groupBox3.Name = "groupBox3";
			this.groupBox3.Size = new System.Drawing.Size(614, 168);
			this.groupBox3.TabIndex = 369;
			this.groupBox3.TabStop = false;
			this.groupBox3.Text = "Estado Físico del Articulo";
			// 
			// rbSerial
			// 
			appearance35.ForeColorDisabled = System.Drawing.Color.Black;
			this.rbSerial.Appearance = appearance35;
			this.rbSerial.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.rbSerial.Enabled = false;
			this.rbSerial.ItemAppearance = appearance36;
			this.rbSerial.ItemOrigin = new System.Drawing.Point(10, 0);
			valueListItem4.DataValue = 1;
			valueListItem4.DisplayText = "B";
			valueListItem5.DataValue = 2;
			valueListItem5.DisplayText = "M";
			valueListItem6.DataValue = 3;
			valueListItem6.DisplayText = "R";
			valueListItem7.DataValue = 4;
			valueListItem7.DisplayText = "ROT";
			valueListItem8.DataValue = 5;
			valueListItem8.DisplayText = "NOV";
			this.rbSerial.Items.Add(valueListItem4);
			this.rbSerial.Items.Add(valueListItem5);
			this.rbSerial.Items.Add(valueListItem6);
			this.rbSerial.Items.Add(valueListItem7);
			this.rbSerial.Items.Add(valueListItem8);
			this.rbSerial.ItemSpacingVertical = 10;
			this.rbSerial.Location = new System.Drawing.Point(392, 72);
			this.rbSerial.Name = "rbSerial";
			this.rbSerial.Size = new System.Drawing.Size(215, 24);
			this.rbSerial.TabIndex = 356;
			// 
			// rbBotones
			// 
			appearance37.ForeColorDisabled = System.Drawing.Color.Black;
			this.rbBotones.Appearance = appearance37;
			this.rbBotones.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.rbBotones.Enabled = false;
			this.rbBotones.ItemAppearance = appearance38;
			this.rbBotones.ItemOrigin = new System.Drawing.Point(10, 0);
			valueListItem9.DataValue = 1;
			valueListItem9.DisplayText = "B";
			valueListItem10.DataValue = 2;
			valueListItem10.DisplayText = "M";
			valueListItem11.DataValue = 3;
			valueListItem11.DisplayText = "R";
			valueListItem12.DataValue = 4;
			valueListItem12.DisplayText = "ROT";
			valueListItem13.DataValue = 5;
			valueListItem13.DisplayText = "NOV";
			this.rbBotones.Items.Add(valueListItem9);
			this.rbBotones.Items.Add(valueListItem10);
			this.rbBotones.Items.Add(valueListItem11);
			this.rbBotones.Items.Add(valueListItem12);
			this.rbBotones.Items.Add(valueListItem13);
			this.rbBotones.ItemSpacingVertical = 10;
			this.rbBotones.Location = new System.Drawing.Point(72, 72);
			this.rbBotones.Name = "rbBotones";
			this.rbBotones.Size = new System.Drawing.Size(215, 24);
			this.rbBotones.TabIndex = 355;
			// 
			// rbPines
			// 
			appearance39.ForeColorDisabled = System.Drawing.Color.Black;
			this.rbPines.Appearance = appearance39;
			this.rbPines.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.rbPines.Enabled = false;
			this.rbPines.ItemAppearance = appearance40;
			this.rbPines.ItemOrigin = new System.Drawing.Point(10, 0);
			valueListItem14.DataValue = 1;
			valueListItem14.DisplayText = "B";
			valueListItem15.DataValue = 2;
			valueListItem15.DisplayText = "M";
			valueListItem16.DataValue = 3;
			valueListItem16.DisplayText = "R";
			valueListItem17.DataValue = 4;
			valueListItem17.DisplayText = "ROT";
			valueListItem18.DataValue = 5;
			valueListItem18.DisplayText = "NOV";
			this.rbPines.Items.Add(valueListItem14);
			this.rbPines.Items.Add(valueListItem15);
			this.rbPines.Items.Add(valueListItem16);
			this.rbPines.Items.Add(valueListItem17);
			this.rbPines.Items.Add(valueListItem18);
			this.rbPines.ItemSpacingVertical = 10;
			this.rbPines.Location = new System.Drawing.Point(72, 46);
			this.rbPines.Name = "rbPines";
			this.rbPines.Size = new System.Drawing.Size(215, 24);
			this.rbPines.TabIndex = 354;
			// 
			// rbSupPosterior
			// 
			appearance41.ForeColorDisabled = System.Drawing.Color.Black;
			this.rbSupPosterior.Appearance = appearance41;
			this.rbSupPosterior.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.rbSupPosterior.Enabled = false;
			this.rbSupPosterior.ItemAppearance = appearance42;
			this.rbSupPosterior.ItemOrigin = new System.Drawing.Point(10, 0);
			valueListItem19.DataValue = 1;
			valueListItem19.DisplayText = "B";
			valueListItem20.DataValue = 2;
			valueListItem20.DisplayText = "M";
			valueListItem21.DataValue = 3;
			valueListItem21.DisplayText = "R";
			valueListItem22.DataValue = 4;
			valueListItem22.DisplayText = "ROT";
			valueListItem23.DataValue = 5;
			valueListItem23.DisplayText = "NOV";
			this.rbSupPosterior.Items.Add(valueListItem19);
			this.rbSupPosterior.Items.Add(valueListItem20);
			this.rbSupPosterior.Items.Add(valueListItem21);
			this.rbSupPosterior.Items.Add(valueListItem22);
			this.rbSupPosterior.Items.Add(valueListItem23);
			this.rbSupPosterior.ItemSpacingVertical = 10;
			this.rbSupPosterior.Location = new System.Drawing.Point(392, 46);
			this.rbSupPosterior.Name = "rbSupPosterior";
			this.rbSupPosterior.Size = new System.Drawing.Size(215, 24);
			this.rbSupPosterior.TabIndex = 353;
			// 
			// rbSupDelantera
			// 
			appearance43.FontData.SizeInPoints = 7F;
			appearance43.ForeColorDisabled = System.Drawing.Color.Black;
			this.rbSupDelantera.Appearance = appearance43;
			this.rbSupDelantera.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.rbSupDelantera.Enabled = false;
			this.rbSupDelantera.ItemAppearance = appearance44;
			this.rbSupDelantera.ItemOrigin = new System.Drawing.Point(10, 0);
			valueListItem24.DataValue = 1;
			valueListItem24.DisplayText = "B";
			valueListItem25.DataValue = 2;
			valueListItem25.DisplayText = "M";
			valueListItem26.DataValue = 3;
			valueListItem26.DisplayText = "R";
			valueListItem27.DataValue = 4;
			valueListItem27.DisplayText = "ROT";
			valueListItem28.DataValue = 5;
			valueListItem28.DisplayText = "NOV";
			this.rbSupDelantera.Items.Add(valueListItem24);
			this.rbSupDelantera.Items.Add(valueListItem25);
			this.rbSupDelantera.Items.Add(valueListItem26);
			this.rbSupDelantera.Items.Add(valueListItem27);
			this.rbSupDelantera.Items.Add(valueListItem28);
			this.rbSupDelantera.ItemSpacingVertical = 10;
			this.rbSupDelantera.Location = new System.Drawing.Point(392, 20);
			this.rbSupDelantera.Name = "rbSupDelantera";
			this.rbSupDelantera.Size = new System.Drawing.Size(215, 24);
			this.rbSupDelantera.TabIndex = 352;
			// 
			// rbBordes
			// 
			appearance45.ForeColorDisabled = System.Drawing.Color.Black;
			this.rbBordes.Appearance = appearance45;
			this.rbBordes.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.rbBordes.Enabled = false;
			this.rbBordes.ItemAppearance = appearance46;
			this.rbBordes.ItemOrigin = new System.Drawing.Point(10, 0);
			valueListItem29.DataValue = 1;
			valueListItem29.DisplayText = "B";
			valueListItem30.DataValue = 2;
			valueListItem30.DisplayText = "M";
			valueListItem31.DataValue = 3;
			valueListItem31.DisplayText = "R";
			valueListItem32.DataValue = 4;
			valueListItem32.DisplayText = "ROT";
			valueListItem33.DataValue = 5;
			valueListItem33.DisplayText = "NOV";
			this.rbBordes.Items.Add(valueListItem29);
			this.rbBordes.Items.Add(valueListItem30);
			this.rbBordes.Items.Add(valueListItem31);
			this.rbBordes.Items.Add(valueListItem32);
			this.rbBordes.Items.Add(valueListItem33);
			this.rbBordes.ItemSpacingVertical = 10;
			this.rbBordes.Location = new System.Drawing.Point(72, 20);
			this.rbBordes.Name = "rbBordes";
			this.rbBordes.Size = new System.Drawing.Size(215, 24);
			this.rbBordes.TabIndex = 351;
			// 
			// chkSerial
			// 
			this.chkSerial.CheckAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.chkSerial.Enabled = false;
			this.chkSerial.Font = new System.Drawing.Font("Tahoma", 7.5F);
			this.chkSerial.Location = new System.Drawing.Point(296, 76);
			this.chkSerial.Name = "chkSerial";
			this.chkSerial.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
			this.chkSerial.Size = new System.Drawing.Size(96, 16);
			this.chkSerial.TabIndex = 349;
			this.chkSerial.Text = "Serial Producto";
			this.chkSerial.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.chkSerial.CheckedChanged += new System.EventHandler(this.chkSerial_CheckedChanged);
			// 
			// chkBotones
			// 
			this.chkBotones.CheckAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.chkBotones.Enabled = false;
			this.chkBotones.Font = new System.Drawing.Font("Tahoma", 7.5F);
			this.chkBotones.Location = new System.Drawing.Point(8, 76);
			this.chkBotones.Name = "chkBotones";
			this.chkBotones.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
			this.chkBotones.Size = new System.Drawing.Size(64, 16);
			this.chkBotones.TabIndex = 347;
			this.chkBotones.Text = "Botones";
			this.chkBotones.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.chkBotones.CheckedChanged += new System.EventHandler(this.chkBotones_CheckedChanged);
			// 
			// chkPines
			// 
			this.chkPines.CheckAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.chkPines.Enabled = false;
			this.chkPines.Font = new System.Drawing.Font("Tahoma", 7.5F);
			this.chkPines.Location = new System.Drawing.Point(8, 50);
			this.chkPines.Name = "chkPines";
			this.chkPines.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
			this.chkPines.Size = new System.Drawing.Size(64, 16);
			this.chkPines.TabIndex = 345;
			this.chkPines.Text = "Pines";
			this.chkPines.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.chkPines.CheckedChanged += new System.EventHandler(this.chkPines_CheckedChanged);
			// 
			// chkSupPosterior
			// 
			this.chkSupPosterior.CheckAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.chkSupPosterior.Enabled = false;
			this.chkSupPosterior.Font = new System.Drawing.Font("Tahoma", 7.5F);
			this.chkSupPosterior.Location = new System.Drawing.Point(296, 50);
			this.chkSupPosterior.Name = "chkSupPosterior";
			this.chkSupPosterior.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
			this.chkSupPosterior.Size = new System.Drawing.Size(96, 16);
			this.chkSupPosterior.TabIndex = 343;
			this.chkSupPosterior.Text = "Sup. Posterior";
			this.chkSupPosterior.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.chkSupPosterior.CheckedChanged += new System.EventHandler(this.chkSupPosterior_CheckedChanged);
			// 
			// chkSupDelantera
			// 
			this.chkSupDelantera.CheckAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.chkSupDelantera.Enabled = false;
			this.chkSupDelantera.Font = new System.Drawing.Font("Tahoma", 7.5F);
			this.chkSupDelantera.Location = new System.Drawing.Point(296, 24);
			this.chkSupDelantera.Name = "chkSupDelantera";
			this.chkSupDelantera.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
			this.chkSupDelantera.Size = new System.Drawing.Size(96, 16);
			this.chkSupDelantera.TabIndex = 341;
			this.chkSupDelantera.Text = "Sup. Delantera";
			this.chkSupDelantera.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.chkSupDelantera.CheckedChanged += new System.EventHandler(this.chkSupDelantera_CheckedChanged);
			// 
			// txtEstadoFisico
			// 
			appearance47.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtEstadoFisico.Appearance = appearance47;
			this.txtEstadoFisico.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtEstadoFisico.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtEstadoFisico.Enabled = false;
			this.txtEstadoFisico.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtEstadoFisico.Location = new System.Drawing.Point(8, 104);
			this.txtEstadoFisico.Multiline = true;
			this.txtEstadoFisico.Name = "txtEstadoFisico";
			this.txtEstadoFisico.Scrollbars = System.Windows.Forms.ScrollBars.Vertical;
			this.txtEstadoFisico.Size = new System.Drawing.Size(600, 60);
			this.txtEstadoFisico.TabIndex = 339;
			this.txtEstadoFisico.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtEstadoFisico_KeyDown);
			// 
			// chkBordes
			// 
			this.chkBordes.CheckAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.chkBordes.Enabled = false;
			this.chkBordes.Font = new System.Drawing.Font("Tahoma", 7.5F);
			this.chkBordes.Location = new System.Drawing.Point(8, 24);
			this.chkBordes.Name = "chkBordes";
			this.chkBordes.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
			this.chkBordes.Size = new System.Drawing.Size(64, 16);
			this.chkBordes.TabIndex = 0;
			this.chkBordes.Text = "Bordes";
			this.chkBordes.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.chkBordes.CheckedChanged += new System.EventHandler(this.chkBordes_CheckedChanged);
			// 
			// groupBox4
			// 
			this.groupBox4.Controls.Add(this.txtAccesorios);
			this.groupBox4.Controls.Add(this.chkAccesorios);
			this.groupBox4.Controls.Add(this.rbAccesorios);
			this.groupBox4.Font = new System.Drawing.Font("Tahoma", 7F);
			this.groupBox4.Location = new System.Drawing.Point(856, 200);
			this.groupBox4.Name = "groupBox4";
			this.groupBox4.Size = new System.Drawing.Size(232, 168);
			this.groupBox4.TabIndex = 369;
			this.groupBox4.TabStop = false;
			// 
			// txtAccesorios
			// 
			appearance48.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtAccesorios.Appearance = appearance48;
			this.txtAccesorios.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtAccesorios.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtAccesorios.Enabled = false;
			this.txtAccesorios.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtAccesorios.Location = new System.Drawing.Point(8, 64);
			this.txtAccesorios.Multiline = true;
			this.txtAccesorios.Name = "txtAccesorios";
			this.txtAccesorios.Scrollbars = System.Windows.Forms.ScrollBars.Vertical;
			this.txtAccesorios.Size = new System.Drawing.Size(216, 100);
			this.txtAccesorios.TabIndex = 339;
			this.txtAccesorios.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtAccesorios_KeyDown);
			// 
			// chkAccesorios
			// 
			this.chkAccesorios.Enabled = false;
			this.chkAccesorios.Font = new System.Drawing.Font("Tahoma", 7.5F);
			this.chkAccesorios.Location = new System.Drawing.Point(0, 2);
			this.chkAccesorios.Name = "chkAccesorios";
			this.chkAccesorios.Size = new System.Drawing.Size(72, 16);
			this.chkAccesorios.TabIndex = 0;
			this.chkAccesorios.Text = "Accesorios";
			this.chkAccesorios.KeyDown += new System.Windows.Forms.KeyEventHandler(this.chkAccesorios_KeyDown);
			this.chkAccesorios.CheckedChanged += new System.EventHandler(this.chkAccesorios_CheckedChanged);
			// 
			// rbAccesorios
			// 
			appearance49.FontData.SizeInPoints = 7.5F;
			appearance49.ForeColorDisabled = System.Drawing.Color.Black;
			this.rbAccesorios.Appearance = appearance49;
			this.rbAccesorios.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.rbAccesorios.Enabled = false;
			this.rbAccesorios.ItemAppearance = appearance50;
			this.rbAccesorios.ItemOrigin = new System.Drawing.Point(10, 0);
			valueListItem34.DataValue = 1;
			valueListItem34.DisplayText = "Bueno";
			valueListItem35.DataValue = 2;
			valueListItem35.DisplayText = "Malo";
			valueListItem36.DataValue = 3;
			valueListItem36.DisplayText = "Novedades";
			this.rbAccesorios.Items.Add(valueListItem34);
			this.rbAccesorios.Items.Add(valueListItem35);
			this.rbAccesorios.Items.Add(valueListItem36);
			this.rbAccesorios.ItemSpacingVertical = 10;
			this.rbAccesorios.Location = new System.Drawing.Point(8, 32);
			this.rbAccesorios.Name = "rbAccesorios";
			this.rbAccesorios.Size = new System.Drawing.Size(216, 24);
			this.rbAccesorios.TabIndex = 384;
			// 
			// groupBox5
			// 
			this.groupBox5.Controls.Add(this.button3);
			this.groupBox5.Controls.Add(this.button2);
			this.groupBox5.Controls.Add(this.button1);
			this.groupBox5.Controls.Add(this.btnImg3);
			this.groupBox5.Controls.Add(this.btnImg2);
			this.groupBox5.Controls.Add(this.btnImg1);
			this.groupBox5.Controls.Add(this.label24);
			this.groupBox5.Controls.Add(this.txtImagen3);
			this.groupBox5.Controls.Add(this.label23);
			this.groupBox5.Controls.Add(this.txtImagen2);
			this.groupBox5.Controls.Add(this.label3);
			this.groupBox5.Controls.Add(this.txtImagen1);
			this.groupBox5.Location = new System.Drawing.Point(416, 376);
			this.groupBox5.Name = "groupBox5";
			this.groupBox5.Size = new System.Drawing.Size(376, 98);
			this.groupBox5.TabIndex = 370;
			this.groupBox5.TabStop = false;
			this.groupBox5.Text = "Imagenes del Articulo";
			// 
			// button3
			// 
			this.button3.Image = ((System.Drawing.Image)(resources.GetObject("button3.Image")));
			this.button3.Location = new System.Drawing.Point(344, 63);
			this.button3.Name = "button3";
			this.button3.Size = new System.Drawing.Size(24, 20);
			this.button3.TabIndex = 280;
			this.toolTip1.SetToolTip(this.button3, "Visualizar imagen 3");
			this.button3.Click += new System.EventHandler(this.button3_Click);
			// 
			// button2
			// 
			this.button2.Image = ((System.Drawing.Image)(resources.GetObject("button2.Image")));
			this.button2.Location = new System.Drawing.Point(344, 41);
			this.button2.Name = "button2";
			this.button2.Size = new System.Drawing.Size(24, 20);
			this.button2.TabIndex = 279;
			this.toolTip1.SetToolTip(this.button2, "Visualizar imagen 2");
			this.button2.Click += new System.EventHandler(this.button2_Click);
			// 
			// button1
			// 
			this.button1.Image = ((System.Drawing.Image)(resources.GetObject("button1.Image")));
			this.button1.Location = new System.Drawing.Point(344, 20);
			this.button1.Name = "button1";
			this.button1.Size = new System.Drawing.Size(24, 20);
			this.button1.TabIndex = 278;
			this.toolTip1.SetToolTip(this.button1, "Visualizar imagen 1");
			this.button1.Click += new System.EventHandler(this.button1_Click);
			// 
			// btnImg3
			// 
			this.btnImg3.Enabled = false;
			this.btnImg3.Image = ((System.Drawing.Image)(resources.GetObject("btnImg3.Image")));
			this.btnImg3.Location = new System.Drawing.Point(320, 63);
			this.btnImg3.Name = "btnImg3";
			this.btnImg3.Size = new System.Drawing.Size(24, 20);
			this.btnImg3.TabIndex = 277;
			this.toolTip1.SetToolTip(this.btnImg3, "Cargar imagen...");
			this.btnImg3.Click += new System.EventHandler(this.btnImg3_Click);
			// 
			// btnImg2
			// 
			this.btnImg2.Enabled = false;
			this.btnImg2.Image = ((System.Drawing.Image)(resources.GetObject("btnImg2.Image")));
			this.btnImg2.Location = new System.Drawing.Point(320, 41);
			this.btnImg2.Name = "btnImg2";
			this.btnImg2.Size = new System.Drawing.Size(24, 20);
			this.btnImg2.TabIndex = 276;
			this.toolTip1.SetToolTip(this.btnImg2, "Cargar Imagen...");
			this.btnImg2.Click += new System.EventHandler(this.btnImg2_Click);
			// 
			// btnImg1
			// 
			this.btnImg1.Enabled = false;
			this.btnImg1.Image = ((System.Drawing.Image)(resources.GetObject("btnImg1.Image")));
			this.btnImg1.Location = new System.Drawing.Point(320, 20);
			this.btnImg1.Name = "btnImg1";
			this.btnImg1.Size = new System.Drawing.Size(24, 20);
			this.btnImg1.TabIndex = 275;
			this.toolTip1.SetToolTip(this.btnImg1, "Cargar Imagen...");
			this.btnImg1.Click += new System.EventHandler(this.btnImg1_Click);
			// 
			// label24
			// 
			this.label24.AutoSize = true;
			this.label24.BackColor = System.Drawing.Color.Transparent;
			this.label24.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label24.Location = new System.Drawing.Point(8, 66);
			this.label24.Name = "label24";
			this.label24.Size = new System.Drawing.Size(47, 15);
			this.label24.TabIndex = 274;
			this.label24.Text = "Imagen 3";
			this.label24.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtImagen3
			// 
			appearance51.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtImagen3.Appearance = appearance51;
			this.txtImagen3.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtImagen3.Enabled = false;
			this.txtImagen3.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtImagen3.Location = new System.Drawing.Point(64, 63);
			this.txtImagen3.MaxLength = 13;
			this.txtImagen3.Name = "txtImagen3";
			this.txtImagen3.Size = new System.Drawing.Size(250, 20);
			this.txtImagen3.TabIndex = 273;
			// 
			// label23
			// 
			this.label23.AutoSize = true;
			this.label23.BackColor = System.Drawing.Color.Transparent;
			this.label23.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label23.Location = new System.Drawing.Point(8, 44);
			this.label23.Name = "label23";
			this.label23.Size = new System.Drawing.Size(47, 15);
			this.label23.TabIndex = 272;
			this.label23.Text = "Imagen 2";
			this.label23.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtImagen2
			// 
			appearance52.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtImagen2.Appearance = appearance52;
			this.txtImagen2.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtImagen2.Enabled = false;
			this.txtImagen2.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtImagen2.Location = new System.Drawing.Point(64, 41);
			this.txtImagen2.MaxLength = 13;
			this.txtImagen2.Name = "txtImagen2";
			this.txtImagen2.Size = new System.Drawing.Size(250, 20);
			this.txtImagen2.TabIndex = 271;
			// 
			// label3
			// 
			this.label3.AutoSize = true;
			this.label3.BackColor = System.Drawing.Color.Transparent;
			this.label3.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label3.Location = new System.Drawing.Point(8, 23);
			this.label3.Name = "label3";
			this.label3.Size = new System.Drawing.Size(47, 15);
			this.label3.TabIndex = 270;
			this.label3.Text = "Imagen 1";
			this.label3.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtImagen1
			// 
			appearance53.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtImagen1.Appearance = appearance53;
			this.txtImagen1.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtImagen1.Enabled = false;
			this.txtImagen1.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtImagen1.Location = new System.Drawing.Point(64, 20);
			this.txtImagen1.MaxLength = 13;
			this.txtImagen1.Name = "txtImagen1";
			this.txtImagen1.Size = new System.Drawing.Size(250, 20);
			this.txtImagen1.TabIndex = 269;
			// 
			// label25
			// 
			this.label25.AutoSize = true;
			this.label25.Font = new System.Drawing.Font("Tahoma", 7F, System.Drawing.FontStyle.Bold);
			this.label25.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(192)), ((System.Byte)(0)), ((System.Byte)(0)));
			this.label25.Location = new System.Drawing.Point(968, 568);
			this.label25.Name = "label25";
			this.label25.Size = new System.Drawing.Size(47, 15);
			this.label25.TabIndex = 371;
			this.label25.Text = "B: Bueno";
			this.label25.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label26
			// 
			this.label26.AutoSize = true;
			this.label26.Font = new System.Drawing.Font("Tahoma", 7F, System.Drawing.FontStyle.Bold);
			this.label26.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(192)), ((System.Byte)(0)), ((System.Byte)(0)));
			this.label26.Location = new System.Drawing.Point(968, 584);
			this.label26.Name = "label26";
			this.label26.Size = new System.Drawing.Size(42, 15);
			this.label26.TabIndex = 372;
			this.label26.Text = "M: Malo";
			this.label26.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label27
			// 
			this.label27.AutoSize = true;
			this.label27.Font = new System.Drawing.Font("Tahoma", 7F, System.Drawing.FontStyle.Bold);
			this.label27.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(192)), ((System.Byte)(0)), ((System.Byte)(0)));
			this.label27.Location = new System.Drawing.Point(968, 600);
			this.label27.Name = "label27";
			this.label27.Size = new System.Drawing.Size(53, 15);
			this.label27.TabIndex = 373;
			this.label27.Text = "R: Rayado";
			this.label27.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label28
			// 
			this.label28.AutoSize = true;
			this.label28.Font = new System.Drawing.Font("Tahoma", 7F, System.Drawing.FontStyle.Bold);
			this.label28.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(192)), ((System.Byte)(0)), ((System.Byte)(0)));
			this.label28.Location = new System.Drawing.Point(968, 616);
			this.label28.Name = "label28";
			this.label28.Size = new System.Drawing.Size(53, 15);
			this.label28.TabIndex = 374;
			this.label28.Text = "ROT: Roto";
			this.label28.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label29
			// 
			this.label29.AutoSize = true;
			this.label29.Font = new System.Drawing.Font("Tahoma", 7F, System.Drawing.FontStyle.Bold);
			this.label29.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(192)), ((System.Byte)(0)), ((System.Byte)(0)));
			this.label29.Location = new System.Drawing.Point(968, 632);
			this.label29.Name = "label29";
			this.label29.Size = new System.Drawing.Size(106, 15);
			this.label29.TabIndex = 375;
			this.label29.Text = "NOV: Con Novedades";
			this.label29.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label31
			// 
			this.label31.AutoSize = true;
			this.label31.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label31.Location = new System.Drawing.Point(432, 11);
			this.label31.Name = "label31";
			this.label31.Size = new System.Drawing.Size(36, 15);
			this.label31.TabIndex = 376;
			this.label31.Text = "Recibe";
			this.label31.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// cmbUsuario
			// 
			this.cmbUsuario.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbUsuario.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbUsuario.DataSource = this.ultraDataSource1;
			ultraGridColumn4.Header.VisiblePosition = 0;
			ultraGridColumn4.Hidden = true;
			ultraGridColumn5.Header.VisiblePosition = 1;
			ultraGridColumn5.Hidden = true;
			ultraGridColumn6.Header.VisiblePosition = 3;
			ultraGridColumn6.Hidden = true;
			ultraGridColumn7.Header.VisiblePosition = 2;
			ultraGridColumn7.Width = 300;
			ultraGridBand2.Columns.AddRange(new object[] {
																										 ultraGridColumn4,
																										 ultraGridColumn5,
																										 ultraGridColumn6,
																										 ultraGridColumn7});
			this.cmbUsuario.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			this.cmbUsuario.DisplayMember = "Nombre";
			this.cmbUsuario.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbUsuario.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbUsuario.Enabled = false;
			this.cmbUsuario.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbUsuario.Location = new System.Drawing.Point(480, 8);
			this.cmbUsuario.MaxDropDownItems = 30;
			this.cmbUsuario.Name = "cmbUsuario";
			this.cmbUsuario.Size = new System.Drawing.Size(200, 20);
			this.cmbUsuario.TabIndex = 377;
			this.cmbUsuario.ValueMember = "idPersonal";
			this.cmbUsuario.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cmbUsuario_KeyDown);
			// 
			// ultraDataSource1
			// 
			ultraDataColumn1.DataType = typeof(int);
			ultraDataColumn2.DataType = typeof(int);
			this.ultraDataSource1.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn1,
																																 ultraDataColumn2,
																																 ultraDataColumn3,
																																 ultraDataColumn4});
			// 
			// groupBox6
			// 
			this.groupBox6.Controls.Add(this.label42);
			this.groupBox6.Controls.Add(this.txtNombreEntrega);
			this.groupBox6.Controls.Add(this.label41);
			this.groupBox6.Controls.Add(this.txtCIEntrega);
			this.groupBox6.Location = new System.Drawing.Point(792, 376);
			this.groupBox6.Name = "groupBox6";
			this.groupBox6.Size = new System.Drawing.Size(296, 98);
			this.groupBox6.TabIndex = 378;
			this.groupBox6.TabStop = false;
			this.groupBox6.Text = "Quien Entrega.?";
			// 
			// label42
			// 
			this.label42.AutoSize = true;
			this.label42.BackColor = System.Drawing.Color.Transparent;
			this.label42.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label42.Location = new System.Drawing.Point(8, 44);
			this.label42.Name = "label42";
			this.label42.Size = new System.Drawing.Size(40, 15);
			this.label42.TabIndex = 383;
			this.label42.Text = "Nombre";
			this.label42.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtNombreEntrega
			// 
			appearance54.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtNombreEntrega.Appearance = appearance54;
			this.txtNombreEntrega.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtNombreEntrega.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNombreEntrega.Enabled = false;
			this.txtNombreEntrega.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtNombreEntrega.Location = new System.Drawing.Point(8, 63);
			this.txtNombreEntrega.Name = "txtNombreEntrega";
			this.txtNombreEntrega.Size = new System.Drawing.Size(280, 20);
			this.txtNombreEntrega.TabIndex = 382;
			this.txtNombreEntrega.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtNombreEntrega_KeyDown);
			this.txtNombreEntrega.Click += new System.EventHandler(this.txtNombreEntrega_Click);
			this.txtNombreEntrega.Enter += new System.EventHandler(this.txtNombreEntrega_Enter);
			// 
			// label41
			// 
			this.label41.AutoSize = true;
			this.label41.BackColor = System.Drawing.Color.Transparent;
			this.label41.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label41.Location = new System.Drawing.Point(8, 23);
			this.label41.Name = "label41";
			this.label41.Size = new System.Drawing.Size(37, 15);
			this.label41.TabIndex = 381;
			this.label41.Text = "Cedula";
			this.label41.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtCIEntrega
			// 
			appearance55.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtCIEntrega.Appearance = appearance55;
			this.txtCIEntrega.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtCIEntrega.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtCIEntrega.Enabled = false;
			this.txtCIEntrega.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtCIEntrega.Location = new System.Drawing.Point(56, 20);
			this.txtCIEntrega.MaxLength = 13;
			this.txtCIEntrega.Name = "txtCIEntrega";
			this.txtCIEntrega.TabIndex = 380;
			this.txtCIEntrega.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtCIEntrega_KeyDown);
			this.txtCIEntrega.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtCIEntrega_KeyPress);
			this.txtCIEntrega.Click += new System.EventHandler(this.txtCIEntrega_Click);
			this.txtCIEntrega.Enter += new System.EventHandler(this.txtCIEntrega_Enter);
			// 
			// label45
			// 
			this.label45.AutoSize = true;
			this.label45.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label45.Location = new System.Drawing.Point(688, 11);
			this.label45.Name = "label45";
			this.label45.Size = new System.Drawing.Size(40, 15);
			this.label45.TabIndex = 379;
			this.label45.Text = "Técnico";
			this.label45.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// cmbAsignaTecnico
			// 
			this.cmbAsignaTecnico.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbAsignaTecnico.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbAsignaTecnico.DataSource = this.ultraDataSource1;
			ultraGridColumn8.Header.VisiblePosition = 0;
			ultraGridColumn8.Hidden = true;
			ultraGridColumn9.Header.VisiblePosition = 1;
			ultraGridColumn9.Hidden = true;
			ultraGridColumn10.Header.VisiblePosition = 3;
			ultraGridColumn10.Hidden = true;
			ultraGridColumn11.Header.VisiblePosition = 2;
			ultraGridColumn11.Width = 300;
			ultraGridBand3.Columns.AddRange(new object[] {
																										 ultraGridColumn8,
																										 ultraGridColumn9,
																										 ultraGridColumn10,
																										 ultraGridColumn11});
			this.cmbAsignaTecnico.DisplayLayout.BandsSerializer.Add(ultraGridBand3);
			this.cmbAsignaTecnico.DisplayMember = "Nombre";
			this.cmbAsignaTecnico.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbAsignaTecnico.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbAsignaTecnico.Enabled = false;
			this.cmbAsignaTecnico.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbAsignaTecnico.Location = new System.Drawing.Point(736, 8);
			this.cmbAsignaTecnico.MaxDropDownItems = 30;
			this.cmbAsignaTecnico.Name = "cmbAsignaTecnico";
			this.cmbAsignaTecnico.Size = new System.Drawing.Size(200, 20);
			this.cmbAsignaTecnico.TabIndex = 380;
			this.cmbAsignaTecnico.ValueMember = "idPersonal";
			this.cmbAsignaTecnico.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cmbAsignaTecnico_KeyDown);
			// 
			// groupBox2
			// 
			this.groupBox2.Location = new System.Drawing.Point(0, 480);
			this.groupBox2.Name = "groupBox2";
			this.groupBox2.Size = new System.Drawing.Size(1096, 8);
			this.groupBox2.TabIndex = 381;
			this.groupBox2.TabStop = false;
			// 
			// frmDTIngreso
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 14);
			this.ClientSize = new System.Drawing.Size(1092, 520);
			this.Controls.Add(this.groupBox2);
			this.Controls.Add(this.label45);
			this.Controls.Add(this.label31);
			this.Controls.Add(this.label29);
			this.Controls.Add(this.label28);
			this.Controls.Add(this.label27);
			this.Controls.Add(this.label26);
			this.Controls.Add(this.label25);
			this.Controls.Add(this.label36);
			this.Controls.Add(this.label33);
			this.Controls.Add(this.label32);
			this.Controls.Add(this.label4);
			this.Controls.Add(this.cmbAsignaTecnico);
			this.Controls.Add(this.groupBox6);
			this.Controls.Add(this.gbProveedor);
			this.Controls.Add(this.cmbUsuario);
			this.Controls.Add(this.groupBox5);
			this.Controls.Add(this.groupBox1);
			this.Controls.Add(this.txtOrigen);
			this.Controls.Add(this.txtIdProveedor);
			this.Controls.Add(this.txtEstado);
			this.Controls.Add(this.label14);
			this.Controls.Add(this.txtDaño);
			this.Controls.Add(this.btnNuevo);
			this.Controls.Add(this.btnConsultar);
			this.Controls.Add(this.btnEditar);
			this.Controls.Add(this.btnImprimir);
			this.Controls.Add(this.btnGuardar);
			this.Controls.Add(this.lblUsuario);
			this.Controls.Add(this.txtBodegaPredef);
			this.Controls.Add(this.cmbBodega);
			this.Controls.Add(this.lblEstado);
			this.Controls.Add(this.lblEgreso);
			this.Controls.Add(this.txtIdIngreso);
			this.Controls.Add(this.lblIngreso);
			this.Controls.Add(this.dtFecha);
			this.Controls.Add(this.btnCancelar);
			this.Controls.Add(this.btnAnular);
			this.Controls.Add(this.btnSalir);
			this.Controls.Add(this.gbProducto);
			this.Controls.Add(this.gbCliente);
			this.Controls.Add(this.txtidCliente);
			this.Controls.Add(this.txtidArticulo);
			this.Controls.Add(this.txtIdSerial);
			this.Controls.Add(this.groupBox3);
			this.Controls.Add(this.groupBox4);
			this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
			this.KeyPreview = true;
			this.MaximizeBox = false;
			this.Name = "frmDTIngreso";
			this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
			this.Text = "Ingreso Departamento Técnico";
			this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.frmDTIngreso_KeyDown);
			this.Closing += new System.ComponentModel.CancelEventHandler(this.frmDTIngreso_Closing);
			this.Load += new System.EventHandler(this.frmDTIngreso_Load);
			this.gbProducto.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.txtLocal)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNumFCVenta)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbFechaVenta)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtSerial)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtTiempoDeGarantia)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCodArticulo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtArticulo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtVendedor)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIngresos)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdSerial)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtidArticulo)).EndInit();
			this.gbCliente.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.txtCelAdicional)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCorreo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtFonoAdicional)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCelular)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtTelefono)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDireccion)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCodigo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNombre)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtidCliente)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDaño)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtBodegaPredef)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbBodega)).EndInit();
			this.gbProveedor.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.txtTiempGarantiaProveedr)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtGrupoProveedor)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbFechaFactura)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtProveedor)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNumeroDeFactura)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdIngreso)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFecha)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtEstado)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdProveedor)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtOrigen)).EndInit();
			this.groupBox1.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.txtEmpqsObservaciones)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.rbEmpaques)).EndInit();
			this.groupBox3.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.rbSerial)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.rbBotones)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.rbPines)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.rbSupPosterior)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.rbSupDelantera)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.rbBordes)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtEstadoFisico)).EndInit();
			this.groupBox4.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.txtAccesorios)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.rbAccesorios)).EndInit();
			this.groupBox5.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.txtImagen3)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtImagen2)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtImagen1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbUsuario)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).EndInit();
			this.groupBox6.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.txtNombreEntrega)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCIEntrega)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbAsignaTecnico)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		int idFVenta = 0;
		int idFCompra = 0;
		int nBodega = 0;
//		int idIngresoDT = 0;
		int iPointCliente = 0; //1 - 2
		int iInternoExterno = 0;
		int idTipoRuc = 0;

		private void btnSalir_Click(object sender, System.EventArgs e)
		{
			this.Close();
		}

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);	
		}

		private void frmDTIngreso_Load(object sender, System.EventArgs e)
		{										
			miAcceso = new Acceso(cdsSeteoF, "1002");		

			if (miAcceso.Nuevo) this.btnNuevo.Enabled = true;
			if (miAcceso.Buscar) this.btnConsultar.Enabled = true;
			if (miAcceso.BVerProveedorDT) this.gbProveedor.Visible = true;
				
			string sSQL = string.Format("Exec BodegaAsignadaPorUsuario {0}", MenuLatinium.IdUsuario);		
			this.txtBodegaPredef.Value = Funcion.iEscalarSQL(cdsSeteoF, sSQL);

			this.dtFecha.CalendarInfo.MaxDate = DateTime.Today;
		}

		private void btnCancelar_Click(object sender, System.EventArgs e)
		{			
			#region General
			bAnulado = false;			
			iInternoExterno = 0;

			this.txtIdIngreso.Value = 0;
			this.txtidCliente.Value = 0;
			this.txtidArticulo.Value = 0;
			this.txtIdSerial.Value = 0;
			this.txtEstado.Value = 0;
			this.txtIdProveedor.Value = 0;
			this.txtOrigen.Value = 0;
			
			this.dtFecha.Value = System.DBNull.Value;
			this.cmbBodega.Value = System.DBNull.Value;
			this.cmbUsuario.Value = System.DBNull.Value;
			this.cmbAsignaTecnico.Value = System.DBNull.Value;
			this.lblIngreso.Text = "";
			this.lblEgreso.Text = "";
			this.lblEstado.Text = "";
			this.lblUsuario.Text = "";

			this.dtFecha.Enabled = false;
			this.cmbBodega.Enabled = false;
			this.cmbUsuario.Enabled = false;
			this.cmbAsignaTecnico.Enabled = false;
			#endregion General

			#region Cliente
			this.txtCodigo.Text = "";
			this.txtNombre.Text = "";
			this.txtDireccion.Text = "";
			this.txtTelefono.Text = "";
			this.txtCelular.Text = "";
			this.txtFonoAdicional.Text = "";
			this.txtCelAdicional.Text = "";
			this.txtCorreo.Text = "";

			this.txtCodigo.Enabled = false;
			this.txtNombre.Enabled = false;
			this.txtDireccion.Enabled = false;
			this.txtTelefono.Enabled = false;
			this.txtCelular.Enabled = false;
			this.txtFonoAdicional.Enabled = false;
			this.txtCelAdicional.Enabled = false;
			this.txtCorreo.Enabled = false;			
			#endregion Cliente

			#region Articulo
			this.txtSerial.Text = "";
			this.txtLocal.Text = "";
			this.txtCodArticulo.Text = "";
			this.txtIngresos.Value = 0;
			this.txtArticulo.Text = "";			
			this.cmbGarantia.Text = "";
			this.cmbGarantia.Items.Clear();
			this.txtTiempoDeGarantia.Value = 0;
			this.txtNumFCVenta.Text = "";
			this.cmbFechaVenta.Value = System.DBNull.Value;	
			this.txtVendedor.Text = "";

			this.txtSerial.Enabled = false;
			this.txtCodArticulo.Enabled = false;
			this.txtArticulo.Enabled = false;
			this.cmbGarantia.Enabled = false;
			this.chkGarantiaExtendida.Enabled = false;
			this.chkGarantiaExtendida.Checked = false;
			#endregion Articulo

			#region Proveedores
			this.txtProveedor.Text = "";
			this.txtNumeroDeFactura.Text = "";
			this.cmbFechaFactura.Value = System.DBNull.Value;
			this.txtGrupoProveedor.Text = "";
			this.cmbGarantiaProveedor.Text = "";
			this.txtTiempGarantiaProveedr.Value = 0;
			#endregion Proveedores

			#region Estado Fisico
			this.chkBordes.Checked = false;			
			this.chkPines.Checked = false;
			this.chkBotones.Checked = false;
			this.chkSupDelantera.Checked = false;
			this.chkSupPosterior.Checked = false;			
			this.chkSerial.Checked = false;

			this.chkBordes.Enabled = false;			
			this.chkPines.Enabled = false;
			this.chkBotones.Enabled = false;
			this.chkSupDelantera.Enabled = false;
			this.chkSupPosterior.Enabled = false;			
			this.chkSerial.Enabled = false;

			this.rbBordes.CheckedIndex = -1;
			this.rbPines.CheckedIndex = -1;
			this.rbBotones.CheckedIndex = -1;
			this.rbSupDelantera.CheckedIndex = -1;
			this.rbSupPosterior.CheckedIndex = -1;			
			this.rbSerial.CheckedIndex = -1;

			this.rbBordes.Enabled = false;
			this.rbPines.Enabled = false;
			this.rbBotones.Enabled = false;
			this.rbSupDelantera.Enabled = false;
			this.rbSupPosterior.Enabled = false;			
			this.rbSerial.Enabled = false;

			this.txtEstadoFisico.Text = "";
			this.txtEstadoFisico.Enabled = false;
			#endregion Estado Fisico

			#region Empaques - Accesorios
			this.chkEmpqs.Checked = false;			
			this.rbEmpaques.CheckedIndex = -1;
			this.txtEmpqsObservaciones.Text = "";

			this.chkEmpqs.Enabled = false;
			this.rbEmpaques.Enabled = false;
			this.txtEmpqsObservaciones.Enabled = false;

			this.chkAccesorios.Checked = false;
			this.rbAccesorios.CheckedIndex = -1;
			this.txtAccesorios.Text = "";

			this.chkAccesorios.Enabled = false;
			this.rbAccesorios.Enabled = false;
			this.txtAccesorios.Enabled = false;
			#endregion Empaques - Accesorios

			this.txtDaño.Text = "";
			this.txtDaño.Enabled = false;

			#region Imagenes
			this.txtImagen1.Text = "";
			this.txtImagen2.Text = "";
			this.txtImagen3.Text = "";

			this.btnImg1.Enabled = false;
			this.btnImg2.Enabled = false;
			this.btnImg3.Enabled = false;

			this.button1.Enabled = false;
			this.button2.Enabled = false;
			this.button3.Enabled = false;
			#endregion Imagenes
						
			#region Entrega
			this.txtCIEntrega.Text = "";
			this.txtNombreEntrega.Text = "";
			this.txtCIEntrega.Enabled = false;
			this.txtNombreEntrega.Enabled = false;
			#endregion Entrega			

			#region Botones
			this.btnCliente.Enabled = false;
			if (miAcceso.Nuevo) this.btnNuevo.Enabled = true;
			if (miAcceso.Buscar) this.btnConsultar.Enabled = true;
			this.btnEditar.Enabled = false;
			this.btnImprimir.Enabled = false;
			this.btnGuardar.Enabled = false;
			this.btnAnular.Enabled = false;
			this.btnCancelar.Enabled = false;		
			#endregion Botones

			bNuevo = false;
			bEdicion = false;

			 nuevoNombreArchivoImgone ="";
			 nuevoNombreArchivoImgTwo ="";
			 nuevoNombreArchivoImgTree ="";

			idFVenta = 0;
			idFCompra = 0;
			nBodega = 0;
			iPointCliente = 0;		
			idTipoRuc = 0;
		}

		private void btnNuevo_Click(object sender, System.EventArgs e)
		{
			this.btnCancelar_Click(sender, e);
			
			using (frmDTNuevoBusquedaArticulos Busqueda = new frmDTNuevoBusquedaArticulos())
			{ 
				if (DialogResult.OK == Busqueda.ShowDialog())
				{									
					#region Valores iniciales
					bNuevo = true;
					bEdicion = true;

					this.txtEstado.Value = 1;
					//FuncionesProcedimientos.FechasDocumentos(this.dtFecha);//this.dtFecha.Value = DateTime.Today;
					FuncionesProcedimientos.RetornaFechaServidor(this.dtFecha, cdsSeteoF, false);
					
					if (miAcceso.CambiarFecha) this.dtFecha.Enabled = true;
					
					if ((int)this.txtBodegaPredef.Value == 0) 
					{
						this.cmbBodega.Value = 20;
						this.cmbBodega.Enabled = true;
					}
					else this.cmbBodega.Value = (int)this.txtBodegaPredef.Value;

					if(miAcceso.BAsignaTecnicos)
					{					
						this.cmbAsignaTecnico.DataSource = Funcion.dvProcedimiento(cdsSeteoF, string.Format("Exec ListaUsuarios {0}", (int)this.cmbBodega.Value));				
						this.cmbAsignaTecnico.Enabled = true;								
					}

					this.cmbUsuario.DataSource = Funcion.dvProcedimiento(cdsSeteoF, string.Format("Exec ListaUsuarios {0}", (int)this.cmbBodega.Value));
					this.cmbUsuario.Enabled = true;					
					#endregion Valores iniciales

					#region Cliente
					if (miAcceso.BEditarNombre || (int)Busqueda.optOpciones.Value == 2) if (iPointCliente == 2)	this.txtCodigo.Enabled = true;
					#endregion Cliente	

					#region Valida tipo de ingreso
					iInternoExterno = (int)Busqueda.optOpciones.Value;

					this.cmbGarantia.Items.Add("NO");
					this.cmbGarantia.Items.Add("POR VERIFICAR");

					if ((int)Busqueda.optOpciones.Value == 1)
					{
						this.cmbGarantia.Text = "POR VERIFICAR";

						if (miAcceso.BModificaGarantia) this.cmbGarantia.Enabled = true;

						BuscaDatosSolicitud(Busqueda.ultraGrid2.ActiveRow.Cells["Serial"].Value.ToString());
					}
					else if ((int)Busqueda.optOpciones.Value == 2)
					{											
						this.cmbGarantia.Text = "NO";						
						this.txtTiempoDeGarantia.Value = 0;

						this.txtCodigo.Enabled = true;
						this.txtSerial.Enabled = true;
						this.txtCodArticulo.Enabled = true;
						this.txtArticulo.Enabled = true;
					
						iPointCliente = 3;
					}
					#endregion Valida tipo de ingreso

					#region Cliente - Estado Fisico - Accesorios
					this.txtFonoAdicional.Enabled = true;
					this.txtCelAdicional.Enabled = true;
					
					this.chkBordes.Enabled = true;
					this.chkSupDelantera.Enabled = true;
					this.chkSupPosterior.Enabled = true;
					this.chkPines.Enabled = true;
					this.chkBotones.Enabled = true;
					this.chkSerial.Enabled = true;

					this.chkBordes.Checked = true;
					this.chkSupDelantera.Checked = true;
					this.chkSupPosterior.Checked = true;
					this.chkPines.Checked = true;
					this.chkBotones.Checked = true;
					this.chkSerial.Checked = true;
					this.chkGarantiaExtendida.Enabled = true;

					this.txtEstadoFisico.Enabled =  true;
					this.chkEmpqs.Enabled = true;
					this.chkAccesorios.Enabled = true;
					this.txtDaño.Enabled = true;		
					#endregion Cliente - Estado Fisico - Accesorios

					#region Imagenes
					this.btnImg1.Enabled = true;
					this.btnImg2.Enabled = true;
					this.btnImg3.Enabled = true;
					this.button1.Enabled = true;
					this.button2.Enabled = true;
					this.button3.Enabled = true;
					#endregion Imagenes

					this.txtCIEntrega.Enabled = true;
					this.txtNombreEntrega.Enabled = true;

					#region Botones
					this.btnCliente.Enabled = true;
					this.btnNuevo.Enabled = false;
					this.btnConsultar.Enabled = false;
					this.btnEditar.Enabled = false;
					this.btnImprimir.Enabled = false;
					this.btnGuardar.Enabled = true;
					this.btnAnular.Enabled = false;
					this.btnCancelar.Enabled = true;
					
					#endregion Botones

					this.cmbUsuario.Focus();
				}
				else this.btnCancelar_Click(sender, e);
			}						 
		}

		private void cmbBodega_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.cmbUsuario.Focus();
		}

		private void dtFecha_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Delete) e.Handled = true;
			if (e.KeyCode == Keys.Enter) this.cmbBodega.Focus();
		}

		private void BuscaDatosSolicitud(string sSerial)
		{
			string sSQL = string.Format("Exec DTDatosArticulo '{0}'", sSerial);
			SqlDataReader dr = Funcion.rEscalarSQL(cdsSeteoF, sSQL, true);
			dr.Read();
			if (dr.HasRows)
			{
				this.txtidCliente.Value = dr.GetInt32(0);
				this.txtCodigo.Text = dr.GetString(1);
				this.txtNombre.Text = dr.GetString(2);
				this.txtDireccion.Text = dr.GetString(3);
				this.txtTelefono.Text = dr.GetString(4);
				this.txtCelular.Text = dr.GetString(5);
				this.txtIdSerial.Value = dr.GetInt32(6);
				this.txtidArticulo.Value = dr.GetInt32(7);
				this.txtSerial.Text = sSerial;
				this.txtCodArticulo.Text = dr.GetString(8);
				this.txtArticulo.Text = dr.GetString(9);
				this.txtTiempoDeGarantia.Value = dr.GetInt32(10);
				this.cmbGarantia.Text = dr.GetString(11);
				this.txtIdProveedor.Value = dr.GetInt32(12);
				this.txtProveedor.Text = dr.GetString(13);
				this.txtNumeroDeFactura.Text = dr.GetString(14);
				if (dr.GetValue(15) != System.DBNull.Value) this.cmbFechaFactura.Value = dr.GetDateTime(15);

				this.txtGrupoProveedor.Text = dr.GetString(16);
				this.txtOrigen.Value = dr.GetInt32(17);
				this.txtNumFCVenta.Text = dr.GetString(18);
				if (dr.GetValue(19) != System.DBNull.Value) this.cmbFechaVenta.Value = dr.GetDateTime(19);
				idFVenta = dr.GetInt32(20);
				idFCompra = dr.GetInt32(21);
				nBodega = dr.GetInt32(22);
				this.cmbGarantiaProveedor.Text = dr.GetString(24);
				this.txtTiempGarantiaProveedr.Value = dr.GetInt32(23);
				this.txtVendedor.Text = dr.GetString(26);
				if(dr.GetString(25).Trim().Equals("")) this.txtLocal.Text = this.cmbBodega.ActiveRow.Cells["Codigo"].Value.ToString();
				else this.txtLocal.Text = dr.GetString(25);
				this.txtCorreo.Text = dr.GetString(27);
				iPointCliente = dr.GetInt32(28);
				idTipoRuc = dr.GetInt32(29);
			}
			dr.Close();

			this.txtIngresos.Value = Funcion.iEscalarSQL(cdsSeteoF, string.Format("Select Count(*) From DTIngreso Where idSerial = {0}", (int)this.txtIdSerial.Value)) + 1;
		}

		private bool ValidaArticuloExterno(string sCodigo, int iBusqueda)
		{
			string sSQL = string.Format("Exec DTValidaArticuloIngresoExterno '{0}', {1}", sCodigo, iBusqueda);
			string sMensaje = Funcion.sEscalarSQL(cdsSeteoF, sSQL);

			if (sMensaje.Trim().Length > 0)
			{
				MessageBox.Show(sMensaje, "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				if (iBusqueda == 1) this.txtSerial.Focus();
				if (iBusqueda == 2) this.txtCodArticulo.Focus();
				return false;
			}

			return true;
		}

		private string ConsumirAPI(string urlBase)
		{
			string respuesta = "";

			try
			{
				// Construir la URL con el parámetro
				string urlCompleta = urlBase ;

				// Crear la solicitud
				HttpWebRequest request = (HttpWebRequest)WebRequest.Create(urlCompleta);
				request.Method = "GET";
				request.Timeout = 10000; // 10 segundos (ajustable)

				// Obtener la respuesta
				using (HttpWebResponse response = (HttpWebResponse)request.GetResponse())
				{
					if (response.StatusCode == HttpStatusCode.OK)
					{
						using (StreamReader reader = new StreamReader(response.GetResponseStream()))
						{
							respuesta = reader.ReadToEnd();
						}
					}
					else
					{
						MessageBox.Show("Error en el servicio: " + response.StatusCode.ToString(), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
						return "Error";
					}
				}
			}
			catch (WebException ex)
			{
				MessageBox.Show("Error de conexión al servicio:\n" + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
				return "Error";
			}
			catch (Exception ex)
			{
				MessageBox.Show("Error inesperado:\n" + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
				return "Error";
			}

			return "OK";
		}


		private void btnGuardar_Click(object sender, System.EventArgs e)
		{
			#region Recibe - Tecnico
			if (!Validacion.vbComboVacio(this.cmbUsuario, "Seleccione el nombre de la persona que recibe el articulo")) return;
			
			if (miAcceso.BAsignaTecnicos && (int)this.cmbBodega.Value == 20)
				if (!Validacion.vbComboVacio(this.cmbAsignaTecnico, "Debe de Asignar el tecnico")) return;
			#endregion Recibe - Tecnico	

			#region Cliente
			if (!Validacion.vbIdentificacion(this.txtCodigo, idTipoRuc, cdsSeteoF)) return;
			
			if (!Validacion.vbValidaTelefono(this.txtTelefono, 1, false))
			{
				FuncionesProcedimientos.AbreFacturacionClientes(idTipoRuc, this.txtCodigo.Text.ToString(), false, false);
				return;
			}

			if (!Validacion.vbValidaTelefono(this.txtCelular, 2, false))
			{
				FuncionesProcedimientos.AbreFacturacionClientes(idTipoRuc, this.txtCodigo.Text.ToString(), false, false);
				CargaCliente();
				return;
			}

			if (!Validacion.bEMail(this.txtCorreo.Text.ToString().Trim()))
			{
				FuncionesProcedimientos.AbreFacturacionClientes(idTipoRuc, this.txtCodigo.Text.ToString(), false, false);
				CargaCliente();
				return;
			}

			if (!Validacion.vbValidaTelefono(this.txtFonoAdicional, 1, false)) return;

			if (!Validacion.vbValidaTelefono(this.txtCelAdicional, 2, false)) return;

			if (this.txtCorreo.Text.ToString().Length > 0) if (!Validacion.bEMail(this.txtCorreo.Text.ToString().Trim())) return;			
			#endregion Cliente

			#region Valida Producto Externo
			if (iInternoExterno == 2)	
			{
				if (!Validacion.vbTexto(this.txtSerial, 5, 30, "Serial")) return;
				if (!Validacion.vbTexto(this.txtCodArticulo, 5, 50, "Codigo articulo")) return;
				if (!Validacion.vbTexto(this.txtArticulo, 5, 50, "Descripcion del articulo")) return;
				if (!ValidaArticuloExterno(this.txtSerial.Text.ToString().Trim(), 1)) return;
				if (!ValidaArticuloExterno(this.txtCodArticulo.Text.ToString().Trim(), 2)) return;
			}
			#endregion Valida Producto Externo

			#region Estado Producto
			if (this.chkBordes.Checked&& this.rbBordes.CheckedIndex == -1) 
			{
				MessageBox.Show("Seleccione el estado de los bordes", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
				this.rbBordes.Focus();
				return;
			} 
			if (this.chkPines.Checked && this.rbPines.CheckedIndex == -1) 
			{
				MessageBox.Show("Seleccione el estado de los pines", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
				this.rbPines.Focus();
				return;
			} 
			if (this.chkBotones.Checked && this.rbBotones.CheckedIndex == -1) 
			{
				MessageBox.Show("Seleccione el estado de los botones", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
				this.rbBotones.Focus();
				return;
			}
			if (this.chkSupDelantera.Checked && this.rbSupDelantera.CheckedIndex == -1) 
			{
				MessageBox.Show("Seleccione el estado de la superficie delantera", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
				this.rbSupDelantera.Focus();
				return;
			}
			if (this.chkSupPosterior.Checked && this.rbSupPosterior.CheckedIndex == -1) 
			{
				MessageBox.Show("Seleccione el estado de la superficie posterior", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
				this.rbSupPosterior.Focus();
				return;
			}			 
			if (this.chkSerial.Checked && this.rbSerial.CheckedIndex == -1) 
			{
				MessageBox.Show("Seleccione el estado de la serie", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
				this.rbSerial.Focus();
				return;
			}
			if (!Validacion.vbTexto(this.txtEstadoFisico, 10, 1000, "Estado fisico del articulo")) return;
			#endregion Estado Producto

			#region Empaques
			if (this.chkEmpqs.Checked)
			{
				if (this.rbEmpaques.CheckedIndex == -1) 
				{
					MessageBox.Show("Seleccione el estado de los empaques", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
					this.rbEmpaques.Focus();
					return;
				}
				if (!Validacion.vbTexto(this.txtEmpqsObservaciones, 10, 1000, "Notas de los empaques de caja")) return;								
			}
			#endregion Empaques
			
			#region Accesorios
			if (this.chkAccesorios.Checked)
			{
				if (this.rbAccesorios.CheckedIndex == -1) 
				{
					MessageBox.Show("Seleccione el estado de los accesorios", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
					this.rbAccesorios.Focus();
					return;
				}
				if (!Validacion.vbTexto(this.txtAccesorios, 10, 1000, "Notas de los accesorios")) return;								
			}
			#endregion Accesorios

			#region Daño - Imagenes
			if (!Validacion.vbTexto(this.txtDaño, 10, 1000, "Daño del articulo")) return;
			
			if (!Validacion.vbTexto(this.txtImagen1, 0, 1000, "Imagen 1")) return;
			if (!Validacion.vbTexto(this.txtImagen2, 0, 1000, "Imagen 2")) return;
			if (!Validacion.vbTexto(this.txtImagen3, 0, 1000, "Imagen 3")) return;
			#endregion Daño - Imagenes			
			
			#region Variables
			DateTime dtFeVenta = DateTime.Today;
			DateTime dtFeCpa = DateTime.Today;
			
			if(this.cmbFechaVenta.Value != System.DBNull.Value) dtFeVenta = (DateTime)this.cmbFechaVenta.Value;
			if(this.cmbFechaFactura.Value != System.DBNull.Value) dtFeCpa = (DateTime)this.cmbFechaFactura.Value;
						
			int tieneAccesorios = 0;
			int tieneEmpaqueCaja = 0;			
			int edoFisicoBordes = 0;
			int edofisicoSuperficieDelantera = 0;
			int edoFiiscoSuperficiePosterior = 0;
			int edoFisicoPines = 0;
			int edoFisicoBotones = 0;
			int edoFisicoSerieProducto = 0;
			int edoAccesorios = 0;
			int edoEmpaqueCaja = 0;	
			int idTecnico = 0;
			#endregion Variables

			#region Asigna Variables Estado Fisico - Empaques - Accesorios
			if (this.rbBordes.CheckedIndex > - 1) edoFisicoBordes = (int)this.rbBordes.Value;

			if(this.rbSupDelantera.CheckedIndex > - 1) edofisicoSuperficieDelantera = (int)this.rbSupDelantera.Value;

			if(this.rbSupPosterior.CheckedIndex > - 1) edoFiiscoSuperficiePosterior = (int)this.rbSupPosterior.Value;

			if(this.rbPines.CheckedIndex > - 1) edoFisicoPines = (int)this.rbPines.Value;

			if(this.rbBotones.CheckedIndex > - 1) edoFisicoBotones = (int)this.rbBotones.Value;

			if(this.rbSerial.CheckedIndex > - 1) edoFisicoSerieProducto = (int)this.rbSerial.Value;			
			
			if (this.rbAccesorios.CheckedIndex > - 1) edoAccesorios = (int)this.rbAccesorios.Value;			

			if (this.rbEmpaques.CheckedIndex > - 1) edoEmpaqueCaja = (int)this.rbEmpaques.Value;		
						
			if(this.cmbAsignaTecnico.ActiveRow != null) idTecnico = (int)this.cmbAsignaTecnico.Value;	

			if (idTecnico == 0)  this.txtEstado.Value = 1; else this.txtEstado.Value = 3;
			#endregion Asigna Variables Estado Fisico - Empaques - Accesorios

			
			if (bNuevo)
			{
				if (DialogResult.Yes == MessageBox.Show("¿Está seguro de guardar? No podrá volver a subir imágenes.", "Point Technology", MessageBoxButtons.YesNo, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2))
				{
					#region Api 
					string url = "http://192.168.1.95:3055/api/v1/point/DepartamentoTecnico/"
						+ nuevoNombreArchivoImgone + "/"
						+ nuevoNombreArchivoImgTwo + "/"
						+ nuevoNombreArchivoImgTree;

					//string url = "http://192.168.1.95:3055/api/v1/point/DepartamentoTecnico/archivo_20251001_160636_508.png/archivo_20251001_160636_508.png/archivo_20251001_160636_508.png";  // La URL base
					string jsonRespuesta = ConsumirAPI(url);

					if (jsonRespuesta == "Error")
					{
						MessageBox.Show("Error en la subida de Imagenes \n\n  Comunicate con el área de Desarrollo.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
						return;
					}

					// Luego puedes parsear el JSON como te mostré antes.
					//MessageBox.Show("Respuesta:\n" + jsonRespuesta, "API", MessageBoxButtons.OK, MessageBoxIcon.Information);
        
					#endregion Api
			
					string baseUrl = "https://storage.googleapis.com/sparta_bucket/DEPARTAMENTOTECNICO/POINT/";
					this.txtImagen1.Text = baseUrl + nuevoNombreArchivoImgone;
					this.txtImagen2.Text = baseUrl + nuevoNombreArchivoImgTwo;
					this.txtImagen3.Text = baseUrl + nuevoNombreArchivoImgTree;
				}
				else
				{
					return;
				
				}
			}

			try
			{
				#region Graba Ingreso
				if (bNuevo) this.lblIngreso.Text = Funcion.sEscalarSQL(cdsSeteoF, string.Format("Exec NumeracionLocales 50, {0}, 0", (int)this.cmbBodega.Value));
		
				string sSQL = string.Format("Exec DTGuardaIngreso {0}, {1}, {2}, {3}, '{4}', '{5}', '{6}', {7}, {8}, '{9}', '{10}', {11}, {12}, '{13}', {14}, {15}, {16}, {17}, '{18}', '{19}', '{20}', {21}, {22}, '{23}', {24}, {25}, '{26}', {27}, {28}, {29}, {30}, {31}, {32}, {33}, {34}, {35}, {36}, {37}, {38}, '{39}', {40}, '{41}', '{42}', '{43}', {44}, '{45}', '{46}', '{47}', '{48}', '{49}', '{50}', {51}, '{52}', '{53}', {54}, {55}",
					(int)this.txtIdIngreso.Value, (int)this.txtOrigen.Value, iInternoExterno, (int)this.cmbBodega.Value, Convert.ToDateTime(this.dtFecha.Value).ToString("yyyyMMdd"), this.lblIngreso.Text.ToString(), this.lblEgreso.Text.ToString(), 
					idFVenta, (int)this.txtidCliente.Value, dtFeVenta.ToString("yyyyMMdd"),this.txtNumFCVenta.Text.ToString(),(int)this.txtidArticulo.Value,
					(int)this.txtIdSerial.Value, this.cmbGarantia.Text.ToString(), (int)this.txtTiempoDeGarantia.Value, (int)this.txtIngresos.Value,
					idFCompra, (int)this.txtIdProveedor.Value, dtFeCpa.ToString("yyyyMMdd"), this.txtNumeroDeFactura.Text.ToString(), this.txtDaño.Value.ToString(),
					this.chkAccesorios.Checked, edoAccesorios, this.txtAccesorios.Text.ToString().Trim(), 
					this.chkEmpqs.Checked, edoEmpaqueCaja, this.txtEmpqsObservaciones.Text.ToString().Trim(),				
					this.chkBordes.Checked, edoFisicoBordes, this.chkSupDelantera.Checked, edofisicoSuperficieDelantera, this.chkSupPosterior.Checked, edoFiiscoSuperficiePosterior,
					this.chkPines.Checked, edoFisicoPines, this.chkBotones.Checked, edoFisicoBotones,
					this.chkSerial.Checked, edoFisicoSerieProducto, this.txtEstadoFisico.Text.ToString().Trim(), 
					(int)this.txtEstado.Value, this.txtImagen1.Text.ToString(), this.txtImagen2.Text.ToString(), this.txtImagen3.Text.ToString(), 
					(int)this.cmbUsuario.Value,	this.txtCIEntrega.Text.ToString(), this.txtNombreEntrega.Text.ToString(), this.txtFonoAdicional.Text.ToString(), 
					this.txtSerial.Text.ToString(),	this.txtCodArticulo.Text.ToString(), this.txtArticulo.Text.ToString(), 
					idTecnico, this.txtCorreo.Text.Trim(), this.txtCelAdicional.Text.Trim(), iPointCliente, this.chkGarantiaExtendida.Checked);
				this.txtIdIngreso.Value = Funcion.iEscalarSQL(cdsSeteoF, sSQL);					
				#endregion Graba Ingreso 

				#region Controles
				#region General
				this.dtFecha.Enabled = false;
				this.cmbBodega.Enabled = false;
				this.cmbUsuario.Enabled = false;
				this.cmbAsignaTecnico.Enabled = false;
				#endregion General

				#region Cliente
				this.txtCodigo.Enabled = false;
				this.txtFonoAdicional.Enabled = false;
				this.txtCelAdicional.Enabled = false;
				this.txtCorreo.Enabled = false;			
				#endregion Cliente

				#region Articulo
				this.txtSerial.Enabled = false;
				this.txtCodArticulo.Enabled = false;
				this.txtArticulo.Enabled = false;
				this.cmbGarantia.Enabled = false;
				this.chkGarantiaExtendida.Enabled = false;
				#endregion Articulo	

				#region Estado Fisico
				this.chkBordes.Enabled = false;			
				this.chkPines.Enabled = false;
				this.chkBotones.Enabled = false;
				this.chkSupDelantera.Enabled = false;
				this.chkSupPosterior.Enabled = false;			
				this.chkSerial.Enabled = false;

				this.rbBordes.Enabled = false;
				this.rbPines.Enabled = false;
				this.rbBotones.Enabled = false;
				this.rbSupDelantera.Enabled = false;
				this.rbSupPosterior.Enabled = false;			
				this.rbSerial.Enabled = false;
			
				this.txtEstadoFisico.Enabled = false;
				#endregion Estado Fisico

				#region Empaques - Accesorios
				this.chkEmpqs.Enabled = false;
				this.rbEmpaques.Enabled = false;
				this.txtEmpqsObservaciones.Enabled = false;

				this.chkAccesorios.Enabled = false;
				this.rbAccesorios.Enabled = false;
				this.txtAccesorios.Enabled = false;
				#endregion Empaques - Accesorios

				this.txtDaño.Enabled = false;

				#region Imagenes
				this.btnImg1.Enabled = false;
				this.btnImg2.Enabled = false;
				this.btnImg3.Enabled = false;

				this.button1.Enabled = false;
				this.button2.Enabled = false;
				this.button3.Enabled = false;
				#endregion Imagenes
						
				#region Entrega
				this.txtCIEntrega.Enabled = false;
				this.txtNombreEntrega.Enabled = false;
				#endregion Entrega
				
				#region Botones
				this.btnCliente.Enabled = false;
				if (miAcceso.Nuevo) this.btnNuevo.Enabled = true;
				if (miAcceso.Buscar) this.btnConsultar.Enabled = true;
				if (miAcceso.Editar) this.btnEditar.Enabled = true;
				if (miAcceso.Imprimir) this.btnImprimir.Enabled = true;
				this.btnGuardar.Enabled = false;
				if (miAcceso.Anular) this.btnAnular.Enabled = true;
				this.btnCancelar.Enabled = true;
				#endregion Botones

				bNuevo = false;
				bEdicion = false;
				nuevoNombreArchivoImgone ="";
				nuevoNombreArchivoImgTwo ="";
				nuevoNombreArchivoImgTree ="";
				#endregion Controles

				#region Guarda Cliente - Serial - Articulo
				//				if ((int)this.rbgInternoExterno.Value == 2)
				//				{
				//					string sSQLCSA = string.Format("Exec DTGuardaClienteArticuloExterno {0}, '{1}', '{2}', '{3}', '{4}', '{5}', {6}, '{7}', {8}, '{9}', '{10}'", 
				//						(int)this.txtidCliente.Value, this.txtCodigo.Text.ToString(), this.txtNombre.Text.ToString(), this.txtDireccion.Text.ToString(), this.txtTelefono.Text.ToString(), this.txtCelular.Text.ToString(),
				//						(int)this.txtIdSerial.Value, this.txtSerial.Text.ToString(), (int)this.txtidArticulo.Value, this.txtCodArticulo.Text.ToString(), this.txtArticulo.Text.ToString());
				//					SqlDataReader dr = Funcion.rEscalarSQL(cdsSeteoF, sSQLCSA, true);
				//					dr.Read();
				//
				//					if (dr.HasRows)
				//					{
				//						this.txtidCliente.Value = dr.GetInt32(0);
				//						this.txtidArticulo.Value = dr.GetInt32(1);
				//						this.txtIdSerial.Value = dr.GetInt32(2);
				//					}
				//					dr.Close();
				//				}
				#endregion Guarda Cliente - Serial - Articulo
			}
			catch (Exception ex)
			{
				MessageBox.Show(string.Format("Exception Type: {0} {1}", ex.GetType(), ex.Message), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);							
			}
		}

		private void dtFecha_BeforeDropDown(object sender, System.ComponentModel.CancelEventArgs e)
		{
		
		}

		private void txtDaño_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.txtCIEntrega.Focus();
		}
		
		private void txtObservaciones_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.btnGuardar.Focus();
		}

		private void txtCodigo_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.txtFonoAdicional.Focus();
		}

		private void txtCodigo_Validating(object sender, System.ComponentModel.CancelEventArgs e)
		{
			if (!bEdicion) return;

			if (!Validacion.vbIdentificacion(this.txtCodigo, idTipoRuc, cdsSeteoF)) e.Cancel = true;
		}

		private void txtCodigo_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			Funcion.CamposNumericos(sender, e);
		}

		private void CargaCliente()
		{
			string[] sParametros = new string[13];

			sParametros = FuncionesProcedimientos.RetornaParametrosClientesProveedores(false, 1, this.txtCodigo.Text.ToString(), cdsSeteoF);
			
			if (sParametros[0] != null) this.txtidCliente.Value = int.Parse(sParametros[0]);			
			if (sParametros[3] != null) this.txtCodigo.Text = sParametros[3];
			if (sParametros[4] != null) this.txtNombre.Text = sParametros[4];	
			if (sParametros[7] != null) this.txtDireccion.Text = sParametros[7];
			if (sParametros[8] != null) this.txtTelefono.Text = sParametros[8];
			if (sParametros[9] != null) this.txtCelular.Text = sParametros[9];
			if (sParametros[12] != null) this.txtCorreo.Text = sParametros[12];
		}

		private void txtCodigo_Validated(object sender, System.EventArgs e)
		{
			if (!bEdicion) return;

			this.txtidCliente.Value = 0;
			this.txtNombre.Text = "";
			this.txtDireccion.Text = "";
			this.txtTelefono.Text = "";
			this.txtCelular.Text = "";
			this.txtFonoAdicional.Text = "";
			this.txtCelAdicional.Text = "";
			this.txtCorreo.Text = "";

			#region Busca Cliente
			CargaCliente();

			if ((int)this.txtidCliente.Value == 0)
			{
				FuncionesProcedimientos.AbreFacturacionClientes(idTipoRuc, this.txtCodigo.Text.ToString(), true, false);

				CargaCliente();
			}
			#endregion Busca Cliente			
		}

		private void txtCelular_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			Funcion.CamposNumericos(sender, e);
		}

		private void txtTelefono_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			Funcion.CamposNumericos(sender, e);
		}

		private void txtNombre_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.txtDireccion.Focus();
		}

		private void txtDireccion_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.txtTelefono.Focus();
		}

		private void txtTelefono_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.txtCelular.Focus();
		}

		private void txtCelular_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.txtFonoAdicional.Focus();
		}

		private void txtCodArticulo_Validated(object sender, System.EventArgs e)
		{
			if (!bEdicion) return;

			if (Funcion.iEscalarSQL(cdsSeteoF, string.Format("Select COUNT(*) From Articulo Where idGrupoArticulo = 28 And Codigo = '{0}'", this.txtCodigo.Text.ToString())) > 0)
			{
				SqlDataReader drArticulo = Funcion.rEscalarSQL(cdsSeteoF, string.Format("Select a.idArticulo, ISNULL(s.idSerial, 0), ISNULL(s.Serial, ''), a.Codigo, a.Articulo From Articulo a Left Outer Join Seriales s On s.idArticulo = a.idArticulo Where idGrupoArticulo = 28 And Codigo = '{0}'", this.txtCodigo.Text.ToString()), true);
				drArticulo.Read();
				if (drArticulo.HasRows)
				{
					this.txtidArticulo.Value = drArticulo.GetInt32(0);
					this.txtIdSerial.Value = drArticulo.GetInt32(1);
					this.txtSerial.Text = drArticulo.GetString(2);
					this.txtCodigo.Text = drArticulo.GetString(3);
					this.txtArticulo.Text = drArticulo.GetString(4);
				}
				drArticulo.Close();

				this.txtIngresos.Value = Funcion.iEscalarSQL(cdsSeteoF, string.Format("Select COUNT(*) From DTIngreso Where idCliente = {0} And idSerial = {1} And Estado > 0", (int)this.txtidCliente.Value, (int)this.txtIdSerial.Value));
			}
		}

		private void txtCodArticulo_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.txtArticulo.Focus();
		}

		private void txtArticulo_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.txtEstadoFisico.Focus();
		}

		private void txtCelular_ValueChanged(object sender, System.EventArgs e)
		{
		
		}

		private void txtCodArticulo_ValueChanged(object sender, System.EventArgs e)
		{
		
		}

		private void txtSerial_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.txtCodArticulo.Focus();
		}

		private void txtSerial_Validated(object sender, System.EventArgs e)
		{
			if (!bEdicion) return;

			if (Funcion.iEscalarSQL(cdsSeteoF, string.Format("Select COUNT(*) From Seriales Where Estado = 20 And Serial = '{0}'", this.txtSerial.Text.ToString())) > 0)
			{
				SqlDataReader drArticulo = Funcion.rEscalarSQL(cdsSeteoF, string.Format("Select a.idArticulo, s.idSerial, a.Codigo, a.Articulo From Articulo a Inner Join Seriales s On s.idArticulo = a.idArticulo Where idGrupoArticulo = 28 And s.Serial = '{0}'", this.txtCodigo.Text.ToString()), true);
				drArticulo.Read();
				if (drArticulo.HasRows)
				{
					this.txtidArticulo.Value = drArticulo.GetInt32(0);
					this.txtIdSerial.Value = drArticulo.GetInt32(1);
					this.txtCodigo.Text = drArticulo.GetString(2);
					this.txtArticulo.Text = drArticulo.GetString(3);
				}
				drArticulo.Close();

				this.txtIngresos.Value = Funcion.iEscalarSQL(cdsSeteoF, string.Format("Select COUNT(*) From DTIngreso Where idCliente = {0} And idSerial = {1} And Estado > 0", (int)this.txtidCliente.Value, (int)this.txtIdSerial.Value));
			}
		}

		private void btnConsultar_Click(object sender, System.EventArgs e)
		{
			this.btnCancelar_Click(sender, e);

			using (frmDTConsultaIngreso Busqueda = new frmDTConsultaIngreso())
			{ 
				if (DialogResult.OK == Busqueda.ShowDialog())
				{
					this.cmbGarantia.Items.Add("SI");
					this.cmbGarantia.Items.Add("NO");
					this.cmbGarantia.Items.Add("POR VERIFICAR");

					this.txtIdIngreso.Value = (int)Busqueda.uGridDTIngreso.ActiveRow.Cells["idIngreso"].Value;					
					this.txtOrigen.Value = (int)Busqueda.uGridDTIngreso.ActiveRow.Cells["Origen"].Value; 
					iInternoExterno = (int)Busqueda.uGridDTIngreso.ActiveRow.Cells["InternoExterno"].Value;
					this.cmbBodega.Value = (int)Busqueda.uGridDTIngreso.ActiveRow.Cells["Bodega"].Value; 
					this.dtFecha.Value = (DateTime)Busqueda.uGridDTIngreso.ActiveRow.Cells["Fecha"].Value;	 
					this.lblIngreso.Text = Busqueda.uGridDTIngreso.ActiveRow.Cells["Ingreso"].Value.ToString();
					this.lblEgreso.Text = Busqueda.uGridDTIngreso.ActiveRow.Cells["Egreso"].Value.ToString();	
					this.txtidCliente.Value = (int)Busqueda.uGridDTIngreso.ActiveRow.Cells["idCliente"].Value;
					this.txtCodigo.Text = Busqueda.uGridDTIngreso.ActiveRow.Cells["Ruc"].Value.ToString();	 
					this.txtNombre.Text = Busqueda.uGridDTIngreso.ActiveRow.Cells["Nombre"].Value.ToString();	 
					this.txtDireccion.Text = Busqueda.uGridDTIngreso.ActiveRow.Cells["Direccion"].Value.ToString();
					this.txtTelefono.Text = Busqueda.uGridDTIngreso.ActiveRow.Cells["Telefono"].Value.ToString();	 
					this.txtCelular.Text = Busqueda.uGridDTIngreso.ActiveRow.Cells["Celular"].Value.ToString();	 
					this.txtIdSerial.Value = (int)Busqueda.uGridDTIngreso.ActiveRow.Cells["idSerial"].Value;
					this.txtSerial.Text = Busqueda.uGridDTIngreso.ActiveRow.Cells["Serial"].Value.ToString();
					this.txtidArticulo.Value = (int)Busqueda.uGridDTIngreso.ActiveRow.Cells["idArticulo"].Value;
					this.txtCodArticulo.Text = Busqueda.uGridDTIngreso.ActiveRow.Cells["Codigo"].Value.ToString();
					this.txtArticulo.Text = Busqueda.uGridDTIngreso.ActiveRow.Cells["Articulo"].Value.ToString();					
					this.cmbGarantia.Text = Busqueda.uGridDTIngreso.ActiveRow.Cells["Garantia"].Value.ToString();
					this.txtTiempoDeGarantia.Value = (int)Busqueda.uGridDTIngreso.ActiveRow.Cells["TiempoGarantia"].Value;
					this.txtIngresos.Value = (int)Busqueda.uGridDTIngreso.ActiveRow.Cells["Ingresos"].Value;
					this.txtIdProveedor.Value = (int)Busqueda.uGridDTIngreso.ActiveRow.Cells["idProveedor"].Value;
					this.txtEstado.Value = (int)Busqueda.uGridDTIngreso.ActiveRow.Cells["Estado"].Value;

//					this.cmbUsuario.DataSource = Funcion.dvProcedimiento(cdsSeteoF, string.Format("Exec ListaUsuarios {0}", (int)this.cmbBodega.Value));
//					this.cmbAsignaTecnico.DataSource = Funcion.dvProcedimiento(cdsSeteoF, string.Format("Exec ListaUsuarios {0}", (int)this.cmbBodega.Value));

					#region Proveedor
					if (iInternoExterno == 1)
					{
						string sSQL = string.Format("Exec DTBuscaProvPorSerial '{0}', {1}", this.txtSerial.Text.ToString(), (int)this.txtOrigen.Value);
						SqlDataReader drProv = Funcion.rEscalarSQL(cdsSeteoF, sSQL, true);
						drProv.Read();
						if (drProv.HasRows)
						{
							if (drProv.GetValue(0) != System.DBNull.Value)this.txtProveedor.Value = drProv.GetString(0);
							if (drProv.GetValue(1) != System.DBNull.Value)this.txtNumeroDeFactura.Value = drProv.GetString(1);
							if (drProv.GetValue(2) != System.DBNull.Value)this.cmbFechaFactura.Value = drProv.GetDateTime(2);
							if (drProv.GetValue(3) != System.DBNull.Value) this.txtGrupoProveedor.Text = drProv.GetString(3);
						}
						drProv.Close();
					}
					#endregion Proveedor

					this.txtDaño.Text = Busqueda.uGridDTIngreso.ActiveRow.Cells["Daño"].Value.ToString();
					//this.txtAccesorios.Text = Busqueda.uGridDTIngreso.ActiveRow.Cells["Accesorios"].Value.ToString();
					//this.txtObservaciones.Text = Busqueda.uGridDTIngreso.ActiveRow.Cells["Observaciones"].Value.ToString();

					/* Buscar resto de opciones de pantalla */
					BDatosSolicitud((int)this.txtIdIngreso.Value);
					          
					if (miAcceso.Editar) this.btnEditar.Enabled = true;
					if (miAcceso.Imprimir) this.btnImprimir.Enabled = true;
					if (miAcceso.Anular) this.btnAnular.Enabled = true;
					this.btnCancelar.Enabled = true;									
				}
			}
		}

		private void BDatosSolicitud(int idIngreso)
		{
			string sSQL = string.Format("Exec DTBuskIngreso {0}", idIngreso);
			SqlDataReader dr = Funcion.rEscalarSQL(cdsSeteoF, sSQL, true);
			dr.Read();
			if (dr.HasRows)
			{
				idFVenta = dr.GetInt32(43);
				idFCompra = dr.GetInt32(42);
				nBodega = dr.GetInt32(3);
				this.txtEstado.Value = dr.GetInt32(59);
				this.chkEmpqs.Checked = dr.GetBoolean(29);
				this.rbEmpaques.Value = dr.GetInt32(28);
				this.txtEmpqsObservaciones.Text = dr.GetString(34).Trim();

				this.chkBordes.Checked = dr.GetBoolean(44);
				this.rbBordes.Value = dr.GetInt32(35);

				this.chkSupDelantera.Checked = dr.GetBoolean(45);
				this.rbSupDelantera.Value = dr.GetInt32(39);

				this.chkSupPosterior.Checked = dr.GetBoolean(46);
				this.rbSupPosterior.Value = dr.GetInt32(40);

				this.chkPines.Checked = dr.GetBoolean(47);
				this.rbPines.Value = dr.GetInt32(37);

				this.chkBotones.Checked = dr.GetBoolean(48);
				this.rbBotones.Value = dr.GetInt32(36);

				this.chkSerial.Checked = dr.GetBoolean(49);
				this.rbSerial.Value = dr.GetInt32(38);
				this.txtEstadoFisico.Text = dr.GetString(41).Trim();

				this.chkAccesorios.Checked = dr.GetBoolean(24);
				this.rbAccesorios.Value = dr.GetInt32(27);
				this.txtAccesorios.Text = dr.GetString(25).Trim();
				
				this.txtImagen1.Text = dr.GetString(50).Trim();
				this.txtImagen2.Text = dr.GetString(51).Trim();
				this.txtImagen3.Text = dr.GetString(52).Trim();
				this.lblEstado.Text = dr.GetString(54).Trim();

				this.txtNumFCVenta.Text = dr.GetString(33).Trim();
				this.txtLocal.Text = dr.GetString(55).Trim();
				this.cmbFechaVenta.Value = dr.GetDateTime(31);
				this.txtVendedor.Text = dr.GetString(56).Trim();
				this.cmbGarantiaProveedor.Text = dr.GetString(57).Trim();
				this.cmbUsuario.Value = dr.GetInt32(58);
				this.txtCIEntrega.Text = dr.GetString(60).Trim();
				this.txtNombreEntrega.Text = dr.GetString(61).Trim();
				this.txtFonoAdicional.Text = dr.GetString(62).Trim();

				this.txtCorreo.Text = dr.GetString(71).Trim();
				this.txtCelAdicional.Text = dr.GetString(72).Trim();
				this.cmbAsignaTecnico.Value = dr.GetInt32(73);
				this.chkGarantiaExtendida.Checked = dr.GetBoolean(74);
			}
			dr.Close();
		}

		private void btnEditar_Click(object sender, System.EventArgs e)
		{	
			this.cmbUsuario.Enabled = true;

			if (miAcceso.BAsignaTecnicos) this.cmbAsignaTecnico.Enabled = true;

			if (miAcceso.BEditarNombre && iInternoExterno == 2) this.txtCodigo.Enabled = true;

			#region Cliente

			this.txtFonoAdicional.Enabled = true;
			this.txtCelAdicional.Enabled = true;
			#endregion Cliente

			#region Articulo
			if (iInternoExterno == 2)
			{	
				this.txtSerial.Enabled = true;
				this.txtCodArticulo.Enabled = true;
				this.txtArticulo.Enabled = true;				
			}
			#endregion Articulo			

			#region Estado Fisico
			this.chkGarantiaExtendida.Enabled = true;

			this.chkBordes.Enabled = true;
			if (this.chkBordes.Checked) this.rbBordes.Enabled = true;

			this.chkPines.Enabled = true;
			if (this.chkPines.Checked) this.rbPines.Enabled = true;

			this.chkBotones.Enabled = true;
			if (this.chkBotones.Checked) this.rbBotones.Enabled = true;

			this.chkSupDelantera.Enabled = true;
			if (this.chkSupDelantera.Checked) this.rbSupDelantera.Enabled = true;

			this.chkSupPosterior.Enabled = true;
			if (this.chkSupPosterior.Checked) this.rbSupPosterior.Enabled = true;
			
			this.chkSerial.Enabled = true;
			if(this.chkSerial.Checked) this.rbSerial.Enabled = true;
					
			this.txtEstadoFisico.Enabled = true;
			#endregion Estado Fisico
			
			#region Empaques - Accesorios
			this.chkEmpqs.Enabled = true;
			if(this.chkEmpqs.Checked)
			{
				this.rbEmpaques.Enabled = true;
				this.txtEmpqsObservaciones.Enabled = true;
			}
			
			this.chkAccesorios.Enabled = true;
			if(this.chkAccesorios.Checked)
			{
				this.rbAccesorios.Enabled = true;
				this.txtAccesorios.Enabled = true;
			}			
			#endregion Empaques - Accesorios

			this.txtDaño.Enabled = true;

			#region Imagenes - Entrega
			this.btnImg1.Enabled = false;
			this.btnImg2.Enabled = false;
			this.btnImg3.Enabled = false;
			this.button1.Enabled = true;
			this.button2.Enabled = true;
			this.button3.Enabled = true;

			this.txtCIEntrega.Enabled = true;
			this.txtNombreEntrega.Enabled = true;
			#endregion Imagenes - Entrega
					
			#region Botones
			this.btnCliente.Enabled = false;
			this.btnNuevo.Enabled = false;
			this.btnConsultar.Enabled = false;
			this.btnEditar.Enabled = false;
			this.btnImprimir.Enabled = false;
			this.btnGuardar.Enabled = true;
			this.btnAnular.Enabled = false;
			#endregion Botones

			bNuevo = false;
			bEdicion = true;

			if (miAcceso.BEditarNombre || iInternoExterno == 2) this.txtCodigo.Focus();
			else this.txtFonoAdicional.Focus();
		}

		private void btnAnular_Click(object sender, System.EventArgs e)
		{			
			if (!miAcceso.Anular)
			{
				MessageBox.Show("No tiene acceso a Anular " + Text.Trim(), "Aviso...!!!", MessageBoxButtons.OK, MessageBoxIcon.Stop);				
			}	
			else if((int)this.txtEstado.Value != 8 && (int)this.txtEstado.Value != 12)
			{
				MessageBox.Show("No se puede Anular Ingreso, el producto ya se encuentra en proceso...!!! ", "Aviso...!!!", MessageBoxButtons.OK, MessageBoxIcon.Stop);				
			}
			else if (DialogResult.Yes == MessageBox.Show("Esta seguro de anular este documento.", "Point Technoligy", MessageBoxButtons.YesNo, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2)) 
			{
				using (frmJustAnulaVenta miNota = new frmJustAnulaVenta())
				{
					if (DialogResult.OK == miNota.ShowDialog())
					{
						try
						{
							string sSQLAnula = string.Format("Update DTIngreso Set EstadoIngreso = 0, Observaciones = Observaciones + CHAR(13) + ' MOTIVO DE ANULACION : ' + '{1}' Where idIngreso = {0}", 
								(int)this.txtIdIngreso.Value, miNota.txtNotas.Text.ToString());
							Funcion.sEscalarSQL(cdsSeteoF, sSQLAnula);

							this.lblEstado.Text = "ANULADO";
							this.txtEstado.Value = 0;					

							this.btnAnular.Enabled = false;

							bNuevo = false;
							bEdicion = false;
						}
						catch(Exception Exc)
						{
							MessageBox.Show(string.Format("Error al anular: {0}", Exc), "Aviso...!!!", MessageBoxButtons.OK, MessageBoxIcon.Error);
						}
					}
				}
			}
		}

		private void Imprimir(bool bEntradaSalida, string stReporte)
		{/*
			Cursor = Cursors.WaitCursor;
			//string stFiltro = "{DTIngreso.idIngreso} = " + this.txtIdIngreso.Value.ToString();
			//string stFiltro = "{DTIngreso.idIngreso} = " + idIngresoDT.ToString().Trim();
			
			#region Imprime el reporte
			//using (Reporte miReporte = new Reporte(stReporte, ""))
			//{		
			if (iTipo == 0) 
			{ // Muestra en pantalla

				ParameterFields paramFields = new ParameterFields ();
				
				ParameterField pfRet = new ParameterField ();
				ParameterDiscreteValue idIngreso = new ParameterDiscreteValue ();
				pfRet.ParameterFieldName = "@idIngrs";
				idIngreso.Value = (int)this.txtIdIngreso.Value;
				pfRet.CurrentValues.Add (idIngreso);
				paramFields.Add (pfRet);
				
				Reporte miRepor = new Reporte("DTInformeIngreso.rpt", "");
				miRepor.MdiParent = this.MdiParent;
				miRepor.Titulo("Ingreso DT");
				miRepor.crVista.ParameterFieldInfo = paramFields;
				//miReporte.ShowDialog();
				miRepor.Show();
			}
			if (iTipo == 1) // Imprime una o mas copias
			{
				ParameterFields paramFields = new ParameterFields ();
				
				ParameterField pfRet = new ParameterField ();
				ParameterDiscreteValue idIngreso = new ParameterDiscreteValue ();
				pfRet.ParameterFieldName = "@idIngrs";
				idIngreso.Value = (int)this.txtIdIngreso.Value;
				pfRet.CurrentValues.Add (idIngreso);
				paramFields.Add (pfRet);
				
				Reporte miRepor = new Reporte("SalidaProductoDt.rpt", "");
				miRepor.MdiParent = this.MdiParent;
				miRepor.Titulo("Ingreso DT");
				miRepor.crVista.ParameterFieldInfo = paramFields;
				//miReporte.ShowDialog();
				miRepor.Show();
			}
			#endregion Imprime el reporte
			Cursor = Cursors.Default;
*/
			#region Imprime el reporte
			Cursor = Cursors.WaitCursor;					
						
			ParameterFields paramFields = new ParameterFields ();
				
			ParameterField pfRet = new ParameterField ();
			ParameterDiscreteValue pidIngreso = new ParameterDiscreteValue ();
			if (!bEntradaSalida) pfRet.ParameterFieldName = "@idIngrs";
			if (bEntradaSalida) pfRet.ParameterFieldName = "@idIngreso";
			pidIngreso.Value = (int)this.txtIdIngreso.Value;
			pfRet.CurrentValues.Add (pidIngreso);
			paramFields.Add (pfRet);
				
			Reporte miRepor = new Reporte(stReporte, "");
			miRepor.MdiParent = this.MdiParent;
			miRepor.Titulo("Ingreso DT");
			miRepor.crVista.ParameterFieldInfo = paramFields;			
			miRepor.Show();	
			
			Cursor = Cursors.Default;
			#endregion Imprime el reporte
		}

		private void btnImprimir_Click(object sender, System.EventArgs e)
		{
			this.mnuImprimir.Show(this.btnImprimir, new Point(1, 25));			
		}

		private void mnuImprimirIngreso_Click(object sender, System.EventArgs e)
		{
			Imprimir(false, "DTInformeIngreso.rpt");
		}

		private void mnuImprimirEgreso_Click(object sender, System.EventArgs e)
		{
			Imprimir(true, "SalidaProductoDt.rpt");
		}

		private void chkEmpqs_CheckedChanged(object sender, System.EventArgs e)
		{
			if (!bEdicion) return;

			if(this.chkEmpqs.Checked)
			{
				this.rbEmpaques.Enabled = true;
				this.txtEmpqsObservaciones.Enabled = true;
				this.rbEmpaques.Value = 3;
				this.txtEmpqsObservaciones.Focus();
			}
			else
			{
				this.rbEmpaques.CheckedIndex = -1;
				this.txtEmpqsObservaciones.Text = "";
				this.rbEmpaques.Enabled = false;
				this.txtEmpqsObservaciones.Enabled = false;
			}
		}
		
		public void VisualizaImagen(string ArchiImagn)
		{
			//string curFile = @"c:\temp\test.txt";
			if(File.Exists(ArchiImagn))
			{
				//using (frmVisualizaImagen miImgn = new frmVisualizaImagen(@"C:\Documents and Settings\Infoelec\Mis documentos\Mis imágenes\hd-computer.jpg"))
				using (frmVisualizaImagen miImgn = new frmVisualizaImagen(ArchiImagn))
				{
					if (DialogResult.OK == miImgn.ShowDialog())
					{
						//BuscaDatosSolicitud(miBusqueda.sSerial);
						//this.txtSerial.Text = miImgn.sSerial;
					}
				}
			}
			else
			{ MessageBox.Show("Archivo no existe.?\nVerifique archivo de imagen....","Aviso...!!!",MessageBoxButtons.OK,MessageBoxIcon.Exclamation); }
		}

		public static void SubirArchivoConCampos(
			string url,
			string filePath,
			string cedula,
			string nombreArchivo,
			string tipo
			)
		{
			try
			{
				string boundary = "---------------------------" + DateTime.Now.Ticks.ToString("x");
				byte[] boundaryBytes = Encoding.ASCII.GetBytes("\r\n--" + boundary + "\r\n");
				byte[] trailer = Encoding.ASCII.GetBytes("\r\n--" + boundary + "--\r\n");

				FileStream fs = new FileStream(filePath, FileMode.Open, FileAccess.Read);
				byte[] fileBytes = new byte[fs.Length];
				fs.Read(fileBytes, 0, fileBytes.Length);
				fs.Close();
				string fileName = Path.GetFileName(filePath);

				HttpWebRequest request = (HttpWebRequest)WebRequest.Create(url);
				request.Method = "PUT";
				request.ContentType = "multipart/form-data; boundary=" + boundary;
				request.KeepAlive = true;

				using (Stream requestStream = request.GetRequestStream())
				{
					// --- Agregar campo cedula ---
					string formItemCedula = "--" + boundary + "\r\n" +
						"Content-Disposition: form-data; name=\"cedula\"\r\n\r\n" +
						cedula + "\r\n";
					byte[] cedulaBytes = Encoding.UTF8.GetBytes(formItemCedula);
					requestStream.Write(cedulaBytes, 0, cedulaBytes.Length);

					// --- Agregar campo nombre_del_archivo ---
					string formItemNombre = "--" + boundary + "\r\n" +
						"Content-Disposition: form-data; name=\"nombre_del_archivo\"\r\n\r\n" +
						nombreArchivo + "\r\n";
					byte[] nombreBytes = Encoding.UTF8.GetBytes(formItemNombre);
					requestStream.Write(nombreBytes, 0, nombreBytes.Length);

					// --- Agregar campo tipo ---
					string formItemTipo = "--" + boundary + "\r\n" +
						"Content-Disposition: form-data; name=\"tipo\"\r\n\r\n" +
						tipo + "\r\n";
					byte[] tipoBytes = Encoding.UTF8.GetBytes(formItemTipo);
					requestStream.Write(tipoBytes, 0, tipoBytes.Length);

					// --- Agregar archivo ---
					string header = "--" + boundary + "\r\n" +
						"Content-Disposition: form-data; name=\"file\"; filename=\"" + fileName + "\"\r\n" +
						"Content-Type: application/pdf\r\n\r\n";
					byte[] headerBytes = Encoding.UTF8.GetBytes(header);
					requestStream.Write(headerBytes, 0, headerBytes.Length);

					// Archivo en sí
					requestStream.Write(fileBytes, 0, fileBytes.Length);

					// Cierre del multipart
					byte[] endBoundaryBytes = Encoding.UTF8.GetBytes("\r\n--" + boundary + "--\r\n");
					requestStream.Write(endBoundaryBytes, 0, endBoundaryBytes.Length);
				}

				using (HttpWebResponse response = (HttpWebResponse)request.GetResponse())
				{
					using (StreamReader reader = new StreamReader(response.GetResponseStream()))
					{
						string responseText = reader.ReadToEnd();
						Console.WriteLine("Respuesta del servidor: " + responseText);
						MessageBox.Show("Archivo y datos enviados correctamente.", "Éxito", MessageBoxButtons.OK, MessageBoxIcon.Information);
					}
				}
			}
			catch (Exception ex)
			{
				MessageBox.Show("Error al subir archivo: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
			}
		}

		private void button1_Click(object sender, System.EventArgs e)
		{
			string sUrlNube = this.txtImagen1.Text;
			if(bNuevo )VisualizaImagen(this.txtImagen1.Text.ToString().Trim());
			if(!bNuevo)
			{
				  System.Diagnostics.Process.Start("chrome.exe", sUrlNube);
			}

		}

		private void btnImg1_Click(object sender, System.EventArgs e)
		{
			/* Cargar imagen */
			OpenFileDialog ofd = new OpenFileDialog();
			ofd.InitialDirectory = @"C:\";
			ofd.Filter = "Archivos de imagen (*.jpg;*.jpeg;*.png;*.gif)|*.jpg;*.jpeg;*.png;*.gif";

			if (ofd.ShowDialog() == DialogResult.OK)
			{
				string archivoSeleccionado = ofd.FileName;
				this.txtImagen1.Text = archivoSeleccionado;

				// Llama al método y obtiene la nueva ruta
				string nuevaRuta = NuevoGoogle(archivoSeleccionado, 1);


				if ( nuevaRuta == null || nuevaRuta == "")
				{
					
					MessageBox.Show("Hubo un problema al subir la imagen.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
				}
				else
				{
					this.txtImagen1.Text = nuevaRuta;
				}
			}
			else
			{
				MessageBox.Show("No se seleccionó una imagen.", "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
			}
		}
     string nuevoNombreArchivoImgone ="";
	  	 string nuevoNombreArchivoImgTwo ="";
		 string nuevoNombreArchivoImgTree ="";

		public string NuevoGoogle(string txtDirectorioarchivo, int tipo )
		{
			if ( !File.Exists(txtDirectorioarchivo))
			{
				MessageBox.Show("El archivo seleccionado no existe.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
				return null;
			}

			try
			{
				// Configuración de la ruta destino
				string carpetaDestino = @"\\192.168.1.10\Proyectos\DTImagenes";

				// Crear carpeta si no existe
				if (!Directory.Exists(carpetaDestino))
				{
					Directory.CreateDirectory(carpetaDestino);
				}

				// Obtener info del archivo original
				FileInfo fi = new FileInfo(txtDirectorioarchivo);
				string extension = fi.Extension;

				// Generar nombre único con timestamp
				string timestamp = DateTime.Now.ToString("yyyyMMdd_HHmmss_fff");
				string nuevoNombreArchivo = "archivo_" + timestamp + extension;
				if (tipo == 1 ) nuevoNombreArchivoImgone =  nuevoNombreArchivo;
				if (tipo == 2 ) nuevoNombreArchivoImgTwo =  nuevoNombreArchivo;
				if (tipo == 3 ) nuevoNombreArchivoImgTree =  nuevoNombreArchivo;



				string rutaDestinoFinal = Path.Combine(carpetaDestino, nuevoNombreArchivo);

				// Si ya existe un archivo con ese nombre, lo elimina (altamente improbable por timestamp)
				if (File.Exists(rutaDestinoFinal))
				{
					File.Delete(rutaDestinoFinal);
				}

				// Conexión a red (forzar acceso)
				Directory.GetDirectories(@"\\192.168.1.10\Proyectos");
				Process.Start("cmd.exe", "/C net use \\\\192.168.1.10 /user:latinium 10dgpv15.p");

				// Mover archivo
				File.Copy(txtDirectorioarchivo, rutaDestinoFinal);

				// Confirmación opcional
				return rutaDestinoFinal;
			}
			catch (Exception ex)
			{
				MessageBox.Show("Error: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
				return null;
			}
		}


		private void btnImg2_Click(object sender, System.EventArgs e)
		{

			/* Cargar imagen */
			OpenFileDialog ofd = new OpenFileDialog();
			ofd.InitialDirectory = @"C:\";
			ofd.Filter = "Archivos de imagen (*.jpg;*.jpeg;*.png;*.gif)|*.jpg;*.jpeg;*.png;*.gif";

			if (ofd.ShowDialog() == DialogResult.OK)
			{
				string archivoSeleccionado = ofd.FileName;
				this.txtImagen2.Text = archivoSeleccionado;

				// Llama al método y obtiene la nueva ruta
				string nuevaRuta = NuevoGoogle(archivoSeleccionado, 2);

				if ( nuevaRuta == null || nuevaRuta == "")
				{
					
					MessageBox.Show("Hubo un problema al subir la imagen.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
				}
				else
				{
					this.txtImagen2.Text = nuevaRuta;
				}
			}
			else
			{
				MessageBox.Show("No se seleccionó una imagen.", "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
			}


			
		}

		private void btnImg3_Click(object sender, System.EventArgs e)
		{
			/* Cargar imagen */
			OpenFileDialog ofd = new OpenFileDialog();
			ofd.InitialDirectory = @"C:\";
			ofd.Filter = "Archivos de imagen (*.jpg;*.jpeg;*.png;*.gif)|*.jpg;*.jpeg;*.png;*.gif";

			if (ofd.ShowDialog() == DialogResult.OK)
			{
				string archivoSeleccionado = ofd.FileName;
				this.txtImagen3.Text = archivoSeleccionado;

				// Llama al método y obtiene la nueva ruta
				string nuevaRuta = NuevoGoogle(archivoSeleccionado, 3);

				if ( nuevaRuta == null || nuevaRuta == "")
				{
					
					MessageBox.Show("Hubo un problema al subir la imagen.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
				}
				else
				{
					this.txtImagen3.Text = nuevaRuta;
				}
			}
			else
			{
				MessageBox.Show("No se seleccionó una imagen.", "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
			}

		}

		private void button2_Click(object sender, System.EventArgs e)
		{
			
			string sUrlNube = this.txtImagen2.Text;
			if(bNuevo )VisualizaImagen(this.txtImagen2.Text.ToString().Trim());
			if(!bNuevo)
			{
				System.Diagnostics.Process.Start("chrome.exe", sUrlNube);
			}
		
		}

		private void button3_Click(object sender, System.EventArgs e)
		{

			string sUrlNube = this.txtImagen3.Text;
			if(bNuevo )VisualizaImagen(this.txtImagen3.Text.ToString().Trim());
			if(!bNuevo)
			{
				System.Diagnostics.Process.Start("chrome.exe", sUrlNube);
			}

		}

		private void txtCIEntrega_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.txtNombreEntrega.Focus();
		}

		private void txtCodigo_Enter(object sender, System.EventArgs e)
		{
			this.txtCodigo.SelectAll();	
		}

		private void txtCodigo_Click(object sender, System.EventArgs e)
		{
			this.txtCodigo.SelectAll();	
		}

		private void txtNombre_Click(object sender, System.EventArgs e)
		{
			this.txtNombre.SelectAll();	
		}

		private void txtNombre_Enter(object sender, System.EventArgs e)
		{
			this.txtNombre.SelectAll();
		}

		private void txtDireccion_Enter(object sender, System.EventArgs e)
		{
			this.txtDireccion.SelectAll();
		}

		private void txtDireccion_Click(object sender, System.EventArgs e)
		{
			this.txtDireccion.SelectAll();
		}

		private void txtTelefono_Click(object sender, System.EventArgs e)
		{
			this.txtTelefono.SelectAll();
		}

		private void txtTelefono_Enter(object sender, System.EventArgs e)
		{
			this.txtTelefono.SelectAll();
		}

		private void txtFonoAdicional_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.txtCelAdicional.Focus();
		}

		private void txtFonoAdicional_Click(object sender, System.EventArgs e)
		{
			this.txtFonoAdicional.SelectAll();		
		}

		private void txtFonoAdicional_Enter(object sender, System.EventArgs e)
		{
			this.txtFonoAdicional.SelectAll();		
		}

		private void txtCelular_Enter(object sender, System.EventArgs e)
		{
			this.txtCelular.SelectAll();
		}

		private void txtCelular_Click(object sender, System.EventArgs e)
		{
			this.txtCelular.SelectAll();
		}

		private void txtCIEntrega_Click(object sender, System.EventArgs e)
		{
			this.txtCIEntrega.SelectAll();
		}

		private void txtCIEntrega_Enter(object sender, System.EventArgs e)
		{
			this.txtCIEntrega.SelectAll();
		}

		private void txtNombreEntrega_Click(object sender, System.EventArgs e)
		{
			this.txtNombreEntrega.SelectAll();
		}

		private void txtNombreEntrega_Enter(object sender, System.EventArgs e)
		{
			this.txtNombreEntrega.SelectAll();
		}	
		
		private void dtFecha_ValueChanged(object sender, System.EventArgs e)
		{
			if (this.dtFecha.Value != System.DBNull.Value)
				FuncionesProcedimientos.ListaBodegas(this.cmbBodega, MenuLatinium.IdUsuario, 50, (DateTime)this.dtFecha.Value, false, cdsSeteoF);
		}

		private void frmDTIngreso_Closing(object sender, System.ComponentModel.CancelEventArgs e)
		{
			if (bEdicion)
			{
				MessageBox.Show("No puede salir mientras edita una ingreso", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				e.Cancel = true;
				return;
			}
		}

		private void cmbBodega_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}

		private void cmbUsuario_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) 
			{
				if(this.cmbAsignaTecnico.Enabled) this.cmbAsignaTecnico.Focus(); 
				else 
				{ 
					if (iInternoExterno == 1) { if (this.txtCodigo.Enabled) this.txtCodigo.Focus(); else this.txtFonoAdicional.Focus();} 
					else this.txtCodigo.Focus();
				}
			}
		}

		private void txtCelAdicional_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) if (iInternoExterno == 1) this.txtEstadoFisico.Focus(); else this.txtSerial.Focus();
		}

		private void txtFonoAdicional_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			Funcion.CamposNumericos(sender, e);
		}

		private void txtCelAdicional_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			Funcion.CamposNumericos(sender, e);
		}

		private void txtCIEntrega_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			Funcion.CamposNumericos(sender, e);
		}

		private void txtNombreEntrega_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.btnGuardar.Focus();
		}

		private void chkAccesorios_CheckedChanged(object sender, System.EventArgs e)
		{
			if (!bEdicion) return;

			if(this.chkAccesorios.Checked)
			{
				this.rbAccesorios.Enabled = true;
				this.txtAccesorios.Enabled = true;
				this.rbAccesorios.Value = 3;
				this.txtAccesorios.Focus();
			}
			else
			{
				this.rbAccesorios.CheckedIndex = -1;
				this.txtAccesorios.Text = "";
				this.rbAccesorios.Enabled = false;
				this.txtAccesorios.Enabled = false;
			}
		}

		private void chkBordes_CheckedChanged(object sender, System.EventArgs e)
		{
			if (!bEdicion) return;

			if(this.chkBordes.Checked)
			{
				this.rbBordes.Enabled = true;
			}
			else
			{
				this.rbBordes.CheckedIndex = -1;
				this.rbBordes.Enabled = false;
			}		
		}

		private void chkPines_CheckedChanged(object sender, System.EventArgs e)
		{
			if (!bEdicion) return;

			if(this.chkPines.Checked)
			{
				this.rbPines.Enabled = true;
			}
			else
			{
				this.rbPines.CheckedIndex = -1;
				this.rbPines.Enabled = false;
			}
		}

		private void chkBotones_CheckedChanged(object sender, System.EventArgs e)
		{
			if (!bEdicion) return;

			if(this.chkBotones.Checked)
			{
				this.rbBotones.Enabled = true;
			}
			else
			{
				this.rbBotones.CheckedIndex = -1;
				this.rbBotones.Enabled = false;
			}
		}

		private void chkSupDelantera_CheckedChanged(object sender, System.EventArgs e)
		{
			if (!bEdicion) return;

			if(this.chkSupDelantera.Checked)
			{
				this.rbSupDelantera.Enabled = true;
			}
			else
			{
				this.rbSupDelantera.CheckedIndex = -1;
				this.rbSupDelantera.Enabled = false;
			}
		}

		private void chkSupPosterior_CheckedChanged(object sender, System.EventArgs e)
		{
			if (!bEdicion) return;

			if(this.chkSupPosterior.Checked)
			{
				this.rbSupPosterior.Enabled = true;
			}
			else
			{
				this.rbSupPosterior.CheckedIndex = -1;
				this.rbSupPosterior.Enabled = false;
			}
		}

		private void chkSerial_CheckedChanged(object sender, System.EventArgs e)
		{
			if (!bEdicion) return;

			if(this.chkSerial.Checked)
			{
				this.rbSerial.Enabled = true;
			}
			else
			{
				this.rbSerial.CheckedIndex = -1;
				this.rbSerial.Enabled = false;
			}
		}

		private void frmDTIngreso_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (!bEdicion)
			{
				if (e.KeyCode == Keys.N) if (miAcceso.Nuevo)	if (this.btnNuevo.Enabled) this.btnNuevo_Click(sender, e);
				if (e.KeyCode == Keys.B) if (miAcceso.Buscar) if (this.btnConsultar.Enabled) this.btnConsultar_Click(sender, e);
				if ((int)this.txtIdIngreso.Value > 0)
				{
					if (e.KeyCode == Keys.E) if (miAcceso.Editar) if (this.btnEditar.Enabled) this.btnEditar_Click(sender, e);					
					if (e.KeyCode == Keys.F5)
					{
						miAcceso = new Acceso(cdsSeteoF, "1002");

						//this.Consultar((int)this.txtNumIdCompra.Value);					
					}
				}

				if (e.KeyCode == Keys.Escape) this.Close();				
			}
			else
			{
				if (e.KeyCode == Keys.Escape)
					if (DialogResult.Yes == MessageBox.Show("Esta seguro de Cancelar la Edición", "Point Technology", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2))
						this.btnCancelar_Click(sender, e);

				if (e.KeyCode == Keys.F12) this.btnGuardar_Click(sender, e);				
			}
		}

		private void cmbAsignaTecnico_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) if (iInternoExterno == 1) { if (miAcceso.BEditarNombre) this.txtCodigo.Focus(); else this.txtFonoAdicional.Focus(); } 
																	 else this.txtCodigo.Focus();
		}

		private void txtEstadoFisico_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) if (this.chkEmpqs.Checked) this.txtEmpqsObservaciones.Focus(); else this.chkEmpqs.Focus();
		}

		private void chkEmpqs_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.chkAccesorios.Focus();
		}

		private void chkAccesorios_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.txtDaño.Focus();
		}

		private void txtEmpqsObservaciones_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) if (this.chkAccesorios.Checked) this.txtAccesorios.Focus(); else this.chkAccesorios.Focus();
		}

		private void txtAccesorios_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.txtDaño.Focus();
		}

		private void dtFecha_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (e.KeyChar == 8) e.Handled = true;
		}

		private void btnCliente_Click(object sender, System.EventArgs e)
		{
			if (this.txtCodigo.Text.Length == 0)
			{
				MessageBox.Show("Ingrese el Numero de Identificacion del Cliente", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.txtCodigo.Focus();
				return;
			}
			
			FuncionesProcedimientos.AbreFacturacionClientes(idTipoRuc, this.txtCodigo.Text.ToString(), true,false);

			CargaCliente();			
		}

		private void cmbBodega_ValueChanged(object sender, System.EventArgs e)
		{
			if (this.cmbBodega.ActiveRow != null)
			{
				if (this.cmbUsuario.ActiveRow != null) this.cmbUsuario.Value = System.DBNull.Value;
				this.cmbUsuario.DataSource = Funcion.dvProcedimiento(cdsSeteoF, string.Format("Exec ListaUsuarios {0}", (int)this.cmbBodega.Value));
				
				if (this.cmbAsignaTecnico.ActiveRow != null) this.cmbAsignaTecnico.Value = System.DBNull.Value;
				this.cmbAsignaTecnico.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Exec ListaUsuarios 20");

				if ((int)this.cmbBodega.Value == 20 || (int)this.cmbBodega.Value == 95) 
				{
					if (miAcceso.BAsignaTecnicos) this.cmbAsignaTecnico.Enabled = true; else this.cmbAsignaTecnico.Enabled = false; 
				}
				else this.cmbAsignaTecnico.Enabled = false;
			}
		}

		private void mnuImprimir_Popup(object sender, System.EventArgs e)
		{
		
		}
		
	}
}

