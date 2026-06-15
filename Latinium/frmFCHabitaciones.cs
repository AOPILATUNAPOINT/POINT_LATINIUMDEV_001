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
using System.Runtime.InteropServices;
using System.Diagnostics;
using System.IO;
using CrystalDecisions;
using CrystalDecisions.CrystalReports.Engine;
using CrystalDecisions.ReportSource;
using System.Text;
using System.Net;
using System.Web;
using System.Xml;
using System.Xml.Schema;
using System.Security.Cryptography.Xml;
using System.Security.Cryptography;
using System.Security.Cryptography.X509Certificates;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de frmFCHabitaciones.
	/// </summary>
	public class frmFCHabitaciones : System.Windows.Forms.Form
	{
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		
		int NumrHab = 0;
		string NomrBodga = "";
		int NBodga = 0;
		int idNhab = 0;
		int idClint = 0;
		int idCpaFC = 0;
		int idProyct = 0;
		decimal PorcentajeIva = 0.00m;
		int idAsient = 0;
		bool bImpreso = false;
		string NumFactura = "";
		private Acceso miAcceso;
		public int RegistroFactura = 0;
		bool bEliminaAlValidar = false;
		string aplicapromocion = "NO";
		int idCpnWeb = 0;
		string respuestapromocion = "";
		int tpopromocion = 0;
		decimal valorpromo = 0.00m;

		#region controles pantalla
		private System.Windows.Forms.GroupBox groupBox1;
		private Infragistics.Win.UltraWinGrid.UltraGrid ugArticulos;
		private System.Windows.Forms.GroupBox groupBox2;
		private System.Windows.Forms.Label label1;
		private System.Windows.Forms.Button button1;
		private System.Windows.Forms.Label label2;
		private System.Windows.Forms.Label label3;
		private System.Windows.Forms.Label label4;
		private System.Windows.Forms.Label label5;
		private System.Windows.Forms.Label label6;
		private System.Windows.Forms.Label label7;
		private System.Windows.Forms.Label label8;
		private Infragistics.Win.UltraWinGrid.UltraGrid ugFacRelizadas;
		private Infragistics.Win.UltraWinGrid.UltraGrid ugDetalleFactura;
		private System.Windows.Forms.Button bntGrabar;
		private System.Windows.Forms.Button btnImprimir;
		private System.Windows.Forms.TextBox txtNumFactura;
		private System.Windows.Forms.TextBox txtNumCedula;
		private System.Windows.Forms.TextBox txtNombre;
		private System.Windows.Forms.TextBox txtTelefono;
		private C1.Data.C1DataSet cdsSeteo;
		private System.Windows.Forms.ToolTip toolTip1;
		private System.Windows.Forms.Button button2;
		private System.Windows.Forms.Button button3;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource1;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource2;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource4;
		private System.Windows.Forms.Label label9;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource3;
		private System.Windows.Forms.CheckBox chkPrecios;
		private System.Windows.Forms.TextBox txtDireccion;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtTotal;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtIva;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtSubTotal;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbFormaPago;
		private System.Windows.Forms.Label label11;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtFecha;
		private System.Windows.Forms.Label label10;
		private System.Windows.Forms.Label label12;
		private System.Windows.Forms.TextBox txtBskArticulo;
		private DevExpress.XtraEditors.SimpleButton btBuscar;
		private System.Windows.Forms.Label label13;
		private System.Windows.Forms.Label label14;
		private System.Windows.Forms.Label label15;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtTiempo;
		private System.Windows.Forms.TextBox txtUnidadTaxi;
		private System.Windows.Forms.Label label16;
		private System.Windows.Forms.Label label17;
		private System.Windows.Forms.Label label18;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbVendedor;
		private System.Windows.Forms.TextBox txtnPromo;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbCobrador;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbCajero;
		private System.Windows.Forms.Label label19;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource5;
		private System.Windows.Forms.Label lblIdentificacion;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbTipoRuc;
		private System.Windows.Forms.Label label20;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtDescnto;
		private System.ComponentModel.IContainer components;
		#endregion controles pantalla

		public frmFCHabitaciones(int BskHabitacion, string NombredeBodega, int idnroHab)
		{
			//
			// Necesario para admitir el Diseñador de Windows Forms
			//
			InitializeComponent();
			NumrHab = BskHabitacion;
			NomrBodga = NombredeBodega;
			//NBodga = NumrBodega;
			idNhab = idnroHab;

			this.label9.Text = NombredeBodega.ToString().Trim()+" - Habitación #:"+NumrHab.ToString().Trim();

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
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmFCHabitaciones));
			Infragistics.Win.Appearance appearance1 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idArticulo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idGrupoArticulo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idSubGrupo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Articulo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Precio1");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Precio2");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("UnboundColumn1", 0);
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idArticulo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idGrupoArticulo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn3 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idSubGrupo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn4 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Articulo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn5 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Precio1");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn6 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Precio2");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn7 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idSubGrupo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn8 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idGrupoArticulo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn9 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("SubGrupo");
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idTipoRuc");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("TipoRuc");
			Infragistics.Win.Appearance appearance11 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand3 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPersonal");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn11 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn12 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn13 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Turno");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn10 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idPersonal");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn11 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn12 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Codigo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn13 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Turno");
			Infragistics.Win.Appearance appearance12 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand4 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn14 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPersonal");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn15 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn16 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.Appearance appearance13 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand5 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn17 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPersonal");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn18 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn19 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.Appearance appearance14 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance15 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton1 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.Appearance appearance16 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand6 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn20 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idFormaPago");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn21 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FormaPago");
			Infragistics.Win.Appearance appearance17 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance18 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance19 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance20 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand7 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn22 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idDetCompra");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn23 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCompra");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn24 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idArticulo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn25 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Articulo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn26 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cantidad");
			Infragistics.Win.Appearance appearance21 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn27 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Precio");
			Infragistics.Win.Appearance appearance22 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn28 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("PDscto");
			Infragistics.Win.Appearance appearance23 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn29 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("SubTotal");
			Infragistics.Win.Appearance appearance24 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance25 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance26 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance27 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance28 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance29 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance30 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn14 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idDetCompra");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn15 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCompra");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn16 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idArticulo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn17 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Articulo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn18 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Cantidad");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn19 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Precio");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn20 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("PDscto");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn21 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("SubTotal");
			Infragistics.Win.Appearance appearance31 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand8 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn30 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCompra");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn31 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Ruc");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn32 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Numero");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn33 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Fecha");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn34 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idVendedor");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn35 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCliente");
			Infragistics.Win.Appearance appearance32 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance33 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance34 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance35 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance36 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance37 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn22 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCompra");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn23 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Ruc");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn24 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Numero");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn25 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Fecha");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn26 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idVendedor");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn27 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCliente");
			this.groupBox1 = new System.Windows.Forms.GroupBox();
			this.btBuscar = new DevExpress.XtraEditors.SimpleButton();
			this.label12 = new System.Windows.Forms.Label();
			this.label9 = new System.Windows.Forms.Label();
			this.button3 = new System.Windows.Forms.Button();
			this.button2 = new System.Windows.Forms.Button();
			this.ugArticulos = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource3 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.txtBskArticulo = new System.Windows.Forms.TextBox();
			this.ultraDataSource1 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.groupBox2 = new System.Windows.Forms.GroupBox();
			this.txtDescnto = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label20 = new System.Windows.Forms.Label();
			this.lblIdentificacion = new System.Windows.Forms.Label();
			this.cmbTipoRuc = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.cmbCajero = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.ultraDataSource5 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.label19 = new System.Windows.Forms.Label();
			this.cmbCobrador = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.cmbVendedor = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.label18 = new System.Windows.Forms.Label();
			this.label17 = new System.Windows.Forms.Label();
			this.txtnPromo = new System.Windows.Forms.TextBox();
			this.label16 = new System.Windows.Forms.Label();
			this.txtTiempo = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label15 = new System.Windows.Forms.Label();
			this.txtUnidadTaxi = new System.Windows.Forms.TextBox();
			this.label14 = new System.Windows.Forms.Label();
			this.label13 = new System.Windows.Forms.Label();
			this.label10 = new System.Windows.Forms.Label();
			this.dtFecha = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.label11 = new System.Windows.Forms.Label();
			this.cmbFormaPago = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.chkPrecios = new System.Windows.Forms.CheckBox();
			this.txtTotal = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtIva = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtSubTotal = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.btnImprimir = new System.Windows.Forms.Button();
			this.bntGrabar = new System.Windows.Forms.Button();
			this.label8 = new System.Windows.Forms.Label();
			this.label7 = new System.Windows.Forms.Label();
			this.label6 = new System.Windows.Forms.Label();
			this.ugDetalleFactura = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource4 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.txtTelefono = new System.Windows.Forms.TextBox();
			this.label5 = new System.Windows.Forms.Label();
			this.txtDireccion = new System.Windows.Forms.TextBox();
			this.label4 = new System.Windows.Forms.Label();
			this.txtNombre = new System.Windows.Forms.TextBox();
			this.label3 = new System.Windows.Forms.Label();
			this.txtNumCedula = new System.Windows.Forms.TextBox();
			this.label2 = new System.Windows.Forms.Label();
			this.button1 = new System.Windows.Forms.Button();
			this.txtNumFactura = new System.Windows.Forms.TextBox();
			this.label1 = new System.Windows.Forms.Label();
			this.ugFacRelizadas = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource2 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.cdsSeteo = new C1.Data.C1DataSet();
			this.toolTip1 = new System.Windows.Forms.ToolTip(this.components);
			this.groupBox1.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.ugArticulos)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource3)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).BeginInit();
			this.groupBox2.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.txtDescnto)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbTipoRuc)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCajero)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource5)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCobrador)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbVendedor)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtTiempo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFecha)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbFormaPago)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtTotal)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIva)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtSubTotal)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ugDetalleFactura)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource4)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ugFacRelizadas)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteo)).BeginInit();
			this.SuspendLayout();
			// 
			// groupBox1
			// 
			this.groupBox1.Controls.Add(this.btBuscar);
			this.groupBox1.Controls.Add(this.label12);
			this.groupBox1.Controls.Add(this.label9);
			this.groupBox1.Controls.Add(this.button3);
			this.groupBox1.Controls.Add(this.button2);
			this.groupBox1.Controls.Add(this.ugArticulos);
			this.groupBox1.Controls.Add(this.txtBskArticulo);
			this.groupBox1.Location = new System.Drawing.Point(0, 8);
			this.groupBox1.Name = "groupBox1";
			this.groupBox1.Size = new System.Drawing.Size(552, 544);
			this.groupBox1.TabIndex = 0;
			this.groupBox1.TabStop = false;
			// 
			// btBuscar
			// 
			this.btBuscar.Image = ((System.Drawing.Image)(resources.GetObject("btBuscar.Image")));
			this.btBuscar.Location = new System.Drawing.Point(312, 64);
			this.btBuscar.Name = "btBuscar";
			this.btBuscar.Size = new System.Drawing.Size(72, 22);
			this.btBuscar.TabIndex = 232;
			this.btBuscar.Text = "&Buscar";
			this.btBuscar.Click += new System.EventHandler(this.btBuscar_Click);
			// 
			// label12
			// 
			this.label12.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label12.Location = new System.Drawing.Point(8, 48);
			this.label12.Name = "label12";
			this.label12.Size = new System.Drawing.Size(64, 16);
			this.label12.TabIndex = 21;
			this.label12.Text = "ARTICULO";
			// 
			// label9
			// 
			this.label9.Font = new System.Drawing.Font("Microsoft Sans Serif", 14F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label9.ForeColor = System.Drawing.Color.Red;
			this.label9.Location = new System.Drawing.Point(16, 16);
			this.label9.Name = "label9";
			this.label9.Size = new System.Drawing.Size(472, 32);
			this.label9.TabIndex = 19;
			this.label9.Text = "label9";
			this.label9.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// button3
			// 
			this.button3.Image = ((System.Drawing.Image)(resources.GetObject("button3.Image")));
			this.button3.Location = new System.Drawing.Point(496, 8);
			this.button3.Name = "button3";
			this.button3.Size = new System.Drawing.Size(48, 32);
			this.button3.TabIndex = 18;
			this.toolTip1.SetToolTip(this.button3, "Alimentos");
			this.button3.Visible = false;
			this.button3.Click += new System.EventHandler(this.button3_Click);
			// 
			// button2
			// 
			this.button2.Image = ((System.Drawing.Image)(resources.GetObject("button2.Image")));
			this.button2.Location = new System.Drawing.Point(496, 40);
			this.button2.Name = "button2";
			this.button2.Size = new System.Drawing.Size(48, 32);
			this.button2.TabIndex = 17;
			this.toolTip1.SetToolTip(this.button2, "Bebidas");
			this.button2.Visible = false;
			this.button2.Click += new System.EventHandler(this.button2_Click);
			// 
			// ugArticulos
			// 
			this.ugArticulos.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.ugArticulos.Cursor = System.Windows.Forms.Cursors.Default;
			this.ugArticulos.DataSource = this.ultraDataSource3;
			appearance1.BackColor = System.Drawing.Color.White;
			this.ugArticulos.DisplayLayout.Appearance = appearance1;
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridColumn2.Hidden = true;
			ultraGridColumn3.Header.VisiblePosition = 4;
			ultraGridColumn3.Hidden = true;
			ultraGridColumn4.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn4.Header.VisiblePosition = 3;
			ultraGridColumn4.Width = 410;
			ultraGridColumn5.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn5.Format = "#,##0.00";
			ultraGridColumn5.Header.Caption = "Precio";
			ultraGridColumn5.Header.VisiblePosition = 5;
			ultraGridColumn5.Width = 80;
			ultraGridColumn6.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn6.Format = "#,##0.00";
			ultraGridColumn6.Header.Caption = "Reserva";
			ultraGridColumn6.Header.VisiblePosition = 6;
			ultraGridColumn6.Hidden = true;
			ultraGridColumn6.Width = 80;
			ultraGridColumn7.ButtonDisplayStyle = Infragistics.Win.UltraWinGrid.ButtonDisplayStyle.Always;
			appearance2.Image = ((object)(resources.GetObject("appearance2.Image")));
			ultraGridColumn7.CellButtonAppearance = appearance2;
			ultraGridColumn7.Header.Caption = "Seleccionar";
			ultraGridColumn7.Header.VisiblePosition = 2;
			ultraGridColumn7.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.Button;
			ultraGridColumn7.Width = 20;
			ultraGridBand1.Columns.AddRange(new object[] {
																										 ultraGridColumn1,
																										 ultraGridColumn2,
																										 ultraGridColumn3,
																										 ultraGridColumn4,
																										 ultraGridColumn5,
																										 ultraGridColumn6,
																										 ultraGridColumn7});
			this.ugArticulos.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			appearance3.ForeColor = System.Drawing.Color.Black;
			appearance3.ForeColorDisabled = System.Drawing.Color.Black;
			this.ugArticulos.DisplayLayout.Override.ActiveRowAppearance = appearance3;
			this.ugArticulos.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.TemplateOnBottom;
			this.ugArticulos.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.ugArticulos.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.False;
			appearance4.BackColor = System.Drawing.Color.Transparent;
			this.ugArticulos.DisplayLayout.Override.CardAreaAppearance = appearance4;
			appearance5.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance5.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance5.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance5.FontData.BoldAsString = "True";
			appearance5.FontData.Name = "Arial";
			appearance5.FontData.SizeInPoints = 10F;
			appearance5.ForeColor = System.Drawing.Color.White;
			appearance5.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.ugArticulos.DisplayLayout.Override.HeaderAppearance = appearance5;
			appearance6.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance6.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance6.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ugArticulos.DisplayLayout.Override.RowAlternateAppearance = appearance6;
			appearance7.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance7.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance7.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ugArticulos.DisplayLayout.Override.RowSelectorAppearance = appearance7;
			appearance8.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance8.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance8.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ugArticulos.DisplayLayout.Override.SelectedRowAppearance = appearance8;
			this.ugArticulos.Font = new System.Drawing.Font("Microsoft Sans Serif", 13F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.ugArticulos.Location = new System.Drawing.Point(8, 88);
			this.ugArticulos.Name = "ugArticulos";
			this.ugArticulos.Size = new System.Drawing.Size(536, 448);
			this.ugArticulos.TabIndex = 16;
			this.ugArticulos.DoubleClick += new System.EventHandler(this.ugArticulos_DoubleClick);
			this.ugArticulos.BeforeSelectChange += new Infragistics.Win.UltraWinGrid.BeforeSelectChangeEventHandler(this.ugArticulos_BeforeSelectChange);
			this.ugArticulos.ClickCellButton += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.ugArticulos_ClickCellButton);
			this.ugArticulos.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.ugArticulos_InitializeLayout);
			// 
			// ultraDataSource3
			// 
			ultraDataColumn1.DataType = typeof(int);
			ultraDataColumn2.DataType = typeof(int);
			ultraDataColumn3.DataType = typeof(int);
			ultraDataColumn5.DataType = typeof(System.Decimal);
			ultraDataColumn6.DataType = typeof(System.Decimal);
			this.ultraDataSource3.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn1,
																																 ultraDataColumn2,
																																 ultraDataColumn3,
																																 ultraDataColumn4,
																																 ultraDataColumn5,
																																 ultraDataColumn6});
			// 
			// txtBskArticulo
			// 
			this.txtBskArticulo.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtBskArticulo.Location = new System.Drawing.Point(8, 65);
			this.txtBskArticulo.Name = "txtBskArticulo";
			this.txtBskArticulo.Size = new System.Drawing.Size(296, 20);
			this.txtBskArticulo.TabIndex = 231;
			this.txtBskArticulo.Text = "";
			this.txtBskArticulo.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtBskArticulo_KeyPress);
			// 
			// ultraDataSource1
			// 
			ultraDataColumn7.DataType = typeof(int);
			ultraDataColumn8.DataType = typeof(int);
			this.ultraDataSource1.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn7,
																																 ultraDataColumn8,
																																 ultraDataColumn9});
			// 
			// groupBox2
			// 
			this.groupBox2.Controls.Add(this.txtDescnto);
			this.groupBox2.Controls.Add(this.label20);
			this.groupBox2.Controls.Add(this.lblIdentificacion);
			this.groupBox2.Controls.Add(this.cmbTipoRuc);
			this.groupBox2.Controls.Add(this.cmbCajero);
			this.groupBox2.Controls.Add(this.label19);
			this.groupBox2.Controls.Add(this.cmbCobrador);
			this.groupBox2.Controls.Add(this.cmbVendedor);
			this.groupBox2.Controls.Add(this.label18);
			this.groupBox2.Controls.Add(this.label17);
			this.groupBox2.Controls.Add(this.txtnPromo);
			this.groupBox2.Controls.Add(this.label16);
			this.groupBox2.Controls.Add(this.txtTiempo);
			this.groupBox2.Controls.Add(this.label15);
			this.groupBox2.Controls.Add(this.txtUnidadTaxi);
			this.groupBox2.Controls.Add(this.label14);
			this.groupBox2.Controls.Add(this.label13);
			this.groupBox2.Controls.Add(this.label10);
			this.groupBox2.Controls.Add(this.dtFecha);
			this.groupBox2.Controls.Add(this.label11);
			this.groupBox2.Controls.Add(this.cmbFormaPago);
			this.groupBox2.Controls.Add(this.chkPrecios);
			this.groupBox2.Controls.Add(this.txtTotal);
			this.groupBox2.Controls.Add(this.txtIva);
			this.groupBox2.Controls.Add(this.txtSubTotal);
			this.groupBox2.Controls.Add(this.btnImprimir);
			this.groupBox2.Controls.Add(this.bntGrabar);
			this.groupBox2.Controls.Add(this.label8);
			this.groupBox2.Controls.Add(this.label7);
			this.groupBox2.Controls.Add(this.label6);
			this.groupBox2.Controls.Add(this.ugDetalleFactura);
			this.groupBox2.Controls.Add(this.txtTelefono);
			this.groupBox2.Controls.Add(this.label5);
			this.groupBox2.Controls.Add(this.txtDireccion);
			this.groupBox2.Controls.Add(this.label4);
			this.groupBox2.Controls.Add(this.txtNombre);
			this.groupBox2.Controls.Add(this.label3);
			this.groupBox2.Controls.Add(this.txtNumCedula);
			this.groupBox2.Controls.Add(this.label2);
			this.groupBox2.Controls.Add(this.button1);
			this.groupBox2.Controls.Add(this.txtNumFactura);
			this.groupBox2.Controls.Add(this.label1);
			this.groupBox2.Controls.Add(this.ugFacRelizadas);
			this.groupBox2.Location = new System.Drawing.Point(552, 8);
			this.groupBox2.Name = "groupBox2";
			this.groupBox2.Size = new System.Drawing.Size(712, 544);
			this.groupBox2.TabIndex = 1;
			this.groupBox2.TabStop = false;
			this.groupBox2.Text = "Facturación HABITACION";
			// 
			// txtDescnto
			// 
			this.txtDescnto.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			appearance9.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtDescnto.Appearance = appearance9;
			this.txtDescnto.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtDescnto.Enabled = false;
			this.txtDescnto.FormatString = "#,##0.00";
			this.txtDescnto.Location = new System.Drawing.Point(608, 454);
			this.txtDescnto.Name = "txtDescnto";
			this.txtDescnto.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtDescnto.PromptChar = ' ';
			this.txtDescnto.Size = new System.Drawing.Size(80, 21);
			this.txtDescnto.TabIndex = 250;
			this.txtDescnto.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtDescnto_KeyPress);
			this.txtDescnto.BeforeExitEditMode += new Infragistics.Win.BeforeExitEditModeEventHandler(this.txtDescnto_BeforeExitEditMode);
			this.txtDescnto.ValueChanged += new System.EventHandler(this.txtDescnto_ValueChanged);
			// 
			// label20
			// 
			this.label20.Location = new System.Drawing.Point(520, 456);
			this.label20.Name = "label20";
			this.label20.Size = new System.Drawing.Size(80, 16);
			this.label20.TabIndex = 249;
			this.label20.Text = "DESCUENTO";
			this.label20.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
			// 
			// lblIdentificacion
			// 
			this.lblIdentificacion.AutoSize = true;
			this.lblIdentificacion.Location = new System.Drawing.Point(16, 136);
			this.lblIdentificacion.Name = "lblIdentificacion";
			this.lblIdentificacion.Size = new System.Drawing.Size(87, 16);
			this.lblIdentificacion.TabIndex = 248;
			this.lblIdentificacion.Text = "Tipo Documento";
			// 
			// cmbTipoRuc
			// 
			appearance10.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbTipoRuc.Appearance = appearance10;
			this.cmbTipoRuc.CausesValidation = false;
			this.cmbTipoRuc.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbTipoRuc.Cursor = System.Windows.Forms.Cursors.Default;
			ultraGridColumn8.Header.VisiblePosition = 0;
			ultraGridColumn8.Hidden = true;
			ultraGridColumn9.Header.VisiblePosition = 1;
			ultraGridColumn9.Width = 233;
			ultraGridBand2.Columns.AddRange(new object[] {
																										 ultraGridColumn8,
																										 ultraGridColumn9});
			this.cmbTipoRuc.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			this.cmbTipoRuc.DisplayMember = "TipoRuc";
			this.cmbTipoRuc.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbTipoRuc.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbTipoRuc.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbTipoRuc.Location = new System.Drawing.Point(8, 152);
			this.cmbTipoRuc.Name = "cmbTipoRuc";
			this.cmbTipoRuc.Size = new System.Drawing.Size(128, 21);
			this.cmbTipoRuc.TabIndex = 247;
			this.cmbTipoRuc.ValueMember = "idTipoRuc";
			// 
			// cmbCajero
			// 
			appearance11.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbCajero.Appearance = appearance11;
			this.cmbCajero.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbCajero.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbCajero.DataSource = this.ultraDataSource5;
			ultraGridColumn10.Header.VisiblePosition = 0;
			ultraGridColumn10.Hidden = true;
			ultraGridColumn11.Header.VisiblePosition = 1;
			ultraGridColumn11.Width = 250;
			ultraGridColumn12.Header.VisiblePosition = 2;
			ultraGridColumn13.Header.VisiblePosition = 3;
			ultraGridColumn13.Width = 30;
			ultraGridBand3.Columns.AddRange(new object[] {
																										 ultraGridColumn10,
																										 ultraGridColumn11,
																										 ultraGridColumn12,
																										 ultraGridColumn13});
			this.cmbCajero.DisplayLayout.BandsSerializer.Add(ultraGridBand3);
			this.cmbCajero.DisplayMember = "Nombre";
			this.cmbCajero.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbCajero.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbCajero.Location = new System.Drawing.Point(256, 216);
			this.cmbCajero.Name = "cmbCajero";
			this.cmbCajero.Size = new System.Drawing.Size(152, 21);
			this.cmbCajero.TabIndex = 246;
			this.cmbCajero.ValueMember = "idPersonal";
			// 
			// ultraDataSource5
			// 
			ultraDataColumn13.DataType = typeof(int);
			this.ultraDataSource5.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn10,
																																 ultraDataColumn11,
																																 ultraDataColumn12,
																																 ultraDataColumn13});
			// 
			// label19
			// 
			this.label19.AutoSize = true;
			this.label19.Location = new System.Drawing.Point(216, 218);
			this.label19.Name = "label19";
			this.label19.Size = new System.Drawing.Size(38, 16);
			this.label19.TabIndex = 245;
			this.label19.Text = "Cajero";
			// 
			// cmbCobrador
			// 
			appearance12.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbCobrador.Appearance = appearance12;
			this.cmbCobrador.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbCobrador.Cursor = System.Windows.Forms.Cursors.Default;
			ultraGridColumn14.Header.VisiblePosition = 0;
			ultraGridColumn14.Hidden = true;
			ultraGridColumn15.Header.VisiblePosition = 1;
			ultraGridColumn15.Width = 335;
			ultraGridColumn16.Header.VisiblePosition = 2;
			ultraGridBand4.Columns.AddRange(new object[] {
																										 ultraGridColumn14,
																										 ultraGridColumn15,
																										 ultraGridColumn16});
			this.cmbCobrador.DisplayLayout.BandsSerializer.Add(ultraGridBand4);
			this.cmbCobrador.DisplayMember = "Nombre";
			this.cmbCobrador.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbCobrador.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbCobrador.Location = new System.Drawing.Point(56, 216);
			this.cmbCobrador.Name = "cmbCobrador";
			this.cmbCobrador.Size = new System.Drawing.Size(152, 21);
			this.cmbCobrador.TabIndex = 244;
			this.cmbCobrador.ValueMember = "idPersonal";
			// 
			// cmbVendedor
			// 
			appearance13.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbVendedor.Appearance = appearance13;
			this.cmbVendedor.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbVendedor.Cursor = System.Windows.Forms.Cursors.Default;
			ultraGridColumn17.Header.VisiblePosition = 0;
			ultraGridColumn17.Hidden = true;
			ultraGridColumn18.Header.VisiblePosition = 1;
			ultraGridColumn18.Width = 335;
			ultraGridColumn19.Header.VisiblePosition = 2;
			ultraGridBand5.Columns.AddRange(new object[] {
																										 ultraGridColumn17,
																										 ultraGridColumn18,
																										 ultraGridColumn19});
			this.cmbVendedor.DisplayLayout.BandsSerializer.Add(ultraGridBand5);
			this.cmbVendedor.DisplayMember = "Nombre";
			this.cmbVendedor.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbVendedor.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbVendedor.Location = new System.Drawing.Point(588, 72);
			this.cmbVendedor.Name = "cmbVendedor";
			this.cmbVendedor.Size = new System.Drawing.Size(116, 21);
			this.cmbVendedor.TabIndex = 243;
			this.cmbVendedor.ValueMember = "idPersonal";
			// 
			// label18
			// 
			this.label18.AutoSize = true;
			this.label18.Location = new System.Drawing.Point(8, 218);
			this.label18.Name = "label18";
			this.label18.Size = new System.Drawing.Size(51, 16);
			this.label18.TabIndex = 242;
			this.label18.Text = "Cobrador";
			// 
			// label17
			// 
			this.label17.AutoSize = true;
			this.label17.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label17.Location = new System.Drawing.Point(588, 56);
			this.label17.Name = "label17";
			this.label17.Size = new System.Drawing.Size(46, 16);
			this.label17.TabIndex = 240;
			this.label17.Text = "Guardia";
			// 
			// txtnPromo
			// 
			this.txtnPromo.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtnPromo.Location = new System.Drawing.Point(552, 152);
			this.txtnPromo.MaxLength = 20;
			this.txtnPromo.Name = "txtnPromo";
			this.txtnPromo.Size = new System.Drawing.Size(144, 20);
			this.txtnPromo.TabIndex = 238;
			this.txtnPromo.Text = "";
			this.txtnPromo.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
			this.txtnPromo.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtnPromo_KeyPress);
			this.txtnPromo.Validated += new System.EventHandler(this.textBox1_Validated);
			this.txtnPromo.TextChanged += new System.EventHandler(this.txtnPromo_TextChanged);
			// 
			// label16
			// 
			this.label16.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label16.ForeColor = System.Drawing.Color.Brown;
			this.label16.Location = new System.Drawing.Point(552, 136);
			this.label16.Name = "label16";
			this.label16.Size = new System.Drawing.Size(128, 16);
			this.label16.TabIndex = 237;
			this.label16.Text = "Registrar PROMOCION";
			// 
			// txtTiempo
			// 
			this.txtTiempo.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			appearance14.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtTiempo.Appearance = appearance14;
			this.txtTiempo.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtTiempo.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtTiempo.FormatString = "#0";
			this.txtTiempo.Location = new System.Drawing.Point(499, 73);
			this.txtTiempo.Name = "txtTiempo";
			this.txtTiempo.PromptChar = ' ';
			this.txtTiempo.Size = new System.Drawing.Size(40, 24);
			this.txtTiempo.TabIndex = 236;
			this.txtTiempo.Click += new System.EventHandler(this.txtTiempo_Click);
			this.txtTiempo.Enter += new System.EventHandler(this.txtTiempo_Enter);
			// 
			// label15
			// 
			this.label15.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label15.Location = new System.Drawing.Point(544, 56);
			this.label15.Name = "label15";
			this.label15.Size = new System.Drawing.Size(40, 16);
			this.label15.TabIndex = 235;
			this.label15.Text = "Unidad";
			// 
			// txtUnidadTaxi
			// 
			this.txtUnidadTaxi.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtUnidadTaxi.Location = new System.Drawing.Point(544, 73);
			this.txtUnidadTaxi.MaxLength = 9;
			this.txtUnidadTaxi.Name = "txtUnidadTaxi";
			this.txtUnidadTaxi.Size = new System.Drawing.Size(40, 20);
			this.txtUnidadTaxi.TabIndex = 234;
			this.txtUnidadTaxi.Text = "";
			this.txtUnidadTaxi.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
			// 
			// label14
			// 
			this.label14.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label14.Location = new System.Drawing.Point(499, 56);
			this.label14.Name = "label14";
			this.label14.Size = new System.Drawing.Size(40, 16);
			this.label14.TabIndex = 233;
			this.label14.Text = "Tmpo";
			// 
			// label13
			// 
			this.label13.AutoSize = true;
			this.label13.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label13.Location = new System.Drawing.Point(8, 456);
			this.label13.Name = "label13";
			this.label13.Size = new System.Drawing.Size(59, 17);
			this.label13.TabIndex = 231;
			this.label13.Text = "NORMAL";
			// 
			// label10
			// 
			this.label10.AutoSize = true;
			this.label10.Font = new System.Drawing.Font("Microsoft Sans Serif", 7F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label10.Location = new System.Drawing.Point(499, 40);
			this.label10.Name = "label10";
			this.label10.Size = new System.Drawing.Size(47, 14);
			this.label10.TabIndex = 230;
			this.label10.Text = "Fecha FC";
			// 
			// dtFecha
			// 
			appearance15.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtFecha.Appearance = appearance15;
			this.dtFecha.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton1.Caption = "Today";
			this.dtFecha.DateButtons.Add(dateButton1);
			this.dtFecha.Enabled = false;
			this.dtFecha.Format = "dd/MM/yyyy";
			this.dtFecha.Location = new System.Drawing.Point(499, 16);
			this.dtFecha.Name = "dtFecha";
			this.dtFecha.NonAutoSizeHeight = 23;
			this.dtFecha.Size = new System.Drawing.Size(104, 21);
			this.dtFecha.SpinButtonsVisible = true;
			this.dtFecha.TabIndex = 229;
			this.dtFecha.Value = ((object)(resources.GetObject("dtFecha.Value")));
			// 
			// label11
			// 
			this.label11.Location = new System.Drawing.Point(544, 175);
			this.label11.Name = "label11";
			this.label11.Size = new System.Drawing.Size(88, 16);
			this.label11.TabIndex = 228;
			this.label11.Text = "Forma de Pago";
			// 
			// cmbFormaPago
			// 
			appearance16.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbFormaPago.Appearance = appearance16;
			this.cmbFormaPago.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbFormaPago.Cursor = System.Windows.Forms.Cursors.Default;
			ultraGridColumn20.Header.VisiblePosition = 0;
			ultraGridColumn20.Hidden = true;
			ultraGridColumn21.Header.VisiblePosition = 1;
			ultraGridColumn21.Width = 152;
			ultraGridBand6.Columns.AddRange(new object[] {
																										 ultraGridColumn20,
																										 ultraGridColumn21});
			this.cmbFormaPago.DisplayLayout.BandsSerializer.Add(ultraGridBand6);
			this.cmbFormaPago.DisplayMember = "FormaPago";
			this.cmbFormaPago.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbFormaPago.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbFormaPago.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbFormaPago.Location = new System.Drawing.Point(544, 191);
			this.cmbFormaPago.Name = "cmbFormaPago";
			this.cmbFormaPago.Size = new System.Drawing.Size(144, 21);
			this.cmbFormaPago.TabIndex = 226;
			this.cmbFormaPago.ValueMember = "idFormaPago";
			// 
			// chkPrecios
			// 
			this.chkPrecios.Checked = true;
			this.chkPrecios.CheckState = System.Windows.Forms.CheckState.Checked;
			this.chkPrecios.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.chkPrecios.Location = new System.Drawing.Point(624, 16);
			this.chkPrecios.Name = "chkPrecios";
			this.chkPrecios.Size = new System.Drawing.Size(80, 24);
			this.chkPrecios.TabIndex = 218;
			this.chkPrecios.Text = "Precio N/R";
			this.chkPrecios.Visible = false;
			// 
			// txtTotal
			// 
			this.txtTotal.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			appearance17.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtTotal.Appearance = appearance17;
			this.txtTotal.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtTotal.Enabled = false;
			this.txtTotal.Font = new System.Drawing.Font("Microsoft Sans Serif", 14F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtTotal.FormatString = "#,##0.00";
			this.txtTotal.Location = new System.Drawing.Point(608, 505);
			this.txtTotal.Name = "txtTotal";
			this.txtTotal.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtTotal.PromptChar = ' ';
			this.txtTotal.Size = new System.Drawing.Size(80, 31);
			this.txtTotal.TabIndex = 217;
			// 
			// txtIva
			// 
			this.txtIva.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			appearance18.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtIva.Appearance = appearance18;
			this.txtIva.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtIva.Enabled = false;
			this.txtIva.FormatString = "#,##0.00";
			this.txtIva.Location = new System.Drawing.Point(608, 476);
			this.txtIva.Name = "txtIva";
			this.txtIva.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtIva.PromptChar = ' ';
			this.txtIva.Size = new System.Drawing.Size(80, 21);
			this.txtIva.TabIndex = 216;
			// 
			// txtSubTotal
			// 
			this.txtSubTotal.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			appearance19.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtSubTotal.Appearance = appearance19;
			this.txtSubTotal.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtSubTotal.Enabled = false;
			this.txtSubTotal.FormatString = "#,##0.00";
			this.txtSubTotal.Location = new System.Drawing.Point(608, 432);
			this.txtSubTotal.Name = "txtSubTotal";
			this.txtSubTotal.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtSubTotal.PromptChar = ' ';
			this.txtSubTotal.Size = new System.Drawing.Size(80, 21);
			this.txtSubTotal.TabIndex = 215;
			// 
			// btnImprimir
			// 
			this.btnImprimir.Image = ((System.Drawing.Image)(resources.GetObject("btnImprimir.Image")));
			this.btnImprimir.Location = new System.Drawing.Point(107, 480);
			this.btnImprimir.Name = "btnImprimir";
			this.btnImprimir.Size = new System.Drawing.Size(72, 40);
			this.btnImprimir.TabIndex = 37;
			this.toolTip1.SetToolTip(this.btnImprimir, "Imprime Factura");
			this.btnImprimir.Visible = false;
			this.btnImprimir.Click += new System.EventHandler(this.btnImprimir_Click);
			// 
			// bntGrabar
			// 
			this.bntGrabar.Image = ((System.Drawing.Image)(resources.GetObject("bntGrabar.Image")));
			this.bntGrabar.Location = new System.Drawing.Point(32, 480);
			this.bntGrabar.Name = "bntGrabar";
			this.bntGrabar.Size = new System.Drawing.Size(75, 40);
			this.bntGrabar.TabIndex = 35;
			this.toolTip1.SetToolTip(this.bntGrabar, "Grabar FACTURA");
			this.bntGrabar.Click += new System.EventHandler(this.bntGrabar_Click);
			// 
			// label8
			// 
			this.label8.Font = new System.Drawing.Font("Microsoft Sans Serif", 14F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label8.Location = new System.Drawing.Point(520, 512);
			this.label8.Name = "label8";
			this.label8.Size = new System.Drawing.Size(80, 16);
			this.label8.TabIndex = 33;
			this.label8.Text = "TOTAL";
			this.label8.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
			// 
			// label7
			// 
			this.label7.Location = new System.Drawing.Point(520, 478);
			this.label7.Name = "label7";
			this.label7.Size = new System.Drawing.Size(80, 16);
			this.label7.TabIndex = 31;
			this.label7.Text = "I.V.A.";
			this.label7.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
			// 
			// label6
			// 
			this.label6.Location = new System.Drawing.Point(520, 434);
			this.label6.Name = "label6";
			this.label6.Size = new System.Drawing.Size(80, 16);
			this.label6.TabIndex = 29;
			this.label6.Text = "SUB TOTAL";
			this.label6.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
			// 
			// ugDetalleFactura
			// 
			this.ugDetalleFactura.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.ugDetalleFactura.Cursor = System.Windows.Forms.Cursors.Default;
			this.ugDetalleFactura.DataSource = this.ultraDataSource4;
			appearance20.BackColor = System.Drawing.Color.White;
			this.ugDetalleFactura.DisplayLayout.Appearance = appearance20;
			ultraGridColumn22.Header.VisiblePosition = 0;
			ultraGridColumn22.Hidden = true;
			ultraGridColumn23.Header.VisiblePosition = 1;
			ultraGridColumn23.Hidden = true;
			ultraGridColumn24.Header.VisiblePosition = 2;
			ultraGridColumn24.Hidden = true;
			ultraGridColumn25.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
			ultraGridColumn25.Header.VisiblePosition = 3;
			ultraGridColumn25.Width = 365;
			appearance21.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn26.CellAppearance = appearance21;
			ultraGridColumn26.Format = "#,##0";
			ultraGridColumn26.Header.VisiblePosition = 4;
			ultraGridColumn26.PromptChar = ' ';
			ultraGridColumn26.Width = 50;
			ultraGridColumn27.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
			appearance22.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn27.CellAppearance = appearance22;
			ultraGridColumn27.Format = "#,##0.000000";
			ultraGridColumn27.Header.VisiblePosition = 5;
			ultraGridColumn27.PromptChar = ' ';
			ultraGridColumn27.Width = 96;
			ultraGridColumn28.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance23.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn28.CellAppearance = appearance23;
			ultraGridColumn28.Format = "#,##0.00";
			ultraGridColumn28.Header.Caption = "%Dscto";
			ultraGridColumn28.Header.VisiblePosition = 6;
			ultraGridColumn28.Width = 40;
			ultraGridColumn29.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
			appearance24.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn29.CellAppearance = appearance24;
			ultraGridColumn29.Format = "#,##0.000000";
			ultraGridColumn29.Header.Caption = "Subtotal";
			ultraGridColumn29.Header.VisiblePosition = 7;
			ultraGridColumn29.PromptChar = ' ';
			ultraGridColumn29.Width = 96;
			ultraGridBand7.Columns.AddRange(new object[] {
																										 ultraGridColumn22,
																										 ultraGridColumn23,
																										 ultraGridColumn24,
																										 ultraGridColumn25,
																										 ultraGridColumn26,
																										 ultraGridColumn27,
																										 ultraGridColumn28,
																										 ultraGridColumn29});
			ultraGridBand7.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.Edit;
			this.ugDetalleFactura.DisplayLayout.BandsSerializer.Add(ultraGridBand7);
			appearance25.ForeColor = System.Drawing.Color.Black;
			appearance25.ForeColorDisabled = System.Drawing.Color.Black;
			this.ugDetalleFactura.DisplayLayout.Override.ActiveRowAppearance = appearance25;
			this.ugDetalleFactura.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.TemplateOnBottom;
			this.ugDetalleFactura.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.True;
			this.ugDetalleFactura.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance26.BackColor = System.Drawing.Color.Transparent;
			this.ugDetalleFactura.DisplayLayout.Override.CardAreaAppearance = appearance26;
			this.ugDetalleFactura.DisplayLayout.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.Edit;
			appearance27.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance27.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance27.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance27.FontData.BoldAsString = "True";
			appearance27.FontData.Name = "Arial";
			appearance27.FontData.SizeInPoints = 10F;
			appearance27.ForeColor = System.Drawing.Color.White;
			appearance27.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.ugDetalleFactura.DisplayLayout.Override.HeaderAppearance = appearance27;
			appearance28.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance28.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance28.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ugDetalleFactura.DisplayLayout.Override.RowAlternateAppearance = appearance28;
			appearance29.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance29.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance29.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ugDetalleFactura.DisplayLayout.Override.RowSelectorAppearance = appearance29;
			appearance30.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance30.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance30.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ugDetalleFactura.DisplayLayout.Override.SelectedRowAppearance = appearance30;
			this.ugDetalleFactura.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.ugDetalleFactura.Location = new System.Drawing.Point(8, 248);
			this.ugDetalleFactura.Name = "ugDetalleFactura";
			this.ugDetalleFactura.Size = new System.Drawing.Size(696, 176);
			this.ugDetalleFactura.TabIndex = 28;
			this.ugDetalleFactura.AfterRowsDeleted += new System.EventHandler(this.ugDetalleFactura_AfterRowsDeleted);
			this.ugDetalleFactura.BeforeRowsDeleted += new Infragistics.Win.UltraWinGrid.BeforeRowsDeletedEventHandler(this.ugDetalleFactura_BeforeRowsDeleted);
			this.ugDetalleFactura.KeyDown += new System.Windows.Forms.KeyEventHandler(this.ugDetalleFactura_KeyDown);
			this.ugDetalleFactura.AfterCellUpdate += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.ugDetalleFactura_AfterCellUpdate);
			this.ugDetalleFactura.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.ugDetalleFactura_InitializeLayout);
			// 
			// ultraDataSource4
			// 
			ultraDataColumn14.DataType = typeof(int);
			ultraDataColumn15.DataType = typeof(int);
			ultraDataColumn16.DataType = typeof(int);
			ultraDataColumn18.DataType = typeof(int);
			ultraDataColumn19.DataType = typeof(System.Decimal);
			ultraDataColumn20.DataType = typeof(System.Decimal);
			ultraDataColumn21.DataType = typeof(System.Decimal);
			this.ultraDataSource4.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn14,
																																 ultraDataColumn15,
																																 ultraDataColumn16,
																																 ultraDataColumn17,
																																 ultraDataColumn18,
																																 ultraDataColumn19,
																																 ultraDataColumn20,
																																 ultraDataColumn21});
			// 
			// txtTelefono
			// 
			this.txtTelefono.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtTelefono.Location = new System.Drawing.Point(440, 192);
			this.txtTelefono.Name = "txtTelefono";
			this.txtTelefono.TabIndex = 27;
			this.txtTelefono.Text = "";
			// 
			// label5
			// 
			this.label5.Location = new System.Drawing.Point(440, 175);
			this.label5.Name = "label5";
			this.label5.Size = new System.Drawing.Size(80, 16);
			this.label5.TabIndex = 26;
			this.label5.Text = "# TELEFONO";
			// 
			// txtDireccion
			// 
			this.txtDireccion.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtDireccion.Location = new System.Drawing.Point(8, 192);
			this.txtDireccion.Name = "txtDireccion";
			this.txtDireccion.Size = new System.Drawing.Size(424, 20);
			this.txtDireccion.TabIndex = 25;
			this.txtDireccion.Text = "";
			// 
			// label4
			// 
			this.label4.Location = new System.Drawing.Point(8, 175);
			this.label4.Name = "label4";
			this.label4.Size = new System.Drawing.Size(80, 16);
			this.label4.TabIndex = 24;
			this.label4.Text = "DIRECCION";
			// 
			// txtNombre
			// 
			this.txtNombre.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtNombre.Location = new System.Drawing.Point(248, 152);
			this.txtNombre.Name = "txtNombre";
			this.txtNombre.Size = new System.Drawing.Size(296, 20);
			this.txtNombre.TabIndex = 23;
			this.txtNombre.Text = "";
			// 
			// label3
			// 
			this.label3.Location = new System.Drawing.Point(248, 136);
			this.label3.Name = "label3";
			this.label3.Size = new System.Drawing.Size(64, 16);
			this.label3.TabIndex = 22;
			this.label3.Text = "NOMBRES";
			// 
			// txtNumCedula
			// 
			this.txtNumCedula.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtNumCedula.Location = new System.Drawing.Point(144, 152);
			this.txtNumCedula.MaxLength = 13;
			this.txtNumCedula.Name = "txtNumCedula";
			this.txtNumCedula.TabIndex = 21;
			this.txtNumCedula.Text = "";
			this.txtNumCedula.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
			this.txtNumCedula.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtNumCedula_KeyDown);
			this.txtNumCedula.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtNumCedula_KeyPress);
			this.txtNumCedula.Validating += new System.ComponentModel.CancelEventHandler(this.txtNumCedula_Validating);
			this.txtNumCedula.Validated += new System.EventHandler(this.txtNumCedula_Validated);
			// 
			// label2
			// 
			this.label2.Location = new System.Drawing.Point(144, 136);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(64, 16);
			this.label2.TabIndex = 20;
			this.label2.Text = "# CEDULA";
			// 
			// button1
			// 
			this.button1.Image = ((System.Drawing.Image)(resources.GetObject("button1.Image")));
			this.button1.Location = new System.Drawing.Point(680, 96);
			this.button1.Name = "button1";
			this.button1.Size = new System.Drawing.Size(24, 23);
			this.button1.TabIndex = 19;
			this.toolTip1.SetToolTip(this.button1, "Registrar Número de Factura");
			this.button1.Click += new System.EventHandler(this.button1_Click);
			// 
			// txtNumFactura
			// 
			this.txtNumFactura.Location = new System.Drawing.Point(588, 96);
			this.txtNumFactura.MaxLength = 9;
			this.txtNumFactura.Name = "txtNumFactura";
			this.txtNumFactura.Size = new System.Drawing.Size(88, 20);
			this.txtNumFactura.TabIndex = 18;
			this.txtNumFactura.Text = "";
			this.txtNumFactura.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
			this.txtNumFactura.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtNumFactura_KeyPress);
			// 
			// label1
			// 
			this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label1.Location = new System.Drawing.Point(512, 98);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(72, 16);
			this.label1.TabIndex = 17;
			this.label1.Text = "# FACTURA";
			// 
			// ugFacRelizadas
			// 
			this.ugFacRelizadas.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.ugFacRelizadas.Cursor = System.Windows.Forms.Cursors.Default;
			this.ugFacRelizadas.DataSource = this.ultraDataSource2;
			appearance31.BackColor = System.Drawing.Color.White;
			this.ugFacRelizadas.DisplayLayout.Appearance = appearance31;
			ultraGridColumn30.Header.Caption = "id";
			ultraGridColumn30.Header.VisiblePosition = 0;
			ultraGridColumn30.Hidden = true;
			ultraGridColumn31.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn31.Header.Caption = "Num. Documento";
			ultraGridColumn31.Header.VisiblePosition = 1;
			ultraGridColumn31.Width = 210;
			ultraGridColumn32.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn32.Header.Caption = "# Factura";
			ultraGridColumn32.Header.VisiblePosition = 2;
			ultraGridColumn32.Width = 140;
			ultraGridColumn33.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn33.Header.VisiblePosition = 3;
			ultraGridColumn33.Width = 100;
			ultraGridColumn34.Header.VisiblePosition = 4;
			ultraGridColumn34.Hidden = true;
			ultraGridColumn35.Header.VisiblePosition = 5;
			ultraGridColumn35.Hidden = true;
			ultraGridBand8.Columns.AddRange(new object[] {
																										 ultraGridColumn30,
																										 ultraGridColumn31,
																										 ultraGridColumn32,
																										 ultraGridColumn33,
																										 ultraGridColumn34,
																										 ultraGridColumn35});
			ultraGridBand8.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.Edit;
			this.ugFacRelizadas.DisplayLayout.BandsSerializer.Add(ultraGridBand8);
			appearance32.ForeColor = System.Drawing.Color.Black;
			appearance32.ForeColorDisabled = System.Drawing.Color.Black;
			this.ugFacRelizadas.DisplayLayout.Override.ActiveRowAppearance = appearance32;
			this.ugFacRelizadas.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.TemplateOnBottom;
			this.ugFacRelizadas.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.True;
			this.ugFacRelizadas.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance33.BackColor = System.Drawing.Color.Transparent;
			this.ugFacRelizadas.DisplayLayout.Override.CardAreaAppearance = appearance33;
			this.ugFacRelizadas.DisplayLayout.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.RowSelect;
			appearance34.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance34.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance34.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance34.FontData.BoldAsString = "True";
			appearance34.FontData.Name = "Arial";
			appearance34.FontData.SizeInPoints = 10F;
			appearance34.ForeColor = System.Drawing.Color.White;
			appearance34.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.ugFacRelizadas.DisplayLayout.Override.HeaderAppearance = appearance34;
			appearance35.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance35.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance35.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ugFacRelizadas.DisplayLayout.Override.RowAlternateAppearance = appearance35;
			appearance36.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance36.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance36.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ugFacRelizadas.DisplayLayout.Override.RowSelectorAppearance = appearance36;
			appearance37.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance37.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance37.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ugFacRelizadas.DisplayLayout.Override.SelectedRowAppearance = appearance37;
			this.ugFacRelizadas.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.ugFacRelizadas.Location = new System.Drawing.Point(8, 16);
			this.ugFacRelizadas.Name = "ugFacRelizadas";
			this.ugFacRelizadas.Size = new System.Drawing.Size(488, 104);
			this.ugFacRelizadas.TabIndex = 16;
			this.ugFacRelizadas.DoubleClick += new System.EventHandler(this.ugFacRelizadas_DoubleClick);
			// 
			// ultraDataSource2
			// 
			ultraDataColumn22.DataType = typeof(int);
			ultraDataColumn25.DataType = typeof(System.DateTime);
			ultraDataColumn26.DataType = typeof(int);
			ultraDataColumn27.DataType = typeof(int);
			this.ultraDataSource2.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn22,
																																 ultraDataColumn23,
																																 ultraDataColumn24,
																																 ultraDataColumn25,
																																 ultraDataColumn26,
																																 ultraDataColumn27});
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
			// frmFCHabitaciones
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(1272, 557);
			this.Controls.Add(this.groupBox2);
			this.Controls.Add(this.groupBox1);
			this.Name = "frmFCHabitaciones";
			this.Text = "Facuración HABITACIONES";
			this.Closing += new System.ComponentModel.CancelEventHandler(this.frmFCHabitaciones_Closing);
			this.Load += new System.EventHandler(this.frmFCHabitaciones_Load);
			this.groupBox1.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.ugArticulos)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource3)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).EndInit();
			this.groupBox2.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.txtDescnto)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbTipoRuc)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCajero)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource5)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCobrador)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbVendedor)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtTiempo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFecha)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbFormaPago)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtTotal)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIva)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtSubTotal)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ugDetalleFactura)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource4)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ugFacRelizadas)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteo)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void frmFCHabitaciones_Load(object sender, System.EventArgs e)
		{
			/* Porcentaje del IVA */
			//string lcPiva = "SELECT Iva FROM SeteoF WHERE idSeteoF=1";
			string lcPiva = "Select dbo.RetornaIVAVigenteLote()";
			PorcentajeIva = Funcion.decEscalarSQL(cdsSeteo, lcPiva); 

			/* Tiempo de habitación */
			string nHoras = "SELECT dbo.Tmpo_Habitaciones()";
			this.txtTiempo.Value = Funcion.iEscalarSQL(cdsSeteo, nHoras); 
			
			//Buscar habitaciones con sus diferentes estados
			string stTipo = "0703";
			string stPermiso = "118";

			//Carga Vendedores, Cajero y Cobrador
			this.cmbVendedor.DataSource = Funcion.dvProcedimiento(cdsSeteo, string.Format("Exec PersonalLotte {0}", NBodga));
			//this.cmbVendedor.DataSource = Funcion.dvProcedimiento(cdsSeteo, string.Format("Exec PersnlLotteTrnoGuardia"));
			this.cmbCobrador.DataSource = Funcion.dvProcedimiento(cdsSeteo, string.Format("Exec PersonalLotte_Cobrador {0}", NBodga));
			//this.cmbCobrador.DataSource = Funcion.dvProcedimiento(cdsSeteo, string.Format("Exec PersnlLotteTrnoCobrador"));
			this.cmbCajero.DataSource = Funcion.dvProcedimiento(cdsSeteo, string.Format("Exec PersonalLotte_Cajero {0}", NBodga));
			//this.cmbCajero.DataSource = Funcion.dvProcedimiento(cdsSeteo, string.Format("Exec PersnlLotteTrnoCajero"));
			this.cmbTipoRuc.DataSource = Funcion.dvProcedimiento(cdsSeteo, "Exec ClienteTipoRucCarga 1");
			this.cmbTipoRuc.Value = 5;

			miAcceso = new Acceso(cdsSeteo, stTipo);
			
			this.cmbFormaPago.Enabled = true;
			this.btnImprimir.Enabled = true;
			this.bntGrabar.Enabled = true;

			this.txtnPromo.Text = "";

			string lcQryBod = "SELECT Bodega From Bodega Where idBodega in (Select idSucursal From Habitaciones Where idHabitacion="+idNhab.ToString()+")";
			int laBodega = Funcion.iEscalarSQL(cdsSeteo, lcQryBod);

			NBodga = laBodega;

			string lcQry = "SELECT IsNull(Ocupado,0) as Ocupado FROM Habitaciones WHERE idHabitacion="+idNhab.ToString();
			int EstadoHabitacion = Funcion.iEscalarSQL(cdsSeteo, lcQry);
			if(EstadoHabitacion!=1) 
			{
				Registra_FC(idNhab);
				// Visualiza información de factura
				int idCpa = 0;
				int idClte = 0;

				foreach(Infragistics.Win.UltraWinGrid.UltraGridRow drFC in this.ugFacRelizadas.Rows.All)
				{	
					idCpa = (int)drFC.Cells["idCompra"].Value;
				}

				/* Campo numero2 utilizado para indicar que habitacion aplica promoción o no */
				string lQry = "SELECT idCompra,idCliente,Numero,IsNull(Numero2,''),idTipoMoneda FROM Compra WHERE Estado not in (6,10) and idCompra="+idCpa.ToString()+" and idRecibe="+idNhab.ToString();
				SqlDataReader dr = Funcion.rEscalarSQL(cdsSeteo, lQry, true);
				dr.Read();

				if(dr.HasRows)
				{
					idCpa = dr.GetInt32(0);
					idCpaFC = dr.GetInt32(0);
					idClte = dr.GetInt32(1);
					idCpnWeb = dr.GetInt32(4);
					NumFactura = dr.GetString(2).Trim();
					aplicapromocion = dr.GetString(3).Trim();
				}
				dr.Close();
				// Visualiza información del Cliente
				if(idClte>0)
				{
					string l2Qry = "SELECT Ruc,Codigo,Nombre,Direccion,Telefono From Cliente Where idCliente="+idClte.ToString();
					SqlDataReader drCl = Funcion.rEscalarSQL(cdsSeteo, l2Qry, true);
					drCl.Read();
					if(drCl.HasRows)
					{
						this.txtNumCedula.Text = drCl.GetString(0);
						this.txtNombre.Text = drCl.GetString(2);
						this.txtDireccion.Text = drCl.GetString(3);
						this.txtTelefono.Text = drCl.GetString(4);
					}
				}
				// Visualiza información detalle compra
				string l3Qry = "SELECT Count(*) as CanT FROM DetCompra WHERE idCompra="+idCpa.ToString();
				int CantDepCpa = Funcion.iEscalarSQL(cdsSeteo, l3Qry);
				if(CantDepCpa>=1)
				{
					string sSQL = string.Format("Select idDetCompra,idCompra,idArticulo,Articulo,Cantidad,Precio,DescuentoPorc as PDscto,SubTotal From DetCompra Where idCompra="+idCpa.ToString());
					ugDetalleFactura.DataSource = Funcion.dvProcedimiento(cdsSeteo, sSQL);
				}
				/*Visualiza cupon*/
				string lCpnQry = "SELECT Cupon From CuponesWeb Where idCuponesWeb="+idCpnWeb.ToString();
				this.txtnPromo.Text = Funcion.sEscalarSQL(cdsSeteo, lCpnQry);

				Total();
			}
			this.cmbFormaPago.DataSource = Funcion.dvProcedimiento(cdsSeteo, "Exec CompraFPagoCarga");
		}

		private void cdsSeteo_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteo.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);
		}

		private void button2_Click(object sender, System.EventArgs e)
		{
			// Cargar Datos de bebidas
			//string sSQL = string.Format("Select idSubGrupo,idGrupoArticulo,SubGrupo From ArticuloSubGrupo Where idGrupoArticulo=40");
			//ugServicios.DataSource = Funcion.dvProcedimiento(cdsSeteo, sSQL);
		}

		private void button3_Click(object sender, System.EventArgs e)
		{
			// Cargar Datos de bebidas
			//string sSQL = string.Format("Select idSubGrupo,idGrupoArticulo,SubGrupo From ArticuloSubGrupo Where idGrupoArticulo=41");
			//ugServicios.DataSource = Funcion.dvProcedimiento(cdsSeteo, sSQL);
		}

		private void ugServicios_Click(object sender, System.EventArgs e)
		{
			// Carga articulos de acuerso al subgrupo
			//int nSubgrp = (int)ugServicios.ActiveRow.Cells["idSubGrupo"].Value;
			//string sSQL = string.Format("Select idArticulo,idSubGrupo,idGrupoArticulo,Articulo,Precio1,Precio2 From Articulo Where idSubGrupo="+nSubgrp.ToString());
			//ugArticulos.DataSource = Funcion.dvProcedimiento(cdsSeteo, sSQL);
		}

		private void ugServicios_BeforeSelectChange(object sender, Infragistics.Win.UltraWinGrid.BeforeSelectChangeEventArgs e)
		{
			// Carga articulos de acuerso al subgrupo
			//int nSubgrp = (int)ugServicios.ActiveRow.Cells["idSubGrupo"].Value;
			//string sSQL = string.Format("Select idArticulo,idSubGrupo,idGrupoArticulo,Articulo,Precio1,Precio2 From Articulo Where idSubGrupo="+nSubgrp.ToString());
			//ugArticulos.DataSource = Funcion.dvProcedimiento(cdsSeteo, sSQL);
		}

		private void button1_Click(object sender, System.EventArgs e)
		{
			// Registrar factura para habitación
			try
			{
				//Validar Numero de Factura
				if(this.cmbVendedor.ActiveRow == null)
				{
					MessageBox.Show("No se a seleccionado Guardia?","Aviso....!!!", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					this.cmbVendedor.Focus();
				}
				else 
				{
					/*Validar si factura no se encuentra repetida*/
					int TmpoHab = (int)this.txtTiempo.Value;
					string lcQry = "SELECT IsNull(Numero,'') as NumFac FROM Compra WHERE numero like "+"'%"+this.txtNumFactura.Text.ToString()+"'"+" And idTipoFactura=1 And Bodega="+NBodga.ToString();
					string nFacHab = Funcion.sEscalarSQL(cdsSeteo, lcQry);
					int valfc = 0;
					if (nFacHab!=null) valfc = Int32.Parse(nFacHab);
					int vlrfc = Int32.Parse(this.txtNumFactura.Text.Trim());
					int idGuadr = (int)this.cmbVendedor.Value;
					if (valfc!=vlrfc)
					{
						/*Verificar Promoción HABITACION*/
						string sSQLPro = string.Format("Exec Busca_Promociones {0}", idNhab);
						SqlDataReader drPromocion = Funcion.rEscalarSQL(cdsSeteo, sSQLPro, true);
						int idPromc = 0;
						int nHabPromocion = 0;
						int nDiaPromocion = 0 ;
						int TipoVlrPromocion = 0;
						int CantHorasPromocion = 0;
						decimal ValorPromocion = 0.00m;
						int SwProm = 0;

						while (drPromocion.Read())
						{
							idPromc = drPromocion.GetInt32(0);
							nHabPromocion = drPromocion.GetInt32(1);
							nDiaPromocion = drPromocion.GetInt32(2);
							TipoVlrPromocion = drPromocion.GetInt32(3);
							CantHorasPromocion = drPromocion.GetInt32(4);
							ValorPromocion = drPromocion.GetDecimal(5);
						}
						drPromocion.Close();

						//this.txtDescnto.Value = ValorPromocion;

						if (idPromc>0 && aplicapromocion.Equals("SI")) 
						{
							SwProm=1;
							if (nHabPromocion>0 || aplicapromocion.Equals("SI")) 
							{
								SwProm=1;
								if (nDiaPromocion>0 || aplicapromocion.Equals("SI")) SwProm=1;
								else SwProm=0;
							} else SwProm=0;
						}
						else SwProm=0;

						/* Verificar si habitacion tiene RESERVA */
						string lcReser = "Exec Numero_Reserva "+idNhab.ToString();
						int Resrv = Funcion.iEscalarSQL(cdsSeteo, lcReser); 

						string SP_Qry = "";

						if (Resrv>0) 
						{
							DialogResult resultado = MessageBox.Show("La Habitación Tiene una RESERVA..\nDesea aplicar RESERVA a la Habitación.?","Aviso...!!!",MessageBoxButtons.YesNo);
							if (DialogResult.Yes == resultado)
								SP_Qry = string.Format("Exec FacturaLotteServicio '{0}',{1},{2},{3},{4},{5},{6},{7}", this.txtNumFactura.Text,NumrHab,idNhab,1,Resrv,TmpoHab,idGuadr,0);
							else SP_Qry = string.Format("Exec FacturaLotteServicio '{0}',{1},{2},{3},{4},{5},{6},{7}", this.txtNumFactura.Text,NumrHab,idNhab,0,0,TmpoHab,idGuadr,0);
						}
						else 
						{
							if (SwProm==1)
							{
								DialogResult resultado = MessageBox.Show("Desea Aplicar PROMOCION.?","Aviso...!!!",MessageBoxButtons.YesNo);
								if (DialogResult.Yes == resultado)
								{
									this.txtTiempo.Value = CantHorasPromocion;
									TmpoHab = CantHorasPromocion;
									SP_Qry = string.Format("Exec FactLSer_Promocion '{0}',{1},{2},{3},{4},{5},{6},{7},{8},'{9}',{10}", 
										this.txtNumFactura.Text,NumrHab,idNhab,0,0,TmpoHab,idGuadr,0,idPromc,aplicapromocion,this.txtnPromo.Text.ToString().Trim());
									aplicapromocion = "SI";
								}
								else
								{
									SP_Qry = string.Format("Exec FacturaLotteServicio '{0}',{1},{2},{3},{4},{5},{6},{7}", this.txtNumFactura.Text,NumrHab,idNhab,0,0,TmpoHab,idGuadr,0);
									ValorPromocion = 0.00m;
									aplicapromocion = "NO";
								}
							}
							else
							{
								SP_Qry = string.Format("Exec FacturaLotteServicio '{0}',{1},{2},{3},{4},{5},{6},{7}", this.txtNumFactura.Text,NumrHab,idNhab,0,0,TmpoHab,idGuadr,0);
								ValorPromocion = 0.00m;
								aplicapromocion = "NO";
							}
						}
						//Funcion.EjecutaSQL(cdsSeteo,SP_Qry);
						idCpaFC = (int)Funcion.iEscalarSQL(cdsSeteo,SP_Qry);
						Registra_FC(idNhab);
						RegistroFactura = 1;
						//Visualiza detalle factura
						string QryDetFc = "SELECT idDetCompra,idCompra,idArticulo,Articulo,Cantidad,Precio,DescuentoPorc as PDscto,SubTotal From DetCompra Where idCompra="+idCpaFC.ToString();
						ugDetalleFactura.DataSource = Funcion.dvProcedimiento(cdsSeteo, QryDetFc);
						Total();
					}
					else
					{
						MessageBox.Show("Factura Registrada en Sistema.?\nPorfavor registre otro número de FACTURA....","Aviso...!!!",MessageBoxButtons.OK);
						return;
					}
				}
			}
			catch (Exception ex2) 
			{
				MessageBox.Show(string.Format("Error: {0} {1}", ex2.GetType(), ex2.Message), "Aviso...!!!", MessageBoxButtons.OK, MessageBoxIcon.Error);
			}
		}

		private void txtNumFactura_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			Funcion.CamposNumericos(sender,e);
		}

		public void Registra_FC(int NumrHab)
		{
			string Qry = string.Format("Exec RegistrodeFacturas {0}", NumrHab);
			ugFacRelizadas.DataSource = Funcion.dvProcedimiento(cdsSeteo,Qry);
		}

		private void ugArticulos_BeforeSelectChange(object sender, Infragistics.Win.UltraWinGrid.BeforeSelectChangeEventArgs e)
		{
			// Carga articulos detalle factura
			//			int swidFC = (int) ugFacRelizadas.ActiveRow.Cells["idCompra"].Value;
			//			if(swidFC>0)
			//			{
			//				int nSubgrp = (int)ugServicios.ActiveRow.Cells["idSubGrupo"].Value;
			//				string sSQL = string.Format("Select idArticulo,idSubGrupo,idGrupoArticulo,Articulo,Precio1,Precio2 From Articulo Where idSubGrupo="+nSubgrp.ToString());
			//				ugArticulos.DataSource = Funcion.dvProcedimiento(cdsSeteo, sSQL);
			//			}
			//			else
			//			{
			//				MessageBox.Show("Registre PRIMERO la FACTURA antes de Registrar Productos...!!!","Aviso...!!!",MessageBoxButtons.OK,MessageBoxIcon.Error);
			//			}
		}

		private void ugArticulos_DoubleClick(object sender, System.EventArgs e)
		{
			//			// Registrar articulo en detalle de factura
			//			try
			//			{
			//				idCpaFC = (int)ugFacRelizadas.ActiveRow.Cells["idCompra"].Value;
			//				int idCltRuc = 0;
			//				if (this.txtNumCedula.Text.Length!=0) 
			//				{
			//					string lcQry = "SELECT idCliente FROM Cliente WHERE Ruc="+this.txtNumCedula.Text.ToString();
			//					idCltRuc = Funcion.iEscalarSQL(cdsSeteo, lcQry); 
			//				}
			//				int idClFc = 0;
			//				if(idCltRuc>0) idClFc = idCltRuc;
			//				int idArtFc = (int)ugArticulos.ActiveRow.Cells["idArticulo"].Value;
			//				decimal preFc = 0;
			//				if(this.chkPrecios.Checked==true)	preFc = (decimal)ugArticulos.ActiveRow.Cells["Precio2"].Value;
			//				else preFc = (decimal)ugArticulos.ActiveRow.Cells["Precio1"].Value;
			//				string SP_Insert = string.Format("Exec GrabaDetalle_FC_Habitacion {0},{1},{2},{3},{4}", idCpaFC,idClFc,idArtFc,1,preFc);
			//				Funcion.EjecutaSQL(cdsSeteo,SP_Insert);
			////				string QryDetFc = "SELECT idDetCompra,idCompra,idArticulo,Articulo,Cantidad,Precio,SubTotal From DetCompra Where idCompra="+idCpaFC.ToString();
			////				ugDetalleFactura.DataSource = Funcion.dvProcedimiento(cdsSeteo, QryDetFc);
			//				Total();
			//			}
			//			catch (Exception ex2) 
			//			{
			//				MessageBox.Show(string.Format("Error: {0} {1}", ex2.GetType(), ex2.Message), "Aviso...!!!", MessageBoxButtons.OK, MessageBoxIcon.Error);
			//			}
		}

		public void BuscaPromocionArticulo(int idArclo)
		{
			/*Verificar si producto se encuentra con promoción*/
			string sSQL = string.Format("Exec BuscaPromoArticulo '{0}'", idArclo);
			SqlDataReader dr = Funcion.rEscalarSQL(cdsSeteo, sSQL, true);
			dr.Read();
			if(dr.HasRows)
			{
				respuestapromocion = dr.GetString(0).Trim();
				tpopromocion = dr.GetInt32(1);
				valorpromo = dr.GetDecimal(2);
			}
			dr.Close();
		}

		private void Total()
		{
			#region Creacion de Variables					
			decimal dDescuentoLineal = 0.00m;
			decimal dIvaLineal = 0.00m;

			#region Calculo Forma Pago Mixto 
			decimal dSubTotal1 = 0.00m;
			decimal dSubTotal2 = 0.00m;
			decimal dSubTotal3 = 0.00m;
			decimal dSubTotal4 = 0.00m;
			#endregion Calculo Forma Pago Mixto 

			#endregion Creacion de Variables
			
			this.txtIva.Value = 0.000000m;
			this.txtSubTotal.Value = 0.000000m;
			this.txtTotal.Value = 0.00m;
			decimal ValDsco = 0.00m;
			decimal ToDscto = 0.00m;

			#region Desplazamiento de grilla
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in ugDetalleFactura.Rows.All)
			{
				if (dr.Cells["Cantidad"].Value != DBNull.Value && dr.Cells["Precio"].Value != DBNull.Value)
				{
					BuscaPromocionArticulo(Convert.ToInt32(dr.Cells["idArticulo"].Value));
					if (respuestapromocion.Equals("SI"))
					{
						if (tpopromocion == 1)
						{
							dr.Cells["Precio"].Value = valorpromo;
							dr.Cells["Subtotal"].Value = Math.Round(Convert.ToDecimal(dr.Cells["Cantidad"].Value) * Convert.ToDecimal(dr.Cells["Precio"].Value), 6);
						}
						else 
						{
							/*Calcular descuento promocion*/
							if (Convert.ToDecimal(dr.Cells["PDscto"].Value)>0)
							{
								dr.Cells["Subtotal"].Value = Math.Round(Convert.ToDecimal(dr.Cells["Cantidad"].Value) *(Convert.ToDecimal(dr.Cells["Precio"].Value)*(Convert.ToDecimal(dr.Cells["PDscto"].Value)/100)), 6);
								ToDscto = ToDscto + Convert.ToDecimal(dr.Cells["Precio"].Value)*(Convert.ToDecimal(dr.Cells["PDscto"].Value)/100);
							}
							else dr.Cells["Subtotal"].Value = Math.Round(Convert.ToDecimal(dr.Cells["Cantidad"].Value) * Convert.ToDecimal(dr.Cells["Precio"].Value), 6);
						}						
					}
					else dr.Cells["Subtotal"].Value = Math.Round(Convert.ToDecimal(dr.Cells["Cantidad"].Value) * Convert.ToDecimal(dr.Cells["Precio"].Value), 6);
					//Actualiza la cantidad de producto
					string Updt = string.Format("Exec Actualiza_DetCpa_Lotte {0},{1}", dr.Cells["idDetCompra"].Value.ToString(),dr.Cells["Cantidad"].Value.ToString());
					Funcion.EjecutaSQL(cdsSeteo,Updt);
					this.txtSubTotal.Value = (decimal)this.txtSubTotal.Value + Math.Round(Convert.ToDecimal(dr.Cells["SubTotal"].Value), 6);
				}
			}

			//Se actualiza valor de descuento en valor de habitación y no total de factura
			//ValDsco = Convert.ToDecimal(this.txtSubTotal.Value)-Convert.ToDecimal(this.txtDescnto.Value);
			ValDsco = Convert.ToDecimal(this.txtSubTotal.Value)-0;
			this.txtDescnto.Value = ToDscto;

			this.txtIva.Value = Math.Round((ValDsco * PorcentajeIva /100.00m), 6);

			//this.txtTotal.Value = (Math.Round(Convert.ToDecimal(this.txtSubTotal.Value), 4) - Math.Round(Convert.ToDecimal(this.txtDescnto.Value), 2))
			//	+ Math.Round(Convert.ToDecimal(this.txtIva.Value), 4);
			this.txtTotal.Value = Math.Round(Convert.ToDecimal(this.txtSubTotal.Value), 6) + Math.Round(Convert.ToDecimal(this.txtIva.Value), 6);

			#endregion Desplazamiento de grilla			
		}

		private void txtNumCedula_Validated(object sender, System.EventArgs e)
		{
			#region Busca Cliente 
			/* Buscar Número de Bodega */
			string lcQryBod = "SELECT Bodega From Bodega Where idBodega in (Select idSucursal From Habitaciones Where idHabitacion="+idNhab.ToString()+")";
			int laBodega = Funcion.iEscalarSQL(cdsSeteo, lcQryBod);
			NBodga = laBodega;
			string lcQryBd = "SELECT idProyecto From Bodega Where idBodega in (Select idSucursal From Habitaciones Where idHabitacion="+idNhab.ToString()+")";
			idProyct = Funcion.iEscalarSQL(cdsSeteo, lcQryBd);
				
			string sSQL = string.Format("Exec BuscaClientePorRUC '{0}', {1}, {2}, {3}", this.txtNumCedula.Text.ToString(), 1, idNhab, false);
			SqlDataReader drC = Funcion.rEscalarSQL(cdsSeteo, sSQL, true);
			drC.Read();
			if (drC.HasRows)
			{
				idClint = (int)drC.GetValue(0);
				
				this.txtNumCedula.Text = (string)drC.GetValue(1);
				this.txtNombre.Text = (string)drC.GetValue(2);
				if (!this.txtNumCedula.Text.ToString().Equals("9999999999"))
				{
					this.txtDireccion.Text = (string)drC.GetValue(7);
					this.txtTelefono.Text = (string)drC.GetValue(8);
				}
				drC.Close();
			}
			else
			{	
				drC.Close();	
				using (VentaCliente miVCliente = new VentaCliente(this.txtNumCedula.Text, 0, 1))
				{
					if (DialogResult.OK == miVCliente.ShowDialog())
					{
						MessageBox.Show("Cliente registrado correctamente.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
						this.txtNombre.Text = miVCliente.txtNombre.Value.ToString().Trim();
						this.txtDireccion.Text = miVCliente.txtDireccion.Value.ToString().Trim();
						this.txtTelefono.Text = miVCliente.txtTelefono.Value.ToString().Trim();

						string sSQLCl = string.Format("Exec BuscaClientePorRUC '{0}', {1}, {2}, {3}", this.txtNumCedula.Text.ToString(), 1, idNhab, false);
						SqlDataReader drCl = Funcion.rEscalarSQL(cdsSeteo, sSQLCl, true);
						drCl.Read();
						if (drCl.HasRows)
						{
							idClint = (int)drCl.GetValue(0);
				
							this.txtNumCedula.Text = (string)drCl.GetValue(1);
							this.txtNombre.Text = (string)drCl.GetValue(2);
							if (!this.txtNumCedula.Text.ToString().Equals("9999999999"))
							{
								this.txtDireccion.Text = (string)drCl.GetValue(7);
								this.txtTelefono.Text = (string)drCl.GetValue(8);
							}
							drCl.Close();
						}
					}
					else this.txtNumCedula.Focus();
				}			
			}
			drC.Close();		
			#endregion Busca Cliente
		}

		private void txtNumCedula_Validating(object sender, System.ComponentModel.CancelEventArgs e)
		{
			#region Valida identificación
			string VlCed = "";
			if(!this.txtNumCedula.Text.ToString().Equals("9999999999"))
			{
				string stExec = string.Format("Select dbo.ValidaIdentificacion ('{0}', {1})", this.txtNumCedula.Text.ToString(), this.cmbTipoRuc.Value);
				VlCed = Funcion.sEscalarSQL(cdsSeteo, stExec);
			}	
			if (VlCed.StartsWith("Error"))
			{
				MessageBox.Show(VlCed, "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.txtNumCedula.Focus();
				e.Cancel = true;			
			}
			#endregion Valida identificación

		}

		private void txtNumCedula_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (e.KeyChar == 13) this.txtNombre.Focus();
		}

		private void txtNumCedula_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.F3)
			{
				using (BuscaClientes miBusqueda = new BuscaClientes(this.txtNumCedula.Text.ToString(), 1, 1, 2, (int)idNhab, false))
				{
					if (DialogResult.OK == miBusqueda.ShowDialog())	
					{
						idClint = miBusqueda.idCliente;
						this.txtNumCedula.Text = miBusqueda.Ruc;
					}
				}
			}
		}

		private void bntGrabar_Click(object sender, System.EventArgs e)
		{
			/* Grabar Factura Habitación */
			DateTime dtFM2015 = DateTime.Parse("01/01/2016");
			int result = 0;

			if(this.dtFecha.Value == DBNull.Value)
			{
				MessageBox.Show("Debe Seleccionar Factura y obtener la Fecha para Grabar la FACTURA.?","Aviso....!!!", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.dtFecha.Focus();
			} 
			else result = DateTime.Compare((DateTime)this.dtFecha.Value, dtFM2015);

			if(this.txtNumCedula.Text.ToString().Length==0)
			{
				MessageBox.Show("Debe Registrar CLIENTE para poder Grabar la FACTURA.?","Aviso....!!!", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.txtNumCedula.Focus();
			}
			else if(this.dtFecha.Value == DBNull.Value)
			{
				MessageBox.Show("Debe Seleccionar Factura y obtener la Fecha para Grabar la FACTURA.?","Aviso....!!!", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.dtFecha.Focus();
			}
			else if(this.cmbCajero.ActiveRow == null)
			{
				MessageBox.Show("No se a seleccionado Cajero?","Aviso....!!!", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.cmbCajero.Focus();
			}
			else if(this.cmbCobrador.ActiveRow == null)
			{
				MessageBox.Show("No se a seleccionado Cobrador?","Aviso....!!!", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.cmbCobrador.Focus();
			}
			else if(this.cmbFormaPago.ActiveRow == null)
			{
				MessageBox.Show("No a seleccionado FORMA DE PAGO?","Aviso....!!!", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.cmbFormaPago.Focus();
			}
			else if(result < 0)
			{
				MessageBox.Show("No se puede GENERAR Facturas con fecha menor a 01/01/2016?","Aviso....!!!", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.txtNumCedula.Focus();
			}
			else if(idClint == 0) 
			{
				MessageBox.Show("Por favor verifique información de Cliente?","Aviso....!!!", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.txtNumCedula.Focus();
			}
			else
			{
				bool Manual = false;
				string sSQL = "";
				decimal dRetenido = 0;
				decimal dRetenidoIVA = 0;
				decimal dRetenidoRenta = 0;
				int idDetFormaPagoRetencion = 0;
				DateTime dtFechaCadR = DateTime.Today;

				#region Validacion
				#region Valida identificación
				string VlCed = "";
				if(!this.txtNumCedula.Text.ToString().Equals("9999999999"))
				{
					string stExec = string.Format("Select dbo.ValidaIdentificacion ('{0}', {1})", this.txtNumCedula.Text.ToString(), this.cmbTipoRuc.Value);
					VlCed = Funcion.sEscalarSQL(cdsSeteo, stExec);
				}
				if (VlCed.StartsWith("Error"))
				{
					MessageBox.Show(VlCed, "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					this.txtNumCedula.Focus();
					return;
				}
				#endregion Valida identificación
				#region Valida Cliente
				string sSQLValCl = string.Format("Exec BuscaClientePorRUC '{0}', {1}, {2}, {3}", this.txtNumCedula.Text.ToString(), 1, (int)idNhab, false);
				SqlDataReader drC = Funcion.rEscalarSQL(cdsSeteo, sSQLValCl, true);
				drC.Read();
				if (!drC.HasRows)
				{
					MessageBox.Show("Cliente no existe.", "Aviso...!!!", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					this.txtNumCedula.Focus();
					return;
				}
				drC.Close();
				#endregion Valida Cliente
				#region Valida Articulos en el grid
				if (this.ugDetalleFactura.Rows.Count == 0)
				{
					MessageBox.Show("Ingrese los productos para crear la Factura", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
					return;
				}
				#endregion Valida Articulos en el grid
				#region Verifica Total Cero
				if ((decimal)txtTotal.Value == 0.00m) 
				{
					MessageBox.Show("No puede grabar la transacción en Cero.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
					return;
				}
				#endregion Verifica Total Cero
				#endregion Validacion

				#region Grabar
				#region Efectivo - Tarjeta 
				DateTime dtFechaFC = (DateTime) this.dtFecha.Value;
				if (Convert.ToDecimal(this.txtTotal.Value)!=0)
					using (VentaFPagoLot miCobro = new VentaFPagoLot(idCpaFC, (int)this.cmbFormaPago.Value, (decimal)this.txtTotal.Value, NumFactura, false, 0, 0, false, 0, false, 0, false, 0, 0, 0, idClint, 0, 0, 0, 0, 0))
					//using (VentaFormaPagoLotte miCobro = new VentaFormaPagoLotte(idCpaFC, 
					//				 (int)this.cmbFormaPago.Value, 
					//				 (decimal)this.txtTotal.Value, 
					//				 this.txtNumFactura.Text.ToString(), false, 0, 0, false, 0, false, 0, false, 0, 0, 5, idClint, 0, 0, 0, 0, 
					//				 NBodga, this.ugDetalleFactura))
					{
						if (DialogResult.OK == miCobro.ShowDialog())
						{
							using (SqlConnection oConexion = new SqlConnection("Password=Bl45o6$9;Persist Security Info=True;User ID= " + MenuLatinium.stUsuario + ";Initial Catalog=" + MenuLatinium.stDirFacturacion + ";Data Source=" + MenuLatinium.stDirServidor))
							{
								oConexion.Open();

								SqlTransaction oTransaction = null;
								SqlCommand oCmdActualiza = oConexion.CreateCommand();
								oCmdActualiza.Connection = oConexion;
								oCmdActualiza.CommandType = CommandType.Text;
								oCmdActualiza.CommandTimeout = 0;
						
								try
								{
									oTransaction = oConexion.BeginTransaction();
									oCmdActualiza.Transaction = oTransaction;

									/*
									if (bNuevo && bNumeracionAutomatica)
									{
										string sSQLNumero = string.Format("Exec NumeracionLocales 1, {0}, 0", idNhab);
										oCmdActualiza.CommandText = sSQLNumero;
										this.txtNumero.Text = oCmdActualiza.ExecuteScalar().ToString(); 
									}	
									*/

									#region Graba Factura
									#region Maestro
									string FechaFC = this.ugFacRelizadas.ActiveRow.Cells["Fecha"].Value.ToString();
									decimal vlrDescuento = 0.00m;
									if (Convert.ToDecimal(this.txtDescnto.Value) != 0) vlrDescuento = Convert.ToDecimal(this.txtDescnto.Value);
									sSQL = string.Format("Exec GrabaMaestroFacturaLotte {0}, {1}, {2}, {3}, {4}, {5}, {6}, {7}, {8}, {9}, {10}, {11}, '{12}', '{13}', {14}, {15}, '{16}', '{17}', '{18}', '{19}', '{20}', '{21}', {22}, '{23}', '{24}', {25}, '{26}', {27}, {28}, {29}, {30}, '{31}', {32}, {33}, {34}, {35}, {36}, {37}, {38}, {39}, {40}, {41}, {42}, 1, 0, 0, '', 0.0, {43}, {44}, {45}, {46}, {47}, {48}, {49}, {50}, {51}, {52}, {53}, '', 0", 
										idCpaFC, 1,	0, 1, 15, 10, false, 1, (int)this.cmbFormaPago.Value, NBodga, idProyct, (int)this.cmbVendedor.Value, NumFactura, "",
										(int)this.cmbTipoRuc.Value, idClint, this.txtNumCedula.Text.ToString(), this.txtNombre.Text.ToString(), dtFechaFC.ToString("yyyyMMdd"), 
										"", "", dtFechaFC.ToString("yyyyMMdd"), false, dtFechaFC.ToString("yyyyMMdd"), dtFechaFC.ToString("yyyyMMdd HH:mm"), 0,
										"", false, 0, 0, 0, dtFechaFC.ToString("yyyyMMdd"), 0, 0, (decimal)this.txtSubTotal.Value, 
										vlrDescuento, (decimal)this.txtIva.Value, 
										(decimal)this.txtTotal.Value,	false, false, false, false, false, 0, 0, 0, 0, false, 0, 0, 0, 0, 0, 0);
									oCmdActualiza.CommandText = sSQL;
									idCpaFC = (int)oCmdActualiza.ExecuteScalar();
									#endregion Maestro

									#region unidad taxi
									string lcQryTaxi = string.Format("Exec GrabaTaxi {0},'{1}'", idCpaFC, this.txtUnidadTaxi.Text.ToString());
									oCmdActualiza.CommandText = lcQryTaxi;
									oCmdActualiza.ExecuteNonQuery();
                  #endregion unidad taxi
									
									#region Actualiza conbrador y cajero
									string lcQryCC = string.Format("Exec ActualizaCajCob {0},{1},{2},1", 
										idCpaFC, (int)this.cmbCobrador.Value, (int)this.cmbCajero.Value);
									oCmdActualiza.CommandText = lcQryCC;
									oCmdActualiza.ExecuteNonQuery();
									#endregion Actualiza conbrador y cajero

									#region cajero vendedor cobrador
									string lcQryCVC = string.Format("Exec GrabaVen_Caj_Cob_Turno {0},{1},{2},{3},1", 
										idCpaFC, (int)this.cmbCobrador.Value, 0, (int)this.cmbCajero.Value);
									oCmdActualiza.CommandText = lcQryCVC;
									oCmdActualiza.ExecuteNonQuery();
									#endregion cajero vendedor cobrador
									#region Detalle
									int iPosicion = 0;
									foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.ugDetalleFactura.Rows.All)
									{	
										iPosicion ++;
										string lcQryCAr = "SELECT Codigo From Articulo Where idArticulo="+dr.Cells["idArticulo"].Value.ToString();
										string CodigoArti = Funcion.sEscalarSQL(cdsSeteo, lcQryCAr);

										string sSQLDetalle = string.Format("Exec GuardaDetalleFacturaLotte {0}, {1}, {2}, {3}, '{4}', {5}, {6}, '{7}', '{8}', {9}, {10}, {11}, {12}, {13}, {14}, {15}, {16}, {17}, {18}, {19}, {20}, {21}, {22}, '{23}'",
											(int)dr.Cells["idDetCompra"].Value, 0, (int)dr.Cells["idCompra"].Value,	false, "", (int)dr.Cells["idArticulo"].Value, 0, 
											CodigoArti, dr.Cells["Articulo"].Value,	(int)dr.Cells["Cantidad"].Value, (decimal)dr.Cells["Precio"].Value, 
											PorcentajeIva, 0, (decimal)dr.Cells["PDscto"].Value, (decimal)dr.Cells["Subtotal"].Value,	NBodga, -1,
											0, false, iPosicion, false, false, 0, "");

										/*string sSQLDetalle = string.Format("Exec GuardaDetalleFacturaLotte {0}, {1}, {2}, {3}, '{4}', {5}, {6}, '{7}', '{8}', {9}, {10}, {11}, {12}, {13}, {14}, {15}, {16}, {17}, {18}, {19}, {20}, {21}, {22}",
											(int)dr.Cells["idDetCompra"].Value, 0, (int)dr.Cells["idCompra"].Value,	false, "", (int)dr.Cells["idArticulo"].Value, 0, 
											CodigoArti, dr.Cells["Articulo"].Value,	(int)dr.Cells["Cantidad"].Value, (decimal)dr.Cells["Precio"].Value, 
											PorcentajeIva, 0, 0, (decimal)dr.Cells["Subtotal"].Value,	NBodga, -1,
											0, false, iPosicion, false, false, 0);*/

										oCmdActualiza.CommandText = sSQLDetalle;
										oCmdActualiza.ExecuteNonQuery();
									}
									#endregion Detalle	
									#endregion Graba Factura

									#region Guarda Cobro de la factura	
									if (miCobro.ultraGrid1.Rows.Count > 0)
									{
										foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in miCobro.ultraGrid1.Rows.All)
										{
											int idDetFormaPago = 0;
											int idBanco = 0;
											int idTarjeta = 0;
											int idCuenta = 0;
											int idPlan = 0;
											int idPlazo = 0;
											decimal dComTarjeta = 0.00m;
											int idDocumento = 0;
											string DesPrenda = "";
											DateTime dtFecCaducidad = DateTime.Today;
							
											if (dr.Cells["idDetFormaPago"].Value != System.DBNull.Value) idDetFormaPago = (int)dr.Cells["idDetFormaPago"].Value;
											if (dr.Cells["idBanco"].Value != System.DBNull.Value) idBanco = (int)dr.Cells["idBanco"].Value;
											if (dr.Cells["idTarjeta"].Value != System.DBNull.Value) idTarjeta = (int)dr.Cells["idTarjeta"].Value;
											if (dr.Cells["idCuenta"].Value != System.DBNull.Value) idCuenta = (int)dr.Cells["idCuenta"].Value;
											if (dr.Cells["idPlan"].Value != System.DBNull.Value) idPlan = (int)dr.Cells["idPlan"].Value;							
											if (dr.Cells["idPlazo"].Value != System.DBNull.Value) idPlazo = (int)dr.Cells["idPlazo"].Value;
											if ((int)dr.Cells["idFormaPago"].Value == 3) dtFecCaducidad = (DateTime)dr.Cells["FechaCaducidad"].Value;
											if (dr.Cells["ValorUsoTarjeta"].Value != System.DBNull.Value) dComTarjeta = (decimal)dr.Cells["ValorUsoTarjeta"].Value;
											if (dr.Cells["idDocumento"].Value != System.DBNull.Value) idDocumento = (int)dr.Cells["idDocumento"].Value;
											if (dr.Cells["DescripcionPrenda"].Value != System.DBNull.Value) DesPrenda = dr.Cells["DescripcionPrenda"].Value.ToString();
								
											string sSQLCobro = string.Format("Exec GuardaDetFormaPagoLotte {0}, {1}, {2}, {3}, {4}, {5}, {6}, '{7}', '{8}', '{9}', '{10}', '{11}', {12}, '{13}', '{14}', {15}, {16}, {17}, {18}, '{19}',{20}", 
												idDetFormaPago, 
												idCpaFC, 1, (int)dr.Cells["idFormaPago"].Value, (decimal)dr.Cells["Valor"].Value,
												idBanco, idTarjeta, dr.Cells["CuentaCorriente"].Value.ToString(), dr.Cells["Numero"].Value.ToString(),
												dtFecCaducidad.ToString("yyyyMMdd"),	dr.Cells["NumVoucher"].Value.ToString(), dr.Cells["Autorizacion"].Value.ToString(),
												idCuenta, dr.Cells["Lote"].Value.ToString(), dr.Cells["Referencia"].Value.ToString(), 0, 0, 0, 0, DesPrenda,0);
											oCmdActualiza.CommandText = sSQLCobro;
											oCmdActualiza.ExecuteNonQuery();									

											/* VALOR DE ANTICIPOS */
											//if ((int)dr.Cells["idFormaPago"].Value == 8) dAnticipo = dAnticipo + Convert.ToDecimal(dr.Cells["Valor"].Value);

											/* ACTUALIZA SALDO DE NOTA DE CREDITO */
											if (idDocumento > 0)
											{
												oCmdActualiza.CommandText = string.Format("Exec ActualizaSaldoFacturas {0}, 5, 4", idDocumento, 1, 4);
												oCmdActualiza.ExecuteNonQuery();										
											}
										}								
									}
									#endregion Guarda Cobro de la factura

									#region Saldo Factura
									if(false)
									{
										if ((int)this.cmbFormaPago.Value == 9)
										{
											string sSQLSaldo = string.Format("Exec ActualizaSaldoFacturas {0}, {1}, {2}", idCpaFC, 1, (int)this.cmbFormaPago.Value);
											oCmdActualiza.CommandText = sSQLSaldo;
											oCmdActualiza.ExecuteNonQuery();								
										}
									}
									#endregion Saldo Factura

									#region Asiento Automatico
									//if (bAsientoAutomatico)
									if (true) //Hay que poner en verdadero cuando este listo el plan de cuentas
									{						
										string sSQLAsiento = string.Format("Exec CreaAsientoDeVenta {0}", idCpaFC.ToString());
										oCmdActualiza.CommandText = sSQLAsiento;
										oCmdActualiza.ExecuteNonQuery();

										string sSQLRetIdAsiento = string.Format("Select ISNULL(idAsiento, 0) From Compra Where idCompra = {0}", idCpaFC.ToString());
										oCmdActualiza.CommandText = sSQLRetIdAsiento; 
										idAsient = (int)oCmdActualiza.ExecuteScalar(); 
								
										Cursor = Cursors.Default;								
									}
									#endregion Asiento Automatico

                  #region Cuadrar factura
									//string sSQLCents = string.Format("Exec CuadraCentavos {0},'{1}'", idCpaFC.ToString(),aplicapromocion);
									//oCmdActualiza.CommandText = sSQLCents;
									//oCmdActualiza.ExecuteNonQuery();
                  #endregion Cuadrar factura

									#region Actualizar Cupon Web
									string sSQLCpon = string.Format("Exec ActualizaCuponWeb {0},{1}", idCpnWeb,idCpaFC);
									oCmdActualiza.CommandText = sSQLCpon;
									oCmdActualiza.ExecuteNonQuery();
									#endregion Actualizar Cupon Web

									oTransaction.Commit();

									#region Impresión
									using (frmMnsajeNumLotte miMsje = new frmMnsajeNumLotte( this.ugFacRelizadas.ActiveRow.Cells["Numero"].Value.ToString(), false) )
									{
										if (DialogResult.OK == miMsje.ShowDialog())
										{
											if (Funcion.bEjecutaSQL(cdsSeteo, "Exec SeteoGExiste 'FVIG'")) Imprimir(0);
										}
									}
									#endregion Impresión

									#region Impresión Promocion
									//Verifica si tiene promocion
									//string bskPro = "Select idRegPromocion From RegistroPromociones Where idCompra="+idCpaFC.ToString();
									//int idRePro = Funcion.iEscalarSQL(cdsSeteo, bskPro); 
									//if (idRePro>0) Imprimir(1);
									#endregion Impresión Promocion

									//Cambia estado de habitación a Limpieza
									string SP_Act = string.Format("Exec Cambia_Estado_Habitacion {0},{1},{2},1", idNhab,3,NBodga);
									Funcion.EjecutaSQL(cdsSeteo,SP_Act);

									this.label13.Text = "CONTABILIZADO";
					
									#region Controles
									this.cmbFormaPago.Enabled = false;
									this.btnImprimir.Enabled = false;
									this.bntGrabar.Enabled = false;
									#endregion Controles
								}
								catch (Exception ex)
								{
									try
									{
										oTransaction.Rollback();

										MessageBox.Show(string.Format("Commit Exception Type : {0} {1}", ex.GetType(), ex.Message), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
									}
									catch (Exception ex2)
									{
										oTransaction.Rollback();

										MessageBox.Show(string.Format("Rollback Exception Type: {0} {1}", ex2.GetType(), ex2.Message), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);							
									}
								}
								finally
								{
									oConexion.Close();
								}
							}
						}
						//else bEdicion = true;
					}
				#endregion Efectivo - Tarjeta 
				#endregion Grabar
			}
		}

		private void ugFacRelizadas_DoubleClick(object sender, System.EventArgs e)
		{
			// obtener la fecha de factura
			try
			{
				this.dtFecha.Value = ugFacRelizadas.ActiveRow.Cells["Fecha"].Value;
				NumFactura = ugFacRelizadas.ActiveRow.Cells["Numero"].Value.ToString();
				this.txtNumCedula.Text = "9999999999";
				this.cmbVendedor.Value = ugFacRelizadas.ActiveRow.Cells["idVendedor"].Value;
				string sSQL = string.Format("Exec BuscaClientePorRUC '{0}', {1}, {2}, {3}", this.txtNumCedula.Text.ToString(), 1, idNhab, false);
				SqlDataReader drC = Funcion.rEscalarSQL(cdsSeteo, sSQL, true);
				drC.Read();
				if (drC.HasRows)
				{
					idClint = (int)drC.GetValue(0);
				
					this.txtNumCedula.Text = (string)drC.GetValue(1);
					this.txtNombre.Text = (string)drC.GetValue(2);
					if (!this.txtNumCedula.Text.ToString().Equals("9999999999"))
					{
						this.txtDireccion.Text = (string)drC.GetValue(7);
						this.txtTelefono.Text = (string)drC.GetValue(8);
					}
					drC.Close();
				}
			}
			catch (Exception ex2) 
			{
				MessageBox.Show(string.Format("Error: {0} {1}", ex2.GetType(), ex2.Message), "Aviso...!!!", MessageBoxButtons.OK, MessageBoxIcon.Error);
			}
		}

		private void Imprimir(int iTipo)
		{
			#region Nombre de Reporte y Filtro
			Cursor = Cursors.WaitCursor;
			string stReporte = "Exec CompraReporteNombre " + idCpaFC.ToString();
			
			stReporte += ", 1";
			string stFiltro = "";
			stReporte = Funcion.sEscalarSQL(cdsSeteo, stReporte);
			if (iTipo==0)	stFiltro = "{Compra.idCompra} = " + idCpaFC.ToString();
			else stFiltro = "{RegistroPromociones.idCompra} = " + idCpaFC.ToString();
			#endregion Nombre de Reporte y Filtro

			#region Imprime el reporte
			//using (Reporte miReporte = new Reporte(stReporte, stFiltro))
			if (iTipo == 0) 
			{
				#region Impresion
				ParameterFields paramFields = new ParameterFields ();
				
				ParameterField pfRet = new ParameterField ();
				ParameterDiscreteValue idCompraRet = new ParameterDiscreteValue ();
				pfRet.ParameterFieldName = "@idCpa";
				idCompraRet.Value = idCpaFC;
				pfRet.CurrentValues.Add (idCompraRet);
				paramFields.Add (pfRet);
				
				Reporte miRepor = new Reporte("FacturaLotteServicio.rpt", "");
				miRepor.MdiParent = this.MdiParent;
				miRepor.Titulo("Retenciones");
				miRepor.crVista.ParameterFieldInfo = paramFields;
				#endregion Impresion
				
				miRepor.Show();

				/*
				using (Reporte miReporte = new Reporte("FacturaLotteServicio", stFiltro))
				{		
					if (iTipo == 0) // Muestra en pantalla
						miReporte.ShowDialog();
					if (iTipo >= 1) // Imprime una o mas copias
					{
						string ImpresoraPredefinida = MenuLatinium.ImpresoraNombre;
						string stImpresora = Funcion.sEscalarSQL(cdsSeteo, "Exec CompraReporteImpresora '" + stReporte + "'");
						if (stImpresora.Length > 0)	MenuLatinium.ImpresoraNombre = stImpresora;
						miReporte.ImprimeRep(1);
						if (stImpresora.Length > 0)	MenuLatinium.ImpresoraNombre = ImpresoraPredefinida;
					}
					if (iTipo == -1) // Archivo PDF
					{
						miReporte.ImprimeRep(2);
					}
				}
				*/
			}
			if (iTipo == 1) 
			{
				using (Reporte miReporte = new Reporte("ImprPromocion", stFiltro))
				{		
					if (iTipo == 1) // Muestra en pantalla
						miReporte.ShowDialog();
					if (iTipo >= 2) // Imprime una o mas copias
					{
						string ImpresoraPredefinida = MenuLatinium.ImpresoraNombre;
						string stImpresora = Funcion.sEscalarSQL(cdsSeteo, "Exec CompraReporteImpresora '" + stReporte + "'");
						if (stImpresora.Length > 0)	MenuLatinium.ImpresoraNombre = stImpresora;
						miReporte.ImprimeRep(1);
						if (stImpresora.Length > 0)	MenuLatinium.ImpresoraNombre = ImpresoraPredefinida;
					}
					if (iTipo == -1) // Archivo PDF
					{
						miReporte.ImprimeRep(2);
					}
				}
			}
			#endregion Imprime el reporte

			#region Marca como impreso y muestra el tiempo
			string stGrabaImprime = "Update Compra Set Impreso = 1 Where idCompra = " + idCpaFC.ToString();
			Funcion.EjecutaSQL(cdsSeteo, stGrabaImprime);
			bImpreso = true;
			if (!miAcceso.ReImprimir) this.btnImprimir.Enabled = false;
			Cursor = Cursors.Default;
			#endregion Marca como impreso y muestra el tiempo
		}

		private void ugArticulos_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}

		private void button4_Click(object sender, System.EventArgs e)
		{
			// Cargar Datos de bebidas
			//string sSQL = string.Format("Select idSubGrupo,idGrupoArticulo,SubGrupo From ArticuloSubGrupo Where idGrupoArticulo=42");
			//ugServicios.DataSource = Funcion.dvProcedimiento(cdsSeteo, sSQL);
		}

		private void ugServicios_ClickCellButton(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			// Carga articulos de acuerso al subgrupo
			//int nSubgrp = (int)ugServicios.ActiveRow.Cells["idSubGrupo"].Value;
			//string sSQL = string.Format("Select idArticulo,idSubGrupo,idGrupoArticulo,Articulo,Precio1,Precio2 From Articulo Where idSubGrupo="+nSubgrp.ToString());
			//ugArticulos.DataSource = Funcion.dvProcedimiento(cdsSeteo, sSQL);

		}

		private void ugArticulos_ClickCellButton(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			// Registrar articulo en detalle de factura
			try
			{
				if (idCpaFC != 0)
				{
					idCpaFC = (int)ugFacRelizadas.ActiveRow.Cells["idCompra"].Value;
				
					int idCltRuc = 0;
					if (this.txtNumCedula.Text.Length!=0) 
					{
						string lcQry = "SELECT idCliente FROM Cliente WHERE Ruc="+this.txtNumCedula.Text.ToString();
						idCltRuc = Funcion.iEscalarSQL(cdsSeteo, lcQry); 
					}
					int idClFc = 0;
					if(idCltRuc>0) idClFc = idCltRuc;
					int idArtFc = (int)ugArticulos.ActiveRow.Cells["idArticulo"].Value;
					decimal preFc = 0;

					/*Verificar si producto se encuentra con promoción*/
					BuscaPromocionArticulo(idArtFc);
					//string sSQL = string.Format("Exec BuscaPromoArticulo '{0}'", idArtFc);
					//SqlDataReader dr = Funcion.rEscalarSQL(cdsSeteo, sSQL, true);
					//dr.Read();
					//if(dr.HasRows)
					//{
					//	respuestapromocion = dr.GetString(0).Trim();
					//	tpopromocion = dr.GetInt32(1);
					//	valorpromo = dr.GetDecimal(2);
					//}
					//dr.Close();

					if(this.chkPrecios.Checked==false)	preFc = (decimal)ugArticulos.ActiveRow.Cells["Precio2"].Value;
					else preFc = (decimal)ugArticulos.ActiveRow.Cells["Precio1"].Value;
				
					//Validar existencia de Producto
					//int Existencia = Funcion.iEscalarSQL(cdsSeteo, string.Format("Select dbo.ExistenciasInventarioLotte({0}, {1})", idArtFc, NBodga));
					int Existencia = Funcion.iEscalarSQL(cdsSeteo, string.Format(" Exec Existencia_Receta {0}, {1}", idArtFc, NBodga));
				
					int iCantidad = 0;
					if (iCantidad > Existencia)
					{					
						if (Existencia < 1)
						{
							//bEliminaAlValidar = true;
							//e.Cell.Row.Delete();
							//bEliminaAlValidar = false;
							MessageBox.Show("Este Articulo No Tiene Existencia", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
						}
						else
						{
							e.Cell.Value = 1;
							MessageBox.Show(string.Format("Existencia Actual : {0}", Existencia), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
						}				
					}
					else if(Existencia==0)
					{
						e.Cell.Value = 1;
						MessageBox.Show(string.Format("Existencia Actual : {0}", Existencia), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					}
					else
					{
						string SP_Insert = string.Format("Exec GrabaDetalle_FCHabitacion {0},{1},{2},{3},{4},'{5}',{6},{7}", idCpaFC,idClFc,idArtFc,1,preFc,respuestapromocion,tpopromocion,valorpromo);
						Funcion.EjecutaSQL(cdsSeteo,SP_Insert);
						string QryDetFc = "SELECT idDetCompra,idCompra,idArticulo,Articulo,Cantidad,Precio,Convert(Decimal(10,2),DescuentoPorc) as PDscto,SubTotal From DetCompra Where idCompra="+idCpaFC.ToString();
						ugDetalleFactura.DataSource = Funcion.dvProcedimiento(cdsSeteo, QryDetFc);
						Total();
					}
				} 
				else 
				{
					MessageBox.Show(string.Format("No se ha REGISTRADO FACTURA...?\nPor favor REGISTRE # Factura para continuar...", "Aviso...!!!", MessageBoxButtons.OK, MessageBoxIcon.Stop));
					this.txtNumFactura.Focus();
				}
			}
			catch (Exception ex2) 
			{
				MessageBox.Show(string.Format("Error: {0} {1}", ex2.GetType(), ex2.Message), "Aviso...!!!", MessageBoxButtons.OK, MessageBoxIcon.Error);
			}
		}

		private void ugDetalleFactura_BeforeRowsDeleted(object sender, Infragistics.Win.UltraWinGrid.BeforeRowsDeletedEventArgs e)
		{
			for (int i = 0; i < e.Rows.Length; i++)
			{
				if ((int)e.Rows[i].Cells["idDetCompra"].Value > 0)
				{	
					if (DialogResult.Yes ==	MessageBox.Show("Desea Eliminar Registro.?", "Point Technology", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2)) 
					{
						Funcion.EjecutaSQL(cdsSeteo, string.Format("Exec GrabaRegistroAnulacion {0}", (int)e.Rows[i].Cells["idDetCompra"].Value));
						Funcion.EjecutaSQL(cdsSeteo, string.Format("Delete From DetCompra Where idDetCompra = {0}", (int)e.Rows[i].Cells["idDetCompra"].Value));
						e.DisplayPromptMsg = false;
						Total();
					}
					else e.Cancel = true;
				}
			}
		}

		private void ugDetalleFactura_AfterCellUpdate(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			#region Cantidad
			if (e.Cell.Column.ToString() == "Cantidad")
			{
				if (e.Cell.Row.Cells["Cantidad"].Value == System.DBNull.Value) e.Cell.Row.Cells["Cantidad"].Value = (int)e.Cell.OriginalValue;
				if (true)
				{
					//if (Funcion.bEscalarSQL(cdsSeteo, string.Format("Declare @Res Bit = 0 If ((Select idTipoGrupo From Articulo Where idArticulo = {0}) In (1, 6) And (Select idGrupoArticulo From Articulo Where idArticulo = {0}) <> 26) Set @Res = 1 Select @Res", (int)e.Cell.Row.Cells["idArticulo"].Value), true))
					//{
					int idArticulo = (int)e.Cell.Row.Cells["idArticulo"].Value;
					//int Existencia = Funcion.iEscalarSQL(cdsSeteo, string.Format("Select dbo.ExistenciasInventarioLotte({0}, {1})", idArticulo, NBodga));
					int Existencia = Funcion.iEscalarSQL(cdsSeteo, string.Format(" Exec Existencia_Receta {0}, {1}", idArticulo, NBodga));
					int iCantidad = 0;

					foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.ugDetalleFactura.Rows.All)
					{
						if (idArticulo == (int)dr.Cells["idArticulo"].Value) iCantidad = iCantidad + (int)dr.Cells["Cantidad"].Value;
					}

					if(idArticulo==3422 || idArticulo==3423) 
					{	
						if (e.Cell.Column.ToString() == "Cantidad") 
							//Calcular nuevo tiempo 
							Funcion.EjecutaSQL(cdsSeteo, string.Format("CalculaNW_Tiempo {0}, {1}", idNhab, iCantidad));
							Total();
					}
					else
					{
						if (iCantidad > Existencia)
						{					
							if (Existencia < 1)
							{
								//bEliminaAlValidar = true;
								//e.Cell.Row.Delete();
								//bEliminaAlValidar = false;
								MessageBox.Show("Este Articulo No Tiene Existencia", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
							}
							else
							{
								e.Cell.Value = 1;
								MessageBox.Show(string.Format("Existencia Actual : {0}", Existencia), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
							}				
						}
						else if (e.Cell.Column.ToString() == "Cantidad") Total();
					}
				}
			}
			#endregion Cantidad
		}

		private void ugDetalleFactura_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (ugDetalleFactura.ActiveCell != null)
			{
				try
				{				
					switch(e.KeyValue)
					{
						#region Enter
						case (int) Keys.Enter:
							ugDetalleFactura.PerformAction(UltraGridAction.ExitEditMode, false, false);
							ugDetalleFactura.PerformAction(UltraGridAction.NextCellByTab, false, false);
							e.Handled = true;
							ugDetalleFactura.PerformAction(UltraGridAction.EnterEditMode, false, false);
							//							if(ultraGrid1.ActiveCell.DroppedDown == false) ultraGrid1.ActiveCell.DroppedDown = true;
							//						ultraGrid1.PerformAction(UltraGridAction.EnterEditModeAndDropdown, false, false);
							break;
						#endregion Enter
					}
				}
				catch(System.Exception ex)
				{
					MessageBox.Show(ex.Message, "Error del sistema.");
				}
			}
		}

		private void frmFCHabitaciones_Closing(object sender, System.ComponentModel.CancelEventArgs e)
		{
			//Validar que tenga servicio de hora
			if (idCpaFC>0)
			{
				#region Valida Articulos en el grid
				//if (this.ugDetalleFactura.Rows.Count == 0)
				//{
				//	MessageBox.Show("Ingrese producto de servicio para CREAR la Factura", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
				//	e.Cancel = true;
				//}
				#endregion Valida Articulos en el grid
				#region Desplazamiento de grilla
				//int registrhora = 0;
				//foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in ugDetalleFactura.Rows.All)
				//{
				//	if ((int)dr.Cells["idArticulo"].Value == 3422)
				//	{
				//		registrhora ++;
				//	}
				//}
				//if (registrhora==0)
				//{
				//	MessageBox.Show("Registre SERVICIO de HORA para grabar factura.!!!", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
				//	e.Cancel = true;
				//}
				#endregion Desplazamiento de grilla
			}

		}

		private void btBuscar_Click(object sender, System.EventArgs e)
		{
			this.Cursor = Cursors.WaitCursor;
			string stBusca = txtBskArticulo.Text.Trim();
			stBusca = stBusca.Replace("'", "''");
				
			//iBodega = (int)this.cmbBodega.Value;			

			DateTime dtFecha = DateTime.Today;

			string sSQL = string.Format("Exec ConsultaInvLotte '{0}'", stBusca);
			ugArticulos.DataSource = Funcion.dvProcedimiento(cdsSeteo, sSQL);
					
			ugArticulos.Focus();
			this.Cursor = Cursors.Default;
		}

		private void btnImprimir_Click(object sender, System.EventArgs e)
		{
		
		}

		private void txtTiempo_Click(object sender, System.EventArgs e)
		{
			this.txtTiempo.SelectAll();
		}

		private void txtTiempo_Enter(object sender, System.EventArgs e)
		{
		  this.txtTiempo.SelectAll();
		}

		private void textBox1_Validated(object sender, System.EventArgs e)
		{
			#region Busca Promocion 
			/* Buscar Número de Bodega */
			string lcQryBod = "SELECT Bodega From Bodega Where idBodega in (Select idSucursal From Habitaciones Where idHabitacion="+idNhab.ToString()+")";
			int laBodega = Funcion.iEscalarSQL(cdsSeteo, lcQryBod);
			NBodga = laBodega;
			string sSQL = string.Format("Exec BuscaPromocion {0}, '{1}'", NBodga, this.txtnPromo.Text.ToString().Trim());
			string Rsltado = Funcion.sEscalarSQL(cdsSeteo, sSQL);
			aplicapromocion = Rsltado;
			if (Rsltado.Equals("NO"))
			{
				MessageBox.Show("Número de promoción no Valida\nPor favor verifique número de promoción.?","Aviso...!!!",MessageBoxButtons.OK, MessageBoxIcon.Information);
				this.txtnPromo.Text = "";
			}
			else
			{
				DialogResult resultado = MessageBox.Show("Desea Aplicar PROMOCION.?","Aviso...!!!",MessageBoxButtons.YesNo);
				if (DialogResult.Yes == resultado)
				{
					string SP_Qryp = string.Format("Exec Fact_Promocion_Indivisual {0},'{1}'", 
						idCpaFC,this.txtnPromo.Text.ToString().Trim());
					Funcion.EjecutaSQL(cdsSeteo,SP_Qryp);
					aplicapromocion = "SI";
					//Visualiza detalle factura
					string QryDetFc = "SELECT idDetCompra,idCompra,idArticulo,Articulo,Cantidad,Precio,DescuentoPorc as PDscto,SubTotal From DetCompra Where idCompra="+idCpaFC.ToString();
					ugDetalleFactura.DataSource = Funcion.dvProcedimiento(cdsSeteo, QryDetFc);
					Total();
				}
				this.txtDireccion.Focus();
			}
			#endregion Busca Promocion
		}

		private void txtnPromo_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (e.KeyChar == 13) this.txtNumFactura.Focus();
		}

		private void ugDetalleFactura_AfterRowsDeleted(object sender, System.EventArgs e)
		{
			Total();
		}

		private void txtBskArticulo_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (e.KeyChar == 13) this.btBuscar_Click(sender, e);
		}

		private void txtDescnto_ValueChanged(object sender, System.EventArgs e)
		{
			//Total();
		}

		private void txtDescnto_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (e.KeyChar == 13)
			{
				/*Calcular descuento de Habitación*/
				string SP_Qry = string.Format("Exec DescuentoFacturaLotteServicio {0},{1}", 
					idCpaFC,Convert.ToDecimal(this.txtDescnto.Value));
				Funcion.EjecutaSQL(cdsSeteo,SP_Qry);
				
				//Visualiza detalle factura
				string QryDetFc = "SELECT idDetCompra,idCompra,idArticulo,Articulo,Cantidad,Precio,SubTotal From DetCompra Where idCompra="+idCpaFC.ToString();
				ugDetalleFactura.DataSource = Funcion.dvProcedimiento(cdsSeteo, QryDetFc);

				Total();
			}
		}

		private void txtDescnto_BeforeExitEditMode(object sender, Infragistics.Win.BeforeExitEditModeEventArgs e)
		{
			/*Calcular descuento de Habitación*/
			string SP_Qry = string.Format("Exec DescuentoFacturaLotteServicio {0},{1}", 
				idCpaFC,Convert.ToDecimal(this.txtDescnto.Value));
			Funcion.EjecutaSQL(cdsSeteo,SP_Qry);
				
			//Visualiza detalle factura
			string QryDetFc = "SELECT idDetCompra,idCompra,idArticulo,Articulo,Cantidad,Precio,SubTotal From DetCompra Where idCompra="+idCpaFC.ToString();
			ugDetalleFactura.DataSource = Funcion.dvProcedimiento(cdsSeteo, QryDetFc);

			Total();
		}

		private void txtnPromo_TextChanged(object sender, System.EventArgs e)
		{
		
		}

		private void ugDetalleFactura_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}

	}
}
