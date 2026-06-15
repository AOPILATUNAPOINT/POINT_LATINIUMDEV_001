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
	/// Descripción breve de frmFacturacion.
	/// </summary>
	public class frmFacturacion: System.Windows.Forms.Form
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
    int idDetFormaPago = 0;
		string niUsuario = "";
		int idGrpClte = 0;
		string DatosDescuento = "";
		string DetalleDePrenda = "";
		decimal TotalDeFactura = 0.00m;
		decimal TotalDeetalleFactura = 0.00m;
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
		private Infragistics.Win.UltraWinGrid.UltraGrid ugDetalleFactura;
		private System.Windows.Forms.Button bntGrabar;
		private System.Windows.Forms.TextBox txtNumFactura;
		private System.Windows.Forms.TextBox txtNumCedula;
		private System.Windows.Forms.TextBox txtNombre;
		private System.Windows.Forms.TextBox txtTelefono;
		private C1.Data.C1DataSet cdsSeteo;
		private System.Windows.Forms.ToolTip toolTip1;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource1;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource2;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource4;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource3;
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
		private System.Windows.Forms.Label label17;
		private System.Windows.Forms.Label label18;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbVendedor;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbCobrador;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbCajero;
		private System.Windows.Forms.Label label19;
		private System.Windows.Forms.Label lblBodega;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbBodega;
		private System.Windows.Forms.Label label9;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource5;
		private Infragistics.Win.UltraWinGrid.UltraGrid ugHabitaciones;
		private System.Windows.Forms.TextBox txtNumHabitacion;
		private System.Windows.Forms.Button btnAnular;
		private System.Windows.Forms.Button button2;
		private System.Windows.Forms.Button button3;
		private System.Windows.Forms.Button btnElimina;
		private System.Windows.Forms.Label lblIdentificacion;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbTipoRuc;
		private System.Windows.Forms.RadioButton rb2;
		private System.Windows.Forms.RadioButton rb1;
		private System.Windows.Forms.Label label15;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtDescnto;
		private System.Windows.Forms.Label label16;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtBaseImponible;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtTotDescuento;
		private System.Windows.Forms.Label label22;
		private System.Windows.Forms.Label label20;
		private System.Windows.Forms.Button btnImprimir;
		private System.ComponentModel.IContainer components;
		#endregion controles pantalla

		public frmFacturacion()
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
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idBodega");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bodega");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idBodega");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Bodega");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn3 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmFacturacion));
			Infragistics.Win.Appearance appearance1 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idArticulo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idGrupoArticulo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idSubGrupo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Articulo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Precio1");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Precio2");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("UnboundColumn1", 0);
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn4 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idArticulo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn5 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idGrupoArticulo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn6 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idSubGrupo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn7 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Articulo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn8 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Precio1");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn9 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Precio2");
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand3 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn11 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idHabitacion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn12 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idSucursal");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn13 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("NumeroHabitacion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn14 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn15 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("PrecioNormal");
			Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance11 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance12 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance13 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance14 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance15 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance16 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn10 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idHabitacion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn11 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idSucursal");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn12 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("NumeroHabitacion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn13 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn14 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("PrecioNormal");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn15 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idSubGrupo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn16 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idGrupoArticulo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn17 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("SubGrupo");
			Infragistics.Win.Appearance appearance17 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance18 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance19 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance20 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand4 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn16 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idTipoRuc");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn17 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("TipoRuc");
			Infragistics.Win.Appearance appearance21 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand5 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn18 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPersonal");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn19 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn20 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.Appearance appearance22 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand6 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn21 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPersonal");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn22 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn23 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.Appearance appearance23 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand7 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn24 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPersonal");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn25 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn26 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.Appearance appearance24 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton1 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.Appearance appearance25 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand8 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn27 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idFormaPago");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn28 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FormaPago");
			Infragistics.Win.Appearance appearance26 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance27 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance28 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance29 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand9 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn29 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idDetCompra");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn30 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCompra");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn31 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idArticulo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn32 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Articulo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn33 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cantidad");
			Infragistics.Win.Appearance appearance30 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn34 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Precio");
			Infragistics.Win.Appearance appearance31 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn35 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("DescuentoPorc");
			Infragistics.Win.Appearance appearance32 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn36 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("SubTotal");
			Infragistics.Win.Appearance appearance33 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn37 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("DescGlobal");
			Infragistics.Win.Appearance appearance34 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance35 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance36 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance37 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance38 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance39 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn18 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idDetCompra");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn19 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCompra");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn20 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idArticulo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn21 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Articulo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn22 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Cantidad");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn23 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Precio");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn24 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("DescuentoPorc");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn25 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("SubTotal");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn26 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("DescGlobal");
			this.groupBox1 = new System.Windows.Forms.GroupBox();
			this.lblBodega = new System.Windows.Forms.Label();
			this.cmbBodega = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.ultraDataSource5 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.btBuscar = new DevExpress.XtraEditors.SimpleButton();
			this.label12 = new System.Windows.Forms.Label();
			this.ugArticulos = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource3 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.txtBskArticulo = new System.Windows.Forms.TextBox();
			this.ugHabitaciones = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource2 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.ultraDataSource1 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.groupBox2 = new System.Windows.Forms.GroupBox();
			this.btnImprimir = new System.Windows.Forms.Button();
			this.txtTotDescuento = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label22 = new System.Windows.Forms.Label();
			this.label20 = new System.Windows.Forms.Label();
			this.txtBaseImponible = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label16 = new System.Windows.Forms.Label();
			this.rb2 = new System.Windows.Forms.RadioButton();
			this.rb1 = new System.Windows.Forms.RadioButton();
			this.label15 = new System.Windows.Forms.Label();
			this.txtDescnto = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.lblIdentificacion = new System.Windows.Forms.Label();
			this.cmbTipoRuc = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.btnElimina = new System.Windows.Forms.Button();
			this.button3 = new System.Windows.Forms.Button();
			this.button2 = new System.Windows.Forms.Button();
			this.btnAnular = new System.Windows.Forms.Button();
			this.txtNumHabitacion = new System.Windows.Forms.TextBox();
			this.label9 = new System.Windows.Forms.Label();
			this.cmbCajero = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.label19 = new System.Windows.Forms.Label();
			this.cmbCobrador = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.cmbVendedor = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.label18 = new System.Windows.Forms.Label();
			this.label17 = new System.Windows.Forms.Label();
			this.label13 = new System.Windows.Forms.Label();
			this.label10 = new System.Windows.Forms.Label();
			this.dtFecha = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.label11 = new System.Windows.Forms.Label();
			this.cmbFormaPago = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.txtTotal = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtIva = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtSubTotal = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
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
			this.cdsSeteo = new C1.Data.C1DataSet();
			this.toolTip1 = new System.Windows.Forms.ToolTip(this.components);
			this.groupBox1.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.cmbBodega)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource5)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ugArticulos)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource3)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ugHabitaciones)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).BeginInit();
			this.groupBox2.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.txtTotDescuento)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtBaseImponible)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDescnto)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbTipoRuc)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCajero)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCobrador)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbVendedor)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFecha)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbFormaPago)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtTotal)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIva)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtSubTotal)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ugDetalleFactura)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource4)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteo)).BeginInit();
			this.SuspendLayout();
			// 
			// groupBox1
			// 
			this.groupBox1.Controls.Add(this.lblBodega);
			this.groupBox1.Controls.Add(this.cmbBodega);
			this.groupBox1.Controls.Add(this.btBuscar);
			this.groupBox1.Controls.Add(this.label12);
			this.groupBox1.Controls.Add(this.ugArticulos);
			this.groupBox1.Controls.Add(this.txtBskArticulo);
			this.groupBox1.Controls.Add(this.ugHabitaciones);
			this.groupBox1.Location = new System.Drawing.Point(0, 8);
			this.groupBox1.Name = "groupBox1";
			this.groupBox1.Size = new System.Drawing.Size(552, 568);
			this.groupBox1.TabIndex = 0;
			this.groupBox1.TabStop = false;
			// 
			// lblBodega
			// 
			this.lblBodega.AutoSize = true;
			this.lblBodega.Location = new System.Drawing.Point(16, 18);
			this.lblBodega.Name = "lblBodega";
			this.lblBodega.Size = new System.Drawing.Size(41, 16);
			this.lblBodega.TabIndex = 251;
			this.lblBodega.Text = "LOCAL";
			this.lblBodega.TextAlign = System.Drawing.ContentAlignment.TopRight;
			// 
			// cmbBodega
			// 
			this.cmbBodega.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbBodega.Cursor = System.Windows.Forms.Cursors.Arrow;
			this.cmbBodega.DataSource = this.ultraDataSource5;
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridColumn3.Header.VisiblePosition = 2;
			ultraGridBand1.Columns.AddRange(new object[] {
																										 ultraGridColumn1,
																										 ultraGridColumn2,
																										 ultraGridColumn3});
			this.cmbBodega.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.cmbBodega.DisplayMember = "Nombre";
			this.cmbBodega.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbBodega.Enabled = false;
			this.cmbBodega.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbBodega.Location = new System.Drawing.Point(64, 16);
			this.cmbBodega.MaxDropDownItems = 30;
			this.cmbBodega.Name = "cmbBodega";
			this.cmbBodega.Size = new System.Drawing.Size(112, 21);
			this.cmbBodega.TabIndex = 250;
			this.cmbBodega.ValueMember = "idBodega";
			this.cmbBodega.ValueChanged += new System.EventHandler(this.cmbBodega_ValueChanged);
			this.cmbBodega.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.cmbBodega_InitializeLayout);
			// 
			// ultraDataSource5
			// 
			ultraDataColumn1.DataType = typeof(int);
			ultraDataColumn2.DataType = typeof(int);
			this.ultraDataSource5.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn1,
																																 ultraDataColumn2,
																																 ultraDataColumn3});
			// 
			// btBuscar
			// 
			this.btBuscar.Image = ((System.Drawing.Image)(resources.GetObject("btBuscar.Image")));
			this.btBuscar.Location = new System.Drawing.Point(312, 152);
			this.btBuscar.Name = "btBuscar";
			this.btBuscar.Size = new System.Drawing.Size(72, 22);
			this.btBuscar.TabIndex = 232;
			this.btBuscar.Text = "&Buscar";
			this.btBuscar.Click += new System.EventHandler(this.btBuscar_Click);
			// 
			// label12
			// 
			this.label12.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label12.Location = new System.Drawing.Point(8, 136);
			this.label12.Name = "label12";
			this.label12.Size = new System.Drawing.Size(64, 16);
			this.label12.TabIndex = 21;
			this.label12.Text = "ARTICULO";
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
			ultraGridColumn4.Header.VisiblePosition = 0;
			ultraGridColumn4.Hidden = true;
			ultraGridColumn5.Header.VisiblePosition = 1;
			ultraGridColumn5.Hidden = true;
			ultraGridColumn6.Header.VisiblePosition = 4;
			ultraGridColumn6.Hidden = true;
			ultraGridColumn7.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn7.Header.VisiblePosition = 3;
			ultraGridColumn7.Width = 410;
			ultraGridColumn8.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn8.Format = "#,##0.00";
			ultraGridColumn8.Header.Caption = "Precio";
			ultraGridColumn8.Header.VisiblePosition = 5;
			ultraGridColumn8.Width = 80;
			ultraGridColumn9.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn9.Format = "#,##0.00";
			ultraGridColumn9.Header.Caption = "Reserva";
			ultraGridColumn9.Header.VisiblePosition = 6;
			ultraGridColumn9.Hidden = true;
			ultraGridColumn9.Width = 80;
			ultraGridColumn10.ButtonDisplayStyle = Infragistics.Win.UltraWinGrid.ButtonDisplayStyle.Always;
			appearance2.Image = ((object)(resources.GetObject("appearance2.Image")));
			ultraGridColumn10.CellButtonAppearance = appearance2;
			ultraGridColumn10.Header.Caption = "Seleccionar";
			ultraGridColumn10.Header.VisiblePosition = 2;
			ultraGridColumn10.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.Button;
			ultraGridColumn10.Width = 20;
			ultraGridBand2.Columns.AddRange(new object[] {
																										 ultraGridColumn4,
																										 ultraGridColumn5,
																										 ultraGridColumn6,
																										 ultraGridColumn7,
																										 ultraGridColumn8,
																										 ultraGridColumn9,
																										 ultraGridColumn10});
			this.ugArticulos.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
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
			this.ugArticulos.Location = new System.Drawing.Point(8, 176);
			this.ugArticulos.Name = "ugArticulos";
			this.ugArticulos.Size = new System.Drawing.Size(536, 384);
			this.ugArticulos.TabIndex = 16;
			this.ugArticulos.DoubleClick += new System.EventHandler(this.ugArticulos_DoubleClick);
			this.ugArticulos.BeforeSelectChange += new Infragistics.Win.UltraWinGrid.BeforeSelectChangeEventHandler(this.ugArticulos_BeforeSelectChange);
			this.ugArticulos.ClickCellButton += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.ugArticulos_ClickCellButton);
			this.ugArticulos.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.ugArticulos_InitializeLayout);
			// 
			// ultraDataSource3
			// 
			ultraDataColumn4.DataType = typeof(int);
			ultraDataColumn5.DataType = typeof(int);
			ultraDataColumn6.DataType = typeof(int);
			ultraDataColumn8.DataType = typeof(System.Decimal);
			ultraDataColumn9.DataType = typeof(System.Decimal);
			this.ultraDataSource3.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn4,
																																 ultraDataColumn5,
																																 ultraDataColumn6,
																																 ultraDataColumn7,
																																 ultraDataColumn8,
																																 ultraDataColumn9});
			// 
			// txtBskArticulo
			// 
			this.txtBskArticulo.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtBskArticulo.Location = new System.Drawing.Point(8, 152);
			this.txtBskArticulo.Name = "txtBskArticulo";
			this.txtBskArticulo.Size = new System.Drawing.Size(296, 20);
			this.txtBskArticulo.TabIndex = 231;
			this.txtBskArticulo.Text = "";
			this.txtBskArticulo.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtBskArticulo_KeyPress);
			// 
			// ugHabitaciones
			// 
			this.ugHabitaciones.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.ugHabitaciones.Cursor = System.Windows.Forms.Cursors.Default;
			this.ugHabitaciones.DataSource = this.ultraDataSource2;
			appearance9.BackColor = System.Drawing.Color.White;
			this.ugHabitaciones.DisplayLayout.Appearance = appearance9;
			ultraGridColumn11.Header.VisiblePosition = 0;
			ultraGridColumn11.Hidden = true;
			ultraGridColumn12.Header.VisiblePosition = 1;
			ultraGridColumn12.Hidden = true;
			ultraGridColumn13.Header.Caption = "#";
			ultraGridColumn13.Header.VisiblePosition = 2;
			ultraGridColumn13.Width = 35;
			ultraGridColumn14.Header.VisiblePosition = 3;
			ultraGridColumn14.Width = 170;
			ultraGridColumn15.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
			appearance10.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn15.CellAppearance = appearance10;
			ultraGridColumn15.Format = "#,##0.00";
			ultraGridColumn15.Header.Caption = "Precio";
			ultraGridColumn15.Header.VisiblePosition = 4;
			ultraGridBand3.Columns.AddRange(new object[] {
																										 ultraGridColumn11,
																										 ultraGridColumn12,
																										 ultraGridColumn13,
																										 ultraGridColumn14,
																										 ultraGridColumn15});
			ultraGridBand3.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.Edit;
			this.ugHabitaciones.DisplayLayout.BandsSerializer.Add(ultraGridBand3);
			appearance11.ForeColor = System.Drawing.Color.Black;
			appearance11.ForeColorDisabled = System.Drawing.Color.Black;
			this.ugHabitaciones.DisplayLayout.Override.ActiveRowAppearance = appearance11;
			this.ugHabitaciones.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.TemplateOnBottom;
			this.ugHabitaciones.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.True;
			this.ugHabitaciones.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance12.BackColor = System.Drawing.Color.Transparent;
			this.ugHabitaciones.DisplayLayout.Override.CardAreaAppearance = appearance12;
			this.ugHabitaciones.DisplayLayout.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.RowSelect;
			appearance13.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance13.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance13.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance13.FontData.BoldAsString = "True";
			appearance13.FontData.Name = "Arial";
			appearance13.FontData.SizeInPoints = 10F;
			appearance13.ForeColor = System.Drawing.Color.White;
			appearance13.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.ugHabitaciones.DisplayLayout.Override.HeaderAppearance = appearance13;
			appearance14.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance14.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance14.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ugHabitaciones.DisplayLayout.Override.RowAlternateAppearance = appearance14;
			appearance15.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance15.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance15.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ugHabitaciones.DisplayLayout.Override.RowSelectorAppearance = appearance15;
			appearance16.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance16.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance16.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ugHabitaciones.DisplayLayout.Override.SelectedRowAppearance = appearance16;
			this.ugHabitaciones.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.ugHabitaciones.Location = new System.Drawing.Point(8, 40);
			this.ugHabitaciones.Name = "ugHabitaciones";
			this.ugHabitaciones.Size = new System.Drawing.Size(336, 88);
			this.ugHabitaciones.TabIndex = 16;
			this.ugHabitaciones.DoubleClick += new System.EventHandler(this.ugFacRelizadas_DoubleClick);
			// 
			// ultraDataSource2
			// 
			ultraDataColumn10.DataType = typeof(int);
			ultraDataColumn11.DataType = typeof(int);
			ultraDataColumn12.DataType = typeof(int);
			ultraDataColumn14.DataType = typeof(System.Decimal);
			this.ultraDataSource2.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn10,
																																 ultraDataColumn11,
																																 ultraDataColumn12,
																																 ultraDataColumn13,
																																 ultraDataColumn14});
			// 
			// ultraDataSource1
			// 
			ultraDataColumn15.DataType = typeof(int);
			ultraDataColumn16.DataType = typeof(int);
			this.ultraDataSource1.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn15,
																																 ultraDataColumn16,
																																 ultraDataColumn17});
			// 
			// groupBox2
			// 
			this.groupBox2.Controls.Add(this.btnImprimir);
			this.groupBox2.Controls.Add(this.txtTotDescuento);
			this.groupBox2.Controls.Add(this.label22);
			this.groupBox2.Controls.Add(this.label20);
			this.groupBox2.Controls.Add(this.txtBaseImponible);
			this.groupBox2.Controls.Add(this.label16);
			this.groupBox2.Controls.Add(this.rb2);
			this.groupBox2.Controls.Add(this.rb1);
			this.groupBox2.Controls.Add(this.label15);
			this.groupBox2.Controls.Add(this.txtDescnto);
			this.groupBox2.Controls.Add(this.lblIdentificacion);
			this.groupBox2.Controls.Add(this.cmbTipoRuc);
			this.groupBox2.Controls.Add(this.btnElimina);
			this.groupBox2.Controls.Add(this.button3);
			this.groupBox2.Controls.Add(this.button2);
			this.groupBox2.Controls.Add(this.btnAnular);
			this.groupBox2.Controls.Add(this.txtNumHabitacion);
			this.groupBox2.Controls.Add(this.label9);
			this.groupBox2.Controls.Add(this.cmbCajero);
			this.groupBox2.Controls.Add(this.label19);
			this.groupBox2.Controls.Add(this.cmbCobrador);
			this.groupBox2.Controls.Add(this.cmbVendedor);
			this.groupBox2.Controls.Add(this.label18);
			this.groupBox2.Controls.Add(this.label17);
			this.groupBox2.Controls.Add(this.label13);
			this.groupBox2.Controls.Add(this.label10);
			this.groupBox2.Controls.Add(this.dtFecha);
			this.groupBox2.Controls.Add(this.label11);
			this.groupBox2.Controls.Add(this.cmbFormaPago);
			this.groupBox2.Controls.Add(this.txtTotal);
			this.groupBox2.Controls.Add(this.txtIva);
			this.groupBox2.Controls.Add(this.txtSubTotal);
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
			this.groupBox2.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.groupBox2.Location = new System.Drawing.Point(552, 8);
			this.groupBox2.Name = "groupBox2";
			this.groupBox2.Size = new System.Drawing.Size(712, 568);
			this.groupBox2.TabIndex = 1;
			this.groupBox2.TabStop = false;
			this.toolTip1.SetToolTip(this.groupBox2, "Generar factura");
			this.groupBox2.Enter += new System.EventHandler(this.groupBox2_Enter);
			// 
			// btnImprimir
			// 
			this.btnImprimir.Enabled = false;
			this.btnImprimir.ForeColor = System.Drawing.Color.Coral;
			this.btnImprimir.Image = ((System.Drawing.Image)(resources.GetObject("btnImprimir.Image")));
			this.btnImprimir.Location = new System.Drawing.Point(192, 512);
			this.btnImprimir.Name = "btnImprimir";
			this.btnImprimir.Size = new System.Drawing.Size(56, 40);
			this.btnImprimir.TabIndex = 280;
			this.toolTip1.SetToolTip(this.btnImprimir, "Imprimir FACTURA");
			this.btnImprimir.Click += new System.EventHandler(this.button4_Click_1);
			// 
			// txtTotDescuento
			// 
			this.txtTotDescuento.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			appearance17.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtTotDescuento.Appearance = appearance17;
			this.txtTotDescuento.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtTotDescuento.Enabled = false;
			this.txtTotDescuento.FormatString = "#,##0.00";
			this.txtTotDescuento.Location = new System.Drawing.Point(440, 440);
			this.txtTotDescuento.Name = "txtTotDescuento";
			this.txtTotDescuento.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtTotDescuento.PromptChar = ' ';
			this.txtTotDescuento.Size = new System.Drawing.Size(80, 21);
			this.txtTotDescuento.TabIndex = 279;
			this.txtTotDescuento.Click += new System.EventHandler(this.txtTotDescuento_Click);
			this.txtTotDescuento.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtTotDescuento_KeyPress);
			this.txtTotDescuento.BeforeExitEditMode += new Infragistics.Win.BeforeExitEditModeEventHandler(this.txtTotDescuento_BeforeExitEditMode);
			this.txtTotDescuento.Enter += new System.EventHandler(this.txtTotDescuento_Enter);
			// 
			// label22
			// 
			this.label22.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label22.ForeColor = System.Drawing.Color.RoyalBlue;
			this.label22.Location = new System.Drawing.Point(304, 442);
			this.label22.Name = "label22";
			this.label22.Size = new System.Drawing.Size(128, 16);
			this.label22.TabIndex = 278;
			this.label22.Text = "Total DESCUENTO:";
			this.label22.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
			// 
			// label20
			// 
			this.label20.ForeColor = System.Drawing.Color.RoyalBlue;
			this.label20.Location = new System.Drawing.Point(520, 442);
			this.label20.Name = "label20";
			this.label20.Size = new System.Drawing.Size(80, 16);
			this.label20.TabIndex = 277;
			this.label20.Text = "DESCUENTO";
			this.label20.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
			// 
			// txtBaseImponible
			// 
			this.txtBaseImponible.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			appearance18.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtBaseImponible.Appearance = appearance18;
			this.txtBaseImponible.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtBaseImponible.Enabled = false;
			this.txtBaseImponible.FormatString = "#,##0.00";
			this.txtBaseImponible.Location = new System.Drawing.Point(608, 472);
			this.txtBaseImponible.Name = "txtBaseImponible";
			this.txtBaseImponible.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtBaseImponible.PromptChar = ' ';
			this.txtBaseImponible.Size = new System.Drawing.Size(80, 21);
			this.txtBaseImponible.TabIndex = 276;
			// 
			// label16
			// 
			this.label16.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label16.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(192)), ((System.Byte)(0)), ((System.Byte)(0)));
			this.label16.Location = new System.Drawing.Point(496, 474);
			this.label16.Name = "label16";
			this.label16.Size = new System.Drawing.Size(104, 16);
			this.label16.TabIndex = 275;
			this.label16.Text = "Base IMPONIBLE:";
			this.label16.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
			// 
			// rb2
			// 
			this.rb2.Location = new System.Drawing.Point(512, 160);
			this.rb2.Name = "rb2";
			this.rb2.Size = new System.Drawing.Size(32, 16);
			this.rb2.TabIndex = 274;
			this.rb2.Text = "2";
			// 
			// rb1
			// 
			this.rb1.Checked = true;
			this.rb1.Location = new System.Drawing.Point(480, 160);
			this.rb1.Name = "rb1";
			this.rb1.Size = new System.Drawing.Size(32, 16);
			this.rb1.TabIndex = 273;
			this.rb1.TabStop = true;
			this.rb1.Text = "1";
			// 
			// label15
			// 
			this.label15.AutoSize = true;
			this.label15.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label15.Location = new System.Drawing.Point(424, 160);
			this.label15.Name = "label15";
			this.label15.Size = new System.Drawing.Size(46, 16);
			this.label15.TabIndex = 272;
			this.label15.Text = "TURNO";
			// 
			// txtDescnto
			// 
			this.txtDescnto.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			appearance19.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtDescnto.Appearance = appearance19;
			this.txtDescnto.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtDescnto.Enabled = false;
			this.txtDescnto.FormatString = "#,##0.00";
			this.txtDescnto.Location = new System.Drawing.Point(608, 440);
			this.txtDescnto.Name = "txtDescnto";
			this.txtDescnto.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtDescnto.PromptChar = ' ';
			this.txtDescnto.Size = new System.Drawing.Size(80, 21);
			this.txtDescnto.TabIndex = 256;
			this.txtDescnto.Click += new System.EventHandler(this.txtDescnto_Click);
			this.txtDescnto.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtDescnto_KeyPress);
			this.txtDescnto.BeforeExitEditMode += new Infragistics.Win.BeforeExitEditModeEventHandler(this.txtDescnto_BeforeExitEditMode);
			this.txtDescnto.ValueChanged += new System.EventHandler(this.txtDescnto_ValueChanged);
			this.txtDescnto.PropertyChanged += new Infragistics.Win.PropertyChangedEventHandler(this.txtDescnto_PropertyChanged);
			this.txtDescnto.Enter += new System.EventHandler(this.txtDescnto_Enter);
			// 
			// lblIdentificacion
			// 
			this.lblIdentificacion.AutoSize = true;
			this.lblIdentificacion.Location = new System.Drawing.Point(8, 72);
			this.lblIdentificacion.Name = "lblIdentificacion";
			this.lblIdentificacion.Size = new System.Drawing.Size(87, 16);
			this.lblIdentificacion.TabIndex = 254;
			this.lblIdentificacion.Text = "Tipo Documento";
			// 
			// cmbTipoRuc
			// 
			appearance20.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbTipoRuc.Appearance = appearance20;
			this.cmbTipoRuc.CausesValidation = false;
			this.cmbTipoRuc.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbTipoRuc.Cursor = System.Windows.Forms.Cursors.Default;
			ultraGridColumn16.Header.VisiblePosition = 0;
			ultraGridColumn16.Hidden = true;
			ultraGridColumn17.Header.VisiblePosition = 1;
			ultraGridColumn17.Width = 233;
			ultraGridBand4.Columns.AddRange(new object[] {
																										 ultraGridColumn16,
																										 ultraGridColumn17});
			this.cmbTipoRuc.DisplayLayout.BandsSerializer.Add(ultraGridBand4);
			this.cmbTipoRuc.DisplayMember = "TipoRuc";
			this.cmbTipoRuc.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbTipoRuc.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbTipoRuc.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbTipoRuc.Location = new System.Drawing.Point(8, 88);
			this.cmbTipoRuc.Name = "cmbTipoRuc";
			this.cmbTipoRuc.Size = new System.Drawing.Size(128, 21);
			this.cmbTipoRuc.TabIndex = 253;
			this.cmbTipoRuc.ValueMember = "idTipoRuc";
			// 
			// btnElimina
			// 
			this.btnElimina.Enabled = false;
			this.btnElimina.Image = ((System.Drawing.Image)(resources.GetObject("btnElimina.Image")));
			this.btnElimina.Location = new System.Drawing.Point(592, 38);
			this.btnElimina.Name = "btnElimina";
			this.btnElimina.Size = new System.Drawing.Size(24, 23);
			this.btnElimina.TabIndex = 252;
			this.toolTip1.SetToolTip(this.btnElimina, "Eliminar FACTURA");
			this.btnElimina.Click += new System.EventHandler(this.btnElimina_Click);
			// 
			// button3
			// 
			this.button3.Image = ((System.Drawing.Image)(resources.GetObject("button3.Image")));
			this.button3.Location = new System.Drawing.Point(24, 512);
			this.button3.Name = "button3";
			this.button3.Size = new System.Drawing.Size(56, 40);
			this.button3.TabIndex = 251;
			this.toolTip1.SetToolTip(this.button3, "Nueva FACTURA");
			this.button3.Click += new System.EventHandler(this.button3_Click_1);
			// 
			// button2
			// 
			this.button2.ForeColor = System.Drawing.Color.Coral;
			this.button2.Image = ((System.Drawing.Image)(resources.GetObject("button2.Image")));
			this.button2.Location = new System.Drawing.Point(136, 512);
			this.button2.Name = "button2";
			this.button2.Size = new System.Drawing.Size(56, 40);
			this.button2.TabIndex = 250;
			this.toolTip1.SetToolTip(this.button2, "Buscar FACTURAS");
			this.button2.Click += new System.EventHandler(this.button2_Click_1);
			// 
			// btnAnular
			// 
			this.btnAnular.ForeColor = System.Drawing.Color.Coral;
			this.btnAnular.Image = ((System.Drawing.Image)(resources.GetObject("btnAnular.Image")));
			this.btnAnular.Location = new System.Drawing.Point(248, 512);
			this.btnAnular.Name = "btnAnular";
			this.btnAnular.Size = new System.Drawing.Size(56, 40);
			this.btnAnular.TabIndex = 249;
			this.toolTip1.SetToolTip(this.btnAnular, "Anular FACTURA");
			this.btnAnular.Click += new System.EventHandler(this.btnAnular_Click);
			// 
			// txtNumHabitacion
			// 
			this.txtNumHabitacion.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtNumHabitacion.Location = new System.Drawing.Point(128, 39);
			this.txtNumHabitacion.Name = "txtNumHabitacion";
			this.txtNumHabitacion.TabIndex = 248;
			this.txtNumHabitacion.Text = "";
			this.txtNumHabitacion.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
			// 
			// label9
			// 
			this.label9.Location = new System.Drawing.Point(128, 24);
			this.label9.Name = "label9";
			this.label9.Size = new System.Drawing.Size(80, 16);
			this.label9.TabIndex = 247;
			this.label9.Text = "# Habitación";
			// 
			// cmbCajero
			// 
			appearance21.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbCajero.Appearance = appearance21;
			this.cmbCajero.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbCajero.Cursor = System.Windows.Forms.Cursors.Default;
			ultraGridColumn18.Header.VisiblePosition = 0;
			ultraGridColumn18.Hidden = true;
			ultraGridColumn19.Header.VisiblePosition = 1;
			ultraGridColumn19.Width = 335;
			ultraGridColumn20.Header.VisiblePosition = 2;
			ultraGridBand5.Columns.AddRange(new object[] {
																										 ultraGridColumn18,
																										 ultraGridColumn19,
																										 ultraGridColumn20});
			this.cmbCajero.DisplayLayout.BandsSerializer.Add(ultraGridBand5);
			this.cmbCajero.DisplayMember = "Nombre";
			this.cmbCajero.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbCajero.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbCajero.Location = new System.Drawing.Point(256, 158);
			this.cmbCajero.Name = "cmbCajero";
			this.cmbCajero.Size = new System.Drawing.Size(152, 21);
			this.cmbCajero.TabIndex = 246;
			this.cmbCajero.ValueMember = "idPersonal";
			// 
			// label19
			// 
			this.label19.AutoSize = true;
			this.label19.Location = new System.Drawing.Point(216, 160);
			this.label19.Name = "label19";
			this.label19.Size = new System.Drawing.Size(38, 16);
			this.label19.TabIndex = 245;
			this.label19.Text = "Cajero";
			// 
			// cmbCobrador
			// 
			appearance22.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbCobrador.Appearance = appearance22;
			this.cmbCobrador.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbCobrador.Cursor = System.Windows.Forms.Cursors.Default;
			ultraGridColumn21.Header.VisiblePosition = 0;
			ultraGridColumn21.Hidden = true;
			ultraGridColumn22.Header.VisiblePosition = 1;
			ultraGridColumn22.Width = 335;
			ultraGridColumn23.Header.VisiblePosition = 2;
			ultraGridBand6.Columns.AddRange(new object[] {
																										 ultraGridColumn21,
																										 ultraGridColumn22,
																										 ultraGridColumn23});
			this.cmbCobrador.DisplayLayout.BandsSerializer.Add(ultraGridBand6);
			this.cmbCobrador.DisplayMember = "Nombre";
			this.cmbCobrador.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbCobrador.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbCobrador.Location = new System.Drawing.Point(56, 158);
			this.cmbCobrador.Name = "cmbCobrador";
			this.cmbCobrador.Size = new System.Drawing.Size(152, 21);
			this.cmbCobrador.TabIndex = 244;
			this.cmbCobrador.ValueMember = "idPersonal";
			// 
			// cmbVendedor
			// 
			appearance23.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbVendedor.Appearance = appearance23;
			this.cmbVendedor.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbVendedor.Cursor = System.Windows.Forms.Cursors.Default;
			ultraGridColumn24.Header.VisiblePosition = 0;
			ultraGridColumn24.Hidden = true;
			ultraGridColumn25.Header.VisiblePosition = 1;
			ultraGridColumn25.Width = 335;
			ultraGridColumn26.Header.VisiblePosition = 2;
			ultraGridBand7.Columns.AddRange(new object[] {
																										 ultraGridColumn24,
																										 ultraGridColumn25,
																										 ultraGridColumn26});
			this.cmbVendedor.DisplayLayout.BandsSerializer.Add(ultraGridBand7);
			this.cmbVendedor.DisplayMember = "Nombre";
			this.cmbVendedor.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbVendedor.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbVendedor.Location = new System.Drawing.Point(232, 39);
			this.cmbVendedor.Name = "cmbVendedor";
			this.cmbVendedor.Size = new System.Drawing.Size(116, 21);
			this.cmbVendedor.TabIndex = 243;
			this.cmbVendedor.ValueMember = "idPersonal";
			// 
			// label18
			// 
			this.label18.AutoSize = true;
			this.label18.Location = new System.Drawing.Point(8, 160);
			this.label18.Name = "label18";
			this.label18.Size = new System.Drawing.Size(51, 16);
			this.label18.TabIndex = 242;
			this.label18.Text = "Cobrador";
			// 
			// label17
			// 
			this.label17.AutoSize = true;
			this.label17.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label17.Location = new System.Drawing.Point(232, 24);
			this.label17.Name = "label17";
			this.label17.Size = new System.Drawing.Size(46, 16);
			this.label17.TabIndex = 240;
			this.label17.Text = "Guardia";
			// 
			// label13
			// 
			this.label13.AutoSize = true;
			this.label13.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label13.ForeColor = System.Drawing.Color.Red;
			this.label13.Location = new System.Drawing.Point(16, 448);
			this.label13.Name = "label13";
			this.label13.Size = new System.Drawing.Size(78, 22);
			this.label13.TabIndex = 231;
			this.label13.Text = "NORMAL";
			// 
			// label10
			// 
			this.label10.AutoSize = true;
			this.label10.Font = new System.Drawing.Font("Microsoft Sans Serif", 8F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label10.Location = new System.Drawing.Point(16, 24);
			this.label10.Name = "label10";
			this.label10.Size = new System.Drawing.Size(54, 16);
			this.label10.TabIndex = 230;
			this.label10.Text = "Fecha FC";
			// 
			// dtFecha
			// 
			appearance24.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtFecha.Appearance = appearance24;
			this.dtFecha.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton1.Caption = "Today";
			this.dtFecha.DateButtons.Add(dateButton1);
			this.dtFecha.Format = "dd/MM/yyyy";
			this.dtFecha.Location = new System.Drawing.Point(16, 39);
			this.dtFecha.Name = "dtFecha";
			this.dtFecha.NonAutoSizeHeight = 23;
			this.dtFecha.Size = new System.Drawing.Size(104, 21);
			this.dtFecha.SpinButtonsVisible = true;
			this.dtFecha.TabIndex = 229;
			this.dtFecha.Value = ((object)(resources.GetObject("dtFecha.Value")));
			// 
			// label11
			// 
			this.label11.Location = new System.Drawing.Point(544, 112);
			this.label11.Name = "label11";
			this.label11.Size = new System.Drawing.Size(88, 16);
			this.label11.TabIndex = 228;
			this.label11.Text = "Forma de Pago";
			// 
			// cmbFormaPago
			// 
			appearance25.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbFormaPago.Appearance = appearance25;
			this.cmbFormaPago.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbFormaPago.Cursor = System.Windows.Forms.Cursors.Default;
			ultraGridColumn27.Header.VisiblePosition = 0;
			ultraGridColumn27.Hidden = true;
			ultraGridColumn28.Header.VisiblePosition = 1;
			ultraGridColumn28.Width = 152;
			ultraGridBand8.Columns.AddRange(new object[] {
																										 ultraGridColumn27,
																										 ultraGridColumn28});
			this.cmbFormaPago.DisplayLayout.BandsSerializer.Add(ultraGridBand8);
			this.cmbFormaPago.DisplayMember = "FormaPago";
			this.cmbFormaPago.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbFormaPago.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbFormaPago.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbFormaPago.Location = new System.Drawing.Point(544, 128);
			this.cmbFormaPago.Name = "cmbFormaPago";
			this.cmbFormaPago.Size = new System.Drawing.Size(144, 21);
			this.cmbFormaPago.TabIndex = 226;
			this.cmbFormaPago.ValueMember = "idFormaPago";
			// 
			// txtTotal
			// 
			this.txtTotal.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			appearance26.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtTotal.Appearance = appearance26;
			this.txtTotal.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtTotal.Enabled = false;
			this.txtTotal.Font = new System.Drawing.Font("Microsoft Sans Serif", 14F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtTotal.FormatString = "#,##0.00";
			this.txtTotal.Location = new System.Drawing.Point(608, 518);
			this.txtTotal.Name = "txtTotal";
			this.txtTotal.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtTotal.PromptChar = ' ';
			this.txtTotal.Size = new System.Drawing.Size(80, 31);
			this.txtTotal.TabIndex = 217;
			// 
			// txtIva
			// 
			this.txtIva.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			appearance27.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtIva.Appearance = appearance27;
			this.txtIva.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtIva.Enabled = false;
			this.txtIva.FormatString = "#,##0.00";
			this.txtIva.Location = new System.Drawing.Point(608, 495);
			this.txtIva.Name = "txtIva";
			this.txtIva.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtIva.PromptChar = ' ';
			this.txtIva.Size = new System.Drawing.Size(80, 21);
			this.txtIva.TabIndex = 216;
			// 
			// txtSubTotal
			// 
			this.txtSubTotal.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			appearance28.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtSubTotal.Appearance = appearance28;
			this.txtSubTotal.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtSubTotal.Enabled = false;
			this.txtSubTotal.FormatString = "#,##0.00";
			this.txtSubTotal.Location = new System.Drawing.Point(608, 416);
			this.txtSubTotal.Name = "txtSubTotal";
			this.txtSubTotal.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtSubTotal.PromptChar = ' ';
			this.txtSubTotal.Size = new System.Drawing.Size(80, 21);
			this.txtSubTotal.TabIndex = 215;
			// 
			// bntGrabar
			// 
			this.bntGrabar.Image = ((System.Drawing.Image)(resources.GetObject("bntGrabar.Image")));
			this.bntGrabar.Location = new System.Drawing.Point(80, 512);
			this.bntGrabar.Name = "bntGrabar";
			this.bntGrabar.Size = new System.Drawing.Size(56, 40);
			this.bntGrabar.TabIndex = 35;
			this.toolTip1.SetToolTip(this.bntGrabar, "Grabar FACTURA");
			this.bntGrabar.Click += new System.EventHandler(this.bntGrabar_Click);
			// 
			// label8
			// 
			this.label8.Font = new System.Drawing.Font("Microsoft Sans Serif", 14F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label8.Location = new System.Drawing.Point(520, 526);
			this.label8.Name = "label8";
			this.label8.Size = new System.Drawing.Size(80, 16);
			this.label8.TabIndex = 33;
			this.label8.Text = "TOTAL";
			this.label8.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
			// 
			// label7
			// 
			this.label7.Location = new System.Drawing.Point(520, 497);
			this.label7.Name = "label7";
			this.label7.Size = new System.Drawing.Size(80, 16);
			this.label7.TabIndex = 31;
			this.label7.Text = "I.V.A.";
			this.label7.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
			// 
			// label6
			// 
			this.label6.Location = new System.Drawing.Point(520, 418);
			this.label6.Name = "label6";
			this.label6.Size = new System.Drawing.Size(80, 16);
			this.label6.TabIndex = 29;
			this.label6.Text = "SUB-TOTAL";
			this.label6.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
			// 
			// ugDetalleFactura
			// 
			this.ugDetalleFactura.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.ugDetalleFactura.Cursor = System.Windows.Forms.Cursors.Default;
			this.ugDetalleFactura.DataSource = this.ultraDataSource4;
			appearance29.BackColor = System.Drawing.Color.White;
			this.ugDetalleFactura.DisplayLayout.Appearance = appearance29;
			ultraGridColumn29.Header.VisiblePosition = 0;
			ultraGridColumn29.Hidden = true;
			ultraGridColumn30.Header.VisiblePosition = 1;
			ultraGridColumn30.Hidden = true;
			ultraGridColumn31.Header.VisiblePosition = 2;
			ultraGridColumn31.Hidden = true;
			ultraGridColumn32.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
			ultraGridColumn32.Header.VisiblePosition = 3;
			ultraGridColumn32.Width = 380;
			appearance30.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn33.CellAppearance = appearance30;
			ultraGridColumn33.Format = "#,##0";
			ultraGridColumn33.Header.VisiblePosition = 4;
			ultraGridColumn33.PromptChar = ' ';
			ultraGridColumn33.Width = 50;
			ultraGridColumn34.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
			appearance31.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn34.CellAppearance = appearance31;
			ultraGridColumn34.Format = "#,##0.000000";
			ultraGridColumn34.Header.VisiblePosition = 5;
			ultraGridColumn34.PromptChar = ' ';
			ultraGridColumn34.Width = 96;
			ultraGridColumn35.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance32.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn35.CellAppearance = appearance32;
			ultraGridColumn35.Format = "#,##0.00";
			ultraGridColumn35.Header.Caption = "%Desc";
			ultraGridColumn35.Header.VisiblePosition = 6;
			ultraGridColumn35.Width = 45;
			ultraGridColumn36.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
			appearance33.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn36.CellAppearance = appearance33;
			ultraGridColumn36.Format = "#,##0.000000";
			ultraGridColumn36.Header.Caption = "Subtotal";
			ultraGridColumn36.Header.VisiblePosition = 7;
			ultraGridColumn36.PromptChar = ' ';
			ultraGridColumn36.Width = 96;
			ultraGridColumn37.Header.VisiblePosition = 8;
			ultraGridColumn37.Hidden = true;
			ultraGridBand9.Columns.AddRange(new object[] {
																										 ultraGridColumn29,
																										 ultraGridColumn30,
																										 ultraGridColumn31,
																										 ultraGridColumn32,
																										 ultraGridColumn33,
																										 ultraGridColumn34,
																										 ultraGridColumn35,
																										 ultraGridColumn36,
																										 ultraGridColumn37});
			ultraGridBand9.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.Edit;
			this.ugDetalleFactura.DisplayLayout.BandsSerializer.Add(ultraGridBand9);
			appearance34.ForeColor = System.Drawing.Color.Black;
			appearance34.ForeColorDisabled = System.Drawing.Color.Black;
			this.ugDetalleFactura.DisplayLayout.Override.ActiveRowAppearance = appearance34;
			this.ugDetalleFactura.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.TemplateOnBottom;
			this.ugDetalleFactura.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.True;
			this.ugDetalleFactura.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance35.BackColor = System.Drawing.Color.Transparent;
			this.ugDetalleFactura.DisplayLayout.Override.CardAreaAppearance = appearance35;
			this.ugDetalleFactura.DisplayLayout.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.Edit;
			appearance36.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance36.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance36.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance36.FontData.BoldAsString = "True";
			appearance36.FontData.Name = "Arial";
			appearance36.FontData.SizeInPoints = 10F;
			appearance36.ForeColor = System.Drawing.Color.White;
			appearance36.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.ugDetalleFactura.DisplayLayout.Override.HeaderAppearance = appearance36;
			appearance37.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance37.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance37.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ugDetalleFactura.DisplayLayout.Override.RowAlternateAppearance = appearance37;
			appearance38.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance38.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance38.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ugDetalleFactura.DisplayLayout.Override.RowSelectorAppearance = appearance38;
			appearance39.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance39.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance39.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ugDetalleFactura.DisplayLayout.Override.SelectedRowAppearance = appearance39;
			this.ugDetalleFactura.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.ugDetalleFactura.Location = new System.Drawing.Point(8, 184);
			this.ugDetalleFactura.Name = "ugDetalleFactura";
			this.ugDetalleFactura.Size = new System.Drawing.Size(696, 224);
			this.ugDetalleFactura.TabIndex = 28;
			this.ugDetalleFactura.AfterRowsDeleted += new System.EventHandler(this.ugDetalleFactura_AfterRowsDeleted);
			this.ugDetalleFactura.BeforeRowsDeleted += new Infragistics.Win.UltraWinGrid.BeforeRowsDeletedEventHandler(this.ugDetalleFactura_BeforeRowsDeleted);
			this.ugDetalleFactura.KeyDown += new System.Windows.Forms.KeyEventHandler(this.ugDetalleFactura_KeyDown);
			this.ugDetalleFactura.AfterCellUpdate += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.ugDetalleFactura_AfterCellUpdate);
			// 
			// ultraDataSource4
			// 
			ultraDataColumn18.DataType = typeof(int);
			ultraDataColumn19.DataType = typeof(int);
			ultraDataColumn20.DataType = typeof(int);
			ultraDataColumn22.DataType = typeof(int);
			ultraDataColumn23.DataType = typeof(System.Decimal);
			ultraDataColumn24.DataType = typeof(System.Decimal);
			ultraDataColumn25.DataType = typeof(System.Decimal);
			ultraDataColumn26.DataType = typeof(System.Decimal);
			this.ultraDataSource4.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn18,
																																 ultraDataColumn19,
																																 ultraDataColumn20,
																																 ultraDataColumn21,
																																 ultraDataColumn22,
																																 ultraDataColumn23,
																																 ultraDataColumn24,
																																 ultraDataColumn25,
																																 ultraDataColumn26});
			// 
			// txtTelefono
			// 
			this.txtTelefono.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtTelefono.Location = new System.Drawing.Point(440, 128);
			this.txtTelefono.Name = "txtTelefono";
			this.txtTelefono.TabIndex = 27;
			this.txtTelefono.Text = "";
			// 
			// label5
			// 
			this.label5.Location = new System.Drawing.Point(440, 112);
			this.label5.Name = "label5";
			this.label5.Size = new System.Drawing.Size(80, 16);
			this.label5.TabIndex = 26;
			this.label5.Text = "# TELEFONO";
			// 
			// txtDireccion
			// 
			this.txtDireccion.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtDireccion.Location = new System.Drawing.Point(8, 128);
			this.txtDireccion.Name = "txtDireccion";
			this.txtDireccion.Size = new System.Drawing.Size(424, 20);
			this.txtDireccion.TabIndex = 25;
			this.txtDireccion.Text = "";
			// 
			// label4
			// 
			this.label4.Location = new System.Drawing.Point(8, 112);
			this.label4.Name = "label4";
			this.label4.Size = new System.Drawing.Size(80, 16);
			this.label4.TabIndex = 24;
			this.label4.Text = "DIRECCION";
			// 
			// txtNombre
			// 
			this.txtNombre.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtNombre.Location = new System.Drawing.Point(248, 88);
			this.txtNombre.Name = "txtNombre";
			this.txtNombre.Size = new System.Drawing.Size(296, 20);
			this.txtNombre.TabIndex = 23;
			this.txtNombre.Text = "";
			// 
			// label3
			// 
			this.label3.Location = new System.Drawing.Point(248, 72);
			this.label3.Name = "label3";
			this.label3.Size = new System.Drawing.Size(64, 16);
			this.label3.TabIndex = 22;
			this.label3.Text = "NOMBRES";
			// 
			// txtNumCedula
			// 
			this.txtNumCedula.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtNumCedula.Location = new System.Drawing.Point(144, 88);
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
			this.label2.Location = new System.Drawing.Point(144, 72);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(64, 16);
			this.label2.TabIndex = 20;
			this.label2.Text = "# CEDULA";
			// 
			// button1
			// 
			this.button1.Image = ((System.Drawing.Image)(resources.GetObject("button1.Image")));
			this.button1.Location = new System.Drawing.Point(544, 38);
			this.button1.Name = "button1";
			this.button1.Size = new System.Drawing.Size(24, 23);
			this.button1.TabIndex = 19;
			this.toolTip1.SetToolTip(this.button1, "Registrar Número de Factura");
			this.button1.Click += new System.EventHandler(this.button1_Click);
			// 
			// txtNumFactura
			// 
			this.txtNumFactura.Location = new System.Drawing.Point(448, 39);
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
			this.label1.Location = new System.Drawing.Point(376, 41);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(72, 16);
			this.label1.TabIndex = 17;
			this.label1.Text = "# FACTURA";
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
			// frmFacturacion
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(1272, 582);
			this.Controls.Add(this.groupBox2);
			this.Controls.Add(this.groupBox1);
			this.Name = "frmFacturacion";
			this.Text = "Facuración HABITACIONES";
			this.Closing += new System.ComponentModel.CancelEventHandler(this.frmFacturacion_Closing);
			this.Load += new System.EventHandler(this.frmFacturacion_Load);
			this.groupBox1.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.cmbBodega)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource5)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ugArticulos)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource3)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ugHabitaciones)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).EndInit();
			this.groupBox2.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.txtTotDescuento)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtBaseImponible)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDescnto)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbTipoRuc)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCajero)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCobrador)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbVendedor)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFecha)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbFormaPago)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtTotal)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIva)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtSubTotal)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ugDetalleFactura)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource4)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteo)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void frmFacturacion_Load(object sender, System.EventArgs e)
		{
			/* Porcentaje del IVA */
			//string lcPiva = "SELECT Iva FROM SeteoF WHERE idSeteoF=1";
			//string lcPiva = "SELECT dbo.RetornaIVAVigenteLote()";
			//PorcentajeIva = Funcion.decEscalarSQL(cdsSeteo, lcPiva); 

			/* Tiempo de habitación */
			string nHoras = "SELECT dbo.Tmpo_Habitaciones()";
			//this.txtTiempo.Value = Funcion.iEscalarSQL(cdsSeteo, nHoras); 
			this.txtNumCedula.Text = "9999999999";
			this.dtFecha.Value = DateTime.Today;

			/* Porcentaje del IVA */
			PorcentajeIva = Funcion.decEscalarSQL(cdsSeteo, string.Format("Select dbo.RetornaImpuestoIVAVigente('{0}', 1)", 
				Convert.ToDateTime(this.dtFecha.Value).ToString("yyyyMMdd")));

			//Buscar habitaciones con sus diferentes estados
			string stTipo = "0703";
			string stPermiso = "118";

			//Carga Vendedores, Cajero y Cobrador
			this.cmbVendedor.DataSource = Funcion.dvProcedimiento(cdsSeteo, string.Format("Exec PersonalLotte {0}", NBodga));
			this.cmbCobrador.DataSource = Funcion.dvProcedimiento(cdsSeteo, string.Format("Exec PersonalLotte_Cobrador {0}", NBodga));
			this.cmbCajero.DataSource = Funcion.dvProcedimiento(cdsSeteo, string.Format("Exec PersonalLotte_Cajero {0}", NBodga));
			string stExecBodega = string.Format("Select idBodega, Bodega, Nombre From Bodega Where Activo = 1");
			this.cmbBodega.DataSource = Funcion.dvProcedimiento(cdsSeteo, stExecBodega);
			this.cmbTipoRuc.DataSource = Funcion.dvProcedimiento(cdsSeteo, "Exec ClienteTipoRucCarga 1");
			this.cmbTipoRuc.Value = 5;

			miAcceso = new Acceso(cdsSeteo, stTipo);
			
			this.cmbFormaPago.Enabled = true;
			this.bntGrabar.Enabled = true;
			this.btnAnular.Enabled = false;

			//string lcQryBod = "SELECT Bodega From Bodega Where idBodega in (Select idSucursal From Habitaciones Where idHabitacion="+idNhab.ToString()+")";
			//int laBodega = Funcion.iEscalarSQL(cdsSeteo, lcQryBod);

			//NBodga = laBodega;

			int codBodHabtcn = Funcion.iEscalarSQL(cdsSeteo, "SELECT Bodega FROM UsuariosBodegas WHERE idUsuario="+MenuLatinium.IdUsuario.ToString());
			int BodaBusk = Funcion.iEscalarSQL(cdsSeteo, "SELECT idBodega FROM Bodega WHERE Bodega="+codBodHabtcn.ToString());
			this.cmbBodega.Value = BodaBusk;

			this.cmbFormaPago.DataSource = Funcion.dvProcedimiento(cdsSeteo, "Exec CompraFPagoCarga_LT");
			niUsuario = MenuLatinium.iNivel.ToString().Trim();
			if(niUsuario.Equals("4") || niUsuario.Equals("1") || niUsuario.Equals("2") || niUsuario.Equals("3")) 
			{
				this.ugDetalleFactura.DisplayLayout.Bands[0].Columns["Precio"].CellActivation = Activation.AllowEdit;
				this.txtTotDescuento.Enabled = true;
				this.cmbBodega.Enabled = true;
				this.btnElimina.Enabled = true;
			}

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
			//Carga articulos de acuerso al subgrupo
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
					//int TmpoHab = (int)this.txtTiempo.Value;
					string lcQry = "SELECT IsNull(Numero,'') as NumFac FROM Compra WHERE numero like "+"'%"+this.txtNumFactura.Text.ToString()+"'"+" And idTipoFactura=1 And Bodega="+NBodga.ToString();
					string nFacHab = Funcion.sEscalarSQL(cdsSeteo, lcQry);
					int valfc = 0;
					if (nFacHab!=null) valfc = Int32.Parse(nFacHab);
					int vlrfc = Int32.Parse(this.txtNumFactura.Text.Trim());
					int idGuadr = (int)this.cmbVendedor.Value;
					if (valfc!=vlrfc)
					{
            string SP_Qry = string.Format("Exec FacturaLotteServicio '{0}',{1},{2},{3},{4},{5},{6},{7}", 
							this.txtNumFactura.Text,
							Int32.Parse(this.txtNumHabitacion.Text.ToString()),
							idNhab,0,0,5,idGuadr,1);
						//Funcion.EjecutaSQL(cdsSeteo,SP_Qry);
						idCpaFC = (int)Funcion.iEscalarSQL(cdsSeteo,SP_Qry);
						Registra_FC(idNhab);
						RegistroFactura = 1;
						//Visualiza detalle factura
						string QryDetFc = "SELECT idDetCompra,idCompra,idArticulo,Articulo,Cantidad,Precio,DescuentoPorc,SubTotal,IsNull(DescGlobal,0) as DescGlobal From DetCompra Where idCompra="+idCpaFC.ToString();
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
			if (e.KeyChar == 13) this.txtNumFactura.Focus();
			//Funcion.CamposNumericos(sender,e);
		}

		public void Registra_FC(int NumrHab)
		{
			string Qry = string.Format("Select idCompra,Ruc,Numero,Fecha From Compra Where idCliente=0 And Estado=0 And idRecibe="+NumrHab.ToString()+" Order By idCompra");
			//ugFacRelizadas.DataSource = Funcion.dvProcedimiento(cdsSeteo,Qry);
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
			decimal ValDsco = 0.00m;
			decimal ToDscto = 0.00m;

			#region Calculo Forma Pago Mixto 
			decimal dSubTotal1 = 0.00m;
			decimal dSubTotal2 = 0.00m;
			decimal dSubTotal3 = 0.00m;
			decimal dSubTotal4 = 0.00m;
			#endregion Calculo Forma Pago Mixto 

			#endregion Creacion de Variables
			
			/* Porcentaje del IVA de acuerdo a la fecha */
			DateTime fechaF = Convert.ToDateTime("31-05-2016").Date;
			DateTime fechaFC = (DateTime)this.dtFecha.Value;
			if (fechaFC <= fechaF) PorcentajeIva = 12;
			else
			{
				//string lcPiva = "SELECT dbo.RetornaImpuestoIVAVigente()";
				//PorcentajeIva = Funcion.decEscalarSQL(cdsSeteo, lcPiva);
				PorcentajeIva = Funcion.decEscalarSQL(cdsSeteo, string.Format("Select dbo.RetornaImpuestoIVAVigente('{0}', 1)", 
					Convert.ToDateTime(this.dtFecha.Value).ToString("yyyyMMdd")));
			}

			this.txtIva.Value = 0.000000m;
			this.txtSubTotal.Value = 0.000000m;
			this.txtTotal.Value = 0.00m;
			
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
							if (Convert.ToDecimal(dr.Cells["DescuentoPorc"].Value)>0)
							{
								dr.Cells["Subtotal"].Value = Math.Round(Convert.ToDecimal(dr.Cells["Cantidad"].Value) *(Convert.ToDecimal(dr.Cells["Precio"].Value)*(Convert.ToDecimal(dr.Cells["DescuentoPorc"].Value)/100)), 6);
								ToDscto = ToDscto + Convert.ToDecimal(dr.Cells["Precio"].Value)*(Convert.ToDecimal(dr.Cells["DescuentoPorc"].Value)/100);
							}
							else dr.Cells["Subtotal"].Value = Math.Round(Convert.ToDecimal(dr.Cells["Cantidad"].Value) * Convert.ToDecimal(dr.Cells["Precio"].Value), 6);
						}
						//dr.Cells["Subtotal"].Value = Math.Round(Convert.ToInt32(dr.Cells["Cantidad"].Value) * Convert.ToDecimal(dr.Cells["Precio"].Value), 6);
						//this.txtSubTotal.Value = (decimal)this.txtSubTotal.Value + Math.Round(Convert.ToDecimal(dr.Cells["SubTotal"].Value), 6);
						//Actualiza la cantidad de producto
						//string Updt = string.Format("Exec Actualiza_DetCpa_Lotte {0},{1}", dr.Cells["idDetCompra"].Value.ToString(),dr.Cells["Cantidad"].Value.ToString());
						//Funcion.EjecutaSQL(cdsSeteo,Updt);
					}
					else dr.Cells["Subtotal"].Value = Math.Round(Convert.ToDecimal(dr.Cells["Cantidad"].Value) * Convert.ToDecimal(dr.Cells["Precio"].Value), 6);
					this.txtSubTotal.Value = (decimal)this.txtSubTotal.Value+(decimal)dr.Cells["Subtotal"].Value;
				}
				if (dr.Cells["DescGlobal"].Value != DBNull.Value && Convert.ToDecimal(dr.Cells["DescGlobal"].Value)>0)
				{
					this.txtTotDescuento.Value = Convert.ToDecimal(dr.Cells["DescGlobal"].Value);
				}
			}
			//this.txtIva.Value = Math.Round((Convert.ToDecimal(this.txtSubTotal.Value) * PorcentajeIva /100.00m), 2);
			//Se actualiza valor de descuento en valor de habitación y no total de factura
			if(idGrpClte==14) ValDsco = Convert.ToDecimal(this.txtSubTotal.Value)-Convert.ToDecimal(this.txtDescnto.Value);
			else ValDsco = Convert.ToDecimal(this.txtSubTotal.Value)-0;
			this.txtDescnto.Value = ToDscto;

			if (Convert.ToDecimal(this.txtTotDescuento.Value) > 0)	this.txtBaseImponible.Value = Math.Round(Convert.ToDecimal(this.txtSubTotal.Value)-Convert.ToDecimal(this.txtDescnto.Value),6);
			else this.txtBaseImponible.Value = Math.Round(Convert.ToDecimal(this.txtSubTotal.Value),6);
			this.txtIva.Value = Math.Round((Convert.ToDecimal(this.txtBaseImponible.Value) * PorcentajeIva /100.00m), 6);

//			if(idGrpClte==14) this.txtTotal.Value = (Math.Round(Convert.ToDecimal(this.txtSubTotal.Value), 6) - Math.Round(Convert.ToDecimal(this.txtDescnto.Value), 2))
//                         + Math.Round(Convert.ToDecimal(this.txtIva.Value), 6);
//			else this.txtTotal.Value = Math.Round(Convert.ToDecimal(this.txtSubTotal.Value), 6) + Math.Round(Convert.ToDecimal(this.txtIva.Value), 6);
			this.txtTotal.Value = Math.Round(Convert.ToDecimal(this.txtBaseImponible.Value), 6) + Math.Round(Convert.ToDecimal(this.txtIva.Value), 6);
			TotalDeetalleFactura = Math.Round(Convert.ToDecimal(this.txtTotal.Value), 2);
			#endregion Desplazamiento de grilla			
		}

		private void txtNumCedula_Validated(object sender, System.EventArgs e)
		{
			#region Busca Cliente 
			/* Buscar Número de Bodega */
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
				idGrpClte = (int)drC.GetValue(3);
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
						idGrpClte = miBusqueda.idGrupoCliente;
						this.txtNumCedula.Text = miBusqueda.Ruc;
					}
				}
			}
		}

		private void bntGrabar_Click(object sender, System.EventArgs e)
		{
			/* Grabar Factura Habitación */
			DateTime dtFM2015 = DateTime.Parse("01/01/2016");
			int result = DateTime.Compare((DateTime)this.dtFecha.Value, dtFM2015);

			if(this.txtNumCedula.Text.ToString().Length==0)
			{
				MessageBox.Show("Debe Registrar CLIENTE para poder Grabar la FACTURA.?","Aviso....!!!", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.txtNumCedula.Focus();
			}
			else if (this.txtNumFactura.Text.ToString().Length==0)
			{
				MessageBox.Show("Debe Registrar # DE FACTURA.?","Aviso....!!!", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.txtNumFactura.Focus();
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
			else if(this.cmbFormaPago.ActiveRow == null)
			{
				MessageBox.Show("No se a seleccionado FORMA DE PAGO?","Aviso....!!!", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.cmbFormaPago.Focus();
			}
			else if(this.cmbCobrador.ActiveRow == null)
			{
				MessageBox.Show("No se a seleccionado Cobrador?","Aviso....!!!", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.cmbCobrador.Focus();
			}
			else if(result < 0)
			{
				MessageBox.Show("No se puede MODIFICAR Facturas con fecha menor a 01/01/2016?","Aviso....!!!", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.button2.Focus();
			}
			else if(idClint ==0) 
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

					/*Registrar razon descuento factura*/
					if(Convert.ToDecimal(this.txtDescnto.Value) > 0)
					{
						using (frmJustDescuento ntaDscto = new frmJustDescuento())
						{
							if (DialogResult.OK == ntaDscto.ShowDialog())
							{
								try
								{
									DatosDescuento = ntaDscto.txtNotas.Text.ToString();
									//string sSQLAnula = string.Format("Exec NotaDescuento {0}, '{1}'", idCpaFC, ntaDscto.txtNotas.Text.ToString());
									//Funcion.EjecutaSQL(cdsSeteo, sSQLAnula);
								}
								catch(Exception Exc)
								{
									MessageBox.Show(string.Format("Error al anular: {0}", Exc), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
								}
							}
						}
				  }

					using (VentaFPagoLot miCobro = new VentaFPagoLot(idCpaFC, (int)this.cmbFormaPago.Value, (decimal)this.txtTotal.Value, NumFactura, false, 0, 0, false, 0, false, 0, false, 0, 0, 0, idClint, 0, 0, 0, 0, 0))
					//using (VentaFormaPagoLotte miCobro = new VentaFormaPagoLotte(idCpaFC, 
					//				 (int)this.cmbFormaPago.Value, 
					//				 (decimal)this.txtTotal.Value, 
					//				 this.txtNumFactura.Text.ToString(), false, 0, 0, false, 0, false, 0, false, 0, 0, 5, idClint, 0, 0, 0, 0, 
					//				 (int)this.cmbBodega.Value, this.ugDetalleFactura))
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
									string FechaFC = this.txtNumFactura.Text.ToString(); //this.ugFacRelizadas.ActiveRow.Cells["Fecha"].Value.ToString();
									decimal vlrDescuento = Convert.ToDecimal(this.txtDescnto.Value);

									if (Convert.ToDecimal(this.txtDescnto.Value) != 0.00m) vlrDescuento = Decimal.Parse(this.txtDescnto.Value.ToString());
									sSQL = string.Format("Exec GrabaMaestroFacturaLotte {0}, {1}, {2}, {3}, {4}, {5}, {6}, {7}, {8}, {9}, {10}, {11}, '{12}', '{13}', {14}, {15}, '{16}', '{17}', '{18}', '{19}', '{20}', '{21}', {22}, '{23}', '{24}', {25}, '{26}', {27}, {28}, {29}, {30}, '{31}', {32}, {33}, {34}, {35}, {36}, {37}, {38}, {39}, {40}, {41}, {42}, 1, 0, 0, '', 0.0, {43}, {44}, {45}, {46}, {47}, {48}, {49}, {50}, {51}, {52}, {53},'', 0", 
										idCpaFC, 1,	0, 1, 15, 10, false, 1, (int)this.cmbFormaPago.Value, NBodga, idProyct, (int)this.cmbVendedor.Value, this.txtNumFactura.Text, "",
										(int)this.cmbTipoRuc.Value, idClint, this.txtNumCedula.Text.ToString(), this.txtNombre.Text.ToString(), dtFechaFC.ToString("yyyyMMdd"), 
										"", "", dtFechaFC.ToString("yyyyMMdd"), false, dtFechaFC.ToString("yyyyMMdd"), dtFechaFC.ToString("yyyyMMdd HH:mm"), 0,
										DatosDescuento, false, 0, 0, 0, dtFechaFC.ToString("yyyyMMdd"), 0, 0, (decimal)this.txtSubTotal.Value, 
										vlrDescuento, (decimal)this.txtIva.Value, 
										(decimal)this.txtTotal.Value,	false, false, false, false, false, 0, 0, 0, 0, false, 0, 0, 0, 0, 0, 0);
									oCmdActualiza.CommandText = sSQL;
									idCpaFC = (int)oCmdActualiza.ExecuteScalar();
									#endregion Maestro
									#region cajero vendedor cobrador
									string lcQryCVC = string.Format("Exec ActualizaCajCob {0},{1},{2},0", 
										idCpaFC, (int)this.cmbCobrador.Value, (int)this.cmbCajero.Value);
									oCmdActualiza.CommandText = lcQryCVC;
									oCmdActualiza.ExecuteNonQuery();
									#endregion cajero vendedor cobrador
									#region Actualiza conbrador y cajero
									string lcQryCC = string.Format("Exec GrabaVen_Caj_Cob_Turno {0},{1},{2},{3},0", 
										idCpaFC, (int)this.cmbCobrador.Value, 0, (int)this.cmbCajero.Value);
									oCmdActualiza.CommandText = lcQryCC;
									oCmdActualiza.ExecuteNonQuery();
									#endregion Actualiza conbrador y cajero
									#region Actualiza TURNO Factura
									int nwTurnoFC = 0;
									if(this.rb1.Checked) nwTurnoFC = 1;
									else if(this.rb2.Checked) nwTurnoFC = 2;
									else nwTurnoFC = 1;
									string lcQryTFC = string.Format("Exec Turno_Manual_Lotte {0},{1}", 
										idCpaFC, nwTurnoFC);
									oCmdActualiza.CommandText = lcQryTFC;
									oCmdActualiza.ExecuteNonQuery();
									#endregion Actualiza TURNO Factura

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
											PorcentajeIva, 0, (decimal)dr.Cells["DescuentoPorc"].Value, (decimal)dr.Cells["Subtotal"].Value,	NBodga, -1,
											0, false, iPosicion, false, false, 0, "");
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
											if ((int)dr.Cells["idFormaPago"].Value==16)
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
									if ((int)this.cmbFormaPago.Value == 6)
									{
										string sSQLSaldo = string.Format("Exec ActualizaSaldoFacturas {0}, {1}, {2}", idCpaFC, 1, (int)this.cmbFormaPago.Value);
										oCmdActualiza.CommandText = sSQLSaldo;
										oCmdActualiza.ExecuteNonQuery();								
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

									if (false)
									{
										string sSQLCents = string.Format("Exec CuadraCentavos {0}", idCpaFC.ToString());
										oCmdActualiza.CommandText = sSQLCents;
										oCmdActualiza.ExecuteNonQuery();
									}

									#endregion Asiento Automatico

									oTransaction.Commit();

									#region Impresión
									//using (frmMensajeNumeración miMsje = new frmMensajeNumeración (this.ugFacRelizadas.ActiveRow.Cells["Numero"].Value.ToString(), false))
									using (frmMnsajeNumLotte miMsje = new frmMnsajeNumLotte(this.txtNumFactura.Text.ToString(), false))
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
									//string SP_Act = string.Format("Exec Cambia_Estado_Habitacion {0},{1},{2},1", idNhab,3,NBodga);
									//Funcion.EjecutaSQL(cdsSeteo,SP_Act);

									this.label13.Text = "CONTABILIZADO";
					
									#region Controles
									//this.cmbFormaPago.Enabled = false;
									this.bntGrabar.Enabled = false;
									this.btnImprimir.Enabled = false;
									#endregion Controles
									TotalDeFactura = Math.Round(Convert.ToDecimal(this.txtTotal.Value), 2);
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
			// obtener numero de habitacion
			try
			{
				this.txtNumHabitacion.Text = ugHabitaciones.ActiveRow.Cells["NumeroHabitacion"].Value.ToString();
				idNhab = (int)ugHabitaciones.ActiveRow.Cells["idHabitacion"].Value;
				//NumFactura = ugFacRelizadas.ActiveRow.Cells["Numero"].Value.ToString();
				//this.txtNumCedula.Text = "9999999999";
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

				#region visualiza reporte
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
				#endregion visualiza reporte
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
				//idCpaFC = (int)ugFacRelizadas.ActiveRow.Cells["idCompra"].Value;
				if (idCpaFC != 0)
				{
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
					preFc = (decimal)ugArticulos.ActiveRow.Cells["Precio1"].Value;
				
					//Validar existencia de Producto
					//int Existencia = Funcion.iEscalarSQL(cdsSeteo, string.Format("Select dbo.ExistenciasInventarioLotte({0}, {1})", idArtFc, NBodga));
					int Existencia = Funcion.iEscalarSQL(cdsSeteo, string.Format("Exec Existencia_Receta {0}, {1}", idArtFc, NBodga));

					BuscaPromocionArticulo(idArtFc);

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
						//string SP_Insert = string.Format("Exec GrabaDetalleFCHabitacion2 {0},{1},{2},{3},{4}", idCpaFC,idClFc,idArtFc,1,preFc);
						string SP_Insert = string.Format("Exec GrabaDetalle_FCHabitacion {0},{1},{2},{3},{4},'{5}',{6},{7}", idCpaFC,idClFc,idArtFc,1,preFc,respuestapromocion,tpopromocion,valorpromo);
						Funcion.EjecutaSQL(cdsSeteo,SP_Insert);
						string QryDetFc = "SELECT idDetCompra,idCompra,idArticulo,Articulo,Cantidad,Precio,DescuentoPorc,SubTotal,IsNull(DescGlobal,0) as DescGlobal From DetCompra Where idCompra="+idCpaFC.ToString();
						ugDetalleFactura.DataSource = Funcion.dvProcedimiento(cdsSeteo, QryDetFc);
						Total();
					}
				} 
				else
				{
					MessageBox.Show("No se ha registrado una factura para poder registrar productos...?\nRegistre un Número de FACTURA para continuar con el proceso...", "Aviso....!!!", MessageBoxButtons.OK, MessageBoxIcon.Stop);
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
						Total();
						e.DisplayPromptMsg = false;
					}
					else e.Cancel = true;
				}
			}
		}

		private void ugDetalleFactura_AfterCellUpdate(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			#region Cantidad
			if (e.Cell.Column.ToString().Equals("Cantidad"))
			{
				if (e.Cell.Row.Cells["Cantidad"].Value == System.DBNull.Value)
				{
					e.Cell.Row.Cells["Cantidad"].Value = (int)e.Cell.OriginalValue;
					//e.Cell.Row.Cells["Precio"].Value = (decimal)e.Cell.OriginalValue;
				}
				if (true)
				{
					int idArticulo = (int)e.Cell.Row.Cells["idArticulo"].Value;
					//int Existencia = Funcion.iEscalarSQL(cdsSeteo, string.Format("Select dbo.ExistenciasInventarioLotte({0}, {1})", idArticulo, NBodga));
					int Existencia = Funcion.iEscalarSQL(cdsSeteo, string.Format("Exec Existencia_Receta {0}, {1}", idArticulo, NBodga));
					int iCantidad = 0;

					foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.ugDetalleFactura.Rows.All)
					{
						if (idArticulo == (int)dr.Cells["idArticulo"].Value) iCantidad = iCantidad + (int)dr.Cells["Cantidad"].Value;
					}

					if(idArticulo==3422 || idArticulo==3423) 
					{	
						if (e.Cell.Column.ToString().Equals("Cantidad")) Total();
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
						else if (e.Cell.Column.ToString().Equals("Cantidad")) Total();
					}
				}
			}
			//Total();
			#endregion Cantidad

			#region Precio
			if (e.Cell.Column.ToString().Equals("Precio"))
			{
				if (e.Cell.Row.Cells["Precio"].Value == System.DBNull.Value)
				{
					e.Cell.Row.Cells["Precio"].Value = (decimal)e.Cell.OriginalValue;
				}
				if (true)
				{
					if (e.Cell.Column.ToString().Equals("Precio")) Total();
				}
			}
			#endregion Precio
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

		private void frmFacturacion_Closing(object sender, System.ComponentModel.CancelEventArgs e)
		{
			//Validar que tenga servicio de hora
			if (idCpaFC>0 && !this.label13.Text.ToString().Equals("ANULADO"))
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
				if(TotalDeFactura!=TotalDeetalleFactura)
				{
					MessageBox.Show("EL TOTAL DE LA FACTURA SE HA CAMBIADO.?"+"\n"+"POR FAVOR ELIMINE O REGISTRE PRODUCTO EN DETALLE DE FACTURA"+"\n"+"O GRABE PARA TOTALIZAR EL NUEVO VALOR DE LA FACTURA!!!", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
					e.Cancel = true;
				}
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

		private void textBox1_Validated(object sender, System.EventArgs e)
		{
			#region Busca Promocion 
			/* Buscar Número de Bodega */
			string lcQryBod = "SELECT Bodega From Bodega Where idBodega in (Select idSucursal From Habitaciones Where idHabitacion="+idNhab.ToString()+")";
			int laBodega = Funcion.iEscalarSQL(cdsSeteo, lcQryBod);
			NBodga = laBodega;
			string sSQL = string.Format("Exec BuscaPromocion {0}, '{1}'", NBodga, "");
			string Rsltado = Funcion.sEscalarSQL(cdsSeteo, sSQL);
			if (Rsltado.Equals("NO"))
			{
				MessageBox.Show("Número de promoción no Valida\nPor favor verifique número de promoción.?","Aviso...!!!",MessageBoxButtons.OK, MessageBoxIcon.Information);
				//this.txtnPromo.Text = "";
			}
			else this.txtDireccion.Focus();
			#endregion Busca Promocion
		}

		private void txtnPromo_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			//if (e.KeyChar == 13) this.txtDireccion.Focus();
		}

		private void cmbBodega_ValueChanged(object sender, System.EventArgs e)
		{
			int idBodg = (int)this.cmbBodega.Value;
			string ugSQL = string.Format("Exec HabitacionesNormales {0}",idBodg);
			ugHabitaciones.DataSource = Funcion.dvProcedimiento(cdsSeteo, ugSQL);
			string nHoras = "SELECT Bodega From Bodega Where idBodega="+idBodg.ToString();
			NBodga = Funcion.iEscalarSQL(cdsSeteo, nHoras);
		}

		private void button2_Click_1(object sender, System.EventArgs e)
		{
			if(this.cmbBodega.ActiveRow == null)
			{
				MessageBox.Show("No se a seleccionado Bodega?","Aviso....!!!", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.cmbBodega.Focus();
			}
			else
			{
				using (frmBuscaFacturas miBusqueda = new frmBuscaFacturas(1, NBodga, miAcceso.BLimiteBusquedaFechas))
				{
					if (DialogResult.OK == miBusqueda.ShowDialog())
					{
						this.Consultar((int)miBusqueda.grdFacturas.ActiveRow.Cells["idCompra"].Value);
						this.btnAnular.Enabled = true;
						this.btnImprimir.Enabled = true;
					}
					else this.btnImprimir.Enabled = false;
				}
			}
		}

		private void Consultar(int IdCompra)
		{
			try
			{	
				#region Maestro
				idCpaFC = IdCompra;
				int iBodega = 0;
				
				string sSQL = string.Format("Exec VentaConsultaIndividual {0}", idCpaFC);
				SqlDataReader dr = Funcion.rEscalarSQL(cdsSeteo, sSQL, true);
				dr.Read();
				if (dr.HasRows)
				{					
					this.label13.Text = dr.GetString(6);
					this.cmbFormaPago.Value = dr.GetInt32(9);
					iBodega = dr.GetInt32(10);
					this.cmbVendedor.Value = dr.GetInt32(12);
					this.txtNumFactura.Text = dr.GetValue(13).ToString();
					idClint = dr.GetInt32(16);
					this.txtNumCedula.Text = dr.GetValue(17).ToString();
					this.txtNombre.Text = dr.GetValue(18).ToString();
					this.dtFecha.Value = dr.GetDateTime(19);
					
					this.txtSubTotal.Value = dr.GetDecimal(36);
					this.txtDescnto.Value = dr.GetDecimal(37);
					this.txtIva.Value = dr.GetDecimal(38);
					this.txtTotal.Value = dr.GetDecimal(39);
					TotalDeFactura = Math.Round(dr.GetDecimal(39),2);
					if(dr.GetInt32(27)==1) this.rb1.Checked = true;
					else this.rb2.Checked = true;

					this.cmbTipoRuc.Value = dr.GetInt32(15);

					//bImpreso = dr.GetBoolean(40);
					//this.txtNumIdAsiento.Value = dr.GetInt32(41);
					//bAnulado = dr.GetBoolean(42);
					//this.lblUsuario.Text = dr.GetString(43);

					//this.txtidBanco.Value = dr.GetInt32(52);
					//this.txtidTarjeta.Value = dr.GetInt32(53);
					//this.txtidPlan.Value = dr.GetInt32(54);
					//this.txtidPlazo.Value = dr.GetInt32(55);
				}
				dr.Close();

				//this.txtBaseImponible.Value = (decimal)this.txtIva.Value - (decimal)this.txtDescIvaTotal.Value;

				if(this.label13.Text.ToString().Equals("ANULADO")) this.bntGrabar.Enabled = false;

				#endregion Maestro

				#region adicionales
				string adSQL = string.Format("Exec DatosAd_Compra {0}",idCpaFC);
				SqlDataReader addr = Funcion.rEscalarSQL(cdsSeteo, adSQL, true);
				addr.Read();
				if (addr.HasRows)
				{					
					this.txtNumHabitacion.Text = addr.GetInt32(0).ToString();
					this.cmbCobrador.Value = addr.GetInt32(2);
					this.cmbCajero.Value = addr.GetInt32(3);
					idDetFormaPago = addr.GetInt32(4);
					DetalleDePrenda = addr.GetString(5).Trim();
				}
				addr.Close();
				#endregion adicionales

				string QryDetFc = "SELECT idDetCompra,idCompra,idArticulo,Articulo,Cantidad,Precio,DescuentoPorc,SubTotal,IsNull(DescGlobal,0) as DescGlobal From DetCompra Where idCompra="+idCpaFC.ToString();
				ugDetalleFactura.DataSource = Funcion.dvProcedimiento(cdsSeteo, QryDetFc);
				idGrpClte = Funcion.iEscalarSQL(cdsSeteo, string.Format("Select idGrupoCliente From Cliente Where idCliente in (Select idCliente From Compra Where idCompra="+idCpaFC.ToString()+")"));
				Total();

				//this.ugDetalleFactura.DataSource = Funcion.dvProcedimiento(cdsSeteo, string.Format("Exec ConsultaDetalleFactura {0}", idCpaFC));

				//if ((int)this.cmbFormaPago.Value == 3) this.lblPlan.Text = Funcion.sEscalarSQL(cdsSeteo, string.Format("Select Descripcion From Planes Where idPlan = {0}", (int)this.txtidPlan.Value));				

				/*
				#region anulado
				if (!bAnulado)
				{
					if ((int)this.txtNumEstado.Value != 6)
					{
						if (!Funcion.ValidaPeriodo((DateTime)this.dtFecha.Value, cdsSeteoF, "Venta"))
						{
							if (miAcceso.Editar) this.btnEditar.Enabled = true;// && ((int)this.cmbFormaPago.Value != 9 || (int)this.cmbFormaPago.Value != 6)) 
							if (miAcceso.Anular && !bAnulado) this.btnAnular.Enabled = true;
							if (miAcceso.BAnulaSoloDelDia && !bAnulado)
							{
								if ((int)this.cmbFormaPago.Value != 9)
								{
									if ((DateTime)this.dtFecha.Value == DateTime.Today)this.btnAnular.Enabled = true;
								}
							}
							if (miAcceso.Eliminar && bAnulado) this.btnAnular.Enabled = true;						
							if (miAcceso.CrearAsiento) this.btnAsiento.Enabled = true;
						}
					}
					else
					{
						if (miAcceso.Editar && ((int)this.cmbFormaPago.Value != 9 || (int)this.cmbFormaPago.Value != 6)) this.btnEditar.Enabled = true;
						if (miAcceso.Anular && !bAnulado) this.btnAnular.Enabled = true;
						if (miAcceso.BAnulaSoloDelDia && !bAnulado)
						{
							if ((int)this.cmbFormaPago.Value != 9)
							{
								if ((DateTime)this.dtFecha.Value == DateTime.Today)this.btnAnular.Enabled = true;
							}
						}
						if (miAcceso.Eliminar && bAnulado) this.btnAnular.Enabled = true;						
						if (miAcceso.CrearAsiento && !bAnulado) this.btnAsiento.Enabled = true;
					}

					if ((int)this.txtNumEstado.Value != 6)
					{
						if (!bImpreso) if (miAcceso.Imprimir) this.btnImprimir.Enabled = true;
						if (bImpreso) if (miAcceso.ReImprimir) this.btnImprimir.Enabled = true;
					}

					if (miAcceso.BBuscarDocumentos) this.btnBuscar.Enabled = true;
					
					if (bCrediPoint) this.btnImprimeDctos.Enabled = true;

					this.btnCliente.Enabled = true;
				}
				else if (miAcceso.Eliminar && bAnulado) this.btnAnular.Enabled = true;

				this.btnCancelar.Enabled = true;

				if (bAnulado) this.btnAnular.Text = "Eliminar";
				#endregion anulado
				*/

			}
			catch(Exception Exc)
			{
				MessageBox.Show(string.Format("Error al Consultar: {0}", Exc), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
			}
		}

		private void button3_Click_1(object sender, System.EventArgs e)
		{
			/*Encerar variable de pantalla*/
			this.txtNumCedula.Text = "";
			this.dtFecha.Value = DateTime.Today;
			this.txtNumHabitacion.Text = "";
			this.txtNumFactura.Text = "";
			this.cmbVendedor.Value = System.DBNull.Value;
			this.cmbFormaPago.Value = System.DBNull.Value;
			this.cmbCobrador.Value = System.DBNull.Value;
			this.cmbCajero.Value = System.DBNull.Value;
			this.txtSubTotal.Value = 0;
			this.txtDescnto.Value = 0;
			this.txtTotal.Value = 0;
			this.txtIva.Value = 0;
			this.txtNombre.Text = "";
			this.txtDireccion.Text = "";
			this.txtTelefono.Text = "";
			this.bntGrabar.Enabled = true;
			this.label13.Text = "";
			this.btnAnular.Enabled = false;
			/*Grilla elimina datos*/
			string QryDetFc = "SELECT idDetCompra,idCompra,idArticulo,Articulo,Cantidad,Precio,SubTotal From DetCompra Where idCompra=0";
			ugDetalleFactura.DataSource = Funcion.dvProcedimiento(cdsSeteo, QryDetFc);
			this.ugDetalleFactura.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			this.ugDetalleFactura.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.True;
			idCpaFC = 0;
			this.txtNumHabitacion.Text = "";
			this.cmbVendedor.Value = 0;
			this.cmbVendedor.Text = "";
			this.txtNumCedula.Text = "";
			this.txtTotDescuento.Value = 0;
			this.btnImprimir.Enabled = false;
		}

		private void btnAnular_Click(object sender, System.EventArgs e)
		{
			string sMensaje = "ANULAR";
			if (true)
			{	
				#region Nuevo Registro	
				if (this.txtNumFactura.Text.ToString().Length == 0)
				{
					MessageBox.Show("Ingrese el número de la factura", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					this.txtNumFactura.Focus();						
				}
				else
				{
					using (frmJustAnulaVenta miNota = new frmJustAnulaVenta())
					{
						if (DialogResult.OK == miNota.ShowDialog())
						{
							#region Maestro
							this.label13.Text = "ANULADO";
							string notasa = miNota.txtNotas.Text.ToString();

							string aSQL = string.Format("Exec AnulaFCLotte {0}, '{1}', {2}", 
								idCpaFC, notasa.ToString(), (int)this.cmbCajero.Value);
							Funcion.EjecutaSQL(cdsSeteo, aSQL, true);
							#endregion Maestro

							string sSQLCobro = string.Format("Exec GuardaDetalleFormaPago 0, {0}, 1, 1, 0, 0, 0, '', '', '', '', '', 0, '', '', 0, 0, 0, 0", idCpaFC);
							Funcion.EjecutaSQL(cdsSeteo, sSQLCobro, true);
							this.button3_Click_1(sender, e);
						}
					}
				}
				#endregion Nuevo Registro
			}
		}

		private void btnElimina_Click(object sender, System.EventArgs e)
		{
			/*Eliminar factura*/
			if (this.txtNumFactura.Text.ToString().Length == 0)
			{
				MessageBox.Show("Ingrese el NUMERO de factura", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.txtNumFactura.Focus();						
			}
			else if(this.cmbBodega.ActiveRow == null)
			{
				MessageBox.Show("No se a seleccionado Bodega?","Aviso....!!!", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.cmbBodega.Focus();
			}
			else
			{	
				string lcQry = "SELECT IsNull(Numero,'') as NumFac FROM Compra WHERE numero like "+"'%"+this.txtNumFactura.Text.ToString()+"'"+" And Bodega="+NBodga.ToString();
				string nFacHab = Funcion.sEscalarSQL(cdsSeteo, lcQry);
				int valfc = 0;
				if (nFacHab!=null) valfc = Int32.Parse(nFacHab);
				int vlrfc = Int32.Parse(this.txtNumFactura.Text.Trim());
				if (valfc==vlrfc)
				{
    
					if (DialogResult.Yes == MessageBox.Show("Esta seguro de ELIMINAR este documento.", "Aviso..!!!", MessageBoxButtons.YesNo, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2))
					{
						string sSQLCobro = string.Format("Exec EliminaFCLotte '{0}',{1}", this.txtNumFactura.Text.ToString(), NBodga);
						Funcion.EjecutaSQL(cdsSeteo, sSQLCobro, true);
						TotalDeFactura = 0.00m;
						TotalDeetalleFactura = 0.00m;
						button3.PerformClick();
					}
				}
				else
				{
					MessageBox.Show("Factura NO se encuentra Registrada en Sistema.?\nPorfavor registre otro número de FACTURA....","Aviso...!!!",MessageBoxButtons.OK);
					return;
				}
			}
		}

		private void ugDetalleFactura_AfterRowsDeleted(object sender, System.EventArgs e)
		{
			Total();
		}

		private void txtBskArticulo_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (e.KeyChar == 13) this.btBuscar_Click(sender, e);
		}

		private void groupBox2_Enter(object sender, System.EventArgs e)
		{
		
		}

		private void txtDescnto_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (e.KeyChar == 13)
			{
				GenerarDescuento();
			}
		}

		public void GenerarDescuento() 
		{
			this.txtDescnto.Value = Math.Round(Convert.ToDecimal(this.txtTotDescuento.Value)/(1+(PorcentajeIva/100)),2);

			/*Calcular descuento de Habitación*/
			string SP_Qry = string.Format("Exec DescuentoFacturaLotteServicio {0},{1}", 
				idCpaFC,Convert.ToDecimal(this.txtDescnto.Value));
		  Funcion.EjecutaSQL(cdsSeteo,SP_Qry);
				
		  //Visualiza detalle factura
		  string QryDetFc = "SELECT idDetCompra,idCompra,idArticulo,Articulo,Cantidad,Precio,SubTotal,IsNull(DescGlobal,0) as DescGlobal From DetCompra Where idCompra="+idCpaFC.ToString();
		  ugDetalleFactura.DataSource = Funcion.dvProcedimiento(cdsSeteo, QryDetFc);

			Total();
			this.bntGrabar.Focus();
		}

		private void txtDescnto_BeforeExitEditMode(object sender, Infragistics.Win.BeforeExitEditModeEventArgs e)
		{
			GenerarDescuento();
			Total();
		}

		private void txtDescnto_Click(object sender, System.EventArgs e)
		{
			this.txtDescnto.SelectAll();
		}

		private void txtDescnto_Enter(object sender, System.EventArgs e)
		{
			this.txtDescnto.SelectAll();
		}

		private void txtDescnto_ValueChanged(object sender, System.EventArgs e)
		{
		}

		private void txtDescnto_PropertyChanged(object sender, Infragistics.Win.PropertyChangedEventArgs e)
		{
		  //GenerarDescuento();
		}

		private void txtTotDescuento_Click(object sender, System.EventArgs e)
		{
			this.txtTotDescuento.SelectAll();
		}

		private void txtTotDescuento_BeforeExitEditMode(object sender, Infragistics.Win.BeforeExitEditModeEventArgs e)
		{
		  this.txtTotDescuento.SelectAll();
		}

		private void txtTotDescuento_Enter(object sender, System.EventArgs e)
		{
			this.txtTotDescuento.SelectAll();
		}

		private void txtTotDescuento_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (e.KeyChar == 13)
			{
				GenerarDescuento();
			}
		}

		private void button4_Click_1(object sender, System.EventArgs e)
		{
			Imprimir(0);
		}

		private void cmbBodega_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}

	}
}
