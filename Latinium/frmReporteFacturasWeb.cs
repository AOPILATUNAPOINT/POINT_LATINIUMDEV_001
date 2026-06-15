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
using System.Threading;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de frmReporteFacturasWeb.
	/// </summary>
	public class frmReporteFacturasWeb : System.Windows.Forms.Form
	{
		private System.Windows.Forms.Label label6;
		private System.Windows.Forms.Label label1;
		private System.Windows.Forms.Label label3;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtNombre;
		private System.Windows.Forms.Button btnConsultar;
		private System.Windows.Forms.GroupBox groupBox1;
		public Infragistics.Win.UltraWinGrid.UltraGrid ultraGrid1;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource1;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtDesde;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtHAsta;
		private System.Windows.Forms.Label label10;
		private System.Windows.Forms.Label label2;
		private System.Windows.Forms.Button btnExportar;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbFormadePago;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbEstado;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource2;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource3;
		private System.Windows.Forms.Label lblContador;
		private C1.Data.C1DataSet cdsSeteoF;
		private Infragistics.Win.UltraWinGrid.UltraGrid ultraGrid2;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource4;
		private System.Windows.Forms.Label label4;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource5;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbEntidad;
		private System.Windows.Forms.Label label5;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource6;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource7;
		public Infragistics.Win.UltraWinGrid.UltraGrid ultraGrid3;
		public Infragistics.Win.UltraWinGrid.UltraGrid ultraGrid4;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource8;
		private Infragistics.Win.UltraWinGrid.UltraDropDown cmbVendedor;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource9;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbEsatdoPedido;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public frmReporteFacturasWeb()
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
			Infragistics.Win.Appearance appearance1 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton1 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmReporteFacturasWeb));
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton2 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Fecha");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idWeb_Cabecera_Pedido");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Ruc");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cliente");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Total");
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idFormaDePagoWeb");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Forma_Pago");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idEstadoWeb");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Estado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Email");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn11 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Direccion_Entrega");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn12 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("telefono");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn13 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Estado_Web");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn14 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCompra");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn15 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("iEWeb");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn16 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Entidad");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn17 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("b_Web");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn18 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idVendedor", -1, "cmbVendedor");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn19 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("btnGestion", 0, null, 0, Infragistics.Win.UltraWinGrid.SortIndicator.Ascending, false);
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn20 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("btnTransferencia", 1);
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn21 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("btnTemporal", 2);
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn22 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("btnSalida", 3);
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn23 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("btnEntregado", 4);
			Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn24 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("btnContabilidad", 5);
			Infragistics.Win.Appearance appearance11 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn25 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("btnVer", 6);
			Infragistics.Win.Appearance appearance12 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings1 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Total", 4, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Total", 4, true);
			Infragistics.Win.Appearance appearance13 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance14 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance15 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance16 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance17 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance18 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance19 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Fecha");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idWeb_Cabecera_Pedido");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn3 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Ruc");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn4 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Cliente");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn5 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Total");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn6 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idFormaDePagoWeb");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn7 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Forma_Pago");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn8 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idEstadoWeb");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn9 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Estado");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn10 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Email");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn11 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Direccion_Entrega");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn12 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("telefono");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn13 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Estado_Web");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn14 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCompra");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn15 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("iEWeb");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn16 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Entidad");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn17 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("b_Web");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn18 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idVendedor");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn26 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idFormaDePagoWeb");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn27 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn19 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idFormaDePagoWeb");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn20 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand3 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn28 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idEstadoWeb");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn29 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn21 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idEstadoWeb");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn22 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.Appearance appearance20 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand4 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn30 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idWeb_Cabecera_Pedido");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn31 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn32 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Articulo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn33 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cantidad");
			Infragistics.Win.Appearance appearance21 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn34 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Precio");
			Infragistics.Win.Appearance appearance22 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn35 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Iva");
			Infragistics.Win.Appearance appearance23 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn36 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Descuento");
			Infragistics.Win.Appearance appearance24 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn37 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Total");
			Infragistics.Win.Appearance appearance25 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance26 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance27 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance28 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance29 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance30 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance31 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn23 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idWeb_Cabecera_Pedido");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn24 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Codigo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn25 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Articulo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn26 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Cantidad");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn27 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Precio");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn28 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Iva");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn29 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Descuento");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn30 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Total");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn31 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Estado");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn32 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand5 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn38 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn39 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Estado");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn33 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn34 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Estado");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn35 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idDetCompra");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn36 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idOrigen");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn37 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCompra");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn38 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idArticulo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn39 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idUnidad");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn40 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Codigo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn41 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Articulo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn42 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Cantidad");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn43 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Precio");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn44 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Impuesto");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn45 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("DescuentoArt");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn46 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("DescuentoPorc");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn47 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("SubTotal");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn48 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("EstadoSerial");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn49 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Tipo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn50 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Entregado");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn51 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Posicion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn52 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Principal");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn53 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Bloqueado");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn54 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCombo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn55 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idArticuloGE");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn56 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Jornadas");
			Infragistics.Win.Appearance appearance32 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand6 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn40 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idDetCompra");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn41 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idOrigen");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn42 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCompra");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn43 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idArticulo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn44 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idUnidad");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn45 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn46 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Articulo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn47 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cantidad");
			Infragistics.Win.Appearance appearance33 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn48 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Precio");
			Infragistics.Win.Appearance appearance34 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn49 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Impuesto");
			Infragistics.Win.Appearance appearance35 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn50 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("DescuentoArt");
			Infragistics.Win.Appearance appearance36 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn51 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("DescuentoPorc");
			Infragistics.Win.Appearance appearance37 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn52 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("SubTotal");
			Infragistics.Win.Appearance appearance38 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn53 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("EstadoSerial");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn54 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Tipo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn55 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Entregado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn56 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Posicion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn57 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Principal");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn58 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bloqueado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn59 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCombo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn60 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idArticuloGE");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn61 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Jornadas");
			Infragistics.Win.Appearance appearance39 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance40 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance41 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance42 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance43 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance44 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance45 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand7 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn62 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idDetCompra");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn63 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idOrigen");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn64 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCompra");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn65 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idArticulo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn66 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idUnidad");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn67 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn68 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Articulo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn69 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cantidad");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn70 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Precio");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn71 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Impuesto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn72 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("DescuentoArt");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn73 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("DescuentoPorc");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn74 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("SubTotal");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn75 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("EstadoSerial");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn76 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Tipo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn77 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Entregado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn78 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Posicion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn79 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Principal");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn80 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bloqueado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn81 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCombo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn82 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idArticuloGE");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn83 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Jornadas");
			Infragistics.Win.Appearance appearance46 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance47 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance48 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance49 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance50 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance51 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn57 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idDetCompra");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn58 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idOrigen");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn59 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCompra");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn60 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idArticulo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn61 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idUnidad");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn62 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Codigo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn63 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Articulo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn64 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Cantidad");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn65 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Precio");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn66 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Impuesto");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn67 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("DescuentoArt");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn68 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("DescuentoPorc");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn69 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("SubTotal");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn70 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("EstadoSerial");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn71 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Tipo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn72 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Entregado");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn73 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Posicion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn74 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Principal");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn75 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Bloqueado");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn76 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCombo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn77 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idArticuloGE");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn78 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Jornadas");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand8 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn84 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPersonal");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn85 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn79 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idPersonal");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn80 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand9 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn86 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Estado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn87 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			this.label6 = new System.Windows.Forms.Label();
			this.dtDesde = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.label1 = new System.Windows.Forms.Label();
			this.dtHAsta = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.label3 = new System.Windows.Forms.Label();
			this.txtNombre = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.btnConsultar = new System.Windows.Forms.Button();
			this.groupBox1 = new System.Windows.Forms.GroupBox();
			this.lblContador = new System.Windows.Forms.Label();
			this.ultraGrid1 = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource1 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.label10 = new System.Windows.Forms.Label();
			this.label2 = new System.Windows.Forms.Label();
			this.cmbFormadePago = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.ultraDataSource3 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.btnExportar = new System.Windows.Forms.Button();
			this.cmbEstado = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.ultraDataSource2 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.ultraGrid2 = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource4 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.ultraDataSource5 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.label4 = new System.Windows.Forms.Label();
			this.cmbEntidad = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.ultraDataSource6 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.label5 = new System.Windows.Forms.Label();
			this.ultraDataSource7 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.ultraGrid3 = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraGrid4 = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource8 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.cmbVendedor = new Infragistics.Win.UltraWinGrid.UltraDropDown();
			this.ultraDataSource9 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.cmbEsatdoPedido = new Infragistics.Win.UltraWinGrid.UltraCombo();
			((System.ComponentModel.ISupportInitialize)(this.dtDesde)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtHAsta)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNombre)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbFormadePago)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource3)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbEstado)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid2)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource4)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource5)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbEntidad)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource6)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource7)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid3)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid4)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource8)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbVendedor)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource9)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbEsatdoPedido)).BeginInit();
			this.SuspendLayout();
			// 
			// label6
			// 
			this.label6.AutoSize = true;
			this.label6.Location = new System.Drawing.Point(16, 16);
			this.label6.Name = "label6";
			this.label6.Size = new System.Drawing.Size(36, 16);
			this.label6.TabIndex = 169;
			this.label6.Text = "Fecha";
			this.label6.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// dtDesde
			// 
			appearance1.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtDesde.Appearance = appearance1;
			this.dtDesde.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton1.Caption = "Today";
			this.dtDesde.DateButtons.Add(dateButton1);
			this.dtDesde.Format = "dd/MM/yyyy";
			this.dtDesde.Location = new System.Drawing.Point(64, 14);
			this.dtDesde.Name = "dtDesde";
			this.dtDesde.NonAutoSizeHeight = 23;
			this.dtDesde.Size = new System.Drawing.Size(112, 21);
			this.dtDesde.SpinButtonsVisible = true;
			this.dtDesde.TabIndex = 168;
			this.dtDesde.Value = ((object)(resources.GetObject("dtDesde.Value")));
			// 
			// label1
			// 
			this.label1.AutoSize = true;
			this.label1.Location = new System.Drawing.Point(184, 16);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(36, 16);
			this.label1.TabIndex = 171;
			this.label1.Text = "Fecha";
			this.label1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// dtHAsta
			// 
			appearance2.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtHAsta.Appearance = appearance2;
			this.dtHAsta.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton2.Caption = "Today";
			this.dtHAsta.DateButtons.Add(dateButton2);
			this.dtHAsta.Format = "dd/MM/yyyy";
			this.dtHAsta.Location = new System.Drawing.Point(232, 14);
			this.dtHAsta.Name = "dtHAsta";
			this.dtHAsta.NonAutoSizeHeight = 23;
			this.dtHAsta.Size = new System.Drawing.Size(112, 21);
			this.dtHAsta.SpinButtonsVisible = true;
			this.dtHAsta.TabIndex = 170;
			this.dtHAsta.Value = ((object)(resources.GetObject("dtHAsta.Value")));
			// 
			// label3
			// 
			this.label3.AutoSize = true;
			this.label3.Location = new System.Drawing.Point(16, 136);
			this.label3.Name = "label3";
			this.label3.Size = new System.Drawing.Size(39, 16);
			this.label3.TabIndex = 246;
			this.label3.Text = "Buscar";
			this.label3.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtNombre
			// 
			appearance3.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtNombre.Appearance = appearance3;
			this.txtNombre.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtNombre.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNombre.Location = new System.Drawing.Point(64, 136);
			this.txtNombre.Name = "txtNombre";
			this.txtNombre.Size = new System.Drawing.Size(280, 21);
			this.txtNombre.TabIndex = 245;
			this.txtNombre.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtNombre_KeyDown);
			// 
			// btnConsultar
			// 
			this.btnConsultar.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(247)), ((System.Byte)(245)), ((System.Byte)(241)));
			this.btnConsultar.CausesValidation = false;
			this.btnConsultar.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.btnConsultar.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(50)), ((System.Byte)(50)), ((System.Byte)(50)));
			this.btnConsultar.Image = ((System.Drawing.Image)(resources.GetObject("btnConsultar.Image")));
			this.btnConsultar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnConsultar.Location = new System.Drawing.Point(360, 136);
			this.btnConsultar.Name = "btnConsultar";
			this.btnConsultar.Size = new System.Drawing.Size(78, 23);
			this.btnConsultar.TabIndex = 247;
			this.btnConsultar.Text = "Consultar";
			this.btnConsultar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnConsultar.Click += new System.EventHandler(this.btnConsultar_Click);
			// 
			// groupBox1
			// 
			this.groupBox1.Location = new System.Drawing.Point(0, 168);
			this.groupBox1.Name = "groupBox1";
			this.groupBox1.Size = new System.Drawing.Size(1168, 8);
			this.groupBox1.TabIndex = 294;
			this.groupBox1.TabStop = false;
			// 
			// lblContador
			// 
			this.lblContador.AutoSize = true;
			this.lblContador.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblContador.ForeColor = System.Drawing.Color.Firebrick;
			this.lblContador.Location = new System.Drawing.Point(456, 136);
			this.lblContador.Name = "lblContador";
			this.lblContador.Size = new System.Drawing.Size(0, 16);
			this.lblContador.TabIndex = 295;
			this.lblContador.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// ultraGrid1
			// 
			this.ultraGrid1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.ultraGrid1.Cursor = System.Windows.Forms.Cursors.Default;
			this.ultraGrid1.DataSource = this.ultraDataSource1;
			appearance4.BackColor = System.Drawing.Color.White;
			appearance4.ForeColor = System.Drawing.Color.Black;
			appearance4.ForeColorDisabled = System.Drawing.Color.Black;
			this.ultraGrid1.DisplayLayout.Appearance = appearance4;
			ultraGridColumn1.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn1.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn1.Header.VisiblePosition = 10;
			ultraGridColumn1.Width = 80;
			ultraGridColumn2.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn2.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn2.Header.VisiblePosition = 0;
			ultraGridColumn2.Hidden = true;
			ultraGridColumn3.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn3.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn3.Header.VisiblePosition = 11;
			ultraGridColumn3.Width = 80;
			ultraGridColumn4.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn4.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn4.Header.VisiblePosition = 12;
			ultraGridColumn4.Width = 250;
			ultraGridColumn5.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn5.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance5.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn5.CellAppearance = appearance5;
			ultraGridColumn5.Format = "#,##0.00";
			ultraGridColumn5.Header.VisiblePosition = 13;
			ultraGridColumn5.Width = 70;
			ultraGridColumn6.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn6.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn6.Header.VisiblePosition = 14;
			ultraGridColumn6.Hidden = true;
			ultraGridColumn7.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn7.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn7.Header.Caption = "Forma De Pago";
			ultraGridColumn7.Header.VisiblePosition = 15;
			ultraGridColumn7.Width = 180;
			ultraGridColumn8.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn8.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn8.Header.VisiblePosition = 17;
			ultraGridColumn8.Hidden = true;
			ultraGridColumn9.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn9.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn9.Header.VisiblePosition = 16;
			ultraGridColumn9.Width = 120;
			ultraGridColumn10.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn10.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn10.Header.VisiblePosition = 18;
			ultraGridColumn10.Width = 200;
			ultraGridColumn11.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn11.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn11.Header.VisiblePosition = 19;
			ultraGridColumn11.Width = 350;
			ultraGridColumn12.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn12.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn12.Header.Caption = "Teléfono";
			ultraGridColumn12.Header.VisiblePosition = 20;
			ultraGridColumn13.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn13.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn13.Header.Caption = "Estado Factura";
			ultraGridColumn13.Header.VisiblePosition = 21;
			ultraGridColumn13.Width = 200;
			ultraGridColumn14.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn14.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn14.Header.VisiblePosition = 22;
			ultraGridColumn14.Hidden = true;
			ultraGridColumn15.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn15.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn15.Header.VisiblePosition = 24;
			ultraGridColumn15.Hidden = true;
			ultraGridColumn16.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn16.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn16.Header.VisiblePosition = 8;
			ultraGridColumn17.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn17.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn17.Header.VisiblePosition = 9;
			ultraGridColumn17.Hidden = true;
			ultraGridColumn17.Width = 100;
			ultraGridColumn18.ButtonDisplayStyle = Infragistics.Win.UltraWinGrid.ButtonDisplayStyle.Always;
			ultraGridColumn18.Header.Caption = "Vendedor";
			ultraGridColumn18.Header.VisiblePosition = 23;
			ultraGridColumn18.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownList;
			ultraGridColumn18.Width = 200;
			ultraGridColumn19.ButtonDisplayStyle = Infragistics.Win.UltraWinGrid.ButtonDisplayStyle.Always;
			appearance6.Image = ((object)(resources.GetObject("appearance6.Image")));
			appearance6.ImageHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn19.CellButtonAppearance = appearance6;
			ultraGridColumn19.Header.Caption = "Gestión";
			ultraGridColumn19.Header.VisiblePosition = 2;
			ultraGridColumn19.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.Button;
			ultraGridColumn19.Width = 30;
			ultraGridColumn20.ButtonDisplayStyle = Infragistics.Win.UltraWinGrid.ButtonDisplayStyle.Always;
			appearance7.Image = ((object)(resources.GetObject("appearance7.Image")));
			appearance7.ImageHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn20.CellButtonAppearance = appearance7;
			ultraGridColumn20.Header.Caption = "Transf.";
			ultraGridColumn20.Header.VisiblePosition = 3;
			ultraGridColumn20.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.Button;
			ultraGridColumn20.Width = 50;
			ultraGridColumn21.ButtonDisplayStyle = Infragistics.Win.UltraWinGrid.ButtonDisplayStyle.Always;
			appearance8.Image = ((object)(resources.GetObject("appearance8.Image")));
			appearance8.ImageHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn21.CellButtonAppearance = appearance8;
			ultraGridColumn21.Header.Caption = "Temporal";
			ultraGridColumn21.Header.VisiblePosition = 5;
			ultraGridColumn21.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.Button;
			ultraGridColumn21.Width = 73;
			ultraGridColumn22.ButtonDisplayStyle = Infragistics.Win.UltraWinGrid.ButtonDisplayStyle.Always;
			appearance9.Image = ((object)(resources.GetObject("appearance9.Image")));
			appearance9.ImageHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn22.CellButtonAppearance = appearance9;
			ultraGridColumn22.Header.Caption = "Enviar";
			ultraGridColumn22.Header.VisiblePosition = 6;
			ultraGridColumn22.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.Button;
			ultraGridColumn22.Width = 50;
			ultraGridColumn23.ButtonDisplayStyle = Infragistics.Win.UltraWinGrid.ButtonDisplayStyle.Always;
			appearance10.Image = ((object)(resources.GetObject("appearance10.Image")));
			appearance10.ImageHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn23.CellButtonAppearance = appearance10;
			ultraGridColumn23.Header.Caption = "OK";
			ultraGridColumn23.Header.VisiblePosition = 7;
			ultraGridColumn23.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.Button;
			ultraGridColumn23.Width = 30;
			ultraGridColumn24.ButtonDisplayStyle = Infragistics.Win.UltraWinGrid.ButtonDisplayStyle.Always;
			appearance11.Image = ((object)(resources.GetObject("appearance11.Image")));
			appearance11.ImageHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn24.CellButtonAppearance = appearance11;
			ultraGridColumn24.Header.Caption = "Contabilidad";
			ultraGridColumn24.Header.VisiblePosition = 4;
			ultraGridColumn24.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.Button;
			ultraGridColumn24.Width = 60;
			ultraGridColumn25.ButtonDisplayStyle = Infragistics.Win.UltraWinGrid.ButtonDisplayStyle.Always;
			appearance12.Image = ((object)(resources.GetObject("appearance12.Image")));
			appearance12.ImageHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn25.CellButtonAppearance = appearance12;
			ultraGridColumn25.Header.Caption = "...";
			ultraGridColumn25.Header.VisiblePosition = 1;
			ultraGridColumn25.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.Button;
			ultraGridColumn25.Width = 30;
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
																										 ultraGridColumn25});
			ultraGridBand1.Header.Caption = "";
			appearance13.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings1.Appearance = appearance13;
			summarySettings1.DisplayFormat = "{0: #,##0.00}";
			summarySettings1.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			ultraGridBand1.Summaries.AddRange(new Infragistics.Win.UltraWinGrid.SummarySettings[] {
																																															summarySettings1});
			this.ultraGrid1.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			appearance14.ForeColor = System.Drawing.Color.Black;
			appearance14.ForeColorDisabled = System.Drawing.Color.Black;
			this.ultraGrid1.DisplayLayout.Override.ActiveRowAppearance = appearance14;
			this.ultraGrid1.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
			this.ultraGrid1.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.True;
			this.ultraGrid1.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance15.BackColor = System.Drawing.Color.Transparent;
			this.ultraGrid1.DisplayLayout.Override.CardAreaAppearance = appearance15;
			appearance16.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance16.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance16.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance16.FontData.BoldAsString = "True";
			appearance16.FontData.Name = "Arial";
			appearance16.FontData.SizeInPoints = 8F;
			appearance16.ForeColor = System.Drawing.Color.White;
			appearance16.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.ultraGrid1.DisplayLayout.Override.HeaderAppearance = appearance16;
			this.ultraGrid1.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti;
			appearance17.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance17.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance17.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.RowAlternateAppearance = appearance17;
			appearance18.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance18.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance18.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.RowSelectorAppearance = appearance18;
			appearance19.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance19.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance19.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.SelectedRowAppearance = appearance19;
			this.ultraGrid1.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.ultraGrid1.Location = new System.Drawing.Point(8, 184);
			this.ultraGrid1.Name = "ultraGrid1";
			this.ultraGrid1.Size = new System.Drawing.Size(1080, 224);
			this.ultraGrid1.TabIndex = 296;
			this.ultraGrid1.BeforeSelectChange += new Infragistics.Win.UltraWinGrid.BeforeSelectChangeEventHandler(this.ultraGrid1_BeforeSelectChange);
			this.ultraGrid1.ClickCellButton += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.ultraGrid1_ClickCellButton);
			this.ultraGrid1.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.ultraGrid1_InitializeLayout);
			// 
			// ultraDataSource1
			// 
			ultraDataColumn1.DataType = typeof(System.DateTime);
			ultraDataColumn2.DataType = typeof(int);
			ultraDataColumn5.DataType = typeof(System.Decimal);
			ultraDataColumn6.DataType = typeof(int);
			ultraDataColumn8.DataType = typeof(int);
			ultraDataColumn14.DataType = typeof(int);
			ultraDataColumn15.DataType = typeof(int);
			ultraDataColumn17.DataType = typeof(int);
			ultraDataColumn18.DataType = typeof(int);
			this.ultraDataSource1.Band.Columns.AddRange(new object[] {
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
																																 ultraDataColumn11,
																																 ultraDataColumn12,
																																 ultraDataColumn13,
																																 ultraDataColumn14,
																																 ultraDataColumn15,
																																 ultraDataColumn16,
																																 ultraDataColumn17,
																																 ultraDataColumn18});
			// 
			// label10
			// 
			this.label10.AutoSize = true;
			this.label10.BackColor = System.Drawing.Color.Transparent;
			this.label10.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label10.Location = new System.Drawing.Point(16, 72);
			this.label10.Name = "label10";
			this.label10.Size = new System.Drawing.Size(38, 17);
			this.label10.TabIndex = 298;
			this.label10.Text = "Estado";
			this.label10.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label2
			// 
			this.label2.AutoSize = true;
			this.label2.BackColor = System.Drawing.Color.Transparent;
			this.label2.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label2.Location = new System.Drawing.Point(16, 104);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(80, 17);
			this.label2.TabIndex = 300;
			this.label2.Text = "Forma de Pago";
			this.label2.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// cmbFormadePago
			// 
			this.cmbFormadePago.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbFormadePago.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbFormadePago.DataSource = this.ultraDataSource3;
			ultraGridColumn26.Header.VisiblePosition = 0;
			ultraGridColumn26.Hidden = true;
			ultraGridColumn27.Header.VisiblePosition = 1;
			ultraGridColumn27.Width = 240;
			ultraGridBand2.Columns.AddRange(new object[] {
																										 ultraGridColumn26,
																										 ultraGridColumn27});
			this.cmbFormadePago.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			this.cmbFormadePago.DisplayMember = "Nombre";
			this.cmbFormadePago.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbFormadePago.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbFormadePago.Location = new System.Drawing.Point(104, 104);
			this.cmbFormadePago.MaxDropDownItems = 30;
			this.cmbFormadePago.Name = "cmbFormadePago";
			this.cmbFormadePago.Size = new System.Drawing.Size(240, 21);
			this.cmbFormadePago.TabIndex = 299;
			this.cmbFormadePago.ValueMember = "idFormaDePagoWeb";
			this.cmbFormadePago.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cmbFormadePago_KeyDown);
			// 
			// ultraDataSource3
			// 
			ultraDataColumn19.DataType = typeof(int);
			this.ultraDataSource3.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn19,
																																 ultraDataColumn20});
			// 
			// btnExportar
			// 
			this.btnExportar.Enabled = false;
			this.btnExportar.Image = ((System.Drawing.Image)(resources.GetObject("btnExportar.Image")));
			this.btnExportar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnExportar.Location = new System.Drawing.Point(1008, 16);
			this.btnExportar.Name = "btnExportar";
			this.btnExportar.Size = new System.Drawing.Size(72, 24);
			this.btnExportar.TabIndex = 742;
			this.btnExportar.Text = "Exportar";
			this.btnExportar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnExportar.Visible = false;
			this.btnExportar.Click += new System.EventHandler(this.btnExportar_Click);
			// 
			// cmbEstado
			// 
			this.cmbEstado.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbEstado.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbEstado.DataSource = this.ultraDataSource2;
			ultraGridColumn28.Header.VisiblePosition = 0;
			ultraGridColumn28.Hidden = true;
			ultraGridColumn28.Width = 240;
			ultraGridColumn29.Header.VisiblePosition = 1;
			ultraGridColumn29.Width = 280;
			ultraGridBand3.Columns.AddRange(new object[] {
																										 ultraGridColumn28,
																										 ultraGridColumn29});
			this.cmbEstado.DisplayLayout.BandsSerializer.Add(ultraGridBand3);
			this.cmbEstado.DisplayMember = "Nombre";
			this.cmbEstado.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbEstado.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbEstado.Location = new System.Drawing.Point(64, 72);
			this.cmbEstado.MaxDropDownItems = 30;
			this.cmbEstado.Name = "cmbEstado";
			this.cmbEstado.Size = new System.Drawing.Size(280, 21);
			this.cmbEstado.TabIndex = 743;
			this.cmbEstado.ValueMember = "idEstadoWeb";
			this.cmbEstado.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cmbEstado_KeyDown);
			this.cmbEstado.ValueChanged += new System.EventHandler(this.cmbEstado_ValueChanged);
			// 
			// ultraDataSource2
			// 
			ultraDataColumn21.DataType = typeof(int);
			this.ultraDataSource2.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn21,
																																 ultraDataColumn22});
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
			// ultraGrid2
			// 
			this.ultraGrid2.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.ultraGrid2.Cursor = System.Windows.Forms.Cursors.Default;
			this.ultraGrid2.DataSource = this.ultraDataSource4;
			appearance20.BackColor = System.Drawing.Color.White;
			appearance20.ForeColor = System.Drawing.Color.Black;
			appearance20.ForeColorDisabled = System.Drawing.Color.Black;
			this.ultraGrid2.DisplayLayout.Appearance = appearance20;
			ultraGridColumn30.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn30.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn30.Header.VisiblePosition = 0;
			ultraGridColumn30.Hidden = true;
			ultraGridColumn31.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn31.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn31.Header.VisiblePosition = 1;
			ultraGridColumn31.Width = 220;
			ultraGridColumn32.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn32.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn32.Header.VisiblePosition = 2;
			ultraGridColumn32.Width = 350;
			ultraGridColumn33.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn33.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance21.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn33.CellAppearance = appearance21;
			ultraGridColumn33.Header.VisiblePosition = 3;
			ultraGridColumn34.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn34.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance22.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn34.CellAppearance = appearance22;
			ultraGridColumn34.Format = "#,##0.00";
			ultraGridColumn34.Header.VisiblePosition = 4;
			ultraGridColumn34.Width = 100;
			ultraGridColumn35.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn35.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance23.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn35.CellAppearance = appearance23;
			ultraGridColumn35.Format = "#,##0.00";
			ultraGridColumn35.Header.VisiblePosition = 5;
			ultraGridColumn35.Width = 100;
			ultraGridColumn36.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn36.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance24.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn36.CellAppearance = appearance24;
			ultraGridColumn36.Format = "#,##0.00";
			ultraGridColumn36.Header.VisiblePosition = 6;
			ultraGridColumn36.Width = 100;
			ultraGridColumn37.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn37.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance25.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn37.CellAppearance = appearance25;
			ultraGridColumn37.Format = "#,##0.00";
			ultraGridColumn37.Header.VisiblePosition = 7;
			ultraGridColumn37.Width = 100;
			ultraGridBand4.Columns.AddRange(new object[] {
																										 ultraGridColumn30,
																										 ultraGridColumn31,
																										 ultraGridColumn32,
																										 ultraGridColumn33,
																										 ultraGridColumn34,
																										 ultraGridColumn35,
																										 ultraGridColumn36,
																										 ultraGridColumn37});
			ultraGridBand4.Header.Caption = "";
			ultraGridBand4.SummaryFooterCaption = "";
			this.ultraGrid2.DisplayLayout.BandsSerializer.Add(ultraGridBand4);
			appearance26.ForeColor = System.Drawing.Color.Black;
			appearance26.ForeColorDisabled = System.Drawing.Color.Black;
			this.ultraGrid2.DisplayLayout.Override.ActiveRowAppearance = appearance26;
			this.ultraGrid2.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.ultraGrid2.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.False;
			appearance27.BackColor = System.Drawing.Color.Transparent;
			this.ultraGrid2.DisplayLayout.Override.CardAreaAppearance = appearance27;
			appearance28.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance28.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance28.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance28.FontData.BoldAsString = "True";
			appearance28.FontData.Name = "Arial";
			appearance28.FontData.SizeInPoints = 8F;
			appearance28.ForeColor = System.Drawing.Color.White;
			appearance28.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.ultraGrid2.DisplayLayout.Override.HeaderAppearance = appearance28;
			appearance29.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance29.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance29.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid2.DisplayLayout.Override.RowAlternateAppearance = appearance29;
			appearance30.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance30.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance30.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid2.DisplayLayout.Override.RowSelectorAppearance = appearance30;
			appearance31.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance31.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance31.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid2.DisplayLayout.Override.SelectedRowAppearance = appearance31;
			this.ultraGrid2.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.ultraGrid2.Location = new System.Drawing.Point(8, 416);
			this.ultraGrid2.Name = "ultraGrid2";
			this.ultraGrid2.Size = new System.Drawing.Size(1080, 176);
			this.ultraGrid2.TabIndex = 744;
			// 
			// ultraDataSource4
			// 
			ultraDataColumn23.DataType = typeof(int);
			ultraDataColumn26.DataType = typeof(int);
			ultraDataColumn27.DataType = typeof(System.Decimal);
			ultraDataColumn28.DataType = typeof(System.Decimal);
			ultraDataColumn29.DataType = typeof(System.Decimal);
			ultraDataColumn30.DataType = typeof(System.Decimal);
			this.ultraDataSource4.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn23,
																																 ultraDataColumn24,
																																 ultraDataColumn25,
																																 ultraDataColumn26,
																																 ultraDataColumn27,
																																 ultraDataColumn28,
																																 ultraDataColumn29,
																																 ultraDataColumn30});
			// 
			// ultraDataSource5
			// 
			ultraDataColumn31.DataType = typeof(int);
			this.ultraDataSource5.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn31,
																																 ultraDataColumn32});
			// 
			// label4
			// 
			this.label4.AutoSize = true;
			this.label4.BackColor = System.Drawing.Color.Transparent;
			this.label4.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label4.Location = new System.Drawing.Point(360, 72);
			this.label4.Name = "label4";
			this.label4.Size = new System.Drawing.Size(75, 17);
			this.label4.TabIndex = 745;
			this.label4.Text = "Estado Pedido";
			this.label4.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.label4.Visible = false;
			// 
			// cmbEntidad
			// 
			this.cmbEntidad.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbEntidad.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbEntidad.DataSource = this.ultraDataSource6;
			ultraGridColumn38.Header.VisiblePosition = 0;
			ultraGridColumn38.Width = 280;
			ultraGridColumn39.Header.VisiblePosition = 1;
			ultraGridColumn39.Hidden = true;
			ultraGridBand5.Columns.AddRange(new object[] {
																										 ultraGridColumn38,
																										 ultraGridColumn39});
			this.cmbEntidad.DisplayLayout.BandsSerializer.Add(ultraGridBand5);
			this.cmbEntidad.DisplayMember = "Nombre";
			this.cmbEntidad.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbEntidad.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbEntidad.Location = new System.Drawing.Point(64, 40);
			this.cmbEntidad.MaxDropDownItems = 30;
			this.cmbEntidad.Name = "cmbEntidad";
			this.cmbEntidad.Size = new System.Drawing.Size(280, 21);
			this.cmbEntidad.TabIndex = 748;
			this.cmbEntidad.ValueMember = "Estado";
			this.cmbEntidad.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cmbEntidad_KeyDown);
			this.cmbEntidad.ValueChanged += new System.EventHandler(this.cmbEntidad_ValueChanged);
			// 
			// ultraDataSource6
			// 
			ultraDataColumn34.DataType = typeof(int);
			this.ultraDataSource6.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn33,
																																 ultraDataColumn34});
			// 
			// label5
			// 
			this.label5.AutoSize = true;
			this.label5.BackColor = System.Drawing.Color.Transparent;
			this.label5.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label5.Location = new System.Drawing.Point(16, 40);
			this.label5.Name = "label5";
			this.label5.Size = new System.Drawing.Size(42, 17);
			this.label5.TabIndex = 747;
			this.label5.Text = "Entidad";
			this.label5.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// ultraDataSource7
			// 
			ultraDataColumn35.DataType = typeof(int);
			ultraDataColumn35.DefaultValue = 0;
			ultraDataColumn36.DataType = typeof(int);
			ultraDataColumn36.DefaultValue = 0;
			ultraDataColumn37.DataType = typeof(int);
			ultraDataColumn37.DefaultValue = 0;
			ultraDataColumn38.DataType = typeof(int);
			ultraDataColumn38.DefaultValue = 0;
			ultraDataColumn39.DataType = typeof(int);
			ultraDataColumn39.DefaultValue = 0;
			ultraDataColumn42.DataType = typeof(int);
			ultraDataColumn42.DefaultValue = 0;
			ultraDataColumn43.DataType = typeof(System.Decimal);
			ultraDataColumn43.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
			ultraDataColumn44.DataType = typeof(System.Decimal);
			ultraDataColumn44.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
			ultraDataColumn45.DataType = typeof(System.Decimal);
			ultraDataColumn45.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
			ultraDataColumn46.DataType = typeof(System.Decimal);
			ultraDataColumn46.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
			ultraDataColumn47.DataType = typeof(System.Decimal);
			ultraDataColumn47.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
			ultraDataColumn48.DataType = typeof(bool);
			ultraDataColumn48.DefaultValue = false;
			ultraDataColumn49.DataType = typeof(int);
			ultraDataColumn49.DefaultValue = 0;
			ultraDataColumn50.DataType = typeof(bool);
			ultraDataColumn50.DefaultValue = false;
			ultraDataColumn51.DataType = typeof(int);
			ultraDataColumn51.DefaultValue = 0;
			ultraDataColumn52.DataType = typeof(bool);
			ultraDataColumn52.DefaultValue = false;
			ultraDataColumn53.DataType = typeof(bool);
			ultraDataColumn53.DefaultValue = false;
			ultraDataColumn54.DataType = typeof(int);
			ultraDataColumn54.DefaultValue = 0;
			ultraDataColumn55.DataType = typeof(int);
			ultraDataColumn55.DefaultValue = 0;
			ultraDataColumn56.DataType = typeof(System.Decimal);
			ultraDataColumn56.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
			this.ultraDataSource7.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn35,
																																 ultraDataColumn36,
																																 ultraDataColumn37,
																																 ultraDataColumn38,
																																 ultraDataColumn39,
																																 ultraDataColumn40,
																																 ultraDataColumn41,
																																 ultraDataColumn42,
																																 ultraDataColumn43,
																																 ultraDataColumn44,
																																 ultraDataColumn45,
																																 ultraDataColumn46,
																																 ultraDataColumn47,
																																 ultraDataColumn48,
																																 ultraDataColumn49,
																																 ultraDataColumn50,
																																 ultraDataColumn51,
																																 ultraDataColumn52,
																																 ultraDataColumn53,
																																 ultraDataColumn54,
																																 ultraDataColumn55,
																																 ultraDataColumn56});
			// 
			// ultraGrid3
			// 
			this.ultraGrid3.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.ultraGrid3.Cursor = System.Windows.Forms.Cursors.Default;
			this.ultraGrid3.DataSource = this.ultraDataSource7;
			appearance32.BackColor = System.Drawing.Color.White;
			this.ultraGrid3.DisplayLayout.Appearance = appearance32;
			this.ultraGrid3.DisplayLayout.AutoFitColumns = true;
			ultraGridColumn40.Header.VisiblePosition = 0;
			ultraGridColumn40.Hidden = true;
			ultraGridColumn41.Header.VisiblePosition = 13;
			ultraGridColumn41.Hidden = true;
			ultraGridColumn42.Header.VisiblePosition = 1;
			ultraGridColumn42.Hidden = true;
			ultraGridColumn42.Width = 56;
			ultraGridColumn43.Header.VisiblePosition = 2;
			ultraGridColumn43.Hidden = true;
			ultraGridColumn43.Width = 66;
			ultraGridColumn44.Header.VisiblePosition = 15;
			ultraGridColumn44.Hidden = true;
			ultraGridColumn45.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn45.Header.VisiblePosition = 3;
			ultraGridColumn45.Width = 238;
			ultraGridColumn46.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn46.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn46.Header.VisiblePosition = 4;
			ultraGridColumn46.Width = 316;
			ultraGridColumn47.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance33.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn47.CellAppearance = appearance33;
			ultraGridColumn47.Format = "";
			ultraGridColumn47.Header.VisiblePosition = 5;
			ultraGridColumn47.PromptChar = ' ';
			ultraGridColumn47.Width = 84;
			appearance34.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn48.CellAppearance = appearance34;
			ultraGridColumn48.Format = "#,##0.00";
			ultraGridColumn48.Header.VisiblePosition = 6;
			ultraGridColumn48.PromptChar = ' ';
			ultraGridColumn48.Width = 100;
			ultraGridColumn49.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance35.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn49.CellAppearance = appearance35;
			ultraGridColumn49.Format = "#,##0.00";
			ultraGridColumn49.Header.Caption = "IVA";
			ultraGridColumn49.Header.VisiblePosition = 7;
			ultraGridColumn49.PromptChar = ' ';
			ultraGridColumn49.Width = 65;
			ultraGridColumn50.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance36.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn50.CellAppearance = appearance36;
			ultraGridColumn50.Format = "#,##0.00";
			ultraGridColumn50.Header.VisiblePosition = 8;
			ultraGridColumn50.Hidden = true;
			ultraGridColumn50.PromptChar = ' ';
			ultraGridColumn50.Width = 90;
			appearance37.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn51.CellAppearance = appearance37;
			ultraGridColumn51.Format = "#,##0.0000";
			ultraGridColumn51.Header.Caption = "% Desc.";
			ultraGridColumn51.Header.VisiblePosition = 9;
			ultraGridColumn51.PromptChar = ' ';
			ultraGridColumn51.Width = 73;
			ultraGridColumn52.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance38.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn52.CellAppearance = appearance38;
			ultraGridColumn52.Format = "#,##0.00";
			ultraGridColumn52.Header.Caption = "Subtotal";
			ultraGridColumn52.Header.VisiblePosition = 10;
			ultraGridColumn52.PromptChar = ' ';
			ultraGridColumn52.Width = 110;
			ultraGridColumn53.Header.VisiblePosition = 14;
			ultraGridColumn53.Hidden = true;
			ultraGridColumn54.Header.VisiblePosition = 12;
			ultraGridColumn54.Hidden = true;
			ultraGridColumn55.Header.VisiblePosition = 11;
			ultraGridColumn55.Hidden = true;
			ultraGridColumn56.Header.VisiblePosition = 16;
			ultraGridColumn56.Hidden = true;
			ultraGridColumn56.Width = 63;
			ultraGridColumn57.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn57.Header.VisiblePosition = 17;
			ultraGridColumn57.Hidden = true;
			ultraGridColumn58.Header.VisiblePosition = 18;
			ultraGridColumn58.Hidden = true;
			ultraGridColumn59.Header.VisiblePosition = 19;
			ultraGridColumn59.Hidden = true;
			ultraGridColumn60.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn60.Header.VisiblePosition = 20;
			ultraGridColumn60.Hidden = true;
			ultraGridColumn61.Header.VisiblePosition = 21;
			ultraGridColumn61.Hidden = true;
			ultraGridBand6.Columns.AddRange(new object[] {
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
																										 ultraGridColumn61});
			ultraGridBand6.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.Edit;
			this.ultraGrid3.DisplayLayout.BandsSerializer.Add(ultraGridBand6);
			appearance39.ForeColor = System.Drawing.Color.Black;
			appearance39.ForeColorDisabled = System.Drawing.Color.Black;
			this.ultraGrid3.DisplayLayout.Override.ActiveRowAppearance = appearance39;
			this.ultraGrid3.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.TemplateOnBottom;
			this.ultraGrid3.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.True;
			this.ultraGrid3.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance40.BackColor = System.Drawing.Color.Transparent;
			this.ultraGrid3.DisplayLayout.Override.CardAreaAppearance = appearance40;
			this.ultraGrid3.DisplayLayout.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.Edit;
			appearance41.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance41.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance41.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance41.FontData.BoldAsString = "True";
			appearance41.FontData.Name = "Arial";
			appearance41.FontData.SizeInPoints = 9F;
			appearance41.ForeColor = System.Drawing.Color.White;
			appearance41.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.ultraGrid3.DisplayLayout.Override.HeaderAppearance = appearance41;
			appearance42.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance42.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance42.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid3.DisplayLayout.Override.RowAlternateAppearance = appearance42;
			appearance43.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance43.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance43.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid3.DisplayLayout.Override.RowSelectorAppearance = appearance43;
			appearance44.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance44.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance44.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid3.DisplayLayout.Override.SelectedRowAppearance = appearance44;
			this.ultraGrid3.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.ultraGrid3.Location = new System.Drawing.Point(45, 203);
			this.ultraGrid3.Name = "ultraGrid3";
			this.ultraGrid3.Size = new System.Drawing.Size(1007, 190);
			this.ultraGrid3.TabIndex = 749;
			// 
			// ultraGrid4
			// 
			this.ultraGrid4.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.ultraGrid4.Cursor = System.Windows.Forms.Cursors.Default;
			this.ultraGrid4.DataSource = this.ultraDataSource8;
			appearance45.BackColor = System.Drawing.Color.White;
			this.ultraGrid4.DisplayLayout.Appearance = appearance45;
			ultraGridColumn62.Header.VisiblePosition = 0;
			ultraGridColumn62.Width = 49;
			ultraGridColumn63.Header.VisiblePosition = 1;
			ultraGridColumn63.Width = 35;
			ultraGridColumn64.Header.VisiblePosition = 2;
			ultraGridColumn64.Width = 39;
			ultraGridColumn65.Header.VisiblePosition = 3;
			ultraGridColumn65.Width = 38;
			ultraGridColumn66.Header.VisiblePosition = 4;
			ultraGridColumn66.Width = 36;
			ultraGridColumn67.Header.VisiblePosition = 5;
			ultraGridColumn67.Width = 61;
			ultraGridColumn68.Header.VisiblePosition = 6;
			ultraGridColumn68.Width = 61;
			ultraGridColumn69.Header.VisiblePosition = 7;
			ultraGridColumn69.Width = 36;
			ultraGridColumn70.Header.VisiblePosition = 8;
			ultraGridColumn70.Width = 49;
			ultraGridColumn71.Header.VisiblePosition = 9;
			ultraGridColumn71.Width = 49;
			ultraGridColumn72.Header.VisiblePosition = 10;
			ultraGridColumn72.Width = 167;
			ultraGridColumn73.Header.VisiblePosition = 11;
			ultraGridColumn73.Width = 150;
			ultraGridColumn74.Header.VisiblePosition = 12;
			ultraGridColumn74.Width = 49;
			ultraGridColumn75.Header.VisiblePosition = 13;
			ultraGridColumn75.Width = 47;
			ultraGridColumn76.Header.VisiblePosition = 14;
			ultraGridColumn76.Width = 35;
			ultraGridColumn77.Header.VisiblePosition = 15;
			ultraGridColumn77.Width = 41;
			ultraGridColumn78.Header.VisiblePosition = 16;
			ultraGridColumn78.Width = 35;
			ultraGridColumn79.Header.VisiblePosition = 17;
			ultraGridColumn79.Width = 35;
			ultraGridColumn80.Header.VisiblePosition = 18;
			ultraGridColumn80.Width = 41;
			ultraGridColumn81.Header.VisiblePosition = 19;
			ultraGridColumn81.Width = 36;
			ultraGridColumn82.Header.VisiblePosition = 20;
			ultraGridColumn82.Width = 46;
			ultraGridColumn83.Header.VisiblePosition = 21;
			ultraGridColumn83.Width = 62;
			ultraGridBand7.Columns.AddRange(new object[] {
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
																										 ultraGridColumn83});
			ultraGridBand7.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.Edit;
			this.ultraGrid4.DisplayLayout.BandsSerializer.Add(ultraGridBand7);
			appearance46.ForeColor = System.Drawing.Color.Black;
			appearance46.ForeColorDisabled = System.Drawing.Color.Black;
			this.ultraGrid4.DisplayLayout.Override.ActiveRowAppearance = appearance46;
			this.ultraGrid4.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.TemplateOnBottom;
			this.ultraGrid4.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.True;
			this.ultraGrid4.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance47.BackColor = System.Drawing.Color.Transparent;
			this.ultraGrid4.DisplayLayout.Override.CardAreaAppearance = appearance47;
			this.ultraGrid4.DisplayLayout.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.Edit;
			appearance48.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance48.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance48.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance48.FontData.BoldAsString = "True";
			appearance48.FontData.Name = "Arial";
			appearance48.FontData.SizeInPoints = 9F;
			appearance48.ForeColor = System.Drawing.Color.White;
			appearance48.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.ultraGrid4.DisplayLayout.Override.HeaderAppearance = appearance48;
			appearance49.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance49.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance49.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid4.DisplayLayout.Override.RowAlternateAppearance = appearance49;
			appearance50.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance50.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance50.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid4.DisplayLayout.Override.RowSelectorAppearance = appearance50;
			appearance51.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance51.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance51.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid4.DisplayLayout.Override.SelectedRowAppearance = appearance51;
			this.ultraGrid4.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.ultraGrid4.Location = new System.Drawing.Point(448, 104);
			this.ultraGrid4.Name = "ultraGrid4";
			this.ultraGrid4.Size = new System.Drawing.Size(640, 56);
			this.ultraGrid4.TabIndex = 750;
			this.ultraGrid4.Visible = false;
			this.ultraGrid4.AfterCellUpdate += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.ultraGrid4_AfterCellUpdate);
			// 
			// ultraDataSource8
			// 
			ultraDataColumn57.DataType = typeof(int);
			ultraDataColumn57.DefaultValue = 0;
			ultraDataColumn58.DataType = typeof(int);
			ultraDataColumn58.DefaultValue = 0;
			ultraDataColumn59.DataType = typeof(int);
			ultraDataColumn59.DefaultValue = 0;
			ultraDataColumn60.DataType = typeof(int);
			ultraDataColumn60.DefaultValue = 0;
			ultraDataColumn61.DataType = typeof(int);
			ultraDataColumn61.DefaultValue = 0;
			ultraDataColumn64.DataType = typeof(int);
			ultraDataColumn64.DefaultValue = 0;
			ultraDataColumn65.DataType = typeof(System.Decimal);
			ultraDataColumn65.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
			ultraDataColumn66.DataType = typeof(System.Decimal);
			ultraDataColumn66.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
			ultraDataColumn67.DataType = typeof(System.Decimal);
			ultraDataColumn67.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
			ultraDataColumn68.DataType = typeof(System.Decimal);
			ultraDataColumn68.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
			ultraDataColumn69.DataType = typeof(System.Decimal);
			ultraDataColumn69.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
			ultraDataColumn70.DataType = typeof(bool);
			ultraDataColumn70.DefaultValue = false;
			ultraDataColumn71.DataType = typeof(int);
			ultraDataColumn71.DefaultValue = 0;
			ultraDataColumn72.DataType = typeof(bool);
			ultraDataColumn72.DefaultValue = false;
			ultraDataColumn73.DataType = typeof(int);
			ultraDataColumn73.DefaultValue = 0;
			ultraDataColumn74.DataType = typeof(bool);
			ultraDataColumn74.DefaultValue = false;
			ultraDataColumn75.DataType = typeof(bool);
			ultraDataColumn75.DefaultValue = false;
			ultraDataColumn76.DataType = typeof(int);
			ultraDataColumn76.DefaultValue = 0;
			ultraDataColumn77.DataType = typeof(int);
			ultraDataColumn77.DefaultValue = 0;
			ultraDataColumn78.DataType = typeof(System.Decimal);
			ultraDataColumn78.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
			this.ultraDataSource8.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn57,
																																 ultraDataColumn58,
																																 ultraDataColumn59,
																																 ultraDataColumn60,
																																 ultraDataColumn61,
																																 ultraDataColumn62,
																																 ultraDataColumn63,
																																 ultraDataColumn64,
																																 ultraDataColumn65,
																																 ultraDataColumn66,
																																 ultraDataColumn67,
																																 ultraDataColumn68,
																																 ultraDataColumn69,
																																 ultraDataColumn70,
																																 ultraDataColumn71,
																																 ultraDataColumn72,
																																 ultraDataColumn73,
																																 ultraDataColumn74,
																																 ultraDataColumn75,
																																 ultraDataColumn76,
																																 ultraDataColumn77,
																																 ultraDataColumn78});
			// 
			// cmbVendedor
			// 
			this.cmbVendedor.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbVendedor.DataSource = this.ultraDataSource9;
			ultraGridColumn84.Header.VisiblePosition = 0;
			ultraGridColumn84.Hidden = true;
			ultraGridColumn85.Header.VisiblePosition = 1;
			ultraGridColumn85.Width = 200;
			ultraGridBand8.Columns.AddRange(new object[] {
																										 ultraGridColumn84,
																										 ultraGridColumn85});
			this.cmbVendedor.DisplayLayout.BandsSerializer.Add(ultraGridBand8);
			this.cmbVendedor.DisplayMember = "Nombre";
			this.cmbVendedor.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbVendedor.Location = new System.Drawing.Point(760, 32);
			this.cmbVendedor.Name = "cmbVendedor";
			this.cmbVendedor.Size = new System.Drawing.Size(160, 56);
			this.cmbVendedor.TabIndex = 986;
			this.cmbVendedor.ValueMember = "idPersonal";
			this.cmbVendedor.Visible = false;
			// 
			// ultraDataSource9
			// 
			ultraDataColumn79.DataType = typeof(int);
			this.ultraDataSource9.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn79,
																																 ultraDataColumn80});
			// 
			// cmbEsatdoPedido
			// 
			this.cmbEsatdoPedido.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbEsatdoPedido.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbEsatdoPedido.DataSource = this.ultraDataSource5;
			ultraGridColumn86.Header.VisiblePosition = 0;
			ultraGridColumn86.Hidden = true;
			ultraGridColumn87.Header.VisiblePosition = 1;
			ultraGridColumn87.Width = 150;
			ultraGridBand9.Columns.AddRange(new object[] {
																										 ultraGridColumn86,
																										 ultraGridColumn87});
			this.cmbEsatdoPedido.DisplayLayout.BandsSerializer.Add(ultraGridBand9);
			this.cmbEsatdoPedido.DisplayMember = "Nombre";
			this.cmbEsatdoPedido.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbEsatdoPedido.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbEsatdoPedido.Location = new System.Drawing.Point(448, 72);
			this.cmbEsatdoPedido.MaxDropDownItems = 30;
			this.cmbEsatdoPedido.Name = "cmbEsatdoPedido";
			this.cmbEsatdoPedido.Size = new System.Drawing.Size(256, 21);
			this.cmbEsatdoPedido.TabIndex = 989;
			this.cmbEsatdoPedido.ValueMember = "Estado";
			this.cmbEsatdoPedido.Visible = false;
			// 
			// frmReporteFacturasWeb
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(1096, 597);
			this.Controls.Add(this.cmbEsatdoPedido);
			this.Controls.Add(this.cmbVendedor);
			this.Controls.Add(this.ultraGrid4);
			this.Controls.Add(this.cmbEntidad);
			this.Controls.Add(this.label5);
			this.Controls.Add(this.label4);
			this.Controls.Add(this.label2);
			this.Controls.Add(this.label10);
			this.Controls.Add(this.label3);
			this.Controls.Add(this.label1);
			this.Controls.Add(this.label6);
			this.Controls.Add(this.ultraGrid2);
			this.Controls.Add(this.cmbEstado);
			this.Controls.Add(this.btnExportar);
			this.Controls.Add(this.cmbFormadePago);
			this.Controls.Add(this.ultraGrid1);
			this.Controls.Add(this.lblContador);
			this.Controls.Add(this.groupBox1);
			this.Controls.Add(this.btnConsultar);
			this.Controls.Add(this.txtNombre);
			this.Controls.Add(this.dtHAsta);
			this.Controls.Add(this.dtDesde);
			this.Controls.Add(this.ultraGrid3);
			this.KeyPreview = true;
			this.Name = "frmReporteFacturasWeb";
			this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
			this.Text = "Ventas WEB";
			this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.frmReporteFacturasWeb_KeyDown);
			this.Load += new System.EventHandler(this.frmReporteFacturasWeb_Load);
			((System.ComponentModel.ISupportInitialize)(this.dtDesde)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtHAsta)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNombre)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbFormadePago)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource3)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbEstado)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid2)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource4)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource5)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbEntidad)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource6)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource7)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid3)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid4)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource8)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbVendedor)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource9)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbEsatdoPedido)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private Acceso miAcceso;

		private void UnloadMe()
		{
			this.Close();
		}

		private void frmReporteFacturasWeb_Load(object sender, System.EventArgs e)
		{
			miAcceso = new Acceso(cdsSeteoF, "200534");

			if (!Funcion.Permiso("1400", cdsSeteoF))
			{				
				MessageBox.Show("No puede ingresar a Ventas WEB", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				BeginInvoke(new MethodInvoker(UnloadMe));
				return;
			}

			if (!miAcceso.Nuevo) this.ultraGrid1.DisplayLayout.Bands[0].Columns["btnGestion"].Hidden = true;
			if (!miAcceso.BRevisadoCE) this.ultraGrid1.DisplayLayout.Bands[0].Columns["btnTransferencia"].Hidden = true;
			if (!miAcceso.BRevisadoCE) this.ultraGrid1.DisplayLayout.Bands[0].Columns["btnContabilidad"].Hidden = true;
			if (!miAcceso.BProcesar) this.ultraGrid1.DisplayLayout.Bands[0].Columns["btnTemporal"].Hidden = true;
			if (!miAcceso.BModificarBodega) this.ultraGrid1.DisplayLayout.Bands[0].Columns["btnSalida"].Hidden = true; 
			if (!miAcceso.BTransferir) this.ultraGrid1.DisplayLayout.Bands[0].Columns["btnEntregado"].Hidden = true;
			if (miAcceso.BExportar) this.btnExportar.Visible = true;
			if (miAcceso.BExportar) this.btnExportar.Enabled = true;


			this.dtDesde.CalendarInfo.MinDate = DateTime.Parse("01/05/2020");
			this.dtHAsta.CalendarInfo.MinDate = DateTime.Parse("01/05/2020");

			this.dtDesde.CalendarInfo.MaxDate = DateTime.Today;
			this.dtHAsta.CalendarInfo.MaxDate = DateTime.Today;

			this.dtDesde.Value = DateTime.Today.AddDays(-30);
			this.dtHAsta.Value = DateTime.Today;
			
			this.cmbFormadePago.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Select idFormaDePagoWeb, Nombre  from FormaDePagoWeb where idFormaDePagoWeb in (1,5) order by nombre ");
			this.cmbEsatdoPedido.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "select Estado, Nombre from Estado_WEB order by nombre ");
			this.cmbEntidad.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "SELECT Nombre, Estado   FROM Tipo_BotonWeb order by Nombre");

			this.cmbVendedor.DataSource = FuncionesProcedimientos.Lista_Vendedores(Convert.ToDateTime(DateTime.Today), 82, cdsSeteoF);
			
		}

		private void btnExportar_Click(object sender, System.EventArgs e)
		{
			FuncionesProcedimientos.ExportaExcel(this.ultraGrid1);
		}

		private void btnConsultar_Click(object sender, System.EventArgs e)
		{
			#region Valida Fechas

			if (this.dtDesde.Value == System.DBNull.Value)
			{
				MessageBox.Show("Seleccione la Fecha Inicial", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.dtDesde.Focus();
				return;
			}

			if (this.dtHAsta.Value == System.DBNull.Value)
			{
				MessageBox.Show("Seleccione la Fecha Final", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.dtHAsta.Focus();
				return;
			}

			if (Convert.ToDateTime(this.dtDesde.Value) > Convert.ToDateTime(this.dtHAsta.Value))
			{
				MessageBox.Show("La Fecha Inicial no puede ser mayor a la Fecha Final", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.dtHAsta.Focus();
				return;
			}
			#endregion Valida Fechas
      
			if (!miAcceso.Nuevo) this.ultraGrid1.DisplayLayout.Bands[0].Columns["btnGestion"].Hidden = true;
			if (!miAcceso.BRevisadoCE) this.ultraGrid1.DisplayLayout.Bands[0].Columns["btnTransferencia"].Hidden = true;
			if (!miAcceso.BProcesar) this.ultraGrid1.DisplayLayout.Bands[0].Columns["btnTemporal"].Hidden = true;
			if (!miAcceso.BModificarBodega) this.ultraGrid1.DisplayLayout.Bands[0].Columns["btnSalida"].Hidden = true; 
			if (!miAcceso.BTransferir) this.ultraGrid1.DisplayLayout.Bands[0].Columns["btnEntregado"].Hidden = true;

			int iEstado = 0;
			int iFormadePago = 0;
			int iEstadoPedido  = 8;
			int iEntidad  = 0;
			if (this.cmbEstado.ActiveRow != null) iEstado = (int)this.cmbEstado.Value;
			if (this.cmbFormadePago.ActiveRow != null) iFormadePago = (int)this.cmbFormadePago.Value;
			if (this.cmbEsatdoPedido.ActiveRow != null) iEstadoPedido = (int)this.cmbEsatdoPedido.Value;
			if (this.cmbEntidad.ActiveRow != null) iEntidad = (int)this.cmbEntidad.Value;
			this.lblContador.Text = "";

			string sSQLPa = string.Format("Exec [ReporteVentasWeb] '{0}', '{1}', {2}, {3}, '{4}', {5}, {6}", 
				Convert.ToDateTime(this.dtDesde.Value).ToString("yyyyMMdd"), Convert.ToDateTime(this.dtHAsta.Value).ToString("yyyyMMdd"), 
				iEstado, iFormadePago, this.txtNombre.Text, iEstadoPedido, iEntidad);
			FuncionesProcedimientos.dsGeneral(sSQLPa, this.ultraGrid1);

			this.lblContador.Text = this.ultraGrid1.Rows.Count + " REGISTROS ENCONTRADOS";

		}

		private void cmbEstado_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter) this.cmbFormadePago.Focus();
		}

		private void cmbFormadePago_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter) this.txtNombre.Focus();
		}

		private void txtNombre_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter) this.btnConsultar_Click (sender, e);
		}

		private void frmReporteFacturasWeb_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Escape) this.Close();		
		}

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);		
		}

		private void ultraGrid1_BeforeSelectChange(object sender, Infragistics.Win.UltraWinGrid.BeforeSelectChangeEventArgs e)
		{
			if (ultraGrid1.ActiveRow == null) return;

			this.ultraGrid4.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Exec ConsultaDetalleFactura 0");
			this.ultraGrid2.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Exec ReporteDetalleVentasWeb 0");
			
			string sSQL = string.Format("Exec ReporteDetalleVentasWeb {0}", (int)this.ultraGrid1.ActiveRow.Cells["idWeb_Cabecera_Pedido"].Value);
			this.ultraGrid2.DataSource = Funcion.dvProcedimiento(cdsSeteoF, sSQL);
      
			
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.ultraGrid2.Rows.All)
			{
				if ( dr.Cells["Codigo"].Value.ToString().Trim().Length  > 0)
				{
					this.ultraGrid4.Rows.Band.AddNew();
				
					this.ultraGrid4.ActiveRow.Cells["Cantidad"].Value = dr.Cells["Cantidad"].Value;
					this.ultraGrid4.ActiveRow.Cells["DescuentoArt"].Value = dr.Cells["Descuento"].Value;
//					this.ultraGrid4.ActiveRow.Cells["SubTotal"].Value = dr.Cells["Total"].Value;
					this.ultraGrid4.ActiveRow.Cells["Precio"].Value = dr.Cells["Precio"].Value;
					this.ultraGrid4.ActiveRow.Cells["Codigo"].Value = dr.Cells["Codigo"].Value;
//					this.ultraGrid4.ActiveRow.Cells["Precio"].Value = dr.Cells["Total"].Value;
				}
			}

			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.ultraGrid4.Rows.All)
			{
				if ( dr.Cells["Codigo"].Value.ToString().Trim().Length  > 0   )
				{
 				dr.Cells["SubTotal"].Value = (decimal )dr.Cells["Precio"].Value * (int)dr.Cells["Cantidad"].Value;
				}
			}
		
		}

	
				
		private void ultraGrid1_ClickCellButton(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{

			#region Gestión WEb 
			if (e.Cell.Column.ToString() == "btnVer")
			{
				if (e.Cell.Row.Cells["idWeb_Cabecera_Pedido"].Value != System.DBNull.Value)
				{
					if ((int)e.Cell.Row.Cells["idWeb_Cabecera_Pedido"].Value > 0)
					{
						if (e.Cell.Row.Cells["idEstadoWeb"].Value != System.DBNull.Value)
						{
							int iConDet= 0;
							iConDet = Funcion.iEscalarSQL(cdsSeteoF, string.Format("EXEC [ValidaDetalleVentasWeb] {0}",Convert.ToInt32 (e.Cell.Row.Cells["idWeb_Cabecera_Pedido"].Value)));	
                 
							if (iConDet == 0)
							{
								MessageBox.Show("El pedido seleccionado no tiene articulos cargados ", "Informacion", MessageBoxButtons.OK, MessageBoxIcon.Information);  
								return;
							}
								
							//							if ((int)e.Cell.Row.Cells["idEstadoWeb"].Value == 3)
							//							{
							using (frmVerWeb web = new frmVerWeb
											 (Convert.ToInt32 (e.Cell.Row.Cells["idWeb_Cabecera_Pedido"].Value) ) )
							{ 
								if (DialogResult.OK == web.ShowDialog( ))
								{

	
								}
								//								}
							}
						}
					}
				}
			}
			#endregion Gestión WEb 

			#region Gestión WEb 
			if (e.Cell.Column.ToString() == "btnGestion")
			{
				if (e.Cell.Row.Cells["idWeb_Cabecera_Pedido"].Value != System.DBNull.Value)
				{
					if ((int)e.Cell.Row.Cells["idWeb_Cabecera_Pedido"].Value > 0)
					{
						if (e.Cell.Row.Cells["idEstadoWeb"].Value != System.DBNull.Value)
						{
//							if ((int)e.Cell.Row.Cells["idEstadoWeb"].Value == 3)
//							{
								using (frmCob_InformeNotCompro Gestion = new frmCob_InformeNotCompro
												 (e.Cell.Row.Cells["Direccion_Entrega"].Value.ToString() ,  0, e.Cell.Row.Cells["telefono"].Value.ToString() , 
												 e.Cell.Row.Cells["Email"].Value.ToString() ,  false , 
												 e.Cell.Row.Cells["Cliente"].Value.ToString() , e.Cell.Row.Cells["Ruc"].Value.ToString() , Convert.ToInt32 (e.Cell.Row.Cells["idWeb_Cabecera_Pedido"].Value),0 )
												 )
								{
									if (DialogResult.OK == Gestion.ShowDialog( ))
									{

	
									}
//								}
							}
						}
					}
				}
			}
      #endregion Gestión WEb 

			#region Gestión WEb trans
			/* tipo 2 confirmacion transferencia */
			if (e.Cell.Column.ToString() == "btnTransferencia")
			{
				if (e.Cell.Row.Cells["idWeb_Cabecera_Pedido"].Value != System.DBNull.Value)
				{
					if ((int)e.Cell.Row.Cells["idWeb_Cabecera_Pedido"].Value > 0)
					{
						if (e.Cell.Row.Cells["idEstadoWeb"].Value != System.DBNull.Value)
						{
							if (((int)e.Cell.Row.Cells["idEstadoWeb"].Value == 2 )&& (int)e.Cell.Row.Cells["idFormaDePagoWeb"].Value == 1 && (int)e.Cell.Row.Cells["iEWeb"].Value  == 1)
							{
								using (frmNotas NotasG = new frmNotas())
								{
									if (DialogResult.OK == NotasG.ShowDialog( ))
									{
										if (DialogResult.Yes == MessageBox.Show("Esta Seguro que el Deposito fue realizado y verificado.", "Point Technology", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2))
										{
											string sSQLWeb = string.Format("Exec [GuardaGestionesWeb] 0, {0},'{1}','{2}', {3}, '{4}', {5}, {6}",
					                                
												Convert.ToInt32 (e.Cell.Row.Cells["idWeb_Cabecera_Pedido"].Value),
												Convert.ToDateTime(DateTime.Today).ToString("yyyyMMdd HH:mm:ss"), 
												Convert.ToString (e.Cell.Row.Cells["telefono"].Value), 
												2,
												NotasG.txtNotas.Text,
												MenuLatinium.IdUsuario, 
												2);
								
											Funcion.EjecutaSQL(cdsSeteoF, sSQLWeb);

											string ssSQL = string.Format("update web_cabecera_pedido set ESTADO = 'P' ,idEstado_WEB = 2 where numero_pedido = {0}", Convert.ToInt32 (e.Cell.Row.Cells["idWeb_Cabecera_Pedido"].Value));
											Funcion.EjecutaSQL(cdsSeteoF, ssSQL);

											this.btnConsultar_Click (sender, e);

										}
									}
								}
							}

						}
					}
				}
			}
			#endregion Gestión WEb trans

			#region Aprueba contabilidad
			if (e.Cell.Column.ToString() == "btnContabilidad")
			{
				if (e.Cell.Row.Cells["idWeb_Cabecera_Pedido"].Value != System.DBNull.Value)
				{
					if ((int)e.Cell.Row.Cells["idWeb_Cabecera_Pedido"].Value > 0)
					{
						if (e.Cell.Row.Cells["idEstadoWeb"].Value != System.DBNull.Value)
						{
							if (((int)e.Cell.Row.Cells["idEstadoWeb"].Value == 4 || (int)e.Cell.Row.Cells["idEstadoWeb"].Value == 1) )
							{
								if (  (int)e.Cell.Row.Cells["idFormaDePagoWeb"].Value == 5 && (int)e.Cell.Row.Cells["iEWeb"].Value  == 0)
								{
									if (	(int)e.Cell.Row.Cells["iEWeb"].Value  == 0 )
									{
										using (frmNotas NotasG = new frmNotas())
										{
											if (DialogResult.OK == NotasG.ShowDialog( ))
											{
												if (DialogResult.Yes == MessageBox.Show("Esta seguro que la transacción fue realizada con éxito para su posterior facturación.", "Point Technology", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2))
												{
													string sSQLWeb = string.Format("Exec [GuardaGestionesWeb] 0, {0},'{1}','{2}', {3}, '{4}', {5}, {6}",
					                                
														Convert.ToInt32 (e.Cell.Row.Cells["idWeb_Cabecera_Pedido"].Value),
														Convert.ToDateTime(DateTime.Today).ToString("yyyyMMdd HH:mm:ss"), 
														Convert.ToString (e.Cell.Row.Cells["telefono"].Value), 
														2,
														NotasG.txtNotas.Text,
														MenuLatinium.IdUsuario, 
														2);
								
													Funcion.EjecutaSQL(cdsSeteoF, sSQLWeb);

													string ssSQL = string.Format("update web_cabecera_pedido set idEstado_WEB = 7 where numero_pedido = {0}", Convert.ToInt32 (e.Cell.Row.Cells["idWeb_Cabecera_Pedido"].Value));
													Funcion.EjecutaSQL(cdsSeteoF, ssSQL);

													this.btnConsultar_Click (sender, e);

												}
											}
										}
										
									}
								}
							}
						}
					}
				}
			}
			#endregion Aprueba contabilidad

			#region Factuta Temporal
			if (e.Cell.Column.ToString() == "btnTemporal")
			{
				if (e.Cell.Row.Cells["idWeb_Cabecera_Pedido"].Value != System.DBNull.Value)
				{
					if ((int)e.Cell.Row.Cells["idWeb_Cabecera_Pedido"].Value > 0)
					{
						if (e.Cell.Row.Cells["idEstadoWeb"].Value != System.DBNull.Value)
						{
							if (((int)e.Cell.Row.Cells["idEstadoWeb"].Value == 4 || (int)e.Cell.Row.Cells["idEstadoWeb"].Value == 1) )
							{
								if ( (int)e.Cell.Row.Cells["idFormaDePagoWeb"].Value == 1 || (int)e.Cell.Row.Cells["idFormaDePagoWeb"].Value == 5 && (int)e.Cell.Row.Cells["iEWeb"].Value  == 7 || (int)e.Cell.Row.Cells["iEWeb"].Value  == 2)
								{
									if (	(int)e.Cell.Row.Cells["iEWeb"].Value  == 7 || (int)e.Cell.Row.Cells["iEWeb"].Value  == 2)
									{
										int iConDet= 0;
										iConDet = Funcion.iEscalarSQL(cdsSeteoF, string.Format("EXEC [ValidaDetalleVentasWeb] {0}",Convert.ToInt32 (e.Cell.Row.Cells["idWeb_Cabecera_Pedido"].Value)));	
                 
										if (iConDet == 0)
										{
											MessageBox.Show("El pedido seleccionado no tiene articulos cargados ", "Informacion", MessageBoxButtons.OK, MessageBoxIcon.Information);  
											return;
										}

							
										
										if (e.Cell.Row.Cells["idVendedor"].Value  == System.DBNull.Value)
										{
											MessageBox.Show("Seleccione el Vendedor", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
											this.ultraGrid1.ActiveRow.Selected = true;
											return;
										}		
				
										if ((int)e.Cell.Row.Cells["idVendedor"].Value  == 0)
										{
											MessageBox.Show("Seleccione el Vendedor", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
											this.ultraGrid1.ActiveRow.Selected = true;
											return;
										}	

								

										if (DialogResult.Yes == MessageBox.Show("Esta Seguro de generar  la Factura ", "Point Technology", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2))
										{
											bool bError = false;
                         								

											#region Guardar
											using (SqlConnection oConexion = new SqlConnection("Password=Bl45o6$9;Persist Security Info=True;User ID= " + MenuLatinium.stUsuario + ";Initial Catalog=" + MenuLatinium.stDirFacturacion + ";Data Source=" + MenuLatinium.stDirServidor + "; max pool size=1024"))
											{
												oConexion.Open();

												SqlTransaction oTransaction = null;
												SqlCommand oCmdActualiza = oConexion.CreateCommand();
												oCmdActualiza.Connection = oConexion;
												oCmdActualiza.CommandType = CommandType.Text;
												oCmdActualiza.CommandTimeout = 120;

												try
												{
													oTransaction = oConexion.BeginTransaction();
													oCmdActualiza.Transaction = oTransaction;

													string sSQLWeb = string.Format("Exec [GeneraFacturaTemporalWeb] {0}, {1} ",
														Convert.ToInt32 (e.Cell.Row.Cells["idWeb_Cabecera_Pedido"].Value), Convert.ToInt32 (e.Cell.Row.Cells["idVendedor"].Value));
													oCmdActualiza.CommandText = sSQLWeb;
													int iDCompra = (int)oCmdActualiza.ExecuteScalar();

													#region Detalle
													foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.ultraGrid4.Rows.All)
													{	
														if (dr.Cells["idArticulo"].Value != System.DBNull.Value)
														{
															if ((int) dr.Cells["idArticulo"].Value >0 )
															{
																string sSQLDetalle = string.Format("Exec GuardaDetalleFactura {0}, {1}, {2}, {3}, {4}, '{5}', '{6}', {7}, {8}, {9}, {10}, {11}, {12}, {13}, {14}, {15}, {16}, {17}, {18}, {19}, {20}",
																	0, 0, iDCompra,
																	(int)dr.Cells["idArticulo"].Value, 
																	(int)dr.Cells["idUnidad"].Value, 
																	dr.Cells["Codigo"].Value.ToString(), 
																	dr.Cells["Articulo"].Value,
																	(int)dr.Cells["Cantidad"].Value,
																	(decimal)dr.Cells["Precio"].Value, 
																	(decimal)dr.Cells["Impuesto"].Value, 
																	(decimal)dr.Cells["DescuentoArt"].Value,
																	(decimal)dr.Cells["DescuentoPorc"].Value, 
																	(decimal)dr.Cells["Subtotal"].Value,
																	false, (int)dr.Cells["Tipo"].Value, 
																	false, 
																	0, 
																	(bool)dr.Cells["Principal"].Value, 
																	false, 
																	(int)dr.Cells["idCombo"].Value, 0);

																oCmdActualiza.CommandText = sSQLDetalle;
																oCmdActualiza.ExecuteNonQuery();
															}
														}
													}
//													oCmdActualiza.ExecuteNonQuery(); ESTE ES EL QUE TE GENERA OTRA FILA
													#endregion Detalle	

													oTransaction.Commit();

													Cursor = Cursors.Default;	

													bError = true;
												
												}
												catch(Exception ex)
												{
													try
													{ 
														oTransaction.Rollback();

														MessageBox.Show(string.Format("Commit Exception Type : {0} {1}", ex.GetType(), ex.Message), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
													}
													catch (Exception ex2)
													{
														MessageBox.Show(string.Format("Rollback Exception Type: {0} {1}", ex2.GetType(), ex2.Message), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);							
													}
												}
												finally
												{
													oConexion.Close();
												}
											}
				
											#endregion Guardar

											if (bError)
											{
												string sSQLWebg = string.Format("Exec [GuardaGestionesWeb] 0, {0},'{1}','{2}', {3}, '{4}', {5}, {6}",
					                                
													Convert.ToInt32 (e.Cell.Row.Cells["idWeb_Cabecera_Pedido"].Value),
													Convert.ToDateTime(DateTime.Today).ToString("yyyyMMdd HH:mm:ss"), 
													Convert.ToString (e.Cell.Row.Cells["telefono"].Value), 
													2,
													"FACTURA TEMPORAL CREADA ",
													MenuLatinium.IdUsuario, 
													3);
								
												Funcion.EjecutaSQL(cdsSeteoF, sSQLWebg);


												string ssSQL = string.Format("update web_cabecera_pedido set idEstado_WEB = 3 where numero_pedido = {0}", Convert.ToInt32 (e.Cell.Row.Cells["idWeb_Cabecera_Pedido"].Value));
												Funcion.EjecutaSQL(cdsSeteoF, ssSQL);
												this.ultraGrid4.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Exec ConsultaDetalleFactura 0");
												this.ultraGrid2.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Exec ReporteDetalleVentasWeb 0");

												MessageBox.Show("Factura Temporal Creada Correctamente ", "Informacion", MessageBoxButtons.OK, MessageBoxIcon.Information);  

											}
											this.btnConsultar_Click (sender, e);

										}
									}
								}
							}
						}
					}
				}
			}
			#endregion Factura Temporal

			#region Gestión WEb 1salida
			/* tipo 2 confirmacion transferencia */
			if (e.Cell.Column.ToString() == "btnSalida")
			{
				if (e.Cell.Row.Cells["idWeb_Cabecera_Pedido"].Value != System.DBNull.Value)
				{
					if ((int)e.Cell.Row.Cells["idWeb_Cabecera_Pedido"].Value > 0)
					{
						if (e.Cell.Row.Cells["idEstadoWeb"].Value != System.DBNull.Value)
						{
							if (((int)e.Cell.Row.Cells["idEstadoWeb"].Value == 4 || (int)e.Cell.Row.Cells["idEstadoWeb"].Value == 1)&& (int)e.Cell.Row.Cells["idFormaDePagoWeb"].Value == 1 || (int)e.Cell.Row.Cells["idFormaDePagoWeb"].Value == 5)
							{
								if (	(int)e.Cell.Row.Cells["iEWeb"].Value  == 4)
								{
									using (frmNotas NotasG = new frmNotas())
									{
										if (DialogResult.OK == NotasG.ShowDialog( ))
										{
											
												string sSQLWeb = string.Format("Exec [GuardaGestionesWeb] 0, {0},'{1}','{2}', {3}, '{4}', {5}, {6}",
					                                
													Convert.ToInt32 (e.Cell.Row.Cells["idWeb_Cabecera_Pedido"].Value),
													Convert.ToDateTime(DateTime.Today).ToString("yyyyMMdd HH:mm:ss"), 
													Convert.ToString (e.Cell.Row.Cells["telefono"].Value), 
													5,
													NotasG.txtNotas.Text,
													MenuLatinium.IdUsuario, 
													5);
								
												Funcion.EjecutaSQL(cdsSeteoF, sSQLWeb);

												string ssSQL = string.Format("update web_cabecera_pedido set idEstado_WEB = 5 where numero_pedido = {0}", Convert.ToInt32 (e.Cell.Row.Cells["idWeb_Cabecera_Pedido"].Value));
												Funcion.EjecutaSQL(cdsSeteoF, ssSQL);

												this.btnConsultar_Click (sender, e);

										}
									}
								}
							}

						}
					}
				}
			}
			#endregion Gestión WEb 1salida

			#region Gestión WEb 1salida
			/* tipo 2 confirmacion transferencia */
			if (e.Cell.Column.ToString() == "btnEntregado")
			{
				if (e.Cell.Row.Cells["idWeb_Cabecera_Pedido"].Value != System.DBNull.Value)
				{
					if ((int)e.Cell.Row.Cells["idWeb_Cabecera_Pedido"].Value > 0)
					{
						if (e.Cell.Row.Cells["idEstadoWeb"].Value != System.DBNull.Value)
						{
							if (((int)e.Cell.Row.Cells["idEstadoWeb"].Value == 4 || (int)e.Cell.Row.Cells["idEstadoWeb"].Value == 1)&& (int)e.Cell.Row.Cells["idFormaDePagoWeb"].Value == 1 || (int)e.Cell.Row.Cells["idFormaDePagoWeb"].Value == 5)
							{
								if (	(int)e.Cell.Row.Cells["iEWeb"].Value  == 5)
								{
									using (frmNotas NotasG = new frmNotas())
									{
										if (DialogResult.OK == NotasG.ShowDialog( ))
										{
											
											string sSQLWeb = string.Format("Exec [GuardaGestionesWeb] 0, {0},'{1}','{2}', {3}, '{4}', {5}, {6}",
					                                
												Convert.ToInt32 (e.Cell.Row.Cells["idWeb_Cabecera_Pedido"].Value),
												Convert.ToDateTime(DateTime.Today).ToString("yyyyMMdd HH:mm:ss"), 
												Convert.ToString (e.Cell.Row.Cells["telefono"].Value), 
												6,
												NotasG.txtNotas.Text,
												MenuLatinium.IdUsuario, 
												6);
								
											Funcion.EjecutaSQL(cdsSeteoF, sSQLWeb);

											string ssSQL = string.Format("update web_cabecera_pedido set idEstado_WEB = 6 where numero_pedido = {0}", Convert.ToInt32 (e.Cell.Row.Cells["idWeb_Cabecera_Pedido"].Value));
											Funcion.EjecutaSQL(cdsSeteoF, ssSQL);

											this.btnConsultar_Click (sender, e);

										}
									}
								}
							}

						}
					}
				}
			}
			#endregion Gestión WEb 1salida


		}

		private void cmbEstado_ValueChanged(object sender, System.EventArgs e)
		{
			if (this.cmbEstado.ActiveRow != null)
			{
				if ((int)this.cmbEstado.Value == 4 || (int)this.cmbEstado.Value == 1 )
				{
					this.cmbEsatdoPedido.Visible = true;
					this.cmbEsatdoPedido.Value = 0;
					this.label4.Visible = true;
				}
				else
				{
					this.cmbEsatdoPedido.Value = 8;
					this.cmbEsatdoPedido.Visible = false;
					this.label4.Visible = false;

				}
			}
			else
			{
				this.cmbEsatdoPedido.Value = 8;
				this.cmbEsatdoPedido.Visible = false;
				this.label4.Visible = false;

			}
		}

		private void cmbEntidad_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter) this.cmbEstado.Focus();
		}

		private void cmbEntidad_ValueChanged(object sender, System.EventArgs e)
		{
			if (this.cmbEntidad.ActiveRow != null)
			{
				if((int)this.cmbEntidad.Value == 2 )
				{
					this.cmbEstado.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Select idEstadoWeb, Nombre from EstadoWeb  where idEstadoWeb not in (1) order by nombre ");
					}
				if((int)this.cmbEntidad.Value == 1)
				{
					this.cmbEstado.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Select idEstadoWeb, Nombre from EstadoWeb where idEstadoWeb not in (4) order by nombre ");
				}
				if((int)this.cmbEntidad.Value == 3  )
				{
					this.cmbEstado.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Select idEstadoWeb, Nombre from EstadoWeb  where idEstadoWeb  in (2,4) order by nombre ");
				}
				
			}
		}

		private void ultraGrid4_AfterCellUpdate(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			if (e.Cell.Column.ToString() == "Codigo")
			{
				FuncionesInventario.FacturacionCargaArticulosWeb(sender, e, this.ultraGrid4, 1, 3, 82, 82, 
					0, 0, 0, 0, DateTime.Today , 1, 
					true, false, false, false, 0 ,0, 0, 0, 0, false, false, 0, cdsSeteoF,0);
			}	
		}

		private void ultraGrid1_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}
	}
}
