using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using DevExpress.XtraEditors;
using System.Data;
using System.Data.SqlClient;
using CrystalDecisions.Shared;
using System.Net;
using System.Web;
using System.Xml;
using System.Xml.Schema;
using System.IO;
using System.Diagnostics;
using Infragistics.Shared;
using Infragistics.Win;
using Infragistics.Win.UltraWinGrid;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de frmReporteCompraProtegidaDocumentoEnviado.
	/// </summary>
	public class frmReporteCompraProtegidaDocumentoEnviado : System.Windows.Forms.Form
	{
		private System.Windows.Forms.Button btnSalir;
		private System.Windows.Forms.Button btnGuardar;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource1;
		private C1.Data.C1DataSet cdsSeteoF;
		private System.Windows.Forms.GroupBox groupBox1;
		private Infragistics.Win.UltraWinGrid.UltraGrid uGridDocumentos;
		private System.Windows.Forms.GroupBox groupBox2;
		private System.Windows.Forms.Label lblBodega;
		private System.Windows.Forms.Label label53;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtDesde;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtHasta;
		private System.Windows.Forms.GroupBox groupBox3;
		private System.Windows.Forms.Label label1;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtTotal;
		private System.Windows.Forms.ToolTip toolTip1;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource udsCuenta;
		private System.Windows.Forms.GroupBox groupBox4;
		private System.Windows.Forms.Label label6;
		private System.Windows.Forms.Label label3;
		private System.Windows.Forms.Label label4;
		public Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtFecha;
		public Infragistics.Win.UltraWinEditors.UltraTextEditor txtCheque;
		public Infragistics.Win.UltraWinEditors.UltraTextEditor txtNotas;
		public Infragistics.Win.UltraWinGrid.UltraCombo cmbCuentas;
		private System.Windows.Forms.Label label5;
		private System.Windows.Forms.GroupBox groupBox5;
		private System.Windows.Forms.Button btnConsultar;
		public System.Windows.Forms.Label label12;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtNumero;
		private System.Windows.Forms.Label label2;
		private System.Windows.Forms.Label lblContador;
		private System.ComponentModel.IContainer components;

		public frmReporteCompraProtegidaDocumentoEnviado()
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
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmReporteCompraProtegidaDocumentoEnviado));
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCompra");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idDetCompra");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn3 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCliente");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn4 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Cedula");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn5 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Cliente");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn6 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Direccion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn7 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Fecha_Factura");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn8 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Numero");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn9 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Almacen");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn10 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("FormaPago");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn11 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Seguro");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn12 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Codigo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn13 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Articulo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn14 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Linea");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn15 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Serial");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn16 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Precio_Contado");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn17 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Costo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn18 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Valor");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn19 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Valor_Seguro");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn20 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Estado");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn21 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("EstadoSeguro");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn22 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("FechaNotificacion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn23 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("TipoNotificacion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn24 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Tipo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn25 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idProforma");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn26 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Valor_Asegurado");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn27 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("TipoCobertura");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn28 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Cobertura");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn29 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("ProformaST");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn30 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Deducible");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn31 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCompraD");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn32 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCompraP");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn33 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("FechaCierre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn34 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idAnticipoD");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn35 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idAnticipoP");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn36 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("DirectorioCP");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn37 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("SelDocumentoEnviado");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn38 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("ValorDocumentoEnviado");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn39 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCompraFacturaAnterior");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn40 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idArticuloGE");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn41 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("CodigoReemplazo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn42 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idArticulo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn43 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idSubGrupo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn44 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Precio3");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn45 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("DescuentoArt");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn46 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idAnticipo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn47 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("bTieneFacturaAnterior");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn48 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("bTieneDesarrolloActual");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn49 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("FechaSeguroLiquidado");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn50 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("bTieneDepositoAnterior");
			Infragistics.Win.Appearance appearance1 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCompra");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idDetCompra");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCliente");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cedula");
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cliente");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Direccion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Fecha_Factura");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Numero");
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Almacen");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FormaPago");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn11 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Seguro");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn12 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn13 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Articulo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn14 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Linea");
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn15 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Serial");
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn16 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Precio_Contado");
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn17 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Costo");
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn18 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Valor");
			Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn19 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Valor_Seguro");
			Infragistics.Win.Appearance appearance11 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn20 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Estado");
			Infragistics.Win.Appearance appearance12 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn21 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("EstadoSeguro");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn22 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaNotificacion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn23 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("TipoNotificacion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn24 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Tipo");
			Infragistics.Win.Appearance appearance13 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn25 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idProforma");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn26 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Valor_Asegurado");
			Infragistics.Win.Appearance appearance14 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn27 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("TipoCobertura");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn28 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cobertura");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn29 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ProformaST");
			Infragistics.Win.Appearance appearance15 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn30 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Deducible");
			Infragistics.Win.Appearance appearance16 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn31 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCompraD");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn32 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCompraP");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn33 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaCierre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn34 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idAnticipoD");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn35 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idAnticipoP");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn36 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("DirectorioCP");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn37 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("SelDocumentoEnviado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn38 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ValorDocumentoEnviado");
			Infragistics.Win.Appearance appearance17 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn39 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCompraFacturaAnterior");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn40 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idArticuloGE");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn41 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CodigoReemplazo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn42 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idArticulo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn43 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idSubGrupo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn44 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Precio3");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn45 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("DescuentoArt");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn46 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idAnticipo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn47 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("bTieneFacturaAnterior");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn48 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("bTieneDesarrolloActual");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn49 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaSeguroLiquidado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn50 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("bTieneDepositoAnterior");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn51 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("btnGestion", 0);
			Infragistics.Win.Appearance appearance18 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn52 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("btnProforma", 1);
			Infragistics.Win.Appearance appearance19 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn53 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Row", 2);
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings1 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "ValorDocumentoEnviado", 37, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "ValorDocumentoEnviado", 37, true);
			Infragistics.Win.Appearance appearance20 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance21 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance22 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance23 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance24 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance25 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance26 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance27 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton1 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.Appearance appearance28 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance29 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance30 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn54 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCuenta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn55 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Descripcion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn51 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCuenta");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn52 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Descripcion");
			Infragistics.Win.Appearance appearance31 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton2 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.Appearance appearance32 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton3 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			this.btnSalir = new System.Windows.Forms.Button();
			this.btnGuardar = new System.Windows.Forms.Button();
			this.ultraDataSource1 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.groupBox1 = new System.Windows.Forms.GroupBox();
			this.groupBox5 = new System.Windows.Forms.GroupBox();
			this.btnConsultar = new System.Windows.Forms.Button();
			this.label12 = new System.Windows.Forms.Label();
			this.txtNumero = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label2 = new System.Windows.Forms.Label();
			this.lblContador = new System.Windows.Forms.Label();
			this.label1 = new System.Windows.Forms.Label();
			this.txtTotal = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.uGridDocumentos = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.groupBox2 = new System.Windows.Forms.GroupBox();
			this.groupBox4 = new System.Windows.Forms.GroupBox();
			this.label6 = new System.Windows.Forms.Label();
			this.label3 = new System.Windows.Forms.Label();
			this.label4 = new System.Windows.Forms.Label();
			this.dtFecha = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.txtCheque = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtNotas = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.cmbCuentas = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.udsCuenta = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.label5 = new System.Windows.Forms.Label();
			this.lblBodega = new System.Windows.Forms.Label();
			this.label53 = new System.Windows.Forms.Label();
			this.dtDesde = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.dtHasta = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.groupBox3 = new System.Windows.Forms.GroupBox();
			this.toolTip1 = new System.Windows.Forms.ToolTip(this.components);
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			this.groupBox1.SuspendLayout();
			this.groupBox5.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.txtNumero)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtTotal)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridDocumentos)).BeginInit();
			this.groupBox2.SuspendLayout();
			this.groupBox4.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.dtFecha)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCheque)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNotas)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCuentas)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.udsCuenta)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtDesde)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtHasta)).BeginInit();
			this.SuspendLayout();
			// 
			// btnSalir
			// 
			this.btnSalir.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btnSalir.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(230)), ((System.Byte)(230)), ((System.Byte)(230)));
			this.btnSalir.CausesValidation = false;
			this.btnSalir.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.btnSalir.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(50)), ((System.Byte)(50)), ((System.Byte)(50)));
			this.btnSalir.Image = ((System.Drawing.Image)(resources.GetObject("btnSalir.Image")));
			this.btnSalir.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnSalir.Location = new System.Drawing.Point(608, 56);
			this.btnSalir.Name = "btnSalir";
			this.btnSalir.Size = new System.Drawing.Size(80, 23);
			this.btnSalir.TabIndex = 982;
			this.btnSalir.Text = "&Salir";
			this.btnSalir.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.toolTip1.SetToolTip(this.btnSalir, "Salir");
			this.btnSalir.Click += new System.EventHandler(this.btnSalir_Click);
			// 
			// btnGuardar
			// 
			this.btnGuardar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btnGuardar.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(230)), ((System.Byte)(230)), ((System.Byte)(230)));
			this.btnGuardar.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(50)), ((System.Byte)(50)), ((System.Byte)(50)));
			this.btnGuardar.Image = ((System.Drawing.Image)(resources.GetObject("btnGuardar.Image")));
			this.btnGuardar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnGuardar.Location = new System.Drawing.Point(608, 16);
			this.btnGuardar.Name = "btnGuardar";
			this.btnGuardar.Size = new System.Drawing.Size(82, 23);
			this.btnGuardar.TabIndex = 981;
			this.btnGuardar.Text = "&Aceptar";
			this.btnGuardar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.toolTip1.SetToolTip(this.btnGuardar, "Aceptar (generará el recibo de cobro de los clientes seleccionados).");
			this.btnGuardar.Click += new System.EventHandler(this.btnGuardar_Click);
			// 
			// ultraDataSource1
			// 
			ultraDataColumn1.DataType = typeof(int);
			ultraDataColumn2.DataType = typeof(int);
			ultraDataColumn3.DataType = typeof(int);
			ultraDataColumn7.DataType = typeof(System.DateTime);
			ultraDataColumn16.DataType = typeof(System.Decimal);
			ultraDataColumn17.DataType = typeof(System.Decimal);
			ultraDataColumn18.DataType = typeof(System.Decimal);
			ultraDataColumn19.DataType = typeof(System.Decimal);
			ultraDataColumn21.DataType = typeof(int);
			ultraDataColumn22.DataType = typeof(System.DateTime);
			ultraDataColumn23.DataType = typeof(int);
			ultraDataColumn25.DataType = typeof(int);
			ultraDataColumn26.DataType = typeof(System.Decimal);
			ultraDataColumn27.DataType = typeof(bool);
			ultraDataColumn29.DataType = typeof(System.Decimal);
			ultraDataColumn31.DataType = typeof(int);
			ultraDataColumn32.DataType = typeof(int);
			ultraDataColumn33.DataType = typeof(System.DateTime);
			ultraDataColumn34.DataType = typeof(int);
			ultraDataColumn35.DataType = typeof(int);
			ultraDataColumn37.DataType = typeof(bool);
			ultraDataColumn37.DefaultValue = false;
			ultraDataColumn38.DataType = typeof(System.Decimal);
			ultraDataColumn39.DataType = typeof(int);
			ultraDataColumn42.DataType = typeof(int);
			ultraDataColumn44.DataType = typeof(System.Decimal);
			ultraDataColumn45.DataType = typeof(System.Decimal);
			ultraDataColumn46.DataType = typeof(int);
			ultraDataColumn47.DataType = typeof(bool);
			ultraDataColumn48.DataType = typeof(bool);
			ultraDataColumn49.DataType = typeof(System.DateTime);
			ultraDataColumn50.DataType = typeof(bool);
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
																																 ultraDataColumn18,
																																 ultraDataColumn19,
																																 ultraDataColumn20,
																																 ultraDataColumn21,
																																 ultraDataColumn22,
																																 ultraDataColumn23,
																																 ultraDataColumn24,
																																 ultraDataColumn25,
																																 ultraDataColumn26,
																																 ultraDataColumn27,
																																 ultraDataColumn28,
																																 ultraDataColumn29,
																																 ultraDataColumn30,
																																 ultraDataColumn31,
																																 ultraDataColumn32,
																																 ultraDataColumn33,
																																 ultraDataColumn34,
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
																																 ultraDataColumn50});
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
			// 
			// groupBox1
			// 
			this.groupBox1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.groupBox1.Controls.Add(this.groupBox5);
			this.groupBox1.Controls.Add(this.label1);
			this.groupBox1.Controls.Add(this.txtTotal);
			this.groupBox1.Controls.Add(this.uGridDocumentos);
			this.groupBox1.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.groupBox1.Location = new System.Drawing.Point(16, 272);
			this.groupBox1.Name = "groupBox1";
			this.groupBox1.Size = new System.Drawing.Size(672, 384);
			this.groupBox1.TabIndex = 983;
			this.groupBox1.TabStop = false;
			this.groupBox1.Text = "Clientes";
			this.groupBox1.Enter += new System.EventHandler(this.groupBox1_Enter);
			// 
			// groupBox5
			// 
			this.groupBox5.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.groupBox5.Controls.Add(this.btnConsultar);
			this.groupBox5.Controls.Add(this.label12);
			this.groupBox5.Controls.Add(this.txtNumero);
			this.groupBox5.Controls.Add(this.label2);
			this.groupBox5.Controls.Add(this.lblContador);
			this.groupBox5.Location = new System.Drawing.Point(16, 24);
			this.groupBox5.Name = "groupBox5";
			this.groupBox5.Size = new System.Drawing.Size(632, 80);
			this.groupBox5.TabIndex = 1055;
			this.groupBox5.TabStop = false;
			this.groupBox5.Text = "Búsqueda";
			this.groupBox5.Enter += new System.EventHandler(this.groupBox5_Enter);
			// 
			// btnConsultar
			// 
			this.btnConsultar.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(230)), ((System.Byte)(230)), ((System.Byte)(230)));
			this.btnConsultar.CausesValidation = false;
			this.btnConsultar.Image = ((System.Drawing.Image)(resources.GetObject("btnConsultar.Image")));
			this.btnConsultar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnConsultar.Location = new System.Drawing.Point(296, 48);
			this.btnConsultar.Name = "btnConsultar";
			this.btnConsultar.Size = new System.Drawing.Size(78, 24);
			this.btnConsultar.TabIndex = 1059;
			this.btnConsultar.Text = "&Consultar";
			this.btnConsultar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.toolTip1.SetToolTip(this.btnConsultar, "Consultar");
			this.btnConsultar.Click += new System.EventHandler(this.btnConsultar_Click);
			// 
			// label12
			// 
			this.label12.AutoSize = true;
			this.label12.Location = new System.Drawing.Point(16, 48);
			this.label12.Name = "label12";
			this.label12.Size = new System.Drawing.Size(39, 16);
			this.label12.TabIndex = 1058;
			this.label12.Text = "Buscar";
			this.label12.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtNumero
			// 
			appearance1.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtNumero.Appearance = appearance1;
			this.txtNumero.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtNumero.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNumero.Location = new System.Drawing.Point(72, 48);
			this.txtNumero.Name = "txtNumero";
			this.txtNumero.Size = new System.Drawing.Size(208, 21);
			this.txtNumero.TabIndex = 1057;
			this.toolTip1.SetToolTip(this.txtNumero, "Busqueda por número, responsable, asignado, título.");
			this.txtNumero.ValueChanged += new System.EventHandler(this.txtNumero_ValueChanged);
			// 
			// label2
			// 
			this.label2.Location = new System.Drawing.Point(8, 24);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(288, 23);
			this.label2.TabIndex = 1055;
			this.label2.Text = "* Seleccionar el cliente para generar su recibo de cobro:";
			// 
			// lblContador
			// 
			this.lblContador.AutoSize = true;
			this.lblContador.Font = new System.Drawing.Font("Tahoma", 10F);
			this.lblContador.ForeColor = System.Drawing.Color.Firebrick;
			this.lblContador.Location = new System.Drawing.Point(392, 48);
			this.lblContador.Name = "lblContador";
			this.lblContador.Size = new System.Drawing.Size(0, 20);
			this.lblContador.TabIndex = 1056;
			this.lblContador.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label1
			// 
			this.label1.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.label1.AutoSize = true;
			this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label1.Location = new System.Drawing.Point(360, 344);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(102, 16);
			this.label1.TabIndex = 983;
			this.label1.Text = "Total seleccionado:";
			this.label1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtTotal
			// 
			this.txtTotal.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			appearance2.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtTotal.Appearance = appearance2;
			this.txtTotal.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtTotal.Enabled = false;
			this.txtTotal.Font = new System.Drawing.Font("Microsoft Sans Serif", 15.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtTotal.FormatString = "#,##0.00";
			this.txtTotal.Location = new System.Drawing.Point(472, 336);
			this.txtTotal.Name = "txtTotal";
			this.txtTotal.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtTotal.PromptChar = ' ';
			this.txtTotal.Size = new System.Drawing.Size(176, 34);
			this.txtTotal.TabIndex = 982;
			// 
			// uGridDocumentos
			// 
			this.uGridDocumentos.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.uGridDocumentos.Cursor = System.Windows.Forms.Cursors.Default;
			this.uGridDocumentos.DataSource = this.ultraDataSource1;
			appearance3.BackColor = System.Drawing.Color.White;
			this.uGridDocumentos.DisplayLayout.Appearance = appearance3;
			ultraGridBand1.AutoPreviewEnabled = true;
			ultraGridColumn1.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn1.Header.VisiblePosition = 3;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn2.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn2.Header.VisiblePosition = 2;
			ultraGridColumn2.Hidden = true;
			ultraGridColumn3.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn3.Header.VisiblePosition = 4;
			ultraGridColumn3.Hidden = true;
			ultraGridColumn4.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance4.ImageHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn4.CellAppearance = appearance4;
			ultraGridColumn4.Header.Caption = "Cédula";
			ultraGridColumn4.Header.VisiblePosition = 5;
			ultraGridColumn4.RowLayoutColumnInfo.OriginX = 4;
			ultraGridColumn4.RowLayoutColumnInfo.OriginY = 0;
			ultraGridColumn4.RowLayoutColumnInfo.PreferredCellSize = new System.Drawing.Size(90, 0);
			ultraGridColumn4.RowLayoutColumnInfo.SpanX = 2;
			ultraGridColumn4.RowLayoutColumnInfo.SpanY = 2;
			ultraGridColumn4.Width = 109;
			ultraGridColumn5.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn5.Header.VisiblePosition = 6;
			ultraGridColumn5.RowLayoutColumnInfo.OriginX = 6;
			ultraGridColumn5.RowLayoutColumnInfo.OriginY = 0;
			ultraGridColumn5.RowLayoutColumnInfo.PreferredCellSize = new System.Drawing.Size(323, 0);
			ultraGridColumn5.RowLayoutColumnInfo.SpanX = 2;
			ultraGridColumn5.RowLayoutColumnInfo.SpanY = 2;
			ultraGridColumn5.Width = 290;
			ultraGridColumn6.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn6.Header.Caption = "Dirección";
			ultraGridColumn6.Header.VisiblePosition = 9;
			ultraGridColumn6.RowLayoutColumnInfo.OriginX = 10;
			ultraGridColumn6.RowLayoutColumnInfo.OriginY = 0;
			ultraGridColumn6.RowLayoutColumnInfo.SpanX = 2;
			ultraGridColumn6.RowLayoutColumnInfo.SpanY = 2;
			ultraGridColumn6.Width = 160;
			ultraGridColumn7.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn7.Header.Caption = "F. Factura";
			ultraGridColumn7.Header.VisiblePosition = 10;
			ultraGridColumn7.RowLayoutColumnInfo.OriginX = 12;
			ultraGridColumn7.RowLayoutColumnInfo.OriginY = 0;
			ultraGridColumn7.RowLayoutColumnInfo.SpanX = 2;
			ultraGridColumn7.RowLayoutColumnInfo.SpanY = 2;
			ultraGridColumn7.Width = 80;
			ultraGridColumn8.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance5.ImageHAlign = Infragistics.Win.HAlign.Center;
			appearance5.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn8.CellAppearance = appearance5;
			ultraGridColumn8.Header.Caption = "Número";
			ultraGridColumn8.Header.VisiblePosition = 11;
			ultraGridColumn8.RowLayoutColumnInfo.OriginX = 14;
			ultraGridColumn8.RowLayoutColumnInfo.OriginY = 0;
			ultraGridColumn8.RowLayoutColumnInfo.SpanX = 2;
			ultraGridColumn8.RowLayoutColumnInfo.SpanY = 2;
			ultraGridColumn8.Width = 90;
			ultraGridColumn9.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn9.Header.Caption = "Local";
			ultraGridColumn9.Header.VisiblePosition = 12;
			ultraGridColumn9.RowLayoutColumnInfo.OriginX = 16;
			ultraGridColumn9.RowLayoutColumnInfo.OriginY = 0;
			ultraGridColumn9.RowLayoutColumnInfo.SpanX = 2;
			ultraGridColumn9.RowLayoutColumnInfo.SpanY = 2;
			ultraGridColumn9.Width = 120;
			ultraGridColumn10.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn10.Header.Caption = "Forma De Pago";
			ultraGridColumn10.Header.VisiblePosition = 13;
			ultraGridColumn10.RowLayoutColumnInfo.OriginX = 18;
			ultraGridColumn10.RowLayoutColumnInfo.OriginY = 0;
			ultraGridColumn10.RowLayoutColumnInfo.SpanX = 2;
			ultraGridColumn10.RowLayoutColumnInfo.SpanY = 2;
			ultraGridColumn10.Width = 140;
			ultraGridColumn11.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn11.Header.VisiblePosition = 14;
			ultraGridColumn11.RowLayoutColumnInfo.OriginX = 20;
			ultraGridColumn11.RowLayoutColumnInfo.OriginY = 0;
			ultraGridColumn11.RowLayoutColumnInfo.SpanX = 2;
			ultraGridColumn11.RowLayoutColumnInfo.SpanY = 2;
			ultraGridColumn12.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn12.Header.Caption = "Código";
			ultraGridColumn12.Header.VisiblePosition = 15;
			ultraGridColumn12.RowLayoutColumnInfo.OriginX = 22;
			ultraGridColumn12.RowLayoutColumnInfo.OriginY = 0;
			ultraGridColumn12.RowLayoutColumnInfo.SpanX = 2;
			ultraGridColumn12.RowLayoutColumnInfo.SpanY = 2;
			ultraGridColumn13.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn13.Header.Caption = "Artículo";
			ultraGridColumn13.Header.VisiblePosition = 16;
			ultraGridColumn13.RowLayoutColumnInfo.OriginX = 24;
			ultraGridColumn13.RowLayoutColumnInfo.OriginY = 0;
			ultraGridColumn13.RowLayoutColumnInfo.SpanX = 2;
			ultraGridColumn13.RowLayoutColumnInfo.SpanY = 2;
			ultraGridColumn14.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance6.ImageHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn14.CellAppearance = appearance6;
			ultraGridColumn14.Header.Caption = "Línea";
			ultraGridColumn14.Header.VisiblePosition = 17;
			ultraGridColumn14.RowLayoutColumnInfo.OriginX = 26;
			ultraGridColumn14.RowLayoutColumnInfo.OriginY = 0;
			ultraGridColumn14.RowLayoutColumnInfo.SpanX = 2;
			ultraGridColumn14.RowLayoutColumnInfo.SpanY = 2;
			ultraGridColumn15.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance7.ImageHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn15.CellAppearance = appearance7;
			ultraGridColumn15.Header.VisiblePosition = 18;
			ultraGridColumn15.RowLayoutColumnInfo.OriginX = 28;
			ultraGridColumn15.RowLayoutColumnInfo.OriginY = 0;
			ultraGridColumn15.RowLayoutColumnInfo.SpanX = 2;
			ultraGridColumn15.RowLayoutColumnInfo.SpanY = 2;
			ultraGridColumn15.Width = 120;
			ultraGridColumn16.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance8.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn16.CellAppearance = appearance8;
			ultraGridColumn16.Format = "#,##0.00";
			ultraGridColumn16.Header.Caption = "Precio Contado";
			ultraGridColumn16.Header.VisiblePosition = 19;
			ultraGridColumn16.RowLayoutColumnInfo.OriginX = 30;
			ultraGridColumn16.RowLayoutColumnInfo.OriginY = 0;
			ultraGridColumn16.RowLayoutColumnInfo.SpanX = 2;
			ultraGridColumn16.RowLayoutColumnInfo.SpanY = 2;
			ultraGridColumn16.Width = 120;
			ultraGridColumn17.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance9.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn17.CellAppearance = appearance9;
			ultraGridColumn17.Format = "#,##0.00";
			ultraGridColumn17.Header.VisiblePosition = 20;
			ultraGridColumn17.RowLayoutColumnInfo.OriginX = 32;
			ultraGridColumn17.RowLayoutColumnInfo.OriginY = 0;
			ultraGridColumn17.RowLayoutColumnInfo.SpanX = 2;
			ultraGridColumn17.RowLayoutColumnInfo.SpanY = 2;
			ultraGridColumn17.Width = 90;
			ultraGridColumn18.AutoEdit = false;
			ultraGridColumn18.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance10.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn18.CellAppearance = appearance10;
			ultraGridColumn18.Format = "#,##0.00";
			ultraGridColumn18.Header.VisiblePosition = 21;
			ultraGridColumn18.RowLayoutColumnInfo.OriginX = 34;
			ultraGridColumn18.RowLayoutColumnInfo.OriginY = 0;
			ultraGridColumn18.RowLayoutColumnInfo.SpanX = 2;
			ultraGridColumn18.RowLayoutColumnInfo.SpanY = 2;
			ultraGridColumn18.Width = 105;
			ultraGridColumn19.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance11.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn19.CellAppearance = appearance11;
			ultraGridColumn19.Format = "#,##0.00";
			ultraGridColumn19.Header.Caption = "Valor Seguro";
			ultraGridColumn19.Header.VisiblePosition = 22;
			ultraGridColumn19.RowLayoutColumnInfo.OriginX = 36;
			ultraGridColumn19.RowLayoutColumnInfo.OriginY = 0;
			ultraGridColumn19.RowLayoutColumnInfo.SpanX = 2;
			ultraGridColumn19.RowLayoutColumnInfo.SpanY = 2;
			ultraGridColumn19.Width = 90;
			ultraGridColumn20.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance12.ImageHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn20.CellAppearance = appearance12;
			ultraGridColumn20.Header.VisiblePosition = 25;
			ultraGridColumn20.Hidden = true;
			ultraGridColumn20.RowLayoutColumnInfo.OriginX = 42;
			ultraGridColumn20.RowLayoutColumnInfo.OriginY = 0;
			ultraGridColumn20.RowLayoutColumnInfo.SpanX = 2;
			ultraGridColumn20.RowLayoutColumnInfo.SpanY = 2;
			ultraGridColumn20.Width = 160;
			ultraGridColumn21.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn21.Header.VisiblePosition = 26;
			ultraGridColumn21.Hidden = true;
			ultraGridColumn22.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn22.Format = "dd/MM/yyyy HH:mm";
			ultraGridColumn22.Header.Caption = "Fecha Notificación";
			ultraGridColumn22.Header.VisiblePosition = 30;
			ultraGridColumn22.RowLayoutColumnInfo.OriginX = 46;
			ultraGridColumn22.RowLayoutColumnInfo.OriginY = 0;
			ultraGridColumn22.RowLayoutColumnInfo.SpanX = 2;
			ultraGridColumn22.RowLayoutColumnInfo.SpanY = 2;
			ultraGridColumn23.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn23.Header.VisiblePosition = 27;
			ultraGridColumn23.Hidden = true;
			ultraGridColumn24.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance13.ImageHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn24.CellAppearance = appearance13;
			ultraGridColumn24.Header.VisiblePosition = 28;
			ultraGridColumn24.RowLayoutColumnInfo.OriginX = 44;
			ultraGridColumn24.RowLayoutColumnInfo.OriginY = 0;
			ultraGridColumn24.RowLayoutColumnInfo.SpanX = 2;
			ultraGridColumn24.RowLayoutColumnInfo.SpanY = 2;
			ultraGridColumn25.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn25.Header.VisiblePosition = 35;
			ultraGridColumn25.Hidden = true;
			ultraGridColumn26.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance14.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn26.CellAppearance = appearance14;
			ultraGridColumn26.Format = "#,##0.00";
			ultraGridColumn26.Header.Caption = "Valor Asegurado";
			ultraGridColumn26.Header.VisiblePosition = 23;
			ultraGridColumn26.RowLayoutColumnInfo.OriginX = 38;
			ultraGridColumn26.RowLayoutColumnInfo.OriginY = 0;
			ultraGridColumn26.RowLayoutColumnInfo.SpanX = 2;
			ultraGridColumn26.RowLayoutColumnInfo.SpanY = 2;
			ultraGridColumn26.Width = 100;
			ultraGridColumn27.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn27.Header.VisiblePosition = 29;
			ultraGridColumn27.Hidden = true;
			ultraGridColumn28.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn28.Header.VisiblePosition = 31;
			ultraGridColumn28.RowLayoutColumnInfo.OriginX = 48;
			ultraGridColumn28.RowLayoutColumnInfo.OriginY = 0;
			ultraGridColumn28.RowLayoutColumnInfo.SpanX = 2;
			ultraGridColumn28.RowLayoutColumnInfo.SpanY = 2;
			ultraGridColumn28.Width = 120;
			ultraGridColumn29.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance15.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn29.CellAppearance = appearance15;
			ultraGridColumn29.Format = "#,##0.00";
			ultraGridColumn29.Header.Caption = "Prof: Serv. Técnico";
			ultraGridColumn29.Header.VisiblePosition = 32;
			ultraGridColumn29.RowLayoutColumnInfo.OriginX = 50;
			ultraGridColumn29.RowLayoutColumnInfo.OriginY = 0;
			ultraGridColumn29.RowLayoutColumnInfo.SpanX = 2;
			ultraGridColumn29.RowLayoutColumnInfo.SpanY = 2;
			ultraGridColumn29.Width = 140;
			ultraGridColumn30.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance16.ImageHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn30.CellAppearance = appearance16;
			ultraGridColumn30.Header.VisiblePosition = 24;
			ultraGridColumn30.Hidden = true;
			ultraGridColumn30.RowLayoutColumnInfo.OriginX = 40;
			ultraGridColumn30.RowLayoutColumnInfo.OriginY = 0;
			ultraGridColumn30.RowLayoutColumnInfo.SpanX = 2;
			ultraGridColumn30.RowLayoutColumnInfo.SpanY = 2;
			ultraGridColumn30.Width = 60;
			ultraGridColumn31.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn31.Header.VisiblePosition = 34;
			ultraGridColumn31.Hidden = true;
			ultraGridColumn32.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn32.Header.VisiblePosition = 36;
			ultraGridColumn32.Hidden = true;
			ultraGridColumn33.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn33.Format = "dd/MM/yyyy HH:mm";
			ultraGridColumn33.Header.Caption = "Fecha Cierre";
			ultraGridColumn33.Header.VisiblePosition = 33;
			ultraGridColumn33.Hidden = true;
			ultraGridColumn33.RowLayoutColumnInfo.OriginX = 52;
			ultraGridColumn33.RowLayoutColumnInfo.OriginY = 0;
			ultraGridColumn33.RowLayoutColumnInfo.SpanX = 2;
			ultraGridColumn33.RowLayoutColumnInfo.SpanY = 2;
			ultraGridColumn34.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn34.Header.VisiblePosition = 37;
			ultraGridColumn34.Hidden = true;
			ultraGridColumn35.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn35.Header.VisiblePosition = 51;
			ultraGridColumn35.Hidden = true;
			ultraGridColumn36.Header.VisiblePosition = 38;
			ultraGridColumn36.Hidden = true;
			ultraGridColumn37.AutoEdit = false;
			ultraGridColumn37.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn37.Header.Caption = "...";
			ultraGridColumn37.Header.VisiblePosition = 8;
			ultraGridColumn37.RowLayoutColumnInfo.OriginX = 2;
			ultraGridColumn37.RowLayoutColumnInfo.OriginY = 0;
			ultraGridColumn37.RowLayoutColumnInfo.PreferredCellSize = new System.Drawing.Size(31, 0);
			ultraGridColumn37.RowLayoutColumnInfo.SpanX = 2;
			ultraGridColumn37.RowLayoutColumnInfo.SpanY = 2;
			ultraGridColumn37.Width = 95;
			appearance17.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn38.CellAppearance = appearance17;
			ultraGridColumn38.Format = "#,##0.00";
			ultraGridColumn38.Header.Caption = "Valor cobro ($)";
			ultraGridColumn38.Header.VisiblePosition = 7;
			ultraGridColumn38.RowLayoutColumnInfo.OriginX = 8;
			ultraGridColumn38.RowLayoutColumnInfo.OriginY = 0;
			ultraGridColumn38.RowLayoutColumnInfo.PreferredCellSize = new System.Drawing.Size(115, 0);
			ultraGridColumn38.RowLayoutColumnInfo.SpanX = 2;
			ultraGridColumn38.RowLayoutColumnInfo.SpanY = 2;
			ultraGridColumn38.Width = 107;
			ultraGridColumn39.Header.VisiblePosition = 39;
			ultraGridColumn39.Hidden = true;
			ultraGridColumn40.Header.VisiblePosition = 40;
			ultraGridColumn40.Hidden = true;
			ultraGridColumn41.Header.VisiblePosition = 41;
			ultraGridColumn41.Hidden = true;
			ultraGridColumn42.Header.VisiblePosition = 42;
			ultraGridColumn42.Hidden = true;
			ultraGridColumn43.Header.VisiblePosition = 43;
			ultraGridColumn43.Hidden = true;
			ultraGridColumn44.Header.VisiblePosition = 44;
			ultraGridColumn44.Hidden = true;
			ultraGridColumn45.Header.VisiblePosition = 45;
			ultraGridColumn45.Hidden = true;
			ultraGridColumn46.Header.VisiblePosition = 46;
			ultraGridColumn46.Hidden = true;
			ultraGridColumn47.Header.VisiblePosition = 47;
			ultraGridColumn47.Hidden = true;
			ultraGridColumn48.Header.VisiblePosition = 48;
			ultraGridColumn48.Hidden = true;
			ultraGridColumn49.Header.VisiblePosition = 49;
			ultraGridColumn49.Hidden = true;
			ultraGridColumn50.Header.VisiblePosition = 50;
			ultraGridColumn50.Hidden = true;
			ultraGridColumn51.ButtonDisplayStyle = Infragistics.Win.UltraWinGrid.ButtonDisplayStyle.Always;
			appearance18.Image = ((object)(resources.GetObject("appearance18.Image")));
			appearance18.ImageHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn51.CellButtonAppearance = appearance18;
			ultraGridColumn51.Header.Caption = "...";
			ultraGridColumn51.Header.VisiblePosition = 0;
			ultraGridColumn51.Hidden = true;
			ultraGridColumn51.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.Button;
			ultraGridColumn51.Width = 30;
			ultraGridColumn52.ButtonDisplayStyle = Infragistics.Win.UltraWinGrid.ButtonDisplayStyle.Always;
			appearance19.Image = ((object)(resources.GetObject("appearance19.Image")));
			appearance19.ImageHAlign = Infragistics.Win.HAlign.Center;
			appearance19.ImageVAlign = Infragistics.Win.VAlign.Middle;
			ultraGridColumn52.CellButtonAppearance = appearance19;
			ultraGridColumn52.Header.Caption = "Proforma";
			ultraGridColumn52.Header.VisiblePosition = 1;
			ultraGridColumn52.Hidden = true;
			ultraGridColumn52.RowLayoutColumnInfo.OriginX = 2;
			ultraGridColumn52.RowLayoutColumnInfo.OriginY = 0;
			ultraGridColumn52.RowLayoutColumnInfo.SpanX = 2;
			ultraGridColumn52.RowLayoutColumnInfo.SpanY = 2;
			ultraGridColumn52.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.Button;
			ultraGridColumn52.Width = 60;
			ultraGridColumn53.AutoEdit = false;
			ultraGridColumn53.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
			ultraGridColumn53.Header.Caption = "N°";
			ultraGridColumn53.Header.VisiblePosition = 52;
			ultraGridColumn53.RowLayoutColumnInfo.OriginX = 0;
			ultraGridColumn53.RowLayoutColumnInfo.OriginY = 0;
			ultraGridColumn53.RowLayoutColumnInfo.PreferredCellSize = new System.Drawing.Size(30, 0);
			ultraGridColumn53.RowLayoutColumnInfo.SpanX = 2;
			ultraGridColumn53.RowLayoutColumnInfo.SpanY = 2;
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
																										 ultraGridColumn53});
			appearance20.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings1.Appearance = appearance20;
			summarySettings1.DisplayFormat = "{0: #,##0.00}";
			summarySettings1.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			ultraGridBand1.Summaries.AddRange(new Infragistics.Win.UltraWinGrid.SummarySettings[] {
																																															summarySettings1});
			ultraGridBand1.SummaryFooterCaption = "Sumatoria:";
			ultraGridBand1.UseRowLayout = true;
			this.uGridDocumentos.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.uGridDocumentos.DisplayLayout.GroupByBox.Hidden = true;
			appearance21.ForeColor = System.Drawing.Color.Black;
			appearance21.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridDocumentos.DisplayLayout.Override.ActiveRowAppearance = appearance21;
			this.uGridDocumentos.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.Yes;
			this.uGridDocumentos.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.True;
			this.uGridDocumentos.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance22.BackColor = System.Drawing.Color.Transparent;
			this.uGridDocumentos.DisplayLayout.Override.CardAreaAppearance = appearance22;
			appearance23.ForeColor = System.Drawing.Color.Black;
			appearance23.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridDocumentos.DisplayLayout.Override.CellAppearance = appearance23;
			appearance24.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance24.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance24.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance24.FontData.BoldAsString = "True";
			appearance24.FontData.Name = "Arial";
			appearance24.FontData.SizeInPoints = 8F;
			appearance24.ForeColor = System.Drawing.Color.White;
			appearance24.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridDocumentos.DisplayLayout.Override.HeaderAppearance = appearance24;
			this.uGridDocumentos.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti;
			appearance25.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance25.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance25.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridDocumentos.DisplayLayout.Override.RowAlternateAppearance = appearance25;
			appearance26.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance26.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance26.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridDocumentos.DisplayLayout.Override.RowSelectorAppearance = appearance26;
			appearance27.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance27.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance27.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridDocumentos.DisplayLayout.Override.SelectedRowAppearance = appearance27;
			this.uGridDocumentos.DisplayLayout.ViewStyleBand = Infragistics.Win.UltraWinGrid.ViewStyleBand.OutlookGroupBy;
			this.uGridDocumentos.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridDocumentos.Location = new System.Drawing.Point(16, 112);
			this.uGridDocumentos.Name = "uGridDocumentos";
			this.uGridDocumentos.RightToLeft = System.Windows.Forms.RightToLeft.No;
			this.uGridDocumentos.RowUpdateCancelAction = Infragistics.Win.UltraWinGrid.RowUpdateCancelAction.RetainDataAndActivation;
			this.uGridDocumentos.Size = new System.Drawing.Size(632, 208);
			this.uGridDocumentos.TabIndex = 981;
			this.uGridDocumentos.InitializeRow += new Infragistics.Win.UltraWinGrid.InitializeRowEventHandler(this.uGridDocumentos_InitializeRow);
			this.uGridDocumentos.AfterRowInsert += new Infragistics.Win.UltraWinGrid.RowEventHandler(this.uGridDocumentos_AfterRowInsert);
			this.uGridDocumentos.AfterCellUpdate += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.uGridDocumentos_AfterCellUpdate);
			this.uGridDocumentos.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.uGridDocumentos_InitializeLayout);
			// 
			// groupBox2
			// 
			this.groupBox2.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.groupBox2.Controls.Add(this.groupBox4);
			this.groupBox2.Controls.Add(this.lblBodega);
			this.groupBox2.Controls.Add(this.label53);
			this.groupBox2.Controls.Add(this.dtDesde);
			this.groupBox2.Controls.Add(this.dtHasta);
			this.groupBox2.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.groupBox2.Location = new System.Drawing.Point(16, 8);
			this.groupBox2.Name = "groupBox2";
			this.groupBox2.Size = new System.Drawing.Size(568, 240);
			this.groupBox2.TabIndex = 989;
			this.groupBox2.TabStop = false;
			this.groupBox2.Text = "Datos principales";
			// 
			// groupBox4
			// 
			this.groupBox4.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.groupBox4.Controls.Add(this.label6);
			this.groupBox4.Controls.Add(this.label3);
			this.groupBox4.Controls.Add(this.label4);
			this.groupBox4.Controls.Add(this.dtFecha);
			this.groupBox4.Controls.Add(this.txtCheque);
			this.groupBox4.Controls.Add(this.txtNotas);
			this.groupBox4.Controls.Add(this.cmbCuentas);
			this.groupBox4.Controls.Add(this.label5);
			this.groupBox4.Location = new System.Drawing.Point(16, 56);
			this.groupBox4.Name = "groupBox4";
			this.groupBox4.Size = new System.Drawing.Size(536, 168);
			this.groupBox4.TabIndex = 1034;
			this.groupBox4.TabStop = false;
			this.groupBox4.Text = "Registro depósito";
			this.groupBox4.Enter += new System.EventHandler(this.groupBox4_Enter);
			// 
			// label6
			// 
			this.label6.AutoSize = true;
			this.label6.Location = new System.Drawing.Point(8, 88);
			this.label6.Name = "label6";
			this.label6.Size = new System.Drawing.Size(39, 16);
			this.label6.TabIndex = 1009;
			this.label6.Text = "* Nota:";
			this.label6.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label3
			// 
			this.label3.AutoSize = true;
			this.label3.Location = new System.Drawing.Point(8, 24);
			this.label3.Name = "label3";
			this.label3.Size = new System.Drawing.Size(43, 16);
			this.label3.TabIndex = 1005;
			this.label3.Text = "* Fecha";
			this.label3.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label4
			// 
			this.label4.AutoSize = true;
			this.label4.Location = new System.Drawing.Point(256, 24);
			this.label4.Name = "label4";
			this.label4.Size = new System.Drawing.Size(55, 16);
			this.label4.TabIndex = 1006;
			this.label4.Text = "* Número:";
			this.label4.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// dtFecha
			// 
			this.dtFecha.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton1.Caption = "Today";
			this.dtFecha.DateButtons.Add(dateButton1);
			this.dtFecha.Enabled = false;
			this.dtFecha.Format = "dd/MM/yyyy";
			this.dtFecha.Location = new System.Drawing.Point(72, 24);
			this.dtFecha.Name = "dtFecha";
			this.dtFecha.NonAutoSizeHeight = 23;
			this.dtFecha.ReadOnly = true;
			this.dtFecha.Size = new System.Drawing.Size(176, 21);
			this.dtFecha.SpinButtonsVisible = true;
			this.dtFecha.TabIndex = 1002;
			this.dtFecha.Value = ((object)(resources.GetObject("dtFecha.Value")));
			// 
			// txtCheque
			// 
			appearance28.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtCheque.Appearance = appearance28;
			this.txtCheque.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtCheque.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtCheque.Location = new System.Drawing.Point(328, 24);
			this.txtCheque.Name = "txtCheque";
			this.txtCheque.Size = new System.Drawing.Size(184, 21);
			this.txtCheque.TabIndex = 1003;
			// 
			// txtNotas
			// 
			appearance29.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtNotas.Appearance = appearance29;
			this.txtNotas.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtNotas.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNotas.Location = new System.Drawing.Point(72, 88);
			this.txtNotas.Multiline = true;
			this.txtNotas.Name = "txtNotas";
			this.txtNotas.Size = new System.Drawing.Size(440, 64);
			this.txtNotas.TabIndex = 1004;
			// 
			// cmbCuentas
			// 
			appearance30.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbCuentas.Appearance = appearance30;
			this.cmbCuentas.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbCuentas.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbCuentas.DataSource = this.udsCuenta;
			ultraGridColumn54.Header.VisiblePosition = 0;
			ultraGridColumn54.Hidden = true;
			ultraGridColumn55.Header.Caption = "Descripción";
			ultraGridColumn55.Header.VisiblePosition = 1;
			ultraGridColumn55.Width = 438;
			ultraGridBand2.Columns.AddRange(new object[] {
																										 ultraGridColumn54,
																										 ultraGridColumn55});
			this.cmbCuentas.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			this.cmbCuentas.DisplayMember = "Descripcion";
			this.cmbCuentas.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbCuentas.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbCuentas.Location = new System.Drawing.Point(72, 56);
			this.cmbCuentas.Name = "cmbCuentas";
			this.cmbCuentas.Size = new System.Drawing.Size(440, 21);
			this.cmbCuentas.TabIndex = 1008;
			this.cmbCuentas.ValueMember = "idCuenta";
			// 
			// udsCuenta
			// 
			ultraDataColumn51.DataType = typeof(int);
			this.udsCuenta.Band.Columns.AddRange(new object[] {
																													ultraDataColumn51,
																													ultraDataColumn52});
			// 
			// label5
			// 
			this.label5.AutoSize = true;
			this.label5.Location = new System.Drawing.Point(8, 56);
			this.label5.Name = "label5";
			this.label5.Size = new System.Drawing.Size(51, 16);
			this.label5.TabIndex = 1007;
			this.label5.Text = "* Cuenta:";
			this.label5.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// lblBodega
			// 
			this.lblBodega.AutoSize = true;
			this.lblBodega.BackColor = System.Drawing.Color.Transparent;
			this.lblBodega.Location = new System.Drawing.Point(24, 24);
			this.lblBodega.Name = "lblBodega";
			this.lblBodega.Size = new System.Drawing.Size(40, 16);
			this.lblBodega.TabIndex = 993;
			this.lblBodega.Text = "Desde:";
			this.lblBodega.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label53
			// 
			this.label53.AutoSize = true;
			this.label53.BackColor = System.Drawing.Color.Transparent;
			this.label53.Location = new System.Drawing.Point(272, 24);
			this.label53.Name = "label53";
			this.label53.Size = new System.Drawing.Size(37, 16);
			this.label53.TabIndex = 991;
			this.label53.Text = "Hasta:";
			this.label53.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// dtDesde
			// 
			appearance31.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtDesde.Appearance = appearance31;
			this.dtDesde.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton2.Caption = "Today";
			this.dtDesde.DateButtons.Add(dateButton2);
			this.dtDesde.Format = "dd/MM/yyyy";
			this.dtDesde.Location = new System.Drawing.Point(88, 24);
			this.dtDesde.Name = "dtDesde";
			this.dtDesde.NonAutoSizeHeight = 23;
			this.dtDesde.ReadOnly = true;
			this.dtDesde.Size = new System.Drawing.Size(176, 21);
			this.dtDesde.SpinButtonsVisible = true;
			this.dtDesde.TabIndex = 989;
			this.dtDesde.Value = ((object)(resources.GetObject("dtDesde.Value")));
			// 
			// dtHasta
			// 
			appearance32.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtHasta.Appearance = appearance32;
			this.dtHasta.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton3.Caption = "Today";
			this.dtHasta.DateButtons.Add(dateButton3);
			this.dtHasta.Format = "dd/MM/yyyy";
			this.dtHasta.Location = new System.Drawing.Point(344, 24);
			this.dtHasta.Name = "dtHasta";
			this.dtHasta.NonAutoSizeHeight = 23;
			this.dtHasta.ReadOnly = true;
			this.dtHasta.Size = new System.Drawing.Size(184, 21);
			this.dtHasta.SpinButtonsVisible = true;
			this.dtHasta.TabIndex = 990;
			this.dtHasta.Value = ((object)(resources.GetObject("dtHasta.Value")));
			// 
			// groupBox3
			// 
			this.groupBox3.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.groupBox3.Location = new System.Drawing.Point(-792, 256);
			this.groupBox3.Name = "groupBox3";
			this.groupBox3.Size = new System.Drawing.Size(2200, 8);
			this.groupBox3.TabIndex = 1032;
			this.groupBox3.TabStop = false;
			// 
			// frmReporteCompraProtegidaDocumentoEnviado
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(704, 670);
			this.Controls.Add(this.groupBox3);
			this.Controls.Add(this.groupBox2);
			this.Controls.Add(this.groupBox1);
			this.Controls.Add(this.btnSalir);
			this.Controls.Add(this.btnGuardar);
			this.Name = "frmReporteCompraProtegidaDocumentoEnviado";
			this.Text = "Documentos enviados";
			this.Load += new System.EventHandler(this.frmReporteCompraProtegidaDocumentoEnviado_Load);
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			this.groupBox1.ResumeLayout(false);
			this.groupBox5.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.txtNumero)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtTotal)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridDocumentos)).EndInit();
			this.groupBox2.ResumeLayout(false);
			this.groupBox4.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.dtFecha)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCheque)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNotas)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCuentas)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.udsCuenta)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtDesde)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtHasta)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void frmReporteCompraProtegidaDocumentoEnviado_Load(object sender, System.EventArgs e)
		{
			this.dtDesde.CalendarInfo.MinDate = DateTime.Parse(frmReporteCompraProtegida.FechaDesde);
			this.dtHasta.CalendarInfo.MinDate = DateTime.Parse(frmReporteCompraProtegida.FechaDesde);

			this.dtDesde.CalendarInfo.MaxDate = DateTime.Today;			
			this.dtHasta.CalendarInfo.MaxDate = DateTime.Today;

			this.dtDesde.Value = DateTime.Parse(frmReporteCompraProtegida.FechaDesde);//"01/" + DateTime.Now.Month.ToString() + "/" + DateTime.Now.Year.ToString();			
			this.dtHasta.Value = DateTime.Today;

			#region Valida Fecha
			if (this.dtDesde.Value == System.DBNull.Value)
			{
				MessageBox.Show("Seleccione la Fecha Inicial", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.dtDesde.Focus();
				return;
			}
     			
			if (this.dtHasta.Value == System.DBNull.Value)
			{
				MessageBox.Show("Seleccione la Fecha Final", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.dtHasta.Focus();
				return;
			}

			if (Convert.ToDateTime(this.dtDesde.Value) > Convert.ToDateTime(this.dtHasta.Value))
			{
				MessageBox.Show("La Fecha Inicial no puede ser mayor a la Fecha Final", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.dtDesde.Focus();
				return;
			}
			#endregion Valida Fecha
			
			Consulta();	
			//Inicialización formulario
			FuncionesProcedimientos.RetornaFechaServidor(this.dtFecha, cdsSeteoF, false);
			this.dtFecha.CalendarInfo.MinDate = Convert.ToDateTime(this.dtFecha.Value).AddDays(-1);			
			this.cmbCuentas.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Select c.idCuenta, Descripcion From Cuenta c Inner Join CuentaCorriente cc On cc.idCuenta = c.idCuenta Union Select idCuenta, Descripcion From Cuenta Where idCuenta = 171 Order By Descripcion");

		}

		public void LimpiarCampos()
		{
			FuncionesProcedimientos.RetornaFechaServidor(this.dtFecha, cdsSeteoF, false);
			this.dtFecha.CalendarInfo.MinDate = Convert.ToDateTime(this.dtFecha.Value).AddDays(-1);
			this.txtCheque.Text = "";
			this.txtNotas.Text = "";
			this.cmbCuentas.Value = System.DBNull.Value;
		}

		private void Consulta()
		{
			string sSQLPa = string.Format("Exec AdministracionCompraProtegida '{0}', '{1}', {2}, '{3}', {4}, {5}", 
				Convert.ToDateTime(this.dtDesde.Value).ToString("yyyyMMdd"), 
				Convert.ToDateTime(this.dtHasta.Value).ToString("yyyyMMdd"), 
				(int)frmReporteCompraProtegida.iEstadoSeguroDocumentoEnviado, 
				this.txtNumero.Text,
				1,
				0); //this.txtBusqueda.Text.ToString()
			FuncionesProcedimientos.dsGeneralModulo(sSQLPa, this.uGridDocumentos);
			//Label de registros
			this.lblContador.Text = this.uGridDocumentos.Rows.Count + " REGISTROS ENCONTRADOS";
			//COntador de filas
			this.uGridDocumentos = Funcion.ContadorFilas(uGridDocumentos, "Row");
			TotalSeleccionado();
		}



		private void btnGuardar_Click(object sender, System.EventArgs e)
		{
			//Eliminar la busqueda.
			this.txtNumero.Text = "";
			btnConsultar_Click(this.btnConsultar, System.EventArgs.Empty);

			if (this.txtCheque.Text.ToString().Length == 0)
			{
				MessageBox.Show("Ingrese el número de depósito.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);					
				this.txtCheque.Focus();
				return;
			}
						
			if (!Validacion.vbComboVacio(this.cmbCuentas, "Seleccione una cuenta bancaria.")) return;

			if (Funcion.iEscalarSQL(cdsSeteoF, string.Format("Select COUNT(*) From AsientoDetalle ad Inner Join Asiento a On a.idAsiento = ad.idAsiento Where a.Cheque = '{0}' And ad.idCuenta In ({1}) And a.idTipoAsiento = 2", this.txtCheque.Text.ToString(), (int)this.cmbCuentas.Value)) > 0)
			{
				MessageBox.Show(string.Format("Numero de Cheque ya esta registrado en la cuenta '{0}'.", this.cmbCuentas.Text.ToString()), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.txtCheque.Focus();				
				return;
			}

			if (!Validacion.vbTexto(this.txtNotas, 10, 100, "Notas")) return;

			#region Validacion
			int ContadorSeleccion = 0;
			bool SelDocumentoEnviado = false;

			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridDocumentos.Rows.All)	
			{
				if(ContadorSeleccion == 0)
				{
					SelDocumentoEnviado = (bool)dr.Cells["SelDocumentoEnviado"].Value;
					if(SelDocumentoEnviado)
					{
						ContadorSeleccion = 1;
					}
				}
			}
			if(ContadorSeleccion == 0)
			{
				MessageBox.Show(string.Format("Se debe seleccionar al menos un valor de cobro."), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return;
			}
			#endregion Validacion

			#region Registro deposito
			int Cuenta = (int)this.cmbCuentas.Value;
			string Cheque = this.txtCheque.Text;
			string Notas = this.txtNotas.Text;
			string Deducible = "";
			int idDetCompra = 0;
			int idCompraP = 0;
			decimal ValorDocumentoEnviado = 0;

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

					foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridDocumentos.Rows.All)	
					{
						idCompraP = (int)dr.Cells["idCompra"].Value;
						Deducible =  dr.Cells["Deducible"].Value.ToString();
						idDetCompra = (int)dr.Cells["idDetCompra"].Value;
						ValorDocumentoEnviado = Convert.ToDecimal(dr.Cells["ValorDocumentoEnviado"].Value);
						SelDocumentoEnviado = (bool)dr.Cells["SelDocumentoEnviado"].Value;

						if(SelDocumentoEnviado)
						{
							if(idCompraP != 0) //839458
							{

								#region Cobro seguro
								string sNumero = Funcion.sEscalarSQL(cdsSeteoF, string.Format("Exec NumeracionLocales 24, 1, 0"));
								string	sSQLg = string.Format("Exec GuardaCobroCompraProtegida {0}, 25, 7, {1}, '{2}', '{3}', '{4}', {5}", 
									idCompraP, Cuenta, Cheque, 
									sNumero, Notas, idDetCompra);
								oCmdActualiza.CommandText = sSQLg;
								int idAnticipo = (int)oCmdActualiza.ExecuteScalar();

								#endregion Cobro seguro
								#region Pregunta deducible
								string squery2 = "";
								//SIEMPRE VA A TENER QUE PASAR POR EL ESTADO SEGURO LIQUIDADO, Y EL ESTADO 8 QUE ES DECUCIBLE LIQUIDADO YA NO SE UTILIZARÍA.
								squery2 = string.Format("Update DetCompra Set EstadoSeguro = 6, idAnticipoP = {1} Where idDetCompra = {0}", 
								idDetCompra, idAnticipo);
								oCmdActualiza.CommandText = squery2;
								oCmdActualiza.ExecuteNonQuery();
								#endregion Pregunta deducible
							}
						}
					}
					oTransaction.Commit();
					oConexion.Close();
					#region Mensaje exito
					MessageBox.Show("Depositos registrados correctamente", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
					#endregion Mensaje exito
					Consulta();
					LimpiarCampos();
					
				}
				catch (Exception ex)
				{
					oTransaction.Rollback();
					MessageBox.Show(Funcion.smensajeerror, "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
					Funcion.LogSistema(frmRequerimiento.sconexionPoint, this.Name + ": "+ this.Text, string.Format("Commit Exception Type : {0} {1}", ex.GetType(), ex.Message), "", Funcion.slogtipoerror);
				}
				finally
				{
					oConexion.Close();
				}

			}
							#endregion Registro depósito
		}

		private void label53_Click(object sender, System.EventArgs e)
		{
		
		}

		private void uGridDocumentos_AfterCellUpdate(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			if (e.Cell.Column.ToString() == "SelDocumentoEnviado")
			{
				bool SelDocumentoEnviado = false;
					try{ SelDocumentoEnviado = (bool)e.Cell.Row.Cells["SelDocumentoEnviado"].Value; }
					catch(Exception ex){ string sms = ex.Message;};
				decimal ValorDocumentoEnviado = 0;
					
					try{ ValorDocumentoEnviado = Convert.ToDecimal(e.Cell.Row.Cells["ValorDocumentoEnviado"].Value);}
					catch(Exception ex){ string sms = ex.Message;};

				if(SelDocumentoEnviado)
				{
					if(ValorDocumentoEnviado == 0)
					{
						MessageBox.Show(string.Format("Para poder seleccionar el registro se debe ubicar un valor de cobro."), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
						e.Cell.Row.Cells["SelDocumentoEnviado"].Value = false;
					}
				}
			}

			if (e.Cell.Column.ToString() == "ValorDocumentoEnviado")
			{
				bool SelDocumentoEnviado = false;
				int getidCompra = 0;
				try{ getidCompra = (int)e.Cell.Row.Cells["idCompra"].Value; }
				catch(Exception ex){ string sms = ex.Message;};
				try{ SelDocumentoEnviado = (bool)e.Cell.Row.Cells["SelDocumentoEnviado"].Value; }
				catch(Exception ex){ string sms = ex.Message;};
				decimal ValorDocumentoEnviado = 0;
				try{ ValorDocumentoEnviado = Convert.ToDecimal(e.Cell.Row.Cells["ValorDocumentoEnviado"].Value);}
				catch(Exception ex){ string sms = ex.Message;};

				if(ValorDocumentoEnviado == 0)
				{
					e.Cell.Row.Cells["SelDocumentoEnviado"].Value = false;
					SelDocumentoEnviado = false;
				}
				else
				{
					e.Cell.Row.Cells["SelDocumentoEnviado"].Value = true;
					SelDocumentoEnviado = true;
				}

				int getIdDetCompra = (int)e.Cell.Row.Cells["idDetCompra"].Value;
				//Guardar en la base de datos
				ReporteCompraProtegidaValoresAct(getidCompra, ValorDocumentoEnviado, SelDocumentoEnviado, getIdDetCompra);
			}
			TotalSeleccionado();

		}

		public void ReporteCompraProtegidaValoresAct(int getidCompra, decimal getValorDocumentoEnviado, bool getSelDocumentoEnviado, int getIdDetCompra)
		{
			#region Guardar datos en detalle compra
			//Actualiza valores en la compra
			string squery = string.Format("EXEC ReporteCompraProtegidaValoresAct {0}, {1}, {2}, {3}", 
				getidCompra, getValorDocumentoEnviado, getSelDocumentoEnviado, getIdDetCompra);
			Funcion.EjecutaSQL(cdsSeteoF, squery);
			#endregion Guardar datos en detalle compra
		}

		private void btnSalir_Click(object sender, System.EventArgs e)
		{
			this.Close();
		}

		private void uGridDocumentos_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
			FuncionesProcedimientos.FormatoGrid(e, "ValorDocumentoEnviado", 2);
		}

		private void uGridDocumentos_InitializeRow(object sender, Infragistics.Win.UltraWinGrid.InitializeRowEventArgs e)
		{
		}

		private void button1_Click(object sender, System.EventArgs e)
		{
		}

		private void uGridDocumentos_AfterRowInsert(object sender, Infragistics.Win.UltraWinGrid.RowEventArgs e)
		{
		
		}

		public void TotalSeleccionado()
		{
//			decimal TotalValorDocumentoEnviado = 0;
//			decimal ValorDocumentoEnviado = 0;
//			bool SelDocumentoEnviado = false;

//			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridDocumentos.Rows.All)	
//			{
//				SelDocumentoEnviado = (bool)dr.Cells["SelDocumentoEnviado"].Value;
//				if(SelDocumentoEnviado)
//				{
//					ValorDocumentoEnviado = Convert.ToDecimal(dr.Cells["ValorDocumentoEnviado"].Value);
//					TotalValorDocumentoEnviado = TotalValorDocumentoEnviado + ValorDocumentoEnviado;
//				}
//			}
			string squery = String.Format("EXEC ReporteCompraProtegidaDocumentoEnviadoTotal");
			this.txtTotal.Value = Funcion.decEscalarSQL(cdsSeteoF, squery);//  TotalValorDocumentoEnviado;
		}

		private void groupBox1_Enter(object sender, System.EventArgs e)
		{
		
		}

		private void groupBox4_Enter(object sender, System.EventArgs e)
		{
		
		}

		private void txtNumero_ValueChanged(object sender, System.EventArgs e)
		{
			//Consulta();
		}

		private void btnConsultar_Click(object sender, System.EventArgs e)
		{
			Consulta();
		}

		private void groupBox5_Enter(object sender, System.EventArgs e)
		{
		
		}


	}
}
