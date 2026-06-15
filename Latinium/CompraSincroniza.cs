using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de CompraSincroniza.
	/// </summary>
	public class CompraSincroniza : System.Windows.Forms.Form
	{
		private Infragistics.Win.UltraWinGrid.UltraGrid ultraGrid1;
		private C1.Data.C1DataSet cdsCompra;
		private System.Windows.Forms.Label label1;
		private System.Windows.Forms.Label label2;
		private System.Windows.Forms.Label label3;
		private Infragistics.Win.Misc.UltraButton btFiltro;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo cmbFechaIni;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo cmbFechaFin;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbTipoFactura;
		private C1.Data.C1DataSet cdCompraNumero;
		private C1.Data.C1DataView cdvCompra;
		private Infragistics.Win.Misc.UltraButton ultraButton1;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkEditado;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkNuevo;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public CompraSincroniza()
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
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Compra", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("AutFactura");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("AutImprenta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("AutPresupuesto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("AutPrincipal");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bodega");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Borrar");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Comprobante");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Contado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ContadoCredito");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Credito");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn11 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Departamento");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn12 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("DerechoDevolucion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn13 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Desc1");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn14 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Desc2");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn15 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Desc3");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn16 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Desc4");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn17 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("DescPorcIva");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn18 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("DescPorcIva0");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn19 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Descuento");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn20 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("DiasCredito");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn21 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("DiasInicio");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn22 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Editar");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn23 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Fecha");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn24 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaIngreso");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn25 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaEntrega");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn26 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaRevision");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn27 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaVencimiento");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn28 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Flete");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn29 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FletePorcentaje");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn30 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Ice");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn31 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idAsiento");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn32 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idBanco");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn33 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCliente");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn34 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCompra");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn35 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idComprobante");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn36 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCredTributario");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn37 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idEmpleado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn38 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idFormaPago");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn39 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idProyecto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn40 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idSubProyecto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn41 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idTipoFactura");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn42 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idTipoMoneda");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn43 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idVendedor");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn44 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Impreso");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn45 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Interes");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn46 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Iva");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn47 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Mensaje1");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn48 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Mensaje2");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn49 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Mensaje3");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn50 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn51 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Notas");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn52 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nuevo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn53 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("NumCuenta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn54 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("NumCuotas");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn55 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("NumCheque");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn56 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Otro");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn57 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("PagosCada");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn58 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("PorcentajeIva");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn59 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Prioridad");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn60 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Refrendo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn61 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Retencion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn62 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Saldo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn63 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("SerieFactura");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn64 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("SubtotalExcento");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn65 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("SubtotalIva");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn66 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Total");
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn67 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("TotalCuotas");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn68 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Usuario");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn69 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ValorCuotas");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn70 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ValorDolar");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn71 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ValorInteres");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn72 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Retenido");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn73 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Entregado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn74 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Revisado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn75 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Descuento0");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn76 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idOrdenProd");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn77 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CentroCosto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn78 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idRecibe");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn79 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("BienServicio");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn80 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ConvenioInternac");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn81 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Vuelto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn82 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idSucursal");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn83 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaEditado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn84 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Numero");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn85 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idAsiRet");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn86 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaCaducidad");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn87 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPorcIce");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn88 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idAsientoPer");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn89 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("IvaServicio");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn90 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CuentaRetDet");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn91 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Comprob1");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn92 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bulto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn93 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Peso");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn94 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Pedido");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn95 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Depart");
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton1 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton2 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("CompraNumero", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn96 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Abrev");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn97 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Automatico");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn98 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn99 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Grupo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn100 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idTipoFactura");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn101 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Name");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn102 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn103 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Numero");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn104 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Signo");
			this.ultraGrid1 = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.cdsCompra = new C1.Data.C1DataSet();
			this.cmbFechaIni = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.label1 = new System.Windows.Forms.Label();
			this.label2 = new System.Windows.Forms.Label();
			this.cmbFechaFin = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.cmbTipoFactura = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.cdCompraNumero = new C1.Data.C1DataSet();
			this.label3 = new System.Windows.Forms.Label();
			this.btFiltro = new Infragistics.Win.Misc.UltraButton();
			this.cdvCompra = new C1.Data.C1DataView();
			this.ultraButton1 = new Infragistics.Win.Misc.UltraButton();
			this.chkEditado = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.chkNuevo = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsCompra)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbFechaIni)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbFechaFin)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbTipoFactura)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdCompraNumero)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvCompra)).BeginInit();
			this.SuspendLayout();
			// 
			// ultraGrid1
			// 
			this.ultraGrid1.Cursor = System.Windows.Forms.Cursors.Default;
			this.ultraGrid1.DataMember = "Compra";
			this.ultraGrid1.DataSource = this.cdsCompra;
			appearance1.BackColor = System.Drawing.Color.White;
			this.ultraGrid1.DisplayLayout.Appearance = appearance1;
			ultraGridColumn1.Header.VisiblePosition = 4;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn2.Header.VisiblePosition = 5;
			ultraGridColumn2.Hidden = true;
			ultraGridColumn3.Header.VisiblePosition = 6;
			ultraGridColumn3.Hidden = true;
			ultraGridColumn4.Header.VisiblePosition = 7;
			ultraGridColumn4.Hidden = true;
			ultraGridColumn5.Header.VisiblePosition = 8;
			ultraGridColumn5.Hidden = true;
			ultraGridColumn6.Header.VisiblePosition = 9;
			ultraGridColumn6.Hidden = true;
			ultraGridColumn7.Header.VisiblePosition = 10;
			ultraGridColumn7.Hidden = true;
			ultraGridColumn8.Header.VisiblePosition = 11;
			ultraGridColumn8.Hidden = true;
			ultraGridColumn9.Header.VisiblePosition = 12;
			ultraGridColumn9.Hidden = true;
			ultraGridColumn10.Header.VisiblePosition = 13;
			ultraGridColumn10.Hidden = true;
			ultraGridColumn11.Header.VisiblePosition = 14;
			ultraGridColumn11.Hidden = true;
			ultraGridColumn12.Header.VisiblePosition = 15;
			ultraGridColumn12.Hidden = true;
			ultraGridColumn13.Header.VisiblePosition = 16;
			ultraGridColumn13.Hidden = true;
			ultraGridColumn14.Header.VisiblePosition = 17;
			ultraGridColumn14.Hidden = true;
			ultraGridColumn15.Header.VisiblePosition = 18;
			ultraGridColumn15.Hidden = true;
			ultraGridColumn16.Header.VisiblePosition = 19;
			ultraGridColumn16.Hidden = true;
			ultraGridColumn17.Header.VisiblePosition = 20;
			ultraGridColumn17.Hidden = true;
			ultraGridColumn18.Header.VisiblePosition = 21;
			ultraGridColumn18.Hidden = true;
			ultraGridColumn19.Header.VisiblePosition = 22;
			ultraGridColumn19.Hidden = true;
			ultraGridColumn20.Header.VisiblePosition = 23;
			ultraGridColumn20.Hidden = true;
			ultraGridColumn21.Header.VisiblePosition = 24;
			ultraGridColumn21.Hidden = true;
			ultraGridColumn22.Header.VisiblePosition = 25;
			ultraGridColumn22.Hidden = true;
			ultraGridColumn23.Format = "dd/MMM/yyyy";
			ultraGridColumn23.Header.VisiblePosition = 1;
			ultraGridColumn23.Width = 79;
			ultraGridColumn24.Header.VisiblePosition = 26;
			ultraGridColumn24.Hidden = true;
			ultraGridColumn25.Header.VisiblePosition = 27;
			ultraGridColumn25.Hidden = true;
			ultraGridColumn26.Header.VisiblePosition = 28;
			ultraGridColumn26.Hidden = true;
			ultraGridColumn27.Header.VisiblePosition = 29;
			ultraGridColumn27.Hidden = true;
			ultraGridColumn28.Header.VisiblePosition = 30;
			ultraGridColumn28.Hidden = true;
			ultraGridColumn29.Header.VisiblePosition = 31;
			ultraGridColumn29.Hidden = true;
			ultraGridColumn30.Header.VisiblePosition = 32;
			ultraGridColumn30.Hidden = true;
			ultraGridColumn31.Header.VisiblePosition = 33;
			ultraGridColumn31.Hidden = true;
			ultraGridColumn32.Header.VisiblePosition = 34;
			ultraGridColumn32.Hidden = true;
			ultraGridColumn33.Header.VisiblePosition = 35;
			ultraGridColumn33.Hidden = true;
			ultraGridColumn34.Header.VisiblePosition = 36;
			ultraGridColumn34.Hidden = true;
			ultraGridColumn35.Header.VisiblePosition = 37;
			ultraGridColumn35.Hidden = true;
			ultraGridColumn36.Header.VisiblePosition = 38;
			ultraGridColumn36.Hidden = true;
			ultraGridColumn37.Header.VisiblePosition = 39;
			ultraGridColumn37.Hidden = true;
			ultraGridColumn38.Header.VisiblePosition = 40;
			ultraGridColumn38.Hidden = true;
			ultraGridColumn39.Header.VisiblePosition = 41;
			ultraGridColumn39.Hidden = true;
			ultraGridColumn40.Header.VisiblePosition = 42;
			ultraGridColumn40.Hidden = true;
			ultraGridColumn41.Header.VisiblePosition = 43;
			ultraGridColumn41.Hidden = true;
			ultraGridColumn42.Header.VisiblePosition = 44;
			ultraGridColumn42.Hidden = true;
			ultraGridColumn43.Header.VisiblePosition = 45;
			ultraGridColumn43.Hidden = true;
			ultraGridColumn44.Header.VisiblePosition = 46;
			ultraGridColumn44.Hidden = true;
			ultraGridColumn45.Header.VisiblePosition = 47;
			ultraGridColumn45.Hidden = true;
			ultraGridColumn46.Header.VisiblePosition = 48;
			ultraGridColumn46.Hidden = true;
			ultraGridColumn47.Header.VisiblePosition = 49;
			ultraGridColumn47.Hidden = true;
			ultraGridColumn48.Header.VisiblePosition = 50;
			ultraGridColumn48.Hidden = true;
			ultraGridColumn49.Header.VisiblePosition = 51;
			ultraGridColumn49.Hidden = true;
			ultraGridColumn50.Header.VisiblePosition = 52;
			ultraGridColumn50.Hidden = true;
			ultraGridColumn51.Header.VisiblePosition = 3;
			ultraGridColumn51.Hidden = true;
			ultraGridColumn52.Header.VisiblePosition = 53;
			ultraGridColumn52.Hidden = true;
			ultraGridColumn53.Header.VisiblePosition = 54;
			ultraGridColumn53.Hidden = true;
			ultraGridColumn54.Header.VisiblePosition = 55;
			ultraGridColumn54.Hidden = true;
			ultraGridColumn55.Header.VisiblePosition = 56;
			ultraGridColumn55.Hidden = true;
			ultraGridColumn56.Header.VisiblePosition = 57;
			ultraGridColumn56.Hidden = true;
			ultraGridColumn57.Header.VisiblePosition = 58;
			ultraGridColumn57.Hidden = true;
			ultraGridColumn58.Header.VisiblePosition = 59;
			ultraGridColumn58.Hidden = true;
			ultraGridColumn59.Header.VisiblePosition = 60;
			ultraGridColumn59.Hidden = true;
			ultraGridColumn60.Header.VisiblePosition = 61;
			ultraGridColumn60.Hidden = true;
			ultraGridColumn61.Header.VisiblePosition = 62;
			ultraGridColumn61.Hidden = true;
			ultraGridColumn62.Header.VisiblePosition = 63;
			ultraGridColumn62.Hidden = true;
			ultraGridColumn63.Header.VisiblePosition = 64;
			ultraGridColumn63.Hidden = true;
			ultraGridColumn64.Header.VisiblePosition = 65;
			ultraGridColumn64.Hidden = true;
			ultraGridColumn65.Header.VisiblePosition = 66;
			ultraGridColumn65.Hidden = true;
			appearance2.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn66.CellAppearance = appearance2;
			ultraGridColumn66.Format = "#,##0.00";
			ultraGridColumn66.Header.VisiblePosition = 2;
			ultraGridColumn66.Width = 66;
			ultraGridColumn67.Header.VisiblePosition = 67;
			ultraGridColumn67.Hidden = true;
			ultraGridColumn68.Header.VisiblePosition = 68;
			ultraGridColumn68.Hidden = true;
			ultraGridColumn69.Header.VisiblePosition = 69;
			ultraGridColumn69.Hidden = true;
			ultraGridColumn70.Header.VisiblePosition = 70;
			ultraGridColumn70.Hidden = true;
			ultraGridColumn71.Header.VisiblePosition = 71;
			ultraGridColumn71.Hidden = true;
			ultraGridColumn72.Header.VisiblePosition = 72;
			ultraGridColumn72.Hidden = true;
			ultraGridColumn73.Header.VisiblePosition = 73;
			ultraGridColumn73.Hidden = true;
			ultraGridColumn74.Header.VisiblePosition = 74;
			ultraGridColumn74.Hidden = true;
			ultraGridColumn75.Header.VisiblePosition = 75;
			ultraGridColumn75.Hidden = true;
			ultraGridColumn76.Header.VisiblePosition = 76;
			ultraGridColumn76.Hidden = true;
			ultraGridColumn77.Header.VisiblePosition = 77;
			ultraGridColumn77.Hidden = true;
			ultraGridColumn78.Header.VisiblePosition = 78;
			ultraGridColumn78.Hidden = true;
			ultraGridColumn79.Header.VisiblePosition = 79;
			ultraGridColumn79.Hidden = true;
			ultraGridColumn80.Header.VisiblePosition = 80;
			ultraGridColumn80.Hidden = true;
			ultraGridColumn81.Header.VisiblePosition = 81;
			ultraGridColumn81.Hidden = true;
			ultraGridColumn82.Header.VisiblePosition = 82;
			ultraGridColumn82.Hidden = true;
			ultraGridColumn83.Format = "dd/MMM/yyyy";
			ultraGridColumn83.Header.Caption = "Editado";
			ultraGridColumn83.Header.VisiblePosition = 83;
			ultraGridColumn83.Width = 79;
			ultraGridColumn84.Header.VisiblePosition = 0;
			ultraGridColumn84.Width = 68;
			ultraGridColumn85.Header.VisiblePosition = 84;
			ultraGridColumn86.Header.VisiblePosition = 85;
			ultraGridColumn87.Header.VisiblePosition = 86;
			ultraGridColumn88.Header.VisiblePosition = 87;
			ultraGridColumn89.Header.VisiblePosition = 88;
			ultraGridColumn90.Header.VisiblePosition = 89;
			ultraGridColumn91.Header.VisiblePosition = 90;
			ultraGridColumn92.Header.VisiblePosition = 91;
			ultraGridColumn93.Header.VisiblePosition = 92;
			ultraGridColumn94.Header.VisiblePosition = 93;
			ultraGridColumn95.Header.VisiblePosition = 94;
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
																										 ultraGridColumn95});
			this.ultraGrid1.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.ultraGrid1.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
			this.ultraGrid1.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.ultraGrid1.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.False;
			appearance3.BackColor = System.Drawing.Color.Transparent;
			this.ultraGrid1.DisplayLayout.Override.CardAreaAppearance = appearance3;
			this.ultraGrid1.DisplayLayout.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.RowSelect;
			this.ultraGrid1.DisplayLayout.Override.ColumnAutoSizeMode = Infragistics.Win.UltraWinGrid.ColumnAutoSizeMode.AllRowsInBand;
			appearance4.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance4.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance4.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance4.FontData.BoldAsString = "True";
			appearance4.FontData.Name = "Arial";
			appearance4.FontData.SizeInPoints = 10F;
			appearance4.ForeColor = System.Drawing.Color.White;
			appearance4.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.ultraGrid1.DisplayLayout.Override.HeaderAppearance = appearance4;
			this.ultraGrid1.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti;
			appearance5.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance5.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance5.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.RowSelectorAppearance = appearance5;
			appearance6.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance6.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance6.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.SelectedRowAppearance = appearance6;
			this.ultraGrid1.DisplayLayout.Override.SelectTypeCell = Infragistics.Win.UltraWinGrid.SelectType.None;
			this.ultraGrid1.DisplayLayout.Override.SelectTypeCol = Infragistics.Win.UltraWinGrid.SelectType.None;
			this.ultraGrid1.DisplayLayout.Override.SelectTypeRow = Infragistics.Win.UltraWinGrid.SelectType.Single;
			this.ultraGrid1.DisplayLayout.ScrollBounds = Infragistics.Win.UltraWinGrid.ScrollBounds.ScrollToFill;
			this.ultraGrid1.DisplayLayout.ScrollStyle = Infragistics.Win.UltraWinGrid.ScrollStyle.Immediate;
			this.ultraGrid1.DisplayLayout.TabNavigation = Infragistics.Win.UltraWinGrid.TabNavigation.NextControl;
			this.ultraGrid1.DisplayLayout.ViewStyle = Infragistics.Win.UltraWinGrid.ViewStyle.SingleBand;
			this.ultraGrid1.Location = new System.Drawing.Point(16, 64);
			this.ultraGrid1.Name = "ultraGrid1";
			this.ultraGrid1.Size = new System.Drawing.Size(312, 152);
			this.ultraGrid1.TabIndex = 0;
			// 
			// cdsCompra
			// 
			this.cdsCompra.BindingContextCtrl = this;
			this.cdsCompra.DataLibrary = "LibFacturacion";
			this.cdsCompra.DataLibraryUrl = "";
			this.cdsCompra.DataSetDef = "dsCompra";
			this.cdsCompra.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsCompra.Name = "cdsCompra";
			this.cdsCompra.SchemaClassName = "LibFacturacion.DataClass";
			this.cdsCompra.SchemaDef = null;
			this.cdsCompra.BeforeFill += new C1.Data.FillEventHandler(this.cdsCompra_BeforeFill);
			// 
			// cmbFechaIni
			// 
			dateButton1.Caption = "Today";
			this.cmbFechaIni.DateButtons.Add(dateButton1);
			this.cmbFechaIni.Format = "dd/MMM/yyyy";
			this.cmbFechaIni.Location = new System.Drawing.Point(64, 8);
			this.cmbFechaIni.Name = "cmbFechaIni";
			this.cmbFechaIni.NonAutoSizeHeight = 23;
			this.cmbFechaIni.Size = new System.Drawing.Size(104, 21);
			this.cmbFechaIni.TabIndex = 1;
			// 
			// label1
			// 
			this.label1.AutoSize = true;
			this.label1.Location = new System.Drawing.Point(16, 8);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(40, 16);
			this.label1.TabIndex = 2;
			this.label1.Text = "Desde:";
			// 
			// label2
			// 
			this.label2.AutoSize = true;
			this.label2.Location = new System.Drawing.Point(184, 8);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(37, 16);
			this.label2.TabIndex = 4;
			this.label2.Text = "Hasta:";
			// 
			// cmbFechaFin
			// 
			dateButton2.Caption = "Today";
			this.cmbFechaFin.DateButtons.Add(dateButton2);
			this.cmbFechaFin.Format = "dd/MMM/yyyy";
			this.cmbFechaFin.Location = new System.Drawing.Point(232, 8);
			this.cmbFechaFin.Name = "cmbFechaFin";
			this.cmbFechaFin.NonAutoSizeHeight = 23;
			this.cmbFechaFin.Size = new System.Drawing.Size(104, 21);
			this.cmbFechaFin.TabIndex = 3;
			// 
			// cmbTipoFactura
			// 
			this.cmbTipoFactura.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbTipoFactura.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbTipoFactura.DataMember = "CompraNumero";
			this.cmbTipoFactura.DataSource = this.cdCompraNumero;
			ultraGridColumn96.Header.VisiblePosition = 1;
			ultraGridColumn96.Hidden = true;
			ultraGridColumn97.Header.VisiblePosition = 2;
			ultraGridColumn97.Hidden = true;
			ultraGridColumn98.Header.VisiblePosition = 3;
			ultraGridColumn98.Hidden = true;
			ultraGridColumn99.Header.VisiblePosition = 4;
			ultraGridColumn99.Hidden = true;
			ultraGridColumn100.Header.VisiblePosition = 5;
			ultraGridColumn100.Hidden = true;
			ultraGridColumn101.Header.VisiblePosition = 6;
			ultraGridColumn101.Hidden = true;
			ultraGridColumn102.Header.VisiblePosition = 0;
			ultraGridColumn103.Header.VisiblePosition = 7;
			ultraGridColumn103.Hidden = true;
			ultraGridColumn104.Header.VisiblePosition = 8;
			ultraGridColumn104.Hidden = true;
			ultraGridBand2.Columns.AddRange(new object[] {
																										 ultraGridColumn96,
																										 ultraGridColumn97,
																										 ultraGridColumn98,
																										 ultraGridColumn99,
																										 ultraGridColumn100,
																										 ultraGridColumn101,
																										 ultraGridColumn102,
																										 ultraGridColumn103,
																										 ultraGridColumn104});
			this.cmbTipoFactura.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			this.cmbTipoFactura.DisplayMember = "Nombre";
			this.cmbTipoFactura.Location = new System.Drawing.Point(400, 8);
			this.cmbTipoFactura.Name = "cmbTipoFactura";
			this.cmbTipoFactura.Size = new System.Drawing.Size(168, 21);
			this.cmbTipoFactura.TabIndex = 5;
			this.cmbTipoFactura.ValueMember = "idTipoFactura";
			// 
			// cdCompraNumero
			// 
			this.cdCompraNumero.BindingContextCtrl = this;
			this.cdCompraNumero.DataLibrary = "LibFacturacion";
			this.cdCompraNumero.DataLibraryUrl = "";
			this.cdCompraNumero.DataSetDef = "dsCompraTabla";
			this.cdCompraNumero.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdCompraNumero.Name = "cdCompraNumero";
			this.cdCompraNumero.SchemaClassName = "LibFacturacion.DataClass";
			this.cdCompraNumero.SchemaDef = null;
			this.cdCompraNumero.BeforeFill += new C1.Data.FillEventHandler(this.cdCompraNumero_BeforeFill);
			// 
			// label3
			// 
			this.label3.AutoSize = true;
			this.label3.Location = new System.Drawing.Point(360, 8);
			this.label3.Name = "label3";
			this.label3.Size = new System.Drawing.Size(29, 16);
			this.label3.TabIndex = 6;
			this.label3.Text = "Tipo:";
			// 
			// btFiltro
			// 
			this.btFiltro.Location = new System.Drawing.Point(568, 8);
			this.btFiltro.Name = "btFiltro";
			this.btFiltro.Size = new System.Drawing.Size(24, 23);
			this.btFiltro.TabIndex = 7;
			this.btFiltro.Text = "...";
			this.btFiltro.Click += new System.EventHandler(this.btFiltro_Click);
			// 
			// cdvCompra
			// 
			this.cdvCompra.BindingContextCtrl = this;
			this.cdvCompra.DataSet = this.cdsCompra;
			this.cdvCompra.TableName = "";
			this.cdvCompra.TableViewName = "Compra";
			// 
			// ultraButton1
			// 
			this.ultraButton1.Location = new System.Drawing.Point(336, 208);
			this.ultraButton1.Name = "ultraButton1";
			this.ultraButton1.TabIndex = 9;
			this.ultraButton1.Text = "ultraButton1";
			this.ultraButton1.Click += new System.EventHandler(this.ultraButton1_Click);
			// 
			// chkEditado
			// 
			this.chkEditado.Checked = true;
			this.chkEditado.CheckState = System.Windows.Forms.CheckState.Indeterminate;
			this.chkEditado.Location = new System.Drawing.Point(96, 32);
			this.chkEditado.Name = "chkEditado";
			this.chkEditado.Size = new System.Drawing.Size(64, 20);
			this.chkEditado.TabIndex = 10;
			this.chkEditado.Text = "Editado";
			this.chkEditado.MouseUp += new System.Windows.Forms.MouseEventHandler(this.chkEditado_MouseDown);
			// 
			// chkNuevo
			// 
			this.chkNuevo.Checked = true;
			this.chkNuevo.CheckState = System.Windows.Forms.CheckState.Indeterminate;
			this.chkNuevo.Location = new System.Drawing.Point(16, 32);
			this.chkNuevo.Name = "chkNuevo";
			this.chkNuevo.Size = new System.Drawing.Size(64, 20);
			this.chkNuevo.TabIndex = 11;
			this.chkNuevo.Text = "Nuevo";
			this.chkNuevo.MouseUp += new System.Windows.Forms.MouseEventHandler(this.chkNuevo_MouseDown);
			// 
			// CompraSincroniza
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(648, 266);
			this.Controls.Add(this.chkNuevo);
			this.Controls.Add(this.chkEditado);
			this.Controls.Add(this.ultraButton1);
			this.Controls.Add(this.btFiltro);
			this.Controls.Add(this.label3);
			this.Controls.Add(this.cmbTipoFactura);
			this.Controls.Add(this.label2);
			this.Controls.Add(this.cmbFechaFin);
			this.Controls.Add(this.label1);
			this.Controls.Add(this.cmbFechaIni);
			this.Controls.Add(this.ultraGrid1);
			this.Name = "CompraSincroniza";
			this.Text = "Sincronizacion de Facturacion";
			this.Load += new System.EventHandler(this.CompraSincroniza_Load);
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsCompra)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbFechaIni)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbFechaFin)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbTipoFactura)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdCompraNumero)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvCompra)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void btFiltro_Click(object sender, System.EventArgs e)
		{
			if (cmbTipoFactura.Value == null) 
			{
				MessageBox.Show("Ingrese Tipo de Factura");
				return;
			}
			DateTime dtFechaIni = (DateTime) this.cmbFechaIni.Value;
			DateTime dtFechaFin = (DateTime) this.cmbFechaFin.Value;
			string strFiltro = "Fecha >= '" + 
				dtFechaIni.ToString("yyyyMMdd") + "'" + 
				" and Fecha < '" + 
				dtFechaFin.AddDays(1).ToString("yyyyMMdd") + "'";

			string stFiltro = strFiltro + " And idTipoFactura = " + cmbTipoFactura.Value.ToString();
			if (chkEditado.CheckState != CheckState.Indeterminate)
			{
				if (chkEditado.Checked)	stFiltro += " And Editar <> 0";
				else stFiltro += " And Editar = 0";
			}
			if (chkNuevo.CheckState != CheckState.Indeterminate)
			{
				if (chkNuevo.Checked)	stFiltro += " And Nuevo <> 0";
				else stFiltro += " And Nuevo = 0";
			}
			MessageBox.Show(stFiltro);
			C1.Data.FilterCondition fcPadre = new 
				C1.Data.FilterCondition(this.cdsCompra, "Compra", stFiltro);
			C1.Data.FilterConditions fcTotal = new C1.Data.FilterConditions();
			fcTotal.Add(fcPadre);

			C1.Data.FilterCondition fcHija = new 
				C1.Data.FilterCondition(this.cdsCompra, "DetCompra", "idCompra = 0");
			fcTotal.Add(fcHija);
			
			this.cdsCompra.Fill(fcTotal, false);
			if (this.cdvCompra.Count == 0)
			{
				MessageBox.Show("No existen registros que cumplan la condicion de busqueda",
					"Información", MessageBoxButtons.OK, MessageBoxIcon.Information);
				return;
			}
		}

		private void cdsCompra_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			this.cdsCompra.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);
		}

		private void cdCompraNumero_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			this.cdCompraNumero.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);
		}

		private void chkNuevo_MouseDown(object sender, System.Windows.Forms.MouseEventArgs e)
		{
			if (e.Button != MouseButtons.Right) return;
      chkNuevo.CheckState = CheckState.Indeterminate;
		}

		private void chkEditado_MouseDown(object sender, System.Windows.Forms.MouseEventArgs e)
		{
			if (e.Button == MouseButtons.Right)
				chkEditado.CheckState = CheckState.Indeterminate;		
		}

		private void ultraButton1_Click(object sender, System.EventArgs e)
		{
		
		}

		private void CompraSincroniza_Load(object sender, System.EventArgs e)
		{
		
		}
	}
}
