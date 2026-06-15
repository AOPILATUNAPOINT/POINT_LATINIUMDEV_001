using System;
using System.Data.SqlClient;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de CompraLotes.
	/// </summary>
	public class CompraLotes : DevExpress.XtraEditors.XtraForm
	{
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtFactura;
		private System.Windows.Forms.Label label1;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbCliente;
		private System.Windows.Forms.Label label2;
		private C1.Data.C1DataSet cdsCliente;
		private C1.Data.C1DataView cdvCliente;
		private Infragistics.Win.Misc.UltraButton btDetalle;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource udsDetalle;
		private C1.Data.C1DataSet cdsDepende;
		private Infragistics.Win.Misc.UltraButton btDepende;
		private Infragistics.Win.UltraWinTabControl.UltraTabControl ultraTabControl1;
		private Infragistics.Win.UltraWinTabControl.UltraTabSharedControlsPage ultraTabSharedControlsPage1;
		private Infragistics.Win.UltraWinTabControl.UltraTabPageControl ultraTabPageControl1;
		private Infragistics.Win.UltraWinTabControl.UltraTabPageControl ultraTabPageControl2;
		private Infragistics.Win.UltraWinGrid.UltraGrid grDetalle;
		private Infragistics.Win.UltraWinGrid.UltraGrid grDependiente;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public CompraLotes()
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
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn63 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Numero");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn64 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Fecha");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn65 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Vencimiento");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn66 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cliente");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn67 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Articulo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn68 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cantidad");
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn69 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Precio");
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn70 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Iva");
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn71 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCliente");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn72 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCompra");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn73 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idArticulo");
			Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance11 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance12 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance13 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Numero");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Fecha");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn3 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Vencimiento");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn4 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Cliente");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn5 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Articulo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn6 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Cantidad");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn7 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Precio");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn8 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Iva");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn9 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCliente");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn10 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCompra");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn11 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idArticulo");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand3 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Cliente", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn74 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn75 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCliente");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn76 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn77 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Proveedor");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn78 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Ruc");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn79 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Ambos");
			Infragistics.Win.UltraWinTabControl.UltraTab ultraTab1 = new Infragistics.Win.UltraWinTabControl.UltraTab();
			Infragistics.Win.UltraWinTabControl.UltraTab ultraTab2 = new Infragistics.Win.UltraWinTabControl.UltraTab();
			this.ultraTabPageControl1 = new Infragistics.Win.UltraWinTabControl.UltraTabPageControl();
			this.grDependiente = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.cdsDepende = new C1.Data.C1DataSet();
			this.ultraTabPageControl2 = new Infragistics.Win.UltraWinTabControl.UltraTabPageControl();
			this.grDetalle = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.udsDetalle = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.txtFactura = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label1 = new System.Windows.Forms.Label();
			this.cmbCliente = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.cdvCliente = new C1.Data.C1DataView();
			this.cdsCliente = new C1.Data.C1DataSet();
			this.label2 = new System.Windows.Forms.Label();
			this.btDetalle = new Infragistics.Win.Misc.UltraButton();
			this.btDepende = new Infragistics.Win.Misc.UltraButton();
			this.ultraTabControl1 = new Infragistics.Win.UltraWinTabControl.UltraTabControl();
			this.ultraTabSharedControlsPage1 = new Infragistics.Win.UltraWinTabControl.UltraTabSharedControlsPage();
			this.ultraTabPageControl1.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.grDependiente)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsDepende)).BeginInit();
			this.ultraTabPageControl2.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.grDetalle)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.udsDetalle)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtFactura)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCliente)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvCliente)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsCliente)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraTabControl1)).BeginInit();
			this.ultraTabControl1.SuspendLayout();
			this.SuspendLayout();
			// 
			// ultraTabPageControl1
			// 
			this.ultraTabPageControl1.Controls.Add(this.grDependiente);
			this.ultraTabPageControl1.Location = new System.Drawing.Point(2, 24);
			this.ultraTabPageControl1.Name = "ultraTabPageControl1";
			this.ultraTabPageControl1.Size = new System.Drawing.Size(636, 262);
			// 
			// grDependiente
			// 
			this.grDependiente.Cursor = System.Windows.Forms.Cursors.Default;
			this.grDependiente.DataMember = "Cliente";
			this.grDependiente.DataSource = this.cdsDepende;
			appearance1.BackColor = System.Drawing.Color.White;
			this.grDependiente.DisplayLayout.Appearance = appearance1;
			ultraGridColumn1.Header.VisiblePosition = 20;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn2.Header.VisiblePosition = 22;
			ultraGridColumn2.Hidden = true;
			ultraGridColumn3.Header.VisiblePosition = 21;
			ultraGridColumn3.Hidden = true;
			ultraGridColumn4.Header.VisiblePosition = 26;
			ultraGridColumn4.Hidden = true;
			ultraGridColumn5.Header.VisiblePosition = 24;
			ultraGridColumn5.Hidden = true;
			ultraGridColumn6.Header.VisiblePosition = 19;
			ultraGridColumn6.Hidden = true;
			ultraGridColumn7.Header.VisiblePosition = 23;
			ultraGridColumn7.Hidden = true;
			ultraGridColumn8.Header.VisiblePosition = 25;
			ultraGridColumn8.Hidden = true;
			ultraGridColumn9.Header.VisiblePosition = 34;
			ultraGridColumn9.Hidden = true;
			ultraGridColumn10.Header.VisiblePosition = 40;
			ultraGridColumn10.Hidden = true;
			ultraGridColumn11.Header.VisiblePosition = 44;
			ultraGridColumn11.Hidden = true;
			ultraGridColumn12.Header.VisiblePosition = 47;
			ultraGridColumn12.Hidden = true;
			ultraGridColumn13.Header.VisiblePosition = 0;
			ultraGridColumn13.Width = 185;
			ultraGridColumn14.Header.VisiblePosition = 9;
			ultraGridColumn14.Hidden = true;
			ultraGridColumn15.Header.VisiblePosition = 1;
			ultraGridColumn15.Width = 176;
			ultraGridColumn16.Header.VisiblePosition = 2;
			ultraGridColumn17.Header.VisiblePosition = 15;
			ultraGridColumn17.Hidden = true;
			ultraGridColumn18.Header.VisiblePosition = 3;
			ultraGridColumn19.Header.VisiblePosition = 29;
			ultraGridColumn19.Hidden = true;
			ultraGridColumn20.Header.VisiblePosition = 32;
			ultraGridColumn20.Hidden = true;
			ultraGridColumn21.Header.VisiblePosition = 14;
			ultraGridColumn21.Hidden = true;
			ultraGridColumn22.Header.VisiblePosition = 8;
			ultraGridColumn22.Hidden = true;
			ultraGridColumn23.Header.VisiblePosition = 10;
			ultraGridColumn23.Hidden = true;
			ultraGridColumn24.Header.VisiblePosition = 18;
			ultraGridColumn24.Hidden = true;
			ultraGridColumn25.Header.VisiblePosition = 30;
			ultraGridColumn25.Hidden = true;
			ultraGridColumn26.Header.VisiblePosition = 28;
			ultraGridColumn26.Hidden = true;
			ultraGridColumn27.Header.VisiblePosition = 4;
			ultraGridColumn27.Hidden = true;
			ultraGridColumn28.Header.VisiblePosition = 31;
			ultraGridColumn28.Hidden = true;
			ultraGridColumn29.Header.VisiblePosition = 5;
			ultraGridColumn29.Hidden = true;
			ultraGridColumn30.Header.VisiblePosition = 6;
			ultraGridColumn30.Hidden = true;
			ultraGridColumn31.Header.VisiblePosition = 27;
			ultraGridColumn31.Hidden = true;
			ultraGridColumn32.Header.VisiblePosition = 13;
			ultraGridColumn32.Hidden = true;
			ultraGridColumn33.Header.VisiblePosition = 33;
			ultraGridColumn33.Hidden = true;
			ultraGridColumn34.Header.VisiblePosition = 16;
			ultraGridColumn34.Hidden = true;
			ultraGridColumn35.Header.VisiblePosition = 17;
			ultraGridColumn35.Hidden = true;
			ultraGridColumn36.Header.VisiblePosition = 11;
			ultraGridColumn36.Hidden = true;
			ultraGridColumn37.Header.VisiblePosition = 12;
			ultraGridColumn37.Hidden = true;
			ultraGridColumn38.Header.VisiblePosition = 7;
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
			ultraGridColumn44.Hidden = true;
			ultraGridColumn45.Header.VisiblePosition = 42;
			ultraGridColumn45.Hidden = true;
			ultraGridColumn46.Header.VisiblePosition = 43;
			ultraGridColumn46.Hidden = true;
			ultraGridColumn47.Header.VisiblePosition = 45;
			ultraGridColumn47.Hidden = true;
			ultraGridColumn48.Header.VisiblePosition = 46;
			ultraGridColumn48.Hidden = true;
			ultraGridColumn49.Header.VisiblePosition = 48;
			ultraGridColumn49.Hidden = true;
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
			this.grDependiente.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.grDependiente.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
			this.grDependiente.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.grDependiente.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.False;
			appearance2.BackColor = System.Drawing.Color.Transparent;
			this.grDependiente.DisplayLayout.Override.CardAreaAppearance = appearance2;
			this.grDependiente.DisplayLayout.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.RowSelect;
			this.grDependiente.DisplayLayout.Override.ColumnAutoSizeMode = Infragistics.Win.UltraWinGrid.ColumnAutoSizeMode.AllRowsInBand;
			appearance3.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance3.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance3.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance3.FontData.BoldAsString = "True";
			appearance3.FontData.Name = "Arial";
			appearance3.FontData.SizeInPoints = 10F;
			appearance3.ForeColor = System.Drawing.Color.White;
			appearance3.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.grDependiente.DisplayLayout.Override.HeaderAppearance = appearance3;
			this.grDependiente.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti;
			appearance4.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance4.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance4.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.grDependiente.DisplayLayout.Override.RowSelectorAppearance = appearance4;
			appearance5.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance5.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance5.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.grDependiente.DisplayLayout.Override.SelectedRowAppearance = appearance5;
			this.grDependiente.DisplayLayout.Override.SelectTypeCell = Infragistics.Win.UltraWinGrid.SelectType.None;
			this.grDependiente.DisplayLayout.Override.SelectTypeCol = Infragistics.Win.UltraWinGrid.SelectType.None;
			this.grDependiente.DisplayLayout.Override.SelectTypeRow = Infragistics.Win.UltraWinGrid.SelectType.Single;
			this.grDependiente.DisplayLayout.ScrollBounds = Infragistics.Win.UltraWinGrid.ScrollBounds.ScrollToFill;
			this.grDependiente.DisplayLayout.ScrollStyle = Infragistics.Win.UltraWinGrid.ScrollStyle.Immediate;
			this.grDependiente.DisplayLayout.TabNavigation = Infragistics.Win.UltraWinGrid.TabNavigation.NextControl;
			this.grDependiente.DisplayLayout.ViewStyle = Infragistics.Win.UltraWinGrid.ViewStyle.SingleBand;
			this.grDependiente.Location = new System.Drawing.Point(16, 16);
			this.grDependiente.Name = "grDependiente";
			this.grDependiente.Size = new System.Drawing.Size(600, 232);
			this.grDependiente.TabIndex = 0;
			// 
			// cdsDepende
			// 
			this.cdsDepende.BindingContextCtrl = this;
			this.cdsDepende.DataLibrary = "LibFacturacion";
			this.cdsDepende.DataLibraryUrl = "";
			this.cdsDepende.DataSetDef = "dsCliente";
			this.cdsDepende.FillOnRequest = false;
			this.cdsDepende.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsDepende.Name = "cdsDepende";
			this.cdsDepende.SchemaClassName = "LibFacturacion.DataClass";
			this.cdsDepende.SchemaDef = null;
			this.cdsDepende.BeforeFill += new C1.Data.FillEventHandler(this.cdsDepende_BeforeFill);
			// 
			// ultraTabPageControl2
			// 
			this.ultraTabPageControl2.Controls.Add(this.grDetalle);
			this.ultraTabPageControl2.Location = new System.Drawing.Point(-10000, -10000);
			this.ultraTabPageControl2.Name = "ultraTabPageControl2";
			this.ultraTabPageControl2.Size = new System.Drawing.Size(636, 262);
			// 
			// grDetalle
			// 
			this.grDetalle.Cursor = System.Windows.Forms.Cursors.Default;
			this.grDetalle.DataSource = this.udsDetalle;
			appearance6.BackColor = System.Drawing.Color.White;
			this.grDetalle.DisplayLayout.Appearance = appearance6;
			ultraGridColumn63.Header.VisiblePosition = 0;
			ultraGridColumn63.Width = 73;
			ultraGridColumn64.Format = "dd/MMM/yyyy";
			ultraGridColumn64.Header.VisiblePosition = 1;
			ultraGridColumn64.Width = 69;
			ultraGridColumn65.Format = "dd/MMM/yyyy";
			ultraGridColumn65.Header.VisiblePosition = 2;
			ultraGridColumn65.Width = 69;
			ultraGridColumn66.Header.VisiblePosition = 3;
			ultraGridColumn66.Width = 121;
			ultraGridColumn67.Header.VisiblePosition = 4;
			appearance7.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn68.CellAppearance = appearance7;
			ultraGridColumn68.Format = "#,##0.00";
			ultraGridColumn68.Header.VisiblePosition = 5;
			ultraGridColumn68.Width = 47;
			appearance8.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn69.CellAppearance = appearance8;
			ultraGridColumn69.Format = "#,##0.00";
			ultraGridColumn69.Header.VisiblePosition = 6;
			ultraGridColumn69.Width = 59;
			appearance9.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn70.CellAppearance = appearance9;
			ultraGridColumn70.Header.VisiblePosition = 7;
			ultraGridColumn70.Width = 38;
			ultraGridColumn71.Header.VisiblePosition = 8;
			ultraGridColumn71.Hidden = true;
			ultraGridColumn72.Header.VisiblePosition = 9;
			ultraGridColumn72.Hidden = true;
			ultraGridColumn73.Header.VisiblePosition = 10;
			ultraGridColumn73.Hidden = true;
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
																										 ultraGridColumn73});
			this.grDetalle.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			this.grDetalle.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
			this.grDetalle.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.grDetalle.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.False;
			appearance10.BackColor = System.Drawing.Color.Transparent;
			this.grDetalle.DisplayLayout.Override.CardAreaAppearance = appearance10;
			this.grDetalle.DisplayLayout.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.RowSelect;
			this.grDetalle.DisplayLayout.Override.ColumnAutoSizeMode = Infragistics.Win.UltraWinGrid.ColumnAutoSizeMode.AllRowsInBand;
			appearance11.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance11.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance11.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance11.FontData.BoldAsString = "True";
			appearance11.FontData.Name = "Arial";
			appearance11.FontData.SizeInPoints = 10F;
			appearance11.ForeColor = System.Drawing.Color.White;
			appearance11.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.grDetalle.DisplayLayout.Override.HeaderAppearance = appearance11;
			this.grDetalle.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti;
			appearance12.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance12.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance12.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.grDetalle.DisplayLayout.Override.RowSelectorAppearance = appearance12;
			appearance13.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance13.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance13.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.grDetalle.DisplayLayout.Override.SelectedRowAppearance = appearance13;
			this.grDetalle.DisplayLayout.Override.SelectTypeCell = Infragistics.Win.UltraWinGrid.SelectType.None;
			this.grDetalle.DisplayLayout.Override.SelectTypeCol = Infragistics.Win.UltraWinGrid.SelectType.None;
			this.grDetalle.DisplayLayout.Override.SelectTypeRow = Infragistics.Win.UltraWinGrid.SelectType.Single;
			this.grDetalle.DisplayLayout.ScrollBounds = Infragistics.Win.UltraWinGrid.ScrollBounds.ScrollToFill;
			this.grDetalle.DisplayLayout.ScrollStyle = Infragistics.Win.UltraWinGrid.ScrollStyle.Immediate;
			this.grDetalle.DisplayLayout.TabNavigation = Infragistics.Win.UltraWinGrid.TabNavigation.NextControl;
			this.grDetalle.DisplayLayout.ViewStyle = Infragistics.Win.UltraWinGrid.ViewStyle.SingleBand;
			this.grDetalle.Location = new System.Drawing.Point(8, 16);
			this.grDetalle.Name = "grDetalle";
			this.grDetalle.Size = new System.Drawing.Size(608, 232);
			this.grDetalle.TabIndex = 5;
			// 
			// udsDetalle
			// 
			ultraDataColumn2.DataType = typeof(System.DateTime);
			ultraDataColumn3.DataType = typeof(System.DateTime);
			ultraDataColumn6.DataType = typeof(System.Double);
			ultraDataColumn6.DefaultValue = 0;
			ultraDataColumn7.DataType = typeof(System.Double);
			ultraDataColumn7.DefaultValue = 0;
			ultraDataColumn8.DataType = typeof(System.Double);
			ultraDataColumn9.DataType = typeof(int);
			ultraDataColumn9.DefaultValue = 0;
			ultraDataColumn10.DataType = typeof(int);
			ultraDataColumn10.DefaultValue = 0;
			ultraDataColumn11.DataType = typeof(int);
			ultraDataColumn11.DefaultValue = 0;
			this.udsDetalle.Band.Columns.AddRange(new object[] {
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
																													 ultraDataColumn11});
			// 
			// txtFactura
			// 
			this.txtFactura.Location = new System.Drawing.Point(96, 32);
			this.txtFactura.Name = "txtFactura";
			this.txtFactura.Size = new System.Drawing.Size(100, 21);
			this.txtFactura.TabIndex = 0;
			// 
			// label1
			// 
			this.label1.AutoSize = true;
			this.label1.Location = new System.Drawing.Point(32, 32);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(43, 16);
			this.label1.TabIndex = 1;
			this.label1.Text = "Factura";
			// 
			// cmbCliente
			// 
			this.cmbCliente.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbCliente.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbCliente.DataSource = this.cdvCliente;
			this.cmbCliente.DisplayLayout.AutoFitColumns = true;
			ultraGridColumn74.Header.VisiblePosition = 2;
			ultraGridColumn74.Hidden = true;
			ultraGridColumn75.Header.VisiblePosition = 3;
			ultraGridColumn75.Hidden = true;
			ultraGridColumn76.Header.VisiblePosition = 0;
			ultraGridColumn76.Width = 254;
			ultraGridColumn77.Header.VisiblePosition = 4;
			ultraGridColumn77.Hidden = true;
			ultraGridColumn78.Header.VisiblePosition = 5;
			ultraGridColumn78.Hidden = true;
			ultraGridColumn79.Header.VisiblePosition = 1;
			ultraGridColumn79.Hidden = true;
			ultraGridBand3.Columns.AddRange(new object[] {
																										 ultraGridColumn74,
																										 ultraGridColumn75,
																										 ultraGridColumn76,
																										 ultraGridColumn77,
																										 ultraGridColumn78,
																										 ultraGridColumn79});
			this.cmbCliente.DisplayLayout.BandsSerializer.Add(ultraGridBand3);
			this.cmbCliente.DisplayMember = "Nombre";
			this.cmbCliente.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbCliente.Location = new System.Drawing.Point(96, 64);
			this.cmbCliente.Name = "cmbCliente";
			this.cmbCliente.Size = new System.Drawing.Size(272, 21);
			this.cmbCliente.TabIndex = 2;
			this.cmbCliente.ValueMember = "idCliente";
			// 
			// cdvCliente
			// 
			this.cdvCliente.BindingContextCtrl = this;
			this.cdvCliente.DataSet = this.cdsCliente;
			this.cdvCliente.RowFilter = "Proveedor = 0 Or Ambos <> 0";
			this.cdvCliente.Sort = "Nombre";
			this.cdvCliente.TableName = "";
			this.cdvCliente.TableViewName = "Cliente";
			// 
			// cdsCliente
			// 
			this.cdsCliente.BindingContextCtrl = this;
			this.cdsCliente.DataLibrary = "LibFacturacion";
			this.cdsCliente.DataLibraryUrl = "";
			this.cdsCliente.DataSetDef = "dsClientes";
			this.cdsCliente.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsCliente.Name = "cdsCliente";
			this.cdsCliente.SchemaClassName = "LibFacturacion.DataClass";
			this.cdsCliente.SchemaDef = null;
			this.cdsCliente.BeforeFill += new C1.Data.FillEventHandler(this.cdsCliente_BeforeFill);
			// 
			// label2
			// 
			this.label2.AutoSize = true;
			this.label2.Location = new System.Drawing.Point(32, 64);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(40, 16);
			this.label2.TabIndex = 3;
			this.label2.Text = "Cliente";
			// 
			// btDetalle
			// 
			this.btDetalle.Location = new System.Drawing.Point(480, 64);
			this.btDetalle.Name = "btDetalle";
			this.btDetalle.TabIndex = 5;
			this.btDetalle.Text = "&Detalle";
			this.btDetalle.Click += new System.EventHandler(this.btDetalle_Click);
			// 
			// btDepende
			// 
			this.btDepende.Location = new System.Drawing.Point(376, 64);
			this.btDepende.Name = "btDepende";
			this.btDepende.Size = new System.Drawing.Size(96, 23);
			this.btDepende.TabIndex = 6;
			this.btDepende.Text = "Dependientes";
			this.btDepende.Click += new System.EventHandler(this.btDepende_Click);
			// 
			// ultraTabControl1
			// 
			this.ultraTabControl1.Controls.Add(this.ultraTabSharedControlsPage1);
			this.ultraTabControl1.Controls.Add(this.ultraTabPageControl1);
			this.ultraTabControl1.Controls.Add(this.ultraTabPageControl2);
			this.ultraTabControl1.Location = new System.Drawing.Point(16, 96);
			this.ultraTabControl1.Name = "ultraTabControl1";
			this.ultraTabControl1.SharedControlsPage = this.ultraTabSharedControlsPage1;
			this.ultraTabControl1.Size = new System.Drawing.Size(640, 288);
			this.ultraTabControl1.TabIndex = 7;
			ultraTab1.TabPage = this.ultraTabPageControl1;
			ultraTab1.Text = "Dependientes";
			ultraTab2.TabPage = this.ultraTabPageControl2;
			ultraTab2.Text = "Detalle";
			this.ultraTabControl1.Tabs.AddRange(new Infragistics.Win.UltraWinTabControl.UltraTab[] {
																																															 ultraTab1,
																																															 ultraTab2});
			// 
			// ultraTabSharedControlsPage1
			// 
			this.ultraTabSharedControlsPage1.Location = new System.Drawing.Point(-10000, -10000);
			this.ultraTabSharedControlsPage1.Name = "ultraTabSharedControlsPage1";
			this.ultraTabSharedControlsPage1.Size = new System.Drawing.Size(636, 262);
			// 
			// CompraLotes
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(672, 414);
			this.Controls.Add(this.ultraTabControl1);
			this.Controls.Add(this.btDepende);
			this.Controls.Add(this.btDetalle);
			this.Controls.Add(this.label2);
			this.Controls.Add(this.label1);
			this.Controls.Add(this.cmbCliente);
			this.Controls.Add(this.txtFactura);
			this.Name = "CompraLotes";
			this.Text = "Facturación por Lotes";
			this.Closing += new System.ComponentModel.CancelEventHandler(this.CompraLotes_Closing);
			this.Load += new System.EventHandler(this.CompraLotes_Load);
			this.ultraTabPageControl1.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.grDependiente)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsDepende)).EndInit();
			this.ultraTabPageControl2.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.grDetalle)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.udsDetalle)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtFactura)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCliente)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvCliente)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsCliente)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraTabControl1)).EndInit();
			this.ultraTabControl1.ResumeLayout(false);
			this.ResumeLayout(false);

		}
		#endregion

		private void btDetalle_Click(object sender, System.EventArgs e)
		{
			if (cmbCliente.Value == System.DBNull.Value 
				|| cmbCliente.Text.Trim().Length == 0)
				return;

			int IdPadre = (int) cmbCliente.Value;
			string stSelect = "SELECT Compra.Numero, Compra.Fecha, "
				+ "Compra.FechaVencimiento, Cliente.Nombre, "
				+ "DetCompra.Cantidad, DetCompra.Precio, Articulo.Articulo, "
				+ "Cliente.idCliente, Compra.idCompra, Articulo.idArticulo "
				+ "FROM Compra INNER JOIN Cliente "
				+ "ON Compra.idCliente = Cliente.idCliente "
				+ "INNER JOIN DetCompra ON Compra.idCompra = DetCompra.idCompra "
				+ "INNER JOIN Articulo ON DetCompra.idArticulo = Articulo.idArticulo "
				+ "WHERE Compra.idTipoFactura = 13 AND Cliente.idPadre = " 
				+ IdPadre.ToString();

			SqlDataReader miReader = Funcion.rEscalarSQL(cdsCliente, stSelect, true);

			int iNumFilas = -1;
			udsDetalle.Rows.Clear();

			while (miReader.Read())
			{
				string stFactura = miReader.GetString(0);
				DateTime dtFecha = miReader.GetDateTime(1);
				DateTime dtVence = miReader.GetDateTime(2);
				string stCliente = miReader.GetString(3);
				double dCantidad = miReader.GetDouble(4);
				double dPrecio = miReader.GetDouble(5);
				string stArticulo = miReader.GetString(6);
				int IdCliente = miReader.GetInt32(7);
				int IdCompra = miReader.GetInt32(8);
				int IdArticulo = miReader.GetInt32(9);
			{
				iNumFilas ++;
				udsDetalle.Rows.Add();
				udsDetalle.Rows[iNumFilas]["Numero"] = stFactura;
				udsDetalle.Rows[iNumFilas]["Fecha"] = dtFecha;
				udsDetalle.Rows[iNumFilas]["Vencimiento"] = dtVence;
				udsDetalle.Rows[iNumFilas]["Cliente"] = stCliente;
				udsDetalle.Rows[iNumFilas]["Articulo"] = stArticulo;
				udsDetalle.Rows[iNumFilas]["Cantidad"] = dCantidad;
				udsDetalle.Rows[iNumFilas]["Precio"] = dPrecio;
				udsDetalle.Rows[iNumFilas]["idCliente"] = IdCliente;
				udsDetalle.Rows[iNumFilas]["idCompra"] = IdCompra;
				udsDetalle.Rows[iNumFilas]["idArticulo"] = IdArticulo;
			}
			}

			cdsCliente.Schema.Connections[0].Close();
		}

		private void btDepende_Click(object sender, System.EventArgs e)
		{
			if (cmbCliente.Value == System.DBNull.Value 
				|| cmbCliente.Text.Trim().Length == 0)
				return;

			int IdPadre = 0;

			IdPadre = (int) cmbCliente.Value;
			string stFiltro = "idPadre = " + IdPadre.ToString();
			C1.Data.FilterCondition fcPadre = new 
				C1.Data.FilterCondition(cdsDepende, "Cliente", stFiltro);
			C1.Data.FilterConditions fcTotal = new C1.Data.FilterConditions();
			fcTotal.Add(fcPadre);
			
			cdsDepende.Fill(fcTotal, false);
		}

		private void cdsCliente_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsCliente.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);
		}

		private void cdsDepende_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsDepende.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);
		}

		private void CompraLotes_Load(object sender, System.EventArgs e)
		{
			string stAudita = "Exec AuditaCrear 53, 6, 'Fact. Lote'";
			Funcion.EjecutaSQL(cdsCliente, stAudita, true);
		}

		private void CompraLotes_Closing(object sender, System.ComponentModel.CancelEventArgs e)
		{
			string stAudita = "Exec AuditaCrear 53, 7, 'Fact. Lote'";
			Funcion.EjecutaSQL(cdsCliente, stAudita, true);
		}
	}
}
