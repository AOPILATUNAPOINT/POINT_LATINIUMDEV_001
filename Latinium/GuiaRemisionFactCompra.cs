using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de GuiaRemisionFact.
	/// </summary>
	public class GuiaRemisionFactCompra : DevExpress.XtraEditors.XtraForm
	{
		private Infragistics.Win.Misc.UltraButton btFiltro;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbCliente;
		private C1.Data.C1DataView cdvCliente;
		private C1.Data.C1DataSet cdsCliente;
		private C1.Data.C1DataView cdvGuiaRemisionG;
		private DevExpress.XtraEditors.CheckEdit chkCliente;
		private DevExpress.XtraEditors.CheckEdit chkFecha;
		private DevExpress.XtraEditors.DateEdit dtFechaDesde;
		private DevExpress.XtraEditors.DateEdit dtFechaHasta;
		private DevExpress.XtraEditors.CheckEdit chkGuia;
		private DevExpress.XtraEditors.CheckEdit chkRuta;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbPuntoPartida;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbPuntoLlegada;
		private C1.Data.C1DataView cdvGuiaPunto;
		private C1.Data.C1DataSet cdsGuiaPunto;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtGuia1;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtGuia2;
		private C1.Data.C1DataSet cdsGuiaRemisionG;
		private Infragistics.Win.UltraWinGrid.UltraGrid ultraGrid1;
		private C1.Data.C1DataSet cdsGuiaRemisionPadre;
		private C1.Data.C1DataView cdvGuiaRemisionPadre;
		private BarraDatoSQL.BarraDatoSQL barraDato1;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtidGuiaRemisionPadre;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtNumero;

		C1.Data.C1DataRow drGuiaPadre;
		private Infragistics.Win.Misc.UltraButton btnAsiento;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbCliente1;
		private System.Windows.Forms.Label label18;
		private System.Windows.Forms.Label label1;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtFecha;
		private System.Windows.Forms.ErrorProvider errorProvider;
		private Infragistics.Win.UltraWinGrid.UltraDropDown cmbLlegada;
		private Infragistics.Win.UltraWinGrid.UltraDropDown cmbPartida;
		int editar=0;
		public GuiaRemisionFactCompra()
		{
			InitializeComponent();
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
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCliente");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Proveedor");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Ruc");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Ambos");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("DesconProv");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("DesconCliente");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Entrega");
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("GuiaRemisionPunto", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idGuiaRemisionPunto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn11 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand3 = new Infragistics.Win.UltraWinGrid.UltraGridBand("GuiaRemisionPunto", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn12 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idGuiaRemisionPunto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn13 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand4 = new Infragistics.Win.UltraWinGrid.UltraGridBand("GuiaRemisionG", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn14 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idGuiaRemisionG");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn15 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("fInicioTraslado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn16 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("fTermTraslado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn17 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("fEmision");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn18 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("NroGuia");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn19 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("venta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn20 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("compra");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn21 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("transformacion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn22 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("consignacion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn23 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("traslado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn24 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("devolucion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn25 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("importacion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn26 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("exportacion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn27 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("otros");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn28 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("fEmision2");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn29 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPuntoPartida");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn30 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("rucCliente");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn31 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPuntoLlegada", -1, "cmbLlegada");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn32 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("nombreTrasnp");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn33 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("rucTransp");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn34 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("observaciones");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn35 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("diesel");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn36 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("adelanto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn37 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("contratado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn38 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idTransportista");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn39 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCiudadEmision");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn40 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Borrar");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn41 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCarro1");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn42 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCarro2");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn43 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCarro3");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn44 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCliente");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn45 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("NombreCliente");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn46 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("TonelajePedido");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn47 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CombAutorizado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn48 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CombConsumido");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn49 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Celular");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn50 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Placa");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn51 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Facturado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn52 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cobrado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn53 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCompra");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn54 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idVenta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn55 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idGuiaRemisionPadre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn56 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("TextoTonelaje");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn57 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Observacion2");
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand5 = new Infragistics.Win.UltraWinGrid.UltraGridBand("GuiaRemisionPunto", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn58 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idGuiaRemisionPunto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn59 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.Appearance appearance11 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance12 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance13 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand6 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Cliente", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn60 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn61 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCliente");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn62 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn63 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Proveedor");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn64 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Ruc");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn65 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Ambos");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn66 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("DesconProv");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn67 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("DesconCliente");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn68 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Entrega");
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton1 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand7 = new Infragistics.Win.UltraWinGrid.UltraGridBand("GuiaRemisionPunto", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn69 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idGuiaRemisionPunto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn70 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			this.btFiltro = new Infragistics.Win.Misc.UltraButton();
			this.cmbCliente = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.cdvCliente = new C1.Data.C1DataView();
			this.cdsCliente = new C1.Data.C1DataSet();
			this.cdvGuiaRemisionG = new C1.Data.C1DataView();
			this.chkGuia = new DevExpress.XtraEditors.CheckEdit();
			this.chkFecha = new DevExpress.XtraEditors.CheckEdit();
			this.chkCliente = new DevExpress.XtraEditors.CheckEdit();
			this.dtFechaHasta = new DevExpress.XtraEditors.DateEdit();
			this.dtFechaDesde = new DevExpress.XtraEditors.DateEdit();
			this.chkRuta = new DevExpress.XtraEditors.CheckEdit();
			this.cmbPuntoPartida = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.cdsGuiaPunto = new C1.Data.C1DataSet();
			this.cmbPuntoLlegada = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.cdvGuiaPunto = new C1.Data.C1DataView();
			this.txtGuia1 = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtGuia2 = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.cdsGuiaRemisionG = new C1.Data.C1DataSet();
			this.ultraGrid1 = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.cdsGuiaRemisionPadre = new C1.Data.C1DataSet();
			this.cmbLlegada = new Infragistics.Win.UltraWinGrid.UltraDropDown();
			this.cdvGuiaRemisionPadre = new C1.Data.C1DataView();
			this.barraDato1 = new BarraDatoSQL.BarraDatoSQL();
			this.txtidGuiaRemisionPadre = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtNumero = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.btnAsiento = new Infragistics.Win.Misc.UltraButton();
			this.cmbCliente1 = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.label18 = new System.Windows.Forms.Label();
			this.label1 = new System.Windows.Forms.Label();
			this.dtFecha = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.errorProvider = new System.Windows.Forms.ErrorProvider();
			this.cmbPartida = new Infragistics.Win.UltraWinGrid.UltraDropDown();
			((System.ComponentModel.ISupportInitialize)(this.cmbCliente)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvCliente)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsCliente)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvGuiaRemisionG)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.chkGuia.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.chkFecha.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.chkCliente.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFechaHasta.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFechaDesde.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.chkRuta.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbPuntoPartida)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsGuiaPunto)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbPuntoLlegada)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvGuiaPunto)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtGuia1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtGuia2)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsGuiaRemisionG)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsGuiaRemisionPadre)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbLlegada)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvGuiaRemisionPadre)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtidGuiaRemisionPadre)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNumero)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCliente1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFecha)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbPartida)).BeginInit();
			this.SuspendLayout();
			// 
			// btFiltro
			// 
			appearance1.Image = 5;
			this.btFiltro.Appearance = appearance1;
			this.btFiltro.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.btFiltro.Location = new System.Drawing.Point(32, 112);
			this.btFiltro.Name = "btFiltro";
			this.btFiltro.Size = new System.Drawing.Size(48, 24);
			this.btFiltro.TabIndex = 10;
			this.btFiltro.Text = "Filtrar";
			this.btFiltro.Click += new System.EventHandler(this.btFiltro_Click);
			// 
			// cmbCliente
			// 
			this.cmbCliente.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbCliente.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbCliente.DataSource = this.cdvCliente;
			ultraGridColumn1.Header.VisiblePosition = 1;
			ultraGridColumn1.Width = 137;
			ultraGridColumn2.Header.VisiblePosition = 3;
			ultraGridColumn2.Hidden = true;
			ultraGridColumn3.Header.VisiblePosition = 0;
			ultraGridColumn3.Width = 322;
			ultraGridColumn4.Header.VisiblePosition = 4;
			ultraGridColumn4.Hidden = true;
			ultraGridColumn5.Header.VisiblePosition = 5;
			ultraGridColumn5.Hidden = true;
			ultraGridColumn6.Header.VisiblePosition = 2;
			ultraGridColumn6.Hidden = true;
			ultraGridColumn7.Header.VisiblePosition = 6;
			ultraGridColumn7.Hidden = true;
			ultraGridColumn8.Header.VisiblePosition = 7;
			ultraGridColumn8.Hidden = true;
			ultraGridColumn9.Header.VisiblePosition = 8;
			ultraGridColumn9.Hidden = true;
			ultraGridBand1.Columns.AddRange(new object[] {
																										 ultraGridColumn1,
																										 ultraGridColumn2,
																										 ultraGridColumn3,
																										 ultraGridColumn4,
																										 ultraGridColumn5,
																										 ultraGridColumn6,
																										 ultraGridColumn7,
																										 ultraGridColumn8,
																										 ultraGridColumn9});
			this.cmbCliente.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.cmbCliente.DisplayMember = "Nombre";
			this.cmbCliente.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbCliente.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbCliente.Location = new System.Drawing.Point(144, 8);
			this.cmbCliente.Name = "cmbCliente";
			this.cmbCliente.Size = new System.Drawing.Size(400, 21);
			this.cmbCliente.TabIndex = 8;
			this.cmbCliente.ValueMember = "idCliente";
			// 
			// cdvCliente
			// 
			this.cdvCliente.BindingContextCtrl = this;
			this.cdvCliente.DataSet = this.cdsCliente;
			this.cdvCliente.RowFilter = "proveedor=1  or ambos=1";
			this.cdvCliente.Sort = "Nombre";
			this.cdvCliente.TableName = "";
			this.cdvCliente.TableViewName = "Cliente";
			// 
			// cdsCliente
			// 
			this.cdsCliente.BindingContextCtrl = this;
			this.cdsCliente.CaseSensitive = false;
			this.cdsCliente.DataLibrary = "LibFacturacion";
			this.cdsCliente.DataLibraryUrl = "";
			this.cdsCliente.DataSetDef = "dsClientes";
			this.cdsCliente.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsCliente.Name = "cdsCliente";
			this.cdsCliente.SchemaClassName = "LibFacturacion.DataClass";
			this.cdsCliente.SchemaDef = null;
			this.cdsCliente.BeforeFill += new C1.Data.FillEventHandler(this.cdsCliente_BeforeFill);
			// 
			// cdvGuiaRemisionG
			// 
			this.cdvGuiaRemisionG.BindingContextCtrl = this;
			this.cdvGuiaRemisionG.DataSet = null;
			this.cdvGuiaRemisionG.TableName = "";
			this.cdvGuiaRemisionG.TableViewName = "GuiaRemisionG";
			// 
			// chkGuia
			// 
			this.chkGuia.EditValue = true;
			this.chkGuia.Location = new System.Drawing.Point(32, 56);
			this.chkGuia.Name = "chkGuia";
			// 
			// chkGuia.Properties
			// 
			this.chkGuia.Properties.Caption = "Toda Guia";
			this.chkGuia.Size = new System.Drawing.Size(88, 19);
			this.chkGuia.TabIndex = 72;
			this.chkGuia.CheckedChanged += new System.EventHandler(this.chkGuia_CheckedChanged);
			// 
			// chkFecha
			// 
			this.chkFecha.EditValue = true;
			this.chkFecha.Location = new System.Drawing.Point(32, 32);
			this.chkFecha.Name = "chkFecha";
			// 
			// chkFecha.Properties
			// 
			this.chkFecha.Properties.Caption = "Toda Fecha";
			this.chkFecha.Size = new System.Drawing.Size(88, 19);
			this.chkFecha.TabIndex = 70;
			this.chkFecha.CheckedChanged += new System.EventHandler(this.chkFecha_CheckedChanged);
			// 
			// chkCliente
			// 
			this.chkCliente.Location = new System.Drawing.Point(32, 8);
			this.chkCliente.Name = "chkCliente";
			// 
			// chkCliente.Properties
			// 
			this.chkCliente.Properties.Caption = "Todo Cliente";
			this.chkCliente.Size = new System.Drawing.Size(96, 19);
			this.chkCliente.TabIndex = 68;
			this.chkCliente.Validated += new System.EventHandler(this.chkCliente_Validated);
			this.chkCliente.CheckedChanged += new System.EventHandler(this.chkCliente_CheckedChanged);
			// 
			// dtFechaHasta
			// 
			this.dtFechaHasta.EditValue = new System.DateTime(2004, 9, 20, 0, 0, 0, 0);
			this.dtFechaHasta.Location = new System.Drawing.Point(280, 32);
			this.dtFechaHasta.Name = "dtFechaHasta";
			// 
			// dtFechaHasta.Properties
			// 
			this.dtFechaHasta.Properties.Buttons.AddRange(new DevExpress.XtraEditors.Controls.EditorButton[] {
																																																				 new DevExpress.XtraEditors.Controls.EditorButton(DevExpress.XtraEditors.Controls.ButtonPredefines.Combo)});
			this.dtFechaHasta.Properties.DisplayFormat.FormatString = "dd/MMM/yyyy";
			this.dtFechaHasta.Properties.DisplayFormat.FormatType = DevExpress.Utils.FormatType.Custom;
			this.dtFechaHasta.Properties.EditFormat.FormatString = "dd/MM/yyyy";
			this.dtFechaHasta.Properties.EditFormat.FormatType = DevExpress.Utils.FormatType.Custom;
			this.dtFechaHasta.Size = new System.Drawing.Size(96, 19);
			this.dtFechaHasta.TabIndex = 75;
			// 
			// dtFechaDesde
			// 
			this.dtFechaDesde.EditValue = new System.DateTime(2004, 9, 20, 0, 0, 0, 0);
			this.dtFechaDesde.Location = new System.Drawing.Point(144, 32);
			this.dtFechaDesde.Name = "dtFechaDesde";
			// 
			// dtFechaDesde.Properties
			// 
			this.dtFechaDesde.Properties.Buttons.AddRange(new DevExpress.XtraEditors.Controls.EditorButton[] {
																																																				 new DevExpress.XtraEditors.Controls.EditorButton(DevExpress.XtraEditors.Controls.ButtonPredefines.Combo)});
			this.dtFechaDesde.Properties.DisplayFormat.FormatString = "dd/MMM/yyyy";
			this.dtFechaDesde.Properties.DisplayFormat.FormatType = DevExpress.Utils.FormatType.Custom;
			this.dtFechaDesde.Properties.EditFormat.FormatString = "dd/MM/yyyy";
			this.dtFechaDesde.Properties.EditFormat.FormatType = DevExpress.Utils.FormatType.Custom;
			this.dtFechaDesde.Size = new System.Drawing.Size(96, 19);
			this.dtFechaDesde.TabIndex = 76;
			// 
			// chkRuta
			// 
			this.chkRuta.EditValue = true;
			this.chkRuta.Location = new System.Drawing.Point(32, 80);
			this.chkRuta.Name = "chkRuta";
			// 
			// chkRuta.Properties
			// 
			this.chkRuta.Properties.Caption = "Toda Ruta";
			this.chkRuta.Size = new System.Drawing.Size(112, 19);
			this.chkRuta.TabIndex = 77;
			this.chkRuta.CheckedChanged += new System.EventHandler(this.chkRuta_CheckedChanged);
			// 
			// cmbPuntoPartida
			// 
			appearance2.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbPuntoPartida.Appearance = appearance2;
			this.cmbPuntoPartida.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbPuntoPartida.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbPuntoPartida.DataSource = this.cdsGuiaPunto;
			ultraGridColumn10.Header.VisiblePosition = 0;
			ultraGridColumn10.Hidden = true;
			ultraGridColumn11.Header.VisiblePosition = 1;
			ultraGridBand2.Columns.AddRange(new object[] {
																										 ultraGridColumn10,
																										 ultraGridColumn11});
			this.cmbPuntoPartida.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			this.cmbPuntoPartida.DisplayMember = "Nombre";
			this.cmbPuntoPartida.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbPuntoPartida.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbPuntoPartida.Enabled = false;
			this.cmbPuntoPartida.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbPuntoPartida.Location = new System.Drawing.Point(144, 80);
			this.cmbPuntoPartida.Name = "cmbPuntoPartida";
			this.cmbPuntoPartida.Size = new System.Drawing.Size(184, 21);
			this.cmbPuntoPartida.TabIndex = 79;
			this.cmbPuntoPartida.ValueMember = "idGuiaRemisionPunto";
			// 
			// cdsGuiaPunto
			// 
			this.cdsGuiaPunto.BindingContextCtrl = this;
			this.cdsGuiaPunto.DataLibrary = "LibFacturacion";
			this.cdsGuiaPunto.DataLibraryUrl = "";
			this.cdsGuiaPunto.DataSetDef = "dsGuiaRemisionPunto";
			this.cdsGuiaPunto.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsGuiaPunto.Name = "cdsGuiaPunto";
			this.cdsGuiaPunto.SchemaClassName = "LibFacturacion.DataClass";
			this.cdsGuiaPunto.SchemaDef = null;
			// 
			// cmbPuntoLlegada
			// 
			appearance3.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbPuntoLlegada.Appearance = appearance3;
			this.cmbPuntoLlegada.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbPuntoLlegada.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbPuntoLlegada.DataSource = this.cdsGuiaPunto;
			ultraGridColumn12.Header.VisiblePosition = 0;
			ultraGridColumn12.Hidden = true;
			ultraGridColumn13.Header.VisiblePosition = 1;
			ultraGridBand3.Columns.AddRange(new object[] {
																										 ultraGridColumn12,
																										 ultraGridColumn13});
			this.cmbPuntoLlegada.DisplayLayout.BandsSerializer.Add(ultraGridBand3);
			this.cmbPuntoLlegada.DisplayMember = "Nombre";
			this.cmbPuntoLlegada.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbPuntoLlegada.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbPuntoLlegada.Enabled = false;
			this.cmbPuntoLlegada.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbPuntoLlegada.Location = new System.Drawing.Point(352, 80);
			this.cmbPuntoLlegada.Name = "cmbPuntoLlegada";
			this.cmbPuntoLlegada.Size = new System.Drawing.Size(192, 21);
			this.cmbPuntoLlegada.TabIndex = 78;
			this.cmbPuntoLlegada.ValueMember = "idGuiaRemisionPunto";
			// 
			// cdvGuiaPunto
			// 
			this.cdvGuiaPunto.BindingContextCtrl = this;
			this.cdvGuiaPunto.DataSet = this.cdsGuiaPunto;
			this.cdvGuiaPunto.Sort = "nombre";
			this.cdvGuiaPunto.TableName = "";
			this.cdvGuiaPunto.TableViewName = "GuiaRemisionPunto";
			// 
			// txtGuia1
			// 
			appearance4.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtGuia1.Appearance = appearance4;
			this.txtGuia1.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtGuia1.Enabled = false;
			this.txtGuia1.FormatString = "";
			this.txtGuia1.Location = new System.Drawing.Point(144, 56);
			this.txtGuia1.MaskInput = "nnnnnnnnn";
			this.txtGuia1.Name = "txtGuia1";
			this.txtGuia1.PromptChar = ' ';
			this.txtGuia1.Size = new System.Drawing.Size(96, 21);
			this.txtGuia1.TabIndex = 120;
			// 
			// txtGuia2
			// 
			appearance5.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtGuia2.Appearance = appearance5;
			this.txtGuia2.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtGuia2.Enabled = false;
			this.txtGuia2.FormatString = "";
			this.txtGuia2.Location = new System.Drawing.Point(280, 56);
			this.txtGuia2.MaskInput = "nnnnnnnnn";
			this.txtGuia2.Name = "txtGuia2";
			this.txtGuia2.PromptChar = ' ';
			this.txtGuia2.Size = new System.Drawing.Size(96, 21);
			this.txtGuia2.TabIndex = 122;
			// 
			// cdsGuiaRemisionG
			// 
			this.cdsGuiaRemisionG.BindingContextCtrl = this;
			this.cdsGuiaRemisionG.DataLibrary = "LibFacturacion";
			this.cdsGuiaRemisionG.DataLibraryUrl = "";
			this.cdsGuiaRemisionG.DataSetDef = "dsGuiaRemisionG";
			this.cdsGuiaRemisionG.EnforceConstraints = false;
			this.cdsGuiaRemisionG.FillOnRequest = false;
			this.cdsGuiaRemisionG.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsGuiaRemisionG.Name = "cdsGuiaRemisionG";
			this.cdsGuiaRemisionG.SchemaClassName = "LibFacturacion.DataClass";
			this.cdsGuiaRemisionG.SchemaDef = null;
			// 
			// ultraGrid1
			// 
			this.ultraGrid1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.ultraGrid1.Cursor = System.Windows.Forms.Cursors.Default;
			this.ultraGrid1.DataMember = "GuiaRemisionG";
			this.ultraGrid1.DataSource = this.cdsGuiaRemisionPadre;
			appearance6.BackColor = System.Drawing.Color.White;
			this.ultraGrid1.DisplayLayout.Appearance = appearance6;
			this.ultraGrid1.DisplayLayout.AutoFitColumns = true;
			ultraGridBand4.AddButtonCaption = "_GuiaRemisionDetalle";
			ultraGridColumn14.Header.VisiblePosition = 0;
			ultraGridColumn14.Hidden = true;
			ultraGridColumn14.Width = 82;
			ultraGridColumn15.Header.VisiblePosition = 1;
			ultraGridColumn15.Hidden = true;
			ultraGridColumn15.Width = 16;
			ultraGridColumn16.Header.VisiblePosition = 11;
			ultraGridColumn16.Hidden = true;
			ultraGridColumn16.Width = 16;
			ultraGridColumn17.Header.VisiblePosition = 13;
			ultraGridColumn17.Hidden = true;
			ultraGridColumn17.Width = 16;
			ultraGridColumn18.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
			ultraGridColumn18.Header.Caption = "Guia";
			ultraGridColumn18.Header.VisiblePosition = 5;
			ultraGridColumn18.Width = 47;
			ultraGridColumn19.Header.VisiblePosition = 14;
			ultraGridColumn19.Hidden = true;
			ultraGridColumn19.Width = 16;
			ultraGridColumn20.Header.VisiblePosition = 15;
			ultraGridColumn20.Hidden = true;
			ultraGridColumn20.Width = 16;
			ultraGridColumn21.Header.VisiblePosition = 16;
			ultraGridColumn21.Hidden = true;
			ultraGridColumn21.Width = 16;
			ultraGridColumn22.Header.VisiblePosition = 17;
			ultraGridColumn22.Hidden = true;
			ultraGridColumn22.Width = 16;
			ultraGridColumn23.Header.VisiblePosition = 18;
			ultraGridColumn23.Hidden = true;
			ultraGridColumn23.Width = 16;
			ultraGridColumn24.Header.VisiblePosition = 19;
			ultraGridColumn24.Hidden = true;
			ultraGridColumn24.Width = 16;
			ultraGridColumn25.Header.VisiblePosition = 20;
			ultraGridColumn25.Hidden = true;
			ultraGridColumn25.Width = 16;
			ultraGridColumn26.Header.VisiblePosition = 21;
			ultraGridColumn26.Hidden = true;
			ultraGridColumn26.Width = 16;
			ultraGridColumn27.Header.VisiblePosition = 22;
			ultraGridColumn27.Hidden = true;
			ultraGridColumn27.Width = 16;
			ultraGridColumn28.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
			ultraGridColumn28.Header.Caption = "Emision";
			ultraGridColumn28.Header.VisiblePosition = 6;
			ultraGridColumn28.Width = 61;
			ultraGridColumn29.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
			ultraGridColumn29.Header.Caption = "Partida";
			ultraGridColumn29.Header.VisiblePosition = 8;
			ultraGridColumn29.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.Edit;
			ultraGridColumn29.Width = 68;
			ultraGridColumn30.Header.VisiblePosition = 40;
			ultraGridColumn30.Hidden = true;
			ultraGridColumn30.Width = 92;
			ultraGridColumn31.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
			ultraGridColumn31.Header.Caption = "Llegada";
			ultraGridColumn31.Header.VisiblePosition = 9;
			ultraGridColumn31.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.Edit;
			ultraGridColumn31.Width = 78;
			ultraGridColumn32.Header.Caption = "Chofer";
			ultraGridColumn32.Header.VisiblePosition = 7;
			ultraGridColumn32.Width = 115;
			ultraGridColumn33.Header.VisiblePosition = 23;
			ultraGridColumn33.Hidden = true;
			ultraGridColumn33.Width = 16;
			ultraGridColumn34.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
			ultraGridColumn34.Header.Caption = "Observaciones";
			ultraGridColumn34.Header.VisiblePosition = 10;
			ultraGridColumn34.Width = 260;
			ultraGridColumn35.Header.VisiblePosition = 24;
			ultraGridColumn35.Hidden = true;
			ultraGridColumn35.Width = 12;
			ultraGridColumn36.Header.VisiblePosition = 25;
			ultraGridColumn36.Hidden = true;
			ultraGridColumn36.Width = 12;
			ultraGridColumn37.Header.VisiblePosition = 26;
			ultraGridColumn37.Hidden = true;
			ultraGridColumn37.Width = 16;
			ultraGridColumn38.Header.VisiblePosition = 27;
			ultraGridColumn38.Hidden = true;
			ultraGridColumn38.Width = 14;
			ultraGridColumn39.Header.VisiblePosition = 28;
			ultraGridColumn39.Hidden = true;
			ultraGridColumn39.Width = 16;
			ultraGridColumn40.Header.VisiblePosition = 29;
			ultraGridColumn40.Hidden = true;
			ultraGridColumn40.Width = 16;
			ultraGridColumn41.Header.VisiblePosition = 30;
			ultraGridColumn41.Hidden = true;
			ultraGridColumn41.Width = 9;
			ultraGridColumn42.Header.VisiblePosition = 31;
			ultraGridColumn42.Hidden = true;
			ultraGridColumn42.Width = 9;
			ultraGridColumn43.Header.VisiblePosition = 32;
			ultraGridColumn43.Hidden = true;
			ultraGridColumn43.Width = 9;
			ultraGridColumn44.Header.VisiblePosition = 33;
			ultraGridColumn44.Hidden = true;
			ultraGridColumn44.Width = 9;
			ultraGridColumn45.Header.Caption = "Cliente";
			ultraGridColumn45.Header.VisiblePosition = 12;
			ultraGridColumn45.Hidden = true;
			ultraGridColumn45.Width = 112;
			ultraGridColumn46.Header.VisiblePosition = 34;
			ultraGridColumn46.Hidden = true;
			ultraGridColumn46.Width = 15;
			ultraGridColumn47.Header.VisiblePosition = 35;
			ultraGridColumn47.Hidden = true;
			ultraGridColumn47.Width = 16;
			ultraGridColumn48.Header.VisiblePosition = 36;
			ultraGridColumn48.Hidden = true;
			ultraGridColumn48.Width = 17;
			ultraGridColumn49.Header.VisiblePosition = 37;
			ultraGridColumn49.Hidden = true;
			ultraGridColumn49.Width = 16;
			ultraGridColumn50.Header.VisiblePosition = 38;
			ultraGridColumn50.Hidden = true;
			ultraGridColumn50.Width = 16;
			ultraGridColumn51.Header.Caption = "Fact";
			ultraGridColumn51.Header.VisiblePosition = 3;
			ultraGridColumn51.Hidden = true;
			ultraGridColumn51.Width = 36;
			ultraGridColumn52.Header.Caption = "Cobro";
			ultraGridColumn52.Header.VisiblePosition = 4;
			ultraGridColumn52.Width = 52;
			ultraGridColumn53.Header.VisiblePosition = 39;
			ultraGridColumn53.Hidden = true;
			ultraGridColumn53.Width = 12;
			ultraGridColumn54.Header.VisiblePosition = 41;
			ultraGridColumn54.Hidden = true;
			ultraGridColumn54.Width = 11;
			ultraGridColumn55.Header.VisiblePosition = 2;
			ultraGridColumn55.Hidden = true;
			ultraGridColumn55.Width = 80;
			ultraGridColumn56.Header.VisiblePosition = 42;
			ultraGridColumn56.Hidden = true;
			ultraGridColumn56.Width = 77;
			ultraGridColumn57.Header.VisiblePosition = 43;
			ultraGridColumn57.Hidden = true;
			ultraGridColumn57.Width = 80;
			ultraGridBand4.Columns.AddRange(new object[] {
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
																										 ultraGridColumn57});
			this.ultraGrid1.DisplayLayout.BandsSerializer.Add(ultraGridBand4);
			this.ultraGrid1.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.TemplateOnBottom;
			this.ultraGrid1.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.True;
			this.ultraGrid1.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance7.BackColor = System.Drawing.Color.Transparent;
			this.ultraGrid1.DisplayLayout.Override.CardAreaAppearance = appearance7;
			appearance8.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance8.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance8.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance8.FontData.BoldAsString = "True";
			appearance8.FontData.Name = "Arial";
			appearance8.FontData.SizeInPoints = 10F;
			appearance8.ForeColor = System.Drawing.Color.White;
			appearance8.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.ultraGrid1.DisplayLayout.Override.HeaderAppearance = appearance8;
			this.ultraGrid1.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti;
			appearance9.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance9.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance9.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.RowSelectorAppearance = appearance9;
			appearance10.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance10.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance10.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.SelectedRowAppearance = appearance10;
			this.ultraGrid1.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.ultraGrid1.Location = new System.Drawing.Point(12, 190);
			this.ultraGrid1.Name = "ultraGrid1";
			this.ultraGrid1.Size = new System.Drawing.Size(702, 152);
			this.ultraGrid1.TabIndex = 125;
			// 
			// cdsGuiaRemisionPadre
			// 
			this.cdsGuiaRemisionPadre.BindingContextCtrl = this;
			this.cdsGuiaRemisionPadre.DataLibrary = "LibFacturacion";
			this.cdsGuiaRemisionPadre.DataLibraryUrl = "";
			this.cdsGuiaRemisionPadre.DataSetDef = "dsGuiaRemisionPadreC";
			this.cdsGuiaRemisionPadre.EnforceConstraints = false;
			this.cdsGuiaRemisionPadre.FillOnRequest = false;
			this.cdsGuiaRemisionPadre.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsGuiaRemisionPadre.Name = "cdsGuiaRemisionPadre";
			this.cdsGuiaRemisionPadre.SchemaClassName = "LibFacturacion.DataClass";
			this.cdsGuiaRemisionPadre.SchemaDef = null;
			// 
			// cmbLlegada
			// 
			this.cmbLlegada.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbLlegada.DataMember = "GuiaRemisionPunto";
			this.cmbLlegada.DataSource = this.cdsGuiaPunto;
			ultraGridBand5.AddButtonCaption = "CompraComprobante";
			ultraGridColumn58.Header.VisiblePosition = 0;
			ultraGridColumn58.Hidden = true;
			ultraGridColumn59.Header.VisiblePosition = 1;
			ultraGridBand5.Columns.AddRange(new object[] {
																										 ultraGridColumn58,
																										 ultraGridColumn59});
			this.cmbLlegada.DisplayLayout.BandsSerializer.Add(ultraGridBand5);
			this.cmbLlegada.DisplayMember = "Nombre";
			this.cmbLlegada.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbLlegada.Location = new System.Drawing.Point(376, 256);
			this.cmbLlegada.Name = "cmbLlegada";
			this.cmbLlegada.Size = new System.Drawing.Size(136, 72);
			this.cmbLlegada.TabIndex = 126;
			this.cmbLlegada.ValueMember = "idGuiaRemisionPunto";
			this.cmbLlegada.Visible = false;
			// 
			// cdvGuiaRemisionPadre
			// 
			this.cdvGuiaRemisionPadre.BindingContextCtrl = this;
			this.cdvGuiaRemisionPadre.DataSet = this.cdsGuiaRemisionPadre;
			this.cdvGuiaRemisionPadre.RowFilter = "tipo=2";
			this.cdvGuiaRemisionPadre.TableName = "";
			this.cdvGuiaRemisionPadre.TableViewName = "GuiaRemisionPadre";
			// 
			// barraDato1
			// 
			this.barraDato1.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.barraDato1.BarraMovimiento = true;
			this.barraDato1.DataMember = "_GuiaRemisionPadreC";
			this.barraDato1.DataNombreId = "idGuiaRemisionPadre";
			this.barraDato1.DataOrden = "";
			this.barraDato1.DataSource = this.cdsGuiaRemisionPadre;
			this.barraDato1.DataTabla = "GuiaRemisionPadre";
			this.barraDato1.DataTablaHija = "GuiaRemisionG";
			this.barraDato1.Location = new System.Drawing.Point(192, 360);
			this.barraDato1.Name = "barraDato1";
			this.barraDato1.Size = new System.Drawing.Size(304, 24);
			this.barraDato1.TabIndex = 127;
			this.barraDato1.VisibleBorrar = true;
			this.barraDato1.VisibleBuscar = true;
			this.barraDato1.VisibleEditar = true;
			this.barraDato1.VisibleImprimir = true;
			this.barraDato1.VisibleNuevo = true;
			this.barraDato1.Borrar += new System.EventHandler(this.barraDato1_Borrar);
			this.barraDato1.Refresca += new System.EventHandler(this.barraDato1_Refresca);
			this.barraDato1.Crear += new System.EventHandler(this.barraDato1_Crear);
			this.barraDato1.Deshacer += new System.EventHandler(this.barraDato1_Deshacer);
			this.barraDato1.Editar += new System.EventHandler(this.barraDato1_Editar);
			this.barraDato1.Grabar += new System.EventHandler(this.barraDato1_Grabar);
			// 
			// txtidGuiaRemisionPadre
			// 
			appearance11.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtidGuiaRemisionPadre.Appearance = appearance11;
			this.txtidGuiaRemisionPadre.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsGuiaRemisionPadre, "_GuiaRemisionPadreC.idGuiaRemisionPadre"));
			this.txtidGuiaRemisionPadre.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtidGuiaRemisionPadre.Enabled = false;
			this.txtidGuiaRemisionPadre.FormatString = "";
			this.txtidGuiaRemisionPadre.Location = new System.Drawing.Point(27, 142);
			this.txtidGuiaRemisionPadre.MaskInput = "nnnnnnnnn";
			this.txtidGuiaRemisionPadre.Name = "txtidGuiaRemisionPadre";
			this.txtidGuiaRemisionPadre.PromptChar = ' ';
			this.txtidGuiaRemisionPadre.Size = new System.Drawing.Size(56, 21);
			this.txtidGuiaRemisionPadre.TabIndex = 128;
			this.txtidGuiaRemisionPadre.ValueChanged += new System.EventHandler(this.ultraNumericEditor2_ValueChanged);
			// 
			// txtNumero
			// 
			appearance12.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtNumero.Appearance = appearance12;
			this.txtNumero.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsGuiaRemisionPadre, "_GuiaRemisionPadreC.Numero"));
			this.txtNumero.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNumero.Enabled = false;
			this.txtNumero.FormatString = "";
			this.txtNumero.Location = new System.Drawing.Point(192, 133);
			this.txtNumero.MaskInput = "nnnnnnnnn";
			this.txtNumero.Name = "txtNumero";
			this.txtNumero.PromptChar = ' ';
			this.txtNumero.Size = new System.Drawing.Size(96, 21);
			this.txtNumero.TabIndex = 129;
			this.txtNumero.ValueChanged += new System.EventHandler(this.ultraNumericEditor3_ValueChanged);
			// 
			// btnAsiento
			// 
			this.btnAsiento.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			appearance13.Image = 5;
			this.btnAsiento.Appearance = appearance13;
			this.btnAsiento.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.btnAsiento.Location = new System.Drawing.Point(504, 360);
			this.btnAsiento.Name = "btnAsiento";
			this.btnAsiento.Size = new System.Drawing.Size(56, 24);
			this.btnAsiento.TabIndex = 132;
			this.btnAsiento.Text = "Compra";
			this.btnAsiento.Click += new System.EventHandler(this.btnAsiento_Click);
			// 
			// cmbCliente1
			// 
			this.cmbCliente1.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbCliente1.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbCliente1.DataSource = this.cdvCliente;
			ultraGridColumn60.Header.VisiblePosition = 1;
			ultraGridColumn60.Width = 137;
			ultraGridColumn61.Header.VisiblePosition = 3;
			ultraGridColumn61.Hidden = true;
			ultraGridColumn62.Header.VisiblePosition = 0;
			ultraGridColumn62.Width = 322;
			ultraGridColumn63.Header.VisiblePosition = 4;
			ultraGridColumn63.Hidden = true;
			ultraGridColumn64.Header.VisiblePosition = 5;
			ultraGridColumn64.Hidden = true;
			ultraGridColumn65.Header.VisiblePosition = 2;
			ultraGridColumn65.Hidden = true;
			ultraGridColumn66.Header.VisiblePosition = 6;
			ultraGridColumn66.Hidden = true;
			ultraGridColumn67.Header.VisiblePosition = 7;
			ultraGridColumn67.Hidden = true;
			ultraGridColumn68.Header.VisiblePosition = 8;
			ultraGridColumn68.Hidden = true;
			ultraGridBand6.Columns.AddRange(new object[] {
																										 ultraGridColumn60,
																										 ultraGridColumn61,
																										 ultraGridColumn62,
																										 ultraGridColumn63,
																										 ultraGridColumn64,
																										 ultraGridColumn65,
																										 ultraGridColumn66,
																										 ultraGridColumn67,
																										 ultraGridColumn68});
			this.cmbCliente1.DisplayLayout.BandsSerializer.Add(ultraGridBand6);
			this.cmbCliente1.DisplayMember = "Nombre";
			this.cmbCliente1.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbCliente1.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbCliente1.Location = new System.Drawing.Point(136, 160);
			this.cmbCliente1.Name = "cmbCliente1";
			this.cmbCliente1.Size = new System.Drawing.Size(400, 21);
			this.cmbCliente1.TabIndex = 133;
			this.cmbCliente1.ValueMember = "idCliente";
			// 
			// label18
			// 
			this.label18.AutoSize = true;
			this.label18.Location = new System.Drawing.Point(141, 135);
			this.label18.Name = "label18";
			this.label18.Size = new System.Drawing.Size(44, 16);
			this.label18.TabIndex = 134;
			this.label18.Text = "Numero";
			this.label18.Click += new System.EventHandler(this.label18_Click);
			// 
			// label1
			// 
			this.label1.AutoSize = true;
			this.label1.Location = new System.Drawing.Point(353, 134);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(36, 16);
			this.label1.TabIndex = 135;
			this.label1.Text = "Fecha";
			// 
			// dtFecha
			// 
			this.dtFecha.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsGuiaRemisionPadre, "_GuiaRemisionPadreC.Fecha"));
			dateButton1.Caption = "Today";
			this.dtFecha.DateButtons.Add(dateButton1);
			this.dtFecha.Format = "dd/MMM/yyyy";
			this.dtFecha.Location = new System.Drawing.Point(450, 132);
			this.dtFecha.Name = "dtFecha";
			this.dtFecha.NonAutoSizeHeight = 23;
			this.dtFecha.Size = new System.Drawing.Size(88, 21);
			this.dtFecha.TabIndex = 139;
			// 
			// errorProvider
			// 
			this.errorProvider.ContainerControl = this;
			this.errorProvider.DataMember = "";
			// 
			// cmbPartida
			// 
			this.cmbPartida.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbPartida.DataMember = "GuiaRemisionPunto";
			this.cmbPartida.DataSource = this.cdsGuiaPunto;
			ultraGridBand7.AddButtonCaption = "CompraComprobante";
			ultraGridColumn69.Header.VisiblePosition = 0;
			ultraGridColumn69.Hidden = true;
			ultraGridColumn70.Header.VisiblePosition = 1;
			ultraGridBand7.Columns.AddRange(new object[] {
																										 ultraGridColumn69,
																										 ultraGridColumn70});
			this.cmbPartida.DisplayLayout.BandsSerializer.Add(ultraGridBand7);
			this.cmbPartida.DisplayMember = "Nombre";
			this.cmbPartida.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbPartida.Location = new System.Drawing.Point(204, 239);
			this.cmbPartida.Name = "cmbPartida";
			this.cmbPartida.Size = new System.Drawing.Size(136, 72);
			this.cmbPartida.TabIndex = 140;
			this.cmbPartida.ValueMember = "idGuiaRemisionPunto";
			this.cmbPartida.Visible = false;
			// 
			// GuiaRemisionFactCompra
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(720, 389);
			this.Controls.Add(this.cmbPartida);
			this.Controls.Add(this.dtFecha);
			this.Controls.Add(this.label1);
			this.Controls.Add(this.label18);
			this.Controls.Add(this.cmbCliente1);
			this.Controls.Add(this.btnAsiento);
			this.Controls.Add(this.txtNumero);
			this.Controls.Add(this.barraDato1);
			this.Controls.Add(this.cmbLlegada);
			this.Controls.Add(this.ultraGrid1);
			this.Controls.Add(this.txtGuia2);
			this.Controls.Add(this.txtGuia1);
			this.Controls.Add(this.cmbPuntoPartida);
			this.Controls.Add(this.cmbPuntoLlegada);
			this.Controls.Add(this.chkRuta);
			this.Controls.Add(this.dtFechaDesde);
			this.Controls.Add(this.dtFechaHasta);
			this.Controls.Add(this.chkGuia);
			this.Controls.Add(this.chkFecha);
			this.Controls.Add(this.chkCliente);
			this.Controls.Add(this.btFiltro);
			this.Controls.Add(this.cmbCliente);
			this.Controls.Add(this.txtidGuiaRemisionPadre);
			this.Name = "GuiaRemisionFactCompra";
			this.Text = "Guia Remision Compras";
			this.Load += new System.EventHandler(this.GuiaRemisionFact_Load);
			((System.ComponentModel.ISupportInitialize)(this.cmbCliente)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvCliente)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsCliente)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvGuiaRemisionG)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.chkGuia.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.chkFecha.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.chkCliente.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFechaHasta.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFechaDesde.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.chkRuta.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbPuntoPartida)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsGuiaPunto)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbPuntoLlegada)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvGuiaPunto)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtGuia1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtGuia2)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsGuiaRemisionG)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsGuiaRemisionPadre)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbLlegada)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvGuiaRemisionPadre)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtidGuiaRemisionPadre)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNumero)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCliente1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFecha)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbPartida)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void label1_Click(object sender, System.EventArgs e)
		{
		
		}

		private void cdsCliente_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			this.cdsCliente.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);
		}

		private void GuiaRemisionFact_Load(object sender, System.EventArgs e)
		{
			txtidGuiaRemisionPadre.Width = 0;
			dtFechaDesde.Enabled =false;
			dtFechaHasta.Enabled =false; 
			dtFecha.Value = DateTime.Today;
		  this.barraDato1.stFiltroExt  = " And tipo = 2";
			this.barraDato1.HabilitaControles(false);  
		  this.barraDato1.PosUltima(); 
		}

		private void cdsGuiaRemisionG_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			this.cdsGuiaRemisionG.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);
		}

		private bool VerificaFiltro()
      {
			bool bOk = true;
			errorProvider.SetError(cmbCliente, "");
			errorProvider.SetError(cmbPartida, "");
			errorProvider.SetError(cmbPuntoLlegada, "");

			if (chkCliente.Checked==false)
				if (cmbCliente.Text.Trim().Length ==0)
				{
					errorProvider.SetError(cmbCliente, "Ingrese Cliente");
					bOk = false;
				}

			if (chkRuta.Checked==false)
{
				if (cmbPartida.Text.Trim().Length ==0)
				{
					errorProvider.SetError(cmbPartida, "Ingrese Partida");
					bOk = false;
				}
				if (cmbPuntoLlegada.Text.Trim().Length ==0)
				{
					errorProvider.SetError(cmbPuntoLlegada, "Ingrese Punto de Llegada");
					bOk = false;
				}
			}


			if (chkGuia.Checked==false)
			{
				if (txtGuia1.Text.Trim().Length ==0)
				{
					errorProvider.SetError(txtGuia1, "Ingrese Guia");
					bOk = false;
				}
			if (txtGuia2.Text.Trim().Length ==0)
			{
				errorProvider.SetError(txtGuia2, "Ingrese Guia");
				bOk = false;
			}
			}

			return bOk;
		}

		private void btFiltro_Click(object sender, System.EventArgs e)
		{
			if (!barraDato1.bEditar) return;
			if (!VerificaFiltro()) return;
			//aqui cobrado
			string  stFiltro = " cobrado=0 and borrar=0 and contratado=1 and idguiaremisionpadre is null ";
			
			if (chkCliente.Checked==false)
					stFiltro += " and idcliente=" + cmbCliente.Value.ToString();

			if (chkRuta.Checked==false)
				stFiltro += " and idpuntopartida=" + cmbPuntoPartida.Value.ToString() + " and idpuntollegada=" + cmbPuntoLlegada.Value.ToString();						

			if (chkGuia.Checked==false)
				stFiltro += " and nroguia>=" + txtGuia1.Value.ToString() + " and nroguia<=" + txtGuia2.Value.ToString() ;
		
			DateTime dtFechaIni = (DateTime) dtFechaDesde.EditValue;
			DateTime dtFechaFin = (DateTime) dtFechaHasta.EditValue;
			dtFechaFin=dtFechaFin.AddDays(1); 
			if (chkFecha.Checked==false)
				stFiltro += " and fEmision2 >='" + dtFechaIni.ToString("yyyyMMdd") + "' and fEmision2 <'" + dtFechaFin.ToString("yyyyMMdd") +"'";
			
//			C1.Data.FilterCondition fcPadre = new 
//				C1.Data.FilterCondition(this.cdsGuiaRemisionG, "GuiaRemisionG", stFiltro);
//			C1.Data.FilterConditions fcTotal = new C1.Data.FilterConditions();
//			fcTotal.Add(fcPadre);
//
//			C1.Data.FilterCondition fcHija = new 
//				C1.Data.FilterCondition(this.cdsGuiaRemisionG , "GuiaRemisionDetalle", "idGuiaRemisionG = 0");
//			fcTotal.Add(fcHija);
//			
//			this.cdsGuiaRemisionG.Fill(fcTotal, false);
			C1.Data.FilterCondition fcPadre = new 
				C1.Data.FilterCondition(this.cdsGuiaRemisionPadre, "GuiaRemisionPadre", "idGuiaRemisionPadre = 0");
			C1.Data.FilterConditions fcTotal = new C1.Data.FilterConditions();
			fcTotal.Add(fcPadre);

			C1.Data.FilterCondition fcHija = new 
				C1.Data.FilterCondition(this.cdsGuiaRemisionPadre , "GuiaRemisionG", stFiltro );
			fcTotal.Add(fcHija);
			
			this.cdsGuiaRemisionPadre.Fill(fcTotal, false);
	
		}

		private void chkCliente_CheckedChanged(object sender, System.EventArgs e)
		{
					this.cmbCliente.Enabled =  !this.chkCliente.Checked;  
		}

		private void chkFecha_CheckedChanged(object sender, System.EventArgs e)
		{
				this.dtFechaDesde.Enabled=!this.chkFecha.Checked;  
				this.dtFechaHasta.Enabled=!this.chkFecha.Checked;
		}

		private void chkGuia_CheckedChanged(object sender, System.EventArgs e)
		{
				this.txtGuia1.Enabled  =  !this.chkGuia.Checked;
				this.txtGuia2.Enabled  =  !this.chkGuia.Checked;
		}

		private void chkRuta_CheckedChanged(object sender, System.EventArgs e)
		{
				this.cmbPuntoPartida.Enabled  =  !this.chkRuta.Checked;
				this.cmbPuntoLlegada.Enabled  =  !this.chkRuta.Checked; 
		}

		private void ultraGrid1_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}

		private void barraDato1_Borrar(object sender, System.EventArgs e)
		{
		
		}

		private void barraDato1_Buscar(object sender, System.EventArgs e)
		{
		

		}

		private void barraDato1_Crear(object sender, System.EventArgs e)
		{
			this.barraDato1.CrearRegistro(); 
			editar=0; 

		}

		private void barraDato1_Deshacer(object sender, System.EventArgs e)
		{
			this.barraDato1.DeshacerRegistro();  
		}

		private void barraDato1_Editar(object sender, System.EventArgs e)
		{
			this.barraDato1.EditarRegistro();  
			editar=1;
		}

		private void barraDato1_Grabar(object sender, System.EventArgs e)
		{

			if (!Verifica()) return;
  		if (editar==1)
			{
				this.barraDato1.GrabaRegistro(); 
			}
			else
			{
				DateTime dtFechaIni1 = (DateTime) dtFecha.Value;
				string stSelect= "insert into guiaremisionpadre (tipo,numero,fecha,idcliente) values(2,"
					+ txtNumero.Value.ToString() + ",'"
					+ dtFechaIni1.ToString("yyyyMMdd") + "',"
					+ cmbCliente1.Value.ToString() + ")";   
				
				this.cdsGuiaRemisionPadre.Update(); 
        
				this.barraDato1.DeshacerRegistro();  
				
				Funcion.EjecutaSQL(cdsGuiaPunto, stSelect, true);	
				this.barraDato1.PosUltima();
				
				stSelect="update guiaremisiong set idguiaremisionpadre=" 
					+ txtidGuiaRemisionPadre.Value.ToString() + " where cobrado=1 and idguiaremisionpadre is null";
				Funcion.EjecutaSQL(cdsGuiaPunto, stSelect, true);	

				this.barraDato1.PosUltima();
			}


		}

		private bool Verifica()
		{
			bool bOk = true;
			BorraErrores();
			if (this.dtFecha.Value == DBNull.Value)
			{
				this.errorProvider.SetError(this.dtFecha , "Ingrese fecha ");
				bOk = false;
			}
			if (this.dtFecha.Text.Trim().Length == 0)
			{
				this.errorProvider.SetError(this.dtFecha , "Ingrese fecha ");
				bOk = false;
			}
			if (this.txtNumero.Text.Trim().Length == 0)
			{
				this.errorProvider.SetError(this.txtNumero , "Ingrese numero de factura");
				bOk = false;
			}
			if (this.cmbCliente1.Text.Length == 0)
			{
				this.errorProvider.SetError(this.cmbCliente1, "Ingrese cliente");
				bOk = false;
			}

			return bOk;
		}

		private void BorraErrores()
		{
			this.errorProvider.SetError(this.dtFecha, "");
			this.errorProvider.SetError(this.txtNumero , "");
			this.errorProvider.SetError(this.cmbCliente1, "");

		}

		private void barraDato1_Refresca(object sender, System.EventArgs e)
		{
			try
			{
				drGuiaPadre = cdsGuiaRemisionG.TableViews["GuiaRemisionPadre"].Rows[0];
			}
			catch{};
		}

		private void ultraNumericEditor2_ValueChanged(object sender, System.EventArgs e)
		{
		
		}

		private void ultraNumericEditor3_ValueChanged(object sender, System.EventArgs e)
		{
		
		}

		private void chkCliente_Validated(object sender, System.EventArgs e)
		{
		}

		private void btnAsiento_Click(object sender, System.EventArgs e)
		{
			if (barraDato1.bEditar) return;
			if (DialogResult.No == MessageBox.Show("¿Desea Generar un Factura de Compra?", 
				"Información", MessageBoxButtons.YesNo, MessageBoxIcon.Question)) return;
			string minota = "VIAJES";		
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.ultraGrid1.Rows.All)
			{
	
				bool bCob = bool.Parse(dr.Cells[38].Value.ToString());
				if (bCob) //esta tomada en cuenta 
				{
					string sel="select nombre from guiaremisionpunto where idguiaremisionpunto="+dr.Cells[15].Value;
					string origen=Funcion.sEscalarSQL(cdsCliente,sel );  
					sel="select nombre from guiaremisionpunto where idguiaremisionpunto="+dr.Cells[17].Value;
					string destino=Funcion.sEscalarSQL(cdsCliente,sel );  
					
					minota = minota + " "+  origen + " - " + destino 	+ " G/R " + dr.Cells[4].Value;
				}
			}
			
			DateTime dtFechaIni = (DateTime) dtFecha.Value;
			string	stExec = "Exec CreaFactura '" + cmbCliente1.Text.Trim().ToString()  + "','001001'"
				+ ",'" + txtNumero.Value.ToString() +"','" + dtFechaIni.ToString("yyyyMMdd") + "',0,0,0,0,0,4,"
				+ "@CONTADOCREDITO=2,@CREAR = 1, @Nota='" + minota.ToString() + "'";

			int IdCompra = Funcion.iEscalarSQL(cdsCliente, stExec, true);
			if (IdCompra == 0)
			{
				MessageBox.Show("Factura de Compra NO Generada porque ya existe", 
					"Información", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
				return;
			}

			MessageBox.Show("Factura de Compra Generada", 
				"Información", MessageBoxButtons.OK, MessageBoxIcon.Information);
		}

		private void label18_Click(object sender, System.EventArgs e)
		{
		
		}
		
	}
}
