using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de gsImpresion.
	/// </summary>
	public class gsImpresion : DevExpress.XtraEditors.XtraForm
	{
		private Infragistics.Win.Misc.UltraLabel ultraLabel1;
		private Infragistics.Win.Misc.UltraLabel ultraLabel2;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkRollo;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtRollo;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtImpresion;
		private Infragistics.Win.Misc.UltraButton btAceptar;
		private Infragistics.Win.Misc.UltraButton btCancelar;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbCliente;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkCliente;
		private C1.Data.C1DataSet cdsCliente;
		private C1.Data.C1DataView cdvCliente;
		private Infragistics.Win.UltraWinEditors.UltraOptionSet optFactura;
		private Infragistics.Win.UltraWinEditors.UltraOptionSet optImprimir;
		private System.Windows.Forms.ErrorProvider errorProvider;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo cmbDesde;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo cmbHasta;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtReporte;
		private System.Windows.Forms.Label label1;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public gsImpresion()
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
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton1 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton2 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Cliente", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Ambos");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("AutorizaFactura");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("AutorizaImprenta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Borrar");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Celular");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Contacto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("DescuentoContado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("DescuentoCredito");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("DiasCredito");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn11 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Direccion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn12 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Email");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn13 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Fax");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn14 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Fecha");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn15 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaIngreso");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn16 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Foto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn17 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCiudad");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn18 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCliente");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn19 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCobrador");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn20 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idGrupoCliente");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn21 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idSector");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn22 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idTipoPrecio");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn23 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idTipoRuc");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn24 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idVendedor");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn25 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("MaximoCredito");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn26 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn27 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Proveedor");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn28 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Representante");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn29 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Ruc");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn30 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("SerieFactura");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn31 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Telefono");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn32 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idSucursal");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn33 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("SaldoCxC");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn34 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("SaldoCxP");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn35 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Editar");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn36 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nuevo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn37 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaEditado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn38 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPadre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn39 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Observacion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn40 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Pais");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn41 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Notas");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn42 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idRetRenta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn43 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idRetIva");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn44 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaCaducidad");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn45 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idFormaPago");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn46 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("AutorizaRet");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn47 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("SerieRet");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn48 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCredTribut");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn49 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ContribEspecial");
			Infragistics.Win.Appearance appearance1 = new Infragistics.Win.Appearance();
			Infragistics.Win.ValueListItem valueListItem1 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem2 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem3 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.ValueListItem valueListItem4 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem5 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem6 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem7 = new Infragistics.Win.ValueListItem();
			this.btAceptar = new Infragistics.Win.Misc.UltraButton();
			this.btCancelar = new Infragistics.Win.Misc.UltraButton();
			this.cmbDesde = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.cmbHasta = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.ultraLabel1 = new Infragistics.Win.Misc.UltraLabel();
			this.ultraLabel2 = new Infragistics.Win.Misc.UltraLabel();
			this.chkRollo = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.txtRollo = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtImpresion = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.cmbCliente = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.chkCliente = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.cdsCliente = new C1.Data.C1DataSet();
			this.cdvCliente = new C1.Data.C1DataView();
			this.optFactura = new Infragistics.Win.UltraWinEditors.UltraOptionSet();
			this.optImprimir = new Infragistics.Win.UltraWinEditors.UltraOptionSet();
			this.errorProvider = new System.Windows.Forms.ErrorProvider();
			this.txtReporte = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label1 = new System.Windows.Forms.Label();
			((System.ComponentModel.ISupportInitialize)(this.cmbDesde)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbHasta)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtRollo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtImpresion)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCliente)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsCliente)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvCliente)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.optFactura)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.optImprimir)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtReporte)).BeginInit();
			this.SuspendLayout();
			// 
			// btAceptar
			// 
			this.btAceptar.Location = new System.Drawing.Point(88, 256);
			this.btAceptar.Name = "btAceptar";
			this.btAceptar.TabIndex = 0;
			this.btAceptar.Text = "&Aceptar";
			this.btAceptar.Click += new System.EventHandler(this.btAceptar_Click);
			// 
			// btCancelar
			// 
			this.btCancelar.Location = new System.Drawing.Point(288, 256);
			this.btCancelar.Name = "btCancelar";
			this.btCancelar.TabIndex = 1;
			this.btCancelar.Text = "&Cancelar";
			this.btCancelar.Click += new System.EventHandler(this.btCancelar_Click);
			// 
			// cmbDesde
			// 
			dateButton1.Caption = "Today";
			this.cmbDesde.DateButtons.Add(dateButton1);
			this.cmbDesde.Format = "dd/MMM/yyyy";
			this.cmbDesde.Location = new System.Drawing.Point(88, 24);
			this.cmbDesde.Name = "cmbDesde";
			this.cmbDesde.NonAutoSizeHeight = 23;
			this.cmbDesde.Size = new System.Drawing.Size(120, 21);
			this.cmbDesde.SpinButtonsVisible = true;
			this.cmbDesde.TabIndex = 2;
			// 
			// cmbHasta
			// 
			dateButton2.Caption = "Today";
			this.cmbHasta.DateButtons.Add(dateButton2);
			this.cmbHasta.Format = "dd/MMM/yyyy";
			this.cmbHasta.Location = new System.Drawing.Point(344, 24);
			this.cmbHasta.Name = "cmbHasta";
			this.cmbHasta.NonAutoSizeHeight = 23;
			this.cmbHasta.Size = new System.Drawing.Size(120, 21);
			this.cmbHasta.SpinButtonsVisible = true;
			this.cmbHasta.TabIndex = 3;
			// 
			// ultraLabel1
			// 
			this.ultraLabel1.AutoSize = true;
			this.ultraLabel1.Location = new System.Drawing.Point(32, 24);
			this.ultraLabel1.Name = "ultraLabel1";
			this.ultraLabel1.Size = new System.Drawing.Size(37, 14);
			this.ultraLabel1.TabIndex = 4;
			this.ultraLabel1.Text = "Desde";
			// 
			// ultraLabel2
			// 
			this.ultraLabel2.AutoSize = true;
			this.ultraLabel2.Location = new System.Drawing.Point(296, 24);
			this.ultraLabel2.Name = "ultraLabel2";
			this.ultraLabel2.Size = new System.Drawing.Size(34, 14);
			this.ultraLabel2.TabIndex = 5;
			this.ultraLabel2.Text = "Hasta";
			// 
			// chkRollo
			// 
			this.chkRollo.Checked = true;
			this.chkRollo.CheckState = System.Windows.Forms.CheckState.Checked;
			this.chkRollo.Location = new System.Drawing.Point(32, 88);
			this.chkRollo.Name = "chkRollo";
			this.chkRollo.Size = new System.Drawing.Size(80, 20);
			this.chkRollo.TabIndex = 6;
			this.chkRollo.Text = "Todo Rollo";
			this.chkRollo.CheckedChanged += new System.EventHandler(this.chkRollo_CheckedChanged);
			// 
			// txtRollo
			// 
			this.txtRollo.Enabled = false;
			this.txtRollo.Location = new System.Drawing.Point(120, 88);
			this.txtRollo.Name = "txtRollo";
			this.txtRollo.Size = new System.Drawing.Size(100, 21);
			this.txtRollo.TabIndex = 7;
			// 
			// txtImpresion
			// 
			this.txtImpresion.Enabled = false;
			this.txtImpresion.Location = new System.Drawing.Point(392, 168);
			this.txtImpresion.Name = "txtImpresion";
			this.txtImpresion.Size = new System.Drawing.Size(100, 21);
			this.txtImpresion.TabIndex = 9;
			// 
			// cmbCliente
			// 
			this.cmbCliente.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbCliente.Cursor = System.Windows.Forms.Cursors.Default;
			ultraGridColumn1.Header.VisiblePosition = 2;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn2.Header.VisiblePosition = 3;
			ultraGridColumn2.Hidden = true;
			ultraGridColumn3.Header.VisiblePosition = 4;
			ultraGridColumn3.Hidden = true;
			ultraGridColumn4.Header.VisiblePosition = 5;
			ultraGridColumn4.Hidden = true;
			ultraGridColumn5.Header.VisiblePosition = 6;
			ultraGridColumn5.Hidden = true;
			ultraGridColumn6.Header.VisiblePosition = 1;
			ultraGridColumn7.Header.VisiblePosition = 7;
			ultraGridColumn7.Hidden = true;
			ultraGridColumn8.Header.VisiblePosition = 8;
			ultraGridColumn8.Hidden = true;
			ultraGridColumn9.Header.VisiblePosition = 9;
			ultraGridColumn9.Hidden = true;
			ultraGridColumn10.Header.VisiblePosition = 10;
			ultraGridColumn10.Hidden = true;
			ultraGridColumn11.Header.VisiblePosition = 11;
			ultraGridColumn11.Hidden = true;
			ultraGridColumn12.Header.VisiblePosition = 12;
			ultraGridColumn12.Hidden = true;
			ultraGridColumn13.Header.VisiblePosition = 13;
			ultraGridColumn13.Hidden = true;
			ultraGridColumn14.Header.VisiblePosition = 14;
			ultraGridColumn14.Hidden = true;
			ultraGridColumn15.Header.VisiblePosition = 15;
			ultraGridColumn15.Hidden = true;
			ultraGridColumn16.Header.VisiblePosition = 16;
			ultraGridColumn16.Hidden = true;
			ultraGridColumn17.Header.VisiblePosition = 17;
			ultraGridColumn17.Hidden = true;
			ultraGridColumn18.Header.VisiblePosition = 18;
			ultraGridColumn18.Hidden = true;
			ultraGridColumn19.Header.VisiblePosition = 19;
			ultraGridColumn19.Hidden = true;
			ultraGridColumn20.Header.VisiblePosition = 20;
			ultraGridColumn20.Hidden = true;
			ultraGridColumn21.Header.VisiblePosition = 21;
			ultraGridColumn21.Hidden = true;
			ultraGridColumn22.Header.VisiblePosition = 22;
			ultraGridColumn22.Hidden = true;
			ultraGridColumn23.Header.VisiblePosition = 23;
			ultraGridColumn23.Hidden = true;
			ultraGridColumn24.Header.VisiblePosition = 24;
			ultraGridColumn24.Hidden = true;
			ultraGridColumn25.Header.VisiblePosition = 25;
			ultraGridColumn25.Hidden = true;
			ultraGridColumn26.Header.VisiblePosition = 0;
			ultraGridColumn27.Header.VisiblePosition = 26;
			ultraGridColumn27.Hidden = true;
			ultraGridColumn28.Header.VisiblePosition = 27;
			ultraGridColumn28.Hidden = true;
			ultraGridColumn29.Header.VisiblePosition = 28;
			ultraGridColumn29.Hidden = true;
			ultraGridColumn30.Header.VisiblePosition = 29;
			ultraGridColumn30.Hidden = true;
			ultraGridColumn31.Header.VisiblePosition = 30;
			ultraGridColumn31.Hidden = true;
			ultraGridColumn32.Header.VisiblePosition = 31;
			ultraGridColumn32.Hidden = true;
			ultraGridColumn33.Header.VisiblePosition = 32;
			ultraGridColumn33.Hidden = true;
			ultraGridColumn34.Header.VisiblePosition = 33;
			ultraGridColumn34.Hidden = true;
			ultraGridColumn35.Header.VisiblePosition = 34;
			ultraGridColumn35.Hidden = true;
			ultraGridColumn36.Header.VisiblePosition = 35;
			ultraGridColumn36.Hidden = true;
			ultraGridColumn37.Header.VisiblePosition = 36;
			ultraGridColumn37.Hidden = true;
			ultraGridColumn38.Header.VisiblePosition = 37;
			ultraGridColumn38.Hidden = true;
			ultraGridColumn39.Header.VisiblePosition = 38;
			ultraGridColumn39.Hidden = true;
			ultraGridColumn40.Header.VisiblePosition = 39;
			ultraGridColumn40.Hidden = true;
			ultraGridColumn41.Header.VisiblePosition = 40;
			ultraGridColumn41.Hidden = true;
			ultraGridColumn42.Header.VisiblePosition = 41;
			ultraGridColumn42.Hidden = true;
			ultraGridColumn43.Header.VisiblePosition = 42;
			ultraGridColumn43.Hidden = true;
			ultraGridColumn44.Header.VisiblePosition = 43;
			ultraGridColumn44.Hidden = true;
			ultraGridColumn45.Header.VisiblePosition = 44;
			ultraGridColumn45.Hidden = true;
			ultraGridColumn46.Header.VisiblePosition = 45;
			ultraGridColumn46.Hidden = true;
			ultraGridColumn47.Header.VisiblePosition = 46;
			ultraGridColumn47.Hidden = true;
			ultraGridColumn48.Header.VisiblePosition = 47;
			ultraGridColumn48.Hidden = true;
			ultraGridColumn49.Header.VisiblePosition = 48;
			ultraGridColumn49.Hidden = true;
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
																										 ultraGridColumn49});
			this.cmbCliente.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.cmbCliente.DisplayMember = "Nombre";
			this.cmbCliente.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbCliente.Enabled = false;
			this.cmbCliente.Location = new System.Drawing.Point(120, 56);
			this.cmbCliente.Name = "cmbCliente";
			this.cmbCliente.Size = new System.Drawing.Size(344, 21);
			this.cmbCliente.TabIndex = 10;
			this.cmbCliente.ValueMember = "idCliente";
			// 
			// chkCliente
			// 
			this.chkCliente.Checked = true;
			this.chkCliente.CheckState = System.Windows.Forms.CheckState.Checked;
			this.chkCliente.Location = new System.Drawing.Point(32, 56);
			this.chkCliente.Name = "chkCliente";
			this.chkCliente.Size = new System.Drawing.Size(88, 20);
			this.chkCliente.TabIndex = 11;
			this.chkCliente.Text = "Todo Cliente";
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
			// 
			// cdvCliente
			// 
			this.cdvCliente.BindingContextCtrl = this;
			this.cdvCliente.DataSet = this.cdsCliente;
			this.cdvCliente.RowFilter = "Proveedor = 0 Or Ambos = 1";
			this.cdvCliente.Sort = "Nombre";
			this.cdvCliente.TableName = "";
			this.cdvCliente.TableViewName = "Cliente";
			// 
			// optFactura
			// 
			this.optFactura.CheckedIndex = 0;
			this.optFactura.ItemAppearance = appearance1;
			this.optFactura.ItemOrigin = new System.Drawing.Point(6, 6);
			valueListItem1.DataValue = 0;
			valueListItem1.DisplayText = "Todo";
			valueListItem2.DataValue = 1;
			valueListItem2.DisplayText = "Sin Facturar";
			valueListItem3.DataValue = 2;
			valueListItem3.DisplayText = "Solo Facturadas";
			this.optFactura.Items.Add(valueListItem1);
			this.optFactura.Items.Add(valueListItem2);
			this.optFactura.Items.Add(valueListItem3);
			this.optFactura.ItemSpacingHorizontal = 4;
			this.optFactura.ItemSpacingVertical = 4;
			this.optFactura.Location = new System.Drawing.Point(56, 128);
			this.optFactura.Name = "optFactura";
			this.optFactura.Size = new System.Drawing.Size(128, 72);
			this.optFactura.TabIndex = 12;
			this.optFactura.Text = "Todo";
			// 
			// optImprimir
			// 
			this.optImprimir.CheckedIndex = 0;
			this.optImprimir.ItemAppearance = appearance2;
			this.optImprimir.ItemOrigin = new System.Drawing.Point(6, 6);
			valueListItem4.DataValue = 0;
			valueListItem4.DisplayText = "Todo";
			valueListItem5.DataValue = 1;
			valueListItem5.DisplayText = "Sin Impresion";
			valueListItem6.DataValue = 2;
			valueListItem6.DisplayText = "Con Impresion";
			valueListItem7.DataValue = 3;
			valueListItem7.DisplayText = "Impresion =";
			this.optImprimir.Items.Add(valueListItem4);
			this.optImprimir.Items.Add(valueListItem5);
			this.optImprimir.Items.Add(valueListItem6);
			this.optImprimir.Items.Add(valueListItem7);
			this.optImprimir.ItemSpacingHorizontal = 4;
			this.optImprimir.ItemSpacingVertical = 4;
			this.optImprimir.Location = new System.Drawing.Point(272, 104);
			this.optImprimir.Name = "optImprimir";
			this.optImprimir.Size = new System.Drawing.Size(112, 96);
			this.optImprimir.TabIndex = 13;
			this.optImprimir.Text = "Todo";
			// 
			// errorProvider
			// 
			this.errorProvider.ContainerControl = this;
			// 
			// txtReporte
			// 
			this.txtReporte.Enabled = false;
			this.txtReporte.Location = new System.Drawing.Point(128, 208);
			this.txtReporte.Name = "txtReporte";
			this.txtReporte.Size = new System.Drawing.Size(100, 21);
			this.txtReporte.TabIndex = 14;
			this.txtReporte.Text = "ProduccionGs_V";
			// 
			// label1
			// 
			this.label1.AutoSize = true;
			this.label1.Location = new System.Drawing.Point(32, 208);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(44, 16);
			this.label1.TabIndex = 15;
			this.label1.Text = "Reporte";
			// 
			// gsImpresion
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(504, 310);
			this.Controls.Add(this.label1);
			this.Controls.Add(this.txtReporte);
			this.Controls.Add(this.optImprimir);
			this.Controls.Add(this.optFactura);
			this.Controls.Add(this.chkCliente);
			this.Controls.Add(this.cmbCliente);
			this.Controls.Add(this.txtImpresion);
			this.Controls.Add(this.txtRollo);
			this.Controls.Add(this.chkRollo);
			this.Controls.Add(this.ultraLabel2);
			this.Controls.Add(this.ultraLabel1);
			this.Controls.Add(this.cmbHasta);
			this.Controls.Add(this.cmbDesde);
			this.Controls.Add(this.btCancelar);
			this.Controls.Add(this.btAceptar);
			this.Name = "gsImpresion";
			this.Text = "Impresion GraphicSource";
			((System.ComponentModel.ISupportInitialize)(this.cmbDesde)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbHasta)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtRollo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtImpresion)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCliente)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsCliente)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvCliente)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.optFactura)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.optImprimir)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtReporte)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void chkRollo_CheckedChanged(object sender, System.EventArgs e)
		{
			txtRollo.Enabled = !chkRollo.Checked;
			if (!chkRollo.Checked) txtRollo.Select();
		}

		private void BorraErrores()
		{
			errorProvider.SetError(txtRollo, "");
			errorProvider.SetError(txtImpresion, "");
			errorProvider.SetError(cmbCliente, "");
		}

		private bool Verifica()
		{
			bool bOk = false;
			BorraErrores();
			if ((int) optImprimir.Value == 3)
				if (txtImpresion.Text.Trim().Length < 1)
				{
					errorProvider.SetError(txtImpresion, "Ingrese Impresion");
					bOk = false;
				}

			if (!chkRollo.Checked) 
				if (txtRollo.Text.Trim().Length < 1)
				{
					errorProvider.SetError(txtRollo, "Ingrese Rollo");
					bOk = false;
				}

			if (!chkCliente.Checked) 
				if (cmbCliente.Text.Trim().Length < 1)
				{
					errorProvider.SetError(cmbCliente, "Ingrese Cliente");
					bOk = false;
				}
			
			return bOk;
		}

		private void btAceptar_Click(object sender, System.EventArgs e)
		{
			if (!Verifica()) return;
			Cursor = Cursors.WaitCursor;
			DateTime dtDesde = (DateTime) cmbDesde.Value;
			DateTime dtHasta = (DateTime) cmbHasta.Value;
			string stFiltro = "{gsProduccion.Fecha} >= #" + dtDesde.ToString("MM/dd/yyyy") 
				+ "# And {gsProduccion.Fecha} < #" + dtHasta.AddDays(1).ToString("MM/dd/yyyy") + "# ";

			if ((int) optImprimir.Value == 1)
				stFiltro = "And Length({gsProduccion.Impresion}) = 0 ";
			if ((int) optImprimir.Value == 2)
				stFiltro = "And Length({gsProduccion.Impresion}) > 0 ";
			if ((int) optImprimir.Value == 3)
				stFiltro = "And {gsProduccion.Impresion} = '"
					+ txtImpresion.Text.Trim() + "' ";
			if (!chkRollo.Checked)
				stFiltro = "And {gsDetProduccion.Rollo} = '"
					+ txtRollo.Text.Trim() + "' ";

			string stReporte = txtReporte.Text.Trim();
			Reporte miRep = new Reporte(stReporte, stFiltro);
			miRep.MdiParent = this.MdiParent;
			miRep.Show();
			Cursor = Cursors.Default;
		}

		private void btCancelar_Click(object sender, System.EventArgs e)
		{
			Close();
		}
	}
}
