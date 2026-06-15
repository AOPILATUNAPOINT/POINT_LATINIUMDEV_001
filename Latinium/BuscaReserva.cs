using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de BuscaReserva.
	/// </summary>
	public class BuscaReserva : DevExpress.XtraEditors.XtraForm
	{
		private Infragistics.Win.Misc.UltraButton btBuscar;
		private Infragistics.Win.Misc.UltraButton btCancelar;
		public Infragistics.Win.UltraWinEditors.UltraNumericEditor txtIdBusca;
		private C1.Data.C1DataSet cdsReserva;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtReserva;
		private Infragistics.Win.UltraWinEditors.UltraOptionSet optTipo;
		private Infragistics.Win.Misc.UltraButton btAceptar;
		private Infragistics.Win.UltraWinGrid.UltraGrid grdReserva;
		private C1.Data.C1DataView cdvReserva;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtPasajero;
		private System.Windows.Forms.ErrorProvider errorProvider;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public BuscaReserva()
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
			Infragistics.Win.ValueListItem valueListItem1 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem2 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Reservacion", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idReservacion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCliente");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Reservacion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Comision");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("PorComision");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Impuestos");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Servicios");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Intermediacion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("IvaIntermediacion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("IvaInterPorcentaje");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn11 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bono");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn12 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Subsidio");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn13 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("TotalSinIva");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn14 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("TotalConIva");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn15 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("TotalServicios");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn16 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("TotalIva");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn17 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("TotalFactura");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn18 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Estado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn19 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idFactura");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn20 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("TipoComision");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn21 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("NumBoletos");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn22 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Pasajeros");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn23 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Notas");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn24 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Numero");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn25 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Fecha");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn26 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Reserva");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn27 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Impuest");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn28 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("NombreBono");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn29 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FormImpuesto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn30 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FormReserva");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn31 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FormSubsidio");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn32 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaSalida");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn33 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaSalida1");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn34 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FormSubsidio1");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn35 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Usuario");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn36 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Porcentaje");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn37 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Borrar");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn38 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("NumeroBono");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn39 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Fee");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn40 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Descuento");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn41 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("VContadoReserva");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn42 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("VCreditoReserva");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn43 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("VContadoContab");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn44 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("VCreditoContab");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn45 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idFacturaAux");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn46 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ContadoCredito");
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			this.txtReserva = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.btBuscar = new Infragistics.Win.Misc.UltraButton();
			this.btCancelar = new Infragistics.Win.Misc.UltraButton();
			this.txtIdBusca = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.cdsReserva = new C1.Data.C1DataSet();
			this.optTipo = new Infragistics.Win.UltraWinEditors.UltraOptionSet();
			this.grdReserva = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.btAceptar = new Infragistics.Win.Misc.UltraButton();
			this.cdvReserva = new C1.Data.C1DataView();
			this.txtPasajero = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.errorProvider = new System.Windows.Forms.ErrorProvider();
			((System.ComponentModel.ISupportInitialize)(this.txtReserva)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdBusca)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsReserva)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.optTipo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.grdReserva)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvReserva)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtPasajero)).BeginInit();
			this.SuspendLayout();
			// 
			// txtReserva
			// 
			this.txtReserva.Location = new System.Drawing.Point(128, 26);
			this.txtReserva.Name = "txtReserva";
			this.txtReserva.PromptChar = ' ';
			this.txtReserva.Size = new System.Drawing.Size(144, 22);
			this.txtReserva.TabIndex = 0;
			// 
			// btBuscar
			// 
			this.btBuscar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btBuscar.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.btBuscar.Location = new System.Drawing.Point(16, 224);
			this.btBuscar.Name = "btBuscar";
			this.btBuscar.Size = new System.Drawing.Size(75, 25);
			this.btBuscar.TabIndex = 2;
			this.btBuscar.Text = "&Buscar";
			this.btBuscar.Click += new System.EventHandler(this.btBuscar_Click);
			// 
			// btCancelar
			// 
			this.btCancelar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btCancelar.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.btCancelar.DialogResult = System.Windows.Forms.DialogResult.Cancel;
			this.btCancelar.Location = new System.Drawing.Point(208, 224);
			this.btCancelar.Name = "btCancelar";
			this.btCancelar.Size = new System.Drawing.Size(75, 25);
			this.btCancelar.TabIndex = 3;
			this.btCancelar.Text = "&Cancelar";
			// 
			// txtIdBusca
			// 
			this.txtIdBusca.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsReserva, "Reservacion.idReservacion"));
			this.txtIdBusca.Location = new System.Drawing.Point(248, 0);
			this.txtIdBusca.Name = "txtIdBusca";
			this.txtIdBusca.PromptChar = ' ';
			this.txtIdBusca.Size = new System.Drawing.Size(40, 22);
			this.txtIdBusca.TabIndex = 4;
			// 
			// cdsReserva
			// 
			this.cdsReserva.BindingContextCtrl = this;
			this.cdsReserva.DataLibrary = "LibFacturacion";
			this.cdsReserva.DataLibraryUrl = "";
			this.cdsReserva.DataSetDef = "dsReservacion";
			this.cdsReserva.FillOnRequest = false;
			this.cdsReserva.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsReserva.Name = "cdsReserva";
			this.cdsReserva.SchemaClassName = "LibFacturacion.DataClass";
			this.cdsReserva.SchemaDef = null;
			// 
			// optTipo
			// 
			this.optTipo.ItemAppearance = appearance1;
			this.optTipo.ItemOrigin = new System.Drawing.Point(10, 4);
			valueListItem1.DataValue = 0;
			valueListItem1.DisplayText = "Número";
			valueListItem2.DataValue = 1;
			valueListItem2.DisplayText = "Reserva";
			this.optTipo.Items.Add(valueListItem1);
			this.optTipo.Items.Add(valueListItem2);
			this.optTipo.ItemSpacingVertical = 10;
			this.optTipo.Location = new System.Drawing.Point(24, 17);
			this.optTipo.Name = "optTipo";
			this.optTipo.Size = new System.Drawing.Size(96, 61);
			this.optTipo.TabIndex = 5;
			// 
			// grdReserva
			// 
			this.grdReserva.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.grdReserva.Cursor = System.Windows.Forms.Cursors.Default;
			this.grdReserva.DataMember = "Reservacion";
			this.grdReserva.DataSource = this.cdsReserva;
			appearance2.BackColor = System.Drawing.Color.White;
			this.grdReserva.DisplayLayout.Appearance = appearance2;
			ultraGridColumn1.Header.VisiblePosition = 3;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn2.Header.VisiblePosition = 4;
			ultraGridColumn2.Hidden = true;
			ultraGridColumn3.Header.VisiblePosition = 2;
			ultraGridColumn4.Header.VisiblePosition = 5;
			ultraGridColumn4.Hidden = true;
			ultraGridColumn5.Header.VisiblePosition = 6;
			ultraGridColumn5.Hidden = true;
			ultraGridColumn6.Header.VisiblePosition = 7;
			ultraGridColumn6.Hidden = true;
			ultraGridColumn7.Header.VisiblePosition = 8;
			ultraGridColumn7.Hidden = true;
			ultraGridColumn8.Header.VisiblePosition = 9;
			ultraGridColumn8.Hidden = true;
			ultraGridColumn9.Header.VisiblePosition = 10;
			ultraGridColumn9.Hidden = true;
			ultraGridColumn10.Header.VisiblePosition = 11;
			ultraGridColumn10.Hidden = true;
			ultraGridColumn11.Header.VisiblePosition = 12;
			ultraGridColumn11.Hidden = true;
			ultraGridColumn12.Header.VisiblePosition = 13;
			ultraGridColumn12.Hidden = true;
			ultraGridColumn13.Header.VisiblePosition = 14;
			ultraGridColumn13.Hidden = true;
			ultraGridColumn14.Header.VisiblePosition = 15;
			ultraGridColumn14.Hidden = true;
			ultraGridColumn15.Header.VisiblePosition = 16;
			ultraGridColumn15.Hidden = true;
			ultraGridColumn16.Header.VisiblePosition = 17;
			ultraGridColumn16.Hidden = true;
			ultraGridColumn17.Header.VisiblePosition = 18;
			ultraGridColumn17.Hidden = true;
			ultraGridColumn18.Header.VisiblePosition = 19;
			ultraGridColumn18.Hidden = true;
			ultraGridColumn19.Header.VisiblePosition = 20;
			ultraGridColumn19.Hidden = true;
			ultraGridColumn20.Header.VisiblePosition = 21;
			ultraGridColumn20.Hidden = true;
			ultraGridColumn21.Header.VisiblePosition = 22;
			ultraGridColumn21.Hidden = true;
			ultraGridColumn22.Header.VisiblePosition = 23;
			ultraGridColumn22.Hidden = true;
			ultraGridColumn23.Header.VisiblePosition = 24;
			ultraGridColumn23.Hidden = true;
			ultraGridColumn24.Header.VisiblePosition = 0;
			ultraGridColumn25.Header.VisiblePosition = 25;
			ultraGridColumn25.Hidden = true;
			ultraGridColumn26.Header.VisiblePosition = 1;
			ultraGridColumn26.Width = 154;
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
			ultraGridColumn39.Header.VisiblePosition = 38;
			ultraGridColumn40.Header.VisiblePosition = 39;
			ultraGridColumn41.Header.VisiblePosition = 40;
			ultraGridColumn42.Header.VisiblePosition = 41;
			ultraGridColumn43.Header.VisiblePosition = 42;
			ultraGridColumn44.Header.VisiblePosition = 43;
			ultraGridColumn45.Header.VisiblePosition = 44;
			ultraGridColumn46.Header.VisiblePosition = 45;
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
																										 ultraGridColumn46});
			this.grdReserva.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			appearance3.BackColor = System.Drawing.Color.Transparent;
			this.grdReserva.DisplayLayout.Override.CardAreaAppearance = appearance3;
			appearance4.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance4.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance4.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance4.FontData.BoldAsString = "True";
			appearance4.FontData.Name = "Arial";
			appearance4.FontData.SizeInPoints = 10F;
			appearance4.ForeColor = System.Drawing.Color.White;
			appearance4.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.grdReserva.DisplayLayout.Override.HeaderAppearance = appearance4;
			appearance5.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance5.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance5.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.grdReserva.DisplayLayout.Override.RowSelectorAppearance = appearance5;
			appearance6.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance6.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance6.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.grdReserva.DisplayLayout.Override.SelectedRowAppearance = appearance6;
			this.grdReserva.Location = new System.Drawing.Point(16, 86);
			this.grdReserva.Name = "grdReserva";
			this.grdReserva.Size = new System.Drawing.Size(264, 121);
			this.grdReserva.TabIndex = 7;
			// 
			// btAceptar
			// 
			this.btAceptar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btAceptar.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.btAceptar.DialogResult = System.Windows.Forms.DialogResult.OK;
			this.btAceptar.Location = new System.Drawing.Point(112, 224);
			this.btAceptar.Name = "btAceptar";
			this.btAceptar.Size = new System.Drawing.Size(75, 25);
			this.btAceptar.TabIndex = 8;
			this.btAceptar.Text = "&Aceptar";
			this.btAceptar.Click += new System.EventHandler(this.btAceptar_Click);
			// 
			// cdvReserva
			// 
			this.cdvReserva.BindingContextCtrl = this;
			this.cdvReserva.DataSet = this.cdsReserva;
			this.cdvReserva.TableName = "";
			this.cdvReserva.TableViewName = "Reservacion";
			// 
			// txtPasajero
			// 
			this.txtPasajero.Location = new System.Drawing.Point(128, 52);
			this.txtPasajero.Name = "txtPasajero";
			this.txtPasajero.Size = new System.Drawing.Size(144, 22);
			this.txtPasajero.TabIndex = 9;
			// 
			// errorProvider
			// 
			this.errorProvider.ContainerControl = this;
			// 
			// BuscaReserva
			// 
			this.AcceptButton = this.btBuscar;
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 14);
			this.CancelButton = this.btCancelar;
			this.ClientSize = new System.Drawing.Size(296, 264);
			this.Controls.Add(this.txtPasajero);
			this.Controls.Add(this.btAceptar);
			this.Controls.Add(this.grdReserva);
			this.Controls.Add(this.optTipo);
			this.Controls.Add(this.txtIdBusca);
			this.Controls.Add(this.btCancelar);
			this.Controls.Add(this.btBuscar);
			this.Controls.Add(this.txtReserva);
			this.Name = "BuscaReserva";
			this.Text = "Busca Reserva";
			this.Load += new System.EventHandler(this.BuscaReserva_Load);
			((System.ComponentModel.ISupportInitialize)(this.txtReserva)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdBusca)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsReserva)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.optTipo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.grdReserva)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvReserva)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtPasajero)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void btBuscar_Click(object sender, System.EventArgs e)
		{
			errorProvider.SetError(txtReserva, "");
			errorProvider.SetError(txtPasajero, "");
			string stFiltro = "";
			string stBusca = "";
			if ((int) optTipo.Value == 0)
			{
				if (txtReserva.Value == null) 
				{
					errorProvider.SetError(txtReserva, "Ingrese Reserva");
					return;
				}
				stBusca = txtReserva.Value.ToString();
				stBusca = stBusca.Replace("'", "''");
				stFiltro = "Numero Like '%" + stBusca + "%'";
			}

			if ((int) optTipo.Value == 1)
			{
				if (txtReserva.Value == null) 
				{
					errorProvider.SetError(txtReserva, "Ingrese Reserva");
					return;
				}
				stBusca = txtPasajero.Value.ToString();
				stBusca = stBusca.Replace("'", "''");
				stFiltro = "Reserva Like '%" + stBusca + "%'";
			}
			C1.Data.FilterCondition fcPadre = new 
				C1.Data.FilterCondition(cdsReserva, "Reservacion", stFiltro);
			C1.Data.FilterConditions fcTotal = new C1.Data.FilterConditions();
			fcTotal.Add(fcPadre);

			cdsReserva.Fill(fcTotal, false);
			if (cdvReserva.Count == 0)
			{
				MessageBox.Show("No existen registros que cumplan la condicion de busqueda",
					"Información", MessageBoxButtons.OK, MessageBoxIcon.Information);
				return;
			}

			if (this.cdvReserva.Count == 1)
			{
				this.grdReserva.Select();
				this.DialogResult = DialogResult.OK;
			}
			else
			{
				this.AcceptButton=this.btAceptar;
				this.grdReserva.Select();
			}

		}

		private void btAceptar_Click(object sender, System.EventArgs e)
		{
		}

		private void BuscaReserva_Load(object sender, System.EventArgs e)
		{
			txtIdBusca.Width = 0;
		}

	}
}
