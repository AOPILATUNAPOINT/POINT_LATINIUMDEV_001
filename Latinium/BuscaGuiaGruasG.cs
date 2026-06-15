using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;
using System.Data.SqlClient;
using System.Data.OleDb;
using System.Globalization;
using CrystalDecisions.Shared;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de BuscaAsiento.
	/// </summary>
	public class BuscaGuiaGruasG : DevExpress.XtraEditors.XtraForm
	{
		public System.Windows.Forms.TextBox idBusca;
		private System.Windows.Forms.ErrorProvider errorProvider;
		private Infragistics.Win.UltraWinGrid.UltraGrid ultraGrid1;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtNumero;
		private Infragistics.Win.UltraWinEditors.UltraOptionSet rdgOpcion;
		private Infragistics.Win.UltraWinEditors.UltraOptionSet rdgTipoBusqueda;
		private DevExpress.XtraEditors.SimpleButton btBuscar;
		private DevExpress.XtraEditors.SimpleButton btAceptar;
		private DevExpress.XtraEditors.SimpleButton btCancelar;
		private C1.Data.C1DataSet cdsGuia;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo cmbFecha;
		private C1.Data.C1DataView cdvGuia;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtNombre;
		private DevExpress.XtraEditors.SimpleButton btReporte;
		private System.ComponentModel.IContainer components=null;

		public BuscaGuiaGruasG()
		{
			InitializeComponent();

			//
			// TODO: Agregar código de constructor después de llamar a InitializeComponent
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

		#region Windows Form Designer generated code
		/// <summary>
		/// Método necesario para admitir el Diseñador, no se puede modificar
		/// el contenido del método con el editor de código.
		/// </summary>
		private void InitializeComponent()
		{
			Infragistics.Win.Appearance appearance1 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("GuiaRemisionG", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idGuiaRemisionG");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("fInicioTraslado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("fTermTraslado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("fEmision");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("NroGuia");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("venta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("compra");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("transformacion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("consignacion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("traslado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn11 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("devolucion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn12 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("importacion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn13 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("exportacion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn14 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("otros");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn15 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("fEmision2");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn16 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPuntoPartida");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn17 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("rucCliente");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn18 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPuntoLlegada");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn19 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("nombreTrasnp");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn20 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("rucTransp");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn21 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("observaciones");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn22 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("diesel");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn23 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("adelanto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn24 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("contratado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn25 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idTransportista");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn26 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCiudadEmision");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn27 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Borrar");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn28 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCarro1");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn29 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCarro2");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn30 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCarro3");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn31 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCliente");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn32 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("NombreCliente");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn33 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("TonelajePedido");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn34 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CombAutorizado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn35 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CombConsumido");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn36 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Celular");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn37 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Placa");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn38 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Facturado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn39 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cobrado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn40 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCompra");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn41 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idVenta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn42 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idGuiaRemisionPadre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn43 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("TextoTonelaje");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn44 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Observacion2");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn45 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idGuiaRemisionPadreC");
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.ValueListItem valueListItem1 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem2 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem3 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
			Infragistics.Win.ValueListItem valueListItem4 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem5 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem6 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton1 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.Appearance appearance11 = new Infragistics.Win.Appearance();
			this.cdsGuia = new C1.Data.C1DataSet();
			this.idBusca = new System.Windows.Forms.TextBox();
			this.errorProvider = new System.Windows.Forms.ErrorProvider();
			this.ultraGrid1 = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.txtNumero = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.rdgOpcion = new Infragistics.Win.UltraWinEditors.UltraOptionSet();
			this.rdgTipoBusqueda = new Infragistics.Win.UltraWinEditors.UltraOptionSet();
			this.btBuscar = new DevExpress.XtraEditors.SimpleButton();
			this.btAceptar = new DevExpress.XtraEditors.SimpleButton();
			this.btCancelar = new DevExpress.XtraEditors.SimpleButton();
			this.cmbFecha = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.cdvGuia = new C1.Data.C1DataView();
			this.txtNombre = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.btReporte = new DevExpress.XtraEditors.SimpleButton();
			((System.ComponentModel.ISupportInitialize)(this.cdsGuia)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNumero)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.rdgOpcion)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.rdgTipoBusqueda)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbFecha)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvGuia)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNombre)).BeginInit();
			this.SuspendLayout();
			// 
			// cdsGuia
			// 
			this.cdsGuia.BindingContextCtrl = this;
			this.cdsGuia.CaseSensitive = false;
			this.cdsGuia.DataLibrary = "LibFacturacion";
			this.cdsGuia.DataLibraryUrl = "";
			this.cdsGuia.DataSetDef = "dsGuiaRemisionG";
			this.cdsGuia.EnforceConstraints = false;
			this.cdsGuia.FillOnRequest = false;
			this.cdsGuia.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsGuia.Name = "cdsGuia";
			this.cdsGuia.SchemaClassName = "LibFacturacion.DataClass";
			this.cdsGuia.SchemaDef = null;
			this.cdsGuia.BeforeFill += new C1.Data.FillEventHandler(this.cdsArticulo_BeforeFill);
			// 
			// idBusca
			// 
			this.idBusca.DataBindings.Add(new System.Windows.Forms.Binding("Text", this.cdsGuia, "GuiaRemisionG.idGuiaRemisionG"));
			this.idBusca.Location = new System.Drawing.Point(480, 17);
			this.idBusca.Name = "idBusca";
			this.idBusca.Size = new System.Drawing.Size(48, 20);
			this.idBusca.TabIndex = 23;
			this.idBusca.Text = "";
			// 
			// errorProvider
			// 
			this.errorProvider.ContainerControl = this;
			this.errorProvider.DataMember = "";
			// 
			// ultraGrid1
			// 
			this.ultraGrid1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.ultraGrid1.Cursor = System.Windows.Forms.Cursors.Default;
			this.ultraGrid1.DataMember = "GuiaRemisionG";
			this.ultraGrid1.DataSource = this.cdsGuia;
			appearance1.BackColor = System.Drawing.Color.White;
			this.ultraGrid1.DisplayLayout.Appearance = appearance1;
			ultraGridBand1.AddButtonCaption = "Articulo";
			ultraGridColumn1.Header.VisiblePosition = 6;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn2.Header.VisiblePosition = 5;
			ultraGridColumn2.Hidden = true;
			ultraGridColumn3.Header.VisiblePosition = 8;
			ultraGridColumn3.Hidden = true;
			ultraGridColumn4.Header.VisiblePosition = 4;
			ultraGridColumn4.Hidden = true;
			ultraGridColumn5.Header.VisiblePosition = 0;
			ultraGridColumn5.Width = 68;
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
			ultraGridColumn15.Header.Caption = "Emision";
			ultraGridColumn15.Header.VisiblePosition = 2;
			ultraGridColumn16.Header.VisiblePosition = 18;
			ultraGridColumn16.Hidden = true;
			ultraGridColumn17.Header.VisiblePosition = 19;
			ultraGridColumn17.Hidden = true;
			ultraGridColumn18.Header.VisiblePosition = 20;
			ultraGridColumn18.Hidden = true;
			ultraGridColumn19.Header.VisiblePosition = 21;
			ultraGridColumn19.Hidden = true;
			ultraGridColumn20.Header.VisiblePosition = 22;
			ultraGridColumn20.Hidden = true;
			ultraGridColumn21.Header.VisiblePosition = 23;
			ultraGridColumn21.Hidden = true;
			ultraGridColumn22.Header.VisiblePosition = 24;
			ultraGridColumn22.Hidden = true;
			ultraGridColumn23.Header.VisiblePosition = 25;
			ultraGridColumn23.Hidden = true;
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
			ultraGridColumn31.Header.Caption = "Nombre";
			ultraGridColumn31.Header.VisiblePosition = 3;
			ultraGridColumn31.Hidden = true;
			ultraGridColumn31.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownList;
			ultraGridColumn31.Width = 337;
			ultraGridColumn32.Header.Caption = "Razon Social";
			ultraGridColumn32.Header.VisiblePosition = 1;
			ultraGridColumn32.Width = 307;
			ultraGridColumn33.Header.VisiblePosition = 33;
			ultraGridColumn33.Hidden = true;
			ultraGridColumn34.Header.VisiblePosition = 34;
			ultraGridColumn34.Hidden = true;
			ultraGridColumn35.Header.VisiblePosition = 35;
			ultraGridColumn35.Hidden = true;
			ultraGridColumn36.Header.VisiblePosition = 36;
			ultraGridColumn36.Hidden = true;
			ultraGridColumn37.Header.VisiblePosition = 37;
			ultraGridColumn37.Hidden = true;
			ultraGridColumn38.Header.VisiblePosition = 38;
			ultraGridColumn38.Hidden = true;
			ultraGridColumn39.Header.VisiblePosition = 39;
			ultraGridColumn39.Hidden = true;
			ultraGridColumn40.Header.VisiblePosition = 40;
			ultraGridColumn40.Hidden = true;
			ultraGridColumn41.Header.VisiblePosition = 41;
			ultraGridColumn41.Hidden = true;
			ultraGridColumn42.Header.VisiblePosition = 7;
			ultraGridColumn42.Hidden = true;
			ultraGridColumn43.Header.VisiblePosition = 42;
			ultraGridColumn43.Hidden = true;
			ultraGridColumn44.Header.VisiblePosition = 43;
			ultraGridColumn44.Hidden = true;
			ultraGridColumn45.Header.VisiblePosition = 44;
			ultraGridColumn45.Hidden = true;
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
																										 ultraGridColumn45});
			this.ultraGrid1.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.ultraGrid1.DisplayLayout.GroupByBox.Prompt = "Arrastre columna a Agrupar";
			this.ultraGrid1.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
			this.ultraGrid1.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.ultraGrid1.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.False;
			appearance2.BackColor = System.Drawing.Color.Transparent;
			this.ultraGrid1.DisplayLayout.Override.CardAreaAppearance = appearance2;
			this.ultraGrid1.DisplayLayout.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.RowSelect;
			this.ultraGrid1.DisplayLayout.Override.ColumnAutoSizeMode = Infragistics.Win.UltraWinGrid.ColumnAutoSizeMode.AllRowsInBand;
			appearance3.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance3.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance3.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance3.FontData.BoldAsString = "True";
			appearance3.FontData.Name = "Arial";
			appearance3.FontData.SizeInPoints = 10F;
			appearance3.ForeColor = System.Drawing.Color.White;
			appearance3.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.ultraGrid1.DisplayLayout.Override.HeaderAppearance = appearance3;
			this.ultraGrid1.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti;
			appearance4.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance4.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance4.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.RowSelectorAppearance = appearance4;
			appearance5.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance5.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance5.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.SelectedRowAppearance = appearance5;
			this.ultraGrid1.DisplayLayout.Override.SelectTypeCell = Infragistics.Win.UltraWinGrid.SelectType.None;
			this.ultraGrid1.DisplayLayout.Override.SelectTypeCol = Infragistics.Win.UltraWinGrid.SelectType.None;
			this.ultraGrid1.DisplayLayout.Override.SelectTypeRow = Infragistics.Win.UltraWinGrid.SelectType.Single;
			this.ultraGrid1.DisplayLayout.ScrollBounds = Infragistics.Win.UltraWinGrid.ScrollBounds.ScrollToFill;
			this.ultraGrid1.DisplayLayout.ScrollStyle = Infragistics.Win.UltraWinGrid.ScrollStyle.Immediate;
			this.ultraGrid1.DisplayLayout.TabNavigation = Infragistics.Win.UltraWinGrid.TabNavigation.NextControl;
			this.ultraGrid1.DisplayLayout.ViewStyle = Infragistics.Win.UltraWinGrid.ViewStyle.SingleBand;
			this.ultraGrid1.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.ultraGrid1.Location = new System.Drawing.Point(16, 160);
			this.ultraGrid1.Name = "ultraGrid1";
			this.ultraGrid1.Size = new System.Drawing.Size(512, 264);
			this.ultraGrid1.TabIndex = 27;
			// 
			// txtNumero
			// 
			appearance6.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtNumero.Appearance = appearance6;
			this.txtNumero.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNumero.Location = new System.Drawing.Point(136, 43);
			this.txtNumero.Name = "txtNumero";
			this.txtNumero.Size = new System.Drawing.Size(128, 22);
			this.txtNumero.TabIndex = 39;
			// 
			// rdgOpcion
			// 
			appearance7.ForeColorDisabled = System.Drawing.Color.Black;
			this.rdgOpcion.Appearance = appearance7;
			this.rdgOpcion.ItemAppearance = appearance8;
			this.rdgOpcion.ItemOrigin = new System.Drawing.Point(6, 0);
			valueListItem1.DataValue = 0;
			valueListItem1.DisplayText = "Nombre";
			valueListItem2.DataValue = 1;
			valueListItem2.DisplayText = "No. Guia";
			valueListItem3.DataValue = 2;
			valueListItem3.DisplayText = "F. Emisiòn";
			this.rdgOpcion.Items.Add(valueListItem1);
			this.rdgOpcion.Items.Add(valueListItem2);
			this.rdgOpcion.Items.Add(valueListItem3);
			this.rdgOpcion.ItemSpacingVertical = 9;
			this.rdgOpcion.Location = new System.Drawing.Point(16, 17);
			this.rdgOpcion.Name = "rdgOpcion";
			this.rdgOpcion.Size = new System.Drawing.Size(112, 86);
			this.rdgOpcion.TabIndex = 40;
			this.rdgOpcion.ValueChanged += new System.EventHandler(this.rdgOpcion_SelectedIndexChanged);
			// 
			// rdgTipoBusqueda
			// 
			appearance9.ForeColorDisabled = System.Drawing.Color.Black;
			this.rdgTipoBusqueda.Appearance = appearance9;
			this.rdgTipoBusqueda.CheckedIndex = 2;
			this.rdgTipoBusqueda.ItemAppearance = appearance10;
			this.rdgTipoBusqueda.ItemOrigin = new System.Drawing.Point(10, 0);
			valueListItem4.DataValue = 0;
			valueListItem4.DisplayText = "Que Empiece";
			valueListItem5.DataValue = 1;
			valueListItem5.DisplayText = "Exacta";
			valueListItem6.DataValue = 2;
			valueListItem6.DisplayText = "En cualquier parte";
			this.rdgTipoBusqueda.Items.Add(valueListItem4);
			this.rdgTipoBusqueda.Items.Add(valueListItem5);
			this.rdgTipoBusqueda.Items.Add(valueListItem6);
			this.rdgTipoBusqueda.ItemSpacingVertical = 10;
			this.rdgTipoBusqueda.Location = new System.Drawing.Point(296, 52);
			this.rdgTipoBusqueda.Name = "rdgTipoBusqueda";
			this.rdgTipoBusqueda.Size = new System.Drawing.Size(136, 86);
			this.rdgTipoBusqueda.TabIndex = 41;
			this.rdgTipoBusqueda.Text = "En cualquier parte";
			// 
			// btBuscar
			// 
			this.btBuscar.Location = new System.Drawing.Point(448, 52);
			this.btBuscar.Name = "btBuscar";
			this.btBuscar.Size = new System.Drawing.Size(75, 24);
			this.btBuscar.TabIndex = 42;
			this.btBuscar.Text = "&Buscar";
			this.btBuscar.Click += new System.EventHandler(this.btBuscar_Click);
			// 
			// btAceptar
			// 
			this.btAceptar.Location = new System.Drawing.Point(448, 86);
			this.btAceptar.Name = "btAceptar";
			this.btAceptar.Size = new System.Drawing.Size(75, 25);
			this.btAceptar.TabIndex = 43;
			this.btAceptar.Text = "&Aceptar";
			this.btAceptar.Click += new System.EventHandler(this.btAceptar_Click);
			// 
			// btCancelar
			// 
			this.btCancelar.DialogResult = System.Windows.Forms.DialogResult.Cancel;
			this.btCancelar.Location = new System.Drawing.Point(448, 121);
			this.btCancelar.Name = "btCancelar";
			this.btCancelar.Size = new System.Drawing.Size(75, 24);
			this.btCancelar.TabIndex = 44;
			this.btCancelar.Text = "&Cancelar";
			// 
			// cmbFecha
			// 
			dateButton1.Caption = "Today";
			this.cmbFecha.DateButtons.Add(dateButton1);
			this.cmbFecha.Format = "dd/MMM/yyyy";
			this.cmbFecha.Location = new System.Drawing.Point(136, 69);
			this.cmbFecha.Name = "cmbFecha";
			this.cmbFecha.NonAutoSizeHeight = 24;
			this.cmbFecha.Size = new System.Drawing.Size(128, 21);
			this.cmbFecha.TabIndex = 45;
			this.cmbFecha.Value = new System.DateTime(2009, 10, 26, 0, 0, 0, 0);
			// 
			// cdvGuia
			// 
			this.cdvGuia.BindingContextCtrl = this;
			this.cdvGuia.DataSet = this.cdsGuia;
			this.cdvGuia.TableName = "";
			this.cdvGuia.TableViewName = "GuiaRemisionG";
			// 
			// txtNombre
			// 
			appearance11.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtNombre.Appearance = appearance11;
			this.txtNombre.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNombre.Location = new System.Drawing.Point(136, 17);
			this.txtNombre.Name = "txtNombre";
			this.txtNombre.Size = new System.Drawing.Size(288, 22);
			this.txtNombre.TabIndex = 46;
			// 
			// btReporte
			// 
			this.btReporte.Location = new System.Drawing.Point(448, 17);
			this.btReporte.Name = "btReporte";
			this.btReporte.Size = new System.Drawing.Size(72, 26);
			this.btReporte.TabIndex = 50;
			this.btReporte.Text = "&Reporte";
			this.btReporte.Click += new System.EventHandler(this.btReporte_Click);
			// 
			// BuscaGuiaGruasG
			// 
			this.AcceptButton = this.btBuscar;
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 14);
			this.CancelButton = this.btCancelar;
			this.ClientSize = new System.Drawing.Size(544, 445);
			this.Controls.Add(this.btReporte);
			this.Controls.Add(this.txtNombre);
			this.Controls.Add(this.cmbFecha);
			this.Controls.Add(this.btCancelar);
			this.Controls.Add(this.btAceptar);
			this.Controls.Add(this.btBuscar);
			this.Controls.Add(this.rdgTipoBusqueda);
			this.Controls.Add(this.rdgOpcion);
			this.Controls.Add(this.txtNumero);
			this.Controls.Add(this.ultraGrid1);
			this.Controls.Add(this.idBusca);
			this.MinimizeBox = false;
			this.Name = "BuscaGuiaGruasG";
			this.Text = "Búsqueda de Artículos";
			this.Resize += new System.EventHandler(this.BuscaArticulo_Resize);
			this.Closing += new System.ComponentModel.CancelEventHandler(this.BuscaArticulo_Closing);
			this.Load += new System.EventHandler(this.BuscaAsiento_Load);
			((System.ComponentModel.ISupportInitialize)(this.cdsGuia)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNumero)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.rdgOpcion)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.rdgTipoBusqueda)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbFecha)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvGuia)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNombre)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void BuscaAsiento_Load(object sender, System.EventArgs e)
		{
			Cursor = Cursors.Default;
			this.idBusca.Width =0;
			this.rdgOpcion.Value = 1;
			cmbFecha.Value = DateTime.Today;
		}

		private bool Verifica()
		{
			this.errorProvider.SetError(this.txtNombre, "");
			this.errorProvider.SetError(this.txtNumero, "");
			this.errorProvider.SetError(this.cmbFecha, "");

			bool bOk = true;
			int iTipo = (int) this.rdgOpcion.Value;

			if (iTipo == 0)
				if (txtNombre.Text.Trim().Length == 0)
				{
					errorProvider.SetError(txtNombre, "Ingrese Nombre");
					bOk = false;
				}

			if (iTipo == 1)
				if (this.txtNumero.Text.Trim().Length == 0)
				{
					this.errorProvider.SetError(this.txtNumero, "Ingrese Numero");
					bOk = false;
				}
			if (iTipo == 2)
				if (this.cmbFecha.Value == null)
				{
					this.errorProvider.SetError(this.cmbFecha, "Ingrese Fecha");
					bOk = false;
				}
			return bOk;
		}

		private void btBuscar_Click(object sender, System.EventArgs e)
		{
			if (!Verifica()) return;
			string strFiltro;
			string strAntes = "";
			string strDespues="'";

//		Opciones de comodines
			if ((int) this.rdgTipoBusqueda.Value == 0) strDespues = "%'";
			if ((int) this.rdgTipoBusqueda.Value == 1) strAntes = "";
			if ((int) this.rdgTipoBusqueda.Value == 2) strDespues = "%'";
			if ((int) this.rdgTipoBusqueda.Value == 2) strAntes = "%";

			strFiltro = "";
			string stBusca = "";
			switch((int) this.rdgOpcion.Value)
			{
				case 0:
					stBusca = this.txtNombre.Value.ToString().Trim();
					stBusca = stBusca.Replace("'", "''");
					strFiltro = "NombreCliente LIKE '" + strAntes + stBusca + strDespues;
					break;
				case 1:
					stBusca = this.txtNumero.Value.ToString().Trim();
					stBusca = stBusca.Replace("'", "''");
					strFiltro = "NroGuia LIKE '" + strAntes + stBusca + strDespues;
					break;
				case 2:
					DateTime dtFecha = (DateTime) cmbFecha.Value;
					strFiltro = "FEmision2 = '" + dtFecha.ToString("yyyyMMdd") + "'";
					break;
			}

			string stFiltro = strFiltro;

			C1.Data.FilterCondition fcPadre = new 
				C1.Data.FilterCondition(this.cdsGuia, "GuiaRemisionG", stFiltro);
			C1.Data.FilterConditions fcTotal = new C1.Data.FilterConditions();
			fcTotal.Add(fcPadre);

			this.cdsGuia.Fill(fcTotal, false);
			if (this.cdvGuia.Count == 0)
			{
				MessageBox.Show("No existen registros que cumplan la condicion de busqueda",
					"Información", MessageBoxButtons.OK, MessageBoxIcon.Information);
				return;
			}

			if (this.cdvGuia.Count == 1)
			{
				this.ultraGrid1.Select();
				this.DialogResult = DialogResult.OK;
			}
			else
			{
				this.AcceptButton=this.btAceptar;
				this.ultraGrid1.Select();
			}
		}


		private void btAceptar_Click(object sender, System.EventArgs e)
		{
			this.DialogResult = DialogResult.OK;
		}

		private void BuscaArticulo_Resize(object sender, System.EventArgs e)
		{
			this.ultraGrid1.DisplayLayout.AutoFitColumns = true;
		}

		private void cdsArticulo_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			this.cdsGuia.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);		
		}


		private void rdgOpcion_SelectedIndexChanged(object sender, System.EventArgs e)
		{
			txtNumero.Enabled = false;
			txtNombre.Enabled = false;
			cmbFecha.Enabled = false;

			switch((int) this.rdgOpcion.Value)
			{
				case 0:
					txtNombre.Enabled = true;
					txtNombre.Select();
					break;
				case 1:
					this.txtNumero.Enabled = true;
					this.txtNumero.Select();
					break;
				case 2:
					this.cmbFecha.Enabled = true;
					this.cmbFecha.Select();
					break;
			}

		}

		private void BuscaArticulo_Closing(object sender, System.ComponentModel.CancelEventArgs e)
		{

		}

		private void btReporte_Click(object sender, System.EventArgs e)
		{
			#region //' reporte con crystal
//			// Definicion de Parametros
//			ParameterFields paramFields = new ParameterFields ();
//			ParameterField paramField = new ParameterField ();
//
//			string stFiltro = "";
//			switch((int) this.rdgOpcion.Value)
//			{
//				case 0:
//					stFiltro = "{GuiaRemisionG.NombreCliente} = " + this.txtNombre.Value.ToString();
//					break;
//				case 1:
//					stFiltro = "{GuiaRemisionG.NroGuia} = " + this.txtNumero.Value.ToString();
//					break;
//				case 2:
//					stFiltro = "{GuiaRemisionG.fEmision2} = " + this.cmbFecha.Value;
//					break;
//			}
//			
//			ParameterDiscreteValue discreteVal = new ParameterDiscreteValue ();
//			paramField.ParameterFieldName = "Filtro";
//			discreteVal.Value = stFiltro;
//			paramField.CurrentValues.Add (discreteVal);
//			paramFields.Add (paramField);
//
//			Cursor = Cursors.WaitCursor;
//			Reporte miReporte = new Reporte("GuiaRemisionGT.rpt",stFiltro);
//			miReporte.crVista.ParameterFieldInfo = paramFields;
//			miReporte.MdiParent = this.MdiParent;
//			miReporte.Show();
//			this.Cursor = Cursors.Default;
# endregion 			

			ReporteGuiasGruas repGG = new ReporteGuiasGruas();			
				repGG.MdiParent = this.MdiParent;
				repGG.Show();			
		}
	}
}
