using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data.SqlClient;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de BuscaAsiento.
	/// </summary>
	public class BuscaPagoLotes : DevExpress.XtraEditors.XtraForm
	{
		public System.Windows.Forms.TextBox idBusca;
		private System.Windows.Forms.ErrorProvider errorProvider;
		private C1.Data.C1DataView cdvCliente;
		private C1.Data.C1DataSet cdsCliente;
		private Infragistics.Win.UltraWinEditors.UltraOptionSet rdgTipoBusqueda;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo txtFecha;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbNombre;
		private Infragistics.Win.UltraWinEditors.UltraOptionSet rdgOpcion;
		private Infragistics.Win.UltraWinGrid.UltraGrid ultraGrid1;
		private Infragistics.Win.UltraWinGrid.UltraDropDown ultraDropDown1;
		private Infragistics.Win.Misc.UltraButton btCancelar;
		private Infragistics.Win.Misc.UltraButton btAceptar;
		private Infragistics.Win.Misc.UltraButton btBuscar;
		private Infragistics.Win.Misc.UltraButton btExcel;
		private C1.Data.C1DataView cdvPago;
		private C1.Data.C1DataSet cdsPago;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource udsPago;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtFactura;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo txtFechaPago;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtNumIngreso;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtComprob;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtPago;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtReferencia;
		public System.Windows.Forms.TextBox txtDeposito;
		private System.ComponentModel.IContainer components=null;

		public BuscaPagoLotes()
		{
			InitializeComponent();
		}
		int IdSigno = 1;
		public BuscaPagoLotes(int idTipoF)
		{
			InitializeComponent();
			IdSigno = idTipoF;
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
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCompra");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idPagoLote");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn3 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Pago");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn4 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn5 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Factura");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn6 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Fecha");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn7 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("FechaPago");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn8 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Concepto");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn9 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Cuenta");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn10 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("NumIngreso");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn11 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Referencia");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn12 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("NumPago");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn13 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Cheque");
			Infragistics.Win.Appearance appearance1 = new Infragistics.Win.Appearance();
			Infragistics.Win.ValueListItem valueListItem1 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem2 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem3 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton1 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton2 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Cliente", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCliente");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Proveedor");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Ruc");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Ambos");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("DesconProv");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("DesconCliente");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ContadoCred");
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.ValueListItem valueListItem4 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem5 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem6 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem7 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem8 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem9 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem10 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem11 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCompra");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn11 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPagoLote");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn12 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Pago");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn13 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn14 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Factura");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn15 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Fecha");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn16 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaPago");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn17 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Concepto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn18 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cuenta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn19 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("NumIngreso");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn20 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Referencia");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn21 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("NumPago");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn22 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cheque");
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand3 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Cliente", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn23 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn24 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCliente");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn25 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn26 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Proveedor");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn27 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Ruc");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn28 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Ambos");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn29 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("DesconProv");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn30 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("DesconCliente");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn31 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ContadoCred");
			this.cdvPago = new C1.Data.C1DataView();
			this.cdsPago = new C1.Data.C1DataSet();
			this.cdvCliente = new C1.Data.C1DataView();
			this.cdsCliente = new C1.Data.C1DataSet();
			this.idBusca = new System.Windows.Forms.TextBox();
			this.udsPago = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.errorProvider = new System.Windows.Forms.ErrorProvider();
			this.rdgTipoBusqueda = new Infragistics.Win.UltraWinEditors.UltraOptionSet();
			this.txtFactura = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtFecha = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.txtFechaPago = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.txtNumIngreso = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtPago = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.cmbNombre = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.rdgOpcion = new Infragistics.Win.UltraWinEditors.UltraOptionSet();
			this.ultraGrid1 = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDropDown1 = new Infragistics.Win.UltraWinGrid.UltraDropDown();
			this.btCancelar = new Infragistics.Win.Misc.UltraButton();
			this.btAceptar = new Infragistics.Win.Misc.UltraButton();
			this.btBuscar = new Infragistics.Win.Misc.UltraButton();
			this.btExcel = new Infragistics.Win.Misc.UltraButton();
			this.txtComprob = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtReferencia = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtDeposito = new System.Windows.Forms.TextBox();
			((System.ComponentModel.ISupportInitialize)(this.cdvPago)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsPago)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvCliente)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsCliente)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.udsPago)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.rdgTipoBusqueda)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtFactura)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtFecha)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtFechaPago)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNumIngreso)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtPago)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbNombre)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.rdgOpcion)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDropDown1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtComprob)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtReferencia)).BeginInit();
			this.SuspendLayout();
			// 
			// cdvPago
			// 
			this.cdvPago.BindingContextCtrl = this;
			this.cdvPago.DataSet = this.cdsPago;
			this.cdvPago.TableName = "";
			this.cdvPago.TableViewName = "Pago";
			// 
			// cdsPago
			// 
			this.cdsPago.BindingContextCtrl = this;
			this.cdsPago.CaseSensitive = false;
			this.cdsPago.DataLibrary = "LibFacturacion";
			this.cdsPago.DataLibraryUrl = "";
			this.cdsPago.DataSetDef = "dsPago";
			this.cdsPago.FillOnRequest = false;
			this.cdsPago.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsPago.Name = "cdsPago";
			this.cdsPago.SchemaClassName = "LibFacturacion.DataClass";
			this.cdsPago.SchemaDef = null;
			this.cdsPago.BeforeFill += new C1.Data.FillEventHandler(this.cdsCompra_BeforeFill);
			// 
			// cdvCliente
			// 
			this.cdvCliente.BindingContextCtrl = this;
			this.cdvCliente.DataSet = this.cdsCliente;
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
			// idBusca
			// 
			this.idBusca.DataBindings.Add(new System.Windows.Forms.Binding("Text", this.udsPago, "Band 0.idPagoLote"));
			this.idBusca.Location = new System.Drawing.Point(480, 17);
			this.idBusca.Name = "idBusca";
			this.idBusca.Size = new System.Drawing.Size(48, 20);
			this.idBusca.TabIndex = 23;
			this.idBusca.Text = "";
			// 
			// udsPago
			// 
			ultraDataColumn1.DataType = typeof(int);
			ultraDataColumn2.DataType = typeof(int);
			ultraDataColumn3.DataType = typeof(System.Double);
			ultraDataColumn6.DataType = typeof(System.DateTime);
			ultraDataColumn7.DataType = typeof(System.DateTime);
			this.udsPago.Band.Columns.AddRange(new object[] {
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
																												ultraDataColumn13});
			// 
			// errorProvider
			// 
			this.errorProvider.ContainerControl = this;
			// 
			// rdgTipoBusqueda
			// 
			this.rdgTipoBusqueda.CheckedIndex = 2;
			this.rdgTipoBusqueda.ItemAppearance = appearance1;
			this.rdgTipoBusqueda.ItemOrigin = new System.Drawing.Point(10, 0);
			valueListItem1.DataValue = 0;
			valueListItem1.DisplayText = "Que Empiece";
			valueListItem2.DataValue = 1;
			valueListItem2.DisplayText = "Exacta";
			valueListItem3.DataValue = 2;
			valueListItem3.DisplayText = "En cualquier parte";
			this.rdgTipoBusqueda.Items.Add(valueListItem1);
			this.rdgTipoBusqueda.Items.Add(valueListItem2);
			this.rdgTipoBusqueda.Items.Add(valueListItem3);
			this.rdgTipoBusqueda.ItemSpacingVertical = 10;
			this.rdgTipoBusqueda.Location = new System.Drawing.Point(288, 60);
			this.rdgTipoBusqueda.Name = "rdgTipoBusqueda";
			this.rdgTipoBusqueda.Size = new System.Drawing.Size(136, 86);
			this.rdgTipoBusqueda.TabIndex = 42;
			this.rdgTipoBusqueda.Text = "En cualquier parte";
			// 
			// txtFactura
			// 
			this.txtFactura.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtFactura.Enabled = false;
			this.txtFactura.Location = new System.Drawing.Point(136, 43);
			this.txtFactura.Name = "txtFactura";
			this.txtFactura.Size = new System.Drawing.Size(128, 22);
			this.txtFactura.TabIndex = 43;
			// 
			// txtFecha
			// 
			this.txtFecha.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton1.Caption = "Today";
			this.txtFecha.DateButtons.Add(dateButton1);
			this.txtFecha.Format = "dd/MMM/yyyy";
			this.txtFecha.Location = new System.Drawing.Point(136, 69);
			this.txtFecha.Name = "txtFecha";
			this.txtFecha.NonAutoSizeHeight = 23;
			this.txtFecha.Size = new System.Drawing.Size(128, 21);
			this.txtFecha.SpinButtonsVisible = true;
			this.txtFecha.TabIndex = 45;
			this.txtFecha.Value = new System.DateTime(2005, 5, 17, 0, 0, 0, 0);
			// 
			// txtFechaPago
			// 
			this.txtFechaPago.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton2.Caption = "Today";
			this.txtFechaPago.DateButtons.Add(dateButton2);
			this.txtFechaPago.Format = "dd/MMM/yyyy";
			this.txtFechaPago.Location = new System.Drawing.Point(136, 95);
			this.txtFechaPago.Name = "txtFechaPago";
			this.txtFechaPago.NonAutoSizeHeight = 23;
			this.txtFechaPago.Size = new System.Drawing.Size(128, 21);
			this.txtFechaPago.SpinButtonsVisible = true;
			this.txtFechaPago.TabIndex = 46;
			this.txtFechaPago.Value = new System.DateTime(2005, 5, 17, 0, 0, 0, 0);
			// 
			// txtNumIngreso
			// 
			this.txtNumIngreso.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNumIngreso.Enabled = false;
			this.txtNumIngreso.Location = new System.Drawing.Point(136, 146);
			this.txtNumIngreso.Name = "txtNumIngreso";
			this.txtNumIngreso.Size = new System.Drawing.Size(128, 22);
			this.txtNumIngreso.TabIndex = 48;
			// 
			// txtPago
			// 
			this.txtPago.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtPago.Enabled = false;
			this.txtPago.Location = new System.Drawing.Point(136, 172);
			this.txtPago.Name = "txtPago";
			this.txtPago.Size = new System.Drawing.Size(128, 22);
			this.txtPago.TabIndex = 49;
			// 
			// cmbNombre
			// 
			this.cmbNombre.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbNombre.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbNombre.DataSource = this.cdvCliente;
			ultraGridColumn1.Header.VisiblePosition = 1;
			ultraGridColumn1.Width = 143;
			ultraGridColumn2.Header.VisiblePosition = 3;
			ultraGridColumn2.Hidden = true;
			ultraGridColumn2.Width = 14;
			ultraGridColumn3.Header.VisiblePosition = 0;
			ultraGridColumn3.Width = 296;
			ultraGridColumn4.Header.VisiblePosition = 4;
			ultraGridColumn4.Hidden = true;
			ultraGridColumn4.Width = 30;
			ultraGridColumn5.Header.VisiblePosition = 5;
			ultraGridColumn5.Hidden = true;
			ultraGridColumn5.Width = 18;
			ultraGridColumn6.Header.VisiblePosition = 2;
			ultraGridColumn6.Hidden = true;
			ultraGridColumn6.Width = 43;
			ultraGridColumn7.Header.VisiblePosition = 6;
			ultraGridColumn7.Hidden = true;
			ultraGridColumn8.Header.VisiblePosition = 7;
			ultraGridColumn8.Hidden = true;
			ultraGridColumn9.Header.VisiblePosition = 8;
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
			this.cmbNombre.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.cmbNombre.DisplayMember = "Nombre";
			this.cmbNombre.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbNombre.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbNombre.Enabled = false;
			this.cmbNombre.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbNombre.Location = new System.Drawing.Point(136, 17);
			this.cmbNombre.Name = "cmbNombre";
			this.cmbNombre.Size = new System.Drawing.Size(328, 21);
			this.cmbNombre.TabIndex = 50;
			this.cmbNombre.ValueMember = "idCliente";
			// 
			// rdgOpcion
			// 
			this.rdgOpcion.ItemAppearance = appearance2;
			this.rdgOpcion.ItemOrigin = new System.Drawing.Point(6, 0);
			valueListItem4.DataValue = 0;
			valueListItem4.DisplayText = "Nombre";
			valueListItem5.DataValue = 1;
			valueListItem5.DisplayText = "Factura";
			valueListItem6.DataValue = 2;
			valueListItem6.DisplayText = "Fecha Fact.";
			valueListItem7.DataValue = 3;
			valueListItem7.DisplayText = "Fecha Pago";
			valueListItem8.DataValue = 4;
			valueListItem8.DisplayText = "Asiento";
			valueListItem9.DataValue = 5;
			valueListItem9.DisplayText = "Código";
			valueListItem10.DataValue = 6;
			valueListItem10.DisplayText = "Cheque";
			valueListItem11.DataValue = 7;
			valueListItem11.DisplayText = "Referencia";
			this.rdgOpcion.Items.Add(valueListItem4);
			this.rdgOpcion.Items.Add(valueListItem5);
			this.rdgOpcion.Items.Add(valueListItem6);
			this.rdgOpcion.Items.Add(valueListItem7);
			this.rdgOpcion.Items.Add(valueListItem8);
			this.rdgOpcion.Items.Add(valueListItem9);
			this.rdgOpcion.Items.Add(valueListItem10);
			this.rdgOpcion.Items.Add(valueListItem11);
			this.rdgOpcion.ItemSpacingVertical = 9;
			this.rdgOpcion.Location = new System.Drawing.Point(16, 17);
			this.rdgOpcion.Name = "rdgOpcion";
			this.rdgOpcion.Size = new System.Drawing.Size(112, 207);
			this.rdgOpcion.TabIndex = 51;
			this.rdgOpcion.ValueChanged += new System.EventHandler(this.rdgOpcion_SelectedIndexChanged);
			// 
			// ultraGrid1
			// 
			this.ultraGrid1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.ultraGrid1.Cursor = System.Windows.Forms.Cursors.Default;
			this.ultraGrid1.DataMember = "Band 0";
			this.ultraGrid1.DataSource = this.udsPago;
			appearance3.BackColor = System.Drawing.Color.White;
			this.ultraGrid1.DisplayLayout.Appearance = appearance3;
			ultraGridBand2.AddButtonCaption = "Cliente";
			ultraGridColumn10.Header.VisiblePosition = 11;
			ultraGridColumn10.Hidden = true;
			ultraGridColumn11.Header.VisiblePosition = 9;
			ultraGridColumn11.Hidden = true;
			ultraGridColumn12.Header.VisiblePosition = 8;
			ultraGridColumn12.Hidden = true;
			ultraGridColumn13.Header.VisiblePosition = 10;
			ultraGridColumn13.Hidden = true;
			ultraGridColumn14.Header.VisiblePosition = 3;
			ultraGridColumn15.Header.VisiblePosition = 2;
			ultraGridColumn16.Header.VisiblePosition = 4;
			ultraGridColumn17.Header.Caption = "Codigo";
			ultraGridColumn17.Header.VisiblePosition = 0;
			ultraGridColumn18.Header.VisiblePosition = 1;
			ultraGridColumn19.Header.VisiblePosition = 6;
			ultraGridColumn20.Header.VisiblePosition = 5;
			ultraGridColumn21.Header.VisiblePosition = 12;
			ultraGridColumn21.Hidden = true;
			ultraGridColumn22.Header.VisiblePosition = 7;
			ultraGridBand2.Columns.AddRange(new object[] {
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
																										 ultraGridColumn22});
			this.ultraGrid1.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			this.ultraGrid1.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
			this.ultraGrid1.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.ultraGrid1.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.False;
			appearance4.BackColor = System.Drawing.Color.Transparent;
			this.ultraGrid1.DisplayLayout.Override.CardAreaAppearance = appearance4;
			this.ultraGrid1.DisplayLayout.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.RowSelect;
			this.ultraGrid1.DisplayLayout.Override.ColumnAutoSizeMode = Infragistics.Win.UltraWinGrid.ColumnAutoSizeMode.AllRowsInBand;
			appearance5.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance5.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance5.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance5.FontData.BoldAsString = "True";
			appearance5.FontData.Name = "Arial";
			appearance5.FontData.SizeInPoints = 10F;
			appearance5.ForeColor = System.Drawing.Color.White;
			appearance5.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.ultraGrid1.DisplayLayout.Override.HeaderAppearance = appearance5;
			this.ultraGrid1.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti;
			appearance6.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance6.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance6.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.RowSelectorAppearance = appearance6;
			appearance7.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance7.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance7.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.SelectedRowAppearance = appearance7;
			this.ultraGrid1.DisplayLayout.Override.SelectTypeCell = Infragistics.Win.UltraWinGrid.SelectType.None;
			this.ultraGrid1.DisplayLayout.Override.SelectTypeCol = Infragistics.Win.UltraWinGrid.SelectType.None;
			this.ultraGrid1.DisplayLayout.Override.SelectTypeRow = Infragistics.Win.UltraWinGrid.SelectType.Single;
			this.ultraGrid1.DisplayLayout.ScrollBounds = Infragistics.Win.UltraWinGrid.ScrollBounds.ScrollToFill;
			this.ultraGrid1.DisplayLayout.ScrollStyle = Infragistics.Win.UltraWinGrid.ScrollStyle.Immediate;
			this.ultraGrid1.DisplayLayout.TabNavigation = Infragistics.Win.UltraWinGrid.TabNavigation.NextControl;
			this.ultraGrid1.DisplayLayout.ViewStyle = Infragistics.Win.UltraWinGrid.ViewStyle.SingleBand;
			this.ultraGrid1.Location = new System.Drawing.Point(16, 241);
			this.ultraGrid1.Name = "ultraGrid1";
			this.ultraGrid1.Size = new System.Drawing.Size(504, 198);
			this.ultraGrid1.TabIndex = 52;
			// 
			// ultraDropDown1
			// 
			this.ultraDropDown1.Cursor = System.Windows.Forms.Cursors.Default;
			this.ultraDropDown1.DataSource = this.cdvCliente;
			ultraGridColumn23.Header.VisiblePosition = 2;
			ultraGridColumn23.Hidden = true;
			ultraGridColumn24.Header.VisiblePosition = 3;
			ultraGridColumn24.Hidden = true;
			ultraGridColumn25.Header.VisiblePosition = 0;
			ultraGridColumn25.Width = 249;
			ultraGridColumn26.Header.VisiblePosition = 4;
			ultraGridColumn26.Hidden = true;
			ultraGridColumn27.Header.VisiblePosition = 5;
			ultraGridColumn27.Hidden = true;
			ultraGridColumn28.Header.VisiblePosition = 1;
			ultraGridColumn28.Hidden = true;
			ultraGridColumn29.Header.VisiblePosition = 6;
			ultraGridColumn30.Header.VisiblePosition = 7;
			ultraGridColumn31.Header.VisiblePosition = 8;
			ultraGridBand3.Columns.AddRange(new object[] {
																										 ultraGridColumn23,
																										 ultraGridColumn24,
																										 ultraGridColumn25,
																										 ultraGridColumn26,
																										 ultraGridColumn27,
																										 ultraGridColumn28,
																										 ultraGridColumn29,
																										 ultraGridColumn30,
																										 ultraGridColumn31});
			this.ultraDropDown1.DisplayLayout.BandsSerializer.Add(ultraGridBand3);
			this.ultraDropDown1.DisplayMember = "Nombre";
			this.ultraDropDown1.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.ultraDropDown1.Location = new System.Drawing.Point(216, 319);
			this.ultraDropDown1.Name = "ultraDropDown1";
			this.ultraDropDown1.Size = new System.Drawing.Size(280, 86);
			this.ultraDropDown1.TabIndex = 53;
			this.ultraDropDown1.Text = "ultraDropDown1";
			this.ultraDropDown1.ValueMember = "idCliente";
			this.ultraDropDown1.Visible = false;
			// 
			// btCancelar
			// 
			this.btCancelar.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.btCancelar.DialogResult = System.Windows.Forms.DialogResult.Cancel;
			this.btCancelar.Location = new System.Drawing.Point(440, 129);
			this.btCancelar.Name = "btCancelar";
			this.btCancelar.Size = new System.Drawing.Size(75, 25);
			this.btCancelar.TabIndex = 57;
			this.btCancelar.Text = "&Cancelar";
			// 
			// btAceptar
			// 
			this.btAceptar.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.btAceptar.Location = new System.Drawing.Point(440, 95);
			this.btAceptar.Name = "btAceptar";
			this.btAceptar.Size = new System.Drawing.Size(75, 25);
			this.btAceptar.TabIndex = 56;
			this.btAceptar.Text = "&Aceptar";
			this.btAceptar.Click += new System.EventHandler(this.btAceptar_Click);
			// 
			// btBuscar
			// 
			this.btBuscar.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.btBuscar.Location = new System.Drawing.Point(440, 60);
			this.btBuscar.Name = "btBuscar";
			this.btBuscar.Size = new System.Drawing.Size(75, 25);
			this.btBuscar.TabIndex = 55;
			this.btBuscar.Text = "&Buscar";
			this.btBuscar.Click += new System.EventHandler(this.btBuscar_Click);
			// 
			// btExcel
			// 
			this.btExcel.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.btExcel.DialogResult = System.Windows.Forms.DialogResult.Cancel;
			this.btExcel.Location = new System.Drawing.Point(440, 172);
			this.btExcel.Name = "btExcel";
			this.btExcel.Size = new System.Drawing.Size(75, 25);
			this.btExcel.TabIndex = 58;
			this.btExcel.Text = "&Excel";
			this.btExcel.Click += new System.EventHandler(this.btExcel_Click);
			// 
			// txtComprob
			// 
			this.txtComprob.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtComprob.Location = new System.Drawing.Point(136, 121);
			this.txtComprob.Name = "txtComprob";
			this.txtComprob.Size = new System.Drawing.Size(128, 22);
			this.txtComprob.TabIndex = 59;
			// 
			// txtReferencia
			// 
			this.txtReferencia.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtReferencia.Enabled = false;
			this.txtReferencia.Location = new System.Drawing.Point(136, 198);
			this.txtReferencia.Name = "txtReferencia";
			this.txtReferencia.Size = new System.Drawing.Size(128, 22);
			this.txtReferencia.TabIndex = 60;
			// 
			// txtDeposito
			// 
			this.txtDeposito.Location = new System.Drawing.Point(288, 155);
			this.txtDeposito.Name = "txtDeposito";
			this.txtDeposito.Size = new System.Drawing.Size(48, 20);
			this.txtDeposito.TabIndex = 61;
			this.txtDeposito.Text = "";
			// 
			// BuscaPagoLotes
			// 
			this.AcceptButton = this.btBuscar;
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 14);
			this.CancelButton = this.btCancelar;
			this.ClientSize = new System.Drawing.Size(536, 445);
			this.Controls.Add(this.txtDeposito);
			this.Controls.Add(this.idBusca);
			this.Controls.Add(this.txtReferencia);
			this.Controls.Add(this.txtComprob);
			this.Controls.Add(this.btExcel);
			this.Controls.Add(this.btCancelar);
			this.Controls.Add(this.btAceptar);
			this.Controls.Add(this.btBuscar);
			this.Controls.Add(this.ultraDropDown1);
			this.Controls.Add(this.ultraGrid1);
			this.Controls.Add(this.rdgOpcion);
			this.Controls.Add(this.cmbNombre);
			this.Controls.Add(this.txtPago);
			this.Controls.Add(this.txtNumIngreso);
			this.Controls.Add(this.txtFechaPago);
			this.Controls.Add(this.txtFecha);
			this.Controls.Add(this.txtFactura);
			this.Controls.Add(this.rdgTipoBusqueda);
			this.MinimizeBox = false;
			this.Name = "BuscaPagoLotes";
			this.Text = "Búsqueda de Pagos";
			this.Load += new System.EventHandler(this.BuscaAsiento_Load);
			((System.ComponentModel.ISupportInitialize)(this.cdvPago)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsPago)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvCliente)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsCliente)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.udsPago)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.rdgTipoBusqueda)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtFactura)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtFecha)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtFechaPago)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNumIngreso)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtPago)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbNombre)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.rdgOpcion)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDropDown1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtComprob)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtReferencia)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void BuscaAsiento_Load(object sender, System.EventArgs e)
		{
			Cursor = Cursors.Default;
			txtDeposito.Width = 0;
			idBusca.Width =0;
			txtFecha.Value = DateTime.Today;
			txtFechaPago.Value = DateTime.Today;
			rdgOpcion.Value = 1;
			if (IdSigno == 1)
				cdvCliente.RowFilter = "Proveedor = 0 Or Ambos = 1";
			else
        cdvCliente.RowFilter = "Proveedor = 1 Or Ambos = 1";
		}

		private void btBuscar_Click(object sender, System.EventArgs e)
		{
			string strFiltro;
			string strAntes = "";
			string strDespues="'";

//		Opciones de comodines
			if ((int) this.rdgTipoBusqueda.Value == 0) strDespues = "%'";
			if ((int) this.rdgTipoBusqueda.Value == 2) strAntes = "%";
			if ((int) this.rdgTipoBusqueda.Value == 2) strDespues = "%'";

			strFiltro = "";
			DateTime dtFecha;
			#region Filtros
			string stBusca = "";
			switch((int) this.rdgOpcion.Value)
			{
				case 0:
					if (this.cmbNombre.Value == null)
					{
						this.errorProvider.SetError(this.cmbNombre, "Ingreses Nombre");
						return;
					}
					this.errorProvider.SetError(this.cmbNombre, "");

					strFiltro = " Compra.idCliente = " + this.cmbNombre.Value.ToString().Trim();
					break;
				case 1:
					if (this.txtFactura.Value == null)
					{
						this.errorProvider.SetError(this.txtFactura, "Ingreses Factura");
						return;
					}
					this.errorProvider.SetError(this.txtFactura, "");
					stBusca = this.txtFactura.Value.ToString().Trim();
					stBusca = stBusca.Replace("'", "''");

					strFiltro = " Compra.Numero LIKE '" + strAntes + stBusca + strDespues;
					break;
				case 2:
					dtFecha = (DateTime) this.txtFecha.Value;
					strFiltro = " Compra.Fecha >= '" + 
						dtFecha.ToString("yyyyMMdd") + "'" + 
						" and Compra.Fecha < '" + 
						dtFecha.AddDays(1).ToString("yyyyMMdd") + "'";
					break;
				case 3:
					dtFecha = (DateTime) this.txtFechaPago.Value;
					strFiltro = " Pago.Fecha >= '" + 
						dtFecha.ToString("yyyyMMdd") + "'" + 
						" and Pago.Fecha < '" + 
						dtFecha.AddDays(1).ToString("yyyyMMdd") + "'";
					break;
				case 4:
					if (this.txtComprob.Value == null)
					{
						this.errorProvider.SetError(this.txtComprob, "Ingrese Comprobante");
						return;
					}
					this.errorProvider.SetError(this.txtComprob, "");
					stBusca = this.txtComprob.Value.ToString().Trim();
					stBusca = stBusca.Replace("'", "''");
					strFiltro = " PagoLote.CodAsiento LIKE '" + strAntes + stBusca + strDespues;
					break;
				case 5:
					if (this.txtNumIngreso.Value == null)
					{
						this.errorProvider.SetError(this.txtNumIngreso, "Ingrese Codigo");
						return;
					}
					this.errorProvider.SetError(this.txtNumIngreso, "");
					stBusca = this.txtNumIngreso.Value.ToString().Trim();
					stBusca = stBusca.Replace("'", "''");
					strFiltro = "PagoLote.Codigo Like '" + strAntes + stBusca + strDespues;
					break;
				case 6:
					if (this.txtPago.Value == null)
					{
						this.errorProvider.SetError(this.txtPago, "Ingrese Cheque");
						return;
					}
					this.errorProvider.SetError(this.txtPago, "");
					stBusca = this.txtPago.Value.ToString().Trim();
					stBusca = stBusca.Replace("'", "''");
					strFiltro = "PagoLote.Cheque Like '" + strAntes + stBusca + strDespues;
					break;
				case 7:
					if (this.txtReferencia.Value == null)
					{
						this.errorProvider.SetError(this.txtReferencia, "Ingrese Referencia");
						return;
					}
					this.errorProvider.SetError(this.txtReferencia, "");
					stBusca = this.txtReferencia.Value.ToString().Trim();
					stBusca = stBusca.Replace("'", "''");
					strFiltro = " PagoLote.Referencia Like '" + strAntes + stBusca + strDespues;
					break;
			}
			#endregion Filtros
			Cursor = Cursors.WaitCursor;
			string stFiltro = strFiltro;

			string stTabla = "Inner Join PagoLote On Pago.idPagoLote = PagoLote.idPagoLote ";
			if (txtDeposito.Text == "d")
				stTabla = "Inner Join PagoLote On Pago.idDeposito = PagoLote.idDeposito ";

			string stSelect = "Select IsNull(Compra.Numero, ''), "
				+ "IsNull(Cliente.Nombre, ''), IsNull(Pago.Pago, 0), "
				+ "IsNull(Compra.idCompra, 0), IsNull(Compra.Fecha, GetDate()), "
				+ "IsNull(Pago.Fecha, GetDate()), IsNull(PagoLote.Codigo, ''), "
				+ "IsNull(PagoLote.Referencia, ''), IsNull(PagoLote.Cheque, ''), "
				+ "IsNull(Pago.Cuenta, ''), IsNull(Pago.Numero, ''), "
				+ "IsNull(Pago.idPagoLote, 0) "
				+ "From Compra Inner Join Cliente On Compra.IdCliente = Cliente.idCliente "
				+ "Inner Join Pago On Pago.idCompra = Compra.idCompra "
				+ stTabla
				+ "Where Compra.idTipoFactura = " + IdSigno.ToString() 
				+ " And Compra.ContadoCredito = 2 "
				+ "And Pago.IdPagoLote > 0 And "
				+ stFiltro;

			if ((int) rdgOpcion.Value == 4 || (int) rdgOpcion.Value == 5
				|| (int) rdgOpcion.Value == 7)
			{
				stSelect = "Select IsNull(Compra.Numero, ''), "
					+ "IsNull(Cliente.Nombre, ''), IsNull(PagoLote.Total, 0), "
					+ "IsNull(Compra.idCompra, 0), IsNull(Compra.Fecha, GetDate()), "
					+ "IsNull(PagoLote.Fecha, GetDate()), IsNull(PagoLote.Codigo, ''), "
					+ "IsNull(PagoLote.Referencia, ''), IsNull(PagoLote.Cheque, ''), "
					+ "IsNull(PagoLote.CodAsiento, ''), IsNull(PagoLote.Referencia, ''), "
					+ "IsNull(PagoLote.idPagoLote, 0) "
					+ "From Compra Inner Join Cliente On Compra.IdCliente = Cliente.idCliente "
					+ "Inner Join Pago On Pago.idCompra = Compra.idCompra "
					+ stTabla
					+ "Where Compra.idTipoFactura = " + IdSigno.ToString() 
					+ " And Compra.ContadoCredito = 2 "
					+ "And PagoLote.IdPagoLote > 0 And "
					+ stFiltro;
			}

			SqlDataReader miReader = Funcion.rEscalarSQL(cdsPago, stSelect, true);
			udsPago.Rows.Clear();
			#region Variables
			int iNumFilas = 0;
			string stNumero;
			string stNombre;
			double dPago;
			int IdCompra;
			DateTime dtFechaC;
			DateTime dtFechaPago;
			string stConcepto;
			string stNumIngreso;
			string stCheque;
			string stCuenta;
			string stNumPago;
			int IdPagoLote;
			#endregion Variables

			while (miReader.Read())
			{
				#region Lectura de Variables
				stNumero = miReader.GetString(0);
				stNombre = miReader.GetString(1);
				dPago = miReader.GetDouble(2);
				IdCompra = miReader.GetInt32(3);
				dtFechaC = miReader.GetDateTime(4);
				dtFechaPago = miReader.GetDateTime(5);
				stConcepto = miReader.GetString(6);
				stNumIngreso = miReader.GetString(7);
				stCheque = miReader.GetString(8);
				stCuenta = miReader.GetString(9);
				stNumPago = miReader.GetString(10);
				IdPagoLote = miReader.GetInt32(11);
				#endregion Lectura de Variables
			
				udsPago.Rows.Add();
				udsPago.Rows[iNumFilas]["Factura"]=stNumero;
				udsPago.Rows[iNumFilas]["Nombre"]=stNombre;
				udsPago.Rows[iNumFilas]["Pago"]=dPago;
				udsPago.Rows[iNumFilas]["Fecha"]=dtFechaC;
				udsPago.Rows[iNumFilas]["FechaPago"]=dtFechaPago;
				udsPago.Rows[iNumFilas]["Concepto"]=stConcepto;
				udsPago.Rows[iNumFilas]["NumIngreso"]=stNumIngreso;
				udsPago.Rows[iNumFilas]["Cheque"]=stCheque;
				udsPago.Rows[iNumFilas]["Cuenta"]=stCuenta;
				udsPago.Rows[iNumFilas]["NumPago"]=stNumPago;
				udsPago.Rows[iNumFilas]["IdPagoLote"]=IdPagoLote;
				iNumFilas ++;
			}
			miReader.Close();

			if (iNumFilas == 1)
			{
				this.ultraGrid1.Select();
				this.DialogResult = DialogResult.OK;
			}
			else
			{
				this.AcceptButton=this.btAceptar;
				this.ultraGrid1.Select();
			}
			Cursor = Cursors.Default;
		}

//		public int IdBuscado()
//		{
//			int iFila = this.gridView1.FocusedRowHandle;
//			return int.Parse(this.cdvSocio[iFila]["idSocio"].ToString());
//		}

		private void btAceptar_Click(object sender, System.EventArgs e)
		{
			this.DialogResult = DialogResult.OK;
		}

		private void btExcel_Click(object sender, System.EventArgs e)
		{
			
		}

		private void cdsCompra_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			this.cdsPago.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);

//			C1.Data.SchemaObjects.DataSetDef dataSetDef = 
//				e.DataSet.Schema.DataSetDefs["dsCompra"];
//
//			string stFiltro = "";
//			if (IdSigno == 1) stFiltro = "";
//			else stFiltro = "";
//			e.Filter.Add(new C1.Data.FilterCondition(dataSetDef.TableViews["Pago"],	stFiltro));
		}

		private void cdsCliente_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			this.cdsCliente.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);
		}

		private void rdgOpcion_SelectedIndexChanged(object sender, System.EventArgs e)
		{
			this.cmbNombre.Enabled = false;
			this.txtFactura.Enabled = false;
			this.txtFecha.Enabled = false;
			this.txtFechaPago.Enabled = false;
			this.txtComprob.Enabled = false;
			this.txtNumIngreso.Enabled = false;
			this.txtPago.Enabled = false;
			this.txtReferencia.Enabled = false;

			switch((int) this.rdgOpcion.Value)
			{
				case 0:
					this.cmbNombre.Enabled = true;
					this.cmbNombre.Select();
					break;
				case 1:
					this.txtFactura.Enabled = true;
					this.txtFactura.Select();
					break;
				case 2:
					this.txtFecha.Enabled = true;
					this.txtFecha.Select();
					break;
				case 3:
					this.txtFechaPago.Enabled = true;
					this.txtFechaPago.Select();
					break;
				case 4:
					this.txtComprob.Enabled = true;
					this.txtComprob.Select();
					break;
				case 5:
					this.txtNumIngreso.Enabled = true;
					this.txtNumIngreso.Select();
					break;
				case 6:
					this.txtPago.Enabled = true;
					this.txtPago.Select();
					break;
				case 7:
					this.txtReferencia.Enabled = true;
					this.txtReferencia.Select();
					break;
			}

		}
	}
}
