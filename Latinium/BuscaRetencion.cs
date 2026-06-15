using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de BuscaAsiento.
	/// </summary>
	public class BuscaRetencion : DevExpress.XtraEditors.XtraForm
	{
		private DevExpress.XtraEditors.RadioGroup rdgOpcion;
		private DevExpress.XtraEditors.RadioGroup rdgTipoBusqueda;
		public System.Windows.Forms.TextBox idBusca;
		private System.Windows.Forms.ErrorProvider errorProvider;
		private DevExpress.XtraEditors.SimpleButton btBuscar;
		private DevExpress.XtraEditors.SimpleButton btAceptar;
		private DevExpress.XtraEditors.SimpleButton btCancelar;
		private C1.Data.C1DataView cdvRetencion;
		private C1.Data.C1DataSet cdsRetencion;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtNombre;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtCodigo;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtNumero;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtTotal;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo txtFecha;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtCheque;
		private Infragistics.Win.UltraWinGrid.UltraGrid ultraGrid1;
		private System.ComponentModel.IContainer components=null;

		public BuscaRetencion()
		{
			InitializeComponent();
		}

		private int IdIngEgr = 0;
		public BuscaRetencion(int idIngEgr)
		{
			InitializeComponent();
			IdIngEgr = idIngEgr;
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
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton1 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.Appearance appearance1 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Retencion", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idRetencion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPagoGrupo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idAsiento");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCuentaCorriente");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idRetRenta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idRetIva");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idRetRenta1");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idRetIva1");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCompra");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idProyecto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn11 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCliente");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn12 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idTipoRetencion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn13 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCuenta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn14 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPagoLote");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn15 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CodigoSri");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn16 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CodigoSri1");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn17 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Fecha");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn18 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Notas");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn19 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cheque");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn20 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("RentaRetenida");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn21 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("IvaRetenido");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn22 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("RentaRetenida1");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn23 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("IvaRetenido1");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn24 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("IvaTotal");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn25 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("IvaTotal1");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn26 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("PorcIva");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn27 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ChequeEfectivo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn28 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("IngresoEgreso");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn29 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FacturaIva0");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn30 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FacturaIva12");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn31 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FacturaIva01");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn32 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FacturaIva121");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn33 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Renta0");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn34 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("TotalFactura");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn35 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Total");
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn36 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("DescIva0");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn37 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("DescIva");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn38 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Confirmado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn39 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Banco");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn40 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("NumFactura");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn41 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn42 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn43 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ContadoCredito");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn44 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("IncluirIva");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn45 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Borrar");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn46 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CodAsiento");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn47 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Beneficiario");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn48 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaIngreso");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn49 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Editar");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn50 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nuevo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn51 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaEditado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn52 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("AutorizaRet");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn53 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("SerieRet");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn54 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Usuario");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn55 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Tarjeta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn56 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("NoAplica");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn57 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaContable");
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			this.cdvRetencion = new C1.Data.C1DataView();
			this.cdsRetencion = new C1.Data.C1DataSet();
			this.rdgOpcion = new DevExpress.XtraEditors.RadioGroup();
			this.rdgTipoBusqueda = new DevExpress.XtraEditors.RadioGroup();
			this.idBusca = new System.Windows.Forms.TextBox();
			this.errorProvider = new System.Windows.Forms.ErrorProvider();
			this.txtFecha = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.btBuscar = new DevExpress.XtraEditors.SimpleButton();
			this.btAceptar = new DevExpress.XtraEditors.SimpleButton();
			this.btCancelar = new DevExpress.XtraEditors.SimpleButton();
			this.txtNombre = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtCodigo = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtNumero = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtTotal = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtCheque = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.ultraGrid1 = new Infragistics.Win.UltraWinGrid.UltraGrid();
			((System.ComponentModel.ISupportInitialize)(this.cdvRetencion)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsRetencion)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.rdgOpcion.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.rdgTipoBusqueda.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtFecha)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNombre)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCodigo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNumero)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtTotal)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCheque)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).BeginInit();
			this.SuspendLayout();
			// 
			// cdvRetencion
			// 
			this.cdvRetencion.BindingContextCtrl = this;
			this.cdvRetencion.DataSet = this.cdsRetencion;
			this.cdvRetencion.TableName = "";
			this.cdvRetencion.TableViewName = "Retencion";
			// 
			// cdsRetencion
			// 
			this.cdsRetencion.BindingContextCtrl = this;
			this.cdsRetencion.CaseSensitive = false;
			this.cdsRetencion.DataLibrary = "LibContabilidad";
			this.cdsRetencion.DataLibraryUrl = "";
			this.cdsRetencion.DataSetDef = "dsRetencion";
			this.cdsRetencion.FillOnRequest = false;
			this.cdsRetencion.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsRetencion.Name = "cdsRetencion";
			this.cdsRetencion.SchemaClassName = "LibContabilidad.DataClass";
			this.cdsRetencion.SchemaDef = null;
			this.cdsRetencion.BeforeFill += new C1.Data.FillEventHandler(this.cdsRetencion_BeforeFill);
			// 
			// rdgOpcion
			// 
			this.rdgOpcion.EditValue = 0;
			this.rdgOpcion.Location = new System.Drawing.Point(16, 9);
			this.rdgOpcion.Name = "rdgOpcion";
			// 
			// rdgOpcion.Properties
			// 
			this.rdgOpcion.Properties.Appearance.BackColor = System.Drawing.Color.Transparent;
			this.rdgOpcion.Properties.Appearance.Options.UseBackColor = true;
			this.rdgOpcion.Properties.Items.AddRange(new DevExpress.XtraEditors.Controls.RadioGroupItem[] {
																																																			new DevExpress.XtraEditors.Controls.RadioGroupItem(0, "Beneficiario"),
																																																			new DevExpress.XtraEditors.Controls.RadioGroupItem(1, "Código Ret."),
																																																			new DevExpress.XtraEditors.Controls.RadioGroupItem(2, "Cód. Asiento"),
																																																			new DevExpress.XtraEditors.Controls.RadioGroupItem(3, "Total"),
																																																			new DevExpress.XtraEditors.Controls.RadioGroupItem(4, "Fecha"),
																																																			new DevExpress.XtraEditors.Controls.RadioGroupItem(5, "Cheque")});
			this.rdgOpcion.Size = new System.Drawing.Size(104, 163);
			this.rdgOpcion.TabIndex = 19;
			this.rdgOpcion.SelectedIndexChanged += new System.EventHandler(this.rdgOpcion_SelectedIndexChanged);
			// 
			// rdgTipoBusqueda
			// 
			this.rdgTipoBusqueda.EditValue = 2;
			this.rdgTipoBusqueda.Location = new System.Drawing.Point(312, 52);
			this.rdgTipoBusqueda.Name = "rdgTipoBusqueda";
			// 
			// rdgTipoBusqueda.Properties
			// 
			this.rdgTipoBusqueda.Properties.Appearance.BackColor = System.Drawing.Color.Transparent;
			this.rdgTipoBusqueda.Properties.Appearance.Options.UseBackColor = true;
			this.rdgTipoBusqueda.Properties.Items.AddRange(new DevExpress.XtraEditors.Controls.RadioGroupItem[] {
																																																						new DevExpress.XtraEditors.Controls.RadioGroupItem(0, "Que empiece "),
																																																						new DevExpress.XtraEditors.Controls.RadioGroupItem(1, "Exacta"),
																																																						new DevExpress.XtraEditors.Controls.RadioGroupItem(2, "En cualquier parte")});
			this.rdgTipoBusqueda.Size = new System.Drawing.Size(152, 86);
			this.rdgTipoBusqueda.TabIndex = 20;
			// 
			// idBusca
			// 
			this.idBusca.DataBindings.Add(new System.Windows.Forms.Binding("Text", this.cdvRetencion, "idRetencion"));
			this.idBusca.Location = new System.Drawing.Point(480, 17);
			this.idBusca.Name = "idBusca";
			this.idBusca.Size = new System.Drawing.Size(48, 20);
			this.idBusca.TabIndex = 23;
			this.idBusca.Text = "";
			// 
			// errorProvider
			// 
			this.errorProvider.ContainerControl = this;
			// 
			// txtFecha
			// 
			this.txtFecha.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton1.Caption = "Today";
			this.txtFecha.DateButtons.Add(dateButton1);
			this.txtFecha.Enabled = false;
			this.txtFecha.Format = "dd/MMM/yyyy";
			this.errorProvider.SetIconAlignment(this.txtFecha, System.Windows.Forms.ErrorIconAlignment.MiddleLeft);
			this.txtFecha.Location = new System.Drawing.Point(136, 121);
			this.txtFecha.Name = "txtFecha";
			this.txtFecha.NonAutoSizeHeight = 23;
			this.txtFecha.Size = new System.Drawing.Size(112, 21);
			this.txtFecha.SpinButtonsVisible = true;
			this.txtFecha.TabIndex = 34;
			this.txtFecha.Value = new System.DateTime(2012, 4, 11, 0, 0, 0, 0);
			// 
			// btBuscar
			// 
			this.btBuscar.ButtonStyle = DevExpress.XtraEditors.Controls.BorderStyles.Office2003;
			this.btBuscar.Location = new System.Drawing.Point(264, 146);
			this.btBuscar.Name = "btBuscar";
			this.btBuscar.Size = new System.Drawing.Size(75, 25);
			this.btBuscar.TabIndex = 27;
			this.btBuscar.Text = "&Buscar";
			this.btBuscar.Click += new System.EventHandler(this.btBuscar_Click);
			// 
			// btAceptar
			// 
			this.btAceptar.ButtonStyle = DevExpress.XtraEditors.Controls.BorderStyles.Office2003;
			this.btAceptar.Location = new System.Drawing.Point(368, 146);
			this.btAceptar.Name = "btAceptar";
			this.btAceptar.Size = new System.Drawing.Size(75, 25);
			this.btAceptar.TabIndex = 28;
			this.btAceptar.Text = "&Aceptar";
			this.btAceptar.Click += new System.EventHandler(this.btAceptar_Click);
			// 
			// btCancelar
			// 
			this.btCancelar.ButtonStyle = DevExpress.XtraEditors.Controls.BorderStyles.Office2003;
			this.btCancelar.DialogResult = System.Windows.Forms.DialogResult.Cancel;
			this.btCancelar.Location = new System.Drawing.Point(448, 146);
			this.btCancelar.Name = "btCancelar";
			this.btCancelar.Size = new System.Drawing.Size(75, 25);
			this.btCancelar.TabIndex = 29;
			this.btCancelar.Text = "&Cancelar";
			// 
			// txtNombre
			// 
			this.txtNombre.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNombre.Location = new System.Drawing.Point(136, 17);
			this.txtNombre.Name = "txtNombre";
			this.txtNombre.Size = new System.Drawing.Size(328, 22);
			this.txtNombre.TabIndex = 30;
			// 
			// txtCodigo
			// 
			this.txtCodigo.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtCodigo.Enabled = false;
			this.txtCodigo.Location = new System.Drawing.Point(136, 43);
			this.txtCodigo.Name = "txtCodigo";
			this.txtCodigo.Size = new System.Drawing.Size(112, 22);
			this.txtCodigo.TabIndex = 31;
			// 
			// txtNumero
			// 
			this.txtNumero.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNumero.Enabled = false;
			this.txtNumero.Location = new System.Drawing.Point(136, 69);
			this.txtNumero.Name = "txtNumero";
			this.txtNumero.Size = new System.Drawing.Size(112, 22);
			this.txtNumero.TabIndex = 32;
			// 
			// txtTotal
			// 
			this.txtTotal.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtTotal.Enabled = false;
			this.txtTotal.FormatString = "#,##0.00";
			this.txtTotal.Location = new System.Drawing.Point(136, 95);
			this.txtTotal.Name = "txtTotal";
			this.txtTotal.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtTotal.PromptChar = ' ';
			this.txtTotal.Size = new System.Drawing.Size(112, 22);
			this.txtTotal.TabIndex = 33;
			// 
			// txtCheque
			// 
			this.txtCheque.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtCheque.Enabled = false;
			this.txtCheque.Location = new System.Drawing.Point(136, 146);
			this.txtCheque.Name = "txtCheque";
			this.txtCheque.Size = new System.Drawing.Size(112, 22);
			this.txtCheque.TabIndex = 35;
			// 
			// ultraGrid1
			// 
			this.ultraGrid1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.ultraGrid1.Cursor = System.Windows.Forms.Cursors.Default;
			this.ultraGrid1.DataSource = this.cdvRetencion;
			appearance1.BackColor = System.Drawing.Color.White;
			this.ultraGrid1.DisplayLayout.Appearance = appearance1;
			ultraGridColumn1.Header.VisiblePosition = 25;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn2.Header.VisiblePosition = 55;
			ultraGridColumn2.Hidden = true;
			ultraGridColumn3.Header.VisiblePosition = 19;
			ultraGridColumn3.Hidden = true;
			ultraGridColumn4.Header.VisiblePosition = 23;
			ultraGridColumn4.Hidden = true;
			ultraGridColumn5.Header.VisiblePosition = 28;
			ultraGridColumn5.Hidden = true;
			ultraGridColumn6.Header.VisiblePosition = 26;
			ultraGridColumn6.Hidden = true;
			ultraGridColumn7.Header.VisiblePosition = 29;
			ultraGridColumn7.Hidden = true;
			ultraGridColumn8.Header.VisiblePosition = 27;
			ultraGridColumn8.Hidden = true;
			ultraGridColumn9.Header.VisiblePosition = 21;
			ultraGridColumn9.Hidden = true;
			ultraGridColumn10.Header.VisiblePosition = 24;
			ultraGridColumn10.Hidden = true;
			ultraGridColumn11.Header.VisiblePosition = 20;
			ultraGridColumn11.Hidden = true;
			ultraGridColumn12.Header.VisiblePosition = 30;
			ultraGridColumn12.Hidden = true;
			ultraGridColumn13.Header.VisiblePosition = 22;
			ultraGridColumn13.Hidden = true;
			ultraGridColumn14.Header.VisiblePosition = 51;
			ultraGridColumn14.Hidden = true;
			ultraGridColumn15.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
			ultraGridColumn15.Header.Caption = "Cod. Sri";
			ultraGridColumn15.Header.VisiblePosition = 6;
			ultraGridColumn15.Width = 73;
			ultraGridColumn16.Header.VisiblePosition = 44;
			ultraGridColumn16.Hidden = true;
			ultraGridColumn17.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
			ultraGridColumn17.Format = "dd/MMM/yyyy";
			ultraGridColumn17.Header.VisiblePosition = 4;
			ultraGridColumn17.Width = 67;
			ultraGridColumn18.Header.VisiblePosition = 38;
			ultraGridColumn18.Hidden = true;
			ultraGridColumn19.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
			ultraGridColumn19.Header.VisiblePosition = 1;
			ultraGridColumn19.Width = 66;
			ultraGridColumn20.Header.VisiblePosition = 41;
			ultraGridColumn20.Hidden = true;
			ultraGridColumn21.Header.VisiblePosition = 33;
			ultraGridColumn21.Hidden = true;
			ultraGridColumn22.Header.VisiblePosition = 42;
			ultraGridColumn22.Hidden = true;
			ultraGridColumn23.Header.VisiblePosition = 34;
			ultraGridColumn23.Hidden = true;
			ultraGridColumn24.Header.VisiblePosition = 35;
			ultraGridColumn24.Hidden = true;
			ultraGridColumn25.Header.VisiblePosition = 36;
			ultraGridColumn25.Hidden = true;
			ultraGridColumn26.Header.VisiblePosition = 39;
			ultraGridColumn26.Hidden = true;
			ultraGridColumn27.Header.VisiblePosition = 10;
			ultraGridColumn27.Hidden = true;
			ultraGridColumn28.Header.VisiblePosition = 32;
			ultraGridColumn28.Hidden = true;
			ultraGridColumn29.Header.VisiblePosition = 15;
			ultraGridColumn29.Hidden = true;
			ultraGridColumn30.Header.VisiblePosition = 17;
			ultraGridColumn30.Hidden = true;
			ultraGridColumn31.Header.VisiblePosition = 16;
			ultraGridColumn31.Hidden = true;
			ultraGridColumn32.Header.VisiblePosition = 18;
			ultraGridColumn32.Hidden = true;
			ultraGridColumn33.Header.VisiblePosition = 40;
			ultraGridColumn33.Hidden = true;
			ultraGridColumn34.Header.VisiblePosition = 43;
			ultraGridColumn34.Hidden = true;
			ultraGridColumn35.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
			appearance2.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn35.CellAppearance = appearance2;
			ultraGridColumn35.Format = "#,##0.00";
			ultraGridColumn35.Header.VisiblePosition = 5;
			ultraGridColumn35.Width = 68;
			ultraGridColumn36.Header.VisiblePosition = 14;
			ultraGridColumn36.Hidden = true;
			ultraGridColumn37.Header.VisiblePosition = 13;
			ultraGridColumn37.Hidden = true;
			ultraGridColumn38.Header.VisiblePosition = 11;
			ultraGridColumn38.Hidden = true;
			ultraGridColumn39.Header.VisiblePosition = 8;
			ultraGridColumn39.Hidden = true;
			ultraGridColumn40.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
			ultraGridColumn40.Header.VisiblePosition = 7;
			ultraGridColumn41.Header.VisiblePosition = 37;
			ultraGridColumn41.Hidden = true;
			ultraGridColumn42.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
			ultraGridColumn42.Header.VisiblePosition = 3;
			ultraGridColumn42.Width = 67;
			ultraGridColumn43.Header.VisiblePosition = 12;
			ultraGridColumn43.Hidden = true;
			ultraGridColumn44.Header.VisiblePosition = 31;
			ultraGridColumn44.Hidden = true;
			ultraGridColumn45.Header.VisiblePosition = 9;
			ultraGridColumn45.Hidden = true;
			ultraGridColumn46.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
			ultraGridColumn46.Header.Caption = "Asiento";
			ultraGridColumn46.Header.VisiblePosition = 2;
			ultraGridColumn46.Width = 62;
			ultraGridColumn47.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
			ultraGridColumn47.Header.VisiblePosition = 0;
			ultraGridColumn48.Header.VisiblePosition = 45;
			ultraGridColumn48.Hidden = true;
			ultraGridColumn49.Header.VisiblePosition = 46;
			ultraGridColumn49.Hidden = true;
			ultraGridColumn50.Header.VisiblePosition = 47;
			ultraGridColumn50.Hidden = true;
			ultraGridColumn51.Header.VisiblePosition = 48;
			ultraGridColumn51.Hidden = true;
			ultraGridColumn52.Header.VisiblePosition = 49;
			ultraGridColumn52.Hidden = true;
			ultraGridColumn53.Header.VisiblePosition = 50;
			ultraGridColumn53.Hidden = true;
			ultraGridColumn54.Header.VisiblePosition = 52;
			ultraGridColumn54.Hidden = true;
			ultraGridColumn55.Header.VisiblePosition = 53;
			ultraGridColumn55.Hidden = true;
			ultraGridColumn56.Header.VisiblePosition = 54;
			ultraGridColumn56.Hidden = true;
			ultraGridColumn57.Header.VisiblePosition = 56;
			ultraGridColumn57.Hidden = true;
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
																										 ultraGridColumn57});
			this.ultraGrid1.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			appearance3.BackColor = System.Drawing.Color.Transparent;
			this.ultraGrid1.DisplayLayout.Override.CardAreaAppearance = appearance3;
			appearance4.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance4.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance4.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance4.FontData.BoldAsString = "True";
			appearance4.FontData.Name = "Arial";
			appearance4.FontData.SizeInPoints = 10F;
			appearance4.ForeColor = System.Drawing.Color.White;
			appearance4.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.ultraGrid1.DisplayLayout.Override.HeaderAppearance = appearance4;
			this.ultraGrid1.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortSingle;
			appearance5.BackColor = System.Drawing.Color.Gainsboro;
			appearance5.BackColor2 = System.Drawing.SystemColors.ControlLight;
			appearance5.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.RowAlternateAppearance = appearance5;
			appearance6.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance6.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance6.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.RowSelectorAppearance = appearance6;
			appearance7.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance7.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance7.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.SelectedRowAppearance = appearance7;
			this.ultraGrid1.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.ultraGrid1.Location = new System.Drawing.Point(8, 181);
			this.ultraGrid1.Name = "ultraGrid1";
			this.ultraGrid1.Size = new System.Drawing.Size(520, 250);
			this.ultraGrid1.TabIndex = 36;
			// 
			// BuscaRetencion
			// 
			this.AcceptButton = this.btBuscar;
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 14);
			this.CancelButton = this.btCancelar;
			this.ClientSize = new System.Drawing.Size(536, 445);
			this.Controls.Add(this.ultraGrid1);
			this.Controls.Add(this.txtCheque);
			this.Controls.Add(this.txtFecha);
			this.Controls.Add(this.txtTotal);
			this.Controls.Add(this.txtNumero);
			this.Controls.Add(this.txtCodigo);
			this.Controls.Add(this.txtNombre);
			this.Controls.Add(this.btCancelar);
			this.Controls.Add(this.btAceptar);
			this.Controls.Add(this.btBuscar);
			this.Controls.Add(this.idBusca);
			this.Controls.Add(this.rdgTipoBusqueda);
			this.Controls.Add(this.rdgOpcion);
			this.MinimizeBox = false;
			this.Name = "BuscaRetencion";
			this.Text = "Búsqueda de Retenciones";
			this.Load += new System.EventHandler(this.BuscaAsiento_Load);
			((System.ComponentModel.ISupportInitialize)(this.cdvRetencion)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsRetencion)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.rdgOpcion.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.rdgTipoBusqueda.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtFecha)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNombre)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCodigo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNumero)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtTotal)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCheque)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void BuscaAsiento_Load(object sender, System.EventArgs e)
		{
			Cursor = Cursors.Default;
			this.idBusca.Width =0;
			this.txtFecha.Value = DateTime.Today;
			this.txtNombre.SelectAll();
		}

		private void btBuscar_Click(object sender, System.EventArgs e)
		{
			string strFiltro;
			string strAntes = "";
			string strDespues="'";

//		Opciones de comodines
			if ((int) this.rdgTipoBusqueda.EditValue == 0) strDespues = "%'";
			if ((int) this.rdgTipoBusqueda.EditValue == 2) 
			{
				strAntes = "%";
				strDespues = "%'";
			}

			strFiltro = "";
			string stBusca = "";
			switch((int) this.rdgOpcion.EditValue)
			{
				case 0:
					if (this.txtNombre.Value == null)
					{
						this.errorProvider.SetError(this.txtNombre, "Ingrese Nombre");
						return;
					}
					this.errorProvider.SetError(this.txtNombre, "");

					stBusca = this.txtNombre.Value.ToString().Trim();
					stBusca = stBusca.Replace("'", "''");
					strFiltro = "Beneficiario LIKE '" + strAntes + stBusca + strDespues;
					break;
				case 1:
					if (this.txtCodigo.Value == null)
					{
						this.errorProvider.SetError(this.txtCodigo, "Ingrese Codigo");
						return;
					}
					this.errorProvider.SetError(this.txtCodigo, "");

					stBusca = this.txtCodigo.Value.ToString().Trim();
					stBusca = stBusca.Replace("'", "''");
					strFiltro = "Codigo LIKE '" + strAntes + stBusca + strDespues;
					break;
				case 2:
					if (this.txtNumero.Value == null)
					{
						this.errorProvider.SetError(this.txtNumero, "Ingrese Codigo de Asiento");
						return;
					}
					this.errorProvider.SetError(this.txtNumero, "");

					stBusca = this.txtNumero.Value.ToString().Trim();
					stBusca = stBusca.Replace("'", "''");
					strFiltro = "CodAsiento LIKE '" + strAntes + stBusca + strDespues;
					break;
				case 3:
					strFiltro = "Total = " + 
						this.txtTotal.Value.ToString().Trim();
					break;
				case 4:
					DateTime dtFecha = (DateTime) this.txtFecha.Value;
					strFiltro = "Fecha >= '" + 
						dtFecha.ToString("yyyyMMdd") + "'"
						+ " and Fecha < '" + 
						dtFecha.AddDays(1).ToString("yyyyMMdd") + "'";
					break;
				case 5:
					if (this.txtCheque.Value == null)
					{
						this.errorProvider.SetError(this.txtCheque, "Ingrese Cheque");
						return;
					}
					this.errorProvider.SetError(this.txtCheque, "");

					stBusca = this.txtCheque.Value.ToString().Trim();
					stBusca = stBusca.Replace("'", "''");

					strFiltro = "Cheque LIKE '" + strAntes + stBusca + strDespues;
					break;
			}

			string stFiltro = strFiltro +
				" And [IngresoEgreso] = " + this.IdIngEgr.ToString().Trim();
			C1.Data.FilterCondition fcPadre = new 
				C1.Data.FilterCondition(this.cdsRetencion, "Retencion", stFiltro);
			C1.Data.FilterConditions fcTotal = new C1.Data.FilterConditions();
			fcTotal.Add(fcPadre);

			this.cdsRetencion.Fill(fcTotal, false);
			if (this.cdvRetencion.Count == 0)
			{
				MessageBox.Show("No existen registros que cumplan la condicion de busqueda", "Información");
				return;
			}	
			if (this.cdvRetencion.Count == 1)
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

//		public int IdBuscado()
//		{
//			int iFila = this.gridView1.FocusedRowHandle;
//			return int.Parse(this.cdvSocio[iFila]["idSocio"].ToString());
//		}

		private void btAceptar_Click(object sender, System.EventArgs e)
		{
			this.DialogResult = DialogResult.OK;
		}

		private void cdsRetencion_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			this.cdsRetencion.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirRetenc);
		}

		private void rdgOpcion_SelectedIndexChanged(object sender, System.EventArgs e)
		{
			this.txtNombre.Enabled = false;
			this.txtCodigo.Enabled = false;
			this.txtNumero.Enabled = false;
			this.txtTotal.Enabled = false;
			this.txtFecha.Enabled = false;
			this.txtCheque.Enabled = false;

			switch((int) this.rdgOpcion.EditValue)
			{
				case 0:
					this.txtNombre.Enabled = true;
					this.txtNombre.Focus();
					this.txtNombre.SelectAll();
					break;
				case 1:
					this.txtCodigo.Enabled = true;
					this.txtCodigo.Focus();
					this.txtCodigo.SelectAll();
					break;
				case 2:
					this.txtNumero.Enabled = true;
					this.txtNumero.Focus();
					this.txtNumero.SelectAll();
					break;
				case 3:
					this.txtTotal.Enabled = true;
					this.txtTotal.Focus();
					this.txtTotal.SelectAll();
					break;
				case 4:
					this.txtFecha.Enabled = true;
					this.txtFecha.Focus();
					this.txtFecha.Select();
					break;
				case 5:
					this.txtCheque.Enabled = true;
					this.txtCheque.Focus();
					this.txtCheque.Select();
					break;
			}

		}
	}
}
