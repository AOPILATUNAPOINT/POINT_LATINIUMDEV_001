using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Globalization;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de BuscaAsiento.
	/// </summary>
	public class BuscaCompra : DevExpress.XtraEditors.XtraForm
	{
		public System.Windows.Forms.TextBox idBusca;
		private System.Windows.Forms.ErrorProvider errorProvider;
		private C1.Data.C1DataView cdvCompra;
		private C1.Data.C1DataSet cdsCompra;
		private Infragistics.Win.UltraWinEditors.UltraOptionSet rdgTipoBusqueda;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtCodigo;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo txtFecha;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo txtVencimiento;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo txtEntrega;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtComprobante;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtOtros;
		private Infragistics.Win.UltraWinEditors.UltraOptionSet rdgOpcion;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtTotal;
		private DevExpress.XtraEditors.SimpleButton btCancelar;
		private DevExpress.XtraEditors.SimpleButton btAceptar;
		private DevExpress.XtraEditors.SimpleButton btBuscar;
		private DevExpress.XtraEditors.SimpleButton btExcel;
		public System.Windows.Forms.CheckBox chkOrden;
		private DevExpress.XtraEditors.SimpleButton btVer;
		private DevExpress.XtraEditors.LookUpEdit cmbNombre;
		private C1.Data.C1DataSet cdsCliente;
		private C1.Data.C1DataView cdvCliente;
		private DevExpress.XtraGrid.GridControl grdBusca;
		private DevExpress.XtraGrid.Views.Grid.GridView grvBusca;
		private DevExpress.XtraGrid.Views.Grid.GridView gridView4;
		private System.ComponentModel.IContainer components=null;

		public BuscaCompra()
		{
			InitializeComponent();
		}
		int IdTipoFactura = 1;
		public BuscaCompra(int idTipoF)
		{
			InitializeComponent();
			IdTipoFactura = idTipoF;
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
			Infragistics.Win.ValueListItem valueListItem1 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem2 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem3 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton1 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton2 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton3 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.ValueListItem valueListItem4 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem5 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem6 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem7 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem8 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem9 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem10 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem11 = new Infragistics.Win.ValueListItem();
			this.cdvCompra = new C1.Data.C1DataView();
			this.cdsCompra = new C1.Data.C1DataSet();
			this.idBusca = new System.Windows.Forms.TextBox();
			this.errorProvider = new System.Windows.Forms.ErrorProvider();
			this.rdgTipoBusqueda = new Infragistics.Win.UltraWinEditors.UltraOptionSet();
			this.txtCodigo = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtFecha = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.txtVencimiento = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.txtEntrega = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.txtComprobante = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtOtros = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.rdgOpcion = new Infragistics.Win.UltraWinEditors.UltraOptionSet();
			this.cdvCliente = new C1.Data.C1DataView();
			this.cdsCliente = new C1.Data.C1DataSet();
			this.txtTotal = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.btCancelar = new DevExpress.XtraEditors.SimpleButton();
			this.btAceptar = new DevExpress.XtraEditors.SimpleButton();
			this.btBuscar = new DevExpress.XtraEditors.SimpleButton();
			this.btExcel = new DevExpress.XtraEditors.SimpleButton();
			this.chkOrden = new System.Windows.Forms.CheckBox();
			this.btVer = new DevExpress.XtraEditors.SimpleButton();
			this.cmbNombre = new DevExpress.XtraEditors.LookUpEdit();
			this.grdBusca = new DevExpress.XtraGrid.GridControl();
			this.grvBusca = new DevExpress.XtraGrid.Views.Grid.GridView();
			this.gridView4 = new DevExpress.XtraGrid.Views.Grid.GridView();
			((System.ComponentModel.ISupportInitialize)(this.cdvCompra)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsCompra)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.rdgTipoBusqueda)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCodigo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtFecha)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtVencimiento)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtEntrega)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtComprobante)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtOtros)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.rdgOpcion)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvCliente)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsCliente)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtTotal)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbNombre.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.grdBusca)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.grvBusca)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.gridView4)).BeginInit();
			this.SuspendLayout();
			// 
			// cdvCompra
			// 
			this.cdvCompra.BindingContextCtrl = this;
			this.cdvCompra.DataSet = this.cdsCompra;
			this.cdvCompra.TableName = "";
			this.cdvCompra.TableViewName = "Compra";
			// 
			// cdsCompra
			// 
			this.cdsCompra.BindingContextCtrl = this;
			this.cdsCompra.CaseSensitive = false;
			this.cdsCompra.DataLibrary = "LibFacturacion";
			this.cdsCompra.DataLibraryUrl = "";
			this.cdsCompra.DataSetDef = "dsCompra";
			this.cdsCompra.FillOnRequest = false;
			this.cdsCompra.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsCompra.Name = "cdsCompra";
			this.cdsCompra.SchemaClassName = "LibFacturacion.DataClass";
			this.cdsCompra.SchemaDef = null;
			this.cdsCompra.BeforeFill += new C1.Data.FillEventHandler(this.cdsCompra_BeforeFill);
			// 
			// idBusca
			// 
			this.idBusca.Location = new System.Drawing.Point(480, 17);
			this.idBusca.Name = "idBusca";
			this.idBusca.Size = new System.Drawing.Size(48, 20);
			this.idBusca.TabIndex = 23;
			this.idBusca.Text = "idBusca";
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
			// txtCodigo
			// 
			this.txtCodigo.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtCodigo.Location = new System.Drawing.Point(136, 43);
			this.txtCodigo.Name = "txtCodigo";
			this.txtCodigo.Size = new System.Drawing.Size(128, 22);
			this.txtCodigo.TabIndex = 43;
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
			// txtVencimiento
			// 
			this.txtVencimiento.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton2.Caption = "Today";
			this.txtVencimiento.DateButtons.Add(dateButton2);
			this.txtVencimiento.Format = "dd/MMM/yyyy";
			this.txtVencimiento.Location = new System.Drawing.Point(136, 95);
			this.txtVencimiento.Name = "txtVencimiento";
			this.txtVencimiento.NonAutoSizeHeight = 23;
			this.txtVencimiento.Size = new System.Drawing.Size(128, 21);
			this.txtVencimiento.SpinButtonsVisible = true;
			this.txtVencimiento.TabIndex = 46;
			this.txtVencimiento.Value = new System.DateTime(2005, 5, 17, 0, 0, 0, 0);
			// 
			// txtEntrega
			// 
			this.txtEntrega.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton3.Caption = "Today";
			this.txtEntrega.DateButtons.Add(dateButton3);
			this.txtEntrega.Format = "dd/MMM/yyyy";
			this.txtEntrega.Location = new System.Drawing.Point(136, 121);
			this.txtEntrega.Name = "txtEntrega";
			this.txtEntrega.NonAutoSizeHeight = 23;
			this.txtEntrega.Size = new System.Drawing.Size(128, 21);
			this.txtEntrega.SpinButtonsVisible = true;
			this.txtEntrega.TabIndex = 47;
			this.txtEntrega.Value = new System.DateTime(2005, 5, 17, 0, 0, 0, 0);
			// 
			// txtComprobante
			// 
			this.txtComprobante.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtComprobante.Location = new System.Drawing.Point(136, 146);
			this.txtComprobante.Name = "txtComprobante";
			this.txtComprobante.Size = new System.Drawing.Size(128, 22);
			this.txtComprobante.TabIndex = 48;
			// 
			// txtOtros
			// 
			this.txtOtros.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtOtros.Location = new System.Drawing.Point(136, 172);
			this.txtOtros.Name = "txtOtros";
			this.txtOtros.Size = new System.Drawing.Size(128, 22);
			this.txtOtros.TabIndex = 49;
			// 
			// rdgOpcion
			// 
			this.rdgOpcion.ItemAppearance = appearance2;
			this.rdgOpcion.ItemOrigin = new System.Drawing.Point(6, 0);
			valueListItem4.DataValue = 0;
			valueListItem4.DisplayText = "Nombre";
			valueListItem5.DataValue = 1;
			valueListItem5.DisplayText = "Código";
			valueListItem6.DataValue = 3;
			valueListItem6.DisplayText = "Fecha";
			valueListItem7.DataValue = 4;
			valueListItem7.DisplayText = "Vencimiento";
			valueListItem8.DataValue = 5;
			valueListItem8.DisplayText = "Entrega";
			valueListItem9.DataValue = 6;
			valueListItem9.DisplayText = "Comprobante";
			valueListItem10.DataValue = 7;
			valueListItem10.DisplayText = "Otros";
			valueListItem11.DataValue = 8;
			valueListItem11.DisplayText = "Total";
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
			this.cdsCliente.FillOnRequest = false;
			this.cdsCliente.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsCliente.Name = "cdsCliente";
			this.cdsCliente.SchemaClassName = "LibFacturacion.DataClass";
			this.cdsCliente.SchemaDef = null;
			this.cdsCliente.BeforeFill += new C1.Data.FillEventHandler(this.cdsCliente_BeforeFill);
			// 
			// txtTotal
			// 
			this.txtTotal.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtTotal.FormatString = "#,##0.00";
			this.txtTotal.Location = new System.Drawing.Point(136, 198);
			this.txtTotal.Name = "txtTotal";
			this.txtTotal.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtTotal.PromptChar = ' ';
			this.txtTotal.Size = new System.Drawing.Size(128, 22);
			this.txtTotal.TabIndex = 54;
			// 
			// btCancelar
			// 
			this.btCancelar.DialogResult = System.Windows.Forms.DialogResult.Cancel;
			this.btCancelar.Location = new System.Drawing.Point(472, 128);
			this.btCancelar.Name = "btCancelar";
			this.btCancelar.Size = new System.Drawing.Size(75, 25);
			this.btCancelar.TabIndex = 57;
			this.btCancelar.Text = "&Cancelar";
			// 
			// btAceptar
			// 
			this.btAceptar.Location = new System.Drawing.Point(472, 88);
			this.btAceptar.Name = "btAceptar";
			this.btAceptar.Size = new System.Drawing.Size(75, 25);
			this.btAceptar.TabIndex = 56;
			this.btAceptar.Text = "&Aceptar";
			this.btAceptar.Click += new System.EventHandler(this.btAceptar_Click);
			// 
			// btBuscar
			// 
			this.btBuscar.Location = new System.Drawing.Point(472, 56);
			this.btBuscar.Name = "btBuscar";
			this.btBuscar.Size = new System.Drawing.Size(75, 25);
			this.btBuscar.TabIndex = 55;
			this.btBuscar.Text = "&Buscar";
			this.btBuscar.Click += new System.EventHandler(this.btBuscar_Click);
			// 
			// btExcel
			// 
			this.btExcel.Location = new System.Drawing.Point(472, 192);
			this.btExcel.Name = "btExcel";
			this.btExcel.Size = new System.Drawing.Size(75, 25);
			this.btExcel.TabIndex = 58;
			this.btExcel.Text = "&Excel";
			this.btExcel.Click += new System.EventHandler(this.btExcel_Click);
			// 
			// chkOrden
			// 
			this.chkOrden.Checked = true;
			this.chkOrden.CheckState = System.Windows.Forms.CheckState.Checked;
			this.chkOrden.Location = new System.Drawing.Point(288, 160);
			this.chkOrden.Name = "chkOrden";
			this.chkOrden.Size = new System.Drawing.Size(152, 26);
			this.chkOrden.TabIndex = 60;
			this.chkOrden.Text = "Mantener Orden al Salir";
			// 
			// btVer
			// 
			this.btVer.Location = new System.Drawing.Point(472, 160);
			this.btVer.Name = "btVer";
			this.btVer.Size = new System.Drawing.Size(75, 24);
			this.btVer.TabIndex = 61;
			this.btVer.Text = "&Ver";
			this.btVer.Click += new System.EventHandler(this.btVer_Click);
			// 
			// cmbNombre
			// 
			this.cmbNombre.Location = new System.Drawing.Point(136, 16);
			this.cmbNombre.Name = "cmbNombre";
			// 
			// cmbNombre.Properties
			// 
			this.cmbNombre.Properties.Buttons.AddRange(new DevExpress.XtraEditors.Controls.EditorButton[] {
																																																			new DevExpress.XtraEditors.Controls.EditorButton(DevExpress.XtraEditors.Controls.ButtonPredefines.Combo)});
			this.cmbNombre.Properties.Columns.AddRange(new DevExpress.XtraEditors.Controls.LookUpColumnInfo[] {
																																																					new DevExpress.XtraEditors.Controls.LookUpColumnInfo("Nombre", "Nombre", 49, DevExpress.Utils.FormatType.None, "", true, DevExpress.Utils.HorzAlignment.Near, DevExpress.Data.ColumnSortOrder.Ascending)});
			this.cmbNombre.Properties.DisplayMember = "Nombre";
			this.cmbNombre.Properties.NullText = "";
			this.cmbNombre.Properties.PopupWidth = 300;
			this.cmbNombre.Properties.ShowFooter = false;
			this.cmbNombre.Properties.ShowHeader = false;
			this.cmbNombre.Properties.ValueMember = "idCliente";
			this.cmbNombre.Size = new System.Drawing.Size(288, 20);
			this.cmbNombre.TabIndex = 64;
			// 
			// grdBusca
			// 
			this.grdBusca.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			// 
			// grdBusca.EmbeddedNavigator
			// 
			this.grdBusca.EmbeddedNavigator.Name = "";
			this.grdBusca.Location = new System.Drawing.Point(8, 232);
			this.grdBusca.MainView = this.grvBusca;
			this.grdBusca.Name = "grdBusca";
			this.grdBusca.Size = new System.Drawing.Size(584, 200);
			this.grdBusca.TabIndex = 65;
			this.grdBusca.ViewCollection.AddRange(new DevExpress.XtraGrid.Views.Base.BaseView[] {
																																														this.grvBusca,
																																														this.gridView4});
			// 
			// grvBusca
			// 
			this.grvBusca.GridControl = this.grdBusca;
			this.grvBusca.Name = "grvBusca";
			this.grvBusca.OptionsBehavior.Editable = false;
			this.grvBusca.OptionsSelection.MultiSelect = true;
			this.grvBusca.OptionsSelection.MultiSelectMode = DevExpress.XtraGrid.Views.Grid.GridMultiSelectMode.CellSelect;
			this.grvBusca.OptionsView.ShowAutoFilterRow = true;
			this.grvBusca.DoubleClick += new System.EventHandler(this.grvBusca_DoubleClick);
			// 
			// gridView4
			// 
			this.gridView4.GridControl = this.grdBusca;
			this.gridView4.Name = "gridView4";
			// 
			// BuscaCompra
			// 
			this.AcceptButton = this.btBuscar;
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 14);
			this.CancelButton = this.btCancelar;
			this.ClientSize = new System.Drawing.Size(600, 445);
			this.Controls.Add(this.grdBusca);
			this.Controls.Add(this.cmbNombre);
			this.Controls.Add(this.btVer);
			this.Controls.Add(this.chkOrden);
			this.Controls.Add(this.idBusca);
			this.Controls.Add(this.btExcel);
			this.Controls.Add(this.btCancelar);
			this.Controls.Add(this.btAceptar);
			this.Controls.Add(this.btBuscar);
			this.Controls.Add(this.txtTotal);
			this.Controls.Add(this.rdgOpcion);
			this.Controls.Add(this.txtOtros);
			this.Controls.Add(this.txtComprobante);
			this.Controls.Add(this.txtEntrega);
			this.Controls.Add(this.txtVencimiento);
			this.Controls.Add(this.txtFecha);
			this.Controls.Add(this.txtCodigo);
			this.Controls.Add(this.rdgTipoBusqueda);
			this.MinimizeBox = false;
			this.Name = "BuscaCompra";
			this.Text = "Búsqueda de Facturas";
			this.Load += new System.EventHandler(this.BuscaAsiento_Load);
			((System.ComponentModel.ISupportInitialize)(this.cdvCompra)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsCompra)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.rdgTipoBusqueda)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCodigo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtFecha)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtVencimiento)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtEntrega)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtComprobante)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtOtros)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.rdgOpcion)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvCliente)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsCliente)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtTotal)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbNombre.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.grdBusca)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.grvBusca)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.gridView4)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void BuscaAsiento_Load(object sender, System.EventArgs e)
		{
			Cursor = Cursors.Default;
			this.idBusca.Width =0;
			this.txtFecha.Value = DateTime.Today;
			this.txtEntrega.Value = DateTime.Today;
			this.txtVencimiento.Value = DateTime.Today;
			this.rdgOpcion.Value = 1;
			string stExec = "Exec ClienteCargaNombre " + IdTipoFactura;
			cmbNombre.Properties.DataSource = Funcion.dvProcedimiento(cdsCompra, stExec);
		}


		private void BorraErrores()
		{
			this.errorProvider.SetError(this.cmbNombre, "");
			this.errorProvider.SetError(this.txtCodigo, "");
		}

		private bool Verifica()
		{
			BorraErrores();
			bool bOk = true;
			switch((int) this.rdgOpcion.Value)
			{
				case 0:
					if (this.cmbNombre.EditValue == null)
					{
						this.errorProvider.SetError(this.cmbNombre, "Ingrese Nombre");
						bOk = false;
					}
					break;
				case 1:
					if (this.txtCodigo.Value == null)
					{
						this.errorProvider.SetError(this.txtCodigo, "Ingrese Codigo");
						bOk = false;
					}
					break;
				case 6:
					if (this.txtComprobante.Value == null)
					{
						this.errorProvider.SetError(this.txtComprobante, "Ingrese Comprobante");
						bOk = false;
					}
					break;
				case 7:
					if (this.txtOtros.Value == null)
					{
						this.errorProvider.SetError(this.txtOtros, "Ingrese Otros");
						bOk = false;
					}
					break;
			}
			return bOk;
		}

		private string stBuscaEnResultado = "";
		CultureInfo us = new CultureInfo("en-US");

		private void btBuscar_Click(object sender, System.EventArgs e)
		{
			if (!Verifica()) return;
			string strFiltro;
			string strAntes = "";
			string strDespues="'";

//		Opciones de comodines
			if ((int) this.rdgTipoBusqueda.Value == 0) strDespues = "%'";
			if ((int) this.rdgTipoBusqueda.Value == 2) strAntes = "%";
			if ((int) this.rdgTipoBusqueda.Value == 2) strDespues = "%'";

			strFiltro = "";
			DateTime dtFecha;
			string stBusca = "";
			switch((int) this.rdgOpcion.Value)
			{
				case 0:
					strFiltro = "cl.idCliente = " + this.cmbNombre.EditValue.ToString().Trim();
					break;
				case 1:
					stBusca = this.txtCodigo.Value.ToString().Trim();
					stBusca = stBusca.Replace("'", "''");
					strFiltro = "Numero LIKE '" + strAntes + stBusca + strDespues;
					break;
				case 2:
					break;
				case 3:
					dtFecha = (DateTime) this.txtFecha.Value;
					strFiltro = "Fecha >= '" + 
						dtFecha.ToString("yyyyMMdd") + "'" + 
						" and Fecha < '" + 
						dtFecha.AddDays(1).ToString("yyyyMMdd") + "'";
					break;
				case 4:
					dtFecha = (DateTime) this.txtVencimiento.Value;
					strFiltro = "FechaVencimiento >= '" + 
						dtFecha.ToString("yyyyMMdd") + "'" + 
						" and FechaVencimiento < '" + 
						dtFecha.AddDays(1).ToString("yyyyMMdd") + "'";
					break;
				case 5:
					dtFecha = (DateTime) this.txtEntrega.Value;
					strFiltro = "FechaEntrega >= '" + 
						dtFecha.ToString("yyyyMMdd") + "'" + 
						" and FechaEntrega < '" + 
						dtFecha.AddDays(1).ToString("yyyyMMdd") + "'";
					break;
				case 6:
					stBusca = this.txtComprobante.Value.ToString().Trim();
					stBusca = stBusca.Replace("'", "''");
					strFiltro = "Comprobante LIKE '" + strAntes + stBusca + strDespues;
					break;
				case 7:
					stBusca = this.txtOtros.Value.ToString().Trim();
					stBusca = stBusca.Replace("'", "''");
					strFiltro = "Otro LIKE '" + strAntes + stBusca + strDespues;
					break;
				case 8:
					double dTotal = double.Parse(txtTotal.Value.ToString());
					strFiltro = "Total >= " + (dTotal-.01).ToString("0.00", us) + " And Total < " + (dTotal + 0.01).ToString("0.00", us);
					break;
			}

			strFiltro += " And idTipoFactura = " + IdTipoFactura.ToString();
			strFiltro = "Exec CompraBuscar '" + strFiltro.Replace("'", "''") + "'";
			grdBusca.DataSource = Funcion.dvProcedimiento(cdsCliente, strFiltro);

			if (grvBusca.RowCount == 0)
			{
				MessageBox.Show("No existen registros que cumplan la condicion de busqueda", "Información", MessageBoxButtons.OK, MessageBoxIcon.Information);
				return;
			}

			IdBuscado();
			if (grvBusca.RowCount == 1)
			{
				this.DialogResult = DialogResult.OK;
			}
			else
			{
				this.AcceptButton=this.btAceptar;
				grdBusca.Select();
			}
		}

		public void GrillaFormato()
		{
			DevExpress.XtraGrid.Views.Base.ColumnView View = 
				grdBusca.MainView as DevExpress.XtraGrid.Views.Base.ColumnView;
			foreach (DevExpress.XtraGrid.Columns.GridColumn col in View.Columns)
			{
				if (col.FieldName.StartsWith("id"))
					col.Visible = false;
				if (col.FieldName.StartsWith("Total") || col.FieldName.StartsWith("SubTotal")
					|| col.FieldName.StartsWith("Desc") || col.FieldName.StartsWith("Iva"))
				{
					col.DisplayFormat.FormatString = "n2";
					col.DisplayFormat.FormatType = DevExpress.Utils.FormatType.Numeric;
				}
				if (col.FieldName.StartsWith("Fecha") || col.FieldName.StartsWith("Vencimiento"))
				{
					col.DisplayFormat.FormatString = "dd/MMM/yyyy";
					col.DisplayFormat.FormatType = DevExpress.Utils.FormatType.DateTime;
				}			
			}
		}

		public void IdBuscado()
		{
			int iFila = grvBusca.FocusedRowHandle;
			if (iFila < 0)
			{
				MessageBox.Show("Seleccione la línea que desa ver la Transacción", "Información");
				return;
			}
		
			int IdCompra = (int) grvBusca.GetRowCellValue(iFila, "idCompra");
			if (IdCompra == 0)
			{
				MessageBox.Show("Fila no contiene Transacción Asociados",	"Información");
				return;
			}

			idBusca.Text = IdCompra.ToString();
			GrillaFormato();
			GrillaFormato();
		}

		private void btAceptar_Click(object sender, System.EventArgs e)
		{
			IdBuscado();
			if (!chkOrden.Checked) stOrden = "";
			this.DialogResult = DialogResult.OK;
		}

		private void btExcel_Click(object sender, System.EventArgs e)
		{
			string stDir = MenuLatinium.stDirInicio + "\\BuscaFactura.xls";
			grdBusca.ExportToXls(stDir, true);
			if (DialogResult.Yes == MessageBox.Show(
				"¿Desea Ver Archivo Generado?", 
				"Información", MessageBoxButtons.YesNo, MessageBoxIcon.Question))
				System.Diagnostics.Process.Start(stDir);		
		}

		private void cdsCompra_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			this.cdsCompra.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);
		}

		public string stOrden = "";
		private void rdgOpcion_SelectedIndexChanged(object sender, System.EventArgs e)
		{
			this.cmbNombre.Enabled = false;
			this.txtCodigo.Enabled = false;
//			this.txtNumero.Enabled = false;
			this.txtFecha.Enabled = false;
			this.txtVencimiento.Enabled = false;
			this.txtEntrega.Enabled = false;
			this.txtComprobante.Enabled = false;
			this.txtOtros.Enabled = false;
			this.txtTotal.Enabled = false;

			switch((int) this.rdgOpcion.Value)
			{
				case 0:
					this.cmbNombre.Enabled = true;
					this.cmbNombre.Select();
					stOrden = "idCliente";
					break;
				case 1:
					this.txtCodigo.Enabled = true;
					this.txtCodigo.Select();
					stOrden = "Numero";
					break;
				case 2:
//					this.txtNumero.Enabled = true;
//					this.txtNumero.Select();
					break;
				case 3:
					this.txtFecha.Enabled = true;
					this.txtFecha.Select();
					stOrden = "Fecha";
					break;
				case 4:
					this.txtVencimiento.Enabled = true;
					this.txtVencimiento.Select();
					break;
				case 5:
					this.txtEntrega.Enabled = true;
					this.txtEntrega.Select();
					break;
				case 6:
					this.txtComprobante.Enabled = true;
					this.txtComprobante.Select();
					break;
				case 7:
					this.txtOtros.Enabled = true;
					this.txtOtros.Select();
					break;
				case 8:
					this.txtTotal.Enabled = true;
					this.txtTotal.Select();
					break;
			}

		}

		private void btVer_Click(object sender, System.EventArgs e)
		{
			IdBuscado();
			string stSelect = "Select idTipoFactura From Compra Where idCompra = " + idBusca.Text;
			int idTipo = Funcion.iEscalarSQL(cdsCompra, stSelect, true);
			int idCompraB = int.Parse(idBusca.Text);
//			Compra miCompra = new Compra (idTipo, idCompraB);
//			miCompra.MdiParent = this.MdiParent;
//			miCompra.Show();
		}

		private void cdsCliente_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			this.cdsCliente.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);
		}

		private void grvBusca_DoubleClick(object sender, System.EventArgs e)
		{
			btAceptar_Click(this, e);
		}
	}
}
