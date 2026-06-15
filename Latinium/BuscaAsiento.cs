using System;
using System.Globalization;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de BuscaAsiento.
	/// </summary>
	public class BuscaAsiento : DevExpress.XtraEditors.XtraForm
	{
		public System.Windows.Forms.TextBox idBusca;
		private System.Windows.Forms.ErrorProvider errorProvider;
		private Infragistics.Win.UltraWinEditors.UltraComboEditor cmbOperador;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtNombre;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtNumero;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtCodigo;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo txtFecha;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtTotal;
		private Infragistics.Win.UltraWinEditors.UltraOptionSet rdgTipoBusqueda;
		private Infragistics.Win.UltraWinEditors.UltraOptionSet rdgOpcion;
		private DevExpress.XtraEditors.SimpleButton btCancelar;
		private DevExpress.XtraEditors.SimpleButton btAceptar;
		private DevExpress.XtraEditors.SimpleButton btBuscar;
		public System.Windows.Forms.CheckBox chkOrden;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtCheque;
		private System.Windows.Forms.ToolTip toolTip1;
		private DevExpress.XtraGrid.GridControl grdBusca;
		private DevExpress.XtraGrid.Views.Grid.GridView grvBusca;
		private DevExpress.XtraGrid.Views.Grid.GridView gridView4;
		private C1.Data.C1DataSet cdsSeteo;
		private System.ComponentModel.IContainer components;

		public BuscaAsiento()
		{
			//
			// Necesario para admitir el Diseñador de Windows Forms
			//
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
			this.components = new System.ComponentModel.Container();
			Infragistics.Win.Appearance appearance1 = new Infragistics.Win.Appearance();
			Infragistics.Win.ValueListItem valueListItem1 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem2 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem3 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem4 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem5 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton1 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.ValueListItem valueListItem6 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem7 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem8 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
			Infragistics.Win.ValueListItem valueListItem9 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem10 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem11 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem12 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem13 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem14 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.Appearance appearance11 = new Infragistics.Win.Appearance();
			this.cdsSeteo = new C1.Data.C1DataSet();
			this.idBusca = new System.Windows.Forms.TextBox();
			this.errorProvider = new System.Windows.Forms.ErrorProvider();
			this.cmbOperador = new Infragistics.Win.UltraWinEditors.UltraComboEditor();
			this.txtNombre = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtNumero = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtCodigo = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtFecha = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.txtTotal = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.rdgTipoBusqueda = new Infragistics.Win.UltraWinEditors.UltraOptionSet();
			this.rdgOpcion = new Infragistics.Win.UltraWinEditors.UltraOptionSet();
			this.btCancelar = new DevExpress.XtraEditors.SimpleButton();
			this.btAceptar = new DevExpress.XtraEditors.SimpleButton();
			this.btBuscar = new DevExpress.XtraEditors.SimpleButton();
			this.chkOrden = new System.Windows.Forms.CheckBox();
			this.txtCheque = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.toolTip1 = new System.Windows.Forms.ToolTip(this.components);
			this.grdBusca = new DevExpress.XtraGrid.GridControl();
			this.grvBusca = new DevExpress.XtraGrid.Views.Grid.GridView();
			this.gridView4 = new DevExpress.XtraGrid.Views.Grid.GridView();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbOperador)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNombre)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNumero)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCodigo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtFecha)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtTotal)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.rdgTipoBusqueda)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.rdgOpcion)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCheque)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.grdBusca)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.grvBusca)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.gridView4)).BeginInit();
			this.SuspendLayout();
			// 
			// cdsSeteo
			// 
			this.cdsSeteo.BindingContextCtrl = this;
			this.cdsSeteo.CaseSensitive = false;
			this.cdsSeteo.DataLibrary = "LibContabilidad";
			this.cdsSeteo.DataLibraryUrl = "";
			this.cdsSeteo.DataSetDef = "dsSeteoC";
			this.cdsSeteo.FillOnRequest = false;
			this.cdsSeteo.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsSeteo.Name = "cdsSeteo";
			this.cdsSeteo.SchemaClassName = "LibContabilidad.DataClass";
			this.cdsSeteo.SchemaDef = null;
			this.cdsSeteo.BeforeFill += new C1.Data.FillEventHandler(this.cdsSeteo_BeforeFill);
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
			this.errorProvider.DataMember = "";
			// 
			// cmbOperador
			// 
			appearance1.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbOperador.Appearance = appearance1;
			this.cmbOperador.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbOperador.DropDownStyle = Infragistics.Win.DropDownStyle.DropDownList;
			this.cmbOperador.Enabled = false;
			valueListItem1.DataValue = ">";
			valueListItem1.DisplayText = "Mayor";
			valueListItem2.DataValue = "<";
			valueListItem2.DisplayText = "Menor";
			valueListItem3.DataValue = "=";
			valueListItem3.DisplayText = "Igual";
			valueListItem4.DataValue = ">=";
			valueListItem4.DisplayText = "Mayor o Igual";
			valueListItem5.DataValue = "<=";
			valueListItem5.DisplayText = "Menor o Igual";
			this.cmbOperador.Items.Add(valueListItem1);
			this.cmbOperador.Items.Add(valueListItem2);
			this.cmbOperador.Items.Add(valueListItem3);
			this.cmbOperador.Items.Add(valueListItem4);
			this.cmbOperador.Items.Add(valueListItem5);
			this.cmbOperador.Location = new System.Drawing.Point(136, 95);
			this.cmbOperador.Name = "cmbOperador";
			this.cmbOperador.Size = new System.Drawing.Size(72, 22);
			this.cmbOperador.TabIndex = 27;
			this.cmbOperador.Text = "=";
			// 
			// txtNombre
			// 
			appearance2.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtNombre.Appearance = appearance2;
			this.txtNombre.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNombre.Location = new System.Drawing.Point(136, 17);
			this.txtNombre.Name = "txtNombre";
			this.txtNombre.Size = new System.Drawing.Size(328, 22);
			this.txtNombre.TabIndex = 38;
			// 
			// txtNumero
			// 
			appearance3.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtNumero.Appearance = appearance3;
			this.txtNumero.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNumero.Location = new System.Drawing.Point(136, 69);
			this.txtNumero.Name = "txtNumero";
			this.txtNumero.Size = new System.Drawing.Size(128, 22);
			this.txtNumero.TabIndex = 41;
			// 
			// txtCodigo
			// 
			appearance4.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtCodigo.Appearance = appearance4;
			this.txtCodigo.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtCodigo.Location = new System.Drawing.Point(136, 43);
			this.txtCodigo.Name = "txtCodigo";
			this.txtCodigo.Size = new System.Drawing.Size(128, 22);
			this.txtCodigo.TabIndex = 40;
			// 
			// txtFecha
			// 
			appearance5.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtFecha.Appearance = appearance5;
			this.txtFecha.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton1.Caption = "Today";
			this.txtFecha.DateButtons.Add(dateButton1);
			this.txtFecha.Format = "dd/MMM/yyyy";
			this.txtFecha.Location = new System.Drawing.Point(136, 121);
			this.txtFecha.Name = "txtFecha";
			this.txtFecha.NonAutoSizeHeight = 23;
			this.txtFecha.Size = new System.Drawing.Size(128, 21);
			this.txtFecha.SpinButtonsVisible = true;
			this.txtFecha.TabIndex = 42;
			this.txtFecha.Value = new System.DateTime(2012, 4, 11, 0, 0, 0, 0);
			// 
			// txtTotal
			// 
			appearance6.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtTotal.Appearance = appearance6;
			this.txtTotal.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.txtTotal.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtTotal.FormatString = "#,##0.00";
			this.txtTotal.Location = new System.Drawing.Point(208, 95);
			this.txtTotal.Name = "txtTotal";
			this.txtTotal.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtTotal.PromptChar = ' ';
			this.txtTotal.Size = new System.Drawing.Size(80, 22);
			this.txtTotal.TabIndex = 43;
			// 
			// rdgTipoBusqueda
			// 
			appearance7.ForeColorDisabled = System.Drawing.Color.Black;
			this.rdgTipoBusqueda.Appearance = appearance7;
			this.rdgTipoBusqueda.CheckedIndex = 2;
			this.rdgTipoBusqueda.ItemAppearance = appearance8;
			this.rdgTipoBusqueda.ItemOrigin = new System.Drawing.Point(10, 0);
			valueListItem6.DataValue = 0;
			valueListItem6.DisplayText = "Que Empiece";
			valueListItem7.DataValue = 1;
			valueListItem7.DisplayText = "Exacta";
			valueListItem8.DataValue = 2;
			valueListItem8.DisplayText = "En cualquier parte";
			this.rdgTipoBusqueda.Items.Add(valueListItem6);
			this.rdgTipoBusqueda.Items.Add(valueListItem7);
			this.rdgTipoBusqueda.Items.Add(valueListItem8);
			this.rdgTipoBusqueda.ItemSpacingVertical = 10;
			this.rdgTipoBusqueda.Location = new System.Drawing.Point(328, 43);
			this.rdgTipoBusqueda.Name = "rdgTipoBusqueda";
			this.rdgTipoBusqueda.Size = new System.Drawing.Size(136, 86);
			this.rdgTipoBusqueda.TabIndex = 44;
			this.rdgTipoBusqueda.Text = "En cualquier parte";
			// 
			// rdgOpcion
			// 
			appearance9.ForeColorDisabled = System.Drawing.Color.Black;
			this.rdgOpcion.Appearance = appearance9;
			this.rdgOpcion.CheckedIndex = 0;
			this.rdgOpcion.ItemAppearance = appearance10;
			this.rdgOpcion.ItemOrigin = new System.Drawing.Point(6, 0);
			valueListItem9.DataValue = 0;
			valueListItem9.DisplayText = "Nombre";
			valueListItem10.DataValue = 1;
			valueListItem10.DisplayText = "Código";
			valueListItem11.DataValue = 2;
			valueListItem11.DisplayText = "Número";
			valueListItem12.DataValue = 3;
			valueListItem12.DisplayText = "Total";
			valueListItem13.DataValue = 4;
			valueListItem13.DisplayText = "Fecha";
			valueListItem14.DataValue = 5;
			valueListItem14.DisplayText = "Cheque";
			this.rdgOpcion.Items.Add(valueListItem9);
			this.rdgOpcion.Items.Add(valueListItem10);
			this.rdgOpcion.Items.Add(valueListItem11);
			this.rdgOpcion.Items.Add(valueListItem12);
			this.rdgOpcion.Items.Add(valueListItem13);
			this.rdgOpcion.Items.Add(valueListItem14);
			this.rdgOpcion.ItemSpacingVertical = 10;
			this.rdgOpcion.Location = new System.Drawing.Point(16, 17);
			this.rdgOpcion.Name = "rdgOpcion";
			this.rdgOpcion.Size = new System.Drawing.Size(112, 164);
			this.rdgOpcion.TabIndex = 45;
			this.rdgOpcion.Text = "Nombre";
			this.rdgOpcion.ValueChanged += new System.EventHandler(this.rdgOpcion_SelectedIndexChanged);
			// 
			// btCancelar
			// 
			this.btCancelar.DialogResult = System.Windows.Forms.DialogResult.Cancel;
			this.btCancelar.Location = new System.Drawing.Point(448, 164);
			this.btCancelar.Name = "btCancelar";
			this.btCancelar.Size = new System.Drawing.Size(75, 24);
			this.btCancelar.TabIndex = 48;
			this.btCancelar.Text = "&Cancelar";
			// 
			// btAceptar
			// 
			this.btAceptar.Location = new System.Drawing.Point(368, 164);
			this.btAceptar.Name = "btAceptar";
			this.btAceptar.Size = new System.Drawing.Size(75, 24);
			this.btAceptar.TabIndex = 47;
			this.btAceptar.Text = "&Aceptar";
			this.toolTip1.SetToolTip(this.btAceptar, "Busca Registro Seleccionado / Abre en nueva Ventana");
			this.btAceptar.Click += new System.EventHandler(this.btAceptar_Click);
			this.btAceptar.MouseDown += new System.Windows.Forms.MouseEventHandler(this.btAceptar_MouseDown);
			// 
			// btBuscar
			// 
			this.btBuscar.Location = new System.Drawing.Point(288, 164);
			this.btBuscar.Name = "btBuscar";
			this.btBuscar.Size = new System.Drawing.Size(75, 24);
			this.btBuscar.TabIndex = 46;
			this.btBuscar.Text = "&Buscar";
			this.btBuscar.Click += new System.EventHandler(this.btBuscar_Click);
			// 
			// chkOrden
			// 
			this.chkOrden.Checked = true;
			this.chkOrden.CheckState = System.Windows.Forms.CheckState.Checked;
			this.chkOrden.Location = new System.Drawing.Point(328, 129);
			this.chkOrden.Name = "chkOrden";
			this.chkOrden.Size = new System.Drawing.Size(144, 26);
			this.chkOrden.TabIndex = 61;
			this.chkOrden.Text = "Mantener Orden al Salir";
			// 
			// txtCheque
			// 
			appearance11.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtCheque.Appearance = appearance11;
			this.txtCheque.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtCheque.Location = new System.Drawing.Point(136, 146);
			this.txtCheque.Name = "txtCheque";
			this.txtCheque.Size = new System.Drawing.Size(128, 22);
			this.txtCheque.TabIndex = 62;
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
			this.grdBusca.Location = new System.Drawing.Point(16, 200);
			this.grdBusca.MainView = this.grvBusca;
			this.grdBusca.Name = "grdBusca";
			this.grdBusca.Size = new System.Drawing.Size(512, 256);
			this.grdBusca.TabIndex = 66;
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
			// BuscaAsiento
			// 
			this.AcceptButton = this.btBuscar;
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 14);
			this.CancelButton = this.btCancelar;
			this.ClientSize = new System.Drawing.Size(536, 471);
			this.Controls.Add(this.grdBusca);
			this.Controls.Add(this.txtCheque);
			this.Controls.Add(this.chkOrden);
			this.Controls.Add(this.idBusca);
			this.Controls.Add(this.btCancelar);
			this.Controls.Add(this.btAceptar);
			this.Controls.Add(this.btBuscar);
			this.Controls.Add(this.rdgOpcion);
			this.Controls.Add(this.rdgTipoBusqueda);
			this.Controls.Add(this.txtTotal);
			this.Controls.Add(this.txtFecha);
			this.Controls.Add(this.txtNumero);
			this.Controls.Add(this.txtCodigo);
			this.Controls.Add(this.txtNombre);
			this.Controls.Add(this.cmbOperador);
			this.MinimizeBox = false;
			this.Name = "BuscaAsiento";
			this.Text = "Búsqueda de Asientos";
			this.Load += new System.EventHandler(this.BuscaAsiento_Load);
			this.Activated += new System.EventHandler(this.BuscaAsiento_Activated);
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbOperador)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNombre)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNumero)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCodigo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtFecha)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtTotal)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.rdgTipoBusqueda)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.rdgOpcion)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCheque)).EndInit();
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
			this.rdgOpcion.Value = 1;
		}

		
		CultureInfo us = new CultureInfo("en-US");
		private void btBuscar_Click(object sender, System.EventArgs e)
		{
			try
			{
				string strFiltro;
				string strAntes = "";
				string strDespues="'";

				//		Opciones de comodines
				if ((int) this.rdgTipoBusqueda.Value == 0) strDespues = "%'";
				if ((int) this.rdgTipoBusqueda.Value == 2) strAntes = "%";
				if ((int) this.rdgTipoBusqueda.Value == 2) strDespues = "%'";
				string stOpera = " LIKE '";
				if ((int) this.rdgTipoBusqueda.Value == 1) stOpera = " = '";

				strFiltro = "";
				string stBusca = "";
				switch((int) this.rdgOpcion.Value)
				{
					case 0:
						if (this.txtNombre.Value == null)
						{
							this.errorProvider.SetError(this.txtNombre, "Ingreses Nombre");
							return;
						}
						this.errorProvider.SetError(this.txtNombre, "");

						stBusca = this.txtNombre.Value.ToString().Trim();
						stBusca = stBusca.Replace("'", "''");
						strFiltro = "Descripcion " + stOpera  + strAntes + stBusca + strDespues;
						break;
					case 1:
						if (this.txtCodigo.Value == null)
						{
							this.errorProvider.SetError(this.txtCodigo, "Ingreses Codigo");
							return;
						}
						this.errorProvider.SetError(this.txtCodigo, "");

						stBusca = this.txtCodigo.Value.ToString().Trim();
						stBusca = stBusca.Replace("'", "''");
						strFiltro = "CodAsiento " + stOpera + strAntes + stBusca + strDespues;
						break;
					case 2:
						stBusca = this.txtNumero.Value.ToString().Trim();
						stBusca = stBusca.Replace("'", "''");
						strFiltro = "Numero " + stOpera + strAntes + stBusca + strDespues;
						break;
					case 3:
						if (this.txtTotal.Value == null)
						{
							this.errorProvider.SetError(this.txtTotal, "Ingrese Total a Buscar");
							return;
						}
						this.errorProvider.SetError(this.txtTotal, "");
						double dTotal = double.Parse(this.txtTotal.Value.ToString());

						double.Parse(this.txtTotal.Value.ToString());
						strFiltro = "Total " + this.cmbOperador.Value.ToString() + dTotal.ToString(us).Trim();
						break;
					case 4:
						DateTime dtFecha = (DateTime) this.txtFecha.Value;
						strFiltro = "Fecha >= '" + 
							dtFecha.ToString("yyyyMMdd", us) + "'" + 
							" and Fecha < '" + 
							dtFecha.AddDays(1).ToString("yyyyMMdd", us) + "'";
						break;
					case 5:
						stBusca = this.txtCheque.Value.ToString().Trim();
						stBusca = stBusca.Replace("'", "''");
						strFiltro = "Cheque " + stOpera + strAntes + stBusca + strDespues;
						break;
				}

				strFiltro = "Exec AsientoBuscar '" + strFiltro.Replace("'", "''") + "'";
				grdBusca.DataSource = Funcion.dvProcedimiento(cdsSeteo, strFiltro);

				if (grvBusca.RowCount == 0)
				{
					MessageBox.Show("No existen registros que cumplan la condicion de busqueda",
						"Información", MessageBoxButtons.OK, MessageBoxIcon.Information);
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
			catch(System.Exception ex)
			{
				MessageBox.Show(ex.Message);
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
		
			int IdAsiento = (int) grvBusca.GetRowCellValue(iFila, "idAsiento");
			if (IdAsiento == 0)
			{
				MessageBox.Show("Fila no contiene Asientos",	"Información");
				return;
			}

			idBusca.Text = IdAsiento.ToString();
			GrillaFormato();
			GrillaFormato();
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



		private void btAceptar_Click(object sender, System.EventArgs e)
		{
			IdBuscado();
			if (!chkOrden.Checked) stOrden = "";
			this.DialogResult = DialogResult.OK;
		}

		public string stOrden = "";
		private void rdgOpcion_SelectedIndexChanged(object sender, System.EventArgs e)
		{
			this.txtNombre.Enabled = false;
			this.txtCodigo.Enabled = false;
			this.txtNumero.Enabled = false;
			this.txtTotal.Enabled = false;
			this.txtFecha.Enabled = false;
			this.txtCheque.Enabled = false;
			this.cmbOperador.Enabled = false;

			switch((int) this.rdgOpcion.Value)
			{
				case 0:
					this.txtNombre.Enabled = true;
					this.txtNombre.Select();
					stOrden = "Descripcion";
					break;
				case 1:
					this.txtCodigo.Enabled = true;
					this.txtCodigo.Select();
					stOrden = "CodAsiento";
					break;
				case 2:
					this.txtNumero.Enabled = true;
					this.txtNumero.Select();
					stOrden = "Numero";
					break;
				case 3:
					this.txtTotal.Enabled = true;
					this.cmbOperador.Enabled = true;
					this.txtTotal.Select();
					stOrden = "Total";
					break;
				case 4:
					this.txtFecha.Enabled = true;
					this.txtFecha.Select();
					stOrden = "Fecha";
					break;
				case 5:
					this.txtCheque.Enabled = true;
					this.txtCheque.Select();
					stOrden = "Cheque";
					break;
			}
		}

		private void BuscaAsiento_Activated(object sender, System.EventArgs e)
		{
			if (this.txtNombre.Enabled) 
			{
				this.txtNombre.Focus();
				this.txtNombre.SelectAll();
			}
		}

		private void btAceptar_MouseDown(object sender, System.Windows.Forms.MouseEventArgs e)
		{
			IdBuscado();
			if (e.Button != MouseButtons.Right) return;
			int IdAsiento = int.Parse(idBusca.Text);
			Asiento miAsiento = new Asiento(IdAsiento);
			miAsiento.MdiParent = this.MdiParent;
			miAsiento.Show();
		}

		private void cdsSeteo_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			this.cdsSeteo.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirConta);

		}

		private void grvBusca_DoubleClick(object sender, System.EventArgs e)
		{
			btAceptar_Click(this, e);
		}
	}
}
