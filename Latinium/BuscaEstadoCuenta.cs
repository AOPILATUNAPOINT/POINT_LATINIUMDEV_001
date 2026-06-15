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
	public class BuscaEstadoCuenta : DevExpress.XtraEditors.XtraForm
	{
		public System.Windows.Forms.TextBox idBusca;
		private System.Windows.Forms.ErrorProvider errorProvider;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtNombre;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtCodigo;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtNumero;
		private Infragistics.Win.UltraWinEditors.UltraOptionSet rdgOpcion;
		private Infragistics.Win.UltraWinEditors.UltraOptionSet rdgTipoBusqueda;
		private DevExpress.XtraEditors.SimpleButton btBuscar;
		private DevExpress.XtraEditors.SimpleButton btAceptar;
		private DevExpress.XtraEditors.SimpleButton btCancelar;
		private C1.Data.C1DataSet cdsEstadoCuenta;
		private Infragistics.Win.UltraWinGrid.UltraGrid ultraGrid1;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo cmbFecha;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtCheque;
		private C1.Data.C1DataView cdvEstadoCuenta;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtValor;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtDepositos;
		private System.ComponentModel.IContainer components=null;

		public BuscaEstadoCuenta()
		{
			InitializeComponent();
		}

		int IdCuenta = 0;
		public BuscaEstadoCuenta(int idCuenta)
		{
			InitializeComponent();
			IdCuenta = idCuenta;
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
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.ValueListItem valueListItem1 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem2 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem3 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem4 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem5 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem6 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem7 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.ValueListItem valueListItem8 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem9 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem10 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("EstadoCuenta", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idDetAsiento");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Confirmado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Debe");
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Haber");
			Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CodAsiento");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Descripcion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Fecha");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Numero");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cheque");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaConf");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn11 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCuenta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn12 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CentroCosto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn13 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CentroCosto_1");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn14 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("TotalH");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn15 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idAsiento");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn16 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaEditado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn17 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Borrar");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn18 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaRef");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn19 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idSucursal");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn20 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idSucursal_1");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn21 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Deposito");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn22 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("fechaComp");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn23 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn24 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Descripcion_1");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn25 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCiudad");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn26 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CentroCostoBR");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn27 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ProyectoBR");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn28 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idNoDeducible");
			Infragistics.Win.Appearance appearance11 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance12 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance13 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance14 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance15 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance16 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance17 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton1 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.Appearance appearance18 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance19 = new Infragistics.Win.Appearance();
			this.idBusca = new System.Windows.Forms.TextBox();
			this.cdsEstadoCuenta = new C1.Data.C1DataSet();
			this.errorProvider = new System.Windows.Forms.ErrorProvider();
			this.txtNombre = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtCodigo = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtNumero = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.rdgOpcion = new Infragistics.Win.UltraWinEditors.UltraOptionSet();
			this.rdgTipoBusqueda = new Infragistics.Win.UltraWinEditors.UltraOptionSet();
			this.btBuscar = new DevExpress.XtraEditors.SimpleButton();
			this.btAceptar = new DevExpress.XtraEditors.SimpleButton();
			this.btCancelar = new DevExpress.XtraEditors.SimpleButton();
			this.ultraGrid1 = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.txtCheque = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.cmbFecha = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.cdvEstadoCuenta = new C1.Data.C1DataView();
			this.txtValor = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtDepositos = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			((System.ComponentModel.ISupportInitialize)(this.cdsEstadoCuenta)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNombre)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCodigo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNumero)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.rdgOpcion)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.rdgTipoBusqueda)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCheque)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbFecha)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvEstadoCuenta)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtValor)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDepositos)).BeginInit();
			this.SuspendLayout();
			// 
			// idBusca
			// 
			this.idBusca.DataBindings.Add(new System.Windows.Forms.Binding("Text", this.cdsEstadoCuenta, "EstadoCuenta.idDetAsiento"));
			this.idBusca.Location = new System.Drawing.Point(480, 0);
			this.idBusca.Name = "idBusca";
			this.idBusca.Size = new System.Drawing.Size(48, 20);
			this.idBusca.TabIndex = 23;
			this.idBusca.Text = "";
			// 
			// cdsEstadoCuenta
			// 
			this.cdsEstadoCuenta.BindingContextCtrl = this;
			this.cdsEstadoCuenta.DataLibrary = "LibContabilidad";
			this.cdsEstadoCuenta.DataLibraryUrl = "";
			this.cdsEstadoCuenta.DataSetDef = "dsEstadoCuenta";
			this.cdsEstadoCuenta.FillOnRequest = false;
			this.cdsEstadoCuenta.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsEstadoCuenta.Name = "cdsEstadoCuenta";
			this.cdsEstadoCuenta.SchemaClassName = "LibContabilidad.DataClass";
			this.cdsEstadoCuenta.SchemaDef = null;
			// 
			// errorProvider
			// 
			this.errorProvider.ContainerControl = this;
			this.errorProvider.DataMember = "";
			// 
			// txtNombre
			// 
			appearance1.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtNombre.Appearance = appearance1;
			this.txtNombre.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNombre.Location = new System.Drawing.Point(136, 17);
			this.txtNombre.Name = "txtNombre";
			this.txtNombre.Size = new System.Drawing.Size(296, 22);
			this.txtNombre.TabIndex = 37;
			// 
			// txtCodigo
			// 
			appearance2.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtCodigo.Appearance = appearance2;
			this.txtCodigo.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtCodigo.Location = new System.Drawing.Point(136, 43);
			this.txtCodigo.Name = "txtCodigo";
			this.txtCodigo.Size = new System.Drawing.Size(128, 22);
			this.txtCodigo.TabIndex = 38;
			// 
			// txtNumero
			// 
			appearance3.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtNumero.Appearance = appearance3;
			this.txtNumero.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNumero.Location = new System.Drawing.Point(136, 95);
			this.txtNumero.Name = "txtNumero";
			this.txtNumero.Size = new System.Drawing.Size(128, 22);
			this.txtNumero.TabIndex = 39;
			// 
			// rdgOpcion
			// 
			appearance4.ForeColorDisabled = System.Drawing.Color.Black;
			this.rdgOpcion.Appearance = appearance4;
			this.rdgOpcion.ItemAppearance = appearance5;
			this.rdgOpcion.ItemOrigin = new System.Drawing.Point(6, 0);
			valueListItem1.DataValue = 0;
			valueListItem1.DisplayText = "Descripción";
			valueListItem2.DataValue = 1;
			valueListItem2.DisplayText = "Código";
			valueListItem3.DataValue = 2;
			valueListItem3.DisplayText = "Cheque";
			valueListItem4.DataValue = 3;
			valueListItem4.DisplayText = "Número";
			valueListItem5.DataValue = 4;
			valueListItem5.DisplayText = "Fecha";
			valueListItem6.DataValue = 5;
			valueListItem6.DisplayText = "Valor";
			valueListItem7.DataValue = 6;
			valueListItem7.DisplayText = "Depósito";
			this.rdgOpcion.Items.Add(valueListItem1);
			this.rdgOpcion.Items.Add(valueListItem2);
			this.rdgOpcion.Items.Add(valueListItem3);
			this.rdgOpcion.Items.Add(valueListItem4);
			this.rdgOpcion.Items.Add(valueListItem5);
			this.rdgOpcion.Items.Add(valueListItem6);
			this.rdgOpcion.Items.Add(valueListItem7);
			this.rdgOpcion.ItemSpacingVertical = 9;
			this.rdgOpcion.Location = new System.Drawing.Point(16, 17);
			this.rdgOpcion.Name = "rdgOpcion";
			this.rdgOpcion.Size = new System.Drawing.Size(112, 190);
			this.rdgOpcion.TabIndex = 40;
			this.rdgOpcion.ValueChanged += new System.EventHandler(this.rdgOpcion_SelectedIndexChanged);
			// 
			// rdgTipoBusqueda
			// 
			appearance6.ForeColorDisabled = System.Drawing.Color.Black;
			this.rdgTipoBusqueda.Appearance = appearance6;
			this.rdgTipoBusqueda.CheckedIndex = 2;
			this.rdgTipoBusqueda.ItemAppearance = appearance7;
			this.rdgTipoBusqueda.ItemOrigin = new System.Drawing.Point(10, 0);
			valueListItem8.DataValue = 0;
			valueListItem8.DisplayText = "Que Empiece";
			valueListItem9.DataValue = 1;
			valueListItem9.DisplayText = "Exacta";
			valueListItem10.DataValue = 2;
			valueListItem10.DisplayText = "En cualquier parte";
			this.rdgTipoBusqueda.Items.Add(valueListItem8);
			this.rdgTipoBusqueda.Items.Add(valueListItem9);
			this.rdgTipoBusqueda.Items.Add(valueListItem10);
			this.rdgTipoBusqueda.ItemSpacingVertical = 10;
			this.rdgTipoBusqueda.Location = new System.Drawing.Point(296, 52);
			this.rdgTipoBusqueda.Name = "rdgTipoBusqueda";
			this.rdgTipoBusqueda.Size = new System.Drawing.Size(136, 86);
			this.rdgTipoBusqueda.TabIndex = 41;
			this.rdgTipoBusqueda.Text = "En cualquier parte";
			// 
			// btBuscar
			// 
			this.btBuscar.Location = new System.Drawing.Point(448, 26);
			this.btBuscar.Name = "btBuscar";
			this.btBuscar.Size = new System.Drawing.Size(75, 25);
			this.btBuscar.TabIndex = 42;
			this.btBuscar.Text = "&Buscar";
			this.btBuscar.Click += new System.EventHandler(this.btBuscar_Click);
			// 
			// btAceptar
			// 
			this.btAceptar.Location = new System.Drawing.Point(448, 60);
			this.btAceptar.Name = "btAceptar";
			this.btAceptar.Size = new System.Drawing.Size(75, 25);
			this.btAceptar.TabIndex = 43;
			this.btAceptar.Text = "&Aceptar";
			this.btAceptar.Click += new System.EventHandler(this.btAceptar_Click);
			// 
			// btCancelar
			// 
			this.btCancelar.DialogResult = System.Windows.Forms.DialogResult.Cancel;
			this.btCancelar.Location = new System.Drawing.Point(448, 95);
			this.btCancelar.Name = "btCancelar";
			this.btCancelar.Size = new System.Drawing.Size(75, 25);
			this.btCancelar.TabIndex = 44;
			this.btCancelar.Text = "&Cancelar";
			// 
			// ultraGrid1
			// 
			this.ultraGrid1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.ultraGrid1.Cursor = System.Windows.Forms.Cursors.Default;
			this.ultraGrid1.DataMember = "EstadoCuenta";
			this.ultraGrid1.DataSource = this.cdsEstadoCuenta;
			appearance8.BackColor = System.Drawing.Color.White;
			this.ultraGrid1.DisplayLayout.Appearance = appearance8;
			ultraGridColumn1.Header.VisiblePosition = 15;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn2.Header.Caption = "Conf.";
			ultraGridColumn2.Header.VisiblePosition = 16;
			ultraGridColumn2.Hidden = true;
			ultraGridColumn2.Width = 47;
			ultraGridColumn3.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
			appearance9.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn3.CellAppearance = appearance9;
			ultraGridColumn3.Format = "#,##0.00";
			ultraGridColumn3.Header.VisiblePosition = 6;
			ultraGridColumn3.Width = 73;
			ultraGridColumn4.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
			appearance10.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn4.CellAppearance = appearance10;
			ultraGridColumn4.Format = "#,##0.00";
			ultraGridColumn4.Header.VisiblePosition = 7;
			ultraGridColumn4.Width = 73;
			ultraGridColumn5.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
			ultraGridColumn5.Header.Caption = "Código";
			ultraGridColumn5.Header.VisiblePosition = 2;
			ultraGridColumn5.Width = 89;
			ultraGridColumn6.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
			ultraGridColumn6.Header.Caption = "Descripción";
			ultraGridColumn6.Header.VisiblePosition = 4;
			ultraGridColumn7.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
			ultraGridColumn7.Format = "dd/MMM/yyyy";
			ultraGridColumn7.Header.VisiblePosition = 1;
			ultraGridColumn8.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
			ultraGridColumn8.Header.VisiblePosition = 5;
			ultraGridColumn9.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
			ultraGridColumn9.Header.VisiblePosition = 3;
			ultraGridColumn9.Width = 88;
			ultraGridColumn10.Format = "dd/MMM/yyyy";
			ultraGridColumn10.Header.Caption = "F. Confirma";
			ultraGridColumn10.Header.VisiblePosition = 0;
			ultraGridColumn11.Header.VisiblePosition = 11;
			ultraGridColumn11.Hidden = true;
			ultraGridColumn12.Header.VisiblePosition = 10;
			ultraGridColumn12.Hidden = true;
			ultraGridColumn13.Header.VisiblePosition = 12;
			ultraGridColumn13.Hidden = true;
			ultraGridColumn14.Header.VisiblePosition = 13;
			ultraGridColumn14.Hidden = true;
			ultraGridColumn15.Header.VisiblePosition = 14;
			ultraGridColumn15.Hidden = true;
			ultraGridColumn16.Header.VisiblePosition = 19;
			ultraGridColumn16.Hidden = true;
			ultraGridColumn17.Header.VisiblePosition = 8;
			ultraGridColumn18.Header.VisiblePosition = 20;
			ultraGridColumn18.Hidden = true;
			ultraGridColumn19.Header.VisiblePosition = 17;
			ultraGridColumn19.Hidden = true;
			ultraGridColumn20.Header.VisiblePosition = 18;
			ultraGridColumn20.Hidden = true;
			ultraGridColumn21.Header.VisiblePosition = 9;
			ultraGridColumn22.Header.VisiblePosition = 21;
			ultraGridColumn22.Hidden = true;
			ultraGridColumn23.Header.VisiblePosition = 22;
			ultraGridColumn24.Header.VisiblePosition = 23;
			ultraGridColumn25.Header.VisiblePosition = 24;
			ultraGridColumn26.Header.VisiblePosition = 25;
			ultraGridColumn27.Header.VisiblePosition = 26;
			ultraGridColumn28.Header.VisiblePosition = 27;
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
																										 ultraGridColumn28});
			this.ultraGrid1.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.ultraGrid1.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
			this.ultraGrid1.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.ultraGrid1.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.False;
			appearance11.BackColor = System.Drawing.Color.Transparent;
			this.ultraGrid1.DisplayLayout.Override.CardAreaAppearance = appearance11;
			appearance12.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance12.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance12.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance12.FontData.BoldAsString = "True";
			appearance12.FontData.Name = "Arial";
			appearance12.FontData.SizeInPoints = 10F;
			appearance12.ForeColor = System.Drawing.Color.White;
			appearance12.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.ultraGrid1.DisplayLayout.Override.HeaderAppearance = appearance12;
			this.ultraGrid1.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortSingle;
			appearance13.BackColor = System.Drawing.Color.Gainsboro;
			appearance13.BackColor2 = System.Drawing.SystemColors.ControlLight;
			appearance13.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.RowAlternateAppearance = appearance13;
			appearance14.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance14.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance14.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.RowSelectorAppearance = appearance14;
			appearance15.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance15.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance15.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.SelectedRowAppearance = appearance15;
			this.ultraGrid1.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.ultraGrid1.Location = new System.Drawing.Point(16, 215);
			this.ultraGrid1.Name = "ultraGrid1";
			this.ultraGrid1.Size = new System.Drawing.Size(512, 224);
			this.ultraGrid1.TabIndex = 45;
			// 
			// txtCheque
			// 
			appearance16.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtCheque.Appearance = appearance16;
			this.txtCheque.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtCheque.Location = new System.Drawing.Point(136, 69);
			this.txtCheque.Name = "txtCheque";
			this.txtCheque.Size = new System.Drawing.Size(128, 22);
			this.txtCheque.TabIndex = 46;
			// 
			// cmbFecha
			// 
			appearance17.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbFecha.Appearance = appearance17;
			dateButton1.Caption = "Today";
			this.cmbFecha.DateButtons.Add(dateButton1);
			this.cmbFecha.Format = "dd/MMM/yyyy";
			this.cmbFecha.Location = new System.Drawing.Point(136, 121);
			this.cmbFecha.Name = "cmbFecha";
			this.cmbFecha.NonAutoSizeHeight = 24;
			this.cmbFecha.Size = new System.Drawing.Size(128, 21);
			this.cmbFecha.SpinButtonsVisible = true;
			this.cmbFecha.TabIndex = 47;
			this.cmbFecha.Value = new System.DateTime(2012, 4, 11, 0, 0, 0, 0);
			// 
			// cdvEstadoCuenta
			// 
			this.cdvEstadoCuenta.BindingContextCtrl = this;
			this.cdvEstadoCuenta.DataSet = this.cdsEstadoCuenta;
			this.cdvEstadoCuenta.TableName = "";
			this.cdvEstadoCuenta.TableViewName = "EstadoCuenta";
			// 
			// txtValor
			// 
			appearance18.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtValor.Appearance = appearance18;
			this.txtValor.FormatString = "#,##0.00";
			this.txtValor.Location = new System.Drawing.Point(136, 146);
			this.txtValor.MaskInput = "{LOC}nn,nnn,nnn.nn";
			this.txtValor.Name = "txtValor";
			this.txtValor.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtValor.PromptChar = ' ';
			this.txtValor.Size = new System.Drawing.Size(128, 22);
			this.txtValor.SpinButtonDisplayStyle = Infragistics.Win.ButtonDisplayStyle.Always;
			this.txtValor.TabIndex = 48;
			// 
			// txtDepositos
			// 
			appearance19.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtDepositos.Appearance = appearance19;
			this.txtDepositos.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtDepositos.Location = new System.Drawing.Point(136, 172);
			this.txtDepositos.Name = "txtDepositos";
			this.txtDepositos.Size = new System.Drawing.Size(128, 22);
			this.txtDepositos.TabIndex = 49;
			// 
			// BuscaEstadoCuenta
			// 
			this.AcceptButton = this.btBuscar;
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 14);
			this.CancelButton = this.btCancelar;
			this.ClientSize = new System.Drawing.Size(544, 445);
			this.Controls.Add(this.txtDepositos);
			this.Controls.Add(this.txtValor);
			this.Controls.Add(this.cmbFecha);
			this.Controls.Add(this.txtCheque);
			this.Controls.Add(this.ultraGrid1);
			this.Controls.Add(this.btCancelar);
			this.Controls.Add(this.btAceptar);
			this.Controls.Add(this.btBuscar);
			this.Controls.Add(this.rdgTipoBusqueda);
			this.Controls.Add(this.rdgOpcion);
			this.Controls.Add(this.txtNumero);
			this.Controls.Add(this.txtCodigo);
			this.Controls.Add(this.txtNombre);
			this.Controls.Add(this.idBusca);
			this.MinimizeBox = false;
			this.Name = "BuscaEstadoCuenta";
			this.Text = "Búsqueda de Estado de Cuenta";
			this.Resize += new System.EventHandler(this.BuscaArticulo_Resize);
			this.Load += new System.EventHandler(this.BuscaAsiento_Load);
			((System.ComponentModel.ISupportInitialize)(this.cdsEstadoCuenta)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNombre)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCodigo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNumero)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.rdgOpcion)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.rdgTipoBusqueda)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCheque)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbFecha)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvEstadoCuenta)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtValor)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDepositos)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void BuscaAsiento_Load(object sender, System.EventArgs e)
		{
			Cursor = Cursors.Default;
			this.idBusca.Width =0;
			this.rdgOpcion.Value = 0;
		}

		private bool Verifica()
		{
			this.errorProvider.SetError(this.txtNombre, "");
			this.errorProvider.SetError(this.txtCodigo, "");
			this.errorProvider.SetError(this.txtNumero, "");
			this.errorProvider.SetError(this.txtCheque, "");
			this.errorProvider.SetError(this.cmbFecha, "");

			bool bOk = true;
			int iTipo = (int) this.rdgOpcion.Value;

			if (iTipo == 0)
				if (this.txtNombre.Text.Trim().Length == 0)
				{
					this.errorProvider.SetError(this.txtNombre, "Ingrese  Descripcion");
					bOk = false;
				}

			if (iTipo == 1)
				if (this.txtCodigo.Text.Trim().Length == 0)
				{
					this.errorProvider.SetError(this.txtCodigo, "Ingrese Codigo");
					bOk = false;
				}
			if (iTipo == 2)
				if (this.txtCheque.Value == null)
				{
					this.errorProvider.SetError(this.txtCheque, "Ingrese Cheque");
					bOk = false;
				}
			if (iTipo == 3)
				if (this.txtNumero.Text.Trim().Length == 0)
				{
					this.errorProvider.SetError(this.txtNumero, "Ingrese Numero");
					bOk = false;
				}
			if (iTipo == 4)
				if (this.cmbFecha.Value == null)
				{
					this.errorProvider.SetError(this.cmbFecha, "Ingrese Fecha");
					bOk = false;
				}
			return bOk;
		}

		CultureInfo us = new CultureInfo("en-US");
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

			strFiltro = "idCuenta = " + IdCuenta.ToString();
			string stBusca = "";
			switch((int) this.rdgOpcion.Value)
			{
				case 0:
					stBusca = txtNombre.Value.ToString().Trim();
					stBusca = stBusca.Replace("'", "''");
					strFiltro += " And Descripcion LIKE '" + strAntes + stBusca + strDespues;
					break;
				case 1:
					stBusca = txtCodigo.Value.ToString().Trim();
					stBusca = stBusca.Replace("'", "''");
					strFiltro += " And CodAsiento LIKE '" + strAntes + stBusca + strDespues;
					break;
				case 2:
					stBusca = txtCheque.Value.ToString().Trim();
					stBusca = stBusca.Replace("'", "''");
					strFiltro += " And Cheque LIKE '" + strAntes + stBusca + strDespues;
					break;
				case 3:
					stBusca = txtNumero.Value.ToString().Trim();
					stBusca = stBusca.Replace("'", "''");
					strFiltro += " And Numero LIKE '" + strAntes + stBusca + strDespues;
					break;
				case 4:
					DateTime dtIni = (DateTime) cmbFecha.Value;
					strFiltro += " And Fecha >= '" + dtIni.ToString("yyyyMMdd")
						+ "' And Fecha < '" + dtIni.AddDays(1).ToString("yyyyMMdd") + "'";
					break;
				case 5:
					double dValor = (double) txtValor.Value;
					strFiltro += " And (Round(Debe, 2) = " + dValor.ToString("0.00", us)
						+ " Or Round(Haber, 2) = " + dValor.ToString("0.00", us) + ")";
					break;
				case 6:
					stBusca = txtDepositos.Value.ToString().Trim();
					stBusca = stBusca.Replace("'", "''");
					strFiltro += " And Deposito Like '" + strAntes + stBusca + strDespues;
					break;
			}

			string stFiltro = strFiltro;
			try
			{
				C1.Data.FilterCondition fcPadre = new 
					C1.Data.FilterCondition(cdsEstadoCuenta, "EstadoCuenta", stFiltro);
				C1.Data.FilterConditions fcTotal = new C1.Data.FilterConditions();
				fcTotal.Add(fcPadre);

				cdsEstadoCuenta.Fill(fcTotal, false);
				if (this.cdvEstadoCuenta.Count == 0)
				{
					MessageBox.Show("No existen registros que cumplan la condicion de busqueda",
						"Información", MessageBoxButtons.OK, MessageBoxIcon.Information);
					return;
				}
			}
			catch(System.Exception ex)
			{
				MessageBox.Show(ex.Message, "Error");
			}

			if (this.cdvEstadoCuenta.Count == 1)
			{
				this.ultraGrid1.Select();
				this.DialogResult = DialogResult.OK;
			}
			else
			{
				ultraGrid1.Refresh();
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

		private void BuscaArticulo_Resize(object sender, System.EventArgs e)
		{
			this.ultraGrid1.DisplayLayout.AutoFitColumns = true;
		}

		private void cdsArticulo_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			this.cdsEstadoCuenta.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);		
		}

		private void rdgOpcion_SelectedIndexChanged(object sender, System.EventArgs e)
		{
			this.txtNombre.Enabled = false;
			this.txtCodigo.Enabled = false;
			this.txtNumero.Enabled = false;
			this.txtCheque.Enabled = false;
			this.cmbFecha.Enabled = false;
			this.txtValor.Enabled = false;
			this.txtDepositos.Enabled = false;

			switch((int) this.rdgOpcion.Value)
			{
				case 0:
					this.txtNombre.Enabled = true;
					this.txtNombre.Select();
					break;
				case 1:
					this.txtCodigo.Enabled = true;
					this.txtCodigo.Select();
					break;
				case 2:
					this.txtCheque.Enabled = true;
					this.txtCheque.Select();
					break;
				case 3:
					this.txtNumero.Enabled = true;
					this.txtNumero.Select();
					break;
				case 4:
					this.cmbFecha.Enabled = true;
					this.cmbFecha.Select();
					break;
				case 5:
					this.txtValor.Enabled = true;
					this.txtValor.Select();
					break;
				case 6:
					this.txtDepositos.Enabled = true;
					this.txtDepositos.Select();
					break;
			}
		}
	}
}
