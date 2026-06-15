using System;
using System.Globalization	;
using System.Data.SqlClient;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using CrystalDecisions.Shared;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de EstadoCuentaImprime.
	/// </summary>
	public class EstadoCuentaImprime : DevExpress.XtraEditors.XtraForm
	{
		public Infragistics.Win.UltraWinGrid.UltraCombo cmbCuentaCorr;
		private C1.Data.C1DataSet cdsCuentaCorr;
		private System.Windows.Forms.Label label1;
		private System.Windows.Forms.Label label2;
		private System.Windows.Forms.Label label3;
		private Infragistics.Win.UltraWinEditors.UltraOptionSet optTipo;
		private System.Windows.Forms.Label label4;
		private System.Windows.Forms.ErrorProvider errorProvider;
		private C1.Data.C1DataSet cdsConciliacion;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtSaldoBanco;
		public DevExpress.XtraEditors.DateEdit dtDesde;
		public DevExpress.XtraEditors.DateEdit dtHasta;
		private DevExpress.XtraGrid.GridControl grdBalances;
		private DevExpress.XtraGrid.Views.Grid.GridView grvBalances;
		private DevExpress.XtraGrid.Views.Grid.GridView gridView4;
		private DevExpress.XtraGrid.Columns.GridColumn colidCuenta_Saldo;
		private DevExpress.XtraGrid.Columns.GridColumn colFechaIni;
		private DevExpress.XtraGrid.Columns.GridColumn colFechaFin;
		private DevExpress.XtraGrid.Columns.GridColumn colidCuenta;
		private DevExpress.XtraGrid.Columns.GridColumn colSaldo;
		private DevExpress.XtraEditors.Repository.RepositoryItemLookUpEdit repositoryItemLookUpEdit1;
		private C1.Data.C1DataView cdvEstadoCuenta;
		private DevExpress.XtraEditors.SimpleButton btCancelar;
		private DevExpress.XtraEditors.SimpleButton btAceptar;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public EstadoCuentaImprime()
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
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("CuentaCorriente", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCuentaCorriente");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idSucursal");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idBanco");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCuenta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("NumCuenta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("NumCheque");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Agencia");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Reporte");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Automatico");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Borrar");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn11 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Asiento");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn12 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCuentaTipo");
			Infragistics.Win.Appearance appearance1 = new Infragistics.Win.Appearance();
			Infragistics.Win.ValueListItem valueListItem1 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem2 = new Infragistics.Win.ValueListItem();
			this.cmbCuentaCorr = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.cdsConciliacion = new C1.Data.C1DataSet();
			this.cdsCuentaCorr = new C1.Data.C1DataSet();
			this.label1 = new System.Windows.Forms.Label();
			this.errorProvider = new System.Windows.Forms.ErrorProvider();
			this.label2 = new System.Windows.Forms.Label();
			this.label3 = new System.Windows.Forms.Label();
			this.optTipo = new Infragistics.Win.UltraWinEditors.UltraOptionSet();
			this.label4 = new System.Windows.Forms.Label();
			this.txtSaldoBanco = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.dtDesde = new DevExpress.XtraEditors.DateEdit();
			this.dtHasta = new DevExpress.XtraEditors.DateEdit();
			this.grdBalances = new DevExpress.XtraGrid.GridControl();
			this.grvBalances = new DevExpress.XtraGrid.Views.Grid.GridView();
			this.colidCuenta_Saldo = new DevExpress.XtraGrid.Columns.GridColumn();
			this.colFechaIni = new DevExpress.XtraGrid.Columns.GridColumn();
			this.colFechaFin = new DevExpress.XtraGrid.Columns.GridColumn();
			this.colidCuenta = new DevExpress.XtraGrid.Columns.GridColumn();
			this.repositoryItemLookUpEdit1 = new DevExpress.XtraEditors.Repository.RepositoryItemLookUpEdit();
			this.colSaldo = new DevExpress.XtraGrid.Columns.GridColumn();
			this.gridView4 = new DevExpress.XtraGrid.Views.Grid.GridView();
			this.cdvEstadoCuenta = new C1.Data.C1DataView();
			this.btCancelar = new DevExpress.XtraEditors.SimpleButton();
			this.btAceptar = new DevExpress.XtraEditors.SimpleButton();
			((System.ComponentModel.ISupportInitialize)(this.cmbCuentaCorr)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsConciliacion)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsCuentaCorr)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.optTipo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtSaldoBanco)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtDesde.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtHasta.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.grdBalances)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.grvBalances)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.repositoryItemLookUpEdit1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.gridView4)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvEstadoCuenta)).BeginInit();
			this.SuspendLayout();
			// 
			// cmbCuentaCorr
			// 
			this.cmbCuentaCorr.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbCuentaCorr.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbCuentaCorr.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsConciliacion, "Cuenta_Saldo.idCuenta"));
			this.cmbCuentaCorr.DataMember = "CuentaCorriente";
			this.cmbCuentaCorr.DataSource = this.cdsCuentaCorr;
			ultraGridColumn1.Header.VisiblePosition = 6;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn2.Header.VisiblePosition = 9;
			ultraGridColumn2.Hidden = true;
			ultraGridColumn3.Header.VisiblePosition = 4;
			ultraGridColumn3.Hidden = true;
			ultraGridColumn4.Header.VisiblePosition = 5;
			ultraGridColumn4.Hidden = true;
			ultraGridColumn5.Header.VisiblePosition = 0;
			ultraGridColumn6.Header.VisiblePosition = 7;
			ultraGridColumn6.Hidden = true;
			ultraGridColumn7.Header.VisiblePosition = 1;
			ultraGridColumn7.Hidden = true;
			ultraGridColumn8.Header.VisiblePosition = 8;
			ultraGridColumn8.Hidden = true;
			ultraGridColumn9.Header.VisiblePosition = 2;
			ultraGridColumn9.Hidden = true;
			ultraGridColumn10.Header.VisiblePosition = 3;
			ultraGridColumn10.Hidden = true;
			ultraGridColumn11.Header.VisiblePosition = 10;
			ultraGridColumn11.Hidden = true;
			ultraGridColumn12.Header.VisiblePosition = 11;
			ultraGridColumn12.Hidden = true;
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
																										 ultraGridColumn12});
			this.cmbCuentaCorr.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.cmbCuentaCorr.DisplayMember = "NumCuenta";
			this.cmbCuentaCorr.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbCuentaCorr.Location = new System.Drawing.Point(104, 96);
			this.cmbCuentaCorr.Name = "cmbCuentaCorr";
			this.cmbCuentaCorr.Size = new System.Drawing.Size(136, 22);
			this.cmbCuentaCorr.TabIndex = 1;
			this.cmbCuentaCorr.ValueMember = "idCuenta";
			// 
			// cdsConciliacion
			// 
			this.cdsConciliacion.BindingContextCtrl = this;
			this.cdsConciliacion.DataLibrary = "LibContabilidad";
			this.cdsConciliacion.DataLibraryUrl = "";
			this.cdsConciliacion.DataSetDef = "dsConciliacion";
			this.cdsConciliacion.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsConciliacion.Name = "cdsConciliacion";
			this.cdsConciliacion.SchemaClassName = "LibContabilidad.DataClass";
			this.cdsConciliacion.SchemaDef = null;
			this.cdsConciliacion.BeforeFill += new C1.Data.FillEventHandler(this.cdsConciliacion_BeforeFill);
			// 
			// cdsCuentaCorr
			// 
			this.cdsCuentaCorr.BindingContextCtrl = this;
			this.cdsCuentaCorr.DataLibrary = "LibContabilidad";
			this.cdsCuentaCorr.DataLibraryUrl = "";
			this.cdsCuentaCorr.DataSetDef = "dsCuentaCorriente";
			this.cdsCuentaCorr.Locale = new System.Globalization.CultureInfo("es-MX");
			this.cdsCuentaCorr.Name = "cdsCuentaCorr";
			this.cdsCuentaCorr.SchemaClassName = "LibContabilidad.DataClass";
			this.cdsCuentaCorr.SchemaDef = null;
			this.cdsCuentaCorr.BeforeFill += new C1.Data.FillEventHandler(this.cdsCuentaCorr_BeforeFill);
			// 
			// label1
			// 
			this.label1.AutoSize = true;
			this.label1.Location = new System.Drawing.Point(24, 32);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(39, 16);
			this.label1.TabIndex = 4;
			this.label1.Text = "Desde:";
			// 
			// errorProvider
			// 
			this.errorProvider.ContainerControl = this;
			// 
			// label2
			// 
			this.label2.AutoSize = true;
			this.label2.Location = new System.Drawing.Point(24, 72);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(37, 16);
			this.label2.TabIndex = 6;
			this.label2.Text = "Hasta:";
			// 
			// label3
			// 
			this.label3.AutoSize = true;
			this.label3.Location = new System.Drawing.Point(24, 104);
			this.label3.Name = "label3";
			this.label3.Size = new System.Drawing.Size(43, 16);
			this.label3.TabIndex = 7;
			this.label3.Text = "Cuenta:";
			// 
			// optTipo
			// 
			this.optTipo.CheckedIndex = 1;
			this.optTipo.ItemAppearance = appearance1;
			this.optTipo.ItemOrigin = new System.Drawing.Point(6, 6);
			valueListItem1.DataValue = 0;
			valueListItem1.DisplayText = "Conciliación Bancaria";
			valueListItem2.DataValue = 3;
			valueListItem2.DisplayText = "Resumen";
			this.optTipo.Items.Add(valueListItem1);
			this.optTipo.Items.Add(valueListItem2);
			this.optTipo.ItemSpacingVertical = 5;
			this.optTipo.Location = new System.Drawing.Point(56, 160);
			this.optTipo.Name = "optTipo";
			this.optTipo.Size = new System.Drawing.Size(160, 86);
			this.optTipo.TabIndex = 8;
			this.optTipo.Text = "Resumen";
			// 
			// label4
			// 
			this.label4.AutoSize = true;
			this.label4.Location = new System.Drawing.Point(24, 128);
			this.label4.Name = "label4";
			this.label4.Size = new System.Drawing.Size(69, 16);
			this.label4.TabIndex = 9;
			this.label4.Text = "Saldo Banco:";
			// 
			// txtSaldoBanco
			// 
			this.txtSaldoBanco.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsConciliacion, "Cuenta_Saldo.Saldo"));
			this.txtSaldoBanco.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtSaldoBanco.FormatString = "#,##0.00";
			this.txtSaldoBanco.Location = new System.Drawing.Point(104, 128);
			this.txtSaldoBanco.MaskInput = "-nn,nnn,nnn.nn";
			this.txtSaldoBanco.Name = "txtSaldoBanco";
			this.txtSaldoBanco.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtSaldoBanco.PromptChar = ' ';
			this.txtSaldoBanco.Size = new System.Drawing.Size(88, 22);
			this.txtSaldoBanco.TabIndex = 10;
			// 
			// dtDesde
			// 
			this.dtDesde.DataBindings.Add(new System.Windows.Forms.Binding("EditValue", this.cdsConciliacion, "Cuenta_Saldo.FechaIni"));
			this.dtDesde.EditValue = null;
			this.dtDesde.Location = new System.Drawing.Point(104, 32);
			this.dtDesde.Name = "dtDesde";
			// 
			// dtDesde.Properties
			// 
			this.dtDesde.Properties.Buttons.AddRange(new DevExpress.XtraEditors.Controls.EditorButton[] {
																																																		new DevExpress.XtraEditors.Controls.EditorButton(DevExpress.XtraEditors.Controls.ButtonPredefines.Combo)});
			this.dtDesde.Properties.DisplayFormat.FormatString = "dd/MMM/yyyy";
			this.dtDesde.Properties.DisplayFormat.FormatType = DevExpress.Utils.FormatType.DateTime;
			this.dtDesde.Properties.EditFormat.FormatString = "dd/MMM/yyyy";
			this.dtDesde.Properties.EditFormat.FormatType = DevExpress.Utils.FormatType.DateTime;
			this.dtDesde.Properties.Mask.EditMask = "dd/MMM/yyyy";
			this.dtDesde.Properties.NullDate = "";
			this.dtDesde.Size = new System.Drawing.Size(104, 20);
			this.dtDesde.TabIndex = 70;
			// 
			// dtHasta
			// 
			this.dtHasta.DataBindings.Add(new System.Windows.Forms.Binding("EditValue", this.cdsConciliacion, "Cuenta_Saldo.FechaFin"));
			this.dtHasta.EditValue = null;
			this.dtHasta.Location = new System.Drawing.Point(104, 64);
			this.dtHasta.Name = "dtHasta";
			// 
			// dtHasta.Properties
			// 
			this.dtHasta.Properties.Buttons.AddRange(new DevExpress.XtraEditors.Controls.EditorButton[] {
																																																		new DevExpress.XtraEditors.Controls.EditorButton(DevExpress.XtraEditors.Controls.ButtonPredefines.Combo)});
			this.dtHasta.Properties.DisplayFormat.FormatString = "dd/MMM/yyyy";
			this.dtHasta.Properties.DisplayFormat.FormatType = DevExpress.Utils.FormatType.DateTime;
			this.dtHasta.Properties.EditFormat.FormatString = "dd/MMM/yyyy";
			this.dtHasta.Properties.EditFormat.FormatType = DevExpress.Utils.FormatType.DateTime;
			this.dtHasta.Properties.Mask.EditMask = "dd/MMM/yyyy";
			this.dtHasta.Properties.NullDate = "";
			this.dtHasta.Size = new System.Drawing.Size(104, 20);
			this.dtHasta.TabIndex = 71;
			// 
			// grdBalances
			// 
			this.grdBalances.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.grdBalances.DataMember = "Cuenta_Saldo";
			this.grdBalances.DataSource = this.cdsConciliacion;
			// 
			// grdBalances.EmbeddedNavigator
			// 
			this.grdBalances.EmbeddedNavigator.Name = "";
			this.grdBalances.Location = new System.Drawing.Point(280, 16);
			this.grdBalances.MainView = this.grvBalances;
			this.grdBalances.Name = "grdBalances";
			this.grdBalances.RepositoryItems.AddRange(new DevExpress.XtraEditors.Repository.RepositoryItem[] {
																																																				 this.repositoryItemLookUpEdit1});
			this.grdBalances.Size = new System.Drawing.Size(328, 288);
			this.grdBalances.TabIndex = 73;
			this.grdBalances.ViewCollection.AddRange(new DevExpress.XtraGrid.Views.Base.BaseView[] {
																																															 this.grvBalances,
																																															 this.gridView4});
			// 
			// grvBalances
			// 
			this.grvBalances.Columns.AddRange(new DevExpress.XtraGrid.Columns.GridColumn[] {
																																											 this.colidCuenta_Saldo,
																																											 this.colFechaIni,
																																											 this.colFechaFin,
																																											 this.colidCuenta,
																																											 this.colSaldo});
			this.grvBalances.GridControl = this.grdBalances;
			this.grvBalances.Name = "grvBalances";
			this.grvBalances.OptionsSelection.MultiSelect = true;
			this.grvBalances.OptionsSelection.MultiSelectMode = DevExpress.XtraGrid.Views.Grid.GridMultiSelectMode.CellSelect;
			this.grvBalances.OptionsView.NewItemRowPosition = DevExpress.XtraGrid.Views.Grid.NewItemRowPosition.Bottom;
			this.grvBalances.OptionsView.ShowGroupPanel = false;
			// 
			// colidCuenta_Saldo
			// 
			this.colidCuenta_Saldo.Caption = "idCuenta_Saldo";
			this.colidCuenta_Saldo.FieldName = "idCuenta_Saldo";
			this.colidCuenta_Saldo.Name = "colidCuenta_Saldo";
			// 
			// colFechaIni
			// 
			this.colFechaIni.Caption = "Desde";
			this.colFechaIni.DisplayFormat.FormatString = "dd/MMM/yyyy";
			this.colFechaIni.DisplayFormat.FormatType = DevExpress.Utils.FormatType.DateTime;
			this.colFechaIni.FieldName = "FechaIni";
			this.colFechaIni.Name = "colFechaIni";
			this.colFechaIni.Visible = true;
			this.colFechaIni.VisibleIndex = 0;
			// 
			// colFechaFin
			// 
			this.colFechaFin.Caption = "Hasta";
			this.colFechaFin.DisplayFormat.FormatString = "dd/MMM/yyyy";
			this.colFechaFin.DisplayFormat.FormatType = DevExpress.Utils.FormatType.DateTime;
			this.colFechaFin.FieldName = "FechaFin";
			this.colFechaFin.Name = "colFechaFin";
			this.colFechaFin.Visible = true;
			this.colFechaFin.VisibleIndex = 1;
			// 
			// colidCuenta
			// 
			this.colidCuenta.Caption = "idCuenta";
			this.colidCuenta.ColumnEdit = this.repositoryItemLookUpEdit1;
			this.colidCuenta.FieldName = "idCuenta";
			this.colidCuenta.Name = "colidCuenta";
			// 
			// repositoryItemLookUpEdit1
			// 
			this.repositoryItemLookUpEdit1.AutoHeight = false;
			this.repositoryItemLookUpEdit1.Buttons.AddRange(new DevExpress.XtraEditors.Controls.EditorButton[] {
																																																					 new DevExpress.XtraEditors.Controls.EditorButton(DevExpress.XtraEditors.Controls.ButtonPredefines.Combo)});
			this.repositoryItemLookUpEdit1.DataSource = this.cdsConciliacion;
			this.repositoryItemLookUpEdit1.Name = "repositoryItemLookUpEdit1";
			// 
			// colSaldo
			// 
			this.colSaldo.Caption = "Saldo";
			this.colSaldo.DisplayFormat.FormatString = "n2";
			this.colSaldo.DisplayFormat.FormatType = DevExpress.Utils.FormatType.Numeric;
			this.colSaldo.FieldName = "Saldo";
			this.colSaldo.Name = "colSaldo";
			this.colSaldo.Visible = true;
			this.colSaldo.VisibleIndex = 2;
			// 
			// gridView4
			// 
			this.gridView4.GridControl = this.grdBalances;
			this.gridView4.Name = "gridView4";
			// 
			// cdvEstadoCuenta
			// 
			this.cdvEstadoCuenta.BindingContextCtrl = this;
			this.cdvEstadoCuenta.DataSet = null;
			this.cdvEstadoCuenta.TableName = "";
			this.cdvEstadoCuenta.TableViewName = "EstadoCuenta";
			// 
			// btCancelar
			// 
			this.btCancelar.DialogResult = System.Windows.Forms.DialogResult.Cancel;
			this.btCancelar.Location = new System.Drawing.Point(168, 264);
			this.btCancelar.Name = "btCancelar";
			this.btCancelar.Size = new System.Drawing.Size(72, 23);
			this.btCancelar.TabIndex = 79;
			this.btCancelar.Text = "&Cancelar";
			this.btCancelar.Click += new System.EventHandler(this.btCancelar_Click);
			// 
			// btAceptar
			// 
			this.btAceptar.Location = new System.Drawing.Point(24, 264);
			this.btAceptar.Name = "btAceptar";
			this.btAceptar.Size = new System.Drawing.Size(72, 23);
			this.btAceptar.TabIndex = 78;
			this.btAceptar.Text = "&Aceptar";
			this.btAceptar.Click += new System.EventHandler(this.btAceptar_Click);
			// 
			// EstadoCuentaImprime
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 14);
			this.ClientSize = new System.Drawing.Size(632, 318);
			this.Controls.Add(this.btCancelar);
			this.Controls.Add(this.btAceptar);
			this.Controls.Add(this.grdBalances);
			this.Controls.Add(this.dtHasta);
			this.Controls.Add(this.dtDesde);
			this.Controls.Add(this.txtSaldoBanco);
			this.Controls.Add(this.label4);
			this.Controls.Add(this.label3);
			this.Controls.Add(this.label2);
			this.Controls.Add(this.label1);
			this.Controls.Add(this.optTipo);
			this.Controls.Add(this.cmbCuentaCorr);
			this.MaximizeBox = false;
			this.Name = "EstadoCuentaImprime";
			this.Text = "Impresión de Estado de Cuenta";
			this.Load += new System.EventHandler(this.EstadoCuentaImprime_Load);
			((System.ComponentModel.ISupportInitialize)(this.cmbCuentaCorr)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsConciliacion)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsCuentaCorr)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.optTipo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtSaldoBanco)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtDesde.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtHasta.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.grdBalances)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.grvBalances)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.repositoryItemLookUpEdit1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.gridView4)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvEstadoCuenta)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void cdsCuentaCorr_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			this.cdsCuentaCorr.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirConta);
		}

		private void EstadoCuentaImprime_Load(object sender, System.EventArgs e)
		{
			Cursor = Cursors.Default;

		}

		private void btCancelar_Click(object sender, System.EventArgs e)
		{
			this.Close();
		}

		CultureInfo us = new CultureInfo("en-US");
		private void cdsConciliacion_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			this.cdsConciliacion.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirConta);		
		}

		private void btAceptar_Click(object sender, System.EventArgs e)
		{
			Cursor = Cursors.WaitCursor;
			
			cdsConciliacion.Update();

			DateTime dtFechaIni = (DateTime) this.dtDesde.EditValue;
			DateTime dtFechaFin = (DateTime) this.dtHasta.EditValue;
			string stSelect = "EXEC CreaConciliacion " + cmbCuentaCorr.Value.ToString()
				+ ", '" + dtFechaIni.ToString("yyyyMMdd") 
				+ "', '" + dtFechaFin.AddDays(1).ToString("yyyyMMdd") + "'";
			Funcion.EjecutaSQL(cdsConciliacion, stSelect, true);

			ParameterFields paramFields = new ParameterFields ();
			ParameterField paramField = new ParameterField ();

			ParameterDiscreteValue discreteVal = new ParameterDiscreteValue ();
			paramField.ParameterFieldName = "dCheque";
			discreteVal.Value = double.Parse(txtSaldoBanco.Value.ToString());
			paramField.CurrentValues.Add (discreteVal);
			paramFields.Add (paramField);

			string stFiltro = "{Conciliacion.idCuenta} = " + cmbCuentaCorr.Value.ToString()
				+ " And {Conciliacion.FechaGeneracion} < #" 
				+ dtFechaFin.AddDays(1).ToString("MM/dd/yyyy") + "#";
			if ((int) this.optTipo.Value == 0 || (int) this.optTipo.Value == 3) 
				stFiltro += " And {Conciliacion.Mayor} = False ";

			if ((int) this.optTipo.Value == 1) stFiltro += " And {Conciliacion.Mayor} = True";
			string stReporte = "Conciliacion.rpt";
			if ((int) optTipo.Value == 3) 
			{
				string stMayor = "";

				string stFechaLinea = "Select IsNull(FechaLinea, 0) From SeteoC";
				bool bFechaLinea = Funcion.bEscalarSQL(cdsCuentaCorr, stFechaLinea, true);
				if (!bFechaLinea)
				{
					stMayor = "Select Sum(IsNull(Debe, 0) - IsNull(Haber, 0)) "
						+ "From AsientoDetalle inner join Asiento On AsientoDetalle.idAsiento = "
						+ "Asiento.idAsiento Where idCuenta = " + cmbCuentaCorr.Value.ToString()
						+ " And Asiento.Fecha < '"
						+ dtFechaFin.AddDays(1).ToString("yyyyMMdd") + "' And Asiento.Borrar = 0";
				}
				else
				{
					stMayor = "Select Sum(IsNull(Debe, 0) - IsNull(Haber, 0)) "
						+ "From AsientoDetalle inner join Asiento On AsientoDetalle.idAsiento = "
						+ "Asiento.idAsiento Where idCuenta = " + cmbCuentaCorr.Value.ToString()
						+ " And AsientoDetalle.FechaComp < '"
						+ dtFechaFin.AddDays(1).ToString("yyyyMMdd") + "' And Asiento.Borrar = 0";
				}

				double dMayor = Funcion.dEscalarSQL(cdsCuentaCorr, stMayor, false);
				ParameterField paramField2 = new ParameterField ();

				ParameterDiscreteValue discreteVal2 = new ParameterDiscreteValue ();
				paramField2.ParameterFieldName = "dMayor";
				discreteVal2.Value = dMayor;
				paramField2.CurrentValues.Add (discreteVal2);
				paramFields.Add (paramField2);

				
				ParameterField paramField1 = new ParameterField ();

				ParameterDiscreteValue discreteVal1 = new ParameterDiscreteValue ();
				paramField1.ParameterFieldName = "dtFechaIni";
				discreteVal1.Value = (DateTime) dtDesde.EditValue;
				paramField1.CurrentValues.Add (discreteVal1);
				paramFields.Add (paramField1);

				stReporte = "ConciliacionResumen1.rpt";
			}
			stSelect = "Update Conciliacion Set Mayor = 0 Where Borrar = 0";
			Funcion.EjecutaSQL(cdsConciliacion, stSelect, true);
			stSelect = "Update Conciliacion Set Mayor = 1 Where Conciliacion.Confirmado = 1 "
				+ "And Conciliacion.FechaConfirmacion < '"
				+ dtFechaFin.AddDays(1).ToString("yyyyMMdd") + "' And Borrar = 0"
				+ " And idCuenta = " + cmbCuentaCorr.Value.ToString();
			Funcion.EjecutaSQL(cdsConciliacion, stSelect, true);
			Reporte miReporte = new Reporte(stReporte, stFiltro);
			miReporte.MdiParent = this.MdiParent;
			miReporte.crVista.ParameterFieldInfo = paramFields;
			string stTitulo = "Conciliación Bancaria desde el "
					+ dtFechaIni.ToString("dd/MMM/yyy") + " hasta el "
					+ dtFechaFin.ToString("dd/MMM/yyy") + " Cuenta: " + cmbCuentaCorr.Text;
      miReporte.Titulo(stTitulo);
			miReporte.Show();
			Cursor = Cursors.Default;
		}

	}
}
