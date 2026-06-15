using System;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de AsientoCuenta.
	/// </summary>
	public class AsientoCuenta : DevExpress.XtraEditors.XtraForm
	{
		private C1.Data.C1DataSet cdsCuentaCorr;
		private C1.Data.C1DataSet cdsCuenta;
		private System.Windows.Forms.Label label1;
		private C1.Data.C1DataView cdvCuentaCod;
		private C1.Data.C1DataView cdvCuenta;
		private System.Windows.Forms.Label label2;
		private System.Windows.Forms.Label label3;
		private System.Windows.Forms.Label label4;
		private C1.Data.C1DataView cdvBanco;
		private System.Windows.Forms.Label label5;
		private DevExpress.XtraEditors.SimpleButton btBanco;
		private DevExpress.XtraEditors.SimpleButton btPlanCuentas;
		private System.Windows.Forms.ErrorProvider errorProvider;
		private System.Windows.Forms.TextBox txtIdCuentaCorr;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor ultraCheckEditor1;
		private C1.Data.C1DataSet cdsSeguridad;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkAsiento;
		private System.Windows.Forms.ToolTip toolTip1;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbCuentaCod;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbCuentaNombre;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtReporte;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbBanco;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtNumCheque;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtNumCuenta;
		private BarraDatoSQL.BarraDatoSQL barraDato1;
		private DevExpress.XtraEditors.SimpleButton btTipo;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbTipo;
		private Infragistics.Win.Misc.UltraLabel lblTipo;
		private C1.Data.C1DataSet cdsSeteo;
		private System.Windows.Forms.Label label6;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor ultraCheckEditor2;
		private DevExpress.XtraEditors.RadioGroup rdgTipo;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkPagoNomina;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkProvisiones;
		private System.Windows.Forms.GroupBox groupBox1;
		private System.Windows.Forms.GroupBox groupBox2;
		private System.Windows.Forms.Label label7;
		private System.Windows.Forms.Label label8;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtDesde;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtHasta;
		private System.Windows.Forms.Label label9;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtDesde;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkCorriente;
		private System.ComponentModel.IContainer components;

		public AsientoCuenta()
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
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Cuenta", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCuenta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Grupo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CodRapido");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Descripcion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CuentaOrigen");
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Cuenta", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCuenta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Grupo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CodRapido");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn11 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Descripcion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn12 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CuentaOrigen");
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand3 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Banco", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn13 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idBanco");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn14 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn15 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Pais");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn16 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn17 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Borrar");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn18 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Hoja");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn19 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ProcesoOk");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn20 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPais");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand4 = new Infragistics.Win.UltraWinGrid.UltraGridBand("CuentaCorrTipo", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn21 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCuentaTipo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn22 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn23 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Tipo");
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance11 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance12 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance13 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance14 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton1 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(AsientoCuenta));
			Infragistics.Win.Appearance appearance15 = new Infragistics.Win.Appearance();
			this.cdsCuentaCorr = new C1.Data.C1DataSet();
			this.cdvBanco = new C1.Data.C1DataView();
			this.cdvCuentaCod = new C1.Data.C1DataView();
			this.cdsCuenta = new C1.Data.C1DataSet();
			this.cdvCuenta = new C1.Data.C1DataView();
			this.label1 = new System.Windows.Forms.Label();
			this.label2 = new System.Windows.Forms.Label();
			this.label3 = new System.Windows.Forms.Label();
			this.label4 = new System.Windows.Forms.Label();
			this.label5 = new System.Windows.Forms.Label();
			this.btBanco = new DevExpress.XtraEditors.SimpleButton();
			this.btPlanCuentas = new DevExpress.XtraEditors.SimpleButton();
			this.errorProvider = new System.Windows.Forms.ErrorProvider();
			this.txtIdCuentaCorr = new System.Windows.Forms.TextBox();
			this.ultraCheckEditor1 = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.cdsSeguridad = new C1.Data.C1DataSet();
			this.chkAsiento = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.toolTip1 = new System.Windows.Forms.ToolTip(this.components);
			this.cmbCuentaCod = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.cmbCuentaNombre = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.txtReporte = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtNumCheque = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtNumCuenta = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.cmbBanco = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.barraDato1 = new BarraDatoSQL.BarraDatoSQL();
			this.cmbTipo = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.lblTipo = new Infragistics.Win.Misc.UltraLabel();
			this.btTipo = new DevExpress.XtraEditors.SimpleButton();
			this.cdsSeteo = new C1.Data.C1DataSet();
			this.label6 = new System.Windows.Forms.Label();
			this.ultraCheckEditor2 = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.rdgTipo = new DevExpress.XtraEditors.RadioGroup();
			this.chkPagoNomina = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.chkProvisiones = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.groupBox1 = new System.Windows.Forms.GroupBox();
			this.groupBox2 = new System.Windows.Forms.GroupBox();
			this.txtDesde = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label7 = new System.Windows.Forms.Label();
			this.txtHasta = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label8 = new System.Windows.Forms.Label();
			this.label9 = new System.Windows.Forms.Label();
			this.dtDesde = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.chkCorriente = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			((System.ComponentModel.ISupportInitialize)(this.cdsCuentaCorr)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvBanco)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvCuentaCod)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsCuenta)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvCuenta)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeguridad)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCuentaCod)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCuentaNombre)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtReporte)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNumCheque)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNumCuenta)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbBanco)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbTipo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.rdgTipo.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDesde)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtHasta)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtDesde)).BeginInit();
			this.SuspendLayout();
			// 
			// cdsCuentaCorr
			// 
			this.cdsCuentaCorr.BindingContextCtrl = this;
			this.cdsCuentaCorr.DataLibrary = "LibContabilidad";
			this.cdsCuentaCorr.DataLibraryUrl = "";
			this.cdsCuentaCorr.DataSetDef = "dsCuentaCorriente";
			this.cdsCuentaCorr.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsCuentaCorr.Name = "cdsCuentaCorr";
			this.cdsCuentaCorr.SchemaClassName = "LibContabilidad.DataClass";
			this.cdsCuentaCorr.SchemaDef = null;
			this.cdsCuentaCorr.BeforeFill += new C1.Data.FillEventHandler(this.cdsCuentaCorr_BeforeFill);
			// 
			// cdvBanco
			// 
			this.cdvBanco.BindingContextCtrl = this;
			this.cdvBanco.DataSet = this.cdsCuentaCorr;
			this.cdvBanco.TableName = "";
			this.cdvBanco.TableViewName = "Banco";
			// 
			// cdvCuentaCod
			// 
			this.cdvCuentaCod.BindingContextCtrl = this;
			this.cdvCuentaCod.DataSet = this.cdsCuenta;
			this.cdvCuentaCod.RowFilter = "Grupo = 0";
			this.cdvCuentaCod.Sort = "Codigo";
			this.cdvCuentaCod.TableName = "";
			this.cdvCuentaCod.TableViewName = "Cuenta";
			// 
			// cdsCuenta
			// 
			this.cdsCuenta.BindingContextCtrl = this;
			this.cdsCuenta.DataLibrary = "LibContabilidad";
			this.cdsCuenta.DataLibraryUrl = "";
			this.cdsCuenta.DataSetDef = "dsCuenta2";
			this.cdsCuenta.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsCuenta.Name = "cdsCuenta";
			this.cdsCuenta.SchemaClassName = "LibContabilidad.DataClass";
			this.cdsCuenta.SchemaDef = null;
			this.cdsCuenta.BeforeFill += new C1.Data.FillEventHandler(this.cdsCuenta_BeforeFill);
			// 
			// cdvCuenta
			// 
			this.cdvCuenta.BindingContextCtrl = this;
			this.cdvCuenta.DataSet = this.cdsCuenta;
			this.cdvCuenta.RowFilter = "Grupo = 0";
			this.cdvCuenta.Sort = "Descripcion";
			this.cdvCuenta.TableName = "";
			this.cdvCuenta.TableViewName = "Cuenta";
			// 
			// label1
			// 
			this.label1.AutoSize = true;
			this.label1.Location = new System.Drawing.Point(11, 7);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(34, 16);
			this.label1.TabIndex = 4;
			this.label1.Text = "Banco";
			// 
			// label2
			// 
			this.label2.AutoSize = true;
			this.label2.Location = new System.Drawing.Point(11, 176);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(86, 16);
			this.label2.TabIndex = 5;
			this.label2.Text = "Código Contable";
			// 
			// label3
			// 
			this.label3.AutoSize = true;
			this.label3.Location = new System.Drawing.Point(11, 51);
			this.label3.Name = "label3";
			this.label3.Size = new System.Drawing.Size(98, 16);
			this.label3.TabIndex = 7;
			this.label3.Text = "Número de Cuenta";
			// 
			// label4
			// 
			this.label4.AutoSize = true;
			this.label4.Location = new System.Drawing.Point(11, 83);
			this.label4.Name = "label4";
			this.label4.Size = new System.Drawing.Size(42, 16);
			this.label4.TabIndex = 9;
			this.label4.Text = "Cheque";
			// 
			// label5
			// 
			this.label5.AutoSize = true;
			this.label5.Location = new System.Drawing.Point(11, 144);
			this.label5.Name = "label5";
			this.label5.Size = new System.Drawing.Size(43, 16);
			this.label5.TabIndex = 11;
			this.label5.Text = "Reporte";
			// 
			// btBanco
			// 
			this.btBanco.Location = new System.Drawing.Point(441, 3);
			this.btBanco.Name = "btBanco";
			this.btBanco.Size = new System.Drawing.Size(24, 25);
			this.btBanco.TabIndex = 9;
			this.btBanco.Text = "...";
			this.btBanco.Click += new System.EventHandler(this.btBanco_Click);
			// 
			// btPlanCuentas
			// 
			this.btPlanCuentas.Location = new System.Drawing.Point(297, 176);
			this.btPlanCuentas.Name = "btPlanCuentas";
			this.btPlanCuentas.Size = new System.Drawing.Size(24, 24);
			this.btPlanCuentas.TabIndex = 10;
			this.btPlanCuentas.Text = "...";
			this.btPlanCuentas.Click += new System.EventHandler(this.btPlanCuentas_Click);
			// 
			// errorProvider
			// 
			this.errorProvider.ContainerControl = this;
			// 
			// txtIdCuentaCorr
			// 
			this.txtIdCuentaCorr.DataBindings.Add(new System.Windows.Forms.Binding("Text", this.cdsCuentaCorr, "CuentaCorriente.idCuentaCorriente"));
			this.txtIdCuentaCorr.Location = new System.Drawing.Point(281, 288);
			this.txtIdCuentaCorr.Name = "txtIdCuentaCorr";
			this.txtIdCuentaCorr.TabIndex = 16;
			this.txtIdCuentaCorr.Text = "";
			// 
			// ultraCheckEditor1
			// 
			appearance1.ForeColorDisabled = System.Drawing.Color.Black;
			this.ultraCheckEditor1.Appearance = appearance1;
			this.ultraCheckEditor1.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.ultraCheckEditor1.DataBindings.Add(new System.Windows.Forms.Binding("CheckedValue", this.cdsCuentaCorr, "CuentaCorriente.Automatico"));
			this.ultraCheckEditor1.Location = new System.Drawing.Point(360, 81);
			this.ultraCheckEditor1.Name = "ultraCheckEditor1";
			this.ultraCheckEditor1.Size = new System.Drawing.Size(128, 21);
			this.ultraCheckEditor1.TabIndex = 7;
			this.ultraCheckEditor1.Text = "Automático";
			// 
			// cdsSeguridad
			// 
			this.cdsSeguridad.BindingContextCtrl = this;
			this.cdsSeguridad.DataLibrary = "LibContabilidad";
			this.cdsSeguridad.DataLibraryUrl = "";
			this.cdsSeguridad.DataSetDef = "dsSeguridad";
			this.cdsSeguridad.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsSeguridad.Name = "cdsSeguridad";
			this.cdsSeguridad.SchemaClassName = "LibContabilidad.DataClass";
			this.cdsSeguridad.SchemaDef = null;
			// 
			// chkAsiento
			// 
			appearance2.ForeColorDisabled = System.Drawing.Color.Black;
			this.chkAsiento.Appearance = appearance2;
			this.chkAsiento.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.chkAsiento.DataBindings.Add(new System.Windows.Forms.Binding("CheckedValue", this.cdsCuentaCorr, "CuentaCorriente.Asiento"));
			this.chkAsiento.Location = new System.Drawing.Point(336, 48);
			this.chkAsiento.Name = "chkAsiento";
			this.chkAsiento.Size = new System.Drawing.Size(152, 22);
			this.chkAsiento.TabIndex = 6;
			this.chkAsiento.Text = "Genera Asiento / Egreso";
			this.toolTip1.SetToolTip(this.chkAsiento, "Si esta activado al realizar un pago y generar un pago se genera un Asiento. Desa" +
				"ctivado genera un Egreso");
			this.chkAsiento.Visible = false;
			// 
			// cmbCuentaCod
			// 
			appearance3.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbCuentaCod.Appearance = appearance3;
			this.cmbCuentaCod.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbCuentaCod.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbCuentaCod.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsCuentaCorr, "CuentaCorriente.idCuenta"));
			this.cmbCuentaCod.DataSource = this.cdvCuentaCod;
			ultraGridColumn1.Header.VisiblePosition = 2;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn2.Header.VisiblePosition = 4;
			ultraGridColumn2.Hidden = true;
			ultraGridColumn3.Header.VisiblePosition = 3;
			ultraGridColumn3.Hidden = true;
			ultraGridColumn4.Header.Caption = "Codigo";
			ultraGridColumn4.Header.VisiblePosition = 0;
			ultraGridColumn4.Width = 151;
			ultraGridColumn5.Header.VisiblePosition = 1;
			ultraGridColumn5.Width = 368;
			ultraGridColumn6.Header.VisiblePosition = 5;
			ultraGridBand1.Columns.AddRange(new object[] {
																										 ultraGridColumn1,
																										 ultraGridColumn2,
																										 ultraGridColumn3,
																										 ultraGridColumn4,
																										 ultraGridColumn5,
																										 ultraGridColumn6});
			this.cmbCuentaCod.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.cmbCuentaCod.DisplayMember = "CodRapido";
			this.cmbCuentaCod.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbCuentaCod.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbCuentaCod.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbCuentaCod.Location = new System.Drawing.Point(115, 176);
			this.cmbCuentaCod.Name = "cmbCuentaCod";
			this.cmbCuentaCod.Size = new System.Drawing.Size(174, 21);
			this.cmbCuentaCod.TabIndex = 4;
			this.cmbCuentaCod.ValueMember = "idCuenta";
			this.cmbCuentaCod.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.cmbCuentaCod_InitializeLayout);
			// 
			// cmbCuentaNombre
			// 
			appearance4.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbCuentaNombre.Appearance = appearance4;
			this.cmbCuentaNombre.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbCuentaNombre.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbCuentaNombre.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsCuentaCorr, "CuentaCorriente.idCuenta"));
			this.cmbCuentaNombre.DataSource = this.cdvCuenta;
			ultraGridColumn7.Header.VisiblePosition = 1;
			ultraGridColumn8.Header.VisiblePosition = 4;
			ultraGridColumn8.Hidden = true;
			ultraGridColumn9.Header.VisiblePosition = 3;
			ultraGridColumn9.Hidden = true;
			ultraGridColumn10.Header.VisiblePosition = 2;
			ultraGridColumn10.Hidden = true;
			ultraGridColumn11.Header.VisiblePosition = 0;
			ultraGridColumn11.Width = 234;
			ultraGridColumn12.Header.VisiblePosition = 5;
			ultraGridBand2.Columns.AddRange(new object[] {
																										 ultraGridColumn7,
																										 ultraGridColumn8,
																										 ultraGridColumn9,
																										 ultraGridColumn10,
																										 ultraGridColumn11,
																										 ultraGridColumn12});
			this.cmbCuentaNombre.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			this.cmbCuentaNombre.DisplayMember = "Descripcion";
			this.cmbCuentaNombre.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbCuentaNombre.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbCuentaNombre.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbCuentaNombre.Location = new System.Drawing.Point(115, 208);
			this.cmbCuentaNombre.Name = "cmbCuentaNombre";
			this.cmbCuentaNombre.Size = new System.Drawing.Size(205, 21);
			this.cmbCuentaNombre.TabIndex = 8;
			this.cmbCuentaNombre.ValueMember = "idCuenta";
			// 
			// txtReporte
			// 
			appearance5.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtReporte.Appearance = appearance5;
			this.txtReporte.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsCuentaCorr, "CuentaCorriente.Reporte"));
			this.txtReporte.Location = new System.Drawing.Point(115, 144);
			this.txtReporte.Name = "txtReporte";
			this.txtReporte.Size = new System.Drawing.Size(176, 22);
			this.txtReporte.TabIndex = 3;
			// 
			// txtNumCheque
			// 
			appearance6.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtNumCheque.Appearance = appearance6;
			this.txtNumCheque.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsCuentaCorr, "CuentaCorriente.NumCheque"));
			this.txtNumCheque.Location = new System.Drawing.Point(115, 80);
			this.txtNumCheque.Name = "txtNumCheque";
			this.txtNumCheque.Size = new System.Drawing.Size(176, 22);
			this.txtNumCheque.TabIndex = 2;
			// 
			// txtNumCuenta
			// 
			appearance7.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtNumCuenta.Appearance = appearance7;
			this.txtNumCuenta.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsCuentaCorr, "CuentaCorriente.NumCuenta"));
			this.txtNumCuenta.Location = new System.Drawing.Point(115, 48);
			this.txtNumCuenta.Name = "txtNumCuenta";
			this.txtNumCuenta.Size = new System.Drawing.Size(176, 22);
			this.txtNumCuenta.TabIndex = 1;
			// 
			// cmbBanco
			// 
			appearance8.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbBanco.Appearance = appearance8;
			this.cmbBanco.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbBanco.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbBanco.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsCuentaCorr, "CuentaCorriente.idBanco"));
			this.cmbBanco.DataSource = this.cdvBanco;
			this.cmbBanco.DisplayLayout.AutoFitColumns = true;
			ultraGridColumn13.Header.VisiblePosition = 1;
			ultraGridColumn13.Hidden = true;
			ultraGridColumn14.Header.VisiblePosition = 0;
			ultraGridColumn14.Width = 299;
			ultraGridColumn15.Header.VisiblePosition = 2;
			ultraGridColumn15.Hidden = true;
			ultraGridColumn16.Header.VisiblePosition = 3;
			ultraGridColumn16.Hidden = true;
			ultraGridColumn17.Header.VisiblePosition = 4;
			ultraGridColumn17.Hidden = true;
			ultraGridColumn18.Header.VisiblePosition = 5;
			ultraGridColumn18.Hidden = true;
			ultraGridColumn18.Width = 60;
			ultraGridColumn19.Header.VisiblePosition = 6;
			ultraGridColumn19.Hidden = true;
			ultraGridColumn19.Width = 62;
			ultraGridColumn20.Header.VisiblePosition = 7;
			ultraGridColumn20.Hidden = true;
			ultraGridColumn20.Width = 50;
			ultraGridBand3.Columns.AddRange(new object[] {
																										 ultraGridColumn13,
																										 ultraGridColumn14,
																										 ultraGridColumn15,
																										 ultraGridColumn16,
																										 ultraGridColumn17,
																										 ultraGridColumn18,
																										 ultraGridColumn19,
																										 ultraGridColumn20});
			this.cmbBanco.DisplayLayout.BandsSerializer.Add(ultraGridBand3);
			this.cmbBanco.DisplayMember = "Nombre";
			this.cmbBanco.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbBanco.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbBanco.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbBanco.Location = new System.Drawing.Point(115, 5);
			this.cmbBanco.Name = "cmbBanco";
			this.cmbBanco.Size = new System.Drawing.Size(318, 21);
			this.cmbBanco.TabIndex = 0;
			this.cmbBanco.ValueMember = "idBanco";
			// 
			// barraDato1
			// 
			this.barraDato1.BarraMovimiento = true;
			this.barraDato1.DataMember = "CuentaCorriente";
			this.barraDato1.DataNombreId = "idCuentaCorriente";
			this.barraDato1.DataOrden = "";
			this.barraDato1.DataSource = this.cdsCuentaCorr;
			this.barraDato1.DataTabla = "CuentaCorriente";
			this.barraDato1.DataTablaHija = "";
			this.barraDato1.Location = new System.Drawing.Point(10, 288);
			this.barraDato1.Name = "barraDato1";
			this.barraDato1.Size = new System.Drawing.Size(256, 26);
			this.barraDato1.TabIndex = 17;
			this.barraDato1.VisibleBorrar = true;
			this.barraDato1.VisibleBuscar = false;
			this.barraDato1.VisibleEditar = true;
			this.barraDato1.VisibleImprimir = false;
			this.barraDato1.VisibleNuevo = true;
			this.barraDato1.Borrar += new System.EventHandler(this.barraDato1_Borrar);
			this.barraDato1.Restaurar += new System.EventHandler(this.barraDatoSQL1_Restaurar);
			this.barraDato1.Refresca += new System.EventHandler(this.barraDatoSQL1_Refresca);
			this.barraDato1.Crear += new System.EventHandler(this.barraDato1_Crear);
			this.barraDato1.Deshacer += new System.EventHandler(this.barraDato1_Deshacer);
			this.barraDato1.Copiar += new System.EventHandler(this.barraDatoSQL1_Copiar);
			this.barraDato1.Editar += new System.EventHandler(this.barraDato1_Editar);
			this.barraDato1.Grabar += new System.EventHandler(this.barraDato1_Grabar);
			// 
			// cmbTipo
			// 
			this.cmbTipo.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbTipo.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbTipo.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsCuentaCorr, "CuentaCorriente.idCuentaTipo"));
			this.cmbTipo.DataMember = "CuentaCorrTipo";
			this.cmbTipo.DataSource = this.cdsCuentaCorr;
			ultraGridColumn21.Header.VisiblePosition = 2;
			ultraGridColumn21.Hidden = true;
			ultraGridColumn22.Header.VisiblePosition = 1;
			ultraGridColumn23.Header.VisiblePosition = 0;
			ultraGridBand4.Columns.AddRange(new object[] {
																										 ultraGridColumn21,
																										 ultraGridColumn22,
																										 ultraGridColumn23});
			this.cmbTipo.DisplayLayout.BandsSerializer.Add(ultraGridBand4);
			this.cmbTipo.DisplayMember = "Tipo";
			this.cmbTipo.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbTipo.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbTipo.Location = new System.Drawing.Point(336, 288);
			this.cmbTipo.Name = "cmbTipo";
			this.cmbTipo.Size = new System.Drawing.Size(86, 22);
			this.cmbTipo.TabIndex = 18;
			this.cmbTipo.ValueMember = "idCuentaTipo";
			this.cmbTipo.Visible = false;
			// 
			// lblTipo
			// 
			this.lblTipo.Location = new System.Drawing.Point(296, 288);
			this.lblTipo.Name = "lblTipo";
			this.lblTipo.Size = new System.Drawing.Size(32, 25);
			this.lblTipo.TabIndex = 19;
			this.lblTipo.Text = "Tipo";
			this.lblTipo.Visible = false;
			// 
			// btTipo
			// 
			this.btTipo.Location = new System.Drawing.Point(432, 288);
			this.btTipo.Name = "btTipo";
			this.btTipo.Size = new System.Drawing.Size(24, 25);
			this.btTipo.TabIndex = 20;
			this.btTipo.Text = "...";
			this.btTipo.Visible = false;
			this.btTipo.Click += new System.EventHandler(this.btTipo_Click);
			// 
			// cdsSeteo
			// 
			this.cdsSeteo.BindingContextCtrl = this;
			this.cdsSeteo.DataLibrary = "LibContabilidad";
			this.cdsSeteo.DataLibraryUrl = "";
			this.cdsSeteo.DataSetDef = "dsSeteoC";
			this.cdsSeteo.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsSeteo.Name = "cdsSeteo";
			this.cdsSeteo.SchemaClassName = "LibContabilidad.DataClass";
			this.cdsSeteo.SchemaDef = null;
			this.cdsSeteo.BeforeFill += new C1.Data.FillEventHandler(this.cdsSeteo_BeforeFill);
			// 
			// label6
			// 
			this.label6.AutoSize = true;
			this.label6.Location = new System.Drawing.Point(11, 208);
			this.label6.Name = "label6";
			this.label6.Size = new System.Drawing.Size(87, 16);
			this.label6.TabIndex = 21;
			this.label6.Text = "Cuenta Contable";
			// 
			// ultraCheckEditor2
			// 
			appearance9.ForeColorDisabled = System.Drawing.Color.Black;
			this.ultraCheckEditor2.Appearance = appearance9;
			this.ultraCheckEditor2.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.ultraCheckEditor2.DataBindings.Add(new System.Windows.Forms.Binding("CheckedValue", this.cdsCuentaCorr, "CuentaCorriente.VisibleLocales"));
			this.ultraCheckEditor2.Location = new System.Drawing.Point(360, 176);
			this.ultraCheckEditor2.Name = "ultraCheckEditor2";
			this.ultraCheckEditor2.Size = new System.Drawing.Size(128, 21);
			this.ultraCheckEditor2.TabIndex = 22;
			this.ultraCheckEditor2.Text = "Visible en locales";
			// 
			// rdgTipo
			// 
			this.rdgTipo.DataBindings.Add(new System.Windows.Forms.Binding("EditValue", this.cdsCuentaCorr, "CuentaCorriente.Tipo"));
			this.rdgTipo.EditValue = 7;
			this.rdgTipo.Location = new System.Drawing.Point(336, 47);
			this.rdgTipo.Name = "rdgTipo";
			// 
			// rdgTipo.Properties
			// 
			this.rdgTipo.Properties.Appearance.BackColor = System.Drawing.Color.Transparent;
			this.rdgTipo.Properties.Appearance.Options.UseBackColor = true;
			this.rdgTipo.Properties.Items.AddRange(new DevExpress.XtraEditors.Controls.RadioGroupItem[] {
																																																		new DevExpress.XtraEditors.Controls.RadioGroupItem(0, "Corriente"),
																																																		new DevExpress.XtraEditors.Controls.RadioGroupItem(1, "Ahorros")});
			this.rdgTipo.Size = new System.Drawing.Size(152, 24);
			this.rdgTipo.TabIndex = 605;
			// 
			// chkPagoNomina
			// 
			appearance10.ForeColorDisabled = System.Drawing.Color.Black;
			this.chkPagoNomina.Appearance = appearance10;
			this.chkPagoNomina.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.chkPagoNomina.DataBindings.Add(new System.Windows.Forms.Binding("CheckedValue", this.cdsCuentaCorr, "CuentaCorriente.PagoNomina"));
			this.chkPagoNomina.Location = new System.Drawing.Point(360, 113);
			this.chkPagoNomina.Name = "chkPagoNomina";
			this.chkPagoNomina.Size = new System.Drawing.Size(128, 21);
			this.chkPagoNomina.TabIndex = 606;
			this.chkPagoNomina.Text = "Pago Nomina";
			// 
			// chkProvisiones
			// 
			appearance11.ForeColorDisabled = System.Drawing.Color.Black;
			this.chkProvisiones.Appearance = appearance11;
			this.chkProvisiones.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.chkProvisiones.DataBindings.Add(new System.Windows.Forms.Binding("CheckedValue", this.cdsCuentaCorr, "CuentaCorriente.Flujo"));
			this.chkProvisiones.Location = new System.Drawing.Point(360, 144);
			this.chkProvisiones.Name = "chkProvisiones";
			this.chkProvisiones.Size = new System.Drawing.Size(128, 21);
			this.chkProvisiones.TabIndex = 607;
			this.chkProvisiones.Text = "Ver En Provisiones";
			// 
			// groupBox1
			// 
			this.groupBox1.Location = new System.Drawing.Point(0, 32);
			this.groupBox1.Name = "groupBox1";
			this.groupBox1.Size = new System.Drawing.Size(504, 8);
			this.groupBox1.TabIndex = 608;
			this.groupBox1.TabStop = false;
			// 
			// groupBox2
			// 
			this.groupBox2.Location = new System.Drawing.Point(0, 272);
			this.groupBox2.Name = "groupBox2";
			this.groupBox2.Size = new System.Drawing.Size(504, 8);
			this.groupBox2.TabIndex = 609;
			this.groupBox2.TabStop = false;
			// 
			// txtDesde
			// 
			appearance12.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtDesde.Appearance = appearance12;
			this.txtDesde.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsCuentaCorr, "CuentaCorriente.Desde"));
			this.txtDesde.Location = new System.Drawing.Point(115, 112);
			this.txtDesde.Name = "txtDesde";
			this.txtDesde.Size = new System.Drawing.Size(80, 22);
			this.txtDesde.TabIndex = 610;
			// 
			// label7
			// 
			this.label7.AutoSize = true;
			this.label7.Location = new System.Drawing.Point(11, 115);
			this.label7.Name = "label7";
			this.label7.Size = new System.Drawing.Size(35, 16);
			this.label7.TabIndex = 611;
			this.label7.Text = "Desde";
			// 
			// txtHasta
			// 
			appearance13.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtHasta.Appearance = appearance13;
			this.txtHasta.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsCuentaCorr, "CuentaCorriente.Hasta"));
			this.txtHasta.Location = new System.Drawing.Point(264, 112);
			this.txtHasta.Name = "txtHasta";
			this.txtHasta.Size = new System.Drawing.Size(80, 22);
			this.txtHasta.TabIndex = 612;
			// 
			// label8
			// 
			this.label8.AutoSize = true;
			this.label8.Location = new System.Drawing.Point(208, 115);
			this.label8.Name = "label8";
			this.label8.Size = new System.Drawing.Size(33, 16);
			this.label8.TabIndex = 613;
			this.label8.Text = "Hasta";
			// 
			// label9
			// 
			this.label9.AutoSize = true;
			this.label9.BackColor = System.Drawing.Color.Transparent;
			this.label9.Location = new System.Drawing.Point(11, 240);
			this.label9.Name = "label9";
			this.label9.Size = new System.Drawing.Size(80, 16);
			this.label9.TabIndex = 615;
			this.label9.Text = "Fecha Apertura";
			this.label9.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// dtDesde
			// 
			appearance14.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtDesde.Appearance = appearance14;
			this.dtDesde.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.dtDesde.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsCuentaCorr, "CuentaCorriente.FechaApertura"));
			dateButton1.Action = Infragistics.Win.UltraWinSchedule.DateButtonAction.AddToSelection;
			dateButton1.Caption = "Today";
			dateButton1.Type = Infragistics.Win.UltraWinSchedule.DateButtonType.Custom;
			this.dtDesde.DateButtons.Add(dateButton1);
			this.dtDesde.Format = "dd/MM/yyyy";
			this.dtDesde.Location = new System.Drawing.Point(115, 240);
			this.dtDesde.Name = "dtDesde";
			this.dtDesde.NonAutoSizeHeight = 23;
			this.dtDesde.Size = new System.Drawing.Size(104, 21);
			this.dtDesde.SpinButtonsVisible = true;
			this.dtDesde.TabIndex = 614;
			this.dtDesde.Value = ((object)(resources.GetObject("dtDesde.Value")));
			// 
			// chkCorriente
			// 
			appearance15.ForeColorDisabled = System.Drawing.Color.Black;
			this.chkCorriente.Appearance = appearance15;
			this.chkCorriente.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.chkCorriente.DataBindings.Add(new System.Windows.Forms.Binding("CheckedValue", this.cdsCuentaCorr, "CuentaCorriente.VisibleLocales"));
			this.chkCorriente.Location = new System.Drawing.Point(360, 208);
			this.chkCorriente.Name = "chkCorriente";
			this.chkCorriente.Size = new System.Drawing.Size(128, 21);
			this.chkCorriente.TabIndex = 616;
			this.chkCorriente.Text = "Cuenta Corriente";
			// 
			// AsientoCuenta
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 14);
			this.ClientSize = new System.Drawing.Size(498, 320);
			this.Controls.Add(this.chkCorriente);
			this.Controls.Add(this.label9);
			this.Controls.Add(this.label8);
			this.Controls.Add(this.label7);
			this.Controls.Add(this.label6);
			this.Controls.Add(this.txtIdCuentaCorr);
			this.Controls.Add(this.label5);
			this.Controls.Add(this.label4);
			this.Controls.Add(this.label3);
			this.Controls.Add(this.label2);
			this.Controls.Add(this.label1);
			this.Controls.Add(this.dtDesde);
			this.Controls.Add(this.txtHasta);
			this.Controls.Add(this.txtDesde);
			this.Controls.Add(this.groupBox2);
			this.Controls.Add(this.groupBox1);
			this.Controls.Add(this.chkProvisiones);
			this.Controls.Add(this.chkPagoNomina);
			this.Controls.Add(this.rdgTipo);
			this.Controls.Add(this.ultraCheckEditor2);
			this.Controls.Add(this.btTipo);
			this.Controls.Add(this.lblTipo);
			this.Controls.Add(this.cmbTipo);
			this.Controls.Add(this.barraDato1);
			this.Controls.Add(this.cmbBanco);
			this.Controls.Add(this.txtNumCuenta);
			this.Controls.Add(this.txtNumCheque);
			this.Controls.Add(this.txtReporte);
			this.Controls.Add(this.cmbCuentaNombre);
			this.Controls.Add(this.cmbCuentaCod);
			this.Controls.Add(this.chkAsiento);
			this.Controls.Add(this.ultraCheckEditor1);
			this.Controls.Add(this.btPlanCuentas);
			this.Controls.Add(this.btBanco);
			this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
			this.MaximizeBox = false;
			this.Name = "AsientoCuenta";
			this.Text = "Cuentas Corrientes";
			this.Closing += new System.ComponentModel.CancelEventHandler(this.AsientoCuenta_Closing);
			this.Load += new System.EventHandler(this.AsientoCuenta_Load);
			((System.ComponentModel.ISupportInitialize)(this.cdsCuentaCorr)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvBanco)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvCuentaCod)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsCuenta)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvCuenta)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeguridad)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCuentaCod)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCuentaNombre)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtReporte)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNumCheque)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNumCuenta)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbBanco)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbTipo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.rdgTipo.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDesde)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtHasta)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtDesde)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void barraDato1_Borrar(object sender, System.EventArgs e)
		{
			if (!miAcceso.Eliminar)
			{
				MessageBox.Show("No tiene Acceso a Eliminar Cuentas Corrientes",
					"Información", MessageBoxButtons.OK, MessageBoxIcon.Information);
				return;
			}

			if (txtIdCuentaCorr.Text.Trim().Length == 0)
			{
				MessageBox.Show("No Existen Cuentas Corrientes para Borrar",
					"Información", MessageBoxButtons.OK, MessageBoxIcon.Information);
				return;
			}

			if (MessageBox.Show("¿Desea Borrar esta Cuenta Corriente?", "Confirmación",
				MessageBoxButtons.YesNo, MessageBoxIcon.Warning,
				MessageBoxDefaultButton.Button2) == DialogResult.No) 
				return;

			this.barraDato1.BorraRegistro();
		}

		private void barraDato1_Crear(object sender, System.EventArgs e)
		{
			if (!miAcceso.Nuevo)
			{
				MessageBox.Show("No tiene acceso a Crear Nuevas Cuentas",
					"Información", MessageBoxButtons.OK, MessageBoxIcon.Information);
				return;
			}

			this.barraDato1.CrearRegistro();
		}

		private void barraDato1_Deshacer(object sender, System.EventArgs e)
		{
			this.BorrarErrores();
			this.barraDato1.DeshacerRegistro();
		}

		private void barraDato1_Editar(object sender, System.EventArgs e)
		{
			if (!miAcceso.Editar)
			{
				MessageBox.Show("No tiene Acceso a Editar Cuentas Corrientes",
					"Información", MessageBoxButtons.OK, MessageBoxIcon.Information);
				return;
			}
			if (txtIdCuentaCorr.Text.Trim().Length == 0)
			{
				MessageBox.Show("No Existen Cuentas Corrientes para Editar",
					"Información", MessageBoxButtons.OK, MessageBoxIcon.Information);
				return;
			}
			this.barraDato1.EditarRegistro();
		}

		private void BorrarErrores()
		{
			this.errorProvider.SetError(this.cmbCuentaCod, "");
			this.errorProvider.SetError(this.txtNumCuenta, "");
		}

		private bool Verifica()
		{
			bool bOk = true;
			this.BorrarErrores();
			this.barraDato1.Focus();
			if (cmbCuentaCod.Value == System.DBNull.Value ||
				cmbCuentaCod.Value == null || 
				(int) cmbCuentaCod.Value < 1)
			{
				errorProvider.SetError(cmbCuentaCod, "Debe ingresar cuenta del plan de cuentas");
				bOk = false;
			}
			else
			{
				// Verificacion de Cuenta Repetida
				this.cdsCuenta.Schema.Connections[0].Open();
				C1.Data.SchemaObjects.Connection miConeccion = this.cdsCuenta.Schema.Connections[0];
				string stSelect = "SELECT COUNT(*) FROM CuentaCorriente WHERE idCuenta = " +
					this.cmbCuentaCod.Value.ToString().Trim();

				if (!this.barraDato1.bNuevo)
					stSelect += " And idCuentaCorriente <> " +
						this.txtIdCuentaCorr.Text.ToString().Trim();

				SqlCommand cmdCodigo = new SqlCommand (
					stSelect, (SqlConnection) miConeccion.DbConnection);
				int iCuenta = (int) cmdCodigo.ExecuteScalar();
				if (iCuenta > 0) 
				{
					this.errorProvider.SetError(this.cmbCuentaCod,"Código de Cuenta ya existe");
					bOk = false;
				}
				this.cdsCuenta.Schema.Connections[0].Close();
				// Fin de verificacion de Código repetido
			}
			
			if (this.txtNumCuenta.Text.Trim().Length == 0)
			{
				this.errorProvider.SetError(this.txtNumCuenta, "Ingrese Número de Cuenta");
				bOk = false;
			}
			else if (this.txtDesde.Text.ToString().Trim().Length == 0)
			{
				this.errorProvider.SetError(this.txtDesde, "Ingrese el numero inicial de la chequera");
				bOk = false;
			}
			else if (this.txtHasta.Text.ToString().Trim().Length == 0)
			{
				this.errorProvider.SetError(this.txtHasta, "Ingrese el numero final de la chequera");
				bOk = false;
			}
			else if (int.Parse(this.txtDesde.Text.ToString().Trim()) > int.Parse(this.txtHasta.Text.ToString().Trim()))
			{
				this.errorProvider.SetError(this.txtHasta, "El numero inicial no puede ser mayor al valor final de la chequera");
				bOk = false;
			}
			else
			{
				// Verificacion de Numero de Cuenta Repetida
				this.cdsCuenta.Schema.Connections[0].Open();
				C1.Data.SchemaObjects.Connection miConeccion = this.cdsCuenta.Schema.Connections[0];
				string stSelect = "SELECT COUNT(*) FROM CuentaCorriente WHERE NumCuenta = '" +
					this.txtNumCuenta.Text.Trim() + "'";

				if (!this.barraDato1.bNuevo)
					stSelect += " And idCuentaCorriente <> " +
						this.txtIdCuentaCorr.Text.ToString().Trim();

				SqlCommand cmdCodigo = new SqlCommand (
					stSelect, (SqlConnection) miConeccion.DbConnection);
				int iCuenta = (int) cmdCodigo.ExecuteScalar();
				if (iCuenta > 0) 
				{
					this.errorProvider.SetError(this.txtNumCuenta,"Número de Cuenta ya existe");
					bOk = false;
				}
				this.cdsCuenta.Schema.Connections[0].Close();
				// Fin de verificacion de Código repetido
			}
			return bOk;
		}

		private void barraDato1_Grabar(object sender, System.EventArgs e)
		{
			if (!Verifica()) return;
			this.barraDato1.GrabaRegistro();
			this.barraDato1.ProximoId(5);
		}

		private void btPlanCuentas_Click(object sender, System.EventArgs e)
		{
			PlanCuentas miCuenta = new PlanCuentas();
			miCuenta.MdiParent = this.MdiParent;
			miCuenta.Show();
		}

		private Acceso miAcceso;
		C1.Data.C1DataRow drSeteo;

		private void AsientoCuenta_Load(object sender, System.EventArgs e)
		{
			Cursor = Cursors.Default;
			string stAudita = "Exec AuditaCrear 53, 6, 'Cuenta Corr'";
			Funcion.EjecutaSQL(cdsCuenta, stAudita, true);
			drSeteo = cdsSeteo.TableViews["SeteoC"].Rows[0];
			if (drSeteo["Gobierno"] != System.DBNull.Value && (bool) drSeteo["Gobierno"] == true)
			{
				lblTipo.Visible = true;
				cmbTipo.Visible = true;
				btTipo.Visible = true;
			}

			miAcceso =new Acceso("040802", cdsSeguridad.StorageDataSet.Tables["Seguridad"]);

			this.txtIdCuentaCorr.Width = 0;
			this.barraDato1.PosUltima();
			this.barraDato1.HabilitaControles(false);
		}

		private void cdsCuentaCorr_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			this.cdsCuentaCorr.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirConta);		
		}

		private void cdsCuenta_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			this.cdsCuenta.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirConta);		
		}

		private void btBanco_Click(object sender, System.EventArgs e)
		{
			Bancos miBanco = new Bancos();
			miBanco.MdiParent = this.MdiParent;
			miBanco.Show();
		}

		private void barraDatoSQL1_Refresca(object sender, System.EventArgs e)
		{
		
		}

		private void barraDatoSQL1_Restaurar(object sender, System.EventArgs e)
		{
		
		}

		private void barraDatoSQL1_Copiar(object sender, System.EventArgs e)
		{
		
		}

		private void btTipo_Click(object sender, System.EventArgs e)
		{
			CuentaCorrTipo miCuenta = new CuentaCorrTipo();
			miCuenta.MdiParent = this.MdiParent;
			miCuenta.Show();
		}

		private void cdsSeteo_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			this.cdsSeteo.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirConta);		
		}

		private void AsientoCuenta_Closing(object sender, System.ComponentModel.CancelEventArgs e)
		{
			string stAudita = "Exec AuditaCrear 53, 7, 'Cuenta. Corr.'";
			Funcion.EjecutaSQL(cdsCuenta, stAudita, true);
		}

		private void cmbCuentaCod_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}
	}
}
