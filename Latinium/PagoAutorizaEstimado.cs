using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;
using System.Data.SqlClient;
using Infragistics.Win.UltraWinGrid;
using CrystalDecisions.Shared;


namespace Latinium
{
	/// <summary>
	/// Descripción breve de PagoAutorizaEstimado.
	/// </summary>
	public class PagoAutorizaEstimado : DevExpress.XtraEditors.XtraForm
	{
		private Infragistics.Win.UltraWinEditors.UltraTextEditor ultraTextEditor14;
		private System.Windows.Forms.Label label2;
		private System.Windows.Forms.Label label3;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtTotal;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtPagos;
		private System.Windows.Forms.Label label6;
		private System.Windows.Forms.Label label7;
		private C1.Data.C1DataView cdvCodCuenta;
		private C1.Data.C1DataSet cdsCuenta;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbCuentaCod;
		private System.Windows.Forms.GroupBox groupBox1;
		private System.Windows.Forms.Label label5;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtPresupuesto;
		private System.Windows.Forms.Label label4;
		private System.Windows.Forms.GroupBox groupBox2;
		private C1.Data.C1DataSet cdsProyecto;
		private DevExpress.XtraEditors.SimpleButton btGenera;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtIngresos;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtEgresos;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkTodoTipo;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbCentroCosto;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo cmbFecha;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo cmbHasta;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtSaldo;
		private System.Windows.Forms.ErrorProvider errorProvider;
		private DevExpress.XtraEditors.SimpleButton btAsiento;
		private DevExpress.XtraEditors.SimpleButton btImprime;
		private System.Windows.Forms.GroupBox groupBox3;
		private C1.Data.C1DataSet cdsPagoAutoriza;
		private Infragistics.Win.UltraWinGrid.UltraGrid ultraGrid1;
		private Infragistics.Win.UltraWinGrid.UltraGrid ultraGrid2;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkCuenta;
		private DevExpress.XtraEditors.SimpleButton btfactura;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public PagoAutorizaEstimado()
		{
			//
			// Necesario para admitir el Diseñador de Windows Forms
			//
			InitializeComponent();

			//
			// TODO: agregar código de constructor después de llamar a InitializeComponent
			//
		}

		private void btSalir_Click(object sender, System.EventArgs e)
		{
			
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
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton1 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Cuenta", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCuenta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idSucursal");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Padre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Descripcion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CodRapido");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CodPresup");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("DescIngles");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Borrar");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Grupo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn11 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Flujo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn12 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Tipo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn13 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaIngreso");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn14 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Editar");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn15 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nuevo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn16 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaEditado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn17 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCuentaFlujo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn18 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCuentaEjecucionP");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn19 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCuentaBalance");
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Proyecto", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn20 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idProyecto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn21 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPadre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn22 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn23 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Lugar");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn24 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Responsable");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn25 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Notas");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn26 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CentroCosto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn27 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaIngreso");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn28 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Visible");
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton2 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance11 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance12 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance13 = new Infragistics.Win.Appearance();
			this.ultraTextEditor14 = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label2 = new System.Windows.Forms.Label();
			this.label3 = new System.Windows.Forms.Label();
			this.txtTotal = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtPagos = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label6 = new System.Windows.Forms.Label();
			this.label7 = new System.Windows.Forms.Label();
			this.cmbFecha = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.cdvCodCuenta = new C1.Data.C1DataView();
			this.cdsCuenta = new C1.Data.C1DataSet();
			this.cmbCuentaCod = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.groupBox1 = new System.Windows.Forms.GroupBox();
			this.txtSaldo = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label5 = new System.Windows.Forms.Label();
			this.txtPresupuesto = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label4 = new System.Windows.Forms.Label();
			this.groupBox2 = new System.Windows.Forms.GroupBox();
			this.ultraGrid1 = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.cdsProyecto = new C1.Data.C1DataSet();
			this.txtIngresos = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtEgresos = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.cmbCentroCosto = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.btGenera = new DevExpress.XtraEditors.SimpleButton();
			this.chkTodoTipo = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.cmbHasta = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.errorProvider = new System.Windows.Forms.ErrorProvider();
			this.btAsiento = new DevExpress.XtraEditors.SimpleButton();
			this.btImprime = new DevExpress.XtraEditors.SimpleButton();
			this.groupBox3 = new System.Windows.Forms.GroupBox();
			this.ultraGrid2 = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.cdsPagoAutoriza = new C1.Data.C1DataSet();
			this.chkCuenta = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.btfactura = new DevExpress.XtraEditors.SimpleButton();
			((System.ComponentModel.ISupportInitialize)(this.ultraTextEditor14)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtTotal)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtPagos)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbFecha)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvCodCuenta)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsCuenta)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCuentaCod)).BeginInit();
			this.groupBox1.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.txtSaldo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtPresupuesto)).BeginInit();
			this.groupBox2.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsProyecto)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIngresos)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtEgresos)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCentroCosto)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbHasta)).BeginInit();
			this.groupBox3.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid2)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsPagoAutoriza)).BeginInit();
			this.SuspendLayout();
			// 
			// ultraTextEditor14
			// 
			appearance1.ForeColorDisabled = System.Drawing.Color.Black;
			this.ultraTextEditor14.Appearance = appearance1;
			this.ultraTextEditor14.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.ultraTextEditor14.Location = new System.Drawing.Point(116, 82);
			this.ultraTextEditor14.Name = "ultraTextEditor14";
			this.ultraTextEditor14.Size = new System.Drawing.Size(128, 21);
			this.ultraTextEditor14.TabIndex = 5;
			// 
			// label2
			// 
			this.label2.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.label2.AutoSize = true;
			this.label2.Location = new System.Drawing.Point(200, 392);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(59, 16);
			this.label2.TabIndex = 49;
			this.label2.Text = "Pagos Aut.";
			// 
			// label3
			// 
			this.label3.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.label3.AutoSize = true;
			this.label3.Location = new System.Drawing.Point(512, 24);
			this.label3.Name = "label3";
			this.label3.Size = new System.Drawing.Size(29, 16);
			this.label3.TabIndex = 50;
			this.label3.Text = "Total";
			// 
			// txtTotal
			// 
			this.txtTotal.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.txtTotal.Enabled = false;
			this.txtTotal.FormatString = "#,##0.00";
			this.txtTotal.Location = new System.Drawing.Point(576, 24);
			this.txtTotal.Name = "txtTotal";
			this.txtTotal.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtTotal.PromptChar = ' ';
			this.txtTotal.Size = new System.Drawing.Size(128, 21);
			this.txtTotal.TabIndex = 54;
			// 
			// txtPagos
			// 
			this.txtPagos.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.txtPagos.Enabled = false;
			this.txtPagos.FormatString = "#,##0.00";
			this.txtPagos.Location = new System.Drawing.Point(288, 384);
			this.txtPagos.Name = "txtPagos";
			this.txtPagos.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtPagos.PromptChar = ' ';
			this.txtPagos.Size = new System.Drawing.Size(88, 21);
			this.txtPagos.TabIndex = 59;
			// 
			// label6
			// 
			this.label6.AutoSize = true;
			this.label6.Location = new System.Drawing.Point(288, 16);
			this.label6.Name = "label6";
			this.label6.Size = new System.Drawing.Size(68, 16);
			this.label6.TabIndex = 65;
			this.label6.Text = "Fecha Hasta";
			// 
			// label7
			// 
			this.label7.AutoSize = true;
			this.label7.Location = new System.Drawing.Point(24, 16);
			this.label7.Name = "label7";
			this.label7.Size = new System.Drawing.Size(71, 16);
			this.label7.TabIndex = 63;
			this.label7.Text = "Fecha Desde";
			// 
			// cmbFecha
			// 
			this.cmbFecha.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton1.Caption = "Today";
			this.cmbFecha.DateButtons.Add(dateButton1);
			this.cmbFecha.Format = "dd/MMM/yyyy";
			this.cmbFecha.Location = new System.Drawing.Point(128, 16);
			this.cmbFecha.Name = "cmbFecha";
			this.cmbFecha.NonAutoSizeHeight = 23;
			this.cmbFecha.Size = new System.Drawing.Size(136, 21);
			this.cmbFecha.SpinButtonsVisible = true;
			this.cmbFecha.TabIndex = 62;
			this.cmbFecha.Value = new System.DateTime(2009, 8, 31, 0, 0, 0, 0);
			// 
			// cdvCodCuenta
			// 
			this.cdvCodCuenta.BindingContextCtrl = this;
			this.cdvCodCuenta.DataSet = this.cdsCuenta;
			this.cdvCodCuenta.RowFilter = "CodPresup > \'\'";
			this.cdvCodCuenta.Sort = "Codigo";
			this.cdvCodCuenta.TableName = "";
			this.cdvCodCuenta.TableViewName = "Cuenta";
			// 
			// cdsCuenta
			// 
			this.cdsCuenta.BindingContextCtrl = this;
			this.cdsCuenta.DataLibrary = "LibContabilidad";
			this.cdsCuenta.DataLibraryUrl = "";
			this.cdsCuenta.DataSetDef = "dsCuenta";
			this.cdsCuenta.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsCuenta.Name = "cdsCuenta";
			this.cdsCuenta.SchemaClassName = "LibContabilidad.DataClass";
			this.cdsCuenta.SchemaDef = null;
			// 
			// cmbCuentaCod
			// 
			appearance2.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbCuentaCod.Appearance = appearance2;
			this.cmbCuentaCod.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbCuentaCod.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbCuentaCod.DataSource = this.cdvCodCuenta;
			ultraGridColumn1.Header.VisiblePosition = 6;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridColumn2.Hidden = true;
			ultraGridColumn3.Header.VisiblePosition = 3;
			ultraGridColumn3.Hidden = true;
			ultraGridColumn4.Header.VisiblePosition = 4;
			ultraGridColumn4.Hidden = true;
			ultraGridColumn5.Header.VisiblePosition = 2;
			ultraGridColumn5.Width = 368;
			ultraGridColumn6.Header.Caption = "Codigo";
			ultraGridColumn6.Header.VisiblePosition = 0;
			ultraGridColumn6.Width = 151;
			ultraGridColumn7.Header.VisiblePosition = 7;
			ultraGridColumn7.Hidden = true;
			ultraGridColumn8.Header.VisiblePosition = 8;
			ultraGridColumn8.Hidden = true;
			ultraGridColumn9.Header.VisiblePosition = 9;
			ultraGridColumn9.Hidden = true;
			ultraGridColumn10.Header.VisiblePosition = 5;
			ultraGridColumn10.Hidden = true;
			ultraGridColumn11.Header.VisiblePosition = 10;
			ultraGridColumn11.Hidden = true;
			ultraGridColumn12.Header.VisiblePosition = 11;
			ultraGridColumn12.Hidden = true;
			ultraGridColumn13.Header.VisiblePosition = 12;
			ultraGridColumn13.Hidden = true;
			ultraGridColumn14.Header.VisiblePosition = 13;
			ultraGridColumn14.Hidden = true;
			ultraGridColumn15.Header.VisiblePosition = 14;
			ultraGridColumn15.Hidden = true;
			ultraGridColumn16.Header.VisiblePosition = 15;
			ultraGridColumn16.Hidden = true;
			ultraGridColumn17.Header.VisiblePosition = 16;
			ultraGridColumn17.Hidden = true;
			ultraGridColumn18.Header.VisiblePosition = 17;
			ultraGridColumn18.Hidden = true;
			ultraGridColumn19.Header.VisiblePosition = 18;
			ultraGridColumn19.Hidden = true;
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
																										 ultraGridColumn19});
			this.cmbCuentaCod.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.cmbCuentaCod.DisplayMember = "Codigo";
			this.cmbCuentaCod.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbCuentaCod.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbCuentaCod.Enabled = false;
			this.cmbCuentaCod.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbCuentaCod.Location = new System.Drawing.Point(128, 56);
			this.cmbCuentaCod.Name = "cmbCuentaCod";
			this.cmbCuentaCod.Size = new System.Drawing.Size(136, 21);
			this.cmbCuentaCod.TabIndex = 67;
			this.cmbCuentaCod.ValueMember = "idCuenta";
			// 
			// groupBox1
			// 
			this.groupBox1.Controls.Add(this.txtSaldo);
			this.groupBox1.Controls.Add(this.label5);
			this.groupBox1.Controls.Add(this.txtPresupuesto);
			this.groupBox1.Controls.Add(this.label4);
			this.groupBox1.Controls.Add(this.label3);
			this.groupBox1.Controls.Add(this.txtTotal);
			this.groupBox1.Location = new System.Drawing.Point(16, 88);
			this.groupBox1.Name = "groupBox1";
			this.groupBox1.Size = new System.Drawing.Size(720, 64);
			this.groupBox1.TabIndex = 71;
			this.groupBox1.TabStop = false;
			this.groupBox1.Enter += new System.EventHandler(this.groupBox1_Enter);
			// 
			// txtSaldo
			// 
			this.txtSaldo.Enabled = false;
			this.txtSaldo.FormatString = "#,##0.00";
			this.txtSaldo.Location = new System.Drawing.Point(384, 24);
			this.txtSaldo.Name = "txtSaldo";
			this.txtSaldo.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtSaldo.PromptChar = ' ';
			this.txtSaldo.Size = new System.Drawing.Size(104, 21);
			this.txtSaldo.TabIndex = 65;
			this.txtSaldo.ValueChanged += new System.EventHandler(this.txtSaldo_ValueChanged);
			// 
			// label5
			// 
			this.label5.AutoSize = true;
			this.label5.Location = new System.Drawing.Point(256, 24);
			this.label5.Name = "label5";
			this.label5.Size = new System.Drawing.Size(79, 16);
			this.label5.TabIndex = 64;
			this.label5.Text = "SaldoContable";
			// 
			// txtPresupuesto
			// 
			this.txtPresupuesto.Enabled = false;
			this.txtPresupuesto.FormatString = "#,##0.00";
			this.txtPresupuesto.Location = new System.Drawing.Point(136, 24);
			this.txtPresupuesto.Name = "txtPresupuesto";
			this.txtPresupuesto.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtPresupuesto.PromptChar = ' ';
			this.txtPresupuesto.Size = new System.Drawing.Size(97, 21);
			this.txtPresupuesto.TabIndex = 63;
			this.txtPresupuesto.Value = 1;
			// 
			// label4
			// 
			this.label4.AutoSize = true;
			this.label4.Location = new System.Drawing.Point(16, 24);
			this.label4.Name = "label4";
			this.label4.Size = new System.Drawing.Size(80, 16);
			this.label4.TabIndex = 62;
			this.label4.Text = "Presupuestado";
			// 
			// groupBox2
			// 
			this.groupBox2.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.groupBox2.Controls.Add(this.ultraGrid1);
			this.groupBox2.Location = new System.Drawing.Point(392, 160);
			this.groupBox2.Name = "groupBox2";
			this.groupBox2.Size = new System.Drawing.Size(344, 216);
			this.groupBox2.TabIndex = 72;
			this.groupBox2.TabStop = false;
			this.groupBox2.Text = "INGRESOS -EGRESOS";
			// 
			// ultraGrid1
			// 
			this.ultraGrid1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.ultraGrid1.Cursor = System.Windows.Forms.Cursors.Default;
			appearance3.BackColor = System.Drawing.Color.White;
			this.ultraGrid1.DisplayLayout.Appearance = appearance3;
			appearance4.BackColor = System.Drawing.Color.Transparent;
			this.ultraGrid1.DisplayLayout.Override.CardAreaAppearance = appearance4;
			appearance5.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance5.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance5.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance5.FontData.BoldAsString = "True";
			appearance5.FontData.Name = "Arial";
			appearance5.FontData.SizeInPoints = 10F;
			appearance5.ForeColor = System.Drawing.Color.White;
			appearance5.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.ultraGrid1.DisplayLayout.Override.HeaderAppearance = appearance5;
			appearance6.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance6.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance6.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.RowSelectorAppearance = appearance6;
			appearance7.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance7.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance7.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.SelectedRowAppearance = appearance7;
			this.ultraGrid1.Location = new System.Drawing.Point(8, 16);
			this.ultraGrid1.Name = "ultraGrid1";
			this.ultraGrid1.Size = new System.Drawing.Size(328, 192);
			this.ultraGrid1.TabIndex = 49;
			// 
			// cdsProyecto
			// 
			this.cdsProyecto.BindingContextCtrl = this;
			this.cdsProyecto.DataLibrary = "LibContabilidad";
			this.cdsProyecto.DataLibraryUrl = "";
			this.cdsProyecto.DataSetDef = "dsProyecto";
			this.cdsProyecto.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsProyecto.Name = "cdsProyecto";
			this.cdsProyecto.SchemaClassName = "LibContabilidad.DataClass";
			this.cdsProyecto.SchemaDef = null;
			// 
			// txtIngresos
			// 
			this.txtIngresos.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.txtIngresos.Enabled = false;
			this.txtIngresos.FormatString = "#,##0.00";
			this.txtIngresos.Location = new System.Drawing.Point(560, 384);
			this.txtIngresos.Name = "txtIngresos";
			this.txtIngresos.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtIngresos.PromptChar = ' ';
			this.txtIngresos.Size = new System.Drawing.Size(80, 21);
			this.txtIngresos.TabIndex = 74;
			this.txtIngresos.Visible = false;
			// 
			// txtEgresos
			// 
			this.txtEgresos.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.txtEgresos.FormatString = "#,##0.00";
			this.txtEgresos.Location = new System.Drawing.Point(640, 384);
			this.txtEgresos.Name = "txtEgresos";
			this.txtEgresos.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtEgresos.PromptChar = ' ';
			this.txtEgresos.Size = new System.Drawing.Size(80, 21);
			this.txtEgresos.TabIndex = 75;
			this.txtEgresos.Visible = false;
			// 
			// cmbCentroCosto
			// 
			appearance8.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbCentroCosto.Appearance = appearance8;
			this.cmbCentroCosto.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbCentroCosto.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbCentroCosto.DataMember = "Proyecto";
			this.cmbCentroCosto.DataSource = this.cdsProyecto;
			ultraGridBand2.AddButtonCaption = "Cuenta";
			ultraGridColumn20.Header.VisiblePosition = 0;
			ultraGridColumn20.Hidden = true;
			ultraGridColumn21.Header.VisiblePosition = 1;
			ultraGridColumn21.Hidden = true;
			ultraGridColumn22.Header.VisiblePosition = 2;
			ultraGridColumn23.Header.VisiblePosition = 3;
			ultraGridColumn23.Hidden = true;
			ultraGridColumn24.Header.VisiblePosition = 4;
			ultraGridColumn24.Hidden = true;
			ultraGridColumn25.Header.VisiblePosition = 5;
			ultraGridColumn25.Hidden = true;
			ultraGridColumn26.Header.VisiblePosition = 6;
			ultraGridColumn27.Header.VisiblePosition = 8;
			ultraGridColumn27.Hidden = true;
			ultraGridColumn28.Header.VisiblePosition = 7;
			ultraGridColumn28.Hidden = true;
			ultraGridBand2.Columns.AddRange(new object[] {
																										 ultraGridColumn20,
																										 ultraGridColumn21,
																										 ultraGridColumn22,
																										 ultraGridColumn23,
																										 ultraGridColumn24,
																										 ultraGridColumn25,
																										 ultraGridColumn26,
																										 ultraGridColumn27,
																										 ultraGridColumn28});
			this.cmbCentroCosto.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			this.cmbCentroCosto.DisplayMember = "CentroCosto";
			this.cmbCentroCosto.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbCentroCosto.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbCentroCosto.Enabled = false;
			this.cmbCentroCosto.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbCentroCosto.Location = new System.Drawing.Point(432, 56);
			this.cmbCentroCosto.Name = "cmbCentroCosto";
			this.cmbCentroCosto.Size = new System.Drawing.Size(136, 21);
			this.cmbCentroCosto.TabIndex = 76;
			this.cmbCentroCosto.ValueMember = "idProyecto";
			// 
			// btGenera
			// 
			this.btGenera.Location = new System.Drawing.Point(640, 16);
			this.btGenera.Name = "btGenera";
			this.btGenera.Size = new System.Drawing.Size(96, 23);
			this.btGenera.TabIndex = 78;
			this.btGenera.Text = "GENERAR";
			this.btGenera.Click += new System.EventHandler(this.btGenera_Click);
			// 
			// chkTodoTipo
			// 
			this.chkTodoTipo.Checked = true;
			this.chkTodoTipo.CheckState = System.Windows.Forms.CheckState.Checked;
			this.chkTodoTipo.Location = new System.Drawing.Point(288, 56);
			this.chkTodoTipo.Name = "chkTodoTipo";
			this.chkTodoTipo.Size = new System.Drawing.Size(128, 16);
			this.chkTodoTipo.TabIndex = 79;
			this.chkTodoTipo.Text = "Todo Centro Costo";
			this.chkTodoTipo.CheckedChanged += new System.EventHandler(this.chkTodoTipo_CheckedChanged);
			// 
			// cmbHasta
			// 
			this.cmbHasta.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton2.Caption = "Today";
			this.cmbHasta.DateButtons.Add(dateButton2);
			this.cmbHasta.Format = "dd/MMM/yyyy";
			this.cmbHasta.Location = new System.Drawing.Point(376, 16);
			this.cmbHasta.Name = "cmbHasta";
			this.cmbHasta.NonAutoSizeHeight = 23;
			this.cmbHasta.Size = new System.Drawing.Size(136, 21);
			this.cmbHasta.SpinButtonsVisible = true;
			this.cmbHasta.TabIndex = 80;
			this.cmbHasta.Value = new System.DateTime(2009, 8, 31, 0, 0, 0, 0);
			// 
			// errorProvider
			// 
			this.errorProvider.ContainerControl = this;
			this.errorProvider.DataMember = "";
			// 
			// btAsiento
			// 
			this.btAsiento.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btAsiento.Location = new System.Drawing.Point(392, 384);
			this.btAsiento.Name = "btAsiento";
			this.btAsiento.Size = new System.Drawing.Size(32, 23);
			this.btAsiento.TabIndex = 81;
			this.btAsiento.Text = "A";
			this.btAsiento.Click += new System.EventHandler(this.btAsiento_Click);
			// 
			// btImprime
			// 
			this.btImprime.Location = new System.Drawing.Point(640, 56);
			this.btImprime.Name = "btImprime";
			this.btImprime.Size = new System.Drawing.Size(96, 23);
			this.btImprime.TabIndex = 82;
			this.btImprime.Text = "IMPRIME";
			this.btImprime.Click += new System.EventHandler(this.simpleButton2_Click);
			// 
			// groupBox3
			// 
			this.groupBox3.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left)));
			this.groupBox3.Controls.Add(this.ultraGrid2);
			this.groupBox3.Location = new System.Drawing.Point(16, 160);
			this.groupBox3.Name = "groupBox3";
			this.groupBox3.Size = new System.Drawing.Size(368, 216);
			this.groupBox3.TabIndex = 83;
			this.groupBox3.TabStop = false;
			this.groupBox3.Text = "PAGOS";
			// 
			// ultraGrid2
			// 
			this.ultraGrid2.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.ultraGrid2.Cursor = System.Windows.Forms.Cursors.Default;
			appearance9.BackColor = System.Drawing.Color.White;
			this.ultraGrid2.DisplayLayout.Appearance = appearance9;
			appearance10.BackColor = System.Drawing.Color.Transparent;
			this.ultraGrid2.DisplayLayout.Override.CardAreaAppearance = appearance10;
			appearance11.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance11.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance11.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance11.FontData.BoldAsString = "True";
			appearance11.FontData.Name = "Arial";
			appearance11.FontData.SizeInPoints = 10F;
			appearance11.ForeColor = System.Drawing.Color.White;
			appearance11.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.ultraGrid2.DisplayLayout.Override.HeaderAppearance = appearance11;
			appearance12.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance12.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance12.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid2.DisplayLayout.Override.RowSelectorAppearance = appearance12;
			appearance13.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance13.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance13.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid2.DisplayLayout.Override.SelectedRowAppearance = appearance13;
			this.ultraGrid2.Location = new System.Drawing.Point(8, 16);
			this.ultraGrid2.Name = "ultraGrid2";
			this.ultraGrid2.Size = new System.Drawing.Size(352, 192);
			this.ultraGrid2.TabIndex = 50;
			// 
			// cdsPagoAutoriza
			// 
			this.cdsPagoAutoriza.BindingContextCtrl = this;
			this.cdsPagoAutoriza.DataLibrary = "LibFacturacion";
			this.cdsPagoAutoriza.DataLibraryUrl = "";
			this.cdsPagoAutoriza.DataSetDef = "dsPagoAutoriza";
			this.cdsPagoAutoriza.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsPagoAutoriza.Name = "cdsPagoAutoriza";
			this.cdsPagoAutoriza.SchemaClassName = "LibFacturacion.DataClass";
			this.cdsPagoAutoriza.SchemaDef = null;
			// 
			// chkCuenta
			// 
			this.chkCuenta.Checked = true;
			this.chkCuenta.CheckState = System.Windows.Forms.CheckState.Checked;
			this.chkCuenta.Location = new System.Drawing.Point(24, 56);
			this.chkCuenta.Name = "chkCuenta";
			this.chkCuenta.Size = new System.Drawing.Size(104, 16);
			this.chkCuenta.TabIndex = 84;
			this.chkCuenta.Text = "Toda Cuenta";
			this.chkCuenta.CheckedChanged += new System.EventHandler(this.chkCuenta_CheckedChanged);
			// 
			// btfactura
			// 
			this.btfactura.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btfactura.Location = new System.Drawing.Point(32, 384);
			this.btfactura.Name = "btfactura";
			this.btfactura.Size = new System.Drawing.Size(32, 23);
			this.btfactura.TabIndex = 85;
			this.btfactura.Text = "F";
			this.btfactura.Click += new System.EventHandler(this.btfactura_Click);
			// 
			// PagoAutorizaEstimado
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(752, 421);
			this.Controls.Add(this.btfactura);
			this.Controls.Add(this.chkCuenta);
			this.Controls.Add(this.groupBox3);
			this.Controls.Add(this.btImprime);
			this.Controls.Add(this.btAsiento);
			this.Controls.Add(this.cmbHasta);
			this.Controls.Add(this.chkTodoTipo);
			this.Controls.Add(this.btGenera);
			this.Controls.Add(this.label6);
			this.Controls.Add(this.label7);
			this.Controls.Add(this.cmbCentroCosto);
			this.Controls.Add(this.txtEgresos);
			this.Controls.Add(this.txtIngresos);
			this.Controls.Add(this.groupBox2);
			this.Controls.Add(this.groupBox1);
			this.Controls.Add(this.cmbCuentaCod);
			this.Controls.Add(this.cmbFecha);
			this.Controls.Add(this.txtPagos);
			this.Controls.Add(this.label2);
			this.Name = "PagoAutorizaEstimado";
			this.Text = "Flujo de Fondos";
			this.Load += new System.EventHandler(this.PagoAutorizaEstimado_Load);
			((System.ComponentModel.ISupportInitialize)(this.ultraTextEditor14)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtTotal)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtPagos)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbFecha)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvCodCuenta)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsCuenta)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCuentaCod)).EndInit();
			this.groupBox1.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.txtSaldo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtPresupuesto)).EndInit();
			this.groupBox2.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsProyecto)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIngresos)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtEgresos)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCentroCosto)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbHasta)).EndInit();
			this.groupBox3.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid2)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsPagoAutoriza)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion
		DataSet miDataSet1 = new DataSet();
		DataView dv1 = new DataView();
		DataSet miDataSetG1 = new DataSet();
		DataView dvG1 = new DataView();
		DataSet miDataSetT1 = new DataSet();
		DataView dvT1 = new DataView();

		DataSet miDataSet = new DataSet();
		DataView dv = new DataView();
		DataSet miDataSetG = new DataSet();
		DataView dvG = new DataView();
		DataSet miDataSetT = new DataSet();
		DataView dvT = new DataView();
		double DPagos = 0.00;
		double dPresupuesto = 0.00;
		double dSaldo = 0.00;
		double dTotal = 0.00;
		int iPagoLote = 0;

		public PagoAutorizaEstimado(double dPagos, int idPagoLote)
		{
			InitializeComponent();
			Funcion miFuncion = new Funcion();
			string strConn = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);	
			DPagos =  dPagos;
			iPagoLote = idPagoLote;
			txtPagos.Value = DPagos;
			#region llena pagos
			miFuncion = new Funcion();
			strConn = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);
			string stProcedimiento = "exec PagoLoteE " + idPagoLote;
			SqlDataAdapter myData = new SqlDataAdapter(stProcedimiento, strConn);
			//			myData.TableMappings.Add("Tabla", "ExcelSube");
			miDataSetG1.Reset();
			miDataSetG1.Clear();
			try
			{
				myData.Fill(miDataSetG1);
			}
			catch(Exception ex)
			{
				MessageBox.Show(ex.Message, "Error al abrir la Hoja A del Archivo: " );
				return;
			}
			dvG1 = miDataSetG1.Tables[0].DefaultView;
			ultraGrid2.DataSource = dvG1;

			#endregion
			#region Visibilidad
			ultraGrid2.DisplayLayout.Bands[0].Columns["idCompra"].Hidden = true;
			#endregion Visibilidad

		}		
		private void PagoAutorizaEstimado_Load(object sender, System.EventArgs e)
		{
		}
		private void btGenera_Click(object sender, System.EventArgs e)
		{
			DateTime dtFechaIni = (DateTime) this.cmbFecha.Value;
			DateTime dtFechaFin = (DateTime) this.cmbHasta.Value;

			int iCuenta = 0;
			int idTipoCC = 0;
//			this.errorProvider.SetError(this.cmbCuentaCod,"");
//			if (cmbCuentaCod.Text.Trim().Length == 0)
//				{
//					errorProvider.SetError(cmbCuentaCod, "Ingrese Cuenta");
//					return;
//				}
			if (!chkCuenta.Checked) iCuenta = (int) this.cmbCuentaCod.Value;
			if (!chkTodoTipo.Checked) idTipoCC = (int) cmbCentroCosto.Value;

			#region presupuesto
			string stBuscaPres = "Select IsNull(sum(IsNull(Debe, 0)) - sum(IsNull(Haber, 0)),0) "
				+ " From PresupuestoDetalle inner join Presupuesto "
				+ " on Presupuesto.idpresupuesto = presupuestoDetalle.idPresupuesto "
				+ " Where (presupuestoDetalle.idcuenta = " + iCuenta.ToString()
				+ " or " +  iCuenta.ToString() + " = 0 )"
				+ " and (presupuesto.idproyecto = " + idTipoCC.ToString()
				+ " or " + idTipoCC.ToString() + "= 0 )";

			dPresupuesto = Funcion.diEscalarSQL(cdsCuenta,stBuscaPres);
			txtPresupuesto.Value = dPresupuesto;
			#endregion 
			#region saldo
			stBuscaPres = "Select isnull(sum(IsNull(Debe,0)) - sum(IsNull(Haber,0)),0) "
				+ "From AsientoDetalle inner join Asiento "
				+ " on Asiento.idAsiento = AsientoDetalle.idAsiento "
				+ " INNER JOIN CUENTA ON CUENTA.IDCUENTA = ASIENTODETALLE.IDCUENTA "
				+ " Where asiento.borrar = 0 and asiento.fecha < '" + dtFechaIni.ToString("yyyyMMdd")+ "'"
				+ " and (asientodetalle.idcuenta = " + iCuenta.ToString() 
				+ " or " +  iCuenta.ToString() + " = 0 )" 
				+ " AND  ISNULL(CUENTA.CodPresup,'') <> '' " 
				+ " and (asiento.idproyecto = " + idTipoCC.ToString() 
				+ " or " + idTipoCC.ToString() + " = 0 )";

			dSaldo = Funcion.diEscalarSQL(cdsCuenta,stBuscaPres);
			#endregion
			#region Ingresos - Egresos
			Funcion miFuncion = new Funcion();
			string strConn = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);
			string stProcedimiento = "exec FujoEfectivoEstimado"
				+ " '" + dtFechaIni.ToString("yyyyMMdd")
				+ "', '" + dtFechaFin.ToString("yyyyMMdd") + "', "
				+ iCuenta.ToString() + ","
				+ idTipoCC.ToString();
			SqlDataAdapter myData = new SqlDataAdapter(stProcedimiento, strConn);
			//			myData.TableMappings.Add("Tabla", "ExcelSube");
			miDataSetG.Reset();
			miDataSetG.Clear();
			try
			{
				myData.Fill(miDataSetG);
			}
			catch(Exception ex)
			{
				MessageBox.Show(ex.Message, "Error al abrir la Hoja A del Archivo: " );
				return;
			}
			dvG = miDataSetG.Tables[0].DefaultView;
			ultraGrid1.DataSource = dvG;
			#endregion
			#region Visibilidad
			ultraGrid1.DisplayLayout.Bands[0].Columns["idAsiento"].Hidden = true;
			#endregion Visibilidad
			#region Total
			//ingresos
			double dIngresos = 0;
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in ultraGrid1.Rows.All)
			{
				if (dr.Cells["Debe"].Value != System.DBNull.Value)
					dIngresos += (double) dr.Cells["Debe"].Value;
			}
			txtIngresos.Value = dIngresos;	
			//egreos
			double dEgresos = 0;
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in ultraGrid1.Rows.All)
			{
				if (dr.Cells["Haber"].Value != System.DBNull.Value)
					dEgresos += (double) dr.Cells["haber"].Value;
			}
			txtEgresos.Value = dEgresos;	
			dSaldo = dSaldo + dIngresos - dEgresos;
			txtSaldo.Value = dSaldo;
			dTotal = dSaldo - DPagos; 
			txtTotal.Value = dTotal;
			#endregion

		MessageBox.Show("Generado");

		}

		private void btAsiento_Click(object sender, System.EventArgs e)
		{
			if (ultraGrid1.ActiveRow == null)
			{
				MessageBox.Show("Seleccione Fila", "Información");
				return;
			}

			int IdCompraReq = (int) ultraGrid1.ActiveRow.Cells["idAsiento"].Value;
			if (IdCompraReq == 0)
			{
				MessageBox.Show("No existe Asiento", "Información");
				return;
			}
			//			int TipoFactura = 36;
			Cursor = Cursors.WaitCursor;
			Asiento miAsiento = new Asiento(IdCompraReq);
			miAsiento.Show();
			Cursor = Cursors.Default;
		}

		private void simpleButton2_Click(object sender, System.EventArgs e)
		{
			DateTime dtDesde = (DateTime) cmbFecha.Value;
			DateTime dtHasta = (DateTime) cmbHasta.Value;

			ParameterFields paramFields = new ParameterFields ();
			ParameterField paramField = new ParameterField ();
			ParameterDiscreteValue discreteVal = new ParameterDiscreteValue ();
			paramField.ParameterFieldName = "@FDesde";
			discreteVal.Value = dtDesde.ToString("yyyy-MM-dd");
			paramField.CurrentValues.Add (discreteVal);
			paramFields.Add (paramField);

			ParameterField paramField1 = new ParameterField ();
			ParameterDiscreteValue discreteVal1 = new ParameterDiscreteValue ();
			paramField1.ParameterFieldName = "@FHasta";
			discreteVal1.Value = dtHasta.ToString("yyyy-MM-dd");
			paramField1.CurrentValues.Add (discreteVal1);
			paramFields.Add (paramField1);

			ParameterField paramField2 = new ParameterField ();
			ParameterDiscreteValue discreteVal2 = new ParameterDiscreteValue ();
			paramField2.ParameterFieldName = "@idCuenta";
			if (chkCuenta.Checked)
				discreteVal2.Value = 0;
			else
				discreteVal2.Value = (int) cmbCuentaCod.Value;
			paramField2.CurrentValues.Add (discreteVal2);
			paramFields.Add (paramField2);

	    ParameterField paramField3 = new ParameterField ();
			ParameterDiscreteValue discreteVal3 = new ParameterDiscreteValue ();
			paramField3.ParameterFieldName = "@idCentroCosto";
			if (chkTodoTipo.Checked)
				discreteVal3.Value = 0;
			else
				discreteVal3.Value = (int) cmbCentroCosto.Value;
			paramField3.CurrentValues.Add (discreteVal3);
			paramFields.Add (paramField3);

			ParameterField paramField4 = new ParameterField ();
			ParameterDiscreteValue discreteVal4 = new ParameterDiscreteValue ();
			paramField4.ParameterFieldName = "@Presupuesto";
			discreteVal4.Value = dPresupuesto;
			paramField4.CurrentValues.Add (discreteVal4);
			paramFields.Add (paramField4);

			ParameterField paramField5 = new ParameterField ();
			ParameterDiscreteValue discreteVal5 = new ParameterDiscreteValue ();
			paramField5.ParameterFieldName = "@SaldoContable";
			discreteVal5.Value = dSaldo;
			paramField5.CurrentValues.Add (discreteVal5);
			paramFields.Add (paramField5);

			ParameterField paramField6 = new ParameterField ();
			ParameterDiscreteValue discreteVal6 = new ParameterDiscreteValue ();
			paramField6.ParameterFieldName = "@PagoAutorizados";
			discreteVal6.Value = DPagos;
			paramField6.CurrentValues.Add (discreteVal6);
			paramFields.Add (paramField6);

			ParameterField paramField7 = new ParameterField ();
			ParameterDiscreteValue discreteVal7 = new ParameterDiscreteValue ();
			paramField7.ParameterFieldName = "@Total";
			discreteVal7.Value = dTotal;
			paramField7.CurrentValues.Add (discreteVal7);
			paramFields.Add (paramField7);

			ParameterField paramField8 = new ParameterField ();
			ParameterDiscreteValue discreteVal8 = new ParameterDiscreteValue ();
			paramField8.ParameterFieldName = "@idPagoLote";
			discreteVal8.Value = (int) iPagoLote;
			paramField8.CurrentValues.Add (discreteVal8);
			paramFields.Add (paramField8);

			string stReporte = "FlujoEfectivoEstimado.rpt";
//			if (stReporte.Trim().Length == 0) stReporte = "FlujoEfectivoEstimado.rpt";
			Reporte miReporte = new Reporte(stReporte, "");
			miReporte.MdiParent = this.MdiParent;
			miReporte.crVista.ParameterFieldInfo = paramFields;
			miReporte.Show();					


		}

		private void groupBox1_Enter(object sender, System.EventArgs e)
		{
		
		}

		private void txtSaldo_ValueChanged(object sender, System.EventArgs e)
		{
		
		}

		private void chkTodoTipo_CheckedChanged(object sender, System.EventArgs e)
		{
			this.cmbCentroCosto.Enabled = !this.chkTodoTipo.Checked;

		}
		private void chkCuenta_CheckedChanged(object sender, System.EventArgs e)
		{
			this.cmbCuentaCod.Enabled = !this.chkCuenta.Checked;

		}

		private void btfactura_Click(object sender, System.EventArgs e)
		{
			if (ultraGrid2.ActiveRow == null)
			{
				MessageBox.Show("Seleccione Fila", "Información");
				return;
			}

			int IdFactura = (int) ultraGrid2.ActiveRow.Cells["IdCompra"].Value;
			if (IdFactura == 0)
			{
				MessageBox.Show("No existe Factura", "Información");
				return;
			}
			//			int TipoFactura = 36;
			Cursor = Cursors.WaitCursor;
//			Compra miCompra = new Compra(4,IdFactura);
//			miCompra.Show();
			Cursor = Cursors.Default;
		}
	}
}





