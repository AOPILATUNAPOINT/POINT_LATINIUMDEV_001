using System;
using System.Data.SqlClient;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de AsientoPeriodo.
	/// </summary>
	public class AsientoPeriodo : DevExpress.XtraEditors.XtraForm
	{
		private System.Windows.Forms.Label label1;
		private System.Windows.Forms.Label label2;
		private Infragistics.Win.UltraWinEditors.UltraOptionSet optTipo;
		private DevExpress.XtraEditors.SimpleButton btGenerar;
		private DevExpress.XtraEditors.SimpleButton btSalir;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor ultraCheckEditor1;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor ultraCheckEditor3;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbProyecto;
		private C1.Data.C1DataSet cdsAsiento;
		private C1.Data.C1DataSet cdsCompra;
		private C1.Data.C1DataSet cdsProyecto;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo cmbDesde;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo cmbHasta;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkTodoProyecto;
		private System.Windows.Forms.ErrorProvider errorProvider;
		private DevExpress.XtraEditors.SimpleButton btCosto;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkAsiRet;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbSubProyecto;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkSubProy;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkCostoMedio;
		private DevExpress.XtraEditors.SimpleButton btIndividual;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkSobrescribe;
		private System.Windows.Forms.GroupBox groupBox1;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public AsientoPeriodo()
		{
			InitializeComponent();
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
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton2 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.Appearance appearance1 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.ValueListItem valueListItem1 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem2 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem3 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem4 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem5 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem6 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem7 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem8 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem9 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem10 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Proyecto", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idProyecto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPadre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Lugar");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Responsable");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Notas");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CentroCosto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaIngreso");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Visible");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idAsiDistGrupo");
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("SubProyecto", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn11 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idSubProyecto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn12 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idProyecto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn13 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn14 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Lugar");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn15 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("RespNombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn16 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Notas");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn17 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CodSubproyecto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn18 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Visible");
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance11 = new Infragistics.Win.Appearance();
			this.cmbDesde = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.label1 = new System.Windows.Forms.Label();
			this.label2 = new System.Windows.Forms.Label();
			this.cmbHasta = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.optTipo = new Infragistics.Win.UltraWinEditors.UltraOptionSet();
			this.btGenerar = new DevExpress.XtraEditors.SimpleButton();
			this.btSalir = new DevExpress.XtraEditors.SimpleButton();
			this.ultraCheckEditor1 = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.chkAsiRet = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.ultraCheckEditor3 = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.chkTodoProyecto = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.cmbProyecto = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.cdsProyecto = new C1.Data.C1DataSet();
			this.cdsAsiento = new C1.Data.C1DataSet();
			this.cdsCompra = new C1.Data.C1DataSet();
			this.errorProvider = new System.Windows.Forms.ErrorProvider();
			this.btCosto = new DevExpress.XtraEditors.SimpleButton();
			this.cmbSubProyecto = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.chkSubProy = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.chkCostoMedio = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.btIndividual = new DevExpress.XtraEditors.SimpleButton();
			this.chkSobrescribe = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.groupBox1 = new System.Windows.Forms.GroupBox();
			((System.ComponentModel.ISupportInitialize)(this.cmbDesde)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbHasta)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.optTipo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbProyecto)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsProyecto)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsAsiento)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsCompra)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbSubProyecto)).BeginInit();
			this.groupBox1.SuspendLayout();
			this.SuspendLayout();
			// 
			// cmbDesde
			// 
			this.cmbDesde.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton1.Caption = "Today";
			this.cmbDesde.DateButtons.Add(dateButton1);
			this.cmbDesde.Format = "dd/MMM/yyyy";
			this.cmbDesde.Location = new System.Drawing.Point(64, 26);
			this.cmbDesde.Name = "cmbDesde";
			this.cmbDesde.NonAutoSizeHeight = 23;
			this.cmbDesde.Size = new System.Drawing.Size(120, 21);
			this.cmbDesde.SpinButtonsVisible = true;
			this.cmbDesde.TabIndex = 0;
			this.cmbDesde.Value = new System.DateTime(2009, 10, 12, 0, 0, 0, 0);
			// 
			// label1
			// 
			this.label1.AutoSize = true;
			this.label1.Location = new System.Drawing.Point(16, 26);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(39, 16);
			this.label1.TabIndex = 1;
			this.label1.Text = "Desde:";
			// 
			// label2
			// 
			this.label2.AutoSize = true;
			this.label2.Location = new System.Drawing.Point(216, 26);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(37, 16);
			this.label2.TabIndex = 3;
			this.label2.Text = "Hasta:";
			// 
			// cmbHasta
			// 
			this.cmbHasta.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton2.Caption = "Today";
			this.cmbHasta.DateButtons.Add(dateButton2);
			this.cmbHasta.Format = "dd/MMM/yyyy";
			this.cmbHasta.Location = new System.Drawing.Point(264, 26);
			this.cmbHasta.Name = "cmbHasta";
			this.cmbHasta.NonAutoSizeHeight = 23;
			this.cmbHasta.Size = new System.Drawing.Size(120, 21);
			this.cmbHasta.SpinButtonsVisible = true;
			this.cmbHasta.TabIndex = 2;
			this.cmbHasta.Value = new System.DateTime(2009, 10, 12, 0, 0, 0, 0);
			// 
			// optTipo
			// 
			appearance1.ForeColorDisabled = System.Drawing.Color.Black;
			this.optTipo.Appearance = appearance1;
			this.optTipo.CheckedIndex = 0;
			this.optTipo.ItemAppearance = appearance2;
			this.optTipo.ItemOrigin = new System.Drawing.Point(10, 6);
			valueListItem1.DataValue = 0;
			valueListItem1.DisplayText = "Ventas";
			valueListItem2.DataValue = 1;
			valueListItem2.DisplayText = "Compras";
			valueListItem3.DataValue = 2;
			valueListItem3.DisplayText = "Dev. Venta";
			valueListItem4.DataValue = 3;
			valueListItem4.DisplayText = "Dev. Compra";
			valueListItem5.DataValue = 4;
			valueListItem5.DisplayText = "Reservación";
			valueListItem6.DataValue = 5;
			valueListItem6.DisplayText = "Egreso de Bodega";
			valueListItem7.DataValue = 6;
			valueListItem7.DisplayText = "Ingreso de Bodega";
			valueListItem8.DataValue = 7;
			valueListItem8.DisplayText = "Pagos Recibidos";
			valueListItem9.DataValue = 8;
			valueListItem9.DisplayText = "Pagos Realizados";
			valueListItem10.DataValue = 9;
			valueListItem10.DisplayText = "Consumo Real";
			this.optTipo.Items.Add(valueListItem1);
			this.optTipo.Items.Add(valueListItem2);
			this.optTipo.Items.Add(valueListItem3);
			this.optTipo.Items.Add(valueListItem4);
			this.optTipo.Items.Add(valueListItem5);
			this.optTipo.Items.Add(valueListItem6);
			this.optTipo.Items.Add(valueListItem7);
			this.optTipo.Items.Add(valueListItem8);
			this.optTipo.Items.Add(valueListItem9);
			this.optTipo.Items.Add(valueListItem10);
			this.optTipo.ItemSpacingVertical = 6;
			this.optTipo.Location = new System.Drawing.Point(24, 69);
			this.optTipo.Name = "optTipo";
			this.optTipo.Size = new System.Drawing.Size(128, 233);
			this.optTipo.TabIndex = 4;
			this.optTipo.Text = "Ventas";
			// 
			// btGenerar
			// 
			this.btGenerar.Location = new System.Drawing.Point(40, 310);
			this.btGenerar.Name = "btGenerar";
			this.btGenerar.Size = new System.Drawing.Size(75, 25);
			this.btGenerar.TabIndex = 5;
			this.btGenerar.Text = "&Generar";
			this.btGenerar.Click += new System.EventHandler(this.btGenerar_Click);
			// 
			// btSalir
			// 
			this.btSalir.Location = new System.Drawing.Point(288, 310);
			this.btSalir.Name = "btSalir";
			this.btSalir.Size = new System.Drawing.Size(75, 25);
			this.btSalir.TabIndex = 6;
			this.btSalir.Text = "&Salir";
			this.btSalir.Click += new System.EventHandler(this.btSalir_Click);
			// 
			// ultraCheckEditor1
			// 
			appearance3.ForeColorDisabled = System.Drawing.Color.Black;
			this.ultraCheckEditor1.Appearance = appearance3;
			this.ultraCheckEditor1.Location = new System.Drawing.Point(288, 52);
			this.ultraCheckEditor1.Name = "ultraCheckEditor1";
			this.ultraCheckEditor1.Size = new System.Drawing.Size(120, 21);
			this.ultraCheckEditor1.TabIndex = 7;
			this.ultraCheckEditor1.Text = "Detalle en Mayor";
			this.ultraCheckEditor1.Visible = false;
			// 
			// chkAsiRet
			// 
			appearance4.ForeColorDisabled = System.Drawing.Color.Black;
			this.chkAsiRet.Appearance = appearance4;
			this.chkAsiRet.Location = new System.Drawing.Point(176, 69);
			this.chkAsiRet.Name = "chkAsiRet";
			this.chkAsiRet.Size = new System.Drawing.Size(136, 21);
			this.chkAsiRet.TabIndex = 8;
			this.chkAsiRet.Text = "Asiento de Retención";
			// 
			// ultraCheckEditor3
			// 
			appearance5.ForeColorDisabled = System.Drawing.Color.Black;
			this.ultraCheckEditor3.Appearance = appearance5;
			this.ultraCheckEditor3.Location = new System.Drawing.Point(208, 0);
			this.ultraCheckEditor3.Name = "ultraCheckEditor3";
			this.ultraCheckEditor3.Size = new System.Drawing.Size(168, 22);
			this.ultraCheckEditor3.TabIndex = 9;
			this.ultraCheckEditor3.Text = "Asiento por Forma de Pago";
			this.ultraCheckEditor3.Visible = false;
			// 
			// chkTodoProyecto
			// 
			appearance6.ForeColorDisabled = System.Drawing.Color.Black;
			this.chkTodoProyecto.Appearance = appearance6;
			this.chkTodoProyecto.Checked = true;
			this.chkTodoProyecto.CheckState = System.Windows.Forms.CheckState.Checked;
			this.chkTodoProyecto.Location = new System.Drawing.Point(176, 103);
			this.chkTodoProyecto.Name = "chkTodoProyecto";
			this.chkTodoProyecto.Size = new System.Drawing.Size(120, 22);
			this.chkTodoProyecto.TabIndex = 10;
			this.chkTodoProyecto.Text = "Todo Centro Costo";
			this.chkTodoProyecto.CheckedChanged += new System.EventHandler(this.chkTodoProyecto_CheckedChanged);
			// 
			// cmbProyecto
			// 
			appearance7.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbProyecto.Appearance = appearance7;
			this.cmbProyecto.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbProyecto.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbProyecto.DataMember = "Proyecto";
			this.cmbProyecto.DataSource = this.cdsProyecto;
			ultraGridColumn1.Header.VisiblePosition = 2;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn2.Header.VisiblePosition = 6;
			ultraGridColumn2.Hidden = true;
			ultraGridColumn3.Header.VisiblePosition = 0;
			ultraGridColumn3.Width = 301;
			ultraGridColumn4.Header.VisiblePosition = 3;
			ultraGridColumn4.Hidden = true;
			ultraGridColumn5.Header.VisiblePosition = 7;
			ultraGridColumn5.Hidden = true;
			ultraGridColumn6.Header.VisiblePosition = 4;
			ultraGridColumn6.Hidden = true;
			ultraGridColumn7.Header.Caption = "Codigo";
			ultraGridColumn7.Header.VisiblePosition = 1;
			ultraGridColumn7.Width = 112;
			ultraGridColumn8.Header.VisiblePosition = 5;
			ultraGridColumn8.Hidden = true;
			ultraGridColumn9.Header.VisiblePosition = 8;
			ultraGridColumn9.Hidden = true;
			ultraGridColumn10.Header.VisiblePosition = 9;
			ultraGridColumn10.Hidden = true;
			ultraGridColumn10.Width = 62;
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
																										 ultraGridColumn10});
			this.cmbProyecto.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.cmbProyecto.DisplayMember = "Nombre";
			this.cmbProyecto.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbProyecto.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbProyecto.Enabled = false;
			this.cmbProyecto.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbProyecto.Location = new System.Drawing.Point(200, 129);
			this.cmbProyecto.Name = "cmbProyecto";
			this.cmbProyecto.Size = new System.Drawing.Size(200, 21);
			this.cmbProyecto.TabIndex = 11;
			this.cmbProyecto.ValueMember = "idProyecto";
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
			this.cdsProyecto.BeforeFill += new C1.Data.FillEventHandler(this.cdsProyecto_BeforeFill);
			// 
			// cdsAsiento
			// 
			this.cdsAsiento.BindingContextCtrl = this;
			this.cdsAsiento.DataLibrary = "LibContabilidad";
			this.cdsAsiento.DataLibraryUrl = "";
			this.cdsAsiento.DataSetDef = "dsAsiento";
			this.cdsAsiento.FillOnRequest = false;
			this.cdsAsiento.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsAsiento.Name = "cdsAsiento";
			this.cdsAsiento.SchemaClassName = "LibContabilidad.DataClass";
			this.cdsAsiento.SchemaDef = null;
			this.cdsAsiento.BeforeFill += new C1.Data.FillEventHandler(this.cdsAsiento_BeforeFill);
			// 
			// cdsCompra
			// 
			this.cdsCompra.BindingContextCtrl = this;
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
			// errorProvider
			// 
			this.errorProvider.ContainerControl = this;
			this.errorProvider.DataMember = "";
			// 
			// btCosto
			// 
			this.btCosto.Location = new System.Drawing.Point(168, 310);
			this.btCosto.Name = "btCosto";
			this.btCosto.Size = new System.Drawing.Size(75, 25);
			this.btCosto.TabIndex = 12;
			this.btCosto.Text = "&Costo";
			this.btCosto.Click += new System.EventHandler(this.btCosto_Click);
			// 
			// cmbSubProyecto
			// 
			appearance8.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbSubProyecto.Appearance = appearance8;
			this.cmbSubProyecto.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbSubProyecto.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbSubProyecto.DataMember = "SubProyecto";
			this.cmbSubProyecto.DataSource = this.cdsProyecto;
			this.cmbSubProyecto.DisplayLayout.AutoFitColumns = true;
			ultraGridBand2.AddButtonCaption = "Proyecto";
			ultraGridColumn11.Header.VisiblePosition = 3;
			ultraGridColumn11.Hidden = true;
			ultraGridColumn11.Width = 107;
			ultraGridColumn12.Header.VisiblePosition = 2;
			ultraGridColumn12.Hidden = true;
			ultraGridColumn13.Header.VisiblePosition = 0;
			ultraGridColumn13.Width = 117;
			ultraGridColumn14.Header.VisiblePosition = 4;
			ultraGridColumn14.Hidden = true;
			ultraGridColumn15.Header.VisiblePosition = 6;
			ultraGridColumn15.Hidden = true;
			ultraGridColumn15.Width = 126;
			ultraGridColumn16.Header.VisiblePosition = 5;
			ultraGridColumn16.Hidden = true;
			ultraGridColumn17.Header.Caption = "Codigo";
			ultraGridColumn17.Header.VisiblePosition = 1;
			ultraGridColumn17.Width = 64;
			ultraGridColumn18.Header.VisiblePosition = 7;
			ultraGridColumn18.Hidden = true;
			ultraGridColumn18.Width = 41;
			ultraGridBand2.Columns.AddRange(new object[] {
																										 ultraGridColumn11,
																										 ultraGridColumn12,
																										 ultraGridColumn13,
																										 ultraGridColumn14,
																										 ultraGridColumn15,
																										 ultraGridColumn16,
																										 ultraGridColumn17,
																										 ultraGridColumn18});
			this.cmbSubProyecto.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			this.cmbSubProyecto.DisplayMember = "Nombre";
			this.cmbSubProyecto.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbSubProyecto.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbSubProyecto.Enabled = false;
			this.cmbSubProyecto.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbSubProyecto.Location = new System.Drawing.Point(200, 181);
			this.cmbSubProyecto.Name = "cmbSubProyecto";
			this.cmbSubProyecto.Size = new System.Drawing.Size(200, 21);
			this.cmbSubProyecto.TabIndex = 14;
			this.cmbSubProyecto.ValueMember = "idSubProyecto";
			// 
			// chkSubProy
			// 
			appearance9.ForeColorDisabled = System.Drawing.Color.Black;
			this.chkSubProy.Appearance = appearance9;
			this.chkSubProy.Checked = true;
			this.chkSubProy.CheckState = System.Windows.Forms.CheckState.Checked;
			this.chkSubProy.Location = new System.Drawing.Point(176, 155);
			this.chkSubProy.Name = "chkSubProy";
			this.chkSubProy.Size = new System.Drawing.Size(120, 22);
			this.chkSubProy.TabIndex = 13;
			this.chkSubProy.Text = "Todo Proyecto";
			this.chkSubProy.CheckedChanged += new System.EventHandler(this.chkSubProy_CheckedChanged);
			// 
			// chkCostoMedio
			// 
			appearance10.ForeColorDisabled = System.Drawing.Color.Black;
			this.chkCostoMedio.Appearance = appearance10;
			this.chkCostoMedio.Location = new System.Drawing.Point(176, 215);
			this.chkCostoMedio.Name = "chkCostoMedio";
			this.chkCostoMedio.Size = new System.Drawing.Size(88, 22);
			this.chkCostoMedio.TabIndex = 15;
			this.chkCostoMedio.Text = "Costo Medio";
			// 
			// btIndividual
			// 
			this.btIndividual.Location = new System.Drawing.Point(152, 17);
			this.btIndividual.Name = "btIndividual";
			this.btIndividual.Size = new System.Drawing.Size(75, 25);
			this.btIndividual.TabIndex = 16;
			this.btIndividual.Text = "&Individual";
			this.btIndividual.Click += new System.EventHandler(this.btIndividual_Click);
			// 
			// chkSobrescribe
			// 
			appearance11.ForeColorDisabled = System.Drawing.Color.Black;
			this.chkSobrescribe.Appearance = appearance11;
			this.chkSobrescribe.Location = new System.Drawing.Point(8, 17);
			this.chkSobrescribe.Name = "chkSobrescribe";
			this.chkSobrescribe.Size = new System.Drawing.Size(144, 22);
			this.chkSobrescribe.TabIndex = 17;
			this.chkSobrescribe.Text = "Sobreecribir Existentes";
			// 
			// groupBox1
			// 
			this.groupBox1.Controls.Add(this.btIndividual);
			this.groupBox1.Controls.Add(this.chkSobrescribe);
			this.groupBox1.Location = new System.Drawing.Point(168, 241);
			this.groupBox1.Name = "groupBox1";
			this.groupBox1.Size = new System.Drawing.Size(240, 52);
			this.groupBox1.TabIndex = 18;
			this.groupBox1.TabStop = false;
			this.groupBox1.Text = "Individual";
			// 
			// AsientoPeriodo
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 14);
			this.ClientSize = new System.Drawing.Size(424, 359);
			this.Controls.Add(this.groupBox1);
			this.Controls.Add(this.chkCostoMedio);
			this.Controls.Add(this.cmbSubProyecto);
			this.Controls.Add(this.chkSubProy);
			this.Controls.Add(this.btCosto);
			this.Controls.Add(this.cmbProyecto);
			this.Controls.Add(this.chkTodoProyecto);
			this.Controls.Add(this.ultraCheckEditor3);
			this.Controls.Add(this.chkAsiRet);
			this.Controls.Add(this.ultraCheckEditor1);
			this.Controls.Add(this.btSalir);
			this.Controls.Add(this.btGenerar);
			this.Controls.Add(this.optTipo);
			this.Controls.Add(this.label2);
			this.Controls.Add(this.label1);
			this.Controls.Add(this.cmbHasta);
			this.Controls.Add(this.cmbDesde);
			this.MaximizeBox = false;
			this.Name = "AsientoPeriodo";
			this.Text = "Asientos por Período";
			this.Closing += new System.ComponentModel.CancelEventHandler(this.AsientoPeriodo_Closing);
			this.Load += new System.EventHandler(this.AsientoPeriodo_Load);
			((System.ComponentModel.ISupportInitialize)(this.cmbDesde)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbHasta)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.optTipo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbProyecto)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsProyecto)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsAsiento)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsCompra)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbSubProyecto)).EndInit();
			this.groupBox1.ResumeLayout(false);
			this.ResumeLayout(false);

		}
		#endregion

		private string TituloAsiento()
		{
			return "";

		}

		private bool Verifica()
		{
			bool bOk = true;
			errorProvider.SetError(cmbSubProyecto, "");
			errorProvider.SetError(cmbProyecto, "");
			if (!chkTodoProyecto.Checked)
				if (cmbProyecto.Text.Trim().Length == 0)
				{
					errorProvider.SetError(cmbProyecto, "Ingrese Proyecto");
					bOk = false;
				}
			if (!chkSubProy.Checked)
				if (cmbSubProyecto.Text.Trim().Length == 0)
				{
					errorProvider.SetError(cmbSubProyecto, "Ingrese SubProyecto");
					bOk = false;
				}
			return bOk;
		}

		private void btGenerar_Click(object sender, System.EventArgs e)
		{
			if (!Verifica()) return;
			DateTime dtFechaIni = (DateTime) this.cmbDesde.Value;
			DateTime dtFechaFin = (DateTime) this.cmbHasta.Value;
			bool bFechaIgual = true;
			if (dtFechaIni != dtFechaFin) bFechaIgual =  false;
			string stDescripcion = this.optTipo.Text.ToString().ToUpper();
			int iTipoF = (int) optTipo.Value;
			
			if (chkAsiRet.Checked) 
			{
				if (iTipoF == 0)
					stDescripcion = "RETENCIONES DE VENTAS";
				else if (iTipoF == 1)
					stDescripcion = "RETENCIONES DE COMPRAS";
				else
				{
					MessageBox.Show("Retenciones solo permitidas en Compras y Ventas",
						"Información", MessageBoxButtons.OK, MessageBoxIcon.Information);
					return;
				}
			}
			if (bFechaIgual)
				stDescripcion = "RESUMEN DE " + stDescripcion + " DEL " + dtFechaIni.ToString("dd/MMM/yyyy");
			else
				stDescripcion = "RESUMEN DE " + stDescripcion + " DESDE EL " + 
					dtFechaIni.ToString("dd/MMM/yyyy") + 
					" HASTA EL " + dtFechaFin.ToString("dd/MMM/yyyy");

			if (DialogResult.No == MessageBox.Show("Desea Generar un " 
				+ stDescripcion + "?", "Confirmación", 
				MessageBoxButtons.YesNo, MessageBoxIcon.Question)) return;

			if (chkAsiRet.Checked) 
			{
				int iTipoRetRes = 1;
				if (iTipoF == 1) iTipoRetRes = 4;
				string stExec = "EXEC AsientoRetencionLote " + iTipoRetRes.ToString() 
					+ ", '" + dtFechaIni.ToString("yyyyMMdd")
					+ "', '" + dtFechaFin.ToString("yyyyMMdd") + "'";
				Funcion.EjecutaSQL(cdsAsiento, stExec, true);

				MessageBox.Show("Asiento de Retencion Creado", "Información");
				return;
			}
			if (iTipoF == 7 || iTipoF == 8)
			{
				if (iTipoF == 7) iTipoF = 1;
				else iTipoF = 4;
				string stExec = "EXEC PagoAsientoLote " + iTipoF.ToString()
					+ ", '" + dtFechaIni.ToString("yyyyMMdd")
					+ "', '" + dtFechaFin.ToString("yyyyMMdd") + "'";
				Funcion.EjecutaSQL(cdsAsiento, stExec, true);

				MessageBox.Show("Asiento de Pagos Creado", "Información");
				return;
			}

			dtFechaFin = dtFechaFin.AddDays(1);
			//if (iTipoF == 2) iTipoF = 4;
			string stProc = "CreaAsientoPeriodo";
			switch(iTipoF)
			{
				case 0: // ventas
					iTipoF = 1;
					break;
				case 1: // Compras
					stProc = "CreaAsientoPerComp";
					iTipoF = 4;
					break;
				case 2: // Devolucion en Ventas
					stProc = "CreaAsientoPerDev";
					iTipoF = 5;
					break;
				case 3: // Devolucion en Compras
					stProc = "CreaAsientoPerDevComp";
					iTipoF = 6;
					break;
				case 4: // Reservacion
					stProc = "CreaAsiPerReserva";
					iTipoF = 11;
					break;
				case 5: // Envio de Transferencia
					stProc = "CreaAsientoPerTrans";
					iTipoF = 8;
					break;
				case 6: // Recepcion de Transferencia
					stProc = "CreaAsientoPerTransRec";
					iTipoF = 9;
					break;
				case 9: // Consumo Real
					stProc = "CreaAsientoPerTrans";
					iTipoF = 10;
					break;
			}


			string stVerifica = String.Format("Exec AsientoVerificaExiste {0}, '{1}', '{2}'",
				iTipoF, dtFechaIni.ToString("yyyyMMdd"), dtFechaFin.ToString("yyyyMMdd"));
			if (!chkTodoProyecto.Checked)
				stVerifica += ", " + cmbProyecto.Value.ToString();
			else stVerifica += ", 0";

			if (!chkSubProy.Checked)
				stVerifica += ", " + cmbSubProyecto.Value.ToString();
			else stVerifica += ", 0";
			
			int iExisteAsi = Funcion.iEscalarSQL(cdsAsiento, stVerifica, true);
			if (iExisteAsi > 0)
			{
				stVerifica = stVerifica.Replace("AsientoVerificaExiste", "AsientoVerificaExisteLista");
				FrmFacturasAsientos miFac = new FrmFacturasAsientos(stVerifica);
				miFac.ShowDialog();

				MessageBox.Show("Generacion Cancelada. Existen Asientos en el periodo: " + iExisteAsi.ToString(), "Información");
				return;
			}
		// Este asiento solo funciona por el momento para las ventas
		// la retencion no se genera
		string stRetencion = ", 0";
			if (chkAsiRet.Checked) stRetencion = ", 1";
			string stSelect = "EXEC " + stProc + " '" 
				+ dtFechaIni.ToString("yyyyMMdd") + "', '" 
				+ dtFechaFin.AddDays(-1).ToString("yyyyMMdd") + "', " + iTipoF.ToString() 
				+ ", '" + stDescripcion + "'" + stRetencion;

			if (!chkTodoProyecto.Checked)
				stSelect += ", " + cmbProyecto.Value.ToString();
			else stSelect += ", 0";

			if (!chkSubProy.Checked)
				stSelect += ", " + cmbSubProyecto.Value.ToString();
			else stSelect += ", 0";

			Funcion.EjecutaSQL(cdsAsiento, stSelect);
			MessageBox.Show("Asiento Creado", "Información");
		}

		private void chkTodoProyecto_CheckedChanged(object sender, System.EventArgs e)
		{
			this.cmbProyecto.Enabled = !this.chkTodoProyecto.Checked;
		}

		private void btSalir_Click(object sender, System.EventArgs e)
		{
			this.Close();
		}

		private void cdsAsiento_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			this.cdsAsiento.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirConta);
		}

		private void cdsCompra_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			this.cdsCompra.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);		
		}

		private void cdsProyecto_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsProyecto.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);
		}

		private void btCosto_Click(object sender, System.EventArgs e)
		{
			int num1 = (int) this.optTipo.Value;
			switch (num1)
			{
				case 0:
					num1 = 1;
					break;

				case 2:
					num1 = 5;
					break;

				case 5:
					num1 = 8;
					break;

				case 6:
					num1 = 9;
					break;
				case 9:
					num1 = 10;
					break;

				default:
					MessageBox.Show("Costo de Venta no aplicable", "Información");
					return;
			}
			DateTime time1 = (DateTime) this.cmbDesde.Value;
			DateTime time2 = (DateTime) this.cmbHasta.Value;
			bool flag1 = true;
			if (time1 != time2)
			{
				flag1 = false;
			}
			string stCostoMedio = " ";
			if (chkCostoMedio.Checked) stCostoMedio = " MEDIO ";
			string text1 = this.optTipo.Text.ToString().ToUpper();
			if (flag1)
			{
				text1 = "COSTO" + stCostoMedio + "DE " + text1 + " DEL " + time1.ToString("dd/MMM/yyyy");
			}
			else
			{
				text1 = "COSTO" + stCostoMedio + "DE " + text1 + " DESDE EL " + time1.ToString("dd/MMM/yyyy") + " HASTA EL " + time2.ToString("dd/MMM/yyyy");
			}

			if (DialogResult.No != MessageBox.Show("¿Desea Generar el " + text1 + "?", "Confirmación", MessageBoxButtons.YesNo, MessageBoxIcon.Question))
			{
				stCostoMedio = "";
				if (chkCostoMedio.Checked) stCostoMedio = ", 1";

				string stProyecto = "0";
				string stSubProy = "0";
				if (!chkTodoProyecto.Checked) stProyecto = cmbProyecto.Value.ToString().Trim();
				if (!chkSubProy.Checked) stSubProy = cmbSubProyecto.Value.ToString().Trim();
				string text2 = "EXEC CreaCostoPeriodo '" + time1.ToString("yyyyMMdd") + "', '" 
					+ time2.AddDays(1).ToString("yyyyMMdd") + "', " + num1.ToString() + ", '" 
					+ text1 + "', " + stProyecto + ", " + stSubProy + stCostoMedio;
				Funcion.EjecutaSQL(this.cdsAsiento, text2, true);
				MessageBox.Show("Asiento Creado", "Información");
			}
		}

		private void AsientoPeriodo_Load(object sender, System.EventArgs e)
		{
			Cursor = Cursors.Default;
			cmbDesde.Value = DateTime.Today;
			cmbHasta.Value = DateTime.Today;
			string stAudita = "Exec AuditaCrear 53, 6, 'Asi. Periodo'";
			Funcion.EjecutaSQL(cdsAsiento, stAudita, true);
		}

		private void chkSubProy_CheckedChanged(object sender, System.EventArgs e)
		{
			this.cmbSubProyecto.Enabled = !this.chkSubProy.Checked;
			if (!chkSubProy.Checked) cmbSubProyecto.Select();
		}

		private void AsientoPeriodo_Closing(object sender, System.ComponentModel.CancelEventArgs e)
		{
			string stAudita = "Exec AuditaCrear 53, 7, 'Asi. Periodo'";
			Funcion.EjecutaSQL(cdsAsiento, stAudita, true);
		}

		private void btIndividual_Click(object sender, System.EventArgs e)
		{
			DateTime dtDesde = (DateTime) cmbDesde.Value;
			DateTime dtHasta = (DateTime) cmbHasta.Value;
			int iTipoF = (int) optTipo.Value;
			string stTitulo = "Desea generar asientos individuales de ";
			string stProc = "AsientosIndividuales";
			switch(iTipoF)
			{
				case 0: // ventas
					iTipoF = 1;
					stTitulo += "Ventas";
					break;
				case 1: // Compras
					iTipoF = 4;
					stTitulo += "Compras";
					break;
				case 2: // Devolucion en Ventas
					stTitulo += "Devolucion en Ventas";
					iTipoF = 5;
					break;
				case 3: // Devolucion en Compras
					stTitulo += "Devolucion en Compras";
					iTipoF = 6;
					break;
				case 4: // Reservacion
					iTipoF = 11;
					stTitulo += "Devolucion en Compras";
					break;
				case 5: // Egreso de Bodega
					stTitulo += "Egreso de Bodega";
					iTipoF = 8;
					break;
				case 6: // Ingreso de Bodega
					stTitulo += "Ingreso de Bodega";
					iTipoF = 9;
					break;
				case 7: // Pagos Recibidos
					stTitulo += "Pagos Recibidos";
          stProc = "PagosIndividuales";
					iTipoF = 1;
					break;
				case 8: // Pagos Realizados
					stTitulo += "Pagos Realizados";
					stProc = "PagosIndividuales";
					iTipoF = 4;
					break;
				case 9: // Consumo Real
					stTitulo += "Consumo Real";
					iTipoF = 10;
					break;
			}

			int iSobrescribe = 0;
			if (chkSobrescribe.Checked) iSobrescribe = 1;
if (DialogResult.No == MessageBox.Show(stTitulo, "Confirmación",
	MessageBoxButtons.YesNo, MessageBoxIcon.Question)) return;

			Cursor = Cursors.WaitCursor;
			string stExec = string.Format("Exec {0} {1}, '{2}', '{3}', {4}", stProc,
				iTipoF, dtDesde.ToString("yyyyMMdd"), dtHasta.ToString("yyyyMMdd"), iSobrescribe);
			Funcion.EjecutaSQL(cdsAsiento, stExec, true);
			Cursor = Cursors.Default;
			MessageBox.Show("Asientos individuales generados", "Información");
		}
	}
}
