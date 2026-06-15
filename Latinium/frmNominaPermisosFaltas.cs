using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de frmNominaPermisosFaltas.
	/// </summary>
	public class frmNominaPermisosFaltas : System.Windows.Forms.Form
	{
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbVendedor;
		private System.Windows.Forms.GroupBox groupBox2;
		private System.Windows.Forms.Label label7;
		private System.Windows.Forms.Button btnCancelar;
		private System.Windows.Forms.Button btnGuardar;
		private System.Windows.Forms.GroupBox groupBox1;
		private System.Windows.Forms.Label label4;
		private System.Windows.Forms.Label label2;
		private System.Windows.Forms.Label lblBodega;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtpHasta;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtpDesde;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbTipo;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource1;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbParto;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource2;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtIdNominaSalidas;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtMotivo;
		private C1.Data.C1DataSet cdsSeteoF;
		private System.Windows.Forms.Label lblHasta;
		private System.Windows.Forms.Label lblDesde;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public frmNominaPermisosFaltas()
		{
			//
			// Necesario para admitir el Diseñador de Windows Forms
			//
			InitializeComponent();

			//
			// TODO: agregar código de constructor después de llamar a InitializeComponent
			//
		}

		int idVendedor = 0;
		public int idCom_AsignacionDeVendedores = 0;
		int iBodega = 0;
		string sPeriodo = "";
		int bPermisosFaltas = 0;
		DateTime dtDesde = DateTime.Today;
		DateTime dtHasta = DateTime.Today;
		DateTime dtUltimoDiaMes = DateTime.Today;
		int idCargo = 0;
		int idCom_Rango = 0;
		int idCom_Estado = 0;

		public frmNominaPermisosFaltas(int IdVendedor, int IdCom_AsignacionDeVendedores, int IBodega, string SPeriodo, int BPermisosFaltas, DateTime DTDesde, DateTime DTHasta, DateTime DTUltimoDiaMes, int IdCargo, int IdCom_Rango, int IdCom_Estado)
		{
			//
			// Necesario para admitir el Diseñador de Windows Forms
			//
			InitializeComponent();
			idVendedor = IdVendedor;
			idCom_AsignacionDeVendedores = IdCom_AsignacionDeVendedores;
			iBodega = IBodega;
			sPeriodo = SPeriodo;
			bPermisosFaltas = BPermisosFaltas;
			dtDesde = DTDesde;
			dtHasta = DTHasta;
			dtUltimoDiaMes = DTUltimoDiaMes;
			idCargo = IdCargo;
			idCom_Rango = IdCom_Rango;
			idCom_Estado = IdCom_Estado;
			//
			// TODO: agregar código de constructor después de llamar a InitializeComponent
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

		#region Código generado por el Diseñador de Windows Forms
		/// <summary>
		/// Método necesario para admitir el Diseñador. No se puede modificar
		/// el contenido del método con el editor de código.
		/// </summary>
		private void InitializeComponent()
		{
			Infragistics.Win.Appearance appearance1 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPersonal");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmNominaPermisosFaltas));
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton1 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton2 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idNominaTiposDeSalidas");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Descripcion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idNominaTiposDeSalidas");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Descripcion");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand3 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idNominaTiposDeParto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Dias");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Descripcion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn3 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idNominaTiposDeParto");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn4 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Dias");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn5 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Descripcion");
			this.cmbVendedor = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.groupBox2 = new System.Windows.Forms.GroupBox();
			this.label7 = new System.Windows.Forms.Label();
			this.btnCancelar = new System.Windows.Forms.Button();
			this.btnGuardar = new System.Windows.Forms.Button();
			this.groupBox1 = new System.Windows.Forms.GroupBox();
			this.txtIdNominaSalidas = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label4 = new System.Windows.Forms.Label();
			this.label2 = new System.Windows.Forms.Label();
			this.lblBodega = new System.Windows.Forms.Label();
			this.dtpHasta = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.lblHasta = new System.Windows.Forms.Label();
			this.dtpDesde = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.lblDesde = new System.Windows.Forms.Label();
			this.txtMotivo = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.cmbTipo = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.ultraDataSource1 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.cmbParto = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.ultraDataSource2 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			((System.ComponentModel.ISupportInitialize)(this.cmbVendedor)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdNominaSalidas)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtpHasta)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtpDesde)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtMotivo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbTipo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbParto)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			this.SuspendLayout();
			// 
			// cmbVendedor
			// 
			appearance1.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbVendedor.Appearance = appearance1;
			this.cmbVendedor.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbVendedor.Cursor = System.Windows.Forms.Cursors.Default;
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridColumn2.Width = 335;
			ultraGridColumn3.Header.VisiblePosition = 2;
			ultraGridBand1.Columns.AddRange(new object[] {
																										 ultraGridColumn1,
																										 ultraGridColumn2,
																										 ultraGridColumn3});
			this.cmbVendedor.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.cmbVendedor.DisplayMember = "Nombre";
			this.cmbVendedor.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbVendedor.Enabled = false;
			this.cmbVendedor.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbVendedor.Location = new System.Drawing.Point(96, 8);
			this.cmbVendedor.Name = "cmbVendedor";
			this.cmbVendedor.Size = new System.Drawing.Size(360, 21);
			this.cmbVendedor.TabIndex = 0;
			this.cmbVendedor.ValueMember = "idPersonal";
			// 
			// groupBox2
			// 
			this.groupBox2.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.groupBox2.Location = new System.Drawing.Point(0, 40);
			this.groupBox2.Name = "groupBox2";
			this.groupBox2.Size = new System.Drawing.Size(464, 8);
			this.groupBox2.TabIndex = 314;
			this.groupBox2.TabStop = false;
			// 
			// label7
			// 
			this.label7.AutoSize = true;
			this.label7.Location = new System.Drawing.Point(8, 10);
			this.label7.Name = "label7";
			this.label7.Size = new System.Drawing.Size(44, 16);
			this.label7.TabIndex = 313;
			this.label7.Text = "Nombre";
			// 
			// btnCancelar
			// 
			this.btnCancelar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnCancelar.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(230)), ((System.Byte)(230)), ((System.Byte)(230)));
			this.btnCancelar.CausesValidation = false;
			this.btnCancelar.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.btnCancelar.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(50)), ((System.Byte)(50)), ((System.Byte)(50)));
			this.btnCancelar.Image = ((System.Drawing.Image)(resources.GetObject("btnCancelar.Image")));
			this.btnCancelar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnCancelar.Location = new System.Drawing.Point(88, 250);
			this.btnCancelar.Name = "btnCancelar";
			this.btnCancelar.Size = new System.Drawing.Size(72, 23);
			this.btnCancelar.TabIndex = 7;
			this.btnCancelar.Text = "Cancelar";
			this.btnCancelar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnCancelar.Click += new System.EventHandler(this.btnCancelar_Click);
			// 
			// btnGuardar
			// 
			this.btnGuardar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnGuardar.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(230)), ((System.Byte)(230)), ((System.Byte)(230)));
			this.btnGuardar.CausesValidation = false;
			this.btnGuardar.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.btnGuardar.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(50)), ((System.Byte)(50)), ((System.Byte)(50)));
			this.btnGuardar.Image = ((System.Drawing.Image)(resources.GetObject("btnGuardar.Image")));
			this.btnGuardar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnGuardar.Location = new System.Drawing.Point(8, 250);
			this.btnGuardar.Name = "btnGuardar";
			this.btnGuardar.Size = new System.Drawing.Size(72, 23);
			this.btnGuardar.TabIndex = 6;
			this.btnGuardar.Text = "Guardar";
			this.btnGuardar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnGuardar.Click += new System.EventHandler(this.btnGuardar_Click);
			// 
			// groupBox1
			// 
			this.groupBox1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.groupBox1.Location = new System.Drawing.Point(0, 240);
			this.groupBox1.Name = "groupBox1";
			this.groupBox1.Size = new System.Drawing.Size(464, 8);
			this.groupBox1.TabIndex = 312;
			this.groupBox1.TabStop = false;
			// 
			// txtIdNominaSalidas
			// 
			this.txtIdNominaSalidas.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.txtIdNominaSalidas.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtIdNominaSalidas.Enabled = false;
			this.txtIdNominaSalidas.Location = new System.Drawing.Point(440, 250);
			this.txtIdNominaSalidas.Name = "txtIdNominaSalidas";
			this.txtIdNominaSalidas.PromptChar = ' ';
			this.txtIdNominaSalidas.Size = new System.Drawing.Size(16, 21);
			this.txtIdNominaSalidas.TabIndex = 311;
			this.txtIdNominaSalidas.Visible = false;
			// 
			// label4
			// 
			this.label4.AutoSize = true;
			this.label4.Location = new System.Drawing.Point(8, 56);
			this.label4.Name = "label4";
			this.label4.Size = new System.Drawing.Size(26, 16);
			this.label4.TabIndex = 303;
			this.label4.Text = "Tipo";
			this.label4.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label2
			// 
			this.label2.AutoSize = true;
			this.label2.Location = new System.Drawing.Point(8, 152);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(38, 16);
			this.label2.TabIndex = 301;
			this.label2.Text = "Motivo";
			this.label2.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// lblBodega
			// 
			this.lblBodega.AutoSize = true;
			this.lblBodega.Location = new System.Drawing.Point(8, 90);
			this.lblBodega.Name = "lblBodega";
			this.lblBodega.Size = new System.Drawing.Size(72, 16);
			this.lblBodega.TabIndex = 300;
			this.lblBodega.Text = "Tipo de Parto";
			this.lblBodega.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// dtpHasta
			// 
			appearance2.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtpHasta.Appearance = appearance2;
			this.dtpHasta.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton1.Caption = "Today";
			this.dtpHasta.DateButtons.Add(dateButton1);
			this.dtpHasta.Format = "dd/MM/yyyy";
			this.dtpHasta.Location = new System.Drawing.Point(280, 120);
			this.dtpHasta.Name = "dtpHasta";
			this.dtpHasta.NonAutoSizeHeight = 23;
			this.dtpHasta.Size = new System.Drawing.Size(112, 21);
			this.dtpHasta.SpinButtonsVisible = true;
			this.dtpHasta.TabIndex = 4;
			this.dtpHasta.Value = ((object)(resources.GetObject("dtpHasta.Value")));
			this.dtpHasta.Visible = false;
			this.dtpHasta.KeyDown += new System.Windows.Forms.KeyEventHandler(this.dtpHasta_KeyDown);
			// 
			// lblHasta
			// 
			this.lblHasta.AutoSize = true;
			this.lblHasta.Location = new System.Drawing.Point(224, 122);
			this.lblHasta.Name = "lblHasta";
			this.lblHasta.Size = new System.Drawing.Size(34, 16);
			this.lblHasta.TabIndex = 299;
			this.lblHasta.Text = "Hasta";
			this.lblHasta.Visible = false;
			// 
			// dtpDesde
			// 
			appearance3.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtpDesde.Appearance = appearance3;
			this.dtpDesde.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton2.Caption = "Today";
			this.dtpDesde.DateButtons.Add(dateButton2);
			this.dtpDesde.Enabled = false;
			this.dtpDesde.Format = "dd/MM/yyyy";
			this.dtpDesde.Location = new System.Drawing.Point(96, 120);
			this.dtpDesde.Name = "dtpDesde";
			this.dtpDesde.NonAutoSizeHeight = 23;
			this.dtpDesde.Size = new System.Drawing.Size(112, 21);
			this.dtpDesde.SpinButtonsVisible = true;
			this.dtpDesde.TabIndex = 3;
			this.dtpDesde.Value = ((object)(resources.GetObject("dtpDesde.Value")));
			this.dtpDesde.KeyDown += new System.Windows.Forms.KeyEventHandler(this.dtpDesde_KeyDown);
			this.dtpDesde.BeforeDropDown += new System.ComponentModel.CancelEventHandler(this.dtpDesde_BeforeDropDown);
			this.dtpDesde.ValueChanged += new System.EventHandler(this.dtpDesde_ValueChanged);
			// 
			// lblDesde
			// 
			this.lblDesde.AutoSize = true;
			this.lblDesde.Location = new System.Drawing.Point(8, 122);
			this.lblDesde.Name = "lblDesde";
			this.lblDesde.Size = new System.Drawing.Size(37, 16);
			this.lblDesde.TabIndex = 298;
			this.lblDesde.Text = "Desde";
			// 
			// txtMotivo
			// 
			appearance4.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtMotivo.Appearance = appearance4;
			this.txtMotivo.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtMotivo.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtMotivo.Location = new System.Drawing.Point(96, 152);
			this.txtMotivo.MaxLength = 300;
			this.txtMotivo.Multiline = true;
			this.txtMotivo.Name = "txtMotivo";
			this.txtMotivo.Size = new System.Drawing.Size(360, 80);
			this.txtMotivo.TabIndex = 5;
			this.txtMotivo.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtMotivo_KeyDown);
			// 
			// cmbTipo
			// 
			this.cmbTipo.CausesValidation = false;
			this.cmbTipo.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbTipo.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbTipo.DataSource = this.ultraDataSource1;
			ultraGridColumn4.Header.VisiblePosition = 0;
			ultraGridColumn4.Hidden = true;
			ultraGridColumn5.Header.VisiblePosition = 1;
			ultraGridColumn5.Width = 296;
			ultraGridBand2.Columns.AddRange(new object[] {
																										 ultraGridColumn4,
																										 ultraGridColumn5});
			this.cmbTipo.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			this.cmbTipo.DisplayMember = "Descripcion";
			this.cmbTipo.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbTipo.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbTipo.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbTipo.Location = new System.Drawing.Point(96, 54);
			this.cmbTipo.MaxDropDownItems = 30;
			this.cmbTipo.Name = "cmbTipo";
			this.cmbTipo.Size = new System.Drawing.Size(296, 21);
			this.cmbTipo.TabIndex = 1;
			this.cmbTipo.ValueMember = "idNominaTiposDeSalidas";
			this.cmbTipo.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cmbTipo_KeyDown);
			this.cmbTipo.ValueChanged += new System.EventHandler(this.cmbTipo_ValueChanged);
			// 
			// ultraDataSource1
			// 
			ultraDataColumn1.DataType = typeof(int);
			this.ultraDataSource1.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn1,
																																 ultraDataColumn2});
			// 
			// cmbParto
			// 
			this.cmbParto.CausesValidation = false;
			this.cmbParto.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbParto.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbParto.DataSource = this.ultraDataSource2;
			ultraGridColumn6.Header.VisiblePosition = 0;
			ultraGridColumn6.Hidden = true;
			ultraGridColumn7.Header.VisiblePosition = 1;
			ultraGridColumn7.Width = 60;
			ultraGridColumn8.Header.VisiblePosition = 2;
			ultraGridColumn8.Width = 296;
			ultraGridBand3.Columns.AddRange(new object[] {
																										 ultraGridColumn6,
																										 ultraGridColumn7,
																										 ultraGridColumn8});
			this.cmbParto.DisplayLayout.BandsSerializer.Add(ultraGridBand3);
			this.cmbParto.DisplayMember = "Descripcion";
			this.cmbParto.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbParto.Enabled = false;
			this.cmbParto.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbParto.Location = new System.Drawing.Point(96, 88);
			this.cmbParto.MaxDropDownItems = 30;
			this.cmbParto.Name = "cmbParto";
			this.cmbParto.Size = new System.Drawing.Size(296, 21);
			this.cmbParto.TabIndex = 2;
			this.cmbParto.ValueMember = "idNominaTiposDeParto";
			// 
			// ultraDataSource2
			// 
			ultraDataColumn3.DataType = typeof(int);
			ultraDataColumn4.DataType = typeof(int);
			this.ultraDataSource2.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn3,
																																 ultraDataColumn4,
																																 ultraDataColumn5});
			// 
			// cdsSeteoF
			// 
			this.cdsSeteoF.BindingContextCtrl = this;
			this.cdsSeteoF.DataLibrary = "LibFacturacion";
			this.cdsSeteoF.DataLibraryUrl = "";
			this.cdsSeteoF.DataSetDef = "dsSeteoF";
			this.cdsSeteoF.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsSeteoF.Name = "cdsSeteoF";
			this.cdsSeteoF.SchemaClassName = "LibFacturacion.DataClass";
			this.cdsSeteoF.SchemaDef = null;
			this.cdsSeteoF.BeforeFill += new C1.Data.FillEventHandler(this.cdsSeteoF_BeforeFill);
			// 
			// frmNominaPermisosFaltas
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(464, 280);
			this.ControlBox = false;
			this.Controls.Add(this.cmbParto);
			this.Controls.Add(this.cmbTipo);
			this.Controls.Add(this.txtMotivo);
			this.Controls.Add(this.cmbVendedor);
			this.Controls.Add(this.groupBox2);
			this.Controls.Add(this.label7);
			this.Controls.Add(this.label4);
			this.Controls.Add(this.label2);
			this.Controls.Add(this.lblBodega);
			this.Controls.Add(this.lblHasta);
			this.Controls.Add(this.lblDesde);
			this.Controls.Add(this.btnCancelar);
			this.Controls.Add(this.btnGuardar);
			this.Controls.Add(this.groupBox1);
			this.Controls.Add(this.txtIdNominaSalidas);
			this.Controls.Add(this.dtpHasta);
			this.Controls.Add(this.dtpDesde);
			this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
			this.KeyPreview = true;
			this.MaximizeBox = false;
			this.Name = "frmNominaPermisosFaltas";
			this.ShowInTaskbar = false;
			this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
			this.Text = "Registro de Permisos y Faltas";
			this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.frmNominaPermisosFaltas_KeyDown);
			this.Load += new System.EventHandler(this.frmNominaPermisosFaltas_Load);
			((System.ComponentModel.ISupportInitialize)(this.cmbVendedor)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdNominaSalidas)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtpHasta)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtpDesde)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtMotivo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbTipo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbParto)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void frmNominaPermisosFaltas_Load(object sender, System.EventArgs e)
		{
			this.cmbVendedor.DataSource = Funcion.dvProcedimiento(cdsSeteoF, string.Format("Exec Com_ListaVendedoresMovimiento '{0}', {1}, 0, 1", DateTime.Today.ToString("yyyyMMdd"), idVendedor));
			this.cmbTipo.DataSource = Funcion.dvProcedimiento(cdsSeteoF, string.Format("Exec Listas 62, {0}", bPermisosFaltas));
			this.cmbParto.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Exec Listas 63, 0");

			this.cmbVendedor.Value = idVendedor;
			
			if (bPermisosFaltas == 0)
			{
				this.lblHasta.Visible = true;
				this.dtpHasta.Visible = true;

				this.dtpDesde.CalendarInfo.MinDate = dtDesde;
				this.dtpDesde.CalendarInfo.MaxDate = dtHasta;		
				this.dtpDesde.Value = DateTime.Today;
				this.dtpHasta.CalendarInfo.MinDate = DateTime.Today;								
				this.dtpHasta.Value = DateTime.Today;

				this.dtpDesde.Enabled = true;
			}
			else if (bPermisosFaltas == 2)
			{
				this.lblHasta.Visible = true;
				this.dtpHasta.Visible = true;
				this.dtpDesde.Enabled = false;
				this.dtpHasta.Enabled = false;
				this.cmbTipo.Enabled=false;
				this.cmbTipo.Value=3;
				this.dtpDesde.Value=dtDesde;
				this.dtpHasta.Value=dtHasta;
			}

			else
			{
				this.lblDesde.Text = "Fecha";				
				this.dtpDesde.Enabled = true;
				this.dtpDesde.CalendarInfo.MinDate = dtDesde.AddDays(-3);
				this.dtpDesde.CalendarInfo.MaxDate = DateTime.Today;				
				this.dtpDesde.Value = DateTime.Today;
			}
		}

		private void btnCancelar_Click(object sender, System.EventArgs e)
		{
			this.DialogResult = DialogResult.No;
		}

		private void cmbTipo_ValueChanged(object sender, System.EventArgs e)
		{
			if (this.cmbTipo.ActiveRow != null)
			{
				if ((int)this.cmbTipo.Value == 4 || (int)this.cmbTipo.Value == 5)
					this.cmbParto.Enabled = true;
				else
				{
					this.cmbParto.Enabled = false;
					this.cmbParto.Value = System.DBNull.Value;
				}
			}
		}

		private void NuevoMovimiento()
		{
			int iUltDiames = FuncionesProcedimientos.iUltimoDiaDelMes(dtUltimoDiaMes.Month, dtUltimoDiaMes.Year);

			if (Convert.ToDateTime(this.dtpHasta.Value) < dtUltimoDiaMes)
			{				
				this.dtpHasta.Value = Convert.ToDateTime(this.dtpHasta.Value).AddDays(1);

				int iNumero = Funcion.iEscalarSQL(cdsSeteoF, string.Format("Select Top 1 Numero From Com_AsignacionDeVendedores Where Periodo = '{0}' Order By Numero Desc", sPeriodo));
				int iNumeroDias = FuncionesProcedimientos.CalculaDias((DateTime)this.dtpHasta.Value, dtUltimoDiaMes, cdsSeteoF);
																
				if (Convert.ToDateTime(this.dtpDesde.Value).Day > 1)
				{
					string sSQLA = string.Format("Exec Com_ActualizaMetasAVSEnMovimientos {0}, '{1}', {2}", 
						idCom_AsignacionDeVendedores, Convert.ToDateTime(this.dtpDesde.Value).AddDays(-1).ToString("yyyyMMdd"), iUltDiames);
					Funcion.EjecutaSQL(cdsSeteoF, sSQLA);

					idCom_AsignacionDeVendedores = 0;
				}

				decimal dSemanal = FuncionesProcedimientos.CalculaMeta(iBodega, Convert.ToDateTime(this.dtpDesde.Value).Month, Convert.ToDateTime(this.dtpDesde.Value).Year, idCargo, iNumeroDias, 1, cdsSeteoF);
				decimal dContado = FuncionesProcedimientos.CalculaMeta(iBodega, Convert.ToDateTime(this.dtpDesde.Value).Month, Convert.ToDateTime(this.dtpDesde.Value).Year, idCargo, iNumeroDias, 2, cdsSeteoF);
				decimal dCredito = FuncionesProcedimientos.CalculaMeta(iBodega, Convert.ToDateTime(this.dtpDesde.Value).Month, Convert.ToDateTime(this.dtpDesde.Value).Year, idCargo, iNumeroDias, 3, cdsSeteoF);
				decimal dGEx = FuncionesProcedimientos.CalculaMeta(iBodega, Convert.ToDateTime(this.dtpDesde.Value).Month, Convert.ToDateTime(this.dtpDesde.Value).Year, idCargo, iNumeroDias, 4, cdsSeteoF);

				string sSQL = string.Format("Exec Com_GuardaAsignacionVendedores {0}, '{1}', {2}, {3}, {4}, {5}, '{6}', '{7}', {8}, {9}, {10}, {11}, {12}, {13}, {14}, {15}, {16}, {17}, '{18}', {19}, {20}, {21}, {22}, {23}",
					idCom_AsignacionDeVendedores, sPeriodo, iBodega, idCargo, (int)this.cmbVendedor.Value, idCom_Rango, //5
					Convert.ToDateTime(this.dtpHasta.Value).ToString("yyyyMMdd"), dtUltimoDiaMes.ToString("yyyyMMdd"), iNumeroDias,  //7
					0.00m, 0.00m, 0.00m, 0.00m, 0.00m, 0.00m, idCom_Estado, 0, false, "20010101", false, dSemanal, dContado, dCredito, dGEx);
				Funcion.EjecutaSQL(cdsSeteoF, sSQL);

				//				string sSQLMV = string.Format("Exec Com_GuardaAsignacionVendedores {0}, '{1}', {2}, {3}, {4}, {5}, '{6}', '{7}', {8}, {9}, {10}, {11}, {12}, {13}, {14}, {15}, {16}",
				//					idCom_AsignacionDeVendedores, sPeriodo, iBodega, idCargo, (int)this.cmbVendedor.Value, idCom_Rango,
				//					Convert.ToDateTime(this.dtpHasta.Value).ToString("yyyyMMdd"), dtUltimoDiaMes.ToString("yyyyMMdd"),
				//					iNumeroDias, dMetaDiaria, 0, dMetaDiariaGE, dMetaGE, dMetaDiariaCarpa, dMetaCarpa, idCom_Estado, iNumero);
				//				Funcion.EjecutaSQL(cdsSeteoF, sSQLMV);
			}
			else
			{
				if (Convert.ToDateTime(this.dtpDesde.Value).Day > 1)
				{
					string sSQLA = string.Format("Exec Com_ActualizaMetasAVSEnMovimientos {0}, '{1}', {2}", 
						idCom_AsignacionDeVendedores, Convert.ToDateTime(this.dtpDesde.Value).AddDays(-1).ToString("yyyyMMdd"), iUltDiames);
					Funcion.EjecutaSQL(cdsSeteoF, sSQLA);					
				}
			}
		}

		private void btnGuardar_Click(object sender, System.EventArgs e)
		{
			if (bPermisosFaltas == 1) this.dtpHasta.Value = Convert.ToDateTime(this.dtpDesde.Value);

			#region Valida campos requeridos
			if (!Validacion.vbComboVacio(this.cmbTipo, "Seleccione el tipo de salida")) return;	
			
			if ((int)this.cmbTipo.Value == 4 || (int)this.cmbTipo.Value == 5)
				if (!Validacion.vbComboVacio(this.cmbParto, "Seleccione el tipo de parto")) return;	
			
			if (!Validacion.vbFechaEnDocumentos(this.dtpDesde, "Seleccione la Fecha Inicial", false, "", cdsSeteoF)) return;
			
			if (!Validacion.vbFechaEnDocumentos(this.dtpHasta, "Seleccione la Fecha Final", false, "", cdsSeteoF)) return;

			if (!Validacion.vbComparaFechas(Convert.ToDateTime(this.dtpDesde.Value), Convert.ToDateTime(this.dtpHasta.Value), this.dtpDesde, "La Fecha Final no puede ser menor a la Fecha Inicial", false)) return;			
			#endregion Valida campos requeridos

			#region Valida Faltas - Permisos - Ventas
			int iCont = 0;
			string sSQLValida = "";

			if (bPermisosFaltas == 1)
			{				
				sSQLValida = string.Format("Exec NominaValidaFaltas {0}, '{1}'", (int)this.cmbVendedor.Value, Convert.ToDateTime(this.dtpDesde.Value).ToString("yyyyMMdd"));
				iCont = Funcion.iEscalarSQL(cdsSeteoF, sSQLValida);
				
				if (iCont > 0)
				{
					MessageBox.Show("Ya existe una falta registrada en la fecha seleccionada", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					this.dtpDesde.Focus();
					return;
				}

				iCont = 0;
				sSQLValida = string.Format("Exec Com_ValidaMovimientosVentas {0}, 0, '{1}', '{2}', 1", 
					(int)this.cmbVendedor.Value, Convert.ToDateTime(this.dtpDesde.Value).ToString("yyyyMMdd"), Convert.ToDateTime(this.dtpHasta.Value).ToString("yyyyMMdd"));
				iCont = Funcion.iEscalarSQL(cdsSeteoF, sSQLValida);
				
				if (iCont > 0)
				{
					MessageBox.Show("El empleado registra Ventas en la fecha seleccionada", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					this.dtpDesde.Focus();
					return;
				}

				iCont = 0;
				sSQLValida = string.Format("Exec Com_ValidaMovimientoPermisos {0}, '{1}', {2}", 
					(int)this.cmbVendedor.Value, Convert.ToDateTime(this.dtpDesde.Value).ToString("yyyyMMdd"), (int)this.txtIdNominaSalidas.Value);
				iCont = Funcion.iEscalarSQL(cdsSeteoF, sSQLValida);
				
				if (iCont > 0)
				{
					MessageBox.Show("El empleado registra permiso (VACACIONES - MATERNIDAD - PATERNIDAD - ENFERMEDAD - FALTAS) en la fecha seleccionada", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					this.dtpDesde.Focus();
					return;
				}
			}
			#endregion Valida Faltas - Permisos - Ventas

			#region Valida Permisos
			iCont = 0;
			sSQLValida = string.Format("Exec Com_ValidaMovimientoPermisos {0}, '{1}', {2}", 
				(int)this.cmbVendedor.Value, Convert.ToDateTime(this.dtpDesde.Value).ToString("yyyyMMdd"), (int)this.txtIdNominaSalidas.Value);
			iCont = Funcion.iEscalarSQL(cdsSeteoF, sSQLValida);
				
			if (iCont > 0)
			{
				MessageBox.Show("El empleado registra permiso (VACACIONES - MATERNIDAD - PATERNIDAD - ENFERMEDAD - FALTAS) en la fecha inicial seleccionada", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.dtpHasta.Focus();
				return;
			}

			iCont = 0;
			sSQLValida = string.Format("Exec Com_ValidaMovimientoPermisos {0}, '{1}', {2}", 
				(int)this.cmbVendedor.Value, Convert.ToDateTime(this.dtpHasta.Value).ToString("yyyyMMdd"), (int)this.txtIdNominaSalidas.Value);
			iCont = Funcion.iEscalarSQL(cdsSeteoF, sSQLValida);
				
			if (iCont > 0)
			{
				MessageBox.Show("El empleado registra permiso (VACACIONES - MATERNIDAD - PATERNIDAD - ENFERMEDAD - FALTAS) en la fecha final seleccionada", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.dtpHasta.Focus();
				return;
			}

			iCont = 0;
			sSQLValida = string.Format("Exec Com_ValidaMovimientosVentas {0}, 0, '{1}', '{2}', 0", 
				(int)this.cmbVendedor.Value, Convert.ToDateTime(this.dtpDesde.Value).ToString("yyyyMMdd"), Convert.ToDateTime(this.dtpHasta.Value).ToString("yyyyMMdd"));
			iCont = Funcion.iEscalarSQL(cdsSeteoF, sSQLValida);

			if (iCont > 0)
			{
				MessageBox.Show("El empleado registra Ventas en el rango de fechas seleccionado", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.dtpHasta.Focus();
				return;
			}
			#endregion Valida Permisos

			int idNomina = Funcion.iEscalarSQL(cdsSeteoF, string.Format("Select idNomina From Nomina Where idPersonal = {0}", (int)this.cmbVendedor.Value));

			int idParto = 0;
			if (this.cmbParto.ActiveRow != null) idParto = (int)this.cmbParto.Value;

			string sSQL = string.Format("Exec NominaGuardaSalidas {0}, {1}, {2}, {3}, '{4}', '{5}', '{6}', '{7}'",
				(int)this.txtIdNominaSalidas.Value, idNomina, (int)this.cmbTipo.Value, idParto, Convert.ToDateTime(this.dtpDesde.Value).ToString("yyyyMMdd"), Convert.ToDateTime(this.dtpHasta.Value).ToString("yyyyMMdd"),
				this.txtMotivo.Text.ToString(), "");
			this.txtIdNominaSalidas.Value = Funcion.iEscalarSQL(cdsSeteoF, sSQL);
						
			#region Genera Movimiento si la fecha final es menor al ultimo dia del mes solo para permisos - vacaciones - no jefes de local
			if (bPermisosFaltas == 0)			
				if (idCargo != 1 || idCargo != 21)
					NuevoMovimiento();						
			#endregion Genera Movimiento si la fecha final es menor al ultimo dia del mes solo para permisos - vacaciones - no jefes de local

			string sSQLMail = string.Format("Exec Com_MailPermisosVacacionesFaltas {0}, '{1}', {2}, {3}", 
				(int)this.txtIdNominaSalidas.Value, sPeriodo, iBodega, (int)this.cmbTipo.Value);
			Funcion.EjecutaSQL(cdsSeteoF, sSQLMail);

			this.DialogResult = DialogResult.OK;	
		}

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);		
		}		

		private void cmbTipo_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter)
			{
				if (this.cmbParto.Enabled) this.cmbParto.Focus(); else this.dtpDesde.Focus();
			}
		}

		private void dtpDesde_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter)
				if (bPermisosFaltas == 1) this.txtMotivo.Focus(); else this.dtpHasta.Focus();			
		}

		private void dtpHasta_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.txtMotivo.Focus();
		}

		private void txtMotivo_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.btnGuardar.Focus();
		}

		private void frmNominaPermisosFaltas_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Escape) this.DialogResult = DialogResult.No;
		}

		private void dtpDesde_BeforeDropDown(object sender, System.ComponentModel.CancelEventArgs e)
		{
		
		}

		private void dtpDesde_ValueChanged(object sender, System.EventArgs e)
		{
			if (bPermisosFaltas == 0)
				if (this.dtpHasta.Value != System.DBNull.Value)
					//if (Convert.ToDateTime(this.dtpDesde.Value) > Convert.ToDateTime(this.dtpHasta.Value))
						this.dtpHasta.CalendarInfo.MinDate = Convert.ToDateTime(this.dtpDesde.Value);
		}
	}
}
