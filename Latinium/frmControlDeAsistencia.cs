using System.Data;
using System.Data.SqlClient;
using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.IO;
using System.Diagnostics;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de frmControlDeAsistencia.
	/// </summary>
	public class frmControlDeAsistencia : System.Windows.Forms.Form
	{
		private System.Windows.Forms.Label label5;
		private System.Windows.Forms.PictureBox pictureBox2;
		private System.Windows.Forms.Label label6;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtFecha;
		private C1.Data.C1DataSet cdsSeteoF;
		private System.Windows.Forms.Timer timer1;
		private System.Windows.Forms.Timer timer2;
		private System.Windows.Forms.Label label1;
		private System.Windows.Forms.Label label2;
		private System.Windows.Forms.Label label3;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtUsuario;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtCedula;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtNombre;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtCargo;
		private System.Windows.Forms.Label label4;
		private System.Windows.Forms.Label label7;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtAlmcen;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtidNomina;
		private System.Windows.Forms.Button btnGuardar;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtIdGrupo;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbEstadoEntrada;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource1;
		private System.Windows.Forms.Label label8;
		private Infragistics.Win.UltraWinGrid.UltraGrid dgVerificacion;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource2;
		private System.Windows.Forms.GroupBox groupBox3;
		public Infragistics.Win.UltraWinEditors.UltraTextEditor txtNotas;
		private System.Windows.Forms.Label label30;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtidBodega;
		private System.Windows.Forms.Button btnNuevo;
		private System.ComponentModel.IContainer components;

		public frmControlDeAsistencia()
		{
			//
			// Necesario para admitir el Diseñador de Windows Forms
			//
			InitializeComponent();

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
			this.components = new System.ComponentModel.Container();
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmControlDeAsistencia));
			Infragistics.Win.Appearance appearance1 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton1 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idEstadoIngreso");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idEstadoIngreso");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idNomina");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Fecha");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("NOMBRE");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ENTRADA");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("SALIDAALMUERZO");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ENTRADAALMUERZO");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("SALIDA");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Observaciones");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn11 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bodega");
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance11 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance12 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn3 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idNomina");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn4 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Fecha");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn5 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("NOMBRE");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn6 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("ENTRADA");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn7 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("SALIDAALMUERZO");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn8 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("ENTRADAALMUERZO");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn9 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("SALIDA");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn10 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Observaciones");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn11 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Bodega");
			Infragistics.Win.Appearance appearance13 = new Infragistics.Win.Appearance();
			this.label5 = new System.Windows.Forms.Label();
			this.pictureBox2 = new System.Windows.Forms.PictureBox();
			this.label6 = new System.Windows.Forms.Label();
			this.dtFecha = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.timer1 = new System.Windows.Forms.Timer(this.components);
			this.timer2 = new System.Windows.Forms.Timer(this.components);
			this.label1 = new System.Windows.Forms.Label();
			this.label2 = new System.Windows.Forms.Label();
			this.label3 = new System.Windows.Forms.Label();
			this.txtUsuario = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtCedula = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtNombre = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtCargo = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label4 = new System.Windows.Forms.Label();
			this.label7 = new System.Windows.Forms.Label();
			this.txtAlmcen = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtidNomina = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.btnGuardar = new System.Windows.Forms.Button();
			this.txtIdGrupo = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.cmbEstadoEntrada = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.ultraDataSource1 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.label8 = new System.Windows.Forms.Label();
			this.dgVerificacion = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource2 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.groupBox3 = new System.Windows.Forms.GroupBox();
			this.txtNotas = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label30 = new System.Windows.Forms.Label();
			this.txtidBodega = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.btnNuevo = new System.Windows.Forms.Button();
			((System.ComponentModel.ISupportInitialize)(this.dtFecha)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtUsuario)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCedula)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNombre)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCargo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtAlmcen)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtidNomina)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdGrupo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbEstadoEntrada)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dgVerificacion)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNotas)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtidBodega)).BeginInit();
			this.SuspendLayout();
			// 
			// label5
			// 
			this.label5.AutoSize = true;
			this.label5.Font = new System.Drawing.Font("Tahoma", 14F, System.Drawing.FontStyle.Bold);
			this.label5.Location = new System.Drawing.Point(144, 8);
			this.label5.Name = "label5";
			this.label5.Size = new System.Drawing.Size(271, 26);
			this.label5.TabIndex = 66;
			this.label5.Text = "Control Asistencia Personal";
			// 
			// pictureBox2
			// 
			this.pictureBox2.Image = ((System.Drawing.Image)(resources.GetObject("pictureBox2.Image")));
			this.pictureBox2.Location = new System.Drawing.Point(456, 48);
			this.pictureBox2.Name = "pictureBox2";
			this.pictureBox2.Size = new System.Drawing.Size(24, 16);
			this.pictureBox2.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage;
			this.pictureBox2.TabIndex = 67;
			this.pictureBox2.TabStop = false;
			// 
			// label6
			// 
			this.label6.AutoSize = true;
			this.label6.Location = new System.Drawing.Point(232, 48);
			this.label6.Name = "label6";
			this.label6.Size = new System.Drawing.Size(36, 16);
			this.label6.TabIndex = 304;
			this.label6.Text = "Fecha";
			// 
			// dtFecha
			// 
			appearance1.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtFecha.Appearance = appearance1;
			this.dtFecha.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton1.Caption = "Today";
			this.dtFecha.DateButtons.Add(dateButton1);
			this.dtFecha.Enabled = false;
			this.dtFecha.Format = "dd/MM/yyyy HH:mm:ss";
			this.dtFecha.Location = new System.Drawing.Point(280, 48);
			this.dtFecha.Name = "dtFecha";
			this.dtFecha.NonAutoSizeHeight = 23;
			this.dtFecha.Size = new System.Drawing.Size(168, 21);
			this.dtFecha.SpinButtonsVisible = true;
			this.dtFecha.TabIndex = 303;
			this.dtFecha.Value = ((object)(resources.GetObject("dtFecha.Value")));
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
			// timer1
			// 
			this.timer1.Enabled = true;
			this.timer1.Tick += new System.EventHandler(this.timer1_Tick);
			// 
			// timer2
			// 
			this.timer2.Interval = 5000;
			// 
			// label1
			// 
			this.label1.AutoSize = true;
			this.label1.Location = new System.Drawing.Point(8, 48);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(43, 16);
			this.label1.TabIndex = 305;
			this.label1.Text = "Usuario";
			// 
			// label2
			// 
			this.label2.AutoSize = true;
			this.label2.Location = new System.Drawing.Point(8, 112);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(40, 16);
			this.label2.TabIndex = 306;
			this.label2.Text = "Cedula";
			// 
			// label3
			// 
			this.label3.AutoSize = true;
			this.label3.Location = new System.Drawing.Point(224, 112);
			this.label3.Name = "label3";
			this.label3.Size = new System.Drawing.Size(44, 16);
			this.label3.TabIndex = 307;
			this.label3.Text = "Nombre";
			// 
			// txtUsuario
			// 
			appearance2.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtUsuario.Appearance = appearance2;
			this.txtUsuario.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtUsuario.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtUsuario.Enabled = false;
			this.txtUsuario.Location = new System.Drawing.Point(64, 48);
			this.txtUsuario.MaxLength = 10;
			this.txtUsuario.Name = "txtUsuario";
			this.txtUsuario.Size = new System.Drawing.Size(150, 21);
			this.txtUsuario.TabIndex = 308;
			// 
			// txtCedula
			// 
			appearance3.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtCedula.Appearance = appearance3;
			this.txtCedula.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtCedula.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtCedula.Location = new System.Drawing.Point(64, 112);
			this.txtCedula.MaxLength = 10;
			this.txtCedula.Name = "txtCedula";
			this.txtCedula.Size = new System.Drawing.Size(150, 21);
			this.txtCedula.TabIndex = 309;
			this.txtCedula.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtCedula_KeyDown);
			this.txtCedula.Validating += new System.ComponentModel.CancelEventHandler(this.txtCedula_Validating);
			this.txtCedula.Validated += new System.EventHandler(this.txtCedula_Validated);
			// 
			// txtNombre
			// 
			appearance4.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtNombre.Appearance = appearance4;
			this.txtNombre.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtNombre.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNombre.Enabled = false;
			this.txtNombre.Location = new System.Drawing.Point(280, 112);
			this.txtNombre.MaxLength = 10;
			this.txtNombre.Name = "txtNombre";
			this.txtNombre.Size = new System.Drawing.Size(240, 21);
			this.txtNombre.TabIndex = 310;
			// 
			// txtCargo
			// 
			appearance5.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtCargo.Appearance = appearance5;
			this.txtCargo.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtCargo.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtCargo.Enabled = false;
			this.txtCargo.Location = new System.Drawing.Point(64, 80);
			this.txtCargo.MaxLength = 10;
			this.txtCargo.Name = "txtCargo";
			this.txtCargo.Size = new System.Drawing.Size(150, 21);
			this.txtCargo.TabIndex = 311;
			// 
			// label4
			// 
			this.label4.AutoSize = true;
			this.label4.Location = new System.Drawing.Point(8, 80);
			this.label4.Name = "label4";
			this.label4.Size = new System.Drawing.Size(35, 16);
			this.label4.TabIndex = 312;
			this.label4.Text = "Cargo";
			// 
			// label7
			// 
			this.label7.AutoSize = true;
			this.label7.Location = new System.Drawing.Point(224, 80);
			this.label7.Name = "label7";
			this.label7.Size = new System.Drawing.Size(48, 16);
			this.label7.TabIndex = 313;
			this.label7.Text = "Almacen";
			// 
			// txtAlmcen
			// 
			appearance6.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtAlmcen.Appearance = appearance6;
			this.txtAlmcen.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtAlmcen.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtAlmcen.Enabled = false;
			this.txtAlmcen.Location = new System.Drawing.Point(280, 80);
			this.txtAlmcen.MaxLength = 10;
			this.txtAlmcen.Name = "txtAlmcen";
			this.txtAlmcen.Size = new System.Drawing.Size(150, 21);
			this.txtAlmcen.TabIndex = 314;
			// 
			// txtidNomina
			// 
			this.txtidNomina.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.txtidNomina.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtidNomina.Enabled = false;
			this.txtidNomina.Location = new System.Drawing.Point(496, 8);
			this.txtidNomina.Name = "txtidNomina";
			this.txtidNomina.PromptChar = ' ';
			this.txtidNomina.Size = new System.Drawing.Size(88, 21);
			this.txtidNomina.SpinButtonAlignment = Infragistics.Win.ButtonAlignment.Left;
			this.txtidNomina.TabIndex = 785;
			this.txtidNomina.Visible = false;
			// 
			// btnGuardar
			// 
			this.btnGuardar.Image = ((System.Drawing.Image)(resources.GetObject("btnGuardar.Image")));
			this.btnGuardar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnGuardar.Location = new System.Drawing.Point(496, 208);
			this.btnGuardar.Name = "btnGuardar";
			this.btnGuardar.Size = new System.Drawing.Size(82, 23);
			this.btnGuardar.TabIndex = 787;
			this.btnGuardar.Text = "&Guardar";
			this.btnGuardar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnGuardar.Click += new System.EventHandler(this.btnGuardar_Click);
			// 
			// txtIdGrupo
			// 
			this.txtIdGrupo.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.txtIdGrupo.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtIdGrupo.Enabled = false;
			this.txtIdGrupo.Location = new System.Drawing.Point(496, 40);
			this.txtIdGrupo.Name = "txtIdGrupo";
			this.txtIdGrupo.PromptChar = ' ';
			this.txtIdGrupo.Size = new System.Drawing.Size(88, 21);
			this.txtIdGrupo.SpinButtonAlignment = Infragistics.Win.ButtonAlignment.Left;
			this.txtIdGrupo.TabIndex = 788;
			this.txtIdGrupo.Visible = false;
			// 
			// cmbEstadoEntrada
			// 
			appearance7.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbEstadoEntrada.Appearance = appearance7;
			this.cmbEstadoEntrada.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.cmbEstadoEntrada.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbEstadoEntrada.DataSource = this.ultraDataSource1;
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridColumn2.Width = 150;
			ultraGridBand1.Columns.AddRange(new object[] {
																										 ultraGridColumn1,
																										 ultraGridColumn2});
			this.cmbEstadoEntrada.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.cmbEstadoEntrada.DisplayMember = "Nombre";
			this.cmbEstadoEntrada.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbEstadoEntrada.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbEstadoEntrada.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbEstadoEntrada.Location = new System.Drawing.Point(64, 144);
			this.cmbEstadoEntrada.Name = "cmbEstadoEntrada";
			this.cmbEstadoEntrada.Size = new System.Drawing.Size(152, 21);
			this.cmbEstadoEntrada.TabIndex = 789;
			this.cmbEstadoEntrada.ValueMember = "idEstadoIngreso";
			// 
			// ultraDataSource1
			// 
			ultraDataColumn1.DataType = typeof(int);
			this.ultraDataSource1.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn1,
																																 ultraDataColumn2});
			// 
			// label8
			// 
			this.label8.AutoSize = true;
			this.label8.Location = new System.Drawing.Point(8, 144);
			this.label8.Name = "label8";
			this.label8.Size = new System.Drawing.Size(26, 16);
			this.label8.TabIndex = 790;
			this.label8.Text = "Tipo";
			// 
			// dgVerificacion
			// 
			this.dgVerificacion.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.dgVerificacion.Cursor = System.Windows.Forms.Cursors.Default;
			this.dgVerificacion.DataSource = this.ultraDataSource2;
			appearance8.BackColor = System.Drawing.Color.White;
			this.dgVerificacion.DisplayLayout.Appearance = appearance8;
			ultraGridColumn3.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn3.Header.VisiblePosition = 0;
			ultraGridColumn3.Hidden = true;
			ultraGridColumn4.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn4.Header.VisiblePosition = 1;
			ultraGridColumn5.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn5.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn5.Header.Caption = "Nombre";
			ultraGridColumn5.Header.VisiblePosition = 2;
			ultraGridColumn5.Width = 200;
			ultraGridColumn6.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn6.Format = "dd/MM/yyy HH:mm:ss";
			ultraGridColumn6.Header.Caption = "Entrada";
			ultraGridColumn6.Header.VisiblePosition = 3;
			ultraGridColumn6.Width = 100;
			ultraGridColumn7.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn7.Format = "dd/MM/yyy HH:mm:ss";
			ultraGridColumn7.Header.Caption = "Salida Almuerzo";
			ultraGridColumn7.Header.VisiblePosition = 4;
			ultraGridColumn7.Width = 100;
			ultraGridColumn8.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn8.Format = "dd/MM/yyy HH:mm:ss";
			ultraGridColumn8.Header.Caption = "Entrada Almuerzo";
			ultraGridColumn8.Header.VisiblePosition = 5;
			ultraGridColumn8.Width = 100;
			ultraGridColumn9.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn9.Format = "dd/MM/yyy HH:mm:ss";
			ultraGridColumn9.Header.Caption = "Salida";
			ultraGridColumn9.Header.VisiblePosition = 6;
			ultraGridColumn9.Width = 100;
			ultraGridColumn10.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn10.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn10.Header.VisiblePosition = 7;
			ultraGridColumn10.Width = 250;
			ultraGridColumn11.Header.VisiblePosition = 8;
			ultraGridColumn11.Hidden = true;
			ultraGridBand2.Columns.AddRange(new object[] {
																										 ultraGridColumn3,
																										 ultraGridColumn4,
																										 ultraGridColumn5,
																										 ultraGridColumn6,
																										 ultraGridColumn7,
																										 ultraGridColumn8,
																										 ultraGridColumn9,
																										 ultraGridColumn10,
																										 ultraGridColumn11});
			this.dgVerificacion.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			this.dgVerificacion.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
			this.dgVerificacion.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.dgVerificacion.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.False;
			appearance9.BackColor = System.Drawing.Color.Transparent;
			this.dgVerificacion.DisplayLayout.Override.CardAreaAppearance = appearance9;
			appearance10.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance10.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance10.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance10.FontData.BoldAsString = "True";
			appearance10.FontData.Name = "Arial";
			appearance10.FontData.SizeInPoints = 8F;
			appearance10.ForeColor = System.Drawing.Color.White;
			appearance10.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.dgVerificacion.DisplayLayout.Override.HeaderAppearance = appearance10;
			appearance11.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance11.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance11.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.dgVerificacion.DisplayLayout.Override.RowSelectorAppearance = appearance11;
			appearance12.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance12.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance12.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.dgVerificacion.DisplayLayout.Override.SelectedRowAppearance = appearance12;
			this.dgVerificacion.Font = new System.Drawing.Font("Tahoma", 7.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.dgVerificacion.Location = new System.Drawing.Point(0, 256);
			this.dgVerificacion.Name = "dgVerificacion";
			this.dgVerificacion.Size = new System.Drawing.Size(592, 136);
			this.dgVerificacion.TabIndex = 791;
			// 
			// ultraDataSource2
			// 
			ultraDataColumn3.DataType = typeof(int);
			ultraDataColumn4.DataType = typeof(System.DateTime);
			ultraDataColumn6.DataType = typeof(System.DateTime);
			ultraDataColumn7.DataType = typeof(System.DateTime);
			ultraDataColumn8.DataType = typeof(System.DateTime);
			ultraDataColumn9.DataType = typeof(System.DateTime);
			ultraDataColumn11.DataType = typeof(int);
			this.ultraDataSource2.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn3,
																																 ultraDataColumn4,
																																 ultraDataColumn5,
																																 ultraDataColumn6,
																																 ultraDataColumn7,
																																 ultraDataColumn8,
																																 ultraDataColumn9,
																																 ultraDataColumn10,
																																 ultraDataColumn11});
			// 
			// groupBox3
			// 
			this.groupBox3.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.groupBox3.Enabled = false;
			this.groupBox3.Location = new System.Drawing.Point(0, 240);
			this.groupBox3.Name = "groupBox3";
			this.groupBox3.Size = new System.Drawing.Size(592, 8);
			this.groupBox3.TabIndex = 797;
			this.groupBox3.TabStop = false;
			// 
			// txtNotas
			// 
			this.txtNotas.AcceptsReturn = true;
			this.txtNotas.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			appearance13.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtNotas.Appearance = appearance13;
			this.txtNotas.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtNotas.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNotas.Location = new System.Drawing.Point(64, 176);
			this.txtNotas.MaxLength = 500;
			this.txtNotas.Multiline = true;
			this.txtNotas.Name = "txtNotas";
			this.txtNotas.Size = new System.Drawing.Size(336, 64);
			this.txtNotas.TabIndex = 798;
			// 
			// label30
			// 
			this.label30.AutoSize = true;
			this.label30.Location = new System.Drawing.Point(8, 184);
			this.label30.Name = "label30";
			this.label30.Size = new System.Drawing.Size(28, 16);
			this.label30.TabIndex = 799;
			this.label30.Text = "Nota";
			// 
			// txtidBodega
			// 
			this.txtidBodega.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.txtidBodega.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtidBodega.Enabled = false;
			this.txtidBodega.Location = new System.Drawing.Point(472, 80);
			this.txtidBodega.Name = "txtidBodega";
			this.txtidBodega.PromptChar = ' ';
			this.txtidBodega.Size = new System.Drawing.Size(104, 21);
			this.txtidBodega.SpinButtonAlignment = Infragistics.Win.ButtonAlignment.Left;
			this.txtidBodega.TabIndex = 800;
			this.txtidBodega.Visible = false;
			// 
			// btnNuevo
			// 
			this.btnNuevo.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnNuevo.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(230)), ((System.Byte)(230)), ((System.Byte)(230)));
			this.btnNuevo.CausesValidation = false;
			this.btnNuevo.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.btnNuevo.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(50)), ((System.Byte)(50)), ((System.Byte)(50)));
			this.btnNuevo.Image = ((System.Drawing.Image)(resources.GetObject("btnNuevo.Image")));
			this.btnNuevo.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnNuevo.Location = new System.Drawing.Point(416, 208);
			this.btnNuevo.Name = "btnNuevo";
			this.btnNuevo.Size = new System.Drawing.Size(72, 23);
			this.btnNuevo.TabIndex = 801;
			this.btnNuevo.Text = "Nuevo";
			this.btnNuevo.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnNuevo.Visible = false;
			this.btnNuevo.Click += new System.EventHandler(this.btnNuevo_Click);
			// 
			// frmControlDeAsistencia
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(592, 390);
			this.Controls.Add(this.btnNuevo);
			this.Controls.Add(this.txtidBodega);
			this.Controls.Add(this.label30);
			this.Controls.Add(this.txtNotas);
			this.Controls.Add(this.groupBox3);
			this.Controls.Add(this.dgVerificacion);
			this.Controls.Add(this.label8);
			this.Controls.Add(this.cmbEstadoEntrada);
			this.Controls.Add(this.txtIdGrupo);
			this.Controls.Add(this.btnGuardar);
			this.Controls.Add(this.txtidNomina);
			this.Controls.Add(this.txtAlmcen);
			this.Controls.Add(this.label7);
			this.Controls.Add(this.label4);
			this.Controls.Add(this.txtCargo);
			this.Controls.Add(this.txtNombre);
			this.Controls.Add(this.txtCedula);
			this.Controls.Add(this.txtUsuario);
			this.Controls.Add(this.label3);
			this.Controls.Add(this.label2);
			this.Controls.Add(this.label1);
			this.Controls.Add(this.label6);
			this.Controls.Add(this.dtFecha);
			this.Controls.Add(this.pictureBox2);
			this.Controls.Add(this.label5);
			this.MaximizeBox = false;
			this.Name = "frmControlDeAsistencia";
			this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
			this.Text = "Control De Asistencia";
			this.Load += new System.EventHandler(this.frmControlDeAsistencia_Load);
			((System.ComponentModel.ISupportInitialize)(this.dtFecha)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtUsuario)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCedula)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNombre)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCargo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtAlmcen)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtidNomina)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdGrupo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbEstadoEntrada)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dgVerificacion)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNotas)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtidBodega)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void timer1_Tick(object sender, System.EventArgs e)
		{
				this.dtFecha.Value = DateTime.Now;
		}

		private void frmControlDeAsistencia_Load(object sender, System.EventArgs e)
		{
//			this.dtFecha.Value = DateTime.Now;
			   this.dtFecha.Value = Funcion.dtEscalarSQL(cdsSeteoF, string.Format(" Select  GETDATE()"));	
			string sUsuario = "";
			string sCedula  = "";
			int iGrupo = 0;
  
			sUsuario = Funcion.sEscalarSQL(cdsSeteoF, string.Format(" Select Nombre from Usuario where idUsuario = {0} ",MenuLatinium.IdUsuario));		
      iGrupo = Funcion.iEscalarSQL(cdsSeteoF, string.Format(" Select idGrupo from Usuario where idUsuario = {0} ",MenuLatinium.IdUsuario));	
     
			if ( iGrupo == 23 || iGrupo == 25)
			{
				this.txtUsuario.Text = sUsuario;
				int iBodega  = Funcion.iEscalarSQL(cdsSeteoF, string.Format(" Select  b.Bodega from UsuariosBodegas U Inner Join Bodega b on b.Bodega = u.Bodega where  b.Credito = 1 and u.idUsuario = {0} ",MenuLatinium.IdUsuario));	
			
				this.btnNuevo.Visible = true;
				this.dgVerificacion.DataSource = Funcion.dvProcedimiento(cdsSeteoF, string.Format("Exec ConsultaGeneralControlAsistencia 0,{0},{1}", iGrupo,iBodega));
       
			}
			else
			{
				 sCedula = Funcion.sEscalarSQL(cdsSeteoF, string.Format(" Select NIdentificacion from Nomina where codigo = '{0}' ",sUsuario));	
         Consulta (sUsuario, sCedula, iGrupo);
				this.txtCedula.Enabled = false;
			 
			}
			Estado ( );
			
		
		}

		private void txtCedula_Validated(object sender, System.EventArgs e)
		{
//			if (!Validacion.vbTexto(this.txtUsuario, 1, 60, "Ingrese el Usuario")) return;
//			if (!Validacion.vbTexto(this.txtCedula, 10, 11, "Ingrese la Cedula")) return;
       Consulta (this.txtUsuario.Text, this.txtCedula.Text, (int)this.txtIdGrupo.Value);
		
			
		}
		private void Estado ( )
		{
			string fecha = DateTime.Now.ToString("HH");
//			if ( Convert.ToInt32(fecha)>= 7 && Convert.ToInt32(fecha)<= 11)
//			{
//        this.cmbEstadoEntrada.DataSource = Funcion.dvProcedimiento(cdsSeteoF,"select idEstadoIngreso, Nombre from EstadoIngreso where idEstadoIngreso = 1");
//			}
      
			if ( Convert.ToInt32(fecha)>= 7 && Convert.ToInt32(fecha)< 15)
			{
				this.cmbEstadoEntrada.DataSource = Funcion.dvProcedimiento(cdsSeteoF,"select idEstadoIngreso, Nombre from EstadoIngreso where idEstadoIngreso in (1,2,3)");
			}

			if ( Convert.ToInt32(fecha)>= 15 && Convert.ToInt32(fecha)<= 24)
			{
				this.cmbEstadoEntrada.DataSource = Funcion.dvProcedimiento(cdsSeteoF,"select idEstadoIngreso, Nombre from EstadoIngreso where idEstadoIngreso in (4)");
			}
			
      
		}

		private void Consulta ( string codigo, string Cedula, int iGrupo )
		{
        	
			string ssqlDo= string.Format ("Exec ConsultaControlAsistencia '{0}','{1}'", codigo,Cedula); 
			SqlDataReader drDO = Funcion.rEscalarSQL(cdsSeteoF,ssqlDo,true);
			drDO.Read();
			if(drDO.HasRows)
			{
				this.txtidNomina.Value = drDO.GetInt32 (0);
				this.txtIdGrupo.Value = drDO.GetInt32  (6);
				if ( (int)this.txtIdGrupo.Value == 23 || (int)this.txtIdGrupo.Value == 25)
				{
					this.txtUsuario.Text = codigo;
					
				}
				else
				{
					this.txtUsuario.Text = drDO.GetString  (2);
				}
				this.txtAlmcen.Text = drDO.GetString  (1);
				this.txtNombre.Text = drDO.GetString  (3);
				this.txtCargo.Text = drDO.GetString  (5);
				this.txtCedula.Text = drDO.GetString  (4);
				this.txtidBodega.Value =  drDO.GetInt32  (7);
       
			}
			else
			{
//				MessageBox.Show("Comuniquese con el departamento de RRHH", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.txtCedula.Text = "";
				this.txtNombre.Text = "";
				this.txtCargo.Text = "";
				this.txtidNomina.Value = 0;
				this.txtAlmcen.Text = "";
				this.txtIdGrupo.Value = 0;
			}
			drDO.Close();	

			if ((int)this.txtidNomina.Value > 0)
			{
				this.dgVerificacion.DataSource = Funcion.dvProcedimiento(cdsSeteoF, string.Format("Exec ConsultaGeneralControlAsistencia {0}, {1}, {2}", 
					(int)this.txtidNomina.Value,(int)this.txtIdGrupo.Value, (int)this.txtidBodega.Value ));
			}
			else
			{
				int iBodega  = Funcion.iEscalarSQL(cdsSeteoF, string.Format(" Select  b.Bodega from UsuariosBodegas U Inner Join Bodega b on b.Bodega = u.Bodega where  b.Credito = 1 and u.idUsuario = {0} ",MenuLatinium.IdUsuario));	
			
//				this.dgVerificacion.DataSource = Funcion.dvProcedimiento(cdsSeteoF, string.Format("Exec ConsultaGeneralControlAsistencia 0,{0},{1}", iGrupo,iBodega ));
			}
		}

		private void txtCedula_Validating(object sender, System.ComponentModel.CancelEventArgs e)
		{
			if (!Validacion.vbIdentificacion(this.txtCedula, 5, cdsSeteoF))
			{
				  this.txtCedula.Text = "";
					this.txtNombre.Text = "";
					this.txtCargo.Text = "";
					this.txtidNomina.Value = 0;
				this.txtAlmcen.Text = "";
				this.txtIdGrupo.Value = 0;
					return;
			}
		}

		private void txtCedula_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter) this.btnGuardar.Focus();
		}

		private void btnGuardar_Click(object sender, System.EventArgs e)
		{
			int iCont = 0;
			if (!Validacion.vbTexto(this.txtCedula, 9, 13, "Cedula")) return;
			if (!Validacion.vbComboVacio(this.cmbEstadoEntrada, "Seleccione el Tipo de Entrada")) return;
      
			string sEstado = "";

			if ((int)this.cmbEstadoEntrada.Value == 1) sEstado ="ENTRADA";
			if ((int)this.cmbEstadoEntrada.Value == 2) sEstado ="SALIDA ALMUERZO";
			if ((int)this.cmbEstadoEntrada.Value == 3) sEstado ="ENTRADA ALMUERZO";
			if ((int)this.cmbEstadoEntrada.Value == 4) sEstado ="SALIDA";
       
			iCont = Funcion.iEscalarSQL(cdsSeteoF, string.Format("select COUNT(*) from ControlAsistencia where DATEPART(YEAR, fecha) = DATEPART(YEAR, GETDATE()) and DATEPART(MONTH, fecha) = DATEPART(MONTH, GETDATE()) and DATEPART(DAY, fecha) = DATEPART(DAY, GETDATE()) and idEstadoIngreso = {0} and idNomina = {1}", (int)this.cmbEstadoEntrada.Value, (int) this.txtidNomina.Value));	
			if(iCont > 0)
			{ 
        MessageBox.Show(string.Format("El tipo de Ingreso '{0}' ya lo a registrado anteriomente ",sEstado ), "Validacion", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return;
			}
			#region Guardar Maestro
			string sSQLl = string.Format("Exec GuardaControlAsistencia {0},'{1}', {2}, {3},'{4}'",
				(int) this.txtidNomina.Value,
				Convert.ToDateTime(this.dtFecha.Value).ToString("yyyyMMdd HH:mm:ss"),
				(int) this.txtidBodega.Value,
				(int) this.cmbEstadoEntrada.Value,
				this.txtNotas.Text.ToString());
			int idControl =Funcion.iEscalarSQL(cdsSeteoF,sSQLl);

			#endregion Guardar Maestro

		  MessageBox.Show("Registro Guardado Correctamente ", "Informacion", MessageBoxButtons.OK, MessageBoxIcon.Information);  
			this.dgVerificacion.DataSource = Funcion.dvProcedimiento(cdsSeteoF, string.Format("Exec ConsultaGeneralControlAsistencia {0}, {1}, {2}", 
				(int)this.txtidNomina.Value,(int)this.txtIdGrupo.Value, (int)this.txtidBodega.Value ));
//			this.Close ();
				
		}

		private void btnNuevo_Click(object sender, System.EventArgs e)
		{
			int iGrupo = Funcion.iEscalarSQL(cdsSeteoF, string.Format(" Select idGrupo from Usuario where idUsuario = {0} ",MenuLatinium.IdUsuario));	
     
			if (iGrupo == 23 || iGrupo == 25)
			{
				this.txtCedula.Text = "";
				this.txtCargo.Text ="";
				this.txtNombre.Text = "";
				this.txtNotas.Text = "";
				this.cmbEstadoEntrada.Value = 0;
			}
		}

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);		
		}
	}
}
