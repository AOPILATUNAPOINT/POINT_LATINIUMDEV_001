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
	/// Descripción breve de frmRolPrestamo.
	/// </summary>
	public class frmRolPrestamo : System.Windows.Forms.Form
	{
		private System.Windows.Forms.Label label7;
		private System.Windows.Forms.Label label8;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtNumero;
		private System.Windows.Forms.Label label6;
		private System.Windows.Forms.Label label5;
		private System.Windows.Forms.Label label4;
		private System.Windows.Forms.Label label3;
		private System.Windows.Forms.Label label1;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtNumCuotas;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtValor;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtConcepto;
		private System.Windows.Forms.Label label2;
		private System.Windows.Forms.Button btnNuevo;
		private System.Windows.Forms.Button btnConsultar;
		private System.Windows.Forms.Button btnEditar;
		private System.Windows.Forms.Button btnGuardar;
		private System.Windows.Forms.Button btnCancelar;
		private System.Windows.Forms.Button btnSalir;
		private System.Windows.Forms.GroupBox groupBox1;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbTipos;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource1;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtFecha;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtFechaCobro;
		private System.Windows.Forms.Button btnEgreso;
		private C1.Data.C1DataSet cdsSeteoF;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource2;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbNombre;
		private System.Windows.Forms.Button btnAnular;
		private System.Windows.Forms.Label lblEstado;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtIdRolPrestamo;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtIdAsiento;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtEstado;
		private System.Windows.Forms.Label lblUsuario;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public frmRolPrestamo()
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
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idRolTipoPrest");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("TipoPrestamo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idRolTipoPrest");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("TipoPrestamo");
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton1 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmRolPrestamo));
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton2 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn3 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idPersonal");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn4 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Personal", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPersonal");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			this.label7 = new System.Windows.Forms.Label();
			this.cmbTipos = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.ultraDataSource1 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.label8 = new System.Windows.Forms.Label();
			this.txtNumero = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label6 = new System.Windows.Forms.Label();
			this.label5 = new System.Windows.Forms.Label();
			this.label4 = new System.Windows.Forms.Label();
			this.label3 = new System.Windows.Forms.Label();
			this.label1 = new System.Windows.Forms.Label();
			this.txtNumCuotas = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.dtFechaCobro = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.dtFecha = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.txtValor = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtConcepto = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label2 = new System.Windows.Forms.Label();
			this.btnNuevo = new System.Windows.Forms.Button();
			this.btnConsultar = new System.Windows.Forms.Button();
			this.btnEditar = new System.Windows.Forms.Button();
			this.btnGuardar = new System.Windows.Forms.Button();
			this.btnCancelar = new System.Windows.Forms.Button();
			this.btnSalir = new System.Windows.Forms.Button();
			this.groupBox1 = new System.Windows.Forms.GroupBox();
			this.btnEgreso = new System.Windows.Forms.Button();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.ultraDataSource2 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.cmbNombre = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.btnAnular = new System.Windows.Forms.Button();
			this.lblEstado = new System.Windows.Forms.Label();
			this.txtIdRolPrestamo = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtIdAsiento = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtEstado = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.lblUsuario = new System.Windows.Forms.Label();
			((System.ComponentModel.ISupportInitialize)(this.cmbTipos)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNumero)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNumCuotas)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFechaCobro)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFecha)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtValor)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtConcepto)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbNombre)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdRolPrestamo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdAsiento)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtEstado)).BeginInit();
			this.SuspendLayout();
			// 
			// label7
			// 
			this.label7.AutoSize = true;
			this.label7.Location = new System.Drawing.Point(8, 10);
			this.label7.Name = "label7";
			this.label7.Size = new System.Drawing.Size(26, 17);
			this.label7.TabIndex = 37;
			this.label7.Text = "Tipo";
			this.label7.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// cmbTipos
			// 
			this.cmbTipos.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbTipos.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbTipos.DataSource = this.ultraDataSource1;
			this.cmbTipos.DisplayLayout.AutoFitColumns = true;
			ultraGridBand1.AddButtonCaption = "Personal";
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn1.Width = 148;
			ultraGridColumn2.Header.Caption = "Tipo";
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridColumn2.Width = 317;
			ultraGridBand1.Columns.AddRange(new object[] {
																										 ultraGridColumn1,
																										 ultraGridColumn2});
			this.cmbTipos.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.cmbTipos.DisplayMember = "TipoPrestamo";
			this.cmbTipos.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbTipos.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbTipos.Enabled = false;
			this.cmbTipos.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbTipos.Location = new System.Drawing.Point(136, 8);
			this.cmbTipos.Name = "cmbTipos";
			this.cmbTipos.Size = new System.Drawing.Size(336, 21);
			this.cmbTipos.TabIndex = 38;
			this.cmbTipos.ValueMember = "idRolTipoPrest";
			this.cmbTipos.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cmbTipos_KeyDown);			
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
			this.label8.Location = new System.Drawing.Point(496, 10);
			this.label8.Name = "label8";
			this.label8.Size = new System.Drawing.Size(44, 17);
			this.label8.TabIndex = 42;
			this.label8.Text = "Número";
			this.label8.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtNumero
			// 
			this.txtNumero.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtNumero.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNumero.Enabled = false;
			this.txtNumero.Location = new System.Drawing.Point(560, 7);
			this.txtNumero.Name = "txtNumero";
			this.txtNumero.Size = new System.Drawing.Size(96, 22);
			this.txtNumero.TabIndex = 41;
			// 
			// label6
			// 
			this.label6.AutoSize = true;
			this.label6.Location = new System.Drawing.Point(264, 107);
			this.label6.Name = "label6";
			this.label6.Size = new System.Drawing.Size(69, 17);
			this.label6.TabIndex = 52;
			this.label6.Text = "N. de Cuotas";
			this.label6.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label5
			// 
			this.label5.AutoSize = true;
			this.label5.Location = new System.Drawing.Point(264, 42);
			this.label5.Name = "label5";
			this.label5.Size = new System.Drawing.Size(82, 17);
			this.label5.TabIndex = 50;
			this.label5.Text = "Fecha de Cobro";
			this.label5.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label4
			// 
			this.label4.AutoSize = true;
			this.label4.Location = new System.Drawing.Point(8, 42);
			this.label4.Name = "label4";
			this.label4.Size = new System.Drawing.Size(95, 17);
			this.label4.TabIndex = 48;
			this.label4.Text = "Fecha de Solicitud";
			this.label4.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label3
			// 
			this.label3.AutoSize = true;
			this.label3.Location = new System.Drawing.Point(8, 107);
			this.label3.Name = "label3";
			this.label3.Size = new System.Drawing.Size(30, 17);
			this.label3.TabIndex = 46;
			this.label3.Text = "Valor";
			this.label3.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label1
			// 
			this.label1.AutoSize = true;
			this.label1.Location = new System.Drawing.Point(8, 170);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(51, 17);
			this.label1.TabIndex = 43;
			this.label1.Text = "Concepto";
			this.label1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtNumCuotas
			// 
			this.txtNumCuotas.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNumCuotas.Enabled = false;
			this.txtNumCuotas.FormatString = "";
			this.txtNumCuotas.Location = new System.Drawing.Point(360, 104);
			this.txtNumCuotas.MaskInput = "nnn";
			this.txtNumCuotas.Name = "txtNumCuotas";
			this.txtNumCuotas.PromptChar = ' ';
			this.txtNumCuotas.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
			this.txtNumCuotas.Size = new System.Drawing.Size(112, 22);
			this.txtNumCuotas.SpinButtonDisplayStyle = Infragistics.Win.ButtonDisplayStyle.Always;
			this.txtNumCuotas.TabIndex = 51;
			this.txtNumCuotas.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtNumCuotas_KeyDown);
			this.txtNumCuotas.Click += new System.EventHandler(this.txtNumCuotas_Click);
			this.txtNumCuotas.Validated += new System.EventHandler(this.txtNumCuotas_Validated);
			this.txtNumCuotas.Enter += new System.EventHandler(this.txtNumCuotas_Enter);
			// 
			// dtFechaCobro
			// 
			this.dtFechaCobro.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton1.Caption = "Today";
			this.dtFechaCobro.DateButtons.Add(dateButton1);
			this.dtFechaCobro.Enabled = false;
			this.dtFechaCobro.Format = "dd/MM/yyyy";
			this.dtFechaCobro.Location = new System.Drawing.Point(360, 40);
			this.dtFechaCobro.Name = "dtFechaCobro";
			this.dtFechaCobro.NonAutoSizeHeight = 23;
			this.dtFechaCobro.Size = new System.Drawing.Size(112, 21);
			this.dtFechaCobro.SpinButtonsVisible = true;
			this.dtFechaCobro.TabIndex = 49;
			this.dtFechaCobro.Value = ((object)(resources.GetObject("dtFechaCobro.Value")));
			this.dtFechaCobro.KeyDown += new System.Windows.Forms.KeyEventHandler(this.dtFechaCobro_KeyDown);
			this.dtFechaCobro.Validated += new System.EventHandler(this.dtFechaCobro_Validated);
			// 
			// dtFecha
			// 
			this.dtFecha.AllowMonthSelection = true;
			this.dtFecha.AutoSelectionUpdate = true;
			this.dtFecha.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton2.Caption = "Today";
			this.dtFecha.DateButtons.Add(dateButton2);
			this.dtFecha.DayDisplayStyle = Infragistics.Win.UltraWinSchedule.DayDisplayStyle.DayNumberAndImage;
			this.dtFecha.DayOfWeekDisplayStyle = Infragistics.Win.UltraWinSchedule.DayOfWeekDisplayStyle.FirstRow;
			this.dtFecha.Enabled = false;
			this.dtFecha.Format = "dd/MM/yyyy";
			this.dtFecha.Location = new System.Drawing.Point(136, 40);
			this.dtFecha.MonthOrientation = Infragistics.Win.UltraWinSchedule.MonthOrientation.DownThenAcross;
			this.dtFecha.Name = "dtFecha";
			this.dtFecha.NonAutoSizeHeight = 23;
			this.dtFecha.Size = new System.Drawing.Size(112, 21);
			this.dtFecha.SpinButtonsVisible = true;
			this.dtFecha.TabIndex = 47;
			this.dtFecha.Value = ((object)(resources.GetObject("dtFecha.Value")));
			this.dtFecha.KeyDown += new System.Windows.Forms.KeyEventHandler(this.dtFecha_KeyDown);
			this.dtFecha.Validated += new System.EventHandler(this.dtFecha_Validated);			
			this.dtFecha.ValueChanged += new System.EventHandler(this.dtFecha_ValueChanged);
			// 
			// txtValor
			// 
			this.txtValor.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtValor.Enabled = false;
			this.txtValor.FormatString = "#,##0.00";
			this.txtValor.Location = new System.Drawing.Point(136, 104);
			this.txtValor.Name = "txtValor";
			this.txtValor.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtValor.PromptChar = ' ';
			this.txtValor.Size = new System.Drawing.Size(112, 22);
			this.txtValor.TabIndex = 45;
			this.txtValor.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtValor_KeyDown);
			this.txtValor.Click += new System.EventHandler(this.txtValor_Click);
			this.txtValor.Validated += new System.EventHandler(this.txtValor_Validated);
			this.txtValor.Enter += new System.EventHandler(this.txtValor_Enter);
			// 
			// txtConcepto
			// 
			this.txtConcepto.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtConcepto.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtConcepto.Enabled = false;
			this.txtConcepto.Location = new System.Drawing.Point(136, 136);
			this.txtConcepto.Multiline = true;
			this.txtConcepto.Name = "txtConcepto";
			this.txtConcepto.Size = new System.Drawing.Size(336, 84);
			this.txtConcepto.TabIndex = 44;
			this.txtConcepto.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtConcepto_KeyDown);
			// 
			// label2
			// 
			this.label2.AutoSize = true;
			this.label2.Location = new System.Drawing.Point(8, 74);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(106, 17);
			this.label2.TabIndex = 53;
			this.label2.Text = "Apellidos y Nombres";
			this.label2.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// btnNuevo
			// 
			this.btnNuevo.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnNuevo.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(230)), ((System.Byte)(230)), ((System.Byte)(230)));
			this.btnNuevo.CausesValidation = false;
			this.btnNuevo.Enabled = false;
			this.btnNuevo.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.btnNuevo.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(50)), ((System.Byte)(50)), ((System.Byte)(50)));
			this.btnNuevo.Image = ((System.Drawing.Image)(resources.GetObject("btnNuevo.Image")));
			this.btnNuevo.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnNuevo.Location = new System.Drawing.Point(8, 248);
			this.btnNuevo.Name = "btnNuevo";
			this.btnNuevo.Size = new System.Drawing.Size(72, 23);
			this.btnNuevo.TabIndex = 279;
			this.btnNuevo.Text = "Nuevo";
			this.btnNuevo.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnNuevo.Click += new System.EventHandler(this.btnNuevo_Click);
			// 
			// btnConsultar
			// 
			this.btnConsultar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnConsultar.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(230)), ((System.Byte)(230)), ((System.Byte)(230)));
			this.btnConsultar.CausesValidation = false;
			this.btnConsultar.Enabled = false;
			this.btnConsultar.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.btnConsultar.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(50)), ((System.Byte)(50)), ((System.Byte)(50)));
			this.btnConsultar.Image = ((System.Drawing.Image)(resources.GetObject("btnConsultar.Image")));
			this.btnConsultar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnConsultar.Location = new System.Drawing.Point(85, 248);
			this.btnConsultar.Name = "btnConsultar";
			this.btnConsultar.Size = new System.Drawing.Size(78, 23);
			this.btnConsultar.TabIndex = 280;
			this.btnConsultar.Text = "Consultar";
			this.btnConsultar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnConsultar.Click += new System.EventHandler(this.btnConsultar_Click);
			// 
			// btnEditar
			// 
			this.btnEditar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnEditar.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(230)), ((System.Byte)(230)), ((System.Byte)(230)));
			this.btnEditar.CausesValidation = false;
			this.btnEditar.Enabled = false;
			this.btnEditar.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.btnEditar.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(50)), ((System.Byte)(50)), ((System.Byte)(50)));
			this.btnEditar.Image = ((System.Drawing.Image)(resources.GetObject("btnEditar.Image")));
			this.btnEditar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnEditar.Location = new System.Drawing.Point(168, 248);
			this.btnEditar.Name = "btnEditar";
			this.btnEditar.Size = new System.Drawing.Size(76, 23);
			this.btnEditar.TabIndex = 282;
			this.btnEditar.Text = "Editar";
			this.btnEditar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnEditar.Click += new System.EventHandler(this.btnEditar_Click);
			// 
			// btnGuardar
			// 
			this.btnGuardar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnGuardar.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(230)), ((System.Byte)(230)), ((System.Byte)(230)));
			this.btnGuardar.Enabled = false;
			this.btnGuardar.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.btnGuardar.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(50)), ((System.Byte)(50)), ((System.Byte)(50)));
			this.btnGuardar.Image = ((System.Drawing.Image)(resources.GetObject("btnGuardar.Image")));
			this.btnGuardar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnGuardar.Location = new System.Drawing.Point(249, 248);
			this.btnGuardar.Name = "btnGuardar";
			this.btnGuardar.Size = new System.Drawing.Size(76, 23);
			this.btnGuardar.TabIndex = 281;
			this.btnGuardar.Text = "Guardar";
			this.btnGuardar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnGuardar.Click += new System.EventHandler(this.btnGuardar_Click);
			// 
			// btnCancelar
			// 
			this.btnCancelar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnCancelar.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(230)), ((System.Byte)(230)), ((System.Byte)(230)));
			this.btnCancelar.CausesValidation = false;
			this.btnCancelar.Enabled = false;
			this.btnCancelar.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.btnCancelar.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(50)), ((System.Byte)(50)), ((System.Byte)(50)));
			this.btnCancelar.Image = ((System.Drawing.Image)(resources.GetObject("btnCancelar.Image")));
			this.btnCancelar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnCancelar.Location = new System.Drawing.Point(491, 248);
			this.btnCancelar.Name = "btnCancelar";
			this.btnCancelar.Size = new System.Drawing.Size(72, 23);
			this.btnCancelar.TabIndex = 283;
			this.btnCancelar.Text = "Cancelar";
			this.btnCancelar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnCancelar.Click += new System.EventHandler(this.btnCancelar_Click);
			// 
			// btnSalir
			// 
			this.btnSalir.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnSalir.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(230)), ((System.Byte)(230)), ((System.Byte)(230)));
			this.btnSalir.CausesValidation = false;
			this.btnSalir.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.btnSalir.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(50)), ((System.Byte)(50)), ((System.Byte)(50)));
			this.btnSalir.Image = ((System.Drawing.Image)(resources.GetObject("btnSalir.Image")));
			this.btnSalir.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnSalir.Location = new System.Drawing.Point(568, 248);
			this.btnSalir.Name = "btnSalir";
			this.btnSalir.TabIndex = 284;
			this.btnSalir.Text = "Salir";
			this.btnSalir.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnSalir.Click += new System.EventHandler(this.btnSalir_Click);
			// 
			// groupBox1
			// 
			this.groupBox1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.groupBox1.Location = new System.Drawing.Point(0, 232);
			this.groupBox1.Name = "groupBox1";
			this.groupBox1.Size = new System.Drawing.Size(666, 8);
			this.groupBox1.TabIndex = 357;
			this.groupBox1.TabStop = false;
			// 
			// btnEgreso
			// 
			this.btnEgreso.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnEgreso.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(230)), ((System.Byte)(230)), ((System.Byte)(230)));
			this.btnEgreso.CausesValidation = false;
			this.btnEgreso.Enabled = false;
			this.btnEgreso.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.btnEgreso.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(50)), ((System.Byte)(50)), ((System.Byte)(50)));
			this.btnEgreso.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnEgreso.ImageIndex = 62;
			this.btnEgreso.Location = new System.Drawing.Point(330, 248);
			this.btnEgreso.Name = "btnEgreso";
			this.btnEgreso.TabIndex = 362;
			this.btnEgreso.Text = "Ver Egreso";
			this.btnEgreso.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnEgreso.Click += new System.EventHandler(this.btnEgreso_Click);
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
			// ultraDataSource2
			// 
			ultraDataColumn3.DataType = typeof(int);
			this.ultraDataSource2.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn3,
																																 ultraDataColumn4});
			this.ultraDataSource2.Band.Key = "Personal";
			// 
			// cmbNombre
			// 
			this.cmbNombre.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbNombre.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbNombre.DataSource = this.ultraDataSource2;
			this.cmbNombre.DisplayLayout.AutoFitColumns = true;
			ultraGridColumn3.Header.VisiblePosition = 1;
			ultraGridColumn3.Hidden = true;
			ultraGridColumn4.Header.VisiblePosition = 0;
			ultraGridColumn4.Width = 317;
			ultraGridBand2.Columns.AddRange(new object[] {
																										 ultraGridColumn3,
																										 ultraGridColumn4});
			this.cmbNombre.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			this.cmbNombre.DisplayMember = "Nombre";
			this.cmbNombre.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbNombre.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbNombre.Enabled = false;
			this.cmbNombre.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbNombre.Location = new System.Drawing.Point(136, 72);
			this.cmbNombre.Name = "cmbNombre";
			this.cmbNombre.Size = new System.Drawing.Size(336, 21);
			this.cmbNombre.TabIndex = 363;
			this.cmbNombre.ValueMember = "idPersonal";
			this.cmbNombre.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cmbNombre_KeyDown);
			// 
			// btnAnular
			// 
			this.btnAnular.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnAnular.CausesValidation = false;
			this.btnAnular.Enabled = false;
			this.btnAnular.Image = ((System.Drawing.Image)(resources.GetObject("btnAnular.Image")));
			this.btnAnular.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnAnular.Location = new System.Drawing.Point(410, 248);
			this.btnAnular.Name = "btnAnular";
			this.btnAnular.Size = new System.Drawing.Size(76, 23);
			this.btnAnular.TabIndex = 364;
			this.btnAnular.Text = "&Anular";
			this.btnAnular.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnAnular.Click += new System.EventHandler(this.btnAnular_Click);
			// 
			// lblEstado
			// 
			this.lblEstado.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.lblEstado.AutoSize = true;
			this.lblEstado.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblEstado.ForeColor = System.Drawing.Color.Firebrick;
			this.lblEstado.Location = new System.Drawing.Point(488, 41);
			this.lblEstado.Name = "lblEstado";
			this.lblEstado.Size = new System.Drawing.Size(0, 19);
			this.lblEstado.TabIndex = 365;
			// 
			// txtIdRolPrestamo
			// 
			this.txtIdRolPrestamo.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtIdRolPrestamo.Enabled = false;
			this.txtIdRolPrestamo.FormatString = "";
			this.txtIdRolPrestamo.Location = new System.Drawing.Point(496, 144);
			this.txtIdRolPrestamo.Name = "txtIdRolPrestamo";
			this.txtIdRolPrestamo.PromptChar = ' ';
			this.txtIdRolPrestamo.Size = new System.Drawing.Size(10, 22);
			this.txtIdRolPrestamo.TabIndex = 366;
			this.txtIdRolPrestamo.Visible = false;
			// 
			// txtIdAsiento
			// 
			this.txtIdAsiento.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtIdAsiento.Enabled = false;
			this.txtIdAsiento.FormatString = "";
			this.txtIdAsiento.Location = new System.Drawing.Point(496, 168);
			this.txtIdAsiento.Name = "txtIdAsiento";
			this.txtIdAsiento.PromptChar = ' ';
			this.txtIdAsiento.Size = new System.Drawing.Size(10, 22);
			this.txtIdAsiento.TabIndex = 367;
			this.txtIdAsiento.Visible = false;
			// 
			// txtEstado
			// 
			this.txtEstado.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtEstado.Enabled = false;
			this.txtEstado.FormatString = "";
			this.txtEstado.Location = new System.Drawing.Point(496, 192);
			this.txtEstado.Name = "txtEstado";
			this.txtEstado.PromptChar = ' ';
			this.txtEstado.Size = new System.Drawing.Size(10, 22);
			this.txtEstado.TabIndex = 368;
			this.txtEstado.Visible = false;
			// 
			// lblUsuario
			// 
			this.lblUsuario.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.lblUsuario.AutoSize = true;
			this.lblUsuario.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblUsuario.ForeColor = System.Drawing.Color.Firebrick;
			this.lblUsuario.Location = new System.Drawing.Point(488, 106);
			this.lblUsuario.Name = "lblUsuario";
			this.lblUsuario.Size = new System.Drawing.Size(0, 19);
			this.lblUsuario.TabIndex = 369;
			// 
			// frmRolPrestamo
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 14);
			this.ClientSize = new System.Drawing.Size(666, 278);
			this.Controls.Add(this.lblUsuario);
			this.Controls.Add(this.txtEstado);
			this.Controls.Add(this.txtIdAsiento);
			this.Controls.Add(this.txtIdRolPrestamo);
			this.Controls.Add(this.lblEstado);
			this.Controls.Add(this.btnAnular);
			this.Controls.Add(this.btnEgreso);
			this.Controls.Add(this.groupBox1);
			this.Controls.Add(this.btnNuevo);
			this.Controls.Add(this.btnConsultar);
			this.Controls.Add(this.btnEditar);
			this.Controls.Add(this.btnGuardar);
			this.Controls.Add(this.btnCancelar);
			this.Controls.Add(this.btnSalir);
			this.Controls.Add(this.label2);
			this.Controls.Add(this.label6);
			this.Controls.Add(this.label5);
			this.Controls.Add(this.label4);
			this.Controls.Add(this.label3);
			this.Controls.Add(this.label1);
			this.Controls.Add(this.label8);
			this.Controls.Add(this.label7);
			this.Controls.Add(this.txtNumCuotas);
			this.Controls.Add(this.dtFechaCobro);
			this.Controls.Add(this.dtFecha);
			this.Controls.Add(this.txtValor);
			this.Controls.Add(this.txtConcepto);
			this.Controls.Add(this.txtNumero);
			this.Controls.Add(this.cmbTipos);
			this.Controls.Add(this.cmbNombre);
			this.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
			this.KeyPreview = true;
			this.MaximizeBox = false;
			this.Name = "frmRolPrestamo";
			this.Text = "Anticipos Empleados";
			this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.frmRolPrestamo_KeyDown);
			this.Closing += new System.ComponentModel.CancelEventHandler(this.frmRolPrestamo_Closing);
			this.Load += new System.EventHandler(this.frmRolPrestamo_Load);
			((System.ComponentModel.ISupportInitialize)(this.cmbTipos)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNumero)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNumCuotas)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFechaCobro)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFecha)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtValor)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtConcepto)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbNombre)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdRolPrestamo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdAsiento)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtEstado)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion
		
		bool bNuevo = false;
		bool bEdicion = false;
		private Acceso miAcceso;
		int idBloqueaTransacciones = 0;

		private void UnloadMe()
		{
			this.Close();
		}

		private void frmRolPrestamo_Load(object sender, System.EventArgs e)
		{
			miAcceso = new Acceso(cdsSeteoF, "0504");
				
			if (!Funcion.Permiso("82", cdsSeteoF))
			{				
				MessageBox.Show("No puede ingresar a Anticipos Empleados", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				BeginInvoke(new MethodInvoker(UnloadMe));				
				return;
			}
			
			if (miAcceso.Nuevo) this.btnNuevo.Enabled = true;
			if (miAcceso.Buscar) this.btnConsultar.Enabled = true;

			this.dtFecha.CalendarInfo.MaxDate = DateTime.Today;
			this.dtFechaCobro.CalendarInfo.MaxDate = DateTime.Today.AddDays(10);
		}

		private void btnNuevo_Click(object sender, System.EventArgs e)
		{
			this.btnCancelar_Click(sender, e);

			this.cmbTipos.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Exec ListaTiposAnticiposEmpleados 1");

			this.cmbTipos.Enabled = true;
			this.cmbNombre.Enabled = true;
			if (miAcceso.CambiarFecha) this.dtFecha.Enabled = true;			
			this.dtFechaCobro.Enabled = true;
			this.txtValor.Enabled = true;
			this.txtNumCuotas.Enabled = true;
			this.txtConcepto.Enabled = true;

			this.btnNuevo.Enabled = false;
			this.btnConsultar.Enabled = false;
			this.btnEditar.Enabled = false;
			this.btnGuardar.Enabled = true;
			this.btnEgreso.Enabled = false;
			this.btnAnular.Enabled = false;
			this.btnCancelar.Enabled = true;

			this.cmbTipos.Value = 1;
			
			this.dtFecha.CalendarInfo.MinDate = DateTime.Today.AddDays(-15);			
			this.dtFecha.Value = DateTime.Today;

			this.dtFechaCobro.CalendarInfo.MinDate = DateTime.Today.AddDays(-15);
			this.dtFechaCobro.CalendarInfo.MaxDate = DateTime.Today.AddDays(10);
			this.dtFechaCobro.Value = DateTime.Today;

			this.txtNumCuotas.Value = 1;
			this.txtNumCuotas.MinValue = 1;
			this.txtNumCuotas.MaxValue = 3;

			bNuevo = true;
			bEdicion = true;

			this.cmbTipos.Focus();
		}

		private void btnCancelar_Click(object sender, System.EventArgs e)
		{
			if (!bNuevo && bEdicion) 
				FuncionesProcedimientos.RegistraBloqueoTransacciones (idBloqueaTransacciones, false, (int)this.txtIdRolPrestamo.Value, 10, cdsSeteoF);

			this.txtIdRolPrestamo.Value = 0;
			this.txtIdAsiento.Value = 0;
			this.txtEstado.Value = 0;
			
			bEdicion = false;
			bNuevo = false;

			this.cmbTipos.Value = System.DBNull.Value;
			this.txtNumero.Text = "";			
			this.dtFecha.Value = System.DBNull.Value;
			this.dtFechaCobro.Value = System.DBNull.Value;

			this.dtFecha.CalendarInfo.MinDate = DateTime.Parse("01/01/2015");			
			this.dtFechaCobro.CalendarInfo.MinDate = DateTime.Parse("01/01/2015");
			
			this.cmbNombre.Value = System.DBNull.Value;			
			this.txtValor.Value = 0.00m;
			this.txtNumCuotas.MinValue = 0;
			this.txtNumCuotas.Value = 0;
			this.txtConcepto.Text = "";
			this.lblEstado.Text = "";
			this.lblUsuario.Text = "";

			this.cmbTipos.Enabled = false;
			this.cmbNombre.Enabled = false;
			this.dtFecha.Enabled = false;
			this.dtFechaCobro.Enabled = false;
			this.txtValor.Enabled = false;
			this.txtNumCuotas.Enabled = false;
			this.txtConcepto.Enabled = false;

			this.btnNuevo.Enabled = true;
			this.btnConsultar.Enabled = true;
			this.btnEditar.Enabled = false;
			this.btnGuardar.Enabled = false;
			this.btnEgreso.Enabled = false;
			this.btnAnular.Enabled = false;
			this.btnCancelar.Enabled = false;

			idBloqueaTransacciones = 0;
		}

		private void btnSalir_Click(object sender, System.EventArgs e)
		{
			this.Close();
		}

		private void Consultar(int IdRolPrestamo)
		{
			try
			{
				this.cmbTipos.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Exec ListaTiposAnticiposEmpleados 1");

				this.btnNuevo.Enabled = false;
				this.btnConsultar.Enabled = false;
				this.btnEditar.Enabled = false;
				this.btnGuardar.Enabled = false;
				this.btnAnular.Enabled = false;
				this.btnEgreso.Enabled = false;
				this.btnCancelar.Enabled = false;

				string sSQL = string.Format("Exec ConsultaIndividualFiniquitos {0}", IdRolPrestamo);
				SqlDataReader dr = Funcion.rEscalarSQL(cdsSeteoF, sSQL, true);
				dr.Read();
				if (dr.HasRows)
				{
					this.txtIdRolPrestamo.Value = IdRolPrestamo;
					this.txtIdAsiento.Value = dr.GetInt32(1);
					this.cmbTipos.Value = dr.GetInt32(2);
					this.txtNumero.Text = dr.GetValue(3).ToString();
					this.cmbNombre.Value = dr.GetInt32(4);
					this.dtFecha.Value = dr.GetDateTime(5);
					this.txtValor.Value = dr.GetDecimal(6);
					this.txtConcepto.Text = dr.GetValue(7).ToString();
					this.txtEstado.Value = dr.GetInt32(8);
					this.lblEstado.Text = dr.GetValue(9).ToString();
					this.dtFechaCobro.Value = dr.GetDateTime(10);
					this.txtNumCuotas.Value = dr.GetInt32(11);
					this.txtIdAsiento.Value = dr.GetInt32(12);
					this.lblUsuario.Text = dr.GetValue(13).ToString();
				}
				dr.Close();
							
				if (miAcceso.Nuevo) this.btnNuevo.Enabled = true;
				if (miAcceso.Buscar) this.btnConsultar.Enabled = true;
				
				if ((int)this.txtEstado.Value == 0)
				{
					if (miAcceso.Editar) this.btnEditar.Enabled = true;
					if (miAcceso.Anular) this.btnAnular.Enabled = true;
				}

				if ((int)this.txtEstado.Value == 2)
					if (miAcceso.CrearAsiento)
						this.btnEgreso.Enabled = true;

				this.btnCancelar.Enabled = true;
			}
			catch(Exception Exc)
			{
				MessageBox.Show(string.Format("Error al Consultar: {0}", Exc), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
			}
		}

		private void btnConsultar_Click(object sender, System.EventArgs e)
		{
			this.btnCancelar_Click(sender, e);

			using (frmBuscaPrestamos Busqueda = new frmBuscaPrestamos(1))
			{
				if (DialogResult.OK == Busqueda.ShowDialog())
					this.Consultar((int)Busqueda.ultraGrid1.ActiveRow.Cells["idRolPrestamo"].Value);
			}		
		}

		private void btnEditar_Click(object sender, System.EventArgs e)
		{
			this.Consultar((int)this.txtIdRolPrestamo.Value);

			if ((int)this.txtEstado.Value == 3)
			{
				MessageBox.Show("El registro esta Anulado", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return;
			}

			if (!miAcceso.Editar)
			{
				MessageBox.Show("No tiene acceso a Editar", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return;
			}

			if ((int)this.txtEstado.Value == 1)
			{
				MessageBox.Show("No puede Editar ya esta Autorizado el Pago", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return;
			}

			if ((int)this.txtEstado.Value == 2)
			{
				MessageBox.Show("No puede Editar ya esta Pagado", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return;
			}
			
			if (Funcion.ValidaBloqueo((int)this.txtIdRolPrestamo.Value, 10, cdsSeteoF)) return;

			FuncionesProcedimientos.RegistraBloqueoTransacciones (idBloqueaTransacciones, true, (int)this.txtIdRolPrestamo.Value, 10, cdsSeteoF);

			this.dtFecha.CalendarInfo.MinDate = Convert.ToDateTime(this.dtFecha.Value).AddDays(-15);			
			this.dtFechaCobro.CalendarInfo.MinDate = Convert.ToDateTime(this.dtFecha.Value).AddDays(-15);
						
			if (miAcceso.CambiarFecha) this.dtFecha.Enabled = true;
      
			this.dtFechaCobro.Enabled = true;
			this.txtValor.Enabled = true;
			this.txtNumCuotas.Enabled = true;
			this.txtConcepto.Enabled = true;

			this.btnNuevo.Enabled = false;
			this.btnConsultar.Enabled = false;
			this.btnEditar.Enabled = false;
			this.btnGuardar.Enabled = true;
			this.btnEgreso.Enabled = false;
			this.btnAnular.Enabled = false;
			this.btnCancelar.Enabled = true;
		}

		private void btnGuardar_Click(object sender, System.EventArgs e)
		{
			if (!Validacion.vbComboVacio(this.cmbTipos, "Seleccione un Tipo")) return;
						
			if (!Validacion.vbFechaEnDocumentos(this.dtFecha, "Seleccione Fecha de solicitud", true, "Roles", cdsSeteoF)) return;

			if (!Validacion.vbFechaEnDocumentos(this.dtFechaCobro, "Seleccione la fecha del primer cobro", false, "Roles", cdsSeteoF)) return;

			if (!Validacion.vbComparaFechas((DateTime)this.dtFecha.Value, (DateTime)this.dtFechaCobro.Value, this.dtFecha, "La fecha de cobro no puede ser anterior a la fecha de emision", false)) return;

			if  (!Validacion.vbComboVacio(this.cmbNombre, "Seleccione el nombre del empleado")) return;

			if (!FuncionesContabilidad.bValidaAnticipoDuplicado(this.cmbTipos, this.cmbNombre, (int)this.txtIdRolPrestamo.Value, cdsSeteoF)) return;

			if (!Validacion.vbCampoDecimalVacio(this.txtValor, "Ingrese el valor", 0.00m, 100000.00m)) return;
			
			if (!Validacion.vbCampoEnteroVacio(this.txtNumCuotas, "Ingrese el numero de cuotas", 0, 3)) return;

			if (!Validacion.vbTexto(this.txtConcepto, 10, 300, "Concepto")) return;
			
			if (bNuevo) this.txtNumero.Value = Funcion.sEscalarSQL(cdsSeteoF, string.Format("Select dbo.SecuencialPrestamoEmpleados({0})", (int)this.cmbTipos.Value));
	
			string sSQL = string.Format("Exec GuardaRolPrestamoAnticiposPrestamos {0}, {1}, '{2}', 2, '{3}', '{4}', {5}, {6}, {7}, '{8}', {9}",
				(int)this.txtIdRolPrestamo.Value, (int)this.cmbTipos.Value, this.txtNumero.Text.ToString(), 
				Convert.ToDateTime(this.dtFecha.Value).ToString("yyyyMMdd"), Convert.ToDateTime(this.dtFechaCobro.Value).ToString("yyyyMMdd"), 
				(int)this.cmbNombre.Value, Convert.ToDecimal(this.txtValor.Value), (int)this.txtNumCuotas.Value, this.txtConcepto.Text.ToString(), (int)this.txtEstado.Value);
			this.txtIdRolPrestamo.Value = Funcion.iEscalarSQL(cdsSeteoF, sSQL);

			sSQL = string.Format("Exec CreaAutorizacionesDePago {0}, {1}, '{2}', '{3}', {4}, 2, '1 DE 1', '{5}', 0", 
				(int)this.cmbTipos.Value, (int)this.txtIdRolPrestamo.Value, 
				Convert.ToDateTime(this.dtFecha.Value).ToString("yyyyMMdd"), Convert.ToDateTime(this.dtFecha.Value).ToString("yyyyMMdd"), 
				Convert.ToDecimal(this.txtValor.Value), this.txtConcepto.Text.ToString());			
			Funcion.EjecutaSQL(cdsSeteoF, sSQL, true);

			FuncionesProcedimientos.RegistraBloqueoTransacciones (idBloqueaTransacciones, false, (int)this.txtIdRolPrestamo.Value, 10, cdsSeteoF);

			this.cmbTipos.Enabled = false;
			this.cmbNombre.Enabled = false;
			this.dtFecha.Enabled = false;
			this.dtFechaCobro.Enabled = false;
			this.txtValor.Enabled = false;
			this.txtNumCuotas.Enabled = false;
			this.txtConcepto.Enabled = false;

			this.btnNuevo.Enabled = true;
			this.btnConsultar.Enabled = true;
			this.btnEditar.Enabled = true;
			this.btnGuardar.Enabled = false;
			this.btnEgreso.Enabled = false;
			this.btnAnular.Enabled = true;
			this.btnCancelar.Enabled = true;

			bEdicion = false;
			bNuevo = false;
		}

		private void btnEgreso_Click(object sender, System.EventArgs e)
		{
			this.Consultar((int)this.txtIdRolPrestamo.Value);

			if ((int)this.txtEstado.Value == 3)
			{
				MessageBox.Show("El registro esta anulado", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return;
			}
			
			if (!miAcceso.CrearAsiento)
			{
				MessageBox.Show("No tiene acceso a ver el egreso", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return;
			}

			if ((int)this.txtEstado.Value < 2)
			{
				MessageBox.Show("No se ha generado el egreso de este registro", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return;
			}
			
			if ((int)this.txtIdAsiento.Value == 0)
			{
				MessageBox.Show("No existe un egreso creado de esta transacción");				
				return;
			}
		
			Cursor = Cursors.WaitCursor;
			frmAsientos Asiento = new frmAsientos((int)this.txtIdAsiento.Value);
			Asiento.MdiParent = MdiParent;
			Asiento.Show();
			Cursor = Cursors.Default;			
		}

		private void cmbTipos_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.dtFecha.Focus();
		}

		private void dtFechaCobro_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.cmbNombre.Focus();
		}

		private void txtValor_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.txtNumCuotas.Focus();
		}

		private void txtNumCuotas_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.txtConcepto.Focus();
		}

		private void txtConcepto_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.btnGuardar.Focus();
		}

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);	
		}

		private void dtFecha_Validated(object sender, System.EventArgs e)
		{
			if (bEdicion)
				if (this.dtFecha.Value == System.DBNull.Value) this.dtFecha.Value = DateTime.Today;
		}

		private void dtFechaCobro_Validated(object sender, System.EventArgs e)
		{
			if (bEdicion)
				if (this.dtFechaCobro.Value == System.DBNull.Value) this.dtFechaCobro.Value = DateTime.Today;
		}

		private void txtValor_Validated(object sender, System.EventArgs e)
		{
			if (bEdicion)
				if (this.txtValor.Value == System.DBNull.Value) this.txtValor.Value = 0.00m;
		}

		private void txtNumCuotas_Validated(object sender, System.EventArgs e)
		{
			if (bEdicion)
				if (this.txtNumCuotas.Value == System.DBNull.Value) this.txtNumCuotas.Value = 1;
		}

		private void cmbNombre_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.txtValor.Focus();
		}

		private void dtFecha_ValueChanged(object sender, System.EventArgs e)
		{
			int idPersonal = 0;
			if (this.cmbNombre.ActiveRow != null) idPersonal = (int)this.cmbNombre.Value;

			if (this.dtFecha.Value != System.DBNull.Value)
				if (this.cmbTipos.ActiveRow != null)
					this.cmbNombre.DataSource = Funcion.dvProcedimiento(cdsSeteoF, string.Format("Exec ListaEmpleadosAnticiposPrestamos '{0}', {1}, 0", 
						Convert.ToDateTime(this.dtFecha.Value).ToString("yyyyMMdd"), (int)this.cmbTipos.Value));
		}

		private void btnAnular_Click(object sender, System.EventArgs e)
		{
			this.Consultar((int)this.txtIdRolPrestamo.Value);

			if ((int)this.txtEstado.Value == 3)
			{
				MessageBox.Show("El registro ya esta Anulado", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return;
			}

			if (!miAcceso.Anular)
			{
				MessageBox.Show("No tiene acceso a Anular", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return;
			}

			if ((int)this.txtEstado.Value == 1)
			{
				MessageBox.Show("No puede Anular esta Autorizado el Pago", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return;
			}

			if ((int)this.txtEstado.Value == 2)
			{
				MessageBox.Show("No puede Anular esta Pagado", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return;
			}

			if (Funcion.ValidaBloqueo((int)this.txtIdRolPrestamo.Value, 10, cdsSeteoF)) return;

			if (DialogResult.Yes == MessageBox.Show("Esta seguro de Anular este registro", "Point Technoligy", MessageBoxButtons.YesNo, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2))
			{				
				using (frmJustAnulaVenta Notas = new frmJustAnulaVenta())
				{
					if (DialogResult.OK == Notas.ShowDialog())
					{
						try
						{
							string sSQL = string.Format("Exec AnulaFiniquitos {0}, {1}, {2}, {3}, '{4}'", 
								(int)this.txtIdRolPrestamo.Value, (int)this.txtIdAsiento.Value, (int)this.cmbTipos.Value, (int)this.txtEstado.Value, Notas.txtNotas.Text.ToString());
							Funcion.EjecutaSQL(cdsSeteoF, sSQL);

							MessageBox.Show("Registro Anulado", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);

							this.Consultar((int)this.txtIdRolPrestamo.Value);
						}
						catch(Exception Exc)
						{
							MessageBox.Show(string.Format("Error al anular: {0}", Exc), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
						}
					}
				}
			}			
		}

		private void dtFecha_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.dtFechaCobro.Focus();
		}

		private void txtValor_Enter(object sender, System.EventArgs e)
		{
			this.txtValor.SelectAll();
		}

		private void txtValor_Click(object sender, System.EventArgs e)
		{
			this.txtValor.SelectAll();
		}

		private void txtNumCuotas_Enter(object sender, System.EventArgs e)
		{
			this.txtNumCuotas.SelectAll();
		}

		private void txtNumCuotas_Click(object sender, System.EventArgs e)
		{
			this.txtNumCuotas.SelectAll();
		}

		private void frmRolPrestamo_Closing(object sender, System.ComponentModel.CancelEventArgs e)
		{
			if (bEdicion)
			{
				MessageBox.Show("No puede salir mientras edita un registro", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				e.Cancel = true;
				return;
			}
		}

		private void frmRolPrestamo_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (!bEdicion)
			{
				if (e.KeyCode == Keys.N) this.btnNuevo_Click(sender, e);
				if (e.KeyCode == Keys.B) this.btnConsultar_Click(sender, e);
				if ((int)this.txtIdRolPrestamo.Value > 0)
				{
					if (e.KeyCode == Keys.E) if ((int)this.txtEstado.Value == 0) if (this.btnEditar.Enabled) this.btnEditar_Click(sender, e);					
					if (e.KeyCode == Keys.A) if ((int)this.txtEstado.Value == 0) if (this.btnAnular.Enabled) this.btnAnular_Click(sender, e);					
					if (e.KeyCode == Keys.F5) this.Consultar((int)this.txtIdRolPrestamo.Value);					
				}
				if (e.KeyCode == Keys.Escape) this.Close();				
			}
			else
			{
				if (e.KeyCode == Keys.Escape) 
					if (DialogResult.Yes == MessageBox.Show("Esta seguro de Cancelar la Edición", "Point Technology", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2))
						this.btnCancelar_Click(sender, e);

				if (e.KeyCode == Keys.F12) this.btnGuardar_Click(sender, e);				
			}
		}
	}
}
