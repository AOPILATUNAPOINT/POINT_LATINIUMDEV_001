using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;
using System.Data.SqlClient;
using DevExpress.XtraEditors;
using CrystalDecisions;
using CrystalDecisions.CrystalReports.Engine;
using CrystalDecisions.ReportSource;
using CrystalDecisions.Shared;


namespace Latinium
{
	/// <summary>
	/// Descripción breve de frmPresupuestoSolicitudAumento.
	/// </summary>
	public class frmPresupuestoSolicitudAumento : System.Windows.Forms.Form
	{
		private System.Windows.Forms.Label label6;
		public Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtFecha;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtidAumentoPresupuesto;
		private System.Windows.Forms.Label lblComprobante;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbProyecto;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtPeriodo;
		private System.Windows.Forms.Label label1;
		private System.Windows.Forms.Label label2;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtMes;
		private System.Windows.Forms.Label label10;
		private System.Windows.Forms.Label label12;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtPresupuesto;
		private System.Windows.Forms.Label label3;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtDisponible;
		private System.Windows.Forms.Label label4;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtiEstado;
		private System.Windows.Forms.Label lblEstado;
		private System.Windows.Forms.Label label9;
		public Infragistics.Win.UltraWinEditors.UltraTextEditor txtNota;
		private System.Windows.Forms.Button btnEditar;
		private System.Windows.Forms.Button btnGuardar;
		private System.Windows.Forms.Button btnCancelar;
		private System.Windows.Forms.Button btnSalir;
		private System.Windows.Forms.Button btnNuevo;
		private System.Windows.Forms.Button btnConsultar;
		private C1.Data.C1DataSet cdsSeteoF;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtValor_Solicitado;
		private System.Windows.Forms.Label lblUsuario;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource1;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource2;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbCuenta;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbTipo;
		private System.Windows.Forms.Label label5;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource3;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public frmPresupuestoSolicitudAumento()
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
			Infragistics.Win.Appearance appearance1 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton1 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmPresupuestoSolicitudAumento));
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bodega");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Bodega");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn3 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCuenta");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn4 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Codigo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn5 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Descripcion");
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCuenta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Descripcion");
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand3 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idTipo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Tipo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn6 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idTipo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn7 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Tipo");
			this.label6 = new System.Windows.Forms.Label();
			this.dtFecha = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.txtidAumentoPresupuesto = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.lblComprobante = new System.Windows.Forms.Label();
			this.cmbProyecto = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.ultraDataSource1 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.txtPeriodo = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label1 = new System.Windows.Forms.Label();
			this.label2 = new System.Windows.Forms.Label();
			this.txtMes = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label10 = new System.Windows.Forms.Label();
			this.label12 = new System.Windows.Forms.Label();
			this.txtPresupuesto = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label3 = new System.Windows.Forms.Label();
			this.txtDisponible = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label4 = new System.Windows.Forms.Label();
			this.txtValor_Solicitado = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtiEstado = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.lblUsuario = new System.Windows.Forms.Label();
			this.lblEstado = new System.Windows.Forms.Label();
			this.label9 = new System.Windows.Forms.Label();
			this.txtNota = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.btnEditar = new System.Windows.Forms.Button();
			this.btnGuardar = new System.Windows.Forms.Button();
			this.btnCancelar = new System.Windows.Forms.Button();
			this.btnSalir = new System.Windows.Forms.Button();
			this.btnNuevo = new System.Windows.Forms.Button();
			this.btnConsultar = new System.Windows.Forms.Button();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.ultraDataSource2 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.cmbCuenta = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.cmbTipo = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.label5 = new System.Windows.Forms.Label();
			this.ultraDataSource3 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			((System.ComponentModel.ISupportInitialize)(this.dtFecha)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtidAumentoPresupuesto)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbProyecto)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtPeriodo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtMes)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtPresupuesto)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDisponible)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtValor_Solicitado)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtiEstado)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNota)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCuenta)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbTipo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource3)).BeginInit();
			this.SuspendLayout();
			// 
			// label6
			// 
			this.label6.AutoSize = true;
			this.label6.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label6.Location = new System.Drawing.Point(16, 16);
			this.label6.Name = "label6";
			this.label6.Size = new System.Drawing.Size(36, 16);
			this.label6.TabIndex = 201;
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
			this.dtFecha.Format = "dd/MM/yyyy";
			this.dtFecha.Location = new System.Drawing.Point(72, 14);
			this.dtFecha.Name = "dtFecha";
			this.dtFecha.NonAutoSizeHeight = 23;
			this.dtFecha.Size = new System.Drawing.Size(112, 21);
			this.dtFecha.SpinButtonsVisible = true;
			this.dtFecha.TabIndex = 200;
			this.dtFecha.Value = ((object)(resources.GetObject("dtFecha.Value")));
			this.dtFecha.KeyDown += new System.Windows.Forms.KeyEventHandler(this.dtFecha_KeyDown);
			this.dtFecha.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.dtFecha_KeyPress);
			this.dtFecha.ValueChanged += new System.EventHandler(this.dtFecha_ValueChanged);
			// 
			// txtidAumentoPresupuesto
			// 
			this.txtidAumentoPresupuesto.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtidAumentoPresupuesto.Enabled = false;
			this.txtidAumentoPresupuesto.Location = new System.Drawing.Point(600, 8);
			this.txtidAumentoPresupuesto.MaxValue = 3600;
			this.txtidAumentoPresupuesto.MinValue = 0;
			this.txtidAumentoPresupuesto.Name = "txtidAumentoPresupuesto";
			this.txtidAumentoPresupuesto.PromptChar = ' ';
			this.txtidAumentoPresupuesto.Size = new System.Drawing.Size(40, 21);
			this.txtidAumentoPresupuesto.TabIndex = 1004;
			this.txtidAumentoPresupuesto.Visible = false;
			// 
			// lblComprobante
			// 
			this.lblComprobante.AutoSize = true;
			this.lblComprobante.BackColor = System.Drawing.Color.Transparent;
			this.lblComprobante.ImeMode = System.Windows.Forms.ImeMode.NoControl;
			this.lblComprobante.Location = new System.Drawing.Point(16, 80);
			this.lblComprobante.Name = "lblComprobante";
			this.lblComprobante.Size = new System.Drawing.Size(49, 16);
			this.lblComprobante.TabIndex = 1006;
			this.lblComprobante.Text = "Proyecto";
			// 
			// cmbProyecto
			// 
			appearance2.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbProyecto.Appearance = appearance2;
			this.cmbProyecto.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbProyecto.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbProyecto.DataSource = this.ultraDataSource1;
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridColumn2.Width = 230;
			ultraGridBand1.Columns.AddRange(new object[] {
																										 ultraGridColumn1,
																										 ultraGridColumn2});
			ultraGridBand1.SummaryFooterCaption = "Grand Summaries";
			this.cmbProyecto.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.cmbProyecto.DisplayMember = "Nombre";
			this.cmbProyecto.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbProyecto.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbProyecto.Enabled = false;
			this.cmbProyecto.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F);
			this.cmbProyecto.Location = new System.Drawing.Point(72, 80);
			this.cmbProyecto.Name = "cmbProyecto";
			this.cmbProyecto.Size = new System.Drawing.Size(232, 21);
			this.cmbProyecto.TabIndex = 1007;
			this.cmbProyecto.ValueMember = "Bodega";
			this.cmbProyecto.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cmbProyecto_KeyDown);
			this.cmbProyecto.ValueChanged += new System.EventHandler(this.cmbProyecto_ValueChanged);
			// 
			// ultraDataSource1
			// 
			ultraDataColumn1.DataType = typeof(int);
			this.ultraDataSource1.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn1,
																																 ultraDataColumn2});
			// 
			// txtPeriodo
			// 
			this.txtPeriodo.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtPeriodo.Enabled = false;
			this.txtPeriodo.Location = new System.Drawing.Point(72, 46);
			this.txtPeriodo.MinValue = 2017;
			this.txtPeriodo.Name = "txtPeriodo";
			this.txtPeriodo.PromptChar = ' ';
			this.txtPeriodo.Size = new System.Drawing.Size(88, 21);
			this.txtPeriodo.SpinButtonDisplayStyle = Infragistics.Win.ButtonDisplayStyle.Always;
			this.txtPeriodo.TabIndex = 1008;
			this.txtPeriodo.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtPeriodo_KeyDown);
			this.txtPeriodo.ValueChanged += new System.EventHandler(this.txtPeriodo_ValueChanged);
			this.txtPeriodo.Enter += new System.EventHandler(this.txtPeriodo_Enter);
			// 
			// label1
			// 
			this.label1.AutoSize = true;
			this.label1.BackColor = System.Drawing.Color.Transparent;
			this.label1.ImeMode = System.Windows.Forms.ImeMode.NoControl;
			this.label1.Location = new System.Drawing.Point(16, 48);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(43, 16);
			this.label1.TabIndex = 1009;
			this.label1.Text = "Periodo";
			// 
			// label2
			// 
			this.label2.AutoSize = true;
			this.label2.BackColor = System.Drawing.Color.Transparent;
			this.label2.ImeMode = System.Windows.Forms.ImeMode.NoControl;
			this.label2.Location = new System.Drawing.Point(176, 48);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(26, 16);
			this.label2.TabIndex = 1011;
			this.label2.Text = "Mes";
			// 
			// txtMes
			// 
			this.txtMes.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtMes.Enabled = false;
			this.txtMes.Location = new System.Drawing.Point(216, 46);
			this.txtMes.MinValue = 1;
			this.txtMes.Name = "txtMes";
			this.txtMes.PromptChar = ' ';
			this.txtMes.Size = new System.Drawing.Size(88, 21);
			this.txtMes.SpinButtonDisplayStyle = Infragistics.Win.ButtonDisplayStyle.Always;
			this.txtMes.TabIndex = 1010;
			this.txtMes.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtMes_KeyDown);
			this.txtMes.ValueChanged += new System.EventHandler(this.txtMes_ValueChanged);
			this.txtMes.Enter += new System.EventHandler(this.txtMes_Enter);
			// 
			// label10
			// 
			this.label10.AutoSize = true;
			this.label10.BackColor = System.Drawing.Color.Transparent;
			this.label10.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label10.Location = new System.Drawing.Point(16, 112);
			this.label10.Name = "label10";
			this.label10.Size = new System.Drawing.Size(39, 17);
			this.label10.TabIndex = 1012;
			this.label10.Text = "Cuenta";
			this.label10.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label12
			// 
			this.label12.AutoSize = true;
			this.label12.Location = new System.Drawing.Point(16, 144);
			this.label12.Name = "label12";
			this.label12.Size = new System.Drawing.Size(68, 16);
			this.label12.TabIndex = 1014;
			this.label12.Text = "Presupuesto";
			// 
			// txtPresupuesto
			// 
			appearance3.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtPresupuesto.Appearance = appearance3;
			this.txtPresupuesto.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtPresupuesto.Enabled = false;
			this.txtPresupuesto.FormatString = "#,##0.00";
			this.txtPresupuesto.Location = new System.Drawing.Point(96, 142);
			this.txtPresupuesto.Name = "txtPresupuesto";
			this.txtPresupuesto.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtPresupuesto.PromptChar = ' ';
			this.txtPresupuesto.Size = new System.Drawing.Size(96, 21);
			this.txtPresupuesto.TabIndex = 1015;
			// 
			// label3
			// 
			this.label3.AutoSize = true;
			this.label3.Location = new System.Drawing.Point(200, 144);
			this.label3.Name = "label3";
			this.label3.Size = new System.Drawing.Size(57, 16);
			this.label3.TabIndex = 1016;
			this.label3.Text = "Disponible";
			// 
			// txtDisponible
			// 
			appearance4.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtDisponible.Appearance = appearance4;
			this.txtDisponible.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtDisponible.Enabled = false;
			this.txtDisponible.FormatString = "#,##0.00";
			this.txtDisponible.Location = new System.Drawing.Point(272, 142);
			this.txtDisponible.Name = "txtDisponible";
			this.txtDisponible.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtDisponible.PromptChar = ' ';
			this.txtDisponible.Size = new System.Drawing.Size(96, 21);
			this.txtDisponible.TabIndex = 1017;
			// 
			// label4
			// 
			this.label4.AutoSize = true;
			this.label4.Location = new System.Drawing.Point(384, 144);
			this.label4.Name = "label4";
			this.label4.Size = new System.Drawing.Size(53, 16);
			this.label4.TabIndex = 1018;
			this.label4.Text = "Solicitado";
			// 
			// txtValor_Solicitado
			// 
			appearance5.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtValor_Solicitado.Appearance = appearance5;
			this.txtValor_Solicitado.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtValor_Solicitado.Enabled = false;
			this.txtValor_Solicitado.FormatString = "#,##0.00";
			this.txtValor_Solicitado.Location = new System.Drawing.Point(448, 142);
			this.txtValor_Solicitado.Name = "txtValor_Solicitado";
			this.txtValor_Solicitado.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtValor_Solicitado.PromptChar = ' ';
			this.txtValor_Solicitado.Size = new System.Drawing.Size(96, 21);
			this.txtValor_Solicitado.TabIndex = 1019;
			this.txtValor_Solicitado.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtValor_Solicitado_KeyDown);
			this.txtValor_Solicitado.Enter += new System.EventHandler(this.txtValor_Solicitado_Enter);
			// 
			// txtiEstado
			// 
			this.txtiEstado.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtiEstado.Enabled = false;
			this.txtiEstado.Location = new System.Drawing.Point(600, 40);
			this.txtiEstado.MaxValue = 3600;
			this.txtiEstado.MinValue = 0;
			this.txtiEstado.Name = "txtiEstado";
			this.txtiEstado.PromptChar = ' ';
			this.txtiEstado.Size = new System.Drawing.Size(40, 21);
			this.txtiEstado.TabIndex = 1020;
			this.txtiEstado.Visible = false;
			// 
			// lblUsuario
			// 
			this.lblUsuario.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.lblUsuario.AutoSize = true;
			this.lblUsuario.Font = new System.Drawing.Font("Microsoft Sans Serif", 15F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblUsuario.ForeColor = System.Drawing.Color.Firebrick;
			this.lblUsuario.Location = new System.Drawing.Point(336, 24);
			this.lblUsuario.Name = "lblUsuario";
			this.lblUsuario.Size = new System.Drawing.Size(0, 26);
			this.lblUsuario.TabIndex = 1021;
			// 
			// lblEstado
			// 
			this.lblEstado.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.lblEstado.AutoSize = true;
			this.lblEstado.Font = new System.Drawing.Font("Microsoft Sans Serif", 15F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblEstado.ForeColor = System.Drawing.Color.Firebrick;
			this.lblEstado.Location = new System.Drawing.Point(336, 72);
			this.lblEstado.Name = "lblEstado";
			this.lblEstado.Size = new System.Drawing.Size(0, 26);
			this.lblEstado.TabIndex = 1022;
			// 
			// label9
			// 
			this.label9.AutoSize = true;
			this.label9.BackColor = System.Drawing.Color.Transparent;
			this.label9.Location = new System.Drawing.Point(16, 192);
			this.label9.Name = "label9";
			this.label9.Size = new System.Drawing.Size(28, 16);
			this.label9.TabIndex = 1023;
			this.label9.Text = "Nota";
			// 
			// txtNota
			// 
			appearance6.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtNota.Appearance = appearance6;
			this.txtNota.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtNota.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNota.Enabled = false;
			this.txtNota.Location = new System.Drawing.Point(72, 176);
			this.txtNota.Multiline = true;
			this.txtNota.Name = "txtNota";
			this.txtNota.Size = new System.Drawing.Size(384, 63);
			this.txtNota.TabIndex = 1024;
			this.txtNota.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtNota_KeyDown);
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
			this.btnEditar.Location = new System.Drawing.Point(192, 264);
			this.btnEditar.Name = "btnEditar";
			this.btnEditar.Size = new System.Drawing.Size(80, 23);
			this.btnEditar.TabIndex = 1028;
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
			this.btnGuardar.Location = new System.Drawing.Point(280, 264);
			this.btnGuardar.Name = "btnGuardar";
			this.btnGuardar.Size = new System.Drawing.Size(80, 23);
			this.btnGuardar.TabIndex = 1027;
			this.btnGuardar.Text = "Guardar";
			this.btnGuardar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnGuardar.Click += new System.EventHandler(this.btnGuardar_Click);
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
			this.btnCancelar.Location = new System.Drawing.Point(368, 264);
			this.btnCancelar.Name = "btnCancelar";
			this.btnCancelar.Size = new System.Drawing.Size(80, 23);
			this.btnCancelar.TabIndex = 1029;
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
			this.btnSalir.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnSalir.ImageIndex = 62;
			this.btnSalir.Location = new System.Drawing.Point(456, 264);
			this.btnSalir.Name = "btnSalir";
			this.btnSalir.Size = new System.Drawing.Size(80, 23);
			this.btnSalir.TabIndex = 1030;
			this.btnSalir.Text = "Salir";
			this.btnSalir.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnSalir.Click += new System.EventHandler(this.btnSalir_Click);
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
			this.btnNuevo.Location = new System.Drawing.Point(16, 264);
			this.btnNuevo.Name = "btnNuevo";
			this.btnNuevo.Size = new System.Drawing.Size(80, 23);
			this.btnNuevo.TabIndex = 1025;
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
			this.btnConsultar.Location = new System.Drawing.Point(104, 264);
			this.btnConsultar.Name = "btnConsultar";
			this.btnConsultar.Size = new System.Drawing.Size(80, 23);
			this.btnConsultar.TabIndex = 1026;
			this.btnConsultar.Text = "Consultar";
			this.btnConsultar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnConsultar.Click += new System.EventHandler(this.btnConsultar_Click);
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
																																 ultraDataColumn4,
																																 ultraDataColumn5});
			// 
			// cmbCuenta
			// 
			appearance7.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbCuenta.Appearance = appearance7;
			this.cmbCuenta.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbCuenta.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbCuenta.DataSource = this.ultraDataSource2;
			ultraGridColumn3.Header.VisiblePosition = 0;
			ultraGridColumn3.Hidden = true;
			ultraGridColumn4.Header.VisiblePosition = 1;
			ultraGridColumn5.Header.VisiblePosition = 2;
			ultraGridColumn5.Width = 232;
			ultraGridBand2.Columns.AddRange(new object[] {
																										 ultraGridColumn3,
																										 ultraGridColumn4,
																										 ultraGridColumn5});
			ultraGridBand2.SummaryFooterCaption = "Grand Summaries";
			this.cmbCuenta.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			this.cmbCuenta.DisplayMember = "Descripcion";
			this.cmbCuenta.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbCuenta.Enabled = false;
			this.cmbCuenta.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F);
			this.cmbCuenta.Location = new System.Drawing.Point(72, 112);
			this.cmbCuenta.Name = "cmbCuenta";
			this.cmbCuenta.Size = new System.Drawing.Size(232, 21);
			this.cmbCuenta.TabIndex = 1032;
			this.cmbCuenta.ValueMember = "idCuenta";
			this.cmbCuenta.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cmbCuenta_KeyDown_1);
			this.cmbCuenta.ValueChanged += new System.EventHandler(this.cmbCuenta_ValueChanged);
			// 
			// cmbTipo
			// 
			appearance8.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbTipo.Appearance = appearance8;
			this.cmbTipo.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbTipo.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbTipo.DataSource = this.ultraDataSource3;
			ultraGridColumn6.Header.VisiblePosition = 0;
			ultraGridColumn6.Hidden = true;
			ultraGridColumn7.Header.VisiblePosition = 1;
			ultraGridBand3.Columns.AddRange(new object[] {
																										 ultraGridColumn6,
																										 ultraGridColumn7});
			ultraGridBand3.SummaryFooterCaption = "Grand Summaries";
			this.cmbTipo.DisplayLayout.BandsSerializer.Add(ultraGridBand3);
			this.cmbTipo.DisplayMember = "Tipo";
			this.cmbTipo.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbTipo.Enabled = false;
			this.cmbTipo.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F);
			this.cmbTipo.Location = new System.Drawing.Point(392, 112);
			this.cmbTipo.Name = "cmbTipo";
			this.cmbTipo.Size = new System.Drawing.Size(152, 21);
			this.cmbTipo.TabIndex = 1033;
			this.cmbTipo.ValueMember = "idTipo";
			// 
			// label5
			// 
			this.label5.AutoSize = true;
			this.label5.Location = new System.Drawing.Point(352, 112);
			this.label5.Name = "label5";
			this.label5.Size = new System.Drawing.Size(26, 16);
			this.label5.TabIndex = 1034;
			this.label5.Text = "Tipo";
			this.label5.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// ultraDataSource3
			// 
			ultraDataColumn6.DataType = typeof(int);
			this.ultraDataSource3.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn6,
																																 ultraDataColumn7});
			// 
			// frmPresupuestoSolicitudAumento
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(672, 302);
			this.Controls.Add(this.label5);
			this.Controls.Add(this.cmbTipo);
			this.Controls.Add(this.cmbCuenta);
			this.Controls.Add(this.btnEditar);
			this.Controls.Add(this.btnGuardar);
			this.Controls.Add(this.btnCancelar);
			this.Controls.Add(this.btnSalir);
			this.Controls.Add(this.btnNuevo);
			this.Controls.Add(this.btnConsultar);
			this.Controls.Add(this.label9);
			this.Controls.Add(this.label4);
			this.Controls.Add(this.label3);
			this.Controls.Add(this.label12);
			this.Controls.Add(this.label10);
			this.Controls.Add(this.label2);
			this.Controls.Add(this.label1);
			this.Controls.Add(this.lblComprobante);
			this.Controls.Add(this.label6);
			this.Controls.Add(this.txtNota);
			this.Controls.Add(this.lblEstado);
			this.Controls.Add(this.lblUsuario);
			this.Controls.Add(this.txtiEstado);
			this.Controls.Add(this.txtValor_Solicitado);
			this.Controls.Add(this.txtDisponible);
			this.Controls.Add(this.txtPresupuesto);
			this.Controls.Add(this.txtMes);
			this.Controls.Add(this.txtPeriodo);
			this.Controls.Add(this.cmbProyecto);
			this.Controls.Add(this.txtidAumentoPresupuesto);
			this.Controls.Add(this.dtFecha);
			this.KeyPreview = true;
			this.Name = "frmPresupuestoSolicitudAumento";
			this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
			this.Text = "Solicitud Aumento Presupuesto";
			this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.frmPresupuestoSolicitudAumento_KeyDown);
			this.Closing += new System.ComponentModel.CancelEventHandler(this.frmPresupuestoSolicitudAumento_Closing);
			this.Load += new System.EventHandler(this.frmPresupuestoSolicitudAumento_Load);
			((System.ComponentModel.ISupportInitialize)(this.dtFecha)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtidAumentoPresupuesto)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbProyecto)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtPeriodo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtMes)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtPresupuesto)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDisponible)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtValor_Solicitado)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtiEstado)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNota)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCuenta)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbTipo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource3)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void btnCancelar_Click(object sender, System.EventArgs e)
		{
			
			bNuevo = false;
			bEdicion = false;

			this.txtidAumentoPresupuesto.Value = 0;
			this.dtFecha.Value = System.DBNull.Value;
      this.cmbProyecto.Value = System.DBNull.Value;
      this.txtPeriodo.Value = 2022;
			this.txtMes.Value = 1;
			this.cmbCuenta.Value = System.DBNull.Value;
			this.cmbTipo.Value = System.DBNull.Value;
			this.txtPresupuesto.Value = 0;
			this.txtDisponible.Value = 0;
			this.txtValor_Solicitado.Value = 0;
			this.txtNota.Text = "";
			this.txtiEstado.Value = 0;
			this.lblEstado.Text ="";
      this.lblUsuario.Text ="";



			this.dtFecha.Enabled = false;
			this.cmbProyecto.Enabled = false;
			this.txtPeriodo.Enabled = false;
			this.txtMes.Enabled = false;
			this.cmbCuenta.Enabled = false;
			this.cmbTipo.Enabled = false;
			this.txtValor_Solicitado.Enabled = false;
			this.txtNota.Enabled = false;


			if (miAcceso.Nuevo) this.btnNuevo.Enabled = true;
			if (miAcceso.Buscar) this.btnConsultar.Enabled = true;
			this.btnEditar.Enabled = false;
			this.btnGuardar.Enabled = false;
			this.btnCancelar.Enabled = false;

		}

		private void btnSalir_Click(object sender, System.EventArgs e)
		{
				this.Close();
		}
	
		bool bNuevo = false;
		bool bEdicion = false;
		bool bImpreso = false;
		private Acceso miAcceso;
		private void UnloadMe()
		{
			this.Close();
		}
		private void frmPresupuestoSolicitudAumento_Load(object sender, System.EventArgs e)
		{
			miAcceso = new Acceso(cdsSeteoF, "2611");

			if (!Funcion.Permiso("1775", cdsSeteoF))
			{				
				MessageBox.Show("No puede ingresar a Solicitud de Aumento de Presupuesto.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				BeginInvoke(new MethodInvoker(UnloadMe));
				return;
			}

			if (miAcceso.Nuevo) this.btnNuevo.Enabled = true;
			if (miAcceso.Buscar) this.btnConsultar.Enabled = true;

			//this.cmbProyecto.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Exec VerCentroDecosto 1");
			this.txtPeriodo.MinValue = 2022;
			this.txtPeriodo.MaxValue = DateTime.Today.Year + 1;
			this.txtPeriodo.Value = DateTime.Today.Year;

			this.txtMes.MinValue = 1;
			this.txtMes.MaxValue = 12;
			this.txtMes.Value = DateTime.Today.Month;
			this.cmbTipo.DataSource = Funcion.dvProcedimiento(cdsSeteoF,"SELECT 1 as idTipo, 'INCREMENTO' AS Tipo union SELECT 2 as idTipo, 'DISMINUIR' AS Tipo");

		}

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);
		
		}

		private void btnNuevo_Click(object sender, System.EventArgs e)
		{
			this.btnCancelar_Click(sender, e);
			FuncionesProcedimientos.RetornaFechaServidor(this.dtFecha, cdsSeteoF, false);
			this.txtiEstado.Value = 1;
			this.lblEstado.Text = "PENDIENTE";
			this.txtPeriodo.Enabled = true;
			this.txtMes.Enabled = true;
			this.cmbProyecto.Enabled = true;
			this.cmbCuenta.Enabled = true;
			this.cmbTipo.Enabled = true;
			this.txtValor_Solicitado.Enabled = true;
			this.txtNota.Enabled = true;
			this.btnNuevo.Enabled = false;
			this.btnConsultar.Enabled = false;
			this.btnEditar.Enabled = false;
			this.btnGuardar.Enabled = true;
			this.btnCancelar.Enabled = true;
			bNuevo = true;
			bEdicion = true;
		}

		private void ListaBodegas (Infragistics.Win.UltraWinGrid.UltraCombo uCombo, int idUsuario,  DateTime dtFecha,  C1.Data.C1DataSet ds)
		{
			if (miAcceso.BVistas)
			{
				uCombo.DataSource = Funcion.dvProcedimiento(ds, string.Format("Exec [VerCentroDecostoSolicitud] {0}", 1));
			}
			else
			{
				uCombo.DataSource = Funcion.dvProcedimiento(ds, string.Format("Exec PresupuestoListaCC {0}, '{1}'", idUsuario,  dtFecha.ToString("yyyyMMdd")));
			}
		}

		

		private void dtFecha_ValueChanged(object sender, System.EventArgs e)
		{
			if (this.dtFecha.Value != System.DBNull.Value)
			{	
				ListaBodegas(this.cmbProyecto, MenuLatinium.IdUsuario,  (DateTime)this.dtFecha.Value,  cdsSeteoF);
			}
		}

		private void dtFecha_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Delete) e.Handled = true;

			if (e.KeyCode == Keys.Enter)
				this.cmbProyecto.Focus();
		}

		private void dtFecha_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
				if (e.KeyChar == 8) e.Handled = true;
		}

		private void frmPresupuestoSolicitudAumento_Closing(object sender, System.ComponentModel.CancelEventArgs e)
		{
			if (bEdicion)
			{
				MessageBox.Show("No puede salir mientras edita una solicitud", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				e.Cancel = true;
				return;
			}
		}

		private void frmPresupuestoSolicitudAumento_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (!bEdicion)
			{
				
				if (e.KeyCode == Keys.N) if (miAcceso.Nuevo)	if (this.btnNuevo.Enabled) this.btnNuevo_Click(sender, e);
				if (e.KeyCode == Keys.B) if (miAcceso.Buscar) if (this.btnConsultar.Enabled) this.btnConsultar_Click(sender, e);
				if ((int)this.txtidAumentoPresupuesto.Value > 0)
				{	
					if (e.KeyCode == Keys.E) if (miAcceso.Editar) if (this.btnEditar.Enabled) this.btnEditar_Click(sender, e);
					if (e.KeyCode == Keys.F5)
					{
						this.Consulta((int)this.txtidAumentoPresupuesto.Value);				
					}
				}
				if (e.KeyCode == Keys.Escape) this.Close();				
			}
			else
			{
				if (e.KeyCode == Keys.Escape) this.btnCancelar_Click(sender, e);
				if (e.KeyCode == Keys.F12) this.btnGuardar_Click(sender, e);				
			}
		}

		private void btnConsultar_Click(object sender, System.EventArgs e)
		{
			this.btnCancelar_Click(sender, e);

			using (frmConusltaGeneralAumentoPresupuesto Busqueda = new frmConusltaGeneralAumentoPresupuesto(miAcceso.BVistas))
			{
				if (DialogResult.OK == Busqueda.ShowDialog())
				{
					this.Consulta((int)Busqueda.uGridCreditos.ActiveRow.Cells["idAumentoPresupuesto"].Value);
				}
			}
		}
		private void Consulta(int idAnticipo)
		{
			try
			{
				this.btnNuevo.Enabled = false;
				this.btnConsultar.Enabled = false;
				this.btnEditar.Enabled = false;
				this.btnGuardar.Enabled = false;
				this.btnCancelar.Enabled = false;				
        int iCuenta = 0;
				string sSQLBusqueda = string.Format("Exec BuscaAumentoPresupuestoIndividual {0}", idAnticipo);
				SqlDataReader dr = Funcion.rEscalarSQL(cdsSeteoF, sSQLBusqueda, true);
				dr.Read();
				if (dr.HasRows)
				{
					this.txtidAumentoPresupuesto.Value = dr.GetInt32(0);	
					this.dtFecha.Value = dr.GetDateTime(1);
					this.cmbProyecto.Value = dr.GetInt32(2);
					this.txtPeriodo.Value = dr.GetInt32(3);
					this.txtMes.Value = dr.GetInt32(4);
					iCuenta = dr.GetInt32(5);
					this.txtValor_Solicitado.Value = dr.GetDecimal(6);
					this.txtiEstado.Value = dr.GetInt32(7);
					this.txtNota.Value = dr.GetString(8);
					this.cmbTipo.Value = dr.GetInt32(9);
				}
				dr.Close();
			  this.cmbCuenta.Value = iCuenta;
				if ((int)this.txtiEstado.Value == 1) this.lblEstado.Text = "PENDIENTE";
				if ((int)this.txtiEstado.Value == 2) this.lblEstado.Text = "APROBADO";
				if ((int)this.txtiEstado.Value == 3) this.lblEstado.Text = "RECHAZADO";
						
				if ((int)this.txtiEstado.Value == 1 )
				{
					if (miAcceso.Editar)	this.btnEditar.Enabled = true;		
				}
				if (miAcceso.Nuevo) this.btnNuevo.Enabled = true;
				if (miAcceso.Buscar) this.btnConsultar.Enabled = true;
				this.btnCancelar.Enabled = true;							
			}
			catch(Exception Exc)
			{
				MessageBox.Show(string.Format("Error al Consultar: {0}", Exc), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
			}			
		}


		private void btnEditar_Click(object sender, System.EventArgs e)
		{
			this.Consulta((int)this.txtidAumentoPresupuesto.Value);
	
			if ((int)this.txtiEstado.Value == 2)
			{
				MessageBox.Show("La solicitud esta Aprobado", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return;
			}

			if ((int)this.txtiEstado.Value == 3)
			{
				MessageBox.Show("La solicitud esta Rechazado", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return;
			}						

			this.txtPeriodo.Enabled = true;
			this.txtMes.Enabled = true;
			this.cmbProyecto.Enabled = true;
			this.cmbCuenta.Enabled = true;
			this.cmbTipo.Enabled = true;
			this.txtValor_Solicitado.Enabled = true;
			this.txtNota.Enabled = true;


			this.btnNuevo.Enabled = false;
			this.btnConsultar.Enabled = false;
			this.btnEditar.Enabled = false;
			this.btnGuardar.Enabled = true;
			this.btnCancelar.Enabled = true;
			
			bNuevo = false;
			bEdicion = true;


		}

		private void btnGuardar_Click(object sender, System.EventArgs e)
		{
			#region Validacion
			if ( (int) this.txtPeriodo.Value < 2022)
			{
				MessageBox.Show(" Selecciona el Periodo", "Point Tehnology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.txtPeriodo.Focus();
				return;
			}

			if ( (int) this.txtMes.Value == 0)
			{
				MessageBox.Show(" El mes no puede ser 0", "Point Tehnology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.txtPeriodo.Focus();
				return;
			}
			if ( (int) this.txtMes.Value > 12)
			{
				MessageBox.Show(" El mes no puede ser mayor 12", "Point Tehnology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.txtMes.Focus();
				return;
			}
			if (!Validacion.vbComboVacio(this.cmbProyecto, "Seleccione el proyecto")) return;
			if (!Validacion.vbComboVacio(this.cmbCuenta, "Seleccione la Cuenta")) return;
			if (!Validacion.vbComboVacio(this.cmbTipo, "Seleccione el tipo de transacción")) return;
			if (!Validacion.vbCampoDecimalVacio (this.txtValor_Solicitado,"Ingrese el Valor a solicitar",1,1000000)) return;
			if (!Validacion.vbTexto(this.txtNota, 30, 299, "Notas")) return;
			if ( (int) this.cmbTipo.Value == 2)
			{
				if ( Convert.ToDecimal (this.txtValor_Solicitado.Value) > Convert.ToDecimal (this.txtDisponible.Value) )
				{
					MessageBox.Show("El valor solicitado no puede ser superior al disponible.", "Point Tehnology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					this.cmbTipo.Focus();
					return;
				}
			}

			string sPeriodoFecha = this.txtPeriodo.Value +"-"+this.txtMes.Value+"-01";

			if (Funcion.ValidaPeriodo(Convert.ToDateTime (sPeriodoFecha), cdsSeteoF, "Presupuesto"))
			{
					DateTime dtFechaVP = Convert.ToDateTime (sPeriodoFecha);
					MessageBox.Show(string.Format(" El Período '{0}' esta Cerrado,\n\n Consulte con Contabilidad", dtFechaVP.ToString("MMMM").ToUpper()), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					return;
			}
    #endregion Validacion

			int icontProeyecto = 0;
			icontProeyecto = Funcion.iEscalarSQL(cdsSeteoF, string.Format("Exec [ValidaPresupuestoPendiente] {0}, {1}, {2}, {3}", 
				 (int) this.txtPeriodo.Value, (int) this.txtMes.Value, (int) this.cmbProyecto.Value, (int) this.cmbCuenta.Value));	
			if ( bNuevo)
			{
				if( icontProeyecto > 0)
				{
					MessageBox.Show(" Ya existe una solicitud pendiente con los mismos parametros.", "Point Tehnology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					this.txtNota.Focus();
					return;
				}
			}
			if ( bEdicion)
			{
				if( icontProeyecto > 1)
				{
					MessageBox.Show(" Ya existe una solicitud pendiente con los mismos parametros.", "Point Tehnology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					this.txtNota.Focus();
					return;
				}
			}

			#region Guardar
			using (SqlConnection oConexion = new SqlConnection("Password=Bl45o6$9;Persist Security Info=True;User ID= " + MenuLatinium.stUsuario + ";Initial Catalog=" + MenuLatinium.stDirFacturacion + ";Data Source=" + MenuLatinium.stDirServidor + "; max pool size=1024"))
			{
				oConexion.Open();

				SqlTransaction oTransaction = null;
				SqlCommand oCmdActualiza = oConexion.CreateCommand();
				oCmdActualiza.Connection = oConexion;
				oCmdActualiza.CommandType = CommandType.Text;
				oCmdActualiza.CommandTimeout = 120;

				try
				{
					oTransaction = oConexion.BeginTransaction();
					oCmdActualiza.Transaction = oTransaction;

					#region Graba Maestro
				
					string  sSQL1 = string.Format("Exec [GuardaAumentaPresupuesto] {0}, '{1}', {2}, {3}, {4}, {5}, {6}, {7}, '{8}', {9}, {10}" ,

						(int)this.txtidAumentoPresupuesto.Value, 
						Convert.ToDateTime(this.dtFecha.Value).ToString("yyyyMMdd"), 
						(int)this.cmbProyecto.Value,
						(int)this.txtPeriodo.Value, 
						(int)this.txtMes.Value,
						(int)this.cmbCuenta.Value,
						Convert.ToDecimal (this.txtPresupuesto.Value), 
						Convert.ToDecimal (this.txtValor_Solicitado.Value), 
						this.txtNota.Text.ToString(), 
						(int)this.txtiEstado.Value,
            (int)this.cmbTipo.Value	);
					oCmdActualiza.CommandText = sSQL1;
					this.txtidAumentoPresupuesto.Value = (int)oCmdActualiza.ExecuteScalar();
					#endregion Graba Maestro
          
					oTransaction.Commit();

					Cursor = Cursors.Default;	

					MessageBox.Show("Registro Guardado Correctamente", "Información", MessageBoxButtons.OK, MessageBoxIcon.Information);

					#region Controles
					this.dtFecha.Enabled = false;
					this.cmbProyecto.Enabled = false;
					this.txtPeriodo.Enabled = false;
					this.txtMes.Enabled = false;
					this.cmbCuenta.Enabled = false;
					this.txtValor_Solicitado.Enabled = false;
					this.txtNota.Enabled = false;
					this.cmbTipo.Enabled = false;
				
					if (miAcceso.Nuevo)this.btnNuevo.Enabled = true;
					if (miAcceso.Buscar)this.btnConsultar.Enabled = true;
					if (miAcceso.Editar)this.btnEditar.Enabled = false;

					this.btnGuardar.Enabled = false;
					this.btnCancelar.Enabled = true;

					this.Consulta((int)this.txtidAumentoPresupuesto.Value);
					bNuevo = false;
					bEdicion = false;
					#endregion Controles			
				}
				catch(Exception ex)
				{
					try
					{ 
						oTransaction.Rollback();

						MessageBox.Show(string.Format("Commit Exception Type : {0} {1}", ex.GetType(), ex.Message), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
					}
					catch (Exception ex2)
					{
						MessageBox.Show(string.Format("Rollback Exception Type: {0} {1}", ex2.GetType(), ex2.Message), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);							
					}
				}
				finally
				{
					oConexion.Close();
				}
			}
				
			#endregion Guardar

		}

		private void txtPeriodo_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter) this.txtMes.Focus();
		}

		private void txtMes_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter) this.cmbProyecto.Focus();
		}

		private void cmbProyecto_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter) this.cmbCuenta.Focus();
		}

		private void cmbCuenta_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter) this.txtValor_Solicitado.Focus();
		}

		private void txtValor_Solicitado_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter) this.txtNota.Focus();
		}

		private void txtNota_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter) this.btnGuardar.Focus();
		}

		private void txtPeriodo_Enter(object sender, System.EventArgs e)
		{
			this.txtPeriodo.SelectAll();
		}

		private void txtMes_Enter(object sender, System.EventArgs e)
		{
			this.txtMes.SelectAll();
		}

		private void txtValor_Solicitado_Enter(object sender, System.EventArgs e)
		{
			this.txtValor_Solicitado.SelectAll();
		}

		private void txtPeriodo_ValueChanged(object sender, System.EventArgs e)
		{
			if (this.txtPeriodo.Value != System.DBNull.Value)
			{
				if ( this.cmbProyecto.ActiveRow != null)
				{	
					ListaCuenta(this.cmbCuenta,(int) this.txtPeriodo.Value, (int) this.cmbProyecto.Value,  cdsSeteoF);
				}
			}
		}

		private void ListaCuenta (Infragistics.Win.UltraWinGrid.UltraCombo uCombo, int iPeriodo,  int idProyecto,  C1.Data.C1DataSet ds)
		{
			uCombo.DataSource = Funcion.dvProcedimiento(ds, string.Format("Exec [VerCuentasPorCC] {0}, {1}", iPeriodo, idProyecto ));
		}

		private void cmbProyecto_ValueChanged(object sender, System.EventArgs e)
		{
			if (this.txtPeriodo.Value != System.DBNull.Value)
			{
				if ( this.cmbProyecto.ActiveRow != null)
				{	
					ListaCuenta(this.cmbCuenta,(int) this.txtPeriodo.Value, (int) this.cmbProyecto.Value,  cdsSeteoF);
				}
			}
			if (this.txtPeriodo.Value != System.DBNull.Value)
			{
				if (this.txtMes.Value != System.DBNull.Value)
				{
					if ( this.cmbProyecto.ActiveRow != null)
					{
						if ( this.cmbCuenta.ActiveRow != null)
						{
							string sFecha = "01/"+ Convert.ToString( this.txtMes.Value)+"/" + Convert.ToString( this.txtPeriodo.Value);

							decimal Presupuesto = 0;
							decimal Disponible = 0;

							string sSQL = string.Format("Exec [ValidaCCPresupuestoAsientoValores] '{0}', {1}, {2}", 
								Convert.ToDateTime(sFecha).ToString("yyyyMMdd"), (int) this.cmbCuenta.Value, (int) this.cmbProyecto.Value );
							SqlDataReader dra = Funcion.rEscalarSQL(cdsSeteoF, sSQL, true);
							dra.Read();
							if (dra.HasRows)
							{
								if (dra.GetValue(0) != System.DBNull.Value) Presupuesto =  dra.GetDecimal (0);
								if (dra.GetValue(2) != System.DBNull.Value) Disponible = dra.GetDecimal (2);
											
							}
							dra.Close();
							this.txtPresupuesto.Value = Presupuesto;
							this.txtDisponible.Value = Disponible;
						}
					}
				}
			}
		}

		private void cmbCuenta_KeyDown_1(object sender, System.Windows.Forms.KeyEventArgs e)
		{
						if(e.KeyCode == Keys.Enter) this.txtValor_Solicitado.Focus();
		}

		private void cmbCuenta_ValueChanged(object sender, System.EventArgs e)
		{
			if (this.txtPeriodo.Value != System.DBNull.Value)
			{
				if (this.txtMes.Value != System.DBNull.Value)
				{
					if ( this.cmbProyecto.ActiveRow != null)
					{
						if ( this.cmbCuenta.ActiveRow != null)
						{
							string sFecha = "01/"+ Convert.ToString( this.txtMes.Value)+"/" + Convert.ToString( this.txtPeriodo.Value);

							decimal Presupuesto = 0;
							decimal Disponible = 0;

							string sSQL = string.Format("Exec [ValidaCCPresupuestoAsientoValores] '{0}', {1}, {2}", 
								Convert.ToDateTime(sFecha).ToString("yyyyMMdd"), (int) this.cmbCuenta.Value, (int) this.cmbProyecto.Value );
							SqlDataReader dra = Funcion.rEscalarSQL(cdsSeteoF, sSQL, true);
							dra.Read();
							if (dra.HasRows)
							{
								if (dra.GetValue(0) != System.DBNull.Value) Presupuesto =  dra.GetDecimal (0);
								if (dra.GetValue(2) != System.DBNull.Value) Disponible = dra.GetDecimal (2);
											
							}
							dra.Close();
							this.txtPresupuesto.Value = Presupuesto;
							this.txtDisponible.Value = Disponible;
						}
					}
				}
			}
		}

		private void txtMes_ValueChanged(object sender, System.EventArgs e)
		{
			if (this.txtPeriodo.Value != System.DBNull.Value)
			{
				if (this.txtMes.Value != System.DBNull.Value)
				{
					if ( this.cmbProyecto.ActiveRow != null)
					{
						if ( this.cmbCuenta.ActiveRow != null)
						{
							string sFecha = "01/"+ Convert.ToString( this.txtMes.Value)+"/" + Convert.ToString( this.txtPeriodo.Value);

							decimal Presupuesto = 0;
							decimal Disponible = 0;

							string sSQL = string.Format("Exec [ValidaCCPresupuestoAsientoValores] '{0}', {1}, {2}", 
								Convert.ToDateTime(sFecha).ToString("yyyyMMdd"), (int) this.cmbCuenta.Value, (int) this.cmbProyecto.Value );
							SqlDataReader dra = Funcion.rEscalarSQL(cdsSeteoF, sSQL, true);
							dra.Read();
							if (dra.HasRows)
							{
								if (dra.GetValue(0) != System.DBNull.Value) Presupuesto =  dra.GetDecimal (0);
								if (dra.GetValue(2) != System.DBNull.Value) Disponible = dra.GetDecimal (2);
											
							}
							dra.Close();
							this.txtPresupuesto.Value = Presupuesto;
							this.txtDisponible.Value = Disponible;
						}
					}
				}
			}
		}
	}
}
