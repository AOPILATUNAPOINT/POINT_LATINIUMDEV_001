using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;
using System.Data.SqlClient;


namespace Latinium
{
	/// <summary>
	/// Descripción breve de frmFactoresArticulos.
	/// </summary>
	public class frmFactoresArticulos : System.Windows.Forms.Form
	{
		private System.Windows.Forms.Label lblContador;
		private System.Windows.Forms.Label label4;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtDescripcion;
		private System.Windows.Forms.Label label2;
		public Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtHasta;
		private Infragistics.Win.UltraWinEditors.UltraOptionSet optEstado;
		private System.Windows.Forms.Label label1;
		public Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtDesde;
		private System.Windows.Forms.Button btnEditar;
		private System.Windows.Forms.Button btnGuardar;
		private System.Windows.Forms.Button btnCancelar;
		private System.Windows.Forms.Button btnNuevo;
		private System.Windows.Forms.Button btnConsultar;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtestado;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtidFactorArticulo;
		private C1.Data.C1DataSet cdsSeteoF;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource3;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource5;
		private System.Windows.Forms.Label lblEstado;
		private System.Windows.Forms.GroupBox groupBox2;
		private System.Windows.Forms.GroupBox groupBox1;
		private Infragistics.Win.UltraWinGrid.UltraGrid uGridInformacion;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource4;
		private Infragistics.Win.UltraWinGrid.UltraDropDown cmbSubGrupoDet;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource1;
		private System.Windows.Forms.Button btnSalir;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public frmFactoresArticulos()
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
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton1 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmFactoresArticulos));
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.ValueListItem valueListItem1 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem2 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem3 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton2 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idArticulo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Codigo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn3 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("IdDetFactoresArticulos");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn4 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("IdFactoresArticulos");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn5 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idArticulo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn6 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Factor1");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn7 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Factor2");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn8 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Factor3");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn9 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Factor4");
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("IdDetFactoresArticulos");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("IdFactoresArticulos");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("IdArticulo", -1, "cmbSubGrupoDet");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Factor1");
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Factor2");
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Factor3");
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Factor4");
			Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance11 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance12 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance13 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance14 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance15 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn10 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("IdDetFactoresArticulos");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn11 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("IdFactoresArticulos");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn12 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("IdArticulo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn13 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Factor1");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn14 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Factor2");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn15 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Factor3");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn16 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Factor4");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idArticulo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn17 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idArticulo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn18 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Codigo");
			this.lblContador = new System.Windows.Forms.Label();
			this.label4 = new System.Windows.Forms.Label();
			this.txtDescripcion = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label2 = new System.Windows.Forms.Label();
			this.dtHasta = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.optEstado = new Infragistics.Win.UltraWinEditors.UltraOptionSet();
			this.label1 = new System.Windows.Forms.Label();
			this.dtDesde = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.btnEditar = new System.Windows.Forms.Button();
			this.btnGuardar = new System.Windows.Forms.Button();
			this.btnCancelar = new System.Windows.Forms.Button();
			this.btnNuevo = new System.Windows.Forms.Button();
			this.btnConsultar = new System.Windows.Forms.Button();
			this.txtestado = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtidFactorArticulo = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.ultraDataSource3 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.ultraDataSource5 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.lblEstado = new System.Windows.Forms.Label();
			this.groupBox2 = new System.Windows.Forms.GroupBox();
			this.groupBox1 = new System.Windows.Forms.GroupBox();
			this.uGridInformacion = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource4 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.cmbSubGrupoDet = new Infragistics.Win.UltraWinGrid.UltraDropDown();
			this.ultraDataSource1 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.btnSalir = new System.Windows.Forms.Button();
			((System.ComponentModel.ISupportInitialize)(this.txtDescripcion)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtHasta)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.optEstado)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtDesde)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtestado)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtidFactorArticulo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource3)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource5)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridInformacion)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource4)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbSubGrupoDet)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).BeginInit();
			this.SuspendLayout();
			// 
			// lblContador
			// 
			this.lblContador.AutoSize = true;
			this.lblContador.Font = new System.Drawing.Font("Tahoma", 12F, System.Drawing.FontStyle.Bold);
			this.lblContador.ForeColor = System.Drawing.Color.Firebrick;
			this.lblContador.Location = new System.Drawing.Point(320, 88);
			this.lblContador.Name = "lblContador";
			this.lblContador.Size = new System.Drawing.Size(0, 23);
			this.lblContador.TabIndex = 968;
			this.lblContador.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// label4
			// 
			this.label4.AutoSize = true;
			this.label4.Location = new System.Drawing.Point(8, 24);
			this.label4.Name = "label4";
			this.label4.Size = new System.Drawing.Size(44, 16);
			this.label4.TabIndex = 962;
			this.label4.Text = "Nombre";
			this.label4.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtDescripcion
			// 
			appearance1.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtDescripcion.Appearance = appearance1;
			this.txtDescripcion.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtDescripcion.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtDescripcion.Enabled = false;
			this.txtDescripcion.Location = new System.Drawing.Point(64, 24);
			this.txtDescripcion.Multiline = true;
			this.txtDescripcion.Name = "txtDescripcion";
			this.txtDescripcion.Size = new System.Drawing.Size(488, 21);
			this.txtDescripcion.TabIndex = 961;
			this.txtDescripcion.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtDescripcion_KeyDown);
			this.txtDescripcion.ValueChanged += new System.EventHandler(this.txtDescripcion_ValueChanged);
			// 
			// label2
			// 
			this.label2.AutoSize = true;
			this.label2.Location = new System.Drawing.Point(200, 58);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(34, 16);
			this.label2.TabIndex = 964;
			this.label2.Text = "Hasta";
			this.label2.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// dtHasta
			// 
			appearance2.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtHasta.Appearance = appearance2;
			this.dtHasta.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton1.Caption = "Today";
			this.dtHasta.DateButtons.Add(dateButton1);
			this.dtHasta.Enabled = false;
			this.dtHasta.Format = "dd/MM/yyyy ";
			this.dtHasta.Location = new System.Drawing.Point(248, 56);
			this.dtHasta.Name = "dtHasta";
			this.dtHasta.NonAutoSizeHeight = 23;
			this.dtHasta.Size = new System.Drawing.Size(124, 21);
			this.dtHasta.SpinButtonsVisible = true;
			this.dtHasta.TabIndex = 966;
			this.dtHasta.Value = ((object)(resources.GetObject("dtHasta.Value")));
			// 
			// optEstado
			// 
			appearance3.ForeColorDisabled = System.Drawing.Color.Black;
			this.optEstado.Appearance = appearance3;
			this.optEstado.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.optEstado.CheckedIndex = 0;
			this.optEstado.Enabled = false;
			this.optEstado.ItemAppearance = appearance4;
			this.optEstado.ItemOrigin = new System.Drawing.Point(10, 0);
			valueListItem1.DataValue = 0;
			valueListItem1.DisplayText = "Pendiente";
			valueListItem2.DataValue = 1;
			valueListItem2.DisplayText = "Activo";
			valueListItem3.DataValue = 2;
			valueListItem3.DisplayText = "Finalizado";
			this.optEstado.Items.Add(valueListItem1);
			this.optEstado.Items.Add(valueListItem2);
			this.optEstado.Items.Add(valueListItem3);
			this.optEstado.ItemSpacingVertical = 8;
			this.optEstado.Location = new System.Drawing.Point(64, 88);
			this.optEstado.Name = "optEstado";
			this.optEstado.Size = new System.Drawing.Size(240, 26);
			this.optEstado.TabIndex = 967;
			this.optEstado.Text = "Pendiente";
			this.optEstado.ValueChanged += new System.EventHandler(this.optEstado_ValueChanged);
			// 
			// label1
			// 
			this.label1.AutoSize = true;
			this.label1.Location = new System.Drawing.Point(8, 58);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(37, 16);
			this.label1.TabIndex = 963;
			this.label1.Text = "Desde";
			this.label1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// dtDesde
			// 
			appearance5.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtDesde.Appearance = appearance5;
			this.dtDesde.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton2.Caption = "Today";
			this.dtDesde.DateButtons.Add(dateButton2);
			this.dtDesde.Enabled = false;
			this.dtDesde.Format = "dd/MM/yyyy ";
			this.dtDesde.Location = new System.Drawing.Point(64, 56);
			this.dtDesde.Name = "dtDesde";
			this.dtDesde.NonAutoSizeHeight = 23;
			this.dtDesde.Size = new System.Drawing.Size(124, 21);
			this.dtDesde.SpinButtonsVisible = true;
			this.dtDesde.TabIndex = 965;
			this.dtDesde.Value = ((object)(resources.GetObject("dtDesde.Value")));
			this.dtDesde.KeyDown += new System.Windows.Forms.KeyEventHandler(this.dtDesde_KeyDown);
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
			this.btnEditar.Location = new System.Drawing.Point(192, 360);
			this.btnEditar.Name = "btnEditar";
			this.btnEditar.Size = new System.Drawing.Size(80, 23);
			this.btnEditar.TabIndex = 988;
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
			this.btnGuardar.Location = new System.Drawing.Point(280, 360);
			this.btnGuardar.Name = "btnGuardar";
			this.btnGuardar.Size = new System.Drawing.Size(80, 23);
			this.btnGuardar.TabIndex = 987;
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
			this.btnCancelar.Location = new System.Drawing.Point(368, 360);
			this.btnCancelar.Name = "btnCancelar";
			this.btnCancelar.Size = new System.Drawing.Size(80, 23);
			this.btnCancelar.TabIndex = 989;
			this.btnCancelar.Text = "Cancelar";
			this.btnCancelar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnCancelar.Click += new System.EventHandler(this.btnCancelar_Click);
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
			this.btnNuevo.Location = new System.Drawing.Point(16, 360);
			this.btnNuevo.Name = "btnNuevo";
			this.btnNuevo.Size = new System.Drawing.Size(80, 23);
			this.btnNuevo.TabIndex = 985;
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
			this.btnConsultar.Location = new System.Drawing.Point(104, 360);
			this.btnConsultar.Name = "btnConsultar";
			this.btnConsultar.Size = new System.Drawing.Size(80, 23);
			this.btnConsultar.TabIndex = 986;
			this.btnConsultar.Text = "Consultar";
			this.btnConsultar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnConsultar.Click += new System.EventHandler(this.btnConsultar_Click);
			// 
			// txtestado
			// 
			this.txtestado.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtestado.Enabled = false;
			this.txtestado.Location = new System.Drawing.Point(624, 360);
			this.txtestado.MaxValue = 3600;
			this.txtestado.MinValue = 0;
			this.txtestado.Name = "txtestado";
			this.txtestado.PromptChar = ' ';
			this.txtestado.Size = new System.Drawing.Size(48, 21);
			this.txtestado.TabIndex = 992;
			this.txtestado.Visible = false;
			// 
			// txtidFactorArticulo
			// 
			this.txtidFactorArticulo.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtidFactorArticulo.Location = new System.Drawing.Point(560, 360);
			this.txtidFactorArticulo.MaxValue = 3600;
			this.txtidFactorArticulo.MinValue = 0;
			this.txtidFactorArticulo.Name = "txtidFactorArticulo";
			this.txtidFactorArticulo.PromptChar = ' ';
			this.txtidFactorArticulo.Size = new System.Drawing.Size(56, 21);
			this.txtidFactorArticulo.TabIndex = 991;
			this.txtidFactorArticulo.Visible = false;
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
			// ultraDataSource3
			// 
			ultraDataColumn1.DataType = typeof(int);
			this.ultraDataSource3.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn1,
																																 ultraDataColumn2});
			// 
			// ultraDataSource5
			// 
			ultraDataColumn3.DataType = typeof(int);
			ultraDataColumn4.DataType = typeof(int);
			ultraDataColumn5.DataType = typeof(int);
			ultraDataColumn6.DataType = typeof(System.Decimal);
			ultraDataColumn7.DataType = typeof(System.Decimal);
			ultraDataColumn8.DataType = typeof(System.Decimal);
			ultraDataColumn9.DataType = typeof(System.Decimal);
			this.ultraDataSource5.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn3,
																																 ultraDataColumn4,
																																 ultraDataColumn5,
																																 ultraDataColumn6,
																																 ultraDataColumn7,
																																 ultraDataColumn8,
																																 ultraDataColumn9});
			// 
			// lblEstado
			// 
			this.lblEstado.AutoSize = true;
			this.lblEstado.Font = new System.Drawing.Font("Tahoma", 12F, System.Drawing.FontStyle.Bold);
			this.lblEstado.ForeColor = System.Drawing.Color.Firebrick;
			this.lblEstado.Location = new System.Drawing.Point(400, 56);
			this.lblEstado.Name = "lblEstado";
			this.lblEstado.Size = new System.Drawing.Size(0, 23);
			this.lblEstado.TabIndex = 993;
			this.lblEstado.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// groupBox2
			// 
			this.groupBox2.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.groupBox2.Location = new System.Drawing.Point(-16, 128);
			this.groupBox2.Name = "groupBox2";
			this.groupBox2.Size = new System.Drawing.Size(704, 8);
			this.groupBox2.TabIndex = 994;
			this.groupBox2.TabStop = false;
			// 
			// groupBox1
			// 
			this.groupBox1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.groupBox1.Location = new System.Drawing.Point(0, 344);
			this.groupBox1.Name = "groupBox1";
			this.groupBox1.Size = new System.Drawing.Size(704, 8);
			this.groupBox1.TabIndex = 995;
			this.groupBox1.TabStop = false;
			// 
			// uGridInformacion
			// 
			this.uGridInformacion.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.uGridInformacion.Cursor = System.Windows.Forms.Cursors.Default;
			this.uGridInformacion.DataSource = this.ultraDataSource4;
			appearance6.BackColor = System.Drawing.Color.White;
			this.uGridInformacion.DisplayLayout.Appearance = appearance6;
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridColumn2.Hidden = true;
			ultraGridColumn3.ButtonDisplayStyle = Infragistics.Win.UltraWinGrid.ButtonDisplayStyle.Always;
			ultraGridColumn3.Header.Caption = "Artículo";
			ultraGridColumn3.Header.VisiblePosition = 2;
			ultraGridColumn3.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDown;
			ultraGridColumn3.Width = 300;
			appearance7.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn4.CellAppearance = appearance7;
			ultraGridColumn4.Format = "#,##0.00";
			ultraGridColumn4.Header.Caption = "13 Semanas";
			ultraGridColumn4.Header.VisiblePosition = 3;
			appearance8.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn5.CellAppearance = appearance8;
			ultraGridColumn5.Format = "#,##0.00";
			ultraGridColumn5.Header.Caption = "26 Semanas";
			ultraGridColumn5.Header.VisiblePosition = 4;
			appearance9.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn6.CellAppearance = appearance9;
			ultraGridColumn6.Format = "#,##0.00";
			ultraGridColumn6.Header.Caption = "39 Semanas";
			ultraGridColumn6.Header.VisiblePosition = 5;
			appearance10.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn7.CellAppearance = appearance10;
			ultraGridColumn7.Format = "#,##0.00";
			ultraGridColumn7.Header.Caption = "52 Semanas";
			ultraGridColumn7.Header.VisiblePosition = 6;
			ultraGridBand1.Columns.AddRange(new object[] {
																										 ultraGridColumn1,
																										 ultraGridColumn2,
																										 ultraGridColumn3,
																										 ultraGridColumn4,
																										 ultraGridColumn5,
																										 ultraGridColumn6,
																										 ultraGridColumn7});
			ultraGridBand1.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.True;
			ultraGridBand1.SummaryFooterCaption = "Totales:";
			this.uGridInformacion.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.uGridInformacion.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.TemplateOnBottom;
			this.uGridInformacion.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.True;
			this.uGridInformacion.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance11.BackColor = System.Drawing.Color.Transparent;
			this.uGridInformacion.DisplayLayout.Override.CardAreaAppearance = appearance11;
			appearance12.ForeColor = System.Drawing.Color.Black;
			appearance12.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridInformacion.DisplayLayout.Override.CellAppearance = appearance12;
			appearance13.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance13.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance13.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance13.FontData.BoldAsString = "True";
			appearance13.FontData.Name = "Arial";
			appearance13.FontData.SizeInPoints = 8F;
			appearance13.ForeColor = System.Drawing.Color.White;
			appearance13.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridInformacion.DisplayLayout.Override.HeaderAppearance = appearance13;
			this.uGridInformacion.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti;
			appearance14.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance14.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance14.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridInformacion.DisplayLayout.Override.RowSelectorAppearance = appearance14;
			appearance15.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance15.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance15.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridInformacion.DisplayLayout.Override.SelectedRowAppearance = appearance15;
			this.uGridInformacion.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridInformacion.Location = new System.Drawing.Point(8, 144);
			this.uGridInformacion.Name = "uGridInformacion";
			this.uGridInformacion.Size = new System.Drawing.Size(664, 192);
			this.uGridInformacion.TabIndex = 996;
			this.uGridInformacion.AfterRowInsert += new Infragistics.Win.UltraWinGrid.RowEventHandler(this.uGridInformacion_AfterRowInsert);
			this.uGridInformacion.BeforeRowsDeleted += new Infragistics.Win.UltraWinGrid.BeforeRowsDeletedEventHandler(this.uGridInformacion_BeforeRowsDeleted);
			this.uGridInformacion.KeyDown += new System.Windows.Forms.KeyEventHandler(this.uGridInformacion_KeyDown);
			this.uGridInformacion.AfterCellUpdate += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.uGridInformacion_AfterCellUpdate);
			// 
			// ultraDataSource4
			// 
			ultraDataColumn10.DataType = typeof(int);
			ultraDataColumn11.DataType = typeof(int);
			ultraDataColumn12.DataType = typeof(int);
			ultraDataColumn13.DataType = typeof(System.Decimal);
			ultraDataColumn14.DataType = typeof(System.Decimal);
			ultraDataColumn15.DataType = typeof(System.Decimal);
			ultraDataColumn16.DataType = typeof(System.Decimal);
			this.ultraDataSource4.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn10,
																																 ultraDataColumn11,
																																 ultraDataColumn12,
																																 ultraDataColumn13,
																																 ultraDataColumn14,
																																 ultraDataColumn15,
																																 ultraDataColumn16});
			// 
			// cmbSubGrupoDet
			// 
			this.cmbSubGrupoDet.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbSubGrupoDet.DataSource = this.ultraDataSource1;
			this.cmbSubGrupoDet.DisplayLayout.AutoFitColumns = true;
			ultraGridColumn8.Header.VisiblePosition = 0;
			ultraGridColumn8.Hidden = true;
			ultraGridColumn8.UseEditorMaskSettings = true;
			ultraGridColumn8.Width = 30;
			ultraGridColumn9.Header.VisiblePosition = 1;
			ultraGridColumn9.Width = 198;
			ultraGridBand2.Columns.AddRange(new object[] {
																										 ultraGridColumn8,
																										 ultraGridColumn9});
			ultraGridBand2.SummaryFooterCaption = "Grand Summaries";
			this.cmbSubGrupoDet.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			this.cmbSubGrupoDet.DisplayMember = "Codigo";
			this.cmbSubGrupoDet.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.cmbSubGrupoDet.Location = new System.Drawing.Point(440, 184);
			this.cmbSubGrupoDet.Name = "cmbSubGrupoDet";
			this.cmbSubGrupoDet.Size = new System.Drawing.Size(200, 104);
			this.cmbSubGrupoDet.TabIndex = 997;
			this.cmbSubGrupoDet.ValueMember = "idArticulo";
			this.cmbSubGrupoDet.Visible = false;
			// 
			// ultraDataSource1
			// 
			ultraDataColumn17.DataType = typeof(int);
			this.ultraDataSource1.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn17,
																																 ultraDataColumn18});
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
			this.btnSalir.Location = new System.Drawing.Point(456, 360);
			this.btnSalir.Name = "btnSalir";
			this.btnSalir.Size = new System.Drawing.Size(80, 23);
			this.btnSalir.TabIndex = 998;
			this.btnSalir.Text = "Salir";
			this.btnSalir.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnSalir.Click += new System.EventHandler(this.btnSalir_Click);
			// 
			// frmFactoresArticulos
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(680, 398);
			this.Controls.Add(this.btnSalir);
			this.Controls.Add(this.cmbSubGrupoDet);
			this.Controls.Add(this.uGridInformacion);
			this.Controls.Add(this.groupBox1);
			this.Controls.Add(this.groupBox2);
			this.Controls.Add(this.lblEstado);
			this.Controls.Add(this.txtestado);
			this.Controls.Add(this.txtidFactorArticulo);
			this.Controls.Add(this.btnEditar);
			this.Controls.Add(this.btnGuardar);
			this.Controls.Add(this.btnCancelar);
			this.Controls.Add(this.btnNuevo);
			this.Controls.Add(this.btnConsultar);
			this.Controls.Add(this.lblContador);
			this.Controls.Add(this.label4);
			this.Controls.Add(this.label2);
			this.Controls.Add(this.label1);
			this.Controls.Add(this.txtDescripcion);
			this.Controls.Add(this.dtHasta);
			this.Controls.Add(this.optEstado);
			this.Controls.Add(this.dtDesde);
			this.KeyPreview = true;
			this.Name = "frmFactoresArticulos";
			this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
			this.Text = "Factor Artículo";
			this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.frmFactoresArticulos_KeyDown);
			this.Load += new System.EventHandler(this.frmFactoresArticulos_Load);
			((System.ComponentModel.ISupportInitialize)(this.txtDescripcion)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtHasta)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.optEstado)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtDesde)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtestado)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtidFactorArticulo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource3)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource5)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridInformacion)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource4)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbSubGrupoDet)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private Acceso miAcceso;
		bool bNuevo = false;
		bool bEdicion = false;

		private void UnloadMe()
		{
			this.Close();
		}

		private void txtDescripcion_ValueChanged(object sender, System.EventArgs e)
		{
		
		}

		private void btnCancelar_Click(object sender, System.EventArgs e)
		{
			bNuevo = false;
			bEdicion = false;

			#region Vacia Campos 

			this.txtidFactorArticulo.Value = 0;
			this.txtDescripcion.Text = "";
			this.dtDesde.Value = System.DBNull.Value;
			this.dtHasta.Value = System.DBNull.Value;
			this.optEstado.Value = 0;
			//this.cmbArticulo.Value = System.DBNull.Value;
			this.txtestado.Value = 0;
			this.lblContador.Text = "";
			this.lblEstado.Text="";

		
			#endregion Vacia Campos 

			#region Bloquea Campo

			this.txtDescripcion.Enabled = false;
			this.dtDesde.Enabled = false;
			this.dtHasta.Enabled = false;
			this.optEstado.Enabled = false;
			FuncionesProcedimientos.EstadoGrids(false, this.uGridInformacion);

			#endregion Bloquea Campo

			this.uGridInformacion.DataSource = Funcion.dvProcedimiento(cdsSeteoF, string.Format("Exec ConsultaIndividualDetFactoresArticulos 0"));

			if (miAcceso.Nuevo) this.btnNuevo.Enabled = true;
			if (miAcceso.Buscar) this.btnConsultar.Enabled = true;
			this.btnEditar.Enabled = false;
			this.btnGuardar.Enabled = false;
			this.btnCancelar.Enabled = false;

		}

		private void frmFactoresArticulos_Load(object sender, System.EventArgs e)
		{
			miAcceso = new Acceso(cdsSeteoF, "0631");

			if (!Funcion.Permiso("1577", cdsSeteoF))
			{				
				MessageBox.Show("No puede ingresar a Factor de Artículos", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				BeginInvoke(new MethodInvoker(UnloadMe));
				return;
			}
			if (miAcceso.Nuevo) this.btnNuevo.Enabled = true;
			if (miAcceso.Buscar) this.btnConsultar.Enabled = true;
			this.dtDesde.CalendarInfo.MinDate = DateTime.Today;
			this.dtHasta.CalendarInfo.MinDate = DateTime.Today;

			this.cmbSubGrupoDet.DataSource = Funcion.dvProcedimiento(cdsSeteoF, " Exec [ConsultaArticuloFactor]");
			FuncionesProcedimientos.EstadoGrids(false, this.uGridInformacion);

		}

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);
		}

		private void optEstado_ValueChanged(object sender, System.EventArgs e)
		{
			if((int) this.optEstado.Value == 1)
			{
				if ((int)this.txtidFactorArticulo.Value == 0 )
				{
					RetornaFechaServidor(this.dtDesde, cdsSeteoF, false);
				}
				this.dtDesde.Enabled = false;
			}
			if((int) this.optEstado.Value == 0)
			{
				this.dtDesde.Enabled = true;
			}
			if((int) this.optEstado.Value == 2)
			{
				this.dtDesde.Enabled = false;
				this.dtHasta.Enabled = false;
			}
			
		}

		private void dtDesde_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter) this.dtHasta.Focus();
		}

		private void txtDescripcion_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter) this.dtDesde.Focus();
		}
		private void Estado ( int IEstado)
		{
			this.lblContador.Text = "";

			if (IEstado == 0)
			{
				this.lblEstado.Text = "PENDIENTE";
			}
			
			if (IEstado == 1)
			{
				this.lblEstado.Text = "ACTIVO";
			}

			if (IEstado == 2)
			{
				this.lblEstado.Text = " FINALIZADO";
			}

		}

		public void RetornaFechaServidor (Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtFecha, C1.Data.C1DataSet ds, bool bHora)
		{
			DateTime dtFechaServ = Funcion.dtEscalarSQL(ds, "Exec RetornaFechaServidor");

			if (!bHora) dtFecha.Value = dtFechaServ.Date;
			if (bHora) dtFecha.Value = dtFechaServ;

			dtFecha.CalendarInfo.MinDate = dtFechaServ.Date;
		} 

		private void btnNuevo_Click(object sender, System.EventArgs e)
		{
			this.btnCancelar_Click(sender, e);


			RetornaFechaServidor(this.dtDesde, cdsSeteoF, false);
			RetornaFechaServidor(this.dtHasta, cdsSeteoF, false);
			

			this.txtDescripcion.Enabled = true;
			this.dtHasta.Enabled = true;
			this.dtDesde.Enabled = true;
			this.optEstado.Enabled = true;
			FuncionesProcedimientos.EstadoGrids(true, this.uGridInformacion);
			Estado ((int) this.txtestado.Value);
   
			this.btnNuevo.Enabled = false;
			this.btnConsultar.Enabled = false;
			this.btnEditar.Enabled = false;
			this.btnGuardar.Enabled = true;
			this.btnCancelar.Enabled = true;	
			
			bNuevo = true;
			bEdicion = true;

			this.lblContador.Text = this.uGridInformacion.Rows.Count + " REGISTROS ENCONTRADOS";

			this.txtDescripcion.Focus();			
		}

		
		private void btnConsultar_Click(object sender, System.EventArgs e)
		{
						this.btnCancelar_Click(sender, e);
			using (frmConsultaFactoresArticulos consultab = new frmConsultaFactoresArticulos())
				if (DialogResult.OK == consultab.ShowDialog())
					this.Consultar((int)consultab.uGridInformacion.ActiveRow.Cells["IdFactoresArticulos"].Value);
		}
		private void Consultar (int IdFactoresArticulos )
		{
			#region Consulta Individual
			string ssql= string.Format ("Exec ConsultaIndividualFactoresArticulos {0}",IdFactoresArticulos); 
			SqlDataReader dr = Funcion.rEscalarSQL(cdsSeteoF,ssql,true);
			dr.Read();
			this.txtidFactorArticulo.Value= dr.GetInt32 (0);
			if (dr.GetValue(1) != System.DBNull.Value)this.txtDescripcion.Text = dr.GetString (1);
			if (dr.GetValue(2) != System.DBNull.Value)this.dtDesde.Value = dr.GetDateTime (2);
			if (dr.GetValue(3) != System.DBNull.Value)this.dtHasta.Value = dr.GetDateTime (3);
			if (dr.GetValue(4) != System.DBNull.Value)this.optEstado.Value = dr.GetInt32 (4);
			if (dr.GetValue(4) != System.DBNull.Value)this.txtestado.Value = dr.GetInt32 (4);
			
			dr.Close();

			this.uGridInformacion.DataSource = Funcion.dvProcedimiento(cdsSeteoF, string.Format("Exec ConsultaIndividualDetFactoresArticulos {0}", IdFactoresArticulos));
			#endregion Consulta Individual

			Estado ((int) this.txtestado.Value);

			this.btnGuardar.Enabled = false;
			this.btnCancelar.Enabled = true;

			if (miAcceso.Nuevo) this.btnNuevo.Enabled = true;
			if (miAcceso.Buscar) this.btnConsultar.Enabled = true;
			if (miAcceso.Editar)
			{
				if ((int)this.txtestado.Value == 0 )
				{
					this.btnEditar.Enabled = true;
				}
			}
			this.lblContador.Text = this.uGridInformacion.Rows.Count + " REGISTROS ENCONTRADOS";

		}


		private void btnGuardar_Click(object sender, System.EventArgs e)
		{
			#region Validacion
			if (!Validacion.vbTexto(this.txtDescripcion, 5, 100, "Descripción")) return;
			if (!Validacion.vbFechaEnDocumentos(this.dtDesde, "Ingrese la fecha Inicial ", true, "VENTA", cdsSeteoF)) return;
			if (!Validacion.vbFechaEnDocumentos(this.dtHasta, "Ingrese la fecha Final ", true, "VENTA", cdsSeteoF)) return;
			if (!Validacion.vbComparaFechas(Convert.ToDateTime(this.dtDesde.Value), Convert.ToDateTime(this.dtHasta.Value), this.dtDesde, "La Fecha Inicial no puede ser mayor a la Fecha Final", false)) return;

			if ( (int)this.optEstado.Value == 2 )
			{
				MessageBox.Show("No puede Ingresar un Factor Finalizado", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return;
			}

			if ( (int)this.uGridInformacion.Rows.Count == 0)
			{
				MessageBox.Show("Ingrese los Articulos de la Promoción ", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return;
			}		
			#endregion Validacion
      
			

			#region Valida precios vacios
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow drr in this.uGridInformacion.Rows.All)
			{
				#region Factor 0
				if (Convert.ToDecimal(drr.Cells["Factor1"].Value) <= 0 )
				{
					this.uGridInformacion.ActiveRow = drr;
					this.uGridInformacion.ActiveRow.Selected = true;
					MessageBox.Show("El Factor 1 no puede ser 0 ", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					return;
				}
				if (Convert.ToDecimal(drr.Cells["Factor2"].Value) <= 0)
				{
					this.uGridInformacion.ActiveRow = drr;
					this.uGridInformacion.ActiveRow.Selected = true;
					MessageBox.Show("El Factor 2 no puede ser 0 ", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					return;
				}
				if (Convert.ToDecimal(drr.Cells["Factor3"].Value) <= 0 )
				{
					this.uGridInformacion.ActiveRow = drr;
					this.uGridInformacion.ActiveRow.Selected = true;
					MessageBox.Show("El Factor 3 no puede ser 0 ", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					return;
				}
				if (Convert.ToDecimal(drr.Cells["Factor4"].Value) <= 0 )
				{
					this.uGridInformacion.ActiveRow = drr;
					this.uGridInformacion.ActiveRow.Selected = true;
					MessageBox.Show("El Factor 4 no puede ser 0 ", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					return;
				}
				#endregion Factor 0

				#region Factor 5
				if (Convert.ToDecimal(drr.Cells["Factor1"].Value) >= 2)
				{
					this.uGridInformacion.ActiveRow = drr;
					this.uGridInformacion.ActiveRow.Selected = true;
					MessageBox.Show("El Factor 1 no puede ser mayor a 1.99 ", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					return;
				}
				if (Convert.ToDecimal(drr.Cells["Factor2"].Value) >= 2)
				{
					this.uGridInformacion.ActiveRow = drr;
					this.uGridInformacion.ActiveRow.Selected = true;
					MessageBox.Show("El Factor 2 no puede ser mayor a 1.99 ", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					return;
				}
				if (Convert.ToDecimal(drr.Cells["Factor3"].Value) >= 2 )
				{
					this.uGridInformacion.ActiveRow = drr;
					this.uGridInformacion.ActiveRow.Selected = true;
					MessageBox.Show("El Factor 3 no puede ser mayor a 1.99 ", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					return;
				}
				if (Convert.ToDecimal(drr.Cells["Factor4"].Value) >= 2 )
				{
					this.uGridInformacion.ActiveRow = drr;
					this.uGridInformacion.ActiveRow.Selected = true;
					MessageBox.Show("El Factor 4 no puede ser mayor a 1.99 ", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					return;
				}
				#endregion Factor 5

			}
			#endregion Valida precios vacios

			#region Valida detalle articulo repetido
			int iArticulo = 0;
			int iCont = 0;

			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridInformacion.Rows.All)
			{         
				iArticulo = Convert.ToInt32(dr.Cells["IdArticulo"].Value); 

				foreach(Infragistics.Win.UltraWinGrid.UltraGridRow drr in this.uGridInformacion.Rows.All)
				{
					if ( iArticulo == Convert.ToInt32(drr.Cells["IdArticulo"].Value))
					{					
						iCont = iCont + 1;
						if (iCont > 1 )
						{
							this.uGridInformacion.ActiveRow = dr;
							this.uGridInformacion.ActiveRow.Selected = true;
						}
					}
				}

				if( iCont <= 1)	iCont = 0;
			}

			if (iCont > 1)
			{
				MessageBox.Show("Existen Productos repetidos", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return;
			}
			#endregion Valida detalle articulo repetido

			if ((int) this.txtidFactorArticulo.Value == 0 )
			{
				DateTime dtFechaServ = Funcion.dtEscalarSQL(cdsSeteoF, "Exec RetornaFechaServidor");
				if ((int) this.optEstado.Value == 0 )
				{
					if( Convert.ToDateTime(this.dtDesde.Value) == Convert.ToDateTime(dtFechaServ.Date))
					{
						this.optEstado.Value = 1;
						this.lblEstado.Text ="ACTIVO";
					}
				}
			}
			if ((int) this.txtidFactorArticulo.Value > 0 )
			{
				DateTime dtFechaServ = Funcion.dtEscalarSQL(cdsSeteoF, "Exec RetornaFechaServidor");
				if ((int) this.optEstado.Value == 1 && (DateTime) this.dtDesde.Value != dtFechaServ.Date)
				{
					this.optEstado.Value = 1;
					this.dtDesde.Value = dtFechaServ;
					this.lblEstado.Text ="ACTIVO";
				}
			}
			if ((int) this.txtidFactorArticulo.Value > 0 )
			{
				DateTime dtFechaServ = Funcion.dtEscalarSQL(cdsSeteoF, "Exec RetornaFechaServidor");
				if ((int) this.optEstado.Value == 0 )
				{
					if( Convert.ToDateTime(this.dtDesde.Value) == Convert.ToDateTime(dtFechaServ.Date))
					{
						this.optEstado.Value = 1;
						this.lblEstado.Text ="ACTIVO";
					}
				}
			}

			int iEstado = (int)this.optEstado.Value;
			this.txtestado.Value = (int)this.optEstado.Value;
			#region Guarda Maestro
			string sSQLl = string.Format("Exec GuardaFactoresArticulos {0}, '{1}' , '{2}' ,'{3}', {4}",
				(int)this.txtidFactorArticulo.Value,
				this.txtDescripcion.Text,
				Convert.ToDateTime(this.dtDesde.Value).ToString("yyyyMMdd"),
				Convert.ToDateTime(this.dtHasta.Value).ToString("yyyyMMdd"),
				iEstado);
			#endregion Guarda Maestro

			this.txtidFactorArticulo.Value=Funcion.iEscalarSQL(cdsSeteoF,sSQLl);

			#region Guarda detalle Maestro
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow drr in this.uGridInformacion.Rows.All)
			{	
				string sSQLDetalle = string.Format("Exec GuardaDetalleFactoresArticulos  {0}, {1}, {2}, {3}, {4}, {5},{6}", 
					(int)drr.Cells["IdDetFactoresArticulos"].Value,
					(int)this.txtidFactorArticulo.Value,
					(int)drr.Cells["idArticulo"].Value,
					(decimal)drr.Cells["Factor1"].Value,
					(decimal)drr.Cells["Factor2"].Value,
					(decimal)drr.Cells["Factor3"].Value,
					(decimal)drr.Cells["Factor4"].Value);
				Funcion.EjecutaSQL(cdsSeteoF, sSQLDetalle);
			}

			
				string sSQL = string.Format("Exec ActualizaPorcentajesFactorArticulo {0}, {1}", iEstado, (int) this.txtidFactorArticulo.Value);
				Funcion.EjecutaSQL(cdsSeteoF, sSQL, true);

			#endregion Guarda detalle Maestro

			#region Controles
			Consultar ((int) this.txtidFactorArticulo.Value);
			//this.txtidFacturacionPromociones.Enabled = false;
			this.txtDescripcion.Enabled = false;
			this.dtDesde.Enabled = false;
			this.dtHasta.Enabled = false;
			this.optEstado.Enabled = false;
			//this.cmbSubGrupoDet.Enabled = false;

			FuncionesProcedimientos.EstadoGrids(false, this.uGridInformacion);

			if (miAcceso.Nuevo) this.btnNuevo.Enabled = true;
			if (miAcceso.Buscar) this.btnConsultar.Enabled = true;
			if (miAcceso.Editar) this.btnEditar.Enabled = true;
			this.btnGuardar.Enabled = false;
			this.btnCancelar.Enabled = true;
	

							
			bNuevo = false;
			bEdicion = false;
			#endregion Controles	

		
		}

		private void uGridDetalle_AfterRowInsert(object sender, Infragistics.Win.UltraWinGrid.RowEventArgs e)
		{
			e.Row.Cells["IdDetFactoresArticulos"].Value = 0;
			e.Row.Cells["IdDetFactoresArticulos"].Value = 0;
			e.Row.Cells["idArticulo"].Value = 0;
			e.Row.Cells["Factor1"].Value = 0.00m; 
			e.Row.Cells["Factor2"].Value = 0.00m; 
			e.Row.Cells["Factor3"].Value = 0.00m; 
			e.Row.Cells["Factor4"].Value = 0.00m; 
		}

		private void uGridDetalle_AfterCellUpdate(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			int CantArt = 0;
			#region Valida serial existente en el grid
			//			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in uGridDetalle.Rows.All)
			//			{
			//				(int) sIdArticulo= dr.Cells["IdArticulo"].Value;
			//				if ((string)this.txtSerial.Text == sIdArticulo)
			//				{
			//					MessageBox.Show("Ya registro esta serial", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
			//					this.txtSerial.Text = "";
			//					this.txtSerial.Focus();
			//					return;
			//				}
			//			}
			#endregion Valida serial existente en el grid

			//	int IdArticulo = (int) this.uGridDetalle.ActiveRow.Cells["IdArticulo"].Value;
			#region Valida cantidad 
			//					int CantArticulo = 0;
			//					CantArticulo = (int)this.uGridDetalle.ActiveRow.Cells["IdArticulo"].Value;
			//												      
			//					foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr1 in this.uGridDetalle.Rows.All)
			//					{
			//						int idArtSerial = (int)dr1.Cells["IdArticulo"].Value;
			//
			//						if (IdArticulo == idArtSerial )
			//						{
			//							CantArt++;
			//						}
			//					}
			//
			//					if (CantArt >= CantArticulo)
			//					{
			//						MessageBox.Show("Ya ha ingresado las seriales para este articulo.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
			//	
			//						return;
			//					}
			//			
			#endregion Valida cantidad 
		}

		private void uGridDetalle_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}

		private void frmFactoresArticulos_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (!bEdicion)
			{
				if (e.KeyCode == Keys.N) this.btnNuevo_Click(sender, e);
				if (e.KeyCode == Keys.B) this.btnConsultar_Click(sender, e);
				if ((int)this.txtidFactorArticulo.Value > 0)
				{
					if (e.KeyCode == Keys.E) if (this.btnEditar.Enabled) this.btnEditar_Click(sender, e);					
					if (e.KeyCode == Keys.F5) this.Consultar((int)this.txtidFactorArticulo.Value);					
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

		private void btnEditar_Click(object sender, System.EventArgs e)
		{
			if ( (int)this.optEstado.Value == 1 )
			{
				MessageBox.Show("No puede Editar un Factor Activo", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return;
			}
			if ( (int)this.optEstado.Value == 2 )
			{
				MessageBox.Show("No puede Editar un factor Finalizado", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return;
			}

				this.txtDescripcion.Enabled = true;
				this.dtHasta.Enabled = true;
				this.dtDesde.Enabled = true;
				this.optEstado.Enabled = true;
				FuncionesProcedimientos.EstadoGrids(true, this.uGridInformacion);


			this.btnNuevo.Enabled = false;
			this.btnConsultar.Enabled = false;
			this.btnEditar.Enabled = false;
			this.btnGuardar.Enabled = true;
			this.btnCancelar.Enabled = true;

			bNuevo = false;
			bEdicion = true;
			this.txtDescripcion.Focus();
		}

		private void uGridInformacion_AfterRowInsert(object sender, Infragistics.Win.UltraWinGrid.RowEventArgs e)
		{
			e.Row.Cells["IdDetFactoresArticulos"].Value = 0;
			e.Row.Cells["IdFactoresArticulos"].Value = 0;
			e.Row.Cells["IdArticulo"].Value = 0;
			e.Row.Cells["Factor1"].Value = 0.00m; 
			e.Row.Cells["Factor2"].Value = 0.00m; 
			e.Row.Cells["Factor3"].Value = 0.00m;
			e.Row.Cells["Factor4"].Value = 0.00m; 
		}

		bool bActualiza = false;

		private void uGridInformacion_AfterCellUpdate(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			if (bActualiza) return;

			if (e.Cell.Column.ToString() == "Factor1")
			{	
				bActualiza = true;

				if ( e.Cell.Row.Cells["Factor1"].Value == System.DBNull.Value)
				{											
					e.Cell.Row.Cells["Factor1"].Value = (decimal)e.Cell.OriginalValue;						
				}	
				bActualiza = false;
			}

			if (e.Cell.Column.ToString() == "Factor2")
			{	
				bActualiza = true;

				if ( e.Cell.Row.Cells["Factor2"].Value == System.DBNull.Value)
				{											
					e.Cell.Row.Cells["Factor2"].Value = (decimal)e.Cell.OriginalValue;						
				}	
				bActualiza = false;
			}

			if (e.Cell.Column.ToString() == "Factor3")
			{	
				bActualiza = true;

				if ( e.Cell.Row.Cells["Factor3"].Value == System.DBNull.Value)
				{											
					e.Cell.Row.Cells["Factor3"].Value = (decimal)e.Cell.OriginalValue;						
				}	
				bActualiza = false;
			}
			if (e.Cell.Column.ToString() == "Factor4")
			{	
				bActualiza = true;

				if ( e.Cell.Row.Cells["Factor4"].Value == System.DBNull.Value)
				{											
					e.Cell.Row.Cells["Factor4"].Value = (decimal)e.Cell.OriginalValue;						
				}	
				bActualiza = false;
			}
		}

		private void btnSalir_Click(object sender, System.EventArgs e)
		{
			this.Close();	
		}

		private void uGridInformacion_BeforeRowsDeleted(object sender, Infragistics.Win.UltraWinGrid.BeforeRowsDeletedEventArgs e)
		{
			if (bEdicion)
			{
				#region Elimina
				if (DialogResult.Yes ==	MessageBox.Show("¿Esta seguro de borrar Los Gastos?", "Point Technology", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2))
				{				
					for (int i=0; i < e.Rows.Length; i++)
					{	
						if ((int)e.Rows[i].Cells["IdDetFactoresArticulos"].Value > 0 && (int) this.txtestado.Value == 0)
						{
							string sSQL = string.Format("Delete From DetalleFactoresArticulos where IdDetFactoresArticulos = {0} ", (int)e.Rows[i].Cells["IdDetFactoresArticulos"].Value);
							Funcion.EjecutaSQL(cdsSeteoF, sSQL, true);

							e.DisplayPromptMsg = false;
						}
						else e.DisplayPromptMsg = false;
					}					
				}
				else e.Cancel = true;
				#endregion Elimina
			}
		}

		private void uGridInformacion_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			FuncionesProcedimientos.DesplazamientoGrids(sender, e, this.uGridInformacion);
		}

	}
}
