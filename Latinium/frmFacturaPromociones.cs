using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;
using System.Data.SqlClient;
using System;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de frmFacturaPromociones.
	/// </summary>
	public class frmFacturaPromociones : System.Windows.Forms.Form
	{
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtidFacturacionPromociones;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtDescripcion;
		private System.Windows.Forms.Label label4;
		private System.Windows.Forms.Label label1;
		private System.Windows.Forms.Label label2;
		private Infragistics.Win.UltraWinEditors.UltraOptionSet optEstado;
		private System.Windows.Forms.GroupBox groupBox3;
		private System.Windows.Forms.Label label3;
		private System.Windows.Forms.Label label6;
		private System.Windows.Forms.Label label47;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbModelo;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbSubGrupo;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource1;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource2;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource3;
		private System.Windows.Forms.Label label5;
		private System.Windows.Forms.GroupBox groupBox2;
		private System.Windows.Forms.Button btnAgregar;
		private Infragistics.Win.UltraWinGrid.UltraGrid uGridInformacion;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource4;
		private Infragistics.Win.UltraWinGrid.UltraDropDown cmbSubGrupoDet;
		private Infragistics.Win.UltraWinGrid.UltraDropDown cmbMarcaDet;
		private Infragistics.Win.UltraWinGrid.UltraDropDown cmbModeloDet;
		private System.Windows.Forms.GroupBox groupBox1;
		private System.Windows.Forms.Button btnEditar;
		private System.Windows.Forms.Button btnGuardar;
		private System.Windows.Forms.Button btnCancelar;
		private System.Windows.Forms.Button btnSalir;
		private System.Windows.Forms.Button btnNuevo;
		private System.Windows.Forms.Button btnConsultar;
		public Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtDesde;
		public Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtHasta;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtPorcentaje;
		private C1.Data.C1DataSet cdsSeteoF;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource5;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbMarca;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtestado;
		private System.Windows.Forms.Label lblContador;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public frmFacturaPromociones()
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
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmFacturaPromociones));
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton2 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.ValueListItem valueListItem1 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem2 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem3 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Modelo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Modelo");
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idSubgrupo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("SubGrupo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idSubgrupo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn3 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("SubGrupo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn4 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idMarca");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn5 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Marca");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn6 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idArticulo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn7 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Modelo");
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand3 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idDetallePromociones");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idFacturacionPromociones");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idSubgrupo", -1, "cmbSubGrupoDet");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idMarca", -1, "cmbMarcaDet");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Modelo", -1, "cmbModeloDet");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Porcentaje");
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings1 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Porcentaje", 5, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Porcentaje", 5, true);
			Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance11 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance12 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance13 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance14 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance15 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn8 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idDetallePromociones");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn9 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idFacturacionPromociones");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn10 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idSubgrupo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn11 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idMarca");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn12 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Modelo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn13 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Porcentaje");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand4 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idSubgrupo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn11 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("SubGrupo");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand5 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn12 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idMarca");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn13 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Marca");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand6 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn14 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idArticulo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn15 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Modelo");
			Infragistics.Win.Appearance appearance16 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand7 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn16 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idMarca");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn17 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Marca");
			this.txtidFacturacionPromociones = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtDescripcion = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label4 = new System.Windows.Forms.Label();
			this.label1 = new System.Windows.Forms.Label();
			this.label2 = new System.Windows.Forms.Label();
			this.dtDesde = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.dtHasta = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.optEstado = new Infragistics.Win.UltraWinEditors.UltraOptionSet();
			this.groupBox3 = new System.Windows.Forms.GroupBox();
			this.lblContador = new System.Windows.Forms.Label();
			this.txtestado = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label3 = new System.Windows.Forms.Label();
			this.label6 = new System.Windows.Forms.Label();
			this.label47 = new System.Windows.Forms.Label();
			this.cmbModelo = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.ultraDataSource5 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.cmbSubGrupo = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.ultraDataSource1 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.ultraDataSource2 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.ultraDataSource3 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.txtPorcentaje = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label5 = new System.Windows.Forms.Label();
			this.groupBox2 = new System.Windows.Forms.GroupBox();
			this.btnAgregar = new System.Windows.Forms.Button();
			this.uGridInformacion = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource4 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.cmbSubGrupoDet = new Infragistics.Win.UltraWinGrid.UltraDropDown();
			this.cmbMarcaDet = new Infragistics.Win.UltraWinGrid.UltraDropDown();
			this.cmbModeloDet = new Infragistics.Win.UltraWinGrid.UltraDropDown();
			this.groupBox1 = new System.Windows.Forms.GroupBox();
			this.btnEditar = new System.Windows.Forms.Button();
			this.btnGuardar = new System.Windows.Forms.Button();
			this.btnCancelar = new System.Windows.Forms.Button();
			this.btnSalir = new System.Windows.Forms.Button();
			this.btnNuevo = new System.Windows.Forms.Button();
			this.btnConsultar = new System.Windows.Forms.Button();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.cmbMarca = new Infragistics.Win.UltraWinGrid.UltraCombo();
			((System.ComponentModel.ISupportInitialize)(this.txtidFacturacionPromociones)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDescripcion)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtDesde)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtHasta)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.optEstado)).BeginInit();
			this.groupBox3.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.txtestado)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbModelo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource5)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbSubGrupo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource3)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtPorcentaje)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridInformacion)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource4)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbSubGrupoDet)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbMarcaDet)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbModeloDet)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbMarca)).BeginInit();
			this.SuspendLayout();
			// 
			// txtidFacturacionPromociones
			// 
			this.txtidFacturacionPromociones.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtidFacturacionPromociones.Location = new System.Drawing.Point(600, 16);
			this.txtidFacturacionPromociones.MaxValue = 3600;
			this.txtidFacturacionPromociones.MinValue = 0;
			this.txtidFacturacionPromociones.Name = "txtidFacturacionPromociones";
			this.txtidFacturacionPromociones.PromptChar = ' ';
			this.txtidFacturacionPromociones.Size = new System.Drawing.Size(72, 22);
			this.txtidFacturacionPromociones.TabIndex = 672;
			this.txtidFacturacionPromociones.Visible = false;
			// 
			// txtDescripcion
			// 
			appearance1.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtDescripcion.Appearance = appearance1;
			this.txtDescripcion.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtDescripcion.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtDescripcion.Enabled = false;
			this.txtDescripcion.Location = new System.Drawing.Point(88, 24);
			this.txtDescripcion.Multiline = true;
			this.txtDescripcion.Name = "txtDescripcion";
			this.txtDescripcion.Size = new System.Drawing.Size(496, 21);
			this.txtDescripcion.TabIndex = 673;
			this.txtDescripcion.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtDescripcion_KeyDown);
			// 
			// label4
			// 
			this.label4.AutoSize = true;
			this.label4.Location = new System.Drawing.Point(16, 26);
			this.label4.Name = "label4";
			this.label4.Size = new System.Drawing.Size(44, 17);
			this.label4.TabIndex = 864;
			this.label4.Text = "Nombre";
			this.label4.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label1
			// 
			this.label1.AutoSize = true;
			this.label1.Location = new System.Drawing.Point(16, 58);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(35, 17);
			this.label1.TabIndex = 865;
			this.label1.Text = "Desde";
			this.label1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label2
			// 
			this.label2.AutoSize = true;
			this.label2.Location = new System.Drawing.Point(272, 58);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(33, 17);
			this.label2.TabIndex = 866;
			this.label2.Text = "Hasta";
			this.label2.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// dtDesde
			// 
			appearance2.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtDesde.Appearance = appearance2;
			this.dtDesde.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton1.Caption = "Today";
			this.dtDesde.DateButtons.Add(dateButton1);
			this.dtDesde.Enabled = false;
			this.dtDesde.Format = "dd/MM/yyyy ";
			this.dtDesde.Location = new System.Drawing.Point(88, 56);
			this.dtDesde.Name = "dtDesde";
			this.dtDesde.NonAutoSizeHeight = 23;
			this.dtDesde.Size = new System.Drawing.Size(124, 21);
			this.dtDesde.SpinButtonsVisible = true;
			this.dtDesde.TabIndex = 867;
			this.dtDesde.Value = ((object)(resources.GetObject("dtDesde.Value")));
			this.dtDesde.KeyDown += new System.Windows.Forms.KeyEventHandler(this.dtDesde_KeyDown);
			// 
			// dtHasta
			// 
			appearance3.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtHasta.Appearance = appearance3;
			this.dtHasta.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton2.Caption = "Today";
			this.dtHasta.DateButtons.Add(dateButton2);
			this.dtHasta.Enabled = false;
			this.dtHasta.Format = "dd/MM/yyyy ";
			this.dtHasta.Location = new System.Drawing.Point(336, 56);
			this.dtHasta.Name = "dtHasta";
			this.dtHasta.NonAutoSizeHeight = 23;
			this.dtHasta.Size = new System.Drawing.Size(124, 21);
			this.dtHasta.SpinButtonsVisible = true;
			this.dtHasta.TabIndex = 868;
			this.dtHasta.Value = ((object)(resources.GetObject("dtHasta.Value")));
			this.dtHasta.KeyDown += new System.Windows.Forms.KeyEventHandler(this.dtHasta_KeyDown);
			// 
			// optEstado
			// 
			appearance4.ForeColorDisabled = System.Drawing.Color.Black;
			this.optEstado.Appearance = appearance4;
			this.optEstado.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.optEstado.CheckedIndex = 0;
			this.optEstado.Enabled = false;
			this.optEstado.ItemAppearance = appearance5;
			this.optEstado.ItemOrigin = new System.Drawing.Point(10, 0);
			valueListItem1.DataValue = 0;
			valueListItem1.DisplayText = "Espera";
			valueListItem2.DataValue = 1;
			valueListItem2.DisplayText = "Activo";
			valueListItem3.DataValue = 2;
			valueListItem3.DisplayText = "Finalizado";
			this.optEstado.Items.Add(valueListItem1);
			this.optEstado.Items.Add(valueListItem2);
			this.optEstado.Items.Add(valueListItem3);
			this.optEstado.ItemSpacingVertical = 8;
			this.optEstado.Location = new System.Drawing.Point(88, 88);
			this.optEstado.Name = "optEstado";
			this.optEstado.Size = new System.Drawing.Size(224, 26);
			this.optEstado.TabIndex = 958;
			this.optEstado.Text = "Espera";
			this.optEstado.ValueChanged += new System.EventHandler(this.optEstado_ValueChanged);
			// 
			// groupBox3
			// 
			this.groupBox3.Controls.Add(this.lblContador);
			this.groupBox3.Controls.Add(this.txtestado);
			this.groupBox3.Controls.Add(this.label4);
			this.groupBox3.Controls.Add(this.txtDescripcion);
			this.groupBox3.Controls.Add(this.label2);
			this.groupBox3.Controls.Add(this.dtHasta);
			this.groupBox3.Controls.Add(this.optEstado);
			this.groupBox3.Controls.Add(this.label1);
			this.groupBox3.Controls.Add(this.dtDesde);
			this.groupBox3.Controls.Add(this.txtidFacturacionPromociones);
			this.groupBox3.Location = new System.Drawing.Point(8, 0);
			this.groupBox3.Name = "groupBox3";
			this.groupBox3.Size = new System.Drawing.Size(688, 128);
			this.groupBox3.TabIndex = 961;
			this.groupBox3.TabStop = false;
			this.groupBox3.Enter += new System.EventHandler(this.groupBox3_Enter);
			// 
			// lblContador
			// 
			this.lblContador.AutoSize = true;
			this.lblContador.Font = new System.Drawing.Font("Tahoma", 12F, System.Drawing.FontStyle.Bold);
			this.lblContador.ForeColor = System.Drawing.Color.Firebrick;
			this.lblContador.Location = new System.Drawing.Point(336, 88);
			this.lblContador.Name = "lblContador";
			this.lblContador.Size = new System.Drawing.Size(0, 23);
			this.lblContador.TabIndex = 960;
			this.lblContador.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// txtestado
			// 
			this.txtestado.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtestado.Enabled = false;
			this.txtestado.Location = new System.Drawing.Point(600, 48);
			this.txtestado.MaxValue = 3600;
			this.txtestado.MinValue = 0;
			this.txtestado.Name = "txtestado";
			this.txtestado.PromptChar = ' ';
			this.txtestado.Size = new System.Drawing.Size(72, 22);
			this.txtestado.TabIndex = 959;
			this.txtestado.Visible = false;
			// 
			// label3
			// 
			this.label3.AutoSize = true;
			this.label3.BackColor = System.Drawing.Color.Transparent;
			this.label3.Location = new System.Drawing.Point(16, 202);
			this.label3.Name = "label3";
			this.label3.Size = new System.Drawing.Size(40, 17);
			this.label3.TabIndex = 969;
			this.label3.Text = "Modelo";
			this.label3.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label6
			// 
			this.label6.AutoSize = true;
			this.label6.BackColor = System.Drawing.Color.Transparent;
			this.label6.Location = new System.Drawing.Point(16, 138);
			this.label6.Name = "label6";
			this.label6.Size = new System.Drawing.Size(35, 17);
			this.label6.TabIndex = 964;
			this.label6.Text = "Grupo";
			this.label6.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label47
			// 
			this.label47.AutoSize = true;
			this.label47.BackColor = System.Drawing.Color.Transparent;
			this.label47.Location = new System.Drawing.Point(16, 170);
			this.label47.Name = "label47";
			this.label47.Size = new System.Drawing.Size(34, 17);
			this.label47.TabIndex = 963;
			this.label47.Text = "Marca";
			this.label47.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// cmbModelo
			// 
			appearance6.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbModelo.Appearance = appearance6;
			this.cmbModelo.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbModelo.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbModelo.DataSource = this.ultraDataSource5;
			this.cmbModelo.DisplayLayout.AutoFitColumns = true;
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn1.Width = 241;
			ultraGridBand1.Columns.AddRange(new object[] {
																										 ultraGridColumn1});
			this.cmbModelo.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.cmbModelo.DisplayMember = "Modelo";
			this.cmbModelo.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbModelo.Enabled = false;
			this.cmbModelo.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbModelo.Location = new System.Drawing.Point(88, 200);
			this.cmbModelo.Name = "cmbModelo";
			this.cmbModelo.Size = new System.Drawing.Size(260, 21);
			this.cmbModelo.TabIndex = 968;
			this.cmbModelo.ValueMember = "";
			this.cmbModelo.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cmbModelo_KeyDown);
			// 
			// ultraDataSource5
			// 
			this.ultraDataSource5.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn1});
			// 
			// cmbSubGrupo
			// 
			appearance7.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbSubGrupo.Appearance = appearance7;
			this.cmbSubGrupo.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbSubGrupo.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbSubGrupo.DataSource = this.ultraDataSource1;
			ultraGridColumn2.Header.VisiblePosition = 0;
			ultraGridColumn2.Hidden = true;
			ultraGridColumn3.Header.VisiblePosition = 1;
			ultraGridColumn3.Width = 245;
			ultraGridBand2.Columns.AddRange(new object[] {
																										 ultraGridColumn2,
																										 ultraGridColumn3});
			this.cmbSubGrupo.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			this.cmbSubGrupo.DisplayMember = "Subgrupo";
			this.cmbSubGrupo.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbSubGrupo.Enabled = false;
			this.cmbSubGrupo.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbSubGrupo.Location = new System.Drawing.Point(88, 136);
			this.cmbSubGrupo.Name = "cmbSubGrupo";
			this.cmbSubGrupo.Size = new System.Drawing.Size(260, 21);
			this.cmbSubGrupo.TabIndex = 965;
			this.cmbSubGrupo.ValueMember = "idSubgrupo";
			this.cmbSubGrupo.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cmbSubGrupo_KeyDown);
			this.cmbSubGrupo.ValueChanged += new System.EventHandler(this.cmbSubGrupo_ValueChanged);
			// 
			// ultraDataSource1
			// 
			ultraDataColumn2.DataType = typeof(int);
			this.ultraDataSource1.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn2,
																																 ultraDataColumn3});
			// 
			// ultraDataSource2
			// 
			ultraDataColumn4.DataType = typeof(int);
			this.ultraDataSource2.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn4,
																																 ultraDataColumn5});
			// 
			// ultraDataSource3
			// 
			ultraDataColumn6.DataType = typeof(int);
			this.ultraDataSource3.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn6,
																																 ultraDataColumn7});
			// 
			// txtPorcentaje
			// 
			this.txtPorcentaje.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtPorcentaje.Enabled = false;
			this.txtPorcentaje.FormatString = "#,##0.00";
			this.txtPorcentaje.Location = new System.Drawing.Point(88, 232);
			this.txtPorcentaje.MaxValue = 5.02;
			this.txtPorcentaje.MinValue = 0;
			this.txtPorcentaje.Name = "txtPorcentaje";
			this.txtPorcentaje.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtPorcentaje.PromptChar = ' ';
			this.txtPorcentaje.Size = new System.Drawing.Size(100, 22);
			this.txtPorcentaje.TabIndex = 970;
			this.txtPorcentaje.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtPorcentaje_KeyDown);
			// 
			// label5
			// 
			this.label5.AutoSize = true;
			this.label5.BackColor = System.Drawing.Color.Transparent;
			this.label5.Location = new System.Drawing.Point(16, 235);
			this.label5.Name = "label5";
			this.label5.Size = new System.Drawing.Size(61, 17);
			this.label5.TabIndex = 971;
			this.label5.Text = "Porcentaje ";
			this.label5.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// groupBox2
			// 
			this.groupBox2.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.groupBox2.Location = new System.Drawing.Point(0, 264);
			this.groupBox2.Name = "groupBox2";
			this.groupBox2.Size = new System.Drawing.Size(704, 8);
			this.groupBox2.TabIndex = 972;
			this.groupBox2.TabStop = false;
			// 
			// btnAgregar
			// 
			this.btnAgregar.Enabled = false;
			this.btnAgregar.Image = ((System.Drawing.Image)(resources.GetObject("btnAgregar.Image")));
			this.btnAgregar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnAgregar.Location = new System.Drawing.Point(268, 231);
			this.btnAgregar.Name = "btnAgregar";
			this.btnAgregar.Size = new System.Drawing.Size(80, 24);
			this.btnAgregar.TabIndex = 973;
			this.btnAgregar.Text = "Agregar";
			this.btnAgregar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnAgregar.Click += new System.EventHandler(this.btnAgregar_Click);
			// 
			// uGridInformacion
			// 
			this.uGridInformacion.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.uGridInformacion.Cursor = System.Windows.Forms.Cursors.Default;
			this.uGridInformacion.DataSource = this.ultraDataSource4;
			appearance8.BackColor = System.Drawing.Color.White;
			this.uGridInformacion.DisplayLayout.Appearance = appearance8;
			ultraGridColumn4.Header.VisiblePosition = 0;
			ultraGridColumn4.Hidden = true;
			ultraGridColumn5.Header.VisiblePosition = 1;
			ultraGridColumn5.Hidden = true;
			ultraGridColumn6.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn6.Header.Caption = "Grupo";
			ultraGridColumn6.Header.VisiblePosition = 2;
			ultraGridColumn6.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDown;
			ultraGridColumn6.Width = 195;
			ultraGridColumn7.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn7.Header.Caption = "Marca";
			ultraGridColumn7.Header.VisiblePosition = 3;
			ultraGridColumn7.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDown;
			ultraGridColumn7.Width = 195;
			ultraGridColumn8.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn8.Header.VisiblePosition = 4;
			ultraGridColumn8.Width = 195;
			ultraGridColumn9.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance9.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn9.CellAppearance = appearance9;
			ultraGridColumn9.Format = "#,##0.00";
			ultraGridColumn9.Header.VisiblePosition = 5;
			ultraGridColumn9.Width = 82;
			ultraGridBand3.Columns.AddRange(new object[] {
																										 ultraGridColumn4,
																										 ultraGridColumn5,
																										 ultraGridColumn6,
																										 ultraGridColumn7,
																										 ultraGridColumn8,
																										 ultraGridColumn9});
			ultraGridBand3.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.True;
			appearance10.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings1.Appearance = appearance10;
			summarySettings1.DisplayFormat = "{0: #,##0.00}";
			summarySettings1.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			ultraGridBand3.Summaries.AddRange(new Infragistics.Win.UltraWinGrid.SummarySettings[] {
																																															summarySettings1});
			ultraGridBand3.SummaryFooterCaption = "Totales:";
			this.uGridInformacion.DisplayLayout.BandsSerializer.Add(ultraGridBand3);
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
			this.uGridInformacion.Location = new System.Drawing.Point(8, 280);
			this.uGridInformacion.Name = "uGridInformacion";
			this.uGridInformacion.Size = new System.Drawing.Size(688, 264);
			this.uGridInformacion.TabIndex = 974;
			this.uGridInformacion.AfterRowInsert += new Infragistics.Win.UltraWinGrid.RowEventHandler(this.uGridInformacion_AfterRowInsert);
			this.uGridInformacion.BeforeRowsDeleted += new Infragistics.Win.UltraWinGrid.BeforeRowsDeletedEventHandler(this.uGridInformacion_BeforeRowsDeleted);
			this.uGridInformacion.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.uGridInformacion_InitializeLayout);
			// 
			// ultraDataSource4
			// 
			ultraDataColumn8.DataType = typeof(int);
			ultraDataColumn9.DataType = typeof(int);
			ultraDataColumn10.DataType = typeof(int);
			ultraDataColumn11.DataType = typeof(int);
			ultraDataColumn13.DataType = typeof(System.Decimal);
			this.ultraDataSource4.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn8,
																																 ultraDataColumn9,
																																 ultraDataColumn10,
																																 ultraDataColumn11,
																																 ultraDataColumn12,
																																 ultraDataColumn13});
			// 
			// cmbSubGrupoDet
			// 
			this.cmbSubGrupoDet.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbSubGrupoDet.DataSource = this.ultraDataSource1;
			this.cmbSubGrupoDet.DisplayLayout.AutoFitColumns = true;
			ultraGridColumn10.Header.VisiblePosition = 0;
			ultraGridColumn10.Hidden = true;
			ultraGridColumn10.Width = 200;
			ultraGridColumn11.Header.VisiblePosition = 1;
			ultraGridColumn11.Width = 98;
			ultraGridBand4.Columns.AddRange(new object[] {
																										 ultraGridColumn10,
																										 ultraGridColumn11});
			ultraGridBand4.SummaryFooterCaption = "Grand Summaries";
			this.cmbSubGrupoDet.DisplayLayout.BandsSerializer.Add(ultraGridBand4);
			this.cmbSubGrupoDet.DisplayMember = "SubGrupo";
			this.cmbSubGrupoDet.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.cmbSubGrupoDet.Location = new System.Drawing.Point(368, 136);
			this.cmbSubGrupoDet.Name = "cmbSubGrupoDet";
			this.cmbSubGrupoDet.Size = new System.Drawing.Size(100, 104);
			this.cmbSubGrupoDet.TabIndex = 975;
			this.cmbSubGrupoDet.ValueMember = "idSubgrupo";
			this.cmbSubGrupoDet.Visible = false;
			// 
			// cmbMarcaDet
			// 
			this.cmbMarcaDet.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbMarcaDet.DataSource = this.ultraDataSource2;
			this.cmbMarcaDet.DisplayLayout.AutoFitColumns = true;
			ultraGridColumn12.Header.VisiblePosition = 0;
			ultraGridColumn12.Hidden = true;
			ultraGridColumn12.Width = 200;
			ultraGridColumn13.Header.VisiblePosition = 1;
			ultraGridColumn13.Width = 98;
			ultraGridBand5.Columns.AddRange(new object[] {
																										 ultraGridColumn12,
																										 ultraGridColumn13});
			ultraGridBand5.SummaryFooterCaption = "Grand Summaries";
			this.cmbMarcaDet.DisplayLayout.BandsSerializer.Add(ultraGridBand5);
			this.cmbMarcaDet.DisplayMember = "Marca";
			this.cmbMarcaDet.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.cmbMarcaDet.Location = new System.Drawing.Point(480, 136);
			this.cmbMarcaDet.Name = "cmbMarcaDet";
			this.cmbMarcaDet.Size = new System.Drawing.Size(100, 104);
			this.cmbMarcaDet.TabIndex = 976;
			this.cmbMarcaDet.ValueMember = "idMarca";
			this.cmbMarcaDet.Visible = false;
			// 
			// cmbModeloDet
			// 
			this.cmbModeloDet.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbModeloDet.DataSource = this.ultraDataSource3;
			this.cmbModeloDet.DisplayLayout.AutoFitColumns = true;
			ultraGridColumn14.Header.VisiblePosition = 0;
			ultraGridColumn14.Hidden = true;
			ultraGridColumn14.Width = 38;
			ultraGridColumn15.Header.VisiblePosition = 1;
			ultraGridColumn15.Width = 98;
			ultraGridBand6.Columns.AddRange(new object[] {
																										 ultraGridColumn14,
																										 ultraGridColumn15});
			ultraGridBand6.SummaryFooterCaption = "Grand Summaries";
			this.cmbModeloDet.DisplayLayout.BandsSerializer.Add(ultraGridBand6);
			this.cmbModeloDet.DisplayMember = "Modelo";
			this.cmbModeloDet.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.cmbModeloDet.Location = new System.Drawing.Point(592, 136);
			this.cmbModeloDet.Name = "cmbModeloDet";
			this.cmbModeloDet.Size = new System.Drawing.Size(100, 104);
			this.cmbModeloDet.TabIndex = 977;
			this.cmbModeloDet.ValueMember = "idArticulo";
			this.cmbModeloDet.Visible = false;
			// 
			// groupBox1
			// 
			this.groupBox1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.groupBox1.Location = new System.Drawing.Point(0, 544);
			this.groupBox1.Name = "groupBox1";
			this.groupBox1.Size = new System.Drawing.Size(704, 8);
			this.groupBox1.TabIndex = 978;
			this.groupBox1.TabStop = false;
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
			this.btnEditar.Location = new System.Drawing.Point(188, 560);
			this.btnEditar.Name = "btnEditar";
			this.btnEditar.Size = new System.Drawing.Size(80, 23);
			this.btnEditar.TabIndex = 982;
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
			this.btnGuardar.Location = new System.Drawing.Point(272, 560);
			this.btnGuardar.Name = "btnGuardar";
			this.btnGuardar.Size = new System.Drawing.Size(80, 23);
			this.btnGuardar.TabIndex = 981;
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
			this.btnCancelar.Location = new System.Drawing.Point(360, 560);
			this.btnCancelar.Name = "btnCancelar";
			this.btnCancelar.Size = new System.Drawing.Size(80, 23);
			this.btnCancelar.TabIndex = 983;
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
			this.btnSalir.Location = new System.Drawing.Point(446, 560);
			this.btnSalir.Name = "btnSalir";
			this.btnSalir.Size = new System.Drawing.Size(80, 23);
			this.btnSalir.TabIndex = 984;
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
			this.btnNuevo.Location = new System.Drawing.Point(16, 560);
			this.btnNuevo.Name = "btnNuevo";
			this.btnNuevo.Size = new System.Drawing.Size(80, 23);
			this.btnNuevo.TabIndex = 979;
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
			this.btnConsultar.Location = new System.Drawing.Point(102, 560);
			this.btnConsultar.Name = "btnConsultar";
			this.btnConsultar.Size = new System.Drawing.Size(80, 23);
			this.btnConsultar.TabIndex = 980;
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
			// cmbMarca
			// 
			appearance16.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbMarca.Appearance = appearance16;
			this.cmbMarca.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbMarca.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbMarca.DataSource = this.ultraDataSource2;
			ultraGridColumn16.Header.VisiblePosition = 0;
			ultraGridColumn16.Hidden = true;
			ultraGridColumn17.Header.VisiblePosition = 1;
			ultraGridColumn17.Width = 248;
			ultraGridBand7.Columns.AddRange(new object[] {
																										 ultraGridColumn16,
																										 ultraGridColumn17});
			this.cmbMarca.DisplayLayout.BandsSerializer.Add(ultraGridBand7);
			this.cmbMarca.DisplayMember = "Marca";
			this.cmbMarca.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbMarca.Enabled = false;
			this.cmbMarca.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbMarca.Location = new System.Drawing.Point(88, 168);
			this.cmbMarca.Name = "cmbMarca";
			this.cmbMarca.Size = new System.Drawing.Size(260, 21);
			this.cmbMarca.TabIndex = 985;
			this.cmbMarca.ValueMember = "idMarca";
			this.cmbMarca.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cmbMarca_KeyDown_1);
			this.cmbMarca.ValueChanged += new System.EventHandler(this.cmbMarca_ValueChanged_1);
			// 
			// frmFacturaPromociones
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 14);
			this.ClientSize = new System.Drawing.Size(704, 590);
			this.Controls.Add(this.groupBox3);
			this.Controls.Add(this.cmbMarca);
			this.Controls.Add(this.btnEditar);
			this.Controls.Add(this.btnGuardar);
			this.Controls.Add(this.btnCancelar);
			this.Controls.Add(this.btnSalir);
			this.Controls.Add(this.btnNuevo);
			this.Controls.Add(this.btnConsultar);
			this.Controls.Add(this.groupBox1);
			this.Controls.Add(this.cmbModeloDet);
			this.Controls.Add(this.cmbMarcaDet);
			this.Controls.Add(this.cmbSubGrupoDet);
			this.Controls.Add(this.uGridInformacion);
			this.Controls.Add(this.btnAgregar);
			this.Controls.Add(this.groupBox2);
			this.Controls.Add(this.label5);
			this.Controls.Add(this.label3);
			this.Controls.Add(this.label6);
			this.Controls.Add(this.label47);
			this.Controls.Add(this.txtPorcentaje);
			this.Controls.Add(this.cmbModelo);
			this.Controls.Add(this.cmbSubGrupo);
			this.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.KeyPreview = true;
			this.Name = "frmFacturaPromociones";
			this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
			this.Text = "Factura Promociones";
			this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.frmFacturaPromociones_KeyDown);
			this.Load += new System.EventHandler(this.frmFacturaPromociones_Load);
			((System.ComponentModel.ISupportInitialize)(this.txtidFacturacionPromociones)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDescripcion)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtDesde)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtHasta)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.optEstado)).EndInit();
			this.groupBox3.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.txtestado)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbModelo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource5)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbSubGrupo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource3)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtPorcentaje)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridInformacion)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource4)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbSubGrupoDet)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbMarcaDet)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbModeloDet)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbMarca)).EndInit();
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

		private void frmFacturaPromociones_Load(object sender, System.EventArgs e)
		{
			miAcceso = new Acceso(cdsSeteoF, "0629");

			if (!Funcion.Permiso("1205", cdsSeteoF))
			{				
				MessageBox.Show("No puede ingresar a Facturas de Promoción", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				BeginInvoke(new MethodInvoker(UnloadMe));
				return;
			}
			if (miAcceso.Nuevo) this.btnNuevo.Enabled = true;
			if (miAcceso.Buscar) this.btnConsultar.Enabled = true;
			//			this.dtDesde.Value = DateTime.Today;
			//this.dtDesde.CalendarInfo.MinDate = DateTime.Today;

			//			this.dtHasta.Value = DateTime.Today;
			this.dtHasta.CalendarInfo.MinDate = DateTime.Today;

			this.cmbSubGrupo.DataSource = FuncionesProcedimientos.dtGeneral("Select idSubgrupo, SubGrupo From ArticuloSubGrupo Order By SubGrupo");
			this.cmbSubGrupoDet.DataSource = FuncionesProcedimientos.dtGeneral("Select idSubgrupo, SubGrupo From ArticuloSubGrupo Order By SubGrupo");
			this.cmbMarcaDet.DataSource = FuncionesProcedimientos.dtGeneral("Select  am.idMarca, Marca From ArticuloMarca am Inner Join SubGrupoMarca sm On sm.idMarca = am.idMarca " );
			this.cmbModeloDet.DataSource = FuncionesProcedimientos.dtGeneral("Select  idArticulo, Modelo From Articulo ");
		}

		private void cmbSubGrupo_ValueChanged(object sender, System.EventArgs e)
		{
			int idSubgrupo = 0;

			if (this.cmbSubGrupo.ActiveRow != null)
			{
				if (this.cmbSubGrupo.ActiveRow != null)  idSubgrupo = (int)this.cmbSubGrupo.Value;
				this.cmbMarca.DataSource = FuncionesProcedimientos.dtGeneral("Select  am.idMarca, Marca From ArticuloMarca am Inner Join SubGrupoMarca sm On sm.idMarca = am.idMarca Where sm.idSubGrupo = '" + idSubgrupo + "'" );
			}
		}

		private void cmbMarca_ValueChanged(object sender, System.EventArgs e)
		{
			int idModelo = 0;
			int idSubgrupo = 0;
			if (this.cmbMarca.ActiveRow != null)
			{
				if (this.cmbSubGrupo.ActiveRow != null)  idSubgrupo = (int)this.cmbSubGrupo.Value;
				if (this.cmbMarca.ActiveRow != null)  idModelo = (int)this.cmbMarca.Value;
				this.cmbModelo.DataSource = FuncionesProcedimientos.dtGeneral("Select Modelo from Articulo Where  idSubGrupo = '" + idSubgrupo + "' And idMarca = '" + idModelo + "'" );			
			}
		}		

		private void btnCancelar_Click(object sender, System.EventArgs e)
		{
			bNuevo = false;
			bEdicion = false;

			#region Vacia Campos 

			this.txtidFacturacionPromociones.Value = 0;
			this.txtDescripcion.Text = "";
			this.dtDesde.Value = System.DBNull.Value;
			this.dtHasta.Value = System.DBNull.Value;
			this.optEstado.Value = 0;
			this.cmbSubGrupo.Value = System.DBNull.Value;
			this.cmbMarca.Value = System.DBNull.Value;
			this.cmbModelo.Value = System.DBNull.Value;
			this.txtPorcentaje.Value = 0;
			this.txtestado.Value = 0;
			this.lblContador.Text = "";

			FuncionesProcedimientos.EstadoGrids(false, this.uGridInformacion);

			#endregion Vacia Campos 

			#region Bloquea Campo

			this.txtidFacturacionPromociones.Enabled = false;
			this.txtDescripcion.Enabled = false;
			this.dtDesde.Enabled = false;
			this.dtHasta.Enabled = false;
			this.optEstado.Enabled = false;
			this.cmbSubGrupo.Enabled = false;
			this.cmbMarca.Enabled = false;
			this.cmbModelo.Enabled = false;
			this.txtPorcentaje.Enabled = false;

			#endregion Bloquea Campo

			this.uGridInformacion.DataSource = Funcion.dvProcedimiento(cdsSeteoF, string.Format("Exec ConsultaIndividualDetallePromociones 0"));

			this.btnNuevo.Enabled = true;
			this.btnConsultar.Enabled = true;
			this.btnEditar.Enabled = false;
			this.btnGuardar.Enabled = false;
			this.btnCancelar.Enabled = false;
			this.btnAgregar.Enabled = false;
		}

		private void btnSalir_Click(object sender, System.EventArgs e)
		{
			this.Close();		
		}

		private void btnNuevo_Click(object sender, System.EventArgs e)
		{
			this.btnCancelar_Click(sender, e);

			this.dtDesde.Value = DateTime.Today;
			this.dtHasta.Value = DateTime.Today;

			this.txtDescripcion.Enabled = true;
			this.dtHasta.Enabled = true;
			this.dtDesde.Enabled = true;
			this.optEstado.Enabled = true;
			this.txtPorcentaje.Enabled = true;
			FuncionesProcedimientos.EstadoGrids(true, this.uGridInformacion);
			Estado ((int) this.txtestado.Value);

			this.cmbSubGrupo.Enabled = true;
			this.cmbModelo.Enabled = true;
			this.cmbMarca.Enabled = true;

			this.btnNuevo.Enabled = false;
			this.btnConsultar.Enabled = false;
			this.btnEditar.Enabled = false;
			this.btnGuardar.Enabled = true;
			this.btnCancelar.Enabled = true;	
			this.btnAgregar.Enabled = true;
			
			bNuevo = true;
			bEdicion = true;
			
			this.txtDescripcion.Focus();			
		}

		private void Estado ( int IEstado)
		{
			if (IEstado == 0)
				this.lblContador.Text = " PENDIENTE";
			
			if (IEstado == 1)
				this.lblContador.Text = " ACTIVO";
			
			if (IEstado == 2)
				this.lblContador.Text = " FINALIZADO";
		}

		private void btnConsultar_Click(object sender, System.EventArgs e)
		{
			this.btnCancelar_Click(sender, e);

			using (frmConsultaFacturaPromocion consultab = new frmConsultaFacturaPromocion())
				if (DialogResult.OK == consultab.ShowDialog())
					this.Consultar((int)consultab.uGridInformacion.ActiveRow.Cells["idFacturacionPromociones"].Value);
		}

		private void Consultar (int IdFacturacionPromociones )
		{
			#region Consulta Individual
			string ssql= string.Format ("Exec ConsultaIndividualFacturaPromociones {0}",IdFacturacionPromociones); 
			SqlDataReader dr = Funcion.rEscalarSQL(cdsSeteoF,ssql,true);
			dr.Read();
			this.txtidFacturacionPromociones.Value= dr.GetInt32 (0);
			if (dr.GetValue(1) != System.DBNull.Value)this.txtDescripcion.Text = dr.GetString (1);
			if (dr.GetValue(2) != System.DBNull.Value)this.dtDesde.Value = dr.GetDateTime (2);
			if (dr.GetValue(3) != System.DBNull.Value)this.dtHasta.Value = dr.GetDateTime (3);
			if (dr.GetValue(4) != System.DBNull.Value)this.optEstado.Value = dr.GetInt32 (4);
			if (dr.GetValue(4) != System.DBNull.Value)this.txtestado.Value = dr.GetInt32 (4);
			
			dr.Close();

			this.uGridInformacion.DataSource = Funcion.dvProcedimiento(cdsSeteoF, string.Format("Exec ConsultaIndividualDetallePromociones {0}", IdFacturacionPromociones));
			#endregion Consulta Individual

			Estado ((int) this.txtestado.Value);

			if (miAcceso.Nuevo)this.btnNuevo.Enabled = true;
			if (miAcceso.Buscar) this.btnConsultar.Enabled = true;
			if (miAcceso.Editar && (int)this.txtestado.Value == 0) this.btnEditar.Enabled = true;
			this.btnGuardar.Enabled = false;
			this.btnCancelar.Enabled = true;
		}

		private void btnEditar_Click(object sender, System.EventArgs e)
		{    
			if ( (int)this.optEstado.Value == 2 )
			{
				MessageBox.Show("No puede Editar una promoción Finalizada", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return;
			}
			
			this.txtDescripcion.Enabled = true;
			this.dtHasta.Enabled = true;
			this.dtDesde.Enabled = true;
			this.optEstado.Enabled = true;
			this.txtPorcentaje.Enabled = true;

			this.cmbSubGrupo.Enabled = true;
			this.cmbModelo.Enabled = true;
			this.cmbMarca.Enabled = true;
		
			this.btnNuevo.Enabled = false;
			this.btnConsultar.Enabled = false;
			this.btnEditar.Enabled = false;
			this.btnGuardar.Enabled = true;
			this.btnCancelar.Enabled = true;
			this.btnAgregar.Enabled = true;

			bEdicion=true;
			this.txtDescripcion.Focus();
		}

		private void btnGuardar_Click(object sender, System.EventArgs e)
		{
			#region Validacion
			if(Convert.ToDateTime (this.dtHasta.Value) == Convert.ToDateTime (DateTime.Today))
			{
				MessageBox.Show("La fecha de inicio debe ser distinta a la actual. ", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return;
			}
			if (!Validacion.vbTexto(this.txtDescripcion, 5, 100, "Descripcion")) return;
			if (!Validacion.vbFechaEnDocumentos(this.dtDesde, "Ingrese la fecha Inicial ", true, "VENTA", cdsSeteoF)) return;
			if (!Validacion.vbFechaEnDocumentos(this.dtHasta, "Ingrese la fecha Final ", true, "VENTA", cdsSeteoF)) return;
			if (!Validacion.vbComparaFechas(Convert.ToDateTime(this.dtDesde.Value), Convert.ToDateTime(this.dtHasta.Value), this.dtDesde, "La Fecha Inicial no puede ser mayor a la Fecha Final", false)) return;
			
			if ( (int)this.uGridInformacion.Rows.Count == 0)
			{
				MessageBox.Show("Ingrese los Articulos de la Promoción ", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return;
			}		
			#endregion Validacion
      
			#region Estado
			if((int) this.txtidFacturacionPromociones.Value > 0 )
			{
				if((int) this.txtestado.Value == 1 && (int) this.optEstado.Value == 0 )
				{
					MessageBox.Show(" La promoción esta activa por favor Finalizarla ", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					return;
				}

				if((int) this.txtestado.Value == 2 && (int) this.optEstado.Value < 2 )
				{
					MessageBox.Show(" La promoción esta Finalizada por favor crear un nueva promoción  ", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					return;
				}
			}
			#endregion Estado

			string sModelo = "";
			int iCont = 0;

			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridInformacion.Rows.All)
			{         
				sModelo = Convert.ToString(dr.Cells["Modelo"].Value); 

				foreach(Infragistics.Win.UltraWinGrid.UltraGridRow drr in this.uGridInformacion.Rows.All)
				{
					if ( sModelo == Convert.ToString(drr.Cells["Modelo"].Value))
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
				this.txtPorcentaje.Focus ();
				return;
			}

			int iEstado = (int)this.optEstado.Value;
			if (bNuevo) iEstado = 0;
			if (!bNuevo && (int)this.optEstado.Value == 2) iEstado = 1;

			#region Guarda Maestro
			string sSQLl = string.Format("Exec GuardaFacturaPromociones {0}, '{1}' , '{2}' ,'{3}', {4}",
				(int) this.txtidFacturacionPromociones.Value,
				this.txtDescripcion.Text,
				Convert.ToDateTime(this.dtDesde.Value).ToString("yyyyMMdd"),
				Convert.ToDateTime(this.dtHasta.Value).ToString("yyyyMMdd"),
				iEstado);
			#endregion Guarda Maestro

			this.txtidFacturacionPromociones.Value=Funcion.iEscalarSQL(cdsSeteoF,sSQLl);

			#region Guarda detalle Maestro
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow drr in this.uGridInformacion.Rows.All)
			{	
				string sSQLDetalle = string.Format("Exec GuardaDetalleFacturaDetallePromociones  {0}, {1}, {2}, {3}, '{4}', {5},0", 
					(int)drr.Cells["idDetallePromociones"].Value,
					(int)this.txtidFacturacionPromociones.Value,
					(int)drr.Cells["idSubgrupo"].Value,
					(int)drr.Cells["idMarca"].Value,
					drr.Cells["Modelo"].Value,
					(decimal)drr.Cells["Porcentaje"].Value);
				Funcion.EjecutaSQL(cdsSeteoF, sSQLDetalle);
			}	
			#endregion Guarda detalle Maestro

			if ((int)this.optEstado.Value == 1)
			{
				string sSQL = string.Format("Exec ActualizaPorcentajesCuotaBarataza '{0}', {1}, {2},0", System.Convert.ToDateTime(this.dtDesde.Value).ToString("yyyyMMdd"), 1, (int)this.txtidFacturacionPromociones.Value);
				Funcion.EjecutaSQL(cdsSeteoF, sSQL, true);       
			}
			else if ((int)this.optEstado.Value == 2)
			{
				string sSQL = string.Format("Exec ActualizaPorcentajesCuotaBarataza '{0}', {1}, {2},0", System.Convert.ToDateTime(this.dtHasta.Value).ToString("yyyyMMdd"), 2, (int)this.txtidFacturacionPromociones.Value);
				Funcion.EjecutaSQL(cdsSeteoF, sSQL, true);
			}

			#region Controles
			Consultar ((int) this.txtidFacturacionPromociones.Value);
			this.txtidFacturacionPromociones.Enabled = false;
			this.txtDescripcion.Enabled = false;
			this.dtDesde.Enabled = false;
			this.dtHasta.Enabled = false;
			this.optEstado.Enabled = false;
			this.cmbSubGrupo.Enabled = false;
			this.cmbMarca.Enabled = false;
			this.cmbModelo.Enabled = false;
			this.txtPorcentaje.Enabled = false;
			FuncionesProcedimientos.EstadoGrids(false, this.uGridInformacion);

			if (miAcceso.Nuevo) this.btnNuevo.Enabled = true;
			if (miAcceso.Buscar) this.btnConsultar.Enabled = true;
			if (miAcceso.Editar) this.btnEditar.Enabled = true;
			this.btnGuardar.Enabled = false;
			this.btnCancelar.Enabled = true;
			this.btnAgregar.Enabled = false;

			
			bNuevo = false;
			bEdicion = false;
			#endregion Controles			
		}

		private void uGridInformacion_AfterRowInsert(object sender, Infragistics.Win.UltraWinGrid.RowEventArgs e)
		{
			e.Row.Cells["idDetallePromociones"].Value = 0;
			e.Row.Cells["idFacturacionPromociones"].Value = 1;
			e.Row.Cells["idSubgrupo"].Value = 0;
			e.Row.Cells["idMarca"].Value = 0; 
			e.Row.Cells["Modelo"].Value = ""; 	
			e.Row.Cells["Porcentaje"].Value = 0.00m; 
		}

		private void btnAgregar_Click(object sender, System.EventArgs e)
		{
			if (!Validacion.vbTexto(this.txtDescripcion, 5, 100, "Descripcion")) return;
			if (!Validacion.vbFechaEnDocumentos(this.dtDesde, "Ingrese la fecha Inicial ", true, "VENTA", cdsSeteoF)) return;
			if (!Validacion.vbFechaEnDocumentos(this.dtHasta, "Ingrese la fecha Final ", true, "VENTA", cdsSeteoF)) return;
			if (!Validacion.vbComparaFechas(Convert.ToDateTime(this.dtDesde.Value), Convert.ToDateTime(this.dtHasta.Value), this.dtDesde, "La Fecha Inicial no puede ser mayor a la Fecha Final", false)) return;
			
			if (!Validacion.vbComboVacio(this.cmbSubGrupo, "Seleccione el Grupo del Articulo")) return;
			if (!Validacion.vbComboVacio(this.cmbMarca, "Seleccione la Marca ")) return;
			if (!Validacion.vbComboVacio(this.cmbModelo, "Seleccione el Modelo ")) return;

			if (!Validacion.vbCampoDecimalVacio(this.txtPorcentaje, " Ingrese el Porcentaje de Descuento ",0,100)) return;

			if ((double)this.txtPorcentaje.Value > 1.99)
			{
				MessageBox.Show("El porcentaje Maximo es de 1.99 ", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.txtPorcentaje.Focus ();
				return;
			}

			if ((double)this.txtPorcentaje.Value < 00.01)
			{
				MessageBox.Show("El porcentaje Minimo es de 00.01 ", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.txtPorcentaje.Focus ();
				return;
			}

			this.uGridInformacion.Rows.Band.AddNew();
			this.uGridInformacion.ActiveRow.Cells["idDetallePromociones"].Value = 0;
			this.uGridInformacion.ActiveRow.Cells["idFacturacionPromociones"].Value = (int)this.txtidFacturacionPromociones.Value;
			this.uGridInformacion.ActiveRow.Cells["idSubgrupo"].Value = (int) this.cmbSubGrupo.Value;
			this.uGridInformacion.ActiveRow.Cells["idMarca"].Value = (int) this.cmbMarca.Value;
			this.uGridInformacion.ActiveRow.Cells["Modelo"].Value = this.cmbModelo.Value.ToString();	
			this.uGridInformacion.ActiveRow.Cells["Porcentaje"].Value = Convert.ToDecimal (this.txtPorcentaje.Value);	

			this.cmbSubGrupo.Value = System.DBNull.Value;
			this.cmbMarca.Value = System.DBNull.Value;
			this.cmbModelo.Value = System.DBNull.Value;
			this.txtPorcentaje.Value = 0;

			this.txtDescripcion.Enabled = false;
			this.dtDesde.Enabled = false;
			this.dtHasta.Enabled = false;
			this.optEstado.Enabled = false;

			bNuevo = true;
			bEdicion = true;
		}

		private void txtDescripcion_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter) this.dtDesde.Focus();
		}

		private void dtDesde_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter) this.dtHasta.Focus();
		}

		private void dtHasta_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter) this.cmbSubGrupo.Focus();
		}

		private void cmbSubGrupo_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter) this.cmbMarca.Focus();
		}

		private void cmbMarca_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter) this.cmbModelo.Focus();
		}

		private void cmbModelo_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter) this.txtPorcentaje.Focus();
		}

		private void txtPorcentaje_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter) this.btnAgregar.Focus();
		}

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);		
		}

		private void uGridInformacion_BeforeRowsDeleted(object sender, Infragistics.Win.UltraWinGrid.BeforeRowsDeletedEventArgs e)
		{
			if (DialogResult.Yes ==	MessageBox.Show("¿Esta seguro de Borrar las Líneas Seleccionadas?", "Point Technology", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2))				
			{
				for (int i = 0; i < e.Rows.Length; i++)
				{
					if (e.Rows[i].Cells["idDetallePromociones"].Value != System.DBNull.Value)
					{									
						if ((int)e.Rows[i].Cells["idDetallePromociones"].Value > 0)
						{
							string sSQL = string.Format("Delete From DetallePromociones Where idDetallePromociones = {0}", (int)e.Rows[i].Cells["idDetallePromociones"].Value);
							Funcion.EjecutaSQL(cdsSeteoF, sSQL, true);
							e.DisplayPromptMsg = false;
						}
						else e.DisplayPromptMsg = false;
					}		
				}
			}		
			else e.Cancel = true;
		}

		private void frmFacturaPromociones_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (!bEdicion)
			{
				if (e.KeyCode == Keys.N) this.btnNuevo_Click(sender, e);
				if (e.KeyCode == Keys.B) this.btnConsultar_Click(sender, e);
				if ((int)this.txtidFacturacionPromociones.Value > 0)
				{
					if (e.KeyCode == Keys.E) if (this.btnEditar.Enabled) this.btnEditar_Click(sender, e);					
					if (e.KeyCode == Keys.F5) this.Consultar((int)this.txtidFacturacionPromociones.Value);					
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

		private void cmbMarca_KeyDown_1(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter) this.cmbModelo.Focus();
		}

		private void optEstado_ValueChanged(object sender, System.EventArgs e)
		{
			if ((int) this.optEstado.Value == 1)
			{
				this.dtDesde.Value = DateTime.Today;
				this.dtDesde.Enabled = false;
			}
			else this.dtDesde.Enabled = true;
		}

		private void cmbMarca_ValueChanged_1(object sender, System.EventArgs e)
		{
			int idModelo = 0;
			int idSubgrupo = 0;
			if (this.cmbMarca.ActiveRow != null)
			{
				if (this.cmbSubGrupo.ActiveRow != null)  idSubgrupo = (int)this.cmbSubGrupo.Value;
				if (this.cmbMarca.ActiveRow != null)  idModelo = (int)this.cmbMarca.Value;
				this.cmbModelo.DataSource = FuncionesProcedimientos.dtGeneral("Select  Modelo From Articulo Where  idSubGrupo = '" + idSubgrupo + "' And idMarca = '" + idModelo + "'" );
			
			}
		}

		private void uGridInformacion_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}

		private void groupBox3_Enter(object sender, System.EventArgs e)
		{
		
		}
	}
}