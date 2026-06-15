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
	/// Descripción breve de frmProformas.
	/// </summary>
	public class frmProformas : System.Windows.Forms.Form
	{
		public Infragistics.Win.UltraWinEditors.UltraNumericEditor txtidProformas;
		private System.Windows.Forms.Label lblNumero;
		private System.Windows.Forms.Label txtNumero;
		private System.Windows.Forms.Label lblFecha;
		public Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtFecha;
		private System.Windows.Forms.Label lblDetalle;
		private System.Windows.Forms.Label lblDescuento;
		private System.Windows.Forms.Label label1;
		private System.Windows.Forms.Label lblCosto;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtDescuento;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtTotal;
		private System.Windows.Forms.Label lblFechaInicia;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtFechaInicia;
		public Infragistics.Win.UltraWinEditors.UltraNumericEditor txtEstado;
		private System.Windows.Forms.Button btnNuevo;
		private System.Windows.Forms.Button btnConsultar;
		private System.Windows.Forms.Button btnEditar;
		private System.Windows.Forms.Button btnCancelar;
		private System.Windows.Forms.Button btnAnular;
		private System.Windows.Forms.Button btnSalir;
		private System.Windows.Forms.GroupBox groupBox3;
		public Infragistics.Win.UltraWinEditors.UltraTextEditor txtDetalle;
		private System.Windows.Forms.Label lblEstado;
		private C1.Data.C1DataSet cdsSeteoF;
		private System.Windows.Forms.Button btnGuardar;
		private System.Windows.Forms.Button btnAutorizar;
		private System.Windows.Forms.Label lblIdentificacion;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtRuc;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource1;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtIdCliente;
		private System.Windows.Forms.Label lblCliente;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtNombre;
		public Infragistics.Win.UltraWinGrid.UltraGrid ultraGrid1;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtidContrato;
		private System.Windows.Forms.GroupBox groupBox1;
		private System.Windows.Forms.GroupBox groupBox2;
		private System.Windows.Forms.GroupBox groupBox4;
		private System.Windows.Forms.Button btnArchivoContrato;
		private System.Windows.Forms.Label lbltitulo;
		private System.Windows.Forms.Label label2;
		private System.Windows.Forms.Label label3;
		public Infragistics.Win.UltraWinEditors.UltraTextEditor txtTitulo;
		private System.Windows.Forms.Button btnArchivoProforma;
		public Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtFechaContrato;
		public Infragistics.Win.UltraWinEditors.UltraTextEditor txtNotasContrato;
		private Infragistics.Win.UltraWinEditors.UltraOptionSet optContrato;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtCosto;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public frmProformas()
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
	
		bool bNuevo = false;
		bool bEdicion = false;		
		string sArchivoProforma;
		string sArchivoContrato;		

		private bool bActualiza = true;
     private Acceso miAcceso;
		private void UnloadMe()
		{
			this.Close();
		}

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
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmProformas));
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.ValueListItem valueListItem1 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem2 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton2 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance11 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idFechaentregaProforma");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idProforma");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn3 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Fecha");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn4 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Descripcion");
			Infragistics.Win.Appearance appearance12 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance13 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idFechaentregaProforma");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idProforma");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Fecha");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Descripcion");
			Infragistics.Win.Appearance appearance14 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance15 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance16 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance17 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance18 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance19 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance20 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton3 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			this.txtidProformas = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.lblNumero = new System.Windows.Forms.Label();
			this.txtNumero = new System.Windows.Forms.Label();
			this.lblFecha = new System.Windows.Forms.Label();
			this.dtFecha = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.txtDetalle = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.lblDetalle = new System.Windows.Forms.Label();
			this.lblCosto = new System.Windows.Forms.Label();
			this.txtCosto = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.lblDescuento = new System.Windows.Forms.Label();
			this.txtDescuento = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label1 = new System.Windows.Forms.Label();
			this.txtTotal = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.optContrato = new Infragistics.Win.UltraWinEditors.UltraOptionSet();
			this.lblFechaInicia = new System.Windows.Forms.Label();
			this.dtFechaInicia = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.txtEstado = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.btnNuevo = new System.Windows.Forms.Button();
			this.btnConsultar = new System.Windows.Forms.Button();
			this.btnEditar = new System.Windows.Forms.Button();
			this.btnGuardar = new System.Windows.Forms.Button();
			this.btnCancelar = new System.Windows.Forms.Button();
			this.btnAnular = new System.Windows.Forms.Button();
			this.btnSalir = new System.Windows.Forms.Button();
			this.groupBox3 = new System.Windows.Forms.GroupBox();
			this.lblEstado = new System.Windows.Forms.Label();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.btnAutorizar = new System.Windows.Forms.Button();
			this.lblIdentificacion = new System.Windows.Forms.Label();
			this.txtRuc = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.btnArchivoContrato = new System.Windows.Forms.Button();
			this.ultraDataSource1 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.txtIdCliente = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.lblCliente = new System.Windows.Forms.Label();
			this.txtNombre = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.ultraGrid1 = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.txtidContrato = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.groupBox1 = new System.Windows.Forms.GroupBox();
			this.groupBox2 = new System.Windows.Forms.GroupBox();
			this.groupBox4 = new System.Windows.Forms.GroupBox();
			this.lbltitulo = new System.Windows.Forms.Label();
			this.label2 = new System.Windows.Forms.Label();
			this.label3 = new System.Windows.Forms.Label();
			this.txtTitulo = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtNotasContrato = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.dtFechaContrato = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.btnArchivoProforma = new System.Windows.Forms.Button();
			((System.ComponentModel.ISupportInitialize)(this.txtidProformas)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFecha)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDetalle)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCosto)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDescuento)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtTotal)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.optContrato)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFechaInicia)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtEstado)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtRuc)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdCliente)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNombre)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtidContrato)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtTitulo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNotasContrato)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFechaContrato)).BeginInit();
			this.SuspendLayout();
			// 
			// txtidProformas
			// 
			appearance1.ForeColor = System.Drawing.Color.Black;
			appearance1.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtidProformas.Appearance = appearance1;
			this.txtidProformas.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtidProformas.Enabled = false;
			this.txtidProformas.FormatString = "###0";
			this.txtidProformas.Location = new System.Drawing.Point(368, 7);
			this.txtidProformas.MaxValue = 100000000;
			this.txtidProformas.MinValue = 0;
			this.txtidProformas.Name = "txtidProformas";
			this.txtidProformas.PromptChar = ' ';
			this.txtidProformas.Size = new System.Drawing.Size(40, 22);
			this.txtidProformas.TabIndex = 699;
			this.txtidProformas.Visible = false;
			// 
			// lblNumero
			// 
			this.lblNumero.AutoSize = true;
			this.lblNumero.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblNumero.Location = new System.Drawing.Point(472, 10);
			this.lblNumero.Name = "lblNumero";
			this.lblNumero.Size = new System.Drawing.Size(44, 17);
			this.lblNumero.TabIndex = 701;
			this.lblNumero.Text = "Numero";
			this.lblNumero.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtNumero
			// 
			this.txtNumero.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
			this.txtNumero.Font = new System.Drawing.Font("Tahoma", 10F, System.Drawing.FontStyle.Bold);
			this.txtNumero.ForeColor = System.Drawing.Color.Firebrick;
			this.txtNumero.Location = new System.Drawing.Point(536, 7);
			this.txtNumero.Name = "txtNumero";
			this.txtNumero.Size = new System.Drawing.Size(168, 22);
			this.txtNumero.TabIndex = 700;
			this.txtNumero.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// lblFecha
			// 
			this.lblFecha.AutoSize = true;
			this.lblFecha.BackColor = System.Drawing.Color.Transparent;
			this.lblFecha.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblFecha.Location = new System.Drawing.Point(16, 10);
			this.lblFecha.Name = "lblFecha";
			this.lblFecha.Size = new System.Drawing.Size(34, 17);
			this.lblFecha.TabIndex = 703;
			this.lblFecha.Text = "Fecha";
			this.lblFecha.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// dtFecha
			// 
			appearance2.FontData.Name = "Tahoma";
			appearance2.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtFecha.Appearance = appearance2;
			this.dtFecha.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton1.Action = Infragistics.Win.UltraWinSchedule.DateButtonAction.AddToSelection;
			dateButton1.Caption = "Today";
			dateButton1.Type = Infragistics.Win.UltraWinSchedule.DateButtonType.Custom;
			this.dtFecha.DateButtons.Add(dateButton1);
			this.dtFecha.Enabled = false;
			this.dtFecha.Format = "dd/MM/yyyy";
			this.dtFecha.Location = new System.Drawing.Point(104, 8);
			this.dtFecha.Name = "dtFecha";
			this.dtFecha.NonAutoSizeHeight = 23;
			this.dtFecha.Size = new System.Drawing.Size(104, 21);
			this.dtFecha.SpinButtonsVisible = true;
			this.dtFecha.TabIndex = 702;
			this.dtFecha.Value = ((object)(resources.GetObject("dtFecha.Value")));
			this.dtFecha.KeyDown += new System.Windows.Forms.KeyEventHandler(this.dtFecha_KeyDown);
			this.dtFecha.ValueChanged += new System.EventHandler(this.dtFecha_ValueChanged);
			// 
			// txtDetalle
			// 
			this.txtDetalle.AcceptsReturn = true;
			appearance3.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtDetalle.Appearance = appearance3;
			this.txtDetalle.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtDetalle.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtDetalle.Enabled = false;
			this.txtDetalle.Location = new System.Drawing.Point(104, 72);
			this.txtDetalle.MaxLength = 1000;
			this.txtDetalle.Multiline = true;
			this.txtDetalle.Name = "txtDetalle";
			this.txtDetalle.Size = new System.Drawing.Size(536, 66);
			this.txtDetalle.TabIndex = 704;
			this.txtDetalle.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtDetalle_KeyDown);
			// 
			// lblDetalle
			// 
			this.lblDetalle.AutoSize = true;
			this.lblDetalle.BackColor = System.Drawing.Color.Transparent;
			this.lblDetalle.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblDetalle.Location = new System.Drawing.Point(16, 97);
			this.lblDetalle.Name = "lblDetalle";
			this.lblDetalle.Size = new System.Drawing.Size(39, 17);
			this.lblDetalle.TabIndex = 705;
			this.lblDetalle.Text = "Detalle";
			this.lblDetalle.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// lblCosto
			// 
			this.lblCosto.AutoSize = true;
			this.lblCosto.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblCosto.Location = new System.Drawing.Point(16, 163);
			this.lblCosto.Name = "lblCosto";
			this.lblCosto.Size = new System.Drawing.Size(45, 17);
			this.lblCosto.TabIndex = 706;
			this.lblCosto.Text = "Subtotal";
			this.lblCosto.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtCosto
			// 
			appearance4.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtCosto.Appearance = appearance4;
			this.txtCosto.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtCosto.Enabled = false;
			this.txtCosto.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtCosto.FormatString = "#,##0.00";
			this.txtCosto.Location = new System.Drawing.Point(104, 160);
			this.txtCosto.Name = "txtCosto";
			this.txtCosto.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtCosto.PromptChar = ' ';
			this.txtCosto.Size = new System.Drawing.Size(100, 22);
			this.txtCosto.TabIndex = 707;
			this.txtCosto.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtcosto_KeyDown);
			this.txtCosto.Click += new System.EventHandler(this.txtcosto_Click);
			this.txtCosto.Validated += new System.EventHandler(this.txtcosto_Validated);
			this.txtCosto.Enter += new System.EventHandler(this.txtcosto_Enter);
			// 
			// lblDescuento
			// 
			this.lblDescuento.AutoSize = true;
			this.lblDescuento.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblDescuento.Location = new System.Drawing.Point(232, 163);
			this.lblDescuento.Name = "lblDescuento";
			this.lblDescuento.Size = new System.Drawing.Size(71, 17);
			this.lblDescuento.TabIndex = 708;
			this.lblDescuento.Text = "% Descuento";
			this.lblDescuento.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtDescuento
			// 
			appearance5.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtDescuento.Appearance = appearance5;
			this.txtDescuento.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtDescuento.Enabled = false;
			this.txtDescuento.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtDescuento.FormatString = "#,##0.00";
			this.txtDescuento.Location = new System.Drawing.Point(320, 160);
			this.txtDescuento.MaxValue = 100;
			this.txtDescuento.MinValue = 0;
			this.txtDescuento.Name = "txtDescuento";
			this.txtDescuento.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtDescuento.PromptChar = ' ';
			this.txtDescuento.Size = new System.Drawing.Size(100, 22);
			this.txtDescuento.TabIndex = 709;
			this.txtDescuento.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtDescuento_KeyDown);
			this.txtDescuento.Click += new System.EventHandler(this.txtDescuento_Click);
			this.txtDescuento.Validated += new System.EventHandler(this.txtDescuento_Validated);
			this.txtDescuento.Enter += new System.EventHandler(this.txtDescuento_Enter);
			// 
			// label1
			// 
			this.label1.AutoSize = true;
			this.label1.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label1.Location = new System.Drawing.Point(448, 163);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(29, 17);
			this.label1.TabIndex = 710;
			this.label1.Text = "Total";
			this.label1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtTotal
			// 
			appearance6.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtTotal.Appearance = appearance6;
			this.txtTotal.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtTotal.Enabled = false;
			this.txtTotal.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtTotal.FormatString = "#,##0.00";
			this.txtTotal.Location = new System.Drawing.Point(496, 160);
			this.txtTotal.Name = "txtTotal";
			this.txtTotal.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtTotal.PromptChar = ' ';
			this.txtTotal.Size = new System.Drawing.Size(100, 22);
			this.txtTotal.TabIndex = 711;
			// 
			// optContrato
			// 
			appearance7.ForeColorDisabled = System.Drawing.Color.Black;
			this.optContrato.Appearance = appearance7;
			this.optContrato.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.optContrato.CheckedIndex = 0;
			this.optContrato.Enabled = false;
			this.optContrato.ItemAppearance = appearance8;
			this.optContrato.ItemOrigin = new System.Drawing.Point(10, 0);
			valueListItem1.DataValue = 0;
			valueListItem1.DisplayText = "Sin Contrato";
			valueListItem2.DataValue = 1;
			valueListItem2.DisplayText = "Contrato";
			this.optContrato.Items.Add(valueListItem1);
			this.optContrato.Items.Add(valueListItem2);
			this.optContrato.ItemSpacingVertical = 10;
			this.optContrato.Location = new System.Drawing.Point(16, 240);
			this.optContrato.Name = "optContrato";
			this.optContrato.Size = new System.Drawing.Size(192, 24);
			this.optContrato.TabIndex = 712;
			this.optContrato.Text = "Sin Contrato";
			this.optContrato.ValueChanged += new System.EventHandler(this.optEstados_ValueChanged);
			// 
			// lblFechaInicia
			// 
			this.lblFechaInicia.AutoSize = true;
			this.lblFechaInicia.BackColor = System.Drawing.Color.Transparent;
			this.lblFechaInicia.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblFechaInicia.Location = new System.Drawing.Point(16, 202);
			this.lblFechaInicia.Name = "lblFechaInicia";
			this.lblFechaInicia.Size = new System.Drawing.Size(78, 17);
			this.lblFechaInicia.TabIndex = 714;
			this.lblFechaInicia.Text = "Inicia Proyecto";
			this.lblFechaInicia.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// dtFechaInicia
			// 
			appearance9.FontData.Name = "Tahoma";
			appearance9.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtFechaInicia.Appearance = appearance9;
			this.dtFechaInicia.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton2.Action = Infragistics.Win.UltraWinSchedule.DateButtonAction.AddToSelection;
			dateButton2.Caption = "Today";
			dateButton2.Type = Infragistics.Win.UltraWinSchedule.DateButtonType.Custom;
			this.dtFechaInicia.DateButtons.Add(dateButton2);
			this.dtFechaInicia.Enabled = false;
			this.dtFechaInicia.Format = "dd/MM/yyyy";
			this.dtFechaInicia.Location = new System.Drawing.Point(104, 200);
			this.dtFechaInicia.Name = "dtFechaInicia";
			this.dtFechaInicia.NonAutoSizeHeight = 23;
			this.dtFechaInicia.Size = new System.Drawing.Size(104, 21);
			this.dtFechaInicia.SpinButtonsVisible = true;
			this.dtFechaInicia.TabIndex = 713;
			this.dtFechaInicia.Value = ((object)(resources.GetObject("dtFechaInicia.Value")));
			this.dtFechaInicia.KeyDown += new System.Windows.Forms.KeyEventHandler(this.dtFechaInicia_KeyDown);
			// 
			// txtEstado
			// 
			appearance10.ForeColor = System.Drawing.Color.Black;
			appearance10.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtEstado.Appearance = appearance10;
			this.txtEstado.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtEstado.Enabled = false;
			this.txtEstado.FormatString = "###0";
			this.txtEstado.Location = new System.Drawing.Point(272, 7);
			this.txtEstado.MaxValue = 100000000;
			this.txtEstado.MinValue = 0;
			this.txtEstado.Name = "txtEstado";
			this.txtEstado.PromptChar = ' ';
			this.txtEstado.Size = new System.Drawing.Size(40, 22);
			this.txtEstado.TabIndex = 715;
			this.txtEstado.Visible = false;
			// 
			// btnNuevo
			// 
			this.btnNuevo.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnNuevo.CausesValidation = false;
			this.btnNuevo.Image = ((System.Drawing.Image)(resources.GetObject("btnNuevo.Image")));
			this.btnNuevo.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnNuevo.Location = new System.Drawing.Point(8, 520);
			this.btnNuevo.Name = "btnNuevo";
			this.btnNuevo.Size = new System.Drawing.Size(80, 23);
			this.btnNuevo.TabIndex = 761;
			this.btnNuevo.Text = "&Nuevo";
			this.btnNuevo.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnNuevo.Click += new System.EventHandler(this.btnNuevo_Click);
			// 
			// btnConsultar
			// 
			this.btnConsultar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnConsultar.CausesValidation = false;
			this.btnConsultar.Image = ((System.Drawing.Image)(resources.GetObject("btnConsultar.Image")));
			this.btnConsultar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnConsultar.Location = new System.Drawing.Point(94, 520);
			this.btnConsultar.Name = "btnConsultar";
			this.btnConsultar.Size = new System.Drawing.Size(84, 23);
			this.btnConsultar.TabIndex = 762;
			this.btnConsultar.Text = "&Consultar";
			this.btnConsultar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnConsultar.Click += new System.EventHandler(this.btnConsultar_Click);
			// 
			// btnEditar
			// 
			this.btnEditar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnEditar.CausesValidation = false;
			this.btnEditar.Enabled = false;
			this.btnEditar.Image = ((System.Drawing.Image)(resources.GetObject("btnEditar.Image")));
			this.btnEditar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnEditar.Location = new System.Drawing.Point(184, 520);
			this.btnEditar.Name = "btnEditar";
			this.btnEditar.Size = new System.Drawing.Size(82, 23);
			this.btnEditar.TabIndex = 765;
			this.btnEditar.Text = "&Editar";
			this.btnEditar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnEditar.Click += new System.EventHandler(this.btnEditar_Click);
			// 
			// btnGuardar
			// 
			this.btnGuardar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnGuardar.Enabled = false;
			this.btnGuardar.Image = ((System.Drawing.Image)(resources.GetObject("btnGuardar.Image")));
			this.btnGuardar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnGuardar.Location = new System.Drawing.Point(272, 520);
			this.btnGuardar.Name = "btnGuardar";
			this.btnGuardar.Size = new System.Drawing.Size(82, 23);
			this.btnGuardar.TabIndex = 763;
			this.btnGuardar.Text = "&Guardar";
			this.btnGuardar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnGuardar.Click += new System.EventHandler(this.btnGuardar_Click);
			// 
			// btnCancelar
			// 
			this.btnCancelar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnCancelar.CausesValidation = false;
			this.btnCancelar.Image = ((System.Drawing.Image)(resources.GetObject("btnCancelar.Image")));
			this.btnCancelar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnCancelar.Location = new System.Drawing.Point(524, 520);
			this.btnCancelar.Name = "btnCancelar";
			this.btnCancelar.Size = new System.Drawing.Size(78, 23);
			this.btnCancelar.TabIndex = 767;
			this.btnCancelar.Text = "&Cancelar";
			this.btnCancelar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnCancelar.Click += new System.EventHandler(this.btnCancelar_Click);
			// 
			// btnAnular
			// 
			this.btnAnular.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnAnular.CausesValidation = false;
			this.btnAnular.Enabled = false;
			this.btnAnular.Image = ((System.Drawing.Image)(resources.GetObject("btnAnular.Image")));
			this.btnAnular.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnAnular.Location = new System.Drawing.Point(440, 520);
			this.btnAnular.Name = "btnAnular";
			this.btnAnular.Size = new System.Drawing.Size(78, 23);
			this.btnAnular.TabIndex = 766;
			this.btnAnular.Text = "&Anular";
			this.btnAnular.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnAnular.Click += new System.EventHandler(this.btnAnular_Click);
			// 
			// btnSalir
			// 
			this.btnSalir.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnSalir.CausesValidation = false;
			this.btnSalir.Image = ((System.Drawing.Image)(resources.GetObject("btnSalir.Image")));
			this.btnSalir.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnSalir.Location = new System.Drawing.Point(608, 520);
			this.btnSalir.Name = "btnSalir";
			this.btnSalir.Size = new System.Drawing.Size(82, 23);
			this.btnSalir.TabIndex = 768;
			this.btnSalir.Text = "&Salir";
			this.btnSalir.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnSalir.Click += new System.EventHandler(this.btnSalir_Click);
			// 
			// groupBox3
			// 
			this.groupBox3.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.groupBox3.Location = new System.Drawing.Point(0, 224);
			this.groupBox3.Name = "groupBox3";
			this.groupBox3.Size = new System.Drawing.Size(712, 8);
			this.groupBox3.TabIndex = 769;
			this.groupBox3.TabStop = false;
			// 
			// lblEstado
			// 
			this.lblEstado.AutoSize = true;
			this.lblEstado.Font = new System.Drawing.Font("Tahoma", 11F, System.Drawing.FontStyle.Bold);
			this.lblEstado.ForeColor = System.Drawing.Color.Firebrick;
			this.lblEstado.ImeMode = System.Windows.Forms.ImeMode.NoControl;
			this.lblEstado.Location = new System.Drawing.Point(528, 200);
			this.lblEstado.Name = "lblEstado";
			this.lblEstado.Size = new System.Drawing.Size(0, 21);
			this.lblEstado.TabIndex = 770;
			this.lblEstado.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
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
			// btnAutorizar
			// 
			this.btnAutorizar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnAutorizar.Enabled = false;
			this.btnAutorizar.Image = ((System.Drawing.Image)(resources.GetObject("btnAutorizar.Image")));
			this.btnAutorizar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnAutorizar.Location = new System.Drawing.Point(360, 520);
			this.btnAutorizar.Name = "btnAutorizar";
			this.btnAutorizar.Size = new System.Drawing.Size(74, 23);
			this.btnAutorizar.TabIndex = 772;
			this.btnAutorizar.Text = "Autorizar";
			this.btnAutorizar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnAutorizar.Click += new System.EventHandler(this.btnAutorizar_Click);
			// 
			// lblIdentificacion
			// 
			this.lblIdentificacion.AutoSize = true;
			this.lblIdentificacion.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblIdentificacion.Location = new System.Drawing.Point(16, 41);
			this.lblIdentificacion.Name = "lblIdentificacion";
			this.lblIdentificacion.Size = new System.Drawing.Size(48, 17);
			this.lblIdentificacion.TabIndex = 774;
			this.lblIdentificacion.Text = "RUC/C.I.";
			this.lblIdentificacion.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtRuc
			// 
			appearance11.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtRuc.Appearance = appearance11;
			this.txtRuc.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtRuc.Enabled = false;
			this.txtRuc.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtRuc.Location = new System.Drawing.Point(104, 38);
			this.txtRuc.MaxLength = 13;
			this.txtRuc.Name = "txtRuc";
			this.txtRuc.Size = new System.Drawing.Size(136, 22);
			this.txtRuc.TabIndex = 773;
			this.txtRuc.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtRuc_KeyDown);
			this.txtRuc.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtRuc_KeyPress);
			this.txtRuc.Validating += new System.ComponentModel.CancelEventHandler(this.txtRuc_Validating);
			this.txtRuc.Validated += new System.EventHandler(this.txtRuc_Validated);
			// 
			// btnArchivoContrato
			// 
			this.btnArchivoContrato.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnArchivoContrato.Enabled = false;
			this.btnArchivoContrato.Image = ((System.Drawing.Image)(resources.GetObject("btnArchivoContrato.Image")));
			this.btnArchivoContrato.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnArchivoContrato.Location = new System.Drawing.Point(584, 272);
			this.btnArchivoContrato.Name = "btnArchivoContrato";
			this.btnArchivoContrato.Size = new System.Drawing.Size(120, 23);
			this.btnArchivoContrato.TabIndex = 778;
			this.btnArchivoContrato.Text = "Archivo Contrato";
			this.btnArchivoContrato.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnArchivoContrato.Click += new System.EventHandler(this.btnArchivo_Click);
			// 
			// ultraDataSource1
			// 
			ultraDataColumn1.DataType = typeof(int);
			ultraDataColumn2.DataType = typeof(int);
			ultraDataColumn3.DataType = typeof(System.DateTime);
			this.ultraDataSource1.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn1,
																																 ultraDataColumn2,
																																 ultraDataColumn3,
																																 ultraDataColumn4});
			// 
			// txtIdCliente
			// 
			this.txtIdCliente.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtIdCliente.Enabled = false;
			this.txtIdCliente.Location = new System.Drawing.Point(320, 7);
			this.txtIdCliente.Name = "txtIdCliente";
			this.txtIdCliente.PromptChar = ' ';
			this.txtIdCliente.Size = new System.Drawing.Size(40, 22);
			this.txtIdCliente.TabIndex = 780;
			this.txtIdCliente.Visible = false;
			// 
			// lblCliente
			// 
			this.lblCliente.AutoSize = true;
			this.lblCliente.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblCliente.Location = new System.Drawing.Point(256, 41);
			this.lblCliente.Name = "lblCliente";
			this.lblCliente.Size = new System.Drawing.Size(55, 17);
			this.lblCliente.TabIndex = 782;
			this.lblCliente.Text = "Proveedor";
			this.lblCliente.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtNombre
			// 
			appearance12.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtNombre.Appearance = appearance12;
			this.txtNombre.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtNombre.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNombre.Enabled = false;
			this.txtNombre.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtNombre.Location = new System.Drawing.Point(328, 38);
			this.txtNombre.Name = "txtNombre";
			this.txtNombre.Size = new System.Drawing.Size(376, 22);
			this.txtNombre.TabIndex = 781;
			// 
			// ultraGrid1
			// 
			this.ultraGrid1.Cursor = System.Windows.Forms.Cursors.Default;
			this.ultraGrid1.DataSource = this.ultraDataSource1;
			appearance13.BackColor = System.Drawing.Color.White;
			this.ultraGrid1.DisplayLayout.Appearance = appearance13;
			this.ultraGrid1.DisplayLayout.AutoFitColumns = true;
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridColumn2.Hidden = true;
			ultraGridColumn3.Header.VisiblePosition = 2;
			ultraGridColumn3.Width = 91;
			ultraGridColumn4.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn4.Header.VisiblePosition = 3;
			ultraGridColumn4.Width = 592;
			ultraGridBand1.Columns.AddRange(new object[] {
																										 ultraGridColumn1,
																										 ultraGridColumn2,
																										 ultraGridColumn3,
																										 ultraGridColumn4});
			ultraGridBand1.Header.Caption = "Fechas de Entrega";
			ultraGridBand1.HeaderVisible = true;
			this.ultraGrid1.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.ultraGrid1.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.TemplateOnBottom;
			this.ultraGrid1.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.ultraGrid1.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance14.BackColor = System.Drawing.Color.Transparent;
			this.ultraGrid1.DisplayLayout.Override.CardAreaAppearance = appearance14;
			appearance15.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance15.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance15.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance15.FontData.BoldAsString = "True";
			appearance15.FontData.Name = "Arial";
			appearance15.FontData.SizeInPoints = 8.5F;
			appearance15.ForeColor = System.Drawing.Color.White;
			appearance15.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.ultraGrid1.DisplayLayout.Override.HeaderAppearance = appearance15;
			appearance16.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance16.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance16.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.RowSelectorAppearance = appearance16;
			appearance17.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance17.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance17.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.SelectedRowAppearance = appearance17;
			this.ultraGrid1.Enabled = false;
			this.ultraGrid1.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.ultraGrid1.Location = new System.Drawing.Point(4, 360);
			this.ultraGrid1.Name = "ultraGrid1";
			this.ultraGrid1.RightToLeft = System.Windows.Forms.RightToLeft.No;
			this.ultraGrid1.Size = new System.Drawing.Size(704, 144);
			this.ultraGrid1.TabIndex = 783;
			this.ultraGrid1.AfterRowInsert += new Infragistics.Win.UltraWinGrid.RowEventHandler(this.ultraGrid1_AfterRowInsert);
			this.ultraGrid1.BeforeRowsDeleted += new Infragistics.Win.UltraWinGrid.BeforeRowsDeletedEventHandler(this.ultraGrid1_BeforeRowsDeleted);
			this.ultraGrid1.KeyDown += new System.Windows.Forms.KeyEventHandler(this.ultraGrid1_KeyDown);
			this.ultraGrid1.AfterCellUpdate += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.ultraGrid1_AfterCellUpdate);
			this.ultraGrid1.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.ultraGrid1_InitializeLayout);
			// 
			// txtidContrato
			// 
			this.txtidContrato.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtidContrato.Enabled = false;
			this.txtidContrato.Location = new System.Drawing.Point(224, 7);
			this.txtidContrato.Name = "txtidContrato";
			this.txtidContrato.PromptChar = ' ';
			this.txtidContrato.Size = new System.Drawing.Size(40, 22);
			this.txtidContrato.TabIndex = 784;
			this.txtidContrato.Visible = false;
			// 
			// groupBox1
			// 
			this.groupBox1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.groupBox1.Location = new System.Drawing.Point(0, 144);
			this.groupBox1.Name = "groupBox1";
			this.groupBox1.Size = new System.Drawing.Size(712, 8);
			this.groupBox1.TabIndex = 785;
			this.groupBox1.TabStop = false;
			// 
			// groupBox2
			// 
			this.groupBox2.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.groupBox2.Location = new System.Drawing.Point(0, 184);
			this.groupBox2.Name = "groupBox2";
			this.groupBox2.Size = new System.Drawing.Size(712, 8);
			this.groupBox2.TabIndex = 786;
			this.groupBox2.TabStop = false;
			// 
			// groupBox4
			// 
			this.groupBox4.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.groupBox4.Location = new System.Drawing.Point(0, 504);
			this.groupBox4.Name = "groupBox4";
			this.groupBox4.Size = new System.Drawing.Size(712, 8);
			this.groupBox4.TabIndex = 787;
			this.groupBox4.TabStop = false;
			// 
			// lbltitulo
			// 
			this.lbltitulo.AutoSize = true;
			this.lbltitulo.Location = new System.Drawing.Point(16, 275);
			this.lbltitulo.Name = "lbltitulo";
			this.lbltitulo.Size = new System.Drawing.Size(32, 17);
			this.lbltitulo.TabIndex = 800;
			this.lbltitulo.Text = "Titulo";
			this.lbltitulo.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label2
			// 
			this.label2.AutoSize = true;
			this.label2.BackColor = System.Drawing.Color.Transparent;
			this.label2.Location = new System.Drawing.Point(16, 318);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(33, 17);
			this.label2.TabIndex = 798;
			this.label2.Text = "Notas";
			this.label2.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label3
			// 
			this.label3.AutoSize = true;
			this.label3.BackColor = System.Drawing.Color.Transparent;
			this.label3.Location = new System.Drawing.Point(256, 244);
			this.label3.Name = "label3";
			this.label3.Size = new System.Drawing.Size(34, 17);
			this.label3.TabIndex = 795;
			this.label3.Text = "Fecha";
			this.label3.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtTitulo
			// 
			appearance18.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtTitulo.Appearance = appearance18;
			this.txtTitulo.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtTitulo.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtTitulo.Enabled = false;
			this.txtTitulo.Location = new System.Drawing.Point(72, 272);
			this.txtTitulo.Name = "txtTitulo";
			this.txtTitulo.Size = new System.Drawing.Size(352, 22);
			this.txtTitulo.TabIndex = 799;
			// 
			// txtNotasContrato
			// 
			this.txtNotasContrato.AcceptsReturn = true;
			appearance19.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtNotasContrato.Appearance = appearance19;
			this.txtNotasContrato.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtNotasContrato.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNotasContrato.Enabled = false;
			this.txtNotasContrato.Location = new System.Drawing.Point(72, 304);
			this.txtNotasContrato.MaxLength = 1000;
			this.txtNotasContrato.Multiline = true;
			this.txtNotasContrato.Name = "txtNotasContrato";
			this.txtNotasContrato.Size = new System.Drawing.Size(352, 44);
			this.txtNotasContrato.TabIndex = 797;
			// 
			// dtFechaContrato
			// 
			appearance20.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtFechaContrato.Appearance = appearance20;
			this.dtFechaContrato.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton3.Action = Infragistics.Win.UltraWinSchedule.DateButtonAction.AddToSelection;
			dateButton3.Caption = "Today";
			dateButton3.Type = Infragistics.Win.UltraWinSchedule.DateButtonType.Custom;
			this.dtFechaContrato.DateButtons.Add(dateButton3);
			this.dtFechaContrato.Enabled = false;
			this.dtFechaContrato.Format = "dd/MM/yyyy";
			this.dtFechaContrato.Location = new System.Drawing.Point(320, 242);
			this.dtFechaContrato.Name = "dtFechaContrato";
			this.dtFechaContrato.NonAutoSizeHeight = 23;
			this.dtFechaContrato.Size = new System.Drawing.Size(104, 21);
			this.dtFechaContrato.SpinButtonsVisible = true;
			this.dtFechaContrato.TabIndex = 796;
			this.dtFechaContrato.Value = ((object)(resources.GetObject("dtFechaContrato.Value")));
			// 
			// btnArchivoProforma
			// 
			this.btnArchivoProforma.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnArchivoProforma.Enabled = false;
			this.btnArchivoProforma.Image = ((System.Drawing.Image)(resources.GetObject("btnArchivoProforma.Image")));
			this.btnArchivoProforma.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnArchivoProforma.Location = new System.Drawing.Point(584, 241);
			this.btnArchivoProforma.Name = "btnArchivoProforma";
			this.btnArchivoProforma.Size = new System.Drawing.Size(120, 23);
			this.btnArchivoProforma.TabIndex = 801;
			this.btnArchivoProforma.Text = "Archivo Proforma";
			this.btnArchivoProforma.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnArchivoProforma.Click += new System.EventHandler(this.btnArchivoProforma_Click);
			// 
			// frmProformas
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 14);
			this.ClientSize = new System.Drawing.Size(712, 550);
			this.Controls.Add(this.btnArchivoProforma);
			this.Controls.Add(this.lbltitulo);
			this.Controls.Add(this.label2);
			this.Controls.Add(this.label3);
			this.Controls.Add(this.lblCliente);
			this.Controls.Add(this.lblIdentificacion);
			this.Controls.Add(this.lblFechaInicia);
			this.Controls.Add(this.label1);
			this.Controls.Add(this.lblDescuento);
			this.Controls.Add(this.lblCosto);
			this.Controls.Add(this.lblDetalle);
			this.Controls.Add(this.lblFecha);
			this.Controls.Add(this.lblNumero);
			this.Controls.Add(this.txtTitulo);
			this.Controls.Add(this.txtNotasContrato);
			this.Controls.Add(this.dtFechaContrato);
			this.Controls.Add(this.groupBox4);
			this.Controls.Add(this.groupBox2);
			this.Controls.Add(this.groupBox1);
			this.Controls.Add(this.txtidContrato);
			this.Controls.Add(this.ultraGrid1);
			this.Controls.Add(this.txtNombre);
			this.Controls.Add(this.txtIdCliente);
			this.Controls.Add(this.btnArchivoContrato);
			this.Controls.Add(this.txtRuc);
			this.Controls.Add(this.btnAutorizar);
			this.Controls.Add(this.lblEstado);
			this.Controls.Add(this.groupBox3);
			this.Controls.Add(this.btnNuevo);
			this.Controls.Add(this.btnConsultar);
			this.Controls.Add(this.btnEditar);
			this.Controls.Add(this.btnGuardar);
			this.Controls.Add(this.btnCancelar);
			this.Controls.Add(this.btnAnular);
			this.Controls.Add(this.btnSalir);
			this.Controls.Add(this.txtEstado);
			this.Controls.Add(this.dtFechaInicia);
			this.Controls.Add(this.optContrato);
			this.Controls.Add(this.txtTotal);
			this.Controls.Add(this.txtDescuento);
			this.Controls.Add(this.txtCosto);
			this.Controls.Add(this.txtDetalle);
			this.Controls.Add(this.dtFecha);
			this.Controls.Add(this.txtNumero);
			this.Controls.Add(this.txtidProformas);
			this.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.KeyPreview = true;
			this.Name = "frmProformas";
			this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
			this.Text = "Proformas";
			this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.frmProformas_KeyDown);
			this.Load += new System.EventHandler(this.frmProformas_Load);
			((System.ComponentModel.ISupportInitialize)(this.txtidProformas)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFecha)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDetalle)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCosto)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDescuento)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtTotal)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.optContrato)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFechaInicia)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtEstado)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtRuc)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdCliente)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNombre)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtidContrato)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtTitulo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNotasContrato)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFechaContrato)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void btnCancelar_Click(object sender, System.EventArgs e)
		{
			bNuevo = false;
			bEdicion = false;

			#region Bloqueo Controles 
			this.txtRuc.Enabled = false;
			this.txtNumero.Enabled= false;
			this.dtFecha.Enabled = false;
			this.txtDetalle.Enabled = false;
			this.txtCosto.Enabled = false;
			this.txtDescuento.Enabled = false;
			this.dtFechaInicia.Enabled = false;
			this.txtEstado.Enabled = false;
			
			this.btnArchivoProforma.Enabled = false;			

			this.optContrato.Enabled = false;
			this.dtFechaContrato.Enabled = false;
			this.txtTitulo.Enabled = false;
			this.txtNotasContrato.Enabled = false;
			this.btnArchivoContrato.Enabled = false;

			FuncionesProcedimientos.EstadoGrids(false, this.ultraGrid1);
			this.ultraGrid1.DataSource = Funcion.dvProcedimiento(cdsSeteoF, string.Format("Exec ConsultaIndividualFechaEntregaProforma 0"));

			#endregion Bloqueo Controles 

			#region Vaciar Campos
			this.txtidProformas.Value = 0;
			this.txtNumero.Text ="";
			this.dtFecha.Value = System.DBNull.Value;
			this.txtIdCliente.Value = 0;
			this.txtRuc.Text = "";			
			this.txtNombre.Text = "";
			this.txtDetalle.Text= "";
			this.txtCosto.Value = 0.00m;
			this.txtDescuento.Value = 0.00m;
			this.txtTotal.Value = 0.00m;			
			this.dtFechaInicia.Value = System.DBNull.Value;
			this.txtEstado.Value = 0;
			this.lblEstado.Text = "";	
			this.optContrato.CheckedIndex = -1;
			this.dtFechaContrato.Text = "";
			this.txtTitulo.Text = "";
			this.txtNotasContrato.Text = "";
			sArchivoContrato = "";
			sArchivoProforma = "";			
			#endregion Vaciar Campos
    
			#region Bloquea Botones 
			if (miAcceso.Nuevo) this.btnNuevo.Enabled = true;
			if (miAcceso.Buscar) this.btnConsultar.Enabled = true;
			this.btnEditar.Enabled = false;
			this.btnGuardar.Enabled = false;
			this.btnCancelar.Enabled = false;
			this.btnAnular.Enabled = false;
			this.btnAutorizar.Enabled = false;
			#endregion Bloquea Botones
		}

		private void frmProformas_Load(object sender, System.EventArgs e)
		{
			miAcceso = new Acceso(cdsSeteoF, "0620");

			if (!Funcion.Permiso("899", cdsSeteoF))
			{				
				MessageBox.Show("No puede ingresar a Proformas", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				BeginInvoke(new MethodInvoker(UnloadMe));
				return;
			}
			if (!miAcceso.Nuevo) this.btnNuevo.Enabled = false;
			if (!miAcceso.Buscar) this.btnConsultar.Enabled = false;
		}

		private void btnNuevo_Click(object sender, System.EventArgs e)
		{
			this.btnCancelar_Click(sender, e);
			this.dtFecha.Value = DateTime.Today;
			this.dtFechaInicia.Value = DateTime.Today;
			this.dtFechaInicia.CalendarInfo.MinDate = Convert.ToDateTime(this.dtFecha.Value);
			this.dtFechaContrato.CalendarInfo.MinDate = Convert.ToDateTime(this.dtFecha.Value);
			this.txtEstado.Value = 1;
			this.optContrato.Value = 0;

			#region Habilita Controles
			this.txtRuc.Enabled = true;
			this.txtDetalle.Enabled = true;
			this.txtCosto.Enabled = true;
			this.txtDescuento.Enabled = true;			
			this.optContrato.Enabled = true;
			this.dtFechaInicia.Enabled = true;
			this.btnArchivoProforma.Enabled = true;

			FuncionesProcedimientos.EstadoGrids(true, this.ultraGrid1);
						 
			this.ultraGrid1.DataSource = Funcion.dvProcedimiento(cdsSeteoF, string.Format("Exec ConsultaIndividualFechaEntregaProforma 0"));
			#endregion Habilita Controles 

			#region Bloquea Botones 
			this.btnNuevo.Enabled = false;
			this.btnConsultar.Enabled = false;
			this.btnEditar.Enabled = false;
			this.btnGuardar.Enabled = true;
			this.btnAnular.Enabled = false;
			this.btnCancelar.Enabled = true;
			this.btnAutorizar.Enabled = false;			
			#endregion Bloquea Botones 

			bNuevo = true;
			bEdicion = true;

			this.txtRuc.Focus();
		}

		private void btnConsultar_Click(object sender, System.EventArgs e)
		{
			this.btnCancelar_Click(sender, e);

			using (frmConsultaProformas consulprofos = new frmConsultaProformas())
				if (DialogResult.OK == consulprofos.ShowDialog())
					this.Consultar((int)consulprofos.ultraGrid1.ActiveRow.Cells["idProforma"].Value);					
		}		

		private void Consultar (int idProforma)
		{
			this.btnNuevo.Enabled = false;
			this.btnConsultar.Enabled = false;
			this.btnEditar.Enabled = false;
			this.btnGuardar.Enabled = false;
			this.btnCancelar.Enabled = false;
			this.btnAnular.Enabled = false;
			this.btnAutorizar.Enabled = false;

			#region Consulta Individual
			string sSQL = string.Format ("Exec ConsultaIndividualProformas {0}", idProforma); 
			SqlDataReader drr = Funcion.rEscalarSQL(cdsSeteoF, sSQL,true);
			drr.Read();

			this.txtidProformas.Value = drr.GetInt32 (0);
			this.txtNumero.Text = drr.GetString (1);
			this.dtFecha.Value = drr.GetDateTime (2);
			this.txtIdCliente.Value= drr.GetInt32(3);
			this.txtRuc.Text= drr.GetString(4);
			this.txtNombre.Text= drr.GetString(5);
			this.txtDetalle.Text= drr.GetString(6);
			this.txtCosto.Value =drr.GetDecimal(7);
			this.txtDescuento.Value = drr.GetDecimal(8);
			this.txtTotal.Value = drr.GetDecimal(9);			
			this.dtFechaInicia.Value = drr.GetDateTime(10);
			this.txtEstado.Value = drr.GetInt32 (11);

			this.optContrato.Value = drr.GetInt32(12);
			if (drr.GetValue(13) != System.DBNull.Value) this.dtFechaContrato.Value = drr.GetDateTime(13);
			if (drr.GetValue(14) != System.DBNull.Value) this.txtTitulo.Text = drr.GetString(14);
			if (drr.GetValue(15) != System.DBNull.Value) this.txtNotasContrato.Text = drr.GetString(15);

			sArchivoProforma = drr.GetString(16);
			sArchivoContrato = drr.GetString(17);			
			drr.Close();

			this.ultraGrid1.DataSource = Funcion.dvProcedimiento(cdsSeteoF, string.Format("Exec ConsultaIndividualFechaEntregaProforma {0}", (int)this.txtidProformas.Value));
			#endregion Consulta Individual

			#region Estados
			if( (int) this.txtEstado.Value <= 1)
			{
				this.lblEstado.Text = "PENDIENTE";
				if (miAcceso.Imprimir) this.btnAutorizar.Enabled = true;				
				if (miAcceso.Editar) this.btnEditar.Enabled = true;
			}
			else if((int)this.txtEstado.Value == 2) this.lblEstado.Text = "AUTORIZADO";	
				else if( (int) this.txtEstado.Value == 3) this.lblEstado.Text = "FACTURADO";
			else if( (int) this.txtEstado.Value == 4) this.lblEstado.Text = "ANULADO";
			#endregion Estados
		
			if((int)this.optContrato.Value == 0)
				this.btnArchivoContrato.Enabled = false;
			else 
				this.btnArchivoContrato.Enabled = true;

			if (miAcceso.Nuevo)this.btnNuevo.Enabled = true;
			if (miAcceso.Buscar)this.btnConsultar.Enabled = true;
			if (miAcceso.Anular) this.btnAnular.Enabled = true;
		
			this.btnCancelar.Enabled = true;			
		}

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);	
		}

		private void btnSalir_Click(object sender, System.EventArgs e)
		{
			this.Close();	
		}

		private void btnGuardar_Click(object sender, System.EventArgs e)
		{
			int iCont = 0;
      		
			#region Validacion General Proforma
			if (!Validacion.vbFechaEnDocumentos(this.dtFecha, "Ingrese la fecha del documento", true, "Cobros", cdsSeteoF)) return;
			
			if (!Validacion.vbClienteExiste(this.txtRuc, true, 4, 6, cdsSeteoF)) return;	
			
			if (!Validacion.vbTexto(this.txtNombre, 10, 1000, "El Nombre")) return;			

			if (!Validacion.vbTexto(this.txtDetalle, 10, 1000, "El Detalle")) return;

			if (!Validacion.vbFechaEnDocumentos(this.dtFecha, "Seleccione la Fecha Inicial", false, "", cdsSeteoF)) return;

			if (!Validacion.vbCampoDecimalVacio(this.txtTotal, "El Total no Puede Tener valor 0 ", 0.01m, 150000.00m)) return; 

			if ((int) this.optContrato.Value < 0)
			{
				MessageBox.Show("Seleccione el tipo de contrato ", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.optContrato.Focus();
				return; 
			}

			if (!Validacion.vbFechaEnDocumentos(this.dtFechaInicia, "Seleccione la Fecha Inicial", false, "", cdsSeteoF)) return;
			//				if (!Validacion.vbTexto(this.txtArchivo, 10, 1000, "Archivo")) return;

			#endregion Validacion General Proforma

			#region Validacion Entrega Proforma
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.ultraGrid1.Rows.All)
				iCont = iCont + 1;

			if(iCont <= 0)
			{
				MessageBox.Show("Ingrese al menos una fecha de entrega", "Validacion", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.ultraGrid1.Focus();
				return;
			}
			
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.ultraGrid1.Rows.All)
			{
				if ((DateTime)dr.Cells["Fecha"].Value < (DateTime)this.dtFechaInicia.Value)
				{
					MessageBox.Show("Las fechas de entrega no pueden ser anteriores a la fecha de inicio", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					this.ultraGrid1.ActiveRow = dr;
					this.ultraGrid1.ActiveRow.Selected = true;
					return;
				}
			}

			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.ultraGrid1.Rows.All)
			{
				if (dr.Cells["Fecha"].Value.ToString().Trim().Length == 0 )
				{
					MessageBox.Show("Seleccione una fecha", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					this.ultraGrid1.ActiveRow = dr;
					this.ultraGrid1.ActiveRow.Selected = true;
					return;
				}
				
				if (dr.Cells["Descripcion"].Value.ToString().Trim().Length == 0)
				{
					MessageBox.Show("Debe Ingresar la descripción de entrega", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);				
					this.ultraGrid1.ActiveRow = dr;
					this.ultraGrid1.ActiveRow.Selected = true;
					return;
				}				

				if (dr.Cells["Descripcion"].Value.ToString().Trim().Length < 10)
				{
					MessageBox.Show("La Descripcion de de entrega debe ser de al menos 10 caracteres", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);				
					this.ultraGrid1.ActiveRow = dr;
					this.ultraGrid1.ActiveRow.Selected = true;
					return;
				}				
			}			
			#endregion Validacion Entrega Proforma

			#region Valida Contrato

			if( (int) this.txtIdCliente.Value < 0)
			{
				MessageBox.Show(" Ingrese el Ruc", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
				return;
			}


			
			//			if( sCArchivo.Length == 0)
			//			{
			//				MessageBox.Show(" Precione el boton Archivo para ingresar la direccion del Archivo  ", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
			//				return;
			//			}
			//
			//			if( sCTitulo.Length == 0)
			//			{
			//				MessageBox.Show(" Precione el boton Archivo para ingresar el Titulo ", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
			//				return;
			//			}
			#endregion Valida Contrato

			#region Numeracion
			if (bNuevo)
			{
				string ssSQL = string.Format("Exec NumeracionLocales 64, 0, 0");
				this.txtNumero.Text = Funcion.sEscalarSQL(cdsSeteoF, ssSQL);			
			}
			#endregion Numeracion

			if (sArchivoProforma.Length > 0)
			{								
				string sArchivoProformaN = @"\\192.168.1.3\doclatinium\" + MenuLatinium.stDirFacturacion + @"\Proforma\" + this.txtNumero.Text.ToString() + ".pdf";

				if (!File.Exists(sArchivoProformaN))
				{
					File.Copy(sArchivoProforma, sArchivoProformaN);

					sArchivoProforma = sArchivoProformaN;
				}
			}

			if (sArchivoContrato.Length > 0)
			{				
				string sArchivoContratoN = @"\\192.168.1.3\doclatinium\" + MenuLatinium.stDirFacturacion + @"\Contratos Proveedores\" + this.txtNumero.Text.ToString() + ".pdf";

				if (!File.Exists(sArchivoContratoN))
				{
					File.Copy(sArchivoContrato, sArchivoContratoN);

					sArchivoContrato = sArchivoContratoN;
				}
			}
			
			DateTime dtFechaContrato = DateTime.Parse("01/01/2000");
			if (this.dtFechaContrato.Value != System.DBNull.Value) dtFechaContrato = Convert.ToDateTime(this.dtFechaContrato.Value);

			#region Guarda Proforma
			string sSQL = string.Format("Exec GuardaProformas {0}, '{1}', '{2}', '{3}', {4}, {5}, {6}, {7}, '{8}', {9}, '{10}', {11}, '{12}', '{13}', '{14}', '{15}', '{16}'",
				(int)this.txtidProformas.Value, 
				this.txtNumero.Text.ToString(),
				Convert.ToDateTime(this.dtFecha.Value).ToString("yyyyMMdd"),
				this.txtDetalle.Text.ToString(),
				Convert.ToDecimal(this.txtCosto.Value),
				Convert.ToDecimal(this.txtDescuento.Value),
				Convert.ToDecimal(this.txtTotal.Value),
				(int) this.optContrato.Value,
				Convert.ToDateTime(this.dtFechaInicia.Value).ToString("yyyyMMdd"),
				(int)this.txtEstado.Value,
				Convert.ToDateTime(this.dtFechaInicia.Value).ToString("yyyyMMdd"),
				(int)this.txtIdCliente.Value,
				dtFechaContrato.ToString("yyyyMMdd"),
				this.txtTitulo.Text.ToString(),
				this.txtNotasContrato.Text.ToString(), sArchivoProforma, sArchivoContrato);
			this.txtidProformas.Value = Funcion.iEscalarSQL(cdsSeteoF, sSQL);
			#endregion Guarda Proforma

			#region Guarda Entrega Proforma 
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow drr in this.ultraGrid1.Rows.All)
			{
				string sSQLDetalle = string.Format("Exec GuardaFechaEntregaProforma  {0},{1},'{2}','{3}' ", 
					drr.Cells["idFechaEntregaProforma"].Value,
					(int)this.txtidProformas.Value ,
					Convert.ToDateTime(	drr.Cells["Fecha"].Value).ToString("yyyyMMdd"),
					drr.Cells["Descripcion"].Value.ToString());     
						
				Funcion.EjecutaSQL(cdsSeteoF, sSQLDetalle);
			}	
			#endregion Guarda Entrega Proforma 

			#region Bloqueo Controles 
			this.txtidProformas.Enabled = false;
			this.txtNumero.Enabled= false;
			this.dtFecha.Enabled = false;
			this.txtRuc.Enabled = false;
			this.txtDetalle.Enabled = false;
			this.txtCosto.Enabled = false;
			this.txtDescuento.Enabled = false;
			this.txtTotal.Enabled = false;
			this.dtFechaInicia.Enabled = false;
			this.btnArchivoProforma.Enabled = false;

			this.optContrato.Enabled = false;
			this.dtFechaContrato.Enabled = false;
			this.txtTitulo.Enabled = false;
			this.txtNotasContrato.Enabled = false;
			this.btnArchivoContrato.Enabled = false;

			FuncionesProcedimientos.EstadoGrids(false, this.ultraGrid1);			
			#endregion Bloqueo Controles 	
			
			Consultar ((int)this.txtidProformas.Value);
						
			bNuevo = false;
			bEdicion = false;	

			//			if ((int) this.optContrato.Value == 1)
			//			{
			//				if (sCTitulo.Length > 0)
			//				{
			//					//					if (dtCFecha.Date != System.DBNull.Value)
			//					//					{
			//					#region Guarda Fecha Entrega
			//					string SQL = string.Format("Exec GuardaContrato {0},{1},{2},'{3}','{4}',{5},'{6}'",
			//						(int)this.txtidContrato.Value, 
			//						(int)this.txtidProformas.Value, 
			//						(int) this.txtIdCliente.Value, 
			//						Convert.ToDateTime(dtCFecha).ToString("yyyyMMdd"),
			//						sCTitulo, 
			//						(int)this.txtEstado.Value,
			//						sCArchivo);
			//					this.txtidContrato.Value =Funcion.iEscalarSQL(cdsSeteoF,SQL);
			//
			//					#endregion Guarda Fecha Entrega
			//					//					}
			//				}				
			//				else 
			//				{
			//					MessageBox.Show(" Ingresa los datos del contrato ", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
			//					return;
			//				}
			//			}								
		}

		private void btnEditar_Click(object sender, System.EventArgs e)
		{
			Consultar ((int) this.txtidProformas.Value);

			#region Estados
			if ((int)this.txtEstado.Value == 2)
			{
				MessageBox.Show("No puede editar proformas autorizadas", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return;
			}

			if ((int)this.txtEstado.Value == 3)
			{
				MessageBox.Show("No puede editar proformas facturadas", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return;
			}

			if ((int)this.txtEstado.Value == 4)
			{
				MessageBox.Show("No puede editar proformas anuladas", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return;
			}

			if ((int)this.txtEstado.Value == 1)
			{				
				this.txtDetalle.Enabled = true;
				this.txtDescuento.Enabled = true;
				this.txtCosto.Enabled = true;
				this.optContrato.Enabled = true;
				this.dtFechaInicia.Enabled = true;

				if ((int)this.optContrato.Value == 1)
				{
					this.dtFechaContrato.Enabled = true;
					this.txtTitulo.Enabled = true;
					this.txtNotasContrato.Enabled = true;
					this.btnArchivoContrato.Enabled = true;
				}

				this.btnNuevo.Enabled = false;
				this.btnConsultar.Enabled = false;
				this.btnEditar.Enabled = false;
				this.btnAutorizar.Enabled = false;
				this.btnAnular.Enabled = false;
				this.btnCancelar.Enabled = true;
				this.btnGuardar.Enabled = true;
				this.btnArchivoProforma.Enabled = true;

				FuncionesProcedimientos.EstadoGrids(true, this.ultraGrid1);

				bNuevo = false;
				bEdicion = true;
			}			
			#endregion Estados			
		}

		private void dtFecha_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.txtRuc.Focus();
		}

		private void dtFecha_ValueChanged(object sender, System.EventArgs e)
		{
			if (bEdicion)
			{
				this.dtFechaInicia.Value = DateTime.Today;
				this.dtFechaInicia.CalendarInfo.MinDate = Convert.ToDateTime(this.dtFecha.Value) ;
			}
		}

		private void cmbTipoRuc_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.txtRuc.Focus();
		}

		private void txtRuc_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{			
			if (e.KeyCode == Keys.F3)
			{
				using (BuscaClientes Busqueda = new BuscaClientes(this.txtRuc.Text.ToString(), 4, 6, 1, 0, false))
				{
					if (DialogResult.OK == Busqueda.ShowDialog())
					{
						this.txtRuc.Text = Busqueda.Ruc;
						this.txtNumero.Focus();
					}
				}
			}

			if (e.KeyCode == Keys.Enter) this.txtDetalle.Focus();
		}

		private void txtRuc_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			Funcion.CamposNumericos(sender, e);
		}

		private void txtRuc_Validated(object sender, System.EventArgs e)
		{
			if (bEdicion)
			{
				this.txtIdCliente.Value = 0;
				this.txtNombre.Text = "";
				
				CargaCliente();

				if ((int)this.txtIdCliente.Value == 0)
				{					
					FuncionesProcedimientos.AbreFacturacionClientes(4, this.txtRuc.Text.ToString(), true, false);
					CargaCliente();
					this.txtDetalle.Focus();
				}
				else this.txtDetalle.Focus();
			}
		}

		private void CargaCliente()
		{
			string[] sParametros = new string[11];
			sParametros = FuncionesProcedimientos.RetornaParametrosClientesProveedores(true, 4, this.txtRuc.Text.ToString(), cdsSeteoF);

			if (sParametros[0] != null) this.txtIdCliente.Value = int.Parse(sParametros[0]);			
			if (sParametros[3] != null) this.txtRuc.Text = sParametros[3];
			if (sParametros[4] != null) this.txtNombre.Text = sParametros[4];			
		}

		private void txtRuc_Validating(object sender, System.ComponentModel.CancelEventArgs e)
		{
			if (bEdicion)
			{
				if (txtRuc.Text.ToString().Trim().Length == 0)
				{
					MessageBox.Show("Ingrese el Número de Identificación del Cliente", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
					e.Cancel = true;
				}
				else
				{					
					string stExec = string.Format("Select dbo.ValidaIdentificacion ('{0}', {1})", txtRuc.Text, 4);
					string VlCed = Funcion.sEscalarSQL(cdsSeteoF, stExec);
					if (VlCed.StartsWith("Error"))
					{
						MessageBox.Show("Número de Identificación Incorrecta", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
						e.Cancel = true;
					}							
				}
			}
		}
		
		private void ultraGrid1_AfterRowInsert(object sender, Infragistics.Win.UltraWinGrid.RowEventArgs e)
		{
			e.Row.Cells["idFechaentregaProforma"].Value = 0;
			e.Row.Cells["idProforma"].Value = 0;
		}

		private void btnArchivo_Click(object sender, System.EventArgs e)
		{
			try
			{				
				if (sArchivoContrato.Length > 0)
				{
					if (File.Exists(sArchivoContrato))
					{							
						if (DialogResult.Yes == MessageBox.Show(string.Format("El archivo del documento {0} ya fue subido\n\n¿Desea Visualizarlo?", sArchivoContrato), "Point Technology", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2))
						{
							Process.Start(sArchivoContrato, "");
							return;
						}
						else return;
					}
				}

				OpenFileDialog oArchivo = new OpenFileDialog();

				oArchivo.Filter = "Archivos pdf (*.pdf)|*.pdf";
				oArchivo.AddExtension = true;
					
				if (oArchivo.ShowDialog() == System.Windows.Forms.DialogResult.OK && oArchivo.FileName.Length > 0)
				{						
					sArchivoContrato = oArchivo.FileName;
												
					MessageBox.Show("Importación Finalizada Correctamente", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
								
					this.Cursor = Cursors.Default;
				}
			}
			catch (Exception ex)
			{
				MessageBox.Show(string.Format("Commit Exception Type : {0} {1}", ex.GetType(), ex.Message), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
				this.Cursor = Cursors.Default;
				return;
			}
//			//			if((int)this.txtidProformas.Value > 0)
//			//			{
//			//				Consultar((int)this.txtidProformas.Value);
//			//			}
//			if((int)this.txtIdCliente.Value > 0)
//			{				
//				using (frmContrato consulprofos = new frmContrato((int)this.txtIdCliente.Value, (int)this.txtEstado.Value, (int) this.txtidProformas.Value, (int)this.txtidContrato.Value, (DateTime) this.dtFecha.Value,
//								 this.txtNombre.Text.ToString()))
//				{
//					if (DialogResult.OK == consulprofos.ShowDialog())
//					{
//						dtCFecha = (DateTime) consulprofos.dtFecha.Value;
//						sCTitulo  = consulprofos.txtTitulo.Text.ToString();
//						sCArchivo = consulprofos.txtArchivo.Text.ToString();           
//					}				
//				}	
//			}
//			else
//			{
//				MessageBox.Show("Ingrese el Numero de Identificacion del Cliente", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
//				this.txtRuc.Focus();
//				return;
//			}
		}

		private void txtDetalle_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.txtCosto.Focus();
		}

		private void txtcosto_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.txtDescuento.Focus();
		}

		private void txtDescuento_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.dtFechaInicia.Focus();
		}

		private void dtFechaInicia_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) 
			{
				this.ultraGrid1.DisplayLayout.Bands[0].AddNew();
				this.ultraGrid1.ActiveCell = this.ultraGrid1.Rows[0].Cells["Fecha"];
			}
		}

		private void ultraGrid1_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			FuncionesProcedimientos.DesplazamientoGrids(sender,e,this.ultraGrid1);
		}

		private void ultraGrid1_BeforeRowsDeleted(object sender, Infragistics.Win.UltraWinGrid.BeforeRowsDeletedEventArgs e)
		{
			if (DialogResult.Yes ==	MessageBox.Show("¿Esta seguro de Borrar las Líneas Seleccionadas?", "Point Technology", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2))				
			{
				for (int i = 0; i < e.Rows.Length; i++)
				{
					if (e.Rows[i].Cells["idFechaEntregaProforma"].Value != System.DBNull.Value)
					{									
						if ((int)e.Rows[i].Cells["idFechaEntregaProforma"].Value > 0)
						{
							string sSQL = string.Format("Delete From FechaEntregaProforma  Where idFechaEntregaProforma = {0}", (int)e.Rows[i].Cells["idFechaEntregaProforma"].Value);
							Funcion.EjecutaSQL(cdsSeteoF, sSQL, true);
						}						
					}
				}
			}

			e.DisplayPromptMsg = false;
		}

		private void txtcosto_Click(object sender, System.EventArgs e)
		{
			this.txtCosto.SelectAll();
		}

		private void txtcosto_Enter(object sender, System.EventArgs e)
		{
			this.txtCosto.SelectAll();
		}

		private void txtcosto_Validated(object sender, System.EventArgs e)
		{			
			if (bEdicion)
			{
				if (this.txtCosto.Value == System.DBNull.Value) this.txtCosto.Value = 0.00m;

				this.txtTotal.Value = Convert.ToDecimal(this.txtCosto.Value);
			}
		}

		private void txtDescuento_Validated(object sender, System.EventArgs e)
		{
			if (bEdicion)
			{
				if (this.txtDescuento.Value == System.DBNull.Value) this.txtDescuento.Value = 0.00m;
				if (this.txtCosto.Value == System.DBNull.Value) this.txtCosto.Value = 0.00m;
				if (Convert.ToDecimal(this.txtDescuento.Value) > 0)
					this.txtTotal.Value = Convert.ToDecimal(this.txtCosto.Value) - ((Convert.ToDecimal(this.txtCosto.Value) * Convert.ToDecimal(this.txtDescuento.Value))/100);
			}
		}

		private void txtDescuento_Click(object sender, System.EventArgs e)
		{
			this.txtDescuento.SelectAll();
		}

		private void txtDescuento_Enter(object sender, System.EventArgs e)
		{
			this.txtDescuento.SelectAll();
		}

		private void btnAnular_Click(object sender, System.EventArgs e)
		{
			Consultar ((int)this.txtidProformas.Value);

			if (!miAcceso.Anular)
			{
				MessageBox.Show("No tiene acceso a Anular", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return;
			}

			if ((int)this.txtEstado.Value == 4)
			{
				MessageBox.Show("La proforma esta Anulada", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return;
			}

			if ((int)this.txtEstado.Value == 3)
			{
				MessageBox.Show("La proforma esta Facturada", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return;
			}
			
			if ((int)this.txtEstado.Value < 3)
			{				
				using (frmNotas Notas = new frmNotas())
				{					
					if (DialogResult.OK == Notas.ShowDialog())
					{
						DateTime Fechas ;
						Fechas = DateTime.Today;
						string sSQL = string.Format("Update Proformas Set Observacion = '{0}', Estado = 4, FechaEstado = GETDATE(), UsuarioEstado = SUSER_SNAME() Where idProforma = {1}", 
							Notas.txtNotas.Text.ToString(), (int)this.txtidProformas.Value);
						Funcion.EjecutaSQL(cdsSeteoF, sSQL);

//						string SQL = string.Format("Update contrato Set Estado = 3 Where idProforma = {0}", 
//							(int)this.txtidProformas.Value);
//						Funcion.EjecutaSQL(cdsSeteoF, SQL);
					
						Consultar ((int) this.txtidProformas.Value);
					}
				}					
			}
		}

		private void btnAutorizar_Click(object sender, System.EventArgs e)
		{
			Consultar ((int)this.txtidProformas.Value);

			if ((int)this.txtEstado.Value == 4)
			{
				MessageBox.Show("La proforma esta Anulada", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return;
			}

			if ((int)this.txtEstado.Value == 2)
			{
				MessageBox.Show("La proforma esta Autorizada", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return;
			}

			if ((int)this.txtEstado.Value == 3)
			{
				MessageBox.Show("La proforma esta Facturada", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return;
			}

			if ((int)this.txtEstado.Value == 1)
			{
				string sSQL = string.Format("Update Proformas Set  Estado = 2 , FechaEstado = GETDATE(), UsuarioEstado = SUSER_SNAME() Where idProforma = {0}", 
					(int)this.txtidProformas.Value);
				Funcion.EjecutaSQL(cdsSeteoF, sSQL);

//				string SQL = string.Format("Update contrato Set Estado = 2 Where idProforma = {0}", (int)this.txtidProformas.Value);
//				Funcion.EjecutaSQL(cdsSeteoF, SQL);					
			
				Consultar ((int)this.txtidProformas.Value);		
			}
		}

		private void frmProformas_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (!bEdicion)
			{
				if (e.KeyCode == Keys.N) if (this.btnNuevo.Enabled) this.btnNuevo_Click(sender, e);
				if (e.KeyCode == Keys.B) if (this.btnConsultar.Enabled) this.btnConsultar_Click(sender, e);
				if (e.KeyCode == Keys.E) if ((int)this.txtEstado.Value < 1) this.btnEditar_Click(sender, e); 																											
				//				if (e.KeyCode == Keys.I)  if (this.btnImprimir.Enabled) this.btnImprimir_Click(sender, e);
				//				if (e.KeyCode == Keys.A)  if ((int)this.txtEstado.Value == 1) this.btnAnular_Click(sender, e); 																																	 
				if (e.KeyCode == Keys.F5)	this.Consultar((int)this.txtidProformas.Value);
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

		private void optEstados_ValueChanged(object sender, System.EventArgs e)
		{
			if (bEdicion)
			{
				this.dtFechaContrato.Value = System.DBNull.Value;
				this.txtTitulo.Text = "";
				this.txtNotasContrato.Text = "";
				sArchivoContrato = "";
				this.btnArchivoContrato.Enabled = false;

				if ((int)this.optContrato.Value == 1)
				{
					this.dtFechaContrato.Enabled = true;
					this.txtTitulo.Enabled = true;
					this.txtNotasContrato.Enabled = true;
					this.btnArchivoContrato.Enabled = true;
				}
			}
		}

		private void ultraGrid1_AfterCellUpdate(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			if (!bActualiza) return;
						
			bActualiza = false;

			if (e.Cell.Column.ToString() == "Fecha")
			{
				if ((DateTime)e.Cell.Row.Cells["Fecha"].Value < (DateTime)this.dtFechaInicia.Value)
				{
					bActualiza = false;
					e.Cell.Row.Cells["Fecha"].Value = (DateTime)this.dtFechaInicia.Value;
					MessageBox.Show("Las fechas de entrega no pueden ser anteriores a la fecha de inicio", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);										
					bActualiza = true;
				}
			}

			bActualiza = true;
		}

		private void btnArchivoProforma_Click(object sender, System.EventArgs e)
		{
			try
			{				
				if (sArchivoProforma.Length > 0)
				{
					if (File.Exists(sArchivoProforma))
					{							
						if (DialogResult.Yes == MessageBox.Show(string.Format("El archivo del documento {0} ya fue subido\n\n¿Desea Visualizarlo?", sArchivoProforma), "Point Technology", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2))
						{
							Process.Start(sArchivoProforma, "");
							return;
						}
						else return;
					}
				}

				OpenFileDialog oArchivo = new OpenFileDialog();

				oArchivo.Filter = "Archivos pdf (*.pdf)|*.pdf";
				oArchivo.AddExtension = true;
					
				if (oArchivo.ShowDialog() == System.Windows.Forms.DialogResult.OK && oArchivo.FileName.Length > 0)
				{			
					sArchivoProforma = oArchivo.FileName;
												
					MessageBox.Show("Importación Finalizada Correctamente", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
								
					this.Cursor = Cursors.Default;
				}
			}
			catch (Exception ex)
			{
				MessageBox.Show(string.Format("Commit Exception Type : {0} {1}", ex.GetType(), ex.Message), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
				this.Cursor = Cursors.Default;
				return;
			}
		}

		private void ultraGrid1_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}		
	}
}
