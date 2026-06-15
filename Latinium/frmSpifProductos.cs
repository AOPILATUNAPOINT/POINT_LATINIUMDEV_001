using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using DevExpress.XtraEditors;
using System.Data;
using System.Data.SqlClient;
using C1.Data;
using Infragistics.Win.UltraWinGrid;
using System.Globalization;


namespace Latinium
{
	/// <summary>
	/// Descripción breve de frmSpifProductos.
	/// </summary>
	public class frmSpifProductos : System.Windows.Forms.Form
	{
		private System.Windows.Forms.Button btnNuevo;
		private System.Windows.Forms.Button btnConsultar;
		private System.Windows.Forms.Button btnEditar;
		private System.Windows.Forms.Button btnGuardar;
		private System.Windows.Forms.Button btnAnular;
		private System.Windows.Forms.Button btnCancelar;
		private System.Windows.Forms.Button btnSalir;
		private C1.Data.C1DataSet cdsSeteoF;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource1;
		private Infragistics.Win.UltraWinGrid.UltraGrid ultraGrid1;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;
		bool bNuevo = false;
		bool bEdicion = false;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtIdSpifProductos;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtEstado;
		private System.Windows.Forms.Label lblEstado;
		private System.Windows.Forms.Label lblContador;
		private System.Windows.Forms.Label label4;
		private System.Windows.Forms.Label label2;
		private System.Windows.Forms.Label label1;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtDescripcion;
		public Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtHasta;
		public Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtDesde;
		private System.Windows.Forms.Button btnExcel;
		bool bEliminaAlValidar = false;

		public frmSpifProductos()
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
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmSpifProductos));
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("IdDetalleSpifProductos");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idArticulo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn3 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Codigo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn4 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Articulo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn5 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Proteccion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn6 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("CostoUltimo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn7 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Tarjeta");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn8 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Total");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn9 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Precio");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn10 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Margen");
			Infragistics.Win.Appearance appearance1 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("IdDetalleSpifProductos");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idArticulo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Articulo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Proteccion");
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CostoUltimo");
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Tarjeta");
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Total");
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Precio");
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Margen");
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance11 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance12 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance13 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance14 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance15 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton1 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.Appearance appearance16 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton2 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			this.btnNuevo = new System.Windows.Forms.Button();
			this.btnConsultar = new System.Windows.Forms.Button();
			this.btnEditar = new System.Windows.Forms.Button();
			this.btnGuardar = new System.Windows.Forms.Button();
			this.btnAnular = new System.Windows.Forms.Button();
			this.btnCancelar = new System.Windows.Forms.Button();
			this.btnSalir = new System.Windows.Forms.Button();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.ultraDataSource1 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.ultraGrid1 = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.txtIdSpifProductos = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtEstado = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.lblEstado = new System.Windows.Forms.Label();
			this.lblContador = new System.Windows.Forms.Label();
			this.label4 = new System.Windows.Forms.Label();
			this.label2 = new System.Windows.Forms.Label();
			this.label1 = new System.Windows.Forms.Label();
			this.txtDescripcion = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.dtHasta = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.dtDesde = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.btnExcel = new System.Windows.Forms.Button();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdSpifProductos)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtEstado)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDescripcion)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtHasta)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtDesde)).BeginInit();
			this.SuspendLayout();
			// 
			// btnNuevo
			// 
			this.btnNuevo.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnNuevo.CausesValidation = false;
			this.btnNuevo.Image = ((System.Drawing.Image)(resources.GetObject("btnNuevo.Image")));
			this.btnNuevo.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnNuevo.Location = new System.Drawing.Point(8, 472);
			this.btnNuevo.Name = "btnNuevo";
			this.btnNuevo.Size = new System.Drawing.Size(76, 23);
			this.btnNuevo.TabIndex = 313;
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
			this.btnConsultar.Location = new System.Drawing.Point(88, 472);
			this.btnConsultar.Name = "btnConsultar";
			this.btnConsultar.Size = new System.Drawing.Size(76, 23);
			this.btnConsultar.TabIndex = 312;
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
			this.btnEditar.Location = new System.Drawing.Point(176, 472);
			this.btnEditar.Name = "btnEditar";
			this.btnEditar.Size = new System.Drawing.Size(76, 23);
			this.btnEditar.TabIndex = 311;
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
			this.btnGuardar.Location = new System.Drawing.Point(264, 472);
			this.btnGuardar.Name = "btnGuardar";
			this.btnGuardar.Size = new System.Drawing.Size(76, 23);
			this.btnGuardar.TabIndex = 310;
			this.btnGuardar.Text = "&Guardar";
			this.btnGuardar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnGuardar.Click += new System.EventHandler(this.btnGuardar_Click);
			// 
			// btnAnular
			// 
			this.btnAnular.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnAnular.CausesValidation = false;
			this.btnAnular.Enabled = false;
			this.btnAnular.Image = ((System.Drawing.Image)(resources.GetObject("btnAnular.Image")));
			this.btnAnular.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnAnular.Location = new System.Drawing.Point(560, 472);
			this.btnAnular.Name = "btnAnular";
			this.btnAnular.Size = new System.Drawing.Size(76, 23);
			this.btnAnular.TabIndex = 309;
			this.btnAnular.Text = "&Anular";
			this.btnAnular.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnAnular.Visible = false;
			// 
			// btnCancelar
			// 
			this.btnCancelar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnCancelar.CausesValidation = false;
			this.btnCancelar.Enabled = false;
			this.btnCancelar.Image = ((System.Drawing.Image)(resources.GetObject("btnCancelar.Image")));
			this.btnCancelar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnCancelar.Location = new System.Drawing.Point(352, 472);
			this.btnCancelar.Name = "btnCancelar";
			this.btnCancelar.Size = new System.Drawing.Size(76, 23);
			this.btnCancelar.TabIndex = 308;
			this.btnCancelar.Text = "&Cancelar";
			this.btnCancelar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnCancelar.Click += new System.EventHandler(this.btnCancelar_Click);
			// 
			// btnSalir
			// 
			this.btnSalir.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnSalir.CausesValidation = false;
			this.btnSalir.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnSalir.Location = new System.Drawing.Point(432, 472);
			this.btnSalir.Name = "btnSalir";
			this.btnSalir.Size = new System.Drawing.Size(76, 23);
			this.btnSalir.TabIndex = 307;
			this.btnSalir.Text = "&Salir";
			this.btnSalir.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnSalir.Click += new System.EventHandler(this.btnSalir_Click);
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
			// ultraDataSource1
			// 
			ultraDataColumn1.DataType = typeof(int);
			ultraDataColumn2.DataType = typeof(int);
			ultraDataColumn5.DataType = typeof(System.Decimal);
			ultraDataColumn5.DefaultValue = new System.Decimal(new int[] {
																																		 0,
																																		 0,
																																		 0,
																																		 0});
			ultraDataColumn6.DataType = typeof(System.Decimal);
			ultraDataColumn6.DefaultValue = new System.Decimal(new int[] {
																																		 0,
																																		 0,
																																		 0,
																																		 0});
			ultraDataColumn7.DataType = typeof(System.Decimal);
			ultraDataColumn8.DataType = typeof(int);
			ultraDataColumn9.DataType = typeof(System.Decimal);
			ultraDataColumn10.DataType = typeof(System.Decimal);
			this.ultraDataSource1.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn1,
																																 ultraDataColumn2,
																																 ultraDataColumn3,
																																 ultraDataColumn4,
																																 ultraDataColumn5,
																																 ultraDataColumn6,
																																 ultraDataColumn7,
																																 ultraDataColumn8,
																																 ultraDataColumn9,
																																 ultraDataColumn10});
			// 
			// ultraGrid1
			// 
			this.ultraGrid1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left)));
			this.ultraGrid1.Cursor = System.Windows.Forms.Cursors.Default;
			this.ultraGrid1.DataSource = this.ultraDataSource1;
			appearance1.BackColor = System.Drawing.Color.White;
			this.ultraGrid1.DisplayLayout.Appearance = appearance1;
			this.ultraGrid1.DisplayLayout.AutoFitColumns = true;
			ultraGridBand1.AddButtonCaption = "DetPromocionVendedores";
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn1.Width = 124;
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridColumn2.Hidden = true;
			ultraGridColumn2.Width = 57;
			ultraGridColumn3.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn3.Header.VisiblePosition = 2;
			ultraGridColumn3.Width = 139;
			ultraGridColumn4.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn4.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn4.Header.VisiblePosition = 3;
			ultraGridColumn4.Width = 278;
			appearance2.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn5.CellAppearance = appearance2;
			ultraGridColumn5.Format = "#,##0.00";
			ultraGridColumn5.Header.Caption = "Protección";
			ultraGridColumn5.Header.VisiblePosition = 5;
			ultraGridColumn5.MaskInput = "";
			ultraGridColumn5.Width = 84;
			ultraGridColumn6.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance3.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn6.CellAppearance = appearance3;
			ultraGridColumn6.Format = "#,##0.00";
			ultraGridColumn6.Header.Caption = "Costo";
			ultraGridColumn6.Header.VisiblePosition = 4;
			ultraGridColumn6.MaskInput = "";
			ultraGridColumn6.MaxValue = 100;
			ultraGridColumn6.MinValue = null;
			ultraGridColumn6.Width = 82;
			ultraGridColumn7.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance4.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn7.CellAppearance = appearance4;
			ultraGridColumn7.Format = "#,##0.00";
			ultraGridColumn7.Header.VisiblePosition = 6;
			ultraGridColumn7.Width = 88;
			ultraGridColumn8.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance5.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn8.CellAppearance = appearance5;
			ultraGridColumn8.Header.Caption = "Stock";
			ultraGridColumn8.Header.VisiblePosition = 7;
			ultraGridColumn8.Width = 94;
			ultraGridColumn9.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance6.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn9.CellAppearance = appearance6;
			ultraGridColumn9.Format = "#,##0.00";
			ultraGridColumn9.Header.Caption = "Nuevo Costo";
			ultraGridColumn9.Header.VisiblePosition = 8;
			ultraGridColumn9.Width = 99;
			ultraGridColumn10.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance7.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn10.CellAppearance = appearance7;
			ultraGridColumn10.Format = "#,##0.00";
			ultraGridColumn10.Header.VisiblePosition = 9;
			ultraGridColumn10.Width = 107;
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
			ultraGridBand1.Header.Caption = "ARTÍCULOS";
			ultraGridBand1.HeaderVisible = true;
			this.ultraGrid1.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			appearance8.ForeColor = System.Drawing.Color.Black;
			appearance8.ForeColorDisabled = System.Drawing.Color.Black;
			this.ultraGrid1.DisplayLayout.Override.ActiveRowAppearance = appearance8;
			this.ultraGrid1.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.TemplateOnBottom;
			this.ultraGrid1.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.True;
			this.ultraGrid1.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance9.BackColor = System.Drawing.Color.Transparent;
			this.ultraGrid1.DisplayLayout.Override.CardAreaAppearance = appearance9;
			appearance10.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance10.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance10.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance10.FontData.BoldAsString = "True";
			appearance10.FontData.Name = "Arial";
			appearance10.FontData.SizeInPoints = 10F;
			appearance10.ForeColor = System.Drawing.Color.White;
			appearance10.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.ultraGrid1.DisplayLayout.Override.HeaderAppearance = appearance10;
			appearance11.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance11.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance11.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.RowAlternateAppearance = appearance11;
			appearance12.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance12.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance12.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.RowSelectorAppearance = appearance12;
			appearance13.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance13.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance13.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.SelectedRowAppearance = appearance13;
			this.ultraGrid1.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.ultraGrid1.Location = new System.Drawing.Point(8, 88);
			this.ultraGrid1.Name = "ultraGrid1";
			this.ultraGrid1.Size = new System.Drawing.Size(992, 376);
			this.ultraGrid1.TabIndex = 314;
			this.ultraGrid1.AfterRowInsert += new Infragistics.Win.UltraWinGrid.RowEventHandler(this.ultraGrid1_AfterRowInsert);
			this.ultraGrid1.AfterColPosChanged += new Infragistics.Win.UltraWinGrid.AfterColPosChangedEventHandler(this.ultraGrid1_AfterColPosChanged);
			this.ultraGrid1.KeyDown += new System.Windows.Forms.KeyEventHandler(this.ultraGrid1_KeyDown);
			this.ultraGrid1.AfterCellUpdate += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.ultraGrid1_AfterCellUpdate);
			this.ultraGrid1.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.ultraGrid1_InitializeLayout);
			// 
			// txtIdSpifProductos
			// 
			this.txtIdSpifProductos.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtIdSpifProductos.Enabled = false;
			this.txtIdSpifProductos.Location = new System.Drawing.Point(944, 16);
			this.txtIdSpifProductos.Name = "txtIdSpifProductos";
			this.txtIdSpifProductos.PromptChar = ' ';
			this.txtIdSpifProductos.Size = new System.Drawing.Size(16, 21);
			this.txtIdSpifProductos.TabIndex = 862;
			this.txtIdSpifProductos.Visible = false;
			// 
			// txtEstado
			// 
			this.txtEstado.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtEstado.Enabled = false;
			this.txtEstado.Location = new System.Drawing.Point(904, 16);
			this.txtEstado.Name = "txtEstado";
			this.txtEstado.PromptChar = ' ';
			this.txtEstado.Size = new System.Drawing.Size(16, 21);
			this.txtEstado.TabIndex = 861;
			this.txtEstado.Visible = false;
			// 
			// lblEstado
			// 
			this.lblEstado.AutoSize = true;
			this.lblEstado.Font = new System.Drawing.Font("Tahoma", 12F, System.Drawing.FontStyle.Bold);
			this.lblEstado.ForeColor = System.Drawing.Color.Firebrick;
			this.lblEstado.Location = new System.Drawing.Point(400, 48);
			this.lblEstado.Name = "lblEstado";
			this.lblEstado.Size = new System.Drawing.Size(0, 23);
			this.lblEstado.TabIndex = 1002;
			this.lblEstado.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// lblContador
			// 
			this.lblContador.AutoSize = true;
			this.lblContador.Font = new System.Drawing.Font("Tahoma", 12F, System.Drawing.FontStyle.Bold);
			this.lblContador.ForeColor = System.Drawing.Color.Firebrick;
			this.lblContador.Location = new System.Drawing.Point(664, 48);
			this.lblContador.Name = "lblContador";
			this.lblContador.Size = new System.Drawing.Size(0, 23);
			this.lblContador.TabIndex = 1001;
			this.lblContador.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// label4
			// 
			this.label4.AutoSize = true;
			this.label4.Location = new System.Drawing.Point(8, 16);
			this.label4.Name = "label4";
			this.label4.Size = new System.Drawing.Size(44, 16);
			this.label4.TabIndex = 995;
			this.label4.Text = "Nombre";
			this.label4.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label2
			// 
			this.label2.AutoSize = true;
			this.label2.Location = new System.Drawing.Point(200, 48);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(34, 16);
			this.label2.TabIndex = 997;
			this.label2.Text = "Hasta";
			this.label2.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label1
			// 
			this.label1.AutoSize = true;
			this.label1.Location = new System.Drawing.Point(8, 48);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(37, 16);
			this.label1.TabIndex = 996;
			this.label1.Text = "Desde";
			this.label1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtDescripcion
			// 
			appearance14.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtDescripcion.Appearance = appearance14;
			this.txtDescripcion.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtDescripcion.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtDescripcion.Enabled = false;
			this.txtDescripcion.Location = new System.Drawing.Point(64, 16);
			this.txtDescripcion.Multiline = true;
			this.txtDescripcion.Name = "txtDescripcion";
			this.txtDescripcion.Size = new System.Drawing.Size(488, 21);
			this.txtDescripcion.TabIndex = 994;
			// 
			// dtHasta
			// 
			appearance15.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtHasta.Appearance = appearance15;
			this.dtHasta.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton1.Caption = "Today";
			this.dtHasta.DateButtons.Add(dateButton1);
			this.dtHasta.Enabled = false;
			this.dtHasta.Format = "dd/MM/yyyy ";
			this.dtHasta.Location = new System.Drawing.Point(248, 48);
			this.dtHasta.Name = "dtHasta";
			this.dtHasta.NonAutoSizeHeight = 23;
			this.dtHasta.Size = new System.Drawing.Size(124, 21);
			this.dtHasta.SpinButtonsVisible = true;
			this.dtHasta.TabIndex = 999;
			this.dtHasta.Value = ((object)(resources.GetObject("dtHasta.Value")));
			// 
			// dtDesde
			// 
			appearance16.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtDesde.Appearance = appearance16;
			this.dtDesde.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton2.Caption = "Today";
			this.dtDesde.DateButtons.Add(dateButton2);
			this.dtDesde.Enabled = false;
			this.dtDesde.Format = "dd/MM/yyyy ";
			this.dtDesde.Location = new System.Drawing.Point(64, 48);
			this.dtDesde.Name = "dtDesde";
			this.dtDesde.NonAutoSizeHeight = 23;
			this.dtDesde.Size = new System.Drawing.Size(124, 21);
			this.dtDesde.SpinButtonsVisible = true;
			this.dtDesde.TabIndex = 998;
			this.dtDesde.Value = ((object)(resources.GetObject("dtDesde.Value")));
			// 
			// btnExcel
			// 
			this.btnExcel.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btnExcel.Image = ((System.Drawing.Image)(resources.GetObject("btnExcel.Image")));
			this.btnExcel.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnExcel.Location = new System.Drawing.Point(912, 48);
			this.btnExcel.Name = "btnExcel";
			this.btnExcel.Size = new System.Drawing.Size(88, 23);
			this.btnExcel.TabIndex = 1023;
			this.btnExcel.Text = "Excel";
			this.btnExcel.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnExcel.Click += new System.EventHandler(this.btnExcel_Click);
			// 
			// frmSpifProductos
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(1008, 502);
			this.Controls.Add(this.btnExcel);
			this.Controls.Add(this.lblEstado);
			this.Controls.Add(this.lblContador);
			this.Controls.Add(this.label4);
			this.Controls.Add(this.label2);
			this.Controls.Add(this.label1);
			this.Controls.Add(this.txtDescripcion);
			this.Controls.Add(this.dtHasta);
			this.Controls.Add(this.dtDesde);
			this.Controls.Add(this.txtIdSpifProductos);
			this.Controls.Add(this.txtEstado);
			this.Controls.Add(this.ultraGrid1);
			this.Controls.Add(this.btnNuevo);
			this.Controls.Add(this.btnConsultar);
			this.Controls.Add(this.btnEditar);
			this.Controls.Add(this.btnGuardar);
			this.Controls.Add(this.btnAnular);
			this.Controls.Add(this.btnCancelar);
			this.Controls.Add(this.btnSalir);
			this.Name = "frmSpifProductos";
			this.Text = "Proteccion de Precios";
			this.Load += new System.EventHandler(this.frmSpifProductos_Load);
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdSpifProductos)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtEstado)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDescripcion)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtHasta)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtDesde)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion
		private Acceso miAcceso;

		private void UnloadMe()
		{
			this.Close();
		}

		private void frmSpifProductos_Load(object sender, System.EventArgs e)
		{
			miAcceso = new Acceso(cdsSeteoF, "200409");

			if (!Funcion.Permiso("1830", cdsSeteoF))
			{				
				MessageBox.Show("No puede ingresar a Proeteccion Precios", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				BeginInvoke(new MethodInvoker(UnloadMe));
				return;		
			}
			ModificaGrilla(false);
		}

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);
		}
		private void ModificaGrilla(bool bModifica)
		{
			ultraGrid1.Enabled = true;
			if (bModifica)
			{
				ultraGrid1.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
				ultraGrid1.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.True;
			}
			else
			{
				ultraGrid1.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.False;
				ultraGrid1.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			}
		}

		private void btnCancelar_Click(object sender, System.EventArgs e)
		{
			this.txtIdSpifProductos.Value = 0;
			this.txtEstado.Value = 0;
			this.dtDesde.Value = System.DBNull.Value;
			this.dtHasta.Value = System.DBNull.Value;
			this.txtDescripcion.Text = "";
			this.lblEstado.Text = "";
			
			this.dtDesde.Enabled = false;
			this.dtHasta.Enabled = false;
			this.txtDescripcion.Enabled = false;
	

			this.btnNuevo.Enabled = true;
			this.btnConsultar.Enabled = true;
			this.btnEditar.Enabled = false;
			this.btnGuardar.Enabled = false;
			this.btnAnular.Enabled = false;
			this.btnCancelar.Enabled = false;
			this.btnExcel.Enabled=true;

			ModificaGrilla(false);
			
			
			this.ultraGrid1.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Exec ConsultaSpifProductos 0");
		

			bNuevo = false;
			bEdicion = false;
      
			this.btnNuevo.Focus();
		}

		private void btnNuevo_Click(object sender, System.EventArgs e)
		{
			this.btnCancelar_Click(sender, e);

			this.dtDesde.Enabled = true;
			this.dtHasta.Enabled = true;
			this.txtDescripcion.Enabled = true;
			this.txtEstado.Value = 0;

			bNuevo = true;
			bEdicion = true;

			this.dtDesde.Value = FuncionesProcedimientos.dtRetornaFechaServidor(cdsSeteoF);	
			this.dtHasta.Value = FuncionesProcedimientos.dtRetornaFechaServidor(cdsSeteoF);	

			DateTime dtInicio = (DateTime)this.dtDesde.Value;
			DateTime dtFin = (DateTime)this.dtHasta.Value;

			ModificaGrilla(true);
			
			this.btnNuevo.Enabled = false;
			this.btnConsultar.Enabled = false;
			this.btnEditar.Enabled = false;
			this.btnGuardar.Enabled = true;
			this.btnAnular.Enabled = false;			
			this.btnCancelar.Enabled = true;

			this.dtDesde.Focus();
		
		}

		private void btnConsultar_Click(object sender, System.EventArgs e)
		{
			this.btnCancelar_Click(sender, e);

			using (frmBuscaSpifProductos consultab = new frmBuscaSpifProductos())
				if (DialogResult.OK == consultab.ShowDialog())
					this.Consultar((int)consultab.ultraGrid1.ActiveRow.Cells["IdMaestroSpifProductos"].Value);

		}
		private void Consultar (int IdMaestroSpifProductos)
		{
			#region Consulta Individual
			string ssql= string.Format ("Exec ConsultaIndividualMaestroSpifProductos {0}",IdMaestroSpifProductos); 
			SqlDataReader dr = Funcion.rEscalarSQL(cdsSeteoF,ssql,true);
			dr.Read();
			this.txtIdSpifProductos.Value= dr.GetInt32 (0);
			if (dr.GetValue(1) != System.DBNull.Value)this.txtDescripcion.Text = dr.GetString (1);
			if (dr.GetValue(2) != System.DBNull.Value)this.dtDesde.Value = dr.GetDateTime (2);
			if (dr.GetValue(3) != System.DBNull.Value)this.dtHasta.Value = dr.GetDateTime (3);
			//if (dr.GetValue(4) != System.DBNull.Value)this.txtestado.Value = dr.GetInt32 (4);			
			dr.Close();

			this.ultraGrid1.DataSource = Funcion.dvProcedimiento(cdsSeteoF, string.Format("Exec ConsultaIndividualDetalleSpifProductos {0}", IdMaestroSpifProductos));
			#endregion Consulta Individual

			//Estado ((int)this.txtestado.Value);

			this.btnNuevo.Enabled = false;
			this.btnConsultar.Enabled = false;
			if (miAcceso.Editar) this.btnEditar.Enabled = true;

			this.btnCancelar.Enabled = true;
			this.btnGuardar.Enabled = false;		

		}

		private void ultraGrid1_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			FuncionesProcedimientos.DesplazamientoGrids(sender, e, this.ultraGrid1);

			#region Codigo
			if (e.KeyCode == Keys.F3)
			{	
				if (this.ultraGrid1.ActiveCell.Column.Key.ToString().ToUpper() == "CODIGO" || this.ultraGrid1.ActiveCell.Column.Key.ToString().ToUpper() == "ARTICULO")
				{																								
					using (BuscaArticulos miBusqueda = new BuscaArticulos("", 1, 0)) 
					{
						if (DialogResult.OK == miBusqueda.ShowDialog())
						{
							this.ultraGrid1.ActiveRow.Cells["Codigo"].Value = miBusqueda.sCodigo;
						}
					}
				}							
				
			}			
			#endregion Codigo

		}

		private void ultraGrid1_AfterColPosChanged(object sender, Infragistics.Win.UltraWinGrid.AfterColPosChangedEventArgs e)
		{
		
		}

		private void ultraGrid1_AfterCellUpdate(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			if (e.Cell.Column.ToString() == "Codigo")
			{
				#region Busca Producto
				string sSQLContador = string.Format("Exec SpifBuscaArticulo '{0}'", e.Cell.Row.Cells["Codigo"].Value.ToString());
				int iContador = Funcion.iEscalarSQL(cdsSeteoF, sSQLContador);

				if (iContador == 1)
				{
					string sSQLInfo = string.Format("Exec SpifInfoArticulo '{0}'", e.Cell.Row.Cells["Codigo"].Value.ToString());
					SqlDataReader dr = Funcion.rEscalarSQL(cdsSeteoF, sSQLInfo, true);
					dr.Read();
					if (dr.HasRows)
					{
						e.Cell.Row.Cells["IdDetalleSpifProductos"].Value = dr.GetInt32(0);
						e.Cell.Row.Cells["idArticulo"].Value = dr.GetInt32(1);
						//e.Cell.Row.Cells["Codigo"].Value = dr.GetString(1);
						e.Cell.Row.Cells["Articulo"].Value = dr.GetString(3);
						e.Cell.Row.Cells["Proteccion"].Value = dr.GetDecimal(4);
						e.Cell.Row.Cells["CostoUltimo"].Value = dr.GetDecimal(5);
						e.Cell.Row.Cells["Tarjeta"].Value = dr.GetDecimal(6);
						e.Cell.Row.Cells["Total"].Value = dr.GetInt32(7);
						e.Cell.Row.Cells["Precio"].Value = dr.GetDecimal(8);
						e.Cell.Row.Cells["Margen"].Value = dr.GetDecimal(9);
						
					}
					dr.Close();

				}
				else 
				{
					using (BuscaArticulos miBusqueda = new BuscaArticulos("", 1, 0)) //(int)this.cmbBodega.Value
					{
						if (DialogResult.OK == miBusqueda.ShowDialog())
						{
							this.ultraGrid1.ActiveRow.Cells["Codigo"].Value = miBusqueda.sCodigo;
						}
					}

				}
				#endregion Busca Cuenta	
			}
			#region Tarjeta
			if (e.Cell.Column.ToString() == "Tarjeta")
			{
				if (e.Cell.Row.Cells["Tarjeta"].Value != System.DBNull.Value)
				{
					if ((decimal)e.Cell.Row.Cells["Tarjeta"].Value > 0.00m)
					{
						Sumatorias(e);
					}
				}
			}
			#endregion Tarjeta

			#region CostoUltimo
			if (e.Cell.Column.ToString() == "CostoUltimo")
			{
				if (e.Cell.Row.Cells["CostoUltimo"].Value != System.DBNull.Value)
				{
					if ((decimal)e.Cell.Row.Cells["CostoUltimo"].Value > 0.00m)
					{
						Sumatorias(e);
					}
				}
			}
			#endregion CostoUltimo

			#region Proteccion
			if (e.Cell.Column.ToString() == "Proteccion")
			{
				if (e.Cell.Row.Cells["Proteccion"].Value != System.DBNull.Value)
				{
					if ((decimal)e.Cell.Row.Cells["Proteccion"].Value > 0.00m)
					{
							Sumatorias(e);
					}
				}
			}
			#endregion Proteccion


	

		}

		private void btnGuardar_Click(object sender, System.EventArgs e)
		{
			#region Validacion
			if (!Validacion.vbTexto(this.txtDescripcion, 5, 100, "Descripción")) return;
			if (!Validacion.vbFechaEnDocumentos(this.dtDesde, "Ingrese la fecha Inicial ", true, "SPIF PRODUCTOS", cdsSeteoF)) return;
			if (!Validacion.vbFechaEnDocumentos(this.dtHasta, "Ingrese la fecha Final ", true, "SPIF PRODUCTOS", cdsSeteoF)) return;
			if (!Validacion.vbComparaFechas(Convert.ToDateTime(this.dtDesde.Value), Convert.ToDateTime(this.dtHasta.Value), this.dtDesde, "La Fecha Inicial no puede ser mayor a la Fecha Final", false)) return;

			if ( (int)this.ultraGrid1.Rows.Count == 0)
			{
				MessageBox.Show("Ingrese los Articulos de la Promoción ", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return;
			}		
			#endregion Validacion

			#region Valida Proteccion vacios
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow drr in this.ultraGrid1.Rows.All)
			{
				if (Convert.ToDecimal(drr.Cells["Proteccion"].Value) <= 0 )
				{
					this.ultraGrid1.ActiveRow = drr;
					this.ultraGrid1.ActiveRow.Selected = true;
					MessageBox.Show("El campo proteccion no puede ser 0 ", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					return;
				}
			}
			#endregion Valida Proteccion vacios

			#region Valida detalle articulo repetido
			int iArticulo = 0;
			int iCont = 0;

			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.ultraGrid1.Rows.All)
			{         
				iArticulo = Convert.ToInt32(dr.Cells["idArticulo"].Value); 

				foreach(Infragistics.Win.UltraWinGrid.UltraGridRow drr in this.ultraGrid1.Rows.All)
				{
					if ( iArticulo == Convert.ToInt32(drr.Cells["idArticulo"].Value))
					{					
						iCont = iCont + 1;
						if (iCont > 1 )
						{
							this.ultraGrid1.ActiveRow = dr;
							this.ultraGrid1.ActiveRow.Selected = true;
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

			#region Guarda Maestro
			string sSQLl = string.Format("Exec GuardaMaestroSpifProductos {0}, '{1}' , '{2}' ,'{3}', {4}",
				(int)this.txtIdSpifProductos.Value,
				this.txtDescripcion.Text,
				Convert.ToDateTime(this.dtDesde.Value).ToString("yyyyMMdd"),
				Convert.ToDateTime(this.dtHasta.Value).ToString("yyyyMMdd"),
				0);
			#endregion Guarda Maestro

				this.txtIdSpifProductos.Value=Funcion.iEscalarSQL(cdsSeteoF,sSQLl);

			#region Guarda detalle 
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow drr in this.ultraGrid1.Rows.All)
			{	
				string sSQLDetalle = string.Format("Exec GuardaDetalleSpifProductos  {0},{1},{2},'{3}','{4}',{5},{6},{7},{8},{9},{10}", 
					(int)drr.Cells["IdDetalleSpifProductos"].Value,
					(int)this.txtIdSpifProductos.Value,
					(int)drr.Cells["idArticulo"].Value,
					(string)drr.Cells["Codigo"].Value.ToString(),
					(string)drr.Cells["Articulo"].Value.ToString(),
					(decimal)drr.Cells["CostoUltimo"].Value,
					(decimal)drr.Cells["Proteccion"].Value,
					(decimal)drr.Cells["Tarjeta"].Value,
					(int)drr.Cells["Total"].Value,
					(decimal)drr.Cells["Precio"].Value,
					(decimal)drr.Cells["Margen"].Value);
				Funcion.EjecutaSQL(cdsSeteoF, sSQLDetalle);
			}

			#endregion Guarda detalle 

			#region Controles
			Consultar ((int) this.txtIdSpifProductos.Value);
			//this.txtidFacturacionPromociones.Enabled = false;
			this.txtDescripcion.Enabled = false;
			this.dtDesde.Enabled = false;
			this.dtHasta.Enabled = false;
			//this.cmbSubGrupoDet.Enabled = false;

			FuncionesProcedimientos.EstadoGrids(false, this.ultraGrid1);

			if (miAcceso.Nuevo) this.btnNuevo.Enabled = true;
			if (miAcceso.Buscar) this.btnConsultar.Enabled = true;
			if (miAcceso.Editar) this.btnEditar.Enabled = true;
			this.btnGuardar.Enabled = false;
			this.btnCancelar.Enabled = true;
	
			bNuevo = false;
			bEdicion = false;
			#endregion Controles	 
		}

		private void btnEditar_Click(object sender, System.EventArgs e)
		{
			this.txtDescripcion.Enabled = true;
			this.dtHasta.Enabled = true;
			this.dtDesde.Enabled = true;
			FuncionesProcedimientos.EstadoGrids(true, this.ultraGrid1);


			this.btnNuevo.Enabled = false;
			this.btnConsultar.Enabled = false;
			this.btnEditar.Enabled = false;
			this.btnGuardar.Enabled = true;
			this.btnCancelar.Enabled = true;

			bNuevo = false;
			bEdicion = true;
			this.txtDescripcion.Focus();
		}
		private void Sumatorias(Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			decimal margen =0.00m;
			e.Cell.Row.Cells["Margen"].Value = 0.00m;

			#region Calculos
			e.Cell.Row.Cells["Precio"].Value = Convert.ToDecimal(e.Cell.Row.Cells["CostoUltimo"].Value)-
				Convert.ToDecimal(e.Cell.Row.Cells["Proteccion"].Value); 

			if (e.Cell.Column.ToString() == "Proteccion")
			{
				e.Cell.Row.Cells["Margen"].Value=Math.Round((Convert.ToDecimal(e.Cell.Row.Cells["Tarjeta"].Value)-
					Convert.ToDecimal(e.Cell.Row.Cells["Precio"].Value))/Convert.ToDecimal(e.Cell.Row.Cells["Tarjeta"].Value), 2);
			}
			#endregion Calculos		
		}

		private void ultraGrid1_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{

		}

		private void ultraGrid1_AfterRowInsert(object sender, Infragistics.Win.UltraWinGrid.RowEventArgs e)
		{
			e.Row.Cells["Proteccion"].Value = 0.00m;
			e.Row.Cells["Precio"].Value = 0.00m;
			e.Row.Cells["Margen"].Value = 0.00m;
			e.Row.Cells["Total"].Value = 0.00m; 
			e.Row.Cells["Tarjeta"].Value = 0.00m; 

//			//e.Cell.Row.Cells["NuevoCosto"].Value = 0.00m;
//			e.Cell.Row.Cells["Proteccion"].Value = 0.00m;
//			e.Cell.Row.Cells["Margen"].Value = 0.00m; 
//			//e.Cell.Row.Cells["Total"].Value = 0.00m; 
//			//e.Cell.Row.Cells["Tarjeta"].Value = 0.00m; 
//			//e.Cell.Row.Cells["CostoUltimo"].Value = 0.00m; 
//			e.Cell.Row.Cells["Precio"].Value = 0.00m;

			
		}

		private void btnSalir_Click(object sender, System.EventArgs e)
		{
			this.Close();	
		}

		private void btnExcel_Click(object sender, System.EventArgs e)
		{
			FuncionesProcedimientos.ExportaExcel(this.ultraGrid1);
		}
	}
}
