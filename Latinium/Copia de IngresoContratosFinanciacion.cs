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
	/// Descripción breve de IngresoContratosFinanciacion.
	/// </summary>
	public class IngresoContratosFinanciacion : DevExpress.XtraEditors.XtraForm
	{
		private Infragistics.Win.Misc.UltraButton btnServicios;
		private Infragistics.Win.UltraWinGrid.UltraGrid ultraGrid2;
		private Infragistics.Win.Misc.UltraButton btnTabla;
		private Infragistics.Win.UltraWinGrid.UltraGrid ultraGrid1;
		private Infragistics.Win.Misc.UltraButton btnFactura;
		private Infragistics.Win.Misc.UltraButton btSalir;
		private Infragistics.Win.UltraWinGrid.UltraDropDown cmbServicio;
		private C1.Data.C1DataSet cdsServicio;
		private C1.Data.C1DataView cdvFinancia;
		private C1.Data.C1DataSet cdsFinancia;
		private C1.Data.C1DataView cdvServicio;
		private System.Windows.Forms.ErrorProvider errorProvider;
		private C1.Data.C1DataSet cdsContratos;
		private C1.Data.C1DataView cdvContratos;
		private Infragistics.Win.Misc.UltraButton btGuardar;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public IngresoContratosFinanciacion()
		{
			//
			// Necesario para admitir el Diseñador de Windows Forms
			//
			InitializeComponent();

			//
			// TODO: agregar código de constructor después de llamar a InitializeComponent
			//
		}
		int vIdContrato=0;
		int VidC=0;
		public IngresoContratosFinanciacion(int idContrato)
		{
			//
			// Necesario para admitir el Diseñador de Windows Forms
			//
			InitializeComponent();
			vIdContrato= idContrato;
			VidC= idContrato;
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
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("ContratosFinanciacion", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idFinanciacionS");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idServicio", -1, "cmbServicio");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cuotas");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Anticipo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Valor");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Porcentaje");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaD");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaH");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idContrato");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("VTrigger");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn11 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("valorU");
			Infragistics.Win.Appearance appearance1 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("ContratosServicios", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn12 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idServicios");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn13 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn14 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Tipo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn15 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Valor");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn16 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idArticulo");
			this.btnServicios = new Infragistics.Win.Misc.UltraButton();
			this.ultraGrid2 = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.cdsFinancia = new C1.Data.C1DataSet();
			this.btnTabla = new Infragistics.Win.Misc.UltraButton();
			this.ultraGrid1 = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.btnFactura = new Infragistics.Win.Misc.UltraButton();
			this.btSalir = new Infragistics.Win.Misc.UltraButton();
			this.cmbServicio = new Infragistics.Win.UltraWinGrid.UltraDropDown();
			this.cdvServicio = new C1.Data.C1DataView();
			this.cdsServicio = new C1.Data.C1DataSet();
			this.cdvFinancia = new C1.Data.C1DataView();
			this.errorProvider = new System.Windows.Forms.ErrorProvider();
			this.cdsContratos = new C1.Data.C1DataSet();
			this.cdvContratos = new C1.Data.C1DataView();
			this.btGuardar = new Infragistics.Win.Misc.UltraButton();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid2)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsFinancia)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbServicio)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvServicio)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsServicio)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvFinancia)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsContratos)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvContratos)).BeginInit();
			this.SuspendLayout();
			// 
			// btnServicios
			// 
			this.btnServicios.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.btnServicios.Location = new System.Drawing.Point(8, 8);
			this.btnServicios.Name = "btnServicios";
			this.btnServicios.Size = new System.Drawing.Size(112, 23);
			this.btnServicios.TabIndex = 125;
			this.btnServicios.Text = "Crea Servicio";
			this.btnServicios.Click += new System.EventHandler(this.btnServicios_Click);
			// 
			// ultraGrid2
			// 
			this.ultraGrid2.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.ultraGrid2.Cursor = System.Windows.Forms.Cursors.Default;
			this.ultraGrid2.DataMember = "ContratosFinanciacion";
			this.ultraGrid2.DataSource = this.cdsFinancia;
			ultraGridBand1.AddButtonCaption = "Pago";
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn2.Header.Caption = "Servicio";
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridColumn2.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownValidate;
			ultraGridColumn2.Width = 99;
			ultraGridColumn3.Header.VisiblePosition = 2;
			ultraGridColumn4.Header.VisiblePosition = 3;
			ultraGridColumn5.Header.Caption = "ValorTotal";
			ultraGridColumn5.Header.VisiblePosition = 5;
			ultraGridColumn5.Width = 115;
			ultraGridColumn6.Header.Caption = "%";
			ultraGridColumn6.Header.VisiblePosition = 6;
			ultraGridColumn6.Width = 60;
			ultraGridColumn7.Header.Caption = "Fecha Inicio";
			ultraGridColumn7.Header.VisiblePosition = 7;
			ultraGridColumn7.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownCalendar;
			ultraGridColumn8.Header.Caption = "Fecha Fin";
			ultraGridColumn8.Header.VisiblePosition = 8;
			ultraGridColumn8.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownCalendar;
			ultraGridColumn9.Header.VisiblePosition = 9;
			ultraGridColumn9.Hidden = true;
			ultraGridColumn10.Header.VisiblePosition = 10;
			ultraGridColumn10.Hidden = true;
			ultraGridColumn11.Header.Caption = "Valor Unitario";
			ultraGridColumn11.Header.VisiblePosition = 4;
			ultraGridColumn11.Width = 105;
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
																										 ultraGridColumn11});
			this.ultraGrid2.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.ultraGrid2.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.TemplateOnBottom;
			this.ultraGrid2.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.True;
			this.ultraGrid2.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance1.BackColor = System.Drawing.Color.Transparent;
			this.ultraGrid2.DisplayLayout.Override.CardAreaAppearance = appearance1;
			appearance2.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance2.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance2.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance2.FontData.BoldAsString = "True";
			appearance2.FontData.Name = "Arial";
			appearance2.FontData.SizeInPoints = 10F;
			appearance2.ForeColor = System.Drawing.Color.White;
			appearance2.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.ultraGrid2.DisplayLayout.Override.HeaderAppearance = appearance2;
			appearance3.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance3.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance3.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid2.DisplayLayout.Override.RowSelectorAppearance = appearance3;
			appearance4.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance4.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance4.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid2.DisplayLayout.Override.SelectedRowAppearance = appearance4;
			this.ultraGrid2.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.ultraGrid2.Location = new System.Drawing.Point(8, 40);
			this.ultraGrid2.Name = "ultraGrid2";
			this.ultraGrid2.Size = new System.Drawing.Size(720, 136);
			this.ultraGrid2.TabIndex = 126;
			this.ultraGrid2.AfterRowInsert += new Infragistics.Win.UltraWinGrid.RowEventHandler(this.ultraGrid2_AfterRowInsert);
			this.ultraGrid2.BeforeRowsDeleted += new Infragistics.Win.UltraWinGrid.BeforeRowsDeletedEventHandler(this.ultraGrid2_BeforeRowsDeleted);
			this.ultraGrid2.AfterCellUpdate += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.ultraGrid2_AfterCellUpdate);
			this.ultraGrid2.AfterCellActivate += new System.EventHandler(this.ultraGrid2_AfterCellActivate);
			// 
			// cdsFinancia
			// 
			this.cdsFinancia.BindingContextCtrl = this;
			this.cdsFinancia.DataLibrary = "LibFacturacion";
			this.cdsFinancia.DataLibraryUrl = "";
			this.cdsFinancia.DataSetDef = "dsContrato";
			this.cdsFinancia.EnforceConstraints = false;
			this.cdsFinancia.Locale = new System.Globalization.CultureInfo("es-MX");
			this.cdsFinancia.Name = "cdsFinancia";
			this.cdsFinancia.SchemaClassName = "LibFacturacion.DataClass";
			this.cdsFinancia.SchemaDef = null;
			// 
			// btnTabla
			// 
			this.btnTabla.Anchor = System.Windows.Forms.AnchorStyles.Right;
			this.btnTabla.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.btnTabla.Location = new System.Drawing.Point(632, 184);
			this.btnTabla.Name = "btnTabla";
			this.btnTabla.Size = new System.Drawing.Size(88, 24);
			this.btnTabla.TabIndex = 127;
			this.btnTabla.Text = "Generar &Tabla";
			this.btnTabla.Click += new System.EventHandler(this.btnTabla_Click);
			// 
			// ultraGrid1
			// 
			this.ultraGrid1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.ultraGrid1.Cursor = System.Windows.Forms.Cursors.Default;
			appearance5.BackColor = System.Drawing.Color.White;
			this.ultraGrid1.DisplayLayout.Appearance = appearance5;
			this.ultraGrid1.DisplayLayout.AutoFitColumns = true;
			this.ultraGrid1.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
			this.ultraGrid1.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.True;
			this.ultraGrid1.DisplayLayout.Override.AllowRowFiltering = Infragistics.Win.DefaultableBoolean.True;
			this.ultraGrid1.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance6.BackColor = System.Drawing.Color.Transparent;
			this.ultraGrid1.DisplayLayout.Override.CardAreaAppearance = appearance6;
			appearance7.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance7.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance7.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance7.FontData.BoldAsString = "True";
			appearance7.FontData.Name = "Arial";
			appearance7.FontData.SizeInPoints = 10F;
			appearance7.ForeColor = System.Drawing.Color.White;
			appearance7.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.ultraGrid1.DisplayLayout.Override.HeaderAppearance = appearance7;
			this.ultraGrid1.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti;
			this.ultraGrid1.DisplayLayout.Override.RowFilterAction = Infragistics.Win.UltraWinGrid.RowFilterAction.HideFilteredOutRows;
			this.ultraGrid1.DisplayLayout.Override.RowFilterMode = Infragistics.Win.UltraWinGrid.RowFilterMode.AllRowsInBand;
			appearance8.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance8.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance8.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.RowSelectorAppearance = appearance8;
			appearance9.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance9.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance9.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.SelectedRowAppearance = appearance9;
			appearance10.BackColorAlpha = Infragistics.Win.Alpha.UseAlphaLevel;
			this.ultraGrid1.DisplayLayout.Override.SummaryValueAppearance = appearance10;
			this.ultraGrid1.DisplayLayout.ViewStyleBand = Infragistics.Win.UltraWinGrid.ViewStyleBand.OutlookGroupBy;
			this.ultraGrid1.Location = new System.Drawing.Point(8, 216);
			this.ultraGrid1.Name = "ultraGrid1";
			this.ultraGrid1.Size = new System.Drawing.Size(720, 184);
			this.ultraGrid1.TabIndex = 128;
			this.ultraGrid1.BeforeRowsDeleted += new Infragistics.Win.UltraWinGrid.BeforeRowsDeletedEventHandler(this.ultraGrid1_BeforeRowsDeleted);
			// 
			// btnFactura
			// 
			this.btnFactura.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnFactura.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.btnFactura.Location = new System.Drawing.Point(16, 408);
			this.btnFactura.Name = "btnFactura";
			this.btnFactura.Size = new System.Drawing.Size(96, 24);
			this.btnFactura.TabIndex = 129;
			this.btnFactura.Text = "Borra Facturado";
			this.btnFactura.Click += new System.EventHandler(this.btnFactura_Click);
			// 
			// btSalir
			// 
			this.btSalir.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.btSalir.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.btSalir.Location = new System.Drawing.Point(640, 408);
			this.btSalir.Name = "btSalir";
			this.btSalir.Size = new System.Drawing.Size(88, 24);
			this.btSalir.TabIndex = 130;
			this.btSalir.Text = "&Salir";
			this.btSalir.Click += new System.EventHandler(this.btSalir_Click);
			// 
			// cmbServicio
			// 
			this.cmbServicio.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbServicio.DataSource = this.cdvServicio;
			ultraGridColumn12.Header.VisiblePosition = 0;
			ultraGridColumn12.Hidden = true;
			ultraGridColumn13.Header.VisiblePosition = 1;
			ultraGridColumn14.Header.VisiblePosition = 2;
			ultraGridColumn14.Hidden = true;
			ultraGridColumn15.Header.VisiblePosition = 3;
			ultraGridColumn15.Hidden = true;
			ultraGridColumn16.Header.VisiblePosition = 4;
			ultraGridColumn16.Hidden = true;
			ultraGridBand2.Columns.AddRange(new object[] {
																										 ultraGridColumn12,
																										 ultraGridColumn13,
																										 ultraGridColumn14,
																										 ultraGridColumn15,
																										 ultraGridColumn16});
			this.cmbServicio.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			this.cmbServicio.DisplayMember = "Nombre";
			this.cmbServicio.Location = new System.Drawing.Point(512, 64);
			this.cmbServicio.Name = "cmbServicio";
			this.cmbServicio.Size = new System.Drawing.Size(120, 96);
			this.cmbServicio.TabIndex = 131;
			this.cmbServicio.ValueMember = "idServicios";
			this.cmbServicio.Visible = false;
			// 
			// cdvServicio
			// 
			this.cdvServicio.BindingContextCtrl = this;
			this.cdvServicio.DataSet = this.cdsServicio;
			this.cdvServicio.TableName = "";
			this.cdvServicio.TableViewName = "ContratosServicios";
			// 
			// cdsServicio
			// 
			this.cdsServicio.BindingContextCtrl = this;
			this.cdsServicio.DataLibrary = "LibFacturacion";
			this.cdsServicio.DataLibraryUrl = "";
			this.cdsServicio.DataSetDef = "dsContrato";
			this.cdsServicio.Locale = new System.Globalization.CultureInfo("es-MX");
			this.cdsServicio.Name = "cdsServicio";
			this.cdsServicio.SchemaClassName = "LibFacturacion.DataClass";
			this.cdsServicio.SchemaDef = null;
			// 
			// cdvFinancia
			// 
			this.cdvFinancia.BindingContextCtrl = this;
			this.cdvFinancia.DataSet = this.cdsFinancia;
			this.cdvFinancia.Sort = "idFinanciacionS";
			this.cdvFinancia.TableName = "";
			this.cdvFinancia.TableViewName = "ContratosFinanciacion";
			// 
			// errorProvider
			// 
			this.errorProvider.ContainerControl = this;
			// 
			// cdsContratos
			// 
			this.cdsContratos.BindingContextCtrl = this;
			this.cdsContratos.DataLibrary = "LibFacturacion";
			this.cdsContratos.DataLibraryUrl = "";
			this.cdsContratos.DataSetDef = "dsContrato";
			this.cdsContratos.EnforceConstraints = false;
			this.cdsContratos.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsContratos.Name = "cdsContratos";
			this.cdsContratos.SchemaClassName = "LibFacturacion.DataClass";
			this.cdsContratos.SchemaDef = null;
			// 
			// cdvContratos
			// 
			this.cdvContratos.BindingContextCtrl = this;
			this.cdvContratos.DataSet = this.cdsContratos;
			this.cdvContratos.Sort = "idContratos";
			this.cdvContratos.TableName = "";
			this.cdvContratos.TableViewName = "Contratos";
			// 
			// btGuardar
			// 
			this.btGuardar.Anchor = System.Windows.Forms.AnchorStyles.Right;
			this.btGuardar.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.btGuardar.Location = new System.Drawing.Point(528, 184);
			this.btGuardar.Name = "btGuardar";
			this.btGuardar.Size = new System.Drawing.Size(88, 24);
			this.btGuardar.TabIndex = 132;
			this.btGuardar.Text = "&Guardar";
			this.btGuardar.Click += new System.EventHandler(this.btGuardar_Click);
			// 
			// IngresoContratosFinanciacion
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(744, 445);
			this.Controls.Add(this.btGuardar);
			this.Controls.Add(this.cmbServicio);
			this.Controls.Add(this.btSalir);
			this.Controls.Add(this.btnFactura);
			this.Controls.Add(this.ultraGrid1);
			this.Controls.Add(this.btnTabla);
			this.Controls.Add(this.ultraGrid2);
			this.Controls.Add(this.btnServicios);
			this.Name = "IngresoContratosFinanciacion";
			this.Text = "Ingreso Contratos Financiación";
			this.Load += new System.EventHandler(this.IngresoContratosFinanciacion_Load);
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid2)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsFinancia)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbServicio)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvServicio)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsServicio)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvFinancia)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsContratos)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvContratos)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void IngresoContratosFinanciacion_Load(object sender, System.EventArgs e)
		{
			ModificaGrilla(true);
			try
			{
				////grilla de ContratoFinanciacion
//				drContrato = cdsContratos.TableViews["Contratos"].Rows[0];			
				C1.Data.FilterCondition fcFinanciacion = new 
					C1.Data.FilterCondition(cdsFinancia, "ContratosFinanciacion", "idContrato = " + vIdContrato);
				/////grilla Amortizacion
				C1.Data.FilterConditions fcTotal = new C1.Data.FilterConditions();
				fcTotal.Add(fcFinanciacion);
				cdsFinancia.Fill(fcTotal, false);
			}
			catch{};
		}
		private void ModificaGrilla(bool bModifica)
		{
			ultraGrid2.Enabled = true;
			if (bModifica)
			{
				ultraGrid2.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
				ultraGrid2.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.True;
			}
			else
			{
				ultraGrid2.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.False;
				ultraGrid2.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			}
		}
		private void btnTabla_Click(object sender, System.EventArgs e)
		{
//			if (vIdContrato <= 0)
//			{
//				txtTriggerC.Value = 0;
//			}

			if(ultraGrid2.Rows.Count < 1 )return;
			Cursor = Cursors.WaitCursor;
			try
			{						
				int Id = (int) ultraGrid2.ActiveRow.Cells["idFinanciacionS"].Value;	
				int idSericio = (int) ultraGrid2.ActiveRow.Cells["idServicio"].Value;
				//				string vTriger = txtTriggerC.Value.ToString();
				if(Id < 0)return;
				generaAmortizacion(Id,1,idSericio,"0");
			}
			catch(Exception ex){MessageBox.Show("Error: "+ex);}
			Cursor = Cursors.Default;

		}

		private void btSalir_Click(object sender, System.EventArgs e)
		{			
			this.DialogResult = DialogResult.OK;
			this.Close();
		}

		private void btnServicios_Click(object sender, System.EventArgs e)
		{
			using (IngresoContratosServicios ingCS = new IngresoContratosServicios())
			{
				ingCS.ShowDialog();
			}
			this.cdsServicio.Clear();
			this.cdsServicio.Fill();	
		}
		DataSet miDataSetG = new DataSet();
		DataView dvG = new DataView();
		
	
		private int verificaDatosAmortiza(int idFinancia)
		{	
			int idCuenta = 0;	
			if(vIdContrato > 0 )
			{
				return idCuenta;
			}
			else
			{			
				string stCuenta = "select count(*)from ContratosAmortizacion where idFinanciacionS = "+ idFinancia;
				idCuenta = Funcion.iEscalarSQL(cdsFinancia, stCuenta, true);
				return idCuenta;
			}
		}
		private void btnFactura_Click(object sender, System.EventArgs e)
		{
			int idAmort = (int) ultraGrid1.ActiveRow.Cells["idAmortizacionC"].Value;
			int idfin = (int) ultraGrid1.ActiveRow.Cells["idFinanciacionS"].Value;
			string stUpdate = "update ContratosAmortizacion set facturado = 0 where idAmortizacionC = "+ idAmort;			
			Funcion.EjecutaSQL(cdsFinancia, stUpdate, true);			
			generaAmortizacion(idfin,2,0,"0");
		}

		private void ultraGrid2_AfterCellActivate(object sender, System.EventArgs e)
		{
			int Id = (int) ultraGrid2.ActiveRow.Cells["idFinanciacionS"].Value;
			if(Id < 0)return;
			

			generaAmortizacion(Id,2,0,"0");
		}

		private void ultraGrid2_AfterRowInsert(object sender, Infragistics.Win.UltraWinGrid.RowEventArgs e)
		{
			//drContrato = cdsContratos.TableViews["Contratos"].Rows[0];
			e.Row.Cells["idContrato"].Value = VidC;
			
		}
C1.Data.C1DataRow drContrato;
		private void ultraGrid2_BeforeRowsDeleted(object sender, Infragistics.Win.UltraWinGrid.BeforeRowsDeletedEventArgs e)
		{
			e.DisplayPromptMsg=false;
			if (DialogResult.No == MessageBox.Show("Desea Eliminar Registro", 
				"Información", MessageBoxButtons.YesNo, MessageBoxIcon.Question)) return;
			try
			{
				this.errorProvider.SetError(this.ultraGrid2, "");								
				
				int idFinac = (int) ultraGrid2.ActiveRow.Cells["idFinanciacionS"].Value;

				string stCuenta = "select count(*)from ContratosAmortizacion where idFinanciacionS = "+ idFinac;
				int numid = Funcion.iEscalarSQL(cdsFinancia, stCuenta, true);
				

				int id= verificaDatosAmortiza(idFinac);
				if(numid > 0)
				{					
					if (DialogResult.No == MessageBox.Show("Registro Tiene Amortizacion Generada \n Desea Eliminar Registro", 
						"Eliminar", MessageBoxButtons.YesNo, MessageBoxIcon.Stop))
					{
						e.Cancel = true;
							return;
					}
					
					string stSelectAmort = "delete ContratosAmortizacion where idFinanciacionS = "+ idFinac;			
					Funcion.EjecutaSQL(cdsFinancia, stSelectAmort, true);
				}
				int idFin = (int) ultraGrid2.ActiveRow.Cells["idFinanciacionS"].Value;
				string stSelect = "delete ContratosFinanciacion where idFinanciacionS = "+ idFin;			
				Funcion.EjecutaSQL(cdsFinancia, stSelect, true);
			}
			catch(Exception ex){MessageBox.Show("Error al borrar: "+ex);}
			return;
		}

		private void ultraGrid1_BeforeRowsDeleted(object sender, Infragistics.Win.UltraWinGrid.BeforeRowsDeletedEventArgs e)
		{
			e.DisplayPromptMsg=false;
			if (DialogResult.No == MessageBox.Show("Desea Eliminar Registro", 
				"Información", MessageBoxButtons.YesNo, MessageBoxIcon.Question)) return;
			borraAmortizacion();
		}
		private void borraAmortizacion()
		{
			int idAmort = (int) ultraGrid1.ActiveRow.Cells["idFinanciacionS"].Value;
			string stCuenta = "delete ContratosAmortizacion where idFinanciacionS = "+ idAmort;			
			Funcion.EjecutaSQL(cdsFinancia, stCuenta, true);
		}
		private void generaAmortizacion(int idFinancia, int controlP, int idServicio, string vTriger)
		{			
			int verif = verificaDatosAmortiza(idFinancia);			
			if (verif > 0 && controlP==1 )return;
			//			if(controlP == 1)
			//			{
			//				if (DialogResult.No == MessageBox.Show("Desea Generar Tabla de Amortizacion", 
			//					"Información", MessageBoxButtons.YesNo, MessageBoxIcon.Question)) return;
			//			}
			
			Funcion miFuncion = new Funcion();
			string strConn = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);	
			string stProcedimiento = "exec ContratosAmortiza "+idFinancia+" ,"+controlP+" ,"+VidC+" ,"+idServicio+" ,"+vTriger;
			SqlDataAdapter myData = new SqlDataAdapter(stProcedimiento, strConn);
			myData.TableMappings.Add("Tabla", "ExcelSube");
			
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
			
			#region Visibilidad
			ultraGrid1.DisplayLayout.Bands[0].Columns["idAmortizacionC"].Hidden = true;
			ultraGrid1.DisplayLayout.Bands[0].Columns["idFinanciacionS"].Hidden = true;
			ultraGrid1.DisplayLayout.Bands[0].Columns["FechaCobro"].Hidden = true;
			ultraGrid1.DisplayLayout.Bands[0].Columns["CapCobrado"].Hidden = true;
			ultraGrid1.DisplayLayout.Bands[0].Columns["Facturado"].CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
			

			#endregion Visibilidad
		}

		private void btGuardar_Click(object sender, System.EventArgs e)
		{
			grabaGrilla2();

		}
		private void grabaGrilla2()
		{
//			int verif = 0;			
//			int Id =0;					
			if(ultraGrid2.Rows.Count<1)return;			
			try
			{
				int cuotas = 0;
				Cursor = Cursors.WaitCursor;
				if(ultraGrid2.ActiveRow.Cells["Cuotas"].Value != System.DBNull.Value)
				{
					cuotas = (int) ultraGrid2.ActiveRow.Cells["Cuotas"].Value;
				}			
				if(ultraGrid2.ActiveRow.Cells["FechaD"].Value == System.DBNull.Value)
				{
					ultraGrid2.ActiveRow.Cells["FechaD"].Value = DateTime.Today;				
				}
				
				if ((ultraGrid2.ActiveRow.Cells["FechaH"].Value == System.DBNull.Value))
				{
					ultraGrid2.ActiveRow.Cells["FechaH"].Value = DateTime.Today.AddMonths(cuotas-1);
					ultraGrid2.ActiveRow.Cells["Cuotas"].Value = cuotas;
				}
				else
				{
					int ret1 = calculaFecha((DateTime)ultraGrid2.ActiveRow.Cells["FechaD"].Value, (DateTime)ultraGrid2.ActiveRow.Cells["FechaH"].Value, 2);
					if(ret1 > 0)
					{
						ultraGrid2.ActiveRow.Cells["Cuotas"].Value = ret1;
					}				
				}
								
				//ultraGrid2.ActiveRow.Cells["VTrigger"].Value = verif;				
			}
			catch (Exception ex)
			{	MessageBox.Show("Error "+ex);	}		
			//			ModificaGrilla(false);
			//			if(vIdContrato > 0) 
			//			{
			//				generaAmortizacion(Id,1);
			//			}	
			this.cdsFinancia.Update();	
			
			Cursor = Cursors.Default;
		}
		private int calculaFecha(DateTime desde, DateTime hasta, int Control)
		{
			int[] monthDay = new int[12] { 31, -1, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31 };
			int year;
			int month;
			int day;
			DateTime dtD; DateTime dtH;
			dtD = desde;
			dtH = hasta;
           
			int increment;

			if (dtD > dtH)
			{
				MessageBox.Show("Fecha Hasta debe ser Mayor", "Fechas", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
				return 0;
			}

			/// 
			/// Day Calculation
			/// 
			increment = 0;

			if (dtD.Day > dtH.Day)
			{
				increment = monthDay[dtD.Month - 1];

			}
			if (increment == -1)
			{
				if (DateTime.IsLeapYear(dtD.Year))
				{
					// leap year february contain 29 days
					increment = 29;
				}
				else
				{
					increment = 28;
				}
			}
			if (increment != 0)
			{
				day = (dtH.Day + increment) - dtD.Day;
				increment = 1;
			}
			else
			{
				day = dtH.Day - dtD.Day;
			}

			///
			///month calculation
			///
			if ((dtD.Month + increment) > dtH.Month)
			{
				month = (dtH.Month + 12) - (dtD.Month + increment);
				increment = 1;
			}
			else
			{
				month = (dtH.Month) - (dtD.Month + increment);
				increment = 0;
			}

			///
			/// year calculation
			///
			year = dtH.Year - (dtD.Year + increment);
			//MessageBox.Show(" anio " + year + " mes " + month + " dia " + day);
			int retorna = 0;
			if(Control == 1)
			{
				//				txtAnio.Value =  year;
				//				txtMes.Value = month;			
				drContrato = cdsContratos.TableViews["Contratos"].Rows[0];				
				drContrato["anio"]=year;
				drContrato["mes"]=month;
		
			}
			else
			{
				retorna = ((year*12)+month);
			}
			return retorna; 
		}

		private void ultraGrid2_AfterCellUpdate(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			
			if (e.Cell.Column.ToString() == "valorU" || e.Cell.Column.ToString() == "Cuotas")
			{

				if (e.Cell.Row.Cells["valorU"].Value == null ||e.Cell.Row.Cells["valorU"].Value.ToString() == "" )
				{
					e.Cell.Row.Cells["valorU"].Value = 0;					
					return;
				}
				if (e.Cell.Row.Cells["Cuotas"].Value == null ||e.Cell.Row.Cells["Cuotas"].Value.ToString() == "" )
				{
					e.Cell.Row.Cells["Cuotas"].Value = 0;					
					return;
				}
				float v = float.Parse(e.Cell.Row.Cells["valorU"].Value.ToString());
				int c = int.Parse(e.Cell.Row.Cells["Cuotas"].Value.ToString());

				e.Cell.Row.Cells["Valor"].Value = v*c;
			}

			

		}	
	}
}
