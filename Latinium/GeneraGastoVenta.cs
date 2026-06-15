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
	/// Descripción breve de GeneraGastoVenta.
	/// </summary>
	public class GeneraGastoVenta : DevExpress.XtraEditors.XtraForm
	{
		private Infragistics.Win.Misc.UltraButton btAceptar;
		private Infragistics.Win.UltraWinGrid.UltraGrid ultraGrid2;
		private Infragistics.Win.UltraWinGrid.UltraGrid ultraGrid1;
		private C1.Data.C1DataSet cdsContratos;
		private System.Windows.Forms.ImageList imageList;
		private System.Windows.Forms.ErrorProvider errorProvider;
		private System.Windows.Forms.ErrorProvider errorProvider1;
		private Infragistics.Win.Misc.UltraButton ultraButton1;
		private Infragistics.Win.Misc.UltraButton ultraButton2;
		private Infragistics.Win.Misc.UltraButton btAgregarTodo;
		private Infragistics.Win.Misc.UltraButton btAgregar;
		private Infragistics.Win.Misc.UltraButton btFiltro;
		private System.Windows.Forms.Label label1;
		private C1.Data.C1DataSet cdsCompra;
		private System.Windows.Forms.Label label5;
		private System.Windows.Forms.Label label6;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo cmbFechaH;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo cmbFechaD;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkCliente;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkCentroC;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkProyecto;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbCliente;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbCentroC;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbProyecto;
		private C1.Data.C1DataView cdvProyecto;
		private C1.Data.C1DataSet cdsProyecto;
		private C1.Data.C1DataSet cdsCliente;
		private C1.Data.C1DataView cdvCliente;
		private C1.Data.C1DataView cdvSubProyecto;
		private System.ComponentModel.IContainer components;

		public GeneraGastoVenta()
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
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton1 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(GeneraGastoVenta));
			Infragistics.Win.Appearance appearance11 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance12 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance13 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance14 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance15 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton2 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.Appearance appearance16 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Cliente", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCliente");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Proveedor");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Ruc");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Ambos");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("DesconProv");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("DesconCliente");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Entrega");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idBanco");
			Infragistics.Win.Appearance appearance17 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Proyecto", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn11 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idProyecto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn12 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPadre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn13 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn14 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Lugar");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn15 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Responsable");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn16 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Notas");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn17 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CentroCosto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn18 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaIngreso");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn19 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Visible");
			Infragistics.Win.Appearance appearance18 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand3 = new Infragistics.Win.UltraWinGrid.UltraGridBand("SubProyecto", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn20 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idSubProyecto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn21 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idProyecto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn22 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn23 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Lugar");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn24 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("RespNombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn25 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Notas");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn26 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CodSubproyecto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn27 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Visible");
			this.btAceptar = new Infragistics.Win.Misc.UltraButton();
			this.ultraGrid2 = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraGrid1 = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.cmbFechaH = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.cdsContratos = new C1.Data.C1DataSet();
			this.imageList = new System.Windows.Forms.ImageList(this.components);
			this.errorProvider = new System.Windows.Forms.ErrorProvider();
			this.errorProvider1 = new System.Windows.Forms.ErrorProvider();
			this.ultraButton1 = new Infragistics.Win.Misc.UltraButton();
			this.ultraButton2 = new Infragistics.Win.Misc.UltraButton();
			this.btAgregarTodo = new Infragistics.Win.Misc.UltraButton();
			this.btAgregar = new Infragistics.Win.Misc.UltraButton();
			this.btFiltro = new Infragistics.Win.Misc.UltraButton();
			this.label1 = new System.Windows.Forms.Label();
			this.cdsCompra = new C1.Data.C1DataSet();
			this.cmbFechaD = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.label5 = new System.Windows.Forms.Label();
			this.label6 = new System.Windows.Forms.Label();
			this.chkCliente = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.chkCentroC = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.chkProyecto = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.cmbCliente = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.cdvCliente = new C1.Data.C1DataView();
			this.cdsCliente = new C1.Data.C1DataSet();
			this.cmbCentroC = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.cdvProyecto = new C1.Data.C1DataView();
			this.cdsProyecto = new C1.Data.C1DataSet();
			this.cmbProyecto = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.cdvSubProyecto = new C1.Data.C1DataView();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid2)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbFechaH)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsContratos)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsCompra)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbFechaD)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCliente)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvCliente)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsCliente)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCentroC)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvProyecto)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsProyecto)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbProyecto)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvSubProyecto)).BeginInit();
			this.SuspendLayout();
			// 
			// btAceptar
			// 
			this.btAceptar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btAceptar.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.btAceptar.Location = new System.Drawing.Point(752, 24);
			this.btAceptar.Name = "btAceptar";
			this.btAceptar.Size = new System.Drawing.Size(75, 32);
			this.btAceptar.TabIndex = 20;
			this.btAceptar.Text = "Convierte a Factura";
			this.btAceptar.Click += new System.EventHandler(this.btAceptar_Click);
			// 
			// ultraGrid2
			// 
			this.ultraGrid2.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.ultraGrid2.Cursor = System.Windows.Forms.Cursors.Default;
			appearance1.BackColor = System.Drawing.Color.White;
			this.ultraGrid2.DisplayLayout.Appearance = appearance1;
			this.ultraGrid2.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
			this.ultraGrid2.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.ultraGrid2.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.False;
			appearance2.BackColor = System.Drawing.Color.Transparent;
			this.ultraGrid2.DisplayLayout.Override.CardAreaAppearance = appearance2;
			this.ultraGrid2.DisplayLayout.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.RowSelect;
			this.ultraGrid2.DisplayLayout.Override.ColumnAutoSizeMode = Infragistics.Win.UltraWinGrid.ColumnAutoSizeMode.AllRowsInBand;
			appearance3.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance3.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance3.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance3.FontData.BoldAsString = "True";
			appearance3.FontData.Name = "Arial";
			appearance3.FontData.SizeInPoints = 10F;
			appearance3.ForeColor = System.Drawing.Color.White;
			appearance3.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.ultraGrid2.DisplayLayout.Override.HeaderAppearance = appearance3;
			this.ultraGrid2.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti;
			appearance4.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance4.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance4.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid2.DisplayLayout.Override.RowSelectorAppearance = appearance4;
			appearance5.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance5.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance5.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid2.DisplayLayout.Override.SelectedRowAppearance = appearance5;
			this.ultraGrid2.DisplayLayout.Override.SelectTypeCell = Infragistics.Win.UltraWinGrid.SelectType.None;
			this.ultraGrid2.DisplayLayout.Override.SelectTypeCol = Infragistics.Win.UltraWinGrid.SelectType.None;
			this.ultraGrid2.DisplayLayout.Override.SelectTypeRow = Infragistics.Win.UltraWinGrid.SelectType.Single;
			this.ultraGrid2.DisplayLayout.ScrollBounds = Infragistics.Win.UltraWinGrid.ScrollBounds.ScrollToFill;
			this.ultraGrid2.DisplayLayout.ScrollStyle = Infragistics.Win.UltraWinGrid.ScrollStyle.Immediate;
			this.ultraGrid2.DisplayLayout.TabNavigation = Infragistics.Win.UltraWinGrid.TabNavigation.NextControl;
			this.ultraGrid2.DisplayLayout.ViewStyle = Infragistics.Win.UltraWinGrid.ViewStyle.SingleBand;
			this.ultraGrid2.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.ultraGrid2.Location = new System.Drawing.Point(464, 96);
			this.ultraGrid2.Name = "ultraGrid2";
			this.ultraGrid2.Size = new System.Drawing.Size(368, 240);
			this.ultraGrid2.TabIndex = 19;
			// 
			// ultraGrid1
			// 
			this.ultraGrid1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left)));
			this.ultraGrid1.Cursor = System.Windows.Forms.Cursors.Default;
			appearance6.BackColor = System.Drawing.Color.White;
			this.ultraGrid1.DisplayLayout.Appearance = appearance6;
			this.ultraGrid1.DisplayLayout.AutoFitColumns = true;
			this.ultraGrid1.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
			this.ultraGrid1.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.ultraGrid1.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.False;
			appearance7.BackColor = System.Drawing.Color.Transparent;
			this.ultraGrid1.DisplayLayout.Override.CardAreaAppearance = appearance7;
			this.ultraGrid1.DisplayLayout.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.RowSelect;
			this.ultraGrid1.DisplayLayout.Override.ColumnAutoSizeMode = Infragistics.Win.UltraWinGrid.ColumnAutoSizeMode.AllRowsInBand;
			appearance8.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance8.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance8.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance8.FontData.BoldAsString = "True";
			appearance8.FontData.Name = "Arial";
			appearance8.FontData.SizeInPoints = 10F;
			appearance8.ForeColor = System.Drawing.Color.White;
			appearance8.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.ultraGrid1.DisplayLayout.Override.HeaderAppearance = appearance8;
			this.ultraGrid1.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti;
			appearance9.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance9.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance9.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.RowSelectorAppearance = appearance9;
			appearance10.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance10.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance10.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.SelectedRowAppearance = appearance10;
			this.ultraGrid1.DisplayLayout.Override.SelectTypeCell = Infragistics.Win.UltraWinGrid.SelectType.None;
			this.ultraGrid1.DisplayLayout.Override.SelectTypeCol = Infragistics.Win.UltraWinGrid.SelectType.None;
			this.ultraGrid1.DisplayLayout.Override.SelectTypeRow = Infragistics.Win.UltraWinGrid.SelectType.Extended;
			this.ultraGrid1.DisplayLayout.ScrollBounds = Infragistics.Win.UltraWinGrid.ScrollBounds.ScrollToFill;
			this.ultraGrid1.DisplayLayout.ScrollStyle = Infragistics.Win.UltraWinGrid.ScrollStyle.Immediate;
			this.ultraGrid1.DisplayLayout.TabNavigation = Infragistics.Win.UltraWinGrid.TabNavigation.NextControl;
			this.ultraGrid1.DisplayLayout.ViewStyle = Infragistics.Win.UltraWinGrid.ViewStyle.SingleBand;
			this.ultraGrid1.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.ultraGrid1.Location = new System.Drawing.Point(16, 96);
			this.ultraGrid1.Name = "ultraGrid1";
			this.ultraGrid1.Size = new System.Drawing.Size(408, 240);
			this.ultraGrid1.TabIndex = 18;
			// 
			// cmbFechaH
			// 
			dateButton1.Caption = "Today";
			this.cmbFechaH.DateButtons.Add(dateButton1);
			this.cmbFechaH.Format = "dd/MM/yyyy";
			this.cmbFechaH.Location = new System.Drawing.Point(112, 32);
			this.cmbFechaH.Name = "cmbFechaH";
			this.cmbFechaH.NonAutoSizeHeight = 23;
			this.cmbFechaH.Size = new System.Drawing.Size(104, 21);
			this.cmbFechaH.SpinButtonsVisible = true;
			this.cmbFechaH.TabIndex = 17;
			this.cmbFechaH.Value = new System.DateTime(2011, 5, 4, 0, 0, 0, 0);
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
			// imageList
			// 
			this.imageList.ColorDepth = System.Windows.Forms.ColorDepth.Depth32Bit;
			this.imageList.ImageSize = new System.Drawing.Size(18, 18);
			this.imageList.ImageStream = ((System.Windows.Forms.ImageListStreamer)(resources.GetObject("imageList.ImageStream")));
			this.imageList.TransparentColor = System.Drawing.Color.Transparent;
			// 
			// errorProvider
			// 
			this.errorProvider.ContainerControl = this;
			// 
			// errorProvider1
			// 
			this.errorProvider1.ContainerControl = this;
			// 
			// ultraButton1
			// 
			this.ultraButton1.Anchor = System.Windows.Forms.AnchorStyles.Left;
			appearance11.Image = 1;
			this.ultraButton1.Appearance = appearance11;
			this.ultraButton1.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.ultraButton1.ImageList = this.imageList;
			this.ultraButton1.Location = new System.Drawing.Point(432, 220);
			this.ultraButton1.Name = "ultraButton1";
			this.ultraButton1.Size = new System.Drawing.Size(24, 23);
			this.ultraButton1.TabIndex = 25;
			this.ultraButton1.Click += new System.EventHandler(this.ultraButton1_Click);
			// 
			// ultraButton2
			// 
			this.ultraButton2.Anchor = System.Windows.Forms.AnchorStyles.Left;
			appearance12.Image = 0;
			this.ultraButton2.Appearance = appearance12;
			this.ultraButton2.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.ultraButton2.ImageList = this.imageList;
			this.ultraButton2.Location = new System.Drawing.Point(432, 252);
			this.ultraButton2.Name = "ultraButton2";
			this.ultraButton2.Size = new System.Drawing.Size(24, 23);
			this.ultraButton2.TabIndex = 24;
			this.ultraButton2.Click += new System.EventHandler(this.ultraButton2_Click);
			// 
			// btAgregarTodo
			// 
			this.btAgregarTodo.Anchor = System.Windows.Forms.AnchorStyles.Left;
			appearance13.Image = 3;
			this.btAgregarTodo.Appearance = appearance13;
			this.btAgregarTodo.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.btAgregarTodo.ImageList = this.imageList;
			this.btAgregarTodo.Location = new System.Drawing.Point(432, 148);
			this.btAgregarTodo.Name = "btAgregarTodo";
			this.btAgregarTodo.Size = new System.Drawing.Size(24, 23);
			this.btAgregarTodo.TabIndex = 23;
			this.btAgregarTodo.Click += new System.EventHandler(this.btAgregarTodo_Click);
			// 
			// btAgregar
			// 
			this.btAgregar.Anchor = System.Windows.Forms.AnchorStyles.Left;
			appearance14.Image = 2;
			this.btAgregar.Appearance = appearance14;
			this.btAgregar.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.btAgregar.ImageList = this.imageList;
			this.btAgregar.Location = new System.Drawing.Point(432, 180);
			this.btAgregar.Name = "btAgregar";
			this.btAgregar.Size = new System.Drawing.Size(24, 23);
			this.btAgregar.TabIndex = 22;
			this.btAgregar.Click += new System.EventHandler(this.btAgregar_Click);
			// 
			// btFiltro
			// 
			appearance15.Image = 5;
			this.btFiltro.Appearance = appearance15;
			this.btFiltro.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.btFiltro.ImageList = this.imageList;
			this.btFiltro.Location = new System.Drawing.Point(24, 64);
			this.btFiltro.Name = "btFiltro";
			this.btFiltro.Size = new System.Drawing.Size(24, 24);
			this.btFiltro.TabIndex = 21;
			this.btFiltro.Click += new System.EventHandler(this.btFiltro_Click);
			// 
			// label1
			// 
			this.label1.Location = new System.Drawing.Point(48, 72);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(152, 16);
			this.label1.TabIndex = 26;
			this.label1.Text = "Buscar Reembolsos Gastos";
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
			// cmbFechaD
			// 
			dateButton2.Caption = "Today";
			this.cmbFechaD.DateButtons.Add(dateButton2);
			this.cmbFechaD.Format = "dd/MM/yyyy";
			this.cmbFechaD.Location = new System.Drawing.Point(112, 8);
			this.cmbFechaD.Name = "cmbFechaD";
			this.cmbFechaD.NonAutoSizeHeight = 23;
			this.cmbFechaD.Size = new System.Drawing.Size(104, 21);
			this.cmbFechaD.SpinButtonsVisible = true;
			this.cmbFechaD.TabIndex = 33;
			this.cmbFechaD.Value = new System.DateTime(2011, 5, 4, 0, 0, 0, 0);
			// 
			// label5
			// 
			this.label5.AutoSize = true;
			this.label5.BackColor = System.Drawing.Color.Transparent;
			this.label5.Location = new System.Drawing.Point(16, 8);
			this.label5.Name = "label5";
			this.label5.Size = new System.Drawing.Size(71, 16);
			this.label5.TabIndex = 34;
			this.label5.Text = "Fecha Desde";
			// 
			// label6
			// 
			this.label6.AutoSize = true;
			this.label6.BackColor = System.Drawing.Color.Transparent;
			this.label6.Location = new System.Drawing.Point(16, 32);
			this.label6.Name = "label6";
			this.label6.Size = new System.Drawing.Size(68, 16);
			this.label6.TabIndex = 35;
			this.label6.Text = "Fecha Hasta";
			// 
			// chkCliente
			// 
			this.chkCliente.BackColor = System.Drawing.Color.Transparent;
			this.chkCliente.CheckAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.chkCliente.Checked = true;
			this.chkCliente.CheckState = System.Windows.Forms.CheckState.Checked;
			this.chkCliente.Location = new System.Drawing.Point(256, 8);
			this.chkCliente.Name = "chkCliente";
			this.chkCliente.Size = new System.Drawing.Size(104, 20);
			this.chkCliente.TabIndex = 36;
			this.chkCliente.Text = "Cliente";
			this.chkCliente.CheckedChanged += new System.EventHandler(this.chkCliente_CheckedChanged);
			// 
			// chkCentroC
			// 
			this.chkCentroC.BackColor = System.Drawing.Color.Transparent;
			this.chkCentroC.CheckAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.chkCentroC.Checked = true;
			this.chkCentroC.CheckState = System.Windows.Forms.CheckState.Checked;
			this.chkCentroC.Location = new System.Drawing.Point(256, 32);
			this.chkCentroC.Name = "chkCentroC";
			this.chkCentroC.Size = new System.Drawing.Size(104, 20);
			this.chkCentroC.TabIndex = 37;
			this.chkCentroC.Text = "Centro de Costo";
			this.chkCentroC.CheckedChanged += new System.EventHandler(this.chkCentroC_CheckedChanged);
			// 
			// chkProyecto
			// 
			this.chkProyecto.BackColor = System.Drawing.Color.Transparent;
			this.chkProyecto.CheckAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.chkProyecto.Checked = true;
			this.chkProyecto.CheckState = System.Windows.Forms.CheckState.Checked;
			this.chkProyecto.Location = new System.Drawing.Point(256, 56);
			this.chkProyecto.Name = "chkProyecto";
			this.chkProyecto.Size = new System.Drawing.Size(104, 20);
			this.chkProyecto.TabIndex = 38;
			this.chkProyecto.Text = "Proyecto";
			this.chkProyecto.CheckedChanged += new System.EventHandler(this.chkProyecto_CheckedChanged);
			// 
			// cmbCliente
			// 
			appearance16.BackColor = System.Drawing.Color.White;
			appearance16.BackColor2 = System.Drawing.Color.White;
			appearance16.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.cmbCliente.Appearance = appearance16;
			this.cmbCliente.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbCliente.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbCliente.DataSource = this.cdvCliente;
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn2.Header.VisiblePosition = 2;
			ultraGridColumn2.Hidden = true;
			ultraGridColumn3.Header.VisiblePosition = 9;
			ultraGridColumn3.Width = 192;
			ultraGridColumn4.Header.VisiblePosition = 1;
			ultraGridColumn4.Hidden = true;
			ultraGridColumn5.Header.VisiblePosition = 3;
			ultraGridColumn5.Hidden = true;
			ultraGridColumn6.Header.VisiblePosition = 4;
			ultraGridColumn6.Hidden = true;
			ultraGridColumn7.Header.VisiblePosition = 5;
			ultraGridColumn7.Hidden = true;
			ultraGridColumn8.Header.VisiblePosition = 6;
			ultraGridColumn8.Hidden = true;
			ultraGridColumn9.Header.VisiblePosition = 7;
			ultraGridColumn9.Hidden = true;
			ultraGridColumn10.Header.VisiblePosition = 8;
			ultraGridColumn10.Hidden = true;
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
			this.cmbCliente.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.cmbCliente.DisplayMember = "Nombre";
			this.cmbCliente.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbCliente.Enabled = false;
			this.cmbCliente.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbCliente.Location = new System.Drawing.Point(368, 8);
			this.cmbCliente.Name = "cmbCliente";
			this.cmbCliente.Size = new System.Drawing.Size(208, 21);
			this.cmbCliente.TabIndex = 39;
			this.cmbCliente.ValueMember = "idCliente";
			// 
			// cdvCliente
			// 
			this.cdvCliente.BindingContextCtrl = this;
			this.cdvCliente.DataSet = this.cdsCliente;
			this.cdvCliente.RowFilter = "Proveedor = 0";
			this.cdvCliente.Sort = "Nombre";
			this.cdvCliente.TableName = "";
			this.cdvCliente.TableViewName = "Cliente";
			// 
			// cdsCliente
			// 
			this.cdsCliente.BindingContextCtrl = this;
			this.cdsCliente.CaseSensitive = false;
			this.cdsCliente.DataLibrary = "LibFacturacion";
			this.cdsCliente.DataLibraryUrl = "";
			this.cdsCliente.DataSetDef = "dsClientes";
			this.cdsCliente.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsCliente.Name = "cdsCliente";
			this.cdsCliente.SchemaClassName = "LibFacturacion.DataClass";
			this.cdsCliente.SchemaDef = null;
			this.cdsCliente.BeforeFill += new C1.Data.FillEventHandler(this.cdsCliente_BeforeFill);
			// 
			// cmbCentroC
			// 
			appearance17.BackColor = System.Drawing.Color.White;
			appearance17.BackColor2 = System.Drawing.Color.White;
			appearance17.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.cmbCentroC.Appearance = appearance17;
			this.cmbCentroC.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbCentroC.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbCentroC.DataSource = this.cdvProyecto;
			ultraGridColumn11.Header.VisiblePosition = 0;
			ultraGridColumn11.Hidden = true;
			ultraGridColumn12.Header.VisiblePosition = 1;
			ultraGridColumn12.Hidden = true;
			ultraGridColumn13.Header.VisiblePosition = 2;
			ultraGridColumn14.Header.VisiblePosition = 3;
			ultraGridColumn14.Hidden = true;
			ultraGridColumn15.Header.VisiblePosition = 4;
			ultraGridColumn15.Hidden = true;
			ultraGridColumn16.Header.VisiblePosition = 5;
			ultraGridColumn16.Hidden = true;
			ultraGridColumn17.Header.VisiblePosition = 6;
			ultraGridColumn18.Header.VisiblePosition = 7;
			ultraGridColumn18.Hidden = true;
			ultraGridColumn19.Header.VisiblePosition = 8;
			ultraGridColumn19.Hidden = true;
			ultraGridBand2.Columns.AddRange(new object[] {
																										 ultraGridColumn11,
																										 ultraGridColumn12,
																										 ultraGridColumn13,
																										 ultraGridColumn14,
																										 ultraGridColumn15,
																										 ultraGridColumn16,
																										 ultraGridColumn17,
																										 ultraGridColumn18,
																										 ultraGridColumn19});
			this.cmbCentroC.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			this.cmbCentroC.DisplayMember = "Nombre";
			this.cmbCentroC.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbCentroC.Enabled = false;
			this.cmbCentroC.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbCentroC.Location = new System.Drawing.Point(368, 32);
			this.cmbCentroC.Name = "cmbCentroC";
			this.cmbCentroC.Size = new System.Drawing.Size(208, 21);
			this.cmbCentroC.TabIndex = 40;
			this.cmbCentroC.ValueMember = "idProyecto";
			// 
			// cdvProyecto
			// 
			this.cdvProyecto.BindingContextCtrl = this;
			this.cdvProyecto.DataSet = this.cdsProyecto;
			this.cdvProyecto.RowFilter = "Visible =1";
			this.cdvProyecto.Sort = "Nombre";
			this.cdvProyecto.TableName = "";
			this.cdvProyecto.TableViewName = "Proyecto";
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
			// cmbProyecto
			// 
			appearance18.BackColor = System.Drawing.Color.White;
			appearance18.BackColor2 = System.Drawing.Color.White;
			appearance18.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.cmbProyecto.Appearance = appearance18;
			this.cmbProyecto.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbProyecto.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbProyecto.DataSource = this.cdvSubProyecto;
			ultraGridColumn20.Header.VisiblePosition = 0;
			ultraGridColumn20.Hidden = true;
			ultraGridColumn21.Header.VisiblePosition = 1;
			ultraGridColumn21.Hidden = true;
			ultraGridColumn22.Header.VisiblePosition = 2;
			ultraGridColumn23.Header.VisiblePosition = 3;
			ultraGridColumn23.Hidden = true;
			ultraGridColumn24.Header.VisiblePosition = 4;
			ultraGridColumn24.Hidden = true;
			ultraGridColumn25.Header.VisiblePosition = 5;
			ultraGridColumn25.Hidden = true;
			ultraGridColumn26.Header.VisiblePosition = 6;
			ultraGridColumn27.Header.VisiblePosition = 7;
			ultraGridColumn27.Hidden = true;
			ultraGridBand3.Columns.AddRange(new object[] {
																										 ultraGridColumn20,
																										 ultraGridColumn21,
																										 ultraGridColumn22,
																										 ultraGridColumn23,
																										 ultraGridColumn24,
																										 ultraGridColumn25,
																										 ultraGridColumn26,
																										 ultraGridColumn27});
			this.cmbProyecto.DisplayLayout.BandsSerializer.Add(ultraGridBand3);
			this.cmbProyecto.DisplayMember = "Nombre";
			this.cmbProyecto.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbProyecto.Enabled = false;
			this.cmbProyecto.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbProyecto.Location = new System.Drawing.Point(368, 56);
			this.cmbProyecto.Name = "cmbProyecto";
			this.cmbProyecto.Size = new System.Drawing.Size(208, 21);
			this.cmbProyecto.TabIndex = 41;
			this.cmbProyecto.ValueMember = "idSubProyecto";
			// 
			// cdvSubProyecto
			// 
			this.cdvSubProyecto.BindingContextCtrl = this;
			this.cdvSubProyecto.DataSet = this.cdsProyecto;
			this.cdvSubProyecto.RowFilter = "Visible = 1";
			this.cdvSubProyecto.Sort = "Nombre";
			this.cdvSubProyecto.TableName = "";
			this.cdvSubProyecto.TableViewName = "SubProyecto";
			// 
			// GeneraGastoVenta
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(840, 341);
			this.Controls.Add(this.cmbProyecto);
			this.Controls.Add(this.cmbCentroC);
			this.Controls.Add(this.cmbCliente);
			this.Controls.Add(this.chkProyecto);
			this.Controls.Add(this.chkCentroC);
			this.Controls.Add(this.chkCliente);
			this.Controls.Add(this.label6);
			this.Controls.Add(this.label5);
			this.Controls.Add(this.cmbFechaD);
			this.Controls.Add(this.label1);
			this.Controls.Add(this.ultraButton1);
			this.Controls.Add(this.ultraButton2);
			this.Controls.Add(this.btAgregarTodo);
			this.Controls.Add(this.btAgregar);
			this.Controls.Add(this.btFiltro);
			this.Controls.Add(this.btAceptar);
			this.Controls.Add(this.ultraGrid2);
			this.Controls.Add(this.ultraGrid1);
			this.Controls.Add(this.cmbFechaH);
			this.Name = "GeneraGastoVenta";
			this.Text = "GeneraGastoVenta";
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid2)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbFechaH)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsContratos)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsCompra)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbFechaD)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCliente)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvCliente)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsCliente)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCentroC)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvProyecto)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsProyecto)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbProyecto)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvSubProyecto)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion
		DataSet miDataSetG = new DataSet();
		DataView dvG = new DataView();
		DataSet miDataSetN = new DataSet();
		DataView dvN = new DataView();
		private void btFiltro_Click(object sender, System.EventArgs e)
		{
			DateTime dtFechaD = (DateTime) cmbFechaD.Value;
			DateTime dtFechaH = (DateTime) cmbFechaH.Value;
			string CC ="0"; string Proy="0"; string Cli="0";
			if(!chkCentroC.Checked && cmbCentroC.Text.Length>0)
				CC= cmbCentroC.Value.ToString();
			if(!chkCliente.Checked && cmbCliente.Text.Length>0)
				Cli = cmbCliente.Value.ToString();
			if(!chkProyecto.Checked && cmbProyecto.Text.Length>0)
				Proy= cmbProyecto.Value.ToString();



			Funcion miFuncion = new Funcion();
			string strConn = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);	
			string stProcedimiento = "exec BuscarReembolsosGastos '"+dtFechaD.ToString("yyyyMMdd")+"' ,'"+dtFechaH.ToString("yyyyMMdd")+"', "
				+Proy+", "+CC+", "+Cli+", 0";

			ultraGrid1.DataSource = Funcion.dvProcedimiento(cdsCliente, stProcedimiento);
			#region Visibilidad
			ultraGrid1.DisplayLayout.Bands[0].Columns["idCompra"].Hidden = true;		
			ultraGrid1.DisplayLayout.Bands[0].Columns["idDetCompra"].Hidden = true;		
			#endregion Visibilidad
		}

		private void btAgregarTodo_Click(object sender, System.EventArgs e)
		{
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr1 in this.ultraGrid1.Rows.All)
			{
				dr1.Activate();
				btAgregar_Click(this, e);
			}
		}

		private void btAgregar_Click(object sender, System.EventArgs e)
		{
			errorProvider.SetError(ultraGrid1, "");

			int IdDetCompra = 0;
			try
			{
				IdDetCompra = (int) ultraGrid1.ActiveRow.Cells["idDetCompra"].Value;
			}
			catch
			{
				errorProvider.SetError(ultraGrid1, "Escoja Transaccion para pasar el detalle");
				return;
			}

			// Verificacion de existencia de factura en la lista de la derecha
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr2 in this.ultraGrid2.Rows.All)
			{
				if ((int) dr2.Cells["idDetCompra"].Value == IdDetCompra)
				{
					MessageBox.Show("Transaccion ya esta pasada", "Información");
					return; 
				}
			}
			string stSelect = "exec BuscarReembolsosGastos 0,0,0,0,0, "+IdDetCompra;
//
			Funcion miFuncion = new Funcion();
			string strConn = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);				
			SqlDataAdapter myData2 = new SqlDataAdapter(stSelect, strConn);
			myData2.TableMappings.Add("Tabla1", "ExcelSube2");			
			//			miDataSetN.Reset();
			//			miDataSetN.Clear();
			try
			{
				myData2.Fill(miDataSetN);
			}
			catch(Exception ex)
			{
				MessageBox.Show(ex.Message, "Error al abrir la Hoja A del Archivo: " );
				return;
			}
			dvN = miDataSetN.Tables[0].DefaultView;
			ultraGrid2.DataSource = dvN;	
//			ultraGrid2.DataSource = Funcion.dvProcedimiento(cdsCliente, stSelect);
			#region Visibilidad
			ultraGrid2.DisplayLayout.Bands[0].Columns["idDetCompra"].Hidden = true;			
			ultraGrid2.DisplayLayout.Bands[0].Columns["idCompra"].Hidden = true;						
			#endregion Visibilidad
		}

		private void ultraButton1_Click(object sender, System.EventArgs e)
		{
			errorProvider.SetError(ultraGrid1, "");
			try
			{
				ultraGrid2.ActiveRow.Delete(false);
			}
			catch(Exception ex)
			{
					Console.Write(ex);
				errorProvider.SetError(ultraGrid1, "Escoja Transaccion para pasar el detalle");
			}
		}

		private void ultraButton2_Click(object sender, System.EventArgs e)
		{
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.ultraGrid2.Rows.All)
			{
				dr.Delete(false);
			}	
		}

		private void btAceptar_Click(object sender, System.EventArgs e)
		{
			errorProvider1.SetError(btAceptar,"");
			if(ultraGrid2.Rows.Count < 1)
			{
				errorProvider1.SetError(this.btAceptar,"No Hay datos para Generar....");
				return;
			}
			if (DialogResult.No == MessageBox.Show("Desea Generar Facturas", 
				"Información", MessageBoxButtons.YesNo, MessageBoxIcon.Question)) return;	
			int vIdCompra= CreaCompra();

			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr3 in ultraGrid2.Rows.All)
			{																
				Cursor = Cursors.WaitCursor;
					int vIdCompraDC = (int) dr3.Cells["idDetCompra"].Value;
				string stCreaFactDet = "Exec CreaVentaReembolsosGastos "+vIdCompra+" ,"+vIdCompraDC;	
				Funcion.EjecutaSQL(cdsCompra, stCreaFactDet, true);																														
			}					
			ultraButton2_Click(this,e);
			btFiltro_Click(this,e);
			Cursor = Cursors.Default;					
		}
		private int CreaCompra()
		{
			string stNumero = Funcion.NumeraFactura(cdsCliente, 1);
			if(stNumero.Length <= 0)stNumero = "0";
			string sCliente ="0";
			string stExec = "Exec CreaFactura '" + sCliente + "', '"
				+ "', '" + stNumero + "', '" + DateTime.Today.ToString("yyyyMMdd") + "',0,0,0,0,0"
				+ ", 1, 0, 2, 1, '', '"
				+ DateTime.Today.ToString("yyyyMMdd") + "'";
		
			int IdCompra = Funcion.iEscalarSQL(cdsCliente, stExec, true);
			return IdCompra;
		}

		private void cdsCompra_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			this.cdsCompra.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);
		}

		private void chkCliente_CheckedChanged(object sender, System.EventArgs e)
		{
			if(!chkCliente.Checked)
				cmbCliente.Enabled = true;
			else
				cmbCliente.Enabled = false;
		}

		private void chkCentroC_CheckedChanged(object sender, System.EventArgs e)
		{
			if(!chkCentroC.Checked)
				cmbCentroC.Enabled = true;
			else
				cmbCentroC.Enabled = false;
		}

		private void chkProyecto_CheckedChanged(object sender, System.EventArgs e)
		{
			if(!chkProyecto.Checked)
				cmbProyecto.Enabled = true;
			else
				cmbProyecto.Enabled = false;
		}

		private void cdsCliente_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			this.cdsCliente.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);
		}

		private void cdsProyecto_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			this.cdsProyecto.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);
		}

		

	}
}
