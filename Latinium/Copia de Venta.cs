using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using Infragistics.Win.UltraWinGrid;
using System.Data;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de Venta.
	/// </summary>
	public class Venta : DevExpress.XtraEditors.XtraForm
	{
		private Infragistics.Win.UltraWinGrid.UltraGrid ultraGrid1;
		private Infragistics.Win.Misc.UltraButton ultraButton1;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtCliente;
		private Infragistics.Win.Misc.UltraButton btCliente;
		private Infragistics.Win.Misc.UltraLabel ultraLabel2;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtNumero;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbComprobante;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbVendedor;
		private Infragistics.Win.Misc.UltraButton btDeshacer;
		private Infragistics.Win.Misc.UltraButton btAnular;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor ultraNumericEditor1;
		private C1.Data.C1DataSet cdsSeteoF;
		private C1.Data.C1DataSet cdsVenta;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtArticulo;
		private Infragistics.Win.Misc.UltraLabel ultraLabel1;
		private Infragistics.Win.UltraWinTabControl.UltraTabControl ultraTabControl1;
		private Infragistics.Win.UltraWinTabControl.UltraTabSharedControlsPage ultraTabSharedControlsPage1;
		private Infragistics.Win.UltraWinTabControl.UltraTabPageControl ultraTabPageControl1;
		private Infragistics.Win.UltraWinTabControl.UltraTabPageControl ultraTabPageControl2;
		private BarraDatoSQL.BarraDatoSQL barraDato1;
		private Infragistics.Win.UltraWinGrid.UltraGrid ultraGrid2;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public Venta()
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
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("_VentaDetalle", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idVentaDetalle");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cantidad");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Precio");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Impuesto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idArticulo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Articulo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idVenta");
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinTabControl.UltraTab ultraTab1 = new Infragistics.Win.UltraWinTabControl.UltraTab();
			Infragistics.Win.UltraWinTabControl.UltraTab ultraTab2 = new Infragistics.Win.UltraWinTabControl.UltraTab();
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("VentaDetalle", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idVentaDetalle");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cantidad");
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn11 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Precio");
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn12 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Impuesto");
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn13 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idArticulo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn14 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Articulo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn15 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn16 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idVenta");
			Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance11 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance12 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance13 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance14 = new Infragistics.Win.Appearance();
			this.ultraTabPageControl1 = new Infragistics.Win.UltraWinTabControl.UltraTabPageControl();
			this.ultraTabPageControl2 = new Infragistics.Win.UltraWinTabControl.UltraTabPageControl();
			this.ultraGrid1 = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.cdsVenta = new C1.Data.C1DataSet();
			this.ultraButton1 = new Infragistics.Win.Misc.UltraButton();
			this.txtNumero = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtCliente = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.btCliente = new Infragistics.Win.Misc.UltraButton();
			this.cmbComprobante = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.cmbVendedor = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.ultraLabel2 = new Infragistics.Win.Misc.UltraLabel();
			this.btDeshacer = new Infragistics.Win.Misc.UltraButton();
			this.btAnular = new Infragistics.Win.Misc.UltraButton();
			this.ultraNumericEditor1 = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.txtArticulo = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.ultraLabel1 = new Infragistics.Win.Misc.UltraLabel();
			this.ultraTabControl1 = new Infragistics.Win.UltraWinTabControl.UltraTabControl();
			this.ultraTabSharedControlsPage1 = new Infragistics.Win.UltraWinTabControl.UltraTabSharedControlsPage();
			this.barraDato1 = new BarraDatoSQL.BarraDatoSQL();
			this.ultraGrid2 = new Infragistics.Win.UltraWinGrid.UltraGrid();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsVenta)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNumero)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCliente)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbComprobante)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbVendedor)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraNumericEditor1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtArticulo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraTabControl1)).BeginInit();
			this.ultraTabControl1.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid2)).BeginInit();
			this.SuspendLayout();
			// 
			// ultraTabPageControl1
			// 
			this.ultraTabPageControl1.Location = new System.Drawing.Point(23, 1);
			this.ultraTabPageControl1.Name = "ultraTabPageControl1";
			this.ultraTabPageControl1.Size = new System.Drawing.Size(270, 84);
			// 
			// ultraTabPageControl2
			// 
			this.ultraTabPageControl2.Location = new System.Drawing.Point(-10000, -10000);
			this.ultraTabPageControl2.Name = "ultraTabPageControl2";
			this.ultraTabPageControl2.Size = new System.Drawing.Size(270, 84);
			// 
			// ultraGrid1
			// 
			this.ultraGrid1.Cursor = System.Windows.Forms.Cursors.Default;
			this.ultraGrid1.DataMember = "_Venta.Venta_Detalle";
			this.ultraGrid1.DataSource = this.cdsVenta;
			appearance1.BackColor = System.Drawing.Color.White;
			this.ultraGrid1.DisplayLayout.Appearance = appearance1;
			ultraGridColumn1.Header.VisiblePosition = 5;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn2.Header.VisiblePosition = 0;
			ultraGridColumn3.Header.VisiblePosition = 3;
			ultraGridColumn4.Header.Caption = "Iva";
			ultraGridColumn4.Header.VisiblePosition = 4;
			ultraGridColumn4.Width = 50;
			ultraGridColumn5.Header.VisiblePosition = 6;
			ultraGridColumn5.Hidden = true;
			ultraGridColumn6.Header.VisiblePosition = 2;
			ultraGridColumn6.Width = 212;
			ultraGridColumn7.Header.VisiblePosition = 1;
			ultraGridColumn8.Header.VisiblePosition = 7;
			ultraGridColumn8.Hidden = true;
			ultraGridBand1.Columns.AddRange(new object[] {
																										 ultraGridColumn1,
																										 ultraGridColumn2,
																										 ultraGridColumn3,
																										 ultraGridColumn4,
																										 ultraGridColumn5,
																										 ultraGridColumn6,
																										 ultraGridColumn7,
																										 ultraGridColumn8});
			this.ultraGrid1.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.ultraGrid1.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.TemplateOnBottom;
			this.ultraGrid1.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.True;
			this.ultraGrid1.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance2.BackColor = System.Drawing.Color.Transparent;
			this.ultraGrid1.DisplayLayout.Override.CardAreaAppearance = appearance2;
			appearance3.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance3.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance3.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance3.FontData.BoldAsString = "True";
			appearance3.FontData.Name = "Arial";
			appearance3.FontData.SizeInPoints = 10F;
			appearance3.ForeColor = System.Drawing.Color.White;
			appearance3.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.ultraGrid1.DisplayLayout.Override.HeaderAppearance = appearance3;
			appearance4.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance4.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance4.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.RowSelectorAppearance = appearance4;
			appearance5.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance5.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance5.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.SelectedRowAppearance = appearance5;
			this.ultraGrid1.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.ultraGrid1.Location = new System.Drawing.Point(104, 104);
			this.ultraGrid1.Name = "ultraGrid1";
			this.ultraGrid1.Size = new System.Drawing.Size(360, 184);
			this.ultraGrid1.TabIndex = 0;
			this.ultraGrid1.UpdateMode = Infragistics.Win.UltraWinGrid.UpdateMode.OnCellChangeOrLostFocus;
			this.ultraGrid1.AfterRowInsert += new Infragistics.Win.UltraWinGrid.RowEventHandler(this.ultraGrid1_AfterRowInsert);
			this.ultraGrid1.AfterCellUpdate += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.ultraGrid1_AfterCellUpdate);
			// 
			// cdsVenta
			// 
			this.cdsVenta.BindingContextCtrl = this;
			this.cdsVenta.DataLibrary = "LibFacturacion";
			this.cdsVenta.DataLibraryUrl = "";
			this.cdsVenta.DataSetDef = "dsVenta";
			this.cdsVenta.FillOnRequest = false;
			this.cdsVenta.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsVenta.Name = "cdsVenta";
			this.cdsVenta.SchemaClassName = "LibFacturacion.DataClass";
			this.cdsVenta.SchemaDef = null;
			// 
			// ultraButton1
			// 
			this.ultraButton1.Location = new System.Drawing.Point(16, 112);
			this.ultraButton1.Name = "ultraButton1";
			this.ultraButton1.Size = new System.Drawing.Size(72, 40);
			this.ultraButton1.TabIndex = 1;
			this.ultraButton1.Text = "Nuevo";
			this.ultraButton1.Click += new System.EventHandler(this.ultraButton1_Click);
			// 
			// txtNumero
			// 
			this.txtNumero.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsVenta, "_Venta.Numero"));
			this.txtNumero.Location = new System.Drawing.Point(16, 16);
			this.txtNumero.Name = "txtNumero";
			this.txtNumero.Size = new System.Drawing.Size(72, 21);
			this.txtNumero.TabIndex = 2;
			// 
			// txtCliente
			// 
			this.txtCliente.Location = new System.Drawing.Point(168, 16);
			this.txtCliente.Name = "txtCliente";
			this.txtCliente.Size = new System.Drawing.Size(224, 21);
			this.txtCliente.TabIndex = 4;
			// 
			// btCliente
			// 
			this.btCliente.Location = new System.Drawing.Point(104, 16);
			this.btCliente.Name = "btCliente";
			this.btCliente.Size = new System.Drawing.Size(56, 24);
			this.btCliente.TabIndex = 5;
			this.btCliente.Text = "Cliente";
			this.btCliente.Click += new System.EventHandler(this.btCliente_Click);
			// 
			// cmbComprobante
			// 
			this.cmbComprobante.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbComprobante.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbComprobante.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsVenta, "_Venta.idComprobante"));
			this.cmbComprobante.DisplayMember = "";
			this.cmbComprobante.Location = new System.Drawing.Point(280, 48);
			this.cmbComprobante.Name = "cmbComprobante";
			this.cmbComprobante.Size = new System.Drawing.Size(112, 21);
			this.cmbComprobante.TabIndex = 6;
			this.cmbComprobante.Text = "ultraCombo1";
			this.cmbComprobante.ValueMember = "";
			// 
			// cmbVendedor
			// 
			this.cmbVendedor.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbVendedor.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbVendedor.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsVenta, "_Venta.idVendedor"));
			this.cmbVendedor.DisplayMember = "";
			this.cmbVendedor.Location = new System.Drawing.Point(168, 48);
			this.cmbVendedor.Name = "cmbVendedor";
			this.cmbVendedor.Size = new System.Drawing.Size(112, 21);
			this.cmbVendedor.TabIndex = 7;
			this.cmbVendedor.Text = "ultraCombo2";
			this.cmbVendedor.ValueMember = "";
			// 
			// ultraLabel2
			// 
			this.ultraLabel2.AutoSize = true;
			this.ultraLabel2.Location = new System.Drawing.Point(104, 48);
			this.ultraLabel2.Name = "ultraLabel2";
			this.ultraLabel2.Size = new System.Drawing.Size(53, 14);
			this.ultraLabel2.TabIndex = 8;
			this.ultraLabel2.Text = "Vendedor";
			// 
			// btDeshacer
			// 
			this.btDeshacer.Location = new System.Drawing.Point(16, 160);
			this.btDeshacer.Name = "btDeshacer";
			this.btDeshacer.Size = new System.Drawing.Size(72, 40);
			this.btDeshacer.TabIndex = 9;
			this.btDeshacer.Text = "Deshacer";
			this.btDeshacer.Click += new System.EventHandler(this.btDeshacer_Click);
			// 
			// btAnular
			// 
			this.btAnular.Location = new System.Drawing.Point(16, 208);
			this.btAnular.Name = "btAnular";
			this.btAnular.Size = new System.Drawing.Size(72, 40);
			this.btAnular.TabIndex = 10;
			this.btAnular.Text = "Anular";
			// 
			// ultraNumericEditor1
			// 
			this.ultraNumericEditor1.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsVenta, "_Venta.Total"));
			this.ultraNumericEditor1.FormatString = "#,##0.00";
			this.ultraNumericEditor1.Location = new System.Drawing.Point(608, 304);
			this.ultraNumericEditor1.Name = "ultraNumericEditor1";
			this.ultraNumericEditor1.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.ultraNumericEditor1.PromptChar = ' ';
			this.ultraNumericEditor1.Size = new System.Drawing.Size(112, 21);
			this.ultraNumericEditor1.TabIndex = 11;
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
			this.cdsSeteoF.BeforeFill += new C1.Data.FillEventHandler(this.cdsProyecto_BeforeFill);
			// 
			// txtArticulo
			// 
			this.txtArticulo.Location = new System.Drawing.Point(168, 72);
			this.txtArticulo.Name = "txtArticulo";
			this.txtArticulo.Size = new System.Drawing.Size(224, 21);
			this.txtArticulo.TabIndex = 12;
			this.txtArticulo.Validating += new System.ComponentModel.CancelEventHandler(this.txtArticulo_Validating);
			// 
			// ultraLabel1
			// 
			this.ultraLabel1.AutoSize = true;
			this.ultraLabel1.Location = new System.Drawing.Point(104, 72);
			this.ultraLabel1.Name = "ultraLabel1";
			this.ultraLabel1.Size = new System.Drawing.Size(42, 14);
			this.ultraLabel1.TabIndex = 13;
			this.ultraLabel1.Text = "Articulo";
			// 
			// ultraTabControl1
			// 
			this.ultraTabControl1.Controls.Add(this.ultraTabSharedControlsPage1);
			this.ultraTabControl1.Controls.Add(this.ultraTabPageControl1);
			this.ultraTabControl1.Controls.Add(this.ultraTabPageControl2);
			this.ultraTabControl1.Location = new System.Drawing.Point(424, 8);
			this.ultraTabControl1.Name = "ultraTabControl1";
			this.ultraTabControl1.SharedControlsPage = this.ultraTabSharedControlsPage1;
			this.ultraTabControl1.Size = new System.Drawing.Size(296, 88);
			this.ultraTabControl1.TabIndex = 14;
			this.ultraTabControl1.TabOrientation = Infragistics.Win.UltraWinTabs.TabOrientation.LeftTop;
			ultraTab1.TabPage = this.ultraTabPageControl1;
			ultraTab1.Text = "tab1";
			ultraTab2.TabPage = this.ultraTabPageControl2;
			ultraTab2.Text = "tab2";
			this.ultraTabControl1.Tabs.AddRange(new Infragistics.Win.UltraWinTabControl.UltraTab[] {
																																															 ultraTab1,
																																															 ultraTab2});
			// 
			// ultraTabSharedControlsPage1
			// 
			this.ultraTabSharedControlsPage1.Location = new System.Drawing.Point(-10000, -10000);
			this.ultraTabSharedControlsPage1.Name = "ultraTabSharedControlsPage1";
			this.ultraTabSharedControlsPage1.Size = new System.Drawing.Size(270, 84);
			// 
			// barraDato1
			// 
			this.barraDato1.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.barraDato1.BarraMovimiento = true;
			this.barraDato1.DataMember = "_Venta";
			this.barraDato1.DataNombreId = "idVenta";
			this.barraDato1.DataOrden = "";
			this.barraDato1.DataSource = this.cdsVenta;
			this.barraDato1.DataTabla = "Venta";
			this.barraDato1.DataTablaHija = "VentaDetalle";
			this.barraDato1.Location = new System.Drawing.Point(24, 296);
			this.barraDato1.Name = "barraDato1";
			this.barraDato1.Size = new System.Drawing.Size(304, 24);
			this.barraDato1.TabIndex = 39;
			this.barraDato1.VisibleBorrar = true;
			this.barraDato1.VisibleBuscar = true;
			this.barraDato1.VisibleEditar = true;
			this.barraDato1.VisibleImprimir = true;
			this.barraDato1.VisibleNuevo = true;
			this.barraDato1.Refresca += new System.EventHandler(this.barraDato1_Refresca);
			this.barraDato1.Crear += new System.EventHandler(this.barraDato1_Crear);
			this.barraDato1.Deshacer += new System.EventHandler(this.barraDato1_Deshacer);
			this.barraDato1.Editar += new System.EventHandler(this.barraDato1_Editar);
			this.barraDato1.Grabar += new System.EventHandler(this.barraDato1_Grabar);
			// 
			// ultraGrid2
			// 
			this.ultraGrid2.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.ultraGrid2.Cursor = System.Windows.Forms.Cursors.Default;
			this.ultraGrid2.DataMember = "VentaDetalle";
			this.ultraGrid2.DataSource = this.cdsVenta;
			appearance6.BackColor = System.Drawing.Color.White;
			this.ultraGrid2.DisplayLayout.Appearance = appearance6;
			ultraGridBand2.AddButtonCaption = "_DetCompra";
			ultraGridColumn9.Header.VisiblePosition = 0;
			appearance7.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn10.CellAppearance = appearance7;
			ultraGridColumn10.Header.VisiblePosition = 2;
			ultraGridColumn10.MaskInput = "{LOC}nn,nnn,nnn.nn";
			ultraGridColumn10.NullText = "0";
			ultraGridColumn10.PromptChar = ' ';
			appearance8.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn11.CellAppearance = appearance8;
			ultraGridColumn11.Format = "";
			ultraGridColumn11.Header.VisiblePosition = 1;
			ultraGridColumn11.MaskInput = "";
			ultraGridColumn11.PromptChar = ' ';
			ultraGridColumn11.Width = 89;
			appearance9.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn12.CellAppearance = appearance9;
			ultraGridColumn12.Header.Caption = "Iva";
			ultraGridColumn12.Header.VisiblePosition = 7;
			ultraGridColumn12.PromptChar = ' ';
			ultraGridColumn12.Width = 34;
			ultraGridColumn13.Header.Caption = "Artículo";
			ultraGridColumn13.Header.VisiblePosition = 3;
			ultraGridColumn13.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownValidate;
			ultraGridColumn13.Width = 216;
			ultraGridColumn14.Header.VisiblePosition = 5;
			ultraGridColumn15.Header.VisiblePosition = 4;
			ultraGridColumn15.Width = 118;
			ultraGridColumn16.Header.VisiblePosition = 6;
			ultraGridBand2.Columns.AddRange(new object[] {
																										 ultraGridColumn9,
																										 ultraGridColumn10,
																										 ultraGridColumn11,
																										 ultraGridColumn12,
																										 ultraGridColumn13,
																										 ultraGridColumn14,
																										 ultraGridColumn15,
																										 ultraGridColumn16});
			this.ultraGrid2.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			this.ultraGrid2.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.TemplateOnBottom;
			this.ultraGrid2.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.True;
			this.ultraGrid2.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance10.BackColor = System.Drawing.Color.Transparent;
			this.ultraGrid2.DisplayLayout.Override.CardAreaAppearance = appearance10;
			appearance11.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance11.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance11.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance11.FontData.BoldAsString = "True";
			appearance11.FontData.Name = "Arial";
			appearance11.FontData.SizeInPoints = 10F;
			appearance11.ForeColor = System.Drawing.Color.White;
			appearance11.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.ultraGrid2.DisplayLayout.Override.HeaderAppearance = appearance11;
			this.ultraGrid2.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortSingle;
			appearance12.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance12.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance12.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid2.DisplayLayout.Override.RowAlternateAppearance = appearance12;
			appearance13.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance13.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance13.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid2.DisplayLayout.Override.RowSelectorAppearance = appearance13;
			appearance14.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance14.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance14.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid2.DisplayLayout.Override.SelectedRowAppearance = appearance14;
			this.ultraGrid2.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.ultraGrid2.Location = new System.Drawing.Point(496, 104);
			this.ultraGrid2.Name = "ultraGrid2";
			this.ultraGrid2.Size = new System.Drawing.Size(232, 184);
			this.ultraGrid2.TabIndex = 40;
			// 
			// Venta
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(744, 342);
			this.Controls.Add(this.ultraGrid2);
			this.Controls.Add(this.barraDato1);
			this.Controls.Add(this.ultraTabControl1);
			this.Controls.Add(this.ultraLabel1);
			this.Controls.Add(this.txtArticulo);
			this.Controls.Add(this.ultraNumericEditor1);
			this.Controls.Add(this.btAnular);
			this.Controls.Add(this.btDeshacer);
			this.Controls.Add(this.ultraLabel2);
			this.Controls.Add(this.cmbVendedor);
			this.Controls.Add(this.cmbComprobante);
			this.Controls.Add(this.btCliente);
			this.Controls.Add(this.txtCliente);
			this.Controls.Add(this.txtNumero);
			this.Controls.Add(this.ultraButton1);
			this.Controls.Add(this.ultraGrid1);
			this.Name = "Venta";
			this.Text = "Punto de Venta";
			this.Load += new System.EventHandler(this.Venta_Load);
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsVenta)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNumero)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCliente)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbComprobante)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbVendedor)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraNumericEditor1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtArticulo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraTabControl1)).EndInit();
			this.ultraTabControl1.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid2)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void cdsProyecto_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			this.cdsSeteoF.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);
		}

		private void txtArticulo_Validating(object sender, System.ComponentModel.CancelEventArgs e)
		{
			string stArticulo = txtArticulo.Text.Trim();
			int idArticulo = Funcion.iEscalarSQL(cdsSeteoF, "Exec VentaArticuloCuenta '" +
				stArticulo + "'", true);
			if (idArticulo > 0)
			{
//				C1.Data.C1DataTable dtCustomers = cdsVenta.Tables["VentaDetalle"];
//				try
//				{
//					C1.Data.C1DataRow drVenta = dtCustomers.AddNew();
//					drVenta["Cantidad"] = 2;
//				}
//				catch(System.Exception ex)
//				{
//					MessageBox.Show(ex.Message);
//				}
				ultraGrid1.Rows.Band.AddNew();
				ultraGrid1.PerformAction(UltraGridAction.EnterEditMode, false, false);
				ultraGrid1.ActiveCell = ultraGrid1.ActiveRow.Cells["idArticulo"];
				ultraGrid1.ActiveCell.Value = idArticulo;
//				string Codigo = Funcion.sEscalarSQL(cdsSeteoF, "Exec VentaArticuloCodigo "
//					+ idArticulo.ToString(), true);
//				int iPos = Codigo.IndexOf('~');
//				string Articulo = Codigo.Substring(0, iPos);
//				Codigo = Codigo.Substring(iPos+1);
//				ultraGrid1.ActiveCell = ultraGrid1.ActiveRow.Cells["Codigo"];
//				ultraGrid1.ActiveCell.Value = Codigo;
//				ultraGrid1.ActiveCell = ultraGrid1.ActiveRow.Cells["Articulo"];
//				ultraGrid1.ActiveCell.Value = Articulo;
				ultraGrid1.PerformAction(UltraGridAction.ExitEditMode, false, false);
//				ultraGrid1.PerformAction(UltraGridAction.AboveCell, false, false);
				ultraGrid1.PerformAction(UltraGridAction.EnterEditMode, false, false);
				//				ultraGrid1.PerformAction(UltraGridAction.ActivateCell, false, false);
				txtArticulo.Text = "";
			}
		}

		private void barraDato1_Refresca(object sender, System.EventArgs e)
		{
		
		}

		private void barraDato1_Crear(object sender, System.EventArgs e)
		{
			barraDato1.CrearRegistro();
		}

		private void barraDato1_Grabar(object sender, System.EventArgs e)
		{
			barraDato1.GrabaRegistro();
		}

		private void barraDato1_Deshacer(object sender, System.EventArgs e)
		{
			barraDato1.DeshacerRegistro();
		}

		private void barraDato1_Editar(object sender, System.EventArgs e)
		{
			barraDato1.EditarRegistro();
		}

		private void ultraGrid1_AfterCellUpdate(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
		}

		private void ultraGrid1_AfterRowInsert(object sender, Infragistics.Win.UltraWinGrid.RowEventArgs e)
		{
		}

		private void Venta_Load(object sender, System.EventArgs e)
		{
			barraDato1.ProximoId(4);
		}

		private void btCliente_Click(object sender, System.EventArgs e)
		{
//			ultraGrid1.Rows.Band.AddNew();
			ultraGrid1.Rows.Band.AddNew();
//			ultraGrid1.PerformAction(UltraGridAction.AboveCell, false, false);
//			ultraGrid1.PerformAction(UltraGridAction.ExitEditMode, false, false);
//			ultraGrid1.PerformAction(UltraGridAction.BelowCell, false, false);
//			ultraGrid1.PerformAction(UltraGridAction.EnterEditMode, false, false);
			ultraGrid1.ActiveCell = ultraGrid1.ActiveRow.Cells["Precio"];
			ultraGrid1.ActiveCell.Value = 2;
			ultraGrid1.PerformAction(UltraGridAction.EnterEditMode, false, false);
			ultraGrid1.Refresh();
//			ultraGrid1.PerformAction(UltraGridAction.ExitEditMode, false, false);
//			ultraGrid1.PerformAction(UltraGridAction.BelowRow, false, false);
			//			ultraGrid1.PerformAction(UltraGridAction.DeleteRows, false, false);
		}

		private void ultraButton1_Click(object sender, System.EventArgs e)
		{
			ultraGrid1.Rows.Band.AddNew();		
		}

		private void btDeshacer_Click(object sender, System.EventArgs e)
		{
			ultraGrid2.Rows.Band.AddNew();
			ultraGrid2.PerformAction(UltraGridAction.EnterEditMode, false, false);
			ultraGrid2.ActiveCell = ultraGrid2.ActiveRow.Cells["Precio"];
			ultraGrid2.ActiveCell.Value = 2;
		}
	}
}
