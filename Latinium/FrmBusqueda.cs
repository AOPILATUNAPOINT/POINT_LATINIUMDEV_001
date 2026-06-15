using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de FrmBusqueda.
	/// </summary>
	public class FrmBusqueda : DevExpress.XtraEditors.XtraForm
	{
		private System.Windows.Forms.TabControl tabControl1;
		private System.Windows.Forms.TabPage tbBasico;
		private System.Windows.Forms.TabPage tbComplemento;
		private System.Windows.Forms.TabPage tbPromocion;
		private System.Windows.Forms.TextBox txtArticulo;
		private C1.Data.C1DataSet cdsSeteo;
		private System.Windows.Forms.Button btArticulo;
		private System.Windows.Forms.TextBox txtCliente;
		private System.Windows.Forms.Button btProveedor;
		private System.Windows.Forms.TextBox textBox1;
		private Infragistics.Win.UltraWinGrid.UltraGrid grdArticulo;
		private System.Windows.Forms.Button btCliente;
		private Infragistics.Win.UltraWinGrid.UltraGrid grdCliente;
		private Infragistics.Win.UltraWinGrid.UltraGrid grdProveedor;
		private System.Windows.Forms.TabPage tbAnalisis;
		private DevExpress.XtraEditors.SimpleButton btKardex;
		private DevExpress.XtraGrid.GridControl grdArticulos;
		private DevExpress.XtraEditors.SimpleButton btArticuloB;
		private DevExpress.XtraEditors.SimpleButton btTransaccion;
		private DevExpress.XtraGrid.Views.Grid.GridView grvArticulos;
		private DevExpress.XtraEditors.DateEdit cmbDesde;
		private DevExpress.XtraEditors.DateEdit cmbHasta;
		private DevExpress.XtraEditors.CheckEdit chkAgrupar;
		private DevExpress.XtraEditors.GridLookUpEdit cmbTipo;
		private DevExpress.XtraGrid.Views.Grid.GridView gridLookUpEdit1View;
		private DevExpress.XtraGrid.Columns.GridColumn gridColumn1;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public FrmBusqueda()
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
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance11 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance12 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance13 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance14 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance15 = new Infragistics.Win.Appearance();
			this.tabControl1 = new System.Windows.Forms.TabControl();
			this.tbBasico = new System.Windows.Forms.TabPage();
			this.grdArticulo = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.btArticulo = new System.Windows.Forms.Button();
			this.txtArticulo = new System.Windows.Forms.TextBox();
			this.tbComplemento = new System.Windows.Forms.TabPage();
			this.grdCliente = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.btCliente = new System.Windows.Forms.Button();
			this.txtCliente = new System.Windows.Forms.TextBox();
			this.tbPromocion = new System.Windows.Forms.TabPage();
			this.grdProveedor = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.btProveedor = new System.Windows.Forms.Button();
			this.textBox1 = new System.Windows.Forms.TextBox();
			this.tbAnalisis = new System.Windows.Forms.TabPage();
			this.grdArticulos = new DevExpress.XtraGrid.GridControl();
			this.grvArticulos = new DevExpress.XtraGrid.Views.Grid.GridView();
			this.btKardex = new DevExpress.XtraEditors.SimpleButton();
			this.cdsSeteo = new C1.Data.C1DataSet();
			this.btArticuloB = new DevExpress.XtraEditors.SimpleButton();
			this.btTransaccion = new DevExpress.XtraEditors.SimpleButton();
			this.cmbDesde = new DevExpress.XtraEditors.DateEdit();
			this.cmbHasta = new DevExpress.XtraEditors.DateEdit();
			this.chkAgrupar = new DevExpress.XtraEditors.CheckEdit();
			this.cmbTipo = new DevExpress.XtraEditors.GridLookUpEdit();
			this.gridLookUpEdit1View = new DevExpress.XtraGrid.Views.Grid.GridView();
			this.gridColumn1 = new DevExpress.XtraGrid.Columns.GridColumn();
			this.tabControl1.SuspendLayout();
			this.tbBasico.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.grdArticulo)).BeginInit();
			this.tbComplemento.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.grdCliente)).BeginInit();
			this.tbPromocion.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.grdProveedor)).BeginInit();
			this.tbAnalisis.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.grdArticulos)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.grvArticulos)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbDesde.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbHasta.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.chkAgrupar.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbTipo.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.gridLookUpEdit1View)).BeginInit();
			this.SuspendLayout();
			// 
			// tabControl1
			// 
			this.tabControl1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.tabControl1.Controls.Add(this.tbBasico);
			this.tabControl1.Controls.Add(this.tbComplemento);
			this.tabControl1.Controls.Add(this.tbPromocion);
			this.tabControl1.Controls.Add(this.tbAnalisis);
			this.tabControl1.Location = new System.Drawing.Point(16, 17);
			this.tabControl1.Name = "tabControl1";
			this.tabControl1.SelectedIndex = 0;
			this.tabControl1.Size = new System.Drawing.Size(688, 407);
			this.tabControl1.TabIndex = 29;
			// 
			// tbBasico
			// 
			this.tbBasico.Controls.Add(this.grdArticulo);
			this.tbBasico.Controls.Add(this.btArticulo);
			this.tbBasico.Controls.Add(this.txtArticulo);
			this.tbBasico.Location = new System.Drawing.Point(4, 22);
			this.tbBasico.Name = "tbBasico";
			this.tbBasico.Size = new System.Drawing.Size(672, 379);
			this.tbBasico.TabIndex = 0;
			this.tbBasico.Text = "Artículo";
			// 
			// grdArticulo
			// 
			this.grdArticulo.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.grdArticulo.Cursor = System.Windows.Forms.Cursors.Hand;
			appearance1.BackColor = System.Drawing.Color.White;
			this.grdArticulo.DisplayLayout.Appearance = appearance1;
			this.grdArticulo.DisplayLayout.AutoFitColumns = true;
			this.grdArticulo.DisplayLayout.Override.AllowRowFiltering = Infragistics.Win.DefaultableBoolean.True;
			appearance2.BackColor = System.Drawing.Color.Transparent;
			this.grdArticulo.DisplayLayout.Override.CardAreaAppearance = appearance2;
			appearance3.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance3.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance3.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance3.FontData.BoldAsString = "True";
			appearance3.FontData.Name = "Arial";
			appearance3.FontData.SizeInPoints = 10F;
			appearance3.ForeColor = System.Drawing.Color.White;
			appearance3.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.grdArticulo.DisplayLayout.Override.HeaderAppearance = appearance3;
			this.grdArticulo.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti;
			appearance4.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance4.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance4.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.grdArticulo.DisplayLayout.Override.RowSelectorAppearance = appearance4;
			appearance5.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance5.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance5.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.grdArticulo.DisplayLayout.Override.SelectedRowAppearance = appearance5;
			this.grdArticulo.DisplayLayout.ViewStyleBand = Infragistics.Win.UltraWinGrid.ViewStyleBand.OutlookGroupBy;
			this.grdArticulo.Location = new System.Drawing.Point(16, 52);
			this.grdArticulo.Name = "grdArticulo";
			this.grdArticulo.Size = new System.Drawing.Size(640, 312);
			this.grdArticulo.TabIndex = 31;
			this.grdArticulo.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.grdArticulo_InitializeLayout);
			// 
			// btArticulo
			// 
			this.btArticulo.Location = new System.Drawing.Point(16, 17);
			this.btArticulo.Name = "btArticulo";
			this.btArticulo.Size = new System.Drawing.Size(75, 25);
			this.btArticulo.TabIndex = 30;
			this.btArticulo.Text = "Artículo";
			this.btArticulo.Click += new System.EventHandler(this.btArticulo_Click);
			// 
			// txtArticulo
			// 
			this.txtArticulo.Location = new System.Drawing.Point(104, 17);
			this.txtArticulo.Name = "txtArticulo";
			this.txtArticulo.Size = new System.Drawing.Size(552, 20);
			this.txtArticulo.TabIndex = 29;
			this.txtArticulo.Text = "";
			// 
			// tbComplemento
			// 
			this.tbComplemento.Controls.Add(this.grdCliente);
			this.tbComplemento.Controls.Add(this.btCliente);
			this.tbComplemento.Controls.Add(this.txtCliente);
			this.tbComplemento.Location = new System.Drawing.Point(4, 22);
			this.tbComplemento.Name = "tbComplemento";
			this.tbComplemento.Size = new System.Drawing.Size(672, 379);
			this.tbComplemento.TabIndex = 1;
			this.tbComplemento.Text = "Cliente";
			this.tbComplemento.Visible = false;
			// 
			// grdCliente
			// 
			this.grdCliente.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.grdCliente.Cursor = System.Windows.Forms.Cursors.Hand;
			appearance6.BackColor = System.Drawing.Color.White;
			this.grdCliente.DisplayLayout.Appearance = appearance6;
			this.grdCliente.DisplayLayout.AutoFitColumns = true;
			this.grdCliente.DisplayLayout.Override.AllowRowFiltering = Infragistics.Win.DefaultableBoolean.True;
			appearance7.BackColor = System.Drawing.Color.Transparent;
			this.grdCliente.DisplayLayout.Override.CardAreaAppearance = appearance7;
			appearance8.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance8.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance8.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance8.FontData.BoldAsString = "True";
			appearance8.FontData.Name = "Arial";
			appearance8.FontData.SizeInPoints = 10F;
			appearance8.ForeColor = System.Drawing.Color.White;
			appearance8.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.grdCliente.DisplayLayout.Override.HeaderAppearance = appearance8;
			this.grdCliente.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti;
			appearance9.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance9.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance9.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.grdCliente.DisplayLayout.Override.RowSelectorAppearance = appearance9;
			appearance10.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance10.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance10.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.grdCliente.DisplayLayout.Override.SelectedRowAppearance = appearance10;
			this.grdCliente.DisplayLayout.ViewStyleBand = Infragistics.Win.UltraWinGrid.ViewStyleBand.OutlookGroupBy;
			this.grdCliente.Location = new System.Drawing.Point(16, 52);
			this.grdCliente.Name = "grdCliente";
			this.grdCliente.Size = new System.Drawing.Size(640, 312);
			this.grdCliente.TabIndex = 34;
			this.grdCliente.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.ultraGrid1_InitializeLayout);
			// 
			// btCliente
			// 
			this.btCliente.Location = new System.Drawing.Point(16, 16);
			this.btCliente.Name = "btCliente";
			this.btCliente.Size = new System.Drawing.Size(75, 25);
			this.btCliente.TabIndex = 33;
			this.btCliente.Text = "Cliente";
			this.btCliente.Click += new System.EventHandler(this.btCliente_Click);
			// 
			// txtCliente
			// 
			this.txtCliente.Location = new System.Drawing.Point(104, 16);
			this.txtCliente.Name = "txtCliente";
			this.txtCliente.Size = new System.Drawing.Size(552, 20);
			this.txtCliente.TabIndex = 32;
			this.txtCliente.Text = "";
			// 
			// tbPromocion
			// 
			this.tbPromocion.Controls.Add(this.grdProveedor);
			this.tbPromocion.Controls.Add(this.btProveedor);
			this.tbPromocion.Controls.Add(this.textBox1);
			this.tbPromocion.Location = new System.Drawing.Point(4, 22);
			this.tbPromocion.Name = "tbPromocion";
			this.tbPromocion.Size = new System.Drawing.Size(672, 379);
			this.tbPromocion.TabIndex = 3;
			this.tbPromocion.Text = "Proveedor";
			this.tbPromocion.Visible = false;
			// 
			// grdProveedor
			// 
			this.grdProveedor.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.grdProveedor.Cursor = System.Windows.Forms.Cursors.Hand;
			appearance11.BackColor = System.Drawing.Color.White;
			this.grdProveedor.DisplayLayout.Appearance = appearance11;
			this.grdProveedor.DisplayLayout.AutoFitColumns = true;
			this.grdProveedor.DisplayLayout.Override.AllowRowFiltering = Infragistics.Win.DefaultableBoolean.True;
			appearance12.BackColor = System.Drawing.Color.Transparent;
			this.grdProveedor.DisplayLayout.Override.CardAreaAppearance = appearance12;
			appearance13.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance13.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance13.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance13.FontData.BoldAsString = "True";
			appearance13.FontData.Name = "Arial";
			appearance13.FontData.SizeInPoints = 10F;
			appearance13.ForeColor = System.Drawing.Color.White;
			appearance13.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.grdProveedor.DisplayLayout.Override.HeaderAppearance = appearance13;
			this.grdProveedor.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti;
			appearance14.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance14.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance14.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.grdProveedor.DisplayLayout.Override.RowSelectorAppearance = appearance14;
			appearance15.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance15.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance15.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.grdProveedor.DisplayLayout.Override.SelectedRowAppearance = appearance15;
			this.grdProveedor.DisplayLayout.ViewStyleBand = Infragistics.Win.UltraWinGrid.ViewStyleBand.OutlookGroupBy;
			this.grdProveedor.Location = new System.Drawing.Point(16, 52);
			this.grdProveedor.Name = "grdProveedor";
			this.grdProveedor.Size = new System.Drawing.Size(640, 312);
			this.grdProveedor.TabIndex = 37;
			this.grdProveedor.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.grdProveedor_InitializeLayout);
			// 
			// btProveedor
			// 
			this.btProveedor.Location = new System.Drawing.Point(16, 16);
			this.btProveedor.Name = "btProveedor";
			this.btProveedor.Size = new System.Drawing.Size(75, 25);
			this.btProveedor.TabIndex = 36;
			this.btProveedor.Text = "Proveedor";
			this.btProveedor.Click += new System.EventHandler(this.btProveedor_Click);
			// 
			// textBox1
			// 
			this.textBox1.Location = new System.Drawing.Point(104, 16);
			this.textBox1.Name = "textBox1";
			this.textBox1.Size = new System.Drawing.Size(552, 20);
			this.textBox1.TabIndex = 35;
			this.textBox1.Text = "";
			// 
			// tbAnalisis
			// 
			this.tbAnalisis.Controls.Add(this.cmbTipo);
			this.tbAnalisis.Controls.Add(this.chkAgrupar);
			this.tbAnalisis.Controls.Add(this.cmbHasta);
			this.tbAnalisis.Controls.Add(this.cmbDesde);
			this.tbAnalisis.Controls.Add(this.btTransaccion);
			this.tbAnalisis.Controls.Add(this.btArticuloB);
			this.tbAnalisis.Controls.Add(this.grdArticulos);
			this.tbAnalisis.Controls.Add(this.btKardex);
			this.tbAnalisis.Location = new System.Drawing.Point(4, 22);
			this.tbAnalisis.Name = "tbAnalisis";
			this.tbAnalisis.Size = new System.Drawing.Size(680, 381);
			this.tbAnalisis.TabIndex = 4;
			this.tbAnalisis.Text = "Analisis";
			// 
			// grdArticulos
			// 
			this.grdArticulos.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			// 
			// grdArticulos.EmbeddedNavigator
			// 
			this.grdArticulos.EmbeddedNavigator.Name = "";
			this.grdArticulos.Location = new System.Drawing.Point(16, 40);
			this.grdArticulos.MainView = this.grvArticulos;
			this.grdArticulos.Name = "grdArticulos";
			this.grdArticulos.Size = new System.Drawing.Size(648, 322);
			this.grdArticulos.TabIndex = 1;
			this.grdArticulos.ViewCollection.AddRange(new DevExpress.XtraGrid.Views.Base.BaseView[] {
																																																this.grvArticulos});
			// 
			// grvArticulos
			// 
			this.grvArticulos.GridControl = this.grdArticulos;
			this.grvArticulos.Name = "grvArticulos";
			this.grvArticulos.OptionsView.ColumnAutoWidth = false;
			this.grvArticulos.OptionsView.ShowAutoFilterRow = true;
			// 
			// btKardex
			// 
			this.btKardex.Location = new System.Drawing.Point(16, 8);
			this.btKardex.Name = "btKardex";
			this.btKardex.Size = new System.Drawing.Size(56, 23);
			this.btKardex.TabIndex = 0;
			this.btKardex.Text = "Kardex";
			this.btKardex.Click += new System.EventHandler(this.btKardex_Click);
			// 
			// cdsSeteo
			// 
			this.cdsSeteo.BindingContextCtrl = this;
			this.cdsSeteo.DataLibrary = "LibFacturacion";
			this.cdsSeteo.DataLibraryUrl = "";
			this.cdsSeteo.DataSetDef = "dsSeteoF";
			this.cdsSeteo.Locale = new System.Globalization.CultureInfo("es-MX");
			this.cdsSeteo.Name = "cdsSeteo";
			this.cdsSeteo.SchemaClassName = "LibFacturacion.DataClass";
			this.cdsSeteo.SchemaDef = null;
			this.cdsSeteo.BeforeFill += new C1.Data.FillEventHandler(this.cdsSeteo_BeforeFill);
			// 
			// btArticuloB
			// 
			this.btArticuloB.Location = new System.Drawing.Point(568, 8);
			this.btArticuloB.Name = "btArticuloB";
			this.btArticuloB.Size = new System.Drawing.Size(40, 23);
			this.btArticuloB.TabIndex = 2;
			this.btArticuloB.Text = "Art.";
			this.btArticuloB.Click += new System.EventHandler(this.btArticuloB_Click);
			// 
			// btTransaccion
			// 
			this.btTransaccion.Location = new System.Drawing.Point(624, 8);
			this.btTransaccion.Name = "btTransaccion";
			this.btTransaccion.Size = new System.Drawing.Size(40, 23);
			this.btTransaccion.TabIndex = 3;
			this.btTransaccion.Text = "Trans.";
			this.btTransaccion.Click += new System.EventHandler(this.btTransaccion_Click);
			// 
			// cmbDesde
			// 
			this.cmbDesde.EditValue = null;
			this.cmbDesde.Location = new System.Drawing.Point(80, 8);
			this.cmbDesde.Name = "cmbDesde";
			// 
			// cmbDesde.Properties
			// 
			this.cmbDesde.Properties.Buttons.AddRange(new DevExpress.XtraEditors.Controls.EditorButton[] {
																																																		 new DevExpress.XtraEditors.Controls.EditorButton(DevExpress.XtraEditors.Controls.ButtonPredefines.Combo)});
			this.cmbDesde.Size = new System.Drawing.Size(80, 20);
			this.cmbDesde.TabIndex = 4;
			// 
			// cmbHasta
			// 
			this.cmbHasta.EditValue = null;
			this.cmbHasta.Location = new System.Drawing.Point(168, 8);
			this.cmbHasta.Name = "cmbHasta";
			// 
			// cmbHasta.Properties
			// 
			this.cmbHasta.Properties.Buttons.AddRange(new DevExpress.XtraEditors.Controls.EditorButton[] {
																																																		 new DevExpress.XtraEditors.Controls.EditorButton(DevExpress.XtraEditors.Controls.ButtonPredefines.Combo)});
			this.cmbHasta.Size = new System.Drawing.Size(80, 20);
			this.cmbHasta.TabIndex = 5;
			// 
			// chkAgrupar
			// 
			this.chkAgrupar.Location = new System.Drawing.Point(256, 8);
			this.chkAgrupar.Name = "chkAgrupar";
			// 
			// chkAgrupar.Properties
			// 
			this.chkAgrupar.Properties.Caption = "Agrupar";
			this.chkAgrupar.Size = new System.Drawing.Size(64, 18);
			this.chkAgrupar.TabIndex = 6;
			// 
			// cmbTipo
			// 
			this.cmbTipo.EditValue = "[EditValueull]";
			this.cmbTipo.Location = new System.Drawing.Point(336, 8);
			this.cmbTipo.Name = "cmbTipo";
			// 
			// cmbTipo.Properties
			// 
			this.cmbTipo.Properties.Buttons.AddRange(new DevExpress.XtraEditors.Controls.EditorButton[] {
																																																		new DevExpress.XtraEditors.Controls.EditorButton(DevExpress.XtraEditors.Controls.ButtonPredefines.Combo)});
			this.cmbTipo.Properties.DisplayMember = "Nombre";
			this.cmbTipo.Properties.NullText = "";
			this.cmbTipo.Properties.ValueMember = "idTipoFactura";
			this.cmbTipo.Properties.View = this.gridLookUpEdit1View;
			this.cmbTipo.Size = new System.Drawing.Size(112, 20);
			this.cmbTipo.TabIndex = 7;
			// 
			// gridLookUpEdit1View
			// 
			this.gridLookUpEdit1View.Columns.AddRange(new DevExpress.XtraGrid.Columns.GridColumn[] {
																																															 this.gridColumn1});
			this.gridLookUpEdit1View.FocusRectStyle = DevExpress.XtraGrid.Views.Grid.DrawFocusRectStyle.RowFocus;
			this.gridLookUpEdit1View.Name = "gridLookUpEdit1View";
			this.gridLookUpEdit1View.OptionsSelection.EnableAppearanceFocusedCell = false;
			this.gridLookUpEdit1View.OptionsView.ShowGroupPanel = false;
			// 
			// gridColumn1
			// 
			this.gridColumn1.Caption = "Nombre";
			this.gridColumn1.FieldName = "Nombre";
			this.gridColumn1.Name = "gridColumn1";
			this.gridColumn1.Visible = true;
			this.gridColumn1.VisibleIndex = 0;
			// 
			// FrmBusqueda
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 14);
			this.ClientSize = new System.Drawing.Size(720, 445);
			this.Controls.Add(this.tabControl1);
			this.Name = "FrmBusqueda";
			this.Text = "Busquedas";
			this.Load += new System.EventHandler(this.FrmBusqueda_Load);
			this.tabControl1.ResumeLayout(false);
			this.tbBasico.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.grdArticulo)).EndInit();
			this.tbComplemento.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.grdCliente)).EndInit();
			this.tbPromocion.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.grdProveedor)).EndInit();
			this.tbAnalisis.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.grdArticulos)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.grvArticulos)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbDesde.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbHasta.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.chkAgrupar.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbTipo.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.gridLookUpEdit1View)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void btArticulo_Click(object sender, System.EventArgs e)
		{
			string stTipo = "Articulo";
			string stCodigo = txtArticulo.Text;
			string stExec = string.Format("Exec Venta{0}Lista '{1}'", stTipo, stCodigo);
			grdArticulo.DataSource = Funcion.dvProcedimiento(cdsSeteo, stExec);

			//			MessageBox.Show("Encontrado");
		}

		private void cdsSeteo_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			this.cdsSeteo.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);
		}

		private void ultraGrid1_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
			e.Layout.Bands[0].Columns["idCliente"].Hidden = true;		
		}

		private void btProveedor_Click(object sender, System.EventArgs e)
		{
			string stTipo = "Proveedor";
			string stCodigo = txtArticulo.Text;
			string stExec = string.Format("Exec Venta{0}Lista '{1}'", stTipo, stCodigo);
			grdProveedor.DataSource = Funcion.dvProcedimiento(cdsSeteo, stExec);
		}

		private void btCliente_Click(object sender, System.EventArgs e)
		{
			string stTipo = "Cliente";
			string stCodigo = txtArticulo.Text;
			string stExec = string.Format("Exec Venta{0}Lista '{1}'", stTipo, stCodigo);
			grdCliente.DataSource = Funcion.dvProcedimiento(cdsSeteo, stExec);
		}

		private void grdArticulo_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
			e.Layout.Bands[0].Columns["idArticulo"].Hidden = true;		
		}

		private void grdProveedor_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
			e.Layout.Bands[0].Columns["idCliente"].Hidden = true;		
		}

		private void FormatoGrilla()
		{
			DevExpress.XtraGrid.Views.Base.ColumnView View = grdArticulos.MainView as DevExpress.XtraGrid.Views.Base.ColumnView;
			foreach (DevExpress.XtraGrid.Columns.GridColumn col in View.Columns)
			{
				col.OptionsColumn.AllowEdit = false;
				if (col.FieldName.StartsWith("id"))
					col.Visible = false;
				if (col.FieldName.StartsWith("Precio") || col.FieldName.StartsWith("Ingreso")
					|| col.FieldName.StartsWith("Egreso") || col.FieldName.StartsWith("Cant"))
				{
					col.DisplayFormat.FormatString = "n2";
					col.DisplayFormat.FormatType = DevExpress.Utils.FormatType.Numeric;
				}
				if (col.FieldName.StartsWith("Fecha") || col.FieldName.StartsWith("Desde")
					|| col.FieldName.StartsWith("Hasta"))
				{
					col.DisplayFormat.FormatString = "dd/MMM/yyyy";
					col.DisplayFormat.FormatType = DevExpress.Utils.FormatType.DateTime;
				}
			}
		}

		private void btKardex_Click(object sender, System.EventArgs e)
		{
			string stFiltro = grvArticulos.FilterPanelText;

			DateTime dtDesde = (DateTime)cmbDesde.EditValue;
			DateTime dtHasta = (DateTime)cmbHasta.EditValue;
			int iAgrupado = 0;
			if (chkAgrupar.Checked) iAgrupado = 1;
			int idTipoFactura = 0;
			if (cmbTipo.Text.Length > 0) idTipoFactura = (int) cmbTipo.EditValue;

			string stExec = string.Format("Exec ArticuloKardex 0, '{0}', '{1}', {2}, {3}", 
				dtDesde.ToString("yyyyMMdd"), dtHasta.ToString("yyyyMMdd"), iAgrupado, idTipoFactura);
			grvArticulos.Columns.Clear();
			grdArticulos.DataSource = Funcion.dvProcedimiento(cdsSeteo, stExec);

			FormatoGrilla();
			FormatoGrilla();
//			grvArticulos.FilterPanelText = stFiltro;
		}

		private void btArticuloB_Click(object sender, System.EventArgs e)
		{
			int iFila = grvArticulos.FocusedRowHandle;
			if (iFila < 0)
			{
				MessageBox.Show("Seleccione la línea que desa ver el Item", "Información");
				return;
			}
		
			Cursor = Cursors.WaitCursor;
			int idArticulo = (int) grvArticulos.GetRowCellValue(iFila, "idArticulo");
			if (idArticulo == 0)
			{
				MessageBox.Show("Fila no contiene Articulo Asociados",	"Información");
				Cursor = Cursors.Default;
				return;
			}
//			Articulo miArt = new Articulo(idArticulo);
//			miArt.MdiParent = MdiParent;
//			miArt.Show();
			Cursor = Cursors.Default;
		}

		private void btTransaccion_Click(object sender, System.EventArgs e)
		{
			int iFila = grvArticulos.FocusedRowHandle;
			if (iFila < 0)
			{
				MessageBox.Show("Seleccione la línea que desa ver la Transacción", "Información");
				return;
			}
		
			Cursor = Cursors.WaitCursor;
			int IdCompra = (int) grvArticulos.GetRowCellValue(iFila, "idCompra");
			if (IdCompra == 0)
			{
				MessageBox.Show("Fila no contiene Transacción Asociados",	"Información");
				Cursor = Cursors.Default;
				return;
			}
//			Compra miCompra = new Compra(-IdCompra);
//			miCompra.MdiParent = this.MdiParent;
//			miCompra.Show();
			Cursor = Cursors.Default;
		}

		private void FrmBusqueda_Load(object sender, System.EventArgs e)
		{
			cmbDesde.EditValue = DateTime.Today.AddDays(1-DateTime.Today.Day);
			cmbHasta.EditValue = DateTime.Today;
			cmbTipo.Properties.DataSource = Funcion.dvProcedimiento(cdsSeteo, 
				"Select Nombre, idTipoFactura From CompraNumero Order by Nombre");
		}
	}
}
