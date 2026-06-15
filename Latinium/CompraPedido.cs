using System;
using System.Globalization;
using System.Data.SqlClient;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de CompraPedido.
	/// </summary>
	public class CompraPedido : DevExpress.XtraEditors.XtraForm
	{
		private C1.Data.C1DataSet cdsCliente;
		private C1.Data.C1DataView cdvCliente;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbCliente;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource udsPendiente;
		private Infragistics.Win.UltraWinGrid.UltraGrid grPendiente;
		private Infragistics.Win.Misc.UltraButton btGenerar;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkCliente;
		private System.Windows.Forms.ErrorProvider errorProvider;
		private Infragistics.Win.Misc.UltraButton btMostrar;
		private Infragistics.Win.Misc.UltraButton btOtros;
		private Infragistics.Win.UltraWinEditors.UltraOptionSet optTipoF;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkNegativos;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public CompraPedido()
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
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Cliente");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Articulo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn3 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Codigo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn4 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Fecha");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn5 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Saldo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn6 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Pedido");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn7 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idArticulo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn8 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCompra");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn9 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCliente");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn10 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Sector");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn11 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idDetCompra");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn12 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Precio");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn13 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Numero");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn14 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idUnidad");
			Infragistics.Win.Appearance appearance1 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cliente");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn11 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Articulo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn12 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn13 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Fecha");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn14 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Saldo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn15 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Pedido");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn16 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idArticulo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn17 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCompra");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn18 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCliente");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn19 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Sector");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn20 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idDetCompra");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn21 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Precio");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn22 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Numero");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn23 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idUnidad");
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.ValueListItem valueListItem1 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem2 = new Infragistics.Win.ValueListItem();
			this.cdsCliente = new C1.Data.C1DataSet();
			this.cdvCliente = new C1.Data.C1DataView();
			this.cmbCliente = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.chkCliente = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.udsPendiente = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.grPendiente = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.btGenerar = new Infragistics.Win.Misc.UltraButton();
			this.errorProvider = new System.Windows.Forms.ErrorProvider();
			this.btMostrar = new Infragistics.Win.Misc.UltraButton();
			this.btOtros = new Infragistics.Win.Misc.UltraButton();
			this.optTipoF = new Infragistics.Win.UltraWinEditors.UltraOptionSet();
			this.chkNegativos = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			((System.ComponentModel.ISupportInitialize)(this.cdsCliente)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvCliente)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCliente)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.udsPendiente)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.grPendiente)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.optTipoF)).BeginInit();
			this.SuspendLayout();
			// 
			// cdsCliente
			// 
			this.cdsCliente.BindingContextCtrl = this;
			this.cdsCliente.DataLibrary = "LibFacturacion";
			this.cdsCliente.DataLibraryUrl = "";
			this.cdsCliente.DataSetDef = "dsClientes";
			this.cdsCliente.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsCliente.Name = "cdsCliente";
			this.cdsCliente.SchemaClassName = "LibFacturacion.DataClass";
			this.cdsCliente.SchemaDef = null;
			this.cdsCliente.BeforeFill += new C1.Data.FillEventHandler(this.cdsCliente_BeforeFill);
			// 
			// cdvCliente
			// 
			this.cdvCliente.BindingContextCtrl = this;
			this.cdvCliente.DataSet = this.cdsCliente;
			this.cdvCliente.Sort = "Nombre";
			this.cdvCliente.TableName = "";
			this.cdvCliente.TableViewName = "Cliente";
			// 
			// cmbCliente
			// 
			this.cmbCliente.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbCliente.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbCliente.DataSource = this.cdvCliente;
			ultraGridColumn1.Header.VisiblePosition = 1;
			ultraGridColumn2.Header.VisiblePosition = 3;
			ultraGridColumn2.Hidden = true;
			ultraGridColumn3.Header.VisiblePosition = 0;
			ultraGridColumn3.Width = 219;
			ultraGridColumn4.Header.VisiblePosition = 4;
			ultraGridColumn4.Hidden = true;
			ultraGridColumn5.Header.VisiblePosition = 5;
			ultraGridColumn5.Hidden = true;
			ultraGridColumn6.Header.VisiblePosition = 2;
			ultraGridColumn6.Hidden = true;
			ultraGridColumn7.Header.VisiblePosition = 6;
			ultraGridColumn7.Hidden = true;
			ultraGridColumn8.Header.VisiblePosition = 7;
			ultraGridColumn8.Hidden = true;
			ultraGridColumn9.Header.VisiblePosition = 8;
			ultraGridColumn9.Hidden = true;
			ultraGridBand1.Columns.AddRange(new object[] {
																										 ultraGridColumn1,
																										 ultraGridColumn2,
																										 ultraGridColumn3,
																										 ultraGridColumn4,
																										 ultraGridColumn5,
																										 ultraGridColumn6,
																										 ultraGridColumn7,
																										 ultraGridColumn8,
																										 ultraGridColumn9});
			this.cmbCliente.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.cmbCliente.DisplayMember = "Nombre";
			this.cmbCliente.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbCliente.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbCliente.Enabled = false;
			this.cmbCliente.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbCliente.Location = new System.Drawing.Point(128, 16);
			this.cmbCliente.Name = "cmbCliente";
			this.cmbCliente.Size = new System.Drawing.Size(200, 21);
			this.cmbCliente.TabIndex = 0;
			this.cmbCliente.ValueMember = "idCliente";
			// 
			// chkCliente
			// 
			this.chkCliente.Checked = true;
			this.chkCliente.CheckState = System.Windows.Forms.CheckState.Checked;
			this.chkCliente.Location = new System.Drawing.Point(32, 16);
			this.chkCliente.Name = "chkCliente";
			this.chkCliente.Size = new System.Drawing.Size(88, 20);
			this.chkCliente.TabIndex = 1;
			this.chkCliente.Text = "Todo Cliente";
			this.chkCliente.CheckedChanged += new System.EventHandler(this.chkCliente_CheckedChanged);
			// 
			// udsPendiente
			// 
			ultraDataColumn4.DataType = typeof(System.DateTime);
			ultraDataColumn5.DataType = typeof(System.Double);
			ultraDataColumn6.DataType = typeof(System.Double);
			ultraDataColumn7.DataType = typeof(int);
			ultraDataColumn8.DataType = typeof(int);
			ultraDataColumn9.DataType = typeof(int);
			ultraDataColumn11.DataType = typeof(int);
			ultraDataColumn12.DataType = typeof(System.Double);
			ultraDataColumn14.DataType = typeof(System.Double);
			this.udsPendiente.Band.Columns.AddRange(new object[] {
																														 ultraDataColumn1,
																														 ultraDataColumn2,
																														 ultraDataColumn3,
																														 ultraDataColumn4,
																														 ultraDataColumn5,
																														 ultraDataColumn6,
																														 ultraDataColumn7,
																														 ultraDataColumn8,
																														 ultraDataColumn9,
																														 ultraDataColumn10,
																														 ultraDataColumn11,
																														 ultraDataColumn12,
																														 ultraDataColumn13,
																														 ultraDataColumn14});
			// 
			// grPendiente
			// 
			this.grPendiente.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.grPendiente.Cursor = System.Windows.Forms.Cursors.Default;
			this.grPendiente.DataSource = this.udsPendiente;
			appearance1.BackColor = System.Drawing.Color.White;
			this.grPendiente.DisplayLayout.Appearance = appearance1;
			this.grPendiente.DisplayLayout.AutoFitColumns = true;
			ultraGridColumn10.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
			ultraGridColumn10.Header.VisiblePosition = 0;
			ultraGridColumn10.Width = 55;
			ultraGridColumn11.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
			ultraGridColumn11.Header.VisiblePosition = 1;
			ultraGridColumn11.Width = 72;
			ultraGridColumn12.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
			ultraGridColumn12.Header.VisiblePosition = 2;
			ultraGridColumn12.Width = 50;
			ultraGridColumn13.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
			ultraGridColumn13.Format = "dd/MMM/yyyy";
			ultraGridColumn13.Header.VisiblePosition = 3;
			ultraGridColumn13.Width = 53;
			ultraGridColumn14.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
			ultraGridColumn14.Header.VisiblePosition = 4;
			ultraGridColumn14.Width = 45;
			ultraGridColumn15.Header.VisiblePosition = 5;
			ultraGridColumn15.Width = 50;
			ultraGridColumn16.Header.VisiblePosition = 6;
			ultraGridColumn16.Hidden = true;
			ultraGridColumn17.Header.VisiblePosition = 7;
			ultraGridColumn17.Hidden = true;
			ultraGridColumn18.Header.VisiblePosition = 8;
			ultraGridColumn18.Hidden = true;
			ultraGridColumn19.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
			ultraGridColumn19.Header.VisiblePosition = 9;
			ultraGridColumn19.Width = 60;
			ultraGridColumn20.Header.VisiblePosition = 10;
			ultraGridColumn20.Hidden = true;
			ultraGridColumn20.Width = 55;
			ultraGridColumn21.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
			ultraGridColumn21.Header.VisiblePosition = 11;
			ultraGridColumn21.Width = 60;
			ultraGridColumn22.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
			ultraGridColumn22.Header.VisiblePosition = 12;
			ultraGridColumn22.Width = 86;
			ultraGridColumn23.Header.Caption = "Unidad";
			ultraGridColumn23.Header.VisiblePosition = 13;
			ultraGridColumn23.Hidden = true;
			ultraGridColumn23.Width = 79;
			ultraGridBand2.Columns.AddRange(new object[] {
																										 ultraGridColumn10,
																										 ultraGridColumn11,
																										 ultraGridColumn12,
																										 ultraGridColumn13,
																										 ultraGridColumn14,
																										 ultraGridColumn15,
																										 ultraGridColumn16,
																										 ultraGridColumn17,
																										 ultraGridColumn18,
																										 ultraGridColumn19,
																										 ultraGridColumn20,
																										 ultraGridColumn21,
																										 ultraGridColumn22,
																										 ultraGridColumn23});
			this.grPendiente.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			this.grPendiente.DisplayLayout.GroupByBox.Prompt = "Arrastre la columna a agrupar";
			appearance2.BackColor = System.Drawing.Color.Transparent;
			this.grPendiente.DisplayLayout.Override.CardAreaAppearance = appearance2;
			appearance3.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance3.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance3.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance3.FontData.BoldAsString = "True";
			appearance3.FontData.Name = "Arial";
			appearance3.FontData.SizeInPoints = 10F;
			appearance3.ForeColor = System.Drawing.Color.White;
			appearance3.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.grPendiente.DisplayLayout.Override.HeaderAppearance = appearance3;
			appearance4.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance4.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance4.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.grPendiente.DisplayLayout.Override.RowSelectorAppearance = appearance4;
			appearance5.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance5.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance5.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.grPendiente.DisplayLayout.Override.SelectedRowAppearance = appearance5;
			this.grPendiente.DisplayLayout.ViewStyleBand = Infragistics.Win.UltraWinGrid.ViewStyleBand.OutlookGroupBy;
			this.grPendiente.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.grPendiente.Location = new System.Drawing.Point(16, 48);
			this.grPendiente.Name = "grPendiente";
			this.grPendiente.Size = new System.Drawing.Size(552, 232);
			this.grPendiente.TabIndex = 2;
			// 
			// btGenerar
			// 
			this.btGenerar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btGenerar.Location = new System.Drawing.Point(112, 288);
			this.btGenerar.Name = "btGenerar";
			this.btGenerar.TabIndex = 3;
			this.btGenerar.Text = "&Generar";
			this.btGenerar.Click += new System.EventHandler(this.btGenerar_Click);
			// 
			// errorProvider
			// 
			this.errorProvider.ContainerControl = this;
			// 
			// btMostrar
			// 
			this.btMostrar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btMostrar.Location = new System.Drawing.Point(16, 288);
			this.btMostrar.Name = "btMostrar";
			this.btMostrar.TabIndex = 4;
			this.btMostrar.Text = "&Mostrar";
			this.btMostrar.Click += new System.EventHandler(this.btMostrar_Click);
			// 
			// btOtros
			// 
			this.btOtros.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.btOtros.Location = new System.Drawing.Point(488, 288);
			this.btOtros.Name = "btOtros";
			this.btOtros.Size = new System.Drawing.Size(80, 23);
			this.btOtros.TabIndex = 5;
			this.btOtros.Text = "&Reportes";
			this.btOtros.Click += new System.EventHandler(this.btOtros_Click);
			this.btOtros.MouseUp += new System.Windows.Forms.MouseEventHandler(this.btOtros_MouseUp);
			// 
			// optTipoF
			// 
			this.optTipoF.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.optTipoF.CheckedIndex = 0;
			this.optTipoF.ItemAppearance = appearance6;
			valueListItem1.DataValue = 1;
			valueListItem1.DisplayText = "Factura";
			valueListItem2.DataValue = 8;
			valueListItem2.DisplayText = "Egreso Bodega";
			this.optTipoF.Items.Add(valueListItem1);
			this.optTipoF.Items.Add(valueListItem2);
			this.optTipoF.Location = new System.Drawing.Point(200, 288);
			this.optTipoF.Name = "optTipoF";
			this.optTipoF.Size = new System.Drawing.Size(112, 32);
			this.optTipoF.TabIndex = 8;
			this.optTipoF.Text = "Factura";
			// 
			// chkNegativos
			// 
			this.chkNegativos.Location = new System.Drawing.Point(360, 16);
			this.chkNegativos.Name = "chkNegativos";
			this.chkNegativos.TabIndex = 9;
			this.chkNegativos.Text = "Mostrar Negativos";
			// 
			// CompraPedido
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(584, 326);
			this.Controls.Add(this.chkNegativos);
			this.Controls.Add(this.optTipoF);
			this.Controls.Add(this.btOtros);
			this.Controls.Add(this.btMostrar);
			this.Controls.Add(this.btGenerar);
			this.Controls.Add(this.grPendiente);
			this.Controls.Add(this.chkCliente);
			this.Controls.Add(this.cmbCliente);
			this.Name = "CompraPedido";
			this.Text = "Pedido Pendientes";
			this.Closing += new System.ComponentModel.CancelEventHandler(this.CompraPedido_Closing);
			this.Load += new System.EventHandler(this.CompraPedido_Load);
			((System.ComponentModel.ISupportInitialize)(this.cdsCliente)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvCliente)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCliente)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.udsPendiente)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.grPendiente)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.optTipoF)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private int iNumLineas = 999;
		private void CompraPedido_Load(object sender, System.EventArgs e)
		{
			string stAudita = "Exec AuditaCrear 53, 6, 'Pedido Compra'";
			Funcion.EjecutaSQL(cdsCliente, stAudita, true);
			string stSelect = "Select Top 1 IsNull(NumeroLineas, 0) From SeteoF";
			iNumLineas = Funcion.iEscalarSQL(cdsCliente, stSelect, true);
		}

		private void btGenerar_Click(object sender, System.EventArgs e)
		{
			errorProvider.SetError(cmbCliente, "");
			if (cmbCliente.Text.Trim().Length == 0)
			{
				errorProvider.SetError(cmbCliente, "Solo se puede genera Facturas de un Cliente");
				return;
			}
			int IdTipoF = (int) optTipoF.Value;
			string stTipoF = optTipoF.Text;
			if (DialogResult.No == MessageBox.Show("¿Desea Generar "
				+ stTipoF + " de los pedidos seleccionados?", "Confirmación",
				MessageBoxButtons.YesNo, MessageBoxIcon.Question)) return;

			int IdCliente = int.Parse(cmbCliente.Value.ToString());
			string stNumero = Funcion.NumeraFactura(cdsCliente, IdTipoF);
			string stSelect = "Select idVendedor From Cliente Where IdCliente = " +
				IdCliente.ToString();
			int IdVendedor = Funcion.iEscalarSQL(cdsCliente, stSelect, true);
			stSelect = "Select DiasCredito From Cliente Where IdCliente = " +
				IdCliente.ToString();
			int iDiasCredito = Funcion.iEscalarSQL(cdsCliente, stSelect, true);
			stSelect = "Insert Into Compra (idCliente, Fecha, Numero, idTipoFactura, idVendedor, "
				+ "ContadoCredito, DiasCredito) " 
				+ " Values (" + IdCliente.ToString() + ", '" + DateTime.Today.ToString("yyyyMMdd")
				+ "', '" + stNumero + "', " + IdTipoF.ToString()  + ", "
				+ IdVendedor.ToString() + ", " + iDiasCredito.ToString() + ", 2)";
			Funcion.EjecutaSQL(cdsCliente, stSelect, true);

			stSelect = "Select Top 1 idCompra FROM Compra Order by IdCompra DESC";
			int IdCompra = Funcion.iEscalarSQL(cdsCliente, stSelect, true);

			int i = 1;
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in grPendiente.Rows.All)
			{
				if (i > iNumLineas) continue;
				else i++;
				int IdDetCompra = (int)	dr.Cells["idDetCompra"].Value;
				int IdArticulo = (int) dr.Cells["idArticulo"].Value;
				double dCantidad = (double) dr.Cells["Pedido"].Value;
				double dSaldo = (double) dr.Cells["Saldo"].Value;
				double dPrecio = (double) dr.Cells["Precio"].Value;
				string stNumPedido = (string) dr.Cells["Numero"].Value;
				string stIdUnidad = dr.Cells["IdUnidad"].Value.ToString();
				if (dCantidad > dSaldo + .000001 && !chkNegativos.Checked)
				{
					string stArt = (string) dr.Cells["Codigo"].Value;
					stArt += " / " + (string) dr.Cells["Articulo"].Value;
					MessageBox.Show("No puede enviar mas que el Saldo" +
						" en Articulo: " + stArt,
						"Información");
					dCantidad = dSaldo;
				}
				if (!chkNegativos.Checked)
					if (dCantidad == 0) continue;

				stSelect = "Update DetCompra SET Pedido = Pedido - " 
					+ dCantidad.ToString("0.00", us) + ", RefNumero = '"
					+ stNumero + "' Where idDetCompra = "
					+ IdDetCompra.ToString();
				Funcion.EjecutaSQL(cdsCliente, stSelect);
				string stCodigo = "Select Codigo From Articulo Where idArticulo = "
					+ IdArticulo.ToString();
				stCodigo = Funcion.sEscalarSQL(cdsCliente, stCodigo);
				string stArticulo = "Select Articulo.Articulo From Articulo Where idArticulo = "
					+ IdArticulo.ToString();
				stArticulo = Funcion.sEscalarSQL(cdsCliente, stArticulo);
				stSelect = "Insert Into DetCompra (idCompra, Cantidad, Unidades, idUnidad, Precio, " 
					+ "idArticulo, RefNumero, Signo, Codigo, Articulo, Impuesto) Values(" + IdCompra.ToString() + ", "
					+ dCantidad.ToString("0.00", us) + ", " + dCantidad.ToString("0.00", us)
					+ ", " + stIdUnidad 
					+ ", " + dPrecio.ToString("0.000000", us) + ", " + IdArticulo.ToString()
					+ ", '" + stNumPedido + "', -1, '" + stCodigo + "', '" + stArticulo + "', 12)";
				Funcion.EjecutaSQL(cdsCliente, stSelect, true);
			}
			stSelect = "Exec FacturaTotal " + IdCompra.ToString();
			Funcion.EjecutaSQL(cdsCliente, stSelect, true);
			MessageBox.Show(stTipoF + " Generada", "Información");
		}

		CultureInfo us = new CultureInfo("en-US");
		private void cdsCliente_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			this.cdsCliente.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);
		}

		private void chkCliente_CheckedChanged(object sender, System.EventArgs e)
		{
			cmbCliente.Enabled = !chkCliente.Checked;
			if (!chkCliente.Checked) cmbCliente.Select();
		}

		private void btMostrar_Click(object sender, System.EventArgs e)
		{
			errorProvider.SetError(cmbCliente, "");
			if (!chkCliente.Checked)
				if (cmbCliente.Text.Trim().Length == 0)
				{
					errorProvider.SetError(cmbCliente, "Ingrese Cliente");
					return;
				}
		
			string stSelect = "Select Cliente.Nombre, Articulo.Articulo, " +
				"Articulo.Codigo, Compra.fecha, Articulo.Saldo, " +
				"DetCompra.Pedido, ClienteSector.Sector, DetCompra.idDetCompra, " +
				"Articulo.idArticulo, Compra.Numero, DetCompra.Precio, " +
				"ClS1.Sector, DetCompra.idUnidad From DetCompra " +
				"Inner join compra on compra.idcompra = detcompra.idcompra " +
				"inner join cliente on compra.idcliente = cliente.idcliente " +
				"Left Outer join clienteSector on ClienteSector.idSector = cliente.idSector " +
				"inner join articulo on articulo.idarticulo = detcompra.idarticulo " +
				"Left Outer join Cliente cl1 on cl1.idCliente = compra.idRecibe " +
				"Left Outer join ClienteSector ClS1 on cl1.idSector = ClS1.idSector " +
				"Where DetCompra.Pedido > 0 And Compra.idTipoFactura = 13 ";
			if (!chkNegativos.Checked) stSelect +=	"And Articulo.Saldo > 0 ";

			if (!chkCliente.Checked) 
				stSelect += " And Compra.idCliente = " + cmbCliente.Value.ToString();

			SqlDataReader dr = Funcion.rEscalarSQL(cdsCliente, stSelect, true);

			udsPendiente.Rows.Clear();
			int i= 0;
			while(dr.Read())
			{
				string stCliente = "";
				if (dr.GetValue(0) != System.DBNull.Value)
					stCliente = dr.GetString(0);

				string stArticulo = dr.GetString(1);
				string stCodigo = dr.GetString(2);
				DateTime dtFecha = dr.GetDateTime(3);
				double dSaldo = dr.GetDouble(4);
				double dPedido = dr.GetDouble(5);
				string stSector = "";
				if (dr.GetValue(6) != System.DBNull.Value)	stSector = dr.GetString(6);
				int IdDetCompra = dr.GetInt32(7);
				int IdArticulo = dr.GetInt32(8);
				string stNumero = "";
				if (dr.GetValue(9) != System.DBNull.Value)	stNumero = dr.GetString(9);
				double dPrecio = dr.GetDouble(10);
				string stSector1 = "";
				if (dr.GetValue(11) != System.DBNull.Value)	stSector1 = dr.GetString(11);
				int IdUnidad = dr.GetInt32(12);

				udsPendiente.Rows.Add(); // coleas@manpoweecuador.com 2567705 503747 ext 125 carmen oleas
				udsPendiente.Rows[i]["Cliente"] = stCliente;
				udsPendiente.Rows[i]["Articulo"] = stArticulo;
				udsPendiente.Rows[i]["Codigo"] = stCodigo;
				udsPendiente.Rows[i]["Fecha"] = dtFecha;
				udsPendiente.Rows[i]["Saldo"] = dSaldo;
				udsPendiente.Rows[i]["Pedido"] = dPedido;
				if (stSector1.Length > 0)	udsPendiente.Rows[i]["Sector"] = stSector1;
				else udsPendiente.Rows[i]["Sector"] = stSector;
				udsPendiente.Rows[i]["idDetCompra"] = IdDetCompra;
				udsPendiente.Rows[i]["idArticulo"] = IdArticulo;
				udsPendiente.Rows[i]["Numero"] = stNumero;
				udsPendiente.Rows[i]["Precio"] = dPrecio;
				udsPendiente.Rows[i]["idUnidad"] = IdUnidad;
				i++;
			}
			dr.Close();
			grPendiente.Refresh();
		}

		private void btOtros_Click(object sender, System.EventArgs e)
		{
			Cursor = Cursors.WaitCursor;
			string stFiltro = "{Compra.idTipoFactura} = 13 ";
			stFiltro += " And {DetCompra.Pedido} > 0";
			Reporte miRep = new Reporte("PedidoPendiente.Rpt", stFiltro);
			miRep.MdiParent = this.MdiParent;
			miRep.Show();
			Cursor = Cursors.Default;
		}

		private void btOtros_MouseUp(object sender, System.Windows.Forms.MouseEventArgs e)
		{
			if (e.Button != MouseButtons.Right) return;
			Cursor = Cursors.WaitCursor;
			string stFiltro = "{Compra.idTipoFactura} = 13 ";
			stFiltro += " And {DetCompra.Pedido} > 0";
			Reporte miRep = new Reporte("PedidoArticulo.Rpt", stFiltro);
			miRep.MdiParent = this.MdiParent;
			miRep.Show();
			Cursor = Cursors.Default;
		}

		private void CompraPedido_Closing(object sender, System.ComponentModel.CancelEventArgs e)
		{
			string stAudita = "Exec AuditaCrear 53, 7, 'Pedido Compra'";
			Funcion.EjecutaSQL(cdsCliente, stAudita, true);
		}

	}
}
