using System;
using System.Globalization;
using System.Data.SqlClient;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;
using CrystalDecisions.Shared;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de CompraPedido.
	/// </summary>
	public class CompraPedido2 : DevExpress.XtraEditors.XtraForm
	{
		private C1.Data.C1DataSet cdsCliente;
		private C1.Data.C1DataView cdvCliente;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbCliente;
		private Infragistics.Win.UltraWinGrid.UltraGrid grPendiente;
		private Infragistics.Win.Misc.UltraButton btGenerar;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkCliente;
		private System.Windows.Forms.ErrorProvider errorProvider;
		private Infragistics.Win.Misc.UltraButton btMostrar;
		private Infragistics.Win.Misc.UltraButton btOtros;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkNegativos;
		private Infragistics.Win.UltraWinEditors.UltraComboEditor cmbFactura;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtPedido;
		private System.Windows.Forms.ToolTip toolTip1;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkPedidos;
		private C1.Data.C1DataSet cdsReporte;
		private C1.Data.C1DataView cdvReporte;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbReporte;
		private System.ComponentModel.IContainer components;

		public CompraPedido2()
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
			Infragistics.Win.Appearance appearance1 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.ValueListItem valueListItem1 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem2 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Reporte", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Detalle");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn11 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idGrupo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn12 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idReporte");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn13 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Reporte");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn14 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Tipo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn15 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Filtro");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn16 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("NoFiltro");
			this.cdsCliente = new C1.Data.C1DataSet();
			this.cdvCliente = new C1.Data.C1DataView();
			this.cmbCliente = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.chkCliente = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.grPendiente = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.btGenerar = new Infragistics.Win.Misc.UltraButton();
			this.errorProvider = new System.Windows.Forms.ErrorProvider();
			this.btMostrar = new Infragistics.Win.Misc.UltraButton();
			this.btOtros = new Infragistics.Win.Misc.UltraButton();
			this.chkNegativos = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.cmbFactura = new Infragistics.Win.UltraWinEditors.UltraComboEditor();
			this.chkPedidos = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.txtPedido = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.toolTip1 = new System.Windows.Forms.ToolTip(this.components);
			this.cmbReporte = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.cdvReporte = new C1.Data.C1DataView();
			this.cdsReporte = new C1.Data.C1DataSet();
			((System.ComponentModel.ISupportInitialize)(this.cdsCliente)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvCliente)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCliente)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.grPendiente)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbFactura)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtPedido)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbReporte)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvReporte)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsReporte)).BeginInit();
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
			// grPendiente
			// 
			this.grPendiente.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.grPendiente.Cursor = System.Windows.Forms.Cursors.Default;
			appearance1.BackColor = System.Drawing.Color.White;
			this.grPendiente.DisplayLayout.Appearance = appearance1;
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
			this.grPendiente.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.grPendiente.Location = new System.Drawing.Point(16, 48);
			this.grPendiente.Name = "grPendiente";
			this.grPendiente.Size = new System.Drawing.Size(792, 232);
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
			this.btMostrar.Location = new System.Drawing.Point(32, 288);
			this.btMostrar.Name = "btMostrar";
			this.btMostrar.TabIndex = 4;
			this.btMostrar.Text = "&Mostrar";
			this.btMostrar.Click += new System.EventHandler(this.btMostrar_Click);
			// 
			// btOtros
			// 
			this.btOtros.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.btOtros.Location = new System.Drawing.Point(728, 288);
			this.btOtros.Name = "btOtros";
			this.btOtros.Size = new System.Drawing.Size(80, 23);
			this.btOtros.TabIndex = 5;
			this.btOtros.Text = "&Reportes";
			this.btOtros.Click += new System.EventHandler(this.btOtros_Click);
			this.btOtros.MouseUp += new System.Windows.Forms.MouseEventHandler(this.btOtros_MouseUp);
			// 
			// chkNegativos
			// 
			this.chkNegativos.Checked = true;
			this.chkNegativos.CheckState = System.Windows.Forms.CheckState.Checked;
			this.chkNegativos.Location = new System.Drawing.Point(360, 16);
			this.chkNegativos.Name = "chkNegativos";
			this.chkNegativos.TabIndex = 9;
			this.chkNegativos.Text = "Mostrar Negativos";
			// 
			// cmbFactura
			// 
			this.cmbFactura.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			valueListItem1.DataValue = 1;
			valueListItem1.DisplayText = "Venta";
			valueListItem2.DataValue = 4;
			valueListItem2.DisplayText = "Compra";
			this.cmbFactura.Items.Add(valueListItem1);
			this.cmbFactura.Items.Add(valueListItem2);
			this.cmbFactura.Location = new System.Drawing.Point(208, 288);
			this.cmbFactura.Name = "cmbFactura";
			this.cmbFactura.Size = new System.Drawing.Size(112, 21);
			this.cmbFactura.TabIndex = 10;
			this.cmbFactura.Text = "Venta";
			// 
			// chkPedidos
			// 
			this.chkPedidos.Checked = true;
			this.chkPedidos.CheckState = System.Windows.Forms.CheckState.Checked;
			this.chkPedidos.Location = new System.Drawing.Point(488, 16);
			this.chkPedidos.Name = "chkPedidos";
			this.chkPedidos.Size = new System.Drawing.Size(104, 20);
			this.chkPedidos.TabIndex = 11;
			this.chkPedidos.Text = "Pedidos Activos";
			// 
			// txtPedido
			// 
			this.txtPedido.Location = new System.Drawing.Point(608, 16);
			this.txtPedido.Name = "txtPedido";
			this.txtPedido.Size = new System.Drawing.Size(80, 21);
			this.txtPedido.TabIndex = 12;
			this.toolTip1.SetToolTip(this.txtPedido, "Codigo del Pedido a Filtrar ");
			// 
			// cmbReporte
			// 
			this.cmbReporte.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.cmbReporte.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbReporte.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbReporte.DataSource = this.cdvReporte;
			this.cmbReporte.DisplayLayout.AutoFitColumns = true;
			ultraGridColumn10.Header.VisiblePosition = 0;
			ultraGridColumn10.Width = 118;
			ultraGridColumn11.Header.VisiblePosition = 1;
			ultraGridColumn11.Hidden = true;
			ultraGridColumn12.Header.VisiblePosition = 2;
			ultraGridColumn12.Hidden = true;
			ultraGridColumn13.Header.VisiblePosition = 3;
			ultraGridColumn13.Hidden = true;
			ultraGridColumn14.Header.VisiblePosition = 4;
			ultraGridColumn14.Hidden = true;
			ultraGridColumn15.Header.VisiblePosition = 5;
			ultraGridColumn15.Hidden = true;
			ultraGridColumn15.Width = 54;
			ultraGridColumn16.Header.VisiblePosition = 6;
			ultraGridColumn16.Hidden = true;
			ultraGridColumn16.Width = 54;
			ultraGridBand2.Columns.AddRange(new object[] {
																										 ultraGridColumn10,
																										 ultraGridColumn11,
																										 ultraGridColumn12,
																										 ultraGridColumn13,
																										 ultraGridColumn14,
																										 ultraGridColumn15,
																										 ultraGridColumn16});
			this.cmbReporte.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			this.cmbReporte.DisplayMember = "Reporte";
			this.cmbReporte.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbReporte.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbReporte.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbReporte.Location = new System.Drawing.Point(568, 290);
			this.cmbReporte.Name = "cmbReporte";
			this.cmbReporte.Size = new System.Drawing.Size(136, 21);
			this.cmbReporte.TabIndex = 45;
			this.toolTip1.SetToolTip(this.cmbReporte, "Grupo 150 Tipo Avanzado");
			this.cmbReporte.ValueMember = "idReporte";
			// 
			// cdvReporte
			// 
			this.cdvReporte.BindingContextCtrl = this;
			this.cdvReporte.DataSet = this.cdsReporte;
			this.cdvReporte.RowFilter = "Tipo = 10 And idGrupo = 150";
			this.cdvReporte.TableName = "";
			this.cdvReporte.TableViewName = "Reporte";
			// 
			// cdsReporte
			// 
			this.cdsReporte.BindingContextCtrl = this;
			this.cdsReporte.DataLibrary = "LibEmpresa";
			this.cdsReporte.DataLibraryUrl = "";
			this.cdsReporte.DataSetDef = "dsReporte";
			this.cdsReporte.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsReporte.Name = "cdsReporte";
			this.cdsReporte.SchemaClassName = "LibEmpresa.DataClass";
			this.cdsReporte.SchemaDef = null;
			// 
			// CompraPedido2
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(824, 325);
			this.Controls.Add(this.cmbReporte);
			this.Controls.Add(this.txtPedido);
			this.Controls.Add(this.chkPedidos);
			this.Controls.Add(this.cmbFactura);
			this.Controls.Add(this.chkNegativos);
			this.Controls.Add(this.btOtros);
			this.Controls.Add(this.btMostrar);
			this.Controls.Add(this.btGenerar);
			this.Controls.Add(this.grPendiente);
			this.Controls.Add(this.chkCliente);
			this.Controls.Add(this.cmbCliente);
			this.Name = "CompraPedido2";
			this.Text = "Pedido Pendientes";
			this.Closing += new System.ComponentModel.CancelEventHandler(this.CompraPedido_Closing);
			this.Load += new System.EventHandler(this.CompraPedido_Load);
			((System.ComponentModel.ISupportInitialize)(this.cdsCliente)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvCliente)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCliente)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.grPendiente)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbFactura)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtPedido)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbReporte)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvReporte)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsReporte)).EndInit();
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
				errorProvider.SetError(cmbCliente, "Solo se puede generar Facturas de un Cliente");
				return;
			}
			int IdTipoF = (int) cmbFactura.Value;
			string stTipoF = cmbFactura.Text;
			if (DialogResult.No == MessageBox.Show("¿Desea Generar "
				+ stTipoF + " de los pedidos seleccionados?", "Confirmación",
				MessageBoxButtons.YesNo, MessageBoxIcon.Question)) return;
			Cursor = Cursors.WaitCursor;
			int IdCliente = (int) cmbCliente.Value;
			int iExiste = Funcion.iEscalarSQL(cdsCliente, 
				string.Format("Exec NumeracionVerifica {0}, {1}, 1, 0", 
					IdTipoF, IdCliente), true);
			if (iExiste > 0)
			{
				MessageBox.Show("Numero ya existe", "Información");
				return;
			}

			string stSelect = "Exec PedidoPendienteFactura " +
				IdCliente.ToString() + ", " + IdTipoF.ToString() + 
				", '" +  txtPedido.Text.Trim() + "'";

			int IdCompra = Funcion.iEscalarSQL(cdsCliente, stSelect, true);
			#region graba el pedido
			string PendienteTipo = "PendienteCliente";
			if (IdTipoF == 4) PendienteTipo = "PendienteProveedor";
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in grPendiente.Rows.All)
			{
				if (dr.Cells[PendienteTipo].Value != System.DBNull.Value)
				{
					int IdDetCompra = (int) dr.Cells["idDetCompra"].Value;
					double dCantidad = (double) dr.Cells[PendienteTipo].Value;
					string stExecP = "Exec PedidoActualizaValor " + IdDetCompra.ToString()
						+ ", " + dCantidad.ToString("0.00", us) + ", " + IdTipoF.ToString();
					Funcion.EjecutaSQL(cdsCliente, stExecP, true);
				}
			}
			#endregion graba el pedido

			stSelect = "Exec PedidoPendienteDetalleF " + IdCompra.ToString() 
				+ ", " + Filtro();
			Funcion.EjecutaSQL(cdsCliente, stSelect, true);

			stSelect = "Exec FacturaTotal " + IdCompra.ToString();
			Funcion.EjecutaSQL(cdsCliente, stSelect, true);


			MessageBox.Show(stTipoF + " Generada", "Información");
//			Compra miCompra = new Compra(IdTipoF, IdCompra);
//			miCompra.MdiParent = this.MdiParent;
//			miCompra.Show();
			Cursor = Cursors.Default;
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
		DataSet miDataSet = new DataSet();
		DataView dv = new DataView();

		private string Filtro()
		{
			string stSelect = "";
			if (!chkCliente.Checked) stSelect += cmbCliente.Value.ToString();
			else stSelect += "0";
			if (!chkNegativos.Checked) stSelect +=	", 0";
			else stSelect +=	", 1";
			if (!chkPedidos.Checked) stSelect +=	", 0";
			else stSelect +=	", 1";
			stSelect += ", '" +  txtPedido.Text.Trim() + "'";
			return stSelect;
		}

		private bool VerificaCliente()
		{
			bool bOk = true;
			errorProvider.SetError(cmbCliente, "");
			if (!chkCliente.Checked)
				if (cmbCliente.Text.Trim().Length == 0)
				{
					errorProvider.SetError(cmbCliente, "Ingrese Cliente");
					bOk = false;
				}
			return bOk;
		}

		private void btMostrar_Click(object sender, System.EventArgs e)
		{
			if (!VerificaCliente()) return;
		
			string stSelect = "Exec PedidoPendienteMostrar " + Filtro();

			Funcion miFuncion = new Funcion();
			string strConn = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);
			SqlDataAdapter myData = new SqlDataAdapter(stSelect, strConn);

			myData.TableMappings.Add("Tabla", "ExcelSube");
			
			miDataSet.Reset();
			miDataSet.Clear();
			try
			{
				myData.Fill(miDataSet);
			}
			catch(Exception ex)
			{
				MessageBox.Show(ex.Message, "Error al abrir la Hoja A del Archivo: " );
				return;
			}
			dv = miDataSet.Tables[0].DefaultView;

			grPendiente.DataSource = dv; // Funcion.LlenaDv(stSelect);
			for (int i=0; i< grPendiente.DisplayLayout.Bands[0].Columns.Count; i++)
			{
				if (grPendiente.DisplayLayout.Bands[0].Columns[i].Key.StartsWith("id"))
					grPendiente.DisplayLayout.Bands[0].Columns[i].Hidden = true;
				else
				{
					grPendiente.DisplayLayout.Bands[0].Columns[i].CellActivation=
						Infragistics.Win.UltraWinGrid.Activation.NoEdit;
					if (grPendiente.DisplayLayout.Bands[0].Columns[i].DataType.ToString()=="System.DateTime")
						grPendiente.DisplayLayout.Bands[0].Columns[i].Format="dd/MMM/yyyy";
				}
			}
			grPendiente.DisplayLayout.Bands[0].Columns["PendienteCliente"].CellActivation=
				Infragistics.Win.UltraWinGrid.Activation.AllowEdit;
			grPendiente.DisplayLayout.Bands[0].Columns["PendienteProveedor"].CellActivation=
				Infragistics.Win.UltraWinGrid.Activation.AllowEdit;
			grPendiente.DataBind();
		}

		private void btOtros_Click(object sender, System.EventArgs e)
		{
			if (!VerificaCliente()) return;
			errorProvider.SetError(cmbReporte, "");
			if (cmbReporte.Text.Length == 0)
			{
				errorProvider.SetError(cmbReporte, "Ingrese Reporte");
				return;
			}
			Cursor = Cursors.WaitCursor;
//			string stTitulo = "Desde " + dtDesde.ToString("dd/MMM/yyyy")
//				+ " hasta " + dtHasta.ToString("dd/MMM/yyyy");
			ParameterFields paramFields = new ParameterFields ();

			#region Parametro Cliente
			ParameterField paramField1 = new ParameterField ();
			ParameterDiscreteValue discreteVal1 = new ParameterDiscreteValue ();
			paramField1.ParameterFieldName = "@idCliente";
			if (!chkCliente.Checked) discreteVal1.Value = cmbCliente.Value;
			else discreteVal1.Value = 0;
			paramField1.CurrentValues.Add (discreteVal1);
			paramFields.Add (paramField1);
		
			ParameterField paramField2 = new ParameterField ();
			ParameterDiscreteValue discreteVal2 = new ParameterDiscreteValue ();
			paramField2.ParameterFieldName = "@Negativo";
			if (!chkNegativos.Checked) discreteVal2.Value =	false;
			else discreteVal2.Value = true;
			paramField2.CurrentValues.Add (discreteVal2);
			paramFields.Add (paramField2);

			ParameterField paramField3 = new ParameterField ();
			ParameterDiscreteValue discreteVal3 = new ParameterDiscreteValue ();
			paramField3.ParameterFieldName = "@TodoPedido";
			if (!chkPedidos.Checked) discreteVal3.Value =	false;
			else discreteVal3.Value =	true;
			paramField3.CurrentValues.Add (discreteVal3);
			paramFields.Add (paramField3);

			ParameterField paramField4 = new ParameterField ();
			ParameterDiscreteValue discreteVal4 = new ParameterDiscreteValue ();
			paramField4.ParameterFieldName = "@CodPedido";
			discreteVal4.Value = txtPedido.Text;
			paramField4.CurrentValues.Add (discreteVal4);
			paramFields.Add (paramField4);

			#endregion Parametro Fecha Desde

			string stReporte = cmbReporte.Text;

			Reporte miRep = new Reporte(stReporte);
			miRep.MdiParent = this.MdiParent;
			miRep.crVista.ParameterFieldInfo = paramFields;
//			miRep.Titulo(stTitulo);
			Cursor = Cursors.Default;
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
