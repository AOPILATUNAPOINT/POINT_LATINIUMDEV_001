using System;
using System.Data.SqlClient;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de BuscaArticulos.
	/// </summary>
	public class BuscaClientes : DevExpress.XtraEditors.XtraForm
	{
		private DevExpress.XtraEditors.SimpleButton btAceptar;
		private DevExpress.XtraEditors.SimpleButton btCancelar;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtArticulo;
		private DevExpress.XtraEditors.SimpleButton btBuscar;
		public Infragistics.Win.UltraWinGrid.UltraGrid ultraGrid1;
		private C1.Data.C1DataSet cdsCliente;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor spnCant;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource1;
		private System.Windows.Forms.Label lblCliente;
		public System.Windows.Forms.ComboBox cmbCriterios;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public BuscaClientes()
		{
			InitializeComponent();
		}

		string stBusca = "";
		int IdTipoFactura = 1;
		int IBodega = 0;
		int IFormaPago = 0;
		int IdTipoRuc = 0;
		bool BConsignacion = false;
		public BuscaClientes(string stBusca1, int idTipoFactura, int iFormaPago, int idTipoRuc, int iBodega, bool bConsignacion)
		{
			InitializeComponent();
			stBusca = stBusca1;
			IdTipoFactura = idTipoFactura;
			IFormaPago = iFormaPago;
			IdTipoRuc = idTipoRuc;
			IBodega = iBodega;
			BConsignacion = bConsignacion;
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
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Ruc");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Comercial");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCliente");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Descuento");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ContadoCred");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idGrupoCliente", 0);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idTipoRuc", 1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("MaximoCredito", 2);
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Ruc");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn3 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Comercial");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn4 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCliente");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn5 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Descuento");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn6 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("ContadoCred");
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			this.cdsCliente = new C1.Data.C1DataSet();
			this.btAceptar = new DevExpress.XtraEditors.SimpleButton();
			this.btCancelar = new DevExpress.XtraEditors.SimpleButton();
			this.txtArticulo = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.btBuscar = new DevExpress.XtraEditors.SimpleButton();
			this.ultraGrid1 = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource1 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.spnCant = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.lblCliente = new System.Windows.Forms.Label();
			this.cmbCriterios = new System.Windows.Forms.ComboBox();
			((System.ComponentModel.ISupportInitialize)(this.cdsCliente)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtArticulo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.spnCant)).BeginInit();
			this.SuspendLayout();
			// 
			// cdsCliente
			// 
			this.cdsCliente.BindingContextCtrl = this;
			this.cdsCliente.CaseSensitive = false;
			this.cdsCliente.DataLibrary = "LibFacturacion";
			this.cdsCliente.DataLibraryUrl = "";
			this.cdsCliente.DataSetDef = "dsClientes";
			this.cdsCliente.FillOnRequest = false;
			this.cdsCliente.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsCliente.Name = "cdsCliente";
			this.cdsCliente.SchemaClassName = "LibFacturacion.DataClass";
			this.cdsCliente.SchemaDef = null;
			this.cdsCliente.BeforeFill += new C1.Data.FillEventHandler(this.cdsArticulo_BeforeFill);
			// 
			// btAceptar
			// 
			this.btAceptar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btAceptar.Location = new System.Drawing.Point(8, 264);
			this.btAceptar.Name = "btAceptar";
			this.btAceptar.Size = new System.Drawing.Size(96, 25);
			this.btAceptar.TabIndex = 4;
			this.btAceptar.Text = "&Aceptar";
			this.btAceptar.Click += new System.EventHandler(this.btAceptar_Click);
			// 
			// btCancelar
			// 
			this.btCancelar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.btCancelar.Location = new System.Drawing.Point(120, 264);
			this.btCancelar.Name = "btCancelar";
			this.btCancelar.Size = new System.Drawing.Size(96, 25);
			this.btCancelar.TabIndex = 5;
			this.btCancelar.Text = "&Cancelar";
			this.btCancelar.Click += new System.EventHandler(this.btCancelar_Click);
			// 
			// txtArticulo
			// 
			appearance1.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtArticulo.Appearance = appearance1;
			this.txtArticulo.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtArticulo.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtArticulo.Location = new System.Drawing.Point(216, 8);
			this.txtArticulo.Name = "txtArticulo";
			this.txtArticulo.Size = new System.Drawing.Size(360, 22);
			this.txtArticulo.TabIndex = 0;
			this.txtArticulo.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtArticulo_KeyPress);
			this.txtArticulo.Validated += new System.EventHandler(this.txtArticulo_Validated);
			this.txtArticulo.ValueChanged += new System.EventHandler(this.txtArticulo_ValueChanged);
			this.txtArticulo.KeyUp += new System.Windows.Forms.KeyEventHandler(this.txtArticulo_KeyUp);
			// 
			// btBuscar
			// 
			this.btBuscar.Location = new System.Drawing.Point(592, 8);
			this.btBuscar.Name = "btBuscar";
			this.btBuscar.Size = new System.Drawing.Size(96, 22);
			this.btBuscar.TabIndex = 2;
			this.btBuscar.Text = "&Buscar";
			this.btBuscar.Click += new System.EventHandler(this.btBuscar_Click);
			// 
			// ultraGrid1
			// 
			this.ultraGrid1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.ultraGrid1.Cursor = System.Windows.Forms.Cursors.Default;
			this.ultraGrid1.DataSource = this.ultraDataSource1;
			appearance2.BackColor = System.Drawing.Color.White;
			appearance2.ForeColor = System.Drawing.Color.Black;
			appearance2.ForeColorDisabled = System.Drawing.Color.Black;
			this.ultraGrid1.DisplayLayout.Appearance = appearance2;
			this.ultraGrid1.DisplayLayout.AutoFitColumns = true;
			ultraGridColumn1.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn1.Header.Caption = "R.U.C.";
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn1.Width = 113;
			ultraGridColumn2.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn2.Header.Caption = "NOMBRE";
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridColumn2.Width = 300;
			ultraGridColumn3.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn3.Header.Caption = "NOMBRE COMERCIAL";
			ultraGridColumn3.Header.VisiblePosition = 2;
			ultraGridColumn3.Width = 300;
			ultraGridColumn4.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn4.Header.Caption = "ID";
			ultraGridColumn4.Header.VisiblePosition = 3;
			ultraGridColumn4.Hidden = true;
			ultraGridColumn4.Width = 178;
			ultraGridColumn5.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn5.Header.VisiblePosition = 4;
			ultraGridColumn5.Hidden = true;
			ultraGridColumn5.Width = 72;
			ultraGridColumn6.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn6.Header.VisiblePosition = 5;
			ultraGridColumn6.Hidden = true;
			ultraGridColumn6.Width = 89;
			ultraGridColumn7.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn7.Header.VisiblePosition = 6;
			ultraGridColumn7.Hidden = true;
			ultraGridColumn7.Width = 96;
			ultraGridColumn8.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn8.Header.VisiblePosition = 7;
			ultraGridColumn8.Hidden = true;
			ultraGridColumn8.Width = 96;
			ultraGridColumn9.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn9.Header.Caption = "Máximo Credito";
			ultraGridColumn9.Header.VisiblePosition = 8;
			ultraGridColumn9.Hidden = true;
			ultraGridColumn9.Width = 95;
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
			this.ultraGrid1.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			appearance3.ForeColor = System.Drawing.Color.Black;
			this.ultraGrid1.DisplayLayout.CaptionAppearance = appearance3;
			this.ultraGrid1.DisplayLayout.GroupByBox.Prompt = "Arrastre columna a Agrupar";
			appearance4.ForeColor = System.Drawing.Color.Black;
			appearance4.ForeColorDisabled = System.Drawing.Color.Black;
			this.ultraGrid1.DisplayLayout.Override.ActiveRowAppearance = appearance4;
			this.ultraGrid1.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
			this.ultraGrid1.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.ultraGrid1.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.False;
			appearance5.BackColor = System.Drawing.Color.Transparent;
			this.ultraGrid1.DisplayLayout.Override.CardAreaAppearance = appearance5;
			this.ultraGrid1.DisplayLayout.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.RowSelect;
			this.ultraGrid1.DisplayLayout.Override.ColumnAutoSizeMode = Infragistics.Win.UltraWinGrid.ColumnAutoSizeMode.AllRowsInBand;
			appearance6.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance6.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance6.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance6.FontData.BoldAsString = "True";
			appearance6.FontData.Name = "Arial";
			appearance6.FontData.SizeInPoints = 8.5F;
			appearance6.ForeColor = System.Drawing.Color.White;
			appearance6.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.ultraGrid1.DisplayLayout.Override.HeaderAppearance = appearance6;
			this.ultraGrid1.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti;
			appearance7.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance7.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance7.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.RowSelectorAppearance = appearance7;
			appearance8.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance8.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance8.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.SelectedRowAppearance = appearance8;
			this.ultraGrid1.DisplayLayout.Override.SelectTypeCell = Infragistics.Win.UltraWinGrid.SelectType.None;
			this.ultraGrid1.DisplayLayout.Override.SelectTypeCol = Infragistics.Win.UltraWinGrid.SelectType.None;
			this.ultraGrid1.DisplayLayout.Override.SelectTypeRow = Infragistics.Win.UltraWinGrid.SelectType.Single;
			this.ultraGrid1.DisplayLayout.ScrollBounds = Infragistics.Win.UltraWinGrid.ScrollBounds.ScrollToFill;
			this.ultraGrid1.DisplayLayout.ScrollStyle = Infragistics.Win.UltraWinGrid.ScrollStyle.Immediate;
			this.ultraGrid1.DisplayLayout.TabNavigation = Infragistics.Win.UltraWinGrid.TabNavigation.NextControl;
			this.ultraGrid1.DisplayLayout.ViewStyle = Infragistics.Win.UltraWinGrid.ViewStyle.SingleBand;
			this.ultraGrid1.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.ultraGrid1.Location = new System.Drawing.Point(5, 39);
			this.ultraGrid1.Name = "ultraGrid1";
			this.ultraGrid1.Size = new System.Drawing.Size(734, 220);
			this.ultraGrid1.TabIndex = 3;
			this.ultraGrid1.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.ultraGrid1_KeyPress);
			this.ultraGrid1.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.ultraGrid1_InitializeLayout);
			// 
			// ultraDataSource1
			// 
			ultraDataColumn5.DataType = typeof(System.Decimal);
			ultraDataColumn5.DefaultValue = new System.Decimal(new int[] {
																																		 0,
																																		 0,
																																		 0,
																																		 0});
			ultraDataColumn6.DataType = typeof(int);
			this.ultraDataSource1.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn1,
																																 ultraDataColumn2,
																																 ultraDataColumn3,
																																 ultraDataColumn4,
																																 ultraDataColumn5,
																																 ultraDataColumn6});
			// 
			// spnCant
			// 
			appearance9.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(247)), ((System.Byte)(245)), ((System.Byte)(241)));
			appearance9.ForeColor = System.Drawing.Color.Black;
			appearance9.ForeColorDisabled = System.Drawing.Color.Black;
			this.spnCant.Appearance = appearance9;
			this.spnCant.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(247)), ((System.Byte)(245)), ((System.Byte)(241)));
			this.spnCant.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.spnCant.ForeColor = System.Drawing.Color.Black;
			this.spnCant.Location = new System.Drawing.Point(696, 8);
			this.spnCant.MaskInput = "nnn";
			this.spnCant.Name = "spnCant";
			this.spnCant.PromptChar = ' ';
			this.spnCant.Size = new System.Drawing.Size(40, 22);
			this.spnCant.SpinButtonDisplayStyle = Infragistics.Win.ButtonDisplayStyle.Always;
			this.spnCant.TabIndex = 32;
			this.spnCant.Value = 10;
			// 
			// lblCliente
			// 
			this.lblCliente.AutoSize = true;
			this.lblCliente.Location = new System.Drawing.Point(8, 11);
			this.lblCliente.Name = "lblCliente";
			this.lblCliente.Size = new System.Drawing.Size(40, 16);
			this.lblCliente.TabIndex = 37;
			this.lblCliente.Text = "Criterio";
			// 
			// cmbCriterios
			// 
			this.cmbCriterios.Location = new System.Drawing.Point(56, 8);
			this.cmbCriterios.Name = "cmbCriterios";
			this.cmbCriterios.Size = new System.Drawing.Size(152, 21);
			this.cmbCriterios.TabIndex = 1;
			this.cmbCriterios.SelectedIndexChanged += new System.EventHandler(this.cmbCriterios_SelectedIndexChanged);
			// 
			// BuscaClientes
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 14);
			this.ClientSize = new System.Drawing.Size(744, 294);
			this.ControlBox = false;
			this.Controls.Add(this.cmbCriterios);
			this.Controls.Add(this.lblCliente);
			this.Controls.Add(this.spnCant);
			this.Controls.Add(this.ultraGrid1);
			this.Controls.Add(this.btBuscar);
			this.Controls.Add(this.txtArticulo);
			this.Controls.Add(this.btCancelar);
			this.Controls.Add(this.btAceptar);
			this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
			this.KeyPreview = true;
			this.MaximizeBox = false;
			this.MinimizeBox = false;
			this.Name = "BuscaClientes";
			this.ShowInTaskbar = false;
			this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
			this.Text = "Buscar";
			this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.BuscaClientes_KeyDown);
			this.Load += new System.EventHandler(this.BuscaArticulos_Load);
			((System.ComponentModel.ISupportInitialize)(this.cdsCliente)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtArticulo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.spnCant)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void btBuscar_Click(object sender, System.EventArgs e)
		{
			string sSQL = string.Format("Exec BusquedaClientes {0}, {1}, '{2}', '{3}', {4}, {5}", IdTipoFactura, IBodega, this.cmbCriterios.Text, this.txtArticulo.Text, IdTipoRuc, BConsignacion);
			ultraGrid1.DataSource = Funcion.dvProcedimiento(cdsCliente, sSQL);

			if (this.ultraGrid1.Rows.Count == 0)
			{
				if (IdTipoFactura == 44)
				{
					MessageBox.Show("No existen clientes que Aplican a Consignación.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
					return;
				}

				if (IdTipoFactura == 2 && BConsignacion)
				{
					MessageBox.Show("No existen proveedores que Aplican a Consignación.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
					return;
				}
			}
			this.ultraGrid1.Focus();
		}

		public string Nombre = "";
		public int idCliente = 0;
		public string Ruc = "";
		public int idGrupoCliente = 0;
		public int idTipoRuc = 0;
		public decimal MaximoCredito = 0;
		public decimal Descuento = 0;
		public int ContadoCred = 0;
		
		private void BuscaArticulos_Load(object sender, System.EventArgs e)
		{
			if (Funcion.EsProveedor(IdTipoFactura)) this.Text = this.Text + " Proveedores";
			else this.Text = this.Text + " Clientes";

			this.cmbCriterios.Text = "DESCRIPCION";
			this.cmbCriterios.Items.Add("CODIGO");
			this.cmbCriterios.Items.Add("DESCRIPCION");

			if (stBusca.Length > 0)
			{
				txtArticulo.Text = stBusca;
				btBuscar_Click(this, e);
			}

			this.txtArticulo.Focus();
		}

		private void cdsArticulo_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsCliente.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);
		}

		private void Aceptar()
		{
			if (ultraGrid1.ActiveRow == null)
			{
				MessageBox.Show("Seleccione la fila deseada", "Información", MessageBoxButtons.OK, MessageBoxIcon.Information);
				txtArticulo.Focus();
				txtArticulo.SelectAll();
				return;
			}
			
			idCliente = (int) ultraGrid1.ActiveRow.Cells["idCliente"].Value;
									
			if (IdTipoFactura == 13 && IFormaPago == 6) 
			{
				int DiasCred = 0;
				decimal MontoCred = 0;

				SqlDataReader dr = Funcion.rEscalarSQL(cdsCliente, string.Format("Select IsNull(DiasCredito, 0), IsNull(MaximoCredito, 0) From Cliente Where idCliente = {0}", idCliente), true);
				dr.Read();
				if (dr.HasRows)
				{
					DiasCred = (int)dr.GetValue(0);
					MontoCred = (decimal)dr.GetValue(1);
				}
				dr.Close();
				if ((int)DiasCred == 0 || (decimal)MontoCred == 0)
				{
					MessageBox.Show("El cliente no tiene Autorización para crédito", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
					return;
				}
			}

			Ruc = ultraGrid1.ActiveRow.Cells["Ruc"].Value.ToString();
			Nombre = ultraGrid1.ActiveRow.Cells["Nombre"].Value.ToString();
			if (ultraGrid1.ActiveRow.Cells["idGrupoCliente"].Value != System.DBNull.Value)idGrupoCliente = (int)ultraGrid1.ActiveRow.Cells["idGrupoCliente"].Value;
			idTipoRuc = (int)ultraGrid1.ActiveRow.Cells["idTipoRuc"].Value;
			MaximoCredito = (decimal)ultraGrid1.ActiveRow.Cells["MaximoCredito"].Value;
			ContadoCred = (int)ultraGrid1.ActiveRow.Cells["ContadoCred"].Value;

			if (IdTipoFactura == 2 || IdTipoFactura == 4) if (ultraGrid1.ActiveRow.Cells["Descuento"].Value != System.DBNull.Value) Descuento = (decimal)ultraGrid1.ActiveRow.Cells["Descuento"].Value;
			this.DialogResult = DialogResult.OK;
		}

		private void btAceptar_Click(object sender, System.EventArgs e)
		{
			Aceptar();
		}

		private void txtArticulo_Validated(object sender, System.EventArgs e)
		{
			
		}

		private void txtArticulo_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (e.KeyChar == 13) btBuscar_Click(this, e);			
		}

		private void txtArticulo_ValueChanged(object sender, System.EventArgs e)
		{
			
		}

		private void txtArticulo_KeyUp(object sender, System.Windows.Forms.KeyEventArgs e)
		{

		}

		private void cmbCriterios_SelectedIndexChanged(object sender, System.EventArgs e)
		{
			this.txtArticulo.Text = "";
			this.txtArticulo.Focus();
		}

		private void BuscaClientes_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Escape) Close();
		}

		private void ultraGrid1_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}

		private void btCancelar_Click(object sender, System.EventArgs e)
		{
			Close();
		}

		private void ultraGrid1_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (e.KeyChar == 13) Aceptar();
		}
	}
}
