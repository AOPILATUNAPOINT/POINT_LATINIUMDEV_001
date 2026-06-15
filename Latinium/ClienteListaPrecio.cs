using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de ClienteListaPrecio.
	/// </summary>
	public class ClienteListaPrecio : DevExpress.XtraEditors.XtraForm
	{
		private Infragistics.Win.UltraWinGrid.UltraGrid ultraGrid1;
		private C1.Data.C1DataSet cdsClientePrecio;
		private Infragistics.Win.UltraWinGrid.UltraDropDown cmbArticulo;
		private C1.Data.C1DataSet cdsArticulo;
		private Infragistics.Win.Misc.UltraButton btAceptar;
		private Infragistics.Win.Misc.UltraButton btCancelar;
		private C1.Data.C1DataSet cdsCliente;
		private Infragistics.Win.Misc.UltraButton btFiltro;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbCliente;
		private System.Windows.Forms.ErrorProvider errorProvider;
		private C1.Data.C1DataView cdvCliente;
		private System.Windows.Forms.Label label1;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public ClienteListaPrecio()
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
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("ClienteListaPrecio", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idClienteListaPrecio");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idArticulo", -1, "cmbArticulo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCliente");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Porcentaje");
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Articulo", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Descontinuado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idArticulo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("dCompra");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("dVenta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idTipoGrupo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Articulo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn11 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn12 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idGrupoArticulo");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand3 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Cliente", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn13 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn14 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCliente");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn15 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn16 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Proveedor");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn17 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Ruc");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn18 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Ambos");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn19 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("DesconProv");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn20 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("DesconCliente");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn21 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Entrega");
			this.ultraGrid1 = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.cdsClientePrecio = new C1.Data.C1DataSet();
			this.btAceptar = new Infragistics.Win.Misc.UltraButton();
			this.btCancelar = new Infragistics.Win.Misc.UltraButton();
			this.cmbArticulo = new Infragistics.Win.UltraWinGrid.UltraDropDown();
			this.cdsArticulo = new C1.Data.C1DataSet();
			this.cdsCliente = new C1.Data.C1DataSet();
			this.cmbCliente = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.btFiltro = new Infragistics.Win.Misc.UltraButton();
			this.errorProvider = new System.Windows.Forms.ErrorProvider();
			this.cdvCliente = new C1.Data.C1DataView();
			this.label1 = new System.Windows.Forms.Label();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsClientePrecio)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbArticulo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsArticulo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsCliente)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCliente)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvCliente)).BeginInit();
			this.SuspendLayout();
			// 
			// ultraGrid1
			// 
			this.ultraGrid1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.ultraGrid1.Cursor = System.Windows.Forms.Cursors.Default;
			this.ultraGrid1.DataMember = "ClienteListaPrecio";
			this.ultraGrid1.DataSource = this.cdsClientePrecio;
			appearance1.BackColor = System.Drawing.Color.White;
			this.ultraGrid1.DisplayLayout.Appearance = appearance1;
			ultraGridColumn1.Header.VisiblePosition = 2;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn2.Header.Caption = "Articulo";
			ultraGridColumn2.Header.VisiblePosition = 0;
			ultraGridColumn2.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownList;
			ultraGridColumn2.Width = 289;
			ultraGridColumn3.Header.VisiblePosition = 3;
			ultraGridColumn3.Hidden = true;
			appearance2.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn4.CellAppearance = appearance2;
			ultraGridColumn4.Format = "#,##0.0000";
			ultraGridColumn4.Header.Caption = "Precio";
			ultraGridColumn4.Header.VisiblePosition = 1;
			ultraGridColumn4.Width = 88;
			ultraGridBand1.Columns.AddRange(new object[] {
																										 ultraGridColumn1,
																										 ultraGridColumn2,
																										 ultraGridColumn3,
																										 ultraGridColumn4});
			this.ultraGrid1.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.ultraGrid1.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.TemplateOnBottom;
			this.ultraGrid1.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.True;
			this.ultraGrid1.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance3.BackColor = System.Drawing.Color.Transparent;
			this.ultraGrid1.DisplayLayout.Override.CardAreaAppearance = appearance3;
			appearance4.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance4.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance4.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance4.FontData.BoldAsString = "True";
			appearance4.FontData.Name = "Arial";
			appearance4.FontData.SizeInPoints = 10F;
			appearance4.ForeColor = System.Drawing.Color.White;
			appearance4.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.ultraGrid1.DisplayLayout.Override.HeaderAppearance = appearance4;
			appearance5.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance5.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance5.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.RowSelectorAppearance = appearance5;
			appearance6.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance6.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance6.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.SelectedRowAppearance = appearance6;
			this.ultraGrid1.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.ultraGrid1.Location = new System.Drawing.Point(16, 48);
			this.ultraGrid1.Name = "ultraGrid1";
			this.ultraGrid1.Size = new System.Drawing.Size(404, 232);
			this.ultraGrid1.TabIndex = 0;
			this.ultraGrid1.AfterRowInsert += new Infragistics.Win.UltraWinGrid.RowEventHandler(this.ultraGrid1_AfterRowInsert);
			this.ultraGrid1.BeforeRowInsert += new Infragistics.Win.UltraWinGrid.BeforeRowInsertEventHandler(this.ultraGrid1_BeforeRowInsert);
			// 
			// cdsClientePrecio
			// 
			this.cdsClientePrecio.BindingContextCtrl = this;
			this.cdsClientePrecio.DataLibrary = "LibFacturacion";
			this.cdsClientePrecio.DataLibraryUrl = "";
			this.cdsClientePrecio.DataSetDef = "dsArticuloPrecio";
			this.cdsClientePrecio.FillOnRequest = false;
			this.cdsClientePrecio.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsClientePrecio.Name = "cdsClientePrecio";
			this.cdsClientePrecio.SchemaClassName = "LibFacturacion.DataClass";
			this.cdsClientePrecio.SchemaDef = null;
			this.cdsClientePrecio.BeforeFill += new C1.Data.FillEventHandler(this.cdsClientePrecio_BeforeFill);
			// 
			// btAceptar
			// 
			this.btAceptar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btAceptar.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.btAceptar.Location = new System.Drawing.Point(48, 300);
			this.btAceptar.Name = "btAceptar";
			this.btAceptar.Size = new System.Drawing.Size(72, 24);
			this.btAceptar.TabIndex = 1;
			this.btAceptar.Text = "&Grabar";
			this.btAceptar.Click += new System.EventHandler(this.btAceptar_Click);
			// 
			// btCancelar
			// 
			this.btCancelar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.btCancelar.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.btCancelar.DialogResult = System.Windows.Forms.DialogResult.Cancel;
			this.btCancelar.Location = new System.Drawing.Point(316, 300);
			this.btCancelar.Name = "btCancelar";
			this.btCancelar.Size = new System.Drawing.Size(72, 24);
			this.btCancelar.TabIndex = 2;
			this.btCancelar.Text = "&Salir";
			this.btCancelar.Click += new System.EventHandler(this.btCancelar_Click);
			// 
			// cmbArticulo
			// 
			this.cmbArticulo.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbArticulo.DataMember = "Articulo";
			this.cmbArticulo.DataSource = this.cdsArticulo;
			ultraGridBand2.AddButtonCaption = "ArticuloPrecios";
			ultraGridColumn5.Header.VisiblePosition = 2;
			ultraGridColumn5.Hidden = true;
			ultraGridColumn6.Header.VisiblePosition = 3;
			ultraGridColumn6.Hidden = true;
			ultraGridColumn7.Header.VisiblePosition = 4;
			ultraGridColumn7.Hidden = true;
			ultraGridColumn8.Header.VisiblePosition = 5;
			ultraGridColumn8.Hidden = true;
			ultraGridColumn9.Header.VisiblePosition = 6;
			ultraGridColumn9.Hidden = true;
			ultraGridColumn10.Header.VisiblePosition = 0;
			ultraGridColumn11.Header.VisiblePosition = 1;
			ultraGridColumn12.Header.VisiblePosition = 7;
			ultraGridBand2.Columns.AddRange(new object[] {
																										 ultraGridColumn5,
																										 ultraGridColumn6,
																										 ultraGridColumn7,
																										 ultraGridColumn8,
																										 ultraGridColumn9,
																										 ultraGridColumn10,
																										 ultraGridColumn11,
																										 ultraGridColumn12});
			this.cmbArticulo.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			this.cmbArticulo.DisplayMember = "Articulo";
			this.cmbArticulo.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbArticulo.Location = new System.Drawing.Point(104, 88);
			this.cmbArticulo.Name = "cmbArticulo";
			this.cmbArticulo.Size = new System.Drawing.Size(136, 88);
			this.cmbArticulo.TabIndex = 3;
			this.cmbArticulo.ValueMember = "idArticulo";
			this.cmbArticulo.Visible = false;
			// 
			// cdsArticulo
			// 
			this.cdsArticulo.BindingContextCtrl = this;
			this.cdsArticulo.DataLibrary = "LibFacturacion";
			this.cdsArticulo.DataLibraryUrl = "";
			this.cdsArticulo.DataSetDef = "dsArticulos";
			this.cdsArticulo.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsArticulo.Name = "cdsArticulo";
			this.cdsArticulo.SchemaClassName = "LibFacturacion.DataClass";
			this.cdsArticulo.SchemaDef = null;
			this.cdsArticulo.BeforeFill += new C1.Data.FillEventHandler(this.cdsArticulo_BeforeFill);
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
			// 
			// cmbCliente
			// 
			this.cmbCliente.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbCliente.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbCliente.DataMember = "Cliente";
			this.cmbCliente.DataSource = this.cdsCliente;
			ultraGridColumn13.Header.VisiblePosition = 1;
			ultraGridColumn14.Header.VisiblePosition = 2;
			ultraGridColumn14.Hidden = true;
			ultraGridColumn15.Header.VisiblePosition = 0;
			ultraGridColumn15.Width = 357;
			ultraGridColumn16.Header.VisiblePosition = 3;
			ultraGridColumn16.Hidden = true;
			ultraGridColumn17.Header.VisiblePosition = 4;
			ultraGridColumn17.Hidden = true;
			ultraGridColumn18.Header.VisiblePosition = 5;
			ultraGridColumn18.Hidden = true;
			ultraGridColumn19.Header.VisiblePosition = 6;
			ultraGridColumn19.Hidden = true;
			ultraGridColumn20.Header.VisiblePosition = 7;
			ultraGridColumn20.Hidden = true;
			ultraGridColumn21.Header.VisiblePosition = 8;
			ultraGridBand3.Columns.AddRange(new object[] {
																										 ultraGridColumn13,
																										 ultraGridColumn14,
																										 ultraGridColumn15,
																										 ultraGridColumn16,
																										 ultraGridColumn17,
																										 ultraGridColumn18,
																										 ultraGridColumn19,
																										 ultraGridColumn20,
																										 ultraGridColumn21});
			this.cmbCliente.DisplayLayout.BandsSerializer.Add(ultraGridBand3);
			this.cmbCliente.DisplayMember = "Nombre";
			this.cmbCliente.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbCliente.Location = new System.Drawing.Point(80, 16);
			this.cmbCliente.Name = "cmbCliente";
			this.cmbCliente.Size = new System.Drawing.Size(240, 21);
			this.cmbCliente.TabIndex = 4;
			this.cmbCliente.ValueMember = "idCliente";
			// 
			// btFiltro
			// 
			this.btFiltro.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.btFiltro.Location = new System.Drawing.Point(360, 16);
			this.btFiltro.Name = "btFiltro";
			this.btFiltro.Size = new System.Drawing.Size(56, 24);
			this.btFiltro.TabIndex = 12;
			this.btFiltro.Text = "&Filtrar";
			this.btFiltro.Click += new System.EventHandler(this.btFiltro_Click);
			// 
			// errorProvider
			// 
			this.errorProvider.ContainerControl = this;
			// 
			// cdvCliente
			// 
			this.cdvCliente.BindingContextCtrl = this;
			this.cdvCliente.DataSet = this.cdsClientePrecio;
			this.cdvCliente.TableName = "";
			this.cdvCliente.TableViewName = "ClienteListaPrecio";
			// 
			// label1
			// 
			this.label1.AutoSize = true;
			this.label1.Location = new System.Drawing.Point(24, 16);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(40, 16);
			this.label1.TabIndex = 13;
			this.label1.Text = "Cliente";
			// 
			// ClienteListaPrecio
			// 
			this.AcceptButton = this.btAceptar;
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.CancelButton = this.btCancelar;
			this.ClientSize = new System.Drawing.Size(432, 333);
			this.Controls.Add(this.label1);
			this.Controls.Add(this.btFiltro);
			this.Controls.Add(this.cmbCliente);
			this.Controls.Add(this.cmbArticulo);
			this.Controls.Add(this.btCancelar);
			this.Controls.Add(this.btAceptar);
			this.Controls.Add(this.ultraGrid1);
			this.Name = "ClienteListaPrecio";
			this.Text = "Lista de Precios de Clientes";
			this.Load += new System.EventHandler(this.ClienteListaPrecio_Load);
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsClientePrecio)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbArticulo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsArticulo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsCliente)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCliente)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvCliente)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void ClienteListaPrecio_Load(object sender, System.EventArgs e)
		{

		}

		private void btCancelar_Click(object sender, System.EventArgs e)
		{
			Close();
		}

		private void cdsArticulo_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			this.cdsArticulo.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);
		}

		private void cdsClientePrecio_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			this.cdsClientePrecio.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);
		}

		private void btAceptar_Click(object sender, System.EventArgs e)
		{
			try
			{
				cdsClientePrecio.Update();
			}
			catch(System.Exception ex)
			{
				MessageBox.Show(ex.Message, "No puedo grabar Precios por cliente");
				return;
			}
			#region Crea Lista de Precios en Empresas Relacionadas
			if (Funcion.bEjecutaSQL(cdsClientePrecio, "Exec SeteoGExiste 'FCLAER'"))
			{
				string stMensaje = Funcion.sEscalarSQL(cdsClientePrecio, "Exec GIM_ClienteListaPrecio "
					+"'"+MenuLatinium.stDirFacturacion+"'");

				if (stMensaje.Length > 0)
				{
					MessageBox.Show(stMensaje, "Información", MessageBoxButtons.OK, MessageBoxIcon.Warning);
					return;
				}
			}
			#endregion
			this.Close();
		}

		private void btFiltro_Click(object sender, System.EventArgs e)
		{
			errorProvider.SetError(cmbCliente, "");
			if (cmbCliente.Value == null)
			{
				errorProvider.SetError(cmbCliente, "Ingrese Cliente");
				return;
			}

			string stFiltro = "idCliente = " + cmbCliente.Value.ToString();

			C1.Data.FilterCondition fcPadre = new 
				C1.Data.FilterCondition(this.cdsClientePrecio, "ClienteListaPrecio", stFiltro);
			C1.Data.FilterConditions fcTotal = new C1.Data.FilterConditions();
			fcTotal.Add(fcPadre);

			cdsClientePrecio.Fill(fcTotal, false);
			if (this.cdvCliente.Count == 0)
			{
				MessageBox.Show("No existen Precios para este Cliente",
					"Información", MessageBoxButtons.OK, MessageBoxIcon.Information);
				return;
			}		
		}

		private void ultraGrid1_AfterRowInsert(object sender, Infragistics.Win.UltraWinGrid.RowEventArgs e)
		{
			e.Row.Cells["idCliente"].Value = (int) cmbCliente.Value;
		}

		private void ultraGrid1_BeforeRowInsert(object sender, Infragistics.Win.UltraWinGrid.BeforeRowInsertEventArgs e)
		{
			errorProvider.SetError(cmbCliente, "");
			if (cmbCliente.Value == null)
			{
				errorProvider.SetError(cmbCliente, "Ingrese Cliente");
				e.Cancel = true;
			}		
		}
	}
}
