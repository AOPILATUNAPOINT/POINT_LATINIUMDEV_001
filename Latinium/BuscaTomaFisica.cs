using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de BuscaAsiento.
	/// </summary>
	public class BuscaTomaFisica : DevExpress.XtraEditors.XtraForm
	{
		public System.Windows.Forms.TextBox idBusca;
		private System.Windows.Forms.ErrorProvider errorProvider;
		private Infragistics.Win.UltraWinGrid.UltraGrid ultraGrid1;
		private Infragistics.Win.UltraWinEditors.UltraOptionSet rdgOpcion;
		private Infragistics.Win.UltraWinEditors.UltraOptionSet rdgTipoBusqueda;
		private Infragistics.Win.Misc.UltraButton btBuscar;
		private Infragistics.Win.Misc.UltraButton btAceptar;
		private Infragistics.Win.Misc.UltraButton btCancelar;
		private C1.Data.C1DataSet cdsTomaFisica;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtTarjeta;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor spnBodega;
		private Infragistics.Win.UltraWinEditors.UltraDateTimeEditor dtFecha;
		private C1.Data.C1DataView cdvTomaFisica;
		private System.ComponentModel.IContainer components=null;

		public BuscaTomaFisica()
		{
			//
			// Necesario para admitir el Diseñador de Windows Forms
			//
			InitializeComponent();

			//
			// TODO: Agregar código de constructor después de llamar a InitializeComponent
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

		#region Windows Form Designer generated code
		/// <summary>
		/// Método necesario para admitir el Diseñador, no se puede modificar
		/// el contenido del método con el editor de código.
		/// </summary>
		private void InitializeComponent()
		{
			Infragistics.Win.Appearance appearance1 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("TmpTomaFisica", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idTmpTomaFisica");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Fecha");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Tarjeta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bodega");
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.ValueListItem valueListItem1 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem2 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem3 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.ValueListItem valueListItem4 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem5 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem6 = new Infragistics.Win.ValueListItem();
			this.cdsTomaFisica = new C1.Data.C1DataSet();
			this.idBusca = new System.Windows.Forms.TextBox();
			this.errorProvider = new System.Windows.Forms.ErrorProvider();
			this.ultraGrid1 = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.cdvTomaFisica = new C1.Data.C1DataView();
			this.rdgOpcion = new Infragistics.Win.UltraWinEditors.UltraOptionSet();
			this.rdgTipoBusqueda = new Infragistics.Win.UltraWinEditors.UltraOptionSet();
			this.btBuscar = new Infragistics.Win.Misc.UltraButton();
			this.btAceptar = new Infragistics.Win.Misc.UltraButton();
			this.btCancelar = new Infragistics.Win.Misc.UltraButton();
			this.txtTarjeta = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.spnBodega = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.dtFecha = new Infragistics.Win.UltraWinEditors.UltraDateTimeEditor();
			((System.ComponentModel.ISupportInitialize)(this.cdsTomaFisica)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvTomaFisica)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.rdgOpcion)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.rdgTipoBusqueda)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtTarjeta)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.spnBodega)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFecha)).BeginInit();
			this.SuspendLayout();
			// 
			// cdsTomaFisica
			// 
			this.cdsTomaFisica.BindingContextCtrl = this;
			this.cdsTomaFisica.CaseSensitive = false;
			this.cdsTomaFisica.DataLibrary = "LibFacturacion";
			this.cdsTomaFisica.DataLibraryUrl = "";
			this.cdsTomaFisica.DataSetDef = "dsTomaFisica";
			this.cdsTomaFisica.FillOnRequest = false;
			this.cdsTomaFisica.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsTomaFisica.Name = "cdsTomaFisica";
			this.cdsTomaFisica.SchemaClassName = "LibFacturacion.DataClass";
			this.cdsTomaFisica.SchemaDef = null;
			this.cdsTomaFisica.BeforeFill += new C1.Data.FillEventHandler(this.cdsArticulo_BeforeFill);
			// 
			// idBusca
			// 
			this.idBusca.DataBindings.Add(new System.Windows.Forms.Binding("Text", this.cdsTomaFisica, "TmpTomaFisica.idTmpTomaFisica"));
			this.idBusca.Location = new System.Drawing.Point(240, 95);
			this.idBusca.Name = "idBusca";
			this.idBusca.Size = new System.Drawing.Size(48, 20);
			this.idBusca.TabIndex = 23;
			this.idBusca.Text = "";
			// 
			// errorProvider
			// 
			this.errorProvider.ContainerControl = this;
			this.errorProvider.DataMember = "";
			// 
			// ultraGrid1
			// 
			this.ultraGrid1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.ultraGrid1.Cursor = System.Windows.Forms.Cursors.Default;
			this.ultraGrid1.DataMember = "TmpTomaFisica";
			this.ultraGrid1.DataSource = this.cdsTomaFisica;
			appearance1.BackColor = System.Drawing.Color.White;
			this.ultraGrid1.DisplayLayout.Appearance = appearance1;
			ultraGridBand1.AddButtonCaption = "Articulo";
			ultraGridColumn1.Header.VisiblePosition = 3;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn2.Header.VisiblePosition = 0;
			ultraGridColumn3.Header.VisiblePosition = 1;
			ultraGridColumn4.Header.VisiblePosition = 2;
			ultraGridBand1.Columns.AddRange(new object[] {
																										 ultraGridColumn1,
																										 ultraGridColumn2,
																										 ultraGridColumn3,
																										 ultraGridColumn4});
			this.ultraGrid1.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.ultraGrid1.DisplayLayout.GroupByBox.Prompt = "Arrastre columna a Agrupar";
			this.ultraGrid1.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
			this.ultraGrid1.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.ultraGrid1.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.False;
			appearance2.BackColor = System.Drawing.Color.Transparent;
			this.ultraGrid1.DisplayLayout.Override.CardAreaAppearance = appearance2;
			this.ultraGrid1.DisplayLayout.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.RowSelect;
			this.ultraGrid1.DisplayLayout.Override.ColumnAutoSizeMode = Infragistics.Win.UltraWinGrid.ColumnAutoSizeMode.AllRowsInBand;
			appearance3.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance3.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance3.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance3.FontData.BoldAsString = "True";
			appearance3.FontData.Name = "Arial";
			appearance3.FontData.SizeInPoints = 10F;
			appearance3.ForeColor = System.Drawing.Color.White;
			appearance3.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.ultraGrid1.DisplayLayout.Override.HeaderAppearance = appearance3;
			this.ultraGrid1.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti;
			appearance4.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance4.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance4.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.RowSelectorAppearance = appearance4;
			appearance5.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance5.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance5.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.SelectedRowAppearance = appearance5;
			this.ultraGrid1.DisplayLayout.Override.SelectTypeCell = Infragistics.Win.UltraWinGrid.SelectType.None;
			this.ultraGrid1.DisplayLayout.Override.SelectTypeCol = Infragistics.Win.UltraWinGrid.SelectType.None;
			this.ultraGrid1.DisplayLayout.Override.SelectTypeRow = Infragistics.Win.UltraWinGrid.SelectType.Single;
			this.ultraGrid1.DisplayLayout.ScrollBounds = Infragistics.Win.UltraWinGrid.ScrollBounds.ScrollToFill;
			this.ultraGrid1.DisplayLayout.ScrollStyle = Infragistics.Win.UltraWinGrid.ScrollStyle.Immediate;
			this.ultraGrid1.DisplayLayout.TabNavigation = Infragistics.Win.UltraWinGrid.TabNavigation.NextControl;
			this.ultraGrid1.DisplayLayout.ViewStyle = Infragistics.Win.UltraWinGrid.ViewStyle.SingleBand;
			this.ultraGrid1.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.ultraGrid1.Location = new System.Drawing.Point(16, 128);
			this.ultraGrid1.Name = "ultraGrid1";
			this.ultraGrid1.Size = new System.Drawing.Size(512, 296);
			this.ultraGrid1.TabIndex = 27;
			// 
			// cdvTomaFisica
			// 
			this.cdvTomaFisica.BindingContextCtrl = this;
			this.cdvTomaFisica.DataSet = this.cdsTomaFisica;
			this.cdvTomaFisica.TableName = "";
			this.cdvTomaFisica.TableViewName = "TmpTomaFisica";
			// 
			// rdgOpcion
			// 
			this.rdgOpcion.ItemAppearance = appearance6;
			this.rdgOpcion.ItemOrigin = new System.Drawing.Point(6, 0);
			valueListItem1.DataValue = 0;
			valueListItem1.DisplayText = "Tarjeta";
			valueListItem2.DataValue = 1;
			valueListItem2.DisplayText = "Bodega";
			valueListItem3.DataValue = 2;
			valueListItem3.DisplayText = "Fecha";
			this.rdgOpcion.Items.Add(valueListItem1);
			this.rdgOpcion.Items.Add(valueListItem2);
			this.rdgOpcion.Items.Add(valueListItem3);
			this.rdgOpcion.ItemSpacingVertical = 9;
			this.rdgOpcion.Location = new System.Drawing.Point(16, 17);
			this.rdgOpcion.Name = "rdgOpcion";
			this.rdgOpcion.Size = new System.Drawing.Size(112, 86);
			this.rdgOpcion.TabIndex = 40;
			this.rdgOpcion.ValueChanged += new System.EventHandler(this.rdgOpcion_SelectedIndexChanged);
			// 
			// rdgTipoBusqueda
			// 
			this.rdgTipoBusqueda.CheckedIndex = 2;
			this.rdgTipoBusqueda.ItemAppearance = appearance7;
			this.rdgTipoBusqueda.ItemOrigin = new System.Drawing.Point(10, 0);
			valueListItem4.DataValue = 0;
			valueListItem4.DisplayText = "Que Empiece";
			valueListItem5.DataValue = 1;
			valueListItem5.DisplayText = "Exacta";
			valueListItem6.DataValue = 2;
			valueListItem6.DisplayText = "En cualquier parte";
			this.rdgTipoBusqueda.Items.Add(valueListItem4);
			this.rdgTipoBusqueda.Items.Add(valueListItem5);
			this.rdgTipoBusqueda.Items.Add(valueListItem6);
			this.rdgTipoBusqueda.ItemSpacingVertical = 10;
			this.rdgTipoBusqueda.Location = new System.Drawing.Point(296, 17);
			this.rdgTipoBusqueda.Name = "rdgTipoBusqueda";
			this.rdgTipoBusqueda.Size = new System.Drawing.Size(136, 86);
			this.rdgTipoBusqueda.TabIndex = 41;
			this.rdgTipoBusqueda.Text = "En cualquier parte";
			// 
			// btBuscar
			// 
			this.btBuscar.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.btBuscar.Location = new System.Drawing.Point(448, 17);
			this.btBuscar.Name = "btBuscar";
			this.btBuscar.Size = new System.Drawing.Size(75, 25);
			this.btBuscar.TabIndex = 42;
			this.btBuscar.Text = "&Buscar";
			this.btBuscar.Click += new System.EventHandler(this.btBuscar_Click);
			// 
			// btAceptar
			// 
			this.btAceptar.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.btAceptar.Location = new System.Drawing.Point(448, 52);
			this.btAceptar.Name = "btAceptar";
			this.btAceptar.Size = new System.Drawing.Size(75, 24);
			this.btAceptar.TabIndex = 43;
			this.btAceptar.Text = "&Aceptar";
			this.btAceptar.Click += new System.EventHandler(this.btAceptar_Click);
			// 
			// btCancelar
			// 
			this.btCancelar.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.btCancelar.DialogResult = System.Windows.Forms.DialogResult.Cancel;
			this.btCancelar.Location = new System.Drawing.Point(448, 86);
			this.btCancelar.Name = "btCancelar";
			this.btCancelar.Size = new System.Drawing.Size(75, 25);
			this.btCancelar.TabIndex = 44;
			this.btCancelar.Text = "&Cancelar";
			// 
			// txtTarjeta
			// 
			this.txtTarjeta.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtTarjeta.Location = new System.Drawing.Point(136, 17);
			this.txtTarjeta.Name = "txtTarjeta";
			this.txtTarjeta.Size = new System.Drawing.Size(128, 22);
			this.txtTarjeta.TabIndex = 45;
			// 
			// spnBodega
			// 
			this.spnBodega.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.spnBodega.Location = new System.Drawing.Point(136, 43);
			this.spnBodega.MaxValue = 999;
			this.spnBodega.MinValue = 1;
			this.spnBodega.Name = "spnBodega";
			this.spnBodega.PromptChar = ' ';
			this.spnBodega.Size = new System.Drawing.Size(80, 22);
			this.spnBodega.SpinButtonDisplayStyle = Infragistics.Win.ButtonDisplayStyle.Always;
			this.spnBodega.TabIndex = 46;
			// 
			// dtFecha
			// 
			this.dtFecha.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.dtFecha.FormatString = "dd/MMM/yyyy";
			this.dtFecha.Location = new System.Drawing.Point(136, 69);
			this.dtFecha.Name = "dtFecha";
			this.dtFecha.Size = new System.Drawing.Size(128, 22);
			this.dtFecha.TabIndex = 47;
			// 
			// BuscaTomaFisica
			// 
			this.AcceptButton = this.btBuscar;
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 14);
			this.CancelButton = this.btCancelar;
			this.ClientSize = new System.Drawing.Size(544, 445);
			this.Controls.Add(this.dtFecha);
			this.Controls.Add(this.spnBodega);
			this.Controls.Add(this.txtTarjeta);
			this.Controls.Add(this.btCancelar);
			this.Controls.Add(this.btAceptar);
			this.Controls.Add(this.btBuscar);
			this.Controls.Add(this.rdgTipoBusqueda);
			this.Controls.Add(this.rdgOpcion);
			this.Controls.Add(this.ultraGrid1);
			this.Controls.Add(this.idBusca);
			this.MinimizeBox = false;
			this.Name = "BuscaTomaFisica";
			this.Text = "Búsqueda de Artículos";
			this.Resize += new System.EventHandler(this.BuscaArticulo_Resize);
			this.Load += new System.EventHandler(this.BuscaAsiento_Load);
			((System.ComponentModel.ISupportInitialize)(this.cdsTomaFisica)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvTomaFisica)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.rdgOpcion)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.rdgTipoBusqueda)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtTarjeta)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.spnBodega)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFecha)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void BuscaAsiento_Load(object sender, System.EventArgs e)
		{
			Cursor = Cursors.Default;
			this.idBusca.Width =0;
			this.rdgOpcion.Value = 0;
		}

		private bool Verifica()
		{
			errorProvider.SetError(txtTarjeta, "");
			errorProvider.SetError(spnBodega, "");
			errorProvider.SetError(dtFecha, "");

			bool bOk = true;
			int iTipo = (int) rdgOpcion.Value;

			if (iTipo == 0)
				if (txtTarjeta.Text.Trim().Length == 0)
				{
					errorProvider.SetError(txtTarjeta, "Ingrese  Articulo");
					bOk = false;
				}

			if (iTipo == 1)
				if (spnBodega.Value == null)
				{
					this.errorProvider.SetError(spnBodega, "Ingrese Bodega");
					bOk = false;
				}
			if (iTipo == 2)
				if (dtFecha.Value == null)
				{
					errorProvider.SetError(dtFecha, "Ingrese Fecha");
					bOk = false;
				}
			return bOk;
		}

		private void btBuscar_Click(object sender, System.EventArgs e)
		{
			if (!Verifica()) return;
			string strFiltro;
			string strAntes = "";
			string strDespues="'";

//		Opciones de comodines
			if ((int) this.rdgTipoBusqueda.Value == 0) strDespues = "%'";
			if ((int) this.rdgTipoBusqueda.Value == 1) strAntes = "";
			if ((int) this.rdgTipoBusqueda.Value == 2) strDespues = "%'";
			if ((int) this.rdgTipoBusqueda.Value == 2) strAntes = "%";

			strFiltro = "";
			string stBusca = "";
			switch((int) this.rdgOpcion.Value)
			{
				case 0:
					if (txtTarjeta.Value == null)
					{
						errorProvider.SetError(txtTarjeta, "Ingrese Tarjeta");
						return;
					}
					errorProvider.SetError(txtTarjeta, "");

					stBusca = this.txtTarjeta.Value.ToString().Trim();
					stBusca = stBusca.Replace("'", "''");
					strFiltro = "Tarjeta LIKE '" + strAntes + stBusca + strDespues;
					break;
				case 1:
					strFiltro = "Bodega = " + spnBodega.Value.ToString();
					break;
				case 2:
					DateTime dtFecha1 = (DateTime) dtFecha.Value;
					strFiltro = "Fecha = '" + dtFecha1.ToString("yyyyMMdd") + "'";
					break;
			}

			string stFiltro = strFiltro;

			C1.Data.FilterCondition fcPadre = new 
				C1.Data.FilterCondition(cdsTomaFisica, "TmpTomaFisica", stFiltro);
			C1.Data.FilterConditions fcTotal = new C1.Data.FilterConditions();
			fcTotal.Add(fcPadre);
			C1.Data.FilterCondition fcHija = new 
				C1.Data.FilterCondition(cdsTomaFisica, "TmpTomaDetalle", "idTmpTomaFisica = 0");
			fcTotal.Add(fcHija);

			cdsTomaFisica.Fill(fcTotal, false);
			if (cdvTomaFisica.Count == 0)
			{
				MessageBox.Show("No existen registros que cumplan la condicion de busqueda",
					"Información", MessageBoxButtons.OK, MessageBoxIcon.Information);
				return;
			}

			if (cdvTomaFisica.Count == 1)
			{
				this.ultraGrid1.Select();
				this.DialogResult = DialogResult.OK;
			}
			else
			{
				this.AcceptButton=this.btAceptar;
				this.ultraGrid1.Select();
			}
		}

//		public int IdBuscado()
//		{
//			int iFila = this.gridView1.FocusedRowHandle;
//			return int.Parse(this.cdvSocio[iFila]["idSocio"].ToString());
//		}

		private void btAceptar_Click(object sender, System.EventArgs e)
		{
			this.DialogResult = DialogResult.OK;
		}

		private void BuscaArticulo_Resize(object sender, System.EventArgs e)
		{
			this.ultraGrid1.DisplayLayout.AutoFitColumns = true;
		}

		private void cdsArticulo_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsTomaFisica.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);		
		}


		private void rdgOpcion_SelectedIndexChanged(object sender, System.EventArgs e)
		{
			txtTarjeta.Enabled = false;
			spnBodega.Enabled = false;
			dtFecha.Enabled = false;

			switch((int) this.rdgOpcion.Value)
			{
				case 0:
					txtTarjeta.Enabled = true;
					txtTarjeta.Select();
					break;
				case 1:
					spnBodega.Enabled = true;
					spnBodega.Select();
					break;
				case 2:
					dtFecha.Enabled = true;
					dtFecha.Select();
					break;
			}

		}


	
	}
}
