using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using Infragistics.Win.UltraWinGrid;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de AsientoGrupoComp.
	/// </summary>
	public class AsientoGrupoComp : DevExpress.XtraEditors.XtraForm
	{
		private DevExpress.XtraEditors.SimpleButton btAceptar;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtComprobante;
		private Infragistics.Win.Misc.UltraLabel ultraLabel1;
		private Infragistics.Win.UltraWinEditors.UltraOptionSet optTipoFact;
		private System.Windows.Forms.ErrorProvider errorProvider;
		private C1.Data.C1DataSet cdsSeteo;
		private C1.Data.C1DataSet cdsCompra;
		private DevExpress.XtraEditors.SimpleButton btFiltro;
		private DevExpress.XtraEditors.SimpleButton btFactura;
		private DevExpress.XtraEditors.SimpleButton btAsiento;
		private C1.Data.C1DataView cdvCompra;
		private DevExpress.XtraEditors.SimpleButton btRetencion;
		private DevExpress.XtraEditors.SimpleButton btIndividual;
		private System.Windows.Forms.ToolTip toolTip1;
		private System.Windows.Forms.CheckBox chkEgreso;
		private DevExpress.XtraGrid.GridControl grdCompras;
		private DevExpress.XtraGrid.Views.Grid.GridView grvCompras;
		private System.ComponentModel.IContainer components;

		public AsientoGrupoComp()
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
			Infragistics.Win.ValueListItem valueListItem1 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem2 = new Infragistics.Win.ValueListItem();
			this.btAceptar = new DevExpress.XtraEditors.SimpleButton();
			this.txtComprobante = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.ultraLabel1 = new Infragistics.Win.Misc.UltraLabel();
			this.optTipoFact = new Infragistics.Win.UltraWinEditors.UltraOptionSet();
			this.errorProvider = new System.Windows.Forms.ErrorProvider();
			this.cdsSeteo = new C1.Data.C1DataSet();
			this.cdsCompra = new C1.Data.C1DataSet();
			this.btFiltro = new DevExpress.XtraEditors.SimpleButton();
			this.btFactura = new DevExpress.XtraEditors.SimpleButton();
			this.btAsiento = new DevExpress.XtraEditors.SimpleButton();
			this.cdvCompra = new C1.Data.C1DataView();
			this.btRetencion = new DevExpress.XtraEditors.SimpleButton();
			this.btIndividual = new DevExpress.XtraEditors.SimpleButton();
			this.toolTip1 = new System.Windows.Forms.ToolTip(this.components);
			this.chkEgreso = new System.Windows.Forms.CheckBox();
			this.grdCompras = new DevExpress.XtraGrid.GridControl();
			this.grvCompras = new DevExpress.XtraGrid.Views.Grid.GridView();
			((System.ComponentModel.ISupportInitialize)(this.txtComprobante)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.optTipoFact)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsCompra)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvCompra)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.grdCompras)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.grvCompras)).BeginInit();
			this.SuspendLayout();
			// 
			// btAceptar
			// 
			this.btAceptar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btAceptar.Location = new System.Drawing.Point(560, 207);
			this.btAceptar.Name = "btAceptar";
			this.btAceptar.Size = new System.Drawing.Size(75, 25);
			this.btAceptar.TabIndex = 0;
			this.btAceptar.Text = "&Resumen";
			this.btAceptar.Click += new System.EventHandler(this.btAceptar_Click);
			// 
			// txtComprobante
			// 
			appearance1.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtComprobante.Appearance = appearance1;
			this.txtComprobante.Location = new System.Drawing.Point(128, 16);
			this.txtComprobante.Name = "txtComprobante";
			this.txtComprobante.Size = new System.Drawing.Size(100, 22);
			this.txtComprobante.TabIndex = 1;
			// 
			// ultraLabel1
			// 
			this.ultraLabel1.AutoSize = true;
			this.ultraLabel1.Location = new System.Drawing.Point(32, 16);
			this.ultraLabel1.Name = "ultraLabel1";
			this.ultraLabel1.Size = new System.Drawing.Size(71, 15);
			this.ultraLabel1.TabIndex = 2;
			this.ultraLabel1.Text = "Comprobante";
			// 
			// optTipoFact
			// 
			appearance2.ForeColorDisabled = System.Drawing.Color.Black;
			this.optTipoFact.Appearance = appearance2;
			this.optTipoFact.CheckedIndex = 1;
			this.optTipoFact.ItemAppearance = appearance3;
			valueListItem1.DataValue = 1;
			valueListItem1.DisplayText = "Venta";
			valueListItem2.DataValue = 4;
			valueListItem2.DisplayText = "Compra";
			this.optTipoFact.Items.Add(valueListItem1);
			this.optTipoFact.Items.Add(valueListItem2);
			this.optTipoFact.ItemSpacingHorizontal = 10;
			this.optTipoFact.ItemSpacingVertical = 10;
			this.optTipoFact.Location = new System.Drawing.Point(264, 8);
			this.optTipoFact.Name = "optTipoFact";
			this.optTipoFact.Size = new System.Drawing.Size(152, 34);
			this.optTipoFact.TabIndex = 3;
			this.optTipoFact.Text = "Compra";
			// 
			// errorProvider
			// 
			this.errorProvider.ContainerControl = this;
			// 
			// cdsSeteo
			// 
			this.cdsSeteo.BindingContextCtrl = this;
			this.cdsSeteo.DataLibrary = "LibFacturacion";
			this.cdsSeteo.DataLibraryUrl = "";
			this.cdsSeteo.DataSetDef = "dsSeteoF";
			this.cdsSeteo.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsSeteo.Name = "cdsSeteo";
			this.cdsSeteo.SchemaClassName = "LibFacturacion.DataClass";
			this.cdsSeteo.SchemaDef = null;
			this.cdsSeteo.BeforeFill += new C1.Data.FillEventHandler(this.cdsSeteo_BeforeFill_1);
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
			// btFiltro
			// 
			this.btFiltro.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btFiltro.Location = new System.Drawing.Point(560, 16);
			this.btFiltro.Name = "btFiltro";
			this.btFiltro.Size = new System.Drawing.Size(75, 25);
			this.btFiltro.TabIndex = 6;
			this.btFiltro.Text = "&Filtro";
			this.btFiltro.Click += new System.EventHandler(this.btFiltro_Click);
			// 
			// btFactura
			// 
			this.btFactura.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btFactura.Location = new System.Drawing.Point(560, 78);
			this.btFactura.Name = "btFactura";
			this.btFactura.Size = new System.Drawing.Size(75, 24);
			this.btFactura.TabIndex = 7;
			this.btFactura.Text = "&Factura";
			this.btFactura.Click += new System.EventHandler(this.btFactura_Click);
			// 
			// btAsiento
			// 
			this.btAsiento.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btAsiento.Location = new System.Drawing.Point(560, 121);
			this.btAsiento.Name = "btAsiento";
			this.btAsiento.Size = new System.Drawing.Size(75, 24);
			this.btAsiento.TabIndex = 8;
			this.btAsiento.Text = "&Asiento";
			this.btAsiento.Click += new System.EventHandler(this.btAsiento_Click);
			this.btAsiento.MouseDown += new System.Windows.Forms.MouseEventHandler(this.btAsiento_MouseDown);
			// 
			// cdvCompra
			// 
			this.cdvCompra.BindingContextCtrl = this;
			this.cdvCompra.DataSet = this.cdsCompra;
			this.cdvCompra.TableName = "";
			this.cdvCompra.TableViewName = "Compra";
			// 
			// btRetencion
			// 
			this.btRetencion.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btRetencion.Location = new System.Drawing.Point(560, 241);
			this.btRetencion.Name = "btRetencion";
			this.btRetencion.Size = new System.Drawing.Size(72, 35);
			this.btRetencion.TabIndex = 9;
			this.btRetencion.Text = "&Asi. Reten.";
			this.btRetencion.Click += new System.EventHandler(this.btRetencion_Click);
			// 
			// btIndividual
			// 
			this.btIndividual.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btIndividual.Location = new System.Drawing.Point(560, 164);
			this.btIndividual.Name = "btIndividual";
			this.btIndividual.Size = new System.Drawing.Size(75, 24);
			this.btIndividual.TabIndex = 10;
			this.btIndividual.Text = "&Individual";
			this.btIndividual.Click += new System.EventHandler(this.btIndividual_Click);
			// 
			// chkEgreso
			// 
			this.chkEgreso.Location = new System.Drawing.Point(440, 16);
			this.chkEgreso.Name = "chkEgreso";
			this.chkEgreso.Size = new System.Drawing.Size(72, 18);
			this.chkEgreso.TabIndex = 11;
			this.chkEgreso.Text = "Egreso";
			// 
			// grdCompras
			// 
			this.grdCompras.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			// 
			// gridControl1.EmbeddedNavigator
			// 
			this.grdCompras.EmbeddedNavigator.Name = "";
			this.grdCompras.Location = new System.Drawing.Point(16, 48);
			this.grdCompras.MainView = this.grvCompras;
			this.grdCompras.Name = "grdCompras";
			this.grdCompras.Size = new System.Drawing.Size(528, 224);
			this.grdCompras.TabIndex = 12;
			this.grdCompras.ViewCollection.AddRange(new DevExpress.XtraGrid.Views.Base.BaseView[] {
																																															this.grvCompras});
			// 
			// grvCompras
			// 
			this.grvCompras.GridControl = this.grdCompras;
			this.grvCompras.Name = "grvCompras";
			this.grvCompras.OptionsView.ShowAutoFilterRow = true;
			// 
			// AsientoGrupoComp
			// 
			this.AcceptButton = this.btAceptar;
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 14);
			this.ClientSize = new System.Drawing.Size(648, 286);
			this.Controls.Add(this.grdCompras);
			this.Controls.Add(this.chkEgreso);
			this.Controls.Add(this.btIndividual);
			this.Controls.Add(this.btRetencion);
			this.Controls.Add(this.btAsiento);
			this.Controls.Add(this.btFactura);
			this.Controls.Add(this.btFiltro);
			this.Controls.Add(this.optTipoFact);
			this.Controls.Add(this.ultraLabel1);
			this.Controls.Add(this.txtComprobante);
			this.Controls.Add(this.btAceptar);
			this.Name = "AsientoGrupoComp";
			this.Text = "Asiento por Comprobante";
			this.Closing += new System.ComponentModel.CancelEventHandler(this.AsientoGrupoComp_Closing);
			this.Load += new System.EventHandler(this.AsientoGrupoComp_Load);
			((System.ComponentModel.ISupportInitialize)(this.txtComprobante)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.optTipoFact)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsCompra)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvCompra)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.grdCompras)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.grvCompras)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void BorraErrores()
		{
			errorProvider.SetError(txtComprobante, "");
			errorProvider.SetError(grdCompras, "");
		}

		private bool Verifica()
		{
			bool bOk = true;
			BorraErrores();
			if (txtComprobante.Text.Trim().Length == 0)
			{
				errorProvider.SetError(txtComprobante, "Ingrese Comprobante");
				bOk = false;
			}
			if (grvCompras.RowCount < 1)
			{
				errorProvider.SetError(grdCompras, "No existen Facturas para generar asiento");
				bOk = false;
			}
			return bOk;
		}

		private void btAceptar_Click(object sender, System.EventArgs e)
		{
			if (!Verifica()) return;
			string stExec = "Exec AsientoComprobanteVerifica '" + txtComprobante.Text.Trim()
				+ "', " + optTipoFact.Value.ToString();
			string stMensaje = Funcion.sEscalarSQL(cdsCompra, stExec + ", 1", true);
			if (stMensaje.Length > 0)
			{
				MessageBox.Show(stMensaje, "Información");
				return;
			}

			int IdAsiento = Funcion.iEscalarSQL(cdsCompra, stExec + ", 2", true);
			string stBorrar = "Select Count(*) From Asiento Where "
				+ " Borrar = 0 And idAsiento = " + IdAsiento.ToString();
			int iCuenta = Funcion.iEscalarSQL(cdsCompra, stBorrar, true);
			if (iCuenta > 0)
			{
				if (DialogResult.No == 	MessageBox.Show("¿Asiento ya existe desea Sobreescribirlo?", 
					"Confirmación", MessageBoxButtons.YesNo, MessageBoxIcon.Question)) return;

				stBorrar = "Delete AsientoDetalle Where idAsiento = "
					+ IdAsiento.ToString();
				Funcion.EjecutaSQL(cdsCompra, stBorrar, true);
			}

			Cursor = Cursors.WaitCursor;
			int iTipo = (int) optTipoFact.Value;
			int iComprobante = 1;
			if (chkEgreso.Checked) iComprobante = 3;
			string stDesc = "Resumen de Comprobantes " + txtComprobante.Text;
      string stSelect = "Exec CreaAsientoCompraComp " + iTipo.ToString()
				+ ", '" + txtComprobante.Text + "', '" + stDesc + "', " + iComprobante.ToString();
			Funcion.EjecutaSQL(cdsSeteo, stSelect, true);
			Cursor = Cursors.Default;
			MessageBox.Show("Asiento Creado", "Información");
		}

		private void cdsSeteo_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteo.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);
		}

		private void cdsSeteo_BeforeFill_1(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteo.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);
		}

		private void btCancelar_Click(object sender, System.EventArgs e)
		{
			Close();
		}

		private void cdsCompra_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsCompra.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);		
		}

		private void btFactura_Click(object sender, System.EventArgs e)
		{
			int iFila = grvCompras.FocusedRowHandle;
			if (iFila < 0)
			{
				MessageBox.Show("Seleccione la línea que desa ver la Transacción", "Información");
				return;
			}
		
			Cursor = Cursors.WaitCursor;
			int IdCompra = (int) grvCompras.GetRowCellValue(iFila, "idCompra");
			if (IdCompra == 0)
			{
				MessageBox.Show("Fila no contiene Transacción Asociados",	"Información");
				Cursor = Cursors.Default;
				return;
			}

			Cursor = Cursors.WaitCursor;
			int IdTipo = (int) optTipoFact.Value;
//			Compra miCompra = new Compra(IdTipo, IdCompra);
//			miCompra.MdiParent = this.MdiParent;
//			miCompra.Show();
			Cursor = Cursors.Default;
		}

		private void FormatoGrilla()
		{
			DevExpress.XtraGrid.Views.Base.ColumnView View = grdCompras.MainView as DevExpress.XtraGrid.Views.Base.ColumnView;
			foreach (DevExpress.XtraGrid.Columns.GridColumn col in View.Columns)
			{
				col.OptionsColumn.AllowEdit = false;
				if (col.FieldName.StartsWith("id"))
					col.Visible = false;
				if (col.FieldName.StartsWith("Total") || col.FieldName.StartsWith("SubTotal")
					|| col.FieldName.StartsWith("Retenido") )
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

		private void btFiltro_Click(object sender, System.EventArgs e)
		{
			BorraErrores();
			if (txtComprobante.Text.Trim().Length == 0)
			{
				errorProvider.SetError(txtComprobante, "Ingrese Comprobante");
				return;
			}

			string stExec = string.Format("Exec CompraOtroFiltro {0}, '{1}'",
        optTipoFact.Value, txtComprobante.Text);
			grdCompras.DataSource = Funcion.dvProcedimiento(cdsCompra, stExec);
			FormatoGrilla();
			FormatoGrilla();
		}

		private void btRetencion_Click(object sender, System.EventArgs e)
		{
			if (!Verifica()) return;
			Cursor = Cursors.WaitCursor;
			int iTipo = (int) optTipoFact.Value;
			string stDesc = "Resumen de Comprobantes " + txtComprobante.Text;
			string stSelect = "Exec AsientoRetComprob " + iTipo.ToString()
				+ ", '" + txtComprobante.Text + "'";
			Funcion.EjecutaSQL(cdsSeteo, stSelect, true);
			Cursor = Cursors.Default;
			MessageBox.Show("Retenciòn Creada", "Información");
		}

		private void btAsiento_Click(object sender, System.EventArgs e)
		{
			int iFila = grvCompras.FocusedRowHandle;
			if (iFila < 0)
			{
				MessageBox.Show("Seleccione la línea que desa ver el Asiento", "Información");
				return;
			}
		
			Cursor = Cursors.WaitCursor;
			int IdAsi = (int) grvCompras.GetRowCellValue(iFila, "idAsiento");
			int IdAsiLote = grvCompras.GetRowCellValue(iFila, "idAsientoLote") == DBNull.Value ? 0 : 
				(int) grvCompras.GetRowCellValue(iFila, "idAsientoLote");
			int IdAsiPer = grvCompras.GetRowCellValue(iFila, "idAsientoPer") == DBNull.Value ? 0 : 
				(int) grvCompras.GetRowCellValue(iFila, "idAsientoPer");

			if (IdAsi == 0) IdAsi = IdAsiLote;
			if (IdAsi == 0) IdAsi = IdAsiPer;

			string stSelect = "Select Count(*) From Asiento Where "
				+ "idAsiento = " + IdAsi.ToString();
			int ICuenta = Funcion.iEscalarSQL(cdsCompra, stSelect, true);

			if (IdAsi == 0)
			{
				MessageBox.Show("No existe un asiento creado");
				return;
			}
			Cursor = Cursors.WaitCursor;
			Asiento miAsiento = new Asiento(IdAsi);
			miAsiento.MdiParent = this.MdiParent;
			miAsiento.Show();
			Cursor = Cursors.Default;
		}

		private int CreaAsiento()
		{
			int iTipo = 1;
			if (chkEgreso.Checked) iTipo = 3;
			string stExec = string.Format("Select Max(Fecha) From Compra Where idTipoFactura = {0} And Otro = '{1}'",
        optTipoFact.Value, txtComprobante.Text);
			DateTime dtFecha = Funcion.dtEscalarSQL(cdsSeteo, stExec);
			string stNumero = Contabilidad.NumeroAsiento(cdsCompra, iTipo);
			string stSelect = string.Format("INSERT INTO ASIENTO (IDTIPOASIENTO, CODASIENTO, DESCRIPCION, Fecha) "
				+ "VALUES ({0}, '{1}', 'Asiento de Comprobante {2} ', '{3}')"
				, iTipo, stNumero, txtComprobante.Text.Trim(), dtFecha.ToString("yyyyMMdd"));
			Funcion.EjecutaSQL(cdsCompra, stSelect, true);
			stSelect = "SELECT TOP 1 IDASIENTO FROM ASIENTO ORDER BY IDASIENTO DESC";
			int idAsiento = Funcion.iEscalarSQL(cdsCompra, stSelect, true);
			return idAsiento;
		}

		private void btIndividual_Click(object sender, System.EventArgs e)
		{
			if (!Verifica()) return;
			string stExec = "Exec AsientoComprobanteVerifica '" + txtComprobante.Text.Trim()
				+ "', " + optTipoFact.Value.ToString();
			string stMensaje = Funcion.sEscalarSQL(cdsCompra, stExec + ", 1", true);
			if (stMensaje.Length > 0)
			{
				MessageBox.Show(stMensaje, "Información");
				return;
			}
			int IdAsiento = Funcion.iEscalarSQL(cdsCompra, stExec + ", 2", true);

			string stBorrar = "Select Count(*) From Asiento Where "
				+ " Borrar = 0 And idAsiento = " + IdAsiento.ToString();
			int iCuenta = Funcion.iEscalarSQL(cdsCompra, stBorrar, true);
			if (iCuenta > 0)
			{
				if (DialogResult.No == 	MessageBox.Show("¿Asiento ya existe desea "
					+ "Sobreescribirlo?", "Confirmación", MessageBoxButtons.YesNo, MessageBoxIcon.Question)) return;

				stBorrar = "Delete AsientoDetalle Where idAsiento = "
					+ IdAsiento.ToString();
				Funcion.EjecutaSQL(cdsCompra, stBorrar, true);
			}
			else // Si esta anulado crea 
			{
				IdAsiento = CreaAsiento();
			}

			Cursor = Cursors.WaitCursor;
			for(int i = 0; i < grvCompras.RowCount; i++)
			{
				int IdCompra = (int) grvCompras.GetRowCellValue(i, "idCompra");

				string stProc = "AsientoIndividual ";
				if ((int) optTipoFact.Value == 1) stProc = "AsientoIndividualV ";
				string stGeneraCompra = "EXEC " + stProc + IdCompra.ToString()
					+ ", " + IdAsiento.ToString();
				Funcion.EjecutaSQL(cdsCompra, stGeneraCompra, true);
			}
			Funcion.EjecutaSQL(cdsCompra, "Exec AsientoTotal " + IdAsiento.ToString(), true);
			Cursor = Cursors.Default;
			btFiltro_Click(this, e);
			MessageBox.Show("Asiento Creado", "Confirmado");
		}

		private void AsientoGrupoComp_Load(object sender, System.EventArgs e)
		{
			string stAudita = "Exec AuditaCrear 53, 6, 'Asi. Comp.'";
			Funcion.EjecutaSQL(cdsCompra, stAudita, true);		
		}

		private void AsientoGrupoComp_Closing(object sender, System.ComponentModel.CancelEventArgs e)
		{
			string stAudita = "Exec AuditaCrear 53, 7, 'Asi. Comp.'";
			Funcion.EjecutaSQL(cdsCompra, stAudita, true);
		}

		private void btAsiento_MouseDown(object sender, System.Windows.Forms.MouseEventArgs e)
		{
			if (e.Button != MouseButtons.Right) return;

			if (txtComprobante.Text.Trim().Length == 0)
			{
				MessageBox.Show("Ingrese Comprobante",
					"Información", MessageBoxButtons.OK, MessageBoxIcon.Information);
				return;
			}
			if (DialogResult.No == MessageBox.Show(
				"¿Desea Quitar toda referencia del asiento creado por lotes?",
				"Información", MessageBoxButtons.YesNo, MessageBoxIcon.Question)) return;
			string stFiltro = "Otro = '" + txtComprobante.Text
				+ "' And idTipoFactura = " + optTipoFact.Value.ToString();
			string stSelect = "Update Compra Set idAsientoPer = 0, idAsiento = 0 Where " + stFiltro; 
			Funcion.EjecutaSQL(cdsCompra, stSelect, true);

			MessageBox.Show("Referencias Eliminadas",
				"Información", MessageBoxButtons.OK, MessageBoxIcon.Information);
		}
	}
}
