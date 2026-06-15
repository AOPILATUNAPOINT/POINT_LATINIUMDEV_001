using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using C1.C1Excel;
using System.Globalization;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de EstadoCuentaAuto.
	/// </summary>
	public class EstadoCuentaAuto : DevExpress.XtraEditors.XtraForm
	{
		private Infragistics.Win.Misc.UltraButton btLeer;
		private Infragistics.Win.Misc.UltraButton btConciliar;
		private C1.Data.C1DataSet cdsTmpConciliacion;
		private C1.C1Excel.C1XLBook c1XLBook1;
		private System.Windows.Forms.OpenFileDialog openFileDialog;
		private Infragistics.Win.Misc.UltraButton btAsiento;
		private System.Windows.Forms.ToolTip toolTip1;
		private Infragistics.Win.Misc.UltraButton btImprimir;
		private Infragistics.Win.Misc.UltraButton btIngreso;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbTipo;
		private DevExpress.XtraGrid.GridControl grdBusca;
		private DevExpress.XtraGrid.Views.Grid.GridView grvBusca;
		private DevExpress.XtraGrid.Views.Grid.GridView gridView4;
		private System.ComponentModel.IContainer components;

		public EstadoCuentaAuto()
		{
			InitializeComponent();
		}

		private int IdCuentaCorr = 0;
		public EstadoCuentaAuto(int idCuentaCorr)
		{
			InitializeComponent();
			IdCuentaCorr = idCuentaCorr;
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
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("AsientoTipo", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idTipoAsiento");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Tipo");
			this.cdsTmpConciliacion = new C1.Data.C1DataSet();
			this.btLeer = new Infragistics.Win.Misc.UltraButton();
			this.btConciliar = new Infragistics.Win.Misc.UltraButton();
			this.c1XLBook1 = new C1.C1Excel.C1XLBook();
			this.openFileDialog = new System.Windows.Forms.OpenFileDialog();
			this.btAsiento = new Infragistics.Win.Misc.UltraButton();
			this.toolTip1 = new System.Windows.Forms.ToolTip(this.components);
			this.cmbTipo = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.btImprimir = new Infragistics.Win.Misc.UltraButton();
			this.btIngreso = new Infragistics.Win.Misc.UltraButton();
			this.grdBusca = new DevExpress.XtraGrid.GridControl();
			this.grvBusca = new DevExpress.XtraGrid.Views.Grid.GridView();
			this.gridView4 = new DevExpress.XtraGrid.Views.Grid.GridView();
			((System.ComponentModel.ISupportInitialize)(this.cdsTmpConciliacion)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbTipo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.grdBusca)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.grvBusca)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.gridView4)).BeginInit();
			this.SuspendLayout();
			// 
			// cdsTmpConciliacion
			// 
			this.cdsTmpConciliacion.BindingContextCtrl = this;
			this.cdsTmpConciliacion.DataLibrary = "LibContabilidad";
			this.cdsTmpConciliacion.DataLibraryUrl = "";
			this.cdsTmpConciliacion.DataSetDef = "dsTmpConcilia";
			this.cdsTmpConciliacion.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsTmpConciliacion.Name = "cdsTmpConciliacion";
			this.cdsTmpConciliacion.SchemaClassName = "LibContabilidad.DataClass";
			this.cdsTmpConciliacion.SchemaDef = null;
			this.cdsTmpConciliacion.BeforeFill += new C1.Data.FillEventHandler(this.cdsTmpConciliacion_BeforeFill);
			// 
			// btLeer
			// 
			this.btLeer.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btLeer.Location = new System.Drawing.Point(24, 344);
			this.btLeer.Name = "btLeer";
			this.btLeer.Size = new System.Drawing.Size(64, 25);
			this.btLeer.TabIndex = 1;
			this.btLeer.Text = "Leer";
			this.toolTip1.SetToolTip(this.btLeer, "Fecha, Referencia, Detalle, Valor");
			this.btLeer.Click += new System.EventHandler(this.btLeer_Click);
			// 
			// btConciliar
			// 
			this.btConciliar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btConciliar.Location = new System.Drawing.Point(104, 344);
			this.btConciliar.Name = "btConciliar";
			this.btConciliar.Size = new System.Drawing.Size(64, 25);
			this.btConciliar.TabIndex = 2;
			this.btConciliar.Text = "Conciliar";
			this.btConciliar.Click += new System.EventHandler(this.btConciliar_Click);
			this.btConciliar.MouseUp += new System.Windows.Forms.MouseEventHandler(this.btConciliar_MouseUp);
			// 
			// btAsiento
			// 
			this.btAsiento.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btAsiento.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.btAsiento.Location = new System.Drawing.Point(176, 344);
			this.btAsiento.Name = "btAsiento";
			this.btAsiento.Size = new System.Drawing.Size(64, 25);
			this.btAsiento.TabIndex = 17;
			this.btAsiento.Text = "Asiento";
			this.btAsiento.Click += new System.EventHandler(this.btAsiento_Click);
			this.btAsiento.MouseUp += new System.Windows.Forms.MouseEventHandler(this.btAsiento_MouseUp);
			// 
			// cmbTipo
			// 
			this.cmbTipo.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			appearance1.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbTipo.Appearance = appearance1;
			this.cmbTipo.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbTipo.Cursor = System.Windows.Forms.Cursors.Default;
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridBand1.Columns.AddRange(new object[] {
																										 ultraGridColumn1,
																										 ultraGridColumn2});
			this.cmbTipo.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.cmbTipo.DisplayMember = "Tipo";
			this.cmbTipo.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbTipo.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbTipo.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbTipo.Location = new System.Drawing.Point(288, 344);
			this.cmbTipo.Name = "cmbTipo";
			this.cmbTipo.Size = new System.Drawing.Size(104, 21);
			this.cmbTipo.TabIndex = 20;
			this.toolTip1.SetToolTip(this.cmbTipo, "Tipo de Asiento");
			this.cmbTipo.ValueMember = "idTipoAsiento";
			// 
			// btImprimir
			// 
			this.btImprimir.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btImprimir.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.btImprimir.Location = new System.Drawing.Point(592, 344);
			this.btImprimir.Name = "btImprimir";
			this.btImprimir.Size = new System.Drawing.Size(64, 25);
			this.btImprimir.TabIndex = 18;
			this.btImprimir.Text = "Imprimir";
			this.btImprimir.Click += new System.EventHandler(this.btImprimir_Click);
			// 
			// btIngreso
			// 
			this.btIngreso.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btIngreso.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.btIngreso.Location = new System.Drawing.Point(392, 344);
			this.btIngreso.Name = "btIngreso";
			this.btIngreso.Size = new System.Drawing.Size(48, 25);
			this.btIngreso.TabIndex = 19;
			this.btIngreso.Text = "Crear";
			this.btIngreso.Click += new System.EventHandler(this.btIngreso_Click);
			// 
			// grdBusca
			// 
			this.grdBusca.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			// 
			// grdBusca.EmbeddedNavigator
			// 
			this.grdBusca.EmbeddedNavigator.Name = "";
			this.grdBusca.Location = new System.Drawing.Point(16, 8);
			this.grdBusca.MainView = this.grvBusca;
			this.grdBusca.Name = "grdBusca";
			this.grdBusca.Size = new System.Drawing.Size(664, 320);
			this.grdBusca.TabIndex = 67;
			this.grdBusca.ViewCollection.AddRange(new DevExpress.XtraGrid.Views.Base.BaseView[] {
																																														this.grvBusca,
																																														this.gridView4});
			this.grdBusca.DoubleClick += new System.EventHandler(this.grdBusca_DoubleClick);
			// 
			// grvBusca
			// 
			this.grvBusca.GridControl = this.grdBusca;
			this.grvBusca.Name = "grvBusca";
			this.grvBusca.OptionsBehavior.Editable = false;
			this.grvBusca.OptionsSelection.MultiSelect = true;
			this.grvBusca.OptionsSelection.MultiSelectMode = DevExpress.XtraGrid.Views.Grid.GridMultiSelectMode.CellSelect;
			this.grvBusca.OptionsView.ShowAutoFilterRow = true;
			// 
			// gridView4
			// 
			this.gridView4.GridControl = this.grdBusca;
			this.gridView4.Name = "gridView4";
			// 
			// EstadoCuentaAuto
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 14);
			this.ClientSize = new System.Drawing.Size(696, 394);
			this.Controls.Add(this.grdBusca);
			this.Controls.Add(this.cmbTipo);
			this.Controls.Add(this.btIngreso);
			this.Controls.Add(this.btImprimir);
			this.Controls.Add(this.btAsiento);
			this.Controls.Add(this.btConciliar);
			this.Controls.Add(this.btLeer);
			this.Name = "EstadoCuentaAuto";
			this.Text = "Estado de Cuenta Automático";
			this.Closing += new System.ComponentModel.CancelEventHandler(this.EstadoCuentaAuto_Closing);
			this.Load += new System.EventHandler(this.EstadoCuentaAuto_Load);
			((System.ComponentModel.ISupportInitialize)(this.cdsTmpConciliacion)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbTipo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.grdBusca)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.grvBusca)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.gridView4)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void cdsTmpConciliacion_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			this.cdsTmpConciliacion.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirConta);
		}

		public void GrillaFormato()
		{
			DevExpress.XtraGrid.Views.Base.ColumnView View = 
				grdBusca.MainView as DevExpress.XtraGrid.Views.Base.ColumnView;
			foreach (DevExpress.XtraGrid.Columns.GridColumn col in View.Columns)
			{
				if (col.FieldName.StartsWith("id"))
					col.Visible = false;
				if (col.FieldName.StartsWith("Total") || col.FieldName.StartsWith("Valor")
					|| col.FieldName.StartsWith("Monto") || col.FieldName.StartsWith("Iva"))
				{
					col.DisplayFormat.FormatString = "n2";
					col.DisplayFormat.FormatType = DevExpress.Utils.FormatType.Numeric;
				}
				if (col.FieldName.StartsWith("Fecha") || col.FieldName.StartsWith("Vencimiento"))
				{
					col.DisplayFormat.FormatString = "dd/MMM/yyyy";
					col.DisplayFormat.FormatType = DevExpress.Utils.FormatType.DateTime;
				}			
			}
		}


		CultureInfo us = new CultureInfo("en-US");
		private void btLeer_Click(object sender, System.EventArgs e)
		{
			string stArchivo = "";
			this.openFileDialog.InitialDirectory = Application.ExecutablePath;
			if (DialogResult.OK == this.openFileDialog.ShowDialog())
			{
				stArchivo = this.openFileDialog.FileName.ToString();
			}
			else
			{
				MessageBox.Show("Cancelado por el usuario");
				return;
			}
			this.Cursor = Cursors.WaitCursor;

			C1.C1Excel.C1XLBook miConcilia = new C1.C1Excel.C1XLBook();
			try
			{
				miConcilia.Load(@stArchivo);
			}
			catch(System.Exception ex)
			{
				MessageBox.Show(ex.Message, "Error Abriendo el Archivo");
				this.Cursor = Cursors.Default;
				return;
			}
			XLSheet sheet = miConcilia.Sheets[0];

			string stSelect = "Delete tmpConciliacion";
			Funcion.EjecutaSQL(cdsTmpConciliacion, stSelect, true);
			for (int i = 1; i < sheet.Rows.Count; i++)
			{
				if (sheet.GetCell(i, 0) == null || sheet.GetCell(i, 0).Value ==  null
					|| sheet.GetCell(i, 1) == null || sheet.GetCell(i, 1).Value ==  null
					|| sheet.GetCell(i, 2) == null || sheet.GetCell(i, 2).Value ==  null
					|| sheet.GetCell(i, 3) == null || sheet.GetCell(i, 3).Value ==  null) 
							continue;
				bool bOk = false;
				stSelect = "insert into tmpConciliacion (idCuenta, ";
				string stValores = " Values (" + IdCuentaCorr + ", ";
				if (sheet.GetCell(i, 0).Value.GetType().Name == "DateTime")
				{
					DateTime dtFecha = (DateTime) sheet.GetCell(i, 0).Value;
					stSelect += "Fecha";
					stValores += "'" + dtFecha.ToString("yyyyMMdd") + "'";
					bOk = true;
				}
				if (sheet.GetCell(i, 1).Value.GetType().Name == "String" || sheet.GetCell(i, 1).Value.GetType().Name == "Double")
				{
					string stComa = "";
					if (bOk) stComa = ", ";
					string stDetalle = sheet.GetCell(i, 1).Value.ToString().Trim();
					stSelect += stComa + "Detalle";
					stValores += stComa + "'" + stDetalle + "'";
					bOk = true;
				}
				
				if (sheet.GetCell(i, 2).Value.GetType().Name == "String" || sheet.GetCell(i, 2).Value.GetType().Name == "Double")
				{
					string stComa = "";
					if (bOk) stComa = ", ";
					string stDetalle = sheet.GetCell(i, 2).Value.ToString().Trim();
					stSelect += stComa + "Referencia";
					stValores += stComa + "'" + stDetalle + "'";
					bOk = true;
				}

				if (sheet.GetCell(i, 3).Value.GetType().Name == "Double")
				{
					string stComa = "";
					if (bOk) stComa = ", ";
					double dValor = (double) sheet.GetCell(i, 3).Value;
					stSelect += stComa + "Valor";
					stValores += stComa + dValor.ToString("0.00", us).Trim();
					bOk = true;
				}

				stSelect += ")" + stValores + ")";
				if (bOk) Funcion.EjecutaSQL(cdsTmpConciliacion, stSelect, true);
			}
      LeerConciliacion();
			this.Cursor = Cursors.Default;
			MessageBox.Show("Archivo de Excel Subido", "Información");
		}

		private void EstadoCuentaAuto_Load(object sender, System.EventArgs e)
		{
			string stAudita = "Exec AuditaCrear 53, 6, 'Est. C. Auto'";
			Funcion.EjecutaSQL(cdsTmpConciliacion, stAudita, true);
			cmbTipo.DataSource = Funcion.dvProcedimiento(cdsTmpConciliacion, "Select idTipoAsiento, Tipo From AsientoTipo Order by idTipoAsiento");
			LeerConciliacion();
		}

		private void LeerConciliacion()
		{
			grdBusca.DataSource = Funcion.dvProcedimiento(cdsTmpConciliacion, 
				"Select Fecha, Detalle as Concepto, Referencia as Documento, Valor as Monto, Errores, idAsiento "
				+ "From TmpConciliacion Where idCuenta = " + IdCuentaCorr);
			GrillaFormato();
			GrillaFormato();
		}
		private void btConciliar_Click(object sender, System.EventArgs e)
		{
			this.Cursor = Cursors.WaitCursor;
			string stSelect = "Exec ConciliacionAutomatica " + IdCuentaCorr.ToString();
			Funcion.EjecutaSQL(cdsTmpConciliacion, stSelect);
			LeerConciliacion();
			this.Cursor = Cursors.Default;
			MessageBox.Show("Conciliacion Automatica generada", "Información");
		}

		private void EstadoCuentaAuto_Closing(object sender, System.ComponentModel.CancelEventArgs e)
		{
			string stAudita = "Exec AuditaCrear 53, 7, 'Est. C. Auto'";
			Funcion.EjecutaSQL(cdsTmpConciliacion, stAudita, true);
		}

		private void btAsiento_Click(object sender, System.EventArgs e)
		{
			int iFila = grvBusca.FocusedRowHandle;
			if (iFila < 0)
			{
				MessageBox.Show("No existen ningun registro en contabilidad con esa Referencia",
					"Información", MessageBoxButtons.OK, MessageBoxIcon.Information);
				return;
			}
		
			int IdAsiento = grvBusca.GetRowCellValue(iFila, "idAsiento") == null ? 0: (int) grvBusca.GetRowCellValue(iFila, "idAsiento");
			if (IdAsiento == 0)
			{
				MessageBox.Show("Fila no contiene Asientos",	"Información");
				return;
			}

			if (IdAsiento == -1) 
			{
				MessageBox.Show("Existen varios registro en contabilidad con esa Referencia",
					"Información", MessageBoxButtons.OK, MessageBoxIcon.Information);
				return;
			}
			Asiento miAsiento = new Asiento(IdAsiento);
			miAsiento.MdiParent = this.MdiParent;
			miAsiento.Show();
		}

		private void btImprimir_Click(object sender, System.EventArgs e)
		{
			Cursor = Cursors.WaitCursor;
			Reporte miRep = new Reporte("ConciliacionAuto.rpt", "");
			miRep.MdiParent = this.MdiParent;
			Cursor = Cursors.Default;
			miRep.Show();
		}

		private void btIngreso_Click(object sender, System.EventArgs e)
		{
			if (grvBusca.SelectedRowsCount == 0)
			{
				MessageBox.Show("Seleccione las filas con la que creará los asiento",
					"Información", MessageBoxButtons.OK, MessageBoxIcon.Information);
				return;
			}
			if (cmbTipo.Text.Length == 0)
			{
				MessageBox.Show("Ingrese tipo de asiento", "Informacion");
				return;
			}
			int idTipoAsiento = (int) cmbTipo.Value;
			string stMensaje = string.Format("Desea Generar {0} {1}",
				grvBusca.SelectedRowsCount, cmbTipo.Text);
			if (DialogResult.No == MessageBox.Show(stMensaje, "Confirmación",
				MessageBoxButtons.YesNo, MessageBoxIcon.Question)) return;

			string stResultado = "";
			foreach(int iFila in grvBusca.GetSelectedRows())
			{
				double dValor = (double) grvBusca.GetRowCellValue(iFila, "Monto");
				string stReferencia = (string) grvBusca.GetRowCellValue(iFila, "Documento");
				string stDetalle = (string) grvBusca.GetRowCellValue(iFila, "Concepto");
				DateTime dtFecha = (DateTime) grvBusca.GetRowCellValue(iFila, "Fecha");
				string stExec = string.Format("Exec AsientoGeneraTodo {0}, '{1}', '{2}', '{3}', {4}, {5}", 
					idTipoAsiento, dtFecha.ToString("yyyyMMdd"), stDetalle.Trim(), stReferencia.Trim(), 
					dValor.ToString("0.00", us), IdCuentaCorr);
				stResultado += Funcion.sEscalarSQL(cdsTmpConciliacion, stExec) + "\n";
			}
			if (stResultado.Length == 0) stResultado = "Asiento Generado";
			MessageBox.Show(stResultado, "Información", MessageBoxButtons.OK, MessageBoxIcon.Information);
		}

		private void btConciliar_MouseUp(object sender, System.Windows.Forms.MouseEventArgs e)
		{
			if (e.Button != MouseButtons.Right) return;
			string stSelect = "Exec ConciliacionAutomatica2 " + IdCuentaCorr.ToString();
			grdBusca.DataSource = Funcion.dvProcedimiento(cdsTmpConciliacion, stSelect);
			GrillaFormato();
			GrillaFormato();
			MessageBox.Show("Conciliación desde Contabilidad Generada", "Información");
		}

		private void grdBusca_DoubleClick(object sender, System.EventArgs e)
		{
			btAsiento_Click(this, e);
		}

		private void btAsiento_MouseUp(object sender, System.Windows.Forms.MouseEventArgs e)
		{
			if (e.Button != MouseButtons.Right) return;

			int iFila = grvBusca.FocusedRowHandle;
			if (iFila < 0)
			{
				MessageBox.Show("Seleccione la fila que desea buscar",
					"Información", MessageBoxButtons.OK, MessageBoxIcon.Information);
				return;
			}
		
			string stCheque = grvBusca.GetRowCellValue(iFila, "Documento") == null ? "": (string) grvBusca.GetRowCellValue(iFila, "Documento");
			double dTotal = grvBusca.GetRowCellValue(iFila, "Monto") == null ? 0: (double) grvBusca.GetRowCellValue(iFila, "Monto");
			DateTime dtFecha = grvBusca.GetRowCellValue(iFila, "Fecha") == null ? DateTime.Today: (DateTime) grvBusca.GetRowCellValue(iFila, "Fecha");
			bool bTodoNumero = true;
			bool bTodoValor = true;
			bool bTodaFecha = true;

			if (grvBusca.FocusedColumn.Caption == "Documento") bTodoNumero = false;
			if (grvBusca.FocusedColumn.Caption == "Monto") bTodoValor = false;
//			if (grvBusca.FocusedColumn.Caption == "Fecha") bTodaFecha = false;

			ConciliacionBusca miBusqueda = new ConciliacionBusca(IdCuentaCorr, stCheque, dTotal, dtFecha, 
				bTodoNumero, bTodoValor, bTodaFecha);
			miBusqueda.MdiParent = this.MdiParent;
			miBusqueda.Show();
		}
	}
}
