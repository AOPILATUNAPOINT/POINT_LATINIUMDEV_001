using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using DevExpress.XtraEditors;
using System.Globalization;

namespace Latinium
{
	/// <summary>
	/// Summary description for ConciliacionBusca.
	/// </summary>
	public class ConciliacionBusca : DevExpress.XtraEditors.XtraForm
	{
		private DevExpress.XtraEditors.Repository.RepositoryItemLookUpEdit repositoryItemLookUpEdit1;
		private DevExpress.XtraGrid.Views.Grid.GridView gridView4;
		private DevExpress.XtraEditors.SimpleButton btAceptar;
		private DevExpress.XtraEditors.SimpleButton btCancelar;
		private DevExpress.XtraEditors.CheckEdit chkCuenta;
		private DevExpress.XtraEditors.TextEdit txtNumero;
		private DevExpress.XtraEditors.CheckEdit chkNumero;
		private DevExpress.XtraEditors.LookUpEdit cmbCuenta;
		private DevExpress.XtraEditors.DateEdit cmbFecha;
		private DevExpress.XtraEditors.SpinEdit spnValor;
		private DevExpress.XtraEditors.CheckEdit chkFecha;
		private DevExpress.XtraEditors.CheckEdit chkValor;
		private C1.Data.C1DataSet cdsTmpConciliacion;
		private DevExpress.XtraEditors.SimpleButton btGrabar;
		private DevExpress.XtraEditors.SimpleButton btAsiento;
		private DevExpress.XtraGrid.GridControl grdBusca;
		private DevExpress.XtraGrid.Views.Grid.GridView grvBusca;
		/// <summary>
		/// Required designer variable.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public ConciliacionBusca(int idCuenta, string stCheque, double dTotal, DateTime dtFecha, 
			bool bTodoNumero, bool bTodoValor, bool bTodaFecha)
		{
			InitializeComponent();
			cmbCuenta.EditValue = idCuenta;
			txtNumero.Text = stCheque;
			spnValor.EditValue = dTotal;
			cmbFecha.EditValue = dtFecha;
			chkNumero.Checked = bTodoNumero;
			chkValor.Checked = bTodoValor;
			chkFecha.Checked = bTodaFecha;
		}

		/// <summary>
		/// Clean up any resources being used.
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
		/// Required method for Designer support - do not modify
		/// the contents of this method with the code editor.
		/// </summary>
		private void InitializeComponent()
		{
			this.grdBusca = new DevExpress.XtraGrid.GridControl();
			this.grvBusca = new DevExpress.XtraGrid.Views.Grid.GridView();
			this.repositoryItemLookUpEdit1 = new DevExpress.XtraEditors.Repository.RepositoryItemLookUpEdit();
			this.gridView4 = new DevExpress.XtraGrid.Views.Grid.GridView();
			this.btAceptar = new DevExpress.XtraEditors.SimpleButton();
			this.btCancelar = new DevExpress.XtraEditors.SimpleButton();
			this.cmbFecha = new DevExpress.XtraEditors.DateEdit();
			this.cmbCuenta = new DevExpress.XtraEditors.LookUpEdit();
			this.spnValor = new DevExpress.XtraEditors.SpinEdit();
			this.chkCuenta = new DevExpress.XtraEditors.CheckEdit();
			this.chkFecha = new DevExpress.XtraEditors.CheckEdit();
			this.txtNumero = new DevExpress.XtraEditors.TextEdit();
			this.chkNumero = new DevExpress.XtraEditors.CheckEdit();
			this.chkValor = new DevExpress.XtraEditors.CheckEdit();
			this.cdsTmpConciliacion = new C1.Data.C1DataSet();
			this.btGrabar = new DevExpress.XtraEditors.SimpleButton();
			this.btAsiento = new DevExpress.XtraEditors.SimpleButton();
			((System.ComponentModel.ISupportInitialize)(this.grdBusca)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.grvBusca)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.repositoryItemLookUpEdit1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.gridView4)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbFecha.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCuenta.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.spnValor.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.chkCuenta.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.chkFecha.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNumero.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.chkNumero.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.chkValor.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsTmpConciliacion)).BeginInit();
			this.SuspendLayout();
			// 
			// grdBusca
			// 
			this.grdBusca.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			// 
			// grdBalances.EmbeddedNavigator
			// 
			this.grdBusca.EmbeddedNavigator.Name = "";
			this.grdBusca.Location = new System.Drawing.Point(24, 56);
			this.grdBusca.MainView = this.grvBusca;
			this.grdBusca.Name = "grdBusca";
			this.grdBusca.RepositoryItems.AddRange(new DevExpress.XtraEditors.Repository.RepositoryItem[] {
																																																			this.repositoryItemLookUpEdit1});
			this.grdBusca.Size = new System.Drawing.Size(832, 288);
			this.grdBusca.TabIndex = 74;
			this.grdBusca.ViewCollection.AddRange(new DevExpress.XtraGrid.Views.Base.BaseView[] {
																																														this.grvBusca,
																																														this.gridView4});
			// 
			// grvBusca
			// 
			this.grvBusca.GridControl = this.grdBusca;
			this.grvBusca.Name = "grvBusca";
			this.grvBusca.OptionsSelection.MultiSelect = true;
			this.grvBusca.OptionsSelection.MultiSelectMode = DevExpress.XtraGrid.Views.Grid.GridMultiSelectMode.CellSelect;
			this.grvBusca.OptionsView.NewItemRowPosition = DevExpress.XtraGrid.Views.Grid.NewItemRowPosition.Bottom;
			this.grvBusca.OptionsView.ShowGroupPanel = false;
			// 
			// repositoryItemLookUpEdit1
			// 
			this.repositoryItemLookUpEdit1.AutoHeight = false;
			this.repositoryItemLookUpEdit1.Buttons.AddRange(new DevExpress.XtraEditors.Controls.EditorButton[] {
																																																					 new DevExpress.XtraEditors.Controls.EditorButton(DevExpress.XtraEditors.Controls.ButtonPredefines.Combo)});
			this.repositoryItemLookUpEdit1.Name = "repositoryItemLookUpEdit1";
			// 
			// gridView4
			// 
			this.gridView4.GridControl = this.grdBusca;
			this.gridView4.Name = "gridView4";
			// 
			// btAceptar
			// 
			this.btAceptar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btAceptar.Location = new System.Drawing.Point(56, 360);
			this.btAceptar.Name = "btAceptar";
			this.btAceptar.Size = new System.Drawing.Size(64, 23);
			this.btAceptar.TabIndex = 76;
			this.btAceptar.Text = "&Generar";
			this.btAceptar.Click += new System.EventHandler(this.btAceptar_Click);
			// 
			// btCancelar
			// 
			this.btCancelar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.btCancelar.DialogResult = System.Windows.Forms.DialogResult.Cancel;
			this.btCancelar.Location = new System.Drawing.Point(728, 360);
			this.btCancelar.Name = "btCancelar";
			this.btCancelar.Size = new System.Drawing.Size(64, 23);
			this.btCancelar.TabIndex = 77;
			this.btCancelar.Text = "&Cancelar";
			this.btCancelar.Click += new System.EventHandler(this.btCancelar_Click);
			// 
			// cmbFecha
			// 
			this.cmbFecha.EditValue = null;
			this.cmbFecha.Location = new System.Drawing.Point(752, 16);
			this.cmbFecha.Name = "cmbFecha";
			// 
			// cmbFecha.Properties
			// 
			this.cmbFecha.Properties.Buttons.AddRange(new DevExpress.XtraEditors.Controls.EditorButton[] {
																																																		 new DevExpress.XtraEditors.Controls.EditorButton(DevExpress.XtraEditors.Controls.ButtonPredefines.Combo)});
			this.cmbFecha.Properties.DisplayFormat.FormatString = "dd/MMM/yyyy";
			this.cmbFecha.Properties.DisplayFormat.FormatType = DevExpress.Utils.FormatType.DateTime;
			this.cmbFecha.Properties.EditFormat.FormatString = "dd/MMM/yyyy";
			this.cmbFecha.Properties.EditFormat.FormatType = DevExpress.Utils.FormatType.DateTime;
			this.cmbFecha.Size = new System.Drawing.Size(96, 20);
			this.cmbFecha.TabIndex = 78;
			// 
			// cmbCuenta
			// 
			this.cmbCuenta.Location = new System.Drawing.Point(120, 16);
			this.cmbCuenta.Name = "cmbCuenta";
			// 
			// cmbCuenta.Properties
			// 
			this.cmbCuenta.Properties.Buttons.AddRange(new DevExpress.XtraEditors.Controls.EditorButton[] {
																																																			new DevExpress.XtraEditors.Controls.EditorButton(DevExpress.XtraEditors.Controls.ButtonPredefines.Combo)});
			this.cmbCuenta.Properties.Columns.AddRange(new DevExpress.XtraEditors.Controls.LookUpColumnInfo[] {
																																																					new DevExpress.XtraEditors.Controls.LookUpColumnInfo("Codigo", "Código", 150),
																																																					new DevExpress.XtraEditors.Controls.LookUpColumnInfo("Descripcion", "Descripción", 300)});
			this.cmbCuenta.Properties.DisplayMember = "Codigo";
			this.cmbCuenta.Properties.ValueMember = "idCuenta";
			this.cmbCuenta.Size = new System.Drawing.Size(120, 20);
			this.cmbCuenta.TabIndex = 79;
			// 
			// spnValor
			// 
			this.spnValor.EditValue = new System.Decimal(new int[] {
																															 0,
																															 0,
																															 0,
																															 0});
			this.spnValor.Location = new System.Drawing.Point(552, 16);
			this.spnValor.Name = "spnValor";
			// 
			// spnValor.Properties
			// 
			this.spnValor.Properties.Buttons.AddRange(new DevExpress.XtraEditors.Controls.EditorButton[] {
																																																		 new DevExpress.XtraEditors.Controls.EditorButton()});
			this.spnValor.Size = new System.Drawing.Size(88, 20);
			this.spnValor.TabIndex = 80;
			// 
			// chkCuenta
			// 
			this.chkCuenta.Location = new System.Drawing.Point(24, 16);
			this.chkCuenta.Name = "chkCuenta";
			// 
			// chkCuenta.Properties
			// 
			this.chkCuenta.Properties.Caption = "Toda Cuenta";
			this.chkCuenta.Size = new System.Drawing.Size(88, 18);
			this.chkCuenta.TabIndex = 81;
			// 
			// chkFecha
			// 
			this.chkFecha.EditValue = true;
			this.chkFecha.Location = new System.Drawing.Point(664, 16);
			this.chkFecha.Name = "chkFecha";
			// 
			// chkFecha.Properties
			// 
			this.chkFecha.Properties.Caption = "Toda Fecha";
			this.chkFecha.Size = new System.Drawing.Size(80, 18);
			this.chkFecha.TabIndex = 82;
			// 
			// txtNumero
			// 
			this.txtNumero.Location = new System.Drawing.Point(352, 16);
			this.txtNumero.Name = "txtNumero";
			this.txtNumero.TabIndex = 83;
			// 
			// chkNumero
			// 
			this.chkNumero.Location = new System.Drawing.Point(256, 16);
			this.chkNumero.Name = "chkNumero";
			// 
			// chkNumero.Properties
			// 
			this.chkNumero.Properties.Caption = "Todo Número";
			this.chkNumero.Size = new System.Drawing.Size(88, 18);
			this.chkNumero.TabIndex = 84;
			// 
			// chkValor
			// 
			this.chkValor.Location = new System.Drawing.Point(472, 16);
			this.chkValor.Name = "chkValor";
			// 
			// chkValor.Properties
			// 
			this.chkValor.Properties.Caption = "Todo Valor";
			this.chkValor.Size = new System.Drawing.Size(72, 18);
			this.chkValor.TabIndex = 85;
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
			// 
			// btGrabar
			// 
			this.btGrabar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btGrabar.Location = new System.Drawing.Point(136, 360);
			this.btGrabar.Name = "btGrabar";
			this.btGrabar.Size = new System.Drawing.Size(64, 23);
			this.btGrabar.TabIndex = 86;
			this.btGrabar.Text = "G&rabar";
			this.btGrabar.Click += new System.EventHandler(this.btGrabar_Click);
			// 
			// btAsiento
			// 
			this.btAsiento.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btAsiento.Location = new System.Drawing.Point(216, 360);
			this.btAsiento.Name = "btAsiento";
			this.btAsiento.Size = new System.Drawing.Size(64, 23);
			this.btAsiento.TabIndex = 87;
			this.btAsiento.Text = "&Asiento";
			this.btAsiento.Click += new System.EventHandler(this.btAsiento_Click);
			// 
			// ConciliacionBusca
			// 
			this.AcceptButton = this.btAceptar;
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 14);
			this.CancelButton = this.btCancelar;
			this.ClientSize = new System.Drawing.Size(880, 398);
			this.Controls.Add(this.btAsiento);
			this.Controls.Add(this.btGrabar);
			this.Controls.Add(this.chkValor);
			this.Controls.Add(this.chkNumero);
			this.Controls.Add(this.txtNumero);
			this.Controls.Add(this.chkFecha);
			this.Controls.Add(this.chkCuenta);
			this.Controls.Add(this.spnValor);
			this.Controls.Add(this.cmbCuenta);
			this.Controls.Add(this.cmbFecha);
			this.Controls.Add(this.btCancelar);
			this.Controls.Add(this.btAceptar);
			this.Controls.Add(this.grdBusca);
			this.Name = "ConciliacionBusca";
			this.Text = "ConciliacionBusca";
			this.Load += new System.EventHandler(this.ConciliacionBusca_Load);
			((System.ComponentModel.ISupportInitialize)(this.grdBusca)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.grvBusca)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.repositoryItemLookUpEdit1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.gridView4)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbFecha.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCuenta.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.spnValor.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.chkCuenta.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.chkFecha.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNumero.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.chkNumero.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.chkValor.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsTmpConciliacion)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void btCancelar_Click(object sender, System.EventArgs e)
		{
			Close();
		}

		CultureInfo us = new CultureInfo("en-US");
		private void btAceptar_Click(object sender, System.EventArgs e)
		{
			int idCuenta = 0;
			string stCheque = "";
			double dTotal = 0;
			DateTime dtFecha = new DateTime(1901, 1, 1);
			if (!chkCuenta.Checked) idCuenta = (int) cmbCuenta.EditValue;
			if (!chkNumero.Checked) stCheque = txtNumero.Text;
			if (!chkValor.Checked) dTotal = double.Parse(spnValor.Text);
			if (!chkFecha.Checked) dtFecha = (DateTime) cmbFecha.EditValue;

			string stExec = string.Format("ConciliacionBusca {0}, '{1}', {2}, '{3}'",
				idCuenta, stCheque, dTotal.ToString("0.00", us), dtFecha.ToString("yyyyMMdd"));
			grdBusca.DataSource = Funcion.dvProcedimiento(cdsTmpConciliacion, stExec);
			GrillaFormato();
			GrillaFormato();
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
					|| col.FieldName.StartsWith("Debe") || col.FieldName.StartsWith("Haber"))
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


		private void ConciliacionBusca_Load(object sender, System.EventArgs e)
		{
			string stSelect = "Select idCuenta, Codigo, Descripcion From Cuenta Where Grupo = 0 Order by Codigo";
			cmbCuenta.Properties.DataSource = Funcion.dvProcedimiento(cdsTmpConciliacion, stSelect);
		}

		private void btAsiento_Click(object sender, System.EventArgs e)
		{
			int iFila = grvBusca.FocusedRowHandle;
			if (iFila < 0)
			{
				MessageBox.Show("Seleccione la fila que desea buscar",
					"Información", MessageBoxButtons.OK, MessageBoxIcon.Information);
				return;
			}
		
			int idAsiento = grvBusca.GetRowCellValue(iFila, "idAsiento") == null ? 0: (int) grvBusca.GetRowCellValue(iFila, "idAsiento");
			Asiento miAsiento = new Asiento(idAsiento);
			miAsiento.MdiParent = this.MdiParent;
			miAsiento.Show();
		}

		private void btGrabar_Click(object sender, System.EventArgs e)
		{
			int iFila = grvBusca.FocusedRowHandle;
			if (iFila < 0)
			{
				MessageBox.Show("Seleccione la fila que desea Grabar",
					"Información", MessageBoxButtons.OK, MessageBoxIcon.Information);
				return;
			}
			int idAsiento = grvBusca.GetRowCellValue(iFila, "idAsiento") == null ? 0: (int) grvBusca.GetRowCellValue(iFila, "idAsiento");
			int idDetAsiento = grvBusca.GetRowCellValue(iFila, "idDetAsiento") == null ? 0: 
				(int) grvBusca.GetRowCellValue(iFila, "idDetAsiento");
			string stCheque = grvBusca.GetRowCellValue(iFila, "Documento") == null ? "": 
				(string) grvBusca.GetRowCellValue(iFila, "Documento");
			bool bConfirmado = grvBusca.GetRowCellValue(iFila, "Confirmado") == null ? false: 
				(bool) grvBusca.GetRowCellValue(iFila, "Confirmado");
			DateTime dtFecha = grvBusca.GetRowCellValue(iFila, "FechaConf") == null ? DateTime.Today: 
				(DateTime) grvBusca.GetRowCellValue(iFila, "FechaConf");

			string stExec = string.Format("Exec ConciliacionGrabar {0}, {1}, '{2}', {3}, '{4}'",
				idAsiento, idDetAsiento, stCheque, bConfirmado, dtFecha.ToString("yyyyMMdd"));
			Funcion.EjecutaSQL(cdsTmpConciliacion, stExec);
			MessageBox.Show("Registro Grabado",
				"Información", MessageBoxButtons.OK, MessageBoxIcon.Information);
		}
	}
}

