using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using DevExpress.XtraEditors;

namespace Latinium
{
	/// <summary>
	/// Summary description for FlujoCaja.
	/// </summary>
	public class FlujoCaja : DevExpress.XtraEditors.XtraForm
	{
		private DevExpress.XtraEditors.LabelControl labelControl1;
		private C1.Data.C1DataSet cdsFlujoCaja;
		private C1.Data.C1DataSet cdsCuenta;
		private DevExpress.XtraEditors.SimpleButton btGrabar;
		private C1.Data.C1DataView cdvCuenta;
		private DevExpress.XtraEditors.SpinEdit spnAnio;
		private DevExpress.XtraGrid.GridControl gridControl2;
		private DevExpress.XtraGrid.Columns.GridColumn gridColumn1;
		private DevExpress.XtraGrid.Columns.GridColumn gridColumn2;
		private DevExpress.XtraEditors.Repository.RepositoryItemLookUpEdit repositoryItemLookUpEdit2;
		private DevExpress.XtraGrid.Columns.GridColumn gridColumn3;
		private DevExpress.XtraGrid.Columns.GridColumn gridColumn4;
		private DevExpress.XtraGrid.Columns.GridColumn gridColumn5;
		private DevExpress.XtraGrid.Columns.GridColumn gridColumn6;
		private DevExpress.XtraGrid.Columns.GridColumn gridColumn7;
		private DevExpress.XtraGrid.Columns.GridColumn gridColumn8;
		private DevExpress.XtraGrid.Columns.GridColumn gridColumn9;
		private DevExpress.XtraGrid.Columns.GridColumn gridColumn10;
		private DevExpress.XtraGrid.Columns.GridColumn gridColumn11;
		private DevExpress.XtraGrid.Columns.GridColumn gridColumn12;
		private DevExpress.XtraGrid.Columns.GridColumn gridColumn13;
		private DevExpress.XtraGrid.Columns.GridColumn gridColumn14;
		private DevExpress.XtraGrid.Columns.GridColumn gridColumn15;
		private C1.Data.C1DataView cdvFlujoCaja;
		private DevExpress.XtraGrid.Views.Grid.GridView gridView1;
		private DevExpress.XtraEditors.LookUpEdit cmbCentro;
		private DevExpress.XtraEditors.SimpleButton btCopiar;
		private DevExpress.XtraEditors.ComboBoxEdit cmbTipo;
		/// <summary>
		/// Required designer variable.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public FlujoCaja()
		{
			//
			// Required for Windows Form Designer support
			//
			InitializeComponent();

			//
			// TODO: Add any constructor code after InitializeComponent call
			//
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
			this.cdsFlujoCaja = new C1.Data.C1DataSet();
			this.cdvCuenta = new C1.Data.C1DataView();
			this.cdsCuenta = new C1.Data.C1DataSet();
			this.spnAnio = new DevExpress.XtraEditors.SpinEdit();
			this.labelControl1 = new DevExpress.XtraEditors.LabelControl();
			this.btGrabar = new DevExpress.XtraEditors.SimpleButton();
			this.gridControl2 = new DevExpress.XtraGrid.GridControl();
			this.gridView1 = new DevExpress.XtraGrid.Views.Grid.GridView();
			this.gridColumn1 = new DevExpress.XtraGrid.Columns.GridColumn();
			this.gridColumn2 = new DevExpress.XtraGrid.Columns.GridColumn();
			this.repositoryItemLookUpEdit2 = new DevExpress.XtraEditors.Repository.RepositoryItemLookUpEdit();
			this.gridColumn3 = new DevExpress.XtraGrid.Columns.GridColumn();
			this.gridColumn4 = new DevExpress.XtraGrid.Columns.GridColumn();
			this.gridColumn5 = new DevExpress.XtraGrid.Columns.GridColumn();
			this.gridColumn6 = new DevExpress.XtraGrid.Columns.GridColumn();
			this.gridColumn7 = new DevExpress.XtraGrid.Columns.GridColumn();
			this.gridColumn8 = new DevExpress.XtraGrid.Columns.GridColumn();
			this.gridColumn9 = new DevExpress.XtraGrid.Columns.GridColumn();
			this.gridColumn10 = new DevExpress.XtraGrid.Columns.GridColumn();
			this.gridColumn11 = new DevExpress.XtraGrid.Columns.GridColumn();
			this.gridColumn12 = new DevExpress.XtraGrid.Columns.GridColumn();
			this.gridColumn13 = new DevExpress.XtraGrid.Columns.GridColumn();
			this.gridColumn14 = new DevExpress.XtraGrid.Columns.GridColumn();
			this.gridColumn15 = new DevExpress.XtraGrid.Columns.GridColumn();
			this.cdvFlujoCaja = new C1.Data.C1DataView();
			this.cmbCentro = new DevExpress.XtraEditors.LookUpEdit();
			this.cmbTipo = new DevExpress.XtraEditors.ComboBoxEdit();
			this.btCopiar = new DevExpress.XtraEditors.SimpleButton();
			((System.ComponentModel.ISupportInitialize)(this.cdsFlujoCaja)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvCuenta)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsCuenta)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.spnAnio.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.gridControl2)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.gridView1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.repositoryItemLookUpEdit2)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvFlujoCaja)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCentro.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbTipo.Properties)).BeginInit();
			this.SuspendLayout();
			// 
			// cdsFlujoCaja
			// 
			this.cdsFlujoCaja.BindingContextCtrl = this;
			this.cdsFlujoCaja.DataLibrary = "LibContabilidad";
			this.cdsFlujoCaja.DataLibraryUrl = "";
			this.cdsFlujoCaja.DataSetDef = "dsFlujoDeCaja";
			this.cdsFlujoCaja.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsFlujoCaja.Name = "cdsFlujoCaja";
			this.cdsFlujoCaja.SchemaClassName = "LibContabilidad.DataClass";
			this.cdsFlujoCaja.SchemaDef = null;
			this.cdsFlujoCaja.BeforeFill += new C1.Data.FillEventHandler(this.cdsFlujoCaja_BeforeFill);
			// 
			// cdvCuenta
			// 
			this.cdvCuenta.BindingContextCtrl = this;
			this.cdvCuenta.DataSet = this.cdsCuenta;
			this.cdvCuenta.Sort = "Codigo";
			this.cdvCuenta.TableName = "";
			this.cdvCuenta.TableViewName = "Cuenta";
			// 
			// cdsCuenta
			// 
			this.cdsCuenta.BindingContextCtrl = this;
			this.cdsCuenta.DataLibrary = "LibContabilidad";
			this.cdsCuenta.DataLibraryUrl = "";
			this.cdsCuenta.DataSetDef = "dsCuenta2";
			this.cdsCuenta.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsCuenta.Name = "cdsCuenta";
			this.cdsCuenta.SchemaClassName = "LibContabilidad.DataClass";
			this.cdsCuenta.SchemaDef = null;
			// 
			// spnAnio
			// 
			this.spnAnio.EditValue = new System.Decimal(new int[] {
																															0,
																															0,
																															0,
																															0});
			this.spnAnio.Location = new System.Drawing.Point(56, 16);
			this.spnAnio.Name = "spnAnio";
			// 
			// spnAnio.Properties
			// 
			this.spnAnio.Properties.Buttons.AddRange(new DevExpress.XtraEditors.Controls.EditorButton[] {
																																																		new DevExpress.XtraEditors.Controls.EditorButton()});
			this.spnAnio.Size = new System.Drawing.Size(64, 20);
			this.spnAnio.TabIndex = 1;
			this.spnAnio.EditValueChanged += new System.EventHandler(this.spnAnio_EditValueChanged);
			// 
			// labelControl1
			// 
			this.labelControl1.AutoSizeMode = DevExpress.XtraEditors.LabelAutoSizeMode.Horizontal;
			this.labelControl1.Location = new System.Drawing.Point(24, 16);
			this.labelControl1.Name = "labelControl1";
			this.labelControl1.Size = new System.Drawing.Size(19, 13);
			this.labelControl1.TabIndex = 2;
			this.labelControl1.Text = "Año";
			// 
			// btGrabar
			// 
			this.btGrabar.Location = new System.Drawing.Point(848, 16);
			this.btGrabar.Name = "btGrabar";
			this.btGrabar.Size = new System.Drawing.Size(72, 24);
			this.btGrabar.TabIndex = 3;
			this.btGrabar.Text = "Grabar";
			this.btGrabar.Click += new System.EventHandler(this.btGrabar_Click);
			// 
			// gridControl2
			// 
			this.gridControl2.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.gridControl2.DataSource = this.cdvFlujoCaja;
			// 
			// gridControl2.EmbeddedNavigator
			// 
			this.gridControl2.EmbeddedNavigator.Name = "";
			this.gridControl2.Location = new System.Drawing.Point(16, 48);
			this.gridControl2.MainView = this.gridView1;
			this.gridControl2.Name = "gridControl2";
			this.gridControl2.RepositoryItems.AddRange(new DevExpress.XtraEditors.Repository.RepositoryItem[] {
																																																					this.repositoryItemLookUpEdit2});
			this.gridControl2.Size = new System.Drawing.Size(912, 368);
			this.gridControl2.TabIndex = 4;
			this.gridControl2.ViewCollection.AddRange(new DevExpress.XtraGrid.Views.Base.BaseView[] {
																																																this.gridView1});
			// 
			// gridView1
			// 
			this.gridView1.Columns.AddRange(new DevExpress.XtraGrid.Columns.GridColumn[] {
																																										 this.gridColumn1,
																																										 this.gridColumn2,
																																										 this.gridColumn3,
																																										 this.gridColumn4,
																																										 this.gridColumn5,
																																										 this.gridColumn6,
																																										 this.gridColumn7,
																																										 this.gridColumn8,
																																										 this.gridColumn9,
																																										 this.gridColumn10,
																																										 this.gridColumn11,
																																										 this.gridColumn12,
																																										 this.gridColumn13,
																																										 this.gridColumn14,
																																										 this.gridColumn15});
			this.gridView1.GridControl = this.gridControl2;
			this.gridView1.Name = "gridView1";
			this.gridView1.OptionsView.NewItemRowPosition = DevExpress.XtraGrid.Views.Grid.NewItemRowPosition.Bottom;
			this.gridView1.OptionsView.ShowGroupPanel = false;
			this.gridView1.InitNewRow += new DevExpress.XtraGrid.Views.Grid.InitNewRowEventHandler(this.gridView1_InitNewRow);
			// 
			// gridColumn1
			// 
			this.gridColumn1.Caption = "idFlujoCaja";
			this.gridColumn1.FieldName = "idFlujoCaja";
			this.gridColumn1.Name = "gridColumn1";
			// 
			// gridColumn2
			// 
			this.gridColumn2.Caption = "Cuenta";
			this.gridColumn2.ColumnEdit = this.repositoryItemLookUpEdit2;
			this.gridColumn2.FieldName = "idCuenta";
			this.gridColumn2.Name = "gridColumn2";
			this.gridColumn2.Visible = true;
			this.gridColumn2.VisibleIndex = 0;
			// 
			// repositoryItemLookUpEdit2
			// 
			this.repositoryItemLookUpEdit2.AutoHeight = false;
			this.repositoryItemLookUpEdit2.Buttons.AddRange(new DevExpress.XtraEditors.Controls.EditorButton[] {
																																																					 new DevExpress.XtraEditors.Controls.EditorButton(DevExpress.XtraEditors.Controls.ButtonPredefines.Combo)});
			this.repositoryItemLookUpEdit2.Columns.AddRange(new DevExpress.XtraEditors.Controls.LookUpColumnInfo[] {
																																																							 new DevExpress.XtraEditors.Controls.LookUpColumnInfo("CodRapido", "CodRapido", 58, DevExpress.Utils.FormatType.None, "", true, DevExpress.Utils.HorzAlignment.Near),
																																																							 new DevExpress.XtraEditors.Controls.LookUpColumnInfo("Descripcion", "Descripcion", 60, DevExpress.Utils.FormatType.None, "", true, DevExpress.Utils.HorzAlignment.Near)});
			this.repositoryItemLookUpEdit2.DataSource = this.cdvCuenta;
			this.repositoryItemLookUpEdit2.DisplayMember = "CodRapido";
			this.repositoryItemLookUpEdit2.Name = "repositoryItemLookUpEdit2";
			this.repositoryItemLookUpEdit2.ValueMember = "idCuenta";
			// 
			// gridColumn3
			// 
			this.gridColumn3.Caption = "Anio";
			this.gridColumn3.FieldName = "Anio";
			this.gridColumn3.Name = "gridColumn3";
			// 
			// gridColumn4
			// 
			this.gridColumn4.Caption = "Enero";
			this.gridColumn4.FieldName = "Enero";
			this.gridColumn4.Name = "gridColumn4";
			this.gridColumn4.Visible = true;
			this.gridColumn4.VisibleIndex = 1;
			// 
			// gridColumn5
			// 
			this.gridColumn5.Caption = "Febrero";
			this.gridColumn5.FieldName = "Febrero";
			this.gridColumn5.Name = "gridColumn5";
			this.gridColumn5.Visible = true;
			this.gridColumn5.VisibleIndex = 2;
			// 
			// gridColumn6
			// 
			this.gridColumn6.Caption = "Marzo";
			this.gridColumn6.FieldName = "Marzo";
			this.gridColumn6.Name = "gridColumn6";
			this.gridColumn6.Visible = true;
			this.gridColumn6.VisibleIndex = 3;
			// 
			// gridColumn7
			// 
			this.gridColumn7.Caption = "Abril";
			this.gridColumn7.FieldName = "Abril";
			this.gridColumn7.Name = "gridColumn7";
			this.gridColumn7.Visible = true;
			this.gridColumn7.VisibleIndex = 4;
			// 
			// gridColumn8
			// 
			this.gridColumn8.Caption = "Mayo";
			this.gridColumn8.FieldName = "Mayo";
			this.gridColumn8.Name = "gridColumn8";
			this.gridColumn8.Visible = true;
			this.gridColumn8.VisibleIndex = 5;
			// 
			// gridColumn9
			// 
			this.gridColumn9.Caption = "Junio";
			this.gridColumn9.FieldName = "Junio";
			this.gridColumn9.Name = "gridColumn9";
			this.gridColumn9.Visible = true;
			this.gridColumn9.VisibleIndex = 6;
			// 
			// gridColumn10
			// 
			this.gridColumn10.Caption = "Julio";
			this.gridColumn10.FieldName = "Julio";
			this.gridColumn10.Name = "gridColumn10";
			this.gridColumn10.Visible = true;
			this.gridColumn10.VisibleIndex = 7;
			// 
			// gridColumn11
			// 
			this.gridColumn11.Caption = "Agosto";
			this.gridColumn11.FieldName = "Agosto";
			this.gridColumn11.Name = "gridColumn11";
			this.gridColumn11.Visible = true;
			this.gridColumn11.VisibleIndex = 8;
			// 
			// gridColumn12
			// 
			this.gridColumn12.Caption = "Septiembre";
			this.gridColumn12.FieldName = "Septiembre";
			this.gridColumn12.Name = "gridColumn12";
			this.gridColumn12.Visible = true;
			this.gridColumn12.VisibleIndex = 9;
			// 
			// gridColumn13
			// 
			this.gridColumn13.Caption = "Octubre";
			this.gridColumn13.FieldName = "Octubre";
			this.gridColumn13.Name = "gridColumn13";
			this.gridColumn13.Visible = true;
			this.gridColumn13.VisibleIndex = 10;
			// 
			// gridColumn14
			// 
			this.gridColumn14.Caption = "Noviembre";
			this.gridColumn14.FieldName = "Noviembre";
			this.gridColumn14.Name = "gridColumn14";
			this.gridColumn14.Visible = true;
			this.gridColumn14.VisibleIndex = 11;
			// 
			// gridColumn15
			// 
			this.gridColumn15.Caption = "Diciembre";
			this.gridColumn15.FieldName = "Diciembre";
			this.gridColumn15.Name = "gridColumn15";
			this.gridColumn15.Visible = true;
			this.gridColumn15.VisibleIndex = 12;
			// 
			// cdvFlujoCaja
			// 
			this.cdvFlujoCaja.BindingContextCtrl = this;
			this.cdvFlujoCaja.DataSet = this.cdsFlujoCaja;
			this.cdvFlujoCaja.RowFilter = "Anio = 2013";
			this.cdvFlujoCaja.Sort = "idFlujoCaja";
			this.cdvFlujoCaja.TableName = "";
			this.cdvFlujoCaja.TableViewName = "FlujoDeCaja";
			// 
			// cmbCentro
			// 
			this.cmbCentro.Location = new System.Drawing.Point(136, 16);
			this.cmbCentro.Name = "cmbCentro";
			// 
			// cmbCentro.Properties
			// 
			this.cmbCentro.Properties.Buttons.AddRange(new DevExpress.XtraEditors.Controls.EditorButton[] {
																																																			new DevExpress.XtraEditors.Controls.EditorButton(DevExpress.XtraEditors.Controls.ButtonPredefines.Combo)});
			this.cmbCentro.Properties.Columns.AddRange(new DevExpress.XtraEditors.Controls.LookUpColumnInfo[] {
																																																					new DevExpress.XtraEditors.Controls.LookUpColumnInfo("Nombre", "Nombre", 49, DevExpress.Utils.FormatType.None, "", true, DevExpress.Utils.HorzAlignment.Near, DevExpress.Data.ColumnSortOrder.Ascending)});
			this.cmbCentro.Properties.DisplayMember = "Nombre";
			this.cmbCentro.Properties.NullText = "";
			this.cmbCentro.Properties.PopupWidth = 300;
			this.cmbCentro.Properties.ShowFooter = false;
			this.cmbCentro.Properties.ShowHeader = false;
			this.cmbCentro.Properties.ValueMember = "idProyecto";
			this.cmbCentro.Size = new System.Drawing.Size(176, 20);
			this.cmbCentro.TabIndex = 62;
			// 
			// cmbTipo
			// 
			this.cmbTipo.Location = new System.Drawing.Point(472, 16);
			this.cmbTipo.Name = "cmbTipo";
			// 
			// cmbTipo.Properties
			// 
			this.cmbTipo.Properties.Buttons.AddRange(new DevExpress.XtraEditors.Controls.EditorButton[] {
																																																		new DevExpress.XtraEditors.Controls.EditorButton(DevExpress.XtraEditors.Controls.ButtonPredefines.Combo)});
			this.cmbTipo.Properties.Items.AddRange(new object[] {
																														"Copiar del Año Pasado",
																														"Copiar del Año Actual",
																														"Copiar Valor de Enero",
																														"Copiar Flujo del Año Pasado"});
			this.cmbTipo.Size = new System.Drawing.Size(128, 20);
			this.cmbTipo.TabIndex = 63;
			// 
			// btCopiar
			// 
			this.btCopiar.Location = new System.Drawing.Point(608, 16);
			this.btCopiar.Name = "btCopiar";
			this.btCopiar.Size = new System.Drawing.Size(24, 24);
			this.btCopiar.TabIndex = 64;
			this.btCopiar.Text = "...";
			this.btCopiar.Click += new System.EventHandler(this.btCopiar_Click);
			// 
			// FlujoCaja
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 14);
			this.ClientSize = new System.Drawing.Size(944, 430);
			this.Controls.Add(this.btCopiar);
			this.Controls.Add(this.cmbTipo);
			this.Controls.Add(this.cmbCentro);
			this.Controls.Add(this.gridControl2);
			this.Controls.Add(this.btGrabar);
			this.Controls.Add(this.labelControl1);
			this.Controls.Add(this.spnAnio);
			this.Name = "FlujoCaja";
			this.Text = "Flujo de Caja";
			this.Load += new System.EventHandler(this.FlujoCaja_Load);
			((System.ComponentModel.ISupportInitialize)(this.cdsFlujoCaja)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvCuenta)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsCuenta)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.spnAnio.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.gridControl2)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.gridView1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.repositoryItemLookUpEdit2)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvFlujoCaja)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCentro.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbTipo.Properties)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void btGrabar_Click(object sender, System.EventArgs e)
		{
			cdsFlujoCaja.Update();
		}

		private void cdsFlujoCaja_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsFlujoCaja.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirConta);
		}

		private void FlujoCaja_Load(object sender, System.EventArgs e)
		{
			spnAnio.EditValue = DateTime.Today.Year;
			cmbCentro.Properties.DataSource = Funcion.dvProcedimiento(cdsCuenta,
				"Select Nombre, idProyecto From Proyecto");
		}

		private void gridView1_InitNewRow(object sender, DevExpress.XtraGrid.Views.Grid.InitNewRowEventArgs e)
		{
			gridView1.SetRowCellValue(gridView1.FocusedRowHandle, "Anio", int.Parse(spnAnio.EditValue.ToString()));
		}

		private void spnAnio_EditValueChanged(object sender, System.EventArgs e)
		{
			cdvFlujoCaja.RowFilter = "Anio = " + spnAnio.EditValue;
		}

		private void btCopiar_Click(object sender, System.EventArgs e)
		{
			int idFlujoCaja = (int) gridView1.GetFocusedRowCellValue("idFlujoCaja");
			string stExec = string.Format("Exec FlujoCajaCopiar {0}, {1}", idFlujoCaja, cmbTipo.SelectedIndex);
			Funcion.EjecutaSQL(cdsCuenta, stExec);
		}
	}
}

