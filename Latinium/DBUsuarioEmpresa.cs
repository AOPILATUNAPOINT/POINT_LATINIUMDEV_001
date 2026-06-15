using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using DevExpress.XtraEditors;

namespace Latinium
{
	/// <summary>
	/// Summary description for DBUsuarioEmpresa.
	/// </summary>
	public class DBUsuarioEmpresa : DevExpress.XtraEditors.XtraForm
	{
		private C1.Data.C1DataSet cdsEmpresa;
		private DevExpress.XtraEditors.LookUpEdit cmbUsuario;
		private DevExpress.XtraGrid.GridControl gridControl1;
		private DevExpress.XtraGrid.Views.Grid.GridView gridView1;
		private DevExpress.XtraGrid.Columns.GridColumn colidEmpresa;
		private DevExpress.XtraGrid.Columns.GridColumn colEmpresa;
		private DevExpress.XtraGrid.Columns.GridColumn colServidor;
		private DevExpress.XtraGrid.Columns.GridColumn colFacturacion;
		private DevExpress.XtraGrid.Columns.GridColumn colReportes;
		private C1.Data.C1DataView cdvUsuario;
		private DevExpress.XtraGrid.GridControl gridControl2;
		private DevExpress.XtraGrid.Views.Grid.GridView gridView2;
		private System.Windows.Forms.Label label1;
		private DevExpress.XtraEditors.SimpleButton btAgregaTodo;
		private DevExpress.XtraEditors.SimpleButton btAgregaUno;
		private DevExpress.XtraEditors.SimpleButton btQuitarUno;
		private DevExpress.XtraEditors.SimpleButton btQuitarTodo;
		private C1.Data.C1DataView cdvUsuarioEmpresa;
		private DevExpress.XtraGrid.Columns.GridColumn colidUsuarioEmpresa;
		private DevExpress.XtraGrid.Columns.GridColumn colidUsuario;
		private DevExpress.XtraGrid.Columns.GridColumn colidEmpresa1;
		private DevExpress.XtraEditors.Repository.RepositoryItemLookUpEdit repositoryItemLookUpEdit1;
		private C1.Data.C1DataView cdvEmpresa;
		private DevExpress.XtraEditors.LabelControl labelControl1;
		/// <summary>
		/// Required designer variable.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public DBUsuarioEmpresa()
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
			this.cmbUsuario = new DevExpress.XtraEditors.LookUpEdit();
			this.cdvUsuario = new C1.Data.C1DataView();
			this.cdsEmpresa = new C1.Data.C1DataSet();
			this.gridControl1 = new DevExpress.XtraGrid.GridControl();
			this.gridView1 = new DevExpress.XtraGrid.Views.Grid.GridView();
			this.colidEmpresa = new DevExpress.XtraGrid.Columns.GridColumn();
			this.colEmpresa = new DevExpress.XtraGrid.Columns.GridColumn();
			this.colServidor = new DevExpress.XtraGrid.Columns.GridColumn();
			this.colFacturacion = new DevExpress.XtraGrid.Columns.GridColumn();
			this.colReportes = new DevExpress.XtraGrid.Columns.GridColumn();
			this.gridControl2 = new DevExpress.XtraGrid.GridControl();
			this.cdvUsuarioEmpresa = new C1.Data.C1DataView();
			this.gridView2 = new DevExpress.XtraGrid.Views.Grid.GridView();
			this.colidUsuarioEmpresa = new DevExpress.XtraGrid.Columns.GridColumn();
			this.colidUsuario = new DevExpress.XtraGrid.Columns.GridColumn();
			this.colidEmpresa1 = new DevExpress.XtraGrid.Columns.GridColumn();
			this.repositoryItemLookUpEdit1 = new DevExpress.XtraEditors.Repository.RepositoryItemLookUpEdit();
			this.cdvEmpresa = new C1.Data.C1DataView();
			this.label1 = new System.Windows.Forms.Label();
			this.btAgregaTodo = new DevExpress.XtraEditors.SimpleButton();
			this.btAgregaUno = new DevExpress.XtraEditors.SimpleButton();
			this.btQuitarUno = new DevExpress.XtraEditors.SimpleButton();
			this.btQuitarTodo = new DevExpress.XtraEditors.SimpleButton();
			this.labelControl1 = new DevExpress.XtraEditors.LabelControl();
			((System.ComponentModel.ISupportInitialize)(this.cmbUsuario.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvUsuario)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsEmpresa)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.gridControl1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.gridView1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.gridControl2)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvUsuarioEmpresa)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.gridView2)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.repositoryItemLookUpEdit1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvEmpresa)).BeginInit();
			this.SuspendLayout();
			// 
			// cmbUsuario
			// 
			this.cmbUsuario.Location = new System.Drawing.Point(470, 37);
			this.cmbUsuario.Name = "cmbUsuario";
			// 
			// cmbUsuario.Properties
			// 
			this.cmbUsuario.Properties.Buttons.AddRange(new DevExpress.XtraEditors.Controls.EditorButton[] {
																																																			 new DevExpress.XtraEditors.Controls.EditorButton(DevExpress.XtraEditors.Controls.ButtonPredefines.Combo)});
			this.cmbUsuario.Properties.Columns.AddRange(new DevExpress.XtraEditors.Controls.LookUpColumnInfo[] {
																																																					 new DevExpress.XtraEditors.Controls.LookUpColumnInfo("Nombre", "Nombre", 43, DevExpress.Utils.FormatType.None, "", true, DevExpress.Utils.HorzAlignment.Near),
																																																					 new DevExpress.XtraEditors.Controls.LookUpColumnInfo("idGrupo", "Grupo", 43, DevExpress.Utils.FormatType.Numeric, "", true, DevExpress.Utils.HorzAlignment.Far, DevExpress.Data.ColumnSortOrder.None)});
			this.cmbUsuario.Properties.DataSource = this.cdvUsuario;
			this.cmbUsuario.Properties.DisplayMember = "Nombre";
			this.cmbUsuario.Properties.NullText = "";
			this.cmbUsuario.Properties.ValueMember = "idUsuario";
			this.cmbUsuario.Size = new System.Drawing.Size(250, 22);
			this.cmbUsuario.TabIndex = 1;
			this.cmbUsuario.EditValueChanged += new System.EventHandler(this.cmbUsuario_EditValueChanged);
			// 
			// cdvUsuario
			// 
			this.cdvUsuario.BindingContextCtrl = this;
			this.cdvUsuario.DataSet = this.cdsEmpresa;
			this.cdvUsuario.TableName = "";
			this.cdvUsuario.TableViewName = "Usuario";
			// 
			// cdsEmpresa
			// 
			this.cdsEmpresa.BindingContextCtrl = this;
			this.cdsEmpresa.DataLibrary = "LibDBEmpresa";
			this.cdsEmpresa.DataLibraryUrl = "";
			this.cdsEmpresa.DataSetDef = "dsEmpresa";
			this.cdsEmpresa.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsEmpresa.Name = "cdsEmpresa";
			this.cdsEmpresa.SchemaClassName = "LibDBEmpresa.DataClass";
			this.cdsEmpresa.SchemaDef = null;
			this.cdsEmpresa.BeforeFill += new C1.Data.FillEventHandler(this.cdsEmpresa_BeforeFill);
			// 
			// gridControl1
			// 
			this.gridControl1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left)));
			this.gridControl1.DataMember = "Empresa";
			this.gridControl1.DataSource = this.cdsEmpresa;
			// 
			// gridControl1.EmbeddedNavigator
			// 
			this.gridControl1.EmbeddedNavigator.Name = "";
			this.gridControl1.Location = new System.Drawing.Point(29, 82);
			this.gridControl1.MainView = this.gridView1;
			this.gridControl1.Name = "gridControl1";
			this.gridControl1.Size = new System.Drawing.Size(317, 265);
			this.gridControl1.TabIndex = 2;
			this.gridControl1.ViewCollection.AddRange(new DevExpress.XtraGrid.Views.Base.BaseView[] {
																																																this.gridView1});
			// 
			// gridView1
			// 
			this.gridView1.Columns.AddRange(new DevExpress.XtraGrid.Columns.GridColumn[] {
																																										 this.colidEmpresa,
																																										 this.colEmpresa,
																																										 this.colServidor,
																																										 this.colFacturacion,
																																										 this.colReportes});
			this.gridView1.GridControl = this.gridControl1;
			this.gridView1.Name = "gridView1";
			// 
			// colidEmpresa
			// 
			this.colidEmpresa.Caption = "idEmpresa";
			this.colidEmpresa.FieldName = "idEmpresa";
			this.colidEmpresa.Name = "colidEmpresa";
			// 
			// colEmpresa
			// 
			this.colEmpresa.Caption = "Empresa";
			this.colEmpresa.FieldName = "Empresa";
			this.colEmpresa.Name = "colEmpresa";
			this.colEmpresa.OptionsColumn.AllowEdit = false;
			this.colEmpresa.Visible = true;
			this.colEmpresa.VisibleIndex = 0;
			// 
			// colServidor
			// 
			this.colServidor.Caption = "Servidor";
			this.colServidor.FieldName = "Servidor";
			this.colServidor.Name = "colServidor";
			// 
			// colFacturacion
			// 
			this.colFacturacion.Caption = "Facturacion";
			this.colFacturacion.FieldName = "Facturacion";
			this.colFacturacion.Name = "colFacturacion";
			// 
			// colReportes
			// 
			this.colReportes.Caption = "Reportes";
			this.colReportes.FieldName = "Reportes";
			this.colReportes.Name = "colReportes";
			// 
			// gridControl2
			// 
			this.gridControl2.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.gridControl2.DataSource = this.cdvUsuarioEmpresa;
			// 
			// gridControl2.EmbeddedNavigator
			// 
			this.gridControl2.EmbeddedNavigator.Name = "";
			this.gridControl2.Location = new System.Drawing.Point(403, 82);
			this.gridControl2.MainView = this.gridView2;
			this.gridControl2.Name = "gridControl2";
			this.gridControl2.RepositoryItems.AddRange(new DevExpress.XtraEditors.Repository.RepositoryItem[] {
																																																					this.repositoryItemLookUpEdit1});
			this.gridControl2.Size = new System.Drawing.Size(317, 265);
			this.gridControl2.TabIndex = 3;
			this.gridControl2.ViewCollection.AddRange(new DevExpress.XtraGrid.Views.Base.BaseView[] {
																																																this.gridView2});
			// 
			// cdvUsuarioEmpresa
			// 
			this.cdvUsuarioEmpresa.BindingContextCtrl = this;
			this.cdvUsuarioEmpresa.DataSet = this.cdsEmpresa;
			this.cdvUsuarioEmpresa.RowFilter = "idEmpresa = 0";
			this.cdvUsuarioEmpresa.TableName = "";
			this.cdvUsuarioEmpresa.TableViewName = "UsuarioEmpresa";
			// 
			// gridView2
			// 
			this.gridView2.Columns.AddRange(new DevExpress.XtraGrid.Columns.GridColumn[] {
																																										 this.colidUsuarioEmpresa,
																																										 this.colidUsuario,
																																										 this.colidEmpresa1});
			this.gridView2.GridControl = this.gridControl2;
			this.gridView2.Name = "gridView2";
			// 
			// colidUsuarioEmpresa
			// 
			this.colidUsuarioEmpresa.Caption = "idUsuarioEmpresa";
			this.colidUsuarioEmpresa.FieldName = "idUsuarioEmpresa";
			this.colidUsuarioEmpresa.Name = "colidUsuarioEmpresa";
			// 
			// colidUsuario
			// 
			this.colidUsuario.Caption = "idUsuario";
			this.colidUsuario.FieldName = "idUsuario";
			this.colidUsuario.Name = "colidUsuario";
			// 
			// colidEmpresa1
			// 
			this.colidEmpresa1.Caption = "Empresa del Usuario";
			this.colidEmpresa1.ColumnEdit = this.repositoryItemLookUpEdit1;
			this.colidEmpresa1.FieldName = "idEmpresa";
			this.colidEmpresa1.Name = "colidEmpresa1";
			this.colidEmpresa1.OptionsColumn.AllowEdit = false;
			this.colidEmpresa1.Visible = true;
			this.colidEmpresa1.VisibleIndex = 0;
			// 
			// repositoryItemLookUpEdit1
			// 
			this.repositoryItemLookUpEdit1.AutoHeight = false;
			this.repositoryItemLookUpEdit1.Buttons.AddRange(new DevExpress.XtraEditors.Controls.EditorButton[] {
																																																					 new DevExpress.XtraEditors.Controls.EditorButton(DevExpress.XtraEditors.Controls.ButtonPredefines.Combo)});
			this.repositoryItemLookUpEdit1.Columns.AddRange(new DevExpress.XtraEditors.Controls.LookUpColumnInfo[] {
																																																							 new DevExpress.XtraEditors.Controls.LookUpColumnInfo("Empresa", "Empresa", 47, DevExpress.Utils.FormatType.None, "", true, DevExpress.Utils.HorzAlignment.Near)});
			this.repositoryItemLookUpEdit1.DataSource = this.cdvEmpresa;
			this.repositoryItemLookUpEdit1.DisplayMember = "Empresa";
			this.repositoryItemLookUpEdit1.Name = "repositoryItemLookUpEdit1";
			this.repositoryItemLookUpEdit1.ValueMember = "idEmpresa";
			// 
			// cdvEmpresa
			// 
			this.cdvEmpresa.BindingContextCtrl = this;
			this.cdvEmpresa.DataSet = this.cdsEmpresa;
			this.cdvEmpresa.TableName = "";
			this.cdvEmpresa.TableViewName = "Empresa";
			// 
			// label1
			// 
			this.label1.Font = new System.Drawing.Font("Tahoma", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label1.Location = new System.Drawing.Point(29, 37);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(230, 26);
			this.label1.TabIndex = 4;
			this.label1.Text = "Listado de Empresas";
			// 
			// btAgregaTodo
			// 
			this.btAgregaTodo.Location = new System.Drawing.Point(355, 146);
			this.btAgregaTodo.Name = "btAgregaTodo";
			this.btAgregaTodo.Size = new System.Drawing.Size(39, 27);
			this.btAgregaTodo.TabIndex = 5;
			this.btAgregaTodo.Text = ">>";
			this.btAgregaTodo.ToolTip = "Agregar todas las empresas al usuario";
			this.btAgregaTodo.Click += new System.EventHandler(this.btAgregaTodo_Click);
			// 
			// btAgregaUno
			// 
			this.btAgregaUno.Location = new System.Drawing.Point(355, 183);
			this.btAgregaUno.Name = "btAgregaUno";
			this.btAgregaUno.Size = new System.Drawing.Size(39, 26);
			this.btAgregaUno.TabIndex = 6;
			this.btAgregaUno.Text = "=>";
			this.btAgregaUno.ToolTip = "Agregar una empresa al usuario";
			this.btAgregaUno.Click += new System.EventHandler(this.btAgregaUno_Click);
			// 
			// btQuitarUno
			// 
			this.btQuitarUno.Location = new System.Drawing.Point(355, 229);
			this.btQuitarUno.Name = "btQuitarUno";
			this.btQuitarUno.Size = new System.Drawing.Size(39, 26);
			this.btQuitarUno.TabIndex = 7;
			this.btQuitarUno.Text = "<=";
			this.btQuitarUno.ToolTip = "Quitar una empresa del usuario";
			this.btQuitarUno.Click += new System.EventHandler(this.btQuitarUno_Click);
			// 
			// btQuitarTodo
			// 
			this.btQuitarTodo.Location = new System.Drawing.Point(355, 274);
			this.btQuitarTodo.Name = "btQuitarTodo";
			this.btQuitarTodo.Size = new System.Drawing.Size(39, 27);
			this.btQuitarTodo.TabIndex = 8;
			this.btQuitarTodo.Text = "<<";
			this.btQuitarTodo.ToolTip = "Quitar todas las empresa del usuario";
			this.btQuitarTodo.Click += new System.EventHandler(this.btQuitarTodo_Click);
			// 
			// labelControl1
			// 
			this.labelControl1.AutoSizeMode = DevExpress.XtraEditors.LabelAutoSizeMode.Horizontal;
			this.labelControl1.Location = new System.Drawing.Point(403, 37);
			this.labelControl1.Name = "labelControl1";
			this.labelControl1.Size = new System.Drawing.Size(43, 16);
			this.labelControl1.TabIndex = 9;
			this.labelControl1.Text = "Usuario";
			// 
			// DBUsuarioEmpresa
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(6, 16);
			this.ClientSize = new System.Drawing.Size(748, 390);
			this.Controls.Add(this.gridControl2);
			this.Controls.Add(this.labelControl1);
			this.Controls.Add(this.btQuitarTodo);
			this.Controls.Add(this.btQuitarUno);
			this.Controls.Add(this.btAgregaUno);
			this.Controls.Add(this.btAgregaTodo);
			this.Controls.Add(this.label1);
			this.Controls.Add(this.gridControl1);
			this.Controls.Add(this.cmbUsuario);
			this.Name = "DBUsuarioEmpresa";
			this.Text = "Empresas a las que accede el usuario";
			this.Load += new System.EventHandler(this.DBUsuarioEmpresa_Load);
			((System.ComponentModel.ISupportInitialize)(this.cmbUsuario.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvUsuario)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsEmpresa)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.gridControl1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.gridView1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.gridControl2)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvUsuarioEmpresa)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.gridView2)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.repositoryItemLookUpEdit1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvEmpresa)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void cdsEmpresa_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsEmpresa.Schema.Connections[0].ConnectionString = miFuncion.AccesoEmpresaBase();

		}

		private void cmbUsuario_EditValueChanged(object sender, System.EventArgs e)
		{
			int idUsuario = (int) cmbUsuario.EditValue;
			cdvUsuarioEmpresa.RowFilter = "idUsuario = " + idUsuario;
		}

		private void btAgregaUno_Click(object sender, System.EventArgs e)
		{
			int iFila = gridView1.FocusedRowHandle;
			if (iFila < 0) 
			{
				MessageBox.Show("Seleccione Fila de la grilla de la izquierda con la empresa a asignar al usuario", "Información",
					MessageBoxButtons.OK, MessageBoxIcon.Information);
				return;
			}
			if (cmbUsuario.EditValue == null || cmbUsuario.Text.Length == 0)
			{
				MessageBox.Show("Seleccione usuario", "Información",
					MessageBoxButtons.OK, MessageBoxIcon.Information);
				return;
			}
			int idUsuario = (int) cmbUsuario.EditValue;
			int idEmpresa = (int) gridView1.GetRowCellValue(iFila, colidEmpresa);
			string stInserta = string.Format("Exec UsuarioEmpresaCrear {0}, {1}",	idUsuario, idEmpresa);
			Funcion.EjecutaSQL(cdsEmpresa, stInserta);
			cdsEmpresa.Update();
			cdsEmpresa.Clear();
			cdsEmpresa.Fill();
		}

		private void btQuitarUno_Click(object sender, System.EventArgs e)
		{
			int iFila = gridView2.FocusedRowHandle;
			if (iFila < 0) 
			{
				MessageBox.Show("Seleccione Fila de la grilla de la derecha con la empresa a quitar al usuario", "Información",
					MessageBoxButtons.OK, MessageBoxIcon.Information);
				return;
			}
			if (cmbUsuario.EditValue == null || cmbUsuario.Text.Length == 0)
			{
				MessageBox.Show("Seleccione usuario", "Información",
					MessageBoxButtons.OK, MessageBoxIcon.Information);
				return;
			}
			int idUsuario = (int) cmbUsuario.EditValue;
			int idEmpresa = (int) gridView2.GetRowCellValue(iFila, colidEmpresa1);
			string stInserta = string.Format("Exec UsuarioEmpresaBorrar {0}, {1}",	idUsuario, idEmpresa);
			Funcion.EjecutaSQL(cdsEmpresa, stInserta);
			cdsEmpresa.Update();
			cdsEmpresa.Clear();
			cdsEmpresa.Fill();
		}

		private void btQuitarTodo_Click(object sender, System.EventArgs e)
		{
			if (gridView2.RowCount == 0)
			{
				MessageBox.Show("No existen empresas para borrar", "Información",
					MessageBoxButtons.OK, MessageBoxIcon.Information);
				return;
			}
			if (cmbUsuario.EditValue == null || cmbUsuario.Text.Length == 0)
			{
				MessageBox.Show("Seleccione usuario", "Información",
					MessageBoxButtons.OK, MessageBoxIcon.Information);
				return;
			}
			int idUsuario = (int) cmbUsuario.EditValue;
			for (int i= 0; i < gridView2.RowCount; i++)
			{
				int idEmpresa = (int) gridView2.GetRowCellValue(i, colidEmpresa1);
				string stInserta = string.Format("Exec UsuarioEmpresaBorrar {0}, {1}",	idUsuario, idEmpresa);
				Funcion.EjecutaSQL(cdsEmpresa, stInserta);
			}
			cdsEmpresa.Update();
			cdsEmpresa.Clear();
			cdsEmpresa.Fill();
		}

		private void btAgregaTodo_Click(object sender, System.EventArgs e)
		{
			if (gridView1.RowCount == 0)
			{
				MessageBox.Show("No existen empresas para pasar", "Información",
					MessageBoxButtons.OK, MessageBoxIcon.Information);
				return;
			}
			if (cmbUsuario.EditValue == null || cmbUsuario.Text.Length == 0)
			{
				MessageBox.Show("Seleccione usuario", "Información",
					MessageBoxButtons.OK, MessageBoxIcon.Information);
				return;
			}
			int idUsuario = (int) cmbUsuario.EditValue;
			for (int i= 0; i < gridView1.RowCount; i++)
			{
				int idEmpresa = (int) gridView1.GetRowCellValue(i, colidEmpresa);
				string stInserta = string.Format("Exec UsuarioEmpresaCrear {0}, {1}",	idUsuario, idEmpresa);
				Funcion.EjecutaSQL(cdsEmpresa, stInserta);
			}
			cdsEmpresa.Update();
			cdsEmpresa.Clear();
			cdsEmpresa.Fill();
		}

		private void DBUsuarioEmpresa_Load(object sender, System.EventArgs e)
		{
		
		}
	}
}
