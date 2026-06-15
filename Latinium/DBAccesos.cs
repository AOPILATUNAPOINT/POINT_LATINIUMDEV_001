using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using DevExpress.XtraEditors;

namespace Latinium
{
	/// <summary>
	/// Summary description for DBAccesos.
	/// </summary>
	public class DBAccesos : DevExpress.XtraEditors.XtraForm
	{
		private DevExpress.XtraGrid.GridControl gridControl1;
		private DevExpress.XtraGrid.Views.Grid.GridView gridView1;
		private C1.Data.C1DataSet cdsEmpresa;
		private DevExpress.XtraGrid.Columns.GridColumn colidUsuarioAcceso;
		private DevExpress.XtraGrid.Columns.GridColumn colNombre;
		private DevExpress.XtraGrid.Columns.GridColumn colEstacion;
		private DevExpress.XtraGrid.Columns.GridColumn colFecha;
		private DevExpress.XtraGrid.Columns.GridColumn colValido;
		private DevExpress.XtraGrid.Columns.GridColumn colMensaje;
		/// <summary>
		/// Required designer variable.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public DBAccesos()
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
			this.gridControl1 = new DevExpress.XtraGrid.GridControl();
			this.cdsEmpresa = new C1.Data.C1DataSet();
			this.gridView1 = new DevExpress.XtraGrid.Views.Grid.GridView();
			this.colidUsuarioAcceso = new DevExpress.XtraGrid.Columns.GridColumn();
			this.colNombre = new DevExpress.XtraGrid.Columns.GridColumn();
			this.colEstacion = new DevExpress.XtraGrid.Columns.GridColumn();
			this.colFecha = new DevExpress.XtraGrid.Columns.GridColumn();
			this.colValido = new DevExpress.XtraGrid.Columns.GridColumn();
			this.colMensaje = new DevExpress.XtraGrid.Columns.GridColumn();
			((System.ComponentModel.ISupportInitialize)(this.gridControl1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsEmpresa)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.gridView1)).BeginInit();
			this.SuspendLayout();
			// 
			// gridControl1
			// 
			this.gridControl1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.gridControl1.DataMember = "UsuarioAcceso";
			this.gridControl1.DataSource = this.cdsEmpresa;
			// 
			// gridControl1.EmbeddedNavigator
			// 
			this.gridControl1.EmbeddedNavigator.Name = "";
			this.gridControl1.Location = new System.Drawing.Point(16, 16);
			this.gridControl1.MainView = this.gridView1;
			this.gridControl1.Name = "gridControl1";
			this.gridControl1.Size = new System.Drawing.Size(592, 368);
			this.gridControl1.TabIndex = 0;
			this.gridControl1.ViewCollection.AddRange(new DevExpress.XtraGrid.Views.Base.BaseView[] {
																																																this.gridView1});
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
			// 
			// gridView1
			// 
			this.gridView1.Columns.AddRange(new DevExpress.XtraGrid.Columns.GridColumn[] {
																																										 this.colidUsuarioAcceso,
																																										 this.colNombre,
																																										 this.colEstacion,
																																										 this.colFecha,
																																										 this.colValido,
																																										 this.colMensaje});
			this.gridView1.GridControl = this.gridControl1;
			this.gridView1.Name = "gridView1";
			// 
			// colidUsuarioAcceso
			// 
			this.colidUsuarioAcceso.Caption = "idUsuarioAcceso";
			this.colidUsuarioAcceso.FieldName = "idUsuarioAcceso";
			this.colidUsuarioAcceso.Name = "colidUsuarioAcceso";
			// 
			// colNombre
			// 
			this.colNombre.Caption = "Nombre";
			this.colNombre.FieldName = "Nombre";
			this.colNombre.Name = "colNombre";
			this.colNombre.OptionsColumn.AllowEdit = false;
			this.colNombre.Visible = true;
			this.colNombre.VisibleIndex = 0;
			// 
			// colEstacion
			// 
			this.colEstacion.Caption = "Estacion";
			this.colEstacion.FieldName = "Estacion";
			this.colEstacion.Name = "colEstacion";
			this.colEstacion.OptionsColumn.AllowEdit = false;
			this.colEstacion.Visible = true;
			this.colEstacion.VisibleIndex = 1;
			// 
			// colFecha
			// 
			this.colFecha.Caption = "Fecha";
			this.colFecha.DisplayFormat.FormatString = "dd/MMM/yyyy HH:mm";
			this.colFecha.DisplayFormat.FormatType = DevExpress.Utils.FormatType.DateTime;
			this.colFecha.FieldName = "Fecha";
			this.colFecha.Name = "colFecha";
			this.colFecha.OptionsColumn.AllowEdit = false;
			this.colFecha.Visible = true;
			this.colFecha.VisibleIndex = 2;
			// 
			// colValido
			// 
			this.colValido.Caption = "Valido";
			this.colValido.FieldName = "Valido";
			this.colValido.Name = "colValido";
			this.colValido.OptionsColumn.AllowEdit = false;
			this.colValido.Visible = true;
			this.colValido.VisibleIndex = 3;
			// 
			// colMensaje
			// 
			this.colMensaje.Caption = "Mensaje";
			this.colMensaje.FieldName = "Mensaje";
			this.colMensaje.Name = "colMensaje";
			this.colMensaje.OptionsColumn.AllowEdit = false;
			this.colMensaje.Visible = true;
			this.colMensaje.VisibleIndex = 4;
			// 
			// DBAccesos
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 14);
			this.ClientSize = new System.Drawing.Size(624, 406);
			this.Controls.Add(this.gridControl1);
			this.Name = "DBAccesos";
			this.Text = "Accesos al sistema";
			((System.ComponentModel.ISupportInitialize)(this.gridControl1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsEmpresa)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.gridView1)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion
	}
}

