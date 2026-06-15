using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using DevExpress.XtraEditors;
using System.Data;
using System.Data.SqlClient;
using System.IO;

namespace Latinium
{
	/// <summary>
	/// Summary description for DBEmpresa.
	/// </summary>
	public class DBEmpresa : DevExpress.XtraEditors.XtraForm
	{
		private DevExpress.XtraGrid.GridControl gridControl1;
		private DevExpress.XtraEditors.SimpleButton btAceptar;
		private DevExpress.XtraEditors.SimpleButton btCancelar;
		private DevExpress.XtraGrid.Views.Grid.GridView grdEmpresa;
		private DevExpress.XtraEditors.SimpleButton btVerificar;
		private C1.Data.C1DataSet cdsEmpresa;
		private DevExpress.XtraGrid.Columns.GridColumn colidEmpresa;
		private DevExpress.XtraGrid.Columns.GridColumn colEmpresa;
		private DevExpress.XtraGrid.Columns.GridColumn colServidor;
		private DevExpress.XtraGrid.Columns.GridColumn colFacturacion;
		private DevExpress.XtraGrid.Columns.GridColumn colReportes;
		/// <summary>
		/// Required designer variable.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public DBEmpresa()
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
			this.grdEmpresa = new DevExpress.XtraGrid.Views.Grid.GridView();
			this.colidEmpresa = new DevExpress.XtraGrid.Columns.GridColumn();
			this.colEmpresa = new DevExpress.XtraGrid.Columns.GridColumn();
			this.colServidor = new DevExpress.XtraGrid.Columns.GridColumn();
			this.colFacturacion = new DevExpress.XtraGrid.Columns.GridColumn();
			this.colReportes = new DevExpress.XtraGrid.Columns.GridColumn();
			this.btAceptar = new DevExpress.XtraEditors.SimpleButton();
			this.btCancelar = new DevExpress.XtraEditors.SimpleButton();
			this.btVerificar = new DevExpress.XtraEditors.SimpleButton();
			((System.ComponentModel.ISupportInitialize)(this.gridControl1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsEmpresa)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.grdEmpresa)).BeginInit();
			this.SuspendLayout();
			// 
			// gridControl1
			// 
			this.gridControl1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.gridControl1.DataMember = "Empresa";
			this.gridControl1.DataSource = this.cdsEmpresa;
			// 
			// gridControl1.EmbeddedNavigator
			// 
			this.gridControl1.EmbeddedNavigator.Name = "";
			this.gridControl1.Location = new System.Drawing.Point(16, 16);
			this.gridControl1.MainView = this.grdEmpresa;
			this.gridControl1.Name = "gridControl1";
			this.gridControl1.Size = new System.Drawing.Size(720, 296);
			this.gridControl1.TabIndex = 0;
			this.gridControl1.ViewCollection.AddRange(new DevExpress.XtraGrid.Views.Base.BaseView[] {
																																																this.grdEmpresa});
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
			// grdEmpresa
			// 
			this.grdEmpresa.Columns.AddRange(new DevExpress.XtraGrid.Columns.GridColumn[] {
																																											this.colidEmpresa,
																																											this.colEmpresa,
																																											this.colServidor,
																																											this.colFacturacion,
																																											this.colReportes});
			this.grdEmpresa.GridControl = this.gridControl1;
			this.grdEmpresa.Name = "grdEmpresa";
			this.grdEmpresa.OptionsView.NewItemRowPosition = DevExpress.XtraGrid.Views.Grid.NewItemRowPosition.Bottom;
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
			this.colEmpresa.Visible = true;
			this.colEmpresa.VisibleIndex = 0;
			// 
			// colServidor
			// 
			this.colServidor.Caption = "Servidor";
			this.colServidor.FieldName = "Servidor";
			this.colServidor.Name = "colServidor";
			this.colServidor.Visible = true;
			this.colServidor.VisibleIndex = 1;
			// 
			// colFacturacion
			// 
			this.colFacturacion.Caption = "Facturacion";
			this.colFacturacion.FieldName = "Facturacion";
			this.colFacturacion.Name = "colFacturacion";
			this.colFacturacion.Visible = true;
			this.colFacturacion.VisibleIndex = 2;
			// 
			// colReportes
			// 
			this.colReportes.Caption = "Reportes";
			this.colReportes.FieldName = "Reportes";
			this.colReportes.Name = "colReportes";
			this.colReportes.Visible = true;
			this.colReportes.VisibleIndex = 3;
			// 
			// btAceptar
			// 
			this.btAceptar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btAceptar.Location = new System.Drawing.Point(72, 328);
			this.btAceptar.Name = "btAceptar";
			this.btAceptar.TabIndex = 1;
			this.btAceptar.Text = "&Grabar";
			this.btAceptar.Click += new System.EventHandler(this.btAceptar_Click);
			// 
			// btCancelar
			// 
			this.btCancelar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.btCancelar.DialogResult = System.Windows.Forms.DialogResult.Cancel;
			this.btCancelar.Location = new System.Drawing.Point(584, 328);
			this.btCancelar.Name = "btCancelar";
			this.btCancelar.TabIndex = 2;
			this.btCancelar.Text = "&Salir";
			this.btCancelar.Click += new System.EventHandler(this.btCancelar_Click);
			// 
			// btVerificar
			// 
			this.btVerificar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btVerificar.Location = new System.Drawing.Point(160, 328);
			this.btVerificar.Name = "btVerificar";
			this.btVerificar.TabIndex = 3;
			this.btVerificar.Text = "&Verificar";
			this.btVerificar.Click += new System.EventHandler(this.btVerificar_Click);
			this.btVerificar.MouseUp += new System.Windows.Forms.MouseEventHandler(this.btVerificar_MouseUp);
			// 
			// DBEmpresa
			// 
			this.AcceptButton = this.btAceptar;
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 14);
			this.CancelButton = this.btCancelar;
			this.ClientSize = new System.Drawing.Size(752, 366);
			this.Controls.Add(this.btVerificar);
			this.Controls.Add(this.btCancelar);
			this.Controls.Add(this.btAceptar);
			this.Controls.Add(this.gridControl1);
			this.Name = "DBEmpresa";
			this.Text = "DBEmpresa";
			this.Load += new System.EventHandler(this.DBEmpresa_Load);
			((System.ComponentModel.ISupportInitialize)(this.gridControl1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsEmpresa)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.grdEmpresa)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void btAceptar_Click(object sender, System.EventArgs e)
		{
			cdsEmpresa.Update();
			MessageBox.Show("Empresas Grabadas", "Información", 
				MessageBoxButtons.OK, MessageBoxIcon.Information);
		}

		private void DBEmpresa_Load(object sender, System.EventArgs e)
		{

		}

		private void VerificaLinea(int i)
		{
			string stMensajes = "";
			#region Empresa
			string stEmpresa = grdEmpresa.GetRowCellValue(i, colEmpresa).ToString();
			if (stEmpresa.Length < 3)
				stMensajes += "Nombre de Empresa tiene que tener un ancho mayor a 3\n";
			#endregion Empresa
			#region Servidor
			string stServidor = grdEmpresa.GetRowCellValue(i, colEmpresa).ToString();
			if (stServidor.Length < 1)
				stMensajes += "Ingrese nombre de servidor\n";
			#endregion Servidor
			#region Reportes
			string stReportes = grdEmpresa.GetRowCellValue(i, colReportes).ToString();
			if (!Directory.Exists(stReportes))
				stMensajes += string.Format("No existe directorio de reportes: [{0}]\n", stReportes);
			#endregion Reportes
			#region Base de Datos
			string stDataBase = grdEmpresa.GetRowCellValue(i, colFacturacion).ToString();
			string stExec = string.Format("Exec EmpresaVerifica '{0}'", stDataBase);
			stMensajes += Funcion.sEscalarSQL(cdsEmpresa, stExec);
			#endregion Base de Datos
			#region Buscar repetidos
			for (int j = 0; j < grdEmpresa.DataRowCount; j++)
			{
				if (j!=i)
				{
					string stEmpresaR = grdEmpresa.GetRowCellValue(j, colEmpresa).ToString();
					if (stEmpresa == stEmpresaR)
						stMensajes += string.Format("Nombre de Empresa Repetida en línea {0}\n", j+1);
					string stDataBaseR = grdEmpresa.GetRowCellValue(j, colFacturacion).ToString();
					if (stDataBase == stDataBaseR)
						stMensajes += string.Format("Base de Datos Repetida en línea {0}\n", j+1);
				}
			}
			#endregion Buscar repetidos
			#region Mostrar Mensajes
			if (stMensajes.Length > 0)
			{
				stMensajes = string.Format("Error en línea {0} Empresa: [{1}]:\n{2}", 
					i+1, stEmpresa, stMensajes);
				MessageBox.Show(stMensajes, "Información", MessageBoxButtons.OK, MessageBoxIcon.Warning);
			}
			#endregion Mostrar Mensajes
		}

		private void btVerificar_Click(object sender, System.EventArgs e)
		{
			VerificaLinea(grdEmpresa.FocusedRowHandle);
		}

		private void btVerificar_MouseUp(object sender, System.Windows.Forms.MouseEventArgs e)
		{
			if (e.Button != MouseButtons.Right) return;
			for (int i = 0; i < grdEmpresa.DataRowCount; i++)
			{
				VerificaLinea(i);
			}
		}

		private void btCancelar_Click(object sender, System.EventArgs e)
		{
			Close();
		}

		private void cdsEmpresa_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsEmpresa.Schema.Connections[0].ConnectionString = miFuncion.AccesoEmpresaBase();
		}
	}
}
