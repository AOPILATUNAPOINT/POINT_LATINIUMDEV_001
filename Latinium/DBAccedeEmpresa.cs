using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data.SqlClient;
using System.IO;
using DevExpress.XtraEditors;
using System.Data;
using System.Data.OleDb;

namespace Latinium
{
	/// <summary>
	/// Summary description for DBAccedeEmpresa.
	/// </summary>
	public class DBAccedeEmpresa : DevExpress.XtraEditors.XtraForm
	{
		private DevExpress.XtraEditors.SimpleButton btCancelar;
		private DevExpress.XtraEditors.SimpleButton btAceptar;
		private DevExpress.XtraEditors.LabelControl labelControl1;
		private DevExpress.XtraEditors.LookUpEdit cmbEmpresa;
		/// <summary>
		/// Required designer variable.
		/// </summary>
		private System.ComponentModel.Container components = null;
		string stConeccion="";
		int vidUsuario = 0;
		string NomUsuario="";
		public string stNombreEmpresa="INFOELECT - LATINIUM";
		public DBAccedeEmpresa(string Conexion, int idUsuario)
		{
			//
			// Required for Windows Form Designer support
			//
			stConeccion = Conexion;
			vidUsuario = idUsuario;
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
			this.btCancelar = new DevExpress.XtraEditors.SimpleButton();
			this.btAceptar = new DevExpress.XtraEditors.SimpleButton();
			this.labelControl1 = new DevExpress.XtraEditors.LabelControl();
			this.cmbEmpresa = new DevExpress.XtraEditors.LookUpEdit();
			((System.ComponentModel.ISupportInitialize)(this.cmbEmpresa.Properties)).BeginInit();
			this.SuspendLayout();
			// 
			// btCancelar
			// 
			this.btCancelar.DialogResult = System.Windows.Forms.DialogResult.Cancel;
			this.btCancelar.Location = new System.Drawing.Point(224, 104);
			this.btCancelar.Name = "btCancelar";
			this.btCancelar.Size = new System.Drawing.Size(90, 27);
			this.btCancelar.TabIndex = 3;
			this.btCancelar.Text = "&Cancelar";
			this.btCancelar.Click += new System.EventHandler(this.btCancelar_Click);
			// 
			// btAceptar
			// 
			this.btAceptar.Location = new System.Drawing.Point(40, 104);
			this.btAceptar.Name = "btAceptar";
			this.btAceptar.Size = new System.Drawing.Size(90, 27);
			this.btAceptar.TabIndex = 2;
			this.btAceptar.Text = "&Ingresar";
			this.btAceptar.Click += new System.EventHandler(this.btAceptar_Click);
			// 
			// labelControl1
			// 
			this.labelControl1.AutoSizeMode = DevExpress.XtraEditors.LabelAutoSizeMode.Horizontal;
			this.labelControl1.Location = new System.Drawing.Point(16, 48);
			this.labelControl1.Name = "labelControl1";
			this.labelControl1.Size = new System.Drawing.Size(56, 16);
			this.labelControl1.TabIndex = 13;
			this.labelControl1.Text = "Empresas";
			// 
			// cmbEmpresa
			// 
			this.cmbEmpresa.Location = new System.Drawing.Point(80, 48);
			this.cmbEmpresa.Name = "cmbEmpresa";
			// 
			// cmbEmpresa.Properties
			// 
			this.cmbEmpresa.Properties.BorderStyle = DevExpress.XtraEditors.Controls.BorderStyles.Style3D;
			this.cmbEmpresa.Properties.Buttons.AddRange(new DevExpress.XtraEditors.Controls.EditorButton[] {
																																																			 new DevExpress.XtraEditors.Controls.EditorButton(DevExpress.XtraEditors.Controls.ButtonPredefines.Combo)});
			this.cmbEmpresa.Properties.Columns.AddRange(new DevExpress.XtraEditors.Controls.LookUpColumnInfo[] {
																																																					 new DevExpress.XtraEditors.Controls.LookUpColumnInfo("Empresa", "Empresa", 20, DevExpress.Utils.FormatType.None, "", true, DevExpress.Utils.HorzAlignment.Default, DevExpress.Data.ColumnSortOrder.None),
																																																					 new DevExpress.XtraEditors.Controls.LookUpColumnInfo("idEmpresa", "idEmpresa", 20, DevExpress.Utils.FormatType.None, "", false, DevExpress.Utils.HorzAlignment.Default, DevExpress.Data.ColumnSortOrder.None)});
			this.cmbEmpresa.Properties.DisplayMember = "Nombre";
			this.cmbEmpresa.Properties.NullText = "";
			this.cmbEmpresa.Properties.ValueMember = "idUsuario";
			this.cmbEmpresa.Size = new System.Drawing.Size(250, 24);
			this.cmbEmpresa.TabIndex = 1;
			// 
			// DBAccedeEmpresa
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(6, 16);
			this.ClientSize = new System.Drawing.Size(360, 152);
			this.Controls.Add(this.labelControl1);
			this.Controls.Add(this.cmbEmpresa);
			this.Controls.Add(this.btCancelar);
			this.Controls.Add(this.btAceptar);
			this.Name = "DBAccedeEmpresa";
			this.Text = "ACCESO A EMPRESA ";
			this.Load += new System.EventHandler(this.DBAccedeEmpresa_Load);
			((System.ComponentModel.ISupportInitialize)(this.cmbEmpresa.Properties)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void DBAccedeEmpresa_Load(object sender, System.EventArgs e)
		{	
				string stSelect = "select u.idEmpresa,Empresa from UsuarioEmpresa u "+
					" inner join empresa e on u.idempresa = e.idempresa "+
					" where idusuario = "+vidUsuario;

				try
				{
					using (SqlConnection conn = new SqlConnection(stConeccion))
					{
						conn.Open();
						SqlCommand cmd2 = new SqlCommand(string.Format("Select isnull(nombre,'') from usuario where idUsuario = {0}",+vidUsuario),conn);
						
						NomUsuario = (string) cmd2.ExecuteScalar();

						DataSet dsArticulo = new DataSet();
						SqlDataAdapter myData = new SqlDataAdapter(stSelect,conn.ConnectionString);
						myData.SelectCommand.CommandTimeout = 0;
						myData.TableMappings.Add("Tabla", "ExcelSube");
						dsArticulo.Reset();
						dsArticulo.Clear();
						try
						{
							myData.Fill(dsArticulo);
						}
						catch (Exception ex)
						{
							MessageBox.Show(ex.Message, "Error al ejecutar procedimiento: ");
						}					
						cmbEmpresa.Properties.DataSource =  (DataView)dsArticulo.Tables[0].DefaultView;
						cmbEmpresa.Properties.DisplayMember = "Empresa";
						cmbEmpresa.Properties.ValueMember = "idEmpresa";						
					}
				}
				catch (SqlException ex)
				{
					MessageBox.Show("Error:"+ ex.Message);
					Close();

				}	
		}

		private void btCancelar_Click(object sender, System.EventArgs e)
		{
			Close();
		}

		

		private void btAceptar_Click(object sender, System.EventArgs e)
		{
			string stDir="";
			string drDir="";
			string drDirRep="";
		
			stNombreEmpresa = cmbEmpresa.Text;
			MenuLatinium.stUsuario = NomUsuario;
			try
			{
				using (SqlConnection conn = new SqlConnection(stConeccion))
				{
//					if (conn)conn.Close();
					conn.Open();

					SqlCommand cmd3 = new SqlCommand(string.Format("select facturacion from empresa where idempresa = {0}",cmbEmpresa.EditValue),conn);
					SqlCommand cmd4 = new SqlCommand(string.Format("select Servidor from empresa where idempresa = {0}",cmbEmpresa.EditValue),conn);
					SqlCommand cmd5 = new SqlCommand(string.Format("select reportes from empresa where idempresa = {0}",cmbEmpresa.EditValue),conn);

					drDir = (string)cmd4.ExecuteScalar() ;
					stDir = (string)cmd3.ExecuteScalar() ;
					drDirRep = (string)cmd5.ExecuteScalar() ;
				}
			}
			catch (SqlException ex)
			{
				Console.Write(ex.Message);
			}
			
			MenuLatinium.stDirReportes = @"c:\Latinium\Reportes\";

			DialogResult = DialogResult.OK;
			MenuLatinium.stDirFacturacion = stDir;
			MenuLatinium.stDirConta = stDir;
			MenuLatinium.stDirServidor = drDir;
			MenuLatinium.stDirNumero = stDir;
			MenuLatinium.stDirRoles = stDir;
			MenuLatinium.stDirRetenc = stDir;
			MenuLatinium.iTipoDB = 3;
			MenuLatinium.stDirReportes = drDirRep;


			Close();	
		}
	}
}

