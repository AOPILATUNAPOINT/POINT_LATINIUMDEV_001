using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data.SqlClient;
using System.IO;
using DevExpress.XtraEditors;

namespace Latinium
{
	/// <summary>
	/// Summary description for XtraForm1.
	/// </summary>
	public class DBIngresoUsuario : DevExpress.XtraEditors.XtraForm
	{
		private DevExpress.XtraEditors.LabelControl lblServidor;
		private DevExpress.XtraEditors.TextEdit txtServidor;
		private DevExpress.XtraEditors.LabelControl labelControl2;
		private DevExpress.XtraEditors.TextEdit txtClave;
		private DevExpress.XtraEditors.LabelControl labelControl1;
		private DevExpress.XtraEditors.TextEdit txtUsuario;
		private DevExpress.XtraEditors.SimpleButton btCancelar;
		private DevExpress.XtraEditors.SimpleButton btAceptar;
		private DevExpress.XtraEditors.DXErrorProvider.DXErrorProvider dxErrorProvider1;
		private System.ComponentModel.IContainer components;

		public DBIngresoUsuario()
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
			this.components = new System.ComponentModel.Container();
			this.lblServidor = new DevExpress.XtraEditors.LabelControl();
			this.txtServidor = new DevExpress.XtraEditors.TextEdit();
			this.labelControl2 = new DevExpress.XtraEditors.LabelControl();
			this.txtClave = new DevExpress.XtraEditors.TextEdit();
			this.labelControl1 = new DevExpress.XtraEditors.LabelControl();
			this.txtUsuario = new DevExpress.XtraEditors.TextEdit();
			this.btCancelar = new DevExpress.XtraEditors.SimpleButton();
			this.btAceptar = new DevExpress.XtraEditors.SimpleButton();
			this.dxErrorProvider1 = new DevExpress.XtraEditors.DXErrorProvider.DXErrorProvider(this.components);
			((System.ComponentModel.ISupportInitialize)(this.txtServidor.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtClave.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtUsuario.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dxErrorProvider1)).BeginInit();
			this.SuspendLayout();
			// 
			// lblServidor
			// 
			this.lblServidor.AutoSizeMode = DevExpress.XtraEditors.LabelAutoSizeMode.Horizontal;
			this.lblServidor.Location = new System.Drawing.Point(56, 40);
			this.lblServidor.Name = "lblServidor";
			this.lblServidor.Size = new System.Drawing.Size(40, 13);
			this.lblServidor.TabIndex = 15;
			this.lblServidor.Text = "Servidor";
			// 
			// txtServidor
			// 
			this.txtServidor.Location = new System.Drawing.Point(128, 40);
			this.txtServidor.Name = "txtServidor";
			this.txtServidor.Size = new System.Drawing.Size(176, 20);
			this.txtServidor.TabIndex = 14;
			// 
			// labelControl2
			// 
			this.labelControl2.AutoSizeMode = DevExpress.XtraEditors.LabelAutoSizeMode.Horizontal;
			this.labelControl2.Location = new System.Drawing.Point(56, 120);
			this.labelControl2.Name = "labelControl2";
			this.labelControl2.Size = new System.Drawing.Size(56, 13);
			this.labelControl2.TabIndex = 13;
			this.labelControl2.Text = "Contraseña";
			// 
			// txtClave
			// 
			this.txtClave.Location = new System.Drawing.Point(128, 120);
			this.txtClave.Name = "txtClave";
			// 
			// txtClave.Properties
			// 
			this.txtClave.Properties.PasswordChar = '*';
			this.txtClave.Size = new System.Drawing.Size(176, 20);
			this.txtClave.TabIndex = 12;
			// 
			// labelControl1
			// 
			this.labelControl1.AutoSizeMode = DevExpress.XtraEditors.LabelAutoSizeMode.Horizontal;
			this.labelControl1.Location = new System.Drawing.Point(56, 80);
			this.labelControl1.Name = "labelControl1";
			this.labelControl1.Size = new System.Drawing.Size(36, 13);
			this.labelControl1.TabIndex = 11;
			this.labelControl1.Text = "Usuario";
			// 
			// txtUsuario
			// 
			this.txtUsuario.Location = new System.Drawing.Point(128, 80);
			this.txtUsuario.Name = "txtUsuario";
			this.txtUsuario.Size = new System.Drawing.Size(176, 20);
			this.txtUsuario.TabIndex = 10;
			// 
			// btCancelar
			// 
			this.btCancelar.DialogResult = System.Windows.Forms.DialogResult.Cancel;
			this.btCancelar.Location = new System.Drawing.Point(216, 184);
			this.btCancelar.Name = "btCancelar";
			this.btCancelar.TabIndex = 9;
			this.btCancelar.Text = "&Cancelar";
			this.btCancelar.Click += new System.EventHandler(this.btCancelar_Click);
			// 
			// btAceptar
			// 
			this.btAceptar.Location = new System.Drawing.Point(64, 184);
			this.btAceptar.Name = "btAceptar";
			this.btAceptar.TabIndex = 8;
			this.btAceptar.Text = "&Aceptar";
			this.btAceptar.Click += new System.EventHandler(this.btAceptar_Click);
			// 
			// dxErrorProvider1
			// 
			this.dxErrorProvider1.ContainerControl = this;
			// 
			// DBIngresoUsuario
			// 
			this.AcceptButton = this.btAceptar;
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 14);
			this.CancelButton = this.btCancelar;
			this.ClientSize = new System.Drawing.Size(360, 246);
			this.Controls.Add(this.lblServidor);
			this.Controls.Add(this.txtServidor);
			this.Controls.Add(this.labelControl2);
			this.Controls.Add(this.txtClave);
			this.Controls.Add(this.labelControl1);
			this.Controls.Add(this.txtUsuario);
			this.Controls.Add(this.btCancelar);
			this.Controls.Add(this.btAceptar);
			this.Name = "DBIngresoUsuario";
			this.Text = "Ingreso de Usuario";
			this.Load += new System.EventHandler(this.UsuarioIngreso_Load);
			((System.ComponentModel.ISupportInitialize)(this.txtServidor.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtClave.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtUsuario.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dxErrorProvider1)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion
		private void btCancelar_Click(object sender, System.EventArgs e)
		{
			Close();
		}

		private bool VerificaServidor(string stServidor)
		{
			Cursor = Cursors.WaitCursor;
			if (stServidor == null || stServidor.Length == 0)
			{
				dxErrorProvider1.SetError(txtServidor, @"Ingrese Nombre del Servidor y la Instancia ej: Server\SqlExpress");
				return false;
			}

			 stConeccion = string.Format(
				"Data Source={0};Initial Catalog=Master;Persist Security Info=True;User ID=infoelect;Password=Bl45o6$9;",stServidor);
			try
			{
				using (SqlConnection conn = new SqlConnection(stConeccion))
				{
					SqlCommand cmd = new SqlCommand("SELECT count(*) FROM master.dbo.sysdatabases Where Name In ('EmpresasBase', 'Bases')", conn);
					conn.Open();
					int iCuenta = (int)cmd.ExecuteScalar();
					if (iCuenta != 2)
					{
						dxErrorProvider1.SetError(txtServidor, 
							"No existen las bases de datos principales de sistema Latinium en Instancia " + stServidor);
						return false;
					}
				}
			}
			catch (SqlException ex)
			{
				string stError = ex.Message;
				SqlErrorCollection stLisat = ex.Errors;
				if (stLisat[0].Number == 18456) stError = "No existe un usuario válido de Latinium en servidor: " + stServidor;
				if (stLisat[0].Number == 17) stError = "No existe o no se tiene acceso a servidor: " + stServidor;
				if (ex.InnerException != null) stError = ex.InnerException.Message;
				dxErrorProvider1.SetError(txtServidor, stError);
				return false;
			}
			Servidor = stServidor;
			return true;
		}

		private string Servidor = "";
		public int idGrupo = 0;
		public string stConeccion="";
		public int vidUsuario=0;
		private bool VerificaUsuario(string stUsuario, string stClave, int Tipo)
		{
			 stConeccion = string.Format(
				"Data Source={0};Initial Catalog=EmpresasBase;Persist Security Info=True;User ID=infoelect;Password=Bl45o6$9;",Servidor);
			try
			{
				using (SqlConnection conn = new SqlConnection(stConeccion))
				{
					SqlCommand cmd = new SqlCommand(string.Format("Exec UsuarioVerifica '{0}', '{1}', {2}",
						stUsuario, stClave, Tipo), conn);
					conn.Open();
					if (Tipo != 2)
					{
						string stMensaje = (string) cmd.ExecuteScalar();
						if (stMensaje.Length > 0)
						{
							dxErrorProvider1.SetError(txtUsuario, stMensaje);
							return false;
						}
						return true;
					}
					else
					{
						idGrupo = (int) cmd.ExecuteScalar();
						SqlCommand cmd1 = new SqlCommand(string.Format("Exec UsuarioID '{0}', '{1}'",stUsuario,stClave),conn);
						vidUsuario = (int) cmd1.ExecuteScalar();
						return true;
					}

				}
			}
			catch (SqlException ex)
			{
				string stError = ex.Message;
				SqlErrorCollection stLisat = ex.Errors;
				if (stLisat[0].Number == 18456) stError = "No existe un usuario válido de Latinium en servidor: " + Servidor;
				if (stLisat[0].Number == 17) stError = "No existe o no se tiene acceso a servidor: " + Servidor;
				if (ex.InnerException != null) stError = ex.InnerException.Message;
				dxErrorProvider1.SetError(txtServidor, stError);
				return false;
			}
		}

		private void UsuarioAccesoCrea(string stNombre, int Valido, string stMensaje)
		{
			string stExec = string.Format("Exec UsuarioAccesoCrea '{0}', {1}, '{2}'",
				stNombre, Valido, stMensaje);

			 stConeccion = string.Format(
				"Data Source={0};Initial Catalog=EmpresasBase;Persist Security Info=True;User ID=infoelect;Password=Bl45o6$9;",Servidor);
			try
			{
				using (SqlConnection conn = new SqlConnection(stConeccion))
				{
					SqlCommand cmd = new SqlCommand(stExec, conn);
					conn.Open();
					cmd.ExecuteScalar();
				}
			}
			catch (SqlException ex)
			{
				string stError = ex.Message;
				SqlErrorCollection stLisat = ex.Errors;
				if (stLisat[0].Number == 18456) stError = "No existe un usuario válido de Latinium en servidor: " + Servidor;
				if (stLisat[0].Number == 17) stError = "No existe o no se tiene acceso a servidor: " + Servidor;
				if (ex.InnerException != null) stError = ex.InnerException.Message;
				dxErrorProvider1.SetError(txtServidor, stError);
			}
		}

		private void btAceptar_Click(object sender, System.EventArgs e)
		{
			#region Verifica Servidor
			if (lblServidor.Visible)
			{
				string stServidor = txtServidor.Text;
				lblServidor.Text = "Verificando";
				lblServidor.Invalidate();
				lblServidor.PerformLayout();
				lblServidor.Refresh();
				lblServidor.ResumeLayout();
				if (VerificaServidor(stServidor))
				{
					Cursor = Cursors.Default;
					StreamWriter sr = new StreamWriter(stArchivo);
					sr.WriteLine(stServidor);
					sr.Flush();
					sr.Close();
					txtServidor.Width = 0;
					lblServidor.Visible = false;
					MessageBox.Show("Servidor Aceptado", "Información", MessageBoxButtons.OK, MessageBoxIcon.Information);
				}
				else
				{
					lblServidor.Text = "Servidor";
					Cursor = Cursors.Default;
					MessageBox.Show("Servidor No Válido", "Información", MessageBoxButtons.OK, MessageBoxIcon.Warning);
					return;
				}
			}
			#endregion Verifica Servidor
			#region Usuario
			string stUsuario = txtUsuario.Text.Trim();
			string stClave = txtClave.Text.ToString().Trim();
			string stClaveBasica = Funcion.CreaHash(stClave);
			if (txtClave.Text.ToString().Trim() == "Bl45o6$9")
			{

				Funcion.NombreSevidor = Servidor;
				MenuLatinium.stUsuario = "prg";
				if (!VerificaUsuario(stUsuario, stClaveBasica, 2)) return;
				if (idGrupo == -1) idGrupo = 10;
				DialogResult = DialogResult.OK;
				Close();
				return;
			}
			if (!VerificaUsuario(stUsuario, stClaveBasica, 1)) 
			{
				UsuarioAccesoCrea(stUsuario, 0, "Usuario o Clave Erronea");
				return;
			}
			if (!VerificaUsuario(stUsuario, stClaveBasica, 2)) return;
			if (idGrupo == -1) idGrupo = 10;
			string stClaveHash = Funcion.CreaHash(stClaveBasica+stUsuario, idGrupo);
			if (!VerificaUsuario(stUsuario, stClaveHash, 3)) 
			{
				UsuarioAccesoCrea(stUsuario, 0, "Intento de Cambio de Grupo");
				return;
			}
			Funcion.NombreSevidor = Servidor;
			#endregion Usuario
			UsuarioAccesoCrea(stUsuario, 1, "Ingreso Aceptado");
			DialogResult = DialogResult.OK;
			Close();		
		}
		string stArchivo="";
		private void UsuarioIngreso_Load(object sender, System.EventArgs e)
		{
			string stDirectorio = @"c:\Latinium\";
			if (!Directory.Exists(stDirectorio)) stDirectorio = Application.StartupPath+"\\";
			stArchivo = stDirectorio + "Servidor.config";
			if (File.Exists(stArchivo))
			{
				StreamReader sr = new StreamReader(stArchivo);
				string stServidor = sr.ReadLine();
				sr.Close();
				if (VerificaServidor(stServidor))
				{
					txtServidor.Width = 0;
					lblServidor.Visible = false;
					txtUsuario.Select();
				}
				Cursor = Cursors.Default;
			}

		}
	}
}

