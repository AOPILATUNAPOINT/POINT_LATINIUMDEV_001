using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de Form1.
	/// </summary>
	public class DBConfiguracion : DevExpress.XtraEditors.XtraForm
	{
		private DevExpress.XtraEditors.SimpleButton btEmpresas;
		private DevExpress.XtraEditors.SimpleButton btUsuarios;
		private DevExpress.XtraEditors.SimpleButton simpleButton1;
		private DevExpress.XtraEditors.SimpleButton simpleButton2;
		private DevExpress.XtraEditors.SimpleButton simpleButton3;
		private DevExpress.XtraEditors.SimpleButton btAccesos;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public DBConfiguracion()
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
				if (components != null) 
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
			this.btEmpresas = new DevExpress.XtraEditors.SimpleButton();
			this.btUsuarios = new DevExpress.XtraEditors.SimpleButton();
			this.simpleButton1 = new DevExpress.XtraEditors.SimpleButton();
			this.simpleButton2 = new DevExpress.XtraEditors.SimpleButton();
			this.simpleButton3 = new DevExpress.XtraEditors.SimpleButton();
			this.btAccesos = new DevExpress.XtraEditors.SimpleButton();
			this.SuspendLayout();
			// 
			// btEmpresas
			// 
			this.btEmpresas.Location = new System.Drawing.Point(86, 39);
			this.btEmpresas.Name = "btEmpresas";
			this.btEmpresas.Size = new System.Drawing.Size(173, 50);
			this.btEmpresas.TabIndex = 0;
			this.btEmpresas.Text = "Empresas";
			this.btEmpresas.Click += new System.EventHandler(this.btEmpresas_Click);
			// 
			// btUsuarios
			// 
			this.btUsuarios.Location = new System.Drawing.Point(384, 39);
			this.btUsuarios.Name = "btUsuarios";
			this.btUsuarios.Size = new System.Drawing.Size(173, 50);
			this.btUsuarios.TabIndex = 1;
			this.btUsuarios.Text = "Crear Usuarios";
			this.btUsuarios.Click += new System.EventHandler(this.btUsuarios_Click);
			// 
			// simpleButton1
			// 
			this.simpleButton1.Location = new System.Drawing.Point(384, 158);
			this.simpleButton1.Name = "simpleButton1";
			this.simpleButton1.Size = new System.Drawing.Size(173, 49);
			this.simpleButton1.TabIndex = 3;
			this.simpleButton1.Text = "Configuración de Claves";
			this.simpleButton1.Click += new System.EventHandler(this.simpleButton1_Click);
			// 
			// simpleButton2
			// 
			this.simpleButton2.Location = new System.Drawing.Point(86, 155);
			this.simpleButton2.Name = "simpleButton2";
			this.simpleButton2.Size = new System.Drawing.Size(173, 50);
			this.simpleButton2.TabIndex = 4;
			this.simpleButton2.Text = "Usuarios con Empresas";
			this.simpleButton2.Click += new System.EventHandler(this.simpleButton2_Click);
			// 
			// simpleButton3
			// 
			this.simpleButton3.Location = new System.Drawing.Point(86, 274);
			this.simpleButton3.Name = "simpleButton3";
			this.simpleButton3.Size = new System.Drawing.Size(173, 49);
			this.simpleButton3.TabIndex = 5;
			this.simpleButton3.Text = "Copiar";
			this.simpleButton3.Visible = false;
			this.simpleButton3.Click += new System.EventHandler(this.simpleButton3_Click);
			// 
			// btAccesos
			// 
			this.btAccesos.Location = new System.Drawing.Point(384, 274);
			this.btAccesos.Name = "btAccesos";
			this.btAccesos.Size = new System.Drawing.Size(173, 49);
			this.btAccesos.TabIndex = 6;
			this.btAccesos.Text = "Accesos";
			this.btAccesos.Click += new System.EventHandler(this.btAccesos_Click);
			// 
			// DBConfiguracion
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(6, 16);
			this.ClientSize = new System.Drawing.Size(624, 384);
			this.Controls.Add(this.btAccesos);
			this.Controls.Add(this.simpleButton3);
			this.Controls.Add(this.simpleButton2);
			this.Controls.Add(this.simpleButton1);
			this.Controls.Add(this.btUsuarios);
			this.Controls.Add(this.btEmpresas);
			this.Name = "DBConfiguracion";
			this.Text = "Configuración de Seguridades";
			this.Load += new System.EventHandler(this.MenuInicio_Load);
			this.ResumeLayout(false);

		}
		#endregion

		/// <summary>
		/// Punto de entrada principal de la aplicación.
		/// </summary>
//		[STAThread]
//		static void Main() 
//		{
//			Application.Run(new MenuInicio());
//		}

		private void MenuInicio_Load(object sender, System.EventArgs e)
		{
//			DBIngresoUsuario miInicio = new DBIngresoUsuario();
//			if (miInicio.ShowDialog() != DialogResult.OK) 
//			{
//				Application.Exit();
//				return;
//			}

		}

		private void btEmpresas_Click(object sender, System.EventArgs e)
		{
			DBEmpresaBD miEmpresa = new DBEmpresaBD();
			miEmpresa.Show();
			
		}

		private void btUsuarios_Click(object sender, System.EventArgs e)
		{
			DBUsuario miUsuario = new DBUsuario();
			miUsuario.Show();
		}

		private void simpleButton1_Click(object sender, System.EventArgs e)
		{
			DBUsuarioSeguridad miUsuario = new DBUsuarioSeguridad();
			miUsuario.Show();
		}


		private void simpleButton2_Click(object sender, System.EventArgs e)
		{
			DBUsuarioEmpresa miEmpresa = new DBUsuarioEmpresa();
			miEmpresa.Show();
		}

		private void simpleButton3_Click(object sender, System.EventArgs e)
		{
			DBCopiar miEmpresa = new DBCopiar();
			miEmpresa.Show();
		}

		private void btAccesos_Click(object sender, System.EventArgs e)
		{
			DBAccesos miAcceso = new DBAccesos();
			miAcceso.Show();
		}

	}
}
