using System;
using System.IO;
using System.Data.SqlClient;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de GeneraSQL.
	/// </summary>
	public class GeneraSQL : DevExpress.XtraEditors.XtraForm
	{
		private System.Windows.Forms.TextBox txtSQL;
		private System.Windows.Forms.TextBox txtSalida;
		private System.Windows.Forms.Button btGenerar;
		private C1.Data.C1DataSet cdsBanco;
		private System.Windows.Forms.Button btReducir;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public GeneraSQL()
		{
			//
			// Necesario para admitir el Diseñador de Windows Forms
			//
			InitializeComponent();

			//
			// TODO: Agregar código de constructor después de llamar a InitializeComponent
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

		#region Windows Form Designer generated code
		/// <summary>
		/// Método necesario para admitir el Diseñador, no se puede modificar
		/// el contenido del método con el editor de código.
		/// </summary>
		private void InitializeComponent()
		{
			this.txtSQL = new System.Windows.Forms.TextBox();
			this.txtSalida = new System.Windows.Forms.TextBox();
			this.btGenerar = new System.Windows.Forms.Button();
			this.cdsBanco = new C1.Data.C1DataSet();
			this.btReducir = new System.Windows.Forms.Button();
			((System.ComponentModel.ISupportInitialize)(this.cdsBanco)).BeginInit();
			this.SuspendLayout();
			// 
			// txtSQL
			// 
			this.txtSQL.Location = new System.Drawing.Point(16, 32);
			this.txtSQL.Multiline = true;
			this.txtSQL.Name = "txtSQL";
			this.txtSQL.Size = new System.Drawing.Size(376, 104);
			this.txtSQL.TabIndex = 0;
			this.txtSQL.Text = "INSERT INTO Banco (idBanco, Nombre) VALUES(100, \'A\')";
			// 
			// txtSalida
			// 
			this.txtSalida.Location = new System.Drawing.Point(16, 144);
			this.txtSalida.Multiline = true;
			this.txtSalida.Name = "txtSalida";
			this.txtSalida.Size = new System.Drawing.Size(376, 112);
			this.txtSalida.TabIndex = 1;
			this.txtSalida.Text = "";
			// 
			// btGenerar
			// 
			this.btGenerar.Location = new System.Drawing.Point(24, 8);
			this.btGenerar.Name = "btGenerar";
			this.btGenerar.TabIndex = 2;
			this.btGenerar.Text = "Generar";
			this.btGenerar.Click += new System.EventHandler(this.btGenerar_Click);
			// 
			// cdsBanco
			// 
			this.cdsBanco.BindingContextCtrl = this;
			this.cdsBanco.DataLibrary = "LibContabilidad";
			this.cdsBanco.DataLibraryUrl = "";
			this.cdsBanco.DataSetDef = "dsCuentaCorriente";
			this.cdsBanco.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsBanco.Name = "cdsBanco";
			this.cdsBanco.SchemaClassName = "LibContabilidad.DataClass";
			this.cdsBanco.SchemaDef = null;
			// 
			// btReducir
			// 
			this.btReducir.Location = new System.Drawing.Point(256, 8);
			this.btReducir.Name = "btReducir";
			this.btReducir.TabIndex = 3;
			this.btReducir.Text = "&Reducir";
			this.btReducir.Click += new System.EventHandler(this.btReducir_Click);
			// 
			// GeneraSQL
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(408, 266);
			this.Controls.Add(this.btReducir);
			this.Controls.Add(this.btGenerar);
			this.Controls.Add(this.txtSalida);
			this.Controls.Add(this.txtSQL);
			this.Name = "GeneraSQL";
			this.Text = "GeneraSQL";
			this.Closing += new System.ComponentModel.CancelEventHandler(this.GeneraSQL_Closing);
			this.Load += new System.EventHandler(this.GeneraSQL_Load);
			((System.ComponentModel.ISupportInitialize)(this.cdsBanco)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void GeneraSQL_Load(object sender, System.EventArgs e)
		{
			Cursor = Cursors.Default;
			string stAudita = "Exec AuditaCrear 53, 6, 'Genera SQL'";
			Funcion.EjecutaSQL(cdsBanco, stAudita, true);
		}

		private void btGenerar_Click(object sender, System.EventArgs e)
		{
			SqlCommand cmdActualiza;
			this.cdsBanco.Schema.Connections[0].Open();
			C1.Data.SchemaObjects.Connection miConeccion = 
				this.cdsBanco.Schema.Connections[0];

			miConeccion = this.cdsBanco.Schema.Connections[0];
			string stSelect = this.txtSQL.Text.ToString();

			cmdActualiza = new SqlCommand (stSelect
				, (SqlConnection) miConeccion.DbConnection);
			cmdActualiza.Transaction = (SqlTransaction)
				miConeccion.DbTransaction;
			try
			{
				cmdActualiza.ExecuteNonQuery();
			}
			catch(System.Exception ex)
			{
				MessageBox.Show(ex.Message);
			}

			this.cdsBanco.Schema.Connections[0].Close();
			MessageBox.Show("Comando Ejecutado");
		
		}

		private void btOsql_Click(object sender, System.EventArgs e)
		{
			string stDir = @"c:\";
			string stOrigen = stDir + "ActualizaVersion.sql";
			string stTemp = stDir + "Entrada.Sql";
			if (!File.Exists(stOrigen))
			{
				MessageBox.Show("No existe archivo " + stOrigen);
				return;
			}
			string stLinea;
			StreamReader sr = new StreamReader(stOrigen);
			StreamWriter sw = new StreamWriter(stTemp);
			while ((stLinea = sr.ReadLine()) != null) 
      {
				stLinea = stLinea.Replace("LatiniumSQL", MenuLatinium.stDirConta);
				sw.WriteLine(stLinea);
      }
			sr.Close();
			sw.Close();
			try
			{
				string stParametro = "/U infoelect /P Bl45o6$9 /d ";
				stParametro += MenuLatinium.stDirConta;
				stParametro += " /i " + stTemp + " /o " + stDir +"salida.txt";
				System.Diagnostics.Process.Start("osql", @stParametro);
			}
			catch(System.Exception ex)
			{
				MessageBox.Show(ex.Message, "Comandos osql");
				return;
			}
//			File.Delete(stTemp);
			MessageBox.Show("Base de Datos Actualizada", "Información");
			System.Diagnostics.Process.Start("NotePad.exe", stDir + "Salida.txt");
		}

		private void btReducir_Click(object sender, System.EventArgs e)
		{
			string stSelect = "DBCC SHRINKDATABASE (" + MenuLatinium.stDirFacturacion + ", 10)";
			Funcion.EjecutaSQL(cdsBanco, stSelect);
		}

		private void GeneraSQL_Closing(object sender, System.ComponentModel.CancelEventArgs e)
		{
			string stAudita = "Exec AuditaCrear 53, 7, 'Genera SQL'";
			Funcion.EjecutaSQL(cdsBanco, stAudita, true);
		}
	}
}
