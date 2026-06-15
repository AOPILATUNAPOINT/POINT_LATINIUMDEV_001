using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Text;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de Form5.
	/// </summary>
	public class Form5 : System.Windows.Forms.Form
	{
		private System.Windows.Forms.Button btnGenerarArchivo;
		private System.Windows.Forms.Button button1;
		private C1.Data.C1DataSet cdsSeteoC;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public Form5()
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
			this.btnGenerarArchivo = new System.Windows.Forms.Button();
			this.button1 = new System.Windows.Forms.Button();
			this.cdsSeteoC = new C1.Data.C1DataSet();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoC)).BeginInit();
			this.SuspendLayout();
			// 
			// btnGenerarArchivo
			// 
			this.btnGenerarArchivo.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.btnGenerarArchivo.Location = new System.Drawing.Point(8, 40);
			this.btnGenerarArchivo.Name = "btnGenerarArchivo";
			this.btnGenerarArchivo.Size = new System.Drawing.Size(184, 23);
			this.btnGenerarArchivo.TabIndex = 725;
			this.btnGenerarArchivo.Text = "&Generar Archivo Pichincha";
			this.btnGenerarArchivo.Click += new System.EventHandler(this.btnGenerarArchivo_Click);
			// 
			// button1
			// 
			this.button1.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.button1.Location = new System.Drawing.Point(200, 40);
			this.button1.Name = "button1";
			this.button1.Size = new System.Drawing.Size(184, 23);
			this.button1.TabIndex = 726;
			this.button1.Text = "&Generar Archivo Internacional";
			this.button1.Click += new System.EventHandler(this.button1_Click);
			// 
			// cdsSeteoC
			// 
			this.cdsSeteoC.BindingContextCtrl = this;
			this.cdsSeteoC.DataLibrary = "LibContabilidad";
			this.cdsSeteoC.DataLibraryUrl = "";
			this.cdsSeteoC.DataSetDef = "dsSeteoC";
			this.cdsSeteoC.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsSeteoC.Name = "cdsSeteoC";
			this.cdsSeteoC.SchemaClassName = "LibContabilidad.DataClass";
			this.cdsSeteoC.SchemaDef = null;
			// 
			// Form5
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(432, 70);
			this.Controls.Add(this.button1);
			this.Controls.Add(this.btnGenerarArchivo);
			this.Name = "Form5";
			this.Text = "Form5";
			this.Load += new System.EventHandler(this.Form5_Load);
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoC)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void ArchivoTXT(int idBanco)
		{
			#region Exporta
			try
			{
				SaveFileDialog sArchivo = new SaveFileDialog();

				sArchivo.Filter = "Documentos de texto (*.txt)|*.txt";
				sArchivo.FileName = sArchivo.FileName;
				sArchivo.OverwritePrompt = true;
				sArchivo.Title = "Guardar Lista Acreditaciones Archivos TXT";		
        
        string sBanco = "PICHINCHA";
        if (idBanco == 5) sBanco = "INTERNACIONAL";         

				if (sArchivo.ShowDialog() == DialogResult.OK)
				{
					using (SqlConnection oConexion = new SqlConnection("Password=Bl45o6$9;Persist Security Info=True;User ID= " + MenuLatinium.stUsuario + ";Initial Catalog=" + MenuLatinium.stDirFacturacion + ";Data Source=" + MenuLatinium.stDirServidor))
					{						
						#region Conexion
						SqlCommand oCmdActualiza = oConexion.CreateCommand();
						oCmdActualiza.Connection = oConexion;
						oCmdActualiza.CommandTimeout = 0;
						//string sSQL = string.Format("Select 'PA' As Pago, 'USD' As Moneda, REPLACE(Valor, '.', '') As Valor, 'CTA' As Titulo, Case idTipoPago When 1 Then 'AHO' When 2 Then 'CTE' End As TipoDeCuenta, NumeroDeCuenta As Cuenta, 'SUELDO' As Mensaje, 'C' As TipoDoc, Cedula, (LTRIM(RTRIM(Nombres)) + ' ' + LTRIM(RTRIM(Apellidos))) As Nombre From Utilidades Where idBanco = {0} And idTipoPago In (1, 2, 3) Order By Nombres, Apellidos", idBanco);
						string sSQL = string.Format("Select 'PA' As Pago, 'USD' As Moneda, REPLACE(ValorD, '.', '') As Valor, 'CTA' As Titulo, TipoCuenta As TipoDeCuenta, Cuenta, 'DECIMO CUARTO' As Mensaje, 'C' As TipoDoc, Cedula, LTRIM(RTRIM(Nombres)) As Nombre From DECIMOS Where Banco = '{0}' Order By Nombres", sBanco);
						oCmdActualiza.CommandText = sSQL;
						SqlDataAdapter oDADetalle = new SqlDataAdapter(oCmdActualiza);
						DataTable oDTDetalle = new DataTable();
						oDADetalle.Fill(oDTDetalle);	
						#endregion Conexion

						#region Pagos
						using (StreamWriter file = new StreamWriter(sArchivo.FileName, true))
						{						
							int i = 1;
							foreach (DataRow dr in oDTDetalle.Rows)
							{							
								string iLinea = dr["Pago"].ToString() + "\t" + i.ToString() + "\t" + dr["Moneda"].ToString() + "\t" + dr["Valor"].ToString() + "\t"
									+ dr["Titulo"].ToString() + "\t" + dr["TipoDeCuenta"].ToString() + "\t" + dr["Cuenta"].ToString() + "\t"
									+ dr["Mensaje"].ToString() + "\t" + dr["TipoDoc"].ToString() + "\t" + dr["Cedula"].ToString() + "\t" + dr["Nombre"].ToString();

								file.WriteLine(iLinea);

								i++;
							}					
						}
						#endregion Pagos					
					}
				}
			}
			catch (Exception exc)
			{
				MessageBox.Show(exc.Message);
			}
			#endregion Exporta
		}

		public Acceso miAcceso;

		private void UnloadMe()
		{
			this.Close();
		}

		private void Form5_Load(object sender, System.EventArgs e)
		{
			miAcceso = new Acceso(cdsSeteoC, "020803");

			if (!Funcion.Permiso("1505",cdsSeteoC))
			{				
				MessageBox.Show("No puede ingresar a Utilitarios", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				BeginInvoke(new MethodInvoker(UnloadMe));				
			}
		}

		private void btnGenerarArchivo_Click(object sender, System.EventArgs e)
		{
			ArchivoTXT(7);
		}

		private void button1_Click(object sender, System.EventArgs e)
		{
			ArchivoTXT(5);
		}
	}
}
