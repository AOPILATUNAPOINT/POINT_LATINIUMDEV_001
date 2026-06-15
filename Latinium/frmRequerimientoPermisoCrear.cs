using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;
using System.Data.SqlClient;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de frmRequerimientoPermisoCrear.
	/// </summary>
	public class frmRequerimientoPermisoCrear : System.Windows.Forms.Form
	{
		private System.Windows.Forms.Button button1;
		private System.Windows.Forms.Button button3;
		private System.Windows.Forms.GroupBox groupBox1;
		public System.Windows.Forms.Label label2;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtDescripcion;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public frmRequerimientoPermisoCrear()
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
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmRequerimientoPermisoCrear));
			Infragistics.Win.Appearance appearance1 = new Infragistics.Win.Appearance();
			this.button1 = new System.Windows.Forms.Button();
			this.button3 = new System.Windows.Forms.Button();
			this.groupBox1 = new System.Windows.Forms.GroupBox();
			this.label2 = new System.Windows.Forms.Label();
			this.txtDescripcion = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.groupBox1.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.txtDescripcion)).BeginInit();
			this.SuspendLayout();
			// 
			// button1
			// 
			this.button1.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.button1.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(230)), ((System.Byte)(230)), ((System.Byte)(230)));
			this.button1.CausesValidation = false;
			this.button1.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.button1.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(50)), ((System.Byte)(50)), ((System.Byte)(50)));
			this.button1.Image = ((System.Drawing.Image)(resources.GetObject("button1.Image")));
			this.button1.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.button1.Location = new System.Drawing.Point(92, 80);
			this.button1.Name = "button1";
			this.button1.Size = new System.Drawing.Size(80, 23);
			this.button1.TabIndex = 1068;
			this.button1.Text = "Salir";
			this.button1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// button3
			// 
			this.button3.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.button3.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(230)), ((System.Byte)(230)), ((System.Byte)(230)));
			this.button3.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.button3.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(50)), ((System.Byte)(50)), ((System.Byte)(50)));
			this.button3.Image = ((System.Drawing.Image)(resources.GetObject("button3.Image")));
			this.button3.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.button3.Location = new System.Drawing.Point(8, 80);
			this.button3.Name = "button3";
			this.button3.Size = new System.Drawing.Size(76, 23);
			this.button3.TabIndex = 1067;
			this.button3.Text = "Guardar";
			this.button3.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.button3.Click += new System.EventHandler(this.button3_Click);
			// 
			// groupBox1
			// 
			this.groupBox1.Controls.Add(this.label2);
			this.groupBox1.Controls.Add(this.txtDescripcion);
			this.groupBox1.Location = new System.Drawing.Point(8, 8);
			this.groupBox1.Name = "groupBox1";
			this.groupBox1.Size = new System.Drawing.Size(344, 60);
			this.groupBox1.TabIndex = 1066;
			this.groupBox1.TabStop = false;
			this.groupBox1.Text = "Información";
			// 
			// label2
			// 
			this.label2.AutoSize = true;
			this.label2.Location = new System.Drawing.Point(16, 24);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(74, 16);
			this.label2.TabIndex = 1055;
			this.label2.Text = "* Descripción:";
			this.label2.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtDescripcion
			// 
			appearance1.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtDescripcion.Appearance = appearance1;
			this.txtDescripcion.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtDescripcion.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtDescripcion.Location = new System.Drawing.Point(128, 24);
			this.txtDescripcion.Name = "txtDescripcion";
			this.txtDescripcion.Size = new System.Drawing.Size(200, 21);
			this.txtDescripcion.TabIndex = 1054;
			// 
			// frmRequerimientoPermisoCrear
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(360, 110);
			this.Controls.Add(this.button1);
			this.Controls.Add(this.button3);
			this.Controls.Add(this.groupBox1);
			this.Name = "frmRequerimientoPermisoCrear";
			this.Text = "Permiso";
			this.Load += new System.EventHandler(this.frmRequerimientoPermisoCrear_Load);
			this.groupBox1.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.txtDescripcion)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void frmRequerimientoPermisoCrear_Load(object sender, System.EventArgs e)
		{
		
		}

		private void button3_Click(object sender, System.EventArgs e)
		{
			#region Validación
			// Valida si el texto de la descripción es válido, con un mínimo de 5 caracteres y un máximo de 100.
			if (!Validacion.vbTexto(this.txtDescripcion, 5, 100, "Ingrese la descripción"))
				return;
			#endregion Validación

			// Obtener el texto de la descripción ingresado
			string descripcion = this.txtDescripcion.Text;

			// Asignar el valor predeterminado para el estado
			bool estado = true; // Por defecto el estado es 1 (true)

			// Usar la cadena de conexión definida en frmRequerimiento
			string connectionString = frmRequerimiento.sconexionPoint;

			// Usar SqlConnection para conectar con la base de datos
			using (SqlConnection connection = new SqlConnection(connectionString))
			{
				try
				{
					// Abrir la conexión
					connection.Open();

					// Crear el comando para ejecutar el procedimiento almacenado
					SqlCommand command = new SqlCommand("RequerimientoPermisoInsertar", connection);
					command.CommandType = CommandType.StoredProcedure;

					// Agregar los parámetros del procedimiento almacenado de manera explícita
					SqlParameter paramDescripcion = new SqlParameter("@Descripcion", SqlDbType.VarChar, 50);
					paramDescripcion.Value = descripcion;
					command.Parameters.Add(paramDescripcion);

					SqlParameter paramEstado = new SqlParameter("@Estado", SqlDbType.Bit);
					paramEstado.Value = estado; // El estado por defecto es true (1)
					command.Parameters.Add(paramEstado);

					// Ejecutar el procedimiento almacenado y obtener el ID generado
					object result = command.ExecuteScalar();

					// Verificar si el registro fue insertado correctamente
					if (result != null)
					{
						int idRequerimientoPermiso = Convert.ToInt32(result);
						MessageBox.Show("Registro insertado correctamente. ID: " + idRequerimientoPermiso);
						this.DialogResult = DialogResult.OK;
					}
					else
					{
						MessageBox.Show("Error al insertar el registro.");
					}
				}
				catch (Exception ex)
				{
					// Manejo de errores
					MessageBox.Show("Error al insertar el registro: " + ex.Message);
				}
			}
		}

	}
}
