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
	/// Descripción breve de frmRequerimientoRolCrear.
	/// </summary>
	public class frmRequerimientoRolCrear : System.Windows.Forms.Form
	{
		private System.Windows.Forms.Button button1;
		private System.Windows.Forms.Button button3;
		private System.Windows.Forms.GroupBox groupBox1;
		public System.Windows.Forms.Label label2;
		public System.Windows.Forms.Label label12;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtDescripcion;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtSubtitulo;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public frmRequerimientoRolCrear()
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
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmRequerimientoRolCrear));
			Infragistics.Win.Appearance appearance1 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			this.button1 = new System.Windows.Forms.Button();
			this.button3 = new System.Windows.Forms.Button();
			this.groupBox1 = new System.Windows.Forms.GroupBox();
			this.label2 = new System.Windows.Forms.Label();
			this.label12 = new System.Windows.Forms.Label();
			this.txtDescripcion = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtSubtitulo = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.groupBox1.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.txtDescripcion)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtSubtitulo)).BeginInit();
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
			this.button1.Location = new System.Drawing.Point(96, 112);
			this.button1.Name = "button1";
			this.button1.Size = new System.Drawing.Size(80, 23);
			this.button1.TabIndex = 1068;
			this.button1.Text = "Salir";
			this.button1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.button1.Click += new System.EventHandler(this.button1_Click);
			// 
			// button3
			// 
			this.button3.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.button3.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(230)), ((System.Byte)(230)), ((System.Byte)(230)));
			this.button3.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.button3.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(50)), ((System.Byte)(50)), ((System.Byte)(50)));
			this.button3.Image = ((System.Drawing.Image)(resources.GetObject("button3.Image")));
			this.button3.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.button3.Location = new System.Drawing.Point(8, 112);
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
			this.groupBox1.Controls.Add(this.label12);
			this.groupBox1.Controls.Add(this.txtDescripcion);
			this.groupBox1.Controls.Add(this.txtSubtitulo);
			this.groupBox1.Location = new System.Drawing.Point(8, 8);
			this.groupBox1.Name = "groupBox1";
			this.groupBox1.Size = new System.Drawing.Size(352, 96);
			this.groupBox1.TabIndex = 1066;
			this.groupBox1.TabStop = false;
			this.groupBox1.Text = "Información";
			this.groupBox1.Enter += new System.EventHandler(this.groupBox1_Enter);
			// 
			// label2
			// 
			this.label2.AutoSize = true;
			this.label2.Location = new System.Drawing.Point(16, 56);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(59, 16);
			this.label2.TabIndex = 1055;
			this.label2.Text = "* Subtitulo:";
			this.label2.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label12
			// 
			this.label12.AutoSize = true;
			this.label12.Location = new System.Drawing.Point(16, 24);
			this.label12.Name = "label12";
			this.label12.Size = new System.Drawing.Size(74, 16);
			this.label12.TabIndex = 1053;
			this.label12.Text = "* Descripción:";
			this.label12.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
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
			this.txtDescripcion.TabIndex = 1052;
			// 
			// txtSubtitulo
			// 
			appearance2.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtSubtitulo.Appearance = appearance2;
			this.txtSubtitulo.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtSubtitulo.Location = new System.Drawing.Point(128, 56);
			this.txtSubtitulo.Name = "txtSubtitulo";
			this.txtSubtitulo.Size = new System.Drawing.Size(200, 21);
			this.txtSubtitulo.TabIndex = 1054;
			// 
			// frmRequerimientoRolCrear
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(376, 142);
			this.Controls.Add(this.button1);
			this.Controls.Add(this.button3);
			this.Controls.Add(this.groupBox1);
			this.Name = "frmRequerimientoRolCrear";
			this.Text = "Rol";
			this.Load += new System.EventHandler(this.frmRequerimientoRolCrear_Load);
			this.groupBox1.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.txtDescripcion)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtSubtitulo)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void frmRequerimientoRolCrear_Load(object sender, System.EventArgs e)
		{
		
		}

		private void groupBox1_Enter(object sender, System.EventArgs e)
		{
		
		}

		private void button1_Click(object sender, System.EventArgs e)
		{
			this.Close();
		}

		private void button3_Click(object sender, EventArgs e)
		{
			#region Validación
			if (!Validacion.vbTexto(this.txtDescripcion, 5, 50, "Ingrese la descripción")) return;
			if (!Validacion.vbTexto(this.txtSubtitulo, 5, 50, "Ingrese el subtítulo")) return;
			#endregion Validación

			// Definir la cadena de conexión
			string connectionString = frmRequerimiento.sconexionPoint;

			// Crear la conexión y el comando
			using (SqlConnection connection = new SqlConnection(connectionString))
			{
				try
				{
					// Crear el comando para llamar al procedimiento almacenado
					using (SqlCommand command = new SqlCommand("RequerimientoRolInsertar", connection))
					{
						command.CommandType = CommandType.StoredProcedure;

						// Añadir los parámetros al comando
						command.Parameters.Add(new SqlParameter("@Descripcion", SqlDbType.VarChar, 50)).Value = this.txtDescripcion.Text;
						command.Parameters.Add(new SqlParameter("@Estado", SqlDbType.Bit)).Value = 1;// Valor predeterminado para Estado 
						command.Parameters.Add(new SqlParameter("@Subtitulo", SqlDbType.VarChar, 50)).Value = this.txtSubtitulo.Text;

						// Abrir la conexión
						connection.Open();

						// Ejecutar el procedimiento almacenado y obtener el ID del nuevo registro
						int idRequerimientoRol = Convert.ToInt32(command.ExecuteScalar());

						// Mostrar un mensaje de éxito
						MessageBox.Show("Registro insertado con éxito. ID: " + idRequerimientoRol, "Éxito", MessageBoxButtons.OK, MessageBoxIcon.Information);
						this.DialogResult = DialogResult.OK;
					}
				}
				catch (SqlException sqlEx)
				{
					// Manejo específico de errores SQL
					MessageBox.Show("Error en la base de datos: " + sqlEx.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
				}
				catch (Exception ex)
				{
					// Manejo general de errores
					MessageBox.Show("Error al insertar el registro: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
				}
			}
		}
	}
}
