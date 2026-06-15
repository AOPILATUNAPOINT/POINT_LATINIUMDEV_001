using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.IO;
using System.Data.SqlClient;
using System.Data;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de frmRequerimientoEstadoCrear.
	/// </summary>
	public class frmRequerimientoEstadoCrear : System.Windows.Forms.Form
	{
		private System.Windows.Forms.Button button1;
		private System.Windows.Forms.Button button3;
		private System.Windows.Forms.GroupBox groupBox1;
		public System.Windows.Forms.Label label12;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtDescripcion;
		private System.Windows.Forms.GroupBox groupBox2;
		private System.Windows.Forms.Button btnCargar;
		private System.Windows.Forms.PictureBox imgEstado;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtImagen;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public frmRequerimientoEstadoCrear()
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
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmRequerimientoEstadoCrear));
			Infragistics.Win.Appearance appearance1 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			this.button1 = new System.Windows.Forms.Button();
			this.button3 = new System.Windows.Forms.Button();
			this.groupBox1 = new System.Windows.Forms.GroupBox();
			this.label12 = new System.Windows.Forms.Label();
			this.txtDescripcion = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.groupBox2 = new System.Windows.Forms.GroupBox();
			this.imgEstado = new System.Windows.Forms.PictureBox();
			this.txtImagen = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.btnCargar = new System.Windows.Forms.Button();
			this.groupBox1.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.txtDescripcion)).BeginInit();
			this.groupBox2.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.txtImagen)).BeginInit();
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
			this.button1.Location = new System.Drawing.Point(104, 160);
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
			this.button3.Location = new System.Drawing.Point(16, 160);
			this.button3.Name = "button3";
			this.button3.Size = new System.Drawing.Size(76, 23);
			this.button3.TabIndex = 1067;
			this.button3.Text = "Guardar";
			this.button3.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.button3.Click += new System.EventHandler(this.button3_Click);
			// 
			// groupBox1
			// 
			this.groupBox1.Controls.Add(this.label12);
			this.groupBox1.Controls.Add(this.txtDescripcion);
			this.groupBox1.Controls.Add(this.groupBox2);
			this.groupBox1.Location = new System.Drawing.Point(16, 8);
			this.groupBox1.Name = "groupBox1";
			this.groupBox1.Size = new System.Drawing.Size(392, 144);
			this.groupBox1.TabIndex = 1066;
			this.groupBox1.TabStop = false;
			this.groupBox1.Text = "Información usuario";
			// 
			// label12
			// 
			this.label12.AutoSize = true;
			this.label12.Location = new System.Drawing.Point(16, 32);
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
			this.txtDescripcion.Location = new System.Drawing.Point(128, 32);
			this.txtDescripcion.Name = "txtDescripcion";
			this.txtDescripcion.Size = new System.Drawing.Size(248, 21);
			this.txtDescripcion.TabIndex = 1052;
			// 
			// groupBox2
			// 
			this.groupBox2.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.groupBox2.Controls.Add(this.imgEstado);
			this.groupBox2.Controls.Add(this.txtImagen);
			this.groupBox2.Controls.Add(this.btnCargar);
			this.groupBox2.Location = new System.Drawing.Point(16, 64);
			this.groupBox2.Name = "groupBox2";
			this.groupBox2.Size = new System.Drawing.Size(360, 64);
			this.groupBox2.TabIndex = 1054;
			this.groupBox2.TabStop = false;
			this.groupBox2.Text = "* Imagen";
			// 
			// imgEstado
			// 
			this.imgEstado.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
			this.imgEstado.Location = new System.Drawing.Point(16, 24);
			this.imgEstado.Name = "imgEstado";
			this.imgEstado.Size = new System.Drawing.Size(232, 24);
			this.imgEstado.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage;
			this.imgEstado.TabIndex = 1055;
			this.imgEstado.TabStop = false;
			// 
			// txtImagen
			// 
			appearance2.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtImagen.Appearance = appearance2;
			this.txtImagen.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtImagen.Enabled = false;
			this.txtImagen.Location = new System.Drawing.Point(16, 24);
			this.txtImagen.MaxLength = 50;
			this.txtImagen.Name = "txtImagen";
			this.txtImagen.Size = new System.Drawing.Size(232, 21);
			this.txtImagen.TabIndex = 1039;
			this.txtImagen.WordWrap = false;
			// 
			// btnCargar
			// 
			this.btnCargar.Location = new System.Drawing.Point(264, 24);
			this.btnCargar.Name = "btnCargar";
			this.btnCargar.Size = new System.Drawing.Size(80, 23);
			this.btnCargar.TabIndex = 1037;
			this.btnCargar.Text = "Cargar";
			this.btnCargar.Click += new System.EventHandler(this.btnCargar_Click);
			// 
			// frmRequerimientoEstadoCrear
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(424, 190);
			this.Controls.Add(this.button1);
			this.Controls.Add(this.button3);
			this.Controls.Add(this.groupBox1);
			this.Name = "frmRequerimientoEstadoCrear";
			this.Text = "Estado";
			this.Load += new System.EventHandler(this.frmRequerimientoEstadoCrear_Load);
			this.groupBox1.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.txtDescripcion)).EndInit();
			this.groupBox2.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.txtImagen)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void frmRequerimientoEstadoCrear_Load(object sender, System.EventArgs e)
		{
		
		}

		private void button3_Click(object sender, System.EventArgs e)
		{
			EstadoRequerimientoInsertar();
		}

		// Variable global para almacenar el nombre del archivo seleccionado
		private string nombreArchivoGlobal;
		// Método del botón de carga
		private void btnCargar_Click(object sender, EventArgs e)
		{
			CargarImagen();
		}

		private void CargarImagen()
		{
			// Crear un OpenFileDialog para seleccionar una imagen
			OpenFileDialog openFileDialog = new OpenFileDialog();
    
			// Configuración del diálogo de selección de archivos
			openFileDialog.Title = "Seleccionar imagen";
			openFileDialog.Filter = "Archivos de imagen|*.jpg;*.jpeg;*.png;*.bmp|Todos los archivos|*.*";
    
			// Mostrar el diálogo de selección de archivo
			if (openFileDialog.ShowDialog() == DialogResult.OK)
			{
				try
				{
					// Cargar la imagen seleccionada en el PictureBox 'imgEstado'
					imgEstado.Image = Image.FromFile(openFileDialog.FileName);
					imgEstado.SizeMode = PictureBoxSizeMode.StretchImage;  // Ajustar la imagen al tamaño del PictureBox

					// Obtener el nombre del archivo seleccionado
					nombreArchivoGlobal = Path.GetFileName(openFileDialog.FileName);

					// Aquí puedes realizar cualquier otra acción que desees después de cargar la imagen
				}
				catch (Exception ex)
				{
					// Manejo de errores si no se puede cargar la imagen
					MessageBox.Show("Error al cargar la imagen: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
				}
			}
		}

		private void button1_Click(object sender, System.EventArgs e)
		{
			this.Close();
		}

		// Método para guardar la imagen en la carpeta compartida y la ruta en la base de datos usando un Stored Procedure
		private void EstadoRequerimientoInsertar()
		{
			#region Validación
			if (!Validacion.vbTexto(this.txtDescripcion, 5, 100, "Ingrese la descripción")) return;
			#endregion Validación

			string Descripcion = "";
			bool Estado = true;

			Descripcion = this.txtDescripcion.Text;

			try
			{
				// Verificar que se haya seleccionado un archivo
				if (nombreArchivoGlobal == null || nombreArchivoGlobal == "" || nombreArchivoGlobal.Length == 0)
				{
					MessageBox.Show("No se ha seleccionado ninguna imagen.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
					return;
				}

				// Ruta de la carpeta compartida
				string rutaCompartida = frmRequerimientoConfiguracion.RutaCarpetaCompartidaImagenesEstados;

				// Generar la ruta completa donde se guardará la imagen
				string rutaGuardado = Path.Combine(rutaCompartida, nombreArchivoGlobal);

				// Guardar la imagen en la carpeta compartida
				File.Copy(nombreArchivoGlobal, rutaGuardado, true);

				// Ahora guardar la ruta de la imagen en la base de datos
				using (SqlConnection conn = new SqlConnection(frmRequerimiento.sconexionPoint))
				{
					conn.Open();

					using (SqlCommand cmd = new SqlCommand("EstadoRequerimientoInsertar", conn))
					{
						cmd.CommandType = CommandType.StoredProcedure;

						// Asumiendo que tienes otros valores como Descripcion y Estado que necesitas pasar al stored procedure
						cmd.Parameters.Add(new SqlParameter("@Descripcion", System.Data.SqlDbType.VarChar, 50)).Value = Descripcion;
						cmd.Parameters.Add(new SqlParameter("@Estado", System.Data.SqlDbType.Bit)).Value = Estado;
						cmd.Parameters.Add(new SqlParameter("@Imagen", System.Data.SqlDbType.VarChar, 500)).Value = frmRequerimientoConfiguracion.RutaCarpetaCompartidaImagenesEstados + nombreArchivoGlobal;

						// Ejecución del procedimiento almacenado
						cmd.ExecuteNonQuery();

						// Mostrar mensaje de éxito
						MessageBox.Show("Estado registrado correctamente", "Éxito", MessageBoxButtons.OK, MessageBoxIcon.Information);
						this.DialogResult = DialogResult.OK;

					}
				}
			}
			catch (Exception ex)
			{
				MessageBox.Show("Error al guardar la ruta en la base de datos: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
			}
		}

	}
}
