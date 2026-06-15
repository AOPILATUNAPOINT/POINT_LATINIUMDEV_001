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
	/// Descripción breve de frmRequerimientoPoliticaCrear.
	/// </summary>
	public class frmRequerimientoPoliticaCrear : System.Windows.Forms.Form
	{
		private System.Windows.Forms.Button button1;
		private System.Windows.Forms.Button button3;
		private System.Windows.Forms.GroupBox groupBox1;
		public System.Windows.Forms.Label label2;
		public System.Windows.Forms.Label label12;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtDescripcion;
		private System.Windows.Forms.Label lblAsignado;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbDepartamento;
		private System.Windows.Forms.Label label4;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtMensaje;
		private System.Windows.Forms.Label label21;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtNumeroPolitica;
		public System.Windows.Forms.Label label1;
		private System.Windows.Forms.CheckBox chkResultado;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource udsRequerimientoDepartamentoCon;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public frmRequerimientoPoliticaCrear()
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
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmRequerimientoPoliticaCrear));
			Infragistics.Win.Appearance appearance1 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idRequerimientoDepartamento");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Descripcion");
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idRequerimientoDepartamento");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Descripcion");
			this.button1 = new System.Windows.Forms.Button();
			this.button3 = new System.Windows.Forms.Button();
			this.groupBox1 = new System.Windows.Forms.GroupBox();
			this.chkResultado = new System.Windows.Forms.CheckBox();
			this.label1 = new System.Windows.Forms.Label();
			this.label21 = new System.Windows.Forms.Label();
			this.txtNumeroPolitica = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.lblAsignado = new System.Windows.Forms.Label();
			this.cmbDepartamento = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.label4 = new System.Windows.Forms.Label();
			this.label2 = new System.Windows.Forms.Label();
			this.label12 = new System.Windows.Forms.Label();
			this.txtDescripcion = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtMensaje = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.udsRequerimientoDepartamentoCon = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.groupBox1.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.txtNumeroPolitica)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbDepartamento)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDescripcion)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtMensaje)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.udsRequerimientoDepartamentoCon)).BeginInit();
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
			this.button1.Location = new System.Drawing.Point(104, 216);
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
			this.button3.Location = new System.Drawing.Point(16, 216);
			this.button3.Name = "button3";
			this.button3.Size = new System.Drawing.Size(76, 23);
			this.button3.TabIndex = 1067;
			this.button3.Text = "Guardar";
			this.button3.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.button3.Click += new System.EventHandler(this.button3_Click);
			// 
			// groupBox1
			// 
			this.groupBox1.Controls.Add(this.chkResultado);
			this.groupBox1.Controls.Add(this.label1);
			this.groupBox1.Controls.Add(this.label21);
			this.groupBox1.Controls.Add(this.txtNumeroPolitica);
			this.groupBox1.Controls.Add(this.lblAsignado);
			this.groupBox1.Controls.Add(this.cmbDepartamento);
			this.groupBox1.Controls.Add(this.label4);
			this.groupBox1.Controls.Add(this.label2);
			this.groupBox1.Controls.Add(this.label12);
			this.groupBox1.Controls.Add(this.txtDescripcion);
			this.groupBox1.Controls.Add(this.txtMensaje);
			this.groupBox1.Location = new System.Drawing.Point(16, 16);
			this.groupBox1.Name = "groupBox1";
			this.groupBox1.Size = new System.Drawing.Size(352, 192);
			this.groupBox1.TabIndex = 1066;
			this.groupBox1.TabStop = false;
			this.groupBox1.Text = "Información";
			// 
			// chkResultado
			// 
			this.chkResultado.Location = new System.Drawing.Point(128, 160);
			this.chkResultado.Name = "chkResultado";
			this.chkResultado.Size = new System.Drawing.Size(16, 24);
			this.chkResultado.TabIndex = 1065;
			// 
			// label1
			// 
			this.label1.AutoSize = true;
			this.label1.Location = new System.Drawing.Point(16, 160);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(66, 16);
			this.label1.TabIndex = 1064;
			this.label1.Text = "* Resultado:";
			this.label1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label21
			// 
			this.label21.AutoSize = true;
			this.label21.BackColor = System.Drawing.Color.Transparent;
			this.label21.Location = new System.Drawing.Point(16, 96);
			this.label21.Name = "label21";
			this.label21.Size = new System.Drawing.Size(93, 16);
			this.label21.TabIndex = 1062;
			this.label21.Text = "* Número política:";
			// 
			// txtNumeroPolitica
			// 
			appearance1.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtNumeroPolitica.Appearance = appearance1;
			this.txtNumeroPolitica.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNumeroPolitica.FormatString = "";
			this.txtNumeroPolitica.Location = new System.Drawing.Point(128, 88);
			this.txtNumeroPolitica.MaskInput = "nn";
			this.txtNumeroPolitica.Name = "txtNumeroPolitica";
			this.txtNumeroPolitica.PromptChar = ' ';
			this.txtNumeroPolitica.Size = new System.Drawing.Size(200, 21);
			this.txtNumeroPolitica.SpinButtonDisplayStyle = Infragistics.Win.ButtonDisplayStyle.Always;
			this.txtNumeroPolitica.TabIndex = 1061;
			// 
			// lblAsignado
			// 
			this.lblAsignado.Location = new System.Drawing.Point(16, 24);
			this.lblAsignado.Name = "lblAsignado";
			this.lblAsignado.Size = new System.Drawing.Size(96, 23);
			this.lblAsignado.TabIndex = 1060;
			this.lblAsignado.Text = "* Departamento:";
			// 
			// cmbDepartamento
			// 
			appearance2.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbDepartamento.Appearance = appearance2;
			this.cmbDepartamento.CausesValidation = false;
			this.cmbDepartamento.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbDepartamento.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbDepartamento.DataSource = this.udsRequerimientoDepartamentoCon;
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn2.Header.Caption = "Descripción";
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridColumn2.Width = 218;
			ultraGridBand1.Columns.AddRange(new object[] {
																										 ultraGridColumn1,
																										 ultraGridColumn2});
			this.cmbDepartamento.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.cmbDepartamento.DisplayMember = "Descripcion";
			this.cmbDepartamento.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbDepartamento.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbDepartamento.Location = new System.Drawing.Point(128, 24);
			this.cmbDepartamento.MaxDropDownItems = 30;
			this.cmbDepartamento.Name = "cmbDepartamento";
			this.cmbDepartamento.Size = new System.Drawing.Size(200, 21);
			this.cmbDepartamento.TabIndex = 1058;
			this.cmbDepartamento.ValueMember = "idRequerimientoDepartamento";
			// 
			// label4
			// 
			this.label4.Location = new System.Drawing.Point(16, 24);
			this.label4.Name = "label4";
			this.label4.Size = new System.Drawing.Size(96, 23);
			this.label4.TabIndex = 1059;
			this.label4.Text = "* Departamento:";
			// 
			// label2
			// 
			this.label2.AutoSize = true;
			this.label2.Location = new System.Drawing.Point(16, 128);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(58, 16);
			this.label2.TabIndex = 1055;
			this.label2.Text = "* Mensaje:";
			this.label2.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label12
			// 
			this.label12.AutoSize = true;
			this.label12.Location = new System.Drawing.Point(16, 56);
			this.label12.Name = "label12";
			this.label12.Size = new System.Drawing.Size(74, 16);
			this.label12.TabIndex = 1053;
			this.label12.Text = "* Descripción:";
			this.label12.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtDescripcion
			// 
			appearance3.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtDescripcion.Appearance = appearance3;
			this.txtDescripcion.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtDescripcion.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtDescripcion.Location = new System.Drawing.Point(128, 56);
			this.txtDescripcion.Name = "txtDescripcion";
			this.txtDescripcion.Size = new System.Drawing.Size(200, 21);
			this.txtDescripcion.TabIndex = 1052;
			// 
			// txtMensaje
			// 
			appearance4.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtMensaje.Appearance = appearance4;
			this.txtMensaje.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtMensaje.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtMensaje.Location = new System.Drawing.Point(128, 128);
			this.txtMensaje.Name = "txtMensaje";
			this.txtMensaje.Size = new System.Drawing.Size(200, 21);
			this.txtMensaje.TabIndex = 1054;
			// 
			// udsRequerimientoDepartamentoCon
			// 
			ultraDataColumn1.DataType = typeof(int);
			this.udsRequerimientoDepartamentoCon.Band.Columns.AddRange(new object[] {
																																								ultraDataColumn1,
																																								ultraDataColumn2});
			// 
			// frmRequerimientoPoliticaCrear
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(384, 254);
			this.Controls.Add(this.button1);
			this.Controls.Add(this.button3);
			this.Controls.Add(this.groupBox1);
			this.Name = "frmRequerimientoPoliticaCrear";
			this.Text = "Política";
			this.Load += new System.EventHandler(this.frmRequerimientoPoliticaCrear_Load);
			this.groupBox1.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.txtNumeroPolitica)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbDepartamento)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDescripcion)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtMensaje)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.udsRequerimientoDepartamentoCon)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void frmRequerimientoPoliticaCrear_Load(object sender, System.EventArgs e)
		{
			frmRequerimientoConfiguracion.RequerimientoDepartamentoCon(this.cmbDepartamento);
		}

		private void button3_Click(object sender, System.EventArgs e)
		{
			#region Validación
			if (!Validacion.vbComboVacio(this.cmbDepartamento, "Ingrese el departamento")) return;
			if (!Validacion.vbTexto(this.txtDescripcion, 5, 500, "Ingrese la descripción")) return;
			if (!Validacion.vbCampoEnteroVacio(this.txtNumeroPolitica, "Ingrese el número de política", 1, 100)) return;
			if (!Validacion.vbTexto(this.txtMensaje, 5, 500, "Ingrese el mensaje")) return;
			#endregion Validación

			// Recoge los valores de los controles del formulario
			int idRequerimientoDepartamento = (int)this.cmbDepartamento.Value;
			string Descripcion = this.txtDescripcion.Text;
			int NumeroPolitica = (int)this.txtNumeroPolitica.Value;  // Cambiado de string a int
			string Mensaje = this.txtMensaje.Text;
			bool Estado = true;  // 1 se convierte en true
			bool Resultado = this.chkResultado.Checked;

			try
			{
				// Crea una conexión a la base de datos
				using (SqlConnection conn = new SqlConnection(frmRequerimiento.sconexionPoint))
				{
					// Abre la conexión
					conn.Open();

					// Define el procedimiento almacenado
					using (SqlCommand cmd = new SqlCommand("RequerimientoPoliticaInsertar", conn))
					{
						cmd.CommandType = CommandType.StoredProcedure;

						// Agrega los parámetros al comando usando SqlParameter
						cmd.Parameters.Add(new SqlParameter("@Descripcion", SqlDbType.VarChar, 500)).Value = Descripcion;
						cmd.Parameters.Add(new SqlParameter("@NumeroPolitica", SqlDbType.Int)).Value = NumeroPolitica;
						cmd.Parameters.Add(new SqlParameter("@Mensaje", SqlDbType.VarChar, 500)).Value = Mensaje;
						cmd.Parameters.Add(new SqlParameter("@Resultado", SqlDbType.Bit)).Value = Resultado;
						cmd.Parameters.Add(new SqlParameter("@Estado", SqlDbType.Bit)).Value = Estado;
						cmd.Parameters.Add(new SqlParameter("@idRequerimientoDepartamento", SqlDbType.Int)).Value = idRequerimientoDepartamento;

						// Ejecuta el procedimiento almacenado y obtiene el ID del nuevo registro
						int idRequerimientoPolitica = Convert.ToInt32(cmd.ExecuteScalar());

						// Muestra un mensaje de éxito
						MessageBox.Show("Registro insertado correctamente. ID: " + idRequerimientoPolitica, "Éxito", MessageBoxButtons.OK, MessageBoxIcon.Information);
						this.DialogResult = DialogResult.OK;
					}
				}
			}
			catch (Exception ex)
			{
				// Muestra un mensaje de error
				MessageBox.Show("Error al insertar el registro: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
			}
		}


		private void button1_Click(object sender, System.EventArgs e)
		{
			this.Close();
		}
	}
}
