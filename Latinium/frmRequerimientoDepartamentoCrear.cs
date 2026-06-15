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
	/// Descripción breve de frmRequerimientoDepartamentoCrear.
	/// </summary>
	public class frmRequerimientoDepartamentoCrear : System.Windows.Forms.Form
	{
		private System.Windows.Forms.GroupBox groupBox1;
		public System.Windows.Forms.Label label12;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtDescripcion;
		private System.Windows.Forms.Label label21;
		private System.Windows.Forms.Label label1;
		private System.Windows.Forms.Label label2;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtGrupo;
		private System.Windows.Forms.CheckBox chkChecklist;
		private System.Windows.Forms.Button btnGuardar;
		private System.Windows.Forms.Button btnSalir;
		private System.Windows.Forms.Label label3;
		private System.Windows.Forms.CheckBox chkPantalla;
		private System.Windows.Forms.Label label4;
		private System.Windows.Forms.Label label5;
		private System.Windows.Forms.CheckBox chkDescripcion;
		private System.Windows.Forms.CheckBox chkAsignacion;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource udsRequerimientoCompraNumeroObtenerTodos;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbTipoFactura;
		private System.Windows.Forms.Label label6;
		private System.Windows.Forms.CheckBox chkTodoUsuario;
		private System.Windows.Forms.Label bDepartamento;
		private System.Windows.Forms.CheckBox chkbDepartamento;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public frmRequerimientoDepartamentoCrear()
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
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmRequerimientoDepartamentoCrear));
			Infragistics.Win.Appearance appearance1 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idTipoFactura");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idTipoFactura");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			this.btnSalir = new System.Windows.Forms.Button();
			this.btnGuardar = new System.Windows.Forms.Button();
			this.groupBox1 = new System.Windows.Forms.GroupBox();
			this.label6 = new System.Windows.Forms.Label();
			this.chkTodoUsuario = new System.Windows.Forms.CheckBox();
			this.cmbTipoFactura = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.udsRequerimientoCompraNumeroObtenerTodos = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.label5 = new System.Windows.Forms.Label();
			this.chkAsignacion = new System.Windows.Forms.CheckBox();
			this.chkDescripcion = new System.Windows.Forms.CheckBox();
			this.label3 = new System.Windows.Forms.Label();
			this.chkPantalla = new System.Windows.Forms.CheckBox();
			this.label2 = new System.Windows.Forms.Label();
			this.chkChecklist = new System.Windows.Forms.CheckBox();
			this.label1 = new System.Windows.Forms.Label();
			this.label21 = new System.Windows.Forms.Label();
			this.txtGrupo = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label12 = new System.Windows.Forms.Label();
			this.txtDescripcion = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label4 = new System.Windows.Forms.Label();
			this.bDepartamento = new System.Windows.Forms.Label();
			this.chkbDepartamento = new System.Windows.Forms.CheckBox();
			this.groupBox1.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.cmbTipoFactura)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.udsRequerimientoCompraNumeroObtenerTodos)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtGrupo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDescripcion)).BeginInit();
			this.SuspendLayout();
			// 
			// btnSalir
			// 
			this.btnSalir.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btnSalir.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(230)), ((System.Byte)(230)), ((System.Byte)(230)));
			this.btnSalir.CausesValidation = false;
			this.btnSalir.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.btnSalir.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(50)), ((System.Byte)(50)), ((System.Byte)(50)));
			this.btnSalir.Image = ((System.Drawing.Image)(resources.GetObject("btnSalir.Image")));
			this.btnSalir.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnSalir.Location = new System.Drawing.Point(384, 48);
			this.btnSalir.Name = "btnSalir";
			this.btnSalir.Size = new System.Drawing.Size(80, 23);
			this.btnSalir.TabIndex = 1072;
			this.btnSalir.Text = "Salir";
			this.btnSalir.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnSalir.Click += new System.EventHandler(this.button1_Click);
			// 
			// btnGuardar
			// 
			this.btnGuardar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btnGuardar.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(230)), ((System.Byte)(230)), ((System.Byte)(230)));
			this.btnGuardar.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.btnGuardar.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(50)), ((System.Byte)(50)), ((System.Byte)(50)));
			this.btnGuardar.Image = ((System.Drawing.Image)(resources.GetObject("btnGuardar.Image")));
			this.btnGuardar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnGuardar.Location = new System.Drawing.Point(384, 16);
			this.btnGuardar.Name = "btnGuardar";
			this.btnGuardar.Size = new System.Drawing.Size(80, 23);
			this.btnGuardar.TabIndex = 1071;
			this.btnGuardar.Text = "Guardar";
			this.btnGuardar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnGuardar.Click += new System.EventHandler(this.btnGuardar_Click);
			// 
			// groupBox1
			// 
			this.groupBox1.Controls.Add(this.chkbDepartamento);
			this.groupBox1.Controls.Add(this.bDepartamento);
			this.groupBox1.Controls.Add(this.label6);
			this.groupBox1.Controls.Add(this.chkTodoUsuario);
			this.groupBox1.Controls.Add(this.cmbTipoFactura);
			this.groupBox1.Controls.Add(this.label5);
			this.groupBox1.Controls.Add(this.chkAsignacion);
			this.groupBox1.Controls.Add(this.chkDescripcion);
			this.groupBox1.Controls.Add(this.label3);
			this.groupBox1.Controls.Add(this.chkPantalla);
			this.groupBox1.Controls.Add(this.label2);
			this.groupBox1.Controls.Add(this.chkChecklist);
			this.groupBox1.Controls.Add(this.label1);
			this.groupBox1.Controls.Add(this.label21);
			this.groupBox1.Controls.Add(this.txtGrupo);
			this.groupBox1.Controls.Add(this.label12);
			this.groupBox1.Controls.Add(this.txtDescripcion);
			this.groupBox1.Controls.Add(this.label4);
			this.groupBox1.Location = new System.Drawing.Point(8, 8);
			this.groupBox1.Name = "groupBox1";
			this.groupBox1.Size = new System.Drawing.Size(360, 176);
			this.groupBox1.TabIndex = 1070;
			this.groupBox1.TabStop = false;
			this.groupBox1.Text = "Información";
			this.groupBox1.Enter += new System.EventHandler(this.groupBox1_Enter);
			// 
			// label6
			// 
			this.label6.AutoSize = true;
			this.label6.BackColor = System.Drawing.Color.Transparent;
			this.label6.Location = new System.Drawing.Point(16, 144);
			this.label6.Name = "label6";
			this.label6.Size = new System.Drawing.Size(63, 16);
			this.label6.TabIndex = 1072;
			this.label6.Text = "Todos usu.:";
			// 
			// chkTodoUsuario
			// 
			this.chkTodoUsuario.Location = new System.Drawing.Point(80, 144);
			this.chkTodoUsuario.Name = "chkTodoUsuario";
			this.chkTodoUsuario.Size = new System.Drawing.Size(16, 24);
			this.chkTodoUsuario.TabIndex = 1071;
			// 
			// cmbTipoFactura
			// 
			appearance1.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbTipoFactura.Appearance = appearance1;
			this.cmbTipoFactura.CausesValidation = false;
			this.cmbTipoFactura.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbTipoFactura.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbTipoFactura.DataSource = this.udsRequerimientoCompraNumeroObtenerTodos;
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridBand1.Columns.AddRange(new object[] {
																										 ultraGridColumn1,
																										 ultraGridColumn2});
			this.cmbTipoFactura.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.cmbTipoFactura.DisplayMember = "Nombre";
			this.cmbTipoFactura.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbTipoFactura.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbTipoFactura.Location = new System.Drawing.Point(96, 56);
			this.cmbTipoFactura.MaxDropDownItems = 30;
			this.cmbTipoFactura.Name = "cmbTipoFactura";
			this.cmbTipoFactura.Size = new System.Drawing.Size(240, 21);
			this.cmbTipoFactura.TabIndex = 1070;
			this.cmbTipoFactura.ValueMember = "idTipoFactura";
			// 
			// udsRequerimientoCompraNumeroObtenerTodos
			// 
			ultraDataColumn1.DataType = typeof(int);
			this.udsRequerimientoCompraNumeroObtenerTodos.Band.Columns.AddRange(new object[] {
																																												 ultraDataColumn1,
																																												 ultraDataColumn2});
			// 
			// label5
			// 
			this.label5.AutoSize = true;
			this.label5.BackColor = System.Drawing.Color.Transparent;
			this.label5.Location = new System.Drawing.Point(256, 120);
			this.label5.Name = "label5";
			this.label5.Size = new System.Drawing.Size(63, 16);
			this.label5.TabIndex = 1069;
			this.label5.Text = "Asignacion:";
			this.label5.Click += new System.EventHandler(this.label5_Click);
			// 
			// chkAsignacion
			// 
			this.chkAsignacion.Location = new System.Drawing.Point(320, 120);
			this.chkAsignacion.Name = "chkAsignacion";
			this.chkAsignacion.Size = new System.Drawing.Size(16, 24);
			this.chkAsignacion.TabIndex = 1068;
			// 
			// chkDescripcion
			// 
			this.chkDescripcion.Location = new System.Drawing.Point(240, 120);
			this.chkDescripcion.Name = "chkDescripcion";
			this.chkDescripcion.Size = new System.Drawing.Size(16, 24);
			this.chkDescripcion.TabIndex = 1066;
			// 
			// label3
			// 
			this.label3.AutoSize = true;
			this.label3.BackColor = System.Drawing.Color.Transparent;
			this.label3.Location = new System.Drawing.Point(96, 120);
			this.label3.Name = "label3";
			this.label3.Size = new System.Drawing.Size(48, 16);
			this.label3.TabIndex = 1065;
			this.label3.Text = "Pantalla:";
			// 
			// chkPantalla
			// 
			this.chkPantalla.Location = new System.Drawing.Point(152, 120);
			this.chkPantalla.Name = "chkPantalla";
			this.chkPantalla.Size = new System.Drawing.Size(16, 24);
			this.chkPantalla.TabIndex = 1064;
			// 
			// label2
			// 
			this.label2.AutoSize = true;
			this.label2.BackColor = System.Drawing.Color.Transparent;
			this.label2.Location = new System.Drawing.Point(16, 120);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(53, 16);
			this.label2.TabIndex = 1063;
			this.label2.Text = "Checklist:";
			// 
			// chkChecklist
			// 
			this.chkChecklist.Location = new System.Drawing.Point(80, 120);
			this.chkChecklist.Name = "chkChecklist";
			this.chkChecklist.Size = new System.Drawing.Size(16, 24);
			this.chkChecklist.TabIndex = 1062;
			// 
			// label1
			// 
			this.label1.AutoSize = true;
			this.label1.BackColor = System.Drawing.Color.Transparent;
			this.label1.Location = new System.Drawing.Point(16, 56);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(74, 16);
			this.label1.TabIndex = 1061;
			this.label1.Text = "* Tipo factura:";
			// 
			// label21
			// 
			this.label21.AutoSize = true;
			this.label21.BackColor = System.Drawing.Color.Transparent;
			this.label21.Location = new System.Drawing.Point(16, 88);
			this.label21.Name = "label21";
			this.label21.Size = new System.Drawing.Size(46, 16);
			this.label21.TabIndex = 1059;
			this.label21.Text = "* Grupo:";
			// 
			// txtGrupo
			// 
			appearance2.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtGrupo.Appearance = appearance2;
			this.txtGrupo.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtGrupo.FormatString = "";
			this.txtGrupo.Location = new System.Drawing.Point(96, 88);
			this.txtGrupo.MaskInput = "nn";
			this.txtGrupo.Name = "txtGrupo";
			this.txtGrupo.PromptChar = ' ';
			this.txtGrupo.Size = new System.Drawing.Size(240, 21);
			this.txtGrupo.SpinButtonDisplayStyle = Infragistics.Win.ButtonDisplayStyle.Always;
			this.txtGrupo.TabIndex = 1058;
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
			appearance3.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtDescripcion.Appearance = appearance3;
			this.txtDescripcion.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtDescripcion.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtDescripcion.Location = new System.Drawing.Point(96, 24);
			this.txtDescripcion.Name = "txtDescripcion";
			this.txtDescripcion.Size = new System.Drawing.Size(240, 21);
			this.txtDescripcion.TabIndex = 1052;
			// 
			// label4
			// 
			this.label4.AutoSize = true;
			this.label4.BackColor = System.Drawing.Color.Transparent;
			this.label4.Location = new System.Drawing.Point(168, 120);
			this.label4.Name = "label4";
			this.label4.Size = new System.Drawing.Size(67, 16);
			this.label4.TabIndex = 1067;
			this.label4.Text = "Descripcion:";
			// 
			// bDepartamento
			// 
			this.bDepartamento.AutoSize = true;
			this.bDepartamento.BackColor = System.Drawing.Color.Transparent;
			this.bDepartamento.Location = new System.Drawing.Point(96, 144);
			this.bDepartamento.Name = "bDepartamento";
			this.bDepartamento.Size = new System.Drawing.Size(44, 16);
			this.bDepartamento.TabIndex = 1074;
			this.bDepartamento.Text = "Depart.:";
			this.bDepartamento.Click += new System.EventHandler(this.bDepartamento_Click);
			// 
			// chkbDepartamento
			// 
			this.chkbDepartamento.Location = new System.Drawing.Point(152, 144);
			this.chkbDepartamento.Name = "chkbDepartamento";
			this.chkbDepartamento.Size = new System.Drawing.Size(16, 24);
			this.chkbDepartamento.TabIndex = 1075;
			// 
			// frmRequerimientoDepartamentoCrear
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(472, 198);
			this.Controls.Add(this.btnSalir);
			this.Controls.Add(this.btnGuardar);
			this.Controls.Add(this.groupBox1);
			this.Name = "frmRequerimientoDepartamentoCrear";
			this.Text = "Departamento";
			this.Load += new System.EventHandler(this.frmRequerimientoDepartamentoCrear_Load);
			this.groupBox1.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.cmbTipoFactura)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.udsRequerimientoCompraNumeroObtenerTodos)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtGrupo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDescripcion)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void frmRequerimientoDepartamentoCrear_Load(object sender, System.EventArgs e)
		{
			RequerimientoCompraNumeroObtenerTodos();
		}

		private void RequerimientoCompraNumeroObtenerTodos()
		{
			string squery = string.Format("Exec RequerimientoCompraNumeroObtenerTodos");
			this.cmbTipoFactura.DataSource = Funcion.dsEjecucion(frmRequerimiento.sconexionPoint, squery);
		}

		private void groupBox1_Enter(object sender, System.EventArgs e)
		{
		
		}

		private void btnGuardar_Click(object sender, System.EventArgs e)
		{
			RequerimientoDepartamentoInsertar();
		}

		private void RequerimientoDepartamentoInsertar()
		{
			#region Validación
			if (!Validacion.vbTexto(this.txtDescripcion, 5, 100, "Ingrese la descripción")) return;
			if (!Validacion.vbCampoEnteroVacio(this.txtGrupo, "Ingrese el grupo", 1, 100)) return;
			if (!Validacion.vbComboVacio(this.cmbTipoFactura, "Ingrese el tipo de factura")) return;
			#endregion Validación

			// Capturamos los valores del formulario
			string Descripcion = this.txtDescripcion.Text;
			int Grupo = (int)this.txtGrupo.Value;
			int idTipoFactura = (int)this.cmbTipoFactura.Value;
			bool Checklist = Convert.ToBoolean(this.chkChecklist.Checked);
			bool Pantalla = Convert.ToBoolean(this.chkPantalla.Checked);
			bool bDescripcion = Convert.ToBoolean(this.chkDescripcion.Checked);
			bool Asignacion = Convert.ToBoolean(this.chkAsignacion.Checked);
			bool TodoUsuario = Convert.ToBoolean(this.chkTodoUsuario.Checked);
			bool bDepartamento = Convert.ToBoolean(this.chkbDepartamento.Checked);

			try
			{
				// Establecer la conexión a la base de datos
				using (SqlConnection conn = new SqlConnection(frmRequerimiento.sconexionPoint))
				{
					conn.Open();

					// Comando para ejecutar el procedimiento almacenado
					using (SqlCommand cmd = new SqlCommand("RequerimientoDepartamentoInsertar", conn))
					{
						cmd.CommandType = CommandType.StoredProcedure;

						// Agregar parámetros al comando utilizando Add y especificando los tipos de datos
						cmd.Parameters.Add("@Descripcion", SqlDbType.VarChar, 500).Value = Descripcion;
						cmd.Parameters.Add("@Grupo", SqlDbType.Int).Value = Grupo;
						cmd.Parameters.Add("@idTipoFactura", SqlDbType.Int).Value = idTipoFactura;
						cmd.Parameters.Add("@Checklist", SqlDbType.Bit).Value = Checklist;
						cmd.Parameters.Add("@Pantalla", SqlDbType.Bit).Value = Pantalla;
						cmd.Parameters.Add("@bDescripcion", SqlDbType.Bit).Value = bDescripcion;
						cmd.Parameters.Add("@Asignacion", SqlDbType.Bit).Value = Asignacion;
						cmd.Parameters.Add("@TodoUsuario", SqlDbType.Bit).Value = Asignacion;
						cmd.Parameters.Add("@bDepartamento", SqlDbType.Bit).Value = bDepartamento;

						// Ejecutar el procedimiento almacenado
						cmd.ExecuteNonQuery();
					}

					// Mostrar mensaje de éxito
					MessageBox.Show("Registro insertado correctamente.", "Éxito", MessageBoxButtons.OK, MessageBoxIcon.Information);
					this.DialogResult = DialogResult.OK;
				}
			}
			catch (Exception ex)
			{
				// Manejo de errores
				MessageBox.Show("Error al insertar el registro: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
			}
		}

		private void button1_Click(object sender, System.EventArgs e)
		{
			this.Close();
		}

		private void label5_Click(object sender, System.EventArgs e)
		{
		
		}

		private void bDepartamento_Click(object sender, System.EventArgs e)
		{
		
		}
	}
}
