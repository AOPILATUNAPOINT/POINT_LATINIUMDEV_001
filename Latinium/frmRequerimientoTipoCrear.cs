using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data; // Para DataTable, DataSet, etc.
using System.Data.SqlClient;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de frmRequerimientoTipoCrear.
	/// </summary>
	public class frmRequerimientoTipoCrear : System.Windows.Forms.Form
	{
		private System.Windows.Forms.Button btnNuevo;
		private System.Windows.Forms.Button button1;
		private System.Windows.Forms.Button button3;
		private System.Windows.Forms.GroupBox groupBox1;
		private System.Windows.Forms.Label lblAsignado;
		public System.Windows.Forms.Label label12;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbDepartamento;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtDescripcion;
		private System.Windows.Forms.Label label4;
		private System.Windows.Forms.Label lblSeleccionDefault;
		private System.Windows.Forms.CheckBox chkSeleccionDefault;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource udsRequerimientoDepartamentoCon;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public frmRequerimientoTipoCrear()
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
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmRequerimientoTipoCrear));
			Infragistics.Win.Appearance appearance1 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idRequerimientoDepartamento");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Descripcion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idRequerimientoDepartamento");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Descripcion");
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			this.btnNuevo = new System.Windows.Forms.Button();
			this.button1 = new System.Windows.Forms.Button();
			this.button3 = new System.Windows.Forms.Button();
			this.groupBox1 = new System.Windows.Forms.GroupBox();
			this.lblSeleccionDefault = new System.Windows.Forms.Label();
			this.chkSeleccionDefault = new System.Windows.Forms.CheckBox();
			this.lblAsignado = new System.Windows.Forms.Label();
			this.label12 = new System.Windows.Forms.Label();
			this.cmbDepartamento = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.udsRequerimientoDepartamentoCon = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.txtDescripcion = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label4 = new System.Windows.Forms.Label();
			this.groupBox1.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.cmbDepartamento)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.udsRequerimientoDepartamentoCon)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDescripcion)).BeginInit();
			this.SuspendLayout();
			// 
			// btnNuevo
			// 
			this.btnNuevo.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btnNuevo.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(230)), ((System.Byte)(230)), ((System.Byte)(230)));
			this.btnNuevo.CausesValidation = false;
			this.btnNuevo.Image = ((System.Drawing.Image)(resources.GetObject("btnNuevo.Image")));
			this.btnNuevo.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnNuevo.Location = new System.Drawing.Point(384, 48);
			this.btnNuevo.Name = "btnNuevo";
			this.btnNuevo.Size = new System.Drawing.Size(80, 24);
			this.btnNuevo.TabIndex = 1073;
			this.btnNuevo.Text = "&Nuevo dep.";
			this.btnNuevo.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnNuevo.Click += new System.EventHandler(this.btnNuevo_Click);
			// 
			// button1
			// 
			this.button1.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.button1.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(230)), ((System.Byte)(230)), ((System.Byte)(230)));
			this.button1.CausesValidation = false;
			this.button1.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.button1.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(50)), ((System.Byte)(50)), ((System.Byte)(50)));
			this.button1.Image = ((System.Drawing.Image)(resources.GetObject("button1.Image")));
			this.button1.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.button1.Location = new System.Drawing.Point(384, 80);
			this.button1.Name = "button1";
			this.button1.Size = new System.Drawing.Size(80, 23);
			this.button1.TabIndex = 1072;
			this.button1.Text = "Salir";
			this.button1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.button1.Click += new System.EventHandler(this.button1_Click);
			// 
			// button3
			// 
			this.button3.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.button3.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(230)), ((System.Byte)(230)), ((System.Byte)(230)));
			this.button3.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.button3.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(50)), ((System.Byte)(50)), ((System.Byte)(50)));
			this.button3.Image = ((System.Drawing.Image)(resources.GetObject("button3.Image")));
			this.button3.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.button3.Location = new System.Drawing.Point(384, 16);
			this.button3.Name = "button3";
			this.button3.Size = new System.Drawing.Size(80, 23);
			this.button3.TabIndex = 1071;
			this.button3.Text = "Guardar";
			this.button3.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.button3.Click += new System.EventHandler(this.button3_Click);
			// 
			// groupBox1
			// 
			this.groupBox1.Controls.Add(this.lblSeleccionDefault);
			this.groupBox1.Controls.Add(this.chkSeleccionDefault);
			this.groupBox1.Controls.Add(this.lblAsignado);
			this.groupBox1.Controls.Add(this.label12);
			this.groupBox1.Controls.Add(this.cmbDepartamento);
			this.groupBox1.Controls.Add(this.txtDescripcion);
			this.groupBox1.Controls.Add(this.label4);
			this.groupBox1.Location = new System.Drawing.Point(16, 16);
			this.groupBox1.Name = "groupBox1";
			this.groupBox1.Size = new System.Drawing.Size(352, 136);
			this.groupBox1.TabIndex = 1070;
			this.groupBox1.TabStop = false;
			this.groupBox1.Text = "Información";
			// 
			// lblSeleccionDefault
			// 
			this.lblSeleccionDefault.AutoSize = true;
			this.lblSeleccionDefault.BackColor = System.Drawing.Color.Transparent;
			this.lblSeleccionDefault.Location = new System.Drawing.Point(24, 104);
			this.lblSeleccionDefault.Name = "lblSeleccionDefault";
			this.lblSeleccionDefault.Size = new System.Drawing.Size(53, 16);
			this.lblSeleccionDefault.TabIndex = 1065;
			this.lblSeleccionDefault.Text = "Checklist:";
			// 
			// chkSeleccionDefault
			// 
			this.chkSeleccionDefault.Location = new System.Drawing.Point(128, 104);
			this.chkSeleccionDefault.Name = "chkSeleccionDefault";
			this.chkSeleccionDefault.Size = new System.Drawing.Size(16, 24);
			this.chkSeleccionDefault.TabIndex = 1064;
			// 
			// lblAsignado
			// 
			this.lblAsignado.Location = new System.Drawing.Point(16, 24);
			this.lblAsignado.Name = "lblAsignado";
			this.lblAsignado.Size = new System.Drawing.Size(96, 23);
			this.lblAsignado.TabIndex = 1042;
			this.lblAsignado.Text = "* Departamento:";
			// 
			// label12
			// 
			this.label12.AutoSize = true;
			this.label12.Location = new System.Drawing.Point(16, 64);
			this.label12.Name = "label12";
			this.label12.Size = new System.Drawing.Size(74, 16);
			this.label12.TabIndex = 1053;
			this.label12.Text = "* Descripción:";
			this.label12.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// cmbDepartamento
			// 
			appearance1.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbDepartamento.Appearance = appearance1;
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
			this.cmbDepartamento.TabIndex = 1041;
			this.cmbDepartamento.ValueMember = "idRequerimientoDepartamento";
			// 
			// udsRequerimientoDepartamentoCon
			// 
			ultraDataColumn1.DataType = typeof(int);
			this.udsRequerimientoDepartamentoCon.Band.Columns.AddRange(new object[] {
																																								ultraDataColumn1,
																																								ultraDataColumn2});
			// 
			// txtDescripcion
			// 
			appearance2.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtDescripcion.Appearance = appearance2;
			this.txtDescripcion.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtDescripcion.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtDescripcion.Location = new System.Drawing.Point(128, 64);
			this.txtDescripcion.Name = "txtDescripcion";
			this.txtDescripcion.Size = new System.Drawing.Size(200, 21);
			this.txtDescripcion.TabIndex = 1052;
			// 
			// label4
			// 
			this.label4.Location = new System.Drawing.Point(16, 24);
			this.label4.Name = "label4";
			this.label4.Size = new System.Drawing.Size(96, 23);
			this.label4.TabIndex = 1042;
			this.label4.Text = "* Departamento:";
			// 
			// frmRequerimientoTipoCrear
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(472, 166);
			this.Controls.Add(this.btnNuevo);
			this.Controls.Add(this.button1);
			this.Controls.Add(this.button3);
			this.Controls.Add(this.groupBox1);
			this.Name = "frmRequerimientoTipoCrear";
			this.Text = "Requerimiento tipo";
			this.Load += new System.EventHandler(this.frmRequerimientoTipoCrear_Load);
			this.groupBox1.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.cmbDepartamento)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.udsRequerimientoDepartamentoCon)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDescripcion)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void btnNuevo_Click(object sender, System.EventArgs e)
		{
			using (frmRequerimientoDepartamentoCrear frmdc = new frmRequerimientoDepartamentoCrear())
			{		
				if (DialogResult.OK == frmdc.ShowDialog())
				{
				}
				//Cargar departamento
				frmRequerimientoConfiguracion.RequerimientoDepartamentoCon(this.cmbDepartamento);
			}
		}

		private void button1_Click(object sender, System.EventArgs e)
		{
			this.Close();
		}

		private void button3_Click(object sender, System.EventArgs e)
		{
			RequerimientoTipoInsertar();
		}

		private void RequerimientoTipoInsertar()
		{
			#region Validación
			if (!Validacion.vbComboVacio(this.cmbDepartamento, "Seleccione el departamento")) return;
			if (!Validacion.vbTexto(this.txtDescripcion, 5, 100, "Ingrese la descripción")) return;
			#endregion Validación

			// Asignar valores a las variables
			string descripcion = txtDescripcion.Text.Trim();
			bool estado = true; // true equivale a 1 en SQL (activo)
			int idRequerimientoDepartamento = Convert.ToInt32(cmbDepartamento.Value); // Asumiendo que el ComboBox contiene el ID
			bool seleccionDefault = chkSeleccionDefault.Checked; // Asumiendo que hay un CheckBox para la selección predeterminada

			// Conexión a la base de datos y ejecución del procedimiento almacenado
			using (SqlConnection conn = new SqlConnection(frmRequerimiento.sconexionPoint))
			{
				try
				{
					conn.Open();
					using (SqlCommand cmd = new SqlCommand("RequerimientoTipoInsertar", conn))
					{
						cmd.CommandType = CommandType.StoredProcedure;

						// Agregar parámetros utilizando SqlParameter
						cmd.Parameters.Add(new SqlParameter("@Descripcion", SqlDbType.VarChar, 500)).Value = descripcion;
						cmd.Parameters.Add(new SqlParameter("@Estado", SqlDbType.Bit)).Value = estado;
						cmd.Parameters.Add(new SqlParameter("@idRequerimientoDepartamento", SqlDbType.Int)).Value = idRequerimientoDepartamento;
						cmd.Parameters.Add(new SqlParameter("@SeleccionDefault", SqlDbType.Bit)).Value = seleccionDefault;

						// Ejecutar el comando
						cmd.ExecuteNonQuery();

						MessageBox.Show("Registro insertado exitosamente", "Éxito", MessageBoxButtons.OK, MessageBoxIcon.Information);
						this.DialogResult = DialogResult.OK;
					}
				}
				catch (Exception ex)
				{
					MessageBox.Show("Error al insertar el registro: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
				}
			}
		}

		private void frmRequerimientoTipoCrear_Load(object sender, System.EventArgs e)
		{
			//Cargar departamento
			frmRequerimientoConfiguracion.RequerimientoDepartamentoCon(this.cmbDepartamento);
		}

	}
}
