using System;
using System.Data.SqlClient; // Necesario para SQL Server
using System.Windows.Forms;

namespace Latinium
{
	public class frmRequerimientoEstadoDepartamento : System.Windows.Forms.Form
	{
		private Button btnGuardar;
		private System.Windows.Forms.GroupBox groupBox1;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbEstado;
		private System.Windows.Forms.Label lblEstado;
		private System.Windows.Forms.Label lblAsignado;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbDepartamento;
		private System.Windows.Forms.Label label4;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource udsEstado;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource udsRequerimientoDepartamentoCon;
		private System.Windows.Forms.Label label1;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtPosicion;
		private System.Windows.Forms.Button btnNvEstado;
		private System.Windows.Forms.Button btnNuevo;
		private Button btnSalir;

		public frmRequerimientoEstadoDepartamento()
		{
			InitializeComponent();
		}

		protected override void Dispose(bool disposing)
		{
			// No necesitas la variable 'components', por lo tanto, puedes quitarla.
			base.Dispose(disposing);
		}

		private void InitializeComponent()
		{
			Infragistics.Win.Appearance appearance1 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idEstadoRequerimiento");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Descripcion");
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idRequerimientoDepartamento");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Descripcion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idEstadoRequerimiento");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Descripcion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn3 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idRequerimientoDepartamento");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn4 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Descripcion");
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmRequerimientoEstadoDepartamento));
			this.btnGuardar = new System.Windows.Forms.Button();
			this.btnSalir = new System.Windows.Forms.Button();
			this.groupBox1 = new System.Windows.Forms.GroupBox();
			this.txtPosicion = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label1 = new System.Windows.Forms.Label();
			this.cmbEstado = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.lblEstado = new System.Windows.Forms.Label();
			this.lblAsignado = new System.Windows.Forms.Label();
			this.cmbDepartamento = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.label4 = new System.Windows.Forms.Label();
			this.udsEstado = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.udsRequerimientoDepartamentoCon = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.btnNvEstado = new System.Windows.Forms.Button();
			this.btnNuevo = new System.Windows.Forms.Button();
			this.groupBox1.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.txtPosicion)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbEstado)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbDepartamento)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.udsEstado)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.udsRequerimientoDepartamentoCon)).BeginInit();
			this.SuspendLayout();
			// 
			// btnGuardar
			// 
			this.btnGuardar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btnGuardar.Location = new System.Drawing.Point(376, 16);
			this.btnGuardar.Name = "btnGuardar";
			this.btnGuardar.Size = new System.Drawing.Size(80, 24);
			this.btnGuardar.TabIndex = 3;
			this.btnGuardar.Text = "Guardar";
			this.btnGuardar.Click += new System.EventHandler(this.btnGuardar_Click);
			// 
			// btnSalir
			// 
			this.btnSalir.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btnSalir.Location = new System.Drawing.Point(376, 112);
			this.btnSalir.Name = "btnSalir";
			this.btnSalir.Size = new System.Drawing.Size(80, 24);
			this.btnSalir.TabIndex = 4;
			this.btnSalir.Text = "Salir";
			this.btnSalir.Click += new System.EventHandler(this.btnSalir_Click);
			// 
			// groupBox1
			// 
			this.groupBox1.Controls.Add(this.txtPosicion);
			this.groupBox1.Controls.Add(this.label1);
			this.groupBox1.Controls.Add(this.cmbEstado);
			this.groupBox1.Controls.Add(this.lblEstado);
			this.groupBox1.Controls.Add(this.lblAsignado);
			this.groupBox1.Controls.Add(this.cmbDepartamento);
			this.groupBox1.Controls.Add(this.label4);
			this.groupBox1.Location = new System.Drawing.Point(16, 8);
			this.groupBox1.Name = "groupBox1";
			this.groupBox1.Size = new System.Drawing.Size(344, 136);
			this.groupBox1.TabIndex = 5;
			this.groupBox1.TabStop = false;
			this.groupBox1.Text = "Información";
			// 
			// txtPosicion
			// 
			appearance1.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtPosicion.Appearance = appearance1;
			this.txtPosicion.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtPosicion.FormatString = "";
			this.txtPosicion.Location = new System.Drawing.Point(120, 88);
			this.txtPosicion.MaskInput = "nn";
			this.txtPosicion.Name = "txtPosicion";
			this.txtPosicion.PromptChar = ' ';
			this.txtPosicion.Size = new System.Drawing.Size(200, 21);
			this.txtPosicion.SpinButtonDisplayStyle = Infragistics.Win.ButtonDisplayStyle.Always;
			this.txtPosicion.TabIndex = 1066;
			// 
			// label1
			// 
			this.label1.Location = new System.Drawing.Point(8, 88);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(96, 23);
			this.label1.TabIndex = 1065;
			this.label1.Text = "* Posición:";
			// 
			// cmbEstado
			// 
			appearance2.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbEstado.Appearance = appearance2;
			this.cmbEstado.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.cmbEstado.Cursor = System.Windows.Forms.Cursors.Default;
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridBand1.Columns.AddRange(new object[] {
																										 ultraGridColumn1,
																										 ultraGridColumn2});
			this.cmbEstado.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.cmbEstado.DisplayMember = "Descripcion";
			this.cmbEstado.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbEstado.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbEstado.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbEstado.Location = new System.Drawing.Point(120, 56);
			this.cmbEstado.Name = "cmbEstado";
			this.cmbEstado.Size = new System.Drawing.Size(200, 21);
			this.cmbEstado.TabIndex = 1064;
			this.cmbEstado.ValueMember = "idEstadoRequerimiento";
			// 
			// lblEstado
			// 
			this.lblEstado.Location = new System.Drawing.Point(8, 56);
			this.lblEstado.Name = "lblEstado";
			this.lblEstado.Size = new System.Drawing.Size(96, 23);
			this.lblEstado.TabIndex = 1063;
			this.lblEstado.Text = "* Estado:";
			// 
			// lblAsignado
			// 
			this.lblAsignado.Location = new System.Drawing.Point(8, 24);
			this.lblAsignado.Name = "lblAsignado";
			this.lblAsignado.Size = new System.Drawing.Size(96, 23);
			this.lblAsignado.TabIndex = 1062;
			this.lblAsignado.Text = "* Departamento:";
			// 
			// cmbDepartamento
			// 
			appearance3.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbDepartamento.Appearance = appearance3;
			this.cmbDepartamento.CausesValidation = false;
			this.cmbDepartamento.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbDepartamento.Cursor = System.Windows.Forms.Cursors.Default;
			ultraGridColumn3.Header.VisiblePosition = 0;
			ultraGridColumn3.Hidden = true;
			ultraGridColumn4.Header.Caption = "Descripción";
			ultraGridColumn4.Header.VisiblePosition = 1;
			ultraGridColumn4.Width = 218;
			ultraGridBand2.Columns.AddRange(new object[] {
																										 ultraGridColumn3,
																										 ultraGridColumn4});
			this.cmbDepartamento.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			this.cmbDepartamento.DisplayMember = "Descripcion";
			this.cmbDepartamento.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbDepartamento.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbDepartamento.Location = new System.Drawing.Point(120, 24);
			this.cmbDepartamento.MaxDropDownItems = 30;
			this.cmbDepartamento.Name = "cmbDepartamento";
			this.cmbDepartamento.Size = new System.Drawing.Size(200, 21);
			this.cmbDepartamento.TabIndex = 1060;
			this.cmbDepartamento.ValueMember = "idRequerimientoDepartamento";
			// 
			// label4
			// 
			this.label4.Location = new System.Drawing.Point(8, 24);
			this.label4.Name = "label4";
			this.label4.Size = new System.Drawing.Size(96, 23);
			this.label4.TabIndex = 1061;
			this.label4.Text = "* Departamento:";
			// 
			// udsEstado
			// 
			ultraDataColumn1.DataType = typeof(int);
			this.udsEstado.Band.Columns.AddRange(new object[] {
																													ultraDataColumn1,
																													ultraDataColumn2});
			// 
			// udsRequerimientoDepartamentoCon
			// 
			ultraDataColumn3.DataType = typeof(int);
			this.udsRequerimientoDepartamentoCon.Band.Columns.AddRange(new object[] {
																																								ultraDataColumn3,
																																								ultraDataColumn4});
			// 
			// btnNvEstado
			// 
			this.btnNvEstado.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btnNvEstado.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(230)), ((System.Byte)(230)), ((System.Byte)(230)));
			this.btnNvEstado.CausesValidation = false;
			this.btnNvEstado.Image = ((System.Drawing.Image)(resources.GetObject("btnNvEstado.Image")));
			this.btnNvEstado.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnNvEstado.Location = new System.Drawing.Point(376, 80);
			this.btnNvEstado.Name = "btnNvEstado";
			this.btnNvEstado.Size = new System.Drawing.Size(80, 24);
			this.btnNvEstado.TabIndex = 1078;
			this.btnNvEstado.Text = "&Nuevo est.";
			this.btnNvEstado.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnNvEstado.Click += new System.EventHandler(this.btnNvEstado_Click);
			// 
			// btnNuevo
			// 
			this.btnNuevo.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btnNuevo.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(230)), ((System.Byte)(230)), ((System.Byte)(230)));
			this.btnNuevo.CausesValidation = false;
			this.btnNuevo.Image = ((System.Drawing.Image)(resources.GetObject("btnNuevo.Image")));
			this.btnNuevo.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnNuevo.Location = new System.Drawing.Point(376, 48);
			this.btnNuevo.Name = "btnNuevo";
			this.btnNuevo.Size = new System.Drawing.Size(80, 24);
			this.btnNuevo.TabIndex = 1077;
			this.btnNuevo.Text = "&Nuevo dep.";
			this.btnNuevo.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnNuevo.Click += new System.EventHandler(this.btnNuevo_Click);
			// 
			// frmRequerimientoEstadoDepartamento
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(464, 158);
			this.Controls.Add(this.btnNvEstado);
			this.Controls.Add(this.btnNuevo);
			this.Controls.Add(this.groupBox1);
			this.Controls.Add(this.btnGuardar);
			this.Controls.Add(this.btnSalir);
			this.Name = "frmRequerimientoEstadoDepartamento";
			this.Text = "Estado Requerimiento Departamento";
			this.Load += new System.EventHandler(this.frmRequerimientoEstadoDepartamento_Load);
			this.groupBox1.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.txtPosicion)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbEstado)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbDepartamento)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.udsEstado)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.udsRequerimientoDepartamentoCon)).EndInit();
			this.ResumeLayout(false);

		}

		private void frmRequerimientoEstadoDepartamento_Load(object sender, EventArgs e)
		{
			// Aquí puedes cargar datos en los combos desde la base de datos.
			CargarCombos();
		}

		private void CargarCombos()
		{
			frmRequerimientoConfiguracion.RequerimientoEstadoObtenerActivos(this.cmbEstado);
			frmRequerimientoConfiguracion.RequerimientoDepartamentoCon(this.cmbDepartamento);
		}

		private void btnGuardar_Click(object sender, EventArgs e)
		{
			#region Validación
			if (!Validacion.vbComboVacio(this.cmbDepartamento, "Seleccione el departamento ")) return;
			if (!Validacion.vbComboVacio(this.cmbEstado, "Seleccione el estado ")) return;
			if (!Validacion.vbCampoEnteroVacio(this.txtPosicion, "Ingrese la posición", 1, 1000)) return;
			#endregion Validación

			// Llamar al método para guardar en la base de datos
			GuardarEstadoRequerimientoDepartamento();
		}

		private void GuardarEstadoRequerimientoDepartamento()
		{
			// Obtener valores de los controles
			int idEstadoRequerimiento = (int)cmbEstado.Value;
			int idRequerimientoDepartamento = (int)cmbDepartamento.Value;
			int posicion = int.Parse(txtPosicion.Text);

			// Cadena de conexión (ajusta según tu configuración)
			string connectionString = frmRequerimiento.sconexionPoint;

			using (SqlConnection conn = new SqlConnection(connectionString))
			{
				try
				{
					conn.Open();
					using (SqlCommand cmd = new SqlCommand("EstadoRequerimientoDepartamentoInsertar", conn))
					{
						cmd.CommandType = System.Data.CommandType.StoredProcedure;
						// Definir los parámetros usando Add y especificar SqlDbType
						cmd.Parameters.Add(new SqlParameter("@idEstadoRequerimiento", System.Data.SqlDbType.Int)).Value = idEstadoRequerimiento;
						cmd.Parameters.Add(new SqlParameter("@idRequerimientoDepartamento", System.Data.SqlDbType.Int)).Value = idRequerimientoDepartamento;
						cmd.Parameters.Add(new SqlParameter("@Posicion", System.Data.SqlDbType.Int)).Value = posicion;
						// Ejecutar el procedimiento almacenado
						int newId = Convert.ToInt32(cmd.ExecuteScalar());
						MessageBox.Show(string.Format("Registro insertado con éxito. Nuevo ID: {0}", newId));
						this.DialogResult = DialogResult.OK;
					}
				}
				catch (Exception ex)
				{
					MessageBox.Show(string.Format("Error al guardar: {0}", ex.Message));
				}
			}
		}

		private void btnSalir_Click(object sender, EventArgs e)
		{
			// Cerrar el formulario
			this.Close();
		}

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

		private void btnNvEstado_Click(object sender, System.EventArgs e)
		{
			using (frmRequerimientoEstadoCrear frmec = new frmRequerimientoEstadoCrear())
			{		
				if (DialogResult.OK == frmec.ShowDialog())
				{
				}
				//Cargar Estados Activos
				frmRequerimientoConfiguracion.RequerimientoEstadoObtenerActivos(this.cmbEstado);
			}	
		}
	}
}
