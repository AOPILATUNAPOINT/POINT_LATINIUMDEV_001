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
	/// Descripción breve de frmRequerimientoRolPersmisoCrear.
	/// </summary>
	public class frmRequerimientoRolPersmisoCrear : System.Windows.Forms.Form
	{
		private System.Windows.Forms.Button button1;
		private System.Windows.Forms.Button button3;
		private System.Windows.Forms.GroupBox groupBox1;
		private System.Windows.Forms.Label lblAsignado;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbPermiso;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbRol;
		private System.Windows.Forms.Label label3;
		private System.Windows.Forms.Label label4;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource udsRequerimientoRolCon;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource udsRequerimientoPermisosObtenerTodos;
		private System.Windows.Forms.Button btnNvRol;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public frmRequerimientoRolPersmisoCrear()
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
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmRequerimientoRolPersmisoCrear));
			Infragistics.Win.Appearance appearance1 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idRequerimientoPermiso");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Descripcion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Estado");
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idRequerimientoRol");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Descripcion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idRequerimientoRol");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Descripcion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn3 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idRequerimientoPermiso");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn4 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Descripcion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn5 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Estado");
			this.button1 = new System.Windows.Forms.Button();
			this.button3 = new System.Windows.Forms.Button();
			this.groupBox1 = new System.Windows.Forms.GroupBox();
			this.lblAsignado = new System.Windows.Forms.Label();
			this.cmbPermiso = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.cmbRol = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.label3 = new System.Windows.Forms.Label();
			this.label4 = new System.Windows.Forms.Label();
			this.udsRequerimientoRolCon = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.udsRequerimientoPermisosObtenerTodos = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.btnNvRol = new System.Windows.Forms.Button();
			this.groupBox1.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.cmbPermiso)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbRol)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.udsRequerimientoRolCon)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.udsRequerimientoPermisosObtenerTodos)).BeginInit();
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
			this.button1.Location = new System.Drawing.Point(376, 80);
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
			this.button3.Location = new System.Drawing.Point(376, 16);
			this.button3.Name = "button3";
			this.button3.Size = new System.Drawing.Size(80, 23);
			this.button3.TabIndex = 1067;
			this.button3.Text = "Guardar";
			this.button3.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.button3.Click += new System.EventHandler(this.button3_Click);
			// 
			// groupBox1
			// 
			this.groupBox1.Controls.Add(this.lblAsignado);
			this.groupBox1.Controls.Add(this.cmbPermiso);
			this.groupBox1.Controls.Add(this.cmbRol);
			this.groupBox1.Controls.Add(this.label3);
			this.groupBox1.Controls.Add(this.label4);
			this.groupBox1.Location = new System.Drawing.Point(16, 16);
			this.groupBox1.Name = "groupBox1";
			this.groupBox1.Size = new System.Drawing.Size(352, 104);
			this.groupBox1.TabIndex = 1066;
			this.groupBox1.TabStop = false;
			this.groupBox1.Text = "Información";
			// 
			// lblAsignado
			// 
			this.lblAsignado.Location = new System.Drawing.Point(16, 24);
			this.lblAsignado.Name = "lblAsignado";
			this.lblAsignado.Size = new System.Drawing.Size(96, 23);
			this.lblAsignado.TabIndex = 1042;
			this.lblAsignado.Text = "* Permiso:";
			// 
			// cmbPermiso
			// 
			appearance1.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbPermiso.Appearance = appearance1;
			this.cmbPermiso.CausesValidation = false;
			this.cmbPermiso.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbPermiso.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbPermiso.DataSource = this.udsRequerimientoPermisosObtenerTodos;
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn2.Header.Caption = "Descripción";
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridColumn2.Width = 218;
			ultraGridColumn3.Header.VisiblePosition = 2;
			ultraGridBand1.Columns.AddRange(new object[] {
																										 ultraGridColumn1,
																										 ultraGridColumn2,
																										 ultraGridColumn3});
			this.cmbPermiso.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.cmbPermiso.DisplayMember = "Descripcion";
			this.cmbPermiso.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbPermiso.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbPermiso.Location = new System.Drawing.Point(128, 24);
			this.cmbPermiso.MaxDropDownItems = 30;
			this.cmbPermiso.Name = "cmbPermiso";
			this.cmbPermiso.Size = new System.Drawing.Size(200, 21);
			this.cmbPermiso.TabIndex = 1041;
			this.cmbPermiso.ValueMember = "idRequerimientoPermiso";
			// 
			// cmbRol
			// 
			appearance2.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbRol.Appearance = appearance2;
			this.cmbRol.CausesValidation = false;
			this.cmbRol.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbRol.Cursor = System.Windows.Forms.Cursors.Default;
			ultraGridColumn4.Header.VisiblePosition = 0;
			ultraGridColumn4.Hidden = true;
			ultraGridColumn5.Header.Caption = "Descripción";
			ultraGridColumn5.Header.VisiblePosition = 1;
			ultraGridColumn5.Width = 218;
			ultraGridBand2.Columns.AddRange(new object[] {
																										 ultraGridColumn4,
																										 ultraGridColumn5});
			this.cmbRol.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			this.cmbRol.DisplayMember = "Descripcion";
			this.cmbRol.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbRol.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbRol.Location = new System.Drawing.Point(128, 64);
			this.cmbRol.MaxDropDownItems = 30;
			this.cmbRol.Name = "cmbRol";
			this.cmbRol.Size = new System.Drawing.Size(200, 21);
			this.cmbRol.TabIndex = 1056;
			this.cmbRol.ValueMember = "idRequerimientoRol";
			// 
			// label3
			// 
			this.label3.Location = new System.Drawing.Point(16, 64);
			this.label3.Name = "label3";
			this.label3.Size = new System.Drawing.Size(96, 23);
			this.label3.TabIndex = 1057;
			this.label3.Text = "* Rol:";
			// 
			// label4
			// 
			this.label4.Location = new System.Drawing.Point(16, 24);
			this.label4.Name = "label4";
			this.label4.Size = new System.Drawing.Size(96, 23);
			this.label4.TabIndex = 1042;
			this.label4.Text = "* Departamento:";
			// 
			// udsRequerimientoRolCon
			// 
			ultraDataColumn1.DataType = typeof(int);
			this.udsRequerimientoRolCon.Band.Columns.AddRange(new object[] {
																																			 ultraDataColumn1,
																																			 ultraDataColumn2});
			// 
			// udsRequerimientoPermisosObtenerTodos
			// 
			ultraDataColumn3.DataType = typeof(int);
			ultraDataColumn5.DataType = typeof(bool);
			this.udsRequerimientoPermisosObtenerTodos.Band.Columns.AddRange(new object[] {
																																										 ultraDataColumn3,
																																										 ultraDataColumn4,
																																										 ultraDataColumn5});
			// 
			// btnNvRol
			// 
			this.btnNvRol.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btnNvRol.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(230)), ((System.Byte)(230)), ((System.Byte)(230)));
			this.btnNvRol.CausesValidation = false;
			this.btnNvRol.Image = ((System.Drawing.Image)(resources.GetObject("btnNvRol.Image")));
			this.btnNvRol.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnNvRol.Location = new System.Drawing.Point(376, 48);
			this.btnNvRol.Name = "btnNvRol";
			this.btnNvRol.Size = new System.Drawing.Size(80, 24);
			this.btnNvRol.TabIndex = 1078;
			this.btnNvRol.Text = "&Nuevo rol";
			this.btnNvRol.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnNvRol.Click += new System.EventHandler(this.btnNvRol_Click);
			// 
			// frmRequerimientoRolPersmisoCrear
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(464, 134);
			this.Controls.Add(this.btnNvRol);
			this.Controls.Add(this.button1);
			this.Controls.Add(this.button3);
			this.Controls.Add(this.groupBox1);
			this.Name = "frmRequerimientoRolPersmisoCrear";
			this.Text = "Permiso - Rol";
			this.Load += new System.EventHandler(this.frmRequerimientoRolPersmisoCrear_Load);
			this.groupBox1.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.cmbPermiso)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbRol)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.udsRequerimientoRolCon)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.udsRequerimientoPermisosObtenerTodos)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void frmRequerimientoRolPersmisoCrear_Load(object sender, System.EventArgs e)
		{
			frmRequerimientoConfiguracion.RequerimientoRolCon(this.cmbRol);
			frmRequerimientoConfiguracion.RequerimientoPermisosObtenerTodos(this.cmbPermiso);
		}

		private void button3_Click(object sender, System.EventArgs e)
		{
			#region Validación
			if (!Validacion.vbComboVacio(this.cmbPermiso, "Seleccione el permiso")) return;
			if (!Validacion.vbComboVacio(this.cmbRol, "Seleccione el rol")) return;
			#endregion Validación

			int idRequerimientoRol = 0;
			int idRequerimientoPermiso = 0;

			// Obtener los valores seleccionados en los combos
			idRequerimientoPermiso = (int)this.cmbPermiso.Value;
			idRequerimientoRol = (int)this.cmbRol.Value;

			// Insertar los datos en la base de datos
			try
			{
				using (SqlConnection conn = new SqlConnection(frmRequerimiento.sconexionPoint))
				{
					conn.Open();
					using (SqlCommand cmd = new SqlCommand("RequerimientoRolPermisoInsertar", conn))
					{
						cmd.CommandType = CommandType.StoredProcedure;
                
						// Agregar parámetros al comando
						cmd.Parameters.Add("@idRequerimientoRol", SqlDbType.Int).Value = idRequerimientoRol;
						cmd.Parameters.Add("@idRequerimientoPermiso", SqlDbType.Int).Value = idRequerimientoPermiso;

						// Ejecutar el comando y obtener el ID del registro recién insertado
						int newId = Convert.ToInt32(cmd.ExecuteScalar());

						// Mostrar mensaje de éxito
						MessageBox.Show("El registro se insertó correctamente con ID: " + newId, "Éxito", MessageBoxButtons.OK, MessageBoxIcon.Information);
						this.DialogResult = DialogResult.OK;
					}
				}
			}
			catch (Exception ex)
			{
				// Manejar excepciones adecuadamente y mostrar mensaje de error
				MessageBox.Show("Error al insertar el registro: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
			}
		}

		private void btnNvRol_Click(object sender, System.EventArgs e)
		{
			using (frmRequerimientoRolCrear frmrrc = new frmRequerimientoRolCrear())
			{		
				if (DialogResult.OK == frmrrc.ShowDialog())
				{
				}
				//Cargar Rol
				frmRequerimientoConfiguracion.RequerimientoRolCon(this.cmbRol);
			}	
		}

	}
}
