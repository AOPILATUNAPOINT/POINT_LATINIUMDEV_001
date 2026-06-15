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
	/// Descripción breve de frmRequerimientoPoliticaEstadoCrear.
	/// </summary>
	public class frmRequerimientoPoliticaEstadoCrear : System.Windows.Forms.Form
	{
		private System.Windows.Forms.Button button1;
		private System.Windows.Forms.Button button3;
		private System.Windows.Forms.GroupBox groupBox1;
		private System.Windows.Forms.Label label3;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbEstado;
		private System.Windows.Forms.Label lblEstado;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbEstadoAnterior;
		private System.Windows.Forms.Label label1;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbPolitica;
		private System.Windows.Forms.Button btnNvEstado;
		private System.Windows.Forms.Button button2;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource udsEstado;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource udsRequerimientoPoliticaObtenerTodos;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public frmRequerimientoPoliticaEstadoCrear()
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
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmRequerimientoPoliticaEstadoCrear));
			Infragistics.Win.Appearance appearance1 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idEstadoRequerimiento");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Descripcion");
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idEstadoRequerimiento");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Descripcion");
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand3 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idRequerimientoPolitica");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Descripcion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("NumeroPolitica");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Mensaje");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Resultado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Estado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn11 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idRequerimientoDepartamento");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idEstadoRequerimiento");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Descripcion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn3 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idRequerimientoPolitica");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn4 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Descripcion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn5 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("NumeroPolitica");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn6 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Mensaje");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn7 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Resultado");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn8 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Estado");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn9 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idRequerimientoDepartamento");
			this.button1 = new System.Windows.Forms.Button();
			this.button3 = new System.Windows.Forms.Button();
			this.groupBox1 = new System.Windows.Forms.GroupBox();
			this.cmbEstadoAnterior = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.label1 = new System.Windows.Forms.Label();
			this.cmbEstado = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.lblEstado = new System.Windows.Forms.Label();
			this.cmbPolitica = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.label3 = new System.Windows.Forms.Label();
			this.btnNvEstado = new System.Windows.Forms.Button();
			this.button2 = new System.Windows.Forms.Button();
			this.udsEstado = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.udsRequerimientoPoliticaObtenerTodos = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.groupBox1.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.cmbEstadoAnterior)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbEstado)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbPolitica)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.udsEstado)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.udsRequerimientoPoliticaObtenerTodos)).BeginInit();
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
			this.button1.Location = new System.Drawing.Point(104, 152);
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
			this.button3.Location = new System.Drawing.Point(16, 152);
			this.button3.Name = "button3";
			this.button3.Size = new System.Drawing.Size(76, 23);
			this.button3.TabIndex = 1067;
			this.button3.Text = "Guardar";
			this.button3.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.button3.Click += new System.EventHandler(this.button3_Click);
			// 
			// groupBox1
			// 
			this.groupBox1.Controls.Add(this.cmbEstadoAnterior);
			this.groupBox1.Controls.Add(this.label1);
			this.groupBox1.Controls.Add(this.cmbEstado);
			this.groupBox1.Controls.Add(this.lblEstado);
			this.groupBox1.Controls.Add(this.cmbPolitica);
			this.groupBox1.Controls.Add(this.label3);
			this.groupBox1.Location = new System.Drawing.Point(16, 16);
			this.groupBox1.Name = "groupBox1";
			this.groupBox1.Size = new System.Drawing.Size(352, 120);
			this.groupBox1.TabIndex = 1066;
			this.groupBox1.TabStop = false;
			this.groupBox1.Text = "Información";
			this.groupBox1.Enter += new System.EventHandler(this.groupBox1_Enter);
			// 
			// cmbEstadoAnterior
			// 
			appearance1.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbEstadoAnterior.Appearance = appearance1;
			this.cmbEstadoAnterior.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.cmbEstadoAnterior.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbEstadoAnterior.DataSource = this.udsEstado;
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridBand1.Columns.AddRange(new object[] {
																										 ultraGridColumn1,
																										 ultraGridColumn2});
			this.cmbEstadoAnterior.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.cmbEstadoAnterior.DisplayMember = "Descripcion";
			this.cmbEstadoAnterior.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbEstadoAnterior.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbEstadoAnterior.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbEstadoAnterior.Location = new System.Drawing.Point(128, 48);
			this.cmbEstadoAnterior.Name = "cmbEstadoAnterior";
			this.cmbEstadoAnterior.Size = new System.Drawing.Size(200, 21);
			this.cmbEstadoAnterior.TabIndex = 1063;
			this.cmbEstadoAnterior.ValueMember = "idEstadoRequerimiento";
			// 
			// label1
			// 
			this.label1.Location = new System.Drawing.Point(16, 16);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(96, 23);
			this.label1.TabIndex = 1062;
			this.label1.Text = "* Estado:";
			// 
			// cmbEstado
			// 
			appearance2.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbEstado.Appearance = appearance2;
			this.cmbEstado.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.cmbEstado.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbEstado.DataSource = this.udsEstado;
			ultraGridColumn3.Header.VisiblePosition = 0;
			ultraGridColumn3.Hidden = true;
			ultraGridColumn4.Header.VisiblePosition = 1;
			ultraGridBand2.Columns.AddRange(new object[] {
																										 ultraGridColumn3,
																										 ultraGridColumn4});
			this.cmbEstado.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			this.cmbEstado.DisplayMember = "Descripcion";
			this.cmbEstado.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbEstado.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbEstado.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbEstado.Location = new System.Drawing.Point(128, 16);
			this.cmbEstado.Name = "cmbEstado";
			this.cmbEstado.Size = new System.Drawing.Size(200, 21);
			this.cmbEstado.TabIndex = 1061;
			this.cmbEstado.ValueMember = "idEstadoRequerimiento";
			this.cmbEstado.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.cmbEstado_InitializeLayout);
			// 
			// lblEstado
			// 
			this.lblEstado.Location = new System.Drawing.Point(16, 48);
			this.lblEstado.Name = "lblEstado";
			this.lblEstado.Size = new System.Drawing.Size(96, 23);
			this.lblEstado.TabIndex = 1060;
			this.lblEstado.Text = "* Estado anterior:";
			// 
			// cmbPolitica
			// 
			appearance3.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbPolitica.Appearance = appearance3;
			this.cmbPolitica.CausesValidation = false;
			this.cmbPolitica.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbPolitica.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbPolitica.DataSource = this.udsRequerimientoPoliticaObtenerTodos;
			ultraGridColumn5.Header.VisiblePosition = 0;
			ultraGridColumn6.Header.Caption = "Descripción";
			ultraGridColumn6.Header.VisiblePosition = 1;
			ultraGridColumn6.Width = 218;
			ultraGridColumn7.Header.VisiblePosition = 2;
			ultraGridColumn8.Header.VisiblePosition = 3;
			ultraGridColumn9.Header.VisiblePosition = 4;
			ultraGridColumn10.Header.VisiblePosition = 5;
			ultraGridColumn11.Header.VisiblePosition = 6;
			ultraGridBand3.Columns.AddRange(new object[] {
																										 ultraGridColumn5,
																										 ultraGridColumn6,
																										 ultraGridColumn7,
																										 ultraGridColumn8,
																										 ultraGridColumn9,
																										 ultraGridColumn10,
																										 ultraGridColumn11});
			this.cmbPolitica.DisplayLayout.BandsSerializer.Add(ultraGridBand3);
			this.cmbPolitica.DisplayMember = "Descripcion";
			this.cmbPolitica.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbPolitica.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbPolitica.Location = new System.Drawing.Point(128, 80);
			this.cmbPolitica.MaxDropDownItems = 30;
			this.cmbPolitica.Name = "cmbPolitica";
			this.cmbPolitica.Size = new System.Drawing.Size(200, 21);
			this.cmbPolitica.TabIndex = 1056;
			this.cmbPolitica.ValueMember = "idRequerimientoPolitica";
			// 
			// label3
			// 
			this.label3.Location = new System.Drawing.Point(16, 80);
			this.label3.Name = "label3";
			this.label3.Size = new System.Drawing.Size(96, 23);
			this.label3.TabIndex = 1057;
			this.label3.Text = "* Política:";
			// 
			// btnNvEstado
			// 
			this.btnNvEstado.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btnNvEstado.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(230)), ((System.Byte)(230)), ((System.Byte)(230)));
			this.btnNvEstado.CausesValidation = false;
			this.btnNvEstado.Image = ((System.Drawing.Image)(resources.GetObject("btnNvEstado.Image")));
			this.btnNvEstado.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnNvEstado.Location = new System.Drawing.Point(192, 152);
			this.btnNvEstado.Name = "btnNvEstado";
			this.btnNvEstado.Size = new System.Drawing.Size(80, 24);
			this.btnNvEstado.TabIndex = 1077;
			this.btnNvEstado.Text = "&Nuevo est.";
			this.btnNvEstado.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnNvEstado.Click += new System.EventHandler(this.btnNvEstado_Click);
			// 
			// button2
			// 
			this.button2.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.button2.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(230)), ((System.Byte)(230)), ((System.Byte)(230)));
			this.button2.CausesValidation = false;
			this.button2.Image = ((System.Drawing.Image)(resources.GetObject("button2.Image")));
			this.button2.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.button2.Location = new System.Drawing.Point(280, 152);
			this.button2.Name = "button2";
			this.button2.Size = new System.Drawing.Size(80, 24);
			this.button2.TabIndex = 1078;
			this.button2.Text = "&Nuevo pol.";
			this.button2.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.button2.Click += new System.EventHandler(this.button2_Click);
			// 
			// udsEstado
			// 
			ultraDataColumn1.DataType = typeof(int);
			this.udsEstado.Band.Columns.AddRange(new object[] {
																													ultraDataColumn1,
																													ultraDataColumn2});
			// 
			// udsRequerimientoPoliticaObtenerTodos
			// 
			ultraDataColumn3.DataType = typeof(int);
			ultraDataColumn5.DataType = typeof(int);
			ultraDataColumn8.DataType = typeof(bool);
			ultraDataColumn9.DataType = typeof(int);
			this.udsRequerimientoPoliticaObtenerTodos.Band.Columns.AddRange(new object[] {
																																										 ultraDataColumn3,
																																										 ultraDataColumn4,
																																										 ultraDataColumn5,
																																										 ultraDataColumn6,
																																										 ultraDataColumn7,
																																										 ultraDataColumn8,
																																										 ultraDataColumn9});
			// 
			// frmRequerimientoPoliticaEstadoCrear
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(384, 190);
			this.Controls.Add(this.button2);
			this.Controls.Add(this.btnNvEstado);
			this.Controls.Add(this.button1);
			this.Controls.Add(this.button3);
			this.Controls.Add(this.groupBox1);
			this.Name = "frmRequerimientoPoliticaEstadoCrear";
			this.Text = "Política - Estado";
			this.Load += new System.EventHandler(this.frmRequerimientoPoliticaEstadoCrear_Load);
			this.groupBox1.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.cmbEstadoAnterior)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbEstado)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbPolitica)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.udsEstado)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.udsRequerimientoPoliticaObtenerTodos)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void frmRequerimientoPoliticaEstadoCrear_Load(object sender, System.EventArgs e)
		{
			frmRequerimientoConfiguracion.RequerimientoEstadoObtenerActivos(this.cmbEstado);
			frmRequerimientoConfiguracion.RequerimientoEstadoObtenerActivos(this.cmbEstadoAnterior);
			frmRequerimientoConfiguracion.RequerimientoPoliticaObtenerTodos(this.cmbPolitica);
		}

		private void groupBox1_Enter(object sender, System.EventArgs e)
		{
		
		}

		private void cmbEstado_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}

		private void button3_Click(object sender, System.EventArgs e)
		{
			#region Validación
			if (!Validacion.vbComboVacio(this.cmbEstado, "Seleccione el estado ")) return;
			if (!Validacion.vbComboVacio(this.cmbEstadoAnterior, "Seleccione el estado anterior ")) return;
			if (!Validacion.vbComboVacio(this.cmbPolitica, "Seleccione la pólitica ")) return;
			#endregion Validación

			int idEstadoRequerimiento = (int)this.cmbEstado.Value;
			int idEstadoRequerimientoAnterior = (int)this.cmbEstadoAnterior.Value;
			int idRequerimientoPolitica = (int)this.cmbPolitica.Value;

			try
			{
				using (SqlConnection conn = new SqlConnection(frmRequerimiento.sconexionPoint))
				{
					conn.Open();
					using (SqlCommand cmd = new SqlCommand("RequerimientoPoliticaEstadoInsertar", conn))
					{
						cmd.CommandType = CommandType.StoredProcedure;

						// Agregar parámetros
						cmd.Parameters.Add(new SqlParameter("@idRequerimientoPolitica", SqlDbType.Int));
						cmd.Parameters["@idRequerimientoPolitica"].Value = idRequerimientoPolitica;

						cmd.Parameters.Add(new SqlParameter("@idEstadoRequerimiento", SqlDbType.Int));
						cmd.Parameters["@idEstadoRequerimiento"].Value = idEstadoRequerimiento;

						cmd.Parameters.Add(new SqlParameter("@idEstadoRequerimientoAnterior", SqlDbType.Int));
						cmd.Parameters["@idEstadoRequerimientoAnterior"].Value = idEstadoRequerimientoAnterior;

						// Ejecutar el procedimiento almacenado y obtener el ID recién insertado
						int idRequerimientoPoliticaEstado = Convert.ToInt32(cmd.ExecuteScalar());

						// Mostrar el ID insertado
						MessageBox.Show("Registro insertado correctamente con ID: " + idRequerimientoPoliticaEstado, "Éxito", MessageBoxButtons.OK, MessageBoxIcon.Information);
						this.DialogResult = DialogResult.OK;
					}
				}
			}
			catch (Exception ex)
			{
				// Manejar cualquier error
				MessageBox.Show("Ocurrió un error al insertar el registro: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
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

		private void button2_Click(object sender, System.EventArgs e)
		{
			using (frmRequerimientoPoliticaCrear frmrpc = new frmRequerimientoPoliticaCrear())
			{		
				if (DialogResult.OK == frmrpc.ShowDialog())
				{
				}
				//Cargar Estados Activos
				frmRequerimientoConfiguracion.RequerimientoPoliticaObtenerTodos(this.cmbPolitica);
			}	
		}

		private void button1_Click(object sender, System.EventArgs e)
		{
			this.Close();
		}
	}
}
