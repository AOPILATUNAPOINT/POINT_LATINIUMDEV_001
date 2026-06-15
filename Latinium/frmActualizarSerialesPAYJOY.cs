using System;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de frmActualizarSerialesPAYJOY.
	/// </summary>
	public class frmActualizarSerialesPAYJOY : System.Windows.Forms.Form
	{
		private System.Windows.Forms.Label lblBodega;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbBodega;
		private System.Windows.Forms.GroupBox groupBox1;
		private System.Windows.Forms.Button btnActualizar;
		private System.Windows.Forms.Label label72;
		private System.Windows.Forms.Button btnBuscar;
		private C1.Data.C1DataSet cdsSeteoF;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtNumero;
		private System.Windows.Forms.Label label2;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtSerial;
		private System.Windows.Forms.Button btnValidaNuevo;
		private System.Windows.Forms.Button btnCancelar;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public frmActualizarSerialesPAYJOY()
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
		/// 		private Acceso miAcceso;
		private Acceso miAcceso;
		private void UnloadMe()
		{
			this.Close();
		}
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
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bodega");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo", 0);
			Infragistics.Win.Appearance appearance1 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			this.lblBodega = new System.Windows.Forms.Label();
			this.cmbBodega = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.groupBox1 = new System.Windows.Forms.GroupBox();
			this.btnActualizar = new System.Windows.Forms.Button();
			this.label72 = new System.Windows.Forms.Label();
			this.btnBuscar = new System.Windows.Forms.Button();
			this.txtNumero = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.label2 = new System.Windows.Forms.Label();
			this.txtSerial = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.btnCancelar = new System.Windows.Forms.Button();
			this.btnValidaNuevo = new System.Windows.Forms.Button();
			((System.ComponentModel.ISupportInitialize)(this.cmbBodega)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNumero)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtSerial)).BeginInit();
			this.SuspendLayout();
			// 
			// lblBodega
			// 
			this.lblBodega.AutoSize = true;
			this.lblBodega.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblBodega.Location = new System.Drawing.Point(7, 11);
			this.lblBodega.Name = "lblBodega";
			this.lblBodega.Size = new System.Drawing.Size(41, 17);
			this.lblBodega.TabIndex = 329;
			this.lblBodega.Text = "Bodega";
			this.lblBodega.TextAlign = System.Drawing.ContentAlignment.TopRight;
			// 
			// cmbBodega
			// 
			this.cmbBodega.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbBodega.Cursor = System.Windows.Forms.Cursors.Default;
			ultraGridColumn1.Header.VisiblePosition = 2;
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridColumn2.Width = 180;
			ultraGridColumn3.Header.VisiblePosition = 0;
			ultraGridBand1.Columns.AddRange(new object[] {
																										 ultraGridColumn1,
																										 ultraGridColumn2,
																										 ultraGridColumn3});
			this.cmbBodega.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.cmbBodega.DisplayMember = "Nombre";
			this.cmbBodega.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbBodega.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbBodega.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbBodega.Location = new System.Drawing.Point(80, 8);
			this.cmbBodega.MaxDropDownItems = 30;
			this.cmbBodega.Name = "cmbBodega";
			this.cmbBodega.Size = new System.Drawing.Size(256, 22);
			this.cmbBodega.TabIndex = 330;
			this.cmbBodega.ValueMember = "Bodega";
			// 
			// groupBox1
			// 
			this.groupBox1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.groupBox1.Location = new System.Drawing.Point(0, 128);
			this.groupBox1.Name = "groupBox1";
			this.groupBox1.Size = new System.Drawing.Size(452, 8);
			this.groupBox1.TabIndex = 328;
			this.groupBox1.TabStop = false;
			// 
			// btnActualizar
			// 
			this.btnActualizar.Enabled = false;
			this.btnActualizar.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.btnActualizar.Location = new System.Drawing.Point(16, 152);
			this.btnActualizar.Name = "btnActualizar";
			this.btnActualizar.Size = new System.Drawing.Size(88, 23);
			this.btnActualizar.TabIndex = 325;
			this.btnActualizar.Text = "Asignar Serie";
			this.btnActualizar.Click += new System.EventHandler(this.btnActualizar_Click);
			// 
			// label72
			// 
			this.label72.AutoSize = true;
			this.label72.BackColor = System.Drawing.Color.Transparent;
			this.label72.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label72.Location = new System.Drawing.Point(7, 56);
			this.label72.Name = "label72";
			this.label72.Size = new System.Drawing.Size(41, 17);
			this.label72.TabIndex = 326;
			this.label72.Text = "Factura";
			// 
			// btnBuscar
			// 
			this.btnBuscar.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.btnBuscar.Location = new System.Drawing.Point(352, 48);
			this.btnBuscar.Name = "btnBuscar";
			this.btnBuscar.Size = new System.Drawing.Size(72, 23);
			this.btnBuscar.TabIndex = 322;
			this.btnBuscar.Text = "&Buscar";
			this.btnBuscar.Click += new System.EventHandler(this.btnBuscar_Click);
			// 
			// txtNumero
			// 
			this.txtNumero.AcceptsTab = true;
			appearance1.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtNumero.Appearance = appearance1;
			this.txtNumero.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtNumero.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNumero.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtNumero.Location = new System.Drawing.Point(80, 48);
			this.txtNumero.Name = "txtNumero";
			this.txtNumero.Size = new System.Drawing.Size(256, 22);
			this.txtNumero.TabIndex = 321;
			this.txtNumero.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtNumero_KeyPress);
			// 
			// cdsSeteoF
			// 
			this.cdsSeteoF.BindingContextCtrl = this;
			this.cdsSeteoF.DataLibrary = "LibFacturacion";
			this.cdsSeteoF.DataLibraryUrl = "";
			this.cdsSeteoF.DataSetDef = "dsSeteoF";
			this.cdsSeteoF.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsSeteoF.Name = "cdsSeteoF";
			this.cdsSeteoF.SchemaClassName = "LibFacturacion.DataClass";
			this.cdsSeteoF.SchemaDef = null;
			this.cdsSeteoF.BeforeFill += new C1.Data.FillEventHandler(this.cdsSeteoF_BeforeFill);
			// 
			// label2
			// 
			this.label2.AutoSize = true;
			this.label2.BackColor = System.Drawing.Color.Transparent;
			this.label2.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label2.Location = new System.Drawing.Point(8, 96);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(32, 17);
			this.label2.TabIndex = 332;
			this.label2.Text = "Serial";
			// 
			// txtSerial
			// 
			this.txtSerial.AcceptsTab = true;
			appearance2.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtSerial.Appearance = appearance2;
			this.txtSerial.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtSerial.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtSerial.Enabled = false;
			this.txtSerial.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtSerial.Location = new System.Drawing.Point(80, 88);
			this.txtSerial.Name = "txtSerial";
			this.txtSerial.Size = new System.Drawing.Size(256, 22);
			this.txtSerial.TabIndex = 331;
			// 
			// btnCancelar
			// 
			this.btnCancelar.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.btnCancelar.Location = new System.Drawing.Point(120, 152);
			this.btnCancelar.Name = "btnCancelar";
			this.btnCancelar.Size = new System.Drawing.Size(72, 23);
			this.btnCancelar.TabIndex = 333;
			this.btnCancelar.Text = "Cancelar";
			this.btnCancelar.Click += new System.EventHandler(this.bntbuscarSerial_Click);
			// 
			// btnValidaNuevo
			// 
			this.btnValidaNuevo.Enabled = false;
			this.btnValidaNuevo.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.btnValidaNuevo.Location = new System.Drawing.Point(352, 88);
			this.btnValidaNuevo.Name = "btnValidaNuevo";
			this.btnValidaNuevo.Size = new System.Drawing.Size(72, 23);
			this.btnValidaNuevo.TabIndex = 334;
			this.btnValidaNuevo.Text = "&Validar";
			this.btnValidaNuevo.Click += new System.EventHandler(this.btnValidaNuevo_Click);
			// 
			// frmActualizarSerialesPAYJOY
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 14);
			this.ClientSize = new System.Drawing.Size(450, 184);
			this.Controls.Add(this.btnValidaNuevo);
			this.Controls.Add(this.btnCancelar);
			this.Controls.Add(this.label2);
			this.Controls.Add(this.txtSerial);
			this.Controls.Add(this.lblBodega);
			this.Controls.Add(this.cmbBodega);
			this.Controls.Add(this.groupBox1);
			this.Controls.Add(this.btnActualizar);
			this.Controls.Add(this.label72);
			this.Controls.Add(this.btnBuscar);
			this.Controls.Add(this.txtNumero);
			this.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
			this.KeyPreview = true;
			this.MaximizeBox = false;
			this.MinimizeBox = false;
			this.Name = "frmActualizarSerialesPAYJOY";
			this.ShowInTaskbar = false;
			this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
			this.Text = "Actualizacion de Seriales PAY-JOY";
			this.Load += new System.EventHandler(this.frmActualizarSerialesPAYJOY_Load);
			((System.ComponentModel.ISupportInitialize)(this.cmbBodega)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNumero)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtSerial)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);
		}

		private void frmActualizarSerialesPAYJOY_Load(object sender, System.EventArgs e)
		{
			miAcceso = new Acceso(cdsSeteoF, "0332");

			if (!Funcion.Permiso("2141", cdsSeteoF))
			{				
				MessageBox.Show("No puede ingresar a Actualizacion de seriales PAYJOY.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				BeginInvoke(new MethodInvoker(UnloadMe));
				return;
			}

			DateTime dFecha = DateTime.Today;

			string sSQL = string.Format("Exec BodegaAsignadaPorUsuario {0}", MenuLatinium.IdUsuario);		

			int iBodega = Funcion.iEscalarSQL(cdsSeteoF, sSQL);

			this.cmbBodega.DataSource = Funcion.dvProcedimiento(
				cdsSeteoF, 
				string.Format("Exec FacturacionListaBodegas {0}, 43, '{1}'", 
				MenuLatinium.IdUsuario, dFecha.ToString("yyyyMMdd"))
				);

			if (iBodega == 0) 
				this.cmbBodega.Enabled = true; 
			else 
				this.cmbBodega.Value = iBodega;
		}
		private int idCompraGuardado = 0;

		private void btnBuscar_Click(object sender, System.EventArgs e)
		{
			if (this.cmbBodega.Text == null || this.cmbBodega.Text.Trim() == "")
			{
				MessageBox.Show("Seleccione una Bodega.", "Advertencia", MessageBoxButtons.OK, MessageBoxIcon.Warning);
				this.cmbBodega.Focus();
				return;
			}
			string sSQL = string.Format("EXEC ValidarFacturaBodega {0}, '{1}'", 
				this.cmbBodega.Value, 
				this.txtNumero.Text);

			SqlDataReader dr = Funcion.rEscalarSQL(cdsSeteoF, sSQL, true);

			if (dr != null && dr.Read()) 
			{
				string mensaje = dr["Mensaje"].ToString();

				int valido = Convert.ToInt32(dr["Valido"]);

				int idCompra = 0;
				if (dr.FieldCount > 2 && !dr.IsDBNull(dr.GetOrdinal("idCompra")))
				{
					idCompra = Convert.ToInt32(dr["idCompra"]);
				}
				idCompraGuardado = idCompra;

				if (valido == 1)
				{
					txtSerial.Enabled = true;
					btnValidaNuevo.Enabled = true;
					btnBuscar.Enabled = false;
					txtNumero.Enabled = false;
				}
				else
				{
					txtSerial.Enabled = false;
					btnActualizar.Enabled = false;
					btnValidaNuevo.Enabled = false;
				}

				dr.Close();

				MessageBox.Show(mensaje, "Resultado", MessageBoxButtons.OK, 
					valido == 1 ? MessageBoxIcon.Information : MessageBoxIcon.Warning);
			}
			else
			{
				MessageBox.Show("No se obtuvo respuesta del servidor.", "Error", 
					MessageBoxButtons.OK, MessageBoxIcon.Error);
			}
		}

		private void btnActualizar_Click(object sender, System.EventArgs e)
		{
			
			if (idCompraGuardado == 0)
			{
				MessageBox.Show("No hay una compra válida para actualizar.", "Error", 
					MessageBoxButtons.OK, MessageBoxIcon.Error);
				return;
			}

			if (txtSerial.Text == null || txtSerial.Text.Trim() == "")
			{
				MessageBox.Show("Debe ingresar un serial nuevo.", "Advertencia", 
					MessageBoxButtons.OK, MessageBoxIcon.Warning);
				return;
			}
			string sSQL = string.Format("Exec ActualizaSeriePayJoy {0}, '{1}'",txtSerial.Text,idCompraGuardado);
			bool resultado = Funcion.EjecutaSQL(cdsSeteoF, sSQL);

			string sSQL2 = string.Format("Exec CreaAsientoDeVenta {0}" ,idCompraGuardado);
			Funcion.EjecutaSQL(cdsSeteoF, sSQL2);

			if (resultado)
			{
				MessageBox.Show("Serial actualizado correctamente.", "Éxito", 
					MessageBoxButtons.OK, MessageBoxIcon.Information);

				idCompraGuardado = 0;
				txtSerial.Text = "";
				txtNumero.Text = "";
				txtSerial.Text= "";

				txtSerial.Enabled = false;
				btnActualizar.Enabled = false;
				txtSerial.Enabled = false;
				bntbuscarSerial_Click(null, EventArgs.Empty);
			}
			else
			{
				MessageBox.Show("No se pudo actualizar el serial.", "Error", 
					MessageBoxButtons.OK, MessageBoxIcon.Error);
			}
		}
		private void bntbuscarSerial_Click(object sender, System.EventArgs e)
		{
			this.txtSerial.Text = "";
			this.txtNumero.Text = "";

			this.btnActualizar.Enabled = false;
			this.btnValidaNuevo.Enabled = false;		
			this.txtSerial.Enabled = false;
			this.btnBuscar.Enabled = true;
			this.txtNumero.Enabled = true;
			this.Close();

		}

		private void btnValidaNuevo_Click(object sender, System.EventArgs e)
		{
			this.btnActualizar.Enabled = false;
			if (!Validacion.vbTexto(this.txtSerial, 4, 1000, "Nueva Serie")) return;
			string sSQL = string.Format("EXEC ValidaSeriaPayJoy {0}, '{1}'", this.cmbBodega.Value, this.txtSerial.Text.ToString());
			SqlDataReader dr = Funcion.rEscalarSQL(cdsSeteoF, sSQL, true);

			if (dr.Read())
			{
				string mensaje = dr[0].ToString();
				MessageBox.Show(mensaje, "Validación de Serie", MessageBoxButtons.OK, MessageBoxIcon.Information);

				if (mensaje == "LA SERIE EXISTE Y ESTA DISPONIBLE")
				{
					this.btnActualizar.Enabled = true;
				}
				else
				{
					this.btnActualizar.Enabled = false;
				}
			}

			dr.Close();
			btnValidaNuevo.Enabled = false;
			txtSerial.Enabled = false;
		}

		private void txtNumero_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar))
			{
				e.Handled = true;
			}
		}
	}
}
