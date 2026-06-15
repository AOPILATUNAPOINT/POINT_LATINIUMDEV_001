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
	/// Descripción breve de Usuario.
	/// </summary>
	public class UsuarioClave : DevExpress.XtraEditors.XtraForm
	{
		private System.Windows.Forms.Label label1;
		private System.Windows.Forms.Label label2;
		private System.Windows.Forms.ErrorProvider errorProvider;
		private System.Windows.Forms.Label label3;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtRepitaClave;
		private System.Windows.Forms.Label label4;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;
		private Infragistics.Win.Misc.UltraButton btAceptar;
		private Infragistics.Win.Misc.UltraButton btCancelar;
		private C1.Data.C1DataView cdvUsuario;
		private System.Windows.Forms.HelpProvider helpProvider;
		private C1.Data.C1DataSet cdsSeteo;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtNombre;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtClaveActual1;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtClaveNueva;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtClaveActual;
		
		public UsuarioClave()
		{
			InitializeComponent();			
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

		#region Windows Form Designer generated code
		/// <summary>
		/// Método necesario para admitir el Diseñador, no se puede modificar
		/// el contenido del método con el editor de código.
		/// </summary>
		private void InitializeComponent()
		{
			this.label1 = new System.Windows.Forms.Label();
			this.label2 = new System.Windows.Forms.Label();
			this.errorProvider = new System.Windows.Forms.ErrorProvider();
			this.btAceptar = new Infragistics.Win.Misc.UltraButton();
			this.btCancelar = new Infragistics.Win.Misc.UltraButton();
			this.txtClaveActual = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtClaveNueva = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label3 = new System.Windows.Forms.Label();
			this.txtRepitaClave = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label4 = new System.Windows.Forms.Label();
			this.cdvUsuario = new C1.Data.C1DataView();
			this.helpProvider = new System.Windows.Forms.HelpProvider();
			this.cdsSeteo = new C1.Data.C1DataSet();
			this.txtNombre = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtClaveActual1 = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			((System.ComponentModel.ISupportInitialize)(this.txtClaveActual)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtClaveNueva)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtRepitaClave)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvUsuario)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNombre)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtClaveActual1)).BeginInit();
			this.SuspendLayout();
			// 
			// label1
			// 
			this.label1.AutoSize = true;
			this.label1.Location = new System.Drawing.Point(8, 11);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(44, 16);
			this.label1.TabIndex = 2;
			this.label1.Text = "Nombre";
			this.label1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label2
			// 
			this.label2.AutoSize = true;
			this.label2.Location = new System.Drawing.Point(8, 43);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(95, 16);
			this.label2.TabIndex = 4;
			this.label2.Text = "Contraseña Actual";
			this.label2.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.label2.Click += new System.EventHandler(this.label2_Click);
			// 
			// errorProvider
			// 
			this.errorProvider.ContainerControl = this;
			this.errorProvider.DataMember = "";
			// 
			// btAceptar
			// 
			this.btAceptar.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.btAceptar.Location = new System.Drawing.Point(8, 136);
			this.btAceptar.Name = "btAceptar";
			this.btAceptar.Size = new System.Drawing.Size(75, 25);
			this.btAceptar.TabIndex = 4;
			this.btAceptar.Text = "&Aceptar";
			this.btAceptar.Click += new System.EventHandler(this.btAceptar_Click);
			// 
			// btCancelar
			// 
			this.btCancelar.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.btCancelar.CausesValidation = false;
			this.btCancelar.Location = new System.Drawing.Point(88, 136);
			this.btCancelar.Name = "btCancelar";
			this.btCancelar.Size = new System.Drawing.Size(75, 25);
			this.btCancelar.TabIndex = 5;
			this.btCancelar.Text = "&Cancelar";
			this.btCancelar.Click += new System.EventHandler(this.btCancelar_Click);
			// 
			// txtClaveActual
			// 
			this.txtClaveActual.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtClaveActual.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtClaveActual.Location = new System.Drawing.Point(144, 40);
			this.txtClaveActual.MaxLength = 12;
			this.txtClaveActual.Name = "txtClaveActual";
			this.txtClaveActual.PasswordChar = '*';
			this.txtClaveActual.Size = new System.Drawing.Size(208, 22);
			this.txtClaveActual.TabIndex = 1;
			this.txtClaveActual.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtClaveActual_KeyDown);
			this.txtClaveActual.Validating += new System.ComponentModel.CancelEventHandler(this.txtClaveActual_Validating);
			this.txtClaveActual.ValueChanged += new System.EventHandler(this.txtClaveActual_ValueChanged);
			// 
			// txtClaveNueva
			// 
			this.txtClaveNueva.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtClaveNueva.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtClaveNueva.Location = new System.Drawing.Point(144, 72);
			this.txtClaveNueva.MaxLength = 12;
			this.txtClaveNueva.Name = "txtClaveNueva";
			this.txtClaveNueva.PasswordChar = '*';
			this.txtClaveNueva.Size = new System.Drawing.Size(208, 22);
			this.txtClaveNueva.TabIndex = 2;
			this.txtClaveNueva.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtClaveNueva_KeyDown);
			this.txtClaveNueva.Validating += new System.ComponentModel.CancelEventHandler(this.txtClaveNueva_Validating);
			this.txtClaveNueva.ValueChanged += new System.EventHandler(this.txtClaveNueva_ValueChanged);
			// 
			// label3
			// 
			this.label3.AutoSize = true;
			this.label3.Location = new System.Drawing.Point(8, 75);
			this.label3.Name = "label3";
			this.label3.Size = new System.Drawing.Size(96, 16);
			this.label3.TabIndex = 9;
			this.label3.Text = "Nueva Contraseña";
			this.label3.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.label3.Click += new System.EventHandler(this.label3_Click);
			// 
			// txtRepitaClave
			// 
			this.txtRepitaClave.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtRepitaClave.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtRepitaClave.Location = new System.Drawing.Point(144, 104);
			this.txtRepitaClave.MaxLength = 12;
			this.txtRepitaClave.Name = "txtRepitaClave";
			this.txtRepitaClave.PasswordChar = '*';
			this.txtRepitaClave.Size = new System.Drawing.Size(208, 22);
			this.txtRepitaClave.TabIndex = 3;
			this.txtRepitaClave.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtRepitaClave_KeyDown);
			this.txtRepitaClave.Validating += new System.ComponentModel.CancelEventHandler(this.txtRepitaClave_Validating);
			this.txtRepitaClave.ValueChanged += new System.EventHandler(this.txtRepitaClave_ValueChanged);
			// 
			// label4
			// 
			this.label4.AutoSize = true;
			this.label4.Location = new System.Drawing.Point(8, 107);
			this.label4.Name = "label4";
			this.label4.Size = new System.Drawing.Size(109, 16);
			this.label4.TabIndex = 11;
			this.label4.Text = "Confirme Contraseña";
			this.label4.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.label4.Click += new System.EventHandler(this.label4_Click);
			// 
			// cdvUsuario
			// 
			this.cdvUsuario.BindingContextCtrl = this;
			this.cdvUsuario.DataSet = null;
			this.cdvUsuario.Sort = "idUsuario";
			this.cdvUsuario.TableName = "";
			this.cdvUsuario.TableViewName = "Usuario";
			// 
			// helpProvider
			// 
			this.helpProvider.HelpNamespace = "C:\\Latinium\\Latinium.chm";
			// 
			// cdsSeteo
			// 
			this.cdsSeteo.BindingContextCtrl = this;
			this.cdsSeteo.DataLibrary = "LibContabilidad";
			this.cdsSeteo.DataLibraryUrl = "";
			this.cdsSeteo.DataSetDef = "dsSeteo";
			this.cdsSeteo.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsSeteo.Name = "cdsSeteo";
			this.cdsSeteo.SchemaClassName = "LibContabilidad.DataClass";
			this.cdsSeteo.SchemaDef = null;
			this.cdsSeteo.BeforeFill += new C1.Data.FillEventHandler(this.cdsSeteo_BeforeFill);
			// 
			// txtNombre
			// 
			this.txtNombre.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtNombre.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNombre.Enabled = false;
			this.txtNombre.Location = new System.Drawing.Point(144, 8);
			this.txtNombre.Name = "txtNombre";
			this.txtNombre.Size = new System.Drawing.Size(208, 22);
			this.txtNombre.TabIndex = 0;
			// 
			// txtClaveActual1
			// 
			this.txtClaveActual1.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtClaveActual1.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtClaveActual1.Enabled = false;
			this.txtClaveActual1.Location = new System.Drawing.Point(176, 136);
			this.txtClaveActual1.Name = "txtClaveActual1";
			this.txtClaveActual1.Size = new System.Drawing.Size(176, 22);
			this.txtClaveActual1.TabIndex = 6;
			this.txtClaveActual1.Visible = false;
			// 
			// UsuarioClave
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 14);
			this.CausesValidation = false;
			this.ClientSize = new System.Drawing.Size(360, 166);
			this.Controls.Add(this.txtClaveActual1);
			this.Controls.Add(this.txtNombre);
			this.Controls.Add(this.txtRepitaClave);
			this.Controls.Add(this.label4);
			this.Controls.Add(this.label3);
			this.Controls.Add(this.label2);
			this.Controls.Add(this.label1);
			this.Controls.Add(this.txtClaveNueva);
			this.Controls.Add(this.txtClaveActual);
			this.Controls.Add(this.btCancelar);
			this.Controls.Add(this.btAceptar);
			this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
			this.helpProvider.SetHelpKeyword(this, "Cambiocontrasena.htm");
			this.helpProvider.SetHelpNavigator(this, System.Windows.Forms.HelpNavigator.Topic);
			this.KeyPreview = true;
			this.MaximizeBox = false;
			this.MinimizeBox = false;
			this.Name = "UsuarioClave";
			this.helpProvider.SetShowHelp(this, true);
			this.ShowInTaskbar = false;
			this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
			this.Text = "Cambio de Contraseña";
			this.Closing += new System.ComponentModel.CancelEventHandler(this.UsuarioClave_Closing);
			this.Load += new System.EventHandler(this.Usuario_Load);
			((System.ComponentModel.ISupportInitialize)(this.txtClaveActual)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtClaveNueva)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtRepitaClave)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvUsuario)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNombre)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtClaveActual1)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void Usuario_Load(object sender, System.EventArgs e)
		{
			string stAudita = "Exec AuditaCrear 53, 6, 'Usuario Clave'";
			Funcion.EjecutaSQL(cdsSeteo, stAudita, true);
			Cursor = Cursors.Default;

			string sSQL = string.Format("Select Nombre, Clave From Usuario Where idUsuario = {0}", MenuLatinium.IdUsuario);
			SqlDataReader drUsuario = Funcion.rEscalarSQL(cdsSeteo, sSQL, true);
			drUsuario.Read();
			if (drUsuario.HasRows)
			{
				this.txtNombre.Text = drUsuario.GetString(0);
				this.txtClaveActual1.Text = drUsuario.GetString(1);
			}
			drUsuario.Close();

			this.txtClaveActual.Focus();
		}

		private void btCancelar_Click(object sender, System.EventArgs e)
		{
			this.Close();
		}

		private void btAceptar_Click(object sender, System.EventArgs e)
		{
			if (this.txtClaveActual.Text.ToString().Length == 0)
			{
				MessageBox.Show("Escriba su Clave Actual", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.txtClaveActual.Focus();				
				return;
			}

			if (this.txtClaveActual.Text.ToString().Trim() != this.txtClaveActual1.Text.ToString().Trim())
			{
				MessageBox.Show("Su Clave Actual no coincide con la Clave registrada en el Sistema", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.txtClaveActual.Focus();			
				return;
			}

			if (this.txtRepitaClave.Text.ToString().Length == 0)
			{
				MessageBox.Show("Repita su Nueva Clave", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.txtRepitaClave.Focus();				
				return;
			}

			if (this.txtClaveNueva.Text.ToString().Trim() != this.txtRepitaClave.Text.ToString().Trim())
			{
				MessageBox.Show("No Coincide su Nueva Clave", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.txtRepitaClave.Focus();			
				return;
			}

			if (this.txtClaveNueva.Text.ToString().Length == 0)
			{
				MessageBox.Show("Escriba su Nueva Clave", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.txtClaveNueva.Focus();			
				return;
			}

			if (this.txtClaveActual.Text.ToString() == this.txtClaveNueva.Text.ToString())
			{
				MessageBox.Show("La Nueva Clave no puede ser igual que la anterior", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.txtClaveNueva.Focus();
				return;
			}

			string sSQL = string.Format("Exec UsuarioModificaClave '{0}', '{1}'", this.txtClaveNueva.Text.ToString(), this.txtNombre.Text.ToString());
			Funcion.EjecutaSQL(cdsSeteo, sSQL);

			string stAudita = "Exec AuditaCrear 60, 2, '" + this.txtClaveActual.Text.ToString() + "'";
			Funcion.EjecutaSQL(cdsSeteo, stAudita, true);

			Funcion.EjecutaSQL(cdsSeteo, string.Format("Insert Into HistorialClaveUsuarios (idUsuario, Contrasena) Values ({0}, '{1}')", MenuLatinium.IdUsuario, this.txtClaveActual.Text.ToString()), true);

			MessageBox.Show("Clave Modificada", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);

			this.Close();
		}

		private void UsuarioClave_Closing(object sender, System.ComponentModel.CancelEventArgs e)
		{
			string stAudita = "Exec AuditaCrear 53, 7, 'Usuario Clave'";
			Funcion.EjecutaSQL(cdsSeteo, stAudita, true);
		}

		private void cdsSeteo_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			this.cdsSeteo.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirConta);
		}

		private void txtClaveActual_Validating(object sender, System.ComponentModel.CancelEventArgs e)
		{
			if (this.txtClaveActual.Text.ToString().Length == 0)
			{
				MessageBox.Show("Escriba su Clave Actual", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.txtClaveActual.Focus();
				e.Cancel = true;
				return;
			}

			if (this.txtClaveActual.Text.ToString().Trim() != this.txtClaveActual1.Text.ToString().Trim())
			{
				MessageBox.Show("Su Clave Actual no coincide con la Clave registrada en el Sistema", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.txtClaveActual.Focus();
				e.Cancel = true;
				return;
			}
		}

		private void txtClaveActual_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.txtClaveNueva.Focus();
		}

		private void txtClaveNueva_Validating(object sender, System.ComponentModel.CancelEventArgs e)
		{
			if (this.txtClaveNueva.Text.ToString().Length == 0)
			{
				MessageBox.Show("Escriba su Nueva Clave", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.txtClaveNueva.Focus();
				e.Cancel = true;
				return;
			}
		}

		private void txtClaveNueva_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.txtRepitaClave.Focus();
		}

		private void txtRepitaClave_Validating(object sender, System.ComponentModel.CancelEventArgs e)
		{
			if (this.txtRepitaClave.Text.ToString().Length == 0)
			{
				MessageBox.Show("Repita su Nueva Clave", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.txtRepitaClave.Focus();
				e.Cancel = true;
				return;
			}

			if (this.txtClaveNueva.Text.ToString().Trim() != this.txtRepitaClave.Text.ToString().Trim())
			{
				MessageBox.Show("No Coincide su Nueva Clave", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.txtRepitaClave.Focus();
				e.Cancel = true;
				return;
			}
		}

		private void txtRepitaClave_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.btAceptar.Focus();
		}

		private void txtRepitaClave_ValueChanged(object sender, System.EventArgs e)
		{
		
		}

		private void label4_Click(object sender, System.EventArgs e)
		{
		
		}

		private void label2_Click(object sender, System.EventArgs e)
		{
		
		}

		private void txtClaveActual_ValueChanged(object sender, System.EventArgs e)
		{
		
		}

		private void txtClaveNueva_ValueChanged(object sender, System.EventArgs e)
		{
		
		}

		private void label3_Click(object sender, System.EventArgs e)
		{
		
		}
	}
}
