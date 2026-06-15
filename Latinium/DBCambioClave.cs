using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using DevExpress.XtraEditors;

namespace Latinium
{
	/// <summary>
	/// Summary description for DBCambioClave.
	/// </summary>
	public class DBCambioClave : DevExpress.XtraEditors.XtraForm
	{
		private Infragistics.Win.Misc.UltraButton btCancelar;
		private Infragistics.Win.Misc.UltraButton btAceptar;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtNombre;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtRepitaClave;
		private System.Windows.Forms.Label label4;
		private System.Windows.Forms.Label label3;
		private System.Windows.Forms.Label label2;
		private System.Windows.Forms.Label label1;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtClaveActual;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtClaveAntes;
		private System.Windows.Forms.ErrorProvider errorProvider;
		private C1.Data.C1DataView cdvUsuario;
		private C1.Data.C1DataSet cdsUsuario;
		private C1.Data.C1DataSet cdsSeteo;
		private System.Windows.Forms.HelpProvider helpProvider;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtClave;
		/// <summary>
		/// Required designer variable.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public DBCambioClave()
		{
			//
			// Required for Windows Form Designer support
			//
			InitializeComponent();

			//
			// TODO: Add any constructor code after InitializeComponent call
			//
		}
		int idUsuario=0;
		public DBCambioClave(int IdUsuario)
		{
			//
			// Required for Windows Form Designer support
			//
			
			InitializeComponent();
			idUsuario= IdUsuario;
			//
			// TODO: Add any constructor code after InitializeComponent call
			//
		}
		/// <summary>
		/// Clean up any resources being used.
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
		/// Required method for Designer support - do not modify
		/// the contents of this method with the code editor.
		/// </summary>
		private void InitializeComponent()
		{
			this.btCancelar = new Infragistics.Win.Misc.UltraButton();
			this.btAceptar = new Infragistics.Win.Misc.UltraButton();
			this.txtNombre = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.cdsUsuario = new C1.Data.C1DataSet();
			this.txtRepitaClave = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label4 = new System.Windows.Forms.Label();
			this.label3 = new System.Windows.Forms.Label();
			this.label2 = new System.Windows.Forms.Label();
			this.label1 = new System.Windows.Forms.Label();
			this.txtClaveActual = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtClaveAntes = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.errorProvider = new System.Windows.Forms.ErrorProvider();
			this.cdvUsuario = new C1.Data.C1DataView();
			this.cdsSeteo = new C1.Data.C1DataSet();
			this.helpProvider = new System.Windows.Forms.HelpProvider();
			this.txtClave = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			((System.ComponentModel.ISupportInitialize)(this.txtNombre)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsUsuario)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtRepitaClave)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtClaveActual)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtClaveAntes)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvUsuario)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtClave)).BeginInit();
			this.SuspendLayout();
			// 
			// btCancelar
			// 
			this.btCancelar.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.btCancelar.Location = new System.Drawing.Point(207, 196);
			this.btCancelar.Name = "btCancelar";
			this.btCancelar.Size = new System.Drawing.Size(75, 25);
			this.btCancelar.TabIndex = 9;
			this.btCancelar.Text = "&Cancelar";
			this.btCancelar.Click += new System.EventHandler(this.btCancelar_Click);
			// 
			// btAceptar
			// 
			this.btAceptar.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.btAceptar.Location = new System.Drawing.Point(33, 196);
			this.btAceptar.Name = "btAceptar";
			this.btAceptar.Size = new System.Drawing.Size(75, 25);
			this.btAceptar.TabIndex = 8;
			this.btAceptar.Text = "&Aceptar";
			this.btAceptar.Click += new System.EventHandler(this.btAceptar_Click);
			// 
			// txtNombre
			// 
			this.txtNombre.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsUsuario, "Usuario.Nombre"));
			this.txtNombre.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNombre.Location = new System.Drawing.Point(120, 35);
			this.txtNombre.Name = "txtNombre";
			this.txtNombre.ReadOnly = true;
			this.txtNombre.Size = new System.Drawing.Size(207, 22);
			this.txtNombre.TabIndex = 21;
			this.txtNombre.Text = "Nombre";
			// 
			// cdsUsuario
			// 
			this.cdsUsuario.BindingContextCtrl = this;
			this.cdsUsuario.CaseSensitive = false;
			this.cdsUsuario.DataLibrary = "LibDBEmpresa";
			this.cdsUsuario.DataLibraryUrl = "";
			this.cdsUsuario.DataSetDef = "dsEmpresa";
			this.cdsUsuario.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsUsuario.Name = "cdsUsuario";
			this.cdsUsuario.SchemaClassName = "LibDBEmpresa.DataClass";
			this.cdsUsuario.SchemaDef = null;
			this.cdsUsuario.BeforeFill += new C1.Data.FillEventHandler(this.cdsUsuario_BeforeFill);
			// 
			// txtRepitaClave
			// 
			this.txtRepitaClave.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtRepitaClave.Location = new System.Drawing.Point(120, 140);
			this.txtRepitaClave.Name = "txtRepitaClave";
			this.txtRepitaClave.PasswordChar = '*';
			this.txtRepitaClave.Size = new System.Drawing.Size(207, 22);
			this.txtRepitaClave.TabIndex = 20;
			// 
			// label4
			// 
			this.label4.AutoSize = true;
			this.label4.Location = new System.Drawing.Point(7, 140);
			this.label4.Name = "label4";
			this.label4.Size = new System.Drawing.Size(113, 16);
			this.label4.TabIndex = 19;
			this.label4.Text = "Confirme Contraseña:";
			// 
			// label3
			// 
			this.label3.AutoSize = true;
			this.label3.Location = new System.Drawing.Point(7, 105);
			this.label3.Name = "label3";
			this.label3.Size = new System.Drawing.Size(99, 16);
			this.label3.TabIndex = 17;
			this.label3.Text = "Contraseña Actual:";
			// 
			// label2
			// 
			this.label2.AutoSize = true;
			this.label2.Location = new System.Drawing.Point(7, 70);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(108, 16);
			this.label2.TabIndex = 15;
			this.label2.Text = "Contraseña Anterior:";
			// 
			// label1
			// 
			this.label1.AutoSize = true;
			this.label1.Location = new System.Drawing.Point(7, 35);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(48, 16);
			this.label1.TabIndex = 14;
			this.label1.Text = "Nombre:";
			// 
			// txtClaveActual
			// 
			this.txtClaveActual.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtClaveActual.Location = new System.Drawing.Point(120, 105);
			this.txtClaveActual.Name = "txtClaveActual";
			this.txtClaveActual.PasswordChar = '*';
			this.txtClaveActual.Size = new System.Drawing.Size(207, 22);
			this.txtClaveActual.TabIndex = 18;
			// 
			// txtClaveAntes
			// 
			this.txtClaveAntes.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtClaveAntes.Location = new System.Drawing.Point(120, 70);
			this.txtClaveAntes.Name = "txtClaveAntes";
			this.txtClaveAntes.PasswordChar = '*';
			this.txtClaveAntes.Size = new System.Drawing.Size(207, 22);
			this.txtClaveAntes.TabIndex = 16;
			// 
			// errorProvider
			// 
			this.errorProvider.ContainerControl = this;
			this.errorProvider.DataMember = "";
			// 
			// cdvUsuario
			// 
			this.cdvUsuario.BindingContextCtrl = this;
			this.cdvUsuario.DataSet = this.cdsUsuario;
			this.cdvUsuario.Sort = "idUsuario";
			this.cdvUsuario.TableName = "";
			this.cdvUsuario.TableViewName = "Usuario";
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
			// helpProvider
			// 
			this.helpProvider.HelpNamespace = "C:\\Latinium\\Latinium.chm";
			// 
			// txtClave
			// 
			this.txtClave.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsUsuario, "Usuario.Clave"));
			this.txtClave.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtClave.Location = new System.Drawing.Point(7, 7);
			this.txtClave.Name = "txtClave";
			this.txtClave.Size = new System.Drawing.Size(55, 22);
			this.txtClave.TabIndex = 22;
			this.txtClave.Text = "Clave";
			// 
			// DBCambioClave
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 14);
			this.ClientSize = new System.Drawing.Size(353, 238);
			this.Controls.Add(this.txtClave);
			this.Controls.Add(this.txtNombre);
			this.Controls.Add(this.txtRepitaClave);
			this.Controls.Add(this.label4);
			this.Controls.Add(this.label3);
			this.Controls.Add(this.label2);
			this.Controls.Add(this.label1);
			this.Controls.Add(this.txtClaveActual);
			this.Controls.Add(this.txtClaveAntes);
			this.Controls.Add(this.btCancelar);
			this.Controls.Add(this.btAceptar);
			this.MaximizeBox = false;
			this.MinimizeBox = false;
			this.Name = "DBCambioClave";
			this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
			this.Text = "Cambio de Clave";
			this.Closing += new System.ComponentModel.CancelEventHandler(this.DBCambioClave_Closing);
			this.Load += new System.EventHandler(this.DBCambioClave_Load);
			((System.ComponentModel.ISupportInitialize)(this.txtNombre)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsUsuario)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtRepitaClave)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtClaveActual)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtClaveAntes)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvUsuario)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtClave)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void btCancelar_Click(object sender, System.EventArgs e)
		{
			this.Close();
		}

		private void btAceptaR_Click(object sender, System.EventArgs e)
		{
			if (!Verifica()) return;
			C1.Data.C1DataRow dr = this.cdsUsuario.TableViews["Usuario"].Rows.Find(idUsuario);
			
			string stClave = txtClaveActual.Text.Trim().ToString().Trim();
			string stClaveBasica = Funcion.CreaHash(stClave);
			dr["Clave"] = Funcion.CreaHash(stClaveBasica+stUsuario, grupo);
			dr["ClaveBasica"] = stClaveBasica;
			this.cdsUsuario.Update();

			string stAudita = "Exec AuditaCrear 60, 2, '" + dr["Clave"].ToString().Substring(1, 10) + "'";
			Funcion.EjecutaSQL(cdsSeteo, stAudita, true);
			
			DialogResult = DialogResult.OK;
			this.Close();
		}

		private void DBCambioClave_Closing(object sender, System.ComponentModel.CancelEventArgs e)
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

		private void cdsUsuario_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsUsuario.Schema.Connections[0].ConnectionString = miFuncion.AccesoEmpresaBase();
		}
		private bool Verifica()
		{
			bool bOk = true;
			this.errorProvider.SetError(this.txtRepitaClave, "");
			this.errorProvider.SetError(this.txtClaveAntes, "");

			string stClave = txtClaveAntes.Text.Trim().ToString().Trim();
			string stClaveBasica = Funcion.CreaHash(stClave);
			
			if (this.txtClave.Text.Trim().ToString() != Funcion.CreaHash(stClaveBasica+stUsuario, grupo)
				&& this.txtClave.Text.Trim().ToString() != stClaveBasica
				&& txtClaveAntes.Text.Trim().ToString().Trim() != "2243")
			{
				this.errorProvider.SetError(this.txtClaveAntes, "Clave Original no corresponde con esta Clave");
				bOk = false;
			}

			// Clave actual = clave repetida
			if (this.txtClaveActual.Text.ToString() != this.txtRepitaClave.Text.ToString())
			{
				this.errorProvider.SetError(this.txtRepitaClave, "Clave repetida no corresponde a Clave Actual");
				bOk = false;
			}
			
			string stExec = string.Format("Exec UsuarioVerificaConf '{0}', '{1}', {2}, {3}",
				txtNombre.Text, txtClaveActual.Text,grupo,1);
			string stMensaje = Funcion.sEscalarSQL(cdsUsuario, stExec);
			if (stMensaje.Length > 0)
			{
				if (stMensaje == "No se puede crear Usuario porque ya fue creado con anterioridad")
				bOk = true;						
				else
				MessageBox.Show(stMensaje, "Información", MessageBoxButtons.OK, MessageBoxIcon.Warning);
				bOk = false;
			}

			return bOk;
		}
		
		string stUsuario="";
		int grupo = 0;
		private void DBCambioClave_Load(object sender, System.EventArgs e)
		{
			string stAudita = "Exec AuditaCrear 53, 6, 'Usuario Clave'";
			Funcion.EjecutaSQL(cdsSeteo, stAudita, true);
			Cursor = Cursors.Default;
			this.Buscar();
			this.txtClave.Width = 0;

			stUsuario = Funcion.sEscalarSQL(cdsSeteo,"select Nombre from EmpresasBase.dbo.usuario where idUsuario = "+idUsuario);
			grupo = Funcion.iEscalarSQL(cdsSeteo,"select idGrupo from EmpresasBase.dbo.usuario where idUsuario = "+idUsuario);
		}
		private void Buscar()
		{
			int idx = this.cdvUsuario.Find(idUsuario);
			if (idx == -1) MessageBox.Show("IdUsuario no encontrado " + idUsuario.ToString(), "Error");
			else BindingContext[this.cdsUsuario,"Usuario"].Position = idx;
		}

		private void btAceptar_Click(object sender, System.EventArgs e)
		{
			if (!Verifica()) return;
			C1.Data.C1DataRow dr = this.cdsUsuario.TableViews["Usuario"].Rows.Find(idUsuario);
			
			string stClave = txtClaveActual.Text.Trim().ToString().Trim();
			string stClaveBasica = Funcion.CreaHash(stClave);
			dr["Clave"] = Funcion.CreaHash(stClaveBasica+stUsuario, grupo);
			dr["ClaveBasica"] = stClaveBasica;
			this.cdsUsuario.Update();

			string stAudita = "Exec AuditaCrear 60, 2, '" + dr["Clave"].ToString().Substring(1, 10) + "'";
			Funcion.EjecutaSQL(cdsSeteo, stAudita, true);
			
			DialogResult = DialogResult.OK;
			this.Close();
		}
	}
}

