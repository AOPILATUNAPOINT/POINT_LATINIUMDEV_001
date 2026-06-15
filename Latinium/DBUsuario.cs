using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using DevExpress.XtraEditors;

namespace Latinium
{
	/// <summary>
	/// Summary description for DBUsuario.
	/// </summary>
	public class DBUsuario : DevExpress.XtraEditors.XtraForm
	{
		private DevExpress.XtraEditors.TextEdit txtUsuario;
		private DevExpress.XtraEditors.LabelControl labelControl1;
		private DevExpress.XtraEditors.SpinEdit spnGrupo;
		private DevExpress.XtraEditors.LabelControl labelControl2;
		private DevExpress.XtraEditors.LabelControl labelControl3;
		private DevExpress.XtraEditors.SimpleButton btAceptar;
		private DevExpress.XtraEditors.TextEdit txtClave;
		private DevExpress.XtraEditors.SimpleButton btCancelar;
		private C1.Data.C1DataSet cdsEmpresa;
		private BarraDatoSQL.BarraDatoSQL barraDato1;
		/// <summary>
		/// Required designer variable.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public DBUsuario()
		{
			//
			// Required for Windows Form Designer support
			//
			InitializeComponent();

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
			this.txtUsuario = new DevExpress.XtraEditors.TextEdit();
			this.cdsEmpresa = new C1.Data.C1DataSet();
			this.labelControl1 = new DevExpress.XtraEditors.LabelControl();
			this.spnGrupo = new DevExpress.XtraEditors.SpinEdit();
			this.labelControl2 = new DevExpress.XtraEditors.LabelControl();
			this.txtClave = new DevExpress.XtraEditors.TextEdit();
			this.labelControl3 = new DevExpress.XtraEditors.LabelControl();
			this.btAceptar = new DevExpress.XtraEditors.SimpleButton();
			this.btCancelar = new DevExpress.XtraEditors.SimpleButton();
			this.barraDato1 = new BarraDatoSQL.BarraDatoSQL();
			((System.ComponentModel.ISupportInitialize)(this.txtUsuario.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsEmpresa)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.spnGrupo.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtClave.Properties)).BeginInit();
			this.SuspendLayout();
			// 
			// txtUsuario
			// 
			this.txtUsuario.DataBindings.Add(new System.Windows.Forms.Binding("Text", this.cdsEmpresa, "Usuario.Nombre"));
			this.txtUsuario.Enabled = false;
			this.txtUsuario.Location = new System.Drawing.Point(134, 27);
			this.txtUsuario.Name = "txtUsuario";
			this.txtUsuario.Size = new System.Drawing.Size(221, 22);
			this.txtUsuario.TabIndex = 0;
			// 
			// cdsEmpresa
			// 
			this.cdsEmpresa.BindingContextCtrl = this;
			this.cdsEmpresa.DataLibrary = "LibDBEmpresa";
			this.cdsEmpresa.DataLibraryUrl = "";
			this.cdsEmpresa.DataSetDef = "dsEmpresa";
			this.cdsEmpresa.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsEmpresa.Name = "cdsEmpresa";
			this.cdsEmpresa.SchemaClassName = "LibDBEmpresa.DataClass";
			this.cdsEmpresa.SchemaDef = null;
			this.cdsEmpresa.BeforeFill += new C1.Data.FillEventHandler(this.cdsUsuario_BeforeFill);
			// 
			// labelControl1
			// 
			this.labelControl1.AutoSizeMode = DevExpress.XtraEditors.LabelAutoSizeMode.Horizontal;
			this.labelControl1.Location = new System.Drawing.Point(29, 27);
			this.labelControl1.Name = "labelControl1";
			this.labelControl1.Size = new System.Drawing.Size(43, 16);
			this.labelControl1.TabIndex = 1;
			this.labelControl1.Text = "Usuario";
			// 
			// spnGrupo
			// 
			this.spnGrupo.DataBindings.Add(new System.Windows.Forms.Binding("EditValue", this.cdsEmpresa, "Usuario.idGrupo"));
			this.spnGrupo.EditValue = new System.Decimal(new int[] {
																															 1,
																															 0,
																															 0,
																															 0});
			this.spnGrupo.Enabled = false;
			this.spnGrupo.Location = new System.Drawing.Point(134, 101);
			this.spnGrupo.Name = "spnGrupo";
			// 
			// spnGrupo.Properties
			// 
			this.spnGrupo.Properties.Buttons.AddRange(new DevExpress.XtraEditors.Controls.EditorButton[] {
																																																		 new DevExpress.XtraEditors.Controls.EditorButton()});
			this.spnGrupo.Properties.IsFloatValue = false;
			this.spnGrupo.Properties.Mask.EditMask = "N00";
			this.spnGrupo.Properties.MaxValue = new System.Decimal(new int[] {
																																				 20,
																																				 0,
																																				 0,
																																				 0});
			this.spnGrupo.Properties.MinValue = new System.Decimal(new int[] {
																																				 1,
																																				 0,
																																				 0,
																																				 0});
			this.spnGrupo.Size = new System.Drawing.Size(68, 22);
			this.spnGrupo.TabIndex = 2;
			// 
			// labelControl2
			// 
			this.labelControl2.AutoSizeMode = DevExpress.XtraEditors.LabelAutoSizeMode.Horizontal;
			this.labelControl2.Location = new System.Drawing.Point(29, 101);
			this.labelControl2.Name = "labelControl2";
			this.labelControl2.Size = new System.Drawing.Size(34, 16);
			this.labelControl2.TabIndex = 3;
			this.labelControl2.Text = "Grupo";
			// 
			// txtClave
			// 
			this.txtClave.DataBindings.Add(new System.Windows.Forms.Binding("Text", this.cdsEmpresa, "Usuario.Clave"));
			this.txtClave.Enabled = false;
			this.txtClave.Location = new System.Drawing.Point(134, 64);
			this.txtClave.Name = "txtClave";
			// 
			// txtClave.Properties
			// 
			this.txtClave.Properties.PasswordChar = '*';
			this.txtClave.Size = new System.Drawing.Size(221, 22);
			this.txtClave.TabIndex = 4;
			// 
			// labelControl3
			// 
			this.labelControl3.AutoSizeMode = DevExpress.XtraEditors.LabelAutoSizeMode.Horizontal;
			this.labelControl3.Location = new System.Drawing.Point(29, 64);
			this.labelControl3.Name = "labelControl3";
			this.labelControl3.Size = new System.Drawing.Size(65, 16);
			this.labelControl3.TabIndex = 5;
			this.labelControl3.Text = "Contraseña";
			// 
			// btAceptar
			// 
			this.btAceptar.Location = new System.Drawing.Point(67, 0);
			this.btAceptar.Name = "btAceptar";
			this.btAceptar.Size = new System.Drawing.Size(87, 26);
			this.btAceptar.TabIndex = 6;
			this.btAceptar.Text = "&Aceptar";
			this.btAceptar.Visible = false;
			this.btAceptar.Click += new System.EventHandler(this.btAceptar_Click);
			// 
			// btCancelar
			// 
			this.btCancelar.DialogResult = System.Windows.Forms.DialogResult.Cancel;
			this.btCancelar.Location = new System.Drawing.Point(278, 0);
			this.btCancelar.Name = "btCancelar";
			this.btCancelar.Size = new System.Drawing.Size(87, 26);
			this.btCancelar.TabIndex = 7;
			this.btCancelar.Text = "&Cancelar";
			this.btCancelar.Visible = false;
			this.btCancelar.Click += new System.EventHandler(this.btCancelar_Click);
			// 
			// barraDato1
			// 
			this.barraDato1.BarraMovimiento = true;
			this.barraDato1.DataMember = "Usuario";
			this.barraDato1.DataNombreId = "idUsuario";
			this.barraDato1.DataOrden = "";
			this.barraDato1.DataSource = this.cdsEmpresa;
			this.barraDato1.DataTabla = "Usuario";
			this.barraDato1.DataTablaHija = "";
			this.barraDato1.Location = new System.Drawing.Point(19, 146);
			this.barraDato1.Name = "barraDato1";
			this.barraDato1.Size = new System.Drawing.Size(365, 30);
			this.barraDato1.TabIndex = 78;
			this.barraDato1.VisibleBorrar = true;
			this.barraDato1.VisibleBuscar = false;
			this.barraDato1.VisibleEditar = true;
			this.barraDato1.VisibleImprimir = false;
			this.barraDato1.VisibleNuevo = true;
			this.barraDato1.Borrar += new System.EventHandler(this.barraDato1_Borrar);
			this.barraDato1.Restaurar += new System.EventHandler(this.barraDato1_Restaurar);
			this.barraDato1.Refresca += new System.EventHandler(this.barraDato1_Refresca);
			this.barraDato1.Crear += new System.EventHandler(this.barraDato1_Crear);
			this.barraDato1.Deshacer += new System.EventHandler(this.barraDato1_Deshacer);
			this.barraDato1.Copiar += new System.EventHandler(this.barraDato1_Copiar);
			this.barraDato1.Editar += new System.EventHandler(this.barraDato1_Editar);
			this.barraDato1.Grabar += new System.EventHandler(this.barraDato1_Grabar);
			// 
			// DBUsuario
			// 
			this.AcceptButton = this.btAceptar;
			this.AutoScaleBaseSize = new System.Drawing.Size(6, 16);
			this.CancelButton = this.btCancelar;
			this.ClientSize = new System.Drawing.Size(412, 180);
			this.Controls.Add(this.barraDato1);
			this.Controls.Add(this.btCancelar);
			this.Controls.Add(this.btAceptar);
			this.Controls.Add(this.labelControl3);
			this.Controls.Add(this.txtClave);
			this.Controls.Add(this.labelControl2);
			this.Controls.Add(this.spnGrupo);
			this.Controls.Add(this.labelControl1);
			this.Controls.Add(this.txtUsuario);
			this.MaximizeBox = false;
			this.Name = "DBUsuario";
			this.Text = "Crear Usuarios";
			this.Load += new System.EventHandler(this.DBUsuario_Load);
			((System.ComponentModel.ISupportInitialize)(this.txtUsuario.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsEmpresa)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.spnGrupo.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtClave.Properties)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void DBUsuario_Load(object sender, System.EventArgs e)
		{
			barraDato1.PosUltima();
		}

		private void btAceptar_Click(object sender, System.EventArgs e)
		{			
			string stExec = string.Format("Exec UsuarioVerificaConf '{0}', '{1}', {2}",
				txtUsuario.Text, txtClave.Text, spnGrupo.Value.ToString());
			string stMensaje = Funcion.sEscalarSQL(cdsEmpresa, stExec);
			if (stMensaje.Length > 0)
			{
				MessageBox.Show(stMensaje, "Información", MessageBoxButtons.OK, MessageBoxIcon.Warning);

				if (stMensaje == "No se puede crear Usuario porque ya fue creado con anterioridad")
					if (DialogResult.Yes == MessageBox.Show("Desea Modificar Usuario o Clave? ", "Confirmación",
						MessageBoxButtons.YesNo, MessageBoxIcon.Question)) 
					{
						grabar();
						return;
					}
				return;
			}
			grabar();			
		}
		private void grabar()
		{
			string stUsuario = txtUsuario.Text.Trim();
			string stClave = txtClave.Text.ToString().Trim();
			string stClaveBasica = Funcion.CreaHash(stClave);
			string stClaveHash = Funcion.CreaHash(stClaveBasica+stUsuario, int.Parse(spnGrupo.EditValue.ToString()));

			string stExec = string.Format("Exec UsuarioCrear '{0}', '{1}', {2}, '{3}'"
				, txtUsuario.Text, stClaveHash, spnGrupo.Value, stClaveBasica);
			Funcion.EjecutaSQL(cdsEmpresa, stExec, true);
			if(editar == 0)
			MessageBox.Show("Nuevo Usuario Creado en la Base de Datos", "Información",
				MessageBoxButtons.OK, MessageBoxIcon.Information);
		}
		private void cdsUsuario_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsEmpresa.Schema.Connections[0].ConnectionString = miFuncion.AccesoEmpresaBase();
		}
		private void btCancelar_Click(object sender, System.EventArgs e)
		{
			Close();
		}

		private void barraDato1_Refresca(object sender, System.EventArgs e)
		{
		
		}

		private void barraDato1_Restaurar(object sender, System.EventArgs e)
		{
		
		}

		int editar = 0;
		private void barraDato1_Crear(object sender, System.EventArgs e)
		{
			barraDato1.CrearRegistro();
		}

		private void barraDato1_Copiar(object sender, System.EventArgs e)
		{
		
		}

		private void barraDato1_Editar(object sender, System.EventArgs e)
		{
			barraDato1.EditarRegistro();
			editar = 1;
		}

		private void barraDato1_Grabar(object sender, System.EventArgs e)
		{	
			if(editar == 0)
			{
				string stExec = string.Format("Exec UsuarioVerificaConf '{0}', '{1}', {2}",
					txtUsuario.Text, txtClave.Text, spnGrupo.Value);
				string stMensaje = Funcion.sEscalarSQL(cdsEmpresa, stExec);
				if (stMensaje.Length > 0)
				{
					  MessageBox.Show(stMensaje, "Información", MessageBoxButtons.OK, MessageBoxIcon.Warning);				
						return;
				}
			}
			if(editar == 1)
			{
				string stExec = string.Format("Exec UsuarioVerificaConf '{0}', '{1}', {2},{3}",
					txtUsuario.Text, txtClave.Text, spnGrupo.Value,1);
				string stMensaje = Funcion.sEscalarSQL(cdsEmpresa, stExec);
				if (stMensaje.Length > 0)
				{
					MessageBox.Show(stMensaje, "Información", MessageBoxButtons.OK, MessageBoxIcon.Warning);				
					return;
				}
			}
			grabar();		
			barraDato1.DeshacerRegistro();
			barraDato1.ProximoId(5);
			editar = 0;
		}

		private void barraDato1_Borrar(object sender, System.EventArgs e)
		{
			if (MessageBox.Show("¿Desea Borrar este Usuario?", 
				"Confirmación", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) == DialogResult.No) 
				return;
			barraDato1.BorraRegistro();
				editar = 0;
		}

		private void barraDato1_Deshacer(object sender, System.EventArgs e)
		{
			barraDato1.DeshacerRegistro();
			editar = 0;
		}
		
	}
}

