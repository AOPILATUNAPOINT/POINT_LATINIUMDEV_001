using System;
using System.Data.OleDb;
using System.Data.SqlClient;
using System.Data;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de Usuario.
	/// </summary>
	public class Usuario : DevExpress.XtraEditors.XtraForm
	{
		private System.Windows.Forms.Label label1;
		private C1.Data.C1DataSet cdsUsuario;
		private System.Windows.Forms.Label label2;
		private System.Windows.Forms.Label label3;
		private System.Windows.Forms.Label label4;
		private System.Windows.Forms.ErrorProvider errorProvider;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtNombre;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtClave;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtInicial;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor spnGrupo;
		private C1.Data.C1DataSet cdsSeteo;
		private Infragistics.Win.UltraWinGrid.UltraGrid ultraGrid1;
		private Infragistics.Win.Misc.UltraButton btAceptar;
		private Infragistics.Win.Misc.UltraButton btSalir;
		private C1.Data.C1DataSet cdsSeguridad;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public Usuario()
		{
			//
			// Necesario para admitir el Diseñador de Windows Forms
			//
			InitializeComponent();

			//
			// TODO: Agregar código de constructor después de llamar a InitializeComponent
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

		#region Windows Form Designer generated code
		/// <summary>
		/// Método necesario para admitir el Diseñador, no se puede modificar
		/// el contenido del método con el editor de código.
		/// </summary>
		private void InitializeComponent()
		{
			Infragistics.Win.Appearance appearance1 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Usuario", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Clave");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idGrupo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idUsuario");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Iniciales");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Sexo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CodArt", 0);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("unidadBotrosa", 1);
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			this.cdsUsuario = new C1.Data.C1DataSet();
			this.label1 = new System.Windows.Forms.Label();
			this.label2 = new System.Windows.Forms.Label();
			this.label3 = new System.Windows.Forms.Label();
			this.label4 = new System.Windows.Forms.Label();
			this.errorProvider = new System.Windows.Forms.ErrorProvider();
			this.txtNombre = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtClave = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtInicial = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.spnGrupo = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.cdsSeteo = new C1.Data.C1DataSet();
			this.ultraGrid1 = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.btAceptar = new Infragistics.Win.Misc.UltraButton();
			this.btSalir = new Infragistics.Win.Misc.UltraButton();
			this.cdsSeguridad = new C1.Data.C1DataSet();
			((System.ComponentModel.ISupportInitialize)(this.cdsUsuario)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNombre)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtClave)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtInicial)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.spnGrupo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeguridad)).BeginInit();
			this.SuspendLayout();
			// 
			// cdsUsuario
			// 
			this.cdsUsuario.BindingContextCtrl = this;
			this.cdsUsuario.CaseSensitive = false;
			this.cdsUsuario.DataLibrary = "LibEmpresa";
			this.cdsUsuario.DataLibraryUrl = "";
			this.cdsUsuario.DataSetDef = "dsUsuario";
			this.cdsUsuario.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsUsuario.Name = "cdsUsuario";
			this.cdsUsuario.SchemaClassName = "LibEmpresa.DataClass";
			this.cdsUsuario.SchemaDef = null;
			this.cdsUsuario.PositionChanged += new C1.Data.PositionChangeEventHandler(this.cdsUsuario_PositionChanged);
			this.cdsUsuario.BeforeFill += new C1.Data.FillEventHandler(this.cdsUsuario_BeforeFill);
			// 
			// label1
			// 
			this.label1.AutoSize = true;
			this.label1.Location = new System.Drawing.Point(8, 32);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(48, 16);
			this.label1.TabIndex = 2;
			this.label1.Text = "Nombre:";
			// 
			// label2
			// 
			this.label2.AutoSize = true;
			this.label2.Location = new System.Drawing.Point(8, 58);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(65, 16);
			this.label2.TabIndex = 4;
			this.label2.Text = "Contraseña:";
			// 
			// label3
			// 
			this.label3.AutoSize = true;
			this.label3.Location = new System.Drawing.Point(8, 7);
			this.label3.Name = "label3";
			this.label3.Size = new System.Drawing.Size(49, 16);
			this.label3.TabIndex = 6;
			this.label3.Text = "Iniciales:";
			this.label3.Visible = false;
			// 
			// label4
			// 
			this.label4.AutoSize = true;
			this.label4.Location = new System.Drawing.Point(8, 84);
			this.label4.Name = "label4";
			this.label4.Size = new System.Drawing.Size(39, 16);
			this.label4.TabIndex = 7;
			this.label4.Text = "Grupo:";
			// 
			// errorProvider
			// 
			this.errorProvider.ContainerControl = this;
			this.errorProvider.DataMember = "";
			// 
			// txtNombre
			// 
			this.txtNombre.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsUsuario, "Usuario.Nombre"));
			this.txtNombre.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNombre.Location = new System.Drawing.Point(96, 32);
			this.txtNombre.Name = "txtNombre";
			this.txtNombre.Size = new System.Drawing.Size(208, 22);
			this.txtNombre.TabIndex = 9;
			this.txtNombre.Validated += new System.EventHandler(this.txtNombre_Validated);
			// 
			// txtClave
			// 
			this.txtClave.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsUsuario, "Usuario.Clave"));
			this.txtClave.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtClave.Location = new System.Drawing.Point(96, 58);
			this.txtClave.Name = "txtClave";
			this.txtClave.PasswordChar = '*';
			this.txtClave.Size = new System.Drawing.Size(208, 22);
			this.txtClave.TabIndex = 10;
			// 
			// txtInicial
			// 
			this.txtInicial.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsUsuario, "Usuario.Iniciales"));
			this.txtInicial.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtInicial.Location = new System.Drawing.Point(96, 7);
			this.txtInicial.Name = "txtInicial";
			this.txtInicial.Size = new System.Drawing.Size(208, 22);
			this.txtInicial.TabIndex = 11;
			// 
			// spnGrupo
			// 
			this.spnGrupo.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsUsuario, "Usuario.idGrupo"));
			this.spnGrupo.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.spnGrupo.Location = new System.Drawing.Point(96, 84);
			this.spnGrupo.MaskInput = "nn";
			this.spnGrupo.MaxValue = 20;
			this.spnGrupo.MinValue = 1;
			this.spnGrupo.Name = "spnGrupo";
			this.spnGrupo.PromptChar = ' ';
			this.spnGrupo.Size = new System.Drawing.Size(88, 22);
			this.spnGrupo.SpinButtonDisplayStyle = Infragistics.Win.ButtonDisplayStyle.Always;
			this.spnGrupo.TabIndex = 12;
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
			// ultraGrid1
			// 
			this.ultraGrid1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.ultraGrid1.Cursor = System.Windows.Forms.Cursors.Default;
			this.ultraGrid1.DataMember = "Usuario";
			this.ultraGrid1.DataSource = this.cdsUsuario;
			appearance1.BackColor = System.Drawing.Color.White;
			this.ultraGrid1.DisplayLayout.Appearance = appearance1;
			this.ultraGrid1.DisplayLayout.AutoFitColumns = true;
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridColumn2.Hidden = true;
			ultraGridColumn3.Header.VisiblePosition = 2;
			ultraGridColumn3.Hidden = true;
			ultraGridColumn4.Header.VisiblePosition = 3;
			ultraGridColumn4.Hidden = true;
			ultraGridColumn5.Header.VisiblePosition = 4;
			ultraGridColumn5.Width = 193;
			ultraGridColumn6.Header.VisiblePosition = 6;
			ultraGridColumn6.Hidden = true;
			ultraGridColumn7.Header.Caption = "Código";
			ultraGridColumn7.Header.VisiblePosition = 5;
			ultraGridColumn7.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownList;
			ultraGridColumn7.Width = 82;
			ultraGridColumn8.Header.Caption = "Unidades";
			ultraGridColumn8.Header.VisiblePosition = 7;
			ultraGridColumn8.Hidden = true;
			ultraGridColumn8.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.Edit;
			ultraGridBand1.Columns.AddRange(new object[] {
																										 ultraGridColumn1,
																										 ultraGridColumn2,
																										 ultraGridColumn3,
																										 ultraGridColumn4,
																										 ultraGridColumn5,
																										 ultraGridColumn6,
																										 ultraGridColumn7,
																										 ultraGridColumn8});
			this.ultraGrid1.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.ultraGrid1.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.TemplateOnBottom;
			this.ultraGrid1.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.True;
			this.ultraGrid1.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance2.BackColor = System.Drawing.Color.Transparent;
			this.ultraGrid1.DisplayLayout.Override.CardAreaAppearance = appearance2;
			appearance3.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance3.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance3.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance3.FontData.BoldAsString = "True";
			appearance3.FontData.Name = "Arial";
			appearance3.FontData.SizeInPoints = 10F;
			appearance3.ForeColor = System.Drawing.Color.White;
			appearance3.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.ultraGrid1.DisplayLayout.Override.HeaderAppearance = appearance3;
			this.ultraGrid1.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortSingle;
			appearance4.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance4.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance4.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.RowAlternateAppearance = appearance4;
			appearance5.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance5.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance5.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.RowSelectorAppearance = appearance5;
			appearance6.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance6.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance6.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.SelectedRowAppearance = appearance6;
			this.ultraGrid1.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.ultraGrid1.Location = new System.Drawing.Point(8, 118);
			this.ultraGrid1.Name = "ultraGrid1";
			this.ultraGrid1.Size = new System.Drawing.Size(296, 200);
			this.ultraGrid1.TabIndex = 14;
			this.ultraGrid1.InitializeRow += new Infragistics.Win.UltraWinGrid.InitializeRowEventHandler(this.ultraGrid1_InitializeRow);
			this.ultraGrid1.BeforeRowsDeleted += new Infragistics.Win.UltraWinGrid.BeforeRowsDeletedEventHandler(this.ultraGrid1_BeforeRowsDeleted);
			// 
			// btAceptar
			// 
			this.btAceptar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			appearance7.Image = 9;
			this.btAceptar.Appearance = appearance7;
			this.btAceptar.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.btAceptar.Location = new System.Drawing.Point(8, 326);
			this.btAceptar.Name = "btAceptar";
			this.btAceptar.Size = new System.Drawing.Size(80, 25);
			this.btAceptar.TabIndex = 17;
			this.btAceptar.Text = "&Grabar";
			this.btAceptar.Click += new System.EventHandler(this.btAceptar_Click);
			this.btAceptar.MouseUp += new System.Windows.Forms.MouseEventHandler(this.btAceptar_MouseUp);
			// 
			// btSalir
			// 
			this.btSalir.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			appearance8.Image = 9;
			this.btSalir.Appearance = appearance8;
			this.btSalir.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.btSalir.DialogResult = System.Windows.Forms.DialogResult.Cancel;
			this.btSalir.Location = new System.Drawing.Point(104, 326);
			this.btSalir.Name = "btSalir";
			this.btSalir.Size = new System.Drawing.Size(80, 25);
			this.btSalir.TabIndex = 18;
			this.btSalir.Text = "&Salir";
			this.btSalir.Click += new System.EventHandler(this.btSalir_Click);
			// 
			// cdsSeguridad
			// 
			this.cdsSeguridad.BindingContextCtrl = this;
			this.cdsSeguridad.DataLibrary = "LibContabilidad";
			this.cdsSeguridad.DataLibraryUrl = "";
			this.cdsSeguridad.DataSetDef = "dsSeguridad";
			this.cdsSeguridad.FillOnRequest = false;
			this.cdsSeguridad.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsSeguridad.Name = "cdsSeguridad";
			this.cdsSeguridad.SchemaClassName = "LibContabilidad.DataClass";
			this.cdsSeguridad.SchemaDef = null;
			// 
			// Usuario
			// 
			this.AcceptButton = this.btAceptar;
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 14);
			this.CancelButton = this.btSalir;
			this.ClientSize = new System.Drawing.Size(312, 358);
			this.Controls.Add(this.btSalir);
			this.Controls.Add(this.btAceptar);
			this.Controls.Add(this.ultraGrid1);
			this.Controls.Add(this.spnGrupo);
			this.Controls.Add(this.txtInicial);
			this.Controls.Add(this.txtClave);
			this.Controls.Add(this.txtNombre);
			this.Controls.Add(this.label4);
			this.Controls.Add(this.label3);
			this.Controls.Add(this.label2);
			this.Controls.Add(this.label1);
			this.MaximizeBox = false;
			this.Name = "Usuario";
			this.Text = "Usuario";
			this.Closing += new System.ComponentModel.CancelEventHandler(this.Usuario_Closing);
			this.Load += new System.EventHandler(this.Usuario_Load);
			((System.ComponentModel.ISupportInitialize)(this.cdsUsuario)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNombre)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtClave)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtInicial)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.spnGrupo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeguridad)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private Acceso miAcceso;
		private void Usuario_Load(object sender, System.EventArgs e)
		{
			string stAudita = "Exec AuditaCrear 53, 6, 'Usuario'";
			Funcion.EjecutaSQL(cdsSeteo, stAudita, true);
			Cursor = Cursors.Default;
			miAcceso = new Acceso(cdsSeteo, "050A2");
			if (!miAcceso.CambiarGrupo) spnGrupo.ReadOnly = true;
			txtInicial.Width = 0;
			if (!miAcceso.Editar) btAceptar.Enabled = false;
		}

		private bool Verifica()
		{
			bool bOk = true;
			if (this.txtNombre.Text.ToString().Trim().Length == 0)
			{
				bOk = false;
				this.errorProvider.SetError(this.txtNombre, "Ingrese nombre de usuario");
			}
			else
			{
				string stNombre = txtNombre.Text.ToString().Trim();
				int iNumUsuarios = 0;
//				if (!barraDatoSQL1.bNuevo) iNumUsuarios = 1;
				
				int iCantidad = 0;
				this.cdsUsuario.Schema.Connections[0].Open();
				C1.Data.SchemaObjects.Connection miConeccion = this.cdsUsuario.Schema.Connections[0];
				OleDbCommand cmdSaldo = new OleDbCommand ("SELECT COUNT(*) FROM Usuario WHERE UPPER(Nombre) = '" + stNombre + "'", (OleDbConnection) miConeccion.DbConnection);

				try
				{
					iCantidad = (int) cmdSaldo.ExecuteScalar();
				}
				catch(System.Exception ex)
				{
					Console.WriteLine(ex.Message + " Funcion Verifica en usuario");
				}
				this.cdsUsuario.Schema.Connections[0].Close();

				//				DataRow[] dr = cdsUsuario.StorageDataSet.Tables["Usuario"].Select(stFiltro);
				if (iCantidad > iNumUsuarios)
				{
					bOk = false;
					this.errorProvider.SetError(this.txtNombre, "Usuario ya existe");
				}
				else
					this.errorProvider.SetError(this.txtNombre, "");
			}

			int iGrupo = 0;
			try
			{
				iGrupo = int.Parse(this.spnGrupo.Value.ToString());
			}
			catch
			{
				bOk = false;
				this.errorProvider.SetError(this.spnGrupo, "Error en Grupo");
			}

			if (iGrupo < 1 || iGrupo > 19 )
			{
				bOk = false;
				this.errorProvider.SetError(this.spnGrupo, "Grupo valido entre 1 y 19");
			}
			else
				this.errorProvider.SetError(this.spnGrupo, "");

			return bOk;
		}

		private void cdsUsuario_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsUsuario.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirInicio + "\\Empresa.Mdb", 1);
		}

		C1.Data.C1DataRow drUsuario;
		string stClaveAntes = "!@$#%#$";
		private void cdsUsuario_PositionChanged(object sender, C1.Data.PositionChangeEventArgs e)
		{
			drUsuario = e.Row;
			stClaveAntes = drUsuario["Clave"].ToString();
		}

		private void Usuario_Closing(object sender, System.ComponentModel.CancelEventArgs e)
		{
			string stAudita = "Exec AuditaCrear 53, 7, 'Usuario'";
			Funcion.EjecutaSQL(cdsSeteo, stAudita, true);
		}

		private void cdsSeteo_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteo.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirConta);
		}

		private void btAceptar_Click(object sender, System.EventArgs e)
		{
			//if (!Verifica()) return;

			try
			{
				string stClave = txtClave.Text.ToString().Trim();
				string stClaveHash = stClave; 
				/*Funcion.CreaHash(stClave);
				if (stClaveAntes != stClave)
				{
					drUsuario["Clave"] = stClaveHash;
					txtClave.Value = stClaveHash;
					txtClave.Select();
					txtClave.Refresh();
				}
				string stInicial = Funcion.CreaHash(spnGrupo.Value.ToString() + "2243");
				txtInicial.Value = stInicial;
				txtInicial.Select();
				txtInicial.Refresh();*/
			
				Validate();
				btAceptar.Select();
				cdsUsuario.Update();
				int iEditar = 2;
//				if (bNuevo) iEditar = 1;
				#region Crea usuario en la base de datos
				stClave = txtClave.Text.ToString().Trim() + spnGrupo.Value.ToString();
//				stClaveHash = Funcion.CreaHash(stClave);*** 

				string stExec = string.Format("Exec UsuarioCrear '{0}', '{1}', {2}, '{3}'", txtNombre.Text, stClaveHash, spnGrupo.Value, txtInicial.Text);
				Funcion.EjecutaSQL(cdsSeteo, stExec, true);
				//MessageBox.Show("Usuario Creado en Base de Datos", "Información");
				#endregion Crea usuario en la base de datos

				#region Crea usuario en el servidor
				string stSelect = string.Format("SELECT Count(*) FROM Master.Dbo.syslogins Where LoginName = '{0}'", this.txtNombre.Text.ToString());
				int iCuentaBases = Funcion.iEscalarSQL(cdsSeguridad, stSelect, true);
				if (iCuentaBases == 0)
				{
					string sSQL = string.Format("Create Login {0} With Password = '{1}'; Exec sp_addsrvrolemember '{0}', 'sysadmin';" 
						+ "Create User {0} For Login {0}", this.txtNombre.Text.ToString(), "Bl45o6$9");
					Funcion.EjecutaSQL(cdsSeguridad, sSQL);
				}
				MessageBox.Show(string.Format("Usuario {0} creado", this.txtNombre.Text.ToString()), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
				#endregion Crea usuario en el servidor
				
        string stAudita = "Exec AuditaCrear 60, " + iEditar.ToString() + ", '" + this.txtNombre.Text.ToString() + "'"; // txtClave.Text.Substring(1, 10) + "'";
				Funcion.EjecutaSQL(cdsSeteo, stAudita, true);
				stAudita = "Exec AuditaCrear 60, " + iEditar.ToString() + ", '" + txtNombre.Text + "'";
				Funcion.EjecutaSQL(cdsSeteo, stAudita, true);
				stAudita = "Exec AuditaCrear 60, " + iEditar.ToString() + ", '" + spnGrupo.Value.ToString() + "'";
				Funcion.EjecutaSQL(cdsSeteo, stAudita, true);
			}
			catch(System.Exception ex)
			{
				MessageBox.Show(ex.Message, "Grabar Usuario");
			}
		}

		private void btAceptar_MouseUp(object sender, System.Windows.Forms.MouseEventArgs e)
		{
			return;
			if (e.Button != MouseButtons.Right) return;
			string stClave = txtClave.Text.ToString().Trim() + spnGrupo.Value.ToString();
			string stClaveHash = Funcion.CreaHash(stClave);

			string stExec = string.Format("Exec UsuarioCrear '{0}', '{1}', {2}, '{3}'"
				, txtNombre.Text, stClaveHash, spnGrupo.Value, txtInicial.Text);
			Funcion.EjecutaSQL(cdsSeteo, stExec, true);
			MessageBox.Show("Usuario Creado en Base de Datos", "Información");
		}

		private void ultraGrid1_BeforeRowsDeleted(object sender, Infragistics.Win.UltraWinGrid.BeforeRowsDeletedEventArgs e)
		{
			if (!miAcceso.Eliminar)
			{
				MessageBox.Show("No tiene acceso a Eliminar Usuarios",
					"Información", MessageBoxButtons.OK, MessageBoxIcon.Information);
				return;
			}

			if (DialogResult.No == MessageBox.Show("¿Desea Borrar Usuario Seleccionado?",
				"Confirmación", MessageBoxButtons.YesNo, MessageBoxIcon.Question))
				e.Cancel = true;

			if (txtNombre.Text.ToUpper().Trim() == "ADMINISTRADOR")
			{
				MessageBox.Show("No se puede Eliminar Usuario Administrador",
					"Información", MessageBoxButtons.OK, MessageBoxIcon.Information);
				return;
			}

			string stAudita = "Exec AuditaCrear 60, 4, '" + txtClave.Text.Substring(1, 10) + "'";
			Funcion.EjecutaSQL(cdsSeteo, stAudita, true);
			stAudita = "Exec AuditaCrear 60, 4, '" + txtNombre.Text + "'";
			Funcion.EjecutaSQL(cdsSeteo, stAudita, true);
			stAudita = "Exec AuditaCrear 60, 4, '" + spnGrupo.Value.ToString() + "'";
			Funcion.EjecutaSQL(cdsSeteo, stAudita, true);
		}

		private void txtNombre_Validated(object sender, System.EventArgs e)
		{
		
		}

		private void btSalir_Click(object sender, System.EventArgs e)
		{
			Close();
		}

		private void ultraGrid1_InitializeRow(object sender, Infragistics.Win.UltraWinGrid.InitializeRowEventArgs e)
		{
//			drUsuario["idGrupo"] = 1;
		}
	}
}
