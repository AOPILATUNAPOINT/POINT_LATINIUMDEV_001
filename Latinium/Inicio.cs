using System;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.IO;
using System.Drawing.Printing;
using System.IO.IsolatedStorage;
using System.Net.Sockets;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de Inicio.
	/// </summary>
	public class Inicio : DevExpress.XtraEditors.XtraForm
	{
		private System.Windows.Forms.Label label2;
		private System.Windows.Forms.HelpProvider helpProvider;
		private System.Windows.Forms.Label label3;
		public Infragistics.Win.UltraWinEditors.UltraTextEditor txtClave;
		private Infragistics.Win.Misc.UltraButton btAceptar;
		private Infragistics.Win.Misc.UltraButton btCancelar;
		private System.Windows.Forms.Label label4;
		public Infragistics.Win.UltraWinEditors.UltraTextEditor txtNombre;
		private System.Windows.Forms.GroupBox groupBox1;
		private System.Windows.Forms.PictureBox pictureBox2;
		private System.Windows.Forms.PictureBox pictureBox3;
		private System.Windows.Forms.Label label5;
		private System.Windows.Forms.Label label6;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbEmpresas;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource1;
		private C1.Data.C1DataSet cdsSeguridad;
		private System.Windows.Forms.Label label1;
		private System.ComponentModel.IContainer components=null;

		public Inicio()
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
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idEmpresa");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Empresa");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idEmpresa");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Empresa");
			this.label2 = new System.Windows.Forms.Label();
			this.helpProvider = new System.Windows.Forms.HelpProvider();
			this.btCancelar = new Infragistics.Win.Misc.UltraButton();
			this.label3 = new System.Windows.Forms.Label();
			this.txtClave = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.btAceptar = new Infragistics.Win.Misc.UltraButton();
			this.txtNombre = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label4 = new System.Windows.Forms.Label();
			this.groupBox1 = new System.Windows.Forms.GroupBox();
			this.label6 = new System.Windows.Forms.Label();
			this.label5 = new System.Windows.Forms.Label();
			this.pictureBox3 = new System.Windows.Forms.PictureBox();
			this.pictureBox2 = new System.Windows.Forms.PictureBox();
			this.cmbEmpresas = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.ultraDataSource1 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.cdsSeguridad = new C1.Data.C1DataSet();
			this.label1 = new System.Windows.Forms.Label();
			((System.ComponentModel.ISupportInitialize)(this.txtClave)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNombre)).BeginInit();
			this.groupBox1.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.cmbEmpresas)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeguridad)).BeginInit();
			this.SuspendLayout();
			// 
			// label2
			// 
			this.label2.AutoSize = true;
			this.label2.BackColor = System.Drawing.Color.Transparent;
			this.label2.Location = new System.Drawing.Point(208, 112);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(61, 16);
			this.label2.TabIndex = 3;
			this.label2.Text = "Contraseña";
			this.label2.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// helpProvider
			// 
			this.helpProvider.HelpNamespace = "C:\\Latinium\\Latinium.chm";
			// 
			// btCancelar
			// 
			this.btCancelar.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.btCancelar.CausesValidation = false;
			this.btCancelar.DialogResult = System.Windows.Forms.DialogResult.Cancel;
			this.helpProvider.SetHelpString(this.btCancelar, "Cierra el Programa");
			this.btCancelar.Location = new System.Drawing.Point(368, 168);
			this.btCancelar.Name = "btCancelar";
			this.helpProvider.SetShowHelp(this.btCancelar, true);
			this.btCancelar.Size = new System.Drawing.Size(88, 25);
			this.btCancelar.TabIndex = 3;
			this.btCancelar.Text = "&Cancelar";
			this.btCancelar.Click += new System.EventHandler(this.btCancelar_Click);
			this.btCancelar.MouseUp += new System.Windows.Forms.MouseEventHandler(this.btCancelar_MouseDown);
			// 
			// label3
			// 
			this.label3.AutoSize = true;
			this.label3.BackColor = System.Drawing.Color.Transparent;
			this.label3.Location = new System.Drawing.Point(208, 136);
			this.label3.Name = "label3";
			this.label3.Size = new System.Drawing.Size(47, 16);
			this.label3.TabIndex = 7;
			this.label3.Text = "Empresa";
			this.label3.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.label3.Visible = false;
			// 
			// txtClave
			// 
			this.txtClave.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtClave.Location = new System.Drawing.Point(280, 112);
			this.txtClave.MaxLength = 12;
			this.txtClave.Name = "txtClave";
			this.txtClave.PasswordChar = '*';
			this.txtClave.Size = new System.Drawing.Size(184, 22);
			this.txtClave.TabIndex = 1;
			this.txtClave.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtClave_KeyPress);
			this.txtClave.Validating += new System.ComponentModel.CancelEventHandler(this.txtClave_Validating);
			this.txtClave.Validated += new System.EventHandler(this.txtClave_Validated);
			this.txtClave.ValueChanged += new System.EventHandler(this.txtClave_ValueChanged);
			// 
			// btAceptar
			// 
			this.btAceptar.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.btAceptar.Location = new System.Drawing.Point(280, 168);
			this.btAceptar.Name = "btAceptar";
			this.btAceptar.Size = new System.Drawing.Size(89, 25);
			this.btAceptar.TabIndex = 2;
			this.btAceptar.Text = "&Aceptar";
			this.btAceptar.Click += new System.EventHandler(this.btAceptar_Click);
			this.btAceptar.MouseUp += new System.Windows.Forms.MouseEventHandler(this.btAceptar_MouseUp);
			// 
			// txtNombre
			// 
			this.txtNombre.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtNombre.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNombre.Location = new System.Drawing.Point(280, 88);
			this.txtNombre.MaxLength = 20;
			this.txtNombre.Name = "txtNombre";
			this.txtNombre.Size = new System.Drawing.Size(184, 22);
			this.txtNombre.TabIndex = 0;
			this.txtNombre.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtNombre_KeyPress);
			this.txtNombre.ValueChanged += new System.EventHandler(this.txtNombre_ValueChanged);
			// 
			// label4
			// 
			this.label4.AutoSize = true;
			this.label4.BackColor = System.Drawing.Color.Transparent;
			this.label4.Location = new System.Drawing.Point(208, 88);
			this.label4.Name = "label4";
			this.label4.Size = new System.Drawing.Size(44, 16);
			this.label4.TabIndex = 60;
			this.label4.Text = "Nombre";
			this.label4.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// groupBox1
			// 
			this.groupBox1.BackColor = System.Drawing.Color.White;
			this.groupBox1.Controls.Add(this.label6);
			this.groupBox1.Controls.Add(this.label5);
			this.groupBox1.Controls.Add(this.pictureBox3);
			this.groupBox1.Location = new System.Drawing.Point(0, -6);
			this.groupBox1.Name = "groupBox1";
			this.groupBox1.Size = new System.Drawing.Size(528, 56);
			this.groupBox1.TabIndex = 62;
			this.groupBox1.TabStop = false;
			this.groupBox1.Enter += new System.EventHandler(this.groupBox1_Enter);
			// 
			// label6
			// 
			this.label6.Location = new System.Drawing.Point(80, 33);
			this.label6.Name = "label6";
			this.label6.Size = new System.Drawing.Size(336, 16);
			this.label6.TabIndex = 66;
			this.label6.Text = "Por favor, introduzca sus credenciales de acceso al sistema.";
			// 
			// label5
			// 
			this.label5.AutoSize = true;
			this.label5.Font = new System.Drawing.Font("Tahoma", 14F, System.Drawing.FontStyle.Bold);
			this.label5.Location = new System.Drawing.Point(78, 12);
			this.label5.Name = "label5";
			this.label5.Size = new System.Drawing.Size(156, 26);
			this.label5.TabIndex = 65;
			this.label5.Text = "Acceso Sistema";
			// 
			// pictureBox3
			// 
			this.pictureBox3.Location = new System.Drawing.Point(16, 8);
			this.pictureBox3.Name = "pictureBox3";
			this.pictureBox3.Size = new System.Drawing.Size(48, 40);
			this.pictureBox3.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage;
			this.pictureBox3.TabIndex = 64;
			this.pictureBox3.TabStop = false;
			// 
			// pictureBox2
			// 
			this.pictureBox2.Location = new System.Drawing.Point(32, 64);
			this.pictureBox2.Name = "pictureBox2";
			this.pictureBox2.Size = new System.Drawing.Size(145, 124);
			this.pictureBox2.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage;
			this.pictureBox2.TabIndex = 63;
			this.pictureBox2.TabStop = false;
			// 
			// cmbEmpresas
			// 
			appearance1.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbEmpresas.Appearance = appearance1;
			this.cmbEmpresas.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbEmpresas.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbEmpresas.DataSource = this.ultraDataSource1;
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridColumn2.Width = 184;
			ultraGridBand1.Columns.AddRange(new object[] {
																										 ultraGridColumn1,
																										 ultraGridColumn2});
			this.cmbEmpresas.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.cmbEmpresas.DisplayMember = "Empresa";
			this.cmbEmpresas.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbEmpresas.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbEmpresas.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbEmpresas.Location = new System.Drawing.Point(280, 136);
			this.cmbEmpresas.Name = "cmbEmpresas";
			this.cmbEmpresas.Size = new System.Drawing.Size(184, 21);
			this.cmbEmpresas.TabIndex = 177;
			this.cmbEmpresas.ValueMember = "idEmpresa";
			this.cmbEmpresas.Visible = false;
			this.cmbEmpresas.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cmbEmpresas_KeyDown);
			this.cmbEmpresas.ValueChanged += new System.EventHandler(this.cmbEmpresas_ValueChanged);
			this.cmbEmpresas.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.cmbEmpresas_InitializeLayout);
			// 
			// ultraDataSource1
			// 
			ultraDataColumn1.DataType = typeof(int);
			this.ultraDataSource1.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn1,
																																 ultraDataColumn2});
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
			// label1
			// 
			this.label1.AutoSize = true;
			this.label1.Font = new System.Drawing.Font("Tahoma", 6F, System.Drawing.FontStyle.Bold);
			this.label1.ForeColor = System.Drawing.Color.Black;
			this.label1.Location = new System.Drawing.Point(472, 120);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(35, 13);
			this.label1.TabIndex = 178;
			this.label1.Text = "Mostrar";
			this.label1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.label1.Click += new System.EventHandler(this.label1_Click);
			// 
			// Inicio
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 14);
			this.ClientSize = new System.Drawing.Size(521, 207);
			this.ControlBox = false;
			this.Controls.Add(this.label1);
			this.Controls.Add(this.label4);
			this.Controls.Add(this.label3);
			this.Controls.Add(this.label2);
			this.Controls.Add(this.cmbEmpresas);
			this.Controls.Add(this.pictureBox2);
			this.Controls.Add(this.groupBox1);
			this.Controls.Add(this.txtNombre);
			this.Controls.Add(this.btCancelar);
			this.Controls.Add(this.btAceptar);
			this.Controls.Add(this.txtClave);
			this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
			this.HelpButton = true;
			this.helpProvider.SetHelpNavigator(this, System.Windows.Forms.HelpNavigator.Topic);
			this.helpProvider.SetHelpString(this, "Adminstrador y Contraseña en blanco");
			this.KeyPreview = true;
			this.MaximizeBox = false;
			this.MinimizeBox = false;
			this.Name = "Inicio";
			this.helpProvider.SetShowHelp(this, true);
			this.ShowInTaskbar = false;
			this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
			this.Text = "Acceso";
			this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.Inicio_KeyDown);
			this.Load += new System.EventHandler(this.Inicio_Load);
			((System.ComponentModel.ISupportInitialize)(this.txtClave)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNombre)).EndInit();
			this.groupBox1.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.cmbEmpresas)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeguridad)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		public static int Intento = 0;
		public string stNombreEmpresa = ""; //"POINT TECHNOLOGY";
		public string stDir = "";
		//public string stServidorReplicacion = "";

		private void btAceptar_Click(object sender, System.EventArgs e)
		{
			#region Valida Campos vacios
			
			if (this.txtNombre.Text.ToString().Trim().Length == 0)
			{
				MessageBox.Show("Escriba el nombre de usuario.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.txtNombre.Focus();
				return;
			}
			if (this.txtClave.Text.ToString().Trim().Length == 0)
			{
				MessageBox.Show("Escriba la clave de usuario.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.txtClave.Focus();
				return;
			}
			#endregion Valida Campos vacios
			
			Funcion.TiempoInicia();
			
			string strHostName = string.Empty;
			strHostName = System.Environment.MachineName;

			MenuLatinium.stUsuario = this.txtNombre.Text.ToString();
						
			string stClave = this.txtClave.Text.Trim();

			SqlConnection Conexion = new SqlConnection(@"Password=Bl45o6$9;Persist Security Info=True;User ID= 'Infoelect';Initial Catalog='INICIO';Data Source='192.168.1.165,56399'");
			
			//SqlConnection Conexion = new SqlConnection(@"Password=Bl45o6$9;Persist Security Info=True;User ID=Infoelect;Initial Catalog=INICIO;Data Source=172.16.10.2\DentalisServer;");
			
			string sSQL = string.Format("Select Empresa, BaseDeDatos, DirReportes, Servidor, ServidorReplicacion,BDReplicacion, DirectorioArchivo, usuarioArchivo, clavearchivo, ApiDepositosPendientes, apiDt, ApiFacturasCompra, stApiRespaldoPayJoy From Empresas Where idEmpresa = {0}", (int)this.cmbEmpresas.Value);			
			SqlCommand command = new SqlCommand(sSQL, Conexion);
			Conexion.Open();
			SqlDataReader reader = command.ExecuteReader();
			reader.Read();
			if (reader.HasRows)
			{
				stNombreEmpresa = (string)reader[0];
				stDir = (string)reader[1];
							
				MenuLatinium.stDirServidor = (string)reader[3];
				MenuLatinium.stServidorReplicacion =(string)reader[4];			
				MenuLatinium.stDirReportes = (string)reader[2];				
				MenuLatinium.stBDReplicacion =(string)reader[5];
				MenuLatinium.DirectorioArchivo =(string)reader[6];
				MenuLatinium.usuarioArchivo =(string)reader[7];
				MenuLatinium.clavearchivo =(string)reader[8];
				MenuLatinium.stApiDepositosPendientes =(string)reader[9];
				MenuLatinium.stapiDt =(string)reader[10];
				MenuLatinium.stApiFacturasCompra =(string)reader[11];
				MenuLatinium.stApiRespaldoPayJoy =(string)reader[12]; 

			}
			reader.Close();
			Conexion.Close();		
			
			MenuLatinium.stDirFacturacion = stDir;
			MenuLatinium.stDirConta = stDir;				
			MenuLatinium.stDirNumero = stDir;
			MenuLatinium.stDirRoles = stDir;
			MenuLatinium.stDirRetenc = stDir;
			MenuLatinium.iTipoDB = 3;	
				
			MenuLatinium.ImpresoraNombre = printDoc.PrinterSettings.PrinterName;
			MenuLatinium.stIdDB = (int)this.cmbEmpresas.Value;
			
			#region Valida Clave De Acceso
			
			string sDataSource = MenuLatinium.stDirServidor;

			if (MenuLatinium.stDirServidor  == @"192.168.1.165\POINTDB")
			{
				sDataSource = @"192.168.1.165,56399";
				MenuLatinium.stDirServidor = sDataSource;
				MenuLatinium.stServidorReplicacion = sDataSource;
			}

			SqlConnection oConexion = new SqlConnection(
				"Password=Bl45o6$9;" +
				"Persist Security Info=True;" +
				"User ID=Infoelect;" +
				"Initial Catalog=" + MenuLatinium.stDirFacturacion + ";" +
				"Data Source=" + sDataSource + ";" +
				"Max Pool Size=1024"
				);
			
			string sSQLId = string.Format("Select idGrupo, idUsuario From Usuario Where Nombre = '{0}' And Clave = '{1}' And Activo = 1", this.txtNombre.Text.ToString(), this.txtClave.Text.ToString());

			if (MenuLatinium.stDirFacturacion == "INICIO") 
				sSQLId = string.Format("Select idGrupo, idUsuario From Usuarios Where Nombre = '{0}' And Clave = '{1}' And Activo = 1", this.txtNombre.Text.ToString(), this.txtClave.Text.ToString());

			SqlCommand oCmdId = new SqlCommand(sSQLId, oConexion);
			oCmdId.CommandTimeout = 120;
			oConexion.Open();
			SqlDataReader oDRId = oCmdId.ExecuteReader();
			oDRId.Read();
			if (oDRId.HasRows)
			{
				iNivel = (int)oDRId[0];
				idUsuario = (int)oDRId[1];
			}
			else
			{
				oConexion.Close();
				MessageBox.Show("Usuario o Clave Incorrecta", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return;
			}
			oConexion.Close();			
			#endregion Valida Clave De Acceso

			if(iNivel == 23 || iNivel == 25)
			{
				MenuLatinium.stDirReportes = @"C:\Latinium\Reportes\";
			}
//				 if ( idUsuario == 396 ) MenuLatinium.stDirReportes = @"C:\Latinium\Reportes\";
//			using (frmActualizacionEmail consultprofos = new frmActualizacionEmail ( this.txtNombre.Text))
//			{
//				if ( DialogResult.OK == consultprofos.ShowDialog ())
//				{
//           this.DialogResult = DialogResult.OK ;
//				}
//			}

			#region Valida Actualizacion Sistema

				 string dFechaAbf ="";

			if (strHostName == "IEXP")
			{
				string dFechaAb = Convert.ToDateTime(dFechaModificacionA).ToString("yyyyMMdd HH:mm:ss");
				dFechaAbf = dFechaAb;

//									string sSQLAsiEs = string.Format("EXEC ActualizaSistema '{0}' ", dFechaAb );			
//									Funcion.EjecutaSQL(cdsSeguridad, sSQLAsiEs, true);  
			}

			#endregion Valida Actualizacion Sistema			

			this.DialogResult = DialogResult.OK;		
		}
		
		public int iNivel = 0;
		public int idUsuario = 0;
		PrintDocument printDoc = new PrintDocument();

		DateTime dFechaModificacion ;
		DateTime dFechaModificacionA ;
		
		private void Inicio_Load(object sender, System.EventArgs e)
		{			
	
			string stDirInicioAc = Application.StartupPath;

//			if (stDirInicioAc.EndsWith("Debug")) stDirInicio = @"C:\Latinium";
			if (stDirInicioAc.EndsWith("Debug"))
			{
				dFechaModificacion = File.GetCreationTime ( @"C:\Latinium\Latinium\bin\Release\Latinium.exe");
				dFechaModificacionA = File.GetLastWriteTimeUtc ( @"C:\Latinium\Latinium\bin\Release\Latinium.exe");
				dFechaModificacionA = File.GetLastWriteTime ( @"C:\Latinium\Latinium\bin\Release\Latinium.exe");
			}
		}

		private void btCancelar_Click(object sender, System.EventArgs e)
		{			
			this.Close();
		}

		private void cdsUsuario_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
		
		}

		private void cdsEmpresa_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
		
		}

		private void btCancelar_MouseDown(object sender, System.Windows.Forms.MouseEventArgs e)
		{
			if (e.Button != MouseButtons.Right) return;
			string stNombre1 = @"c:\bd11.mdb"; 
			string stNombre2 = @"c:\bd21.mdb"; 
			stNombre1 = @"c:\ex1.xls"; 
			stNombre2 = @"c:\ex2.xls"; 
			#region No es necesario
			StreamReader sr1 = null;
			string stArch1 = "";
			if (File.Exists(stNombre1))
			{
				sr1 = new StreamReader(stNombre1,System.Text.Encoding.UTF7);
				stArch1  = sr1.ReadToEnd();
			}
			#endregion No es necesario
			string stComp = stNombre2;
			if (txtClave.Text.Length > 0)
				stComp = @"c:\" + txtClave.Text.Trim() + ".mdb";
			if (!File.Exists(stComp))
			{
				MessageBox.Show("No existe Archivo: " + stComp, "Información");
				return;
			}

			StreamReader sr2 = new StreamReader(stComp,System.Text.Encoding.UTF7);
			string stArch2  = sr2.ReadToEnd();
			char [] stClave1 = stArch1.ToCharArray();
			char [] stClave2 = stArch2.ToCharArray();

			string stMensaje = "";
			int j = 0;
			int [] iClavePub = new int[16];// 
			iClavePub[0] = 134;
			iClavePub[0] = 134;iClavePub[1] = 251;iClavePub[2] = 236;iClavePub[3] = 55;iClavePub[4] = 93;
			iClavePub[5] = 68;iClavePub[6] = 156;iClavePub[7] = 250;iClavePub[8] = 198;iClavePub[9] = 94;
			iClavePub[10] = 40;iClavePub[11] = 230;iClavePub[12] = 19;iClavePub[13] = 182;
			iClavePub[14] = 138;iClavePub[15] = 96;
			int iIni = 0;
			int iFin = 1000;
			if (!File.Exists(stNombre1))
			{
				iIni = 66;
				iFin = 82;
			}

			int iCompara = 0;
			string stAscii = "\nASCII: ";
			for (int i = iIni; i< iFin; i++)
			{
				if (iIni > 0)	iCompara = iClavePub[j];
				else iCompara = stClave1[i];
				if (iCompara != stClave2[i])
				{
					Console.WriteLine("{0}\t{1}\t{2}\t{3}", i, iCompara, 
						Convert.ToInt32(stClave2[i]),iCompara ^ stClave2[i]);
					stMensaje += Convert.ToChar(iCompara ^ stClave2[i]);
					stAscii += Convert.ToInt32(iCompara ^ stClave2[i]) + " ";
					if (j >= 16) i=31000;
				}
				j++;
			}
			Console.WriteLine("Fin {0} : {1}", stComp, stMensaje);
			if (iIni == 0)	sr1.Close();
			sr2.Close();
			MessageBox.Show("Clave: " + stMensaje + stAscii, "Decodificado");	
		}

		private void btAceptar_MouseUp(object sender, System.Windows.Forms.MouseEventArgs e)
		{
//			if (e.Button != MouseButtons.Right) return;
//			if (DialogResult.No == MessageBox.Show("¿Desea usar accesos seguros?", "Confirmación", MessageBoxButtons.YesNo, MessageBoxIcon.Question)) return;
		}

		private void txtNombre_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (e.KeyChar == 13) this.txtClave.Focus();
		}

		private void txtClave_ValueChanged(object sender, System.EventArgs e)
		{
			if (this.cmbEmpresas.Visible) this.cmbEmpresas.Visible = false;
		}

		private void txtClave_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (e.KeyChar == 13) this.btAceptar.Focus();
		}

		private void Inicio_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Escape) this.Close();
		}

		private void cmbEmpresas_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}

		private void cmbEmpresas_ValueChanged(object sender, System.EventArgs e)
		{

		}

		private void cmbEmpresas_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.btAceptar.Focus();
		}

		private void txtClave_Validated(object sender, System.EventArgs e)
		{
			if (this.txtNombre.Text.Length > 0 && this.txtClave.Text.Length > 0)
			{				
				try
				{
				//	SqlConnection oConexion = new SqlConnection(@"Password=Bl45o6$9;Persist Security Info=True;User ID= 'Infoelect';Initial Catalog='INICIO';Data Source='192.168.1.19'");
					SqlConnection oConexion = new SqlConnection(@"Password=Bl45o6$9;Persist Security Info=True;User ID= 'Infoelect';Initial Catalog='INICIO';Data Source='192.168.1.165,56399'");
					//SqlConnection oConexion = new SqlConnection(@"Password=Bl45o6$9;Persist Security Info=True;User ID= 'Infoelect';Initial Catalog='INICIO';Data Source='192.168.1.15'");
//         	SqlConnection oConexion = new SqlConnection(@"Password=Bl45o6$9;User ID=Infoelect;Initial Catalog=INICIO;Data Source=172.16.10.2\DentalisServer;Max Pool Size = 2048;");
						
					SqlCommand oCmd = new SqlCommand();
					oCmd.Connection = oConexion;
					oCmd.CommandType = CommandType.Text;
					oCmd.CommandText = string.Format("Select e.idEmpresa, e.Empresa From Empresas e Inner Join EmpresasUsuarios eu On eu.idEmpresa = e.idEmpresa Where e.Estado = 1 And eu.idUsuario = {0} Order By e.Empresa", idUsuario);
					SqlDataAdapter oDA = new SqlDataAdapter(oCmd.CommandText, oConexion);
				
					DataSet oDS = new DataSet();
					oDA.Fill(oDS);
					this.cmbEmpresas.DataSource = oDS;
				
					SqlCommand oCmdEmp = new SqlCommand(string.Format("Select Top 1 e.idEmpresa From Empresas e Inner Join EmpresasUsuarios eu On eu.idEmpresa = e.idEmpresa Where e.Estado = 1 And eu.idUsuario = {0} Order By e.idEmpresa", idUsuario), oConexion);
					oConexion.Open();
				
					this.cmbEmpresas.Value = (int)oCmdEmp.ExecuteScalar();
					oConexion.Close();


					SqlCommand oCmdEmpvER = new SqlCommand(string.Format("select versionamiento  from Seteo "), oConexion);
					oConexion.Open();
				
					MenuLatinium.sVersionSeteo = (string)oCmdEmpvER.ExecuteScalar();
					oConexion.Close();


          MenuLatinium.IdUsuarioBDDInicio =  idUsuario;
					SqlCommand oCmdEmpPass = new SqlCommand(string.Format("Exec ValidaEstadoContraseñaAcceso '{0}', '{1}', {2}", 
						this.txtNombre.Text.ToString(), this.txtClave.Text.ToString(), MenuLatinium.IdUsuarioBDDInicio), oConexion);
					oConexion.Open();
				
					MenuLatinium.sMensajePasswordAcceso = (string)oCmdEmpPass.ExecuteScalar();
					oConexion.Close();



				}
				catch(Exception ex)
				{
					MessageBox.Show(ex.ToString(), "VALIDADO");
				}
				if (this.cmbEmpresas.Rows.Count > 1) 
				{
					this.cmbEmpresas.Visible = true;
					this.cmbEmpresas.Focus();
				}
				else 
				{
					this.cmbEmpresas.Visible = false;
					this.btAceptar.Focus();
				}
			}
		}

		private void txtClave_Validating(object sender, System.ComponentModel.CancelEventArgs e)
		{
			if (this.txtClave.Text.Length > 0)
			{
				#region Valida Clave De Acceso
				try
				{//Persist Security Info=True; 172.16.10.2
					
					//SqlConnection Conexion = new SqlConnection(@"Password=Bl45o6$9;Persist Security Info=True;User ID= 'Infoelect';Initial Catalog='INICIO';Data Source='192.168.1.19'");
					
					
					SqlConnection Conexion = new SqlConnection(@"Password=Bl45o6$9;Persist Security Info=True;User ID= 'Infoelect';Initial Catalog='INICIO';Data Source='192.168.1.165,56399'");
					//SqlConnection Conexion = new SqlConnection(@"Password=Bl45o6$9;Persist Security Info=True;User ID= 'Infoelect';Initial Catalog='INICIO';Data Source='192.168.1.15'");
//					SqlConnection Conexion = new SqlConnection(@"Password=Bl45o6$9;User ID=Infoelect;Initial Catalog=INICIO;Data Source=172.16.10.2\DentalisServer;Max Pool Size = 2048;");
									
					string sSQL = string.Format("Select idGrupo, idUsuario From Usuarios Where Nombre = '{0}' And Clave = '{1}' And Activo = 1", this.txtNombre.Text.ToString(), this.txtClave.Text.ToString());
					SqlCommand command = new SqlCommand(sSQL, Conexion);
					Conexion.Open();
					
					SqlDataReader reader = command.ExecuteReader();
					reader.Read();
					if (reader.HasRows)
					{
						iNivel = (int)reader[0];
						idUsuario = (int)reader[1];
					}
					reader.Close();
					Conexion.Close();
				}
				catch(Exception ex)
				{
					MessageBox.Show(ex.ToString(), "VALIDACION");
					return;
				}
	
				if (iNivel == 0)
				{
					MessageBox.Show("Usuario o Clave incorrecta", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);				
					this.txtNombre.Focus();
					e.Cancel = true;
				}
				else
				{
					
					MenuLatinium.stUsuario = this.txtNombre.Text.ToString().Trim();								
				}
				#endregion Valida Clave De Acceso
			}
		}

		private void groupBox1_Enter(object sender, System.EventArgs e)
		{
		
		}

		private void txtNombre_ValueChanged(object sender, System.EventArgs e)
		{
		
		}
    private bool claveVisible = false;
		private void label1_Click(object sender, System.EventArgs e)
		{
			if (claveVisible)
			{
				// Ocultar clave
				this.txtClave.PasswordChar  = '*';
				this.label1.Text = "Mostrar";
				claveVisible = false;
			}
			else
			{
				// Mostrar clave
				this.txtClave.PasswordChar = '\0'; // Sin carácter = muestra texto plano
				this.label1.Text = "Ocultar";
				claveVisible = true;
			}
		}

		
	}
}
