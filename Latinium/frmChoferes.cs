using System;
using System.Globalization;
using System.Data;
using System.Data.SqlClient;
using C1.Data;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using Infragistics.Win.UltraWinGrid;
using System.Web.Mail;
using CrystalDecisions.Shared;
using System.Runtime.InteropServices;
using System.Diagnostics;
using System.IO;
using CrystalDecisions;
using CrystalDecisions.CrystalReports.Engine;
using CrystalDecisions.ReportSource;
using System.Net;
using System.Web;
using System.Xml;
using System.Xml.Schema;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de frmChoferes.
	/// </summary>
	public class frmChoferes : System.Windows.Forms.Form
	{
		private System.Windows.Forms.Button btnSalir;
		private System.Windows.Forms.Button btnNuevo;
		private System.Windows.Forms.Button btnConsultar;
		private System.Windows.Forms.Button btnEditar;
		public System.Windows.Forms.Button btnGuardar;
		private System.Windows.Forms.Button btnCancelar;
		public Infragistics.Win.UltraWinEditors.UltraNumericEditor txtidChofer;
		public Infragistics.Win.UltraWinEditors.UltraTextEditor txtCedula;
		private System.Windows.Forms.Label label1;
		private System.Windows.Forms.Label label3;
		private System.Windows.Forms.Label label2;
		public Infragistics.Win.UltraWinEditors.UltraTextEditor txtPlaca;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkInventario;
		private System.Windows.Forms.Label label4;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtNombre;
		private System.Windows.Forms.Label lblEstado;
		private C1.Data.C1DataSet cdsSeteoF;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public frmChoferes()
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
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmChoferes));
			Infragistics.Win.Appearance appearance1 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			this.btnSalir = new System.Windows.Forms.Button();
			this.btnNuevo = new System.Windows.Forms.Button();
			this.btnConsultar = new System.Windows.Forms.Button();
			this.btnEditar = new System.Windows.Forms.Button();
			this.btnGuardar = new System.Windows.Forms.Button();
			this.btnCancelar = new System.Windows.Forms.Button();
			this.txtidChofer = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtCedula = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label1 = new System.Windows.Forms.Label();
			this.label3 = new System.Windows.Forms.Label();
			this.txtNombre = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label2 = new System.Windows.Forms.Label();
			this.txtPlaca = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.chkInventario = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.label4 = new System.Windows.Forms.Label();
			this.lblEstado = new System.Windows.Forms.Label();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			((System.ComponentModel.ISupportInitialize)(this.txtidChofer)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCedula)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNombre)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtPlaca)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			this.SuspendLayout();
			// 
			// btnSalir
			// 
			this.btnSalir.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnSalir.CausesValidation = false;
			this.btnSalir.Image = ((System.Drawing.Image)(resources.GetObject("btnSalir.Image")));
			this.btnSalir.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnSalir.Location = new System.Drawing.Point(408, 128);
			this.btnSalir.Name = "btnSalir";
			this.btnSalir.TabIndex = 1017;
			this.btnSalir.Text = "&Salir";
			this.btnSalir.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// btnNuevo
			// 
			this.btnNuevo.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnNuevo.CausesValidation = false;
			this.btnNuevo.Enabled = false;
			this.btnNuevo.Image = ((System.Drawing.Image)(resources.GetObject("btnNuevo.Image")));
			this.btnNuevo.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnNuevo.Location = new System.Drawing.Point(8, 128);
			this.btnNuevo.Name = "btnNuevo";
			this.btnNuevo.Size = new System.Drawing.Size(72, 24);
			this.btnNuevo.TabIndex = 1015;
			this.btnNuevo.Text = "&Nuevo";
			this.btnNuevo.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnNuevo.Click += new System.EventHandler(this.btnNuevo_Click);
			// 
			// btnConsultar
			// 
			this.btnConsultar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnConsultar.CausesValidation = false;
			this.btnConsultar.Enabled = false;
			this.btnConsultar.Image = ((System.Drawing.Image)(resources.GetObject("btnConsultar.Image")));
			this.btnConsultar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnConsultar.Location = new System.Drawing.Point(88, 128);
			this.btnConsultar.Name = "btnConsultar";
			this.btnConsultar.Size = new System.Drawing.Size(78, 24);
			this.btnConsultar.TabIndex = 1014;
			this.btnConsultar.Text = "&Consultar";
			this.btnConsultar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnConsultar.Click += new System.EventHandler(this.btnConsultar_Click);
			// 
			// btnEditar
			// 
			this.btnEditar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnEditar.CausesValidation = false;
			this.btnEditar.Enabled = false;
			this.btnEditar.Image = ((System.Drawing.Image)(resources.GetObject("btnEditar.Image")));
			this.btnEditar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnEditar.Location = new System.Drawing.Point(168, 128);
			this.btnEditar.Name = "btnEditar";
			this.btnEditar.Size = new System.Drawing.Size(76, 24);
			this.btnEditar.TabIndex = 1013;
			this.btnEditar.Text = "&Editar";
			this.btnEditar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnEditar.Click += new System.EventHandler(this.btnEditar_Click);
			// 
			// btnGuardar
			// 
			this.btnGuardar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnGuardar.Enabled = false;
			this.btnGuardar.Image = ((System.Drawing.Image)(resources.GetObject("btnGuardar.Image")));
			this.btnGuardar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnGuardar.Location = new System.Drawing.Point(248, 128);
			this.btnGuardar.Name = "btnGuardar";
			this.btnGuardar.Size = new System.Drawing.Size(76, 24);
			this.btnGuardar.TabIndex = 1012;
			this.btnGuardar.Text = "&Guardar";
			this.btnGuardar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnGuardar.Click += new System.EventHandler(this.btnGuardar_Click);
			// 
			// btnCancelar
			// 
			this.btnCancelar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnCancelar.CausesValidation = false;
			this.btnCancelar.Enabled = false;
			this.btnCancelar.Image = ((System.Drawing.Image)(resources.GetObject("btnCancelar.Image")));
			this.btnCancelar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnCancelar.Location = new System.Drawing.Point(328, 128);
			this.btnCancelar.Name = "btnCancelar";
			this.btnCancelar.Size = new System.Drawing.Size(72, 24);
			this.btnCancelar.TabIndex = 1011;
			this.btnCancelar.Text = "&Cancelar";
			this.btnCancelar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnCancelar.Click += new System.EventHandler(this.btnCancelar_Click);
			// 
			// txtidChofer
			// 
			this.txtidChofer.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtidChofer.Enabled = false;
			this.txtidChofer.Location = new System.Drawing.Point(424, 8);
			this.txtidChofer.Name = "txtidChofer";
			this.txtidChofer.PromptChar = ' ';
			this.txtidChofer.Size = new System.Drawing.Size(56, 21);
			this.txtidChofer.SpinButtonAlignment = Infragistics.Win.ButtonAlignment.Left;
			this.txtidChofer.TabIndex = 1018;
			this.txtidChofer.Visible = false;
			// 
			// txtCedula
			// 
			appearance1.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtCedula.Appearance = appearance1;
			this.txtCedula.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtCedula.Enabled = false;
			this.txtCedula.Location = new System.Drawing.Point(64, 16);
			this.txtCedula.MaxLength = 10;
			this.txtCedula.Name = "txtCedula";
			this.txtCedula.Size = new System.Drawing.Size(136, 21);
			this.txtCedula.TabIndex = 1020;
			this.txtCedula.Validating += new System.ComponentModel.CancelEventHandler(this.txtCedula_Validating);
			// 
			// label1
			// 
			this.label1.AutoSize = true;
			this.label1.Location = new System.Drawing.Point(8, 18);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(40, 16);
			this.label1.TabIndex = 1021;
			this.label1.Text = "Cedula";
			// 
			// label3
			// 
			this.label3.AutoSize = true;
			this.label3.Location = new System.Drawing.Point(8, 48);
			this.label3.Name = "label3";
			this.label3.Size = new System.Drawing.Size(44, 16);
			this.label3.TabIndex = 1022;
			this.label3.Text = "Nombre";
			this.label3.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtNombre
			// 
			appearance2.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtNombre.Appearance = appearance2;
			this.txtNombre.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtNombre.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNombre.Enabled = false;
			this.txtNombre.Location = new System.Drawing.Point(64, 46);
			this.txtNombre.Name = "txtNombre";
			this.txtNombre.Size = new System.Drawing.Size(424, 21);
			this.txtNombre.TabIndex = 1023;
			// 
			// label2
			// 
			this.label2.AutoSize = true;
			this.label2.BackColor = System.Drawing.Color.Transparent;
			this.label2.Location = new System.Drawing.Point(8, 82);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(33, 16);
			this.label2.TabIndex = 1025;
			this.label2.Text = "Placa";
			// 
			// txtPlaca
			// 
			appearance3.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.txtPlaca.Appearance = appearance3;
			this.txtPlaca.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtPlaca.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtPlaca.Enabled = false;
			this.txtPlaca.Location = new System.Drawing.Point(64, 80);
			this.txtPlaca.MaxLength = 7;
			this.txtPlaca.Name = "txtPlaca";
			this.txtPlaca.Size = new System.Drawing.Size(136, 21);
			this.txtPlaca.TabIndex = 1024;
			// 
			// chkInventario
			// 
			appearance4.ForeColorDisabled = System.Drawing.Color.Black;
			this.chkInventario.Appearance = appearance4;
			this.chkInventario.BackColor = System.Drawing.Color.Transparent;
			this.chkInventario.Enabled = false;
			this.chkInventario.Location = new System.Drawing.Point(264, 80);
			this.chkInventario.Name = "chkInventario";
			this.chkInventario.Size = new System.Drawing.Size(16, 21);
			this.chkInventario.TabIndex = 1026;
			// 
			// label4
			// 
			this.label4.AutoSize = true;
			this.label4.BackColor = System.Drawing.Color.Transparent;
			this.label4.Location = new System.Drawing.Point(216, 82);
			this.label4.Name = "label4";
			this.label4.Size = new System.Drawing.Size(35, 16);
			this.label4.TabIndex = 1027;
			this.label4.Text = "Activo";
			// 
			// lblEstado
			// 
			this.lblEstado.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.lblEstado.AutoSize = true;
			this.lblEstado.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblEstado.ForeColor = System.Drawing.Color.Firebrick;
			this.lblEstado.Location = new System.Drawing.Point(296, 79);
			this.lblEstado.Name = "lblEstado";
			this.lblEstado.Size = new System.Drawing.Size(0, 22);
			this.lblEstado.TabIndex = 1028;
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
			// frmChoferes
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(496, 158);
			this.Controls.Add(this.lblEstado);
			this.Controls.Add(this.label4);
			this.Controls.Add(this.chkInventario);
			this.Controls.Add(this.label2);
			this.Controls.Add(this.txtPlaca);
			this.Controls.Add(this.label3);
			this.Controls.Add(this.txtNombre);
			this.Controls.Add(this.label1);
			this.Controls.Add(this.txtCedula);
			this.Controls.Add(this.txtidChofer);
			this.Controls.Add(this.btnSalir);
			this.Controls.Add(this.btnNuevo);
			this.Controls.Add(this.btnConsultar);
			this.Controls.Add(this.btnEditar);
			this.Controls.Add(this.btnGuardar);
			this.Controls.Add(this.btnCancelar);
			this.KeyPreview = true;
			this.Name = "frmChoferes";
			this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
			this.Text = "Ingreso Chofer";
			this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.frmChoferes_KeyDown);
			this.Load += new System.EventHandler(this.frmChoferes_Load);
			((System.ComponentModel.ISupportInitialize)(this.txtidChofer)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCedula)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNombre)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtPlaca)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void txtCedula_Validating(object sender, System.ComponentModel.CancelEventArgs e)
		{
			if (bEdicion)
				if (!Validacion.vbIdentificacion(this.txtCedula, 5, cdsSeteoF)) e.Cancel = true;
		}
		private Acceso miAcceso;
		private void UnloadMe()
		{
			this.Close();
		}

		bool bNuevo = false;
		bool bEdicion = false;
		bool bAnulado = false;

		private void frmChoferes_Load(object sender, System.EventArgs e)
		{
			miAcceso = new Acceso(cdsSeteoF, "200112");

			if (!Funcion.Permiso("1893", cdsSeteoF))
			{				
				MessageBox.Show("No puede ingresar a Registro de Choferes", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				BeginInvoke(new MethodInvoker(UnloadMe));
				return;
			}

			if (miAcceso.Nuevo) this.btnNuevo.Enabled = true;
			if (miAcceso.Buscar) this.btnConsultar.Enabled = true;


		}

		private void btnGuardar_Click(object sender, System.EventArgs e)
		{
			if (!Validacion.vbIdentificacion(this.txtCedula, 5, cdsSeteoF)) return;
			if (!Validacion.vbTexto(this.txtNombre, 8, 280, "Nombres")) return;
			if (!Validacion.vbValidaPlaca(this.txtPlaca)) return;

			#region Guardar

			//			DateTime dtInicia = DateTime.Today;
			//			DateTime dtRetorna = DateTime.Today;
			//				dtInicia = (DateTime)this.dtInicia.Value;
			//				dtRetorna = (DateTime)this.dtRetorna.Value;

			using (SqlConnection oConexion = new SqlConnection("Password=Bl45o6$9;Persist Security Info=True;User ID= " + MenuLatinium.stUsuario + ";Initial Catalog=" + MenuLatinium.stDirFacturacion + ";Data Source=" + MenuLatinium.stDirServidor + "; max pool size=1024"))
			{
				oConexion.Open();

				SqlTransaction oTransaction = null;
				SqlCommand oCmdActualiza = oConexion.CreateCommand();
				oCmdActualiza.Connection = oConexion;
				oCmdActualiza.CommandType = CommandType.Text;
				oCmdActualiza.CommandTimeout = 120;

				try
				{
					oTransaction = oConexion.BeginTransaction();
					oCmdActualiza.Transaction = oTransaction;

					

					#region Graba Maestro

					string  sSQL1 = string.Format("Exec [GuardaChoferes] {0}, '{1}', '{2}', '{3}', {4}" ,
						(int)this.txtidChofer.Value, //0
						(string) this.txtCedula.Text, //7
						(string) this.txtNombre.Text,//8
						(string) this.txtPlaca.Text,
						(bool) this.chkInventario.Checked//9
						);
					oCmdActualiza.CommandText = sSQL1;
					this.txtidChofer.Value = (int)oCmdActualiza.ExecuteScalar();
					#endregion Graba Maestro
          
					oTransaction.Commit();

					Cursor = Cursors.Default;	

					MessageBox.Show("Registro Guardado Correctamente", "Información", MessageBoxButtons.OK, MessageBoxIcon.Information);

					#region Controles
					this.txtCedula.Enabled = false;
					this.txtNombre.Enabled = false;
					this.txtPlaca.Enabled = false;	
					this.chkInventario.Enabled = false;

					if (miAcceso.Nuevo)this.btnNuevo.Enabled = true;
					if (miAcceso.Buscar)this.btnConsultar.Enabled = true;
					if (miAcceso.Editar)this.btnEditar.Enabled = true;
					this.btnGuardar.Enabled = false;
					this.btnCancelar.Enabled = true;

					this.Consulta((int)this.txtidChofer.Value);
					
					bNuevo = false;
					bEdicion = false;
					#endregion Controles			
				}
				catch(Exception ex)
				{
					try
					{ 
						oTransaction.Rollback();

						MessageBox.Show(string.Format("Commit Exception Type : {0} {1}", ex.GetType(), ex.Message), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
					}
					catch (Exception ex2)
					{
						MessageBox.Show(string.Format("Rollback Exception Type: {0} {1}", ex2.GetType(), ex2.Message), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);							
					}
				}
				finally
				{
					oConexion.Close();
				}
			}
				
			#endregion Guardar

		}

		private void btnCancelar_Click(object sender, System.EventArgs e)
		{
			bNuevo = false;
			bEdicion = false;

			#region Vacia Campos

			this.txtidChofer.Value = 0;
			this.txtCedula.Text = "";	
			this.txtNombre.Text =  "";
			this.txtPlaca.Text = "";	
			this.chkInventario.Checked=false;
			this.lblEstado.Text = "";	

			#endregion Vacia Campos

			#region Bloquea Controles
			this.txtCedula.Enabled = false;
			this.txtNombre.Enabled = false;
			this.txtPlaca.Enabled = false;	
			this.chkInventario.Enabled = false;
			#endregion Bloquea Controles

			this.btnNuevo.Enabled = true;
			this.btnConsultar.Enabled = true;
			this.btnEditar.Enabled = false;
			this.btnGuardar.Enabled = false;
			this.btnCancelar.Enabled = false;
		}

		private void frmChoferes_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			
			if (!bEdicion)
			{
				if (e.KeyCode == Keys.N) if (this.btnNuevo.Enabled) this.btnNuevo_Click(sender, e);
				if (e.KeyCode == Keys.B) if (this.btnConsultar.Enabled) this.btnConsultar_Click(sender, e);
				if (e.KeyCode == Keys.E) if (this.btnEditar.Enabled) this.btnEditar_Click(sender, e); 																																																										 
				if (e.KeyCode == Keys.F5)	this.Consulta((int)this.txtidChofer.Value);
				if (e.KeyCode == Keys.Escape) this.Close();					
			}
			else
			{
				if (e.KeyCode == Keys.Escape) 
					if (DialogResult.Yes == MessageBox.Show("Esta seguro de Cancelar la Edición", "Point Technology", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2))
						this.btnCancelar_Click(sender, e);

				if (e.KeyCode == Keys.F12) this.btnGuardar_Click(sender, e);								
			}
		}

		private void btnNuevo_Click(object sender, System.EventArgs e)
		{
			this.btnCancelar_Click(sender, e);
		

			#region bloqueo

			this.txtCedula.Enabled = true;
			this.txtNombre.Enabled = true;
			this.txtPlaca.Enabled = true;
			this.chkInventario.Enabled = true;
			#endregion bloqueo

			this.lblEstado.Text = "";


			#region controles
			this.btnNuevo.Enabled = false;
			this.btnConsultar.Enabled = false;
			this.btnEditar.Enabled = false;
			this.btnGuardar.Enabled = true;
			this.btnCancelar.Enabled = true;	
			#endregion controles
			
			bNuevo = true;
			bEdicion = true;

			this.txtCedula.Focus();		
		}

		private void btnEditar_Click(object sender, System.EventArgs e)
		{
			Consulta ((int)this.txtidChofer.Value);


				#region bloqueo

				this.txtNombre.Enabled = true;
				this.txtPlaca.Enabled = true;
				this.chkInventario.Enabled = true;
				
				#endregion bloqueo
			this.btnNuevo.Enabled = false;
			this.btnConsultar.Enabled = false;
			this.btnEditar.Enabled = false;			
			this.btnGuardar.Enabled = true;			
			this.btnCancelar.Enabled = true;


			bNuevo = false;
			bEdicion = true;
		}

		private void btnConsultar_Click(object sender, System.EventArgs e)
		{
			this.btnCancelar_Click(sender, e);

			using (frmConsultaChofer Busqueda = new frmConsultaChofer())
			{
				if (DialogResult.OK == Busqueda.ShowDialog())
				{
					this.Consulta((int)Busqueda.uGridCreditos.ActiveRow.Cells["idChoferes"].Value);
				}
			}
		}
		private void Consulta (int Id)
		{			
			string sSQL = string.Format("Exec [ConsultaIndividualChoferes] {0}", Id);
			SqlDataReader dr = Funcion.rEscalarSQL(cdsSeteoF, sSQL, true);
			dr.Read();
			if (dr.HasRows)
			{
				this.txtidChofer.Value = Id;
				if (dr.GetValue(1) != System.DBNull.Value) this.txtCedula.Text = dr.GetString(1);
				if (dr.GetValue(2) != System.DBNull.Value) this.txtNombre.Text = dr.GetString(2);
				if (dr.GetValue(3) != System.DBNull.Value) this.txtPlaca.Text = dr.GetString(3); 
				if (dr.GetValue(4) != System.DBNull.Value) this.chkInventario.Checked= dr.GetBoolean(4);
				if (dr.GetValue(5) != System.DBNull.Value) this.lblEstado.Text = dr.GetString(5);
				
			}
			dr.Close();

			this.btnGuardar.Enabled = false;
			this.btnCancelar.Enabled = true;

			if (miAcceso.Nuevo) this.btnNuevo.Enabled = true;
			if (miAcceso.Buscar) this.btnConsultar.Enabled = true;
			if (miAcceso.Editar) this.btnEditar.Enabled = true; 
				
		}

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);
		
		}
	}
}
