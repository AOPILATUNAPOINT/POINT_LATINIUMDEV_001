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
	/// Descripción breve de frmDescripcionArticuloWeb.
	/// </summary>
	public class frmDescripcionArticuloWeb : System.Windows.Forms.Form
	{
		public Infragistics.Win.UltraWinEditors.UltraTextEditor txtObservaciones;
		private System.Windows.Forms.Button btnArticulo;
		private System.Windows.Forms.Button btnGuardar;
		private System.Windows.Forms.Button btnCancelar;
		private System.Windows.Forms.Button btnSalir;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtCodigo;
		private System.Windows.Forms.Label label1;
		private C1.Data.C1DataSet cdsSeteoF;
		private System.Windows.Forms.Label lblNumero;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtProducto;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;


		int idArticulo = 0;
		string sArticulo = "";
		string sCodigo = "";

		private Acceso miAcceso;

		public frmDescripcionArticuloWeb()
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
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmDescripcionArticuloWeb));
			Infragistics.Win.Appearance appearance1 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			this.btnArticulo = new System.Windows.Forms.Button();
			this.txtObservaciones = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.btnGuardar = new System.Windows.Forms.Button();
			this.btnCancelar = new System.Windows.Forms.Button();
			this.btnSalir = new System.Windows.Forms.Button();
			this.txtCodigo = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label1 = new System.Windows.Forms.Label();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.lblNumero = new System.Windows.Forms.Label();
			this.txtProducto = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			((System.ComponentModel.ISupportInitialize)(this.txtObservaciones)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCodigo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtProducto)).BeginInit();
			this.SuspendLayout();
			// 
			// btnArticulo
			// 
			this.btnArticulo.Location = new System.Drawing.Point(264, 8);
			this.btnArticulo.Name = "btnArticulo";
			this.btnArticulo.Size = new System.Drawing.Size(24, 22);
			this.btnArticulo.TabIndex = 237;
			this.btnArticulo.Text = "...";
			this.btnArticulo.Visible = false;
			// 
			// txtObservaciones
			// 
			this.txtObservaciones.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.txtObservaciones.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtObservaciones.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtObservaciones.Location = new System.Drawing.Point(16, 72);
			this.txtObservaciones.Multiline = true;
			this.txtObservaciones.Name = "txtObservaciones";
			this.txtObservaciones.Size = new System.Drawing.Size(440, 84);
			this.txtObservaciones.TabIndex = 232;
			// 
			// btnGuardar
			// 
			this.btnGuardar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnGuardar.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(230)), ((System.Byte)(230)), ((System.Byte)(230)));
			this.btnGuardar.Enabled = false;
			this.btnGuardar.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.btnGuardar.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(50)), ((System.Byte)(50)), ((System.Byte)(50)));
			this.btnGuardar.Image = ((System.Drawing.Image)(resources.GetObject("btnGuardar.Image")));
			this.btnGuardar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnGuardar.Location = new System.Drawing.Point(16, 176);
			this.btnGuardar.Name = "btnGuardar";
			this.btnGuardar.Size = new System.Drawing.Size(76, 23);
			this.btnGuardar.TabIndex = 286;
			this.btnGuardar.Text = "Guardar";
			this.btnGuardar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnGuardar.Click += new System.EventHandler(this.btnGuardar_Click);
			// 
			// btnCancelar
			// 
			this.btnCancelar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnCancelar.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(230)), ((System.Byte)(230)), ((System.Byte)(230)));
			this.btnCancelar.CausesValidation = false;
			this.btnCancelar.Enabled = false;
			this.btnCancelar.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.btnCancelar.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(50)), ((System.Byte)(50)), ((System.Byte)(50)));
			this.btnCancelar.Image = ((System.Drawing.Image)(resources.GetObject("btnCancelar.Image")));
			this.btnCancelar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnCancelar.Location = new System.Drawing.Point(106, 176);
			this.btnCancelar.Name = "btnCancelar";
			this.btnCancelar.Size = new System.Drawing.Size(72, 23);
			this.btnCancelar.TabIndex = 288;
			this.btnCancelar.Text = "Cancelar";
			this.btnCancelar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnCancelar.Click += new System.EventHandler(this.btnCancelar_Click);
			// 
			// btnSalir
			// 
			this.btnSalir.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnSalir.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(230)), ((System.Byte)(230)), ((System.Byte)(230)));
			this.btnSalir.CausesValidation = false;
			this.btnSalir.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.btnSalir.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(50)), ((System.Byte)(50)), ((System.Byte)(50)));
			this.btnSalir.Image = ((System.Drawing.Image)(resources.GetObject("btnSalir.Image")));
			this.btnSalir.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnSalir.Location = new System.Drawing.Point(192, 176);
			this.btnSalir.Name = "btnSalir";
			this.btnSalir.TabIndex = 289;
			this.btnSalir.Text = "Salir";
			this.btnSalir.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnSalir.Click += new System.EventHandler(this.btnSalir_Click);
			// 
			// txtCodigo
			// 
			this.txtCodigo.AcceptsTab = true;
			appearance1.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtCodigo.Appearance = appearance1;
			this.txtCodigo.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtCodigo.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtCodigo.Enabled = false;
			this.txtCodigo.Location = new System.Drawing.Point(80, 40);
			this.txtCodigo.Name = "txtCodigo";
			this.txtCodigo.Size = new System.Drawing.Size(376, 21);
			this.txtCodigo.TabIndex = 320;
			// 
			// label1
			// 
			this.label1.AutoSize = true;
			this.label1.BackColor = System.Drawing.Color.Transparent;
			this.label1.Location = new System.Drawing.Point(16, 40);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(42, 16);
			this.label1.TabIndex = 321;
			this.label1.Text = "Articulo";
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
			// lblNumero
			// 
			this.lblNumero.AutoSize = true;
			this.lblNumero.BackColor = System.Drawing.Color.Transparent;
			this.lblNumero.Location = new System.Drawing.Point(16, 8);
			this.lblNumero.Name = "lblNumero";
			this.lblNumero.Size = new System.Drawing.Size(63, 16);
			this.lblNumero.TabIndex = 236;
			this.lblNumero.Text = "Codigo [F3]";
			this.lblNumero.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.lblNumero.Click += new System.EventHandler(this.lblNumero_Click);
			// 
			// txtProducto
			// 
			appearance2.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtProducto.Appearance = appearance2;
			this.txtProducto.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtProducto.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtProducto.Location = new System.Drawing.Point(80, 8);
			this.txtProducto.Name = "txtProducto";
			this.txtProducto.Size = new System.Drawing.Size(139, 21);
			this.txtProducto.TabIndex = 322;
			this.txtProducto.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtProducto_KeyDown);
			this.txtProducto.Validated += new System.EventHandler(this.txtProducto_Validated);
			// 
			// frmDescripcionArticuloWeb
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(474, 216);
			this.Controls.Add(this.txtProducto);
			this.Controls.Add(this.txtCodigo);
			this.Controls.Add(this.label1);
			this.Controls.Add(this.btnGuardar);
			this.Controls.Add(this.btnCancelar);
			this.Controls.Add(this.btnSalir);
			this.Controls.Add(this.btnArticulo);
			this.Controls.Add(this.lblNumero);
			this.Controls.Add(this.txtObservaciones);
			this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
			this.Name = "frmDescripcionArticuloWeb";
			this.Text = "Descripción  Articulo Web";
			this.Load += new System.EventHandler(this.frmDescripcionArticuloWeb_Load);
			((System.ComponentModel.ISupportInitialize)(this.txtObservaciones)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCodigo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtProducto)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void UnloadMe()
		{
			this.Close();
		}

		private void frmDescripcionArticuloWeb_Load(object sender, System.EventArgs e)
		{
			miAcceso = new Acceso(cdsSeteoF, "200413");
			
			if (!Funcion.Permiso("2043", cdsSeteoF))
			{				
				MessageBox.Show("No cuenta con acceso a Descripcion Articulo Web.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				BeginInvoke(new MethodInvoker(UnloadMe));				

			}
		}

		private void lblNumero_Click(object sender, System.EventArgs e)
		{
		
		}

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);	
		}

		private void btnArticulo_Click(object sender, System.EventArgs e)
		{
			
		}

		private void txtProducto_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.txtObservaciones.Focus();

			if (e.KeyCode == Keys.F3)
			{
				using (BuscaArticulos miBusqueda = new BuscaArticulos(this.txtProducto.Text.ToString(), 0, 43, 9, DateTime.Today, false))
				{
					if (DialogResult.OK == miBusqueda.ShowDialog())
					{
						this.txtProducto.Text = miBusqueda.sCodigo;
						idArticulo = (int)miBusqueda.idArticulo;
						this.txtCodigo.Text =miBusqueda.sArticulo;
					}
				}
			}
		}

		private void txtProducto_Validated(object sender, System.EventArgs e)
		{
			if (this.txtProducto.Text.ToString().Length > 0)
			{
				string sSQL = string.Format("Select idArticulo, Codigo,Articulo,ISNULL(DescripcionWeb,'') AS DescripcionWeb From Articulo Where (Codigo = '{0}' Or Articulo = '{0}')", this.txtProducto.Text.ToString().Trim()); 
				SqlDataReader drArticulo = Funcion.rEscalarSQL(cdsSeteoF, sSQL, true);
				drArticulo.Read();
				if (drArticulo.HasRows)
				{
					idArticulo = drArticulo.GetInt32(0);
					this.txtProducto.Text = drArticulo.GetString(1);
					this.txtCodigo.Text = drArticulo.GetString(2);
					this.txtObservaciones.Text = drArticulo.GetString(3);
					drArticulo.Close();	

					this.btnGuardar.Enabled =true;
					this.btnCancelar.Enabled =true;
				}										
				else
				{
					MessageBox.Show("No existe el Producto", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
					idArticulo = 0;
				}
			}
		}

		private void btnGuardar_Click(object sender, System.EventArgs e)
		{
			#region Actualiza Descripcion Web
			if (idArticulo > 0)
			{
				if (DialogResult.Yes == MessageBox.Show("¿Está usted seguro que quiere guardar los cambios?", "Point Technology", 
					MessageBoxButtons.YesNo, MessageBoxIcon.Question))
				{					
					string sSQL= string.Format("Exec ActualizaDescripcionWeb {0},'{1}'",idArticulo, this.txtObservaciones.Text.ToString());
					Funcion.EjecutaSQL(cdsSeteoF, sSQL, true);

					this.txtProducto.Value="";
					this.txtCodigo.Value="";
					this.txtObservaciones.Value="";
					idArticulo=0;

				

					MessageBox.Show("Description Web Actualizado Correctamente", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
				}
			}
			#endregion Actualiza Descripcion Web
		}

		private void btnCancelar_Click(object sender, System.EventArgs e)
		{
			this.txtProducto.Value="";
			this.txtCodigo.Value="";
			this.txtObservaciones.Value="";
			idArticulo=0;
			this.btnGuardar.Enabled =false;
			this.btnCancelar.Enabled =false;
		}

		private void btnSalir_Click(object sender, System.EventArgs e)
		{
			this.Close();
		}
		
	}
}
