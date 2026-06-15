using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de frmVentaClaveDescuento.
	/// </summary>
	public class frmVentaClaveDescuento : System.Windows.Forms.Form
	{
		private System.Windows.Forms.Label lblCliente;
		private System.Windows.Forms.Button btnNuevo;
		private System.Windows.Forms.Button btnCancelar;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtClave;
		private C1.Data.C1DataSet cdsSeteoF;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public frmVentaClaveDescuento()
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
			Infragistics.Win.Appearance appearance1 = new Infragistics.Win.Appearance();
			this.lblCliente = new System.Windows.Forms.Label();
			this.txtClave = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.btnNuevo = new System.Windows.Forms.Button();
			this.btnCancelar = new System.Windows.Forms.Button();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			((System.ComponentModel.ISupportInitialize)(this.txtClave)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			this.SuspendLayout();
			// 
			// lblCliente
			// 
			this.lblCliente.AutoSize = true;
			this.lblCliente.Location = new System.Drawing.Point(8, 8);
			this.lblCliente.Name = "lblCliente";
			this.lblCliente.Size = new System.Drawing.Size(86, 16);
			this.lblCliente.TabIndex = 162;
			this.lblCliente.Text = "Ingrese la Clave";
			// 
			// txtClave
			// 
			this.txtClave.AcceptsTab = true;
			appearance1.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtClave.Appearance = appearance1;
			this.txtClave.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtClave.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtClave.Location = new System.Drawing.Point(8, 32);
			this.txtClave.Name = "txtClave";
			this.txtClave.PasswordChar = '*';
			this.txtClave.Size = new System.Drawing.Size(160, 21);
			this.txtClave.TabIndex = 161;
			this.txtClave.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtClave_KeyDown);
			// 
			// btnNuevo
			// 
			this.btnNuevo.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnNuevo.CausesValidation = false;
			this.btnNuevo.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnNuevo.Location = new System.Drawing.Point(8, 64);
			this.btnNuevo.Name = "btnNuevo";
			this.btnNuevo.TabIndex = 163;
			this.btnNuevo.Text = "&Aceptar";
			this.btnNuevo.Click += new System.EventHandler(this.btnNuevo_Click);
			// 
			// btnCancelar
			// 
			this.btnCancelar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnCancelar.CausesValidation = false;
			this.btnCancelar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnCancelar.Location = new System.Drawing.Point(88, 64);
			this.btnCancelar.Name = "btnCancelar";
			this.btnCancelar.TabIndex = 164;
			this.btnCancelar.Text = "&Cancelar";
			this.btnCancelar.Click += new System.EventHandler(this.btnCancelar_Click);
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
			// frmVentaClaveDescuento
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(178, 96);
			this.Controls.Add(this.btnNuevo);
			this.Controls.Add(this.btnCancelar);
			this.Controls.Add(this.lblCliente);
			this.Controls.Add(this.txtClave);
			this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
			this.KeyPreview = true;
			this.MaximizeBox = false;
			this.MinimizeBox = false;
			this.Name = "frmVentaClaveDescuento";
			this.ShowInTaskbar = false;
			this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
			((System.ComponentModel.ISupportInitialize)(this.txtClave)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void btnCancelar_Click(object sender, System.EventArgs e)
		{
			this.DialogResult = DialogResult.No;
		}

		int iConteo = 1;		

		private void btnNuevo_Click(object sender, System.EventArgs e)
		{
			if (this.txtClave.Text.ToString().Length == 0)
			{
				MessageBox.Show("Ingrese la Clave", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);

				iConteo++;
			}
			else
			{
				string sSQL = string.Format("Exec ValidaClaveDescuento '{0}', {1}", this.txtClave.Text.ToString(), MenuLatinium.IdUsuario);
				bool bRes = Funcion.bEscalarSQL(cdsSeteoF, sSQL, true);
	
				if (!bRes)
				{
					MessageBox.Show("Contraseña Incorrecta", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);

					iConteo++;

					if (iConteo == 3) this.DialogResult = DialogResult.No;
				}
				else 
				{
					this.DialogResult = DialogResult.OK;
				}
			}			
		}

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);	
		}

		private void txtClave_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.btnNuevo.Focus();
		}
	}
}
