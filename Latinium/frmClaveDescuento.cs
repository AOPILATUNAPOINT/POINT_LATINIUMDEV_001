using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using DevExpress.XtraEditors;

namespace Latinium
{
	/// <summary>
	/// Summary description for frmClaveDescuento.
	/// </summary>
	public class frmClaveDescuento : DevExpress.XtraEditors.XtraForm
	{
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtRepitaClave;
		private System.Windows.Forms.Label label4;
		private System.Windows.Forms.Label label3;
		private System.Windows.Forms.Label label2;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtClaveNueva;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtClaveActual;
		private Infragistics.Win.Misc.UltraButton btCancelar;
		private Infragistics.Win.Misc.UltraButton btAceptar;
		private C1.Data.C1DataSet cdsSeteoF;
		/// <summary>
		/// Required designer variable.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public frmClaveDescuento()
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
			this.txtRepitaClave = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label4 = new System.Windows.Forms.Label();
			this.label3 = new System.Windows.Forms.Label();
			this.label2 = new System.Windows.Forms.Label();
			this.txtClaveNueva = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtClaveActual = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.btCancelar = new Infragistics.Win.Misc.UltraButton();
			this.btAceptar = new Infragistics.Win.Misc.UltraButton();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			((System.ComponentModel.ISupportInitialize)(this.txtRepitaClave)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtClaveNueva)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtClaveActual)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			this.SuspendLayout();
			// 
			// txtRepitaClave
			// 
			this.txtRepitaClave.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtRepitaClave.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtRepitaClave.Location = new System.Drawing.Point(144, 72);
			this.txtRepitaClave.Name = "txtRepitaClave";
			this.txtRepitaClave.PasswordChar = '*';
			this.txtRepitaClave.Size = new System.Drawing.Size(208, 22);
			this.txtRepitaClave.TabIndex = 2;
			this.txtRepitaClave.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtRepitaClave_KeyDown);
			this.txtRepitaClave.Validating += new System.ComponentModel.CancelEventHandler(this.txtRepitaClave_Validating);
			// 
			// label4
			// 
			this.label4.AutoSize = true;
			this.label4.Location = new System.Drawing.Point(8, 72);
			this.label4.Name = "label4";
			this.label4.Size = new System.Drawing.Size(109, 16);
			this.label4.TabIndex = 19;
			this.label4.Text = "Confirme Contraseña";
			this.label4.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label3
			// 
			this.label3.AutoSize = true;
			this.label3.Location = new System.Drawing.Point(8, 40);
			this.label3.Name = "label3";
			this.label3.Size = new System.Drawing.Size(96, 16);
			this.label3.TabIndex = 18;
			this.label3.Text = "Nueva Contraseña";
			this.label3.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label2
			// 
			this.label2.AutoSize = true;
			this.label2.Location = new System.Drawing.Point(8, 8);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(95, 16);
			this.label2.TabIndex = 16;
			this.label2.Text = "Contraseña Actual";
			this.label2.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtClaveNueva
			// 
			this.txtClaveNueva.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtClaveNueva.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtClaveNueva.Location = new System.Drawing.Point(144, 40);
			this.txtClaveNueva.Name = "txtClaveNueva";
			this.txtClaveNueva.PasswordChar = '*';
			this.txtClaveNueva.Size = new System.Drawing.Size(208, 22);
			this.txtClaveNueva.TabIndex = 1;
			this.txtClaveNueva.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtClaveNueva_KeyDown);
			this.txtClaveNueva.Validating += new System.ComponentModel.CancelEventHandler(this.txtClaveNueva_Validating);
			this.txtClaveNueva.ValueChanged += new System.EventHandler(this.txtClaveNueva_ValueChanged);
			// 
			// txtClaveActual
			// 
			this.txtClaveActual.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtClaveActual.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtClaveActual.Location = new System.Drawing.Point(144, 8);
			this.txtClaveActual.Name = "txtClaveActual";
			this.txtClaveActual.PasswordChar = '*';
			this.txtClaveActual.Size = new System.Drawing.Size(208, 22);
			this.txtClaveActual.TabIndex = 0;
			this.txtClaveActual.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtClaveActual_KeyDown);
			this.txtClaveActual.Validating += new System.ComponentModel.CancelEventHandler(this.txtClaveActual_Validating);
			this.txtClaveActual.ValueChanged += new System.EventHandler(this.txtClaveActual_ValueChanged);
			// 
			// btCancelar
			// 
			this.btCancelar.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.btCancelar.CausesValidation = false;
			this.btCancelar.Location = new System.Drawing.Point(88, 104);
			this.btCancelar.Name = "btCancelar";
			this.btCancelar.Size = new System.Drawing.Size(75, 25);
			this.btCancelar.TabIndex = 4;
			this.btCancelar.Text = "&Cancelar";
			this.btCancelar.Click += new System.EventHandler(this.btCancelar_Click);
			// 
			// btAceptar
			// 
			this.btAceptar.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.btAceptar.Location = new System.Drawing.Point(8, 104);
			this.btAceptar.Name = "btAceptar";
			this.btAceptar.Size = new System.Drawing.Size(75, 25);
			this.btAceptar.TabIndex = 3;
			this.btAceptar.Text = "&Aceptar";
			this.btAceptar.Click += new System.EventHandler(this.btAceptar_Click);
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
			// frmClaveDescuento
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 14);
			this.CausesValidation = false;
			this.ClientSize = new System.Drawing.Size(360, 134);
			this.Controls.Add(this.txtRepitaClave);
			this.Controls.Add(this.label4);
			this.Controls.Add(this.label3);
			this.Controls.Add(this.label2);
			this.Controls.Add(this.txtClaveNueva);
			this.Controls.Add(this.txtClaveActual);
			this.Controls.Add(this.btCancelar);
			this.Controls.Add(this.btAceptar);
			this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
			this.MaximizeBox = false;
			this.Name = "frmClaveDescuento";
			this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
			this.Text = "Clave Descuento";
			this.Load += new System.EventHandler(this.frmClaveDescuento_Load);
			((System.ComponentModel.ISupportInitialize)(this.txtRepitaClave)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtClaveNueva)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtClaveActual)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void txtClaveActual_Validating(object sender, System.ComponentModel.CancelEventArgs e)
		{
			if (Funcion.iEscalarSQL(cdsSeteoF, string.Format("Select Count(*) From ClaveDescuento Where idUsuario = {0}", MenuLatinium.IdUsuario)) > 0)
			{
				if (this.txtClaveActual.Text.ToString().Length == 0)
				{
					MessageBox.Show("Ingrese la Clave Actual", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					this.txtClaveActual.Focus();
					e.Cancel = true;				
				}
				else 
				{
				
					string sSQL = string.Format("Exec ValidaClaveDescuento '{0}', {1}", this.txtClaveActual.Text.ToString(), MenuLatinium.IdUsuario);
					bool bRes = Funcion.bEscalarSQL(cdsSeteoF, sSQL, true);

					if (!bRes)
					{
						MessageBox.Show("La Clave no existe", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
						this.txtClaveActual.Focus();
						e.Cancel = true;
					}
				}
			}
		}

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);	
		}

		private void txtClaveNueva_Validating(object sender, System.ComponentModel.CancelEventArgs e)
		{
			if (this.txtClaveNueva.Text.ToString().Length == 0)
			{
				MessageBox.Show("Ingrese la Nueva Clave", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.txtClaveNueva.Focus();
				e.Cancel = true;				
			}
			else 
			{
				string sSQL = string.Format("Select Count(*) From ClaveDescuento Where Clave = '{0}' And idUsuario = {1}", this.txtClaveNueva.Text.ToString(), MenuLatinium.IdUsuario);
				int iRes = Funcion.iEscalarSQL(cdsSeteoF, sSQL, true);

				if (iRes > 0)
				{
					MessageBox.Show("La Clave ya fue usada anteriormente, Ingrese otra Clave", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					this.txtClaveNueva.Focus();
					e.Cancel = true;
				}
			}
		}

		private void txtRepitaClave_Validating(object sender, System.ComponentModel.CancelEventArgs e)
		{
			if (this.txtRepitaClave.Text.ToString().Length == 0)
			{
				MessageBox.Show("Repita su Nueva Clave", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.txtRepitaClave.Focus();
				e.Cancel = true;				
			}
			else if (this.txtClaveNueva.Text.ToString().Trim() != this.txtRepitaClave.Text.ToString().Trim())
			{
				MessageBox.Show("No Coincide su Nueva Clave", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.txtRepitaClave.Focus();
				e.Cancel = true;				
			}
		}

		private void txtClaveActual_ValueChanged(object sender, System.EventArgs e)
		{
		
		}

		private void txtClaveNueva_ValueChanged(object sender, System.EventArgs e)
		{
		
		}

		private void btCancelar_Click(object sender, System.EventArgs e)
		{
			this.Close();
		}

		private void btAceptar_Click(object sender, System.EventArgs e)
		{
			string sSQL = string.Format("Exec GuardaClaveDescuento '{0}', {1}", this.txtClaveNueva.Text.ToString(), MenuLatinium.IdUsuario);
			Funcion.EjecutaSQL(cdsSeteoF, sSQL);

			MessageBox.Show("Clave Registrada", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);

			this.Close();
		}

		private void txtClaveActual_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.txtClaveNueva.Focus();
		}

		private void txtClaveNueva_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.txtRepitaClave.Focus();
		}

		private void txtRepitaClave_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.btAceptar.Focus();
		}

		private void frmClaveDescuento_Load(object sender, System.EventArgs e)
		{
		
		}
	}
}

