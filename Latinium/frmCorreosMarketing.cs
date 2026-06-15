using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de frmCorreosMarketing.
	/// </summary>
	public class frmCorreosMarketing : System.Windows.Forms.Form
	{
		private System.Windows.Forms.Label label3;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtCorreo;
		private System.Windows.Forms.Button btnCliente;
		private System.Windows.Forms.Button button1;
		private C1.Data.C1DataSet cdsSeteoF;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public frmCorreosMarketing()
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
			this.label3 = new System.Windows.Forms.Label();
			this.txtCorreo = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.btnCliente = new System.Windows.Forms.Button();
			this.button1 = new System.Windows.Forms.Button();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			((System.ComponentModel.ISupportInitialize)(this.txtCorreo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			this.SuspendLayout();
			// 
			// label3
			// 
			this.label3.AutoSize = true;
			this.label3.Location = new System.Drawing.Point(8, 10);
			this.label3.Name = "label3";
			this.label3.Size = new System.Drawing.Size(39, 16);
			this.label3.TabIndex = 246;
			this.label3.Text = "Correo";
			this.label3.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtCorreo
			// 
			appearance1.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtCorreo.Appearance = appearance1;
			this.txtCorreo.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtCorreo.Location = new System.Drawing.Point(64, 8);
			this.txtCorreo.Name = "txtCorreo";
			this.txtCorreo.Size = new System.Drawing.Size(232, 21);
			this.txtCorreo.TabIndex = 245;
			this.txtCorreo.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtCorreo_KeyPress);
			// 
			// btnCliente
			// 
			this.btnCliente.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Underline, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.btnCliente.Location = new System.Drawing.Point(64, 40);
			this.btnCliente.Name = "btnCliente";
			this.btnCliente.Size = new System.Drawing.Size(64, 22);
			this.btnCliente.TabIndex = 247;
			this.btnCliente.Text = "Aceptar ";
			this.btnCliente.Click += new System.EventHandler(this.btnCliente_Click);
			// 
			// button1
			// 
			this.button1.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Underline, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.button1.Location = new System.Drawing.Point(144, 40);
			this.button1.Name = "button1";
			this.button1.Size = new System.Drawing.Size(72, 22);
			this.button1.TabIndex = 248;
			this.button1.Text = "Cancelar ";
			this.button1.Click += new System.EventHandler(this.button1_Click_1);
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
			// frmCorreosMarketing
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(304, 70);
			this.Controls.Add(this.button1);
			this.Controls.Add(this.btnCliente);
			this.Controls.Add(this.label3);
			this.Controls.Add(this.txtCorreo);
			this.KeyPreview = true;
			this.MaximizeBox = false;
			this.MinimizeBox = false;
			this.Name = "frmCorreosMarketing";
			this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
			this.Text = "Correo";
			this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.frmCorreosMarketing_KeyDown);
			((System.ComponentModel.ISupportInitialize)(this.txtCorreo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion
    
		public int  idCorreo = 0 ;
		int idCompra;
		private void button1_Click(object sender, System.EventArgs e)
		{
		
		}

		private void btnRegistrar_Click(object sender, System.EventArgs e)
		{
		
		}

		private void btnCliente_Click(object sender, System.EventArgs e)
		{
		 string sNumero;
			
			idCorreo = Funcion.iEscalarSQL(cdsSeteoF, string.Format(" If Not Exists(Select idCorreosMarketing From ListaCorreosMarketing Where Email = '{0}') Select CONVERT(Int , -1) Else If Exists(Select idCorreosMarketing From ListaCorreosMarketing Where Email = '{0}' ) Select idCorreosMarketing From ListaCorreosMarketing Where Email = '{0}'",
				this.txtCorreo.Text.ToString()));		
			
			if (idCorreo < 0 )
			{
				MessageBox.Show("El Correo ingresado no existe", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.txtCorreo.Focus();
			
				return;
			} 
			
				idCompra  = Funcion.iEscalarSQL(cdsSeteoF, string.Format(" If Exists(Select idcompra From ListaCorreosMarketing Where Email = '{0}' ) Select idcompra From ListaCorreosMarketing Where Email = '{0}'else Select Convert(int, 0)",this.txtCorreo.Text.ToString()));	
					
				if ( idCompra > 0)

				   {
              
					string  ssql = Funcion.sEscalarSQL(cdsSeteoF,string.Format("Select Numero from Compra where idCompra = {0}",idCompra));
					sNumero = ssql;
			
					MessageBox.Show("El Correo ingresado ya fue utilizado \n en la factura número " + sNumero+ "", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					this.txtCorreo.Focus();
					return;
				} 
					
	this.DialogResult = DialogResult.OK;
	}

	private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);
		}

		private void txtCorreo_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (e.KeyChar == 13) this.btnCliente.Focus();
		}

		private void button1_Click_1(object sender, System.EventArgs e)
		{
					this.Close();
		}

		private void frmCorreosMarketing_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Escape) this.Close();			
		}
	}
}
