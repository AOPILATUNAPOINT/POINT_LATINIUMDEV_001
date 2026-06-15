using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de frmAvisoFactura.
	/// </summary>
	public class frmAvisoFactura : System.Windows.Forms.Form
	{
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public frmAvisoFactura()
		{
			//
			// Necesario para admitir el Diseñador de Windows Forms
			//
			InitializeComponent();

			//
			// TODO: agregar código de constructor después de llamar a InitializeComponent
			//
		}

		private System.Windows.Forms.Label lblEstado;
		private System.Windows.Forms.Button btnAceptar;
		private System.Windows.Forms.Button btnCancelar;
		private C1.Data.C1DataSet cdsSeteoF;
		private System.Windows.Forms.Label lblEstado2;
		private System.Windows.Forms.Label lblEstado3;
		private System.Windows.Forms.Label lblEstado4;
		int iCompra = 0;
		public frmAvisoFactura(int ICompra )
		{
			//
			// Necesario para admitir el Diseñador de Windows Forms
			//
			InitializeComponent();
      iCompra =  ICompra ;
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
			this.lblEstado = new System.Windows.Forms.Label();
			this.btnAceptar = new System.Windows.Forms.Button();
			this.btnCancelar = new System.Windows.Forms.Button();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.lblEstado2 = new System.Windows.Forms.Label();
			this.lblEstado3 = new System.Windows.Forms.Label();
			this.lblEstado4 = new System.Windows.Forms.Label();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			this.SuspendLayout();
			// 
			// lblEstado
			// 
			this.lblEstado.Anchor = System.Windows.Forms.AnchorStyles.Left;
			this.lblEstado.AutoSize = true;
			this.lblEstado.Font = new System.Drawing.Font("Microsoft Sans Serif", 25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblEstado.ForeColor = System.Drawing.Color.Firebrick;
			this.lblEstado.Location = new System.Drawing.Point(16, 8);
			this.lblEstado.Name = "lblEstado";
			this.lblEstado.Size = new System.Drawing.Size(941, 41);
			this.lblEstado.TabIndex = 274;
			this.lblEstado.Text = "ESTA SEGURO DE IMPRIMIR LOS DATOS DEL SR/SRA:";
			this.lblEstado.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// btnAceptar
			// 
			this.btnAceptar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnAceptar.Location = new System.Drawing.Point(8, 240);
			this.btnAceptar.Name = "btnAceptar";
			this.btnAceptar.Size = new System.Drawing.Size(88, 23);
			this.btnAceptar.TabIndex = 275;
			this.btnAceptar.Text = "&Aceptar";
			this.btnAceptar.Click += new System.EventHandler(this.btnAceptar_Click);
			// 
			// btnCancelar
			// 
			this.btnCancelar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnCancelar.Location = new System.Drawing.Point(104, 240);
			this.btnCancelar.Name = "btnCancelar";
			this.btnCancelar.Size = new System.Drawing.Size(88, 23);
			this.btnCancelar.TabIndex = 276;
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
			// lblEstado2
			// 
			this.lblEstado2.Anchor = System.Windows.Forms.AnchorStyles.Left;
			this.lblEstado2.AutoSize = true;
			this.lblEstado2.Font = new System.Drawing.Font("Microsoft Sans Serif", 25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblEstado2.ForeColor = System.Drawing.Color.Firebrick;
			this.lblEstado2.Location = new System.Drawing.Point(16, 64);
			this.lblEstado2.Name = "lblEstado2";
			this.lblEstado2.Size = new System.Drawing.Size(0, 41);
			this.lblEstado2.TabIndex = 277;
			this.lblEstado2.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// lblEstado3
			// 
			this.lblEstado3.Anchor = System.Windows.Forms.AnchorStyles.Left;
			this.lblEstado3.AutoSize = true;
			this.lblEstado3.Font = new System.Drawing.Font("Microsoft Sans Serif", 25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblEstado3.ForeColor = System.Drawing.Color.Firebrick;
			this.lblEstado3.Location = new System.Drawing.Point(16, 120);
			this.lblEstado3.Name = "lblEstado3";
			this.lblEstado3.Size = new System.Drawing.Size(0, 41);
			this.lblEstado3.TabIndex = 278;
			this.lblEstado3.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// lblEstado4
			// 
			this.lblEstado4.Anchor = System.Windows.Forms.AnchorStyles.Left;
			this.lblEstado4.AutoSize = true;
			this.lblEstado4.Font = new System.Drawing.Font("Microsoft Sans Serif", 25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblEstado4.ForeColor = System.Drawing.Color.Firebrick;
			this.lblEstado4.Location = new System.Drawing.Point(16, 184);
			this.lblEstado4.Name = "lblEstado4";
			this.lblEstado4.Size = new System.Drawing.Size(0, 41);
			this.lblEstado4.TabIndex = 279;
			this.lblEstado4.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// frmAvisoFactura
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(928, 270);
			this.ControlBox = false;
			this.Controls.Add(this.lblEstado4);
			this.Controls.Add(this.lblEstado3);
			this.Controls.Add(this.lblEstado2);
			this.Controls.Add(this.btnAceptar);
			this.Controls.Add(this.btnCancelar);
			this.Controls.Add(this.lblEstado);
			this.KeyPreview = true;
			this.Name = "frmAvisoFactura";
			this.Text = "AVISO";
			this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.frmAvisoFactura_KeyDown);
			this.Load += new System.EventHandler(this.frmAvisoFactura_Load);
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void frmAvisoFactura_Load(object sender, System.EventArgs e)
		{
			string SRSRNombreuc = Funcion.sEscalarSQL(cdsSeteoF, string.Format(" Select cl.nombre  from COMPRA c Inner Join Cliente cl on cl.idcliente  =  c.idcliente where idcompra = {0}", iCompra));	
			string sruc = Funcion.sEscalarSQL(cdsSeteoF, string.Format(" Select RUC from COMPRA  where idcompra = {0}", iCompra));	
			string snUM = Funcion.sEscalarSQL(cdsSeteoF, string.Format(" Select NUMERO from COMPRA  where idcompra = {0}", iCompra));	

			this.lblEstado2.Text = SRSRNombreuc ; 
			this.lblEstado3.Text = "CON CI:"+ sruc ; 
			this.lblEstado4.Text = "CON NUMERO DE FACTURA" + snUM ; 
		}

		private void btnAceptar_Click(object sender, System.EventArgs e)
		{
				this.DialogResult = DialogResult.OK;
		}

		private void btnCancelar_Click(object sender, System.EventArgs e)
		{
			this.DialogResult = DialogResult.No;
		}

		private void frmAvisoFactura_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
					if (e.KeyCode == Keys.Escape) this.DialogResult = DialogResult.No;
		}

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);		
		}
	}
}
