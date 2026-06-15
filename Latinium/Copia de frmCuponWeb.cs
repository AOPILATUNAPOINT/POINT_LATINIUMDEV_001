using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de frmCuponWeb.
	/// </summary>
	public class frmCuponWeb : System.Windows.Forms.Form
	{
		private System.Windows.Forms.Button btnCancelar;
		private System.Windows.Forms.Label lblCliente;
		private System.Windows.Forms.GroupBox groupBox1;
		private C1.Data.C1DataSet cdsSeteoF;
		public Infragistics.Win.UltraWinEditors.UltraTextEditor txtCupon;
		private System.Windows.Forms.Button btnAceptar;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public frmCuponWeb()
		{
			//
			// Necesario para admitir el Diseñador de Windows Forms
			//
			InitializeComponent();

			//
			// TODO: agregar código de constructor después de llamar a InitializeComponent
			//
		}

		int idArticulo = 0;
		bool bCuponPromocional = false;

		public frmCuponWeb(int IdArticulo, bool BCuponPromocional)
		{
			//
			// Necesario para admitir el Diseñador de Windows Forms
			//
			InitializeComponent();

			idArticulo = IdArticulo;
			bCuponPromocional = BCuponPromocional;
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
			this.txtCupon = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.btnAceptar = new System.Windows.Forms.Button();
			this.btnCancelar = new System.Windows.Forms.Button();
			this.lblCliente = new System.Windows.Forms.Label();
			this.groupBox1 = new System.Windows.Forms.GroupBox();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			((System.ComponentModel.ISupportInitialize)(this.txtCupon)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			this.SuspendLayout();
			// 
			// txtCupon
			// 
			appearance1.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtCupon.Appearance = appearance1;
			this.txtCupon.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtCupon.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtCupon.Location = new System.Drawing.Point(56, 8);
			this.txtCupon.MaxLength = 15;
			this.txtCupon.Name = "txtCupon";
			this.txtCupon.Size = new System.Drawing.Size(288, 22);
			this.txtCupon.TabIndex = 0;
			this.txtCupon.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtCupon_KeyPress);
			// 
			// btnAceptar
			// 
			this.btnAceptar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnAceptar.CausesValidation = false;
			this.btnAceptar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnAceptar.Location = new System.Drawing.Point(8, 56);
			this.btnAceptar.Name = "btnAceptar";
			this.btnAceptar.TabIndex = 1;
			this.btnAceptar.Text = "&Aceptar";
			this.btnAceptar.Click += new System.EventHandler(this.btnAceptar_Click);
			// 
			// btnCancelar
			// 
			this.btnCancelar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnCancelar.CausesValidation = false;
			this.btnCancelar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnCancelar.Location = new System.Drawing.Point(88, 56);
			this.btnCancelar.Name = "btnCancelar";
			this.btnCancelar.TabIndex = 2;
			this.btnCancelar.Text = "&Cancelar";
			this.btnCancelar.Click += new System.EventHandler(this.btnCancelar_Click);
			// 
			// lblCliente
			// 
			this.lblCliente.AutoSize = true;
			this.lblCliente.Location = new System.Drawing.Point(8, 10);
			this.lblCliente.Name = "lblCliente";
			this.lblCliente.Size = new System.Drawing.Size(36, 17);
			this.lblCliente.TabIndex = 165;
			this.lblCliente.Text = "Cupón";
			// 
			// groupBox1
			// 
			this.groupBox1.Location = new System.Drawing.Point(0, 40);
			this.groupBox1.Name = "groupBox1";
			this.groupBox1.Size = new System.Drawing.Size(352, 8);
			this.groupBox1.TabIndex = 168;
			this.groupBox1.TabStop = false;
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
			// frmCuponWeb
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 14);
			this.ClientSize = new System.Drawing.Size(354, 86);
			this.ControlBox = false;
			this.Controls.Add(this.groupBox1);
			this.Controls.Add(this.btnAceptar);
			this.Controls.Add(this.btnCancelar);
			this.Controls.Add(this.lblCliente);
			this.Controls.Add(this.txtCupon);
			this.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
			this.KeyPreview = true;
			this.Name = "frmCuponWeb";
			this.ShowInTaskbar = false;
			this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
			this.Text = "Cupón de Descuento Para Ventas Digitales";
			this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.frmCuponWeb_KeyDown);
			this.Load += new System.EventHandler(this.frmCuponWeb_Load);
			((System.ComponentModel.ISupportInitialize)(this.txtCupon)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		public int idCuponWeb = 0;

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);
		}

		private void frmCuponWeb_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Escape) this.DialogResult = DialogResult.No;
		}

		private void btnCancelar_Click(object sender, System.EventArgs e)
		{
			this.DialogResult = DialogResult.No;
		}

		private void btnAceptar_Click(object sender, System.EventArgs e)
		{
			if (!Validacion.vbTexto(this.txtCupon, 6, 20, "Cupon")) return;

			string sMensaje = "El Cupón no existe";
			int iResValid = 0;

			if (!bCuponPromocional)
			{
				iResValid = Funcion.iEscalarSQL(cdsSeteoF, string.Format("Exec FacturacionValidaCuponWeb '{0}'", this.txtCupon.Text.ToString().Trim()));

				if (iResValid == 1) sMensaje = "El Cupón esta siendo utilizado en otro documento";
				if (iResValid == 2) sMensaje = "El Cupón ya fue utilizado en otro documento";
			}

			if (iResValid != 0)
			{
				MessageBox.Show(sMensaje, "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.txtCupon.Focus();
				return;
			}

			if (bCuponPromocional)
			{
				if (Funcion.iEscalarSQL(cdsSeteoF, string.Format("Select COUNT(*) From Articulo Where Cupon = '{0}' And idArticulo = {1}", this.txtCupon.Text.ToString(), idArticulo)) == 0)
				{
					MessageBox.Show("El cupon registrado no existe", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					this.txtCupon.Focus();
					return;
				}
			}
			
			if (!bCuponPromocional)
				idCuponWeb = Funcion.iEscalarSQL(cdsSeteoF, string.Format("Select idCuponesWeb From CuponesWeb Where Estado = 0 And Cupon = '{0}'", this.txtCupon.Text.ToString().Trim()));
			
			this.DialogResult = DialogResult.OK;
		}

		private void txtCupon_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (e.KeyChar == 13) this.btnAceptar_Click(sender, e);
		}

		private void frmCuponWeb_Load(object sender, System.EventArgs e)
		{
		
		}
	}
}
