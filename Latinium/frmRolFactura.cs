using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de frmRolFactura.
	/// </summary>
	public class frmRolFactura : System.Windows.Forms.Form
	{
		private C1.Data.C1DataSet cdsSeteoF;
		private System.Windows.Forms.Button btnCancelar;
		private System.Windows.Forms.Button btnAceptar;
		public Infragistics.Win.UltraWinEditors.UltraTextEditor txtObservaciones;
		private System.Windows.Forms.Label lblNumero;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtNumero;
		public Infragistics.Win.UltraWinEditors.UltraNumericEditor txtNumIdCompra;
		private System.Windows.Forms.Button btnCliente;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public frmRolFactura()
		{
			//
			// Necesario para admitir el Diseñador de Windows Forms
			//
			InitializeComponent();

			//
			// TODO: agregar código de constructor después de llamar a InitializeComponent
			//
		}

		int idCompra = 0;
		string sObservaciones = "";
		public frmRolFactura(int IdCompra, string SObservaciones)
		{
			//
			// Necesario para admitir el Diseñador de Windows Forms
			//
			InitializeComponent();
			
			idCompra = IdCompra;
			sObservaciones = SObservaciones;
			
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
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.btnCancelar = new System.Windows.Forms.Button();
			this.btnAceptar = new System.Windows.Forms.Button();
			this.txtObservaciones = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.lblNumero = new System.Windows.Forms.Label();
			this.txtNumero = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtNumIdCompra = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.btnCliente = new System.Windows.Forms.Button();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtObservaciones)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNumero)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNumIdCompra)).BeginInit();
			this.SuspendLayout();
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
			// 
			// btnCancelar
			// 
			this.btnCancelar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnCancelar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnCancelar.Location = new System.Drawing.Point(89, 137);
			this.btnCancelar.Name = "btnCancelar";
			this.btnCancelar.Size = new System.Drawing.Size(75, 24);
			this.btnCancelar.TabIndex = 5;
			this.btnCancelar.Text = "Cancelar";
			this.btnCancelar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnCancelar.Click += new System.EventHandler(this.btnCancelar_Click);
			// 
			// btnAceptar
			// 
			this.btnAceptar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnAceptar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnAceptar.Location = new System.Drawing.Point(9, 137);
			this.btnAceptar.Name = "btnAceptar";
			this.btnAceptar.Size = new System.Drawing.Size(75, 24);
			this.btnAceptar.TabIndex = 4;
			this.btnAceptar.Text = "Aceptar";
			this.btnAceptar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnAceptar.Click += new System.EventHandler(this.btnAceptar_Click);
			// 
			// txtObservaciones
			// 
			this.txtObservaciones.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.txtObservaciones.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtObservaciones.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtObservaciones.Location = new System.Drawing.Point(9, 41);
			this.txtObservaciones.Multiline = true;
			this.txtObservaciones.Name = "txtObservaciones";
			this.txtObservaciones.Size = new System.Drawing.Size(440, 84);
			this.txtObservaciones.TabIndex = 3;
			// 
			// lblNumero
			// 
			this.lblNumero.AutoSize = true;
			this.lblNumero.BackColor = System.Drawing.Color.Transparent;
			this.lblNumero.Location = new System.Drawing.Point(9, 11);
			this.lblNumero.Name = "lblNumero";
			this.lblNumero.Size = new System.Drawing.Size(44, 17);
			this.lblNumero.TabIndex = 192;
			this.lblNumero.Text = "Número";
			this.lblNumero.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtNumero
			// 
			appearance1.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtNumero.Appearance = appearance1;
			this.txtNumero.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNumero.Enabled = false;
			this.txtNumero.Location = new System.Drawing.Point(72, 9);
			this.txtNumero.MaxLength = 9;
			this.txtNumero.Name = "txtNumero";
			this.txtNumero.Size = new System.Drawing.Size(128, 22);
			this.txtNumero.TabIndex = 191;
			// 
			// txtNumIdCompra
			// 
			this.txtNumIdCompra.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.txtNumIdCompra.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNumIdCompra.Enabled = false;
			this.txtNumIdCompra.Location = new System.Drawing.Point(24, -144);
			this.txtNumIdCompra.Name = "txtNumIdCompra";
			this.txtNumIdCompra.PromptChar = ' ';
			this.txtNumIdCompra.Size = new System.Drawing.Size(16, 22);
			this.txtNumIdCompra.TabIndex = 230;
			this.txtNumIdCompra.Visible = false;
			// 
			// btnCliente
			// 
			this.btnCliente.Location = new System.Drawing.Point(256, 8);
			this.btnCliente.Name = "btnCliente";
			this.btnCliente.Size = new System.Drawing.Size(24, 22);
			this.btnCliente.TabIndex = 231;
			this.btnCliente.Text = "...";
			this.btnCliente.Click += new System.EventHandler(this.btnCliente_Click);
			// 
			// frmRolFactura
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 14);
			this.ClientSize = new System.Drawing.Size(458, 168);
			this.ControlBox = false;
			this.Controls.Add(this.btnCliente);
			this.Controls.Add(this.txtNumIdCompra);
			this.Controls.Add(this.lblNumero);
			this.Controls.Add(this.txtNumero);
			this.Controls.Add(this.btnCancelar);
			this.Controls.Add(this.btnAceptar);
			this.Controls.Add(this.txtObservaciones);
			this.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
			this.KeyPreview = true;
			this.Name = "frmRolFactura";
			this.ShowInTaskbar = false;
			this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
			this.Text = "Factura";
			this.Load += new System.EventHandler(this.frmRolFactura_Load);
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtObservaciones)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNumero)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNumIdCompra)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void btnCliente_Click(object sender, System.EventArgs e)
		{
			using (frmBuscaFacturas Busqueda = new frmBuscaFacturas(1, 0, false))
			{
				if (DialogResult.OK == Busqueda.ShowDialog())
				{
					this.txtNumIdCompra.Value = (int)Busqueda.grdFacturas.ActiveRow.Cells["idCompra"].Value;
					this.txtNumero.Text = Busqueda.grdFacturas.ActiveRow.Cells["Numero"].Value.ToString(); 
				}
			}
		}

		private void frmRolFactura_Load(object sender, System.EventArgs e)
		{
			if (idCompra > 0)
			{
				this.txtNumIdCompra.Value = idCompra;
				this.txtNumero.Text = Funcion.sEscalarSQL(cdsSeteoF, string.Format("Select b.Codigo + '-' + Numero From Compra c Inner Join Bodega b On b.Bodega = c.Bodega Where idCompra = {0}", idCompra));
				this.txtObservaciones.Text = sObservaciones;
			}
		}

		private void btnCancelar_Click(object sender, System.EventArgs e)
		{
			this.DialogResult = DialogResult.No;
		}

		private void btnAceptar_Click(object sender, System.EventArgs e)
		{
			if ((int)this.txtNumIdCompra.Value == 0)
			{
				MessageBox.Show("Seleccione una factura", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.txtObservaciones.Focus();
				return;
			}

			if (this.txtObservaciones.Text.ToString().Trim().Length == 0)
			{
				MessageBox.Show("Ingrese una Observación", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.txtObservaciones.Focus();
				return;
			}

			this.DialogResult = DialogResult.OK;
		}
	}
}
