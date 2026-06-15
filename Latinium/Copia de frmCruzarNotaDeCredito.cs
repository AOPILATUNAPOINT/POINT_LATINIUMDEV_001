using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Data.SqlClient;
using System.Windows.Forms;
using C1.Data;

namespace Latinium
{
	/// <summary>
	/// Summary description for frmCruzarNotaDeCredito.
	/// </summary>
	public class frmCruzarNotaDeCredito : DevExpress.XtraEditors.XtraForm
	{
		private System.Windows.Forms.Label label3;
		private System.Windows.Forms.Label label1;
		private System.Windows.Forms.Label label2;
		public Infragistics.Win.UltraWinEditors.UltraNumericEditor txtSaldo;
		public Infragistics.Win.UltraWinEditors.UltraTextEditor txtNumero;
		/// <summary>
		/// Required designer variable.
		/// </summary>
		private System.ComponentModel.Container components = null;
		private C1.Data.C1DataSet cdsSeteoF;

		int IdCliente = 0;
		public int IdCompra = 0;
		string SNumero = "";
		decimal DValor = 0.00m;
		decimal DSaldo = 0.00m;
		private System.Windows.Forms.Button btnAceptar;
		private System.Windows.Forms.Button btnCancelar;
		private System.Windows.Forms.GroupBox groupBox1;
		public Infragistics.Win.UltraWinEditors.UltraNumericEditor txtValor;
		decimal DValorFact = 0.00m;

		public frmCruzarNotaDeCredito(int idCliente, int idCompra, string sNumero, decimal dValor, decimal dSaldo, decimal dValorFact)
		{
			//
			// Required for Windows Form Designer support
			//
			InitializeComponent();
			IdCliente = idCliente;
			IdCompra = idCompra;
			SNumero = sNumero;
			DValor = dValor;
			DSaldo = dSaldo;
			DValorFact = dValorFact;

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
			Infragistics.Win.Appearance appearance1 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			this.label3 = new System.Windows.Forms.Label();
			this.txtNumero = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label1 = new System.Windows.Forms.Label();
			this.label2 = new System.Windows.Forms.Label();
			this.txtSaldo = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.btnAceptar = new System.Windows.Forms.Button();
			this.btnCancelar = new System.Windows.Forms.Button();
			this.groupBox1 = new System.Windows.Forms.GroupBox();
			this.txtValor = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			((System.ComponentModel.ISupportInitialize)(this.txtNumero)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtSaldo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtValor)).BeginInit();
			this.SuspendLayout();
			// 
			// label3
			// 
			this.label3.AutoSize = true;
			this.label3.Location = new System.Drawing.Point(12, 11);
			this.label3.Name = "label3";
			this.label3.Size = new System.Drawing.Size(44, 16);
			this.label3.TabIndex = 179;
			this.label3.Text = "Número";
			this.label3.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtNumero
			// 
			appearance1.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtNumero.Appearance = appearance1;
			this.txtNumero.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtNumero.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNumero.Location = new System.Drawing.Point(76, 8);
			this.txtNumero.MaxLength = 9;
			this.txtNumero.Name = "txtNumero";
			this.txtNumero.Size = new System.Drawing.Size(136, 22);
			this.txtNumero.TabIndex = 0;
			this.txtNumero.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtNumero_KeyPress);
			this.txtNumero.Validating += new System.ComponentModel.CancelEventHandler(this.txtNumero_Validating);
			this.txtNumero.Validated += new System.EventHandler(this.txtNumero_Validated);
			// 
			// label1
			// 
			this.label1.AutoSize = true;
			this.label1.Location = new System.Drawing.Point(12, 43);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(30, 16);
			this.label1.TabIndex = 182;
			this.label1.Text = "Valor";
			this.label1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label2
			// 
			this.label2.AutoSize = true;
			this.label2.Location = new System.Drawing.Point(12, 75);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(31, 16);
			this.label2.TabIndex = 184;
			this.label2.Text = "Saldo";
			this.label2.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtSaldo
			// 
			appearance2.BackColorDisabled = System.Drawing.Color.FromArgb(((System.Byte)(224)), ((System.Byte)(224)), ((System.Byte)(224)));
			appearance2.BackColorDisabled2 = System.Drawing.Color.FromArgb(((System.Byte)(224)), ((System.Byte)(224)), ((System.Byte)(224)));
			appearance2.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtSaldo.Appearance = appearance2;
			this.txtSaldo.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtSaldo.Enabled = false;
			this.txtSaldo.FormatString = "#,##0.00";
			this.txtSaldo.Location = new System.Drawing.Point(76, 72);
			this.txtSaldo.Name = "txtSaldo";
			this.txtSaldo.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtSaldo.PromptChar = ' ';
			this.txtSaldo.Size = new System.Drawing.Size(136, 22);
			this.txtSaldo.TabIndex = 183;
			this.txtSaldo.Click += new System.EventHandler(this.txtSaldo_Click);
			this.txtSaldo.Enter += new System.EventHandler(this.txtSaldo_Enter);
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
			// btnAceptar
			// 
			this.btnAceptar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnAceptar.Enabled = false;
			this.btnAceptar.Location = new System.Drawing.Point(8, 120);
			this.btnAceptar.Name = "btnAceptar";
			this.btnAceptar.TabIndex = 185;
			this.btnAceptar.Text = "&Aceptar";
			this.btnAceptar.Click += new System.EventHandler(this.btnAceptar_Click);
			// 
			// btnCancelar
			// 
			this.btnCancelar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnCancelar.CausesValidation = false;
			this.btnCancelar.Location = new System.Drawing.Point(88, 120);
			this.btnCancelar.Name = "btnCancelar";
			this.btnCancelar.TabIndex = 186;
			this.btnCancelar.Text = "&Cancelar";
			this.btnCancelar.Click += new System.EventHandler(this.btnCancelar_Click);
			// 
			// groupBox1
			// 
			this.groupBox1.Location = new System.Drawing.Point(0, 104);
			this.groupBox1.Name = "groupBox1";
			this.groupBox1.Size = new System.Drawing.Size(224, 8);
			this.groupBox1.TabIndex = 187;
			this.groupBox1.TabStop = false;
			// 
			// txtValor
			// 
			this.txtValor.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtValor.FormatString = "#,##0.00";
			this.txtValor.Location = new System.Drawing.Point(76, 40);
			this.txtValor.Name = "txtValor";
			this.txtValor.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtValor.PromptChar = ' ';
			this.txtValor.Size = new System.Drawing.Size(136, 22);
			this.txtValor.TabIndex = 188;
			this.txtValor.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtValor_KeyDown);
			this.txtValor.Click += new System.EventHandler(this.txtValor_Click);
			this.txtValor.Validated += new System.EventHandler(this.txtValor_Validated);
			this.txtValor.Enter += new System.EventHandler(this.txtValor_Enter);
			// 
			// frmCruzarNotaDeCredito
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 14);
			this.CausesValidation = false;
			this.ClientSize = new System.Drawing.Size(224, 144);
			this.ControlBox = false;
			this.Controls.Add(this.txtValor);
			this.Controls.Add(this.groupBox1);
			this.Controls.Add(this.btnAceptar);
			this.Controls.Add(this.btnCancelar);
			this.Controls.Add(this.label2);
			this.Controls.Add(this.label1);
			this.Controls.Add(this.label3);
			this.Controls.Add(this.txtSaldo);
			this.Controls.Add(this.txtNumero);
			this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
			this.MaximizeBox = false;
			this.MinimizeBox = false;
			this.Name = "frmCruzarNotaDeCredito";
			this.ShowInTaskbar = false;
			this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
			this.Text = "Notas de Crédito";
			this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.frmCruzarNotaDeCredito_KeyDown);
			this.Load += new System.EventHandler(this.frmCruzarNotaDeCredito_Load);
			((System.ComponentModel.ISupportInitialize)(this.txtNumero)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtSaldo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtValor)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		
		private void frmCruzarNotaDeCredito_Load(object sender, System.EventArgs e)
		{			
			SqlDataReader dr = Funcion.rEscalarSQL(cdsSeteoF, string.Format("Select idCompra, Numero, ISNULL(SaldoAnticipo, 0) From Compra Where idCompra = {0} And idTipoFactura = 5", IdCompra), true);
			dr.Read();
			if (dr.HasRows)
			{
				IdCompra = dr.GetInt32(0);
				this.txtNumero.Text = dr.GetString(1);
				this.txtValor.Value = dr.GetDecimal(2);
				this.txtSaldo.Value = dr.GetDecimal(2);
		
				if (Convert.ToDecimal(this.txtSaldo.Value) > 0.00m) 
					this.btnAceptar.Enabled = true;
			}			
			dr.Close();			
		}

		private void txtNumero_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{			
			Funcion.CamposNumericos(sender, e);

			if (e.KeyChar == 13) this.txtValor.Focus();
		}

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);		
		}

		private void txtNumero_Validating(object sender, System.ComponentModel.CancelEventArgs e)
		{
			if (this.txtNumero.TextLength == 0)
			{
				MessageBox.Show("Escriba el numero del documento", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.txtNumero.Focus();	
				e.Cancel = true;
				return;
			}
			
			char Pad = '0';
			this.txtNumero.Text = this.txtNumero.Text.ToString().PadLeft(9, Pad);

			string sSQL = string.Format("Select COUNT(*) From Compra Where Numero = '{0}' And idCliente = {1} And idTipoFactura = 5", this.txtNumero.Text.ToString(), IdCliente);
			int iCount = Funcion.iEscalarSQL(cdsSeteoF, sSQL);

			if (iCount == 0)
			{
				MessageBox.Show("El Documento no existe", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.txtNumero.Focus();	
				e.Cancel = true;
				return;
			}
						
			sSQL = string.Format("Select ISNULL(SaldoAnticipo, 0) From Compra Where Numero = '{0}' And idCliente = {1} And idTipoFactura = 5", this.txtNumero.Text.ToString(), IdCliente);
			decimal dSaldo = Funcion.decEscalarSQL(cdsSeteoF, sSQL);		
           
			if (dSaldo <= 0.00m)
			{
				MessageBox.Show("El documento no tiene saldo", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.txtNumero.Focus();				
				e.Cancel = true;
				return;
			}
		}

		private void txtNumero_Validated(object sender, System.EventArgs e)
		{
			this.btnAceptar.Enabled = false;
			IdCompra = 0;

			SqlDataReader dr = Funcion.rEscalarSQL(cdsSeteoF, string.Format("Select idCompra, Numero, ISNULL(SaldoAnticipo, 0) From Compra Where Numero = '{0}' And idCliente = {1} And idTipoFactura = 5", 
				this.txtNumero.Text.ToString(), IdCliente), true);
			dr.Read();
			if (dr.HasRows)
			{
				IdCompra = dr.GetInt32(0);
				this.txtNumero.Text = dr.GetString(1);
				this.txtValor.Value = dr.GetDecimal(2);
				this.txtSaldo.Value = dr.GetDecimal(2);
				
				if (Convert.ToDecimal(this.txtSaldo.Value) > 0.00m) 
					this.btnAceptar.Enabled = true;
			}	
			dr.Close();
		}

		private void btnAceptar_Click(object sender, System.EventArgs e)
		{	
			if (this.txtNumero.Text.ToString().Trim().Length == 0)
			{
				MessageBox.Show("Ingrese el numero de documento", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.txtNumero.Focus();								
				return;
			}

			if (IdCompra == 0)
			{
				MessageBox.Show("El documento no existe", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.txtNumero.Focus();								
				return;
			}

			if (Convert.ToDecimal(this.txtSaldo.Value) <= 0.00m)
			{
				MessageBox.Show("El documento no tiene saldo", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.txtNumero.Focus();								
				return;
			}

			if (Convert.ToDecimal(this.txtValor.Value) == 0.00m)
			{
				MessageBox.Show("El valor a cruzar no pueder ser CERO", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.txtNumero.Focus();								
				return;
			}			

			if (Convert.ToDecimal(this.txtValor.Value) > Convert.ToDecimal(this.txtSaldo.Value))
			{
				MessageBox.Show("El valor a cruzar no pueder ser mayor al saldo", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.txtNumero.Focus();								
				return;
			}

			this.DialogResult = DialogResult.OK;			
		}

		private void btnCancelar_Click(object sender, System.EventArgs e)
		{
			this.DialogResult = DialogResult.No;
		}

		private void frmCruzarNotaDeCredito_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Escape) this.DialogResult = DialogResult.No;
		}

		private void txtValor_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.btnAceptar.Focus();
		}

		private void txtSaldo_Enter(object sender, System.EventArgs e)
		{
			this.txtSaldo.SelectAll();
		}

		private void txtSaldo_Click(object sender, System.EventArgs e)
		{
			this.txtSaldo.SelectAll();
		}

		private void txtValor_Enter(object sender, System.EventArgs e)
		{
			this.txtValor.SelectAll();
		}

		private void txtValor_Click(object sender, System.EventArgs e)
		{
			this.txtValor.SelectAll();
		}

		private void txtValor_Validated(object sender, System.EventArgs e)
		{
			if (this.txtValor.Value == System.DBNull.Value) 
				this.txtValor.Value = Convert.ToDecimal(this.txtSaldo.Value);
		}
	}
}

