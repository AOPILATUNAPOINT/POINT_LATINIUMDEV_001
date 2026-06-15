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
	/// Descripción breve de frmCodigoGifCard.
	/// </summary>
	public class frmCodigoGifCard : System.Windows.Forms.Form
	{
		private System.Windows.Forms.GroupBox groupBox1;
		private System.Windows.Forms.Button btnAceptar;
		private System.Windows.Forms.Button btnCancelar;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;
		public Infragistics.Win.UltraWinEditors.UltraNumericEditor txtValor;
		private System.Windows.Forms.Label label2;
		private System.Windows.Forms.Label label1;
		private System.Windows.Forms.Label label3;
		public Infragistics.Win.UltraWinEditors.UltraNumericEditor txtSaldo;
		public Infragistics.Win.UltraWinEditors.UltraTextEditor txtNumero;
		private C1.Data.C1DataSet cdsSeteoF;
    
		int idCliente = 0 ;
		public frmCodigoGifCard()
		{
			//
			// Necesario para admitir el Diseñador de Windows Forms
			//
			InitializeComponent();

			//
			// TODO: agregar código de constructor después de llamar a InitializeComponent
			//
		}

		public int idGifCards = 0;
		decimal dValorFactura = 0.00m;
		public frmCodigoGifCard(int IdCliente, decimal DValorFactura)
		{
			//
			// Necesario para admitir el Diseñador de Windows Forms
			//
			InitializeComponent();
			idCliente = IdCliente;
			dValorFactura = DValorFactura;
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
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			this.groupBox1 = new System.Windows.Forms.GroupBox();
			this.btnAceptar = new System.Windows.Forms.Button();
			this.btnCancelar = new System.Windows.Forms.Button();
			this.txtValor = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label2 = new System.Windows.Forms.Label();
			this.label1 = new System.Windows.Forms.Label();
			this.label3 = new System.Windows.Forms.Label();
			this.txtSaldo = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtNumero = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			((System.ComponentModel.ISupportInitialize)(this.txtValor)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtSaldo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNumero)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			this.SuspendLayout();
			// 
			// groupBox1
			// 
			this.groupBox1.Location = new System.Drawing.Point(0, 104);
			this.groupBox1.Name = "groupBox1";
			this.groupBox1.Size = new System.Drawing.Size(224, 8);
			this.groupBox1.TabIndex = 988;
			this.groupBox1.TabStop = false;
			// 
			// btnAceptar
			// 
			this.btnAceptar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnAceptar.Enabled = false;
			this.btnAceptar.Location = new System.Drawing.Point(16, 120);
			this.btnAceptar.Name = "btnAceptar";
			this.btnAceptar.TabIndex = 986;
			this.btnAceptar.Text = "&Aceptar";
			this.btnAceptar.Click += new System.EventHandler(this.btnAceptar_Click);
			// 
			// btnCancelar
			// 
			this.btnCancelar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnCancelar.CausesValidation = false;
			this.btnCancelar.Location = new System.Drawing.Point(96, 120);
			this.btnCancelar.Name = "btnCancelar";
			this.btnCancelar.TabIndex = 987;
			this.btnCancelar.Text = "&Cancelar";
			this.btnCancelar.Click += new System.EventHandler(this.btnCancelar_Click);
			// 
			// txtValor
			// 
			this.txtValor.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtValor.FormatString = "#,##0.00";
			this.txtValor.Location = new System.Drawing.Point(72, 40);
			this.txtValor.Name = "txtValor";
			this.txtValor.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtValor.PromptChar = ' ';
			this.txtValor.Size = new System.Drawing.Size(136, 21);
			this.txtValor.TabIndex = 994;
			this.txtValor.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtValor_KeyDown);
			this.txtValor.Click += new System.EventHandler(this.txtValor_Click);
			this.txtValor.Validated += new System.EventHandler(this.txtValor_Validated);
			this.txtValor.Enter += new System.EventHandler(this.txtValor_Enter);
			// 
			// label2
			// 
			this.label2.AutoSize = true;
			this.label2.Location = new System.Drawing.Point(8, 80);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(33, 16);
			this.label2.TabIndex = 993;
			this.label2.Text = "Saldo";
			this.label2.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label1
			// 
			this.label1.AutoSize = true;
			this.label1.Location = new System.Drawing.Point(8, 48);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(31, 16);
			this.label1.TabIndex = 991;
			this.label1.Text = "Valor";
			this.label1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label3
			// 
			this.label3.AutoSize = true;
			this.label3.Location = new System.Drawing.Point(8, 16);
			this.label3.Name = "label3";
			this.label3.Size = new System.Drawing.Size(40, 16);
			this.label3.TabIndex = 990;
			this.label3.Text = "Codigo";
			this.label3.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtSaldo
			// 
			appearance1.BackColorDisabled = System.Drawing.Color.FromArgb(((System.Byte)(224)), ((System.Byte)(224)), ((System.Byte)(224)));
			appearance1.BackColorDisabled2 = System.Drawing.Color.FromArgb(((System.Byte)(224)), ((System.Byte)(224)), ((System.Byte)(224)));
			appearance1.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtSaldo.Appearance = appearance1;
			this.txtSaldo.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtSaldo.Enabled = false;
			this.txtSaldo.FormatString = "#,##0.00";
			this.txtSaldo.Location = new System.Drawing.Point(72, 72);
			this.txtSaldo.Name = "txtSaldo";
			this.txtSaldo.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtSaldo.PromptChar = ' ';
			this.txtSaldo.Size = new System.Drawing.Size(136, 21);
			this.txtSaldo.TabIndex = 992;
			// 
			// txtNumero
			// 
			appearance2.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtNumero.Appearance = appearance2;
			this.txtNumero.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtNumero.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNumero.Location = new System.Drawing.Point(72, 8);
			this.txtNumero.MaxLength = 9;
			this.txtNumero.Name = "txtNumero";
			this.txtNumero.Size = new System.Drawing.Size(136, 21);
			this.txtNumero.TabIndex = 989;
			this.txtNumero.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtNumero_KeyDown);
			this.txtNumero.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtNumero_KeyPress);
			this.txtNumero.Validating += new System.ComponentModel.CancelEventHandler(this.txtNumero_Validating);
			this.txtNumero.Validated += new System.EventHandler(this.txtNumero_Validated);
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
			// frmCodigoGifCard
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(224, 150);
			this.ControlBox = false;
			this.Controls.Add(this.txtValor);
			this.Controls.Add(this.label2);
			this.Controls.Add(this.label1);
			this.Controls.Add(this.label3);
			this.Controls.Add(this.txtSaldo);
			this.Controls.Add(this.txtNumero);
			this.Controls.Add(this.groupBox1);
			this.Controls.Add(this.btnAceptar);
			this.Controls.Add(this.btnCancelar);
			this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
			this.KeyPreview = true;
			this.Name = "frmCodigoGifCard";
			this.ShowInTaskbar = false;
			this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
			this.Text = "Codigo Gif Card";
			this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.frmCodigoGifCard_KeyDown);
			this.Load += new System.EventHandler(this.frmCodigoGifCard_Load);
			((System.ComponentModel.ISupportInitialize)(this.txtValor)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtSaldo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNumero)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void frmCodigoGifCard_Load(object sender, System.EventArgs e)
		{
			
		}

		private void txtNumero_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
	
		}

		private void txtNumero_Validating(object sender, System.ComponentModel.CancelEventArgs e)
		{
			 #region Valida Documento
			if (this.txtNumero.TextLength == 0)
			{
				MessageBox.Show("Ingrese el Codigo ", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.txtNumero.Focus();	
				e.Cancel = true;
				return;
			}
			
			string sSQL = string.Format("Select COUNT(*) From GifCards Where Codigo = '{0}' And idPersonal = {1} ", this.txtNumero.Text.ToString(), idCliente);
			int iCount = Funcion.iEscalarSQL(cdsSeteoF, sSQL);

			if (iCount == 0)
			{
				MessageBox.Show("El Documento no existe", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.txtNumero.Focus();	
				e.Cancel = true;
				return;
			}
						
			sSQL = string.Format("Select ISNULL(Saldo, 0) From GifCards Where Codigo = '{0}' And idPersonal = {1} ", this.txtNumero.Text.ToString(), idCliente);
			decimal dSaldo = Funcion.decEscalarSQL(cdsSeteoF, sSQL);		
           
			if (dSaldo <= 0.00m)
			{
				MessageBox.Show("El documento no tiene saldo", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.txtNumero.Focus();				
				e.Cancel = true;
				return;
			}
			#endregion Valida Documento
		}

		private void txtNumero_Validated(object sender, System.EventArgs e)
		{
			this.btnAceptar.Enabled = false;
			 idGifCards = 0;

			SqlDataReader dr = Funcion.rEscalarSQL(cdsSeteoF, string.Format("Select idGifCards, Codigo, ISNULL(Saldo, 0) From GifCards Where Codigo = '{0}' And idPersonal = {1} ", 
				this.txtNumero.Text.ToString(), idCliente), true);
			dr.Read();
			if (dr.HasRows)
			{
				idGifCards = dr.GetInt32(0);
				this.txtNumero.Text = dr.GetString(1);
				this.txtValor.Value = dr.GetDecimal(2);
				this.txtSaldo.Value = dr.GetDecimal(2);
				
				if (Convert.ToDecimal(this.txtSaldo.Value) > 0.00m) 
					this.btnAceptar.Enabled = true;
			}	
			dr.Close();
		}

		private void txtValor_Validated(object sender, System.EventArgs e)
		{
			if (this.txtValor.Value == System.DBNull.Value) 
				this.txtValor.Value = Convert.ToDecimal(this.txtSaldo.Value);
		}

		private void txtValor_Click(object sender, System.EventArgs e)
		{
			this.txtValor.SelectAll();
		}

		private void txtValor_Enter(object sender, System.EventArgs e)
		{
			this.txtValor.SelectAll();
		}

		private void btnCancelar_Click(object sender, System.EventArgs e)
		{
			this.DialogResult = DialogResult.No;
		}

		private void txtValor_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.btnAceptar.Focus();
		}

		private void frmCodigoGifCard_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Escape) this.DialogResult = DialogResult.No;
		}

		private void btnAceptar_Click(object sender, System.EventArgs e)
		{
			if (this.txtNumero.Text.ToString().Trim().Length == 0)
			{
				MessageBox.Show("Ingrese el numero de documento", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.txtNumero.Focus();								
				return;
			}

			if (idGifCards == 0)
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

			if (Convert.ToDecimal(this.txtValor.Value) > dValorFactura)
			{
				MessageBox.Show("El valor a cruzar no pueder ser mayor al valor de la factura", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.txtNumero.Focus();								
				return;
			}
			
			this.DialogResult = DialogResult.OK;	
		}

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);	
			
		}

		private void txtNumero_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (e.KeyChar == 13) this.txtValor.Focus();
		}
	}
		
}

