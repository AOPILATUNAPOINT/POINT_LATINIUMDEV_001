using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de frmAnticipoLotte.
	/// </summary>
	public class frmAnticipoLotte : System.Windows.Forms.Form
	{
		private System.Windows.Forms.Label lblLote;
		private System.Windows.Forms.Button btCancelar;
		private System.Windows.Forms.Button btGrabar;
		private C1.Data.C1DataSet cdsSeteoF;
		public Infragistics.Win.UltraWinEditors.UltraTextEditor txtNumAnticipo;
		private System.Windows.Forms.Label label1;
		public Infragistics.Win.UltraWinEditors.UltraNumericEditor txtSaldo;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public frmAnticipoLotte()
		{
			//
			// Necesario para admitir el Diseñador de Windows Forms
			//
			InitializeComponent();

			//
			// TODO: agregar código de constructor después de llamar a InitializeComponent
			//
		}

		int IdCliente = 0;
		decimal DValor = 0.00m;
		int ILocal = 0;

		public frmAnticipoLotte(int idCliente, decimal dValor, int iLocal)
		{
			//
			// Necesario para admitir el Diseñador de Windows Forms
			//
			IdCliente = idCliente;
			DValor = dValor;
			ILocal = iLocal;

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
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmAnticipoLotte));
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			this.lblLote = new System.Windows.Forms.Label();
			this.txtNumAnticipo = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.btCancelar = new System.Windows.Forms.Button();
			this.btGrabar = new System.Windows.Forms.Button();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.label1 = new System.Windows.Forms.Label();
			this.txtSaldo = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			((System.ComponentModel.ISupportInitialize)(this.txtNumAnticipo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtSaldo)).BeginInit();
			this.SuspendLayout();
			// 
			// lblLote
			// 
			this.lblLote.AutoSize = true;
			this.lblLote.BackColor = System.Drawing.Color.Transparent;
			this.lblLote.Location = new System.Drawing.Point(8, 10);
			this.lblLote.Name = "lblLote";
			this.lblLote.Size = new System.Drawing.Size(60, 16);
			this.lblLote.TabIndex = 174;
			this.lblLote.Text = "N° Anticipo";
			// 
			// txtNumAnticipo
			// 
			appearance1.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtNumAnticipo.Appearance = appearance1;
			this.txtNumAnticipo.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNumAnticipo.Location = new System.Drawing.Point(80, 8);
			this.txtNumAnticipo.MaxLength = 15;
			this.txtNumAnticipo.Name = "txtNumAnticipo";
			this.txtNumAnticipo.Size = new System.Drawing.Size(128, 21);
			this.txtNumAnticipo.TabIndex = 173;
			this.txtNumAnticipo.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtNumAnticipo_KeyPress);
			// 
			// btCancelar
			// 
			this.btCancelar.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.btCancelar.Image = ((System.Drawing.Image)(resources.GetObject("btCancelar.Image")));
			this.btCancelar.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btCancelar.Location = new System.Drawing.Point(112, 72);
			this.btCancelar.Name = "btCancelar";
			this.btCancelar.Size = new System.Drawing.Size(91, 23);
			this.btCancelar.TabIndex = 178;
			this.btCancelar.Text = "&Cancelar";
			this.btCancelar.Click += new System.EventHandler(this.btCancelar_Click);
			// 
			// btGrabar
			// 
			this.btGrabar.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.btGrabar.Image = ((System.Drawing.Image)(resources.GetObject("btGrabar.Image")));
			this.btGrabar.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btGrabar.Location = new System.Drawing.Point(8, 72);
			this.btGrabar.Name = "btGrabar";
			this.btGrabar.Size = new System.Drawing.Size(91, 24);
			this.btGrabar.TabIndex = 177;
			this.btGrabar.Text = "&Grabar";
			this.btGrabar.Click += new System.EventHandler(this.btGrabar_Click);
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
			// label1
			// 
			this.label1.AutoSize = true;
			this.label1.BackColor = System.Drawing.Color.Transparent;
			this.label1.Location = new System.Drawing.Point(8, 43);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(33, 16);
			this.label1.TabIndex = 180;
			this.label1.Text = "Saldo";
			// 
			// txtSaldo
			// 
			appearance2.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtSaldo.Appearance = appearance2;
			this.txtSaldo.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtSaldo.Enabled = false;
			this.txtSaldo.FormatString = "#,##0.00";
			this.txtSaldo.Location = new System.Drawing.Point(80, 41);
			this.txtSaldo.Name = "txtSaldo";
			this.txtSaldo.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtSaldo.PromptChar = ' ';
			this.txtSaldo.Size = new System.Drawing.Size(128, 21);
			this.txtSaldo.TabIndex = 308;
			// 
			// frmAnticipoLotte
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(216, 102);
			this.Controls.Add(this.txtSaldo);
			this.Controls.Add(this.label1);
			this.Controls.Add(this.btCancelar);
			this.Controls.Add(this.btGrabar);
			this.Controls.Add(this.lblLote);
			this.Controls.Add(this.txtNumAnticipo);
			this.MaximizeBox = false;
			this.MinimizeBox = false;
			this.Name = "frmAnticipoLotte";
			this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
			this.Text = "Anticipos";
			this.Load += new System.EventHandler(this.frmAnticipoLotte_Load);
			((System.ComponentModel.ISupportInitialize)(this.txtNumAnticipo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtSaldo)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		public decimal dSaldo = 0.00m;
		public int iValAnticipo = 0;
		public string sNumDocumento = "";
		public int iAnticipo = 0;

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);		
		}

		private void frmAnticipoLotte_Load(object sender, System.EventArgs e)
		{
		
		}

		private void ValAnticipo()
		{
			if (this.txtNumAnticipo.TextLength > 0)
			{
				string sSQLNumero = string.Format("Exec ValAnticipo {0}, '{1}', {2}", IdCliente, this.txtNumAnticipo.Text.ToString(), ILocal);
				iValAnticipo = Funcion.iEscalarSQL(cdsSeteoF, sSQLNumero);

				if (iValAnticipo == 0)
				{
					this.txtSaldo.Value = 0.00m;
					MessageBox.Show(string.Format("El Anticipo no existe"), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					this.txtNumAnticipo.Focus();
					return;
				}
				else
				{
					string sSQLVal = string.Format("Exec ValSaldoAnticipo {0}, '{1}', {2}", IdCliente, this.txtNumAnticipo.Text.ToString(), ILocal);
					this.txtSaldo.Value = Funcion.dEscalarSQL(cdsSeteoF, sSQLVal);
					dSaldo = Convert.ToDecimal(this.txtSaldo.Value);	
					
					if (Convert.ToDecimal(this.txtSaldo.Value) <= 0.00m)
					{
						MessageBox.Show("El documento no tiene saldo", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
						this.txtNumAnticipo.Focus();								
						return;
					}

					if (Convert.ToDecimal(this.txtSaldo.Value) == 0.00m)
					{
						MessageBox.Show("El valor a cruzar no pueder ser CERO", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
						this.txtNumAnticipo.Focus();								
						return;
					}

					this.DialogResult = DialogResult.OK;
				}
			}
			else
			{
				MessageBox.Show(string.Format("Ingrese Numero de Documento"), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.txtNumAnticipo.Focus();
				return;
			}
		}

		private void btGrabar_Click(object sender, System.EventArgs e)
		{
			ValAnticipo();
		}
		private void Verificacion()
		{
			if (this.txtNumAnticipo.TextLength > 0)
			{
				string sSQLNumero = string.Format("Exec ValAnticipo {0}, '{1}', {2}", IdCliente, this.txtNumAnticipo.Text.ToString(), ILocal);
				iValAnticipo = Funcion.iEscalarSQL(cdsSeteoF, sSQLNumero);

				if (iValAnticipo == 0)
				{
					this.txtSaldo.Value = 0.00m;
					MessageBox.Show(string.Format("El Anticipo no existe"), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					this.txtNumAnticipo.Focus();
					return;
				}
				else
				{
					string sSQLVal = string.Format("Exec ValSaldoAnticipo {0}, '{1}', {2}", IdCliente, this.txtNumAnticipo.Text.ToString(), ILocal);
					this.txtSaldo.Value = Funcion.dEscalarSQL(cdsSeteoF, sSQLVal);
					sNumDocumento = this.txtNumAnticipo.Text;
				}
			}
			else
			{
				MessageBox.Show(string.Format("Ingrese Numero de Documento"), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.txtNumAnticipo.Focus();
				return;
			}
		}

		private void txtNumAnticipo_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (e.KeyChar == 13) this.Verificacion();
		}

		private void btCancelar_Click(object sender, System.EventArgs e)
		{
			dSaldo = 0.00m;
			this.Close();
		}
	}
}
