using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using DevExpress.XtraEditors;

namespace Latinium
{
	/// <summary>
	/// Summary description for frmNumeroNotaDeDebito.
	/// </summary>
	public class frmNumeroNotaDeDebito : DevExpress.XtraEditors.XtraForm
	{
		private System.Windows.Forms.Label label8;
		private System.Windows.Forms.Label label4;
		public Infragistics.Win.UltraWinSchedule.UltraCalendarCombo txtFecha;
		private System.Windows.Forms.Button btnAceptar;
		private System.Windows.Forms.Button btnCancelar;
		private C1.Data.C1DataSet cdsSeteoF;
		private System.Windows.Forms.Label label1;
		public Infragistics.Win.UltraWinEditors.UltraTextEditor txtDebitoCapital;
		public Infragistics.Win.UltraWinEditors.UltraTextEditor txtDebitoInteres;
		private System.Windows.Forms.GroupBox groupBox1;
		/// <summary>
		/// Required designer variable.
		/// </summary>
		private System.ComponentModel.Container components = null;

		DateTime dtVence = DateTime.Today;

		public frmNumeroNotaDeDebito(DateTime DTVence)
		{
			//
			// Required for Windows Form Designer support
			//
			InitializeComponent();

			dtVence = DTVence;

			//
			// TODO: Add any constructor code after InitializeComponent call
			//
		}

		public frmNumeroNotaDeDebito()
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
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton1 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmNumeroNotaDeDebito));
			this.label8 = new System.Windows.Forms.Label();
			this.txtDebitoInteres = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label4 = new System.Windows.Forms.Label();
			this.txtFecha = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.btnAceptar = new System.Windows.Forms.Button();
			this.btnCancelar = new System.Windows.Forms.Button();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.label1 = new System.Windows.Forms.Label();
			this.txtDebitoCapital = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.groupBox1 = new System.Windows.Forms.GroupBox();
			((System.ComponentModel.ISupportInitialize)(this.txtDebitoInteres)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtFecha)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDebitoCapital)).BeginInit();
			this.SuspendLayout();
			// 
			// label8
			// 
			this.label8.AutoSize = true;
			this.label8.Location = new System.Drawing.Point(8, 43);
			this.label8.Name = "label8";
			this.label8.Size = new System.Drawing.Size(120, 16);
			this.label8.TabIndex = 281;
			this.label8.Text = "Nota De Debito Interes";
			this.label8.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtDebitoInteres
			// 
			this.txtDebitoInteres.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtDebitoInteres.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtDebitoInteres.Location = new System.Drawing.Point(160, 40);
			this.txtDebitoInteres.MaxLength = 15;
			this.txtDebitoInteres.Name = "txtDebitoInteres";
			this.txtDebitoInteres.Size = new System.Drawing.Size(136, 22);
			this.txtDebitoInteres.TabIndex = 1;
			this.txtDebitoInteres.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtDebitoInteres_KeyDown);
			this.txtDebitoInteres.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtDebitoInteres_KeyPress);
			// 
			// label4
			// 
			this.label4.AutoSize = true;
			this.label4.Location = new System.Drawing.Point(8, 8);
			this.label4.Name = "label4";
			this.label4.Size = new System.Drawing.Size(34, 16);
			this.label4.TabIndex = 283;
			this.label4.Text = "Fecha";
			this.label4.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtFecha
			// 
			this.txtFecha.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton1.Caption = "Today";
			this.txtFecha.DateButtons.Add(dateButton1);
			this.txtFecha.Format = "dd/MM/yyyy";
			this.txtFecha.Location = new System.Drawing.Point(80, 8);
			this.txtFecha.Name = "txtFecha";
			this.txtFecha.NonAutoSizeHeight = 23;
			this.txtFecha.Size = new System.Drawing.Size(112, 21);
			this.txtFecha.SpinButtonsVisible = true;
			this.txtFecha.TabIndex = 0;
			this.txtFecha.Value = ((object)(resources.GetObject("txtFecha.Value")));
			this.txtFecha.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtFecha_KeyDown);
			// 
			// btnAceptar
			// 
			this.btnAceptar.Location = new System.Drawing.Point(8, 128);
			this.btnAceptar.Name = "btnAceptar";
			this.btnAceptar.TabIndex = 3;
			this.btnAceptar.Text = "Aceptar";
			this.btnAceptar.Click += new System.EventHandler(this.btnAceptar_Click);
			// 
			// btnCancelar
			// 
			this.btnCancelar.Location = new System.Drawing.Point(88, 128);
			this.btnCancelar.Name = "btnCancelar";
			this.btnCancelar.TabIndex = 4;
			this.btnCancelar.Text = "Cancelar";
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
			// label1
			// 
			this.label1.AutoSize = true;
			this.label1.Location = new System.Drawing.Point(8, 75);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(118, 16);
			this.label1.TabIndex = 287;
			this.label1.Text = "Nota De Debito Capital";
			this.label1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtDebitoCapital
			// 
			this.txtDebitoCapital.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtDebitoCapital.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtDebitoCapital.Location = new System.Drawing.Point(160, 72);
			this.txtDebitoCapital.MaxLength = 15;
			this.txtDebitoCapital.Name = "txtDebitoCapital";
			this.txtDebitoCapital.Size = new System.Drawing.Size(136, 22);
			this.txtDebitoCapital.TabIndex = 2;
			this.txtDebitoCapital.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtDebitoCapital_KeyDown);
			this.txtDebitoCapital.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtDebitoCapital_KeyPress);
			// 
			// groupBox1
			// 
			this.groupBox1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.groupBox1.Location = new System.Drawing.Point(0, 112);
			this.groupBox1.Name = "groupBox1";
			this.groupBox1.Size = new System.Drawing.Size(312, 8);
			this.groupBox1.TabIndex = 288;
			this.groupBox1.TabStop = false;
			// 
			// frmNumeroNotaDeDebito
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 14);
			this.ClientSize = new System.Drawing.Size(306, 160);
			this.ControlBox = false;
			this.Controls.Add(this.groupBox1);
			this.Controls.Add(this.label1);
			this.Controls.Add(this.txtDebitoCapital);
			this.Controls.Add(this.btnCancelar);
			this.Controls.Add(this.btnAceptar);
			this.Controls.Add(this.label4);
			this.Controls.Add(this.txtFecha);
			this.Controls.Add(this.label8);
			this.Controls.Add(this.txtDebitoInteres);
			this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
			this.KeyPreview = true;
			this.MaximizeBox = false;
			this.MinimizeBox = false;
			this.Name = "frmNumeroNotaDeDebito";
			this.ShowInTaskbar = false;
			this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
			this.Text = "Registro de Pago";
			this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.frmNumeroNotaDeDebito_KeyDown);
			this.Load += new System.EventHandler(this.frmNumeroNotaDeDebito_Load);
			((System.ComponentModel.ISupportInitialize)(this.txtDebitoInteres)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtFecha)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDebitoCapital)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void frmNumeroNotaDeDebito_Load(object sender, System.EventArgs e)
		{
			this.txtFecha.CalendarInfo.MinDate = dtVence.AddDays(-10);
			this.txtFecha.CalendarInfo.MaxDate = DateTime.Today;
			this.txtFecha.Value = DateTime.Today;			
		}

		private void btnCancelar_Click(object sender, System.EventArgs e)
		{
			this.DialogResult = DialogResult.No;
		}

		private void frmNumeroNotaDeDebito_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Escape) this.DialogResult = DialogResult.No;
		}

		private void btnAceptar_Click(object sender, System.EventArgs e)
		{
			if (!Validacion.vbFechaEnDocumentos(this.txtFecha, "Ingrese la Fecha", true, "Bancos", cdsSeteoF)) return;
			
			if (this.txtDebitoInteres.Text.ToString().Length == 0)
			{
				MessageBox.Show("Escriba el numero de la nota de debito para interes", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.txtDebitoInteres.Focus();				
				return;
			}

			if (this.txtDebitoCapital.Text.ToString().Length == 0)
			{
				MessageBox.Show("Escriba el numero de la nota de debito para capital", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.txtDebitoCapital.Focus();				
				return;
			}
//			if (Funcion.iEscalarSQL(cdsSeteoF, string.Format("Select COUNT(*) From Asiento Where CodAsiento = '{0}' And idTipoAsiento = 4 And LEN(CodAsiento) > 0", this.txtDebitoInteres.Text.ToString())) > 0)
//			{
//				MessageBox.Show(string.Format("Numero de Documento '{0}' ya Existe", this.txtNumero.Text.ToString()), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
//				this.txtNumero.Focus();				
//				return;
//			}
			
			this.DialogResult = DialogResult.OK;
		}

		private void txtFecha_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.txtDebitoInteres.Focus();
		}

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);
		}

		private void txtDebitoInteres_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			Funcion.CamposNumericos(sender, e);
		}

		private void txtDebitoCapital_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			Funcion.CamposNumericos(sender, e);
		}

		private void txtDebitoInteres_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.txtDebitoCapital.Focus();
		}

		private void txtDebitoCapital_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.btnAceptar.Focus();
		}
	}
}

