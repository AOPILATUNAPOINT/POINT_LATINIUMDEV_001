using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using DevExpress.XtraEditors;

namespace Latinium
{
	/// <summary>
	/// Summary description for frmNumeracionCheque.
	/// </summary>
	public class frmNumeracionCheque : DevExpress.XtraEditors.XtraForm
	{
		/// <summary>
		/// Required designer variable.
		/// </summary>
		private System.ComponentModel.Container components = null;
		private DevExpress.XtraEditors.SimpleButton btImprimir;
		private DevExpress.XtraEditors.SimpleButton btCancelar;

		string SNumCheque = "";
		string SBeneficiario = "";
		int IdCuenta = 0;
		int IdFormaPago = 0;
		public Infragistics.Win.UltraWinEditors.UltraTextEditor txtNumCheque;
		private System.Windows.Forms.Label label1;
		public Infragistics.Win.UltraWinEditors.UltraTextEditor txtNombre;
		private System.Windows.Forms.Label label6;
		public Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtFecha;
		private C1.Data.C1DataSet cdsSeteoF;
		private Infragistics.Win.Misc.UltraLabel ultraLabel1;
		public Infragistics.Win.UltraWinEditors.UltraTextEditor txtNumEgreso;
		private System.Windows.Forms.Label label2;
		public Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtFechaCheque;
		private Infragistics.Win.Misc.UltraLabel lblCheque;
		DateTime dtFechaPago = DateTime.Today;

		public frmNumeracionCheque(string sNumCheque, string sBeneficiario, int idCuenta, int idFormaPago, DateTime DTFechaPago)
		{
			//
			// Required for Windows Form Designer support
			//
			InitializeComponent();
			SNumCheque = sNumCheque;
			SBeneficiario = sBeneficiario;
			IdCuenta = idCuenta;
			IdFormaPago = idFormaPago;
			dtFechaPago = DTFechaPago;
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
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton1 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton2 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			this.btImprimir = new DevExpress.XtraEditors.SimpleButton();
			this.btCancelar = new DevExpress.XtraEditors.SimpleButton();
			this.txtNumCheque = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.lblCheque = new Infragistics.Win.Misc.UltraLabel();
			this.label1 = new System.Windows.Forms.Label();
			this.txtNombre = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label6 = new System.Windows.Forms.Label();
			this.dtFecha = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.txtNumEgreso = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.ultraLabel1 = new Infragistics.Win.Misc.UltraLabel();
			this.label2 = new System.Windows.Forms.Label();
			this.dtFechaCheque = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			((System.ComponentModel.ISupportInitialize)(this.txtNumCheque)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNombre)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFecha)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNumEgreso)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFechaCheque)).BeginInit();
			this.SuspendLayout();
			// 
			// btImprimir
			// 
			this.btImprimir.Location = new System.Drawing.Point(8, 176);
			this.btImprimir.Name = "btImprimir";
			this.btImprimir.Size = new System.Drawing.Size(72, 26);
			this.btImprimir.TabIndex = 2;
			this.btImprimir.Text = "&Imprimir";
			this.btImprimir.ToolTip = "Asigna la cuenta a los pagos seleccionados";
			this.btImprimir.Click += new System.EventHandler(this.btImprimir_Click);
			// 
			// btCancelar
			// 
			this.btCancelar.Location = new System.Drawing.Point(88, 176);
			this.btCancelar.Name = "btCancelar";
			this.btCancelar.Size = new System.Drawing.Size(72, 26);
			this.btCancelar.TabIndex = 1;
			this.btCancelar.Text = "&Cancelar";
			this.btCancelar.Click += new System.EventHandler(this.btCancelar_Click);
			// 
			// txtNumCheque
			// 
			appearance1.ForeColor = System.Drawing.Color.Red;
			this.txtNumCheque.Appearance = appearance1;
			this.txtNumCheque.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtNumCheque.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNumCheque.Font = new System.Drawing.Font("Tahoma", 18F);
			this.txtNumCheque.Location = new System.Drawing.Point(184, 72);
			this.txtNumCheque.Name = "txtNumCheque";
			this.txtNumCheque.Size = new System.Drawing.Size(280, 39);
			this.txtNumCheque.TabIndex = 0;
			this.txtNumCheque.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtNumCheque_KeyPress);
			// 
			// lblCheque
			// 
			this.lblCheque.AutoSize = true;
			this.lblCheque.BackColor = System.Drawing.Color.Transparent;
			this.lblCheque.Font = new System.Drawing.Font("Tahoma", 12F, System.Drawing.FontStyle.Bold);
			this.lblCheque.Location = new System.Drawing.Point(8, 80);
			this.lblCheque.Name = "lblCheque";
			this.lblCheque.Size = new System.Drawing.Size(91, 22);
			this.lblCheque.TabIndex = 210;
			this.lblCheque.Text = "Cheque N.";
			// 
			// label1
			// 
			this.label1.AutoSize = true;
			this.label1.BackColor = System.Drawing.Color.Transparent;
			this.label1.Location = new System.Drawing.Point(8, 11);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(48, 16);
			this.label1.TabIndex = 211;
			this.label1.Text = "Nombre:";
			this.label1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtNombre
			// 
			appearance2.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtNombre.Appearance = appearance2;
			this.txtNombre.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtNombre.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNombre.Location = new System.Drawing.Point(72, 8);
			this.txtNombre.Name = "txtNombre";
			this.txtNombre.Size = new System.Drawing.Size(392, 22);
			this.txtNombre.TabIndex = 212;
			this.txtNombre.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtNombre_KeyPress);
			// 
			// label6
			// 
			this.label6.AutoSize = true;
			this.label6.Location = new System.Drawing.Point(8, 43);
			this.label6.Name = "label6";
			this.label6.Size = new System.Drawing.Size(91, 16);
			this.label6.TabIndex = 214;
			this.label6.Text = "Fecha Del Egreso";
			// 
			// dtFecha
			// 
			appearance3.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtFecha.Appearance = appearance3;
			this.dtFecha.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton1.Caption = "Today";
			this.dtFecha.DateButtons.Add(dateButton1);
			this.dtFecha.Format = "dd/MM/yyyy";
			this.dtFecha.Location = new System.Drawing.Point(120, 40);
			this.dtFecha.Name = "dtFecha";
			this.dtFecha.NonAutoSizeHeight = 23;
			this.dtFecha.Size = new System.Drawing.Size(104, 21);
			this.dtFecha.SpinButtonsVisible = true;
			this.dtFecha.TabIndex = 213;
			this.dtFecha.Value = new System.DateTime(2005, 6, 8, 0, 0, 0, 0);
			this.dtFecha.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.dtFecha_KeyPress);
			this.dtFecha.Validated += new System.EventHandler(this.dtFecha_Validated);
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
			// txtNumEgreso
			// 
			appearance4.ForeColor = System.Drawing.Color.Red;
			this.txtNumEgreso.Appearance = appearance4;
			this.txtNumEgreso.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtNumEgreso.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNumEgreso.Enabled = false;
			this.txtNumEgreso.Font = new System.Drawing.Font("Tahoma", 18F);
			this.txtNumEgreso.Location = new System.Drawing.Point(184, 120);
			this.txtNumEgreso.Name = "txtNumEgreso";
			this.txtNumEgreso.Size = new System.Drawing.Size(280, 39);
			this.txtNumEgreso.TabIndex = 215;
			this.txtNumEgreso.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtNumEgreso_KeyPress);
			// 
			// ultraLabel1
			// 
			this.ultraLabel1.AutoSize = true;
			this.ultraLabel1.BackColor = System.Drawing.Color.Transparent;
			this.ultraLabel1.Font = new System.Drawing.Font("Tahoma", 12F, System.Drawing.FontStyle.Bold);
			this.ultraLabel1.Location = new System.Drawing.Point(8, 128);
			this.ultraLabel1.Name = "ultraLabel1";
			this.ultraLabel1.Size = new System.Drawing.Size(85, 22);
			this.ultraLabel1.TabIndex = 216;
			this.ultraLabel1.Text = "Egreso N.";
			// 
			// label2
			// 
			this.label2.AutoSize = true;
			this.label2.Location = new System.Drawing.Point(240, 43);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(84, 16);
			this.label2.TabIndex = 218;
			this.label2.Text = "Fecha De Cobro";
			// 
			// dtFechaCheque
			// 
			appearance5.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtFechaCheque.Appearance = appearance5;
			this.dtFechaCheque.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton2.Caption = "Today";
			this.dtFechaCheque.DateButtons.Add(dateButton2);
			this.dtFechaCheque.Format = "dd/MM/yyyy";
			this.dtFechaCheque.Location = new System.Drawing.Point(360, 40);
			this.dtFechaCheque.Name = "dtFechaCheque";
			this.dtFechaCheque.NonAutoSizeHeight = 23;
			this.dtFechaCheque.Size = new System.Drawing.Size(104, 21);
			this.dtFechaCheque.SpinButtonsVisible = true;
			this.dtFechaCheque.TabIndex = 217;
			this.dtFechaCheque.Value = new System.DateTime(2005, 6, 8, 0, 0, 0, 0);
			this.dtFechaCheque.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.dtFechaCheque_KeyPress);
			this.dtFechaCheque.Validated += new System.EventHandler(this.dtFechaCheque_Validated);
			// 
			// frmNumeracionCheque
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 14);
			this.ClientSize = new System.Drawing.Size(474, 208);
			this.ControlBox = false;
			this.Controls.Add(this.label2);
			this.Controls.Add(this.dtFechaCheque);
			this.Controls.Add(this.txtNumEgreso);
			this.Controls.Add(this.ultraLabel1);
			this.Controls.Add(this.label6);
			this.Controls.Add(this.label1);
			this.Controls.Add(this.dtFecha);
			this.Controls.Add(this.txtNombre);
			this.Controls.Add(this.txtNumCheque);
			this.Controls.Add(this.lblCheque);
			this.Controls.Add(this.btImprimir);
			this.Controls.Add(this.btCancelar);
			this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
			this.KeyPreview = true;
			this.MaximizeBox = false;
			this.MinimizeBox = false;
			this.Name = "frmNumeracionCheque";
			this.ShowInTaskbar = false;
			this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
			this.Text = "Impresion de Pagos";
			this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.frmNumeracionCheque_KeyDown);
			this.Load += new System.EventHandler(this.frmNumeracionCheque_Load);
			((System.ComponentModel.ISupportInitialize)(this.txtNumCheque)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNombre)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFecha)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNumEgreso)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFechaCheque)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void frmNumeracionCheque_Load(object sender, System.EventArgs e)
		{			
			this.txtNombre.Text = SBeneficiario;
			this.dtFecha.Value = DateTime.Today;
			this.dtFechaCheque.Value = DateTime.Today;			

			if (IdFormaPago == 7) this.lblCheque.Text = "TRANSFERENCIA N. ";
			if (IdFormaPago == 2) this.lblCheque.Text = "CHEQUE N. ";

			if (Funcion.bEscalarSQL(cdsSeteoF, "Select Automatico From AsientoTipo Where idTipoAsiento = 3", true))
				this.txtNumEgreso.Text = Funcion.sEscalarSQL(cdsSeteoF, "Exec AsientoNumeracion 3");			
			else this.txtNumEgreso.Enabled = true;

			this.txtNumCheque.Text = SNumCheque;
			this.txtNumCheque.Focus();
		}

		private void btCancelar_Click(object sender, System.EventArgs e)
		{
			this.DialogResult = DialogResult.No;			
		}

		private void frmNumeracionCheque_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Escape) this.DialogResult = DialogResult.No;
		}

		private void btImprimir_Click(object sender, System.EventArgs e)
		{
			if (this.txtNombre.Text.ToString().Trim().Length == 0)
			{
				MessageBox.Show("Escriba el nombre del beneficiario.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
				this.txtNombre.Focus();
				return;
			}
			if (this.dtFecha.Text.ToString().Trim().Length == 0)
			{
				MessageBox.Show("Escriba la fecha.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);			
				this.dtFecha.Focus();
				return;
			}
			
			if (Funcion.ValidaPeriodo((DateTime)this.dtFecha.Value, cdsSeteoF, "Pagos"))
			{
				DateTime dtFechaVP = (DateTime)this.dtFecha.Value;
				MessageBox.Show(string.Format("No puede ingresar Pagos en esta Fecha, El Período '{0}' esta Cerrado,\n\n Consulte al Administrador", dtFechaVP.ToString("MMMM").ToUpper()), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.dtFecha.Focus();				
				return;
			}
			
			if (this.txtNumCheque.Text.ToString().Trim().Length == 0)
			{
				MessageBox.Show("Escriba el número.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);				
				this.txtNumCheque.Focus();
				return;
			}
			
			string sSQL = string.Format("Select COUNT(*) From AsientoDetalle ad Inner Join Asiento a On a.idAsiento = ad.idAsiento Where a.Cheque = '{0}' And ad.idCuenta In ({1}) And a.idTipoAsiento = 3", this.txtNumCheque.Text.ToString(), IdCuenta);
			int iCuenta = Funcion.iEscalarSQL(cdsSeteoF, sSQL);

			if (IdFormaPago == 2 && iCuenta > 0)
			{
				MessageBox.Show(string.Format("El Cheque '{0}' ya ha sido utilizado.", this.txtNumCheque.Text.ToString()), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
				this.txtNumCheque.Focus();					
				return;
			}
			
			if (!Funcion.bEscalarSQL(cdsSeteoF, "Select Automatico From AsientoTipo Where idTipoAsiento = 3", true))
			{
				if (this.txtNumEgreso.Text.ToString().Trim().Length == 0)
				{
					MessageBox.Show("Escriba el Número de Comprobante de Egreso.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);					
					this.txtNumEgreso.Focus();
					return;
				}
			
				char Pad = '0';
				this.txtNumEgreso.Text = this.txtNumEgreso.Value.ToString().PadLeft(7, Pad);
				
				sSQL = string.Format("Select COUNT(*) From Asiento Where idTipoAsiento = 3 And CodAsiento = '{0}'", this.txtNumEgreso.Text.ToString());
				int iEgreso = Funcion.iEscalarSQL(cdsSeteoF, sSQL);
				
				if (iEgreso > 0)
				{
					MessageBox.Show("El Comprobante de Egreso Número " + this.txtNumEgreso.Text.ToString() + " ya ha sido utilizado.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
					this.txtNumEgreso.Focus();						
					return;
				}
			}
			else Funcion.EjecutaSQL(cdsSeteoF, string.Format("Update AsientoTipo Set Numero = {0} + 1 WHERE idTipoAsiento = 3", int.Parse(this.txtNumEgreso.Text.ToString())));
			
			this.DialogResult = DialogResult.OK;			
		}

		private void txtNumCheque_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			Funcion.CamposNumericos(sender, e);
			if (e.KeyChar == 13) this.txtNumEgreso.Focus();
		}

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);	
		}

		private void txtNumEgreso_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			Funcion.CamposNumericos(sender, e);
			if (e.KeyChar == 13) this.btImprimir.Focus();
		}

		private void txtNombre_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (e.KeyChar == 13) this.dtFecha.Focus();
		}

		private void dtFecha_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (e.KeyChar == 13) this.dtFechaCheque.Focus();
		}
		
		private void dtFechaCheque_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (e.KeyChar == 13) this.txtNumCheque.Focus();
		}

		private void dtFecha_Validated(object sender, System.EventArgs e)
		{
			if (this.dtFecha.Value == System.DBNull.Value) this.dtFecha.Value = DateTime.Today;
		}

		private void dtFechaCheque_Validated(object sender, System.EventArgs e)
		{
			if (this.dtFechaCheque.Value == System.DBNull.Value) this.dtFechaCheque.Value = DateTime.Today;
		}
	}
}

