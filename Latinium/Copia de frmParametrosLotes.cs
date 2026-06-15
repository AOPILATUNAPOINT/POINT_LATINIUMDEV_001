using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using DevExpress.XtraEditors;
using System.Data.SqlClient;

namespace Latinium
{
	/// <summary>
	/// Summary description for frmParametrosLotes.
	/// </summary>
	public class frmParametrosLotes : DevExpress.XtraEditors.XtraForm
	{
		private System.Windows.Forms.Label label14;
		private System.Windows.Forms.Label label13;
		public DevExpress.XtraEditors.SimpleButton btnAceptar;
		private System.Windows.Forms.Label label6;
		private C1.Data.C1DataSet cdsSeteoF;
		public Infragistics.Win.UltraWinEditors.UltraNumericEditor txtFondoGarantia;
		public Infragistics.Win.UltraWinEditors.UltraNumericEditor txtComision;
		public Infragistics.Win.UltraWinGrid.UltraCombo cmbCuentas;
		public DevExpress.XtraEditors.SimpleButton btnCancelar;
		/// <summary>
		/// Required designer variable.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public frmParametrosLotes()
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
			Infragistics.Win.Appearance appearance1 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Cuenta", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCuenta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Grupo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CodRapido");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Descripcion");
			this.label14 = new System.Windows.Forms.Label();
			this.txtFondoGarantia = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label13 = new System.Windows.Forms.Label();
			this.txtComision = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.btnCancelar = new DevExpress.XtraEditors.SimpleButton();
			this.btnAceptar = new DevExpress.XtraEditors.SimpleButton();
			this.label6 = new System.Windows.Forms.Label();
			this.cmbCuentas = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			((System.ComponentModel.ISupportInitialize)(this.txtFondoGarantia)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtComision)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCuentas)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			this.SuspendLayout();
			// 
			// label14
			// 
			this.label14.AutoSize = true;
			this.label14.BackColor = System.Drawing.Color.Transparent;
			this.label14.Location = new System.Drawing.Point(6, 43);
			this.label14.Name = "label14";
			this.label14.Size = new System.Drawing.Size(96, 16);
			this.label14.TabIndex = 29;
			this.label14.Text = "Fondo de Garantía";
			// 
			// txtFondoGarantia
			// 
			this.txtFondoGarantia.FormatString = "#,##0.00";
			this.txtFondoGarantia.Location = new System.Drawing.Point(174, 40);
			this.txtFondoGarantia.MaskInput = "nn.nn";
			this.txtFondoGarantia.Name = "txtFondoGarantia";
			this.txtFondoGarantia.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtFondoGarantia.PromptChar = ' ';
			this.txtFondoGarantia.Size = new System.Drawing.Size(112, 22);
			this.txtFondoGarantia.TabIndex = 1;
			this.txtFondoGarantia.Click += new System.EventHandler(this.txtFondoGarantia_Click);
			this.txtFondoGarantia.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtFondoGarantia_KeyPress);
			this.txtFondoGarantia.Validating += new System.ComponentModel.CancelEventHandler(this.txtFondoGarantia_Validating);
			this.txtFondoGarantia.Enter += new System.EventHandler(this.txtFondoGarantia_Enter);
			// 
			// label13
			// 
			this.label13.AutoSize = true;
			this.label13.BackColor = System.Drawing.Color.Transparent;
			this.label13.Location = new System.Drawing.Point(6, 11);
			this.label13.Name = "label13";
			this.label13.Size = new System.Drawing.Size(49, 16);
			this.label13.TabIndex = 27;
			this.label13.Text = "Comisión";
			// 
			// txtComision
			// 
			this.txtComision.FormatString = "#,##0.00";
			this.txtComision.Location = new System.Drawing.Point(174, 8);
			this.txtComision.MaskInput = "nn.nn";
			this.txtComision.Name = "txtComision";
			this.txtComision.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtComision.PromptChar = ' ';
			this.txtComision.Size = new System.Drawing.Size(112, 22);
			this.txtComision.TabIndex = 0;
			this.txtComision.Click += new System.EventHandler(this.txtComision_Click);
			this.txtComision.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtComision_KeyPress);
			this.txtComision.Validating += new System.ComponentModel.CancelEventHandler(this.txtComision_Validating);
			this.txtComision.Enter += new System.EventHandler(this.txtComision_Enter);
			// 
			// btnCancelar
			// 
			this.btnCancelar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btnCancelar.Location = new System.Drawing.Point(150, 136);
			this.btnCancelar.Name = "btnCancelar";
			this.btnCancelar.Size = new System.Drawing.Size(96, 22);
			this.btnCancelar.TabIndex = 4;
			this.btnCancelar.Text = "&Cancelar";
			this.btnCancelar.Click += new System.EventHandler(this.btnCancelar_Click);
			// 
			// btnAceptar
			// 
			this.btnAceptar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btnAceptar.Location = new System.Drawing.Point(46, 136);
			this.btnAceptar.Name = "btnAceptar";
			this.btnAceptar.Size = new System.Drawing.Size(96, 22);
			this.btnAceptar.TabIndex = 3;
			this.btnAceptar.Text = "&Aceptar";
			this.btnAceptar.Click += new System.EventHandler(this.btnAceptar_Click);
			// 
			// label6
			// 
			this.label6.AutoSize = true;
			this.label6.Location = new System.Drawing.Point(8, 80);
			this.label6.Name = "label6";
			this.label6.Size = new System.Drawing.Size(87, 16);
			this.label6.TabIndex = 33;
			this.label6.Text = "Cuenta Contable";
			// 
			// cmbCuentas
			// 
			appearance1.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbCuentas.Appearance = appearance1;
			this.cmbCuentas.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbCuentas.Cursor = System.Windows.Forms.Cursors.Default;
			ultraGridColumn1.Header.VisiblePosition = 1;
			ultraGridColumn2.Header.VisiblePosition = 4;
			ultraGridColumn2.Hidden = true;
			ultraGridColumn3.Header.VisiblePosition = 3;
			ultraGridColumn3.Hidden = true;
			ultraGridColumn4.Header.VisiblePosition = 2;
			ultraGridColumn4.Hidden = true;
			ultraGridColumn5.Header.VisiblePosition = 0;
			ultraGridColumn5.Width = 234;
			ultraGridBand1.Columns.AddRange(new object[] {
																										 ultraGridColumn1,
																										 ultraGridColumn2,
																										 ultraGridColumn3,
																										 ultraGridColumn4,
																										 ultraGridColumn5});
			this.cmbCuentas.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.cmbCuentas.DisplayMember = "Descripcion";
			this.cmbCuentas.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbCuentas.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbCuentas.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbCuentas.Location = new System.Drawing.Point(6, 104);
			this.cmbCuentas.Name = "cmbCuentas";
			this.cmbCuentas.Size = new System.Drawing.Size(280, 21);
			this.cmbCuentas.TabIndex = 2;
			this.cmbCuentas.ValueMember = "idCuenta";
			this.cmbCuentas.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.cmbCuentas_KeyPress);
			this.cmbCuentas.Validating += new System.ComponentModel.CancelEventHandler(this.cmbCuentas_Validating);
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
			// frmParametrosLotes
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 14);
			this.ClientSize = new System.Drawing.Size(292, 168);
			this.Controls.Add(this.label6);
			this.Controls.Add(this.cmbCuentas);
			this.Controls.Add(this.btnCancelar);
			this.Controls.Add(this.btnAceptar);
			this.Controls.Add(this.label14);
			this.Controls.Add(this.txtFondoGarantia);
			this.Controls.Add(this.label13);
			this.Controls.Add(this.txtComision);
			this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
			this.KeyPreview = true;
			this.MaximizeBox = false;
			this.MinimizeBox = false;
			this.Name = "frmParametrosLotes";
			this.Text = "Parametros";
			this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.frmParametrosLotes_KeyDown);
			this.Load += new System.EventHandler(this.frmParametrosLotes_Load);
			((System.ComponentModel.ISupportInitialize)(this.txtFondoGarantia)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtComision)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCuentas)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void frmParametrosLotes_Load(object sender, System.EventArgs e)
		{	
			this.cmbCuentas.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Exec ListaCuentasCierreCaja");
			this.cmbCuentas.DisplayLayout.Bands[0].Columns["Descripcion"].Width = 200;
			this.cmbCuentas.DisplayLayout.Bands[0].Columns["idCuenta"].Hidden = true;

			SqlDataReader dr = Funcion.rEscalarSQL(cdsSeteoF, "Select Comision, FondoDeGarantia, idCuenta From ParametrosCruceCP", true);
			dr.Read();
			if (dr.HasRows)
			{
				this.txtComision.Value = dr.GetDouble(0);
				this.txtFondoGarantia.Value = dr.GetDouble(1);
				this.cmbCuentas.Value = dr.GetInt32(2);
			}
			dr.Close();

			this.txtComision.Select();
		}

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);
		}

		private void btnCancelar_Click(object sender, System.EventArgs e)
		{
			this.DialogResult = DialogResult.No;
		}

		private void frmParametrosLotes_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Escape) this.DialogResult = DialogResult.No;
		}

		private void btnAceptar_Click(object sender, System.EventArgs e)
		{
			this.DialogResult = DialogResult.OK;
		}

		private void txtComision_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (e.KeyChar == 13) this.txtFondoGarantia.Focus();
		}

		private void txtFondoGarantia_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (e.KeyChar == 13) this.cmbCuentas.Focus();
		}

		private void cmbCuentas_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (e.KeyChar == 13) this.btnAceptar.Focus();
		}

		private void txtComision_Click(object sender, System.EventArgs e)
		{
			this.txtComision.SelectAll();
		}

		private void txtComision_Enter(object sender, System.EventArgs e)
		{
			this.txtComision.SelectAll();
		}

		private void txtFondoGarantia_Click(object sender, System.EventArgs e)
		{
			this.txtFondoGarantia.SelectAll();
		}

		private void txtFondoGarantia_Enter(object sender, System.EventArgs e)
		{
			this.txtFondoGarantia.SelectAll();
		}

		private void txtComision_Validating(object sender, System.ComponentModel.CancelEventArgs e)
		{
			if (this.txtComision.Value == System.DBNull.Value)
			{
				MessageBox.Show("Ingrese Porcentaje para Comisión.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.txtComision.Focus();
				e.Cancel = true;
			}
		}

		private void txtFondoGarantia_Validating(object sender, System.ComponentModel.CancelEventArgs e)
		{
			if (this.txtFondoGarantia.Value == System.DBNull.Value)
			{
				MessageBox.Show("Ingrese Porcentaje para Fondo de Garantía.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.txtFondoGarantia.Focus();
				e.Cancel = true;
			}
		}

		private void cmbCuentas_Validating(object sender, System.ComponentModel.CancelEventArgs e)
		{
			if (this.cmbCuentas.Value == System.DBNull.Value)
			{
				MessageBox.Show("Seleccione una Cuenta.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.cmbCuentas.Focus();
				e.Cancel = true;
			}
		}
	}
}

