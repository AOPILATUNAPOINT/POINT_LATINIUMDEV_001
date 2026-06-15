using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de frmRolPorcentajes.
	/// </summary>
	public class frmRolPorcentajes : System.Windows.Forms.Form
	{
		private System.Windows.Forms.Label lblSubTotal;
		private System.Windows.Forms.Label lblIva;
		private System.Windows.Forms.Label lblDescuento;
		private System.Windows.Forms.GroupBox groupBox1;
		private System.Windows.Forms.Button btnAceptar;
		private System.Windows.Forms.Button btnCancelar;
		public Infragistics.Win.UltraWinEditors.UltraNumericEditor txtHO;
		public Infragistics.Win.UltraWinEditors.UltraNumericEditor txtHE;
		public Infragistics.Win.UltraWinEditors.UltraNumericEditor txtBonos;
		private C1.Data.C1DataSet cdsSeteoF;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		DateTime dtPeriodo = DateTime.Today;

		public frmRolPorcentajes(DateTime DTPeriodo)
		{
			//
			// Necesario para admitir el Diseñador de Windows Forms
			//
			InitializeComponent();
			dtPeriodo = DTPeriodo;

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
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			this.lblSubTotal = new System.Windows.Forms.Label();
			this.lblIva = new System.Windows.Forms.Label();
			this.lblDescuento = new System.Windows.Forms.Label();
			this.txtHO = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtHE = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtBonos = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.groupBox1 = new System.Windows.Forms.GroupBox();
			this.btnCancelar = new System.Windows.Forms.Button();
			this.btnAceptar = new System.Windows.Forms.Button();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			((System.ComponentModel.ISupportInitialize)(this.txtHO)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtHE)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtBonos)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			this.SuspendLayout();
			// 
			// lblSubTotal
			// 
			this.lblSubTotal.AutoSize = true;
			this.lblSubTotal.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblSubTotal.Location = new System.Drawing.Point(6, 10);
			this.lblSubTotal.Name = "lblSubTotal";
			this.lblSubTotal.Size = new System.Drawing.Size(81, 16);
			this.lblSubTotal.TabIndex = 224;
			this.lblSubTotal.Text = "Bono/Comision";
			// 
			// lblIva
			// 
			this.lblIva.AutoSize = true;
			this.lblIva.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblIva.Location = new System.Drawing.Point(6, 58);
			this.lblIva.Name = "lblIva";
			this.lblIva.Size = new System.Drawing.Size(90, 16);
			this.lblIva.TabIndex = 223;
			this.lblIva.Text = "Horas Ordinarias";
			this.lblIva.Click += new System.EventHandler(this.lblIva_Click);
			// 
			// lblDescuento
			// 
			this.lblDescuento.AutoSize = true;
			this.lblDescuento.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblDescuento.Location = new System.Drawing.Point(6, 34);
			this.lblDescuento.Name = "lblDescuento";
			this.lblDescuento.Size = new System.Drawing.Size(116, 16);
			this.lblDescuento.TabIndex = 222;
			this.lblDescuento.Text = "Horas ExtraOrdinarias";
			this.lblDescuento.Click += new System.EventHandler(this.lblDescuento_Click);
			// 
			// txtHO
			// 
			appearance1.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtHO.Appearance = appearance1;
			this.txtHO.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtHO.FormatString = "#,##0.00";
			this.txtHO.Location = new System.Drawing.Point(142, 56);
			this.txtHO.MaxValue = 100;
			this.txtHO.MinValue = 0;
			this.txtHO.Name = "txtHO";
			this.txtHO.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtHO.PromptChar = ' ';
			this.txtHO.Size = new System.Drawing.Size(98, 21);
			this.txtHO.TabIndex = 2;
			this.txtHO.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtHO_KeyDown);
			this.txtHO.Click += new System.EventHandler(this.txtHO_Click);
			this.txtHO.Validated += new System.EventHandler(this.txtHO_Validated);
			this.txtHO.Enter += new System.EventHandler(this.txtHO_Enter);
			// 
			// txtHE
			// 
			appearance2.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtHE.Appearance = appearance2;
			this.txtHE.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtHE.FormatString = "#,##0.00";
			this.txtHE.Location = new System.Drawing.Point(142, 32);
			this.txtHE.MaxValue = 100;
			this.txtHE.MinValue = 0;
			this.txtHE.Name = "txtHE";
			this.txtHE.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtHE.PromptChar = ' ';
			this.txtHE.Size = new System.Drawing.Size(98, 21);
			this.txtHE.TabIndex = 1;
			this.txtHE.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtHE_KeyDown);
			this.txtHE.Click += new System.EventHandler(this.txtHE_Click);
			this.txtHE.Validated += new System.EventHandler(this.txtHE_Validated);
			this.txtHE.Enter += new System.EventHandler(this.txtHE_Enter);
			// 
			// txtBonos
			// 
			appearance3.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtBonos.Appearance = appearance3;
			this.txtBonos.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtBonos.FormatString = "#,##0.00";
			this.txtBonos.Location = new System.Drawing.Point(142, 8);
			this.txtBonos.MaxValue = 100;
			this.txtBonos.MinValue = 0;
			this.txtBonos.Name = "txtBonos";
			this.txtBonos.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtBonos.PromptChar = ' ';
			this.txtBonos.Size = new System.Drawing.Size(98, 21);
			this.txtBonos.TabIndex = 0;
			this.txtBonos.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtBonos_KeyDown);
			this.txtBonos.Click += new System.EventHandler(this.txtBonos_Click);
			this.txtBonos.Validated += new System.EventHandler(this.txtBonos_Validated);
			this.txtBonos.Enter += new System.EventHandler(this.txtBonos_Enter);
			// 
			// groupBox1
			// 
			this.groupBox1.Location = new System.Drawing.Point(0, 88);
			this.groupBox1.Name = "groupBox1";
			this.groupBox1.Size = new System.Drawing.Size(256, 8);
			this.groupBox1.TabIndex = 228;
			this.groupBox1.TabStop = false;
			// 
			// btnCancelar
			// 
			this.btnCancelar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnCancelar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnCancelar.Location = new System.Drawing.Point(88, 112);
			this.btnCancelar.Name = "btnCancelar";
			this.btnCancelar.Size = new System.Drawing.Size(75, 24);
			this.btnCancelar.TabIndex = 687;
			this.btnCancelar.Text = "Cancelar";
			this.btnCancelar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnCancelar.Click += new System.EventHandler(this.btnCancelar_Click);
			// 
			// btnAceptar
			// 
			this.btnAceptar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnAceptar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnAceptar.Location = new System.Drawing.Point(8, 112);
			this.btnAceptar.Name = "btnAceptar";
			this.btnAceptar.Size = new System.Drawing.Size(75, 24);
			this.btnAceptar.TabIndex = 686;
			this.btnAceptar.Text = "Aceptar";
			this.btnAceptar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnAceptar.Click += new System.EventHandler(this.btnAceptar_Click);
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
			// frmRolPorcentajes
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(250, 144);
			this.ControlBox = false;
			this.Controls.Add(this.btnCancelar);
			this.Controls.Add(this.btnAceptar);
			this.Controls.Add(this.groupBox1);
			this.Controls.Add(this.lblSubTotal);
			this.Controls.Add(this.lblIva);
			this.Controls.Add(this.lblDescuento);
			this.Controls.Add(this.txtHO);
			this.Controls.Add(this.txtHE);
			this.Controls.Add(this.txtBonos);
			this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
			this.KeyPreview = true;
			this.Name = "frmRolPorcentajes";
			this.ShowInTaskbar = false;
			this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
			this.Text = "Porcentajes";
			this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.frmRolPorcentajes_KeyDown);
			this.Load += new System.EventHandler(this.frmRolPorcentajes_Load);
			((System.ComponentModel.ISupportInitialize)(this.txtHO)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtHE)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtBonos)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void frmRolPorcentajes_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Escape) this.DialogResult = DialogResult.No;
		}

		private void btnAceptar_Click(object sender, System.EventArgs e)
		{
			if (Convert.ToDecimal(this.txtBonos.Value) == 0.00m)
			{
				MessageBox.Show("Ingrese el Porcentaje para el Bono", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.txtBonos.Focus();
				return;
			}

			if (Convert.ToDecimal(this.txtHE.Value) == 0.00m)
			{
				MessageBox.Show("Ingrese el Porcentaje para el Bono", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.txtHE.Focus();
				return;
			}

			if (Convert.ToDecimal(this.txtHO.Value) == 0.00m)
			{
				MessageBox.Show("Ingrese el Porcentaje para el Bono", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.txtHO.Focus();
				return;
			}

			decimal dPorcentaje = Convert.ToDecimal(this.txtBonos.Value) + Convert.ToDecimal(this.txtHE.Value) + Convert.ToDecimal(this.txtHO.Value);
			if (dPorcentaje != 100)
			{
				MessageBox.Show("La suma de porcentajes debe ser igual a 100", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.txtBonos.Focus();
				return;
			}

			string sSQL = string.Format("Exec NominaValidaPorcentajes {0}, '{1}'", Convert.ToDecimal(this.txtBonos.Value), dtPeriodo.ToString("yyyyMMdd"));
			string sMensaje = Funcion.sEscalarSQL(cdsSeteoF, sSQL);

			if (sMensaje.Length > 0)
			{
				MessageBox.Show(sMensaje, "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.txtBonos.Focus();
				return;
			}

			this.DialogResult = DialogResult.OK;
		}

		private void btnCancelar_Click(object sender, System.EventArgs e)
		{
			this.DialogResult = DialogResult.No;
		}

		private void txtBonos_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.txtHE.Focus();
		}

		private void txtHE_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.txtHO.Focus();
		}

		private void txtHO_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.btnAceptar.Focus();
		}

		private void txtBonos_Click(object sender, System.EventArgs e)
		{
			this.txtBonos.SelectAll();
		}

		private void txtBonos_Enter(object sender, System.EventArgs e)
		{
			this.txtBonos.SelectAll();
		}

		private void txtHE_Click(object sender, System.EventArgs e)
		{
			this.txtHE.SelectAll();
		}

		private void txtHE_Enter(object sender, System.EventArgs e)
		{
			this.txtHE.SelectAll();
		}

		private void txtHO_Click(object sender, System.EventArgs e)
		{
			this.txtHO.SelectAll();
		}

		private void txtHO_Enter(object sender, System.EventArgs e)
		{
			this.txtHO.SelectAll();
		}

		private void lblDescuento_Click(object sender, System.EventArgs e)
		{
		
		}

		private void lblIva_Click(object sender, System.EventArgs e)
		{
		
		}

		private void txtBonos_Validated(object sender, System.EventArgs e)
		{
			if (this.txtBonos.Value == System.DBNull.Value) this.txtBonos.Value = 0.00m;
		}

		private void txtHE_Validated(object sender, System.EventArgs e)
		{
			if (this.txtHE.Value == System.DBNull.Value) this.txtHE.Value = 0.00m;
		}

		private void txtHO_Validated(object sender, System.EventArgs e)
		{
			if (this.txtHO.Value == System.DBNull.Value) this.txtHO.Value = 0.00m;
		}

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);		
		}

		private void frmRolPorcentajes_Load(object sender, System.EventArgs e)
		{
		
		}
	}
}
