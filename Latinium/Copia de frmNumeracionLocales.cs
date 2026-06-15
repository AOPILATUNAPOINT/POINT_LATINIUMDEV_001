using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using DevExpress.XtraEditors;
using System.Data;
using System.Data.SqlClient;

namespace Latinium
{
	/// <summary>
	/// Summary description for frmNumeracionLocales.
	/// </summary>
	public class frmNumeracionLocales : DevExpress.XtraEditors.XtraForm
	{
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbBodega;
		private System.Windows.Forms.Label lblBodega;
		private System.Windows.Forms.Label lblNumero;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtNumero;
		private System.Windows.Forms.Label label1;
		private System.Windows.Forms.Label label2;
		private DevExpress.XtraEditors.SimpleButton btGuardar;
		private DevExpress.XtraEditors.SimpleButton btSalir;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtReciboCobro;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtDT;
		private C1.Data.C1DataSet cdsSeteoF;
		private DevExpress.XtraEditors.SimpleButton btCancelar;
		private System.Windows.Forms.GroupBox groupBox1;
		private System.Windows.Forms.GroupBox groupBox2;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtGuiaDeRemision;
		private System.Windows.Forms.Label label3;
		/// <summary>
		/// Required designer variable.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public frmNumeracionLocales()
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
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bodega");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo", 0);
			Infragistics.Win.Appearance appearance1 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			this.cmbBodega = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.lblBodega = new System.Windows.Forms.Label();
			this.lblNumero = new System.Windows.Forms.Label();
			this.txtNumero = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label1 = new System.Windows.Forms.Label();
			this.txtReciboCobro = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label2 = new System.Windows.Forms.Label();
			this.txtDT = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.btSalir = new DevExpress.XtraEditors.SimpleButton();
			this.btGuardar = new DevExpress.XtraEditors.SimpleButton();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.btCancelar = new DevExpress.XtraEditors.SimpleButton();
			this.groupBox1 = new System.Windows.Forms.GroupBox();
			this.groupBox2 = new System.Windows.Forms.GroupBox();
			this.txtGuiaDeRemision = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label3 = new System.Windows.Forms.Label();
			((System.ComponentModel.ISupportInitialize)(this.cmbBodega)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNumero)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtReciboCobro)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDT)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtGuiaDeRemision)).BeginInit();
			this.SuspendLayout();
			// 
			// cmbBodega
			// 
			this.cmbBodega.CausesValidation = false;
			this.cmbBodega.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbBodega.Cursor = System.Windows.Forms.Cursors.Default;
			ultraGridColumn1.Header.VisiblePosition = 2;
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridColumn2.Width = 180;
			ultraGridColumn3.Header.VisiblePosition = 0;
			ultraGridBand1.Columns.AddRange(new object[] {
																										 ultraGridColumn1,
																										 ultraGridColumn2,
																										 ultraGridColumn3});
			this.cmbBodega.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.cmbBodega.DisplayMember = "Nombre";
			this.cmbBodega.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbBodega.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbBodega.Enabled = false;
			this.cmbBodega.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbBodega.Location = new System.Drawing.Point(104, 8);
			this.cmbBodega.MaxDropDownItems = 30;
			this.cmbBodega.Name = "cmbBodega";
			this.cmbBodega.Size = new System.Drawing.Size(288, 21);
			this.cmbBodega.TabIndex = 245;
			this.cmbBodega.ValueMember = "Bodega";
			this.cmbBodega.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.cmbBodega_KeyPress);
			this.cmbBodega.ValueChanged += new System.EventHandler(this.cmbBodega_ValueChanged);
			this.cmbBodega.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.cmbBodega_InitializeLayout);
			// 
			// lblBodega
			// 
			this.lblBodega.AutoSize = true;
			this.lblBodega.Location = new System.Drawing.Point(8, 8);
			this.lblBodega.Name = "lblBodega";
			this.lblBodega.Size = new System.Drawing.Size(30, 16);
			this.lblBodega.TabIndex = 244;
			this.lblBodega.Text = "Local";
			this.lblBodega.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// lblNumero
			// 
			this.lblNumero.AutoSize = true;
			this.lblNumero.BackColor = System.Drawing.Color.Transparent;
			this.lblNumero.Location = new System.Drawing.Point(8, 56);
			this.lblNumero.Name = "lblNumero";
			this.lblNumero.Size = new System.Drawing.Size(41, 16);
			this.lblNumero.TabIndex = 247;
			this.lblNumero.Text = "Factura";
			// 
			// txtNumero
			// 
			appearance1.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtNumero.Appearance = appearance1;
			this.txtNumero.CausesValidation = false;
			this.txtNumero.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNumero.Location = new System.Drawing.Point(120, 56);
			this.txtNumero.Name = "txtNumero";
			this.txtNumero.Size = new System.Drawing.Size(104, 22);
			this.txtNumero.TabIndex = 246;
			this.txtNumero.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtNumero_KeyPress);
			// 
			// label1
			// 
			this.label1.AutoSize = true;
			this.label1.BackColor = System.Drawing.Color.Transparent;
			this.label1.Location = new System.Drawing.Point(8, 88);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(86, 16);
			this.label1.TabIndex = 249;
			this.label1.Text = "Recibo de Cobro";
			// 
			// txtReciboCobro
			// 
			appearance2.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtReciboCobro.Appearance = appearance2;
			this.txtReciboCobro.CausesValidation = false;
			this.txtReciboCobro.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtReciboCobro.Location = new System.Drawing.Point(120, 88);
			this.txtReciboCobro.Name = "txtReciboCobro";
			this.txtReciboCobro.Size = new System.Drawing.Size(104, 22);
			this.txtReciboCobro.TabIndex = 248;
			this.txtReciboCobro.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtReciboCobro_KeyPress);
			this.txtReciboCobro.ValueChanged += new System.EventHandler(this.txtReciboCobro_ValueChanged);
			// 
			// label2
			// 
			this.label2.AutoSize = true;
			this.label2.BackColor = System.Drawing.Color.Transparent;
			this.label2.Location = new System.Drawing.Point(8, 152);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(70, 16);
			this.label2.TabIndex = 250;
			this.label2.Text = "Ingreso D. T.";
			// 
			// txtDT
			// 
			appearance3.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtDT.Appearance = appearance3;
			this.txtDT.CausesValidation = false;
			this.txtDT.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtDT.Enabled = false;
			this.txtDT.Location = new System.Drawing.Point(120, 152);
			this.txtDT.Name = "txtDT";
			this.txtDT.Size = new System.Drawing.Size(104, 22);
			this.txtDT.TabIndex = 251;
			this.txtDT.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtDT_KeyPress);
			// 
			// btSalir
			// 
			this.btSalir.Location = new System.Drawing.Point(184, 200);
			this.btSalir.Name = "btSalir";
			this.btSalir.Size = new System.Drawing.Size(80, 26);
			this.btSalir.TabIndex = 253;
			this.btSalir.Text = "&Salir";
			this.btSalir.Click += new System.EventHandler(this.btSalir_Click);
			// 
			// btGuardar
			// 
			this.btGuardar.Location = new System.Drawing.Point(8, 200);
			this.btGuardar.Name = "btGuardar";
			this.btGuardar.Size = new System.Drawing.Size(80, 26);
			this.btGuardar.TabIndex = 252;
			this.btGuardar.Text = "&Guardar";
			this.btGuardar.Click += new System.EventHandler(this.btGuardar_Click);
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
			// btCancelar
			// 
			this.btCancelar.Location = new System.Drawing.Point(96, 200);
			this.btCancelar.Name = "btCancelar";
			this.btCancelar.Size = new System.Drawing.Size(80, 26);
			this.btCancelar.TabIndex = 254;
			this.btCancelar.Text = "&Cancelar";
			this.btCancelar.Click += new System.EventHandler(this.btCancelar_Click);
			// 
			// groupBox1
			// 
			this.groupBox1.Location = new System.Drawing.Point(0, 184);
			this.groupBox1.Name = "groupBox1";
			this.groupBox1.Size = new System.Drawing.Size(400, 8);
			this.groupBox1.TabIndex = 255;
			this.groupBox1.TabStop = false;
			// 
			// groupBox2
			// 
			this.groupBox2.Location = new System.Drawing.Point(0, 40);
			this.groupBox2.Name = "groupBox2";
			this.groupBox2.Size = new System.Drawing.Size(400, 8);
			this.groupBox2.TabIndex = 256;
			this.groupBox2.TabStop = false;
			// 
			// txtGuiaDeRemision
			// 
			appearance4.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtGuiaDeRemision.Appearance = appearance4;
			this.txtGuiaDeRemision.CausesValidation = false;
			this.txtGuiaDeRemision.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtGuiaDeRemision.Location = new System.Drawing.Point(120, 120);
			this.txtGuiaDeRemision.Name = "txtGuiaDeRemision";
			this.txtGuiaDeRemision.Size = new System.Drawing.Size(104, 22);
			this.txtGuiaDeRemision.TabIndex = 258;
			this.txtGuiaDeRemision.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtGuiaDeRemision_KeyPress);
			this.txtGuiaDeRemision.ValueChanged += new System.EventHandler(this.txtGuiaDeRemision_ValueChanged);
			// 
			// label3
			// 
			this.label3.AutoSize = true;
			this.label3.BackColor = System.Drawing.Color.Transparent;
			this.label3.Location = new System.Drawing.Point(8, 120);
			this.label3.Name = "label3";
			this.label3.Size = new System.Drawing.Size(91, 16);
			this.label3.TabIndex = 257;
			this.label3.Text = "Guia de Remisión";
			// 
			// frmNumeracionLocales
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 14);
			this.ClientSize = new System.Drawing.Size(400, 232);
			this.Controls.Add(this.txtGuiaDeRemision);
			this.Controls.Add(this.label3);
			this.Controls.Add(this.groupBox2);
			this.Controls.Add(this.groupBox1);
			this.Controls.Add(this.btCancelar);
			this.Controls.Add(this.btSalir);
			this.Controls.Add(this.btGuardar);
			this.Controls.Add(this.txtDT);
			this.Controls.Add(this.label2);
			this.Controls.Add(this.label1);
			this.Controls.Add(this.lblNumero);
			this.Controls.Add(this.lblBodega);
			this.Controls.Add(this.txtReciboCobro);
			this.Controls.Add(this.txtNumero);
			this.Controls.Add(this.cmbBodega);
			this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
			this.KeyPreview = true;
			this.MaximizeBox = false;
			this.MinimizeBox = false;
			this.Name = "frmNumeracionLocales";
			this.Text = "Numeración de los Locales";
			this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.frmNumeracionLocales_KeyDown);
			this.Load += new System.EventHandler(this.frmNumeracionLocales_Load);
			((System.ComponentModel.ISupportInitialize)(this.cmbBodega)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNumero)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtReciboCobro)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDT)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtGuiaDeRemision)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void frmNumeracionLocales_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Escape) this.Close();
		}

		private void btGuardar_Click(object sender, System.EventArgs e)
		{
			if (this.txtNumero.Text.ToString().Trim().Length == 0) 
			{
				MessageBox.Show("Ingrese el numero para Facturas.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
				this.txtNumero.Focus();
				return;
			}

			if (this.txtReciboCobro.Text.ToString().Trim().Length == 0)
			{
				MessageBox.Show("Ingrese el numero para Recibos de cobro.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
				this.txtReciboCobro.Focus();
				return;
			}

			if (this.txtGuiaDeRemision.Text.ToString().Trim().Length == 0)
			{
				MessageBox.Show("Ingrese el numero para Guias de Remisión.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
				this.txtGuiaDeRemision.Focus();
				return;
			}

			if (DialogResult.No == MessageBox.Show(string.Format("Esta seguro de modificar los secuenciales del local '{0}'?.", this.cmbBodega.Text.ToString().Trim()), "Point Technology", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2)) return;
			
			string sSQL = string.Format("Update Bodega Set NumFactura = '{0}', NumAnticipo = '{1}', NumGuia = '{2}' Where Bodega = {3}", 
				this.txtNumero.Text.ToString(), this.txtReciboCobro.Text.ToString(), this.txtGuiaDeRemision.Text.ToString(), (int)this.cmbBodega.Value);
			Funcion.EjecutaSQL(cdsSeteoF, sSQL, true);
			
			MessageBox.Show(string.Format("Se modifico la Numeración del local '{0}'.", this.cmbBodega.Text.ToString().Trim()), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
		}

		private void btSalir_Click(object sender, System.EventArgs e)
		{
			this.Close();
		}

		private void cmbBodega_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (e.KeyChar == 13) this.txtNumero.Focus();
		}

		private void txtNumero_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (e.KeyChar == 13) this.txtReciboCobro.Focus();
		}

		private void txtReciboCobro_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (e.KeyChar == 13) this.txtGuiaDeRemision.Focus();
		}

		private void txtDT_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (e.KeyChar == 13) this.btGuardar.Focus();
		}

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);		
		}

		private void frmNumeracionLocales_Load(object sender, System.EventArgs e)
		{
			string sSQL = string.Format("Exec ListaBodegas {0}, 1", MenuLatinium.IdUsuario);
			this.cmbBodega.DataSource = Funcion.dvProcedimiento(cdsSeteoF, sSQL);

			sSQL = string.Format("Exec BodegaAsignadaPorUsuario {0}", MenuLatinium.IdUsuario);		
			int IBodega = Funcion.iEscalarSQL(cdsSeteoF, sSQL); 
			if (IBodega == 0)
			{
				this.cmbBodega.Enabled = true;
				this.cmbBodega.Value = 1;
				this.cmbBodega.Focus();
			}
			else
			{
				this.cmbBodega.Value = IBodega;
				this.txtNumero.Focus();
			}
		}

		private void cmbBodega_ValueChanged(object sender, System.EventArgs e)
		{
			string sSQL = string.Format("Select ISNULL(NumFactura, ''), ISNULL(NumAnticipo, ''), ISNULL(NumGuia, '') From Bodega Where Bodega = {0}", (int)this.cmbBodega.Value);
			SqlDataReader dr = Funcion.rEscalarSQL(cdsSeteoF, sSQL, true);
			dr.Read();
			if (dr.HasRows)
			{
				this.txtNumero.Text = dr.GetString(0);
				this.txtReciboCobro.Text = dr.GetString(1);
				this.txtGuiaDeRemision.Text = dr.GetString(2);
			}
			dr.Close();
		}

		private void cmbBodega_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}

		private void btCancelar_Click(object sender, System.EventArgs e)
		{
			string sSQL = string.Format("Select NumFactura, NumAnticipo From Bodega Where Bodega = {0}", (int)this.cmbBodega.Value);
			SqlDataReader dr = Funcion.rEscalarSQL(cdsSeteoF, sSQL, true);
			dr.Read();
			if (dr.HasRows)
			{
				this.txtNumero.Text = dr.GetString(0);
				this.txtReciboCobro.Text = dr.GetString(1);
			}
			dr.Close();
		}

		private void txtReciboCobro_ValueChanged(object sender, System.EventArgs e)
		{
		
		}

		private void txtGuiaDeRemision_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (e.KeyChar == 13) this.btGuardar.Focus();
		}

		private void txtGuiaDeRemision_ValueChanged(object sender, System.EventArgs e)
		{
		
		}
	}
}

