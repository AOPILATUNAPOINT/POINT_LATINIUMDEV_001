using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using DevExpress.XtraEditors;

namespace Latinium
{
	/// <summary>
	/// Summary description for frmRegistraConfigRetenciones.
	/// </summary>
	public class frmRegistraConfigRetenciones : DevExpress.XtraEditors.XtraForm
	{
		private System.Windows.Forms.Label lblIce;
		private System.Windows.Forms.Label label6;
		private System.Windows.Forms.Label label1;
		private System.Windows.Forms.Label label2;
		private System.Windows.Forms.Button btnGuardar;
		private System.Windows.Forms.Button btnCancelar;
		private C1.Data.C1DataSet cdsSeteoF;
		private System.Windows.Forms.GroupBox groupBox1;
		public Infragistics.Win.UltraWinEditors.UltraNumericEditor txtPorcBienes;
		public Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtDesde;
		public Infragistics.Win.UltraWinEditors.UltraNumericEditor txtPorcServicios;
		public Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtHasta;
		/// <summary>
		/// Required designer variable.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public frmRegistraConfigRetenciones()
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
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton1 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmRegistraConfigRetenciones));
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton2 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			this.lblIce = new System.Windows.Forms.Label();
			this.label6 = new System.Windows.Forms.Label();
			this.txtPorcBienes = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.dtDesde = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.label1 = new System.Windows.Forms.Label();
			this.label2 = new System.Windows.Forms.Label();
			this.txtPorcServicios = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.dtHasta = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.btnGuardar = new System.Windows.Forms.Button();
			this.btnCancelar = new System.Windows.Forms.Button();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.groupBox1 = new System.Windows.Forms.GroupBox();
			((System.ComponentModel.ISupportInitialize)(this.txtPorcBienes)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtDesde)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtPorcServicios)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtHasta)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			this.SuspendLayout();
			// 
			// lblIce
			// 
			this.lblIce.AutoSize = true;
			this.lblIce.Location = new System.Drawing.Point(16, 72);
			this.lblIce.Name = "lblIce";
			this.lblIce.Size = new System.Drawing.Size(51, 16);
			this.lblIce.TabIndex = 320;
			this.lblIce.Text = "% Bienes";
			this.lblIce.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label6
			// 
			this.label6.AutoSize = true;
			this.label6.Location = new System.Drawing.Point(16, 8);
			this.label6.Name = "label6";
			this.label6.Size = new System.Drawing.Size(35, 16);
			this.label6.TabIndex = 319;
			this.label6.Text = "Desde";
			this.label6.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtPorcBienes
			// 
			appearance1.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtPorcBienes.Appearance = appearance1;
			this.txtPorcBienes.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtPorcBienes.FormatString = "#,##0.00";
			this.txtPorcBienes.Location = new System.Drawing.Point(112, 72);
			this.txtPorcBienes.Name = "txtPorcBienes";
			this.txtPorcBienes.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtPorcBienes.PromptChar = ' ';
			this.txtPorcBienes.Size = new System.Drawing.Size(112, 22);
			this.txtPorcBienes.TabIndex = 2;
			this.txtPorcBienes.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtPorcBienes_KeyDown);
			this.txtPorcBienes.Click += new System.EventHandler(this.txtPorcBienes_Click);
			this.txtPorcBienes.Validated += new System.EventHandler(this.txtPorcBienes_Validated);
			this.txtPorcBienes.Enter += new System.EventHandler(this.txtPorcBienes_Enter);
			// 
			// dtDesde
			// 
			appearance2.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtDesde.Appearance = appearance2;
			this.dtDesde.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton1.Caption = "Today";
			this.dtDesde.DateButtons.Add(dateButton1);
			this.dtDesde.Enabled = false;
			this.dtDesde.Format = "dd/MM/yyyy";
			this.dtDesde.Location = new System.Drawing.Point(112, 8);
			this.dtDesde.Name = "dtDesde";
			this.dtDesde.NonAutoSizeHeight = 23;
			this.dtDesde.Size = new System.Drawing.Size(112, 21);
			this.dtDesde.SpinButtonsVisible = true;
			this.dtDesde.TabIndex = 0;
			this.dtDesde.Value = ((object)(resources.GetObject("dtDesde.Value")));
			// 
			// label1
			// 
			this.label1.AutoSize = true;
			this.label1.Location = new System.Drawing.Point(16, 104);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(62, 16);
			this.label1.TabIndex = 324;
			this.label1.Text = "% Servicios";
			this.label1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label2
			// 
			this.label2.AutoSize = true;
			this.label2.Location = new System.Drawing.Point(16, 40);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(33, 16);
			this.label2.TabIndex = 323;
			this.label2.Text = "Hasta";
			this.label2.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtPorcServicios
			// 
			appearance3.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtPorcServicios.Appearance = appearance3;
			this.txtPorcServicios.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtPorcServicios.FormatString = "#,##0.00";
			this.txtPorcServicios.Location = new System.Drawing.Point(112, 104);
			this.txtPorcServicios.Name = "txtPorcServicios";
			this.txtPorcServicios.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtPorcServicios.PromptChar = ' ';
			this.txtPorcServicios.Size = new System.Drawing.Size(112, 22);
			this.txtPorcServicios.TabIndex = 3;
			this.txtPorcServicios.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtPorcServicios_KeyDown);
			this.txtPorcServicios.Click += new System.EventHandler(this.txtPorcServicios_Click);
			this.txtPorcServicios.Validated += new System.EventHandler(this.txtPorcServicios_Validated);
			this.txtPorcServicios.ValueChanged += new System.EventHandler(this.txtPorcServicios_ValueChanged);
			this.txtPorcServicios.Enter += new System.EventHandler(this.txtPorcServicios_Enter);
			// 
			// dtHasta
			// 
			appearance4.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtHasta.Appearance = appearance4;
			this.dtHasta.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton2.Caption = "Today";
			this.dtHasta.DateButtons.Add(dateButton2);
			this.dtHasta.Format = "dd/MM/yyyy";
			this.dtHasta.Location = new System.Drawing.Point(112, 40);
			this.dtHasta.Name = "dtHasta";
			this.dtHasta.NonAutoSizeHeight = 23;
			this.dtHasta.Size = new System.Drawing.Size(112, 21);
			this.dtHasta.SpinButtonsVisible = true;
			this.dtHasta.TabIndex = 1;
			this.dtHasta.Value = ((object)(resources.GetObject("dtHasta.Value")));
			this.dtHasta.KeyDown += new System.Windows.Forms.KeyEventHandler(this.dtHasta_KeyDown);
			this.dtHasta.BeforeDropDown += new System.ComponentModel.CancelEventHandler(this.dtHasta_BeforeDropDown);
			// 
			// btnGuardar
			// 
			this.btnGuardar.Image = ((System.Drawing.Image)(resources.GetObject("btnGuardar.Image")));
			this.btnGuardar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnGuardar.Location = new System.Drawing.Point(44, 152);
			this.btnGuardar.Name = "btnGuardar";
			this.btnGuardar.Size = new System.Drawing.Size(76, 23);
			this.btnGuardar.TabIndex = 4;
			this.btnGuardar.Text = "Guardar";
			this.btnGuardar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnGuardar.Click += new System.EventHandler(this.btnGuardar_Click);
			// 
			// btnCancelar
			// 
			this.btnCancelar.CausesValidation = false;
			this.btnCancelar.Image = ((System.Drawing.Image)(resources.GetObject("btnCancelar.Image")));
			this.btnCancelar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnCancelar.Location = new System.Drawing.Point(124, 152);
			this.btnCancelar.Name = "btnCancelar";
			this.btnCancelar.Size = new System.Drawing.Size(72, 23);
			this.btnCancelar.TabIndex = 5;
			this.btnCancelar.Text = "Cancelar";
			this.btnCancelar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
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
			// groupBox1
			// 
			this.groupBox1.Location = new System.Drawing.Point(0, 136);
			this.groupBox1.Name = "groupBox1";
			this.groupBox1.Size = new System.Drawing.Size(240, 8);
			this.groupBox1.TabIndex = 341;
			this.groupBox1.TabStop = false;
			// 
			// frmRegistraConfigRetenciones
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 14);
			this.ClientSize = new System.Drawing.Size(240, 182);
			this.Controls.Add(this.groupBox1);
			this.Controls.Add(this.btnGuardar);
			this.Controls.Add(this.btnCancelar);
			this.Controls.Add(this.label1);
			this.Controls.Add(this.label2);
			this.Controls.Add(this.txtPorcServicios);
			this.Controls.Add(this.dtHasta);
			this.Controls.Add(this.lblIce);
			this.Controls.Add(this.label6);
			this.Controls.Add(this.txtPorcBienes);
			this.Controls.Add(this.dtDesde);
			this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
			this.KeyPreview = true;
			this.MaximizeBox = false;
			this.MinimizeBox = false;
			this.Name = "frmRegistraConfigRetenciones";
			this.ShowInTaskbar = false;
			this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
			this.Text = "Nuevo";
			this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.frmRegistraConfigRetenciones_KeyDown);
			this.Load += new System.EventHandler(this.frmRegistraConfigRetenciones_Load);
			((System.ComponentModel.ISupportInitialize)(this.txtPorcBienes)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtDesde)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtPorcServicios)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtHasta)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void btnCancelar_Click(object sender, System.EventArgs e)
		{
			this.DialogResult = DialogResult.No;
		}

		private void btnGuardar_Click(object sender, System.EventArgs e)
		{
			if (this.dtHasta.Value == System.DBNull.Value)
			{
				MessageBox.Show("Seleccione la Fecha Final Recuerde que no puede ser menor a la Fecha Inicial", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
			}
			else if ((double)this.txtPorcBienes.Value == 0)
			{
				MessageBox.Show("Ingrese el Porcentaje Para Bienes", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
			}
			else if ((double)this.txtPorcServicios.Value == 0)
			{
				MessageBox.Show("Ingrese el Porcentaje Para Servicios", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
			}
			else this.DialogResult = DialogResult.OK;
		}

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);		
		}

		private void frmRegistraConfigRetenciones_Load(object sender, System.EventArgs e)
		{
			this.dtDesde.Value = Funcion.dtEscalarSQL(cdsSeteoF, "Exec RetornaUltimaFechaConfRetIVA");

			this.dtHasta.CalendarInfo.MinDate = (DateTime)this.dtDesde.Value;			
		}

		private void dtHasta_BeforeDropDown(object sender, System.ComponentModel.CancelEventArgs e)
		{
		
		}

		private void dtHasta_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.txtPorcBienes.Focus();
		}

		private void txtPorcBienes_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.txtPorcServicios.Focus();
		}

		private void txtPorcServicios_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.btnGuardar.Focus();
		}

		private void txtPorcBienes_Validated(object sender, System.EventArgs e)
		{
			if (this.txtPorcBienes.Value == System.DBNull.Value) this.txtPorcBienes.Value = 0;			
		}

		private void txtPorcServicios_Validated(object sender, System.EventArgs e)
		{
			if (this.txtPorcServicios.Value == System.DBNull.Value) this.txtPorcServicios.Value = 0;
		}

		private void txtPorcBienes_Click(object sender, System.EventArgs e)
		{
			this.txtPorcBienes.SelectAll();
		}

		private void txtPorcBienes_Enter(object sender, System.EventArgs e)
		{
			this.txtPorcBienes.SelectAll();
		}

		private void txtPorcServicios_Enter(object sender, System.EventArgs e)
		{
			this.txtPorcServicios.SelectAll();
		}

		private void txtPorcServicios_Click(object sender, System.EventArgs e)
		{
			this.txtPorcServicios.SelectAll();
		}

		private void frmRegistraConfigRetenciones_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Escape) this.DialogResult = DialogResult.No;
		}

		private void txtPorcServicios_ValueChanged(object sender, System.EventArgs e)
		{
		
		}
	}
}

