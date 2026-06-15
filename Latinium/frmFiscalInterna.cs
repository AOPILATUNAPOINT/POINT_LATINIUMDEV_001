using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using DevExpress.XtraEditors;

namespace Latinium
{
	/// <summary>
	/// Summary description for frmFiscalInterna.
	/// </summary>
	public class frmFiscalInterna : DevExpress.XtraEditors.XtraForm
	{
		private System.Windows.Forms.Button btnAceptar;
		private System.Windows.Forms.Button btnCancelar;
		public Infragistics.Win.UltraWinEditors.UltraCheckEditor chkConsignacion;
		public Infragistics.Win.UltraWinEditors.UltraOptionSet optFiltro;
		private System.Windows.Forms.GroupBox groupBox1;
		/// <summary>
		/// Required designer variable.
		/// </summary>
		private System.ComponentModel.Container components = null;
		public Infragistics.Win.UltraWinEditors.UltraCheckEditor chkAnulado;

		bool bFiscal = false;
		public frmFiscalInterna(bool BFiscal)
		{
			//
			// Required for Windows Form Designer support
			//
			InitializeComponent();

			bFiscal = BFiscal;
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
			Infragistics.Win.ValueListItem valueListItem1 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem2 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			this.btnAceptar = new System.Windows.Forms.Button();
			this.btnCancelar = new System.Windows.Forms.Button();
			this.chkConsignacion = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.optFiltro = new Infragistics.Win.UltraWinEditors.UltraOptionSet();
			this.groupBox1 = new System.Windows.Forms.GroupBox();
			this.chkAnulado = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			((System.ComponentModel.ISupportInitialize)(this.optFiltro)).BeginInit();
			this.SuspendLayout();
			// 
			// btnAceptar
			// 
			this.btnAceptar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnAceptar.CausesValidation = false;
			this.btnAceptar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnAceptar.Location = new System.Drawing.Point(8, 88);
			this.btnAceptar.Name = "btnAceptar";
			this.btnAceptar.TabIndex = 3;
			this.btnAceptar.Text = "&Aceptar";
			this.btnAceptar.Click += new System.EventHandler(this.btnAceptar_Click);
			// 
			// btnCancelar
			// 
			this.btnCancelar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnCancelar.CausesValidation = false;
			this.btnCancelar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnCancelar.Location = new System.Drawing.Point(88, 88);
			this.btnCancelar.Name = "btnCancelar";
			this.btnCancelar.TabIndex = 4;
			this.btnCancelar.Text = "&Cancelar";
			this.btnCancelar.Click += new System.EventHandler(this.btnCancelar_Click);
			// 
			// chkConsignacion
			// 
			appearance1.ForeColorDisabled = System.Drawing.Color.Black;
			this.chkConsignacion.Appearance = appearance1;
			this.chkConsignacion.BackColor = System.Drawing.Color.Transparent;
			this.chkConsignacion.CausesValidation = false;
			this.chkConsignacion.Location = new System.Drawing.Point(149, 40);
			this.chkConsignacion.Name = "chkConsignacion";
			this.chkConsignacion.Size = new System.Drawing.Size(96, 22);
			this.chkConsignacion.TabIndex = 627;
			this.chkConsignacion.Text = "Consignación";
			// 
			// optFiltro
			// 
			appearance2.ForeColorDisabled = System.Drawing.Color.Black;
			this.optFiltro.Appearance = appearance2;
			this.optFiltro.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.optFiltro.CheckedIndex = 1;
			this.optFiltro.Enabled = false;
			this.optFiltro.ItemAppearance = appearance3;
			this.optFiltro.ItemOrigin = new System.Drawing.Point(10, 0);
			valueListItem1.DataValue = true;
			valueListItem1.DisplayText = "Documento Fiscal";
			valueListItem2.DataValue = false;
			valueListItem2.DisplayText = "Documento Interno";
			this.optFiltro.Items.Add(valueListItem1);
			this.optFiltro.Items.Add(valueListItem2);
			this.optFiltro.ItemSpacingVertical = 10;
			this.optFiltro.Location = new System.Drawing.Point(9, 8);
			this.optFiltro.Name = "optFiltro";
			this.optFiltro.Size = new System.Drawing.Size(248, 24);
			this.optFiltro.TabIndex = 628;
			this.optFiltro.Text = "Documento Interno";
			this.optFiltro.ValueChanged += new System.EventHandler(this.optFiltro_ValueChanged);
			// 
			// groupBox1
			// 
			this.groupBox1.Location = new System.Drawing.Point(0, 72);
			this.groupBox1.Name = "groupBox1";
			this.groupBox1.Size = new System.Drawing.Size(272, 8);
			this.groupBox1.TabIndex = 629;
			this.groupBox1.TabStop = false;
			// 
			// chkAnulado
			// 
			appearance4.ForeColorDisabled = System.Drawing.Color.Black;
			this.chkAnulado.Appearance = appearance4;
			this.chkAnulado.BackColor = System.Drawing.Color.Transparent;
			this.chkAnulado.CausesValidation = false;
			this.chkAnulado.Enabled = false;
			this.chkAnulado.Location = new System.Drawing.Point(21, 40);
			this.chkAnulado.Name = "chkAnulado";
			this.chkAnulado.Size = new System.Drawing.Size(96, 22);
			this.chkAnulado.TabIndex = 630;
			this.chkAnulado.Text = "Anulado";
			this.chkAnulado.Visible = false;
			this.chkAnulado.CheckedChanged += new System.EventHandler(this.chkAnulado_CheckedChanged);
			// 
			// frmFiscalInterna
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 14);
			this.ClientSize = new System.Drawing.Size(266, 120);
			this.ControlBox = false;
			this.Controls.Add(this.chkAnulado);
			this.Controls.Add(this.groupBox1);
			this.Controls.Add(this.optFiltro);
			this.Controls.Add(this.chkConsignacion);
			this.Controls.Add(this.btnAceptar);
			this.Controls.Add(this.btnCancelar);
			this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
			this.KeyPreview = true;
			this.MaximizeBox = false;
			this.MinimizeBox = false;
			this.Name = "frmFiscalInterna";
			this.ShowInTaskbar = false;
			this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
			this.Text = "Tipo de Documento";
			this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.frmFiscalInterna_KeyDown);
			this.Load += new System.EventHandler(this.frmFiscalInterna_Load);
			((System.ComponentModel.ISupportInitialize)(this.optFiltro)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void frmFiscalInterna_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			this.DialogResult = DialogResult.No;
		}

		private void btnAceptar_Click(object sender, System.EventArgs e)
		{
			this.DialogResult = DialogResult.OK;
		}

		private void btnCancelar_Click(object sender, System.EventArgs e)
		{
			this.DialogResult = DialogResult.No;
		}

		private void optFiltro_ValueChanged(object sender, System.EventArgs e)
		{
			if ((bool)this.optFiltro.Value) 
			{
				//this.chkAnulado.Enabled = true;
				this.chkConsignacion.Checked = false;
				this.chkConsignacion.Enabled = false;
			}
			else 
			{
				this.chkConsignacion.Enabled = true;
//				this.chkAnulado.Checked = false;
//				this.chkAnulado.Enabled = false;				
			}
		}

		private void frmFiscalInterna_Load(object sender, System.EventArgs e)
		{
			if (bFiscal) this.optFiltro.Enabled = true;			
		}

		private void chkAnulado_CheckedChanged(object sender, System.EventArgs e)
		{
			
		}
	}
}

