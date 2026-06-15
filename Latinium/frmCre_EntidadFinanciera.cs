using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de frmCre_EntidadFinanciera.
	/// </summary>
	public class frmCre_EntidadFinanciera : System.Windows.Forms.Form
	{
		private System.Windows.Forms.Label label6;
		public Infragistics.Win.UltraWinGrid.UltraCombo cmbEntidadFinanciera;
		private System.Windows.Forms.Button btnCancelar;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;
		private System.Windows.Forms.Button btnAceptar;
		public Infragistics.Win.UltraWinEditors.UltraTextEditor txtNotas;
		private System.Windows.Forms.Label label1;
		private System.Windows.Forms.GroupBox groupBox1;

		int idEntidadFinanciera = 0;

		public frmCre_EntidadFinanciera(int IdEntidadFinanciera)
		{
			//
			// Necesario para admitir el Diseñador de Windows Forms
			//
			InitializeComponent();

			idEntidadFinanciera = IdEntidadFinanciera;

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
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idEntidadFinanciera");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			this.label6 = new System.Windows.Forms.Label();
			this.cmbEntidadFinanciera = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.btnAceptar = new System.Windows.Forms.Button();
			this.btnCancelar = new System.Windows.Forms.Button();
			this.txtNotas = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label1 = new System.Windows.Forms.Label();
			this.groupBox1 = new System.Windows.Forms.GroupBox();
			((System.ComponentModel.ISupportInitialize)(this.cmbEntidadFinanciera)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNotas)).BeginInit();
			this.SuspendLayout();
			// 
			// label6
			// 
			this.label6.AutoSize = true;
			this.label6.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label6.Location = new System.Drawing.Point(8, 8);
			this.label6.Name = "label6";
			this.label6.Size = new System.Drawing.Size(34, 17);
			this.label6.TabIndex = 677;
			this.label6.Text = "Banco";
			this.label6.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// cmbEntidadFinanciera
			// 
			appearance1.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbEntidadFinanciera.Appearance = appearance1;
			this.cmbEntidadFinanciera.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbEntidadFinanciera.Cursor = System.Windows.Forms.Cursors.Default;
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn2.Header.Caption = "Banco";
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridColumn2.Width = 200;
			ultraGridBand1.Columns.AddRange(new object[] {
																										 ultraGridColumn1,
																										 ultraGridColumn2});
			this.cmbEntidadFinanciera.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.cmbEntidadFinanciera.DisplayMember = "Nombre";
			this.cmbEntidadFinanciera.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbEntidadFinanciera.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbEntidadFinanciera.Location = new System.Drawing.Point(64, 8);
			this.cmbEntidadFinanciera.Name = "cmbEntidadFinanciera";
			this.cmbEntidadFinanciera.Size = new System.Drawing.Size(216, 21);
			this.cmbEntidadFinanciera.TabIndex = 678;
			this.cmbEntidadFinanciera.ValueMember = "idEntidadFinanciera";
			// 
			// btnAceptar
			// 
			this.btnAceptar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnAceptar.CausesValidation = false;
			this.btnAceptar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnAceptar.Location = new System.Drawing.Point(8, 152);
			this.btnAceptar.Name = "btnAceptar";
			this.btnAceptar.Size = new System.Drawing.Size(77, 23);
			this.btnAceptar.TabIndex = 679;
			this.btnAceptar.Text = "&Aceptar";
			this.btnAceptar.Click += new System.EventHandler(this.btnAceptar_Click);
			// 
			// btnCancelar
			// 
			this.btnCancelar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnCancelar.CausesValidation = false;
			this.btnCancelar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnCancelar.Location = new System.Drawing.Point(96, 152);
			this.btnCancelar.Name = "btnCancelar";
			this.btnCancelar.Size = new System.Drawing.Size(77, 23);
			this.btnCancelar.TabIndex = 680;
			this.btnCancelar.Text = "&Cancelar";
			this.btnCancelar.Click += new System.EventHandler(this.btnCancelar_Click);
			// 
			// txtNotas
			// 
			appearance2.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtNotas.Appearance = appearance2;
			this.txtNotas.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtNotas.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNotas.Location = new System.Drawing.Point(64, 40);
			this.txtNotas.MaxLength = 8000;
			this.txtNotas.Multiline = true;
			this.txtNotas.Name = "txtNotas";
			this.txtNotas.Size = new System.Drawing.Size(400, 90);
			this.txtNotas.TabIndex = 681;
			// 
			// label1
			// 
			this.label1.AutoSize = true;
			this.label1.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label1.Location = new System.Drawing.Point(8, 40);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(33, 17);
			this.label1.TabIndex = 682;
			this.label1.Text = "Notas";
			this.label1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// groupBox1
			// 
			this.groupBox1.Location = new System.Drawing.Point(0, 136);
			this.groupBox1.Name = "groupBox1";
			this.groupBox1.Size = new System.Drawing.Size(472, 8);
			this.groupBox1.TabIndex = 683;
			this.groupBox1.TabStop = false;
			// 
			// frmCre_EntidadFinanciera
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(472, 182);
			this.ControlBox = false;
			this.Controls.Add(this.groupBox1);
			this.Controls.Add(this.label1);
			this.Controls.Add(this.txtNotas);
			this.Controls.Add(this.btnAceptar);
			this.Controls.Add(this.btnCancelar);
			this.Controls.Add(this.label6);
			this.Controls.Add(this.cmbEntidadFinanciera);
			this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
			this.KeyPreview = true;
			this.Name = "frmCre_EntidadFinanciera";
			this.ShowInTaskbar = false;
			this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
			this.Text = "Seleccione Un Banco";
			this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.frmCre_EntidadFinanciera_KeyDown);
			this.Load += new System.EventHandler(this.frmCre_EntidadFinanciera_Load);
			((System.ComponentModel.ISupportInitialize)(this.cmbEntidadFinanciera)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNotas)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void frmCre_EntidadFinanciera_Load(object sender, System.EventArgs e)
		{
			this.cmbEntidadFinanciera.DataSource = FuncionesProcedimientos.dtGeneral(string.Format("Select idEntidadFinanciera, Nombre From Cre_EntidadFinanciera Where Venta = 1 And idEntidadFinanciera <> {0}", idEntidadFinanciera));
		}

		private void btnCancelar_Click(object sender, System.EventArgs e)
		{
			this.DialogResult = DialogResult.No;
		}

		private void btnAceptar_Click(object sender, System.EventArgs e)
		{
			if (!Validacion.vbComboVacio(this.cmbEntidadFinanciera, "Seleccione un nuevo banco")) return;
			
			if (!Validacion.vbTexto(this.txtNotas, 10, 8000, "La observación debe contener como minimo 20 caracteres")) return;

			this.DialogResult = DialogResult.OK;
		}

		private void frmCre_EntidadFinanciera_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Escape) this.DialogResult = DialogResult.No;
		}
	}
}
