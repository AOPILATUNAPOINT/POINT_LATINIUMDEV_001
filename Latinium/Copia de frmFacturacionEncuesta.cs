using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de frmFacturacionEncuesta.
	/// </summary>
	public class frmFacturacionEncuesta : System.Windows.Forms.Form
	{
		public Infragistics.Win.UltraWinEditors.UltraOptionSet optEncuesta;
		private System.Windows.Forms.Button btnAceptar;
		private System.Windows.Forms.Button btnCancelar;
		private System.Windows.Forms.GroupBox groupBox1;
		public Infragistics.Win.UltraWinEditors.UltraTextEditor txtOtros;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public frmFacturacionEncuesta()
		{
			//
			// Necesario para admitir el Diseñador de Windows Forms
			//
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
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.ValueListItem valueListItem1 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem2 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem3 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem4 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem5 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem6 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem7 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem8 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem9 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem10 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem11 = new Infragistics.Win.ValueListItem();
			this.optEncuesta = new Infragistics.Win.UltraWinEditors.UltraOptionSet();
			this.btnAceptar = new System.Windows.Forms.Button();
			this.btnCancelar = new System.Windows.Forms.Button();
			this.txtOtros = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.groupBox1 = new System.Windows.Forms.GroupBox();
			((System.ComponentModel.ISupportInitialize)(this.optEncuesta)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtOtros)).BeginInit();
			this.SuspendLayout();
			// 
			// optEncuesta
			// 
			appearance1.ForeColorDisabled = System.Drawing.Color.Black;
			this.optEncuesta.Appearance = appearance1;
			this.optEncuesta.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.optEncuesta.ItemAppearance = appearance2;
			this.optEncuesta.ItemOrigin = new System.Drawing.Point(10, 0);
			valueListItem1.DataValue = 1;
			valueListItem1.DisplayText = "FACEBOOK";
			valueListItem2.DataValue = 2;
			valueListItem2.DisplayText = "CATALOGO";
			valueListItem3.DataValue = 3;
			valueListItem3.DisplayText = "RADIO";
			valueListItem4.DataValue = 4;
			valueListItem4.DisplayText = "RECOMENDADO";
			valueListItem5.DataValue = 5;
			valueListItem5.DisplayText = "REITERATIVO";
			valueListItem6.DataValue = 7;
			valueListItem6.DisplayText = "PASO POR EL LOCAL";
			valueListItem7.DataValue = 8;
			valueListItem7.DisplayText = "PUERTEO";
			valueListItem8.DataValue = 9;
			valueListItem8.DisplayText = "SMS";
			valueListItem9.DataValue = 10;
			valueListItem9.DisplayText = "WEB";
			valueListItem10.DataValue = 11;
			valueListItem10.DisplayText = "WHATSAPP";
			valueListItem11.DataValue = 6;
			valueListItem11.DisplayText = "OTROS";
			this.optEncuesta.Items.Add(valueListItem1);
			this.optEncuesta.Items.Add(valueListItem2);
			this.optEncuesta.Items.Add(valueListItem3);
			this.optEncuesta.Items.Add(valueListItem4);
			this.optEncuesta.Items.Add(valueListItem5);
			this.optEncuesta.Items.Add(valueListItem6);
			this.optEncuesta.Items.Add(valueListItem7);
			this.optEncuesta.Items.Add(valueListItem8);
			this.optEncuesta.Items.Add(valueListItem9);
			this.optEncuesta.Items.Add(valueListItem10);
			this.optEncuesta.Items.Add(valueListItem11);
			this.optEncuesta.ItemSpacingVertical = 10;
			this.optEncuesta.Location = new System.Drawing.Point(8, 8);
			this.optEncuesta.Name = "optEncuesta";
			this.optEncuesta.Size = new System.Drawing.Size(184, 264);
			this.optEncuesta.TabIndex = 0;
			this.optEncuesta.ValueChanged += new System.EventHandler(this.optEncuesta_ValueChanged);
			// 
			// btnAceptar
			// 
			this.btnAceptar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnAceptar.Location = new System.Drawing.Point(8, 318);
			this.btnAceptar.Name = "btnAceptar";
			this.btnAceptar.Size = new System.Drawing.Size(88, 23);
			this.btnAceptar.TabIndex = 2;
			this.btnAceptar.Text = "&Aceptar";
			this.btnAceptar.Click += new System.EventHandler(this.btnAceptar_Click);
			// 
			// btnCancelar
			// 
			this.btnCancelar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnCancelar.Location = new System.Drawing.Point(104, 318);
			this.btnCancelar.Name = "btnCancelar";
			this.btnCancelar.Size = new System.Drawing.Size(88, 23);
			this.btnCancelar.TabIndex = 3;
			this.btnCancelar.Text = "&Cancelar";
			this.btnCancelar.Click += new System.EventHandler(this.btnCancelar_Click);
			// 
			// txtOtros
			// 
			this.txtOtros.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtOtros.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtOtros.Enabled = false;
			this.txtOtros.Location = new System.Drawing.Point(8, 280);
			this.txtOtros.MaxLength = 20;
			this.txtOtros.Name = "txtOtros";
			this.txtOtros.Size = new System.Drawing.Size(184, 21);
			this.txtOtros.TabIndex = 1;
			this.txtOtros.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtOtros_KeyDown);
			// 
			// groupBox1
			// 
			this.groupBox1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.groupBox1.Location = new System.Drawing.Point(0, 304);
			this.groupBox1.Name = "groupBox1";
			this.groupBox1.Size = new System.Drawing.Size(202, 8);
			this.groupBox1.TabIndex = 141;
			this.groupBox1.TabStop = false;
			// 
			// frmFacturacionEncuesta
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(202, 344);
			this.ControlBox = false;
			this.Controls.Add(this.groupBox1);
			this.Controls.Add(this.btnAceptar);
			this.Controls.Add(this.btnCancelar);
			this.Controls.Add(this.txtOtros);
			this.Controls.Add(this.optEncuesta);
			this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
			this.KeyPreview = true;
			this.Name = "frmFacturacionEncuesta";
			this.ShowInTaskbar = false;
			this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
			this.Text = "Encuesta";
			this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.frmFacturacionEncuesta_KeyDown);
			this.Load += new System.EventHandler(this.frmFacturacionEncuesta_Load);
			((System.ComponentModel.ISupportInitialize)(this.optEncuesta)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtOtros)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void frmFacturacionEncuesta_Load(object sender, System.EventArgs e)
		{
			this.optEncuesta.Value = System.DBNull.Value;
		}

		private void btnCancelar_Click(object sender, System.EventArgs e)
		{
			this.DialogResult = DialogResult.No;
		}

		private void frmFacturacionEncuesta_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Escape) this.DialogResult = DialogResult.No;
		}

		private void btnAceptar_Click(object sender, System.EventArgs e)
		{
			if (this.optEncuesta.CheckedIndex == - 1)
			{
				MessageBox.Show("Seleccione una opción", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return;
			}

			if ((int)this.optEncuesta.Value == 3 || (int)this.optEncuesta.Value == 6 || (int)this.optEncuesta.Value == 9)
				if (!Validacion.vbTexto(this.txtOtros, 10, 20, "Notas")) return;			
				
			this.DialogResult = DialogResult.OK;
		}

		private void optEncuesta_ValueChanged(object sender, System.EventArgs e)
		{
			if ((int)this.optEncuesta.Value == 3 || (int)this.optEncuesta.Value == 6 || (int)this.optEncuesta.Value == 9) 
			{
				this.txtOtros.Enabled = true;
				this.txtOtros.Focus();
			}
			else
			{
				this.txtOtros.Text = "";
				this.txtOtros.Enabled = false;
			}
		}

		private void txtOtros_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Escape) this.btnAceptar.Focus();
		}
	}
}
