using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de IngresoContratoAdemdumdet.
	/// </summary>
	public class IngresoContratoAdemdumdet : DevExpress.XtraEditors.XtraForm
	{
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtNotas;
		private Infragistics.Win.Misc.UltraButton btnVerAdemdums;
		private Infragistics.Win.Misc.UltraLabel ultraLabel15;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public IngresoContratoAdemdumdet()
		{
			//
			// Necesario para admitir el Diseñador de Windows Forms
			//
			InitializeComponent();

			//
			// TODO: agregar código de constructor después de llamar a InitializeComponent
			//
		}
		int idContrato=0;
		public IngresoContratoAdemdumdet(int idC)
		{
			//
			// Necesario para admitir el Diseñador de Windows Forms
			//
			InitializeComponent();
			idContrato = idC;

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
			this.txtNotas = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.btnVerAdemdums = new Infragistics.Win.Misc.UltraButton();
			this.ultraLabel15 = new Infragistics.Win.Misc.UltraLabel();
			((System.ComponentModel.ISupportInitialize)(this.txtNotas)).BeginInit();
			this.SuspendLayout();
			// 
			// txtNotas
			// 
			this.txtNotas.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			appearance1.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtNotas.Appearance = appearance1;
			this.txtNotas.Location = new System.Drawing.Point(26, 32);
			this.txtNotas.Multiline = true;
			this.txtNotas.Name = "txtNotas";
			this.txtNotas.Size = new System.Drawing.Size(240, 64);
			this.txtNotas.TabIndex = 84;
			// 
			// btnVerAdemdums
			// 
			this.btnVerAdemdums.BackColor = System.Drawing.Color.Transparent;
			this.btnVerAdemdums.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.btnVerAdemdums.Location = new System.Drawing.Point(88, 104);
			this.btnVerAdemdums.Name = "btnVerAdemdums";
			this.btnVerAdemdums.Size = new System.Drawing.Size(96, 23);
			this.btnVerAdemdums.TabIndex = 144;
			this.btnVerAdemdums.Text = "Ingresar";
			this.btnVerAdemdums.Click += new System.EventHandler(this.btnVerAdemdums_Click);
			// 
			// ultraLabel15
			// 
			this.ultraLabel15.BackColor = System.Drawing.Color.Transparent;
			this.ultraLabel15.Location = new System.Drawing.Point(24, 8);
			this.ultraLabel15.Name = "ultraLabel15";
			this.ultraLabel15.Size = new System.Drawing.Size(128, 16);
			this.ultraLabel15.TabIndex = 145;
			this.ultraLabel15.Text = "Descripcion";
			// 
			// IngresoContratoAdemdumdet
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(292, 133);
			this.Controls.Add(this.ultraLabel15);
			this.Controls.Add(this.btnVerAdemdums);
			this.Controls.Add(this.txtNotas);
			this.MaximizeBox = false;
			this.MinimizeBox = false;
			this.Name = "IngresoContratoAdemdumdet";
			this.Text = "IngresoContratoAdemdumdet";
			this.Load += new System.EventHandler(this.IngresoContratoAdemdumdet_Load);
			((System.ComponentModel.ISupportInitialize)(this.txtNotas)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion
public string texto ="";
		private void btnVerAdemdums_Click(object sender, System.EventArgs e)
		{
			if(txtNotas.Text.Length < 1 )return;
			texto = txtNotas.Text;			
			this.DialogResult = DialogResult.OK;
			this.Close();
		}

		private void IngresoContratoAdemdumdet_Load(object sender, System.EventArgs e)
		{
		
		}
	}
}
