using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de frmdescuento.
	/// </summary>
	public class frmdescuento : System.Windows.Forms.Form
	{
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtnumero;
		private System.Windows.Forms.Button btnNuevo;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtdesc;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txttotal;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public frmdescuento()
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
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmdescuento));
			this.txtnumero = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtdesc = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txttotal = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.btnNuevo = new System.Windows.Forms.Button();
			((System.ComponentModel.ISupportInitialize)(this.txtnumero)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtdesc)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txttotal)).BeginInit();
			this.SuspendLayout();
			// 
			// txtnumero
			// 
			appearance1.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtnumero.Appearance = appearance1;
			this.txtnumero.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtnumero.Location = new System.Drawing.Point(48, 24);
			this.txtnumero.Name = "txtnumero";
			this.txtnumero.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtnumero.PromptChar = ' ';
			this.txtnumero.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
			this.txtnumero.Size = new System.Drawing.Size(80, 21);
			this.txtnumero.SpinButtonAlignment = Infragistics.Win.ButtonAlignment.Left;
			this.txtnumero.TabIndex = 348;
			// 
			// txtdesc
			// 
			appearance2.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtdesc.Appearance = appearance2;
			this.txtdesc.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtdesc.Location = new System.Drawing.Point(48, 64);
			this.txtdesc.Name = "txtdesc";
			this.txtdesc.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtdesc.PromptChar = ' ';
			this.txtdesc.Size = new System.Drawing.Size(80, 21);
			this.txtdesc.TabIndex = 349;
			this.txtdesc.ValueChanged += new System.EventHandler(this.ultraNumericEditor2_ValueChanged);
			// 
			// txttotal
			// 
			appearance3.ForeColorDisabled = System.Drawing.Color.Black;
			this.txttotal.Appearance = appearance3;
			this.txttotal.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txttotal.Location = new System.Drawing.Point(48, 104);
			this.txttotal.Name = "txttotal";
			this.txttotal.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txttotal.PromptChar = ' ';
			this.txttotal.Size = new System.Drawing.Size(80, 21);
			this.txttotal.TabIndex = 350;
			// 
			// btnNuevo
			// 
			this.btnNuevo.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnNuevo.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(230)), ((System.Byte)(230)), ((System.Byte)(230)));
			this.btnNuevo.CausesValidation = false;
			this.btnNuevo.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.btnNuevo.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(50)), ((System.Byte)(50)), ((System.Byte)(50)));
			this.btnNuevo.Image = ((System.Drawing.Image)(resources.GetObject("btnNuevo.Image")));
			this.btnNuevo.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnNuevo.Location = new System.Drawing.Point(48, 168);
			this.btnNuevo.Name = "btnNuevo";
			this.btnNuevo.Size = new System.Drawing.Size(72, 23);
			this.btnNuevo.TabIndex = 351;
			this.btnNuevo.Text = "Nuevo";
			this.btnNuevo.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnNuevo.Click += new System.EventHandler(this.btnNuevo_Click);
			// 
			// frmdescuento
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(528, 446);
			this.Controls.Add(this.btnNuevo);
			this.Controls.Add(this.txttotal);
			this.Controls.Add(this.txtdesc);
			this.Controls.Add(this.txtnumero);
			this.Name = "frmdescuento";
			this.Text = "frmdescuento";
			((System.ComponentModel.ISupportInitialize)(this.txtnumero)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtdesc)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txttotal)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void ultraNumericEditor2_ValueChanged(object sender, System.EventArgs e)
		{
		
		}

		private void btnNuevo_Click(object sender, System.EventArgs e)
		{
//			this.txtver.Value = 125.90;
//			
//			double num = 125.49;
//			int p_ent = (int) num;
//			double p_dec = num - p_ent;
//			if(p_dec < 0.50)
//			{
//				this.txtprobar.Value= p_ent;}
//			else 
//			{
//				this.txtprobar.Value =p_ent +1;
//			}


     double num = (double) this.txtnumero.Value;
			int p_ente = (int) num;
			double p_dec = num - p_ente;
			if( p_dec <0.50)
			{
				this.txtnumero.Value =p_ente;
			}
			else{
			this.txtnumero.Value =p_ente + 1;
			} 
      this.txttotal.Value= ( (double)this.txtnumero.Value * (double) this.txtdesc.Value )/100; 
      double iVerificar; 

		}
	}
}
