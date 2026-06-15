using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de frmHabitacionExpres.
	/// </summary>
	public class frmHabitacionExpres : System.Windows.Forms.Form
	{
		private System.Windows.Forms.Button btAceptar;
		private System.Windows.Forms.Button btCancelar;
		private System.Windows.Forms.GroupBox groupBox1;
		private Infragistics.Win.UltraWinEditors.UltraOptionSet optTarifa;
		private System.Windows.Forms.Label label17;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public frmHabitacionExpres()
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
			this.btAceptar = new System.Windows.Forms.Button();
			this.btCancelar = new System.Windows.Forms.Button();
			this.groupBox1 = new System.Windows.Forms.GroupBox();
			this.optTarifa = new Infragistics.Win.UltraWinEditors.UltraOptionSet();
			this.label17 = new System.Windows.Forms.Label();
			((System.ComponentModel.ISupportInitialize)(this.optTarifa)).BeginInit();
			this.SuspendLayout();
			// 
			// btAceptar
			// 
			this.btAceptar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btAceptar.CausesValidation = false;
			this.btAceptar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btAceptar.Location = new System.Drawing.Point(8, 64);
			this.btAceptar.Name = "btAceptar";
			this.btAceptar.Size = new System.Drawing.Size(72, 23);
			this.btAceptar.TabIndex = 259;
			this.btAceptar.Text = "&Aceptar";
			this.btAceptar.Click += new System.EventHandler(this.btAceptar_Click);
			// 
			// btCancelar
			// 
			this.btCancelar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btCancelar.CausesValidation = false;
			this.btCancelar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btCancelar.Location = new System.Drawing.Point(88, 64);
			this.btCancelar.Name = "btCancelar";
			this.btCancelar.Size = new System.Drawing.Size(78, 23);
			this.btCancelar.TabIndex = 258;
			this.btCancelar.Text = "&Cancelar";
			this.btCancelar.Click += new System.EventHandler(this.btCancelar_Click);
			// 
			// groupBox1
			// 
			this.groupBox1.Location = new System.Drawing.Point(-488, 48);
			this.groupBox1.Name = "groupBox1";
			this.groupBox1.Size = new System.Drawing.Size(1264, 8);
			this.groupBox1.TabIndex = 445;
			this.groupBox1.TabStop = false;
			// 
			// optTarifa
			// 
			appearance1.ForeColor = System.Drawing.Color.Black;
			appearance1.ForeColorDisabled = System.Drawing.Color.Black;
			this.optTarifa.Appearance = appearance1;
			this.optTarifa.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.optTarifa.CheckedIndex = 0;
			this.optTarifa.ItemAppearance = appearance2;
			this.optTarifa.ItemOrigin = new System.Drawing.Point(10, 0);
			valueListItem1.DataValue = 1;
			valueListItem1.DisplayText = "Normal";
			valueListItem2.DataValue = 2;
			valueListItem2.DisplayText = "Express";
			this.optTarifa.Items.Add(valueListItem1);
			this.optTarifa.Items.Add(valueListItem2);
			this.optTarifa.ItemSpacingVertical = 10;
			this.optTarifa.Location = new System.Drawing.Point(80, 12);
			this.optTarifa.Name = "optTarifa";
			this.optTarifa.Size = new System.Drawing.Size(144, 24);
			this.optTarifa.TabIndex = 704;
			this.optTarifa.Text = "Normal";
			// 
			// label17
			// 
			this.label17.AutoSize = true;
			this.label17.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label17.Location = new System.Drawing.Point(8, 16);
			this.label17.Name = "label17";
			this.label17.Size = new System.Drawing.Size(60, 17);
			this.label17.TabIndex = 705;
			this.label17.Text = "Seleccione:";
			this.label17.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// frmHabitacionExpres
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(240, 94);
			this.Controls.Add(this.label17);
			this.Controls.Add(this.optTarifa);
			this.Controls.Add(this.groupBox1);
			this.Controls.Add(this.btAceptar);
			this.Controls.Add(this.btCancelar);
			this.MaximizeBox = false;
			this.MinimizeBox = false;
			this.Name = "frmHabitacionExpres";
			this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
			this.Text = "Tarifa";
			((System.ComponentModel.ISupportInitialize)(this.optTarifa)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		public int iTarifa = 0;

		private void btCancelar_Click(object sender, System.EventArgs e)
		{
			this.DialogResult = DialogResult.No;
		}

		private void btAceptar_Click(object sender, System.EventArgs e)
		{
			Aceptar();
		}

		private void Aceptar()
		{
			if ((int)optTarifa.Value == 1) iTarifa = 1;
			if ((int)optTarifa.Value == 2) iTarifa = 2;
			DialogResult = DialogResult.OK;
		}		
	}
}
