
using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;
using System.Data.SqlClient;
using System.Globalization;
using System.Runtime.InteropServices;
using System.Diagnostics;
using System.IO;
using Infragistics.Win.UltraWinGrid;
using CrystalDecisions;
using CrystalDecisions.CrystalReports.Engine;
using CrystalDecisions.ReportSource;
using CrystalDecisions.Shared;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de frmCre_RespuestaCedulas.
	/// </summary>
	public class frmCre_RespuestaCedulas : System.Windows.Forms.Form
	{
		public Infragistics.Win.UltraWinEditors.UltraOptionSet optEstados;
		private System.Windows.Forms.Label lblBodega;
		private System.Windows.Forms.GroupBox groupBox1;
		public Infragistics.Win.UltraWinEditors.UltraTextEditor txtObservaciones;
		private System.Windows.Forms.Button btnAceptar;
		private System.Windows.Forms.Button btnCancelar;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public frmCre_RespuestaCedulas()
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
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			this.optEstados = new Infragistics.Win.UltraWinEditors.UltraOptionSet();
			this.lblBodega = new System.Windows.Forms.Label();
			this.txtObservaciones = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.groupBox1 = new System.Windows.Forms.GroupBox();
			this.btnAceptar = new System.Windows.Forms.Button();
			this.btnCancelar = new System.Windows.Forms.Button();
			((System.ComponentModel.ISupportInitialize)(this.optEstados)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtObservaciones)).BeginInit();
			this.SuspendLayout();
			// 
			// optEstados
			// 
			appearance1.ForeColorDisabled = System.Drawing.Color.Black;
			this.optEstados.Appearance = appearance1;
			this.optEstados.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.optEstados.CheckedIndex = 0;
			this.optEstados.ItemAppearance = appearance2;
			this.optEstados.ItemOrigin = new System.Drawing.Point(10, 0);
			valueListItem1.DataValue = 1;
			valueListItem1.DisplayText = "Aplica";
			valueListItem2.DataValue = 3;
			valueListItem2.DisplayText = "No Aplica";
			this.optEstados.Items.Add(valueListItem1);
			this.optEstados.Items.Add(valueListItem2);
			this.optEstados.ItemSpacingVertical = 10;
			this.optEstados.Location = new System.Drawing.Point(8, 8);
			this.optEstados.Name = "optEstados";
			this.optEstados.Size = new System.Drawing.Size(152, 32);
			this.optEstados.TabIndex = 832;
			this.optEstados.Text = "Aplica";
			// 
			// lblBodega
			// 
			this.lblBodega.AutoSize = true;
			this.lblBodega.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblBodega.Location = new System.Drawing.Point(8, 56);
			this.lblBodega.Name = "lblBodega";
			this.lblBodega.Size = new System.Drawing.Size(76, 17);
			this.lblBodega.TabIndex = 833;
			this.lblBodega.Text = "Observaciones";
			this.lblBodega.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtObservaciones
			// 
			this.txtObservaciones.AcceptsReturn = true;
			appearance3.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtObservaciones.Appearance = appearance3;
			this.txtObservaciones.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtObservaciones.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtObservaciones.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtObservaciones.Location = new System.Drawing.Point(8, 80);
			this.txtObservaciones.Multiline = true;
			this.txtObservaciones.Name = "txtObservaciones";
			this.txtObservaciones.Size = new System.Drawing.Size(351, 77);
			this.txtObservaciones.TabIndex = 834;
			// 
			// groupBox1
			// 
			this.groupBox1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.groupBox1.Location = new System.Drawing.Point(0, 160);
			this.groupBox1.Name = "groupBox1";
			this.groupBox1.Size = new System.Drawing.Size(368, 7);
			this.groupBox1.TabIndex = 835;
			this.groupBox1.TabStop = false;
			// 
			// btnAceptar
			// 
			this.btnAceptar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnAceptar.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(230)), ((System.Byte)(230)), ((System.Byte)(230)));
			this.btnAceptar.CausesValidation = false;
			this.btnAceptar.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.btnAceptar.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(50)), ((System.Byte)(50)), ((System.Byte)(50)));
			this.btnAceptar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnAceptar.ImageIndex = 62;
			this.btnAceptar.Location = new System.Drawing.Point(8, 176);
			this.btnAceptar.Name = "btnAceptar";
			this.btnAceptar.Size = new System.Drawing.Size(72, 23);
			this.btnAceptar.TabIndex = 836;
			this.btnAceptar.Text = "Aceptar";
			this.btnAceptar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnAceptar.Click += new System.EventHandler(this.btnAceptar_Click);
			// 
			// btnCancelar
			// 
			this.btnCancelar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnCancelar.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(230)), ((System.Byte)(230)), ((System.Byte)(230)));
			this.btnCancelar.CausesValidation = false;
			this.btnCancelar.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.btnCancelar.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(50)), ((System.Byte)(50)), ((System.Byte)(50)));
			this.btnCancelar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnCancelar.ImageIndex = 62;
			this.btnCancelar.Location = new System.Drawing.Point(88, 176);
			this.btnCancelar.Name = "btnCancelar";
			this.btnCancelar.Size = new System.Drawing.Size(72, 23);
			this.btnCancelar.TabIndex = 837;
			this.btnCancelar.Text = "Cancelar";
			this.btnCancelar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnCancelar.Click += new System.EventHandler(this.btnCancelar_Click);
			// 
			// frmCre_RespuestaCedulas
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(366, 206);
			this.ControlBox = false;
			this.Controls.Add(this.btnCancelar);
			this.Controls.Add(this.btnAceptar);
			this.Controls.Add(this.groupBox1);
			this.Controls.Add(this.txtObservaciones);
			this.Controls.Add(this.lblBodega);
			this.Controls.Add(this.optEstados);
			this.Name = "frmCre_RespuestaCedulas";
			this.Text = "Respuesta Envio de Cedulas";
			this.Load += new System.EventHandler(this.frmCre_RespuestaCedulas_Load);
			((System.ComponentModel.ISupportInitialize)(this.optEstados)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtObservaciones)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void btnAceptar_Click(object sender, System.EventArgs e)
		{
			if (this.txtObservaciones.Text.Length == 0)
			{
				MessageBox.Show("Ingrese una Observación", "Validacion", MessageBoxButtons.OK, MessageBoxIcon.Error);
				this.txtObservaciones.Focus();
				return;
			}

			this.DialogResult = DialogResult.OK;
		}

		private void btnCancelar_Click(object sender, System.EventArgs e)
		{
			this.DialogResult = DialogResult.No;
		}

		private void frmCre_RespuestaCedulas_Load(object sender, System.EventArgs e)
		{
		
		}
	}
}
