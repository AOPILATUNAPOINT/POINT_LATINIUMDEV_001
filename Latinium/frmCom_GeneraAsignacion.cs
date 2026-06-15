using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de frmCom_GeneraAsignacion.
	/// </summary>
	public class frmCom_GeneraAsignacion : System.Windows.Forms.Form
	{
		private Infragistics.Win.UltraWinEditors.UltraOptionSet optFiltro;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public frmCom_GeneraAsignacion()
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
			this.optFiltro = new Infragistics.Win.UltraWinEditors.UltraOptionSet();
			((System.ComponentModel.ISupportInitialize)(this.optFiltro)).BeginInit();
			this.SuspendLayout();
			// 
			// optFiltro
			// 
			appearance1.ForeColorDisabled = System.Drawing.Color.Black;
			this.optFiltro.Appearance = appearance1;
			this.optFiltro.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.optFiltro.ItemAppearance = appearance2;
			this.optFiltro.ItemOrigin = new System.Drawing.Point(10, 0);
			valueListItem1.DataValue = 1;
			valueListItem1.DisplayText = "(F7) Con Existencia";
			valueListItem2.DataValue = 2;
			valueListItem2.DisplayText = "(F8) Sin Existencia";
			valueListItem3.DataValue = 3;
			valueListItem3.DisplayText = "(F9) Todos";
			this.optFiltro.Items.Add(valueListItem1);
			this.optFiltro.Items.Add(valueListItem2);
			this.optFiltro.Items.Add(valueListItem3);
			this.optFiltro.ItemSpacingVertical = 10;
			this.optFiltro.Location = new System.Drawing.Point(8, 16);
			this.optFiltro.Name = "optFiltro";
			this.optFiltro.Size = new System.Drawing.Size(336, 26);
			this.optFiltro.TabIndex = 137;
			// 
			// frmCom_GeneraAsignacion
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(336, 158);
			this.Controls.Add(this.optFiltro);
			this.KeyPreview = true;
			this.MaximizeBox = false;
			this.MinimizeBox = false;
			this.Name = "frmCom_GeneraAsignacion";
			this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
			((System.ComponentModel.ISupportInitialize)(this.optFiltro)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion
	}
}
