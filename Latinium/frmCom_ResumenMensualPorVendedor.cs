using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using Infragistics.Win.UltraWinGrid;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de frmCom_ResumenMensualPorVendedor.
	/// </summary>
	public class frmCom_ResumenMensualPorVendedor : System.Windows.Forms.Form
	{
		public Infragistics.Win.UltraWinGrid.UltraGrid uGridResumen;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public frmCom_ResumenMensualPorVendedor()
		{
			//
			// Necesario para admitir el Diseñador de Windows Forms
			//
			InitializeComponent();

			//
			// TODO: agregar código de constructor después de llamar a InitializeComponent
			//
		}

		int idPersonal = 0;

		public frmCom_ResumenMensualPorVendedor(int IdPersonal)
		{
			//
			// Necesario para admitir el Diseñador de Windows Forms
			//
			InitializeComponent();
			
			idPersonal = IdPersonal;

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
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			this.uGridResumen = new Infragistics.Win.UltraWinGrid.UltraGrid();
			((System.ComponentModel.ISupportInitialize)(this.uGridResumen)).BeginInit();
			this.SuspendLayout();
			// 
			// uGridResumen
			// 
			this.uGridResumen.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.uGridResumen.Cursor = System.Windows.Forms.Cursors.Default;
			appearance1.BackColor = System.Drawing.Color.White;
			this.uGridResumen.DisplayLayout.Appearance = appearance1;
			appearance2.ForeColor = System.Drawing.Color.Black;
			appearance2.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridResumen.DisplayLayout.Override.ActiveRowAppearance = appearance2;
			this.uGridResumen.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.True;
			appearance3.BackColor = System.Drawing.Color.Transparent;
			this.uGridResumen.DisplayLayout.Override.CardAreaAppearance = appearance3;
			this.uGridResumen.DisplayLayout.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.Edit;
			appearance4.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance4.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance4.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance4.FontData.BoldAsString = "True";
			appearance4.FontData.Name = "Arial";
			appearance4.FontData.SizeInPoints = 8F;
			appearance4.ForeColor = System.Drawing.Color.White;
			appearance4.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridResumen.DisplayLayout.Override.HeaderAppearance = appearance4;
			this.uGridResumen.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti;
			appearance5.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance5.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance5.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridResumen.DisplayLayout.Override.RowSelectorAppearance = appearance5;
			appearance6.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance6.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance6.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridResumen.DisplayLayout.Override.SelectedRowAppearance = appearance6;
			this.uGridResumen.DisplayLayout.Override.SelectTypeRow = Infragistics.Win.UltraWinGrid.SelectType.Extended;
			this.uGridResumen.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridResumen.Location = new System.Drawing.Point(8, 8);
			this.uGridResumen.Name = "uGridResumen";
			this.uGridResumen.Size = new System.Drawing.Size(736, 136);
			this.uGridResumen.TabIndex = 291;
			// 
			// frmCom_ResumenMensualPorVendedor
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(752, 150);
			this.Controls.Add(this.uGridResumen);
			this.KeyPreview = true;
			this.MaximizeBox = false;
			this.MinimizeBox = false;
			this.Name = "frmCom_ResumenMensualPorVendedor";
			this.ShowInTaskbar = false;
			this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
			this.Text = "Historial";
			this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.frmCom_ResumenMensualPorVendedor_KeyDown);
			this.Load += new System.EventHandler(this.frmCom_ResumenMensualPorVendedor_Load);
			((System.ComponentModel.ISupportInitialize)(this.uGridResumen)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void DiseñoGrid()
		{
			#region Diseño Grid
			if (this.uGridResumen.Rows.Count > 0)
			{				
				this.uGridResumen.DisplayLayout.Bands[0].Columns["Codigo"].Width = 100;
				this.uGridResumen.DisplayLayout.Bands[0].Columns["Codigo"].CellActivation = Activation.ActivateOnly;
				this.uGridResumen.DisplayLayout.Bands[0].Columns["Vendedor"].Width = 220;
				this.uGridResumen.DisplayLayout.Bands[0].Columns["Vendedor"].CellActivation = Activation.ActivateOnly;
				this.uGridResumen.DisplayLayout.Bands[0].Columns["Entrada"].Width = 75;				
				this.uGridResumen.DisplayLayout.Bands[0].Columns["Entrada"].CellActivation = Activation.ActivateOnly;
				this.uGridResumen.DisplayLayout.Bands[0].Columns["FechaTerminaPrueba"].Header.Caption = "Vence";
				this.uGridResumen.DisplayLayout.Bands[0].Columns["FechaTerminaPrueba"].Width = 75;				
				this.uGridResumen.DisplayLayout.Bands[0].Columns["FechaTerminaPrueba"].CellActivation = Activation.ActivateOnly;
				this.uGridResumen.DisplayLayout.Bands[0].Columns["Estado"].Width = 70;
				this.uGridResumen.DisplayLayout.Bands[0].Columns["Estado"].CellActivation = Activation.ActivateOnly;
				this.uGridResumen.DisplayLayout.Bands[0].Columns["Local"].Width = 120;
				this.uGridResumen.DisplayLayout.Bands[0].Columns["Local"].CellActivation = Activation.ActivateOnly;				
				this.uGridResumen.DisplayLayout.Bands[0].Columns["Cargo"].Width = 190;
				this.uGridResumen.DisplayLayout.Bands[0].Columns["Cargo"].CellActivation = Activation.ActivateOnly;
				this.uGridResumen.DisplayLayout.Bands[0].Columns["Sector"].Width = 200;
				this.uGridResumen.DisplayLayout.Bands[0].Columns["Sector"].CellActivation = Activation.ActivateOnly;				
				this.uGridResumen.DisplayLayout.Bands[0].Columns["Contador"].Hidden = true;
				
				if (this.uGridResumen.Rows.Count > 0)
				{
					string nombrCol = "";
							
					int iColumnas = (int)this.uGridResumen.DisplayLayout.Bands[0].Columns.Count;
				
					for (int j = 9; j < iColumnas; j++)
					{
						nombrCol = this.uGridResumen.DisplayLayout.Bands[0].Columns[j].Header.Caption;
						
						this.uGridResumen.DisplayLayout.Bands[0].Columns[nombrCol].CellActivation = Activation.ActivateOnly;
						this.uGridResumen.DisplayLayout.Bands[0].Columns[nombrCol].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center;
						this.uGridResumen.DisplayLayout.Bands[0].Columns[nombrCol].Width = 50;
					}
				}				
			}

			this.uGridResumen.DisplayLayout.Bands[0].Header.Caption = this.uGridResumen.Rows.Count + " ARTÍCULOS ENCONTRADOS.";
			#endregion Diseño Grid
		}

		private void frmCom_ResumenMensualPorVendedor_Load(object sender, System.EventArgs e)
		{
			string sSQL = string.Format("Exec Com_ResumenMensualPorVendedor {0}", idPersonal);
			FuncionesProcedimientos.dsGeneral(sSQL, this.uGridResumen);

			DiseñoGrid();
		}

		private void frmCom_ResumenMensualPorVendedor_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Escape) this.DialogResult = DialogResult.No;
		}
	}
}
