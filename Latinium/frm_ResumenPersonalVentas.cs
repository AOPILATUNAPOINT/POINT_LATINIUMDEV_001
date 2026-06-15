using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using Infragistics.Win.UltraWinGrid;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de frm_ResumenPersonalVentas.
	/// </summary>
	public class frm_ResumenPersonalVentas : System.Windows.Forms.Form
	{
		public Infragistics.Win.UltraWinGrid.UltraGrid uGridVentas;
		public Infragistics.Win.UltraWinGrid.UltraGrid uGridResultados;
		private System.Windows.Forms.Label lblPersonal;
		public Infragistics.Win.UltraWinGrid.UltraGrid uGridMetas;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public frm_ResumenPersonalVentas()
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
		string sVendedor = "";

		public frm_ResumenPersonalVentas(int IdPersonal, string SVendedor)
		{
			//
			// Necesario para admitir el Diseñador de Windows Forms
			//
			InitializeComponent();
			
			idPersonal = IdPersonal;
			sVendedor = SVendedor;
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
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance11 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance12 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance13 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance14 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance15 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance16 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance17 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance18 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance19 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance20 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance21 = new Infragistics.Win.Appearance();
			this.uGridVentas = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.uGridResultados = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.lblPersonal = new System.Windows.Forms.Label();
			this.uGridMetas = new Infragistics.Win.UltraWinGrid.UltraGrid();
			((System.ComponentModel.ISupportInitialize)(this.uGridVentas)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridResultados)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridMetas)).BeginInit();
			this.SuspendLayout();
			// 
			// uGridVentas
			// 
			this.uGridVentas.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.uGridVentas.Cursor = System.Windows.Forms.Cursors.Default;
			appearance1.BackColor = System.Drawing.Color.White;
			this.uGridVentas.DisplayLayout.Appearance = appearance1;
			appearance2.ForeColor = System.Drawing.Color.Black;
			appearance2.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridVentas.DisplayLayout.Override.ActiveRowAppearance = appearance2;
			this.uGridVentas.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
			this.uGridVentas.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.uGridVentas.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.False;
			appearance3.BackColor = System.Drawing.Color.Transparent;
			this.uGridVentas.DisplayLayout.Override.CardAreaAppearance = appearance3;
			this.uGridVentas.DisplayLayout.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.RowSelect;
			appearance4.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance4.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance4.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance4.FontData.BoldAsString = "True";
			appearance4.FontData.Name = "Arial";
			appearance4.FontData.SizeInPoints = 8.75F;
			appearance4.ForeColor = System.Drawing.Color.White;
			appearance4.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridVentas.DisplayLayout.Override.HeaderAppearance = appearance4;
			this.uGridVentas.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortSingle;
			appearance5.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance5.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance5.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridVentas.DisplayLayout.Override.RowAlternateAppearance = appearance5;
			appearance6.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance6.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance6.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridVentas.DisplayLayout.Override.RowSelectorAppearance = appearance6;
			appearance7.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance7.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance7.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridVentas.DisplayLayout.Override.SelectedRowAppearance = appearance7;
			this.uGridVentas.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridVentas.Location = new System.Drawing.Point(8, 120);
			this.uGridVentas.Name = "uGridVentas";
			this.uGridVentas.Size = new System.Drawing.Size(624, 64);
			this.uGridVentas.TabIndex = 677;
			// 
			// uGridResultados
			// 
			this.uGridResultados.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.uGridResultados.Cursor = System.Windows.Forms.Cursors.Default;
			appearance8.BackColor = System.Drawing.Color.White;
			this.uGridResultados.DisplayLayout.Appearance = appearance8;
			appearance9.ForeColor = System.Drawing.Color.Black;
			appearance9.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridResultados.DisplayLayout.Override.ActiveRowAppearance = appearance9;
			this.uGridResultados.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
			this.uGridResultados.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.uGridResultados.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.False;
			appearance10.BackColor = System.Drawing.Color.Transparent;
			this.uGridResultados.DisplayLayout.Override.CardAreaAppearance = appearance10;
			this.uGridResultados.DisplayLayout.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.RowSelect;
			appearance11.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance11.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance11.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance11.FontData.BoldAsString = "True";
			appearance11.FontData.Name = "Arial";
			appearance11.FontData.SizeInPoints = 8.75F;
			appearance11.ForeColor = System.Drawing.Color.White;
			appearance11.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridResultados.DisplayLayout.Override.HeaderAppearance = appearance11;
			this.uGridResultados.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortSingle;
			appearance12.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance12.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance12.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridResultados.DisplayLayout.Override.RowAlternateAppearance = appearance12;
			appearance13.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance13.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance13.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridResultados.DisplayLayout.Override.RowSelectorAppearance = appearance13;
			appearance14.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance14.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance14.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridResultados.DisplayLayout.Override.SelectedRowAppearance = appearance14;
			this.uGridResultados.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridResultados.Location = new System.Drawing.Point(8, 192);
			this.uGridResultados.Name = "uGridResultados";
			this.uGridResultados.Size = new System.Drawing.Size(624, 64);
			this.uGridResultados.TabIndex = 678;
			// 
			// lblPersonal
			// 
			this.lblPersonal.AutoSize = true;
			this.lblPersonal.Font = new System.Drawing.Font("Microsoft Sans Serif", 11F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Underline))), System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblPersonal.ForeColor = System.Drawing.Color.Firebrick;
			this.lblPersonal.Location = new System.Drawing.Point(8, 8);
			this.lblPersonal.Name = "lblPersonal";
			this.lblPersonal.Size = new System.Drawing.Size(0, 20);
			this.lblPersonal.TabIndex = 679;
			this.lblPersonal.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// uGridMetas
			// 
			this.uGridMetas.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.uGridMetas.Cursor = System.Windows.Forms.Cursors.Default;
			appearance15.BackColor = System.Drawing.Color.White;
			this.uGridMetas.DisplayLayout.Appearance = appearance15;
			appearance16.ForeColor = System.Drawing.Color.Black;
			appearance16.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridMetas.DisplayLayout.Override.ActiveRowAppearance = appearance16;
			this.uGridMetas.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
			this.uGridMetas.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.uGridMetas.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.False;
			appearance17.BackColor = System.Drawing.Color.Transparent;
			this.uGridMetas.DisplayLayout.Override.CardAreaAppearance = appearance17;
			this.uGridMetas.DisplayLayout.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.RowSelect;
			appearance18.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance18.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance18.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance18.FontData.BoldAsString = "True";
			appearance18.FontData.Name = "Arial";
			appearance18.FontData.SizeInPoints = 8.75F;
			appearance18.ForeColor = System.Drawing.Color.White;
			appearance18.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridMetas.DisplayLayout.Override.HeaderAppearance = appearance18;
			this.uGridMetas.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortSingle;
			appearance19.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance19.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance19.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridMetas.DisplayLayout.Override.RowAlternateAppearance = appearance19;
			appearance20.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance20.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance20.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridMetas.DisplayLayout.Override.RowSelectorAppearance = appearance20;
			appearance21.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance21.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance21.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridMetas.DisplayLayout.Override.SelectedRowAppearance = appearance21;
			this.uGridMetas.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridMetas.Location = new System.Drawing.Point(8, 48);
			this.uGridMetas.Name = "uGridMetas";
			this.uGridMetas.Size = new System.Drawing.Size(624, 64);
			this.uGridMetas.TabIndex = 680;
			// 
			// frm_ResumenPersonalVentas
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(640, 264);
			this.Controls.Add(this.uGridMetas);
			this.Controls.Add(this.lblPersonal);
			this.Controls.Add(this.uGridResultados);
			this.Controls.Add(this.uGridVentas);
			this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
			this.KeyPreview = true;
			this.MaximizeBox = false;
			this.MinimizeBox = false;
			this.Name = "frm_ResumenPersonalVentas";
			this.ShowInTaskbar = false;
			this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
			this.Text = "Resumen Rendimiento Personal de Ventas";
			this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.frm_ResumenPersonalVentas_KeyDown);
			this.Load += new System.EventHandler(this.frm_ResumenPersonalVentas_Load);
			((System.ComponentModel.ISupportInitialize)(this.uGridVentas)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridResultados)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridMetas)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void frm_ResumenPersonalVentas_Load(object sender, System.EventArgs e)
		{
			this.lblPersonal.Text = sVendedor;

			FuncionesProcedimientos.dsGeneralModulo(string.Format("Exec Com_ResumenRendimientoVendedores {0}, 1", idPersonal), this.uGridMetas);
			FuncionesProcedimientos.dsGeneralModulo(string.Format("Exec Com_ResumenRendimientoVendedores {0}, 2", idPersonal), this.uGridVentas);
			FuncionesProcedimientos.dsGeneralModulo(string.Format("Exec Com_ResumenRendimientoVendedores {0}, 3", idPersonal), this.uGridResultados);

			DiseñoGrid();
		}

		private void DiseñoGrid()
		{
			#region Diseño Grid
			if (this.uGridMetas.Rows.Count > 0)
			{
				string nombrCol = "";
							
				int iColumnas = (int)this.uGridMetas.DisplayLayout.Bands[0].Columns.Count;
				
				for (int j = 0; j < iColumnas; j++)
				{
					nombrCol = this.uGridMetas.DisplayLayout.Bands[0].Columns[j].Header.Caption;
												
					this.uGridMetas.DisplayLayout.Bands[0].Columns[nombrCol].CellActivation = Activation.ActivateOnly;
					this.uGridMetas.DisplayLayout.Bands[0].Columns[nombrCol].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right;
					this.uGridMetas.DisplayLayout.Bands[0].Columns[nombrCol].Width = 85;						
				}					
			}

			if (this.uGridVentas.Rows.Count > 0)
			{
				this.uGridVentas.DisplayLayout.Bands[0].Columns["Total"].Width = 85;
				this.uGridVentas.DisplayLayout.Bands[0].Columns["Total"].CellActivation = Activation.NoEdit;
				this.uGridVentas.DisplayLayout.Bands[0].Columns["Total"].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right;
				this.uGridVentas.DisplayLayout.Bands[0].Columns["Total"].CellAppearance.BackColor = Color.Silver;
				
				this.uGridVentas.DisplayLayout.Bands[0].Columns["Promedio"].Width = 85;
				this.uGridVentas.DisplayLayout.Bands[0].Columns["Promedio"].CellActivation = Activation.NoEdit;
				this.uGridVentas.DisplayLayout.Bands[0].Columns["Promedio"].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right;
				this.uGridVentas.DisplayLayout.Bands[0].Columns["Promedio"].CellAppearance.BackColor = Color.Silver;
				
				if (this.uGridVentas.Rows.Count > 0)
				{
					string nombrCol = "";
							
					int iColumnas = (int)this.uGridVentas.DisplayLayout.Bands[0].Columns.Count;
				
					for (int j = 2; j < iColumnas; j++)
					{
						nombrCol = this.uGridVentas.DisplayLayout.Bands[0].Columns[j].Header.Caption;
												
						this.uGridVentas.DisplayLayout.Bands[0].Columns[nombrCol].CellActivation = Activation.ActivateOnly;
						this.uGridVentas.DisplayLayout.Bands[0].Columns[nombrCol].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right;
						this.uGridVentas.DisplayLayout.Bands[0].Columns[nombrCol].Width = 85;						
					}					
				}
			}

			if (this.uGridResultados.Rows.Count > 0)
			{
				string nombrCol = "";
							
				int iColumnas = (int)this.uGridResultados.DisplayLayout.Bands[0].Columns.Count;
				
				for (int j = 0; j < iColumnas; j++)
				{
					nombrCol = this.uGridResultados.DisplayLayout.Bands[0].Columns[j].Header.Caption;
												
					this.uGridResultados.DisplayLayout.Bands[0].Columns[nombrCol].CellActivation = Activation.ActivateOnly;
					this.uGridResultados.DisplayLayout.Bands[0].Columns[nombrCol].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right;
					this.uGridResultados.DisplayLayout.Bands[0].Columns[nombrCol].Width = 85;						
				}					
			}			
			#endregion Diseño Grid
		}

		private void frm_ResumenPersonalVentas_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Escape) this.Close();
		}
	}
}
