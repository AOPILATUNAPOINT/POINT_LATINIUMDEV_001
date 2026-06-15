using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using Infragistics.Win.UltraWinGrid;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de TipoCaso.
	/// </summary>
	public class TipoCaso : System.Windows.Forms.Form
	{
		private Infragistics.Win.UltraWinGrid.UltraGrid ultraGrid1;
		private Infragistics.Win.Misc.UltraButton btCancelar;
		private Infragistics.Win.Misc.UltraButton btAceptar;
		private Infragistics.Win.Misc.UltraButton ultraButton1;
		private Infragistics.Win.Misc.UltraButton ultraButton2;
		private C1.Data.C1DataSet cdsTipoCaso;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public TipoCaso()
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
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("P_TipoCaso", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idTipoCaso");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("TipoCaso");
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			this.ultraGrid1 = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.btCancelar = new Infragistics.Win.Misc.UltraButton();
			this.btAceptar = new Infragistics.Win.Misc.UltraButton();
			this.ultraButton1 = new Infragistics.Win.Misc.UltraButton();
			this.ultraButton2 = new Infragistics.Win.Misc.UltraButton();
			this.cdsTipoCaso = new C1.Data.C1DataSet();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsTipoCaso)).BeginInit();
			this.SuspendLayout();
			// 
			// ultraGrid1
			// 
			this.ultraGrid1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.ultraGrid1.Cursor = System.Windows.Forms.Cursors.Default;
			this.ultraGrid1.DataMember = "P_TipoCaso";
			this.ultraGrid1.DataSource = this.cdsTipoCaso;
			appearance1.BackColor = System.Drawing.Color.White;
			this.ultraGrid1.DisplayLayout.Appearance = appearance1;
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridBand1.Columns.AddRange(new object[] {
																										 ultraGridColumn1,
																										 ultraGridColumn2});
			this.ultraGrid1.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.ultraGrid1.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.TemplateOnBottom;
			this.ultraGrid1.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.True;
			this.ultraGrid1.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance2.BackColor = System.Drawing.Color.Transparent;
			this.ultraGrid1.DisplayLayout.Override.CardAreaAppearance = appearance2;
			appearance3.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance3.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance3.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance3.FontData.BoldAsString = "True";
			appearance3.FontData.Name = "Arial";
			appearance3.FontData.SizeInPoints = 10F;
			appearance3.ForeColor = System.Drawing.Color.White;
			appearance3.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.ultraGrid1.DisplayLayout.Override.HeaderAppearance = appearance3;
			this.ultraGrid1.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti;
			appearance4.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance4.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance4.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.RowSelectorAppearance = appearance4;
			appearance5.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance5.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance5.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.SelectedRowAppearance = appearance5;
			this.ultraGrid1.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.ultraGrid1.Location = new System.Drawing.Point(24, 16);
			this.ultraGrid1.Name = "ultraGrid1";
			this.ultraGrid1.Size = new System.Drawing.Size(224, 192);
			this.ultraGrid1.TabIndex = 1;
			this.ultraGrid1.KeyDown += new System.Windows.Forms.KeyEventHandler(this.ultraGrid1_KeyDown);
			// 
			// btCancelar
			// 
			this.btCancelar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.btCancelar.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.btCancelar.Location = new System.Drawing.Point(308, 224);
			this.btCancelar.Name = "btCancelar";
			this.btCancelar.Size = new System.Drawing.Size(40, 23);
			this.btCancelar.TabIndex = 4;
			this.btCancelar.Text = "&Cancelar";
			// 
			// btAceptar
			// 
			this.btAceptar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btAceptar.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.btAceptar.Location = new System.Drawing.Point(-40, 224);
			this.btAceptar.Name = "btAceptar";
			this.btAceptar.Size = new System.Drawing.Size(40, 23);
			this.btAceptar.TabIndex = 3;
			this.btAceptar.Text = "&Aceptar";
			// 
			// ultraButton1
			// 
			this.ultraButton1.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.ultraButton1.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.ultraButton1.Location = new System.Drawing.Point(112, 224);
			this.ultraButton1.Name = "ultraButton1";
			this.ultraButton1.TabIndex = 6;
			this.ultraButton1.Text = "&Cancelar";
			this.ultraButton1.Click += new System.EventHandler(this.ultraButton1_Click);
			this.ultraButton1.MouseUp += new System.Windows.Forms.MouseEventHandler(this.ultraButton1_MouseUp);
			// 
			// ultraButton2
			// 
			this.ultraButton2.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.ultraButton2.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.ultraButton2.Location = new System.Drawing.Point(24, 224);
			this.ultraButton2.Name = "ultraButton2";
			this.ultraButton2.TabIndex = 5;
			this.ultraButton2.Text = "&Aceptar";
			this.ultraButton2.Click += new System.EventHandler(this.ultraButton2_Click);
			// 
			// cdsTipoCaso
			// 
			this.cdsTipoCaso.BindingContextCtrl = this;
			this.cdsTipoCaso.DataLibrary = "LibFacturacion";
			this.cdsTipoCaso.DataLibraryUrl = "";
			this.cdsTipoCaso.DataSetDef = "dsTipoCaso";
			this.cdsTipoCaso.Locale = new System.Globalization.CultureInfo("es-MX");
			this.cdsTipoCaso.Name = "cdsTipoCaso";
			this.cdsTipoCaso.SchemaClassName = "LibFacturacion.DataClass";
			this.cdsTipoCaso.SchemaDef = null;
			this.cdsTipoCaso.BeforeFill += new C1.Data.FillEventHandler(this.cdsTipoCaso_BeforeFill);
			// 
			// TipoCaso
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(288, 273);
			this.Controls.Add(this.ultraButton1);
			this.Controls.Add(this.ultraButton2);
			this.Controls.Add(this.btCancelar);
			this.Controls.Add(this.btAceptar);
			this.Controls.Add(this.ultraGrid1);
			this.Name = "TipoCaso";
			this.Text = "TipoCaso";
			this.Load += new System.EventHandler(this.TipoCaso_Load);
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsTipoCaso)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void cdsTipoCaso_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			//			Funcion miFuncion = new Funcion();
			//			this.cdsTipoCaso.Schema.Connections[0].ConnectionString = 
			//				miFuncion.AccesoBase(MenuLatinium.stDirConta);
		}

		private void TipoCaso_Load(object sender, System.EventArgs e)
		{
		
		}

		private void ultraGrid1_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (ultraGrid1.ActiveCell == null) return;
			try
			{
				if (ultraGrid1.ActiveCell.Column.Key.ToString().ToUpper() == "NOTAS") return;
				switch(e.KeyValue)
				{
					case (int) Keys.PageDown:
						if(ultraGrid1.ActiveCell.DroppedDown == false)
							ultraGrid1.ActiveCell.DroppedDown = true;
						break;

					case (int) Keys.Up:
						if (ultraGrid1.ActiveCell.DroppedDown) return;

						ultraGrid1.PerformAction(UltraGridAction.ExitEditMode, false, false);
						ultraGrid1.PerformAction(UltraGridAction.AboveCell, false, false);
						e.Handled = true;
						ultraGrid1.PerformAction(UltraGridAction.EnterEditMode, false, false);
						break;
					case (int) Keys.Left:
						ultraGrid1.PerformAction(UltraGridAction.ExitEditMode, false, false);
						ultraGrid1.PerformAction(UltraGridAction.PrevCellByTab, false, false);
						e.Handled = true;
						ultraGrid1.PerformAction(UltraGridAction.EnterEditMode, false, false);
						break;
					case (int) Keys.Down:
						if (ultraGrid1.ActiveCell.DroppedDown) return;
			
						ultraGrid1.PerformAction(UltraGridAction.ExitEditMode, false, false);
						ultraGrid1.PerformAction(UltraGridAction.BelowCell, false, false);
						e.Handled = true;
						ultraGrid1.PerformAction(UltraGridAction.EnterEditMode, false, false);
						break;
					case (int) Keys.Enter:
						ultraGrid1.PerformAction(UltraGridAction.ExitEditMode, false, false);
						ultraGrid1.PerformAction(UltraGridAction.NextCellByTab, false, false);
						e.Handled = true;
						ultraGrid1.PerformAction(UltraGridAction.EnterEditMode, false, false);
						if(ultraGrid1.ActiveCell.DroppedDown == false)
							ultraGrid1.ActiveCell.DroppedDown = true;
						//						ultraGrid1.PerformAction(UltraGridAction.EnterEditModeAndDropdown, false, false);
						break;
				}
			}
			catch(System.Exception ex)
			{
				Console.WriteLine(ex.Message);
				//				MessageBox.Show(ex.Message, "Tecla");
			}

		}

		private void ultraButton1_MouseUp(object sender, System.Windows.Forms.MouseEventArgs e)
		{
			if (e.Button == MouseButtons.Right)
				Funcion.CreaFormulario(this);
		}

		private void ultraButton1_Click(object sender, System.EventArgs e)
		{
			this.Close();
		}

		private void ultraButton2_Click(object sender, System.EventArgs e)
		{
			try
			{
				this.cdsTipoCaso.Update();
			}
			catch(System.Exception ex)
			{
				MessageBox.Show(ex.Message);
			}
			this.Close();
		}
	}
}
