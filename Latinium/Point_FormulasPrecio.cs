using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using DevExpress.XtraEditors;

namespace Latinium
{
	/// <summary>
	/// Summary description for Point_FormulasPrecio.
	/// </summary>
	public class Point_FormulasPrecio : DevExpress.XtraEditors.XtraForm
	{
		private Infragistics.Win.UltraWinGrid.UltraGrid grdFormulaPrecio;
		private C1.Data.C1DataSet cdsFormulaPrecio;
		private Infragistics.Win.Misc.UltraButton btCancelar;
		private Infragistics.Win.Misc.UltraButton btAceptar;
		/// <summary>
		/// Required designer variable.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public Point_FormulasPrecio()
		{
			//
			// Required for Windows Form Designer support
			//
			InitializeComponent();

			//
			// TODO: Add any constructor code after InitializeComponent call
			//
		}

		/// <summary>
		/// Clean up any resources being used.
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

		#region Windows Form Designer generated code
		/// <summary>
		/// Required method for Designer support - do not modify
		/// the contents of this method with the code editor.
		/// </summary>
		private void InitializeComponent()
		{
			Infragistics.Win.Appearance appearance1 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Point_Formula", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idFormula");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Formula");
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			this.grdFormulaPrecio = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.cdsFormulaPrecio = new C1.Data.C1DataSet();
			this.btCancelar = new Infragistics.Win.Misc.UltraButton();
			this.btAceptar = new Infragistics.Win.Misc.UltraButton();
			((System.ComponentModel.ISupportInitialize)(this.grdFormulaPrecio)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsFormulaPrecio)).BeginInit();
			this.SuspendLayout();
			// 
			// grdFormulaPrecio
			// 
			this.grdFormulaPrecio.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.grdFormulaPrecio.Cursor = System.Windows.Forms.Cursors.Default;
			this.grdFormulaPrecio.DataBindings.Add(new System.Windows.Forms.Binding("Text", this.cdsFormulaPrecio, "Point_Formula.idFormula"));
			this.grdFormulaPrecio.DataSource = this.cdsFormulaPrecio;
			appearance1.BackColor = System.Drawing.Color.White;
			this.grdFormulaPrecio.DisplayLayout.Appearance = appearance1;
			this.grdFormulaPrecio.DisplayLayout.AutoFitColumns = true;
			ultraGridBand1.AddButtonCaption = "FacturaNumero";
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn1.Width = 187;
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridColumn2.Width = 280;
			ultraGridBand1.Columns.AddRange(new object[] {
																										 ultraGridColumn1,
																										 ultraGridColumn2});
			this.grdFormulaPrecio.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.grdFormulaPrecio.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.TemplateOnBottom;
			this.grdFormulaPrecio.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.True;
			this.grdFormulaPrecio.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance2.BackColor = System.Drawing.Color.Transparent;
			this.grdFormulaPrecio.DisplayLayout.Override.CardAreaAppearance = appearance2;
			appearance3.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance3.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance3.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance3.FontData.BoldAsString = "True";
			appearance3.FontData.Name = "Arial";
			appearance3.FontData.SizeInPoints = 10F;
			appearance3.ForeColor = System.Drawing.Color.White;
			appearance3.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.grdFormulaPrecio.DisplayLayout.Override.HeaderAppearance = appearance3;
			appearance4.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance4.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance4.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.grdFormulaPrecio.DisplayLayout.Override.RowSelectorAppearance = appearance4;
			appearance5.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance5.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance5.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.grdFormulaPrecio.DisplayLayout.Override.SelectedRowAppearance = appearance5;
			this.grdFormulaPrecio.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.grdFormulaPrecio.Location = new System.Drawing.Point(16, 8);
			this.grdFormulaPrecio.Name = "grdFormulaPrecio";
			this.grdFormulaPrecio.Size = new System.Drawing.Size(488, 288);
			this.grdFormulaPrecio.TabIndex = 1;
			// 
			// cdsFormulaPrecio
			// 
			this.cdsFormulaPrecio.BindingContextCtrl = this;
			this.cdsFormulaPrecio.DataLibrary = "LibFacturacion";
			this.cdsFormulaPrecio.DataLibraryUrl = "";
			this.cdsFormulaPrecio.DataSetDef = "dsFormulaPrecio";
			this.cdsFormulaPrecio.EnforceConstraints = false;
			this.cdsFormulaPrecio.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsFormulaPrecio.Name = "cdsFormulaPrecio";
			this.cdsFormulaPrecio.SchemaClassName = "LibFacturacion.DataClass";
			this.cdsFormulaPrecio.SchemaDef = null;
			// 
			// btCancelar
			// 
			this.btCancelar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.btCancelar.DialogResult = System.Windows.Forms.DialogResult.Cancel;
			this.btCancelar.Location = new System.Drawing.Point(320, 320);
			this.btCancelar.Name = "btCancelar";
			this.btCancelar.Size = new System.Drawing.Size(104, 25);
			this.btCancelar.TabIndex = 95;
			this.btCancelar.Text = "&Cancelar";
			this.btCancelar.Click += new System.EventHandler(this.btCancelar_Click);
			// 
			// btAceptar
			// 
			this.btAceptar.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.btAceptar.Location = new System.Drawing.Point(88, 320);
			this.btAceptar.Name = "btAceptar";
			this.btAceptar.Size = new System.Drawing.Size(75, 25);
			this.btAceptar.TabIndex = 96;
			this.btAceptar.Text = "&Aceptar";
			this.btAceptar.Click += new System.EventHandler(this.btAceptar_Click);
			// 
			// Point_FormulasPrecio
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 14);
			this.ClientSize = new System.Drawing.Size(528, 366);
			this.Controls.Add(this.btAceptar);
			this.Controls.Add(this.btCancelar);
			this.Controls.Add(this.grdFormulaPrecio);
			this.Name = "Point_FormulasPrecio";
			this.Text = "Point_FormulasPrecio";
			((System.ComponentModel.ISupportInitialize)(this.grdFormulaPrecio)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsFormulaPrecio)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void btCancelar_Click(object sender, System.EventArgs e)
		{
			Close();
		}

		private void btAceptar_Click(object sender, System.EventArgs e)
		{
			this.cdsFormulaPrecio.Update();

			// Actualiza Clave Principal
//			string stSelect = "EXEC InicioIDP " + this.spnEmpresa.Value.ToString();
//			Funcion.EjecutaSQL(cdsSeteo, stSelect);

			this.Close();
		}
	}
}

