using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de P_TablaComisiones.
	/// </summary>
	public class P_TablaComisiones : DevExpress.XtraEditors.XtraForm
	{
		private Infragistics.Win.Misc.UltraButton btCancelar;
		private Infragistics.Win.Misc.UltraButton btAceptar;
		private C1.Data.C1DataSet cdsTabla;
		private System.Windows.Forms.Label label1;
		private System.Windows.Forms.Label label2;
		private Infragistics.Win.UltraWinGrid.UltraGrid grTablaComision;
		private Infragistics.Win.UltraWinGrid.UltraGrid grTablaComisionP;
		private C1.Data.C1DataSet cdsTabla1;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public P_TablaComisiones()
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
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("P_TablaComision", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idTablaComision");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Porcentaje");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("dias");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Penalidad");
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("P_TablaComision", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idTablaComision");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Porcentaje");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("dias");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Penalidad");
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
			this.btCancelar = new Infragistics.Win.Misc.UltraButton();
			this.btAceptar = new Infragistics.Win.Misc.UltraButton();
			this.cdsTabla = new C1.Data.C1DataSet();
			this.label1 = new System.Windows.Forms.Label();
			this.label2 = new System.Windows.Forms.Label();
			this.grTablaComision = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.grTablaComisionP = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.cdsTabla1 = new C1.Data.C1DataSet();
			((System.ComponentModel.ISupportInitialize)(this.cdsTabla)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.grTablaComision)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.grTablaComisionP)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsTabla1)).BeginInit();
			this.SuspendLayout();
			// 
			// btCancelar
			// 
			this.btCancelar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.btCancelar.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.btCancelar.Location = new System.Drawing.Point(464, 292);
			this.btCancelar.Name = "btCancelar";
			this.btCancelar.TabIndex = 14;
			this.btCancelar.Text = "&Cancelar";
			this.btCancelar.Click += new System.EventHandler(this.btCancelar_Click);
			// 
			// btAceptar
			// 
			this.btAceptar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btAceptar.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.btAceptar.Location = new System.Drawing.Point(32, 292);
			this.btAceptar.Name = "btAceptar";
			this.btAceptar.TabIndex = 13;
			this.btAceptar.Text = "&Aceptar";
			this.btAceptar.Click += new System.EventHandler(this.btAceptar_Click);
			// 
			// cdsTabla
			// 
			this.cdsTabla.BindingContextCtrl = this;
			this.cdsTabla.DataLibrary = "LibFacturacion";
			this.cdsTabla.DataLibraryUrl = "";
			this.cdsTabla.DataSetDef = "dsTablaComision";
			this.cdsTabla.Locale = new System.Globalization.CultureInfo("es-MX");
			this.cdsTabla.Name = "cdsTabla";
			this.cdsTabla.SchemaClassName = "LibFacturacion.DataClass";
			this.cdsTabla.SchemaDef = null;
			this.cdsTabla.BeforeFill += new C1.Data.FillEventHandler(this.cdsTabla_BeforeFill);
			// 
			// label1
			// 
			this.label1.Location = new System.Drawing.Point(72, 8);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(200, 16);
			this.label1.TabIndex = 16;
			this.label1.Text = "CALCULO COMISIONES";
			// 
			// label2
			// 
			this.label2.Location = new System.Drawing.Point(384, 8);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(72, 16);
			this.label2.TabIndex = 17;
			this.label2.Text = "PENALIDAD";
			// 
			// grTablaComision
			// 
			this.grTablaComision.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left)));
			this.grTablaComision.Cursor = System.Windows.Forms.Cursors.Default;
			this.grTablaComision.DataMember = "P_TablaComision";
			this.grTablaComision.DataSource = this.cdsTabla;
			appearance1.BackColor = System.Drawing.Color.White;
			this.grTablaComision.DisplayLayout.Appearance = appearance1;
			this.grTablaComision.DisplayLayout.AutoFitColumns = true;
			ultraGridBand1.AddButtonCaption = "SubProyecto";
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn1.Width = 74;
			ultraGridColumn2.Format = "0.00";
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridColumn2.Width = 136;
			ultraGridColumn3.Header.VisiblePosition = 2;
			ultraGridColumn3.Width = 99;
			ultraGridColumn4.Header.VisiblePosition = 3;
			ultraGridColumn4.Hidden = true;
			ultraGridColumn4.Width = 172;
			ultraGridBand1.Columns.AddRange(new object[] {
																										 ultraGridColumn1,
																										 ultraGridColumn2,
																										 ultraGridColumn3,
																										 ultraGridColumn4});
			this.grTablaComision.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.grTablaComision.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.TemplateOnBottom;
			this.grTablaComision.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.True;
			this.grTablaComision.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance2.BackColor = System.Drawing.Color.Transparent;
			this.grTablaComision.DisplayLayout.Override.CardAreaAppearance = appearance2;
			appearance3.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance3.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance3.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance3.FontData.BoldAsString = "True";
			appearance3.FontData.Name = "Arial";
			appearance3.FontData.SizeInPoints = 10F;
			appearance3.ForeColor = System.Drawing.Color.White;
			appearance3.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.grTablaComision.DisplayLayout.Override.HeaderAppearance = appearance3;
			this.grTablaComision.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti;
			appearance4.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance4.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance4.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.grTablaComision.DisplayLayout.Override.RowSelectorAppearance = appearance4;
			appearance5.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance5.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance5.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.grTablaComision.DisplayLayout.Override.SelectedRowAppearance = appearance5;
			this.grTablaComision.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.grTablaComision.Location = new System.Drawing.Point(8, 32);
			this.grTablaComision.Name = "grTablaComision";
			this.grTablaComision.Size = new System.Drawing.Size(256, 248);
			this.grTablaComision.TabIndex = 18;
			this.grTablaComision.InitializeRow += new Infragistics.Win.UltraWinGrid.InitializeRowEventHandler(this.grTablaComision_InitializeRow);
			// 
			// grTablaComisionP
			// 
			this.grTablaComisionP.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left)));
			this.grTablaComisionP.Cursor = System.Windows.Forms.Cursors.Default;
			this.grTablaComisionP.DataMember = "P_TablaComision";
			this.grTablaComisionP.DataSource = this.cdsTabla1;
			appearance6.BackColor = System.Drawing.Color.White;
			this.grTablaComisionP.DisplayLayout.Appearance = appearance6;
			this.grTablaComisionP.DisplayLayout.AutoFitColumns = true;
			ultraGridBand2.AddButtonCaption = "SubProyecto";
			ultraGridColumn5.Header.VisiblePosition = 0;
			ultraGridColumn5.Hidden = true;
			ultraGridColumn5.Width = 74;
			ultraGridColumn6.Format = "0.00";
			ultraGridColumn6.Header.VisiblePosition = 1;
			ultraGridColumn6.Width = 133;
			ultraGridColumn7.Header.VisiblePosition = 2;
			ultraGridColumn7.Width = 94;
			ultraGridColumn8.Header.VisiblePosition = 3;
			ultraGridColumn8.Hidden = true;
			ultraGridColumn8.Width = 75;
			ultraGridBand2.Columns.AddRange(new object[] {
																										 ultraGridColumn5,
																										 ultraGridColumn6,
																										 ultraGridColumn7,
																										 ultraGridColumn8});
			this.grTablaComisionP.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			this.grTablaComisionP.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.TemplateOnBottom;
			this.grTablaComisionP.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.True;
			this.grTablaComisionP.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance7.BackColor = System.Drawing.Color.Transparent;
			this.grTablaComisionP.DisplayLayout.Override.CardAreaAppearance = appearance7;
			appearance8.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance8.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance8.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance8.FontData.BoldAsString = "True";
			appearance8.FontData.Name = "Arial";
			appearance8.FontData.SizeInPoints = 10F;
			appearance8.ForeColor = System.Drawing.Color.White;
			appearance8.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.grTablaComisionP.DisplayLayout.Override.HeaderAppearance = appearance8;
			this.grTablaComisionP.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti;
			appearance9.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance9.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance9.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.grTablaComisionP.DisplayLayout.Override.RowSelectorAppearance = appearance9;
			appearance10.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance10.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance10.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.grTablaComisionP.DisplayLayout.Override.SelectedRowAppearance = appearance10;
			this.grTablaComisionP.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.grTablaComisionP.Location = new System.Drawing.Point(296, 32);
			this.grTablaComisionP.Name = "grTablaComisionP";
			this.grTablaComisionP.Size = new System.Drawing.Size(248, 248);
			this.grTablaComisionP.TabIndex = 19;
			this.grTablaComisionP.InitializeRow += new Infragistics.Win.UltraWinGrid.InitializeRowEventHandler(this.grTablaComisionP_InitializeRow);
			// 
			// cdsTabla1
			// 
			this.cdsTabla1.BindingContextCtrl = this;
			this.cdsTabla1.DataLibrary = "LibFacturacion";
			this.cdsTabla1.DataLibraryUrl = "";
			this.cdsTabla1.DataSetDef = "dsTablaComision";
			this.cdsTabla1.Locale = new System.Globalization.CultureInfo("es-MX");
			this.cdsTabla1.Name = "cdsTabla1";
			this.cdsTabla1.SchemaClassName = "LibFacturacion.DataClass";
			this.cdsTabla1.SchemaDef = null;
			this.cdsTabla1.BeforeFill += new C1.Data.FillEventHandler(this.cdsTabla1_BeforeFill);
			// 
			// P_TablaComisiones
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(560, 325);
			this.Controls.Add(this.grTablaComisionP);
			this.Controls.Add(this.grTablaComision);
			this.Controls.Add(this.label2);
			this.Controls.Add(this.label1);
			this.Controls.Add(this.btCancelar);
			this.Controls.Add(this.btAceptar);
			this.Name = "P_TablaComisiones";
			this.Text = "Tabla de Comisiones";
			this.Load += new System.EventHandler(this.P_TablaComisiones_Load);
			((System.ComponentModel.ISupportInitialize)(this.cdsTabla)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.grTablaComision)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.grTablaComisionP)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsTabla1)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void btCancelar_Click(object sender, System.EventArgs e)
		{
			Close();
		}

		private void btAceptar_Click(object sender, System.EventArgs e)
		{				
			try
			{
				cdsTabla.Update();
				cdsTabla1.Update();
				
			}
			catch(Exception ex)
			{
				MessageBox.Show("No se pudo Grabar...."+ex.Message);
			}

		}	

		private void cdsTabla_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsTabla.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);
		}

		private void P_TablaComisiones_Load(object sender, System.EventArgs e)
		{
			C1.Data.FilterCondition fcPadre = new 
				C1.Data.FilterCondition(this.cdsTabla, "P_TablaComision", "Penalidad = 0");
			C1.Data.FilterConditions fcTotal = new C1.Data.FilterConditions();
			fcTotal.Add(fcPadre);
			this.cdsTabla.Fill(fcTotal, false);

			C1.Data.FilterCondition fcPadre1 = new 
				C1.Data.FilterCondition(this.cdsTabla1, "P_TablaComision", "Penalidad = 1");
			C1.Data.FilterConditions fcTotal1 = new C1.Data.FilterConditions();
			fcTotal1.Add(fcPadre1);
			this.cdsTabla1.Fill(fcTotal1, false);

		}


		private void cdsTabla1_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsTabla1.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);
		}

		private void grTablaComision_InitializeRow(object sender, Infragistics.Win.UltraWinGrid.InitializeRowEventArgs e)
		{
			e.Row.Cells["Penalidad"].Value = 0;		
		}

		private void grTablaComisionP_InitializeRow(object sender, Infragistics.Win.UltraWinGrid.InitializeRowEventArgs e)
		{
				e.Row.Cells["Penalidad"].Value = 1;
		}
	}
}
