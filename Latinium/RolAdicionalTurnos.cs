using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de RolAdicionalTurnos.
	/// </summary>
	public class RolAdicionalTurnos : System.Windows.Forms.Form
	{
		private C1.Data.C1DataSet cdsRolAdicional;
		private Infragistics.Win.UltraWinGrid.UltraGrid grdRetRenta;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public RolAdicionalTurnos()
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
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("_RolAdicionalesDetalle", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idRolAdicionalDetalle");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idRolAdicionales");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPersonal");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Valor1");
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Valor2");
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Valor3");
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("HoraEntrada");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("HoraSalida");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idHorario");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Libre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn11 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CL");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn12 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Diferencia");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn13 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Valor1Ext");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn14 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Valor2Ext");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn15 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Valor3Ext");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn16 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Observaciones");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn17 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("HorasTotales");
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			this.cdsRolAdicional = new C1.Data.C1DataSet();
			this.grdRetRenta = new Infragistics.Win.UltraWinGrid.UltraGrid();
			((System.ComponentModel.ISupportInitialize)(this.cdsRolAdicional)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.grdRetRenta)).BeginInit();
			this.SuspendLayout();
			// 
			// cdsRolAdicional
			// 
			this.cdsRolAdicional.BindingContextCtrl = this;
			this.cdsRolAdicional.DataLibrary = "LibPersonal";
			this.cdsRolAdicional.DataLibraryUrl = "";
			this.cdsRolAdicional.DataSetDef = "dsRolAdicional";
			this.cdsRolAdicional.EnforceConstraints = false;
			this.cdsRolAdicional.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsRolAdicional.Name = "cdsRolAdicional";
			this.cdsRolAdicional.SchemaClassName = "LibPersonal.DataClass";
			this.cdsRolAdicional.SchemaDef = null;
			// 
			// grdRetRenta
			// 
			this.grdRetRenta.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.grdRetRenta.Cursor = System.Windows.Forms.Cursors.Default;
			this.grdRetRenta.DataMember = "_RolAdicionales.RolAdicionales - RolAdicionalesDetalle";
			this.grdRetRenta.DataSource = this.cdsRolAdicional;
			appearance1.BackColor = System.Drawing.Color.White;
			this.grdRetRenta.DisplayLayout.Appearance = appearance1;
			this.grdRetRenta.DisplayLayout.AutoFitColumns = true;
			ultraGridBand1.AddButtonCaption = "RetDetalle";
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn1.Width = 94;
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridColumn2.Hidden = true;
			ultraGridColumn2.Width = 75;
			ultraGridColumn3.Header.Caption = "Personal";
			ultraGridColumn3.Header.VisiblePosition = 2;
			ultraGridColumn3.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownList;
			ultraGridColumn3.Width = 151;
			appearance2.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn4.CellAppearance = appearance2;
			ultraGridColumn4.Format = "n2";
			ultraGridColumn4.Header.VisiblePosition = 5;
			ultraGridColumn4.Hidden = true;
			ultraGridColumn4.Width = 51;
			appearance3.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn5.CellAppearance = appearance3;
			ultraGridColumn5.Format = "n2";
			ultraGridColumn5.Header.VisiblePosition = 6;
			ultraGridColumn5.Hidden = true;
			ultraGridColumn5.Width = 46;
			appearance4.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn6.CellAppearance = appearance4;
			ultraGridColumn6.Format = "n2";
			ultraGridColumn6.Header.VisiblePosition = 7;
			ultraGridColumn6.Hidden = true;
			ultraGridColumn6.Width = 53;
			ultraGridColumn7.Format = "HH:mm";
			ultraGridColumn7.Header.VisiblePosition = 3;
			ultraGridColumn7.VertScrollBar = true;
			ultraGridColumn7.Width = 63;
			ultraGridColumn8.Format = "HH:mm";
			ultraGridColumn8.Header.VisiblePosition = 4;
			ultraGridColumn8.VertScrollBar = true;
			ultraGridColumn8.Width = 76;
			ultraGridColumn9.Header.VisiblePosition = 8;
			ultraGridColumn9.Hidden = true;
			ultraGridColumn9.Width = 42;
			ultraGridColumn10.Header.Caption = "Doblar";
			ultraGridColumn10.Header.VisiblePosition = 9;
			ultraGridColumn10.Width = 34;
			ultraGridColumn11.Header.Caption = "CLibre";
			ultraGridColumn11.Header.VisiblePosition = 10;
			ultraGridColumn11.Width = 34;
			ultraGridColumn12.Header.VisiblePosition = 11;
			ultraGridColumn12.Hidden = true;
			ultraGridColumn12.Width = 40;
			ultraGridColumn13.Header.VisiblePosition = 13;
			ultraGridColumn13.Width = 53;
			ultraGridColumn14.Header.VisiblePosition = 14;
			ultraGridColumn14.Width = 53;
			ultraGridColumn15.Header.VisiblePosition = 15;
			ultraGridColumn15.Width = 59;
			ultraGridColumn16.Header.VisiblePosition = 16;
			ultraGridColumn16.Width = 92;
			ultraGridColumn17.Header.VisiblePosition = 12;
			ultraGridColumn17.Width = 100;
			ultraGridBand1.Columns.AddRange(new object[] {
																										 ultraGridColumn1,
																										 ultraGridColumn2,
																										 ultraGridColumn3,
																										 ultraGridColumn4,
																										 ultraGridColumn5,
																										 ultraGridColumn6,
																										 ultraGridColumn7,
																										 ultraGridColumn8,
																										 ultraGridColumn9,
																										 ultraGridColumn10,
																										 ultraGridColumn11,
																										 ultraGridColumn12,
																										 ultraGridColumn13,
																										 ultraGridColumn14,
																										 ultraGridColumn15,
																										 ultraGridColumn16,
																										 ultraGridColumn17});
			this.grdRetRenta.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.grdRetRenta.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.TemplateOnBottom;
			this.grdRetRenta.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.True;
			this.grdRetRenta.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance5.BackColor = System.Drawing.Color.Transparent;
			this.grdRetRenta.DisplayLayout.Override.CardAreaAppearance = appearance5;
			appearance6.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance6.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance6.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance6.FontData.BoldAsString = "True";
			appearance6.FontData.Name = "Arial";
			appearance6.FontData.SizeInPoints = 10F;
			appearance6.ForeColor = System.Drawing.Color.White;
			appearance6.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.grdRetRenta.DisplayLayout.Override.HeaderAppearance = appearance6;
			appearance7.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance7.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance7.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.grdRetRenta.DisplayLayout.Override.RowSelectorAppearance = appearance7;
			appearance8.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance8.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance8.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.grdRetRenta.DisplayLayout.Override.SelectedRowAppearance = appearance8;
			this.grdRetRenta.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.grdRetRenta.Location = new System.Drawing.Point(24, 14);
			this.grdRetRenta.Name = "grdRetRenta";
			this.grdRetRenta.Size = new System.Drawing.Size(736, 288);
			this.grdRetRenta.TabIndex = 28;
			// 
			// RolAdicionalTurnos
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(784, 317);
			this.Controls.Add(this.grdRetRenta);
			this.Name = "RolAdicionalTurnos";
			this.Text = "RolAdicionalTurnos";
			((System.ComponentModel.ISupportInitialize)(this.cdsRolAdicional)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.grdRetRenta)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion
	}
}
