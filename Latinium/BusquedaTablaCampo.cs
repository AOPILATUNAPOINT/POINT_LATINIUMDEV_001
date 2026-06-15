using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de BusquedaTablaCampo.
	/// </summary>
	public class BusquedaTablaCampo : DevExpress.XtraEditors.XtraForm
	{
		private Infragistics.Win.UltraWinGrid.UltraGrid ultraGrid1;
		private C1.Data.C1DataSet cdsTablaInstruccion;
		private Infragistics.Win.Misc.UltraButton ultraButton1;
		private Infragistics.Win.Misc.UltraButton ultraButton2;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource udsTipoCampo;
		private Infragistics.Win.UltraWinGrid.UltraDropDown cmbTipo;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public BusquedaTablaCampo()
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
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("TablaCampo", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idTablaCampo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Tabla");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Campo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("TipoCampo", -1, "cmbTipo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Busqueda");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Pantalla");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPrincipal");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Visible");
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idTipoCampo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("TipoCampo");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idTipoCampo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("TipoCampo");
			this.ultraGrid1 = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.cdsTablaInstruccion = new C1.Data.C1DataSet();
			this.ultraButton1 = new Infragistics.Win.Misc.UltraButton();
			this.ultraButton2 = new Infragistics.Win.Misc.UltraButton();
			this.udsTipoCampo = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.cmbTipo = new Infragistics.Win.UltraWinGrid.UltraDropDown();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsTablaInstruccion)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.udsTipoCampo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbTipo)).BeginInit();
			this.SuspendLayout();
			// 
			// ultraGrid1
			// 
			this.ultraGrid1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.ultraGrid1.Cursor = System.Windows.Forms.Cursors.Default;
			this.ultraGrid1.DataMember = "TablaCampo";
			this.ultraGrid1.DataSource = this.cdsTablaInstruccion;
			appearance1.BackColor = System.Drawing.Color.White;
			this.ultraGrid1.DisplayLayout.Appearance = appearance1;
			this.ultraGrid1.DisplayLayout.AutoFitColumns = true;
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridColumn2.Width = 93;
			ultraGridColumn3.Header.VisiblePosition = 2;
			ultraGridColumn3.Width = 91;
			ultraGridColumn4.Header.VisiblePosition = 3;
			ultraGridColumn4.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownValidate;
			ultraGridColumn4.Width = 69;
			ultraGridColumn5.Header.VisiblePosition = 4;
			ultraGridColumn5.Width = 58;
			ultraGridColumn6.Header.VisiblePosition = 5;
			ultraGridColumn6.Width = 50;
			ultraGridColumn7.Header.VisiblePosition = 6;
			ultraGridColumn7.Width = 66;
			ultraGridColumn8.Header.VisiblePosition = 7;
			ultraGridColumn8.Width = 48;
			ultraGridBand1.Columns.AddRange(new object[] {
																										 ultraGridColumn1,
																										 ultraGridColumn2,
																										 ultraGridColumn3,
																										 ultraGridColumn4,
																										 ultraGridColumn5,
																										 ultraGridColumn6,
																										 ultraGridColumn7,
																										 ultraGridColumn8});
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
			appearance4.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance4.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance4.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.RowSelectorAppearance = appearance4;
			appearance5.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance5.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance5.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.SelectedRowAppearance = appearance5;
			this.ultraGrid1.DisplayLayout.ViewStyleBand = Infragistics.Win.UltraWinGrid.ViewStyleBand.OutlookGroupBy;
			this.ultraGrid1.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.ultraGrid1.Location = new System.Drawing.Point(24, 16);
			this.ultraGrid1.Name = "ultraGrid1";
			this.ultraGrid1.Size = new System.Drawing.Size(496, 208);
			this.ultraGrid1.TabIndex = 0;
			// 
			// cdsTablaInstruccion
			// 
			this.cdsTablaInstruccion.BindingContextCtrl = this;
			this.cdsTablaInstruccion.DataLibrary = "LibFacturacion";
			this.cdsTablaInstruccion.DataLibraryUrl = "";
			this.cdsTablaInstruccion.DataSetDef = "dsTablaBuscar";
			this.cdsTablaInstruccion.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsTablaInstruccion.Name = "cdsTablaInstruccion";
			this.cdsTablaInstruccion.SchemaClassName = "LibFacturacion.DataClass";
			this.cdsTablaInstruccion.SchemaDef = null;
			// 
			// ultraButton1
			// 
			this.ultraButton1.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.ultraButton1.Location = new System.Drawing.Point(48, 232);
			this.ultraButton1.Name = "ultraButton1";
			this.ultraButton1.Size = new System.Drawing.Size(72, 24);
			this.ultraButton1.TabIndex = 1;
			this.ultraButton1.Text = "&Grabar";
			this.ultraButton1.Click += new System.EventHandler(this.ultraButton1_Click);
			// 
			// ultraButton2
			// 
			this.ultraButton2.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.ultraButton2.Location = new System.Drawing.Point(408, 232);
			this.ultraButton2.Name = "ultraButton2";
			this.ultraButton2.Size = new System.Drawing.Size(72, 24);
			this.ultraButton2.TabIndex = 2;
			this.ultraButton2.Text = "&Salir";
			this.ultraButton2.Click += new System.EventHandler(this.ultraButton2_Click);
			// 
			// udsTipoCampo
			// 
			this.udsTipoCampo.Band.Columns.AddRange(new object[] {
																														 ultraDataColumn1,
																														 ultraDataColumn2});
			this.udsTipoCampo.Rows.AddRange(new object[] {
																										 new Infragistics.Win.UltraWinDataSource.UltraDataRow(new object[] {
																																																												 "idTipoCampo",
																																																												 "1",
																																																												 "TipoCampo",
																																																												 "Bit"}, new object[0]),
																										 new Infragistics.Win.UltraWinDataSource.UltraDataRow(new object[] {
																																																												 "idTipoCampo",
																																																												 "2",
																																																												 "TipoCampo",
																																																												 "Entero"}, new object[0]),
																										 new Infragistics.Win.UltraWinDataSource.UltraDataRow(new object[] {
																																																												 "idTipoCampo",
																																																												 "3",
																																																												 "TipoCampo",
																																																												 "Float"}, new object[0]),
																										 new Infragistics.Win.UltraWinDataSource.UltraDataRow(new object[] {
																																																												 "idTipoCampo",
																																																												 "4",
																																																												 "TipoCampo",
																																																												 "Fecha"}, new object[0]),
																										 new Infragistics.Win.UltraWinDataSource.UltraDataRow(new object[] {
																																																												 "idTipoCampo",
																																																												 "5",
																																																												 "TipoCampo",
																																																												 "Caracter"}, new object[0])});
			// 
			// cmbTipo
			// 
			this.cmbTipo.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbTipo.DataSource = this.udsTipoCampo;
			ultraGridColumn9.Header.VisiblePosition = 0;
			ultraGridColumn9.Hidden = true;
			ultraGridColumn10.Header.VisiblePosition = 1;
			ultraGridColumn10.Width = 128;
			ultraGridBand2.Columns.AddRange(new object[] {
																										 ultraGridColumn9,
																										 ultraGridColumn10});
			this.cmbTipo.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			this.cmbTipo.DisplayMember = "TipoCampo";
			this.cmbTipo.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbTipo.Location = new System.Drawing.Point(168, 96);
			this.cmbTipo.Name = "cmbTipo";
			this.cmbTipo.Size = new System.Drawing.Size(144, 88);
			this.cmbTipo.TabIndex = 3;
			this.cmbTipo.Text = "ultraDropDown1";
			this.cmbTipo.ValueMember = "idTipoCampo";
			this.cmbTipo.Visible = false;
			// 
			// BusquedaTablaCampo
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(536, 266);
			this.Controls.Add(this.cmbTipo);
			this.Controls.Add(this.ultraButton2);
			this.Controls.Add(this.ultraButton1);
			this.Controls.Add(this.ultraGrid1);
			this.Name = "BusquedaTablaCampo";
			this.Text = "Búsqueda Tabla Campo";
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsTablaInstruccion)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.udsTipoCampo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbTipo)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void ultraButton2_Click(object sender, System.EventArgs e)
		{
			Close();
		}

		private void ultraButton1_Click(object sender, System.EventArgs e)
		{
			cdsTablaInstruccion.Update();
			Close();
		}
	}
}
