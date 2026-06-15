using System;
using System.IO;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using Infragistics.Win.UltraWinGrid;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de ReporteLista.
	/// </summary>
	public class ReporteLista : DevExpress.XtraEditors.XtraForm
	{
		private C1.Data.C1DataSet cdsReporte;
		private Infragistics.Win.UltraWinGrid.UltraGrid ultraGrid1;
		private Infragistics.Win.Misc.UltraButton btAceptar;
		private Infragistics.Win.Misc.UltraButton btCancelar;
		private Infragistics.Win.UltraWinGrid.UltraDropDown cmbTipo;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource2;
		private Infragistics.Win.Misc.UltraButton btReportes;
		private C1.Data.C1DataSet cdsSeteo;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public ReporteLista()
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
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Reporte", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Detalle");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idGrupo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idReporte");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Reporte");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Tipo", -1, "cmbTipo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Filtro");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("NoFiltro");
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Reporte");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idRep");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Reporte");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idRep");
			this.cdsReporte = new C1.Data.C1DataSet();
			this.ultraGrid1 = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.btAceptar = new Infragistics.Win.Misc.UltraButton();
			this.btCancelar = new Infragistics.Win.Misc.UltraButton();
			this.cmbTipo = new Infragistics.Win.UltraWinGrid.UltraDropDown();
			this.ultraDataSource2 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.btReportes = new Infragistics.Win.Misc.UltraButton();
			this.cdsSeteo = new C1.Data.C1DataSet();
			((System.ComponentModel.ISupportInitialize)(this.cdsReporte)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbTipo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteo)).BeginInit();
			this.SuspendLayout();
			// 
			// cdsReporte
			// 
			this.cdsReporte.BindingContextCtrl = this;
			this.cdsReporte.DataLibrary = "LibEmpresa";
			this.cdsReporte.DataLibraryUrl = "";
			this.cdsReporte.DataSetDef = "dsReporte";
			this.cdsReporte.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsReporte.Name = "cdsReporte";
			this.cdsReporte.SchemaClassName = "LibEmpresa.DataClass";
			this.cdsReporte.SchemaDef = null;
			// 
			// ultraGrid1
			// 
			this.ultraGrid1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.ultraGrid1.Cursor = System.Windows.Forms.Cursors.Default;
			this.ultraGrid1.DataMember = "Reporte";
			this.ultraGrid1.DataSource = this.cdsReporte;
			appearance1.BackColor = System.Drawing.Color.White;
			this.ultraGrid1.DisplayLayout.Appearance = appearance1;
			this.ultraGrid1.DisplayLayout.AutoFitColumns = true;
			ultraGridColumn1.AllowGroupBy = Infragistics.Win.DefaultableBoolean.False;
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn1.Width = 104;
			ultraGridColumn2.Header.Caption = "Grupo";
			ultraGridColumn2.Header.VisiblePosition = 3;
			ultraGridColumn2.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownList;
			ultraGridColumn2.Width = 56;
			ultraGridColumn3.AllowGroupBy = Infragistics.Win.DefaultableBoolean.False;
			ultraGridColumn3.Header.VisiblePosition = 4;
			ultraGridColumn3.Hidden = true;
			ultraGridColumn3.Width = 59;
			ultraGridColumn4.AllowGroupBy = Infragistics.Win.DefaultableBoolean.False;
			ultraGridColumn4.Header.VisiblePosition = 1;
			ultraGridColumn4.Width = 97;
			ultraGridColumn5.Header.VisiblePosition = 2;
			ultraGridColumn5.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownList;
			ultraGridColumn5.Width = 56;
			ultraGridColumn6.Header.VisiblePosition = 5;
			ultraGridColumn6.Width = 78;
			ultraGridColumn7.Header.VisiblePosition = 6;
			ultraGridColumn7.Width = 84;
			ultraGridBand1.Columns.AddRange(new object[] {
																										 ultraGridColumn1,
																										 ultraGridColumn2,
																										 ultraGridColumn3,
																										 ultraGridColumn4,
																										 ultraGridColumn5,
																										 ultraGridColumn6,
																										 ultraGridColumn7});
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
			this.ultraGrid1.DisplayLayout.ViewStyleBand = Infragistics.Win.UltraWinGrid.ViewStyleBand.OutlookGroupBy;
			this.ultraGrid1.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.ultraGrid1.Location = new System.Drawing.Point(16, 16);
			this.ultraGrid1.Name = "ultraGrid1";
			this.ultraGrid1.Size = new System.Drawing.Size(496, 200);
			this.ultraGrid1.TabIndex = 0;
			this.ultraGrid1.KeyDown += new System.Windows.Forms.KeyEventHandler(this.ultraGrid1_KeyDown);
			// 
			// btAceptar
			// 
			this.btAceptar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btAceptar.Location = new System.Drawing.Point(56, 232);
			this.btAceptar.Name = "btAceptar";
			this.btAceptar.TabIndex = 1;
			this.btAceptar.Text = "&Aceptar";
			this.btAceptar.Click += new System.EventHandler(this.btAceptar_Click);
			// 
			// btCancelar
			// 
			this.btCancelar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.btCancelar.Location = new System.Drawing.Point(360, 232);
			this.btCancelar.Name = "btCancelar";
			this.btCancelar.TabIndex = 2;
			this.btCancelar.Text = "&Cancelar";
			this.btCancelar.Click += new System.EventHandler(this.btCancelar_Click);
			// 
			// cmbTipo
			// 
			this.cmbTipo.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbTipo.DataMember = "Band 0";
			this.cmbTipo.DataSource = this.ultraDataSource2;
			this.cmbTipo.DisplayLayout.AutoFitColumns = true;
			ultraGridColumn8.Header.VisiblePosition = 0;
			ultraGridColumn8.Width = 181;
			ultraGridColumn9.Header.VisiblePosition = 1;
			ultraGridColumn9.Hidden = true;
			ultraGridBand2.Columns.AddRange(new object[] {
																										 ultraGridColumn8,
																										 ultraGridColumn9});
			this.cmbTipo.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			this.cmbTipo.DisplayMember = "Reporte";
			this.cmbTipo.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbTipo.Location = new System.Drawing.Point(208, 112);
			this.cmbTipo.Name = "cmbTipo";
			this.cmbTipo.Size = new System.Drawing.Size(200, 80);
			this.cmbTipo.TabIndex = 3;
			this.cmbTipo.ValueMember = "idRep";
			this.cmbTipo.Visible = false;
			// 
			// ultraDataSource2
			// 
			ultraDataColumn2.DataType = typeof(int);
			ultraDataColumn2.DefaultValue = 0;
			this.ultraDataSource2.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn1,
																																 ultraDataColumn2});
			this.ultraDataSource2.Rows.AddRange(new object[] {
																												 new Infragistics.Win.UltraWinDataSource.UltraDataRow(new object[] {
																																																														 "Reporte",
																																																														 "Resumenes"}, new object[0]),
																												 new Infragistics.Win.UltraWinDataSource.UltraDataRow(new object[] {
																																																														 "Reporte",
																																																														 "Articulos",
																																																														 "idRep",
																																																														 1}, new object[0]),
																												 new Infragistics.Win.UltraWinDataSource.UltraDataRow(new object[] {
																																																														 "Reporte",
																																																														 "Produccion",
																																																														 "idRep",
																																																														 2}, new object[0]),
																												 new Infragistics.Win.UltraWinDataSource.UltraDataRow(new object[] {
																																																														 "Reporte",
																																																														 "Inventario",
																																																														 "idRep",
																																																														 3}, new object[0]),
																												 new Infragistics.Win.UltraWinDataSource.UltraDataRow(new object[] {
																																																														 "Reporte",
																																																														 "Contabilidad",
																																																														 "idRep",
																																																														 4}, new object[0]),
																												 new Infragistics.Win.UltraWinDataSource.UltraDataRow(new object[] {
																																																														 "Reporte",
																																																														 "Facturacion",
																																																														 "idRep",
																																																														 5}, new object[0]),
																												 new Infragistics.Win.UltraWinDataSource.UltraDataRow(new object[] {
																																																														 "Reporte",
																																																														 "Roles",
																																																														 "idRep",
																																																														 6}, new object[0]),
																												 new Infragistics.Win.UltraWinDataSource.UltraDataRow(new object[] {
																																																														 "Reporte",
																																																														 "Presupuesto",
																																																														 "idRep",
																																																														 7}, new object[0]),
																												 new Infragistics.Win.UltraWinDataSource.UltraDataRow(new object[] {
																																																														 "Reporte",
																																																														 "NIIFF",
																																																														 "idRep",
																																																														 8}, new object[0]),
																												 new Infragistics.Win.UltraWinDataSource.UltraDataRow(new object[] {
																																																														 "Reporte",
																																																														 "Personal",
																																																														 "idRep",
																																																														 9}, new object[0]),
																												 new Infragistics.Win.UltraWinDataSource.UltraDataRow(new object[] {
																																																														 "Reporte",
																																																														 "Avanzado",
																																																														 "idRep",
																																																														 10}, new object[0])});
			// 
			// btReportes
			// 
			this.btReportes.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btReportes.Location = new System.Drawing.Point(160, 232);
			this.btReportes.Name = "btReportes";
			this.btReportes.TabIndex = 4;
			this.btReportes.Text = "&Reportes";
			this.btReportes.Click += new System.EventHandler(this.btReportes_Click);
			// 
			// cdsSeteo
			// 
			this.cdsSeteo.BindingContextCtrl = this;
			this.cdsSeteo.DataLibrary = "LibContabilidad";
			this.cdsSeteo.DataLibraryUrl = "";
			this.cdsSeteo.DataSetDef = "dsSeteo";
			this.cdsSeteo.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsSeteo.Name = "cdsSeteo";
			this.cdsSeteo.SchemaClassName = "LibContabilidad.DataClass";
			this.cdsSeteo.SchemaDef = null;
			this.cdsSeteo.BeforeFill += new C1.Data.FillEventHandler(this.cdsSeteo_BeforeFill);
			// 
			// ReporteLista
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(528, 266);
			this.Controls.Add(this.btReportes);
			this.Controls.Add(this.cmbTipo);
			this.Controls.Add(this.btCancelar);
			this.Controls.Add(this.btAceptar);
			this.Controls.Add(this.ultraGrid1);
			this.Name = "ReporteLista";
			this.Text = "Lista de Reportes";
			this.Closing += new System.ComponentModel.CancelEventHandler(this.ReporteLista_Closing);
			this.Load += new System.EventHandler(this.ReporteLista_Load);
			((System.ComponentModel.ISupportInitialize)(this.cdsReporte)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbTipo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteo)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void btAceptar_Click(object sender, System.EventArgs e)
		{
			try
			{
				this.cdsReporte.Update();
			}
			catch(System.Exception ex)
			{
				MessageBox.Show(ex.Message, "Información");
				return;
			}
			this.Close();
		}

		private void btCancelar_Click(object sender, System.EventArgs e)
		{
			this.Close();
		}

		private void ReporteLista_Load(object sender, System.EventArgs e)
		{
			string stAudita = "Exec AuditaCrear 53, 6, 'ReporteLista'";
			Funcion.EjecutaSQL(cdsSeteo, stAudita, true);
			Cursor = Cursors.Default;
		}

		private void btReportes_Click(object sender, System.EventArgs e)
		{
			try
			{
				Cursor = Cursors.WaitCursor;
				string stArchivo = MenuLatinium.stDirReportes.Trim();
				if (!stArchivo.EndsWith("\\")) stArchivo += "\\";
				if (!Directory.Exists(stArchivo))
				{
					MessageBox.Show("No existe Directorio: " + stArchivo,
						"Información", MessageBoxButtons.OK, MessageBoxIcon.Information);
					Cursor = Cursors.Default;
					return;
				}
				stArchivo += ultraGrid1.ActiveRow.Cells["Reporte"].Value.ToString().Trim();
				if (!stArchivo.ToUpper().EndsWith(".RPT")) stArchivo += ".RPT";

				if (!File.Exists(stArchivo))
				{
					MessageBox.Show("No existe Archivo: " + stArchivo,
						"Información", MessageBoxButtons.OK, MessageBoxIcon.Information);
					Cursor = Cursors.Default;
					return;
				}
				System.Diagnostics.Process.Start(stArchivo);
				Cursor = Cursors.Default;
			}
			catch(System.Exception ex)
			{
				MessageBox.Show("No Existe Editor de Reportes crystal Reports", ex.Message);
			}
		}

		private void ReporteLista_Closing(object sender, System.ComponentModel.CancelEventArgs e)
		{
			string stAudita = "Exec AuditaCrear 53, 7, 'ReporteLista'";
			Funcion.EjecutaSQL(cdsSeteo, stAudita, true);
		}

		private void cdsSeteo_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteo.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirConta);
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

	}
}
