using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de frmSeteoFechasCumplimientos.
	/// </summary>
	public class frmSeteoFechasCumplimientos : System.Windows.Forms.Form
	{
		private Infragistics.Win.UltraWinGrid.UltraGrid uGridFechas;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource1;
		private Infragistics.Win.Misc.UltraButton btCancelar;
		private Infragistics.Win.Misc.UltraButton btAceptar;
		private C1.Data.C1DataSet cdsSeteoF;

		int idCumplimiento = 0;
		int IdPlanificacion = 0;

		public frmSeteoFechasCumplimientos(int IdCumplimiento, int idPlanificacion)
		{
			//
			// Necesario para admitir el Diseñador de Windows Forms
			//
			InitializeComponent();

			idCumplimiento = IdCumplimiento;
			IdPlanificacion = idPlanificacion;

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
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idFechasCumplimientos");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Dia");
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Mes");
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Año");
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idFechasCumplimientos");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Dia");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn3 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Mes");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn4 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Año");
			this.uGridFechas = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource1 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.btCancelar = new Infragistics.Win.Misc.UltraButton();
			this.btAceptar = new Infragistics.Win.Misc.UltraButton();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			((System.ComponentModel.ISupportInitialize)(this.uGridFechas)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			this.SuspendLayout();
			// 
			// uGridFechas
			// 
			this.uGridFechas.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.uGridFechas.Cursor = System.Windows.Forms.Cursors.Default;
			this.uGridFechas.DataSource = this.ultraDataSource1;
			appearance1.BackColor = System.Drawing.Color.White;
			this.uGridFechas.DisplayLayout.Appearance = appearance1;
			this.uGridFechas.DisplayLayout.AutoFitColumns = true;
			ultraGridBand1.AddButtonCaption = "ClienteTipoRuc";
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn1.Width = 91;
			appearance2.TextHAlign = Infragistics.Win.HAlign.Center;
			appearance2.TextVAlign = Infragistics.Win.VAlign.Middle;
			ultraGridColumn2.CellAppearance = appearance2;
			ultraGridColumn2.Format = "##0";
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridColumn2.Width = 70;
			appearance3.TextHAlign = Infragistics.Win.HAlign.Center;
			appearance3.TextVAlign = Infragistics.Win.VAlign.Middle;
			ultraGridColumn3.CellAppearance = appearance3;
			ultraGridColumn3.Format = "##0";
			ultraGridColumn3.Header.VisiblePosition = 2;
			ultraGridColumn3.Width = 68;
			appearance4.TextHAlign = Infragistics.Win.HAlign.Center;
			appearance4.TextVAlign = Infragistics.Win.VAlign.Middle;
			ultraGridColumn4.CellAppearance = appearance4;
			ultraGridColumn4.Format = "###0";
			ultraGridColumn4.Header.VisiblePosition = 3;
			ultraGridColumn4.Width = 73;
			ultraGridBand1.Columns.AddRange(new object[] {
																										 ultraGridColumn1,
																										 ultraGridColumn2,
																										 ultraGridColumn3,
																										 ultraGridColumn4});
			this.uGridFechas.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.uGridFechas.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.TemplateOnBottom;
			this.uGridFechas.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.True;
			this.uGridFechas.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance5.BackColor = System.Drawing.Color.Transparent;
			this.uGridFechas.DisplayLayout.Override.CardAreaAppearance = appearance5;
			appearance6.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance6.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance6.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance6.FontData.BoldAsString = "True";
			appearance6.FontData.Name = "Arial";
			appearance6.FontData.SizeInPoints = 8.75F;
			appearance6.ForeColor = System.Drawing.Color.White;
			appearance6.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridFechas.DisplayLayout.Override.HeaderAppearance = appearance6;
			this.uGridFechas.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortSingle;
			appearance7.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance7.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance7.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridFechas.DisplayLayout.Override.RowSelectorAppearance = appearance7;
			appearance8.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance8.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance8.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridFechas.DisplayLayout.Override.SelectedRowAppearance = appearance8;
			this.uGridFechas.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridFechas.Location = new System.Drawing.Point(0, 0);
			this.uGridFechas.Name = "uGridFechas";
			this.uGridFechas.Size = new System.Drawing.Size(232, 248);
			this.uGridFechas.TabIndex = 3;
			this.uGridFechas.AfterRowInsert += new Infragistics.Win.UltraWinGrid.RowEventHandler(this.uGridFechas_AfterRowInsert);
			this.uGridFechas.BeforeRowInsert += new Infragistics.Win.UltraWinGrid.BeforeRowInsertEventHandler(this.uGridFechas_BeforeRowInsert);
			this.uGridFechas.KeyDown += new System.Windows.Forms.KeyEventHandler(this.uGridFechas_KeyDown);
			this.uGridFechas.AfterCellUpdate += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.uGridFechas_AfterCellUpdate);
			this.uGridFechas.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.uGridFechas_InitializeLayout);
			// 
			// ultraDataSource1
			// 
			ultraDataColumn1.DataType = typeof(int);
			ultraDataColumn2.DataType = typeof(int);
			ultraDataColumn3.DataType = typeof(int);
			ultraDataColumn4.DataType = typeof(int);
			this.ultraDataSource1.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn1,
																																 ultraDataColumn2,
																																 ultraDataColumn3,
																																 ultraDataColumn4});
			// 
			// btCancelar
			// 
			this.btCancelar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btCancelar.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.btCancelar.DialogResult = System.Windows.Forms.DialogResult.Cancel;
			this.btCancelar.Location = new System.Drawing.Point(88, 256);
			this.btCancelar.Name = "btCancelar";
			this.btCancelar.Size = new System.Drawing.Size(75, 25);
			this.btCancelar.TabIndex = 21;
			this.btCancelar.Text = "&Cancelar";
			this.btCancelar.Click += new System.EventHandler(this.btCancelar_Click);
			// 
			// btAceptar
			// 
			this.btAceptar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btAceptar.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.btAceptar.Location = new System.Drawing.Point(8, 256);
			this.btAceptar.Name = "btAceptar";
			this.btAceptar.Size = new System.Drawing.Size(75, 25);
			this.btAceptar.TabIndex = 20;
			this.btAceptar.Text = "&Aceptar";
			this.btAceptar.Click += new System.EventHandler(this.btAceptar_Click);
			// 
			// cdsSeteoF
			// 
			this.cdsSeteoF.BindingContextCtrl = this;
			this.cdsSeteoF.DataLibrary = "LibFacturacion";
			this.cdsSeteoF.DataLibraryUrl = "";
			this.cdsSeteoF.DataSetDef = "dsSeteoF";
			this.cdsSeteoF.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsSeteoF.Name = "cdsSeteoF";
			this.cdsSeteoF.SchemaClassName = "LibFacturacion.DataClass";
			this.cdsSeteoF.SchemaDef = null;
			this.cdsSeteoF.BeforeFill += new C1.Data.FillEventHandler(this.cdsSeteoF_BeforeFill);
			// 
			// frmSeteoFechasCumplimientos
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(232, 286);
			this.ControlBox = false;
			this.Controls.Add(this.btCancelar);
			this.Controls.Add(this.btAceptar);
			this.Controls.Add(this.uGridFechas);
			this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
			this.KeyPreview = true;
			this.Name = "frmSeteoFechasCumplimientos";
			this.ShowInTaskbar = false;
			this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
			this.Text = "Fechas";
			this.Load += new System.EventHandler(this.frmSeteoFechasCumplimientos_Load);
			((System.ComponentModel.ISupportInitialize)(this.uGridFechas)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		int iFilas = 0;

		private void frmSeteoFechasCumplimientos_Load(object sender, System.EventArgs e)
		{
			this.uGridFechas.DataSource = Funcion.dvProcedimiento(cdsSeteoF, string.Format("Select idFechasCumplimientos, Dia, Mes, Año From FechasCumplimientos Where idCumplimiento = {0}", idCumplimiento));

			if (IdPlanificacion == 1) iFilas = 1;
			if (IdPlanificacion == 2) iFilas = 4;
			if (IdPlanificacion == 3) iFilas = 2;
			if (IdPlanificacion == 4) iFilas = 1;

			if (IdPlanificacion == 1) 
			{
				this.uGridFechas.DisplayLayout.Bands[0].Columns["Mes"].Hidden = true;
				this.uGridFechas.DisplayLayout.Bands[0].Columns["Año"].Hidden = true;
			}

			if (IdPlanificacion == 2 || IdPlanificacion == 3 || IdPlanificacion == 4) 
				this.uGridFechas.DisplayLayout.Bands[0].Columns["Año"].Hidden = true;
		}

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);		
		}

		private void uGridFechas_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
			FuncionesProcedimientos.FormatoGrid(e, "Dia", 0);
			FuncionesProcedimientos.FormatoGrid(e, "Mes", 0);
			FuncionesProcedimientos.FormatoGrid(e, "Año", 0);
		}

		private void btCancelar_Click(object sender, System.EventArgs e)
		{
			this.DialogResult = DialogResult.No;
		}

		private void btAceptar_Click(object sender, System.EventArgs e)
		{
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridFechas.Rows.All)
			{
				string sSQL = string.Format("Exec SeteoGuardaFechasCumplimientos {0}, {1}, {2}, {3}, {4}", 
					(int)dr.Cells["idFechasCumplimientos"].Value, idCumplimiento, (int)dr.Cells["Dia"].Value, (int)dr.Cells["Mes"].Value, (int)dr.Cells["Año"].Value);
				Funcion.EjecutaSQL(cdsSeteoF, sSQL, true);
			}

			this.DialogResult = DialogResult.OK;
		}

		private void uGridFechas_BeforeRowInsert(object sender, Infragistics.Win.UltraWinGrid.BeforeRowInsertEventArgs e)
		{
			if (this.uGridFechas.Rows.Count > iFilas - 1) e.Cancel = true;
		}

		private void uGridFechas_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			FuncionesProcedimientos.DesplazamientoGrids(sender, e, this.uGridFechas);
		}

		private void uGridFechas_AfterCellUpdate(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			if (e.Cell.Column.ToString() == "Dia") if ((int)e.Cell.Row.Cells["Dia"].Value > 31) e.Cell.Row.Cells["Dia"].Value = 1;
			if (e.Cell.Column.ToString() == "Mes") if ((int)e.Cell.Row.Cells["Mes"].Value > 12) e.Cell.Row.Cells["Mes"].Value = 1;
		}

		private void uGridFechas_AfterRowInsert(object sender, Infragistics.Win.UltraWinGrid.RowEventArgs e)
		{
			e.Row.Cells["idFechasCumplimientos"].Value = 0;
			e.Row.Cells["Dia"].Value = 1;
			if (IdPlanificacion == 1) e.Row.Cells["Mes"].Value = 0; else e.Row.Cells["Mes"].Value = 1;
			e.Row.Cells["Año"].Value = 0;
		}
	}
}
