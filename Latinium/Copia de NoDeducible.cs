using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de NoDeducible.
	/// </summary>
	public class NoDeducible : DevExpress.XtraEditors.XtraForm
	{
		private DevExpress.XtraEditors.SimpleButton btGenerar;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo cmdHasta;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo cmdDesde;
		private Infragistics.Win.UltraWinGrid.UltraGrid grdNoDeducible;
		private C1.Data.C1DataSet cdsSeteoC;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbVendedor;
		private DevExpress.XtraEditors.SimpleButton btAsignar;
		private DevExpress.XtraEditors.SimpleButton btImprimir;
		private DevExpress.XtraEditors.SimpleButton btAsiento;
		private System.Windows.Forms.Label label1;
		private System.Windows.Forms.Label label2;
		private System.Windows.Forms.Label label3;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public NoDeducible()
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
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton1 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton2 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idNoDeducible");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Detalle");
			this.grdNoDeducible = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.btGenerar = new DevExpress.XtraEditors.SimpleButton();
			this.cmdHasta = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.cmdDesde = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.cdsSeteoC = new C1.Data.C1DataSet();
			this.cmbVendedor = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.btAsignar = new DevExpress.XtraEditors.SimpleButton();
			this.btImprimir = new DevExpress.XtraEditors.SimpleButton();
			this.btAsiento = new DevExpress.XtraEditors.SimpleButton();
			this.label1 = new System.Windows.Forms.Label();
			this.label2 = new System.Windows.Forms.Label();
			this.label3 = new System.Windows.Forms.Label();
			((System.ComponentModel.ISupportInitialize)(this.grdNoDeducible)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmdHasta)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmdDesde)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoC)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbVendedor)).BeginInit();
			this.SuspendLayout();
			// 
			// grdNoDeducible
			// 
			this.grdNoDeducible.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.grdNoDeducible.Cursor = System.Windows.Forms.Cursors.Default;
			appearance1.BackColor = System.Drawing.Color.White;
			this.grdNoDeducible.DisplayLayout.Appearance = appearance1;
			this.grdNoDeducible.DisplayLayout.AutoFitColumns = true;
			appearance2.BackColor = System.Drawing.Color.Transparent;
			this.grdNoDeducible.DisplayLayout.Override.CardAreaAppearance = appearance2;
			appearance3.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance3.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance3.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance3.FontData.BoldAsString = "True";
			appearance3.FontData.Name = "Arial";
			appearance3.FontData.SizeInPoints = 10F;
			appearance3.ForeColor = System.Drawing.Color.White;
			appearance3.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.grdNoDeducible.DisplayLayout.Override.HeaderAppearance = appearance3;
			this.grdNoDeducible.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti;
			appearance4.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance4.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance4.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.grdNoDeducible.DisplayLayout.Override.RowSelectorAppearance = appearance4;
			appearance5.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance5.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance5.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.grdNoDeducible.DisplayLayout.Override.SelectedRowAppearance = appearance5;
			this.grdNoDeducible.DisplayLayout.ViewStyleBand = Infragistics.Win.UltraWinGrid.ViewStyleBand.OutlookGroupBy;
			this.grdNoDeducible.Location = new System.Drawing.Point(16, 96);
			this.grdNoDeducible.Name = "grdNoDeducible";
			this.grdNoDeducible.Size = new System.Drawing.Size(736, 360);
			this.grdNoDeducible.TabIndex = 2;
			// 
			// btGenerar
			// 
			this.btGenerar.Location = new System.Drawing.Point(376, 16);
			this.btGenerar.Name = "btGenerar";
			this.btGenerar.TabIndex = 3;
			this.btGenerar.Text = "Generar";
			this.btGenerar.Click += new System.EventHandler(this.btGenerar_Click);
			// 
			// cmdHasta
			// 
			appearance6.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmdHasta.Appearance = appearance6;
			this.cmdHasta.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton1.Caption = "Today";
			this.cmdHasta.DateButtons.Add(dateButton1);
			this.cmdHasta.Format = "dd/MMM/yyyy";
			this.cmdHasta.Location = new System.Drawing.Point(256, 16);
			this.cmdHasta.Name = "cmdHasta";
			this.cmdHasta.NonAutoSizeHeight = 23;
			this.cmdHasta.Size = new System.Drawing.Size(112, 21);
			this.cmdHasta.SpinButtonsVisible = true;
			this.cmdHasta.TabIndex = 6;
			this.cmdHasta.Value = new System.DateTime(2005, 6, 8, 0, 0, 0, 0);
			// 
			// cmdDesde
			// 
			appearance7.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmdDesde.Appearance = appearance7;
			this.cmdDesde.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton2.Caption = "Today";
			this.cmdDesde.DateButtons.Add(dateButton2);
			this.cmdDesde.Format = "dd/MMM/yyyy";
			this.cmdDesde.Location = new System.Drawing.Point(72, 16);
			this.cmdDesde.Name = "cmdDesde";
			this.cmdDesde.NonAutoSizeHeight = 23;
			this.cmdDesde.Size = new System.Drawing.Size(112, 21);
			this.cmdDesde.SpinButtonsVisible = true;
			this.cmdDesde.TabIndex = 7;
			this.cmdDesde.Value = new System.DateTime(2005, 6, 8, 0, 0, 0, 0);
			// 
			// cdsSeteoC
			// 
			this.cdsSeteoC.BindingContextCtrl = this;
			this.cdsSeteoC.DataLibrary = "LibContabilidad";
			this.cdsSeteoC.DataLibraryUrl = "";
			this.cdsSeteoC.DataSetDef = "dsSeteoC";
			this.cdsSeteoC.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsSeteoC.Name = "cdsSeteoC";
			this.cdsSeteoC.SchemaClassName = "LibContabilidad.DataClass";
			this.cdsSeteoC.SchemaDef = null;
			this.cdsSeteoC.BeforeFill += new C1.Data.FillEventHandler(this.cdsSeteoC_BeforeFill);
			// 
			// cmbVendedor
			// 
			appearance8.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbVendedor.Appearance = appearance8;
			this.cmbVendedor.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbVendedor.Cursor = System.Windows.Forms.Cursors.Default;
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridColumn2.Width = 389;
			ultraGridBand1.Columns.AddRange(new object[] {
																										 ultraGridColumn1,
																										 ultraGridColumn2});
			this.cmbVendedor.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.cmbVendedor.DisplayMember = "Detalle";
			this.cmbVendedor.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbVendedor.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbVendedor.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbVendedor.Location = new System.Drawing.Point(144, 56);
			this.cmbVendedor.Name = "cmbVendedor";
			this.cmbVendedor.Size = new System.Drawing.Size(224, 21);
			this.cmbVendedor.TabIndex = 91;
			this.cmbVendedor.ValueMember = "idNoDeducible";
			this.cmbVendedor.MouseUp += new System.Windows.Forms.MouseEventHandler(this.cmbVendedor_MouseUp);
			// 
			// btAsignar
			// 
			this.btAsignar.Location = new System.Drawing.Point(376, 56);
			this.btAsignar.Name = "btAsignar";
			this.btAsignar.TabIndex = 92;
			this.btAsignar.Text = "Asignar";
			this.btAsignar.Click += new System.EventHandler(this.btAsignar_Click);
			// 
			// btImprimir
			// 
			this.btImprimir.Location = new System.Drawing.Point(672, 16);
			this.btImprimir.Name = "btImprimir";
			this.btImprimir.Size = new System.Drawing.Size(72, 23);
			this.btImprimir.TabIndex = 93;
			this.btImprimir.Text = "Imprimir";
			this.btImprimir.Click += new System.EventHandler(this.btImprimir_Click);
			// 
			// btAsiento
			// 
			this.btAsiento.Location = new System.Drawing.Point(672, 56);
			this.btAsiento.Name = "btAsiento";
			this.btAsiento.TabIndex = 94;
			this.btAsiento.Text = "Asiento";
			this.btAsiento.Click += new System.EventHandler(this.btAsiento_Click);
			// 
			// label1
			// 
			this.label1.AutoSize = true;
			this.label1.Location = new System.Drawing.Point(24, 16);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(37, 16);
			this.label1.TabIndex = 95;
			this.label1.Text = "Desde";
			// 
			// label2
			// 
			this.label2.AutoSize = true;
			this.label2.Location = new System.Drawing.Point(24, 56);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(113, 16);
			this.label2.TabIndex = 96;
			this.label2.Text = "Tipo de No Deducible";
			// 
			// label3
			// 
			this.label3.AutoSize = true;
			this.label3.Location = new System.Drawing.Point(208, 16);
			this.label3.Name = "label3";
			this.label3.Size = new System.Drawing.Size(34, 16);
			this.label3.TabIndex = 97;
			this.label3.Text = "Hasta";
			// 
			// NoDeducible
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(768, 478);
			this.Controls.Add(this.label3);
			this.Controls.Add(this.label2);
			this.Controls.Add(this.label1);
			this.Controls.Add(this.btAsiento);
			this.Controls.Add(this.btImprimir);
			this.Controls.Add(this.btAsignar);
			this.Controls.Add(this.cmbVendedor);
			this.Controls.Add(this.cmdDesde);
			this.Controls.Add(this.cmdHasta);
			this.Controls.Add(this.btGenerar);
			this.Controls.Add(this.grdNoDeducible);
			this.Name = "NoDeducible";
			this.Text = "Conciliación Tributaria";
			this.Load += new System.EventHandler(this.NoDeducible_Load);
			((System.ComponentModel.ISupportInitialize)(this.grdNoDeducible)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmdHasta)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmdDesde)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoC)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbVendedor)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void btGenerar_Click(object sender, System.EventArgs e)
		{
			Cursor = Cursors.WaitCursor;
			DateTime dtDesde = (DateTime) cmdDesde.Value;
			DateTime dtHasta = (DateTime) cmdHasta.Value;
			string stExec = string.Format("Exec NoDeducibleGenera '{0}', '{1}', 0", dtDesde, dtHasta);
			grdNoDeducible.DataSource  = Funcion.dvProcedimiento(cdsSeteoC, stExec);
			grdNoDeducible.DisplayLayout.Bands[0].Columns["idDetAsiento"].Hidden = true;										
			grdNoDeducible.DisplayLayout.Bands[0].Columns["idAsiento"].Hidden = true;
			MessageBox.Show("Listado Generado", "Información");
			Cursor = Cursors.Default;
		}

		private void NoDeducible_Load(object sender, System.EventArgs e)
		{
			int iAnio = DateTime.Today.Year - 1;
			cmdDesde.Value = new DateTime(iAnio, 1, 1);
			cmdHasta.Value = new DateTime(iAnio, 12, 31);

			cmbVendedor.DataSource  = Funcion.dvProcedimiento(cdsSeteoC, "Select idNiif_TipoDTDP As idNoDeducible, Concepto AS Detalle from dbo.Niif_TipoDTDP");
			cmbVendedor.DisplayMember = "Detalle";
			cmbVendedor.ValueMember = "idNoDeducible";
		}

		private void cdsSeteoC_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoC.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirConta);
		}

		private void btAsignar_Click(object sender, System.EventArgs e)
		{
			Cursor = Cursors.WaitCursor;
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in grdNoDeducible.Selected.Rows)
			{
				int idAsientoDetalle = (int) dr.Cells["idDetAsiento"].Value;
				int idNoDeducible = 0;
				if (cmbVendedor.Value != null) idNoDeducible = (int) cmbVendedor.Value;
				string stUpdate = string.Format("Update AsientoDetalle Set idNoDeducible = {0} Where idDetAsiento = {1}",
					idNoDeducible, idAsientoDetalle);
				Funcion.EjecutaSQL(cdsSeteoC, stUpdate);
				dr.Cells["Detalle"].Value = cmbVendedor.Text;
			}
			Cursor = Cursors.Default;
		}

		private void btImprimir_Click(object sender, System.EventArgs e)
		{
			Cursor = Cursors.WaitCursor;
			DateTime dtDesde = (DateTime) cmdDesde.Value;
			DateTime dtHasta = (DateTime) cmdHasta.Value;

			string stFiltro = string.Format(
			"{{Asiento.Borrar}} = False And {{Asiento.Fecha}} >= #{0}# And {{Asiento.Fecha}} < #{1}# And {{AsientoDetalle.idNodeducible}} > 0",
				dtDesde.ToString("MM/dd/yyyy"), dtHasta.AddDays(1).ToString("MM/dd/yyyy"));
			Reporte miRep = new Reporte("ConciliacionTributaria.rpt", stFiltro);
			miRep.Titulo(string.Format("Desde {0} Hasta {1}", 
				dtDesde.ToString("dd/MMM/yyyy"), dtHasta.ToString("dd/MMM/yyyy")));
			miRep.MdiParent = this.MdiParent;
			miRep.Show();
			Cursor = Cursors.Default;
		}

		private void cmbVendedor_MouseUp(object sender, System.Windows.Forms.MouseEventArgs e)
		{
			if (e.Button != MouseButtons.Right) return;
			cmbVendedor.Value = 0;
			btAsiento.Select();
			cmbVendedor.Select();
		}

		private void btAsiento_Click(object sender, System.EventArgs e)
		{
			if (grdNoDeducible.ActiveRow == null ||	grdNoDeducible.ActiveRow.Cells["idAsiento"].Value == DBNull.Value)
			{
				MessageBox.Show("Seleccione el Asiento que desea ver", "Información",
					MessageBoxButtons.OK, MessageBoxIcon.Information);
				return;
			}
			Cursor = Cursors.WaitCursor;
			int idAsiento = (int) grdNoDeducible.ActiveRow.Cells["idAsiento"].Value;
			Asiento miAsiento = new Asiento(idAsiento);
			miAsiento.MdiParent = this.MdiParent;
			miAsiento.Show();
		}
	}
}
