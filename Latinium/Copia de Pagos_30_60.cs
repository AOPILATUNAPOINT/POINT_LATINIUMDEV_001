using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data.SqlClient;
using System.IO;
using Infragistics.Win.UltraWinGrid;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de Pagos_30_60.
	/// </summary>
	public class Pagos_30_60 : DevExpress.XtraEditors.XtraForm
	{
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor spnValor1;
		private System.Windows.Forms.Label label1;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor spnValor2;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor spnValor4;
		private System.Windows.Forms.Label label2;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor spnValor3;
		private System.Windows.Forms.Label label3;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor spnValor5;
		private System.Windows.Forms.Label label4;
		private Infragistics.Win.Misc.UltraButton btGenerar;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo cmbFecha;
		private C1.Data.C1DataSet cdsSeteoF;
		private Infragistics.Win.UltraWinGrid.UltraGrid ultraGrid1;
		private Infragistics.Win.Misc.UltraButton btFactura;
		private System.Windows.Forms.Label label5;
		private Infragistics.Win.Misc.UltraButton btImprimir;
		private Infragistics.Win.UltraWinEditors.UltraOptionSet optCobrar;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtReporte;
		private System.Windows.Forms.ErrorProvider errorProvider;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public Pagos_30_60()
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
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton1 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.Appearance appearance1 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.ValueListItem valueListItem1 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem2 = new Infragistics.Win.ValueListItem();
			this.cmbFecha = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.spnValor1 = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label1 = new System.Windows.Forms.Label();
			this.spnValor2 = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.spnValor4 = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label2 = new System.Windows.Forms.Label();
			this.spnValor3 = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label3 = new System.Windows.Forms.Label();
			this.spnValor5 = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label4 = new System.Windows.Forms.Label();
			this.btGenerar = new Infragistics.Win.Misc.UltraButton();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.ultraGrid1 = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.btFactura = new Infragistics.Win.Misc.UltraButton();
			this.label5 = new System.Windows.Forms.Label();
			this.btImprimir = new Infragistics.Win.Misc.UltraButton();
			this.optCobrar = new Infragistics.Win.UltraWinEditors.UltraOptionSet();
			this.txtReporte = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.errorProvider = new System.Windows.Forms.ErrorProvider();
			((System.ComponentModel.ISupportInitialize)(this.cmbFecha)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.spnValor1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.spnValor2)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.spnValor4)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.spnValor3)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.spnValor5)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.optCobrar)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtReporte)).BeginInit();
			this.SuspendLayout();
			// 
			// cmbFecha
			// 
			this.cmbFecha.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			dateButton1.Caption = "Today";
			this.cmbFecha.DateButtons.Add(dateButton1);
			this.cmbFecha.Location = new System.Drawing.Point(560, 144);
			this.cmbFecha.Name = "cmbFecha";
			this.cmbFecha.NonAutoSizeHeight = 24;
			this.cmbFecha.Size = new System.Drawing.Size(104, 21);
			this.cmbFecha.TabIndex = 0;
			// 
			// spnValor1
			// 
			this.spnValor1.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.spnValor1.Enabled = false;
			this.spnValor1.Location = new System.Drawing.Point(592, 16);
			this.spnValor1.MaskInput = "nnnn";
			this.spnValor1.Name = "spnValor1";
			this.spnValor1.PromptChar = ' ';
			this.spnValor1.Size = new System.Drawing.Size(72, 21);
			this.spnValor1.SpinButtonDisplayStyle = Infragistics.Win.ButtonDisplayStyle.Always;
			this.spnValor1.TabIndex = 1;
			// 
			// label1
			// 
			this.label1.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.label1.Location = new System.Drawing.Point(512, 32);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(64, 16);
			this.label1.TabIndex = 2;
			this.label1.Text = "Periodo 1";
			// 
			// spnValor2
			// 
			this.spnValor2.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.spnValor2.Location = new System.Drawing.Point(592, 40);
			this.spnValor2.MaskInput = "nnnn";
			this.spnValor2.Name = "spnValor2";
			this.spnValor2.PromptChar = ' ';
			this.spnValor2.Size = new System.Drawing.Size(72, 21);
			this.spnValor2.SpinButtonDisplayStyle = Infragistics.Win.ButtonDisplayStyle.Always;
			this.spnValor2.TabIndex = 3;
			this.spnValor2.Value = 30;
			// 
			// spnValor4
			// 
			this.spnValor4.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.spnValor4.Location = new System.Drawing.Point(592, 88);
			this.spnValor4.MaskInput = "nnnn";
			this.spnValor4.Name = "spnValor4";
			this.spnValor4.PromptChar = ' ';
			this.spnValor4.Size = new System.Drawing.Size(72, 21);
			this.spnValor4.SpinButtonDisplayStyle = Infragistics.Win.ButtonDisplayStyle.Always;
			this.spnValor4.TabIndex = 6;
			this.spnValor4.Value = 90;
			// 
			// label2
			// 
			this.label2.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.label2.Location = new System.Drawing.Point(512, 56);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(64, 16);
			this.label2.TabIndex = 5;
			this.label2.Text = "Periodo 2";
			// 
			// spnValor3
			// 
			this.spnValor3.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.spnValor3.Location = new System.Drawing.Point(592, 64);
			this.spnValor3.MaskInput = "nnnn";
			this.spnValor3.Name = "spnValor3";
			this.spnValor3.PromptChar = ' ';
			this.spnValor3.Size = new System.Drawing.Size(72, 21);
			this.spnValor3.SpinButtonDisplayStyle = Infragistics.Win.ButtonDisplayStyle.Always;
			this.spnValor3.TabIndex = 4;
			this.spnValor3.Value = 60;
			// 
			// label3
			// 
			this.label3.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.label3.Location = new System.Drawing.Point(512, 80);
			this.label3.Name = "label3";
			this.label3.Size = new System.Drawing.Size(64, 16);
			this.label3.TabIndex = 8;
			this.label3.Text = "Periodo 3";
			// 
			// spnValor5
			// 
			this.spnValor5.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.spnValor5.Location = new System.Drawing.Point(592, 112);
			this.spnValor5.MaskInput = "nnnn";
			this.spnValor5.Name = "spnValor5";
			this.spnValor5.PromptChar = ' ';
			this.spnValor5.Size = new System.Drawing.Size(72, 21);
			this.spnValor5.SpinButtonDisplayStyle = Infragistics.Win.ButtonDisplayStyle.Always;
			this.spnValor5.TabIndex = 7;
			this.spnValor5.Value = 120;
			// 
			// label4
			// 
			this.label4.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.label4.Location = new System.Drawing.Point(512, 104);
			this.label4.Name = "label4";
			this.label4.Size = new System.Drawing.Size(64, 16);
			this.label4.TabIndex = 10;
			this.label4.Text = "Periodo 4";
			// 
			// btGenerar
			// 
			this.btGenerar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btGenerar.Location = new System.Drawing.Point(512, 232);
			this.btGenerar.Name = "btGenerar";
			this.btGenerar.TabIndex = 11;
			this.btGenerar.Text = "Generar";
			this.btGenerar.Click += new System.EventHandler(this.btGenerar_Click);
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
			// ultraGrid1
			// 
			this.ultraGrid1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.ultraGrid1.Cursor = System.Windows.Forms.Cursors.Default;
			appearance1.BackColor = System.Drawing.Color.White;
			this.ultraGrid1.DisplayLayout.Appearance = appearance1;
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
			this.ultraGrid1.Location = new System.Drawing.Point(8, 16);
			this.ultraGrid1.Name = "ultraGrid1";
			this.ultraGrid1.Size = new System.Drawing.Size(488, 316);
			this.ultraGrid1.TabIndex = 12;
			this.ultraGrid1.KeyDown += new System.Windows.Forms.KeyEventHandler(this.ultraGrid1_KeyDown);
			// 
			// btFactura
			// 
			this.btFactura.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btFactura.Location = new System.Drawing.Point(600, 232);
			this.btFactura.Name = "btFactura";
			this.btFactura.TabIndex = 13;
			this.btFactura.Text = "Factura";
			this.btFactura.Click += new System.EventHandler(this.btFactura_Click);
			// 
			// label5
			// 
			this.label5.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.label5.AutoSize = true;
			this.label5.Location = new System.Drawing.Point(512, 144);
			this.label5.Name = "label5";
			this.label5.Size = new System.Drawing.Size(36, 16);
			this.label5.TabIndex = 14;
			this.label5.Text = "Fecha";
			// 
			// btImprimir
			// 
			this.btImprimir.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btImprimir.Location = new System.Drawing.Point(512, 264);
			this.btImprimir.Name = "btImprimir";
			this.btImprimir.TabIndex = 15;
			this.btImprimir.Text = "Imprimir";
			this.btImprimir.Click += new System.EventHandler(this.btImprimir_Click);
			// 
			// optCobrar
			// 
			this.optCobrar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.optCobrar.CheckedIndex = 0;
			this.optCobrar.ItemAppearance = appearance6;
			this.optCobrar.ItemOrigin = new System.Drawing.Point(6, 4);
			valueListItem1.DataValue = 1;
			valueListItem1.DisplayText = "Cuentas por Cobrar";
			valueListItem2.DataValue = 4;
			valueListItem2.DisplayText = "Cuentas por Pagar";
			this.optCobrar.Items.Add(valueListItem1);
			this.optCobrar.Items.Add(valueListItem2);
			this.optCobrar.ItemSpacingHorizontal = 10;
			this.optCobrar.ItemSpacingVertical = 4;
			this.optCobrar.Location = new System.Drawing.Point(512, 176);
			this.optCobrar.Name = "optCobrar";
			this.optCobrar.Size = new System.Drawing.Size(152, 48);
			this.optCobrar.TabIndex = 16;
			this.optCobrar.Text = "Cuentas por Cobrar";
			// 
			// txtReporte
			// 
			this.txtReporte.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.txtReporte.Location = new System.Drawing.Point(520, 296);
			this.txtReporte.Name = "txtReporte";
			this.txtReporte.Size = new System.Drawing.Size(128, 21);
			this.txtReporte.TabIndex = 17;
			this.txtReporte.Text = "CxCPeriodo";
			// 
			// errorProvider
			// 
			this.errorProvider.ContainerControl = this;
			// 
			// Pagos_30_60
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(680, 350);
			this.Controls.Add(this.txtReporte);
			this.Controls.Add(this.optCobrar);
			this.Controls.Add(this.btImprimir);
			this.Controls.Add(this.label5);
			this.Controls.Add(this.btFactura);
			this.Controls.Add(this.ultraGrid1);
			this.Controls.Add(this.btGenerar);
			this.Controls.Add(this.label4);
			this.Controls.Add(this.label3);
			this.Controls.Add(this.spnValor5);
			this.Controls.Add(this.spnValor4);
			this.Controls.Add(this.label2);
			this.Controls.Add(this.spnValor3);
			this.Controls.Add(this.spnValor2);
			this.Controls.Add(this.label1);
			this.Controls.Add(this.spnValor1);
			this.Controls.Add(this.cmbFecha);
			this.Name = "Pagos_30_60";
			this.Text = "Pagos a 30, 60, 90 dias";
			this.Closing += new System.ComponentModel.CancelEventHandler(this.Pagos_30_60_Closing);
			this.Load += new System.EventHandler(this.Pagos_30_60_Load);
			((System.ComponentModel.ISupportInitialize)(this.cmbFecha)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.spnValor1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.spnValor2)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.spnValor4)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.spnValor3)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.spnValor5)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.optCobrar)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtReporte)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private bool Verifica()
		{
			bool bOk = true;
      return bOk;
		}

		private void btGenerar_Click(object sender, System.EventArgs e)
		{
			if (!Verifica()) return;
			Cursor = Cursors.WaitCursor;
			DateTime dtFecha = (DateTime) cmbFecha.Value;
			int optTipo = (int) optCobrar.Value;
			string stSelect = "EXEC CxC30 '" + dtFecha.ToString("yyyyMMdd") 
				+ "', " + spnValor2.Value.ToString() + ", " + spnValor3.Value.ToString()
				+ ", " + spnValor4.Value.ToString() + ", " + spnValor5.Value.ToString()
				+ ", " + optTipo.ToString();
			ultraGrid1.DataSource = Funcion.dvProcedimiento(cdsSeteoF, stSelect);
			Cursor = Cursors.Default;
		}

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			this.cdsSeteoF.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);
		}

		private void btFactura_Click(object sender, System.EventArgs e)
		{
			if (ultraGrid1.ActiveRow == null)
			{
				MessageBox.Show("Seleccione la Fila de la cual desea ver su compra", "Información");
				return;
			}
			Cursor = Cursors.WaitCursor;
			int IdCompra = (int) ultraGrid1.ActiveRow.Cells["idCompra"].Value;
			string stSelect = "Select idTipoFactura From Compra "
				+ "Where idCompra = " + IdCompra.ToString();
			int idTipo = Funcion.iEscalarSQL(cdsSeteoF, stSelect, true);
//			Compra miCompra = new Compra(idTipo, IdCompra);
//			miCompra.MdiParent = this.MdiParent;
//			miCompra.Show();
			Cursor = Cursors.Default;
		}

		private void btImprimir_Click(object sender, System.EventArgs e)
		{
			string stNombreRep = txtReporte.Text.Trim() + ".rpt";
			stNombreRep = MenuLatinium.stDirReportes + stNombreRep;
			errorProvider.SetError(txtReporte, "");
			if (!File.Exists(stNombreRep))
			{
				errorProvider.SetError(txtReporte, "Reporte No existe");
				return;
			}
			stNombreRep = txtReporte.Text.Trim() + ".rpt";
			Cursor = Cursors.WaitCursor;
			Reporte miRep = new Reporte(stNombreRep);
			miRep.MdiParent = this.MdiParent;
			miRep.Show();
			Cursor = Cursors.Default;
		}

		private void Pagos_30_60_Closing(object sender, System.ComponentModel.CancelEventArgs e)
		{
			string stAudita = "Exec AuditaCrear 53, 7, 'Pagos_30_60'";
			Funcion.EjecutaSQL(cdsSeteoF, stAudita, true);
		}

		private void Pagos_30_60_Load(object sender, System.EventArgs e)
		{
			string stAudita = "Exec AuditaCrear 53, 6, 'Pagos_30_60'";
			Funcion.EjecutaSQL(cdsSeteoF, stAudita, true);
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
