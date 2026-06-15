using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using CrystalDecisions;
using CrystalDecisions.CrystalReports.Engine;
using CrystalDecisions.ReportSource;
using CrystalDecisions.Shared;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de frmReporteDeTransferencias.
	/// </summary>
	public class frmReporteDeTransferencias : System.Windows.Forms.Form
	{
		private System.Windows.Forms.Button btnVer;
		private System.Windows.Forms.Label label53;
		private System.Windows.Forms.Label lblBodega;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtHasta;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtDesde;
		private System.Windows.Forms.Label label1;
		private System.Windows.Forms.Label label;
		public Infragistics.Win.UltraWinGrid.UltraCombo cmbBodega;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbTipoFactura;
		private C1.Data.C1DataSet cdsSeteoF;
		private Infragistics.Win.UltraWinEditors.UltraOptionSet optEstadosTF;
		private Infragistics.Win.UltraWinEditors.UltraOptionSet optEstados;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public frmReporteDeTransferencias()
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
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmReporteDeTransferencias));
			Infragistics.Win.Appearance appearance1 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton1 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton2 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.ValueListItem valueListItem1 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem2 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem3 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem4 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem5 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem6 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bodega");
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idTipoFactura");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.ValueListItem valueListItem7 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem8 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem9 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem10 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem11 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem12 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem13 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem14 = new Infragistics.Win.ValueListItem();
			this.btnVer = new System.Windows.Forms.Button();
			this.label53 = new System.Windows.Forms.Label();
			this.lblBodega = new System.Windows.Forms.Label();
			this.dtHasta = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.dtDesde = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.optEstadosTF = new Infragistics.Win.UltraWinEditors.UltraOptionSet();
			this.label1 = new System.Windows.Forms.Label();
			this.label = new System.Windows.Forms.Label();
			this.cmbBodega = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.cmbTipoFactura = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.optEstados = new Infragistics.Win.UltraWinEditors.UltraOptionSet();
			((System.ComponentModel.ISupportInitialize)(this.dtHasta)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtDesde)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.optEstadosTF)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbBodega)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbTipoFactura)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.optEstados)).BeginInit();
			this.SuspendLayout();
			// 
			// btnVer
			// 
			this.btnVer.Image = ((System.Drawing.Image)(resources.GetObject("btnVer.Image")));
			this.btnVer.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnVer.Location = new System.Drawing.Point(328, 70);
			this.btnVer.Name = "btnVer";
			this.btnVer.Size = new System.Drawing.Size(75, 24);
			this.btnVer.TabIndex = 403;
			this.btnVer.Text = "Ver";
			this.btnVer.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnVer.Click += new System.EventHandler(this.btnVer_Click);
			// 
			// label53
			// 
			this.label53.AutoSize = true;
			this.label53.BackColor = System.Drawing.Color.Transparent;
			this.label53.Location = new System.Drawing.Point(216, 10);
			this.label53.Name = "label53";
			this.label53.Size = new System.Drawing.Size(34, 16);
			this.label53.TabIndex = 402;
			this.label53.Text = "Hasta";
			this.label53.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// lblBodega
			// 
			this.lblBodega.AutoSize = true;
			this.lblBodega.BackColor = System.Drawing.Color.Transparent;
			this.lblBodega.Location = new System.Drawing.Point(8, 10);
			this.lblBodega.Name = "lblBodega";
			this.lblBodega.Size = new System.Drawing.Size(37, 16);
			this.lblBodega.TabIndex = 401;
			this.lblBodega.Text = "Desde";
			this.lblBodega.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// dtHasta
			// 
			appearance1.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtHasta.Appearance = appearance1;
			this.dtHasta.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton1.Caption = "Today";
			this.dtHasta.DateButtons.Add(dateButton1);
			this.dtHasta.Format = "dd/MM/yyyy";
			this.dtHasta.Location = new System.Drawing.Point(272, 8);
			this.dtHasta.Name = "dtHasta";
			this.dtHasta.NonAutoSizeHeight = 23;
			this.dtHasta.Size = new System.Drawing.Size(112, 21);
			this.dtHasta.SpinButtonsVisible = true;
			this.dtHasta.TabIndex = 400;
			this.dtHasta.Value = ((object)(resources.GetObject("dtHasta.Value")));
			// 
			// dtDesde
			// 
			appearance2.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtDesde.Appearance = appearance2;
			this.dtDesde.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton2.Caption = "Today";
			this.dtDesde.DateButtons.Add(dateButton2);
			this.dtDesde.Format = "dd/MM/yyyy";
			this.dtDesde.Location = new System.Drawing.Point(80, 8);
			this.dtDesde.Name = "dtDesde";
			this.dtDesde.NonAutoSizeHeight = 23;
			this.dtDesde.Size = new System.Drawing.Size(112, 21);
			this.dtDesde.SpinButtonsVisible = true;
			this.dtDesde.TabIndex = 399;
			this.dtDesde.Value = ((object)(resources.GetObject("dtDesde.Value")));
			// 
			// optEstadosTF
			// 
			appearance3.ForeColorDisabled = System.Drawing.Color.Black;
			this.optEstadosTF.Appearance = appearance3;
			this.optEstadosTF.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.optEstadosTF.CheckedIndex = 0;
			this.optEstadosTF.ItemAppearance = appearance4;
			this.optEstadosTF.ItemOrigin = new System.Drawing.Point(10, 0);
			valueListItem1.DataValue = -1;
			valueListItem1.DisplayText = "Todos";
			valueListItem2.DataValue = 0;
			valueListItem2.DisplayText = "En Espera";
			valueListItem3.DataValue = 1;
			valueListItem3.DisplayText = "Pendiente";
			valueListItem4.DataValue = 3;
			valueListItem4.DisplayText = "En Tránsito";
			valueListItem5.DataValue = 4;
			valueListItem5.DisplayText = "Procesado";
			valueListItem6.DataValue = 5;
			valueListItem6.DisplayText = "Anulado";
			this.optEstadosTF.Items.Add(valueListItem1);
			this.optEstadosTF.Items.Add(valueListItem2);
			this.optEstadosTF.Items.Add(valueListItem3);
			this.optEstadosTF.Items.Add(valueListItem4);
			this.optEstadosTF.Items.Add(valueListItem5);
			this.optEstadosTF.Items.Add(valueListItem6);
			this.optEstadosTF.ItemSpacingVertical = 10;
			this.optEstadosTF.Location = new System.Drawing.Point(8, 104);
			this.optEstadosTF.Name = "optEstadosTF";
			this.optEstadosTF.Size = new System.Drawing.Size(480, 24);
			this.optEstadosTF.TabIndex = 404;
			this.optEstadosTF.Text = "Todos";
			// 
			// label1
			// 
			this.label1.AutoSize = true;
			this.label1.BackColor = System.Drawing.Color.Transparent;
			this.label1.Location = new System.Drawing.Point(8, 42);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(43, 16);
			this.label1.TabIndex = 408;
			this.label1.Text = "Bodega";
			// 
			// label
			// 
			this.label.AutoSize = true;
			this.label.BackColor = System.Drawing.Color.Transparent;
			this.label.Location = new System.Drawing.Point(8, 74);
			this.label.Name = "label";
			this.label.Size = new System.Drawing.Size(62, 16);
			this.label.TabIndex = 407;
			this.label.Text = "Documento";
			// 
			// cmbBodega
			// 
			appearance5.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbBodega.Appearance = appearance5;
			this.cmbBodega.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbBodega.Cursor = System.Windows.Forms.Cursors.Default;
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn1.Width = 60;
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridColumn2.Width = 250;
			ultraGridColumn3.Header.VisiblePosition = 2;
			ultraGridColumn3.Width = 50;
			ultraGridBand1.Columns.AddRange(new object[] {
																										 ultraGridColumn1,
																										 ultraGridColumn2,
																										 ultraGridColumn3});
			this.cmbBodega.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.cmbBodega.DisplayMember = "Nombre";
			this.cmbBodega.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbBodega.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbBodega.Enabled = false;
			this.cmbBodega.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbBodega.Location = new System.Drawing.Point(80, 40);
			this.cmbBodega.Name = "cmbBodega";
			this.cmbBodega.Size = new System.Drawing.Size(240, 21);
			this.cmbBodega.TabIndex = 405;
			this.cmbBodega.ValueMember = "Bodega";
			// 
			// cmbTipoFactura
			// 
			appearance6.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbTipoFactura.Appearance = appearance6;
			this.cmbTipoFactura.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbTipoFactura.Cursor = System.Windows.Forms.Cursors.Default;
			ultraGridColumn4.Header.VisiblePosition = 0;
			ultraGridColumn4.Hidden = true;
			ultraGridColumn5.Header.VisiblePosition = 1;
			ultraGridColumn5.Width = 250;
			ultraGridBand2.Columns.AddRange(new object[] {
																										 ultraGridColumn4,
																										 ultraGridColumn5});
			this.cmbTipoFactura.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			this.cmbTipoFactura.DisplayLayout.Override.NullText = "";
			this.cmbTipoFactura.DisplayMember = "Nombre";
			this.cmbTipoFactura.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbTipoFactura.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbTipoFactura.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbTipoFactura.Location = new System.Drawing.Point(80, 72);
			this.cmbTipoFactura.Name = "cmbTipoFactura";
			this.cmbTipoFactura.Size = new System.Drawing.Size(240, 21);
			this.cmbTipoFactura.TabIndex = 406;
			this.cmbTipoFactura.ValueMember = "idTipoFactura";
			this.cmbTipoFactura.ValueChanged += new System.EventHandler(this.cmbTipoFactura_ValueChanged);
			this.cmbTipoFactura.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.cmbTipoFactura_InitializeLayout);
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
			// optEstados
			// 
			appearance7.ForeColorDisabled = System.Drawing.Color.Black;
			this.optEstados.Appearance = appearance7;
			this.optEstados.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.optEstados.CheckedIndex = 0;
			this.optEstados.ItemAppearance = appearance8;
			this.optEstados.ItemOrigin = new System.Drawing.Point(10, 0);
			valueListItem7.DataValue = 0;
			valueListItem7.DisplayText = "Todos";
			valueListItem8.DataValue = 1;
			valueListItem8.DisplayText = "Pendiente";
			valueListItem9.DataValue = 3;
			valueListItem9.DisplayText = "Procesado";
			valueListItem10.DataValue = 4;
			valueListItem10.DisplayText = "Rechazado";
			valueListItem11.DataValue = 5;
			valueListItem11.DisplayText = "Aprobado";
			valueListItem12.DataValue = 7;
			valueListItem12.DisplayText = "Liquidado";
			valueListItem13.DataValue = 9;
			valueListItem13.DisplayText = "Facturado";
			valueListItem14.DataValue = 16;
			valueListItem14.DisplayText = "Anulado";
			this.optEstados.Items.Add(valueListItem7);
			this.optEstados.Items.Add(valueListItem8);
			this.optEstados.Items.Add(valueListItem9);
			this.optEstados.Items.Add(valueListItem10);
			this.optEstados.Items.Add(valueListItem11);
			this.optEstados.Items.Add(valueListItem12);
			this.optEstados.Items.Add(valueListItem13);
			this.optEstados.Items.Add(valueListItem14);
			this.optEstados.ItemSpacingVertical = 10;
			this.optEstados.Location = new System.Drawing.Point(8, 104);
			this.optEstados.Name = "optEstados";
			this.optEstados.Size = new System.Drawing.Size(640, 24);
			this.optEstados.TabIndex = 409;
			this.optEstados.Text = "Todos";
			this.optEstados.Visible = false;
			// 
			// frmReporteDeTransferencias
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(658, 136);
			this.Controls.Add(this.optEstados);
			this.Controls.Add(this.label1);
			this.Controls.Add(this.label);
			this.Controls.Add(this.cmbBodega);
			this.Controls.Add(this.cmbTipoFactura);
			this.Controls.Add(this.optEstadosTF);
			this.Controls.Add(this.btnVer);
			this.Controls.Add(this.label53);
			this.Controls.Add(this.lblBodega);
			this.Controls.Add(this.dtHasta);
			this.Controls.Add(this.dtDesde);
			this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
			this.KeyPreview = true;
			this.MaximizeBox = false;
			this.Name = "frmReporteDeTransferencias";
			this.Text = "Movimiento de Inventarios";
			this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.frmReporteDeTransferencias_KeyDown);
			this.Load += new System.EventHandler(this.frmReporteDeTransferencias_Load);
			((System.ComponentModel.ISupportInitialize)(this.dtHasta)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtDesde)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.optEstadosTF)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbBodega)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbTipoFactura)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.optEstados)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void btnVer_Click(object sender, System.EventArgs e)
		{
			#region Valida Fechas
			if (this.dtDesde.Value == System.DBNull.Value)
			{
				MessageBox.Show("Seleccione la Fecha Inicial", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.dtDesde.Focus();
				return;
			}

			if (this.dtHasta.Value == System.DBNull.Value)
			{
				MessageBox.Show("Seleccione la Fecha Final", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.dtHasta.Focus();
				return;
			}

			if (Convert.ToDateTime(this.dtDesde.Value) > Convert.ToDateTime(this.dtHasta.Value))
			{
				MessageBox.Show("La Fecha Inicial no puede ser mayor a la Fecha Final", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.dtHasta.Focus();
				return;
			}
			#endregion Valida Fechas

			#region Parametros
			ParameterFields paramFields = new ParameterFields ();
											
			ParameterField paramFieldCompra = new ParameterField ();
			ParameterDiscreteValue ValIdCompra = new ParameterDiscreteValue ();
			paramFieldCompra.ParameterFieldName = "@Desde";
			ValIdCompra.Value = Convert.ToDateTime(this.dtDesde.Value);
			paramFieldCompra.CurrentValues.Add (ValIdCompra);
			paramFields.Add (paramFieldCompra);
							
			ParameterField paramFieldFecha = new ParameterField ();
			ParameterDiscreteValue ValIdFecha = new ParameterDiscreteValue ();
			paramFieldFecha.ParameterFieldName = "@Hasta";
			ValIdFecha.Value = Convert.ToDateTime(this.dtHasta.Value);
			paramFieldFecha.CurrentValues.Add (ValIdFecha);
			paramFields.Add (paramFieldFecha);
			
			int iEstado = (int)this.optEstadosTF.Value;
			string sReporte = "ReporteDeTransferencias.rpt";

			if ((int)this.cmbTipoFactura.Value != 42)
			{
				ParameterField paramFieldidTipoFactura = new ParameterField ();
				ParameterDiscreteValue ValIdidTipoFactura = new ParameterDiscreteValue ();
				paramFieldidTipoFactura.ParameterFieldName = "@idTipoFactura";
				ValIdidTipoFactura.Value = Convert.ToInt32(this.cmbTipoFactura.Value);
				paramFieldidTipoFactura.CurrentValues.Add (ValIdidTipoFactura);
				paramFields.Add (paramFieldidTipoFactura);

				iEstado = (int)this.optEstados.Value;
				sReporte = "ReporteDeMovimientosInventario.rpt";
			}

			ParameterField paramFieldEstado = new ParameterField ();
			ParameterDiscreteValue ValIdEstado = new ParameterDiscreteValue ();
			paramFieldEstado.ParameterFieldName = "@Estado";
			ValIdEstado.Value = Convert.ToInt32(iEstado);
			paramFieldEstado.CurrentValues.Add (ValIdEstado);
			paramFields.Add (paramFieldEstado);
			
			int iBodegas = 0;
			if (this.cmbBodega.ActiveRow != null) iBodegas = (int)this.cmbBodega.Value;

			ParameterField paramFieldBodega = new ParameterField ();
			ParameterDiscreteValue ValIdBodega = new ParameterDiscreteValue ();
			paramFieldBodega.ParameterFieldName = "@Bodega";
			ValIdBodega.Value = iBodegas;
			paramFieldBodega.CurrentValues.Add (ValIdBodega);
			paramFields.Add (paramFieldBodega);
			#endregion Parametros		
			
			Reporte Rep = new Reporte(sReporte, "");
			Rep.MdiParent = this.MdiParent;			
			Rep.crVista.ParameterFieldInfo = paramFields;
			Rep.Show();
		}

		private void frmReporteDeTransferencias_Load(object sender, System.EventArgs e)
		{
			this.dtDesde.CalendarInfo.MaxDate = DateTime.Today;			
			this.dtHasta.CalendarInfo.MaxDate = DateTime.Today;

			this.dtDesde.Value = "01/" + DateTime.Now.Month.ToString() + "/" + DateTime.Now.Year.ToString();			
			this.dtHasta.Value = DateTime.Today;

			Funcion.Resolucion(this, this.Height, this.Width);

			string sSQL = string.Format("Exec ListaBodegas {0}, 48", MenuLatinium.IdUsuario);
			this.cmbBodega.DataSource = Funcion.dvProcedimiento(cdsSeteoF, sSQL);

			sSQL = string.Format("Exec BodegaAsignadaPorUsuario {0}", MenuLatinium.IdUsuario);
			int iBodega = Funcion.iEscalarSQL(cdsSeteoF, sSQL);

			if (iBodega == 0) this.cmbBodega.Enabled = true;

			this.cmbTipoFactura.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Select idTipoFactura, Nombre From CompraNumero Where idTipoFactura In (2, 13, 42, 44)");

			this.cmbTipoFactura.Value = 42;
		}

		private void frmReporteDeTransferencias_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Escape) this.Close();
		}

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);	
		}

		private void cmbTipoFactura_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}

		private void cmbTipoFactura_ValueChanged(object sender, System.EventArgs e)
		{
			if (this.cmbTipoFactura.ActiveRow != null)
			{
				if ((int)this.cmbTipoFactura.Value == 42)
				{
					this.optEstadosTF.Visible = true;
					this.optEstados.Visible = false;
					this.optEstadosTF.Value = -1;
				}	
				else				
				{
					this.optEstadosTF.Visible = false;
					this.optEstados.Visible = true;
					this.optEstados.Value = 0;
				}	
			}
		}
	}
}
