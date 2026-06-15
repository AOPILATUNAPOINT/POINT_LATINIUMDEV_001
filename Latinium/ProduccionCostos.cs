using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using CrystalDecisions.Shared;
using System.Globalization;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de ProduccionCostos.
	/// </summary>
	public class ProduccionCostos : DevExpress.XtraEditors.XtraForm
	{
		private System.Windows.Forms.Label label1;
		private System.Windows.Forms.Label label2;
		private Infragistics.Win.Misc.UltraButton btManoObra;
		private System.Windows.Forms.Label label3;
		private System.Windows.Forms.Label label4;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkResumen;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtManoObra;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtPorcDistrib;
		private Infragistics.Win.UltraWinEditors.UltraOptionSet optReporte;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtGastosFabricacion;
		private System.Windows.Forms.Label label5;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo cmbDesde;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo cmbHasta;
		private Infragistics.Win.Misc.UltraButton btSemiterminado;
		private Infragistics.Win.Misc.UltraButton btTerminado;
		private C1.Data.C1DataSet cdsSeteo;
		private System.Windows.Forms.Label label6;
		private System.Windows.Forms.Label label7;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtHora;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtInd;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public ProduccionCostos()
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
			Infragistics.Win.ValueListItem valueListItem1 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem2 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton1 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton2 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			this.label1 = new System.Windows.Forms.Label();
			this.txtManoObra = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtPorcDistrib = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label2 = new System.Windows.Forms.Label();
			this.btManoObra = new Infragistics.Win.Misc.UltraButton();
			this.label3 = new System.Windows.Forms.Label();
			this.label4 = new System.Windows.Forms.Label();
			this.chkResumen = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.optReporte = new Infragistics.Win.UltraWinEditors.UltraOptionSet();
			this.txtGastosFabricacion = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label5 = new System.Windows.Forms.Label();
			this.cmbDesde = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.cmbHasta = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.btSemiterminado = new Infragistics.Win.Misc.UltraButton();
			this.btTerminado = new Infragistics.Win.Misc.UltraButton();
			this.cdsSeteo = new C1.Data.C1DataSet();
			this.label6 = new System.Windows.Forms.Label();
			this.txtHora = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label7 = new System.Windows.Forms.Label();
			this.txtInd = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			((System.ComponentModel.ISupportInitialize)(this.txtManoObra)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtPorcDistrib)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.optReporte)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtGastosFabricacion)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbDesde)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbHasta)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtHora)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtInd)).BeginInit();
			this.SuspendLayout();
			// 
			// label1
			// 
			this.label1.Location = new System.Drawing.Point(24, 64);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(112, 23);
			this.label1.TabIndex = 0;
			this.label1.Text = "Mano de Obra Mes";
			// 
			// txtManoObra
			// 
			this.txtManoObra.FormatString = "#,##0.00";
			this.txtManoObra.Location = new System.Drawing.Point(168, 64);
			this.txtManoObra.Name = "txtManoObra";
			this.txtManoObra.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtManoObra.PromptChar = ' ';
			this.txtManoObra.Size = new System.Drawing.Size(96, 21);
			this.txtManoObra.TabIndex = 1;
			this.txtManoObra.Value = 25000;
			// 
			// txtPorcDistrib
			// 
			this.txtPorcDistrib.FormatString = "#,##0.00";
			this.txtPorcDistrib.Location = new System.Drawing.Point(200, 120);
			this.txtPorcDistrib.Name = "txtPorcDistrib";
			this.txtPorcDistrib.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtPorcDistrib.PromptChar = ' ';
			this.txtPorcDistrib.Size = new System.Drawing.Size(64, 21);
			this.txtPorcDistrib.TabIndex = 2;
			this.txtPorcDistrib.Value = 20;
			this.txtPorcDistrib.ValueChanged += new System.EventHandler(this.ultraNumericEditor2_ValueChanged);
			// 
			// label2
			// 
			this.label2.Location = new System.Drawing.Point(24, 120);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(120, 23);
			this.label2.TabIndex = 3;
			this.label2.Text = "Porcentaje en Proceso";
			// 
			// btManoObra
			// 
			this.btManoObra.Location = new System.Drawing.Point(24, 152);
			this.btManoObra.Name = "btManoObra";
			this.btManoObra.Size = new System.Drawing.Size(104, 23);
			this.btManoObra.TabIndex = 4;
			this.btManoObra.Text = "Reporte";
			this.btManoObra.Click += new System.EventHandler(this.btManoObra_Click);
			// 
			// label3
			// 
			this.label3.AutoSize = true;
			this.label3.Location = new System.Drawing.Point(32, 8);
			this.label3.Name = "label3";
			this.label3.Size = new System.Drawing.Size(37, 16);
			this.label3.TabIndex = 7;
			this.label3.Text = "Desde";
			// 
			// label4
			// 
			this.label4.AutoSize = true;
			this.label4.Location = new System.Drawing.Point(32, 40);
			this.label4.Name = "label4";
			this.label4.Size = new System.Drawing.Size(34, 16);
			this.label4.TabIndex = 8;
			this.label4.Text = "Hasta";
			// 
			// chkResumen
			// 
			this.chkResumen.Location = new System.Drawing.Point(40, 184);
			this.chkResumen.Name = "chkResumen";
			this.chkResumen.Size = new System.Drawing.Size(88, 20);
			this.chkResumen.TabIndex = 9;
			this.chkResumen.Text = "Resumen";
			// 
			// optReporte
			// 
			this.optReporte.CheckedIndex = 0;
			this.optReporte.ItemAppearance = appearance1;
			this.optReporte.ItemOrigin = new System.Drawing.Point(6, 4);
			valueListItem1.DataValue = 0;
			valueListItem1.DisplayText = "Mano de Obra";
			valueListItem2.DataValue = 1;
			valueListItem2.DisplayText = "Reparto";
			this.optReporte.Items.Add(valueListItem1);
			this.optReporte.Items.Add(valueListItem2);
			this.optReporte.ItemSpacingVertical = 4;
			this.optReporte.Location = new System.Drawing.Point(160, 152);
			this.optReporte.Name = "optReporte";
			this.optReporte.Size = new System.Drawing.Size(104, 48);
			this.optReporte.TabIndex = 10;
			this.optReporte.Text = "Mano de Obra";
			// 
			// txtGastosFabricacion
			// 
			this.txtGastosFabricacion.FormatString = "#,##0.00";
			this.txtGastosFabricacion.Location = new System.Drawing.Point(168, 88);
			this.txtGastosFabricacion.Name = "txtGastosFabricacion";
			this.txtGastosFabricacion.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtGastosFabricacion.PromptChar = ' ';
			this.txtGastosFabricacion.Size = new System.Drawing.Size(96, 21);
			this.txtGastosFabricacion.TabIndex = 12;
			this.txtGastosFabricacion.Value = 30000;
			// 
			// label5
			// 
			this.label5.Location = new System.Drawing.Point(24, 88);
			this.label5.Name = "label5";
			this.label5.Size = new System.Drawing.Size(112, 23);
			this.label5.TabIndex = 11;
			this.label5.Text = "Gastos Fabricación";
			// 
			// cmbDesde
			// 
			dateButton1.Caption = "Today";
			this.cmbDesde.DateButtons.Add(dateButton1);
			this.cmbDesde.Format = "dd/MMM/yyyy";
			this.cmbDesde.Location = new System.Drawing.Point(144, 8);
			this.cmbDesde.Name = "cmbDesde";
			this.cmbDesde.NonAutoSizeHeight = 24;
			this.cmbDesde.Size = new System.Drawing.Size(120, 21);
			this.cmbDesde.SpinButtonIncrement = Infragistics.Win.UltraWinSchedule.SpinIncrementUnit.Months;
			this.cmbDesde.SpinButtonsVisible = true;
			this.cmbDesde.TabIndex = 13;
			this.cmbDesde.Value = new System.DateTime(2009, 7, 5, 0, 0, 0, 0);
			// 
			// cmbHasta
			// 
			dateButton2.Caption = "Today";
			this.cmbHasta.DateButtons.Add(dateButton2);
			this.cmbHasta.Format = "dd/MMM/yyyy";
			this.cmbHasta.Location = new System.Drawing.Point(144, 32);
			this.cmbHasta.Name = "cmbHasta";
			this.cmbHasta.NonAutoSizeHeight = 24;
			this.cmbHasta.Size = new System.Drawing.Size(120, 21);
			this.cmbHasta.SpinButtonIncrement = Infragistics.Win.UltraWinSchedule.SpinIncrementUnit.Months;
			this.cmbHasta.SpinButtonsVisible = true;
			this.cmbHasta.TabIndex = 14;
			this.cmbHasta.Value = new System.DateTime(2009, 7, 5, 0, 0, 0, 0);
			// 
			// btSemiterminado
			// 
			this.btSemiterminado.Location = new System.Drawing.Point(32, 256);
			this.btSemiterminado.Name = "btSemiterminado";
			this.btSemiterminado.Size = new System.Drawing.Size(104, 23);
			this.btSemiterminado.TabIndex = 15;
			this.btSemiterminado.Text = "Semiterminado";
			this.btSemiterminado.Click += new System.EventHandler(this.btSemiterminado_Click);
			// 
			// btTerminado
			// 
			this.btTerminado.Location = new System.Drawing.Point(160, 256);
			this.btTerminado.Name = "btTerminado";
			this.btTerminado.Size = new System.Drawing.Size(104, 23);
			this.btTerminado.TabIndex = 16;
			this.btTerminado.Text = "Terminado";
			this.btTerminado.Click += new System.EventHandler(this.btTerminado_Click);
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
			// label6
			// 
			this.label6.Location = new System.Drawing.Point(24, 224);
			this.label6.Name = "label6";
			this.label6.Size = new System.Drawing.Size(64, 23);
			this.label6.TabIndex = 18;
			this.label6.Text = "Costo min";
			// 
			// txtHora
			// 
			this.txtHora.FormatString = "#,##0.0000";
			this.txtHora.Location = new System.Drawing.Point(88, 224);
			this.txtHora.MaskInput = "{LOC}nnnnnnn.nnnn";
			this.txtHora.Name = "txtHora";
			this.txtHora.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtHora.PromptChar = ' ';
			this.txtHora.Size = new System.Drawing.Size(64, 21);
			this.txtHora.TabIndex = 17;
			this.txtHora.Value = 1;
			// 
			// label7
			// 
			this.label7.Location = new System.Drawing.Point(160, 224);
			this.label7.Name = "label7";
			this.label7.Size = new System.Drawing.Size(40, 23);
			this.label7.TabIndex = 20;
			this.label7.Text = "% Ind.";
			// 
			// txtInd
			// 
			this.txtInd.FormatString = "#,##0.00";
			this.txtInd.Location = new System.Drawing.Point(208, 224);
			this.txtInd.Name = "txtInd";
			this.txtInd.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtInd.PromptChar = ' ';
			this.txtInd.Size = new System.Drawing.Size(64, 21);
			this.txtInd.TabIndex = 19;
			this.txtInd.Value = 5;
			// 
			// ProduccionCostos
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(292, 294);
			this.Controls.Add(this.label7);
			this.Controls.Add(this.txtInd);
			this.Controls.Add(this.label6);
			this.Controls.Add(this.txtHora);
			this.Controls.Add(this.btTerminado);
			this.Controls.Add(this.btSemiterminado);
			this.Controls.Add(this.cmbHasta);
			this.Controls.Add(this.cmbDesde);
			this.Controls.Add(this.txtGastosFabricacion);
			this.Controls.Add(this.label5);
			this.Controls.Add(this.optReporte);
			this.Controls.Add(this.chkResumen);
			this.Controls.Add(this.label4);
			this.Controls.Add(this.label3);
			this.Controls.Add(this.btManoObra);
			this.Controls.Add(this.label2);
			this.Controls.Add(this.txtPorcDistrib);
			this.Controls.Add(this.txtManoObra);
			this.Controls.Add(this.label1);
			this.MaximizeBox = false;
			this.Name = "ProduccionCostos";
			this.Text = "Costos de Producción";
			this.Load += new System.EventHandler(this.ProduccionCostos_Load);
			((System.ComponentModel.ISupportInitialize)(this.txtManoObra)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtPorcDistrib)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.optReporte)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtGastosFabricacion)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbDesde)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbHasta)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtHora)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtInd)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void ultraNumericEditor2_ValueChanged(object sender, System.EventArgs e)
		{
		
		}

		private void btManoObra_Click(object sender, System.EventArgs e)
		{
			int iReporte = (int) optReporte.Value;
			string stReporte = "ProduccionCostoReparto.rpt";
			if (iReporte == 0) stReporte = "ProduccionCosto.rpt";
			Cursor = Cursors.WaitCursor;
			DateTime dtDesde = (DateTime) cmbDesde.Value;
			DateTime dtHasta = (DateTime) cmbHasta.Value;
			Reporte miRep = new Reporte(stReporte, "");
			string stTitulo = "Calculo de Minutos de Mano de Obra y Gastos de Fabricacion "
				+ "del " + dtDesde.ToString("dd/MMM/yyyy") + " al "
				+ dtHasta.ToString("dd/MMM/yyyy");

			#region Parametros
			ParameterFields paramFields = new ParameterFields ();
			ParameterField paramField = new ParameterField ();
			ParameterDiscreteValue discreteVal = new ParameterDiscreteValue ();

			paramField.ParameterFieldName = "@FechaDesde";
			discreteVal.Value = dtDesde.ToString("yyyy-MM-dd");
			paramField.CurrentValues.Add (discreteVal);
			paramFields.Add (paramField);

			ParameterDiscreteValue discreteVal1 = new ParameterDiscreteValue ();
			ParameterField paramField1 = new ParameterField ();
			paramField1.ParameterFieldName = "@FechaHasta";
			discreteVal1.Value = dtHasta.AddDays(1).ToString("yyyy-MM-dd");
			paramField1.CurrentValues.Add (discreteVal1);
			paramFields.Add (paramField1);
			
			ParameterDiscreteValue discreteVal2 = new ParameterDiscreteValue ();
			ParameterField paramField2 = new ParameterField ();
			paramField2.ParameterFieldName = "@Porc";
			discreteVal2.Value = (double) txtPorcDistrib.Value;
			paramField2.CurrentValues.Add (discreteVal2);
			paramFields.Add (paramField2);

			ParameterDiscreteValue discreteVal3 = new ParameterDiscreteValue ();
			ParameterField paramField3 = new ParameterField ();
			paramField3.ParameterFieldName = "ValorManoObra";
			discreteVal3.Value = (double) txtManoObra.Value;
			paramField3.CurrentValues.Add (discreteVal3);
			paramFields.Add (paramField3);

			if (iReporte == 1)
			{
				ParameterDiscreteValue discreteVal4 = new ParameterDiscreteValue ();
				ParameterField paramField4 = new ParameterField ();
				paramField4.ParameterFieldName = "GastosFabricacion";
				discreteVal4.Value = (double) txtGastosFabricacion.Value;
				paramField4.CurrentValues.Add (discreteVal4);
				paramFields.Add (paramField4);
			}
			miRep.crVista.ParameterFieldInfo = paramFields;
			#endregion Parametros
			
			if (chkResumen.Checked) miRep.Resumen();
			miRep.Titulo(stTitulo);
			miRep.MdiParent = this.MdiParent;
			miRep.Show();
			Cursor = Cursors.Default;
		}

		private void ProduccionCostos_Load(object sender, System.EventArgs e)
		{
			cmbDesde.Value = DateTime.Today.AddDays(1 -DateTime.Today.Day);
			cmbHasta.Value = DateTime.Today;
		}

		private void btSemiterminado_Click(object sender, System.EventArgs e)
		{
			DateTime dtDesde = (DateTime) cmbDesde.Value;
			DateTime dtHasta = (DateTime) cmbHasta.Value;
			string stExec = "Exec CostoSemiterminado '"
				+ dtDesde.ToString("yyyyMMdd") + "', '"
				+ dtHasta.ToString("yyyyMMdd") + "'";
			Funcion.EjecutaSQL(cdsSeteo, stExec, true);
			MessageBox.Show("Costo de Semiterminado Generado",
				"Información");
		}

		private void cdsSeteo_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			this.cdsSeteo.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirConta);		
		}

		CultureInfo us = new CultureInfo("en-US");
		private void btTerminado_Click(object sender, System.EventArgs e)
		{
			double dHora = (double) txtHora.Value;
			double dInd = (double) txtInd.Value;
			DateTime dtDesde = (DateTime) cmbDesde.Value;
			DateTime dtHasta = (DateTime) cmbHasta.Value;
			string stExec = "Exec CostoTerminado '"
				+ dtDesde.ToString("yyyyMMdd") + "', '"
				+ dtHasta.ToString("yyyyMMdd") + "', "
				+ dHora.ToString("0.00000", us) + ", " + dInd.ToString("0.00000", us);
			Funcion.EjecutaSQL(cdsSeteo, stExec, true);
			MessageBox.Show("Costo de Terminado Generado",
				"Información");
		}
	}
}
