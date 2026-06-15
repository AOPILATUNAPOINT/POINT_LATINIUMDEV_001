using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using DevExpress.XtraEditors;
using System.Data;
using System.Data.SqlClient;
using CrystalDecisions.Shared;
using System.Net;
using System.Web;
using System.Xml;
using System.Xml.Schema;
using System.IO;
using System.Diagnostics;


namespace Latinium
{
	/// <summary>
	/// Descripción breve de frmCuentaAuxiliarRepote.
	/// </summary>
	public class frmCuentaAuxiliarRepote : System.Windows.Forms.Form
	{
		private C1.Data.C1DataSet cdsSeteoF;
		private System.Windows.Forms.Label lbl_registros;
		private Infragistics.Win.UltraWinEditors.UltraOptionSet optTipoResultado;
		private System.Windows.Forms.Label label1;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtDesde;
		private System.Windows.Forms.Label label18;
		private System.Windows.Forms.Label label17;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtHasta;
		private System.Windows.Forms.GroupBox groupBox2;
		private Infragistics.Win.UltraWinEditors.UltraOptionSet optReportes;
		private System.Windows.Forms.Button button1;
		private System.Windows.Forms.Button btnSalir;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbPeriodo;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource udsPeriodo;
		/// <summary>
		/// Variable del diseñador requerida.
		int idCuentaAuxiliarPeriodoGenerado = 0;
		/// </summary>
		private System.ComponentModel.Container components = null;

		public frmCuentaAuxiliarRepote()
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
			Infragistics.Win.ValueListItem valueListItem1 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem2 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCuentaAuxiliarPeriodoGenerado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Periodo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCuentaAuxiliarPeriodoGenerado");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Periodo");
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton1 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmCuentaAuxiliarRepote));
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton2 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.ValueListItem valueListItem3 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem4 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem5 = new Infragistics.Win.ValueListItem();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.lbl_registros = new System.Windows.Forms.Label();
			this.optTipoResultado = new Infragistics.Win.UltraWinEditors.UltraOptionSet();
			this.label1 = new System.Windows.Forms.Label();
			this.cmbPeriodo = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.udsPeriodo = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.dtDesde = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.label18 = new System.Windows.Forms.Label();
			this.label17 = new System.Windows.Forms.Label();
			this.dtHasta = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.groupBox2 = new System.Windows.Forms.GroupBox();
			this.optReportes = new Infragistics.Win.UltraWinEditors.UltraOptionSet();
			this.button1 = new System.Windows.Forms.Button();
			this.btnSalir = new System.Windows.Forms.Button();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.optTipoResultado)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbPeriodo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.udsPeriodo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtDesde)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtHasta)).BeginInit();
			this.groupBox2.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.optReportes)).BeginInit();
			this.SuspendLayout();
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
			// lbl_registros
			// 
			this.lbl_registros.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lbl_registros.ForeColor = System.Drawing.Color.Red;
			this.lbl_registros.Location = new System.Drawing.Point(8, 120);
			this.lbl_registros.Name = "lbl_registros";
			this.lbl_registros.Size = new System.Drawing.Size(224, 23);
			this.lbl_registros.TabIndex = 758;
			this.lbl_registros.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// optTipoResultado
			// 
			appearance1.ForeColorDisabled = System.Drawing.Color.Black;
			this.optTipoResultado.Appearance = appearance1;
			this.optTipoResultado.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.optTipoResultado.Enabled = false;
			this.optTipoResultado.ItemAppearance = appearance2;
			this.optTipoResultado.ItemOrigin = new System.Drawing.Point(10, 0);
			valueListItem1.DataValue = 0;
			valueListItem1.DisplayText = "Del Mes";
			valueListItem2.DataValue = 1;
			valueListItem2.DisplayText = "Acumulado";
			this.optTipoResultado.Items.Add(valueListItem1);
			this.optTipoResultado.Items.Add(valueListItem2);
			this.optTipoResultado.ItemSpacingVertical = 10;
			this.optTipoResultado.Location = new System.Drawing.Point(8, 88);
			this.optTipoResultado.Name = "optTipoResultado";
			this.optTipoResultado.Size = new System.Drawing.Size(224, 24);
			this.optTipoResultado.TabIndex = 759;
			// 
			// label1
			// 
			this.label1.Location = new System.Drawing.Point(8, 24);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(80, 23);
			this.label1.TabIndex = 1043;
			this.label1.Text = "* Período";
			// 
			// cmbPeriodo
			// 
			appearance3.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbPeriodo.Appearance = appearance3;
			this.cmbPeriodo.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbPeriodo.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbPeriodo.DataSource = this.udsPeriodo;
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn2.Header.Caption = "Período";
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridColumn2.Width = 154;
			ultraGridBand1.Columns.AddRange(new object[] {
																										 ultraGridColumn1,
																										 ultraGridColumn2});
			this.cmbPeriodo.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.cmbPeriodo.DisplayMember = "Periodo";
			this.cmbPeriodo.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbPeriodo.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbPeriodo.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbPeriodo.Location = new System.Drawing.Point(88, 24);
			this.cmbPeriodo.Name = "cmbPeriodo";
			this.cmbPeriodo.Size = new System.Drawing.Size(144, 21);
			this.cmbPeriodo.TabIndex = 1044;
			this.cmbPeriodo.ValueMember = "idCuentaAuxiliarPeriodoGenerado";
			this.cmbPeriodo.ValueChanged += new System.EventHandler(this.cmbPeriodo_ValueChanged);
			// 
			// udsPeriodo
			// 
			ultraDataColumn1.DataType = typeof(int);
			ultraDataColumn1.ReadOnly = Infragistics.Win.DefaultableBoolean.False;
			ultraDataColumn2.ReadOnly = Infragistics.Win.DefaultableBoolean.False;
			this.udsPeriodo.Band.Columns.AddRange(new object[] {
																													 ultraDataColumn1,
																													 ultraDataColumn2});
			// 
			// dtDesde
			// 
			appearance4.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtDesde.Appearance = appearance4;
			this.dtDesde.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton1.Caption = "Today";
			this.dtDesde.DateButtons.Add(dateButton1);
			this.dtDesde.Format = "dd/MM/yyyy";
			this.dtDesde.Location = new System.Drawing.Point(72, 176);
			this.dtDesde.Name = "dtDesde";
			this.dtDesde.NonAutoSizeHeight = 23;
			this.dtDesde.Size = new System.Drawing.Size(112, 21);
			this.dtDesde.SpinButtonsVisible = true;
			this.dtDesde.TabIndex = 751;
			this.dtDesde.Value = ((object)(resources.GetObject("dtDesde.Value")));
			this.dtDesde.Visible = false;
			// 
			// label18
			// 
			this.label18.AutoSize = true;
			this.label18.Location = new System.Drawing.Point(16, 176);
			this.label18.Name = "label18";
			this.label18.Size = new System.Drawing.Size(40, 16);
			this.label18.TabIndex = 753;
			this.label18.Text = "Desde:";
			this.label18.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.label18.Visible = false;
			// 
			// label17
			// 
			this.label17.AutoSize = true;
			this.label17.Location = new System.Drawing.Point(16, 208);
			this.label17.Name = "label17";
			this.label17.Size = new System.Drawing.Size(37, 16);
			this.label17.TabIndex = 754;
			this.label17.Text = "Hasta:";
			this.label17.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.label17.Visible = false;
			// 
			// dtHasta
			// 
			appearance5.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtHasta.Appearance = appearance5;
			this.dtHasta.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton2.Caption = "Today";
			this.dtHasta.DateButtons.Add(dateButton2);
			this.dtHasta.Format = "dd/MM/yyyy";
			this.dtHasta.Location = new System.Drawing.Point(72, 208);
			this.dtHasta.Name = "dtHasta";
			this.dtHasta.NonAutoSizeHeight = 23;
			this.dtHasta.Size = new System.Drawing.Size(112, 21);
			this.dtHasta.SpinButtonsVisible = true;
			this.dtHasta.TabIndex = 752;
			this.dtHasta.Value = ((object)(resources.GetObject("dtHasta.Value")));
			this.dtHasta.Visible = false;
			// 
			// groupBox2
			// 
			this.groupBox2.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.groupBox2.Controls.Add(this.optReportes);
			this.groupBox2.Controls.Add(this.optTipoResultado);
			this.groupBox2.Controls.Add(this.label1);
			this.groupBox2.Controls.Add(this.cmbPeriodo);
			this.groupBox2.Controls.Add(this.lbl_registros);
			this.groupBox2.Location = new System.Drawing.Point(16, 8);
			this.groupBox2.Name = "groupBox2";
			this.groupBox2.Size = new System.Drawing.Size(248, 160);
			this.groupBox2.TabIndex = 1045;
			this.groupBox2.TabStop = false;
			this.groupBox2.Text = "Datos principales";
			// 
			// optReportes
			// 
			appearance6.ForeColorDisabled = System.Drawing.Color.Black;
			this.optReportes.Appearance = appearance6;
			this.optReportes.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.optReportes.ItemAppearance = appearance7;
			this.optReportes.ItemOrigin = new System.Drawing.Point(10, 0);
			valueListItem3.DataValue = 0;
			valueListItem3.DisplayText = "Activos";
			valueListItem4.DataValue = 3;
			valueListItem4.DisplayText = "Pasivos";
			valueListItem5.DataValue = 1;
			valueListItem5.DisplayText = "Resultados";
			this.optReportes.Items.Add(valueListItem3);
			this.optReportes.Items.Add(valueListItem4);
			this.optReportes.Items.Add(valueListItem5);
			this.optReportes.ItemSpacingVertical = 10;
			this.optReportes.Location = new System.Drawing.Point(8, 56);
			this.optReportes.Name = "optReportes";
			this.optReportes.Size = new System.Drawing.Size(224, 24);
			this.optReportes.TabIndex = 1045;
			// 
			// button1
			// 
			this.button1.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.button1.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(230)), ((System.Byte)(230)), ((System.Byte)(230)));
			this.button1.CausesValidation = false;
			this.button1.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.button1.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(50)), ((System.Byte)(50)), ((System.Byte)(50)));
			this.button1.Image = ((System.Drawing.Image)(resources.GetObject("button1.Image")));
			this.button1.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.button1.Location = new System.Drawing.Point(280, 16);
			this.button1.Name = "button1";
			this.button1.Size = new System.Drawing.Size(88, 24);
			this.button1.TabIndex = 1046;
			this.button1.Text = "Imprimir";
			this.button1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.button1.Click += new System.EventHandler(this.button1_Click);
			// 
			// btnSalir
			// 
			this.btnSalir.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btnSalir.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(230)), ((System.Byte)(230)), ((System.Byte)(230)));
			this.btnSalir.CausesValidation = false;
			this.btnSalir.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.btnSalir.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(50)), ((System.Byte)(50)), ((System.Byte)(50)));
			this.btnSalir.Image = ((System.Drawing.Image)(resources.GetObject("btnSalir.Image")));
			this.btnSalir.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnSalir.Location = new System.Drawing.Point(280, 48);
			this.btnSalir.Name = "btnSalir";
			this.btnSalir.Size = new System.Drawing.Size(88, 23);
			this.btnSalir.TabIndex = 1047;
			this.btnSalir.Text = "Salir";
			this.btnSalir.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnSalir.Click += new System.EventHandler(this.btnSalir_Click);
			// 
			// frmCuentaAuxiliarRepote
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(376, 238);
			this.Controls.Add(this.btnSalir);
			this.Controls.Add(this.button1);
			this.Controls.Add(this.groupBox2);
			this.Controls.Add(this.dtDesde);
			this.Controls.Add(this.label17);
			this.Controls.Add(this.dtHasta);
			this.Controls.Add(this.label18);
			this.Name = "frmCuentaAuxiliarRepote";
			this.Text = "Reporte estados financieros";
			this.Load += new System.EventHandler(this.frmCuentaAuxiliarRepote_Load);
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.optTipoResultado)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbPeriodo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.udsPeriodo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtDesde)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtHasta)).EndInit();
			this.groupBox2.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.optReportes)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void frmCuentaAuxiliarRepote_Load(object sender, System.EventArgs e)
		{
			string ambiente = "DEV"; //PRO

			if(ambiente == "PRO")
			{
				if (!Funcion.Permiso("2008", cdsSeteoF))
				{
					MessageBox.Show("No puede ingresar a Configuración Estados Financieros", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					BeginInvoke(new MethodInvoker(UnloadMe));
					return;
				}
			}

			this.dtDesde.Value = "01/" + DateTime.Now.Month.ToString() + "/" + DateTime.Now.Year.ToString();
			this.dtHasta.Value = DateTime.DaysInMonth(DateTime.Now.Year, DateTime.Now.Month).ToString() + "/" + DateTime.Now.Month.ToString() + "/" + DateTime.Now.Year.ToString();
			DateTime dtFecha = FuncionesProcedimientos.dtRetornaFechaServidor(cdsSeteoF);
			this.dtDesde.CalendarInfo.MinDate = DateTime.Parse("01/01/2014");
			this.dtHasta.CalendarInfo.MinDate = DateTime.Parse("01/01/2014");
			this.dtDesde.CalendarInfo.MaxDate = dtFecha;
			this.dtHasta.CalendarInfo.MaxDate = dtFecha;
			this.dtDesde.Value = "01/" + dtFecha.Month.ToString() + "/" + dtFecha.Year.ToString();
			this.dtHasta.Value = dtFecha.AddDays(-1);
			this.optReportes.CheckedIndex = 0;
			this.optTipoResultado.CheckedIndex = 0;

			//Cargar la tabla de periodos, para el reporte
			this.cmbPeriodo.DataSource = Funcion.dvProcedimiento(cdsSeteoF, String.Format("EXEC CuentaAuxiliarPeriodoGeneradoCon"));


		}

		private void UnloadMe()
		{
			this.Close();
		}


		private void groupBox1_Enter(object sender, System.EventArgs e)
		{
		
		}
		string TipoReporte=null;
		string Tipo=null;

		private void button1_Click(object sender, System.EventArgs e)
		{
			#region Validacion
			if (!Validacion.vbComboVacio(this.cmbPeriodo, "Ingrese un período")) return;
			if (!Validacion.vbValidaFechasDesdeHasta(this.dtDesde, this.dtHasta)) return;
			#endregion Validacion

			RegistrosReporteFinal();

			string sReporte = "InformeCuentasContables2.rpt";
			if ((int)this.optReportes.Value == 0)
			{
				Tipo="ACTIVOS";
				TipoReporte = "1.01.";
						
			}
			if ((int)this.optReportes.Value == 3)
			{
				Tipo="PASIVOS Y PATRIMONIO";
				TipoReporte = "2.01.";			

			}
			if ((int)this.optReportes.Value == 1)
			{
				Tipo="ESTADO DE RESULTADOS DEL MES";
				TipoReporte = "3.01.";	
				sReporte = "InformeResultados2.rpt";
				
				if((int)this.optTipoResultado.Value == 1)
				{
					Tipo="ESTADO DE RESULTADOS ACUMULADO";
					TipoReporte = "3.01.";	
					sReporte = "InformeResultadosAcumulado.rpt";
				}
			}

			ParameterFields paramFields = new ParameterFields ();
			ParameterField paramFieldTipo = new ParameterField ();
			ParameterDiscreteValue discreteValTipo = new ParameterDiscreteValue ();
			paramFieldTipo.ParameterFieldName = "@Tipo";
			discreteValTipo.Value = Tipo;
			paramFieldTipo.CurrentValues.Add (discreteValTipo);
			paramFields.Add (paramFieldTipo);

			ParameterField paramFieldReporte = new ParameterField ();
			ParameterDiscreteValue discreteValReporte = new ParameterDiscreteValue ();
			paramFieldReporte.ParameterFieldName = "@reporte";
			discreteValReporte.Value = TipoReporte;
			paramFieldReporte.CurrentValues.Add (discreteValReporte);
			paramFields.Add (paramFieldReporte);

			ParameterField paramFieldestado = new ParameterField ();
			ParameterDiscreteValue discreteValestado = new ParameterDiscreteValue ();
			paramFieldestado.ParameterFieldName = "@estado";
			discreteValestado.Value = 1;
			paramFieldestado.CurrentValues.Add (discreteValestado);
			paramFields.Add (paramFieldestado);

			ParameterField paramFieldHasta = new ParameterField ();
			ParameterDiscreteValue discreteValHasta = new ParameterDiscreteValue ();
			paramFieldHasta.ParameterFieldName = "@Hasta";
			discreteValHasta.Value = Convert.ToDateTime(this.dtHasta.Value);
			paramFieldHasta.CurrentValues.Add (discreteValHasta);
			paramFields.Add (paramFieldHasta);

			ParameterField paramFieldDesde = new ParameterField ();
			ParameterDiscreteValue discreteValDesde  = new ParameterDiscreteValue ();
			paramFieldDesde.ParameterFieldName = "@Desde";
			discreteValDesde.Value = Convert.ToDateTime(this.dtDesde.Value);
			paramFieldDesde.CurrentValues.Add (discreteValDesde);
			paramFields.Add (paramFieldDesde);
			
			
			ParameterField paramFieldEmpresa = new ParameterField ();
			ParameterDiscreteValue discreteValEmpresa = new ParameterDiscreteValue ();
			paramFieldEmpresa.ParameterFieldName = "@idEmpresa";
			discreteValEmpresa.Value = MenuLatinium.stIdDB; 
			paramFieldEmpresa.CurrentValues.Add (discreteValEmpresa);
			paramFields.Add (paramFieldEmpresa);

			Reporte Reporte = new Reporte(sReporte, "");
			Reporte.MdiParent = this.MdiParent;			
			Reporte.crVista.ParameterFieldInfo = paramFields;
			Reporte.Show();
			
			int iCont=0;
			iCont=Funcion.iEscalarSQL(cdsSeteoF, string.Format("EXEC CuentaAuxiliarListaConTipo {0}", (int)this.optReportes.Value));
			this.lbl_registros.Text = iCont.ToString()+ " GENERADOS";
			MessageBox.Show(string.Format("Proceso Finalizado Correctamente, {0} Generados.", iCont), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);

		}

		private void optReportes_ValueChanged(object sender, System.EventArgs e)
		{
			this.optTipoResultado.Enabled = false;
			this.optTipoResultado.CheckedIndex = 0;

			if ((int)optReportes.Value == 1)
			{
				this.optTipoResultado.Enabled = true;
			}

		}

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);

		}

		private void RegistrosReporteFinal()
		{
			string sCuentas = "";

			string sSQLVD = "";
			try
			{		
				string sTipo = "";
				if ((int)this.optReportes.Value == 0)
				{
					sTipo = "1";
						
				}
				if ((int)this.optReportes.Value == 3)
				{
					sTipo = "2";		
				}
				//Resultados
				if ((int)this.optReportes.Value == 1)
				{
					sTipo = "3";		
				}

				string sSQL = string.Format("Exec CuentaAuxiliarCrearReporteFinal '{0}',{1},{2}",
					sTipo,
					(int)this.optTipoResultado.Value,
					idCuentaAuxiliarPeriodoGenerado
					);
				Funcion.EjecutaSQL(cdsSeteoF, sSQL);

			}
			catch (Exception ex)
			{
				MessageBox.Show(string.Format("Commit Exception Type : {0} {1}", ex.GetType(), ex.Message), "Demo", MessageBoxButtons.OK, MessageBoxIcon.Error);					
				this.Cursor = Cursors.Default;							
			}
		}

		private void btnSalir_Click(object sender, System.EventArgs e)
		{
			this.Close();
		}

		private void cmbPeriodo_ValueChanged(object sender, System.EventArgs e)
		{
			idCuentaAuxiliarPeriodoGenerado = (int)this.cmbPeriodo.Value;
			DateTime Desde = Funcion.dtEscalarSQL(cdsSeteoF,  String.Format("EXEC CuentaAuxiliarPeriodoGeneradoModel '{0}',{1}",
				"Desde",
				idCuentaAuxiliarPeriodoGenerado
				));
			DateTime Hasta = Funcion.dtEscalarSQL(cdsSeteoF,  String.Format("EXEC CuentaAuxiliarPeriodoGeneradoModel '{0}',{1}",
				"Hasta",
				idCuentaAuxiliarPeriodoGenerado
				));
			this.dtDesde.Value = Desde;
			this.dtHasta.Value = Hasta;
		}

	}
}
