using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de EstadoCuentaAntes.
	/// </summary>
	public class EstadoCuentaAntes : DevExpress.XtraEditors.XtraForm
	{
		private System.Windows.Forms.Label label1;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtDebe;
		private C1.Data.C1DataSet cdsConfirmado;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtHaber;
		private System.Windows.Forms.Label label2;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkConfirmado;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtDetalle;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtCheque;
		private System.Windows.Forms.Label label3;
		private System.Windows.Forms.Label label4;
		private System.Windows.Forms.Label label5;
		private System.Windows.Forms.Label label6;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtFConfirmacion;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtFechaGenerado;
		private C1.Data.C1DataSet cdsCuentaCorr;
		private System.Windows.Forms.Label label7;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbCuentaCorr;
		private System.Windows.Forms.Label label8;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbProyecto;
		private C1.Data.C1DataSet cdsProyecto;
		private BarraDatoSQL.BarraDatoSQL barraDato1;
		private System.Windows.Forms.ErrorProvider errorProvider1;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public EstadoCuentaAntes()
		{
			InitializeComponent();
		}
		int IdCuenta = 0;
		public EstadoCuentaAntes(int idCuenta)
		{
			InitializeComponent();
			IdCuenta = idCuenta;
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

		#region Windows Form Designer generated code
		/// <summary>
		/// Método necesario para admitir el Diseñador, no se puede modificar
		/// el contenido del método con el editor de código.
		/// </summary>
		private void InitializeComponent()
		{
			Infragistics.Win.Appearance appearance1 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton1 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton2 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("CuentaCorriente", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCuentaCorriente");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idSucursal");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idBanco");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCuenta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("NumCuenta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("NumCheque");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Agencia");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Reporte");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Automatico");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Borrar");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn11 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Asiento");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn12 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCuentaTipo");
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Proyecto", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn13 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idProyecto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn14 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPadre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn15 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn16 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Lugar");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn17 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Responsable");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn18 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Notas");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn19 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CentroCosto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn20 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaIngreso");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn21 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Visible");
			this.label1 = new System.Windows.Forms.Label();
			this.txtDebe = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.cdsConfirmado = new C1.Data.C1DataSet();
			this.txtHaber = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label2 = new System.Windows.Forms.Label();
			this.chkConfirmado = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.txtDetalle = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtCheque = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label3 = new System.Windows.Forms.Label();
			this.label4 = new System.Windows.Forms.Label();
			this.label5 = new System.Windows.Forms.Label();
			this.label6 = new System.Windows.Forms.Label();
			this.dtFConfirmacion = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.dtFechaGenerado = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.cmbCuentaCorr = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.cdsCuentaCorr = new C1.Data.C1DataSet();
			this.label7 = new System.Windows.Forms.Label();
			this.label8 = new System.Windows.Forms.Label();
			this.cmbProyecto = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.cdsProyecto = new C1.Data.C1DataSet();
			this.barraDato1 = new BarraDatoSQL.BarraDatoSQL();
			this.errorProvider1 = new System.Windows.Forms.ErrorProvider();
			((System.ComponentModel.ISupportInitialize)(this.txtDebe)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsConfirmado)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtHaber)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDetalle)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCheque)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFConfirmacion)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFechaGenerado)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCuentaCorr)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsCuentaCorr)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbProyecto)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsProyecto)).BeginInit();
			this.SuspendLayout();
			// 
			// label1
			// 
			this.label1.AutoSize = true;
			this.label1.Location = new System.Drawing.Point(24, 16);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(34, 16);
			this.label1.TabIndex = 0;
			this.label1.Text = "Debe:";
			// 
			// txtDebe
			// 
			appearance1.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtDebe.Appearance = appearance1;
			this.txtDebe.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsConfirmado, "Conciliacion.Debe"));
			this.txtDebe.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtDebe.FormatProvider = new System.Globalization.CultureInfo("en-US");
			this.txtDebe.FormatString = "#,##0.00";
			this.txtDebe.Location = new System.Drawing.Point(88, 16);
			this.txtDebe.Name = "txtDebe";
			this.txtDebe.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtDebe.PromptChar = ' ';
			this.txtDebe.Size = new System.Drawing.Size(120, 21);
			this.txtDebe.TabIndex = 1;
			// 
			// cdsConfirmado
			// 
			this.cdsConfirmado.BindingContextCtrl = this;
			this.cdsConfirmado.DataLibrary = "LibContabilidad";
			this.cdsConfirmado.DataLibraryUrl = "";
			this.cdsConfirmado.DataSetDef = "dsConciliacion";
			this.cdsConfirmado.EnforceConstraints = false;
			this.cdsConfirmado.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsConfirmado.Name = "cdsConfirmado";
			this.cdsConfirmado.SchemaClassName = "LibContabilidad.DataClass";
			this.cdsConfirmado.SchemaDef = null;
			this.cdsConfirmado.PositionChanged += new C1.Data.PositionChangeEventHandler(this.cdsConfirmado_PositionChanged);
			this.cdsConfirmado.BeforeFill += new C1.Data.FillEventHandler(this.cdsConfirmado_BeforeFill);
			// 
			// txtHaber
			// 
			appearance2.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtHaber.Appearance = appearance2;
			this.txtHaber.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsConfirmado, "Conciliacion.Haber"));
			this.txtHaber.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtHaber.FormatProvider = new System.Globalization.CultureInfo("en-US");
			this.txtHaber.FormatString = "#,##0.00";
			this.txtHaber.Location = new System.Drawing.Point(312, 16);
			this.txtHaber.Name = "txtHaber";
			this.txtHaber.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtHaber.PromptChar = ' ';
			this.txtHaber.Size = new System.Drawing.Size(120, 21);
			this.txtHaber.TabIndex = 3;
			// 
			// label2
			// 
			this.label2.AutoSize = true;
			this.label2.Location = new System.Drawing.Point(256, 16);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(38, 16);
			this.label2.TabIndex = 2;
			this.label2.Text = "Haber:";
			// 
			// chkConfirmado
			// 
			this.chkConfirmado.BackColor = System.Drawing.Color.Transparent;
			this.chkConfirmado.ButtonStyle = Infragistics.Win.UIElementButtonStyle.WindowsXPCommandButton;
			this.chkConfirmado.DataBindings.Add(new System.Windows.Forms.Binding("CheckedValue", this.cdsConfirmado, "Conciliacion.Anterior"));
			this.chkConfirmado.Location = new System.Drawing.Point(88, 80);
			this.chkConfirmado.Name = "chkConfirmado";
			this.chkConfirmado.Size = new System.Drawing.Size(88, 20);
			this.chkConfirmado.TabIndex = 4;
			this.chkConfirmado.Text = "Confirmado";
			// 
			// txtDetalle
			// 
			appearance3.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtDetalle.Appearance = appearance3;
			this.txtDetalle.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsConfirmado, "Conciliacion.Detalle"));
			this.txtDetalle.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtDetalle.Location = new System.Drawing.Point(88, 48);
			this.txtDetalle.Name = "txtDetalle";
			this.txtDetalle.Size = new System.Drawing.Size(344, 21);
			this.txtDetalle.TabIndex = 5;
			// 
			// txtCheque
			// 
			appearance4.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtCheque.Appearance = appearance4;
			this.txtCheque.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsConfirmado, "Conciliacion.Cheque"));
			this.txtCheque.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtCheque.Location = new System.Drawing.Point(88, 112);
			this.txtCheque.Name = "txtCheque";
			this.txtCheque.Size = new System.Drawing.Size(120, 21);
			this.txtCheque.TabIndex = 6;
			// 
			// label3
			// 
			this.label3.AutoSize = true;
			this.label3.Location = new System.Drawing.Point(24, 112);
			this.label3.Name = "label3";
			this.label3.Size = new System.Drawing.Size(47, 16);
			this.label3.TabIndex = 8;
			this.label3.Text = "Cheque:";
			// 
			// label4
			// 
			this.label4.AutoSize = true;
			this.label4.Location = new System.Drawing.Point(24, 48);
			this.label4.Name = "label4";
			this.label4.Size = new System.Drawing.Size(43, 16);
			this.label4.TabIndex = 7;
			this.label4.Text = "Detalle:";
			// 
			// label5
			// 
			this.label5.AutoSize = true;
			this.label5.Location = new System.Drawing.Point(216, 80);
			this.label5.Name = "label5";
			this.label5.Size = new System.Drawing.Size(106, 16);
			this.label5.TabIndex = 9;
			this.label5.Text = "Fecha Confirmación";
			// 
			// label6
			// 
			this.label6.AutoSize = true;
			this.label6.Location = new System.Drawing.Point(216, 112);
			this.label6.Name = "label6";
			this.label6.Size = new System.Drawing.Size(97, 16);
			this.label6.TabIndex = 11;
			this.label6.Text = "Fecha Generación";
			// 
			// dtFConfirmacion
			// 
			appearance5.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtFConfirmacion.Appearance = appearance5;
			this.dtFConfirmacion.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.dtFConfirmacion.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsConfirmado, "Conciliacion.FechaConfirmacion"));
			dateButton1.Caption = "Today";
			this.dtFConfirmacion.DateButtons.Add(dateButton1);
			this.dtFConfirmacion.Format = "dd/MMM/yyyy";
			this.dtFConfirmacion.Location = new System.Drawing.Point(328, 80);
			this.dtFConfirmacion.Name = "dtFConfirmacion";
			this.dtFConfirmacion.NonAutoSizeHeight = 23;
			this.dtFConfirmacion.Size = new System.Drawing.Size(104, 21);
			this.dtFConfirmacion.TabIndex = 14;
			this.dtFConfirmacion.Value = new System.DateTime(2006, 9, 4, 0, 0, 0, 0);
			// 
			// dtFechaGenerado
			// 
			appearance6.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtFechaGenerado.Appearance = appearance6;
			this.dtFechaGenerado.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.dtFechaGenerado.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsConfirmado, "Conciliacion.FechaGeneracion"));
			dateButton2.Caption = "Today";
			this.dtFechaGenerado.DateButtons.Add(dateButton2);
			this.dtFechaGenerado.Format = "dd/MMM/yyyy";
			this.dtFechaGenerado.Location = new System.Drawing.Point(328, 112);
			this.dtFechaGenerado.Name = "dtFechaGenerado";
			this.dtFechaGenerado.NonAutoSizeHeight = 23;
			this.dtFechaGenerado.Size = new System.Drawing.Size(104, 21);
			this.dtFechaGenerado.TabIndex = 15;
			this.dtFechaGenerado.Value = new System.DateTime(2006, 9, 4, 0, 0, 0, 0);
			// 
			// cmbCuentaCorr
			// 
			appearance7.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbCuentaCorr.Appearance = appearance7;
			this.cmbCuentaCorr.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbCuentaCorr.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbCuentaCorr.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsConfirmado, "Conciliacion.idCuenta"));
			this.cmbCuentaCorr.DataMember = "CuentaCorriente";
			this.cmbCuentaCorr.DataSource = this.cdsCuentaCorr;
			this.cmbCuentaCorr.DisplayLayout.AutoFitColumns = true;
			ultraGridColumn1.Header.VisiblePosition = 6;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn2.Header.VisiblePosition = 9;
			ultraGridColumn2.Hidden = true;
			ultraGridColumn3.Header.VisiblePosition = 4;
			ultraGridColumn3.Hidden = true;
			ultraGridColumn4.Header.VisiblePosition = 5;
			ultraGridColumn4.Hidden = true;
			ultraGridColumn5.Header.VisiblePosition = 0;
			ultraGridColumn5.Width = 215;
			ultraGridColumn6.Header.VisiblePosition = 7;
			ultraGridColumn6.Hidden = true;
			ultraGridColumn7.Header.VisiblePosition = 1;
			ultraGridColumn7.Hidden = true;
			ultraGridColumn8.Header.VisiblePosition = 8;
			ultraGridColumn8.Hidden = true;
			ultraGridColumn9.Header.VisiblePosition = 2;
			ultraGridColumn9.Hidden = true;
			ultraGridColumn10.Header.VisiblePosition = 3;
			ultraGridColumn10.Hidden = true;
			ultraGridColumn11.Header.VisiblePosition = 10;
			ultraGridColumn11.Hidden = true;
			ultraGridColumn11.Width = 29;
			ultraGridColumn12.Header.VisiblePosition = 11;
			ultraGridColumn12.Hidden = true;
			ultraGridColumn12.Width = 71;
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
																										 ultraGridColumn12});
			this.cmbCuentaCorr.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.cmbCuentaCorr.DisplayMember = "NumCuenta";
			this.cmbCuentaCorr.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbCuentaCorr.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbCuentaCorr.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbCuentaCorr.Location = new System.Drawing.Point(128, 144);
			this.cmbCuentaCorr.Name = "cmbCuentaCorr";
			this.cmbCuentaCorr.Size = new System.Drawing.Size(304, 21);
			this.cmbCuentaCorr.TabIndex = 16;
			this.cmbCuentaCorr.ValueMember = "idCuenta";
			// 
			// cdsCuentaCorr
			// 
			this.cdsCuentaCorr.BindingContextCtrl = this;
			this.cdsCuentaCorr.DataLibrary = "LibContabilidad";
			this.cdsCuentaCorr.DataLibraryUrl = "";
			this.cdsCuentaCorr.DataSetDef = "dsCuentaCorriente";
			this.cdsCuentaCorr.EnforceConstraints = false;
			this.cdsCuentaCorr.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsCuentaCorr.Name = "cdsCuentaCorr";
			this.cdsCuentaCorr.SchemaClassName = "LibContabilidad.DataClass";
			this.cdsCuentaCorr.SchemaDef = null;
			this.cdsCuentaCorr.BeforeFill += new C1.Data.FillEventHandler(this.cdsCuentaCorr_BeforeFill);
			// 
			// label7
			// 
			this.label7.AutoSize = true;
			this.label7.Location = new System.Drawing.Point(24, 144);
			this.label7.Name = "label7";
			this.label7.Size = new System.Drawing.Size(90, 16);
			this.label7.TabIndex = 17;
			this.label7.Text = "Cuenta Corriente";
			// 
			// label8
			// 
			this.label8.AutoSize = true;
			this.label8.Location = new System.Drawing.Point(24, 176);
			this.label8.Name = "label8";
			this.label8.Size = new System.Drawing.Size(86, 16);
			this.label8.TabIndex = 19;
			this.label8.Text = "Centro de Costo";
			// 
			// cmbProyecto
			// 
			appearance8.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbProyecto.Appearance = appearance8;
			this.cmbProyecto.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbProyecto.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbProyecto.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsConfirmado, "Conciliacion.idProyecto"));
			this.cmbProyecto.DataMember = "Proyecto";
			this.cmbProyecto.DataSource = this.cdsProyecto;
			this.cmbProyecto.DisplayLayout.AutoFitColumns = true;
			ultraGridBand2.AddButtonCaption = "CuentaCorriente";
			ultraGridColumn13.Header.VisiblePosition = 1;
			ultraGridColumn13.Hidden = true;
			ultraGridColumn14.Header.VisiblePosition = 6;
			ultraGridColumn14.Hidden = true;
			ultraGridColumn15.Header.VisiblePosition = 0;
			ultraGridColumn15.Width = 286;
			ultraGridColumn16.Header.VisiblePosition = 2;
			ultraGridColumn16.Hidden = true;
			ultraGridColumn17.Header.VisiblePosition = 7;
			ultraGridColumn17.Hidden = true;
			ultraGridColumn18.Header.VisiblePosition = 3;
			ultraGridColumn18.Hidden = true;
			ultraGridColumn19.Header.VisiblePosition = 4;
			ultraGridColumn19.Hidden = true;
			ultraGridColumn20.Header.VisiblePosition = 5;
			ultraGridColumn20.Hidden = true;
			ultraGridColumn21.Header.VisiblePosition = 8;
			ultraGridColumn21.Hidden = true;
			ultraGridColumn21.Width = 27;
			ultraGridBand2.Columns.AddRange(new object[] {
																										 ultraGridColumn13,
																										 ultraGridColumn14,
																										 ultraGridColumn15,
																										 ultraGridColumn16,
																										 ultraGridColumn17,
																										 ultraGridColumn18,
																										 ultraGridColumn19,
																										 ultraGridColumn20,
																										 ultraGridColumn21});
			this.cmbProyecto.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			this.cmbProyecto.DisplayMember = "Nombre";
			this.cmbProyecto.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbProyecto.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbProyecto.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbProyecto.Location = new System.Drawing.Point(128, 176);
			this.cmbProyecto.Name = "cmbProyecto";
			this.cmbProyecto.Size = new System.Drawing.Size(304, 21);
			this.cmbProyecto.TabIndex = 18;
			this.cmbProyecto.ValueMember = "idProyecto";
			// 
			// cdsProyecto
			// 
			this.cdsProyecto.BindingContextCtrl = this;
			this.cdsProyecto.DataLibrary = "LibContabilidad";
			this.cdsProyecto.DataLibraryUrl = "";
			this.cdsProyecto.DataSetDef = "dsProyecto";
			this.cdsProyecto.EnforceConstraints = false;
			this.cdsProyecto.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsProyecto.Name = "cdsProyecto";
			this.cdsProyecto.SchemaClassName = "LibContabilidad.DataClass";
			this.cdsProyecto.SchemaDef = null;
			this.cdsProyecto.BeforeFill += new C1.Data.FillEventHandler(this.cdsProyecto_BeforeFill);
			// 
			// barraDato1
			// 
			this.barraDato1.BarraMovimiento = true;
			this.barraDato1.DataMember = "Conciliacion";
			this.barraDato1.DataNombreId = "idConciliacion";
			this.barraDato1.DataOrden = "";
			this.barraDato1.DataSource = this.cdsConfirmado;
			this.barraDato1.DataTabla = "Conciliacion";
			this.barraDato1.DataTablaHija = "";
			this.barraDato1.Location = new System.Drawing.Point(88, 224);
			this.barraDato1.Name = "barraDato1";
			this.barraDato1.Size = new System.Drawing.Size(256, 24);
			this.barraDato1.TabIndex = 20;
			this.barraDato1.VisibleBorrar = true;
			this.barraDato1.VisibleBuscar = false;
			this.barraDato1.VisibleEditar = true;
			this.barraDato1.VisibleImprimir = false;
			this.barraDato1.VisibleNuevo = true;
			this.barraDato1.Borrar += new System.EventHandler(this.barraDato1_Borrar);
			this.barraDato1.Refresca += new System.EventHandler(this.barraDatoSQL1_Refresca);
			this.barraDato1.Crear += new System.EventHandler(this.barraDato1_Crear);
			this.barraDato1.Deshacer += new System.EventHandler(this.barraDato1_Deshacer);
			this.barraDato1.Editar += new System.EventHandler(this.barraDato1_Editar);
			this.barraDato1.Grabar += new System.EventHandler(this.barraDato1_Grabar);
			// 
			// errorProvider1
			// 
			this.errorProvider1.ContainerControl = this;
			// 
			// EstadoCuentaAntes
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(464, 261);
			this.Controls.Add(this.barraDato1);
			this.Controls.Add(this.label8);
			this.Controls.Add(this.label7);
			this.Controls.Add(this.label6);
			this.Controls.Add(this.label5);
			this.Controls.Add(this.label3);
			this.Controls.Add(this.label4);
			this.Controls.Add(this.label2);
			this.Controls.Add(this.label1);
			this.Controls.Add(this.cmbProyecto);
			this.Controls.Add(this.cmbCuentaCorr);
			this.Controls.Add(this.dtFechaGenerado);
			this.Controls.Add(this.dtFConfirmacion);
			this.Controls.Add(this.txtCheque);
			this.Controls.Add(this.txtDetalle);
			this.Controls.Add(this.chkConfirmado);
			this.Controls.Add(this.txtHaber);
			this.Controls.Add(this.txtDebe);
			this.MaximizeBox = false;
			this.Name = "EstadoCuentaAntes";
			this.Text = "Cheques del Año pasado";
			this.Closing += new System.ComponentModel.CancelEventHandler(this.EstadoCuentaAntes_Closing);
			this.Load += new System.EventHandler(this.EstadoCuentaAntes_Load);
			((System.ComponentModel.ISupportInitialize)(this.txtDebe)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsConfirmado)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtHaber)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDetalle)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCheque)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFConfirmacion)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFechaGenerado)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCuentaCorr)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsCuentaCorr)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbProyecto)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsProyecto)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void cdsConfirmado_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			this.cdsConfirmado.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirConta);
		}

		private void barraDato1_Borrar(object sender, System.EventArgs e)
		{
			if (MessageBox.Show("¿Desea Borrar este registro?",
				"Confirmación", MessageBoxButtons.YesNo, 
				MessageBoxIcon.Question, 
				MessageBoxDefaultButton.Button2) == DialogResult.No) return;
			this.barraDato1.BorraRegistro();
		}

		private void barraDato1_Crear(object sender, System.EventArgs e)
		{
			this.barraDato1.CrearRegistro();
			this.drConciliacion["FechaConfirmacion"] = DateTime.Today;
			this.drConciliacion["FechaGeneracion"] = DateTime.Today;
		}

		private void barraDato1_Deshacer(object sender, System.EventArgs e)
		{
			this.barraDato1.DeshacerRegistro();
		}

		private void barraDato1_Editar(object sender, System.EventArgs e)
		{
			this.barraDato1.EditarRegistro();
		}

		private void EstadoCuentaAntes_Load(object sender, System.EventArgs e)
		{
			string stAudita = "Exec AuditaCrear 53, 6, 'Cheques Anteriores'";
			Funcion.EjecutaSQL(cdsConfirmado, stAudita, true);
			Cursor = Cursors.Default;
			barraDato1.stFiltroExt = " And Borrar = 0 ";
			if (IdCuenta > 0)
				barraDato1.stFiltroExt += " And idCuenta = " + IdCuenta.ToString();
			this.barraDato1.PosUltima();
			this.barraDato1.HabilitaControles(false);
		}

		private bool Verifica()
		{
			bool bOk = true;
			errorProvider1.SetError(cmbCuentaCorr, "");
			errorProvider1.SetError(txtCheque, "");
			errorProvider1.SetError(txtDetalle, "");
			if (cmbCuentaCorr.Text.Trim().Length == 0)
			{
        errorProvider1.SetError(cmbCuentaCorr, "Ingrese cuenta corriente");
				bOk = false;
			}
			if (txtCheque.Text.Trim().Length == 0)
			{
				errorProvider1.SetError(txtCheque, "Ingrese cheque");
				bOk = false;
			}
			if (txtDetalle.Text.Trim().Length == 0)
			{
				errorProvider1.SetError(txtDetalle, "Ingrese detalle");
				bOk = false;
			}
			return bOk;
		}

		private void barraDato1_Grabar(object sender, System.EventArgs e)
		{
			if (!Verifica()) return;
			this.barraDato1.GrabaRegistro();
		}

		C1.Data.C1DataRow drConciliacion;
		private void cdsConfirmado_PositionChanged(object sender, C1.Data.PositionChangeEventArgs e)
		{
			drConciliacion = e.Row;
		}

		private void cdsProyecto_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			this.cdsProyecto.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirConta);
		}

		private void cdsCuentaCorr_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			this.cdsCuentaCorr.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirConta);		
		}

		private void barraDatoSQL1_Refresca(object sender, System.EventArgs e)
		{
		
		}

		private void EstadoCuentaAntes_Closing(object sender, System.ComponentModel.CancelEventArgs e)
		{
			string stAudita = "Exec AuditaCrear 53, 6, 'Cheques Anteriores'";
			Funcion.EjecutaSQL(cdsConfirmado, stAudita, true);
		}
	}
}
