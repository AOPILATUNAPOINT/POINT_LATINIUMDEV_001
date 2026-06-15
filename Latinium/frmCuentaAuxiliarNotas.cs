using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de frmCuentaAuxiliarNotas.
	/// </summary>
	public class frmCuentaAuxiliarNotas : System.Windows.Forms.Form
	{
		private System.Windows.Forms.GroupBox groupBox2;
		private System.Windows.Forms.Label label17;
		private System.Windows.Forms.Label label18;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtHasta;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtDesde;
		private System.Windows.Forms.GroupBox grpActualizacion;
		private System.Windows.Forms.Label lblDescripcion;
		private System.Windows.Forms.Button btnGrabar;
		private System.Windows.Forms.Button btnCancelar;
		private System.Windows.Forms.Button btnEditar;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource udsNotas;
		private C1.Data.C1DataSet cdsSeteoF;
		private Infragistics.Win.UltraWinGrid.UltraGrid udgNotas;
		/// <summary>
		/// Variable del diseñador requerida.
		int idCuentaAuxiliarNotasCuentaEstadosAB = 0;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtNotas;
		private System.Windows.Forms.Button btnBuscar;
		/// </summary>
		private System.ComponentModel.Container components = null;

		public frmCuentaAuxiliarNotas()
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
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton1 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmCuentaAuxiliarNotas));
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton2 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCuentaAuxiliarNotasCuentaEstadosAB");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Desde");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Hasta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Id_CuentaEstadosABs");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Notas");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("TipoCuentaAuxiliar");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("TipoCuenta");
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance11 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCuentaAuxiliarNotasCuentaEstadosAB");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Desde");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn3 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Hasta");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn4 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Id_CuentaEstadosABs");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn5 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Notas");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn6 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("TipoCuentaAuxiliar");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn7 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("TipoCuenta");
			this.groupBox2 = new System.Windows.Forms.GroupBox();
			this.label17 = new System.Windows.Forms.Label();
			this.label18 = new System.Windows.Forms.Label();
			this.dtHasta = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.dtDesde = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.grpActualizacion = new System.Windows.Forms.GroupBox();
			this.txtNotas = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.lblDescripcion = new System.Windows.Forms.Label();
			this.btnGrabar = new System.Windows.Forms.Button();
			this.btnCancelar = new System.Windows.Forms.Button();
			this.btnEditar = new System.Windows.Forms.Button();
			this.udgNotas = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.udsNotas = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.btnBuscar = new System.Windows.Forms.Button();
			this.groupBox2.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.dtHasta)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtDesde)).BeginInit();
			this.grpActualizacion.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.txtNotas)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.udgNotas)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.udsNotas)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			this.SuspendLayout();
			// 
			// groupBox2
			// 
			this.groupBox2.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.groupBox2.Controls.Add(this.btnBuscar);
			this.groupBox2.Controls.Add(this.label17);
			this.groupBox2.Controls.Add(this.label18);
			this.groupBox2.Controls.Add(this.dtHasta);
			this.groupBox2.Controls.Add(this.dtDesde);
			this.groupBox2.Location = new System.Drawing.Point(16, 16);
			this.groupBox2.Name = "groupBox2";
			this.groupBox2.Size = new System.Drawing.Size(512, 56);
			this.groupBox2.TabIndex = 755;
			this.groupBox2.TabStop = false;
			this.groupBox2.Text = "Período";
			// 
			// label17
			// 
			this.label17.AutoSize = true;
			this.label17.Location = new System.Drawing.Point(200, 16);
			this.label17.Name = "label17";
			this.label17.Size = new System.Drawing.Size(44, 16);
			this.label17.TabIndex = 1051;
			this.label17.Text = "* Hasta:";
			this.label17.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label18
			// 
			this.label18.AutoSize = true;
			this.label18.Location = new System.Drawing.Point(8, 16);
			this.label18.Name = "label18";
			this.label18.Size = new System.Drawing.Size(47, 16);
			this.label18.TabIndex = 1050;
			this.label18.Text = "* Desde:";
			this.label18.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// dtHasta
			// 
			appearance1.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtHasta.Appearance = appearance1;
			this.dtHasta.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton1.Caption = "Today";
			this.dtHasta.DateButtons.Add(dateButton1);
			this.dtHasta.Format = "dd/MM/yyyy";
			this.dtHasta.Location = new System.Drawing.Point(256, 16);
			this.dtHasta.Name = "dtHasta";
			this.dtHasta.NonAutoSizeHeight = 23;
			this.dtHasta.ReadOnly = true;
			this.dtHasta.Size = new System.Drawing.Size(136, 21);
			this.dtHasta.SpinButtonsVisible = true;
			this.dtHasta.TabIndex = 1049;
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
			this.dtDesde.Location = new System.Drawing.Point(56, 16);
			this.dtDesde.Name = "dtDesde";
			this.dtDesde.NonAutoSizeHeight = 23;
			this.dtDesde.ReadOnly = true;
			this.dtDesde.Size = new System.Drawing.Size(128, 21);
			this.dtDesde.SpinButtonsVisible = true;
			this.dtDesde.TabIndex = 1048;
			this.dtDesde.Value = ((object)(resources.GetObject("dtDesde.Value")));
			// 
			// grpActualizacion
			// 
			this.grpActualizacion.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.grpActualizacion.Controls.Add(this.txtNotas);
			this.grpActualizacion.Controls.Add(this.lblDescripcion);
			this.grpActualizacion.Controls.Add(this.btnGrabar);
			this.grpActualizacion.Controls.Add(this.btnCancelar);
			this.grpActualizacion.Controls.Add(this.btnEditar);
			this.grpActualizacion.Location = new System.Drawing.Point(16, 80);
			this.grpActualizacion.Name = "grpActualizacion";
			this.grpActualizacion.Size = new System.Drawing.Size(512, 96);
			this.grpActualizacion.TabIndex = 756;
			this.grpActualizacion.TabStop = false;
			this.grpActualizacion.Text = "Mantenimiento de notas";
			// 
			// txtNotas
			// 
			appearance3.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtNotas.Appearance = appearance3;
			this.txtNotas.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNotas.Enabled = false;
			this.txtNotas.Location = new System.Drawing.Point(104, 24);
			this.txtNotas.Name = "txtNotas";
			this.txtNotas.Size = new System.Drawing.Size(240, 21);
			this.txtNotas.TabIndex = 1039;
			// 
			// lblDescripcion
			// 
			this.lblDescripcion.Location = new System.Drawing.Point(8, 24);
			this.lblDescripcion.Name = "lblDescripcion";
			this.lblDescripcion.TabIndex = 751;
			this.lblDescripcion.Text = "* Notas:";
			// 
			// btnGrabar
			// 
			this.btnGrabar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnGrabar.Enabled = false;
			this.btnGrabar.Image = ((System.Drawing.Image)(resources.GetObject("btnGrabar.Image")));
			this.btnGrabar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnGrabar.Location = new System.Drawing.Point(104, 64);
			this.btnGrabar.Name = "btnGrabar";
			this.btnGrabar.Size = new System.Drawing.Size(76, 23);
			this.btnGrabar.TabIndex = 749;
			this.btnGrabar.Text = "&Guardar";
			this.btnGrabar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnGrabar.Click += new System.EventHandler(this.btnGrabar_Click);
			// 
			// btnCancelar
			// 
			this.btnCancelar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnCancelar.CausesValidation = false;
			this.btnCancelar.Enabled = false;
			this.btnCancelar.Image = ((System.Drawing.Image)(resources.GetObject("btnCancelar.Image")));
			this.btnCancelar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnCancelar.Location = new System.Drawing.Point(192, 64);
			this.btnCancelar.Name = "btnCancelar";
			this.btnCancelar.Size = new System.Drawing.Size(72, 23);
			this.btnCancelar.TabIndex = 750;
			this.btnCancelar.Text = "&Cancelar";
			this.btnCancelar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnCancelar.Click += new System.EventHandler(this.btnCancelar_Click);
			// 
			// btnEditar
			// 
			this.btnEditar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnEditar.CausesValidation = false;
			this.btnEditar.Enabled = false;
			this.btnEditar.Image = ((System.Drawing.Image)(resources.GetObject("btnEditar.Image")));
			this.btnEditar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnEditar.Location = new System.Drawing.Point(16, 64);
			this.btnEditar.Name = "btnEditar";
			this.btnEditar.Size = new System.Drawing.Size(76, 23);
			this.btnEditar.TabIndex = 1036;
			this.btnEditar.Text = "&Editar";
			this.btnEditar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnEditar.Click += new System.EventHandler(this.btnEditar_Click);
			// 
			// udgNotas
			// 
			this.udgNotas.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.udgNotas.Cursor = System.Windows.Forms.Cursors.Default;
			this.udgNotas.DataSource = this.udsNotas;
			appearance4.BackColor = System.Drawing.Color.White;
			this.udgNotas.DisplayLayout.Appearance = appearance4;
			this.udgNotas.DisplayLayout.AutoFitColumns = true;
			ultraGridColumn1.Header.Caption = "idCuentaAuxiliarNotas";
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn1.Width = 190;
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridColumn2.Hidden = true;
			ultraGridColumn3.Header.VisiblePosition = 2;
			ultraGridColumn3.Hidden = true;
			ultraGridColumn4.Header.VisiblePosition = 3;
			ultraGridColumn4.Hidden = true;
			ultraGridColumn5.Header.VisiblePosition = 6;
			ultraGridColumn5.Width = 182;
			ultraGridColumn6.Header.Caption = "Cuenta";
			ultraGridColumn6.Header.VisiblePosition = 4;
			ultraGridColumn6.Width = 226;
			ultraGridColumn7.Header.Caption = "Tipo";
			ultraGridColumn7.Header.VisiblePosition = 5;
			ultraGridColumn7.Width = 83;
			ultraGridBand1.Columns.AddRange(new object[] {
																										 ultraGridColumn1,
																										 ultraGridColumn2,
																										 ultraGridColumn3,
																										 ultraGridColumn4,
																										 ultraGridColumn5,
																										 ultraGridColumn6,
																										 ultraGridColumn7});
			ultraGridBand1.Header.Caption = "Detalle";
			ultraGridBand1.HeaderVisible = true;
			this.udgNotas.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.udgNotas.DisplayLayout.GroupByBox.Hidden = true;
			appearance5.ForeColor = System.Drawing.Color.Black;
			appearance5.ForeColorDisabled = System.Drawing.Color.Black;
			this.udgNotas.DisplayLayout.Override.ActiveRowAppearance = appearance5;
			appearance6.BackColor = System.Drawing.Color.Transparent;
			this.udgNotas.DisplayLayout.Override.CardAreaAppearance = appearance6;
			appearance7.ForeColor = System.Drawing.Color.Black;
			appearance7.ForeColorDisabled = System.Drawing.Color.Black;
			this.udgNotas.DisplayLayout.Override.GroupByRowAppearance = appearance7;
			appearance8.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance8.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance8.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance8.FontData.BoldAsString = "True";
			appearance8.FontData.Name = "Arial";
			appearance8.FontData.SizeInPoints = 8.5F;
			appearance8.ForeColor = System.Drawing.Color.White;
			appearance8.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.udgNotas.DisplayLayout.Override.HeaderAppearance = appearance8;
			appearance9.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance9.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance9.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.udgNotas.DisplayLayout.Override.RowAlternateAppearance = appearance9;
			appearance10.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance10.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance10.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.udgNotas.DisplayLayout.Override.RowSelectorAppearance = appearance10;
			appearance11.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance11.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance11.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.udgNotas.DisplayLayout.Override.SelectedRowAppearance = appearance11;
			this.udgNotas.DisplayLayout.ViewStyleBand = Infragistics.Win.UltraWinGrid.ViewStyleBand.OutlookGroupBy;
			this.udgNotas.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.udgNotas.Location = new System.Drawing.Point(16, 192);
			this.udgNotas.Name = "udgNotas";
			this.udgNotas.Size = new System.Drawing.Size(512, 216);
			this.udgNotas.TabIndex = 757;
			this.udgNotas.Click += new System.EventHandler(this.udgNotas_Click);
			// 
			// udsNotas
			// 
			ultraDataColumn1.DataType = typeof(int);
			ultraDataColumn1.ReadOnly = Infragistics.Win.DefaultableBoolean.False;
			ultraDataColumn2.DataType = typeof(System.DateTime);
			ultraDataColumn2.ReadOnly = Infragistics.Win.DefaultableBoolean.False;
			ultraDataColumn3.DataType = typeof(System.DateTime);
			ultraDataColumn4.DataType = typeof(System.UInt32);
			this.udsNotas.Band.Columns.AddRange(new object[] {
																												 ultraDataColumn1,
																												 ultraDataColumn2,
																												 ultraDataColumn3,
																												 ultraDataColumn4,
																												 ultraDataColumn5,
																												 ultraDataColumn6,
																												 ultraDataColumn7});
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
			// 
			// btnBuscar
			// 
			this.btnBuscar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnBuscar.CausesValidation = false;
			this.btnBuscar.Image = ((System.Drawing.Image)(resources.GetObject("btnBuscar.Image")));
			this.btnBuscar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnBuscar.Location = new System.Drawing.Point(424, 16);
			this.btnBuscar.Name = "btnBuscar";
			this.btnBuscar.Size = new System.Drawing.Size(78, 23);
			this.btnBuscar.TabIndex = 1052;
			this.btnBuscar.Text = "&Consultar";
			this.btnBuscar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnBuscar.Click += new System.EventHandler(this.btnBuscar_Click);
			// 
			// frmCuentaAuxiliarNotas
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(552, 454);
			this.Controls.Add(this.udgNotas);
			this.Controls.Add(this.grpActualizacion);
			this.Controls.Add(this.groupBox2);
			this.Name = "frmCuentaAuxiliarNotas";
			this.Text = "Configuración de notas";
			this.Load += new System.EventHandler(this.frmCuentaAuxiliarNotas_Load);
			this.groupBox2.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.dtHasta)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtDesde)).EndInit();
			this.grpActualizacion.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.txtNotas)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.udgNotas)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.udsNotas)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void frmCuentaAuxiliarNotas_Load(object sender, System.EventArgs e)
		{
			#region Fechas por defecto
			this.dtDesde.Value = "01/" + DateTime.Now.Month.ToString() + "/" + DateTime.Now.Year.ToString();
			this.dtHasta.Value = DateTime.DaysInMonth(DateTime.Now.Year, DateTime.Now.Month).ToString() + "/" + DateTime.Now.Month.ToString() + "/" + DateTime.Now.Year.ToString();
			DateTime dtFecha = FuncionesProcedimientos.dtRetornaFechaServidor(cdsSeteoF);
			this.dtDesde.CalendarInfo.MinDate = DateTime.Parse("01/01/2014");
			this.dtHasta.CalendarInfo.MinDate = DateTime.Parse("01/01/2014");
			this.dtDesde.CalendarInfo.MaxDate = dtFecha;
			this.dtHasta.CalendarInfo.MaxDate = dtFecha;
			this.dtDesde.Value = frmCuentaSubPlan.Desde;
			this.dtHasta.Value = frmCuentaSubPlan.Hasta;
			#endregion Fechas por defecto
			//En listar
			CuentaAuxiliarNotasCon();
		}

		public void CuentaAuxiliarNotasCon()
		{
				this.udgNotas.DataSource = Funcion.dvProcedimiento(cdsSeteoF, String.Format("EXEC CuentaAuxiliarNotasCon {0}",
				frmCuentaSubPlan.idCuentaAuxiliarPeriodoGenerado)
				);
		}


//		public void CuentaAuxiliarNotasCon()
//		{
//			this.udgNotas.DataSource = 
//		}

		private void btnGenerar_Click(object sender, System.EventArgs e)
		{
		
		}

		private void udgNotas_Click(object sender, System.EventArgs e)
		{
			 idCuentaAuxiliarNotasCuentaEstadosAB = 0;
			 idCuentaAuxiliarNotasCuentaEstadosAB =	(int)this.udgNotas.ActiveRow.Cells["idCuentaAuxiliarNotasCuentaEstadosAB"].Value;
			 this.txtNotas.Text =	this.udgNotas.ActiveRow.Cells["Notas"].Value.ToString();

			 this.btnEditar.Enabled = true;
			 this.btnCancelar.Enabled = true;
			
		}

		private void btnGrabar_Click(object sender, System.EventArgs e)
		{
			bool bFunciones = Funcion.bEscalarSQL(cdsSeteoF, String.Format("EXEC CuentaAuxiliarNotasActualizar {0}, '{1}'",
				idCuentaAuxiliarNotasCuentaEstadosAB,
				this.txtNotas.Text), false
				);
			if(bFunciones)
			{
				idCuentaAuxiliarNotasCuentaEstadosAB = 0;
				this.btnEditar.Enabled = false;
				this.btnCancelar.Enabled = false;
				this.txtNotas.Text = "";
				this.txtNotas.Enabled = false;
				this.btnGrabar.Enabled = false;
			}
			//En listar
			CuentaAuxiliarNotasCon();
		}

		private void btnCancelar_Click(object sender, System.EventArgs e)
		{
			idCuentaAuxiliarNotasCuentaEstadosAB = 0;
			this.btnEditar.Enabled = false;
			this.btnCancelar.Enabled = false;
			this.txtNotas.Text = "";
			this.txtNotas.Enabled = false;
		}

		private void btnEditar_Click(object sender, System.EventArgs e)
		{
			this.btnGrabar.Enabled = true;
			this.txtNotas.Enabled = true;
		}

		private void btnBuscar_Click(object sender, System.EventArgs e)
		{
			//En listar
			CuentaAuxiliarNotasCon();
		}
	}
}
