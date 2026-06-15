using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Text;
using CrystalDecisions;
using CrystalDecisions.CrystalReports.Engine;
using CrystalDecisions.ReportSource;
using CrystalDecisions.Shared;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de frmRolAcreditacion.
	/// </summary>
	public class frmRolAcreditacion : System.Windows.Forms.Form
	{
		private C1.Data.C1DataSet cdsSeteoF;
		private System.Windows.Forms.Label lblCuenta;
		public Infragistics.Win.UltraWinGrid.UltraCombo cmbCuenta;
		private System.Windows.Forms.Label label10;
		private System.Windows.Forms.Button btnCancelar;
		private System.Windows.Forms.GroupBox groupBox1;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtpFecha;
		private System.Windows.Forms.Button btnAcreditar;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource2;
		private Infragistics.Win.UltraWinGrid.UltraGrid uGridLoja;
		private Infragistics.Win.UltraWinGrid.ExcelExport.UltraGridExcelExporter ultraGridExcelExporter1;
		private System.Windows.Forms.SaveFileDialog saveFileDialog1;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource1;
		private System.Windows.Forms.Button btnGenerarArchivo;
		private System.Windows.Forms.Label lblMensaje;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtIdAsiento;
		private System.Windows.Forms.Label lblEstado;
		public System.Windows.Forms.Button btnImprimir;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public frmRolAcreditacion()
		{
			//
			// Necesario para admitir el Diseñador de Windows Forms
			//
			InitializeComponent();

			//
			// TODO: agregar código de constructor después de llamar a InitializeComponent
			//
		}

		int idRolNomina = 0;
		int idAsiento = 0;
		DateTime dtFecha = DateTime.Today;
		int iRolComisionDecimo = 0;

		public frmRolAcreditacion(int IdRolNomina, int IdAsiento, DateTime DTFecha, int IRolComisionDecimo)
		{
			//
			// Necesario para admitir el Diseñador de Windows Forms
			//
			InitializeComponent();

			idRolNomina = IdRolNomina;
			idAsiento = IdAsiento;
			dtFecha = DTFecha;
			iRolComisionDecimo = IRolComisionDecimo;
			//
			// TODO: agregar código de constructor después de llamar a InitializeComponent
			//
		}

		int idBanco = 0;
		int iPosicion = 0;
		int iCantidad = 0;
		decimal dValor = 0.00m;
		string sNumeroDeCuenta = "";

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
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCuenta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Descripcion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idBanco");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Posicion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCuenta");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Descripcion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn3 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idBanco");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn4 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Posicion");
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton1 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmRolAcreditacion));
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn5 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("CED. IDENTIDAD");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn6 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("CUENTA BANCARIA");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn7 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("APELLIDO");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn8 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("NOMBRE");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn9 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("TOTAL A PAGAR");
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CED. IDENTIDAD");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CUENTA BANCARIA");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("APELLIDO");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("NOMBRE");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("TOTAL A PAGAR");
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings1 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "TOTAL A PAGAR", 4, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "TOTAL A PAGAR", 4, true);
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.lblCuenta = new System.Windows.Forms.Label();
			this.cmbCuenta = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.ultraDataSource1 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.label10 = new System.Windows.Forms.Label();
			this.dtpFecha = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.btnCancelar = new System.Windows.Forms.Button();
			this.btnAcreditar = new System.Windows.Forms.Button();
			this.groupBox1 = new System.Windows.Forms.GroupBox();
			this.ultraDataSource2 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.uGridLoja = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraGridExcelExporter1 = new Infragistics.Win.UltraWinGrid.ExcelExport.UltraGridExcelExporter();
			this.saveFileDialog1 = new System.Windows.Forms.SaveFileDialog();
			this.btnGenerarArchivo = new System.Windows.Forms.Button();
			this.lblMensaje = new System.Windows.Forms.Label();
			this.txtIdAsiento = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.lblEstado = new System.Windows.Forms.Label();
			this.btnImprimir = new System.Windows.Forms.Button();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCuenta)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtpFecha)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridLoja)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdAsiento)).BeginInit();
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
			// lblCuenta
			// 
			this.lblCuenta.AutoSize = true;
			this.lblCuenta.BackColor = System.Drawing.Color.Transparent;
			this.lblCuenta.Location = new System.Drawing.Point(8, 34);
			this.lblCuenta.Name = "lblCuenta";
			this.lblCuenta.Size = new System.Drawing.Size(41, 16);
			this.lblCuenta.TabIndex = 170;
			this.lblCuenta.Text = "Cuenta";
			// 
			// cmbCuenta
			// 
			appearance1.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbCuenta.Appearance = appearance1;
			this.cmbCuenta.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbCuenta.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbCuenta.DataSource = this.ultraDataSource1;
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridColumn2.Width = 352;
			ultraGridColumn3.Header.VisiblePosition = 2;
			ultraGridColumn3.Hidden = true;
			ultraGridColumn4.Header.VisiblePosition = 3;
			ultraGridColumn4.Hidden = true;
			ultraGridBand1.Columns.AddRange(new object[] {
																										 ultraGridColumn1,
																										 ultraGridColumn2,
																										 ultraGridColumn3,
																										 ultraGridColumn4});
			this.cmbCuenta.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.cmbCuenta.DisplayMember = "Descripcion";
			this.cmbCuenta.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbCuenta.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbCuenta.Location = new System.Drawing.Point(72, 32);
			this.cmbCuenta.Name = "cmbCuenta";
			this.cmbCuenta.Size = new System.Drawing.Size(352, 21);
			this.cmbCuenta.TabIndex = 169;
			this.cmbCuenta.ValueMember = "idCuenta";
			this.cmbCuenta.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cmbCuenta_KeyDown);
			this.cmbCuenta.ValueChanged += new System.EventHandler(this.cmbCuenta_ValueChanged);
			this.cmbCuenta.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.cmbCuenta_InitializeLayout);
			// 
			// ultraDataSource1
			// 
			ultraDataColumn1.DataType = typeof(int);
			ultraDataColumn3.DataType = typeof(int);
			ultraDataColumn3.DefaultValue = 0;
			ultraDataColumn4.DataType = typeof(int);
			ultraDataColumn4.DefaultValue = 0;
			this.ultraDataSource1.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn1,
																																 ultraDataColumn2,
																																 ultraDataColumn3,
																																 ultraDataColumn4});
			// 
			// label10
			// 
			this.label10.AutoSize = true;
			this.label10.BackColor = System.Drawing.Color.Transparent;
			this.label10.Location = new System.Drawing.Point(8, 8);
			this.label10.Name = "label10";
			this.label10.Size = new System.Drawing.Size(36, 16);
			this.label10.TabIndex = 172;
			this.label10.Text = "Fecha";
			this.label10.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// dtpFecha
			// 
			this.dtpFecha.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton1.Caption = "Today";
			this.dtpFecha.DateButtons.Add(dateButton1);
			this.dtpFecha.Format = "dd/MM/yyyy";
			this.dtpFecha.Location = new System.Drawing.Point(72, 6);
			this.dtpFecha.Name = "dtpFecha";
			this.dtpFecha.NonAutoSizeHeight = 23;
			this.dtpFecha.Size = new System.Drawing.Size(104, 21);
			this.dtpFecha.TabIndex = 171;
			this.dtpFecha.Value = ((object)(resources.GetObject("dtpFecha.Value")));
			this.dtpFecha.KeyDown += new System.Windows.Forms.KeyEventHandler(this.dtpFecha_KeyDown);
			// 
			// btnCancelar
			// 
			this.btnCancelar.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.btnCancelar.Location = new System.Drawing.Point(104, 184);
			this.btnCancelar.Name = "btnCancelar";
			this.btnCancelar.Size = new System.Drawing.Size(88, 23);
			this.btnCancelar.TabIndex = 175;
			this.btnCancelar.Text = "&Salir";
			this.btnCancelar.Click += new System.EventHandler(this.btnCancelar_Click);
			// 
			// btnAcreditar
			// 
			this.btnAcreditar.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.btnAcreditar.Location = new System.Drawing.Point(8, 184);
			this.btnAcreditar.Name = "btnAcreditar";
			this.btnAcreditar.Size = new System.Drawing.Size(88, 23);
			this.btnAcreditar.TabIndex = 173;
			this.btnAcreditar.Text = "&Acreditar";
			this.btnAcreditar.Click += new System.EventHandler(this.btnAcreditar_Click);
			// 
			// groupBox1
			// 
			this.groupBox1.Location = new System.Drawing.Point(0, 168);
			this.groupBox1.Name = "groupBox1";
			this.groupBox1.Size = new System.Drawing.Size(440, 8);
			this.groupBox1.TabIndex = 176;
			this.groupBox1.TabStop = false;
			// 
			// ultraDataSource2
			// 
			ultraDataColumn9.DataType = typeof(System.Decimal);
			ultraDataColumn9.DefaultValue = new System.Decimal(new int[] {
																																		 0,
																																		 0,
																																		 0,
																																		 0});
			this.ultraDataSource2.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn5,
																																 ultraDataColumn6,
																																 ultraDataColumn7,
																																 ultraDataColumn8,
																																 ultraDataColumn9});
			// 
			// uGridLoja
			// 
			this.uGridLoja.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.uGridLoja.Cursor = System.Windows.Forms.Cursors.Default;
			this.uGridLoja.DataSource = this.ultraDataSource2;
			appearance2.BackColor = System.Drawing.Color.White;
			this.uGridLoja.DisplayLayout.Appearance = appearance2;
			ultraGridColumn5.Header.VisiblePosition = 0;
			ultraGridColumn5.Width = 100;
			ultraGridColumn6.Header.VisiblePosition = 1;
			ultraGridColumn6.Width = 150;
			ultraGridColumn7.Header.VisiblePosition = 2;
			ultraGridColumn7.Width = 160;
			ultraGridColumn8.Header.VisiblePosition = 3;
			ultraGridColumn8.Width = 160;
			appearance3.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn9.CellAppearance = appearance3;
			ultraGridColumn9.Format = "#,##0.00";
			ultraGridColumn9.Header.VisiblePosition = 4;
			ultraGridColumn9.Width = 120;
			ultraGridBand2.Columns.AddRange(new object[] {
																										 ultraGridColumn5,
																										 ultraGridColumn6,
																										 ultraGridColumn7,
																										 ultraGridColumn8,
																										 ultraGridColumn9});
			ultraGridBand2.Header.Caption = "SUPERMERCADO DE COMPUTADORES COMPUBUSSINES CIA. LTDA.";
			ultraGridBand2.HeaderVisible = true;
			appearance4.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings1.Appearance = appearance4;
			summarySettings1.DisplayFormat = "{0: #,##0.00}";
			summarySettings1.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			ultraGridBand2.Summaries.AddRange(new Infragistics.Win.UltraWinGrid.SummarySettings[] {
																																															summarySettings1});
			ultraGridBand2.SummaryFooterCaption = "";
			this.uGridLoja.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			this.uGridLoja.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
			this.uGridLoja.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.uGridLoja.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.False;
			appearance5.BackColor = System.Drawing.Color.Transparent;
			this.uGridLoja.DisplayLayout.Override.CardAreaAppearance = appearance5;
			appearance6.ForeColor = System.Drawing.Color.Black;
			appearance6.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridLoja.DisplayLayout.Override.CellAppearance = appearance6;
			appearance7.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance7.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance7.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance7.FontData.BoldAsString = "True";
			appearance7.FontData.Name = "Arial";
			appearance7.FontData.SizeInPoints = 10F;
			appearance7.ForeColor = System.Drawing.Color.White;
			appearance7.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridLoja.DisplayLayout.Override.HeaderAppearance = appearance7;
			appearance8.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance8.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance8.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridLoja.DisplayLayout.Override.RowSelectorAppearance = appearance8;
			appearance9.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance9.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance9.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridLoja.DisplayLayout.Override.SelectedRowAppearance = appearance9;
			this.uGridLoja.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridLoja.Location = new System.Drawing.Point(-120, 216);
			this.uGridLoja.Name = "uGridLoja";
			this.uGridLoja.Size = new System.Drawing.Size(712, 242);
			this.uGridLoja.TabIndex = 721;
			this.uGridLoja.Visible = false;
			// 
			// btnGenerarArchivo
			// 
			this.btnGenerarArchivo.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.btnGenerarArchivo.Location = new System.Drawing.Point(8, 136);
			this.btnGenerarArchivo.Name = "btnGenerarArchivo";
			this.btnGenerarArchivo.Size = new System.Drawing.Size(104, 23);
			this.btnGenerarArchivo.TabIndex = 722;
			this.btnGenerarArchivo.Text = "&Generar Archivo";
			this.btnGenerarArchivo.Click += new System.EventHandler(this.btnGenerarArchivo_Click);
			// 
			// lblMensaje
			// 
			this.lblMensaje.AutoSize = true;
			this.lblMensaje.BackColor = System.Drawing.Color.Transparent;
			this.lblMensaje.Font = new System.Drawing.Font("Microsoft Sans Serif", 11F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblMensaje.ForeColor = System.Drawing.Color.Firebrick;
			this.lblMensaje.Location = new System.Drawing.Point(8, 64);
			this.lblMensaje.Name = "lblMensaje";
			this.lblMensaje.Size = new System.Drawing.Size(0, 20);
			this.lblMensaje.TabIndex = 723;
			this.lblMensaje.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// txtIdAsiento
			// 
			this.txtIdAsiento.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.txtIdAsiento.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtIdAsiento.Enabled = false;
			this.txtIdAsiento.Location = new System.Drawing.Point(408, 8);
			this.txtIdAsiento.Name = "txtIdAsiento";
			this.txtIdAsiento.PromptChar = ' ';
			this.txtIdAsiento.Size = new System.Drawing.Size(16, 21);
			this.txtIdAsiento.TabIndex = 733;
			this.txtIdAsiento.Visible = false;
			// 
			// lblEstado
			// 
			this.lblEstado.AutoSize = true;
			this.lblEstado.BackColor = System.Drawing.Color.Transparent;
			this.lblEstado.Font = new System.Drawing.Font("Microsoft Sans Serif", 11F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblEstado.ForeColor = System.Drawing.Color.Firebrick;
			this.lblEstado.Location = new System.Drawing.Point(8, 96);
			this.lblEstado.Name = "lblEstado";
			this.lblEstado.Size = new System.Drawing.Size(0, 20);
			this.lblEstado.TabIndex = 734;
			this.lblEstado.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// btnImprimir
			// 
			this.btnImprimir.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btnImprimir.Image = ((System.Drawing.Image)(resources.GetObject("btnImprimir.Image")));
			this.btnImprimir.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnImprimir.Location = new System.Drawing.Point(120, 136);
			this.btnImprimir.Name = "btnImprimir";
			this.btnImprimir.Size = new System.Drawing.Size(88, 23);
			this.btnImprimir.TabIndex = 735;
			this.btnImprimir.Text = "&Imprimir";
			this.btnImprimir.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnImprimir.Click += new System.EventHandler(this.btnImprimir_Click);
			// 
			// frmRolAcreditacion
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(440, 216);
			this.ControlBox = false;
			this.Controls.Add(this.btnImprimir);
			this.Controls.Add(this.lblEstado);
			this.Controls.Add(this.txtIdAsiento);
			this.Controls.Add(this.lblMensaje);
			this.Controls.Add(this.btnGenerarArchivo);
			this.Controls.Add(this.groupBox1);
			this.Controls.Add(this.btnCancelar);
			this.Controls.Add(this.btnAcreditar);
			this.Controls.Add(this.label10);
			this.Controls.Add(this.lblCuenta);
			this.Controls.Add(this.dtpFecha);
			this.Controls.Add(this.cmbCuenta);
			this.Controls.Add(this.uGridLoja);
			this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
			this.KeyPreview = true;
			this.Name = "frmRolAcreditacion";
			this.ShowInTaskbar = false;
			this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
			this.Text = "Acreditación Nomina";
			this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.frmRolAcreditacion_KeyDown);
			this.Load += new System.EventHandler(this.frmRolAcreditacion_Load);
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCuenta)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtpFecha)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridLoja)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdAsiento)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);		
		}

		private void frmRolAcreditacion_Load(object sender, System.EventArgs e)
		{
			this.dtpFecha.CalendarInfo.MinDate = dtFecha;
			this.dtpFecha.CalendarInfo.MaxDate = DateTime.Today;
			this.dtpFecha.Value = DateTime.Today;

			this.cmbCuenta.DataSource = Funcion.dvProcedimiento(cdsSeteoF, string.Format("Exec RolListaCuentasAcreditacion {0}, {1}", idRolNomina, iRolComisionDecimo));
		}

		private void ArchivoTXTInternacionalCompleto()
		{
			#region Exporta
			try
			{
				SaveFileDialog sArchivo = new SaveFileDialog();

				sArchivo.Filter = "Documentos de texto (*.txt)|*.txt";
				sArchivo.FileName = sArchivo.FileName;
				sArchivo.OverwritePrompt = true;
				sArchivo.Title = "Guardar Lista Acreditaciones Archivos TXT";		
                                
				if (sArchivo.ShowDialog() == DialogResult.OK)
				{
					using (SqlConnection oConexion = new SqlConnection("Password=Bl45o6$9;Persist Security Info=True;User ID= " + MenuLatinium.stUsuario + ";Initial Catalog=" + MenuLatinium.stDirFacturacion + ";Data Source=" + MenuLatinium.stDirServidor))
					{						
						#region Conexion
						SqlCommand oCmdActualiza = oConexion.CreateCommand();
						oCmdActualiza.Connection = oConexion;
						oCmdActualiza.CommandTimeout = 0;
						string sSQL = string.Format("Exec RolListaPersonalInternacionalCompleto {0}, {1}, {2}", idRolNomina, idBanco, iRolComisionDecimo);
						oCmdActualiza.CommandText = sSQL;
						SqlDataAdapter oDADetalle = new SqlDataAdapter(oCmdActualiza);
						DataTable oDTDetalle = new DataTable();
						oDADetalle.Fill(oDTDetalle);	
						#endregion Conexion				
	
						#region Pagos
						using (StreamWriter file = new StreamWriter(sArchivo.FileName, true, System.Text.Encoding.Default))
						{						
							int i = 1;

							foreach (DataRow dr in oDTDetalle.Rows)
							{							
								string iLinea = dr["Pago"].ToString() + "\t" + sNumeroDeCuenta + "\t" + i.ToString() + "\t" + dr["Comprobante"].ToString() + "\t" + dr["Contrapartida"].ToString() + "\t" 
									+ dr["Moneda"].ToString() + "\t" + dr["Valor"].ToString() + "\t"
									+ dr["Titulo"].ToString() + "\t" + "32" + "\t" + dr["TipoDeCuenta"].ToString() + "\t" + dr["Cuenta"].ToString() + "\t"
									+ dr["TipoDoc"].ToString() + "\t" + dr["Cedula"].ToString() + "\t" + dr["Nombre"].ToString() + "\t"
									+ dr["DirBenef"].ToString() + "\t" + dr["CiudBenef"].ToString() + "\t" + dr["TelBenef"].ToString() + "\t"
									+ dr["LocalBenef"].ToString() + "\t" + dr["Referencia"].ToString() + "\t" + "|" + dr["RefAdicional"].ToString();

								file.WriteLine(iLinea);

								i++;
							}					
						}
						#endregion Pagos					
					}
				}
			}
			catch (Exception exc)
			{
				MessageBox.Show(exc.Message);
			}
			#endregion Exporta
		}

		private void ArchivoTXTDecimos()
		{
			#region Exporta
			try
			{
				SaveFileDialog sArchivo = new SaveFileDialog();

				sArchivo.Filter = "Documentos de texto (*.txt)|*.txt";
				sArchivo.FileName = sArchivo.FileName;
				sArchivo.OverwritePrompt = true;
				sArchivo.Title = "Guardar Lista Acreditaciones Archivos TXT";		
                                
				if (sArchivo.ShowDialog() == DialogResult.OK)
				{
					using (SqlConnection oConexion = new SqlConnection("Password=Bl45o6$9;Persist Security Info=True;User ID= " + MenuLatinium.stUsuario + ";Initial Catalog=" + MenuLatinium.stDirFacturacion + ";Data Source=" + MenuLatinium.stDirServidor))
					{						
						#region Conexion
						SqlCommand oCmdActualiza = oConexion.CreateCommand();
						oCmdActualiza.Connection = oConexion;
						oCmdActualiza.CommandTimeout = 0;
						string sSQL = string.Format("Exec [RolAcreditacionDecimoCuartoSueldo] {0}", idBanco);
						oCmdActualiza.CommandText = sSQL;
						SqlDataAdapter oDADetalle = new SqlDataAdapter(oCmdActualiza);
						DataTable oDTDetalle = new DataTable();
						oDADetalle.Fill(oDTDetalle);	
						#endregion Conexion

						#region Pagos
						using (StreamWriter file = new StreamWriter(sArchivo.FileName, true))
						{						
							int i = 1;
							foreach (DataRow dr in oDTDetalle.Rows)
							{							
								string iLinea = dr["Pago"].ToString() + "\t" + i.ToString() + "\t" + dr["Moneda"].ToString() + "\t" + dr["Valor"].ToString() + "\t"
									+ dr["Titulo"].ToString() + "\t" + dr["TipoDeCuenta"].ToString() + "\t" + dr["Cuenta"].ToString() + "\t"
									+ dr["Mensaje"].ToString() + "\t" + dr["TipoDoc"].ToString() + "\t" + dr["Cedula"].ToString() + "\t" + dr["Nombre"].ToString();

								file.WriteLine(iLinea);

								i++;
							}					
						}
						#endregion Pagos					
					}
				}
			}
			catch (Exception exc)
			{
				MessageBox.Show(exc.Message);
			}
			#endregion Exporta
		}

		private void ArchivoTXT()
		{
			#region Exporta
			try
			{
				SaveFileDialog sArchivo = new SaveFileDialog();

				sArchivo.Filter = "Documentos de texto (*.txt)|*.txt";
				sArchivo.FileName = sArchivo.FileName;
				sArchivo.OverwritePrompt = true;
				sArchivo.Title = "Guardar Lista Acreditaciones Archivos TXT";		
                                
				if (sArchivo.ShowDialog() == DialogResult.OK)
				{
					using (SqlConnection oConexion = new SqlConnection("Password=Bl45o6$9;Persist Security Info=True;User ID= " + MenuLatinium.stUsuario + ";Initial Catalog=" + MenuLatinium.stDirFacturacion + ";Data Source=" + MenuLatinium.stDirServidor))
					{						
						#region Conexion
						SqlCommand oCmdActualiza = oConexion.CreateCommand();
						oCmdActualiza.Connection = oConexion;
						oCmdActualiza.CommandTimeout = 0;
						string sSQL = string.Format("Exec RolListaPersonalInternacional {0}, {1}, {2}", idRolNomina, idBanco, iRolComisionDecimo);
						oCmdActualiza.CommandText = sSQL;
						SqlDataAdapter oDADetalle = new SqlDataAdapter(oCmdActualiza);
						DataTable oDTDetalle = new DataTable();
						oDADetalle.Fill(oDTDetalle);	
						#endregion Conexion

						#region Pagos
						using (StreamWriter file = new StreamWriter(sArchivo.FileName, true))
						{						
							int i = 1;
							foreach (DataRow dr in oDTDetalle.Rows)
							{							
								string iLinea = dr["Pago"].ToString() + "\t" + i.ToString() + "\t" + dr["Moneda"].ToString() + "\t" + dr["Valor"].ToString() + "\t"
									+ dr["Titulo"].ToString() + "\t" + dr["TipoDeCuenta"].ToString() + "\t" + dr["Cuenta"].ToString() + "\t"
									+ dr["Mensaje"].ToString() + "\t" + dr["TipoDoc"].ToString() + "\t" + dr["Cedula"].ToString() + "\t" + dr["Nombre"].ToString();

								file.WriteLine(iLinea);

								i++;
							}					
						}
						#endregion Pagos					
					}
				}
			}
			catch (Exception exc)
			{
				MessageBox.Show(exc.Message);
			}
			#endregion Exporta
		}

		private void HojaExcel()
		{			
			SaveFileDialog sArchivo = new SaveFileDialog();

			sArchivo.Filter = "Microsoft Excel (*.xls)|*.xls";
			sArchivo.FileName = sArchivo.FileName;
			sArchivo.OverwritePrompt = true;
			sArchivo.Title = "Guardar Lista Acreditaciones Banco De Loja";
                                
			if (sArchivo.ShowDialog() == DialogResult.OK)
			{
				this.uGridLoja.DataSource = Funcion.dvProcedimiento(cdsSeteoF, string.Format("Exec RolListaPersonalLoja {0}, 1, {1}", idRolNomina, iRolComisionDecimo));

				ultraGridExcelExporter1.Export(this.uGridLoja, sArchivo.FileName);
				System.Diagnostics.Process.Start(sArchivo.FileName);
			}
		}

		private void btnAcreditar_Click(object sender, System.EventArgs e)
		{
			#region Validacion
			if (this.dtpFecha.Value == System.DBNull.Value)
			{
				MessageBox.Show("Ingrese la Fecha", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.dtpFecha.Focus();
				return;
			}
			
			if (!Validacion.vbComboVacio(this.cmbCuenta, "Seleccione una cuenta")) return;
			#endregion Validacion

			#region Asiento - Estado Rol
			if ((int)this.txtIdAsiento.Value > 0 && DialogResult.No == MessageBox.Show(string.Format("Ya Existe una acreditación de la cuenta '{0}'\n\n¿Desea sobre escribirla?", 
				this.cmbCuenta.Text.ToString()), "Point Technology", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2)) return;
			
			string sSQL = "";
			
			if (iRolComisionDecimo == 0) sSQL = string.Format("Exec AsientoRolAcreditacion {0}, {1}, {2}, '{3}'", (int)this.txtIdAsiento.Value, idRolNomina, (int)this.cmbCuenta.Value, Convert.ToDateTime(this.dtpFecha.Value).ToString("yyyyMMdd"));
			if (iRolComisionDecimo == 1) sSQL = string.Format("Exec Com_AsientoAcreditacion {0}, {1}, {2}, '{3}'", (int)this.txtIdAsiento.Value, idRolNomina, (int)this.cmbCuenta.Value, Convert.ToDateTime(this.dtpFecha.Value).ToString("yyyyMMdd"));
      if (iRolComisionDecimo == 2) sSQL = string.Format("Exec AsientoDecimoTerSueldo {0}, {1}, {2}, '{3}'", (int)this.txtIdAsiento.Value, idRolNomina, (int)this.cmbCuenta.Value, Convert.ToDateTime(this.dtpFecha.Value).ToString("yyyyMMdd"));
			int idAsiento = Funcion.iEscalarSQL(cdsSeteoF, sSQL);
	
			Funcion.EjecutaSQL(cdsSeteoF, string.Format("Exec NominaActualizaEstadoAcreditado {0}, {1}", idRolNomina, iRolComisionDecimo));

			MessageBox.Show(string.Format("Acreditación de la Cuenta '{0}' Generada Correctamente", this.cmbCuenta.Text.ToString()), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);

			this.cmbCuenta_ValueChanged(sender, e);
			#endregion Asiento - Estado Rol

			#region Impresion
			string strFiltro = "{Asiento.idAsiento} = " + idAsiento.ToString();
				
			Reporte Reporte = new Reporte("Asientos.rpt", strFiltro);
			Reporte.MdiParent = this.MdiParent;				
			Reporte.Show();
			#endregion Impresion
		}

		private void btnCancelar_Click(object sender, System.EventArgs e)
		{
			this.Close();
		}

		private void dtpFecha_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.cmbCuenta.Focus();
		}

		private void cmbCuenta_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.btnAcreditar.Focus();
		}

		private void frmRolAcreditacion_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Escape) this.Close();
		}

		private void btnGenerarArchivo_Click(object sender, System.EventArgs e)
		{
			if (this.dtpFecha.Value == System.DBNull.Value)
			{
				MessageBox.Show("Ingrese la Fecha", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.dtpFecha.Focus();
				return;
			}
			
			if (!Validacion.vbComboVacio(this.cmbCuenta, "Seleccione una cuenta")) return;

			if (idBanco == 1) this.HojaExcel();

			if (idBanco == 5 || idBanco == 7 || idBanco == 21)
				if (MenuLatinium.stIdDB == 8) this.ArchivoTXTInternacionalCompleto(); 
				else this.ArchivoTXT();

//			ArchivoTXTDecimos();
		}

		private void cmbCuenta_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}

		private void cmbCuenta_ValueChanged(object sender, System.EventArgs e)
		{
			iPosicion = (int)this.cmbCuenta.ActiveRow.Cells["Posicion"].Value;
			idBanco = (int)this.cmbCuenta.ActiveRow.Cells["idBanco"].Value;
			sNumeroDeCuenta = Funcion.sEscalarSQL(cdsSeteoF, string.Format("Select NumCuenta From CuentaCorriente Where idCuenta = {0}", (int)this.cmbCuenta.Value));
			
			iCantidad = 0;
			dValor = 0.00m;

			this.txtIdAsiento.Value = 0;
			this.lblEstado.Text = "";

			string sSQL = string.Format("Exec RolValoresPorBanco {0}, {1}, {2}", idRolNomina, idBanco, iRolComisionDecimo);
			SqlDataReader dr = Funcion.rEscalarSQL(cdsSeteoF, sSQL, true);
			dr.Read();
			if (dr.HasRows)
			{
				iCantidad = dr.GetInt32(0);
				dValor = dr.GetDecimal(1);
			}
			dr.Close();

			this.lblMensaje.Text = string.Format("Empleados : {0} - Acreditación : {1}", iCantidad, dValor);

			sSQL = string.Format("Exec RolRetornaIdAsientoacreditacion {0}, {1}, {2}", idRolNomina, iPosicion, iRolComisionDecimo);
			this.txtIdAsiento.Value = Funcion.iEscalarSQL(cdsSeteoF, sSQL);

			if ((int)this.txtIdAsiento.Value > 0) this.lblEstado.Text = "ACREDITACIÓN CONTABILIZADA"; else this.lblEstado.Text = "ACREDITACIÓN PENDIENTE";
		}

		private void btnImprimir_Click(object sender, System.EventArgs e)
		{
			#region Impresion Pantalla Bloqueada
			ParameterFields paramFields = new ParameterFields ();
																				
			ParameterField paramFieldidRolNomina = new ParameterField ();
			ParameterDiscreteValue ValidRolNomina = new ParameterDiscreteValue ();
			paramFieldidRolNomina.ParameterFieldName = "@idRolNomina";
			ValidRolNomina.Value = idRolNomina;
			paramFieldidRolNomina.CurrentValues.Add (ValidRolNomina);
			paramFields.Add (paramFieldidRolNomina);
																
			ParameterField paramFieldRolComision = new ParameterField ();
			ParameterDiscreteValue ValIdRolComision = new ParameterDiscreteValue ();
			paramFieldRolComision.ParameterFieldName = "@RolComision";
			ValIdRolComision.Value = iRolComisionDecimo;
			paramFieldRolComision.CurrentValues.Add (ValIdRolComision);
			paramFields.Add (paramFieldRolComision);
																			
			ParameterField paramFieldidTipoDePago = new ParameterField ();
			ParameterDiscreteValue ValidTipoDePago = new ParameterDiscreteValue ();
			paramFieldidTipoDePago.ParameterFieldName = "@idTipoDePago";
			ValidTipoDePago.Value = 1;
			paramFieldidTipoDePago.CurrentValues.Add (ValidTipoDePago);
			paramFields.Add (paramFieldidTipoDePago);
																
			ParameterField paramFieldidBanco = new ParameterField ();
			ParameterDiscreteValue ValIdidBanco = new ParameterDiscreteValue ();
			paramFieldidBanco.ParameterFieldName = "@idBanco";
			ValIdidBanco.Value = idBanco;
			paramFieldidBanco.CurrentValues.Add (ValIdidBanco);
			paramFields.Add (paramFieldidBanco);					

			Reporte miRep = new Reporte("ListaAcreditacionCheques.rpt", "");
			miRep.crVista.ParameterFieldInfo = paramFields;
			miRep.MdiParent = this.MdiParent;						
			miRep.Show();		
			#endregion Impresion Pantalla Bloqueada
		}
	}
}
