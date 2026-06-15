using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;
using System.Data.SqlClient;
using DevExpress.XtraEditors;
using CrystalDecisions;
using CrystalDecisions.CrystalReports.Engine;
using CrystalDecisions.ReportSource;
using CrystalDecisions.Shared;
using Infragistics.Win.UltraWinGrid;



namespace Latinium
{
	/// <summary>
	/// Descripción breve de frmComicionesSemanales.
	/// </summary>
	public class frmComicionesSemanales : System.Windows.Forms.Form
	{
		private System.Windows.Forms.Button btnSalir;
		private Infragistics.Win.UltraWinGrid.UltraGrid uGridInformacion;
		private System.Windows.Forms.GroupBox groupBox1;
		private System.Windows.Forms.GroupBox groupBox2;
		private System.Windows.Forms.Label lblEstado;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtestado;
		private System.Windows.Forms.Button btnEditar;
		private System.Windows.Forms.Button btnGuardar;
		private System.Windows.Forms.Button btnCancelar;
		private System.Windows.Forms.Button btnNuevo;
		private System.Windows.Forms.Button btnConsultar;
		private System.Windows.Forms.Label lblContador;
		private System.Windows.Forms.Label label4;
		private System.Windows.Forms.Label label2;
		private System.Windows.Forms.Label label1;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtDescripcion;
		public Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtHasta;
		public Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtDesde;
		private System.Windows.Forms.Button btnImportar;
		private DevExpress.XtraEditors.SimpleButton btnTipoBono;
		private System.Windows.Forms.Label label3;
		private C1.Data.C1DataSet cdsSeteoF;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource4;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtidComicionesSemanales;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbTipoBono;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtidMaestroSubida;
		private System.Windows.Forms.Button btnAgregar;
		private System.Windows.Forms.Button btnAcreditacion;
		private System.Windows.Forms.Button btnExportar;
		private System.Windows.Forms.Label label10;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtpFecha;
		private System.Windows.Forms.ContextMenu mnuAcreditacion;
		private System.Windows.Forms.MenuItem mnuAcreditacionTransferencias;
		private System.Windows.Forms.MenuItem mnuAcreditacionCheques;
		private System.Windows.Forms.Label lblEstado2;
		private System.Windows.Forms.Button btnAnular;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public frmComicionesSemanales()
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
		
		private void UnloadMe()
		{
			this.Close();
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
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("N", -1, null, 0, Infragistics.Win.UltraWinGrid.SortIndicator.Ascending, false);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idNomina");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cedula");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Apellidos");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombres");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("TipoDePago");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("TipoDepagoM");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idBanco");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Banco");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("NumeroDeCuenta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn11 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idProyecto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn12 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Centro_De_Costo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn13 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idDepartamento");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn14 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Departamento");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn15 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cargo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn16 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bono");
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn17 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("TipoBanco");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn18 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Agregar", 0);
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmComicionesSemanales));
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings1 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Bono", 15, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Bono", 15, true);
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("N");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idNomina");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn3 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Cedula");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn4 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Apellidos");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn5 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombres");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn6 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("TipoDePago");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn7 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("TipoDepagoM");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn8 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idBanco");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn9 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Banco");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn10 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("NumeroDeCuenta");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn11 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idProyecto");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn12 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Centro_De_Costo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn13 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idDepartamento");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn14 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Departamento");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn15 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Cargo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn16 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Bono");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn17 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("TipoBanco");
			Infragistics.Win.Appearance appearance11 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance12 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton1 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.Appearance appearance13 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton2 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.Appearance appearance14 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn19 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("IdTipoBono");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn20 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton3 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			this.btnSalir = new System.Windows.Forms.Button();
			this.uGridInformacion = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource4 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.groupBox1 = new System.Windows.Forms.GroupBox();
			this.groupBox2 = new System.Windows.Forms.GroupBox();
			this.lblEstado = new System.Windows.Forms.Label();
			this.txtestado = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtidComicionesSemanales = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.btnEditar = new System.Windows.Forms.Button();
			this.btnGuardar = new System.Windows.Forms.Button();
			this.btnCancelar = new System.Windows.Forms.Button();
			this.btnNuevo = new System.Windows.Forms.Button();
			this.btnConsultar = new System.Windows.Forms.Button();
			this.lblContador = new System.Windows.Forms.Label();
			this.label4 = new System.Windows.Forms.Label();
			this.label2 = new System.Windows.Forms.Label();
			this.label1 = new System.Windows.Forms.Label();
			this.txtDescripcion = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.dtHasta = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.dtDesde = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.btnImportar = new System.Windows.Forms.Button();
			this.btnTipoBono = new DevExpress.XtraEditors.SimpleButton();
			this.label3 = new System.Windows.Forms.Label();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.cmbTipoBono = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.txtidMaestroSubida = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.btnAgregar = new System.Windows.Forms.Button();
			this.btnAcreditacion = new System.Windows.Forms.Button();
			this.btnExportar = new System.Windows.Forms.Button();
			this.label10 = new System.Windows.Forms.Label();
			this.dtpFecha = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.mnuAcreditacion = new System.Windows.Forms.ContextMenu();
			this.mnuAcreditacionTransferencias = new System.Windows.Forms.MenuItem();
			this.mnuAcreditacionCheques = new System.Windows.Forms.MenuItem();
			this.lblEstado2 = new System.Windows.Forms.Label();
			this.btnAnular = new System.Windows.Forms.Button();
			((System.ComponentModel.ISupportInitialize)(this.uGridInformacion)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource4)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtestado)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtidComicionesSemanales)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDescripcion)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtHasta)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtDesde)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbTipoBono)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtidMaestroSubida)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtpFecha)).BeginInit();
			this.SuspendLayout();
			// 
			// btnSalir
			// 
			this.btnSalir.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnSalir.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(230)), ((System.Byte)(230)), ((System.Byte)(230)));
			this.btnSalir.CausesValidation = false;
			this.btnSalir.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.btnSalir.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(50)), ((System.Byte)(50)), ((System.Byte)(50)));
			this.btnSalir.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnSalir.ImageIndex = 62;
			this.btnSalir.Location = new System.Drawing.Point(520, 437);
			this.btnSalir.Name = "btnSalir";
			this.btnSalir.Size = new System.Drawing.Size(80, 22);
			this.btnSalir.TabIndex = 1019;
			this.btnSalir.Text = "Salir";
			this.btnSalir.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnSalir.Click += new System.EventHandler(this.btnSalir_Click);
			// 
			// uGridInformacion
			// 
			this.uGridInformacion.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.uGridInformacion.Cursor = System.Windows.Forms.Cursors.Default;
			this.uGridInformacion.DataSource = this.ultraDataSource4;
			appearance1.BackColor = System.Drawing.Color.White;
			this.uGridInformacion.DisplayLayout.Appearance = appearance1;
			ultraGridColumn1.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn1.Header.VisiblePosition = 1;
			ultraGridColumn1.Width = 25;
			ultraGridColumn2.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn2.Header.VisiblePosition = 2;
			ultraGridColumn2.Hidden = true;
			ultraGridColumn3.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn3.Header.VisiblePosition = 3;
			ultraGridColumn4.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn4.Header.VisiblePosition = 4;
			ultraGridColumn5.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn5.Header.VisiblePosition = 5;
			ultraGridColumn6.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn6.Header.VisiblePosition = 6;
			ultraGridColumn6.Hidden = true;
			ultraGridColumn7.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn7.Header.Caption = "Tipo Pago";
			ultraGridColumn7.Header.VisiblePosition = 7;
			ultraGridColumn8.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn8.Header.VisiblePosition = 8;
			ultraGridColumn8.Hidden = true;
			ultraGridColumn9.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn9.Header.VisiblePosition = 9;
			ultraGridColumn10.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn10.Header.Caption = "Numero De Cuenta";
			ultraGridColumn10.Header.VisiblePosition = 10;
			ultraGridColumn11.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn11.Header.VisiblePosition = 11;
			ultraGridColumn11.Hidden = true;
			ultraGridColumn12.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn12.Header.Caption = "Centro De Costo";
			ultraGridColumn12.Header.VisiblePosition = 12;
			ultraGridColumn13.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn13.Header.VisiblePosition = 13;
			ultraGridColumn13.Hidden = true;
			ultraGridColumn14.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn14.Header.VisiblePosition = 14;
			ultraGridColumn15.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn15.Header.VisiblePosition = 15;
			ultraGridColumn16.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance2.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn16.CellAppearance = appearance2;
			ultraGridColumn16.Format = "#,##0.00";
			ultraGridColumn16.Header.VisiblePosition = 16;
			ultraGridColumn17.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn17.Header.VisiblePosition = 17;
			ultraGridColumn17.Hidden = true;
			ultraGridColumn18.ButtonDisplayStyle = Infragistics.Win.UltraWinGrid.ButtonDisplayStyle.Always;
			ultraGridColumn18.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance3.ImageHAlign = Infragistics.Win.HAlign.Center;
			appearance3.ImageVAlign = Infragistics.Win.VAlign.Middle;
			ultraGridColumn18.CellAppearance = appearance3;
			appearance4.Image = ((object)(resources.GetObject("appearance4.Image")));
			appearance4.ImageHAlign = Infragistics.Win.HAlign.Center;
			appearance4.ImageVAlign = Infragistics.Win.VAlign.Middle;
			ultraGridColumn18.CellButtonAppearance = appearance4;
			ultraGridColumn18.Header.Caption = "...";
			ultraGridColumn18.Header.VisiblePosition = 0;
			ultraGridColumn18.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.Button;
			ultraGridColumn18.Width = 25;
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
																										 ultraGridColumn12,
																										 ultraGridColumn13,
																										 ultraGridColumn14,
																										 ultraGridColumn15,
																										 ultraGridColumn16,
																										 ultraGridColumn17,
																										 ultraGridColumn18});
			appearance5.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings1.Appearance = appearance5;
			summarySettings1.DisplayFormat = "{0: #,##0.00}";
			summarySettings1.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			ultraGridBand1.Summaries.AddRange(new Infragistics.Win.UltraWinGrid.SummarySettings[] {
																																															summarySettings1});
			this.uGridInformacion.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.uGridInformacion.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.TemplateOnBottom;
			this.uGridInformacion.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.True;
			this.uGridInformacion.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance6.BackColor = System.Drawing.Color.Transparent;
			this.uGridInformacion.DisplayLayout.Override.CardAreaAppearance = appearance6;
			appearance7.ForeColor = System.Drawing.Color.Black;
			appearance7.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridInformacion.DisplayLayout.Override.CellAppearance = appearance7;
			this.uGridInformacion.DisplayLayout.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.Edit;
			appearance8.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance8.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance8.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance8.FontData.BoldAsString = "True";
			appearance8.FontData.Name = "Arial";
			appearance8.FontData.SizeInPoints = 8F;
			appearance8.ForeColor = System.Drawing.Color.White;
			appearance8.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridInformacion.DisplayLayout.Override.HeaderAppearance = appearance8;
			this.uGridInformacion.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti;
			appearance9.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance9.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance9.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridInformacion.DisplayLayout.Override.RowSelectorAppearance = appearance9;
			appearance10.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance10.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance10.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridInformacion.DisplayLayout.Override.SelectedRowAppearance = appearance10;
			this.uGridInformacion.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridInformacion.Location = new System.Drawing.Point(7, 83);
			this.uGridInformacion.Name = "uGridInformacion";
			this.uGridInformacion.Size = new System.Drawing.Size(1111, 336);
			this.uGridInformacion.TabIndex = 1017;
			this.uGridInformacion.AfterRowInsert += new Infragistics.Win.UltraWinGrid.RowEventHandler(this.uGridInformacion_AfterRowInsert);
			this.uGridInformacion.BeforeCellUpdate += new Infragistics.Win.UltraWinGrid.BeforeCellUpdateEventHandler(this.uGridInformacion_BeforeCellUpdate);
			this.uGridInformacion.ClickCellButton += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.uGridInformacion_ClickCellButton);
			this.uGridInformacion.AfterCellUpdate += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.uGridInformacion_AfterCellUpdate);
			// 
			// ultraDataSource4
			// 
			ultraDataColumn1.DataType = typeof(int);
			ultraDataColumn2.DataType = typeof(int);
			ultraDataColumn6.DataType = typeof(int);
			ultraDataColumn8.DataType = typeof(int);
			ultraDataColumn11.DataType = typeof(int);
			ultraDataColumn13.DataType = typeof(int);
			ultraDataColumn16.DataType = typeof(System.Decimal);
			ultraDataColumn17.DataType = typeof(int);
			this.ultraDataSource4.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn1,
																																 ultraDataColumn2,
																																 ultraDataColumn3,
																																 ultraDataColumn4,
																																 ultraDataColumn5,
																																 ultraDataColumn6,
																																 ultraDataColumn7,
																																 ultraDataColumn8,
																																 ultraDataColumn9,
																																 ultraDataColumn10,
																																 ultraDataColumn11,
																																 ultraDataColumn12,
																																 ultraDataColumn13,
																																 ultraDataColumn14,
																																 ultraDataColumn15,
																																 ultraDataColumn16,
																																 ultraDataColumn17});
			// 
			// groupBox1
			// 
			this.groupBox1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.groupBox1.Location = new System.Drawing.Point(8, 424);
			this.groupBox1.Name = "groupBox1";
			this.groupBox1.Size = new System.Drawing.Size(832, 8);
			this.groupBox1.TabIndex = 1016;
			this.groupBox1.TabStop = false;
			this.groupBox1.Enter += new System.EventHandler(this.groupBox1_Enter);
			// 
			// groupBox2
			// 
			this.groupBox2.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.groupBox2.Location = new System.Drawing.Point(16, 72);
			this.groupBox2.Name = "groupBox2";
			this.groupBox2.Size = new System.Drawing.Size(1160, 8);
			this.groupBox2.TabIndex = 1015;
			this.groupBox2.TabStop = false;
			// 
			// lblEstado
			// 
			this.lblEstado.AutoSize = true;
			this.lblEstado.Font = new System.Drawing.Font("Tahoma", 12F, System.Drawing.FontStyle.Bold);
			this.lblEstado.ForeColor = System.Drawing.Color.Firebrick;
			this.lblEstado.Location = new System.Drawing.Point(688, 48);
			this.lblEstado.Name = "lblEstado";
			this.lblEstado.Size = new System.Drawing.Size(0, 23);
			this.lblEstado.TabIndex = 1014;
			this.lblEstado.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// txtestado
			// 
			this.txtestado.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtestado.Enabled = false;
			this.txtestado.Location = new System.Drawing.Point(993, 444);
			this.txtestado.MaxValue = 3600;
			this.txtestado.MinValue = 0;
			this.txtestado.Name = "txtestado";
			this.txtestado.PromptChar = ' ';
			this.txtestado.Size = new System.Drawing.Size(49, 21);
			this.txtestado.TabIndex = 1013;
			this.txtestado.Visible = false;
			// 
			// txtidComicionesSemanales
			// 
			this.txtidComicionesSemanales.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtidComicionesSemanales.Location = new System.Drawing.Point(933, 444);
			this.txtidComicionesSemanales.MaxValue = 3600;
			this.txtidComicionesSemanales.MinValue = 0;
			this.txtidComicionesSemanales.Name = "txtidComicionesSemanales";
			this.txtidComicionesSemanales.PromptChar = ' ';
			this.txtidComicionesSemanales.Size = new System.Drawing.Size(56, 21);
			this.txtidComicionesSemanales.TabIndex = 1012;
			this.txtidComicionesSemanales.Visible = false;
			// 
			// btnEditar
			// 
			this.btnEditar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnEditar.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(230)), ((System.Byte)(230)), ((System.Byte)(230)));
			this.btnEditar.CausesValidation = false;
			this.btnEditar.Enabled = false;
			this.btnEditar.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.btnEditar.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(50)), ((System.Byte)(50)), ((System.Byte)(50)));
			this.btnEditar.Image = ((System.Drawing.Image)(resources.GetObject("btnEditar.Image")));
			this.btnEditar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnEditar.Location = new System.Drawing.Point(267, 437);
			this.btnEditar.Name = "btnEditar";
			this.btnEditar.Size = new System.Drawing.Size(80, 22);
			this.btnEditar.TabIndex = 1010;
			this.btnEditar.Text = "Editar";
			this.btnEditar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnEditar.Click += new System.EventHandler(this.btnEditar_Click);
			// 
			// btnGuardar
			// 
			this.btnGuardar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnGuardar.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(230)), ((System.Byte)(230)), ((System.Byte)(230)));
			this.btnGuardar.Enabled = false;
			this.btnGuardar.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.btnGuardar.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(50)), ((System.Byte)(50)), ((System.Byte)(50)));
			this.btnGuardar.Image = ((System.Drawing.Image)(resources.GetObject("btnGuardar.Image")));
			this.btnGuardar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnGuardar.Location = new System.Drawing.Point(353, 437);
			this.btnGuardar.Name = "btnGuardar";
			this.btnGuardar.Size = new System.Drawing.Size(80, 22);
			this.btnGuardar.TabIndex = 1009;
			this.btnGuardar.Text = "Guardar";
			this.btnGuardar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnGuardar.Click += new System.EventHandler(this.btnGuardar_Click);
			// 
			// btnCancelar
			// 
			this.btnCancelar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnCancelar.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(230)), ((System.Byte)(230)), ((System.Byte)(230)));
			this.btnCancelar.CausesValidation = false;
			this.btnCancelar.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.btnCancelar.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(50)), ((System.Byte)(50)), ((System.Byte)(50)));
			this.btnCancelar.Image = ((System.Drawing.Image)(resources.GetObject("btnCancelar.Image")));
			this.btnCancelar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnCancelar.Location = new System.Drawing.Point(182, 437);
			this.btnCancelar.Name = "btnCancelar";
			this.btnCancelar.Size = new System.Drawing.Size(80, 22);
			this.btnCancelar.TabIndex = 1011;
			this.btnCancelar.Text = "Cancelar";
			this.btnCancelar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnCancelar.Click += new System.EventHandler(this.btnCancelar_Click);
			// 
			// btnNuevo
			// 
			this.btnNuevo.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnNuevo.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(230)), ((System.Byte)(230)), ((System.Byte)(230)));
			this.btnNuevo.CausesValidation = false;
			this.btnNuevo.Enabled = false;
			this.btnNuevo.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.btnNuevo.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(50)), ((System.Byte)(50)), ((System.Byte)(50)));
			this.btnNuevo.Image = ((System.Drawing.Image)(resources.GetObject("btnNuevo.Image")));
			this.btnNuevo.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnNuevo.Location = new System.Drawing.Point(8, 437);
			this.btnNuevo.Name = "btnNuevo";
			this.btnNuevo.Size = new System.Drawing.Size(80, 22);
			this.btnNuevo.TabIndex = 1007;
			this.btnNuevo.Text = "Nuevo";
			this.btnNuevo.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnNuevo.Click += new System.EventHandler(this.btnNuevo_Click);
			// 
			// btnConsultar
			// 
			this.btnConsultar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnConsultar.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(230)), ((System.Byte)(230)), ((System.Byte)(230)));
			this.btnConsultar.CausesValidation = false;
			this.btnConsultar.Enabled = false;
			this.btnConsultar.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.btnConsultar.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(50)), ((System.Byte)(50)), ((System.Byte)(50)));
			this.btnConsultar.Image = ((System.Drawing.Image)(resources.GetObject("btnConsultar.Image")));
			this.btnConsultar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnConsultar.Location = new System.Drawing.Point(95, 437);
			this.btnConsultar.Name = "btnConsultar";
			this.btnConsultar.Size = new System.Drawing.Size(80, 22);
			this.btnConsultar.TabIndex = 1008;
			this.btnConsultar.Text = "Consultar";
			this.btnConsultar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnConsultar.Click += new System.EventHandler(this.btnConsultar_Click);
			// 
			// lblContador
			// 
			this.lblContador.AutoSize = true;
			this.lblContador.Font = new System.Drawing.Font("Tahoma", 12F, System.Drawing.FontStyle.Bold);
			this.lblContador.ForeColor = System.Drawing.Color.Firebrick;
			this.lblContador.Location = new System.Drawing.Point(688, 48);
			this.lblContador.Name = "lblContador";
			this.lblContador.Size = new System.Drawing.Size(0, 23);
			this.lblContador.TabIndex = 1006;
			this.lblContador.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// label4
			// 
			this.label4.AutoSize = true;
			this.label4.Location = new System.Drawing.Point(16, 19);
			this.label4.Name = "label4";
			this.label4.Size = new System.Drawing.Size(44, 16);
			this.label4.TabIndex = 1000;
			this.label4.Text = "Nombre";
			this.label4.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label2
			// 
			this.label2.AutoSize = true;
			this.label2.Location = new System.Drawing.Point(400, 48);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(34, 16);
			this.label2.TabIndex = 1002;
			this.label2.Text = "Hasta";
			this.label2.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label1
			// 
			this.label1.AutoSize = true;
			this.label1.Location = new System.Drawing.Point(208, 48);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(37, 16);
			this.label1.TabIndex = 1001;
			this.label1.Text = "Desde";
			this.label1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtDescripcion
			// 
			appearance11.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtDescripcion.Appearance = appearance11;
			this.txtDescripcion.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtDescripcion.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtDescripcion.Enabled = false;
			this.txtDescripcion.Location = new System.Drawing.Point(72, 17);
			this.txtDescripcion.Multiline = true;
			this.txtDescripcion.Name = "txtDescripcion";
			this.txtDescripcion.Size = new System.Drawing.Size(312, 21);
			this.txtDescripcion.TabIndex = 999;
			// 
			// dtHasta
			// 
			appearance12.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtHasta.Appearance = appearance12;
			this.dtHasta.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton1.Caption = "Today";
			this.dtHasta.DateButtons.Add(dateButton1);
			this.dtHasta.Enabled = false;
			this.dtHasta.Format = "dd/MM/yyyy ";
			this.dtHasta.Location = new System.Drawing.Point(448, 48);
			this.dtHasta.Name = "dtHasta";
			this.dtHasta.NonAutoSizeHeight = 23;
			this.dtHasta.Size = new System.Drawing.Size(124, 21);
			this.dtHasta.SpinButtonsVisible = true;
			this.dtHasta.TabIndex = 1004;
			this.dtHasta.Value = ((object)(resources.GetObject("dtHasta.Value")));
			// 
			// dtDesde
			// 
			appearance13.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtDesde.Appearance = appearance13;
			this.dtDesde.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton2.Caption = "Today";
			this.dtDesde.DateButtons.Add(dateButton2);
			this.dtDesde.Enabled = false;
			this.dtDesde.Format = "dd/MM/yyyy ";
			this.dtDesde.Location = new System.Drawing.Point(264, 48);
			this.dtDesde.Name = "dtDesde";
			this.dtDesde.NonAutoSizeHeight = 23;
			this.dtDesde.Size = new System.Drawing.Size(124, 21);
			this.dtDesde.SpinButtonsVisible = true;
			this.dtDesde.TabIndex = 1003;
			this.dtDesde.Value = ((object)(resources.GetObject("dtDesde.Value")));
			// 
			// btnImportar
			// 
			this.btnImportar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btnImportar.CausesValidation = false;
			this.btnImportar.Enabled = false;
			this.btnImportar.Image = ((System.Drawing.Image)(resources.GetObject("btnImportar.Image")));
			this.btnImportar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnImportar.Location = new System.Drawing.Point(824, 16);
			this.btnImportar.Name = "btnImportar";
			this.btnImportar.Size = new System.Drawing.Size(80, 23);
			this.btnImportar.TabIndex = 1020;
			this.btnImportar.Text = "Importar";
			this.btnImportar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnImportar.Click += new System.EventHandler(this.btnImportar_Click);
			// 
			// btnTipoBono
			// 
			this.btnTipoBono.CausesValidation = false;
			this.btnTipoBono.Enabled = false;
			this.btnTipoBono.Image = ((System.Drawing.Image)(resources.GetObject("btnTipoBono.Image")));
			this.btnTipoBono.Location = new System.Drawing.Point(616, 16);
			this.btnTipoBono.Name = "btnTipoBono";
			this.btnTipoBono.Size = new System.Drawing.Size(24, 22);
			this.btnTipoBono.TabIndex = 1024;
			this.btnTipoBono.ToolTip = "Grupo de artículos";
			this.btnTipoBono.Visible = false;
			this.btnTipoBono.Click += new System.EventHandler(this.btnTipoBono_Click);
			// 
			// label3
			// 
			this.label3.AutoSize = true;
			this.label3.BackColor = System.Drawing.SystemColors.Control;
			this.label3.Location = new System.Drawing.Point(392, 19);
			this.label3.Name = "label3";
			this.label3.Size = new System.Drawing.Size(56, 16);
			this.label3.TabIndex = 1023;
			this.label3.Text = "Tipo Bono";
			this.label3.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
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
			// cmbTipoBono
			// 
			appearance14.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbTipoBono.Appearance = appearance14;
			this.cmbTipoBono.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.cmbTipoBono.Cursor = System.Windows.Forms.Cursors.Default;
			ultraGridColumn19.Header.VisiblePosition = 0;
			ultraGridColumn19.Hidden = true;
			ultraGridColumn20.Header.VisiblePosition = 1;
			ultraGridColumn20.Width = 150;
			ultraGridBand2.Columns.AddRange(new object[] {
																										 ultraGridColumn19,
																										 ultraGridColumn20});
			this.cmbTipoBono.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			this.cmbTipoBono.DisplayMember = "Nombre";
			this.cmbTipoBono.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbTipoBono.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbTipoBono.Enabled = false;
			this.cmbTipoBono.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbTipoBono.Location = new System.Drawing.Point(456, 17);
			this.cmbTipoBono.Name = "cmbTipoBono";
			this.cmbTipoBono.Size = new System.Drawing.Size(150, 21);
			this.cmbTipoBono.TabIndex = 1026;
			this.cmbTipoBono.ValueMember = "IdTipoBono";
			// 
			// txtidMaestroSubida
			// 
			this.txtidMaestroSubida.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtidMaestroSubida.Location = new System.Drawing.Point(1053, 444);
			this.txtidMaestroSubida.MaxValue = 3600;
			this.txtidMaestroSubida.MinValue = 0;
			this.txtidMaestroSubida.Name = "txtidMaestroSubida";
			this.txtidMaestroSubida.PromptChar = ' ';
			this.txtidMaestroSubida.Size = new System.Drawing.Size(56, 21);
			this.txtidMaestroSubida.TabIndex = 1027;
			this.txtidMaestroSubida.Visible = false;
			// 
			// btnAgregar
			// 
			this.btnAgregar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btnAgregar.CausesValidation = false;
			this.btnAgregar.Enabled = false;
			this.btnAgregar.Image = ((System.Drawing.Image)(resources.GetObject("btnAgregar.Image")));
			this.btnAgregar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnAgregar.Location = new System.Drawing.Point(728, 16);
			this.btnAgregar.Name = "btnAgregar";
			this.btnAgregar.Size = new System.Drawing.Size(88, 23);
			this.btnAgregar.TabIndex = 1030;
			this.btnAgregar.Text = "Agregar";
			this.btnAgregar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnAgregar.Visible = false;
			this.btnAgregar.Click += new System.EventHandler(this.btnAgregar_Click);
			// 
			// btnAcreditacion
			// 
			this.btnAcreditacion.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btnAcreditacion.CausesValidation = false;
			this.btnAcreditacion.Enabled = false;
			this.btnAcreditacion.Image = ((System.Drawing.Image)(resources.GetObject("btnAcreditacion.Image")));
			this.btnAcreditacion.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnAcreditacion.Location = new System.Drawing.Point(912, 16);
			this.btnAcreditacion.Name = "btnAcreditacion";
			this.btnAcreditacion.Size = new System.Drawing.Size(88, 23);
			this.btnAcreditacion.TabIndex = 1029;
			this.btnAcreditacion.Text = "Acreditación";
			this.btnAcreditacion.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnAcreditacion.Click += new System.EventHandler(this.btnAcreditacion_Click);
			// 
			// btnExportar
			// 
			this.btnExportar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btnExportar.CausesValidation = false;
			this.btnExportar.Enabled = false;
			this.btnExportar.Image = ((System.Drawing.Image)(resources.GetObject("btnExportar.Image")));
			this.btnExportar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnExportar.Location = new System.Drawing.Point(1008, 16);
			this.btnExportar.Name = "btnExportar";
			this.btnExportar.Size = new System.Drawing.Size(88, 23);
			this.btnExportar.TabIndex = 1028;
			this.btnExportar.Text = "Exportar";
			this.btnExportar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnExportar.Click += new System.EventHandler(this.btnExportar_Click);
			// 
			// label10
			// 
			this.label10.AutoSize = true;
			this.label10.BackColor = System.Drawing.Color.Transparent;
			this.label10.Location = new System.Drawing.Point(16, 48);
			this.label10.Name = "label10";
			this.label10.Size = new System.Drawing.Size(36, 16);
			this.label10.TabIndex = 1032;
			this.label10.Text = "Fecha";
			this.label10.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// dtpFecha
			// 
			this.dtpFecha.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton3.Caption = "Today";
			this.dtpFecha.DateButtons.Add(dateButton3);
			this.dtpFecha.Enabled = false;
			this.dtpFecha.Format = "dd/MM/yyyy";
			this.dtpFecha.Location = new System.Drawing.Point(80, 48);
			this.dtpFecha.Name = "dtpFecha";
			this.dtpFecha.NonAutoSizeHeight = 23;
			this.dtpFecha.Size = new System.Drawing.Size(104, 21);
			this.dtpFecha.TabIndex = 1031;
			this.dtpFecha.Value = ((object)(resources.GetObject("dtpFecha.Value")));
			// 
			// mnuAcreditacion
			// 
			this.mnuAcreditacion.MenuItems.AddRange(new System.Windows.Forms.MenuItem[] {
																																										this.mnuAcreditacionTransferencias,
																																										this.mnuAcreditacionCheques});
			this.mnuAcreditacion.Popup += new System.EventHandler(this.mnuAcreditacion_Popup);
			// 
			// mnuAcreditacionTransferencias
			// 
			this.mnuAcreditacionTransferencias.Index = 0;
			this.mnuAcreditacionTransferencias.Text = "Transferencias";
			this.mnuAcreditacionTransferencias.Click += new System.EventHandler(this.mnuAcreditacionTransferencias_Click);
			// 
			// mnuAcreditacionCheques
			// 
			this.mnuAcreditacionCheques.Index = 1;
			this.mnuAcreditacionCheques.Text = "Cheques";
			this.mnuAcreditacionCheques.Click += new System.EventHandler(this.mnuAcreditacionCheques_Click);
			// 
			// lblEstado2
			// 
			this.lblEstado2.AutoSize = true;
			this.lblEstado2.Font = new System.Drawing.Font("Tahoma", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblEstado2.ForeColor = System.Drawing.Color.Firebrick;
			this.lblEstado2.Location = new System.Drawing.Point(950, 48);
			this.lblEstado2.Name = "lblEstado2";
			this.lblEstado2.Size = new System.Drawing.Size(0, 23);
			this.lblEstado2.TabIndex = 1033;
			this.lblEstado2.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// btnAnular
			// 
			this.btnAnular.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnAnular.CausesValidation = false;
			this.btnAnular.Enabled = false;
			this.btnAnular.Image = ((System.Drawing.Image)(resources.GetObject("btnAnular.Image")));
			this.btnAnular.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnAnular.Location = new System.Drawing.Point(440, 437);
			this.btnAnular.Name = "btnAnular";
			this.btnAnular.Size = new System.Drawing.Size(76, 22);
			this.btnAnular.TabIndex = 1034;
			this.btnAnular.Text = "&Anular";
			this.btnAnular.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnAnular.Click += new System.EventHandler(this.btnAnular_Click);
			// 
			// frmComicionesSemanales
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(1111, 469);
			this.Controls.Add(this.btnAnular);
			this.Controls.Add(this.lblEstado2);
			this.Controls.Add(this.label10);
			this.Controls.Add(this.label3);
			this.Controls.Add(this.label4);
			this.Controls.Add(this.label2);
			this.Controls.Add(this.label1);
			this.Controls.Add(this.dtpFecha);
			this.Controls.Add(this.btnAgregar);
			this.Controls.Add(this.btnAcreditacion);
			this.Controls.Add(this.btnExportar);
			this.Controls.Add(this.txtidMaestroSubida);
			this.Controls.Add(this.cmbTipoBono);
			this.Controls.Add(this.btnTipoBono);
			this.Controls.Add(this.btnImportar);
			this.Controls.Add(this.btnSalir);
			this.Controls.Add(this.uGridInformacion);
			this.Controls.Add(this.groupBox1);
			this.Controls.Add(this.groupBox2);
			this.Controls.Add(this.lblEstado);
			this.Controls.Add(this.txtestado);
			this.Controls.Add(this.txtidComicionesSemanales);
			this.Controls.Add(this.btnEditar);
			this.Controls.Add(this.btnGuardar);
			this.Controls.Add(this.btnCancelar);
			this.Controls.Add(this.btnNuevo);
			this.Controls.Add(this.btnConsultar);
			this.Controls.Add(this.lblContador);
			this.Controls.Add(this.txtDescripcion);
			this.Controls.Add(this.dtHasta);
			this.Controls.Add(this.dtDesde);
			this.Name = "frmComicionesSemanales";
			this.Text = "Comisiones Semanales";
			this.Load += new System.EventHandler(this.frmComicionesSemanales_Load);
			((System.ComponentModel.ISupportInitialize)(this.uGridInformacion)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource4)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtestado)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtidComicionesSemanales)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDescripcion)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtHasta)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtDesde)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbTipoBono)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtidMaestroSubida)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtpFecha)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private Acceso miAcceso;
		bool bNuevo = false;
		bool bEdicion = false;

		private void frmComicionesSemanales_Load(object sender, System.EventArgs e)
		{
			miAcceso = new Acceso(cdsSeteoF, "0526");

			if (!Funcion.Permiso("1787", cdsSeteoF))
			{				
				MessageBox.Show("No puede ingresar a Bonos Semanales", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				BeginInvoke(new MethodInvoker(UnloadMe));
				return;
			}
			if (miAcceso.Nuevo) this.btnNuevo.Enabled = true;
			if (miAcceso.Buscar) this.btnConsultar.Enabled = true;
			this.dtpFecha.Value = FuncionesProcedimientos.dtRetornaFechaServidor(cdsSeteoF);	

			this.cmbTipoBono.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Select IdTipoBono, Nombre From TipoBono Order By Nombre");

		}
		private void Estado ( int IEstado)
		{
			this.lblContador.Text = "";

			if (IEstado == 0)
			{
				this.lblEstado2.Text = "PENDIENTE";
			}
			
			if (IEstado == 1)
			{
				this.lblEstado2.Text = "GENERADO";
			}

			if (IEstado == 2)
			{
				this.lblEstado2.Text = "ACREDITADO";
			}
			if (IEstado == 3)
			{
				this.lblEstado2.Text = "ANULADO";
			}

		}


		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);
		}

		private void btnTipoBono_Click(object sender, System.EventArgs e)
		{
			using (frmTipoBonos TipoBonos = new frmTipoBonos())
			{
				if (DialogResult.OK == TipoBonos.ShowDialog())
				{
					this.cmbTipoBono.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Select IdTipoBono, Nombre From TipoBono Order By Nombre");
				}
			}
		}

		private void btnCancelar_Click(object sender, System.EventArgs e)
		{
			bNuevo = false;
			bEdicion = false;

			#region Vacia Campos 

			this.txtidComicionesSemanales.Value = 0;
			this.txtidMaestroSubida.Value = 0;
			this.txtDescripcion.Text = "";
			this.dtDesde.Value = System.DBNull.Value;
			this.dtHasta.Value = System.DBNull.Value;
			this.cmbTipoBono.Value = System.DBNull.Value;
			this.txtestado.Value = 0;
			this.lblContador.Text = "";
			this.lblEstado2.Text="";

		
			#endregion Vacia Campos 

			#region Bloquea Campo

			this.txtDescripcion.Enabled = false;
			this.dtDesde.Enabled = false;
			this.dtHasta.Enabled = false;
			this.cmbTipoBono.Enabled = false;
	
			FuncionesProcedimientos.EstadoGrids(false, this.uGridInformacion);

			#endregion Bloquea Campo

			this.uGridInformacion.DataSource = Funcion.dvProcedimiento(cdsSeteoF, string.Format("Exec ConsultaIndividualDetalleBonoSemanal 0"));
			this.uGridInformacion.DisplayLayout.Bands[0].Columns["Agregar"].CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;

			if (miAcceso.Nuevo) this.btnNuevo.Enabled = true;
			if (miAcceso.Buscar) this.btnConsultar.Enabled = true;
			if (miAcceso.BImportar)   this.btnImportar.Enabled = true;
			if (miAcceso.BAcreditar) this.btnAcreditacion.Enabled = true;
			this.btnEditar.Enabled = false;
			this.btnGuardar.Enabled = false;
			this.btnCancelar.Enabled = false;
			this.btnAgregar.Enabled = true;			
			this.btnTipoBono.Enabled = false;

		}

		private void btnNuevo_Click(object sender, System.EventArgs e)
		{
			this.btnCancelar_Click(sender, e);

			this.txtDescripcion.Enabled = true;
			this.dtHasta.Enabled = true;
			this.dtDesde.Enabled = true;

			FuncionesProcedimientos.EstadoGrids(true, this.uGridInformacion);

			Estado ((int) this.txtestado.Value);

			this.btnNuevo.Enabled = false;
			this.btnConsultar.Enabled = false;
			this.btnEditar.Enabled = false;
			this.btnGuardar.Enabled = true;
			this.btnCancelar.Enabled = true;		
			this.cmbTipoBono.Enabled = true;
			if (miAcceso.BTipoBono) 
			{
				this.btnTipoBono.Visible = true;	
				this.btnTipoBono.Enabled = true;
			}
			
			bNuevo = true;
			bEdicion = true;

			this.lblContador.Text = this.uGridInformacion.Rows.Count + " REGISTROS ENCONTRADOS";

			this.txtDescripcion.Focus();	
		}

		private void btnImportar_Click(object sender, System.EventArgs e)
		{
			int iCont = 0;

			#region Validacion
			if (!Validacion.vbTexto(this.txtDescripcion, 5, 100, "Descripción")) return;
			if (!Validacion.vbFechaEnDocumentos(this.dtDesde, "Ingrese la fecha Desde ", true, "COMISIONES SEMANALES", cdsSeteoF)) return;
			if (!Validacion.vbFechaEnDocumentos(this.dtHasta, "Ingrese la fecha Hasta ", true, "COMISIONES SEMANALES", cdsSeteoF)) return;
			if (!Validacion.vbComparaFechas(Convert.ToDateTime(this.dtDesde.Value), Convert.ToDateTime(this.dtHasta.Value), this.dtDesde, "La Fecha desde no puede ser mayor a la Fecha hasta", false)) return;
			if (!Validacion.vbComboVacio(this.cmbTipoBono, "Seleccione un Bono")) return;
			
			if (Funcion.bEscalarSQL(cdsSeteoF, string.Format("Exec NominaValidaRolPagado '{0}'", Convert.ToDateTime(this.dtpFecha.Value).ToString("yyyyMMdd")), true))
			{
				MessageBox.Show("El Rol Del Periodo se encuentra Generado", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return;
			}
			#endregion Validacion
			this.uGridInformacion.DataSource = Funcion.dvProcedimiento(cdsSeteoF, string.Format("Exec ConsultaIndividualDetalleBonoSemanal 0"));
			this.uGridInformacion.DisplayLayout.Bands[0].Columns["Agregar"].CellActivation = Infragistics.Win.UltraWinGrid.Activation.AllowEdit;

			try
			{
				this.txtestado.Value = 1;
				OpenFileDialog sArchivo = new OpenFileDialog();

				sArchivo.Filter = "Microsoft Excel (*.xls)|*.xls";

				if (sArchivo.ShowDialog() == System.Windows.Forms.DialogResult.OK && sArchivo.FileName.Length > 0)
				{
					this.Cursor = Cursors.WaitCursor;                                  
					
					iCont = 0;

					#region Guarda Maestro
					string sSQLl = string.Format("Exec GuardaMaestroSubidaBonos {0}, '{1}' , '{2}' ,'{3}', {4},'{5}',{6}",
						(int)this.txtidMaestroSubida.Value,
						this.txtDescripcion.Text,
						Convert.ToDateTime(this.dtDesde.Value).ToString("yyyyMMdd"),
						Convert.ToDateTime(this.dtHasta.Value).ToString("yyyyMMdd"),
						(int)this.txtestado.Value,Convert.ToDateTime(this.dtpFecha.Value).ToString("yyyyMMdd"),(int)this.cmbTipoBono.Value);
					this.txtidMaestroSubida.Value=Funcion.iEscalarSQL(cdsSeteoF,sSQLl);
					#endregion Guarda Maestro

					#region Sube archivo
					foreach (DataRow row in FuncionesProcedimientos.ImportaDeExcel(sArchivo.FileName, "SELECT * FROM [DATOS$]", "DATOS").Tables[0].Rows)
					{						
						#region Guarda Informacion
						string sSQLArchivo = string.Format("Exec GuardaDetalleSubidaBonos '{0}', {1}, {2}",row["NOMBRE"].ToString(),decimal.Parse(row["VALOR"].ToString()),(int)this.txtidMaestroSubida.Value);
						Funcion.EjecutaSQL(cdsSeteoF, sSQLArchivo);
						#endregion Guarda Informacion

						iCont++;
					}				
					#endregion Sube archivo

					#region Genera Bonos  
					string sSQL = string.Format("Exec GeneraBonos {0}",(int)this.txtidMaestroSubida.Value);
					FuncionesProcedimientos.dsGeneral(sSQL,uGridInformacion);
					this.lblContador.Text = this.uGridInformacion.Rows.Count + " REGISTROS ENCONTRADOS";
					Estado ((int) this.txtestado.Value);
					this.txtDescripcion.Enabled = false;
					this.dtDesde.Enabled = false;
					this.dtHasta.Enabled = false;
					this.cmbTipoBono.Enabled = false;

					FuncionesProcedimientos.EstadoGrids(false,this.uGridInformacion);
	

					MessageBox.Show(string.Format("Proceso Finalizado Correctamente"), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
					#endregion Genera Bonos  

					using (frmValidaSubidaBonos RAC = new frmValidaSubidaBonos((int)this.txtidMaestroSubida.Value))
					{
						if (DialogResult.OK == RAC.ShowDialog())
						{
		
						}
					}	
					this.btnEditar.Enabled=true;

	
					this.Cursor = Cursors.Default;
				}
			}
			catch (Exception ex)
			{
				MessageBox.Show("Fila : " + iCont.ToString());
				MessageBox.Show(string.Format("Commit Exception Type : {0} {1}", ex.GetType(), ex.Message), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
				this.Cursor = Cursors.Default;
				return;
			}
		}

		private void btnGuardar_Click(object sender, System.EventArgs e)
		{
			#region Validacion
			if (!Validacion.vbTexto(this.txtDescripcion, 5, 100, "Descripción")) return;
			if (!Validacion.vbFechaEnDocumentos(this.dtDesde, "Ingrese la fecha Desde ", true, "COMISIONES SEMANALES", cdsSeteoF)) return;
			if (!Validacion.vbFechaEnDocumentos(this.dtHasta, "Ingrese la fecha Hasta ", true, "COMISIONES SEMANALES", cdsSeteoF)) return;
			if (!Validacion.vbComparaFechas(Convert.ToDateTime(this.dtDesde.Value), Convert.ToDateTime(this.dtHasta.Value), this.dtDesde, "La Fecha desde no puede ser mayor a la Fecha hasta", false)) return;
			if (!Validacion.vbComboVacio(this.cmbTipoBono, "Seleccione un Bono")) return;
			
			if (Funcion.bEscalarSQL(cdsSeteoF, string.Format("Exec NominaValidaRolPagado '{0}'", Convert.ToDateTime(this.dtpFecha.Value).ToString("yyyyMMdd")), true))
			{
				MessageBox.Show("El Rol Del Periodo se encuentra Generado", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return;
			}
			if ( (int)this.uGridInformacion.Rows.Count == 0)
			{
				MessageBox.Show("Ingrese registros ", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return;
			}		
			#endregion Validacion

			#region Valida Proteccion vacios
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow drr in this.uGridInformacion.Rows.All)
			{
				if (Convert.ToDecimal(drr.Cells["Bono"].Value) <= 0 )
				{
					this.uGridInformacion.ActiveRow = drr;
					this.uGridInformacion.ActiveRow.Selected = true;
					MessageBox.Show("El campo Bono no puede ser 0 ", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					return;
				}
			}
			#endregion Valida Proteccion vacios

			#region Guarda detalle 
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow drr in this.uGridInformacion.Rows.All)
			{	
				string sSQLDetalle = string.Format("Exec GuardaDetalleBonoSemanales  0,{0},{1},{2},'{3}','{4}','{5}',{6},'{7}',{8},'{9}','{10}',{11},'{12}',{13},'{14}','{15}',{16},{17}", 
					/* 0  */(int)drr.Cells["N"].Value,
					/* 1  */(int)this.txtidMaestroSubida.Value,
					/* 2  */(int)drr.Cells["idNomina"].Value,
					/* 3  */(string)drr.Cells["Cedula"].Value.ToString(),
					/* 4  */(string)drr.Cells["Apellidos"].Value.ToString(),
					/* 5  */(string)drr.Cells["Nombres"].Value.ToString(),
					/* 6  */(int)drr.Cells["TipoDePago"].Value,
					/* 7  */(string)drr.Cells["TipoDepagoM"].Value.ToString(),
					/* 8  */(int)drr.Cells["idBanco"].Value,
					/* 9  */(string)drr.Cells["Banco"].Value.ToString(),
					/* 10 */(string)drr.Cells["NumeroDeCuenta"].Value.ToString(),
					/* 11 */(int)drr.Cells["idProyecto"].Value,
					/* 12 */(string)drr.Cells["Centro_De_Costo"].Value.ToString(),
					/* 13 */(int)drr.Cells["idDepartamento"].Value,
					/* 14 */(string)drr.Cells["Departamento"].Value.ToString(),
					/* 15 */(string)drr.Cells["Cargo"].Value.ToString(),
					/* 16 */(decimal)drr.Cells["Bono"].Value,
					/* 17 */(int)drr.Cells["TipoBanco"].Value);
				Funcion.EjecutaSQL(cdsSeteoF, sSQLDetalle);
				
			}
			FuncionesProcedimientos.EstadoGrids(false,this.uGridInformacion);
			this.txtDescripcion.Enabled = false;
			this.dtDesde.Enabled = false;
			this.dtHasta.Enabled = false;
			this.cmbTipoBono.Enabled = false;
			this.btnGuardar.Enabled=false;
			this.btnImportar.Enabled=false;

			#endregion Guarda detalle 

		}

		private void btnAcreditacion_Click(object sender, System.EventArgs e)
		{
			this.mnuAcreditacion.Show(this.btnAcreditacion, new Point(1, 25));
		}

		private void btnExportar_Click(object sender, System.EventArgs e)
		{
			FuncionesProcedimientos.ExportaExcel(this.uGridInformacion);
		}

		private void mnuAcreditacionTransferencias_Click(object sender, System.EventArgs e)
		{
			DateTime dtFecha = (DateTime)this.dtpFecha.Value;

			using (frmRolAcreditacion RAC = new frmRolAcreditacion((int)this.txtidMaestroSubida.Value, 0, dtFecha, 4))
			{
				if (DialogResult.OK == RAC.ShowDialog())
				{
		
				}
			}		
		}

		private void mnuAcreditacionCheques_Click(object sender, System.EventArgs e)
		{
			DateTime dtFecha = (DateTime)this.dtpFecha.Value;

			using (frmRolAutorizacionPagoCheques RAPC = new frmRolAutorizacionPagoCheques((int)this.txtidMaestroSubida.Value, dtFecha, 4))
			{
				if (DialogResult.OK == RAPC.ShowDialog())
				{
		
				}
			}
		}

		private void btnConsultar_Click(object sender, System.EventArgs e)
		{
			this.btnCancelar_Click(sender, e);

			using (frmConsultaBonoSemananal consultab = new frmConsultaBonoSemananal())
				if (DialogResult.OK == consultab.ShowDialog())
					this.Consultar((int)consultab.uGridInformacion.ActiveRow.Cells["IdMaestroSubidaDatosBonos"].Value);
		}
		private void Consultar (int IdMaestroSubidaDatosBonos)
		{
			#region Consulta Individual
			string ssql= string.Format ("Exec ConsultaIndividualMaestroBonos {0}",IdMaestroSubidaDatosBonos); 
			SqlDataReader dr = Funcion.rEscalarSQL(cdsSeteoF,ssql,true);
			dr.Read();
			this.txtidMaestroSubida.Value= dr.GetInt32 (0);
			if (dr.GetValue(1) != System.DBNull.Value)this.txtDescripcion.Text = dr.GetString (1);
			if (dr.GetValue(2) != System.DBNull.Value)this.dtDesde.Value = dr.GetDateTime (2);
			if (dr.GetValue(3) != System.DBNull.Value)this.dtHasta.Value = dr.GetDateTime (3);
			if (dr.GetValue(4) != System.DBNull.Value)this.txtestado.Value = dr.GetInt32 (4);		
			if (dr.GetValue(5) != System.DBNull.Value)this.dtpFecha.Value = dr.GetDateTime (5);
			if (dr.GetValue(6) != System.DBNull.Value)this.cmbTipoBono.Value = dr.GetInt32 (6);
			dr.Close();

			this.uGridInformacion.DataSource = Funcion.dvProcedimiento(cdsSeteoF, string.Format("Exec ConsultaIndividualDetalleBonoSemanal {0}", IdMaestroSubidaDatosBonos));
			#endregion Consulta Individual

			Estado ((int)this.txtestado.Value);


			this.btnNuevo.Enabled = true;
			this.btnConsultar.Enabled = true;
			this.btnExportar.Enabled = true;

			if ((int)this.txtestado.Value == 1) 
			{
				this.btnImportar.Enabled = false;
				this.btnEditar.Enabled = true;
				this.btnAnular.Enabled=true;
			}
			if ((int)this.txtestado.Value == 2) 
			{
				this.btnImportar.Enabled = false;
				this.btnAcreditacion.Enabled = false;
				this.btnEditar.Enabled = true;
				this.btnAnular.Enabled=false;
			}
			if ((int)this.txtestado.Value == 3) 
			{
				this.btnImportar.Enabled = false;
				this.btnAcreditacion.Enabled = false;
				this.btnEditar.Enabled = true;
				this.btnAnular.Enabled=false;
				this.btnExportar.Enabled=false;
			}
			this.btnCancelar.Enabled = true;

			this.lblContador.Text = this.uGridInformacion.Rows.Count + " REGISTROS ENCONTRADOS";
		}

		private void btnAgregar_Click(object sender, System.EventArgs e)
		{
		
		}

		private void btnSalir_Click(object sender, System.EventArgs e)
		{
					this.Close();
		}

		private void btnAnular_Click(object sender, System.EventArgs e)
		{
			if ((int)this.txtestado.Value  == 2)
			{
				MessageBox.Show("El Documento ya esta acreditado.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);				
				return;
			}
			string Nota ="";
			bEdicion=true;
			using (frmAnulaBonosSemanales miNota = new frmAnulaBonosSemanales(bEdicion))
			{
				if (DialogResult.OK == miNota.ShowDialog())
				{			
						Nota = miNota.txtObservaciones.Text.ToString().Trim();
					Funcion.EjecutaSQL(cdsSeteoF, string.Format("Exec AnulaBonosSemanales {0}, '{1}'", (int)this.txtidMaestroSubida.Value,Nota.ToString() ), true);
					this.Consultar((int)this.txtidMaestroSubida.Value);
										
				}
			}	
		}

		private void uGridInformacion_BeforeCellUpdate(object sender, Infragistics.Win.UltraWinGrid.BeforeCellUpdateEventArgs e)
		{

		}

		private void uGridInformacion_ClickCellButton(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			if (bEdicion)	
			{
				if (e.Cell.Column.ToString() == "Agregar")
				{
					if ((int)e.Cell.Row.Cells["idNomina"].Value > 0) 
					{
						//this.uGridInformacion.DisplayLayout.Bands[0].Columns["idNomina"].CellActivation = Activation.NoEdit;

						MessageBox.Show("No se puede remplazar datos cargados registradados", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
						return;
					}

				}
	
					///this.uGridInformacion.DisplayLayout.Bands[0].Columns["Agregar"].CellActivation = Infragistics.Win.UltraWinGrid.Activation.AllowEdit;
					this.uGridInformacion.DisplayLayout.Bands[0].Columns["Cedula"].CellActivation = Infragistics.Win.UltraWinGrid.Activation.AllowEdit;
					this.uGridInformacion.DisplayLayout.Bands[0].Columns["Bono"].CellActivation = Infragistics.Win.UltraWinGrid.Activation.AllowEdit;
		
			}
		}

		private void btnEditar_Click(object sender, System.EventArgs e)
		{
			if ((int)this.txtestado.Value  == 2)
			{
				MessageBox.Show("El Documento ya esta acreditado.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);				
				return;
			}
			if ((int)this.txtestado.Value  == 3)
			{
				MessageBox.Show("El Documento ya esta anulado.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);				
				return;
			}
			FuncionesProcedimientos.EstadoGrids(true,this.uGridInformacion);
			bEdicion=true;

			this.uGridInformacion.DisplayLayout.Bands[0].Columns["Agregar"].CellActivation = Infragistics.Win.UltraWinGrid.Activation.AllowEdit;
		}

		private void uGridInformacion_AfterRowInsert(object sender, Infragistics.Win.UltraWinGrid.RowEventArgs e)
		{
			if (bEdicion)
			{
				e.Row.Cells["idNomina"].Value = 0;
//				e.Row.Cells["Cedula"].Value = "";
//				e.Row.Cells["Apellidos"].Value = "";
//				e.Row.Cells["Nombres"].Value = "" ;	
//				e.Row.Cells["TipoDePago"].Value = 0 ;	
//				e.Row.Cells["TipoDepagoM"].Value = "";
//
//				e.Row.Cells["idBanco"].Value = 0;
//				e.Row.Cells["Banco"].Value = "";
//				e.Row.Cells["NumeroDeCuenta"].Value = "";
//				e.Row.Cells["idProyecto"].Value = 0;
//				e.Row.Cells["Centro_De_Costo"].Value = "";
//				e.Row.Cells["idDepartamento"].Value = 0;
//
//				e.Row.Cells["Departamento"].Value = "";
//				e.Row.Cells["Cargo"].Value = "";
				e.Row.Cells["Bono"].Value = 0.00;
//				e.Row.Cells["TipoBanco"].Value = 0;
			}

		}

		private void uGridInformacion_AfterCellUpdate(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
//			if (e.Cell.Column.ToString() == "Cedula" || e.Cell.Column.ToString() == "Bono" )
//			{
//				if ((int)e.Cell.Row.Cells["idNomina"].Value > 0) 
//				{
////					//this.uGridInformacion.DisplayLayout.Bands[0].Columns["idNomina"].CellActivation = Activation.NoEdit;
////
////					MessageBox.Show("No se puede remplazar datos cargados registradados", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
////					return;
//
//					this.uGridInformacion.DisplayLayout.Bands[0].Columns["Cedula"].CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
//					this.uGridInformacion.DisplayLayout.Bands[0].Columns["Apellidos"].CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
//					this.uGridInformacion.DisplayLayout.Bands[0].Columns["Nombres"].CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
//					this.uGridInformacion.DisplayLayout.Bands[0].Columns["TipoDePagoM"].CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
//					this.uGridInformacion.DisplayLayout.Bands[0].Columns["Banco"].CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
//					this.uGridInformacion.DisplayLayout.Bands[0].Columns["NumeroDeCuenta"].CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
//					this.uGridInformacion.DisplayLayout.Bands[0].Columns["Centro_De_Costo"].CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
//					this.uGridInformacion.DisplayLayout.Bands[0].Columns["Departamento"].CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
//					this.uGridInformacion.DisplayLayout.Bands[0].Columns["Cargo"].CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
//					this.uGridInformacion.DisplayLayout.Bands[0].Columns["Bono"].CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
//					this.uGridInformacion.DisplayLayout.Bands[0].Columns["TipoBanco"].CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
//				}
//			}

			if (e.Cell.Column.ToString() == "Cedula")
			{
				#region Busca Producto
				string sSQLContador = string.Format("Exec BuscaEmpleadoBono '{0}'", e.Cell.Row.Cells["Cedula"].Value.ToString());
				int iContador = Funcion.iEscalarSQL(cdsSeteoF, sSQLContador);

				if (iContador == 1)
				{
					string sSQLInfo = string.Format("Exec InfoEmpleadoBonos '{0}'", e.Cell.Row.Cells["Cedula"].Value.ToString());
					SqlDataReader dr = Funcion.rEscalarSQL(cdsSeteoF, sSQLInfo, true);
					dr.Read();
					if (dr.HasRows)
					{
						e.Cell.Row.Cells["N"].Value = dr.GetInt32(0);
						e.Cell.Row.Cells["idNomina"].Value = dr.GetInt32(1);
						//e.Cell.Row.Cells["Codigo"].Value = dr.GetString(1);
						e.Cell.Row.Cells["Apellidos"].Value = dr.GetString(3);
						e.Cell.Row.Cells["Nombres"].Value = dr.GetString(4);
						e.Cell.Row.Cells["TipoDePago"].Value = dr.GetInt32(5);
						e.Cell.Row.Cells["TipoDePagoM"].Value = dr.GetString(6);
						e.Cell.Row.Cells["idBanco"].Value = dr.GetInt32(7);
						e.Cell.Row.Cells["Banco"].Value = dr.GetString(8);
						e.Cell.Row.Cells["NumeroDeCuenta"].Value = dr.GetString(9);

						e.Cell.Row.Cells["idProyecto"].Value = dr.GetInt32(10);
						e.Cell.Row.Cells["Centro_De_Costo"].Value = dr.GetString(11);
						e.Cell.Row.Cells["idDepartamento"].Value = dr.GetInt32(12);
						e.Cell.Row.Cells["Departamento"].Value = dr.GetString(13);
						e.Cell.Row.Cells["Cargo"].Value = dr.GetString(14);
						e.Cell.Row.Cells["Bono"].Value = dr.GetDecimal(15);
						e.Cell.Row.Cells["TipoBanco"].Value = dr.GetInt32(16);
		
					}
					dr.Close();

				}
				else 
				{
	
					using (frmConsultaNomina Busqueda = new frmConsultaNomina())
					{
						if (DialogResult.OK == Busqueda.ShowDialog())
						{
							this.uGridInformacion.ActiveRow.Cells["Cedula"].Value = Busqueda.uGrid.ActiveRow.Cells["NIdentificacion"].Value;
						}
					}

				}
				#endregion Busca Cuenta	
			}
		}

		private void groupBox1_Enter(object sender, System.EventArgs e)
		{
		
		}

		private void mnuAcreditacion_Popup(object sender, System.EventArgs e)
		{
		
		}

	}
}
