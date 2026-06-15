using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Globalization;
using System.Runtime.InteropServices;
using System.Diagnostics;
using System.IO;
using Infragistics.Win.UltraWinGrid;
using System.Data;
using System.Data.SqlClient;
using CrystalDecisions;
using CrystalDecisions.CrystalReports.Engine;
using CrystalDecisions.ReportSource;
using CrystalDecisions.Shared;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de frmReporteAuditoriaLocal1.
	/// </summary>
	public class frmReporteAuditoriaLocal1 : System.Windows.Forms.Form
	{
		private System.Windows.Forms.Label lblEstado;
		public Infragistics.Win.UltraWinEditors.UltraNumericEditor txtTotal;
		public Infragistics.Win.UltraWinEditors.UltraNumericEditor txtIdAuditoriaLocales;
		private System.Windows.Forms.Button btnConsultar;
		private System.Windows.Forms.Label label2;
		private System.Windows.Forms.Label label1;
		private System.Windows.Forms.Label label7;
		private System.Windows.Forms.Label lblBodega;
		private System.Windows.Forms.Label lblfecha;
		public Infragistics.Win.UltraWinGrid.UltraCombo cmbAuditor;
		public Infragistics.Win.UltraWinGrid.UltraCombo cmbVendedor;
		public Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtFecha;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource1;
		private C1.Data.C1DataSet cdsSeteoF;
		public System.Windows.Forms.Button btnEditar;
		public System.Windows.Forms.Button btnCancelar;
		public System.Windows.Forms.Button btnGuardar;
		public System.Windows.Forms.Button btnNuevo;
		private System.Windows.Forms.GroupBox groupBox3;
		public Infragistics.Win.UltraWinGrid.UltraGrid grdReporteAuditoria;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource2;
		int Bodegas = 0;
		public System.Windows.Forms.Button btnImprimir;
		public Infragistics.Win.UltraWinGrid.UltraCombo cmbBodega;
		DateTime dtFechas;
		int iEstadoConteo = 0;
		int Local = 0;

		public frmReporteAuditoriaLocal1(int iBodega, DateTime Fecha, int IEstadoConteo,int iLocal)
		{
			//
			// Necesario para admitir el Diseñador de Windows Forms
			//
			InitializeComponent();
			Bodegas = iBodega;
			dtFechas = Fecha;
			iEstadoConteo = IEstadoConteo;
			Local = iLocal;
     
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
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmReporteAuditoriaLocal1));
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPersonal");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPersonal");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Bodega");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Codigo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn3 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton1 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn4 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("IdDetalleAuditoria");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn5 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("IdAuditoriaLocales");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn6 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Descripcion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn7 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Puntaje");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn8 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("PuntajeObtenido");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn9 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Observaciones");
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand3 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("IdDetalleAuditoria");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("IdAuditoriaLocales");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Descripcion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Puntaje");
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn11 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("PuntajeObtenido");
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn12 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Observaciones");
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings1 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Puntaje", 3, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, null, -1, false);
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings2 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "PuntajeObtenido", 4, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, null, -1, false);
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance11 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance12 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance13 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance14 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance15 = new Infragistics.Win.Appearance();
			this.lblEstado = new System.Windows.Forms.Label();
			this.txtTotal = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtIdAuditoriaLocales = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.btnConsultar = new System.Windows.Forms.Button();
			this.label2 = new System.Windows.Forms.Label();
			this.label1 = new System.Windows.Forms.Label();
			this.label7 = new System.Windows.Forms.Label();
			this.lblBodega = new System.Windows.Forms.Label();
			this.lblfecha = new System.Windows.Forms.Label();
			this.cmbAuditor = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.cmbVendedor = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.ultraDataSource2 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.dtFecha = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.ultraDataSource1 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.btnEditar = new System.Windows.Forms.Button();
			this.btnCancelar = new System.Windows.Forms.Button();
			this.btnGuardar = new System.Windows.Forms.Button();
			this.btnNuevo = new System.Windows.Forms.Button();
			this.groupBox3 = new System.Windows.Forms.GroupBox();
			this.grdReporteAuditoria = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.btnImprimir = new System.Windows.Forms.Button();
			this.cmbBodega = new Infragistics.Win.UltraWinGrid.UltraCombo();
			((System.ComponentModel.ISupportInitialize)(this.txtTotal)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdAuditoriaLocales)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbAuditor)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbVendedor)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFecha)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.grdReporteAuditoria)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbBodega)).BeginInit();
			this.SuspendLayout();
			// 
			// lblEstado
			// 
			this.lblEstado.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.lblEstado.AutoSize = true;
			this.lblEstado.BackColor = System.Drawing.Color.Transparent;
			this.lblEstado.Font = new System.Drawing.Font("Tahoma", 10F, System.Drawing.FontStyle.Bold);
			this.lblEstado.ForeColor = System.Drawing.Color.Blue;
			this.lblEstado.Location = new System.Drawing.Point(600, 72);
			this.lblEstado.Name = "lblEstado";
			this.lblEstado.Size = new System.Drawing.Size(0, 20);
			this.lblEstado.TabIndex = 795;
			this.lblEstado.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtTotal
			// 
			this.txtTotal.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Right)));
			appearance1.BackColor = System.Drawing.Color.White;
			appearance1.ForeColor = System.Drawing.Color.Red;
			this.txtTotal.Appearance = appearance1;
			this.txtTotal.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtTotal.Enabled = false;
			this.txtTotal.Font = new System.Drawing.Font("Microsoft Sans Serif", 14.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtTotal.FormatString = "##0";
			this.txtTotal.Location = new System.Drawing.Point(504, 67);
			this.txtTotal.Name = "txtTotal";
			this.txtTotal.PromptChar = ' ';
			this.txtTotal.Size = new System.Drawing.Size(80, 31);
			this.txtTotal.SpinButtonAlignment = Infragistics.Win.ButtonAlignment.Left;
			this.txtTotal.TabIndex = 794;
			this.txtTotal.ValueChanged += new System.EventHandler(this.txtTotal_ValueChanged);
			// 
			// txtIdAuditoriaLocales
			// 
			this.txtIdAuditoriaLocales.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.txtIdAuditoriaLocales.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtIdAuditoriaLocales.Location = new System.Drawing.Point(760, 8);
			this.txtIdAuditoriaLocales.Name = "txtIdAuditoriaLocales";
			this.txtIdAuditoriaLocales.PromptChar = ' ';
			this.txtIdAuditoriaLocales.Size = new System.Drawing.Size(40, 22);
			this.txtIdAuditoriaLocales.SpinButtonAlignment = Infragistics.Win.ButtonAlignment.Left;
			this.txtIdAuditoriaLocales.TabIndex = 793;
			this.txtIdAuditoriaLocales.Visible = false;
			// 
			// btnConsultar
			// 
			this.btnConsultar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnConsultar.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(230)), ((System.Byte)(230)), ((System.Byte)(230)));
			this.btnConsultar.CausesValidation = false;
			this.btnConsultar.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.btnConsultar.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(50)), ((System.Byte)(50)), ((System.Byte)(50)));
			this.btnConsultar.Image = ((System.Drawing.Image)(resources.GetObject("btnConsultar.Image")));
			this.btnConsultar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnConsultar.Location = new System.Drawing.Point(680, 472);
			this.btnConsultar.Name = "btnConsultar";
			this.btnConsultar.Size = new System.Drawing.Size(40, 23);
			this.btnConsultar.TabIndex = 792;
			this.btnConsultar.Text = "Consultar";
			this.btnConsultar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnConsultar.Visible = false;
			this.btnConsultar.Click += new System.EventHandler(this.btnConsultar_Click);
			// 
			// label2
			// 
			this.label2.AutoSize = true;
			this.label2.Enabled = false;
			this.label2.Location = new System.Drawing.Point(400, 74);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(92, 17);
			this.label2.TabIndex = 791;
			this.label2.Text = "Calificacion Total:";
			// 
			// label1
			// 
			this.label1.AutoSize = true;
			this.label1.Enabled = false;
			this.label1.Location = new System.Drawing.Point(8, 104);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(44, 17);
			this.label1.TabIndex = 790;
			this.label1.Text = "Auditor:";
			// 
			// label7
			// 
			this.label7.AutoSize = true;
			this.label7.Enabled = false;
			this.label7.Location = new System.Drawing.Point(8, 72);
			this.label7.Name = "label7";
			this.label7.Size = new System.Drawing.Size(58, 17);
			this.label7.TabIndex = 788;
			this.label7.Text = "Jefe Local:";
			// 
			// lblBodega
			// 
			this.lblBodega.AutoSize = true;
			this.lblBodega.Enabled = false;
			this.lblBodega.Location = new System.Drawing.Point(8, 40);
			this.lblBodega.Name = "lblBodega";
			this.lblBodega.Size = new System.Drawing.Size(34, 17);
			this.lblBodega.TabIndex = 785;
			this.lblBodega.Text = "Local:";
			this.lblBodega.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// lblfecha
			// 
			this.lblfecha.AutoSize = true;
			this.lblfecha.Enabled = false;
			this.lblfecha.Location = new System.Drawing.Point(8, 8);
			this.lblfecha.Name = "lblfecha";
			this.lblfecha.Size = new System.Drawing.Size(34, 17);
			this.lblfecha.TabIndex = 783;
			this.lblfecha.Text = "Fecha";
			this.lblfecha.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// cmbAuditor
			// 
			appearance2.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbAuditor.Appearance = appearance2;
			this.cmbAuditor.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbAuditor.Cursor = System.Windows.Forms.Cursors.Default;
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridColumn2.Width = 335;
			ultraGridColumn3.Header.VisiblePosition = 2;
			ultraGridBand1.Columns.AddRange(new object[] {
																										 ultraGridColumn1,
																										 ultraGridColumn2,
																										 ultraGridColumn3});
			this.cmbAuditor.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.cmbAuditor.DisplayMember = "Nombre";
			this.cmbAuditor.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbAuditor.Enabled = false;
			this.cmbAuditor.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbAuditor.Location = new System.Drawing.Point(80, 104);
			this.cmbAuditor.Name = "cmbAuditor";
			this.cmbAuditor.Size = new System.Drawing.Size(250, 21);
			this.cmbAuditor.TabIndex = 789;
			this.cmbAuditor.ValueMember = "idPersonal";
			this.cmbAuditor.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cmbAuditor_KeyDown);
			this.cmbAuditor.ValueChanged += new System.EventHandler(this.cmbAuditor_ValueChanged);
			// 
			// cmbVendedor
			// 
			appearance3.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbVendedor.Appearance = appearance3;
			this.cmbVendedor.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbVendedor.Cursor = System.Windows.Forms.Cursors.Default;
			ultraGridColumn4.Header.VisiblePosition = 0;
			ultraGridColumn4.Hidden = true;
			ultraGridColumn5.Header.VisiblePosition = 1;
			ultraGridColumn5.Width = 335;
			ultraGridColumn6.Header.VisiblePosition = 2;
			ultraGridBand2.Columns.AddRange(new object[] {
																										 ultraGridColumn4,
																										 ultraGridColumn5,
																										 ultraGridColumn6});
			this.cmbVendedor.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			this.cmbVendedor.DisplayMember = "Nombre";
			this.cmbVendedor.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbVendedor.Enabled = false;
			this.cmbVendedor.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbVendedor.Location = new System.Drawing.Point(80, 72);
			this.cmbVendedor.Name = "cmbVendedor";
			this.cmbVendedor.Size = new System.Drawing.Size(250, 21);
			this.cmbVendedor.TabIndex = 787;
			this.cmbVendedor.ValueMember = "idPersonal";
			this.cmbVendedor.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cmbVendedor_KeyDown);
			this.cmbVendedor.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.cmbVendedor_InitializeLayout);
			// 
			// ultraDataSource2
			// 
			ultraDataColumn1.DataType = typeof(int);
			this.ultraDataSource2.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn1,
																																 ultraDataColumn2,
																																 ultraDataColumn3});
			// 
			// dtFecha
			// 
			appearance4.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtFecha.Appearance = appearance4;
			this.dtFecha.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton1.Caption = "Today";
			this.dtFecha.DateButtons.Add(dateButton1);
			this.dtFecha.Enabled = false;
			this.dtFecha.Format = "dd/MM/yyyy";
			this.dtFecha.Location = new System.Drawing.Point(80, 8);
			this.dtFecha.Name = "dtFecha";
			this.dtFecha.NonAutoSizeHeight = 23;
			this.dtFecha.Size = new System.Drawing.Size(112, 21);
			this.dtFecha.SpinButtonsVisible = true;
			this.dtFecha.TabIndex = 784;
			this.dtFecha.Value = ((object)(resources.GetObject("dtFecha.Value")));
			this.dtFecha.KeyDown += new System.Windows.Forms.KeyEventHandler(this.dtFecha_KeyDown);
			this.dtFecha.BeforeDropDown += new System.ComponentModel.CancelEventHandler(this.dtFecha_BeforeDropDown);
			this.dtFecha.ValueChanged += new System.EventHandler(this.dtFecha_ValueChanged);
			// 
			// ultraDataSource1
			// 
			ultraDataColumn4.DataType = typeof(int);
			ultraDataColumn5.DataType = typeof(int);
			ultraDataColumn7.DataType = typeof(int);
			ultraDataColumn8.DataType = typeof(int);
			this.ultraDataSource1.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn4,
																																 ultraDataColumn5,
																																 ultraDataColumn6,
																																 ultraDataColumn7,
																																 ultraDataColumn8,
																																 ultraDataColumn9});
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
			this.btnEditar.Location = new System.Drawing.Point(728, 472);
			this.btnEditar.Name = "btnEditar";
			this.btnEditar.Size = new System.Drawing.Size(16, 23);
			this.btnEditar.TabIndex = 800;
			this.btnEditar.Text = "Editar";
			this.btnEditar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnEditar.Visible = false;
			this.btnEditar.Click += new System.EventHandler(this.btnEditar_Click);
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
			this.btnCancelar.Location = new System.Drawing.Point(88, 473);
			this.btnCancelar.Name = "btnCancelar";
			this.btnCancelar.Size = new System.Drawing.Size(72, 23);
			this.btnCancelar.TabIndex = 799;
			this.btnCancelar.Text = "Cancelar";
			this.btnCancelar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnCancelar.Click += new System.EventHandler(this.btnCancelar_Click);
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
			this.btnGuardar.Location = new System.Drawing.Point(8, 473);
			this.btnGuardar.Name = "btnGuardar";
			this.btnGuardar.Size = new System.Drawing.Size(72, 23);
			this.btnGuardar.TabIndex = 798;
			this.btnGuardar.Text = "Aceptar";
			this.btnGuardar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnGuardar.Click += new System.EventHandler(this.btnGuardar_Click);
			// 
			// btnNuevo
			// 
			this.btnNuevo.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnNuevo.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(230)), ((System.Byte)(230)), ((System.Byte)(230)));
			this.btnNuevo.CausesValidation = false;
			this.btnNuevo.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.btnNuevo.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(50)), ((System.Byte)(50)), ((System.Byte)(50)));
			this.btnNuevo.Image = ((System.Drawing.Image)(resources.GetObject("btnNuevo.Image")));
			this.btnNuevo.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnNuevo.Location = new System.Drawing.Point(752, 472);
			this.btnNuevo.Name = "btnNuevo";
			this.btnNuevo.Size = new System.Drawing.Size(48, 23);
			this.btnNuevo.TabIndex = 797;
			this.btnNuevo.Text = "Nuevo";
			this.btnNuevo.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnNuevo.Visible = false;
			this.btnNuevo.Click += new System.EventHandler(this.btnNuevo_Click);
			// 
			// groupBox3
			// 
			this.groupBox3.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.groupBox3.Enabled = false;
			this.groupBox3.Location = new System.Drawing.Point(0, 460);
			this.groupBox3.Name = "groupBox3";
			this.groupBox3.Size = new System.Drawing.Size(816, 8);
			this.groupBox3.TabIndex = 796;
			this.groupBox3.TabStop = false;
			// 
			// grdReporteAuditoria
			// 
			this.grdReporteAuditoria.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.grdReporteAuditoria.Cursor = System.Windows.Forms.Cursors.Default;
			this.grdReporteAuditoria.DataSource = this.ultraDataSource1;
			appearance5.BackColor = System.Drawing.Color.White;
			this.grdReporteAuditoria.DisplayLayout.Appearance = appearance5;
			this.grdReporteAuditoria.DisplayLayout.AutoFitColumns = true;
			ultraGridColumn7.Header.VisiblePosition = 0;
			ultraGridColumn7.Hidden = true;
			ultraGridColumn8.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn8.Header.VisiblePosition = 1;
			ultraGridColumn8.Hidden = true;
			ultraGridColumn9.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn9.Header.VisiblePosition = 2;
			ultraGridColumn9.Width = 271;
			ultraGridColumn10.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance6.BackColor = System.Drawing.SystemColors.Control;
			appearance6.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn10.CellAppearance = appearance6;
			ultraGridColumn10.Header.VisiblePosition = 3;
			ultraGridColumn10.Width = 75;
			appearance7.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn11.CellAppearance = appearance7;
			ultraGridColumn11.Header.Caption = "Calificación";
			ultraGridColumn11.Header.VisiblePosition = 4;
			ultraGridColumn11.Width = 75;
			ultraGridColumn12.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn12.Header.VisiblePosition = 5;
			ultraGridColumn12.Width = 350;
			ultraGridBand3.Columns.AddRange(new object[] {
																										 ultraGridColumn7,
																										 ultraGridColumn8,
																										 ultraGridColumn9,
																										 ultraGridColumn10,
																										 ultraGridColumn11,
																										 ultraGridColumn12});
			ultraGridBand3.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			appearance8.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings1.Appearance = appearance8;
			summarySettings1.DisplayFormat = "{00}";
			summarySettings1.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance9.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings2.Appearance = appearance9;
			summarySettings2.DisplayFormat = "{00}";
			summarySettings2.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			ultraGridBand3.Summaries.AddRange(new Infragistics.Win.UltraWinGrid.SummarySettings[] {
																																															summarySettings1,
																																															summarySettings2});
			ultraGridBand3.SummaryFooterCaption = "TOTALES";
			this.grdReporteAuditoria.DisplayLayout.BandsSerializer.Add(ultraGridBand3);
			this.grdReporteAuditoria.DisplayLayout.GroupByBox.Hidden = true;
			appearance10.ForeColor = System.Drawing.Color.Black;
			appearance10.ForeColorDisabled = System.Drawing.Color.Black;
			this.grdReporteAuditoria.DisplayLayout.Override.ActiveRowAppearance = appearance10;
			this.grdReporteAuditoria.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.grdReporteAuditoria.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance11.BackColor = System.Drawing.Color.Transparent;
			this.grdReporteAuditoria.DisplayLayout.Override.CardAreaAppearance = appearance11;
			appearance12.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance12.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance12.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance12.FontData.BoldAsString = "True";
			appearance12.FontData.Name = "Arial";
			appearance12.FontData.SizeInPoints = 8F;
			appearance12.ForeColor = System.Drawing.Color.White;
			appearance12.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.grdReporteAuditoria.DisplayLayout.Override.HeaderAppearance = appearance12;
			this.grdReporteAuditoria.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti;
			appearance13.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance13.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance13.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.grdReporteAuditoria.DisplayLayout.Override.RowAlternateAppearance = appearance13;
			appearance14.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance14.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance14.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.grdReporteAuditoria.DisplayLayout.Override.RowSelectorAppearance = appearance14;
			appearance15.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance15.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance15.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.grdReporteAuditoria.DisplayLayout.Override.SelectedRowAppearance = appearance15;
			this.grdReporteAuditoria.DisplayLayout.ViewStyleBand = Infragistics.Win.UltraWinGrid.ViewStyleBand.OutlookGroupBy;
			this.grdReporteAuditoria.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.grdReporteAuditoria.Location = new System.Drawing.Point(8, 136);
			this.grdReporteAuditoria.Name = "grdReporteAuditoria";
			this.grdReporteAuditoria.Size = new System.Drawing.Size(792, 320);
			this.grdReporteAuditoria.TabIndex = 804;
			this.grdReporteAuditoria.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.grdReporteAuditoria_KeyPress);
			this.grdReporteAuditoria.KeyDown += new System.Windows.Forms.KeyEventHandler(this.grdReporteAuditoria_KeyDown);
			this.grdReporteAuditoria.AfterCellUpdate += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.grdReporteAuditoria_AfterCellUpdate);
			this.grdReporteAuditoria.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.grdReporteAuditoria_InitializeLayout);
			// 
			// btnImprimir
			// 
			this.btnImprimir.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnImprimir.CausesValidation = false;
			this.btnImprimir.Enabled = false;
			this.btnImprimir.Image = ((System.Drawing.Image)(resources.GetObject("btnImprimir.Image")));
			this.btnImprimir.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnImprimir.Location = new System.Drawing.Point(168, 472);
			this.btnImprimir.Name = "btnImprimir";
			this.btnImprimir.Size = new System.Drawing.Size(77, 23);
			this.btnImprimir.TabIndex = 806;
			this.btnImprimir.Text = "&Imprimir";
			this.btnImprimir.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnImprimir.Click += new System.EventHandler(this.btnImprimir_Click);
			// 
			// cmbBodega
			// 
			this.cmbBodega.CausesValidation = false;
			this.cmbBodega.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbBodega.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbBodega.DisplayMember = "Nombre";
			this.cmbBodega.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbBodega.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbBodega.Enabled = false;
			this.cmbBodega.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbBodega.Location = new System.Drawing.Point(80, 40);
			this.cmbBodega.MaxDropDownItems = 30;
			this.cmbBodega.Name = "cmbBodega";
			this.cmbBodega.Size = new System.Drawing.Size(180, 22);
			this.cmbBodega.TabIndex = 807;
			this.cmbBodega.ValueMember = "Bodega";
			this.cmbBodega.ValueChanged += new System.EventHandler(this.cmbBodega_ValueChanged);
			this.cmbBodega.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.cmbBodega_InitializeLayout);
			// 
			// frmReporteAuditoriaLocal1
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 14);
			this.ClientSize = new System.Drawing.Size(808, 502);
			this.ControlBox = false;
			this.Controls.Add(this.cmbBodega);
			this.Controls.Add(this.btnImprimir);
			this.Controls.Add(this.grdReporteAuditoria);
			this.Controls.Add(this.btnEditar);
			this.Controls.Add(this.btnCancelar);
			this.Controls.Add(this.btnGuardar);
			this.Controls.Add(this.btnNuevo);
			this.Controls.Add(this.groupBox3);
			this.Controls.Add(this.lblEstado);
			this.Controls.Add(this.txtTotal);
			this.Controls.Add(this.txtIdAuditoriaLocales);
			this.Controls.Add(this.btnConsultar);
			this.Controls.Add(this.label2);
			this.Controls.Add(this.label1);
			this.Controls.Add(this.label7);
			this.Controls.Add(this.lblBodega);
			this.Controls.Add(this.lblfecha);
			this.Controls.Add(this.cmbAuditor);
			this.Controls.Add(this.cmbVendedor);
			this.Controls.Add(this.dtFecha);
			this.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
			this.KeyPreview = true;
			this.MaximizeBox = false;
			this.MinimizeBox = false;
			this.Name = "frmReporteAuditoriaLocal1";
			this.ShowInTaskbar = false;
			this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
			this.Text = "Calificación Estado Local";
			this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.frmReporteAuditoriaLocal1_KeyDown);
			this.Load += new System.EventHandler(this.frmReporteAuditoriaLocal1_Load);
			((System.ComponentModel.ISupportInitialize)(this.txtTotal)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdAuditoriaLocales)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbAuditor)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbVendedor)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFecha)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.grdReporteAuditoria)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbBodega)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion
		 


		private void frmReporteAuditoriaLocal_Load(object sender, System.EventArgs e)
		{
		
		}

		private void btnNuevo_Click(object sender, System.EventArgs e)
		{
			//			this.btnCancelar_Click(sender, e);
			//			this.dtFecha.CalendarInfo.MaxDate = DateTime.Today;
			//			this.dtFecha.Value = DateTime.Today;
			//			this.dtFecha.Enabled=true;
			//			this.cmbBodega.Focus();
			//			this.txtIdAuditoriaLocales.Enabled=true;
			//			this.cmbBodega.Enabled=true;
			//			this.cmbAuditor.Enabled=true;
			//			this.cmbVendedor.Enabled=true;
			//			FuncionesProcedimientos.EstadoGrids(true, this.grdReporteAuditoria);
			//	
			//			string sSQL = string.Format("Exec DetalleAuditoriaLocales");
			//			this.grdReporteAuditoria.DataSource = Funcion.dvProcedimiento(cdsSeteoF, sSQL);
			//		
			//			this.btnNuevo.Enabled = false;
			//			this.btnConsultar.Enabled = false;
			//			this.btnGuardar.Enabled = true;
			//			this.btnCancelar.Enabled = true;	
			//
			//					ModificaGrilla(true);

		}

		private void btnGuardar_Click(object sender, System.EventArgs e)
		{
			#region valida
			if (!Validacion.vbComboVacio(this.cmbBodega, "Seleccione la bodega")) return;
			if (!Validacion.vbComboVacio(this.cmbVendedor, "Seleccione Jefe de Local")) return;
			if (!Validacion.vbComboVacio(this.cmbAuditor, "Seleccione Auditor")) return;
			#endregion valida

			#region Calificacion
			int iCont =0;
			
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.grdReporteAuditoria.Rows.All)
			{	
				if((int)dr.Cells["PuntajeObtenido"].Value > 10 || (int)dr.Cells["PuntajeObtenido"].Value == 0 ) 
				{
					this.grdReporteAuditoria.ActiveRow = dr;
					this.grdReporteAuditoria.ActiveRow.Selected = true;	
					iCont = iCont + 1 ;
				}								
			}
			if(iCont > 0) {MessageBox.Show("No se permite valores mayores a 10 o iguales a 0" , "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop); return;} 
			#endregion Calificacion

			#region Valida Observaciones
			int iContobs = 0;
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow drDTi in this.grdReporteAuditoria.Rows.All)
			{
				if(drDTi.Cells["Observaciones"].Value.ToString().Length < 10)
				{
					iContobs = iContobs + 1;					
					this.grdReporteAuditoria.ActiveRow = drDTi;
					this.grdReporteAuditoria.ActiveRow.Selected = true;
					
					break;
				}
			}
			if (iContobs > 0) {MessageBox.Show("El campo Observaciones debe tener minimo 10 caracteres ", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop); return;}
			#endregion Valida Observaciones

			#region Guarda registro			
			string sSQLl = string.Format("Exec GuardaAuditoriaLocales {0},'{1}',{2},{3},{4},{5}",
				(int) this.txtIdAuditoriaLocales.Value, Convert.ToDateTime(this.dtFecha.Value).ToString("yyyyMMdd"),
				(int) this.cmbBodega.Value, (int) this.cmbVendedor.Value,
				(int) this.cmbAuditor.Value, (int) this.txtTotal.Value);
			this.txtIdAuditoriaLocales.Value=Funcion.iEscalarSQL(cdsSeteoF,sSQLl);
			
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.grdReporteAuditoria.Rows.All)
			{							
				string sSQLDetalle = string.Format("Exec GuardaDetalleAuditoriaLocales {0},{1},'{2}',{3},{4},'{5}'", 
					(int)dr.Cells["IdDetalleAuditoria"].Value,
					(int)this.txtIdAuditoriaLocales.Value,
					dr.Cells["Descripcion"].Value.ToString(),
					(int)dr.Cells["Puntaje"].Value,
					(int)dr.Cells["PuntajeObtenido"].Value,
					dr.Cells["Observaciones"].Value.ToString());
				Funcion.EjecutaSQL(cdsSeteoF, sSQLDetalle,true);
			}			
			#endregion Guarda registro

			#region controles
			this.dtFecha.Enabled=false;
			this.txtIdAuditoriaLocales.Enabled=false;
			this.cmbBodega.Enabled=false;
			this.cmbAuditor.Enabled=false;
			this.cmbVendedor.Enabled=false;

			FuncionesProcedimientos.EstadoGrids(false, this.grdReporteAuditoria);

			this.btnNuevo.Enabled = true;
			this.btnConsultar.Enabled = true;
			this.btnEditar.Enabled = true;
			this.btnGuardar.Enabled = false;
			this.btnCancelar.Enabled = true;
			this.btnImprimir.Enabled = true;			
			#endregion controles 		
		}

		private void btnEditar_Click(object sender, System.EventArgs e)
		{
			#region controles
			this.dtFecha.CalendarInfo.MaxDate = DateTime.Today;
			this.dtFecha.Value = DateTime.Today;
			this.dtFecha.Enabled=true;
			this.cmbBodega.Focus();
			this.txtIdAuditoriaLocales.Enabled=true;
			this.cmbBodega.Enabled=true;
			this.cmbAuditor.Enabled=true;
			this.cmbVendedor.Enabled=true;
			FuncionesProcedimientos.EstadoGrids(true, this.grdReporteAuditoria);


			this.btnNuevo.Enabled = false;
			this.btnConsultar.Enabled = false;
			this.btnEditar.Enabled = false;
			this.btnGuardar.Enabled = true;
			this.btnCancelar.Enabled = true;

			#endregion controles
			
			//ModificaGrilla(true);
		}

		private void btnCancelar_Click(object sender, System.EventArgs e)
		{
			this.dtFecha.Value=System.DBNull.Value;
			this.txtIdAuditoriaLocales.Value=0;
			this.cmbBodega.Value=System.DBNull.Value;
			this.cmbAuditor.Value=System.DBNull.Value;
			this.cmbVendedor.Value=System.DBNull.Value;
			this.txtTotal.Value=0;
			this.txtTotal.Value=0;
			this.lblEstado.Text=" ";
			this.dtFecha.Enabled=false;
			this.txtIdAuditoriaLocales.Enabled=false;
			this.cmbBodega.Enabled=false;
			this.cmbAuditor.Enabled=false;
			this.cmbVendedor.Enabled=false;
			this.txtTotal.Enabled=false;
			//			FuncionesProcedimientos.EstadoGrids(false, this.grdReporteAuditoria);
			//			//			FuncionesProcedimientos.EstadoGrids(true, this.grdReporteAuditoria);
			//			string sSQL = string.Format("Exec DetalleAuditoriaLocales 0");
			//			this.grdReporteAuditoria.DataSource = Funcion.dvProcedimiento(cdsSeteoF, sSQL);
			this.btnNuevo.Enabled = true;
			this.btnConsultar.Enabled = true;
			this.btnEditar.Enabled = false;
			this.btnGuardar.Enabled = false;
			this.btnCancelar.Enabled = false;
			//			ModificaGrilla(false);	
			//this.Close();	
			this.DialogResult = DialogResult.No;
		}

		private void sumatoria()
		{
			int iPuntajeObtenido = 0;

			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.grdReporteAuditoria.Rows.All)
			{	
				iPuntajeObtenido = iPuntajeObtenido + (int)dr.Cells["PuntajeObtenido"].Value;				
			}

			this.txtTotal.Value = iPuntajeObtenido;

			if(iPuntajeObtenido < 90)
			{
				lblEstado.Text="REGULAR";
				lblEstado.ForeColor=Color.Red;
			}
			else
			{
				lblEstado.Text="EXCELENTE";
				lblEstado.ForeColor=Color.Blue;
			}
		}

		private void rango()
		{
			#region rango
			int iCont =0;
			int iCont1 =0;
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.grdReporteAuditoria.Rows.All)
			{	
				if((int)dr.Cells["PuntajeObtenido"].Value > 10) 
				{
					this.grdReporteAuditoria.ActiveRow = dr;
					this.grdReporteAuditoria.ActiveRow.Selected = true;	
					iCont = iCont + 1 ;
				}
								
			}
			if(iCont > 0) {MessageBox.Show("No se permite valores mayores a 10 " , "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop); return;} 

			#endregion rango
		}
		
		private void dtFecha_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter) this.cmbBodega.Focus();
		}

		private void cmbBodega_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter) this.cmbVendedor.Focus();
		}

		private void cmbVendedor_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter) this.cmbAuditor.Focus();
		}

		private void cmbAuditor_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter) this.grdReporteAuditoria.Focus();
		}

		private void grdReporteAuditoria_AfterCellUpdate(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			if (e.Cell.Column.ToString() == "PuntajeObtenido")
			{				
				if (e.Cell.Row.Cells["PuntajeObtenido"].Value == System.DBNull.Value)
				{
					e.Cell.Row.Cells["PuntajeObtenido"].Value = 0;
					return;
				}

				sumatoria();
			}			
		}

		private void btnConsultar_Click(object sender, System.EventArgs e)
		{
			this.btnNuevo.Enabled = true;
			this.btnConsultar.Enabled = true;
			this.btnEditar.Enabled = true;
			this.btnGuardar.Enabled = false;
			this.btnCancelar.Enabled = true;
		}

		private void dtFecha_ValueChanged(object sender, System.EventArgs e)
		{
			
		}
      
		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);
		}

		private void frmReporteAuditoriaLocal1_Load(object sender, System.EventArgs e)
		{	
		
			string sSQLA = string.Format("Exec ConsultaAuditoriaBodegaReamtes {0}", Bodegas);
			this.cmbBodega.DataSource = Funcion.dvProcedimiento(cdsSeteoF, sSQLA);

			this.cmbAuditor.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "SELECT idPersonal, ApellidoPaterno + ' '+ ApellidoMaterno+ ' '+ PrimerNombre+ ' '+ SegundoNombre As Nombre  FROM Nomina WHERE idNomina IN ( 2613, 2596 )");
			
			int iBodega = 0;
	
			string sSQL = string.Format("Exec ConsultaAuditoriaLocales {0}, '{1}'", Bodegas, dtFechas.ToString("yyyyMMdd"));
			SqlDataReader dr = Funcion.rEscalarSQL(cdsSeteoF, sSQL, true);
			dr.Read();
			if (dr.HasRows)
			{
				this.txtIdAuditoriaLocales.Value = dr.GetInt32(0);
				this.dtFecha.Value = dr.GetDateTime(1);
				iBodega = dr.GetInt32(2);
				this.cmbVendedor.Value = dr.GetInt32(3);
				this.cmbAuditor.Value = dr.GetInt32(4);
				this.txtTotal.Value = dr.GetInt32(5);				
			}
			dr.Close();

			this.cmbBodega.Value = iBodega;

			string sSQLDet = string.Format("Exec DetalleAuditoriaLocales {0}", (int)this.txtIdAuditoriaLocales.Value);
			this.grdReporteAuditoria.DataSource = Funcion.dvProcedimiento(cdsSeteoF, sSQLDet);
		
			if ((int)this.txtIdAuditoriaLocales.Value == 0)
			{				
				this.dtFecha.Value = dtFechas;
				this.cmbBodega.Value = Bodegas;
			}
			else this.btnImprimir.Enabled = true;

			if (iEstadoConteo == 1)
			{				
				this.cmbAuditor.Enabled = true;

				FuncionesProcedimientos.EstadoGrids(true, this.grdReporteAuditoria);
				
				this.btnGuardar.Enabled = true;
			}
			else FuncionesProcedimientos.EstadoGrids(false, this.grdReporteAuditoria);
		}

		private void DiseñoGrid()
		{
			#region Diseño Grid
			if (this.grdReporteAuditoria.Rows.Count > 0)
			{
				this.grdReporteAuditoria.DisplayLayout.Bands[0].Columns["Descripcion"].Width = 277;
				this.grdReporteAuditoria.DisplayLayout.Bands[0].Columns["Descripcion"].CellActivation = Activation.ActivateOnly;
				this.grdReporteAuditoria.DisplayLayout.Bands[0].Columns["Puntaje"].Width = 60;
				this.grdReporteAuditoria.DisplayLayout.Bands[0].Columns["Puntaje"].CellActivation = Activation.ActivateOnly;
				this.grdReporteAuditoria.DisplayLayout.Bands[0].Columns["PuntajeObtenido"].Width = 98;
				this.grdReporteAuditoria.DisplayLayout.Bands[0].Columns["PuntajeObtenido"].CellActivation = Activation.AllowEdit;
				this.grdReporteAuditoria.DisplayLayout.Bands[0].Columns["Observaciones"].Width = 343;
				this.grdReporteAuditoria.DisplayLayout.Bands[0].Columns["Observaciones"].CellActivation = Activation.AllowEdit;
			}
			#endregion Diseño Grid
		}

		private void grdReporteAuditoria_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
			
		}

		private void grdReporteAuditoria_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (grdReporteAuditoria.ActiveCell.Column.Key.ToString().ToUpper() == "PUNTAJEOBTENIDO")
				Funcion.CamposNumericos(sender, e);			
		}

		private void grdReporteAuditoria_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			FuncionesProcedimientos.DesplazamientoGrids(sender,e,this.grdReporteAuditoria);			
		}

		private void cmbAuditor_ValueChanged(object sender, System.EventArgs e)
		{
		
		}

		private void txtTotal_ValueChanged(object sender, System.EventArgs e)
		{
		
		}

		private void btnImprimir_Click(object sender, System.EventArgs e)
		{
			ParameterFields paramFields = new ParameterFields ();
				
			ParameterField paramFieldId = new ParameterField ();
			ParameterDiscreteValue ValIdId = new ParameterDiscreteValue ();
			paramFieldId.ParameterFieldName = "@idAudirotiaLocales";
			ValIdId.Value = int.Parse(this.txtIdAuditoriaLocales.Value.ToString());
			paramFieldId.CurrentValues.Add (ValIdId);
			paramFields.Add (paramFieldId);


			Reporte miRep = new Reporte("ReporteAuditoriaLocales.rpt", "");
			miRep.MdiParent = this.MdiParent;
			miRep.Titulo("Recepción Mercaderia de Retiro");
			miRep.crVista.ParameterFieldInfo = paramFields;
			miRep.Show();
		}
		
		private void cmbBodega_ValueChanged(object sender, System.EventArgs e)
		{		
			if (this.cmbBodega.ActiveRow != null)
			{
				int ibodega = (int)this.cmbBodega.Value;
							
				//				this.cmbVendedor.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Exec [ConsultaAuditoriaLocalesPersonalBodega] {0}, '{1}'", ibodega, Convert.ToDateTime(this.dtFecha.Value).ToString("yyyyMMdd")  );
				this.cmbVendedor.DataSource = Funcion.dvProcedimiento(cdsSeteoF, string.Format("Exec [ConsultaAuditoriaLocalesPersonalBodega] {0}, '{1}' ", ibodega, Convert.ToDateTime(this.dtFecha.Value).ToString("yyyyMMdd")) );
				this.cmbVendedor.Value = Funcion.iEscalarSQL(cdsSeteoF, string.Format("Exec [ConsultaAuditoriaLocalesPersonalBodega] {0}, '{1}'", ibodega, Convert.ToDateTime(this.dtFecha.Value).ToString("yyyyMMdd") ));					
			}
		}

		private void cmbBodega_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}

		private void frmReporteAuditoriaLocal1_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Escape) this.DialogResult = DialogResult.No;
		}

		private void dtFecha_BeforeDropDown(object sender, System.ComponentModel.CancelEventArgs e)
		{
		
		}

		private void cmbVendedor_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}
		
	}
}
