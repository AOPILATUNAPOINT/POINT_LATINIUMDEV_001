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
	/// Descripción breve de frmRolAutorizacionPagoCheques.
	/// </summary>
	public class frmRolAutorizacionPagoCheques : System.Windows.Forms.Form
	{
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource1;
		public Infragistics.Win.UltraWinGrid.UltraDropDown cmbCuenta;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource2;
		private C1.Data.C1DataSet cdsSeteoF;
		private System.Windows.Forms.Button btnAutorizar;
		private Infragistics.Win.UltraWinGrid.UltraGrid uGridListaCheques;
		private System.Windows.Forms.Button btnSalir;
		private System.Windows.Forms.Label lblContador;
		private System.Windows.Forms.GroupBox groupBox1;
		private System.Windows.Forms.SaveFileDialog saveFileDialog1;
		private Infragistics.Win.UltraWinGrid.ExcelExport.UltraGridExcelExporter ultraGridExcelExporter1;
		private System.Windows.Forms.Button btnExportar;
		public System.Windows.Forms.Button btnImprimir;
		private System.Windows.Forms.Label lblCuenta;
		private System.Windows.Forms.Button btnAsignar;
		public Infragistics.Win.UltraWinGrid.UltraCombo cmbCuentaAsigna;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public frmRolAutorizacionPagoCheques()
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
		DateTime dtFecha = DateTime.Today;
		int iRolComisionDecimo = 0;

		public frmRolAutorizacionPagoCheques(int IdRolNomina, DateTime DTFecha, int IRolComisionDecimo)
		{
			//
			// Necesario para admitir el Diseñador de Windows Forms
			//
			InitializeComponent();

			idRolNomina = IdRolNomina;
			dtFecha = DTFecha;
			iRolComisionDecimo = IRolComisionDecimo;

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
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idDetalleRolNomina");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cedula");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idNomina");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPersonal");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idRolPrestamo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Apellidos");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombres");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCuenta", -1, "cmbCuenta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Total");
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings1 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Total", 8, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Total", 8, true);
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idDetalleRolNomina");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Cedula");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn3 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idNomina");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn4 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idPersonal");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn5 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idRolPrestamo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn6 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Apellidos");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn7 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombres");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn8 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCuenta");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn9 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Total");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCuenta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn11 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn12 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Descripcion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn10 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCuenta");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn11 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Codigo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn12 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Descripcion");
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmRolAutorizacionPagoCheques));
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand3 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn13 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCuenta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn14 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn15 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Descripcion");
			this.uGridListaCheques = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource1 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.cmbCuenta = new Infragistics.Win.UltraWinGrid.UltraDropDown();
			this.ultraDataSource2 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.btnAutorizar = new System.Windows.Forms.Button();
			this.btnSalir = new System.Windows.Forms.Button();
			this.lblContador = new System.Windows.Forms.Label();
			this.groupBox1 = new System.Windows.Forms.GroupBox();
			this.saveFileDialog1 = new System.Windows.Forms.SaveFileDialog();
			this.ultraGridExcelExporter1 = new Infragistics.Win.UltraWinGrid.ExcelExport.UltraGridExcelExporter();
			this.btnExportar = new System.Windows.Forms.Button();
			this.btnImprimir = new System.Windows.Forms.Button();
			this.lblCuenta = new System.Windows.Forms.Label();
			this.cmbCuentaAsigna = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.btnAsignar = new System.Windows.Forms.Button();
			((System.ComponentModel.ISupportInitialize)(this.uGridListaCheques)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCuenta)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCuentaAsigna)).BeginInit();
			this.SuspendLayout();
			// 
			// uGridListaCheques
			// 
			this.uGridListaCheques.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.uGridListaCheques.Cursor = System.Windows.Forms.Cursors.Default;
			this.uGridListaCheques.DataSource = this.ultraDataSource1;
			appearance1.BackColor = System.Drawing.Color.White;
			this.uGridListaCheques.DisplayLayout.Appearance = appearance1;
			this.uGridListaCheques.DisplayLayout.AutoFitColumns = true;
			ultraGridBand1.AddButtonCaption = "RolReingreso";
			ultraGridColumn1.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn2.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridColumn3.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn3.Header.VisiblePosition = 2;
			ultraGridColumn3.Hidden = true;
			ultraGridColumn4.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn4.Header.VisiblePosition = 4;
			ultraGridColumn4.Hidden = true;
			ultraGridColumn5.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn5.Header.VisiblePosition = 5;
			ultraGridColumn5.Hidden = true;
			ultraGridColumn6.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn6.Header.VisiblePosition = 3;
			ultraGridColumn6.Width = 180;
			ultraGridColumn7.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn7.Header.VisiblePosition = 6;
			ultraGridColumn7.Width = 180;
			ultraGridColumn8.Header.Caption = "Cuenta Bancaria";
			ultraGridColumn8.Header.VisiblePosition = 7;
			ultraGridColumn8.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDown;
			ultraGridColumn8.Width = 200;
			ultraGridColumn9.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance2.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn9.CellAppearance = appearance2;
			ultraGridColumn9.Format = "#,##0.00";
			ultraGridColumn9.Header.Caption = "Monto";
			ultraGridColumn9.Header.VisiblePosition = 8;
			ultraGridColumn9.Width = 100;
			ultraGridBand1.Columns.AddRange(new object[] {
																										 ultraGridColumn1,
																										 ultraGridColumn2,
																										 ultraGridColumn3,
																										 ultraGridColumn4,
																										 ultraGridColumn5,
																										 ultraGridColumn6,
																										 ultraGridColumn7,
																										 ultraGridColumn8,
																										 ultraGridColumn9});
			ultraGridBand1.Header.Caption = "";
			appearance3.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings1.Appearance = appearance3;
			summarySettings1.DisplayFormat = "{0: #,##0.00}";
			summarySettings1.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			ultraGridBand1.Summaries.AddRange(new Infragistics.Win.UltraWinGrid.SummarySettings[] {
																																															summarySettings1});
			ultraGridBand1.SummaryFooterCaption = "";
			this.uGridListaCheques.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.uGridListaCheques.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
			this.uGridListaCheques.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.uGridListaCheques.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance4.BackColor = System.Drawing.Color.Transparent;
			this.uGridListaCheques.DisplayLayout.Override.CardAreaAppearance = appearance4;
			appearance5.ForeColor = System.Drawing.Color.Black;
			appearance5.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridListaCheques.DisplayLayout.Override.CellAppearance = appearance5;
			appearance6.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance6.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance6.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance6.FontData.BoldAsString = "True";
			appearance6.FontData.Name = "Arial";
			appearance6.FontData.SizeInPoints = 10F;
			appearance6.ForeColor = System.Drawing.Color.White;
			appearance6.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridListaCheques.DisplayLayout.Override.HeaderAppearance = appearance6;
			appearance7.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance7.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance7.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridListaCheques.DisplayLayout.Override.RowSelectorAppearance = appearance7;
			appearance8.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance8.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance8.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridListaCheques.DisplayLayout.Override.SelectedRowAppearance = appearance8;
			this.uGridListaCheques.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridListaCheques.Location = new System.Drawing.Point(8, 88);
			this.uGridListaCheques.Name = "uGridListaCheques";
			this.uGridListaCheques.Size = new System.Drawing.Size(781, 320);
			this.uGridListaCheques.TabIndex = 721;
			this.uGridListaCheques.KeyDown += new System.Windows.Forms.KeyEventHandler(this.uGridListaCheques_KeyDown);
			this.uGridListaCheques.AfterCellUpdate += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.uGridListaCheques_AfterCellUpdate);
			this.uGridListaCheques.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.uGridListaCheques_InitializeLayout);
			// 
			// ultraDataSource1
			// 
			ultraDataColumn1.DataType = typeof(int);
			ultraDataColumn3.DataType = typeof(int);
			ultraDataColumn4.DataType = typeof(int);
			ultraDataColumn5.DataType = typeof(int);
			ultraDataColumn8.DataType = typeof(int);
			ultraDataColumn9.DataType = typeof(System.Decimal);
			ultraDataColumn9.DefaultValue = new System.Decimal(new int[] {
																																		 0,
																																		 0,
																																		 0,
																																		 0});
			this.ultraDataSource1.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn1,
																																 ultraDataColumn2,
																																 ultraDataColumn3,
																																 ultraDataColumn4,
																																 ultraDataColumn5,
																																 ultraDataColumn6,
																																 ultraDataColumn7,
																																 ultraDataColumn8,
																																 ultraDataColumn9});
			// 
			// cmbCuenta
			// 
			this.cmbCuenta.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbCuenta.DataSource = this.ultraDataSource2;
			ultraGridBand2.AddButtonCaption = "PersParentesco";
			ultraGridColumn10.Header.VisiblePosition = 0;
			ultraGridColumn10.Hidden = true;
			ultraGridColumn11.Header.VisiblePosition = 1;
			ultraGridColumn11.Width = 120;
			ultraGridColumn12.Header.VisiblePosition = 2;
			ultraGridColumn12.Width = 300;
			ultraGridBand2.Columns.AddRange(new object[] {
																										 ultraGridColumn10,
																										 ultraGridColumn11,
																										 ultraGridColumn12});
			this.cmbCuenta.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			this.cmbCuenta.DisplayMember = "Descripcion";
			this.cmbCuenta.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbCuenta.Location = new System.Drawing.Point(360, 264);
			this.cmbCuenta.Name = "cmbCuenta";
			this.cmbCuenta.Size = new System.Drawing.Size(352, 104);
			this.cmbCuenta.TabIndex = 722;
			this.cmbCuenta.ValueMember = "idCuenta";
			this.cmbCuenta.Visible = false;
			// 
			// ultraDataSource2
			// 
			ultraDataColumn10.DataType = typeof(int);
			this.ultraDataSource2.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn10,
																																 ultraDataColumn11,
																																 ultraDataColumn12});
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
			// btnAutorizar
			// 
			this.btnAutorizar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btnAutorizar.CausesValidation = false;
			this.btnAutorizar.Image = ((System.Drawing.Image)(resources.GetObject("btnAutorizar.Image")));
			this.btnAutorizar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnAutorizar.Location = new System.Drawing.Point(416, 8);
			this.btnAutorizar.Name = "btnAutorizar";
			this.btnAutorizar.Size = new System.Drawing.Size(88, 23);
			this.btnAutorizar.TabIndex = 723;
			this.btnAutorizar.Text = "Autorizar";
			this.btnAutorizar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnAutorizar.Click += new System.EventHandler(this.btnAutorizar_Click);
			// 
			// btnSalir
			// 
			this.btnSalir.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btnSalir.CausesValidation = false;
			this.btnSalir.Image = ((System.Drawing.Image)(resources.GetObject("btnSalir.Image")));
			this.btnSalir.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnSalir.Location = new System.Drawing.Point(704, 8);
			this.btnSalir.Name = "btnSalir";
			this.btnSalir.Size = new System.Drawing.Size(88, 23);
			this.btnSalir.TabIndex = 724;
			this.btnSalir.Text = "Salir";
			this.btnSalir.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnSalir.Click += new System.EventHandler(this.btnSalir_Click);
			// 
			// lblContador
			// 
			this.lblContador.AutoSize = true;
			this.lblContador.BackColor = System.Drawing.Color.Transparent;
			this.lblContador.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblContador.ForeColor = System.Drawing.Color.Firebrick;
			this.lblContador.Location = new System.Drawing.Point(8, 8);
			this.lblContador.Name = "lblContador";
			this.lblContador.Size = new System.Drawing.Size(0, 22);
			this.lblContador.TabIndex = 725;
			this.lblContador.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// groupBox1
			// 
			this.groupBox1.Location = new System.Drawing.Point(0, 72);
			this.groupBox1.Name = "groupBox1";
			this.groupBox1.Size = new System.Drawing.Size(800, 8);
			this.groupBox1.TabIndex = 726;
			this.groupBox1.TabStop = false;
			// 
			// btnExportar
			// 
			this.btnExportar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btnExportar.CausesValidation = false;
			this.btnExportar.Image = ((System.Drawing.Image)(resources.GetObject("btnExportar.Image")));
			this.btnExportar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnExportar.Location = new System.Drawing.Point(608, 8);
			this.btnExportar.Name = "btnExportar";
			this.btnExportar.Size = new System.Drawing.Size(88, 23);
			this.btnExportar.TabIndex = 727;
			this.btnExportar.Text = "Exportar";
			this.btnExportar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnExportar.Click += new System.EventHandler(this.btnExportar_Click);
			// 
			// btnImprimir
			// 
			this.btnImprimir.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btnImprimir.Image = ((System.Drawing.Image)(resources.GetObject("btnImprimir.Image")));
			this.btnImprimir.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnImprimir.Location = new System.Drawing.Point(512, 8);
			this.btnImprimir.Name = "btnImprimir";
			this.btnImprimir.Size = new System.Drawing.Size(88, 23);
			this.btnImprimir.TabIndex = 728;
			this.btnImprimir.Text = "&Imprimir";
			this.btnImprimir.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnImprimir.Click += new System.EventHandler(this.btnImprimir_Click);
			// 
			// lblCuenta
			// 
			this.lblCuenta.AutoSize = true;
			this.lblCuenta.BackColor = System.Drawing.Color.Transparent;
			this.lblCuenta.Location = new System.Drawing.Point(360, 42);
			this.lblCuenta.Name = "lblCuenta";
			this.lblCuenta.Size = new System.Drawing.Size(41, 16);
			this.lblCuenta.TabIndex = 730;
			this.lblCuenta.Text = "Cuenta";
			// 
			// cmbCuentaAsigna
			// 
			appearance9.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbCuentaAsigna.Appearance = appearance9;
			this.cmbCuentaAsigna.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbCuentaAsigna.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbCuentaAsigna.DataSource = this.ultraDataSource2;
			ultraGridColumn13.Header.VisiblePosition = 0;
			ultraGridColumn13.Hidden = true;
			ultraGridColumn14.Header.VisiblePosition = 1;
			ultraGridColumn15.Header.VisiblePosition = 2;
			ultraGridColumn15.Width = 280;
			ultraGridBand3.Columns.AddRange(new object[] {
																										 ultraGridColumn13,
																										 ultraGridColumn14,
																										 ultraGridColumn15});
			this.cmbCuentaAsigna.DisplayLayout.BandsSerializer.Add(ultraGridBand3);
			this.cmbCuentaAsigna.DisplayMember = "Descripcion";
			this.cmbCuentaAsigna.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbCuentaAsigna.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbCuentaAsigna.Location = new System.Drawing.Point(416, 40);
			this.cmbCuentaAsigna.Name = "cmbCuentaAsigna";
			this.cmbCuentaAsigna.Size = new System.Drawing.Size(280, 21);
			this.cmbCuentaAsigna.TabIndex = 732;
			this.cmbCuentaAsigna.ValueMember = "idCuenta";
			// 
			// btnAsignar
			// 
			this.btnAsignar.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.btnAsignar.Location = new System.Drawing.Point(704, 39);
			this.btnAsignar.Name = "btnAsignar";
			this.btnAsignar.Size = new System.Drawing.Size(88, 23);
			this.btnAsignar.TabIndex = 733;
			this.btnAsignar.Text = "&Asignar";
			this.btnAsignar.Click += new System.EventHandler(this.btnAsignar_Click);
			// 
			// frmRolAutorizacionPagoCheques
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(797, 416);
			this.ControlBox = false;
			this.Controls.Add(this.btnAsignar);
			this.Controls.Add(this.cmbCuentaAsigna);
			this.Controls.Add(this.lblCuenta);
			this.Controls.Add(this.btnImprimir);
			this.Controls.Add(this.groupBox1);
			this.Controls.Add(this.lblContador);
			this.Controls.Add(this.btnSalir);
			this.Controls.Add(this.btnAutorizar);
			this.Controls.Add(this.cmbCuenta);
			this.Controls.Add(this.uGridListaCheques);
			this.Controls.Add(this.btnExportar);
			this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
			this.KeyPreview = true;
			this.MaximizeBox = false;
			this.Name = "frmRolAutorizacionPagoCheques";
			this.ShowInTaskbar = false;
			this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
			this.Text = "Autorización de Pago Nómina en Cheques";
			this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.frmRolAutorizacionPagoCheques_KeyDown);
			this.Load += new System.EventHandler(this.frmRolAutorizacionPagoCheques_Load);
			((System.ComponentModel.ISupportInitialize)(this.uGridListaCheques)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCuenta)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCuentaAsigna)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);	
		}

		private void Consulta ()
		{
			this.uGridListaCheques.DataSource = Funcion.dvProcedimiento(cdsSeteoF, string.Format("Exec RolListaPersonalCheques {0}, {1}", idRolNomina, iRolComisionDecimo));

			this.lblContador.Text = this.uGridListaCheques.Rows.Count + " Registros Encontrados";
		}

		private void frmRolAutorizacionPagoCheques_Load(object sender, System.EventArgs e)
		{
			this.cmbCuentaAsigna.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Select c.idCuenta, c.Codigo, c.Descripcion From Cuenta c Inner Join CuentaCorriente cc On cc.idCuenta = c.idCuenta Order By Descripcion");
			this.cmbCuenta.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Select c.idCuenta, c.Codigo, c.Descripcion From Cuenta c Inner Join CuentaCorriente cc On cc.idCuenta = c.idCuenta Order By Descripcion");

			Consulta ();
		}

		private void frmRolAutorizacionPagoCheques_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Escape) this.DialogResult = DialogResult.No;
		}

		private void uGridListaCheques_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			FuncionesProcedimientos.DesplazamientoGrids(sender, e, this.uGridListaCheques);
		}

		private void uGridListaCheques_AfterCellUpdate(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			if (e.Cell.Column.ToString() == "idCuenta")
			{
				int idCuenta = 0;
				if (e.Cell.Row.Cells["idCuenta"].Value != System.DBNull.Value) idCuenta = (int)e.Cell.Row.Cells["idCuenta"].Value;

				if (iRolComisionDecimo == 0)
				{
					string sSQL = string.Format("Update DetalleRolNomina Set idCuenta = {0} Where idDetalleRolNomina = {1}", idCuenta, (int)e.Cell.Row.Cells["idDetalleRolNomina"].Value);
					Funcion.EjecutaSQL(cdsSeteoF, sSQL);
				}
				else if (iRolComisionDecimo == 1)
				{
					string sSQL = string.Format("Update DetalleComisiones Set idCuenta = {0} Where idDetalleComision = {1}", idCuenta, (int)e.Cell.Row.Cells["idDetalleRolNomina"].Value);
					Funcion.EjecutaSQL(cdsSeteoF, sSQL);
				}
				else if (iRolComisionDecimo == 2)
				{
					string sSQL = string.Format("Update NominaDetalleDecimoTercerSueldo Set idCuenta = {0} Where idNominaDetalleDecimoTercerSueldo = {1}", idCuenta, (int)e.Cell.Row.Cells["idDetalleRolNomina"].Value);
					Funcion.EjecutaSQL(cdsSeteoF, sSQL);
				}
			}
		}

		private void btnSalir_Click(object sender, System.EventArgs e)
		{
			this.DialogResult = DialogResult.No;
		}

		private void btnAutorizar_Click(object sender, System.EventArgs e)
		{
			#region Validacion
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridListaCheques.Rows.All)
			{
				if (dr.Cells["idCuenta"].Value == System.DBNull.Value) 
				{
					MessageBox.Show(string.Format("Seleccione una Cuenta Bancaria, '{0} {1}' ", dr.Cells["Apellidos"].Value.ToString(), dr.Cells["Nombres"].Value.ToString()), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					this.uGridListaCheques.ActiveRow = dr;
					this.uGridListaCheques.ActiveRow.Selected = true;					
					return;
				}
				if ((int)dr.Cells["idCuenta"].Value == 0) 
				{
					MessageBox.Show(string.Format("Seleccione una Cuenta Bancaria, '{0} {1}' ", dr.Cells["Apellidos"].Value.ToString(), dr.Cells["Nombres"].Value.ToString()), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					this.uGridListaCheques.ActiveRow = dr;
					this.uGridListaCheques.ActiveRow.Selected = true;					
					return;
				}

				
				if (dr.Cells["Total"].Value == System.DBNull.Value) 
				{
					MessageBox.Show(string.Format("No puede generar pago con valor 0, '{0} {1}' ", dr.Cells["Apellidos"].Value.ToString(), dr.Cells["Nombres"].Value.ToString()), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					this.uGridListaCheques.ActiveRow = dr;
					this.uGridListaCheques.ActiveRow.Selected = true;					
					return;
				}

				if (Convert.ToDecimal(dr.Cells["Total"].Value) == 0.00m)
				{
					MessageBox.Show(string.Format("No puede generar pago con valor 0, '{0} {1}' ", dr.Cells["Apellidos"].Value.ToString(), dr.Cells["Nombres"].Value.ToString()), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					this.uGridListaCheques.ActiveRow = dr;
					this.uGridListaCheques.ActiveRow.Selected = true;					
					return;
				}
			}
			#endregion Validacion

			if (DialogResult.Yes == MessageBox.Show("Esta Seguro de Autorizar los Pagos", "Point Technology", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2)) 
			{
				this.Cursor = Cursors.WaitCursor;				
				foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridListaCheques.Rows.All)
				{
					if ((int)dr.Cells["idRolPrestamo"].Value == 0)
					{
						int idTipo = 18;
						if (iRolComisionDecimo == 1) idTipo = 1;
						if (iRolComisionDecimo == 2) idTipo = 20;

						string sNumero = Funcion.sEscalarSQL(cdsSeteoF, "Select dbo.SecuencialPrestamoEmpleados(18)");
						string sMensaje = "";
						if (iRolComisionDecimo == 0) sMensaje = "SUELDOS POR PAGAR DEL PERIODO : " + dtFecha.ToString("MMMM/yyyy");
						else if (iRolComisionDecimo == 1) sMensaje = "COMISIONES POR PAGAR DEL PERIODO : " + dtFecha.ToString("MMMM/yyyy");
						else if (iRolComisionDecimo == 2) sMensaje = "DECIMO TERCER SUELDO POR PAGAR DEL PERIODO : " + dtFecha.ToString("yyyy");

						string sSQL = string.Format("Exec GuardaRolPrestamo {0}, {1}, '{2}', 2, {3}, '{4}', 0, 0, {5}, '{6}', 1", 
							(int)dr.Cells["idRolPrestamo"].Value, idTipo, sNumero, (int)dr.Cells["idPersonal"].Value, dtFecha.ToString("yyyyMMdd"), 
							Convert.ToDecimal(dr.Cells["Total"].Value), sMensaje);
						dr.Cells["idRolPrestamo"].Value = Funcion.iEscalarSQL(cdsSeteoF, sSQL);
					
						string sSQLAut = string.Format("Exec CreaAutorizacionesDePago {0}, {1}, '{2}', '{3}', {4}, 2, '1 DE 1', '{5}', {6}", 
							idTipo, (int)dr.Cells["idRolPrestamo"].Value, dtFecha.ToString("yyyyMMdd"), dtFecha.ToString("yyyyMMdd"), Convert.ToDecimal(dr.Cells["Total"].Value), sMensaje, (int)dr.Cells["idCuenta"].Value);
						Funcion.EjecutaSQL(cdsSeteoF, sSQLAut, true);

						Funcion.EjecutaSQL(cdsSeteoF, string.Format("Update AutorizacionDePagos Set Estado = 1, Autorizado = Valor, Saldo = 0, FechaAut = GETDATE() Where idTipo = {1} And idOrigen = {0}", (int)dr.Cells["idRolPrestamo"].Value, idTipo));

						Funcion.EjecutaSQL(cdsSeteoF, string.Format("Exec NominaActualizaIdAnticipo {0}, {1}, {2}", (int)dr.Cells["idRolPrestamo"].Value, (int)dr.Cells["idDetalleRolNomina"].Value, iRolComisionDecimo));
					}
				}

				Funcion.EjecutaSQL(cdsSeteoF, string.Format("Exec NominaActualizaEstadoAcreditado {0}, {1}", idRolNomina, iRolComisionDecimo));

				MessageBox.Show("Autorización Generada Correctamente", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);

				Consulta ();				

				this.Cursor = Cursors.Default;
			}
		}

		private void uGridListaCheques_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}

		private void btnExportar_Click(object sender, System.EventArgs e)
		{
			if (this.uGridListaCheques.Rows.Count == 0)
			{
				MessageBox.Show("No hay Filas para Exportar", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);				
			}
			else
			{
				this.saveFileDialog1.Filter = "Microsoft Excel (*.xls)|*.xls";

				if(saveFileDialog1.ShowDialog() == System.Windows.Forms.DialogResult.OK && saveFileDialog1.FileName.Length > 0) 
				{
					ultraGridExcelExporter1.Export(this.uGridListaCheques, saveFileDialog1.FileName);
					System.Diagnostics.Process.Start(saveFileDialog1.FileName);
				}
			}
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
			ValidTipoDePago.Value = 4;
			paramFieldidTipoDePago.CurrentValues.Add (ValidTipoDePago);
			paramFields.Add (paramFieldidTipoDePago);
																
			ParameterField paramFieldidBanco = new ParameterField ();
			ParameterDiscreteValue ValIdidBanco = new ParameterDiscreteValue ();
			paramFieldidBanco.ParameterFieldName = "@idBanco";
			ValIdidBanco.Value = 0;
			paramFieldidBanco.CurrentValues.Add (ValIdidBanco);
			paramFields.Add (paramFieldidBanco);					
			/*ImprimeRolListaPersonalCheques]
	@idRolNomina Int, @RolComision Bit, @idTipoDePago Int, @idCuenta Int*/
			Reporte miRep = new Reporte("ListaAcreditacionCheques.rpt", "");
			miRep.crVista.ParameterFieldInfo = paramFields;
			miRep.MdiParent = this.MdiParent;						
			miRep.Show();		
			#endregion Impresion Pantalla Bloqueada
		}

		private void btnAsignar_Click(object sender, System.EventArgs e)
		{
			if (this.uGridListaCheques.Rows.Count == 0)
			{
				MessageBox.Show("No existen registros para asignarles una cuenta", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
				return;
			}

			if (this.cmbCuentaAsigna.ActiveRow == null)
			{
				MessageBox.Show("Seleccione una cuenta", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
				this.cmbCuentaAsigna.Focus();
				return;
			}
			
			int i = 0;

			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in uGridListaCheques.Rows.All)
			{
				if ((int)dr.Cells["idCuenta"].Value > 0) i++;
			}

			if (i > 0 && DialogResult.No == MessageBox.Show("Ya esta asignada una Cuenta.\n\n¿Desea Modificarla?", "Point Technology", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2)) return;
			
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in uGridListaCheques.Rows.All)
			{
				dr.Cells["idCuenta"].Value = (int)this.cmbCuentaAsigna.Value;				
			}

			MessageBox.Show("Asignación Generada Correctamente", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
		}
	}
}
