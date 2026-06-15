using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de frmPresupuestoCobradoresSupervisores.
	/// </summary>
	public class frmPresupuestoCobradoresSupervisores : System.Windows.Forms.Form
	{
		private System.Windows.Forms.DateTimePicker dtpPeriodo;
		private System.Windows.Forms.Button btnGuardar;
		private System.Windows.Forms.GroupBox groupBox1;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbVendedor;
		private System.Windows.Forms.Label label7;
		private System.Windows.Forms.Button btnVer;
		private System.Windows.Forms.Label label4;
		private Infragistics.Win.UltraWinGrid.UltraGrid uGridAsignacion;
		private C1.Data.C1DataSet cdsSeteoF;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource1;
		private System.Windows.Forms.Button btnComandos;
		private System.Windows.Forms.Button button2;
		private System.Windows.Forms.ContextMenu mnuAcciones;
		private System.Windows.Forms.MenuItem mnuAccionesCopiarUltimoMes;
		private System.Windows.Forms.MenuItem mnuAccionesAgregarNuevoLocal;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public frmPresupuestoCobradoresSupervisores()
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
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmPresupuestoCobradoresSupervisores));
			Infragistics.Win.Appearance appearance1 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPersonal");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPresupuestoCobradoresSupervisores");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPersonal");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Porc_Cumplimiento");
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Porc_Comision");
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Meta_Desde");
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Meta_Hasta");
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Porc_Bono");
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn11 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bono");
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings1 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Bono", 7, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Bono", 7, true);
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance11 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance12 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance13 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance14 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance15 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance16 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idPresupuestoCobradoresSupervisores");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idPersonal");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn3 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Porc_Cumplimiento");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn4 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Porc_Comision");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn5 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Meta_Desde");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn6 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Meta_Hasta");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn7 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Porc_Bono");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn8 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Bono");
			this.dtpPeriodo = new System.Windows.Forms.DateTimePicker();
			this.btnGuardar = new System.Windows.Forms.Button();
			this.groupBox1 = new System.Windows.Forms.GroupBox();
			this.cmbVendedor = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.label7 = new System.Windows.Forms.Label();
			this.btnVer = new System.Windows.Forms.Button();
			this.label4 = new System.Windows.Forms.Label();
			this.uGridAsignacion = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource1 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.btnComandos = new System.Windows.Forms.Button();
			this.button2 = new System.Windows.Forms.Button();
			this.mnuAcciones = new System.Windows.Forms.ContextMenu();
			this.mnuAccionesCopiarUltimoMes = new System.Windows.Forms.MenuItem();
			this.mnuAccionesAgregarNuevoLocal = new System.Windows.Forms.MenuItem();
			((System.ComponentModel.ISupportInitialize)(this.cmbVendedor)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridAsignacion)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			this.SuspendLayout();
			// 
			// dtpPeriodo
			// 
			this.dtpPeriodo.CustomFormat = "MM/yyyy";
			this.dtpPeriodo.Format = System.Windows.Forms.DateTimePickerFormat.Custom;
			this.dtpPeriodo.Location = new System.Drawing.Point(69, 7);
			this.dtpPeriodo.MinDate = new System.DateTime(2016, 5, 1, 0, 0, 0, 0);
			this.dtpPeriodo.Name = "dtpPeriodo";
			this.dtpPeriodo.ShowUpDown = true;
			this.dtpPeriodo.Size = new System.Drawing.Size(96, 20);
			this.dtpPeriodo.TabIndex = 744;
			this.dtpPeriodo.Value = new System.DateTime(2016, 5, 1, 0, 0, 0, 0);
			this.dtpPeriodo.ValueChanged += new System.EventHandler(this.dtpPeriodo_ValueChanged);
			// 
			// btnGuardar
			// 
			this.btnGuardar.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(230)), ((System.Byte)(230)), ((System.Byte)(230)));
			this.btnGuardar.CausesValidation = false;
			this.btnGuardar.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.btnGuardar.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(50)), ((System.Byte)(50)), ((System.Byte)(50)));
			this.btnGuardar.Image = ((System.Drawing.Image)(resources.GetObject("btnGuardar.Image")));
			this.btnGuardar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnGuardar.Location = new System.Drawing.Point(485, 40);
			this.btnGuardar.Name = "btnGuardar";
			this.btnGuardar.Size = new System.Drawing.Size(72, 23);
			this.btnGuardar.TabIndex = 742;
			this.btnGuardar.Text = "Guardar";
			this.btnGuardar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnGuardar.Visible = false;
			this.btnGuardar.Click += new System.EventHandler(this.btnGuardar_Click);
			// 
			// groupBox1
			// 
			this.groupBox1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.groupBox1.Location = new System.Drawing.Point(-3, 71);
			this.groupBox1.Name = "groupBox1";
			this.groupBox1.Size = new System.Drawing.Size(872, 8);
			this.groupBox1.TabIndex = 741;
			this.groupBox1.TabStop = false;
			// 
			// cmbVendedor
			// 
			appearance1.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbVendedor.Appearance = appearance1;
			this.cmbVendedor.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbVendedor.Cursor = System.Windows.Forms.Cursors.Default;
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridColumn2.Width = 335;
			ultraGridColumn3.Header.VisiblePosition = 2;
			ultraGridBand1.Columns.AddRange(new object[] {
																										 ultraGridColumn1,
																										 ultraGridColumn2,
																										 ultraGridColumn3});
			this.cmbVendedor.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.cmbVendedor.DisplayMember = "Nombre";
			this.cmbVendedor.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbVendedor.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbVendedor.Location = new System.Drawing.Point(69, 39);
			this.cmbVendedor.Name = "cmbVendedor";
			this.cmbVendedor.Size = new System.Drawing.Size(320, 21);
			this.cmbVendedor.TabIndex = 739;
			this.cmbVendedor.ValueMember = "idPersonal";
			this.cmbVendedor.ValueChanged += new System.EventHandler(this.cmbVendedor_ValueChanged);
			// 
			// label7
			// 
			this.label7.AutoSize = true;
			this.label7.Location = new System.Drawing.Point(5, 41);
			this.label7.Name = "label7";
			this.label7.Size = new System.Drawing.Size(44, 16);
			this.label7.TabIndex = 740;
			this.label7.Text = "Nombre";
			// 
			// btnVer
			// 
			this.btnVer.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(230)), ((System.Byte)(230)), ((System.Byte)(230)));
			this.btnVer.CausesValidation = false;
			this.btnVer.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.btnVer.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(50)), ((System.Byte)(50)), ((System.Byte)(50)));
			this.btnVer.Image = ((System.Drawing.Image)(resources.GetObject("btnVer.Image")));
			this.btnVer.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnVer.Location = new System.Drawing.Point(405, 40);
			this.btnVer.Name = "btnVer";
			this.btnVer.Size = new System.Drawing.Size(72, 23);
			this.btnVer.TabIndex = 738;
			this.btnVer.Text = "Ver";
			this.btnVer.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnVer.Click += new System.EventHandler(this.btnVer_Click);
			// 
			// label4
			// 
			this.label4.AutoSize = true;
			this.label4.BackColor = System.Drawing.Color.Transparent;
			this.label4.Location = new System.Drawing.Point(5, 9);
			this.label4.Name = "label4";
			this.label4.Size = new System.Drawing.Size(43, 16);
			this.label4.TabIndex = 737;
			this.label4.Text = "Periodo";
			this.label4.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// uGridAsignacion
			// 
			this.uGridAsignacion.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.uGridAsignacion.Cursor = System.Windows.Forms.Cursors.Default;
			this.uGridAsignacion.DataSource = this.ultraDataSource1;
			appearance2.BackColor = System.Drawing.Color.White;
			this.uGridAsignacion.DisplayLayout.Appearance = appearance2;
			this.uGridAsignacion.DisplayLayout.AutoFitColumns = true;
			ultraGridBand2.AddButtonCaption = "BodegaPromocion";
			ultraGridColumn4.Header.VisiblePosition = 0;
			ultraGridColumn4.Hidden = true;
			ultraGridColumn4.Width = 218;
			ultraGridColumn5.Header.VisiblePosition = 1;
			ultraGridColumn5.Hidden = true;
			ultraGridColumn5.Width = 70;
			appearance3.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn6.CellAppearance = appearance3;
			ultraGridColumn6.Format = "#,##0.00";
			ultraGridColumn6.Header.Caption = "% Cumplimiento";
			ultraGridColumn6.Header.VisiblePosition = 2;
			ultraGridColumn6.Width = 180;
			appearance4.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn7.CellAppearance = appearance4;
			ultraGridColumn7.Format = "#,##0.00";
			ultraGridColumn7.Header.Caption = "% Comisión";
			ultraGridColumn7.Header.VisiblePosition = 3;
			ultraGridColumn7.Width = 142;
			appearance5.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn8.CellAppearance = appearance5;
			ultraGridColumn8.Format = "#,##0.00";
			ultraGridColumn8.Header.Caption = "Meta Desde";
			ultraGridColumn8.Header.VisiblePosition = 4;
			ultraGridColumn8.Width = 127;
			appearance6.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn9.CellAppearance = appearance6;
			ultraGridColumn9.Format = "#,##0.00";
			ultraGridColumn9.Header.Caption = "Meta Hasta";
			ultraGridColumn9.Header.VisiblePosition = 5;
			ultraGridColumn9.Width = 127;
			appearance7.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn10.CellAppearance = appearance7;
			ultraGridColumn10.Format = "#,##0.00";
			ultraGridColumn10.Header.Caption = "% Bono";
			ultraGridColumn10.Header.VisiblePosition = 6;
			ultraGridColumn10.Width = 127;
			appearance8.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn11.CellAppearance = appearance8;
			ultraGridColumn11.Format = "#,##0.00";
			ultraGridColumn11.Header.VisiblePosition = 7;
			ultraGridColumn11.Width = 127;
			ultraGridBand2.Columns.AddRange(new object[] {
																										 ultraGridColumn4,
																										 ultraGridColumn5,
																										 ultraGridColumn6,
																										 ultraGridColumn7,
																										 ultraGridColumn8,
																										 ultraGridColumn9,
																										 ultraGridColumn10,
																										 ultraGridColumn11});
			ultraGridBand2.Header.Caption = "";
			appearance9.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings1.Appearance = appearance9;
			summarySettings1.DisplayFormat = "{0:#,##0.00}";
			summarySettings1.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			ultraGridBand2.Summaries.AddRange(new Infragistics.Win.UltraWinGrid.SummarySettings[] {
																																															summarySettings1});
			ultraGridBand2.SummaryFooterCaption = "";
			this.uGridAsignacion.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			appearance10.ForeColor = System.Drawing.Color.Black;
			appearance10.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridAsignacion.DisplayLayout.Override.ActiveRowAppearance = appearance10;
			this.uGridAsignacion.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.TemplateOnBottom;
			this.uGridAsignacion.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			appearance11.BackColor = System.Drawing.Color.Transparent;
			this.uGridAsignacion.DisplayLayout.Override.CardAreaAppearance = appearance11;
			appearance12.ForeColor = System.Drawing.Color.Black;
			appearance12.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridAsignacion.DisplayLayout.Override.CellAppearance = appearance12;
			appearance13.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance13.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance13.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance13.FontData.BoldAsString = "True";
			appearance13.FontData.Name = "Arial";
			appearance13.FontData.SizeInPoints = 8.5F;
			appearance13.ForeColor = System.Drawing.Color.White;
			appearance13.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridAsignacion.DisplayLayout.Override.HeaderAppearance = appearance13;
			appearance14.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance14.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance14.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridAsignacion.DisplayLayout.Override.RowAlternateAppearance = appearance14;
			appearance15.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance15.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance15.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridAsignacion.DisplayLayout.Override.RowSelectorAppearance = appearance15;
			appearance16.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance16.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance16.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridAsignacion.DisplayLayout.Override.SelectedRowAppearance = appearance16;
			this.uGridAsignacion.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridAsignacion.Location = new System.Drawing.Point(5, 87);
			this.uGridAsignacion.Name = "uGridAsignacion";
			this.uGridAsignacion.Size = new System.Drawing.Size(851, 432);
			this.uGridAsignacion.TabIndex = 736;
			this.uGridAsignacion.AfterRowInsert += new Infragistics.Win.UltraWinGrid.RowEventHandler(this.uGridAsignacion_AfterRowInsert);
			this.uGridAsignacion.BeforeRowsDeleted += new Infragistics.Win.UltraWinGrid.BeforeRowsDeletedEventHandler(this.uGridAsignacion_BeforeRowsDeleted);
			this.uGridAsignacion.AfterCellUpdate += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.uGridAsignacion_AfterCellUpdate);
			// 
			// ultraDataSource1
			// 
			ultraDataColumn1.DataType = typeof(int);
			ultraDataColumn2.DataType = typeof(int);
			ultraDataColumn3.DataType = typeof(System.Decimal);
			ultraDataColumn4.DataType = typeof(System.Decimal);
			ultraDataColumn5.DataType = typeof(System.Decimal);
			ultraDataColumn6.DataType = typeof(System.Decimal);
			ultraDataColumn7.DataType = typeof(System.Decimal);
			ultraDataColumn8.DataType = typeof(System.Decimal);
			this.ultraDataSource1.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn1,
																																 ultraDataColumn2,
																																 ultraDataColumn3,
																																 ultraDataColumn4,
																																 ultraDataColumn5,
																																 ultraDataColumn6,
																																 ultraDataColumn7,
																																 ultraDataColumn8});
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
			// btnComandos
			// 
			this.btnComandos.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btnComandos.Image = ((System.Drawing.Image)(resources.GetObject("btnComandos.Image")));
			this.btnComandos.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnComandos.Location = new System.Drawing.Point(776, 8);
			this.btnComandos.Name = "btnComandos";
			this.btnComandos.TabIndex = 978;
			this.btnComandos.Text = "Acciones";
			this.btnComandos.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnComandos.Visible = false;
			this.btnComandos.Click += new System.EventHandler(this.btnComandos_Click);
			// 
			// button2
			// 
			this.button2.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.button2.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(230)), ((System.Byte)(230)), ((System.Byte)(230)));
			this.button2.CausesValidation = false;
			this.button2.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.button2.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(50)), ((System.Byte)(50)), ((System.Byte)(50)));
			this.button2.Image = ((System.Drawing.Image)(resources.GetObject("button2.Image")));
			this.button2.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.button2.Location = new System.Drawing.Point(776, 40);
			this.button2.Name = "button2";
			this.button2.TabIndex = 979;
			this.button2.Text = "Salir";
			this.button2.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.button2.Click += new System.EventHandler(this.button2_Click);
			// 
			// mnuAcciones
			// 
			this.mnuAcciones.MenuItems.AddRange(new System.Windows.Forms.MenuItem[] {
																																								this.mnuAccionesCopiarUltimoMes,
																																								this.mnuAccionesAgregarNuevoLocal});
			// 
			// mnuAccionesCopiarUltimoMes
			// 
			this.mnuAccionesCopiarUltimoMes.Index = 0;
			this.mnuAccionesCopiarUltimoMes.Text = "Copiar Ultimo Mes";
			this.mnuAccionesCopiarUltimoMes.Click += new System.EventHandler(this.mnuAccionesCopiarUltimoMes_Click);
			// 
			// mnuAccionesAgregarNuevoLocal
			// 
			this.mnuAccionesAgregarNuevoLocal.Index = 1;
			this.mnuAccionesAgregarNuevoLocal.Text = "Exportar";
			this.mnuAccionesAgregarNuevoLocal.Click += new System.EventHandler(this.mnuAccionesAgregarNuevoLocal_Click);
			// 
			// frmPresupuestoCobradoresSupervisores
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(867, 526);
			this.Controls.Add(this.button2);
			this.Controls.Add(this.btnComandos);
			this.Controls.Add(this.dtpPeriodo);
			this.Controls.Add(this.btnGuardar);
			this.Controls.Add(this.groupBox1);
			this.Controls.Add(this.cmbVendedor);
			this.Controls.Add(this.label7);
			this.Controls.Add(this.btnVer);
			this.Controls.Add(this.label4);
			this.Controls.Add(this.uGridAsignacion);
			this.KeyPreview = true;
			this.Name = "frmPresupuestoCobradoresSupervisores";
			this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
			this.Text = "Asignaciones Mensual De Jefes Y Supervisores De cobranzas";
			this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.frmPresupuestoCobradoresSupervisores_KeyDown);
			this.Load += new System.EventHandler(this.frmPresupuestoCobradoresSupervisores_Load);
			((System.ComponentModel.ISupportInitialize)(this.cmbVendedor)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridAsignacion)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private Acceso miAcceso;
		
		private void UnloadMe()
		{
			this.Close();
		}


		private void frmPresupuestoCobradoresSupervisores_Load(object sender, System.EventArgs e)
		{
			miAcceso = new Acceso(cdsSeteoF, "083609");

			if (!Funcion.Permiso("1841", cdsSeteoF))
			{		
				MessageBox.Show("No tiene Acceso a Porcentaje de Supervisores", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				BeginInvoke(new MethodInvoker(UnloadMe));
				return;
			}

			if(miAcceso.BExportar)this.btnComandos.Visible = true;
			if(miAcceso.Nuevo)this.btnGuardar.Visible = true;

			this.dtpPeriodo.MaxDate = DateTime.Parse(DateTime.DaysInMonth(DateTime.Now.Year, DateTime.Now.Month).ToString() + "/" + DateTime.Now.Month.ToString() + "/" + DateTime.Now.Year.ToString());
			this.dtpPeriodo.Value = DateTime.Parse("01/" + DateTime.Now.Month.ToString() + "/" + DateTime.Now.Year.ToString());
		}

		private void cmbVendedor_ValueChanged(object sender, System.EventArgs e)
		{
			Consultar ();
		}

		private void Consultar ()
		{
			int idVendedor = 0;
			if (this.cmbVendedor.ActiveRow != null) idVendedor = (int)this.cmbVendedor.Value;

			string sSQL = string.Format("Exec [ConsultaGeneralListaCobradores] '{0}', {1}", this.dtpPeriodo.Text.ToString(), idVendedor);
			this.uGridAsignacion.DataSource = FuncionesProcedimientos.dtGeneral(sSQL);

		}


		private void btnVer_Click(object sender, System.EventArgs e)
		{
			if (this.dtpPeriodo.Text.ToString().Trim().Length == 0)
			{
				MessageBox.Show("Seleccione un periodo", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.dtpPeriodo.Focus();
				return;
			}

			if (!Validacion.vbComboVacio(this.cmbVendedor, "Seleccione el Vendedor")) return;
			
			Consultar();
		}

		private void dtpPeriodo_ValueChanged(object sender, System.EventArgs e)
		{
			Consultar ();

			DateTime dtDesde = DateTime.Parse("01/" + DateTime.Now.Month.ToString() + "/" + DateTime.Now.Year.ToString());
			DateTime dtHasta = DateTime.Parse(DateTime.DaysInMonth(DateTime.Now.Year, DateTime.Now.Month).ToString() + "/" + DateTime.Now.Month.ToString() + "/" + DateTime.Now.Year.ToString());
			this.cmbVendedor.DataSource = Funcion.dvProcedimiento(cdsSeteoF, string.Format("Exec [ListaCobradores] "));

		}

		private void btnGuardar_Click(object sender, System.EventArgs e)
		{
			if (this.dtpPeriodo.Text.ToString().Trim().Length == 0)
			{
				MessageBox.Show("Seleccione un periodo", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.dtpPeriodo.Focus();
				return;
			}
			if (!Validacion.vbComboVacio(this.cmbVendedor, "Seleccione el Vendedor")) return;
			
			DateTime dtDesde = DateTime.Parse("01/" +  "/" + this.dtpPeriodo.Value.Month.ToString()+  "/" + this.dtpPeriodo.Value.Year.ToString());
			//DateTime dtHasta = DateTime.Parse(DateTime.DaysInMonth(DateTime.Now.Year, DateTime.Now.Month).ToString() + "/" + DateTime.Now.Month.ToString() + "/" + DateTime.Now.Year.ToString());
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridAsignacion.Rows.All)
			{
				if (dr.Cells["Porc_Cumplimiento"].Value == System.DBNull.Value)
				{
					MessageBox.Show("Ingrese el % Cumplimiento ", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					this.uGridAsignacion.ActiveRow = dr;
					this.uGridAsignacion.ActiveRow.Selected = true;
					return;
				}
				if (dr.Cells["Porc_Comision"].Value == System.DBNull.Value)
				{
					MessageBox.Show("Ingrese el % Comision ", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					this.uGridAsignacion.ActiveRow = dr;
					this.uGridAsignacion.ActiveRow.Selected = true;
					return;
				}
				if (dr.Cells["Meta_Desde"].Value == System.DBNull.Value)
				{
					MessageBox.Show("Ingrese la Meta Desde ", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					this.uGridAsignacion.ActiveRow = dr;
					this.uGridAsignacion.ActiveRow.Selected = true;
					return;
				}
				if (dr.Cells["Meta_Hasta"].Value == System.DBNull.Value)
				{
					MessageBox.Show("Ingrese la Meta Hasta ", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					this.uGridAsignacion.ActiveRow = dr;
					this.uGridAsignacion.ActiveRow.Selected = true;
					return;
				}
				if (dr.Cells["Porc_Bono"].Value == System.DBNull.Value)
				{
					MessageBox.Show("Ingrese el % Bono ", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					this.uGridAsignacion.ActiveRow = dr;
					this.uGridAsignacion.ActiveRow.Selected = true;
					return;
				}
				if (dr.Cells["Bono"].Value == System.DBNull.Value)
				{
					MessageBox.Show("Ingrese el Bono ", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					this.uGridAsignacion.ActiveRow = dr;
					this.uGridAsignacion.ActiveRow.Selected = true;
					return;
				}

				if (((decimal)dr.Cells["Porc_Cumplimiento"].Value + (decimal) dr.Cells["Porc_Comision"].Value ) > 100 )
				{
					MessageBox.Show("La suma % Comisión y Cumplimiento no puede ser mayor a 100. ", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					this.uGridAsignacion.ActiveRow = dr;
					this.uGridAsignacion.ActiveRow.Selected = true;
					return;
				}
				
			}
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridAsignacion.Rows.All)
			{
				string sSQL = string.Format("Exec [Com_GuardaAsignacionJefesCobradores] {0}, '{1}', {2}, '{3}', '{4}', {5}, {6}, {7}, {8}, {9}, {10}", 
					(int)dr.Cells["idPresupuestoCobradoresSupervisores"].Value, this.dtpPeriodo.Text.ToString(), (int)this.cmbVendedor.Value, 
					Convert.ToDateTime(dtDesde).ToString("yyyyMMdd"), 
					Convert.ToDateTime(dtDesde).ToString("yyyyMMdd"),
					Convert.ToDecimal(dr.Cells["Porc_Cumplimiento"].Value),
					Convert.ToDecimal(dr.Cells["Porc_Comision"].Value),
					Convert.ToDecimal(dr.Cells["Meta_Desde"].Value),
					Convert.ToDecimal(dr.Cells["Meta_Hasta"].Value),
					Convert.ToDecimal(dr.Cells["Porc_Bono"].Value),
					Convert.ToDecimal(dr.Cells["Bono"].Value)
					);
				Funcion.EjecutaSQL(cdsSeteoF, sSQL);
			}

			this.btnVer_Click(sender, e);
			
			MessageBox.Show("Asignacion Guardada Correctamente", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);

		}

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);
		}

		private void uGridAsignacion_AfterRowInsert(object sender, Infragistics.Win.UltraWinGrid.RowEventArgs e)
		{
			e.Row.Cells["idPresupuestoCobradoresSupervisores"].Value = 0;
			//e.Row.Cells["Periodo"].Value = "";
			e.Row.Cells["idPersonal"].Value = 0;
			e.Row.Cells["Porc_Cumplimiento"].Value = 0.00m;  
			e.Row.Cells["Porc_Comision"].Value = 0.00m; 	
			e.Row.Cells["Meta_Desde"].Value = 0.00m;
			e.Row.Cells["Meta_Hasta"].Value = 0.00m; 
			e.Row.Cells["Porc_Bono"].Value = 0.00m; 
			e.Row.Cells["Bono"].Value = 0.00m; 
		}

		private void frmPresupuestoCobradoresSupervisores_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Escape) this.Close();
		}

		private void button2_Click(object sender, System.EventArgs e)
		{
			this.Close();
		}

		private void mnuAccionesAgregarNuevoLocal_Click(object sender, System.EventArgs e)
		{
			FuncionesProcedimientos.ExportaExcel(this.uGridAsignacion);
	
		}

		private void uGridAsignacion_BeforeRowsDeleted(object sender, Infragistics.Win.UltraWinGrid.BeforeRowsDeletedEventArgs e)
		{
			if (!miAcceso.BEliminaSeriales) 
			{
				e.Cancel = true;
			}
			else
			{
				if (DialogResult.Yes ==	MessageBox.Show("¿Esta seguro de Borrar las Líneas Seleccionadas?", "Point Technology", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2))				
				{
					for (int i = 0; i < e.Rows.Length; i++)
					{
						if (e.Rows[i].Cells["idPresupuestoCobradoresSupervisores"].Value != System.DBNull.Value)
						{									
							if ((int)e.Rows[i].Cells["idPresupuestoCobradoresSupervisores"].Value > 0)
							{
								string sSQL = string.Format("Delete From PresupuestoCobradoresSupervisores Where idPresupuestoCobradoresSupervisores = {0}", (int)e.Rows[i].Cells["idPresupuestoCobradoresSupervisores"].Value);
								Funcion.EjecutaSQL(cdsSeteoF, sSQL, true);
							}
						}			
					}
					e.DisplayPromptMsg = false;
				}
				else
				{
					e.Cancel = true;
				}
			
			}
		}

		private void mnuAccionesCopiarUltimoMes_Click(object sender, System.EventArgs e)
		{
			int iRetorna = 0;
			if (this.dtpPeriodo.Text.ToString().Trim().Length == 0)
			{
				MessageBox.Show("Seleccione un periodo", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.dtpPeriodo.Focus();
				return;
			}

			string sPeriodo = Funcion.sEscalarSQL(cdsSeteoF, "Select top 1 Periodo from dbo.PresupuestoCobradoresSupervisores order by  idPresupuestoCobradoresSupervisores desc");

			if (DialogResult.Yes == MessageBox.Show("¿Esta Seguro De Copiar La Información del anterior Mes?", "Point Technology", MessageBoxButtons.YesNo, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2))
			{	

				iRetorna = Funcion.iEscalarSQL(cdsSeteoF, string.Format("Exec CopiaAsignacioncobradoresPorcentajeSupervisor "));
				//	this.uGridDepositos.DataSource = FuncionesProcedimientos.dtGeneral(sSQLPa);
				if ( iRetorna == 1)
				{

					MessageBox.Show(string.Format("Se copio la Información del Periodo {0} correctamente", sPeriodo), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
					//this.btnVer_Click(sender, e);
				}
				else
				{
					MessageBox.Show(string.Format("No existen Datos en el Mes anterior {0} o ya se genero una copia en el mes actual ", sPeriodo), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
					//this.btnVer_Click(sender, e);

				}
			}
		}
		bool bActualiza = false;

		private void uGridAsignacion_AfterCellUpdate(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			if (bActualiza) return;

			if (e.Cell.Column.ToString() == "Porc_Cumplimiento")
			{	
				bActualiza = true;

				if ( e.Cell.Row.Cells["Porc_Cumplimiento"].Value == System.DBNull.Value)
				{											
					e.Cell.Row.Cells["Porc_Cumplimiento"].Value = (decimal)e.Cell.OriginalValue;						
				}	
				bActualiza = false;
			}

			if (e.Cell.Column.ToString() == "Porc_Comision")
			{	
				bActualiza = true;

				if ( e.Cell.Row.Cells["Porc_Comision"].Value == System.DBNull.Value)
				{											
					e.Cell.Row.Cells["Porc_Comision"].Value = (decimal)e.Cell.OriginalValue;						
				}	
				bActualiza = false;
			}

			if (e.Cell.Column.ToString() == "Meta_Desde")
			{	
				bActualiza = true;

				if ( e.Cell.Row.Cells["Meta_Desde"].Value == System.DBNull.Value)
				{											
					e.Cell.Row.Cells["Meta_Desde"].Value = (decimal)e.Cell.OriginalValue;						
				}	
				bActualiza = false;
			}

			if (e.Cell.Column.ToString() == "Meta_Hasta")
			{	
				bActualiza = true;

				if ( e.Cell.Row.Cells["Meta_Hasta"].Value == System.DBNull.Value)
				{											
					e.Cell.Row.Cells["Meta_Hasta"].Value = (decimal)e.Cell.OriginalValue;						
				}	
				bActualiza = false;
			}

			if (e.Cell.Column.ToString() == "Porc_Bono")
			{	
				bActualiza = true;

				if ( e.Cell.Row.Cells["Porc_Bono"].Value == System.DBNull.Value)
				{											
					e.Cell.Row.Cells["Porc_Bono"].Value = (decimal)e.Cell.OriginalValue;						
				}	
				bActualiza = false;
			}

			if (e.Cell.Column.ToString() == "Bono")
			{	
				bActualiza = true;

				if ( e.Cell.Row.Cells["Bono"].Value == System.DBNull.Value)
				{											
					e.Cell.Row.Cells["Bono"].Value = (decimal)e.Cell.OriginalValue;						
				}	
				bActualiza = false;
			}
		}

		private void btnComandos_Click(object sender, System.EventArgs e)
		{
		this.mnuAcciones.Show(this.btnComandos, new Point(1, 25));
		}
	}
}
