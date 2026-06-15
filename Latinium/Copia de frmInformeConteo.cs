using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using DevExpress.XtraEditors;
using System.Data;
using System.Data.SqlClient;
using Infragistics.Win.UltraWinGrid;
using CrystalDecisions;
using CrystalDecisions.CrystalReports.Engine;
using CrystalDecisions.ReportSource;
using CrystalDecisions.Shared;

namespace Latinium
{
	/// <summary>
	/// Summary description for frmInformeConteo.
	/// </summary>
	public class frmInformeConteo : DevExpress.XtraEditors.XtraForm
	{
		private System.Windows.Forms.Button btnGuardar;
		private C1.Data.C1DataSet cdsSeteoF;
		private Infragistics.Win.UltraWinGrid.UltraGrid uGridInforme;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource1;
		private Infragistics.Win.UltraWinGrid.UltraDropDown cmbEstados;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource2;
		public Infragistics.Win.UltraWinEditors.UltraNumericEditor txtidCierre;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtEstado;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkPaso1;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkPaso2;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkPaso3;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkPaso4;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkPaso5;
		private System.Windows.Forms.Label lblContador;
		private System.Windows.Forms.Button btnExcel;
		private System.Windows.Forms.Button btnImprimir;
		private System.Windows.Forms.GroupBox groupBox1;
		private System.Windows.Forms.Label lblEstado;
		private System.Windows.Forms.Button btnParticipantes;
		private System.Windows.Forms.Label lblFecha;
		public Infragistics.Win.UltraWinEditors.UltraTextEditor txtNotas;
		private System.Windows.Forms.Button btnCalificar;
		/// <summary>
		/// Required designer variable.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public frmInformeConteo()
		{
			//
			// Required for Windows Form Designer support
			//
			InitializeComponent();

			//
			// TODO: Add any constructor code after InitializeComponent call
			//
		}
		
		DateTime dtFecha = DateTime.Now;
		int Bodega = 0;

		public frmInformeConteo(DateTime Fecha, int iBodega)
		{
			//
			// Required for Windows Form Designer support
			//
			InitializeComponent();

			dtFecha = Fecha;
			Bodega = iBodega;
			//
			// TODO: Add any constructor code after InitializeComponent call
			//
		}

		/// <summary>
		/// Clean up any resources being used.
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
		/// Required method for Designer support - do not modify
		/// the contents of this method with the code editor.
		/// </summary>
		private void InitializeComponent()
		{
			Infragistics.Win.Appearance appearance1 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idLogAjusteCantidades");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCierreConteoInventario");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idArticulo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Existencia");
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Conteo");
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ConteoExistencia");
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CantidadJustifica");
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idEstado", -1, "cmbEstados");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nota");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn11 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("NotasDelSistema");
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idLogAjusteCantidades");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCierreConteoInventario");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn3 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idArticulo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn4 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Codigo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn5 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Existencia");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn6 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Conteo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn7 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("ConteoExistencia");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn8 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("CantidadJustifica");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn9 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idEstado");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn10 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nota");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn11 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("NotasDelSistema");
			Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance11 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance12 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance13 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance14 = new Infragistics.Win.Appearance();
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmInformeConteo));
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn12 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idEstado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn13 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Estado");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn12 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idEstado");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn13 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Estado");
			Infragistics.Win.Appearance appearance15 = new Infragistics.Win.Appearance();
			this.uGridInforme = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource1 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.chkPaso1 = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.chkPaso2 = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.chkPaso4 = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.chkPaso3 = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.chkPaso5 = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.btnGuardar = new System.Windows.Forms.Button();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.cmbEstados = new Infragistics.Win.UltraWinGrid.UltraDropDown();
			this.ultraDataSource2 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.txtidCierre = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtEstado = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.lblContador = new System.Windows.Forms.Label();
			this.btnExcel = new System.Windows.Forms.Button();
			this.btnImprimir = new System.Windows.Forms.Button();
			this.groupBox1 = new System.Windows.Forms.GroupBox();
			this.lblEstado = new System.Windows.Forms.Label();
			this.btnParticipantes = new System.Windows.Forms.Button();
			this.lblFecha = new System.Windows.Forms.Label();
			this.txtNotas = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.btnCalificar = new System.Windows.Forms.Button();
			((System.ComponentModel.ISupportInitialize)(this.uGridInforme)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbEstados)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtidCierre)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtEstado)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNotas)).BeginInit();
			this.SuspendLayout();
			// 
			// uGridInforme
			// 
			this.uGridInforme.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.uGridInforme.Cursor = System.Windows.Forms.Cursors.Default;
			this.uGridInforme.DataSource = this.ultraDataSource1;
			appearance1.BackColor = System.Drawing.Color.White;
			this.uGridInforme.DisplayLayout.Appearance = appearance1;
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn1.Width = 152;
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridColumn2.Hidden = true;
			ultraGridColumn2.Width = 111;
			ultraGridColumn3.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn3.Header.VisiblePosition = 2;
			ultraGridColumn3.Hidden = true;
			ultraGridColumn3.Width = 44;
			ultraGridColumn4.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn4.Header.Caption = "Articulo";
			ultraGridColumn4.Header.VisiblePosition = 3;
			ultraGridColumn4.Width = 180;
			ultraGridColumn5.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance2.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn5.CellAppearance = appearance2;
			ultraGridColumn5.Header.VisiblePosition = 4;
			ultraGridColumn5.Width = 75;
			ultraGridColumn6.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance3.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn6.CellAppearance = appearance3;
			ultraGridColumn6.Header.VisiblePosition = 5;
			ultraGridColumn6.Width = 75;
			ultraGridColumn7.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance4.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn7.CellAppearance = appearance4;
			ultraGridColumn7.Header.Caption = "Diferencias";
			ultraGridColumn7.Header.VisiblePosition = 6;
			ultraGridColumn7.Width = 75;
			appearance5.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn8.CellAppearance = appearance5;
			ultraGridColumn8.Header.Caption = "Cant Cruce";
			ultraGridColumn8.Header.VisiblePosition = 8;
			ultraGridColumn8.Width = 75;
			ultraGridColumn9.Header.Caption = "Estado";
			ultraGridColumn9.Header.VisiblePosition = 7;
			ultraGridColumn9.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownList;
			ultraGridColumn9.Width = 130;
			ultraGridColumn10.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn10.Header.VisiblePosition = 9;
			ultraGridColumn10.Width = 307;
			ultraGridColumn11.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn11.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn11.Header.Caption = "Nota Del Sistema";
			ultraGridColumn11.Header.VisiblePosition = 10;
			ultraGridColumn11.Width = 300;
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
																										 ultraGridColumn11});
			this.uGridInforme.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.uGridInforme.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
			this.uGridInforme.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			appearance6.BackColor = System.Drawing.Color.Transparent;
			this.uGridInforme.DisplayLayout.Override.CardAreaAppearance = appearance6;
			appearance7.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance7.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance7.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance7.FontData.BoldAsString = "True";
			appearance7.FontData.Name = "Tahoma";
			appearance7.FontData.SizeInPoints = 8F;
			appearance7.ForeColor = System.Drawing.Color.White;
			appearance7.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridInforme.DisplayLayout.Override.HeaderAppearance = appearance7;
			this.uGridInforme.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti;
			appearance8.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance8.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance8.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridInforme.DisplayLayout.Override.RowSelectorAppearance = appearance8;
			appearance9.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance9.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance9.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridInforme.DisplayLayout.Override.SelectedRowAppearance = appearance9;
			this.uGridInforme.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridInforme.Location = new System.Drawing.Point(8, 152);
			this.uGridInforme.Name = "uGridInforme";
			this.uGridInforme.Size = new System.Drawing.Size(938, 328);
			this.uGridInforme.TabIndex = 367;
			this.uGridInforme.DoubleClick += new System.EventHandler(this.uGridInforme_DoubleClick);
			this.uGridInforme.BeforeCellUpdate += new Infragistics.Win.UltraWinGrid.BeforeCellUpdateEventHandler(this.uGridInforme_BeforeCellUpdate);
			this.uGridInforme.KeyDown += new System.Windows.Forms.KeyEventHandler(this.uGridInforme_KeyDown);
			this.uGridInforme.AfterCellUpdate += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.uGridInforme_AfterCellUpdate);
			this.uGridInforme.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.uGridInforme_InitializeLayout);
			// 
			// ultraDataSource1
			// 
			ultraDataColumn1.DataType = typeof(int);
			ultraDataColumn2.DataType = typeof(int);
			ultraDataColumn3.DataType = typeof(int);
			ultraDataColumn5.DataType = typeof(int);
			ultraDataColumn6.DataType = typeof(int);
			ultraDataColumn7.DataType = typeof(int);
			ultraDataColumn8.DataType = typeof(int);
			ultraDataColumn9.DataType = typeof(int);
			this.ultraDataSource1.Band.Columns.AddRange(new object[] {
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
																																 ultraDataColumn11});
			// 
			// chkPaso1
			// 
			appearance10.FontData.Name = "Tahoma";
			appearance10.ForeColorDisabled = System.Drawing.Color.Black;
			this.chkPaso1.Appearance = appearance10;
			this.chkPaso1.BackColor = System.Drawing.Color.Transparent;
			this.chkPaso1.Location = new System.Drawing.Point(16, 8);
			this.chkPaso1.Name = "chkPaso1";
			this.chkPaso1.Size = new System.Drawing.Size(336, 22);
			this.chkPaso1.TabIndex = 368;
			this.chkPaso1.Text = "Recuente Faltantes y Sobrantes Detectados";
			this.chkPaso1.CheckedChanged += new System.EventHandler(this.chkPaso1_CheckedChanged);
			// 
			// chkPaso2
			// 
			appearance11.FontData.Name = "Tahoma";
			appearance11.ForeColorDisabled = System.Drawing.Color.Black;
			this.chkPaso2.Appearance = appearance11;
			this.chkPaso2.BackColor = System.Drawing.Color.Transparent;
			this.chkPaso2.Enabled = false;
			this.chkPaso2.Location = new System.Drawing.Point(16, 32);
			this.chkPaso2.Name = "chkPaso2";
			this.chkPaso2.Size = new System.Drawing.Size(336, 22);
			this.chkPaso2.TabIndex = 369;
			this.chkPaso2.Text = "Verifique Productos que no estan en Conteo y Capture la Serial";
			this.chkPaso2.CheckedChanged += new System.EventHandler(this.chkPaso2_CheckedChanged);
			// 
			// chkPaso4
			// 
			appearance12.FontData.Name = "Tahoma";
			appearance12.ForeColorDisabled = System.Drawing.Color.Black;
			this.chkPaso4.Appearance = appearance12;
			this.chkPaso4.BackColor = System.Drawing.Color.Transparent;
			this.chkPaso4.Enabled = false;
			this.chkPaso4.Location = new System.Drawing.Point(16, 80);
			this.chkPaso4.Name = "chkPaso4";
			this.chkPaso4.Size = new System.Drawing.Size(336, 22);
			this.chkPaso4.TabIndex = 371;
			this.chkPaso4.Text = "Verifique si hay Faltantes Justificados";
			this.chkPaso4.CheckedChanged += new System.EventHandler(this.chkPaso4_CheckedChanged);
			// 
			// chkPaso3
			// 
			appearance13.FontData.Name = "Tahoma";
			appearance13.ForeColorDisabled = System.Drawing.Color.Black;
			this.chkPaso3.Appearance = appearance13;
			this.chkPaso3.BackColor = System.Drawing.Color.Transparent;
			this.chkPaso3.Enabled = false;
			this.chkPaso3.Location = new System.Drawing.Point(16, 56);
			this.chkPaso3.Name = "chkPaso3";
			this.chkPaso3.Size = new System.Drawing.Size(336, 22);
			this.chkPaso3.TabIndex = 370;
			this.chkPaso3.Text = "Errores en Captura Cierre y Vuelva a Capturar";
			this.chkPaso3.CheckedChanged += new System.EventHandler(this.chkPaso3_CheckedChanged);
			// 
			// chkPaso5
			// 
			appearance14.FontData.Name = "Tahoma";
			appearance14.ForeColorDisabled = System.Drawing.Color.Black;
			this.chkPaso5.Appearance = appearance14;
			this.chkPaso5.BackColor = System.Drawing.Color.Transparent;
			this.chkPaso5.Enabled = false;
			this.chkPaso5.Location = new System.Drawing.Point(16, 104);
			this.chkPaso5.Name = "chkPaso5";
			this.chkPaso5.Size = new System.Drawing.Size(336, 22);
			this.chkPaso5.TabIndex = 372;
			this.chkPaso5.Text = "Actualice Estados de Cruces y Faltantes";
			this.chkPaso5.CheckedChanged += new System.EventHandler(this.chkPaso5_CheckedChanged);
			// 
			// btnGuardar
			// 
			this.btnGuardar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnGuardar.Enabled = false;
			this.btnGuardar.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.btnGuardar.Image = ((System.Drawing.Image)(resources.GetObject("btnGuardar.Image")));
			this.btnGuardar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnGuardar.Location = new System.Drawing.Point(832, 104);
			this.btnGuardar.Name = "btnGuardar";
			this.btnGuardar.Size = new System.Drawing.Size(120, 23);
			this.btnGuardar.TabIndex = 373;
			this.btnGuardar.Text = "Cerrar Inventario";
			this.btnGuardar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnGuardar.Click += new System.EventHandler(this.btnGuardar_Click);
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
			// cmbEstados
			// 
			this.cmbEstados.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbEstados.DataSource = this.ultraDataSource2;
			this.cmbEstados.DisplayLayout.AutoFitColumns = true;
			ultraGridBand2.AddButtonCaption = "ImportacionTipo";
			ultraGridColumn12.Header.VisiblePosition = 0;
			ultraGridColumn12.Hidden = true;
			ultraGridColumn12.Width = 139;
			ultraGridColumn13.Header.VisiblePosition = 1;
			ultraGridColumn13.Width = 278;
			ultraGridBand2.Columns.AddRange(new object[] {
																										 ultraGridColumn12,
																										 ultraGridColumn13});
			this.cmbEstados.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			this.cmbEstados.DisplayMember = "Estado";
			this.cmbEstados.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbEstados.Location = new System.Drawing.Point(520, 328);
			this.cmbEstados.Name = "cmbEstados";
			this.cmbEstados.Size = new System.Drawing.Size(280, 112);
			this.cmbEstados.TabIndex = 374;
			this.cmbEstados.ValueMember = "idEstado";
			this.cmbEstados.Visible = false;
			// 
			// ultraDataSource2
			// 
			this.ultraDataSource2.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn12,
																																 ultraDataColumn13});
			// 
			// txtidCierre
			// 
			this.txtidCierre.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.txtidCierre.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtidCierre.Enabled = false;
			this.txtidCierre.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtidCierre.Location = new System.Drawing.Point(480, 8);
			this.txtidCierre.Name = "txtidCierre";
			this.txtidCierre.PromptChar = ' ';
			this.txtidCierre.Size = new System.Drawing.Size(64, 22);
			this.txtidCierre.TabIndex = 375;
			this.txtidCierre.Visible = false;
			// 
			// txtEstado
			// 
			this.txtEstado.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.txtEstado.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtEstado.Enabled = false;
			this.txtEstado.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtEstado.Location = new System.Drawing.Point(552, 8);
			this.txtEstado.Name = "txtEstado";
			this.txtEstado.PromptChar = ' ';
			this.txtEstado.Size = new System.Drawing.Size(64, 22);
			this.txtEstado.TabIndex = 376;
			this.txtEstado.Visible = false;
			// 
			// lblContador
			// 
			this.lblContador.AutoSize = true;
			this.lblContador.Font = new System.Drawing.Font("Tahoma", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblContador.ForeColor = System.Drawing.Color.Firebrick;
			this.lblContador.Location = new System.Drawing.Point(368, 8);
			this.lblContador.Name = "lblContador";
			this.lblContador.Size = new System.Drawing.Size(0, 23);
			this.lblContador.TabIndex = 377;
			this.lblContador.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// btnExcel
			// 
			this.btnExcel.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnExcel.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.btnExcel.Image = ((System.Drawing.Image)(resources.GetObject("btnExcel.Image")));
			this.btnExcel.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnExcel.Location = new System.Drawing.Point(568, 104);
			this.btnExcel.Name = "btnExcel";
			this.btnExcel.Size = new System.Drawing.Size(76, 23);
			this.btnExcel.TabIndex = 379;
			this.btnExcel.Text = "Excel";
			this.btnExcel.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnExcel.Click += new System.EventHandler(this.btnExcel_Click);
			// 
			// btnImprimir
			// 
			this.btnImprimir.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnImprimir.Enabled = false;
			this.btnImprimir.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.btnImprimir.Image = ((System.Drawing.Image)(resources.GetObject("btnImprimir.Image")));
			this.btnImprimir.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnImprimir.Location = new System.Drawing.Point(488, 104);
			this.btnImprimir.Name = "btnImprimir";
			this.btnImprimir.Size = new System.Drawing.Size(76, 23);
			this.btnImprimir.TabIndex = 380;
			this.btnImprimir.Text = "Imprimir";
			this.btnImprimir.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnImprimir.Click += new System.EventHandler(this.btnImprimir_Click);
			// 
			// groupBox1
			// 
			this.groupBox1.Location = new System.Drawing.Point(0, 136);
			this.groupBox1.Name = "groupBox1";
			this.groupBox1.Size = new System.Drawing.Size(960, 8);
			this.groupBox1.TabIndex = 381;
			this.groupBox1.TabStop = false;
			// 
			// lblEstado
			// 
			this.lblEstado.AutoSize = true;
			this.lblEstado.Font = new System.Drawing.Font("Tahoma", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblEstado.ForeColor = System.Drawing.Color.Firebrick;
			this.lblEstado.Location = new System.Drawing.Point(368, 48);
			this.lblEstado.Name = "lblEstado";
			this.lblEstado.Size = new System.Drawing.Size(0, 23);
			this.lblEstado.TabIndex = 382;
			this.lblEstado.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// btnParticipantes
			// 
			this.btnParticipantes.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnParticipantes.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.btnParticipantes.Image = ((System.Drawing.Image)(resources.GetObject("btnParticipantes.Image")));
			this.btnParticipantes.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnParticipantes.Location = new System.Drawing.Point(648, 104);
			this.btnParticipantes.Name = "btnParticipantes";
			this.btnParticipantes.Size = new System.Drawing.Size(88, 23);
			this.btnParticipantes.TabIndex = 383;
			this.btnParticipantes.Text = "Participantes";
			this.btnParticipantes.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnParticipantes.Click += new System.EventHandler(this.btnParticipantes_Click);
			// 
			// lblFecha
			// 
			this.lblFecha.AutoSize = true;
			this.lblFecha.Font = new System.Drawing.Font("Tahoma", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblFecha.ForeColor = System.Drawing.Color.Firebrick;
			this.lblFecha.Location = new System.Drawing.Point(368, 88);
			this.lblFecha.Name = "lblFecha";
			this.lblFecha.Size = new System.Drawing.Size(0, 23);
			this.lblFecha.TabIndex = 384;
			this.lblFecha.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// txtNotas
			// 
			appearance15.BackColorDisabled = System.Drawing.Color.Transparent;
			appearance15.BackColorDisabled2 = System.Drawing.Color.Transparent;
			appearance15.ForeColor = System.Drawing.Color.Black;
			appearance15.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtNotas.Appearance = appearance15;
			this.txtNotas.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtNotas.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNotas.Enabled = false;
			this.txtNotas.Location = new System.Drawing.Point(640, 8);
			this.txtNotas.MaxLength = 1000;
			this.txtNotas.Multiline = true;
			this.txtNotas.Name = "txtNotas";
			this.txtNotas.Size = new System.Drawing.Size(312, 84);
			this.txtNotas.TabIndex = 385;
			// 
			// btnCalificar
			// 
			this.btnCalificar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnCalificar.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.btnCalificar.Image = ((System.Drawing.Image)(resources.GetObject("btnCalificar.Image")));
			this.btnCalificar.ImageAlign = System.Drawing.ContentAlignment.BottomRight;
			this.btnCalificar.Location = new System.Drawing.Point(740, 104);
			this.btnCalificar.Name = "btnCalificar";
			this.btnCalificar.Size = new System.Drawing.Size(88, 23);
			this.btnCalificar.TabIndex = 386;
			this.btnCalificar.Text = "Calificar";
			this.btnCalificar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnCalificar.Click += new System.EventHandler(this.btnCalificar_Click);
			// 
			// frmInformeConteo
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 14);
			this.ClientSize = new System.Drawing.Size(954, 486);
			this.Controls.Add(this.btnCalificar);
			this.Controls.Add(this.txtNotas);
			this.Controls.Add(this.lblFecha);
			this.Controls.Add(this.btnParticipantes);
			this.Controls.Add(this.lblEstado);
			this.Controls.Add(this.groupBox1);
			this.Controls.Add(this.btnImprimir);
			this.Controls.Add(this.btnExcel);
			this.Controls.Add(this.lblContador);
			this.Controls.Add(this.txtEstado);
			this.Controls.Add(this.txtidCierre);
			this.Controls.Add(this.cmbEstados);
			this.Controls.Add(this.btnGuardar);
			this.Controls.Add(this.chkPaso5);
			this.Controls.Add(this.chkPaso4);
			this.Controls.Add(this.chkPaso3);
			this.Controls.Add(this.chkPaso2);
			this.Controls.Add(this.chkPaso1);
			this.Controls.Add(this.uGridInforme);
			this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
			this.KeyPreview = true;
			this.MaximizeBox = false;
			this.MinimizeBox = false;
			this.Name = "frmInformeConteo";
			this.ShowInTaskbar = false;
			this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
			this.Text = "Informe Conteo De Inventario";
			this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.frmInformeConteo_KeyDown);
			this.Load += new System.EventHandler(this.frmInformeConteo_Load);
			((System.ComponentModel.ISupportInitialize)(this.uGridInforme)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbEstados)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtidCierre)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtEstado)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNotas)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);
		}

		bool bLoad = false;

		private void frmInformeConteo_Load(object sender, System.EventArgs e)
		{
			bLoad = true;

			this.lblFecha.Text = dtFecha.ToString("dd/MM/yyyy");

			this.cmbEstados.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Select idEstado, Estado From EstadoAjusteConteo Where idEstado In (2, 3, 4) Order By idEstado");

			string sSQLGuardaCierre = string.Format("Exec GuardaCierreConteoInventario {0}, '{1}'", Bodega, dtFecha.ToString("yyyyMMdd"));
			this.txtidCierre.Value = Funcion.iEscalarSQL(cdsSeteoF, sSQLGuardaCierre, true);

			string sSQLConsultaCierre = string.Format("Exec ConsultaCierreInventario {0}", (int)this.txtidCierre.Value);
			SqlDataReader dr = Funcion.rEscalarSQL(cdsSeteoF, sSQLConsultaCierre, true);
			dr.Read();
			if (dr.HasRows)
			{				
				this.chkPaso1.Checked = dr.GetBoolean(0);
				this.chkPaso2.Checked = dr.GetBoolean(1);
				this.chkPaso3.Checked = dr.GetBoolean(2);
				this.chkPaso4.Checked = dr.GetBoolean(3);
				this.chkPaso5.Checked = dr.GetBoolean(4);
				this.txtEstado.Value = dr.GetInt32(5);
				this.txtNotas.Value = dr.GetString(6);
			}
			dr.Close();

			if ((int)this.txtEstado.Value == 1)
			{
				string sSQL = string.Format("Exec	DiferenciasConteoInventario '{0}', {1}, {2}", dtFecha.ToString("yyyyMMdd"), Bodega, (int)this.txtidCierre.Value);
				FuncionesProcedimientos.dsGeneral(sSQL, this.uGridInforme);
				//this.uGridInforme.DataSource = Funcion.dvProcedimiento(cdsSeteoF, sSQL);	

				this.lblEstado.Text = "ABIERTO";	

				this.txtNotas.Enabled = true;
				this.btnGuardar.Enabled = true;
			}
			else
			{
				string sSQL = string.Format("Exec	ConsultaDiferenciasInventario '{0}', {1}, {2}", dtFecha.ToString("yyyyMMdd"), Bodega, (int)this.txtidCierre.Value);
				FuncionesProcedimientos.dsGeneral(sSQL, this.uGridInforme);
				//this.uGridInforme.DataSource = Funcion.dvProcedimiento(cdsSeteoF, sSQL);		

				this.lblEstado.Text = "CERRADO";

				this.txtNotas.Enabled = false;
				this.btnGuardar.Enabled = false;
			}

			this.lblContador.Text = this.uGridInforme.Rows.Count + " Registros";

			bLoad = false;		
		}

		private void ActualizaPasos(int Paso, bool bValor)
		{
			string sSQL = string.Format("Exec ActualizaCierreConteo {0}, {1}, {2}", Paso, bValor, (int)this.txtidCierre.Value);
			Funcion.EjecutaSQL(cdsSeteoF, sSQL, true);
		}

		private void chkPaso1_CheckedChanged(object sender, System.EventArgs e)
		{
			if (!bLoad) ActualizaPasos(1, this.chkPaso1.Checked);	

			if (this.chkPaso1.Checked) this.chkPaso2.Enabled = true;
			else this.chkPaso2.Enabled = false;
		}

		private void chkPaso2_CheckedChanged(object sender, System.EventArgs e)
		{
			if (!bLoad) ActualizaPasos(2, this.chkPaso2.Checked);

			if (this.chkPaso2.Checked)
			{
				this.chkPaso4.Enabled = false;
				this.chkPaso3.Enabled = true;				
				this.chkPaso1.Enabled = false;
			}
			else if (!this.chkPaso2.Checked)
			{
				this.btnGuardar.Enabled = false;

				this.chkPaso5.Enabled = false;
				this.chkPaso4.Enabled = false;				
				this.chkPaso3.Enabled = false;
				this.chkPaso1.Enabled = true;
			}
		}

		private void chkPaso3_CheckedChanged(object sender, System.EventArgs e)
		{
			if (!bLoad) ActualizaPasos(3, this.chkPaso3.Checked);

			if (this.chkPaso3.Checked)
			{
				this.chkPaso5.Enabled = false;
				this.chkPaso4.Enabled = true;				
				this.chkPaso2.Enabled = false;
				this.chkPaso1.Enabled = false;
			}
			else if (!this.chkPaso3.Checked)
			{
				this.btnGuardar.Enabled = false;
				this.chkPaso5.Enabled = false;
				this.chkPaso4.Enabled = false;
				this.chkPaso2.Enabled = true;
				this.chkPaso1.Enabled = false;
			}
		}

		private void chkPaso4_CheckedChanged(object sender, System.EventArgs e)
		{
			if (!bLoad) ActualizaPasos(4, this.chkPaso4.Checked);

			if (this.chkPaso4.Checked) 
			{
				if ((int)this.txtEstado.Value == 1) this.chkPaso5.Enabled = true;	else this.chkPaso5.Enabled = false;
				this.chkPaso3.Enabled = false;
				this.chkPaso2.Enabled = false;
				this.chkPaso1.Enabled = false;
			}
			else if (!this.chkPaso4.Checked)
			{
				this.btnGuardar.Enabled = false;
				this.chkPaso5.Enabled = false;				
				this.chkPaso3.Enabled = true;
				this.chkPaso2.Enabled = false;
				this.chkPaso1.Enabled = false;
			}
		}

		private void chkPaso5_CheckedChanged(object sender, System.EventArgs e)
		{
			if (!bLoad) ActualizaPasos(5, this.chkPaso5.Checked);

			if (this.chkPaso5.Checked) 
			{
				if ((int)this.txtEstado.Value == 1)	this.btnGuardar.Enabled = true;
				this.btnImprimir.Enabled = true;
				
				this.chkPaso4.Enabled = false;
				this.chkPaso3.Enabled = false;
				this.chkPaso2.Enabled = false;
				this.chkPaso1.Enabled = false;
				
				FuncionesProcedimientos.EstadoGrids(true, this.uGridInforme);
			}
			else if (!this.chkPaso5.Checked)
			{
				this.btnGuardar.Enabled = false;
				this.btnImprimir.Enabled = false;
				
				this.chkPaso4.Enabled = true;
				this.chkPaso3.Enabled = false;
				this.chkPaso2.Enabled = false;
				this.chkPaso1.Enabled = false;

				FuncionesProcedimientos.EstadoGrids(false, this.uGridInforme);				
			}
		}

		private void frmInformeConteo_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Escape) this.DialogResult = DialogResult.No;
		}

		private void btnGuardar_Click(object sender, System.EventArgs e)
		{	
//			if (Funcion.iEscalarSQL(cdsSeteoF, string.Format("Select COUNT(*) From AuditoriaLocales Where Bodega = {0} And CONVERT(Date, Fecha) = '{1}'", Bodega, dtFecha)) == 0)
//			{
//				MessageBox.Show("No ha ingresado la calificacion del estado del almacen", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
//				return;
//			}

			if (DialogResult.Yes == MessageBox.Show("Antes de Cerrar el Inventario verique que haya seguido los pasos correctamente \n\nEste Proceso no puede ser revertido \n\n\n¿Seguro de Cerrar el Inventario?", "Point Technology", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2))	
			{
				this.DialogResult = DialogResult.OK;
			}
		}		

		private void uGridInforme_AfterCellUpdate(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			#region Cantidad Cruce
			if (e.Cell.Column.ToString() == "CantidadJustifica")
			{
				if ((int)e.Cell.Row.Cells["idEstado"].Value == 3)
				{
					if ((int)e.Cell.Row.Cells["CantidadJustifica"].Value > 0)
					{
						e.Cell.Value = 0;
						MessageBox.Show("Seleccione Estado CRUCE antes de poner la Cantidad a Cruzar", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					}
				}

				if ((int)e.Cell.Row.Cells["CantidadJustifica"].Value > Math.Abs((int)e.Cell.Row.Cells["ConteoExistencia"].Value))
				{
					e.Cell.Value = 0;
					MessageBox.Show("La Cantidad a Cruzar no puede ser Mayor al Faltante", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				}
								
				string sSQL = string.Format("Exec ValidaCantidaTotalCruceConteo {0}, {1}, '{2}'", (int)e.Cell.Row.Cells["idArticulo"].Value, Bodega, dtFecha.ToString("yyyyMMdd"));
				int iCantidad = Funcion.iEscalarSQL(cdsSeteoF, sSQL);		
				if ((int)e.Cell.Row.Cells["CantidadJustifica"].Value > iCantidad)
				{
					e.Cell.Value = 0;
					MessageBox.Show(string.Format("El Maximo a Cruzar es {0}", iCantidad), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				}		
		
				string sSQLAC = string.Format("Update LogAjusteCantidades Set CantidadJustifica = {0} Where idLogAjusteCantidades = {1}", (int)e.Cell.Row.Cells["CantidadJustifica"].Value, (int)e.Cell.Row.Cells["idLogAjusteCantidades"].Value);
				Funcion.EjecutaSQL(cdsSeteoF, sSQLAC, true);
			}
			#endregion Cantidad Cruce

			#region Estado
			if (e.Cell.Column.ToString() == "idEstado")
			{
				if ((int)e.Cell.Row.Cells["idEstado"].Value == 2)
				{
					if ((int)e.Cell.Row.Cells["ConteoExistencia"].Value < 0)
					{
						e.Cell.Value = 3;
						MessageBox.Show("No puede Cambiar a Estado SOBRANTE Cuando la DIFERENCIA ES NEGATIVA O SEA FALTANTE", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					}
				}
				
				if ((int)e.Cell.Row.Cells["idEstado"].Value == 3)
				{
					if ((int)e.Cell.Row.Cells["ConteoExistencia"].Value > 0)
					{
						e.Cell.Value = 2;
						MessageBox.Show("No puede Cambiar a Estado FALTANTE Cuando la DIFERENCIA ES POSITIVA O SEA SOBRANTE", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					}
				}

				if ((int)e.Cell.Row.Cells["idEstado"].Value == 4)
				{
					if ((int)e.Cell.Row.Cells["ConteoExistencia"].Value > 0)
					{
						e.Cell.Value = 2;
						MessageBox.Show("Para generar un CRUCE Seleccione un FALTANTE", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					}
				}

				string sSQLAE = string.Format("Update LogAjusteCantidades Set idEstado = {0} Where idLogAjusteCantidades = {1}", (int)e.Cell.Row.Cells["idEstado"].Value, (int)e.Cell.Row.Cells["idLogAjusteCantidades"].Value);
				Funcion.EjecutaSQL(cdsSeteoF, sSQLAE, true);
			}
			#endregion Estado

			#region Notas
			if (e.Cell.Column.ToString() == "Nota")
			{
				string sSQLAN = string.Format("Update LogAjusteCantidades Set Nota = '{0}' Where idLogAjusteCantidades = {1}", e.Cell.Row.Cells["Nota"].Value.ToString(), (int)e.Cell.Row.Cells["idLogAjusteCantidades"].Value);
				Funcion.EjecutaSQL(cdsSeteoF, sSQLAN, true);
			}
			#endregion Notas
		}

		private void uGridInforme_DoubleClick(object sender, System.EventArgs e)
		{
			if (this.chkPaso5.Checked)
			{
				if ((int)this.uGridInforme.ActiveRow.Cells["idEstado"].Value == 2)
				{
					MessageBox.Show("Solo FALTANTES", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				}
				else if ((int)this.uGridInforme.ActiveRow.Cells["idEstado"].Value != 4)
				{
					MessageBox.Show("Seleccione Estado CRUCE", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				}
				else if ((int)this.uGridInforme.ActiveRow.Cells["CantidadJustifica"].Value == 0)
				{
					MessageBox.Show("Ingrese la Cantidad a Ajustar", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				}				
				else if ((int)this.uGridInforme.ActiveRow.Cells["idEstado"].Value == 4)
				{
					if (Funcion.iEscalarSQL(cdsSeteoF, string.Format("Exec ValidaCierreConteoCruce {0}, {1}, '{2}'", (int)this.uGridInforme.ActiveRow.Cells["idArticulo"].Value, Bodega, dtFecha.ToString("yyyyMMdd"))) == 0)
					{
						MessageBox.Show("No existen Coincidencias Para Cruzar", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					}
					else
					{
						using (frmCruceConteoInventario Cruce = new frmCruceConteoInventario((int)this.uGridInforme.ActiveRow.Cells["idArticulo"].Value, (int)this.uGridInforme.ActiveRow.Cells["idArticulo"].Value, Bodega, dtFecha, this.uGridInforme.ActiveRow.Cells["Codigo"].Value.ToString(), (int)this.uGridInforme.ActiveRow.Cells["ConteoExistencia"].Value, (int)this.uGridInforme.ActiveRow.Cells["CantidadJustifica"].Value))
						{
							if (DialogResult.OK == Cruce.ShowDialog())
							{
								string sSQL = string.Format("Exec	ConsultaDiferenciasInventario '{0}', {1}, {2}", dtFecha.ToString("yyyyMMdd"), Bodega, (int)this.txtidCierre.Value);
								this.uGridInforme.DataSource = Funcion.dvProcedimiento(cdsSeteoF, sSQL);	
							}
						}
					}
				}
			}
		}

		private void uGridInforme_BeforeCellUpdate(object sender, Infragistics.Win.UltraWinGrid.BeforeCellUpdateEventArgs e)
		{
			if (e.Cell.Column.ToString() == "idEstado")
			{
				if (Funcion.iEscalarSQL(cdsSeteoF, string.Format("Exec ValidaCierreConteoCruce {0}, {1}, '{2}'", (int)this.uGridInforme.ActiveRow.Cells["idArticulo"].Value, Bodega, dtFecha.ToString("yyyyMMdd"))) == 0)
				{
					MessageBox.Show("No existen Coincidencias Para Cruzar", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					e.Cancel = true;
				}
			}

			if (e.Cell.Column.ToString() == "CantidadJustifica")	
			{
				if ((int)e.Cell.Row.Cells["idEstado"].Value == 2)
				{
					MessageBox.Show("Solo FALTANTES", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					e.Cancel = true;
				}
				else if (Funcion.iEscalarSQL(cdsSeteoF, string.Format("Exec ValidaCierreConteoCruce {0}, {1}, '{2}'", (int)this.uGridInforme.ActiveRow.Cells["idArticulo"].Value, Bodega, dtFecha.ToString("yyyyMMdd"))) == 0)
				{
					MessageBox.Show("No existen Coincidencias Para Cruzar", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					e.Cancel = true;
				}
			}
		}

		private void uGridInforme_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			FuncionesProcedimientos.DesplazamientoGrids(sender, e, this.uGridInforme);
		}

		private void btnExcel_Click(object sender, System.EventArgs e)
		{
			FuncionesProcedimientos.ExportaExcel(this.uGridInforme);			
		}

		private void btnImprimir_Click(object sender, System.EventArgs e)
		{
			ParameterFields paramFields = new ParameterFields ();
				
			ParameterField pfCapital = new ParameterField ();
			ParameterDiscreteValue ValCapital = new ParameterDiscreteValue ();
			pfCapital.ParameterFieldName = "@Fecha";
			ValCapital.Value = dtFecha;
			pfCapital.CurrentValues.Add (ValCapital);
			paramFields.Add (pfCapital);
				
			ParameterField pfCuotas = new ParameterField ();
			ParameterDiscreteValue ValCuotas = new ParameterDiscreteValue ();
			pfCuotas.ParameterFieldName = "@Bodega";
			ValCuotas.Value = Bodega;
			pfCuotas.CurrentValues.Add (ValCuotas);
			paramFields.Add (pfCuotas);
				
			ParameterField pfIdCliente = new ParameterField ();
			ParameterDiscreteValue ValIdCliente = new ParameterDiscreteValue ();
			pfIdCliente.ParameterFieldName = "@idCierreConteoInventario";
			ValIdCliente.Value = (int)this.txtidCierre.Value;
			pfIdCliente.CurrentValues.Add (ValIdCliente);
			paramFields.Add (pfIdCliente);
			
			Reporte miRepor = new Reporte("ResumenDiferenciasInventario.rpt", "");
			miRepor.MdiParent = this.MdiParent;			
			miRepor.crVista.ParameterFieldInfo = paramFields;
			miRepor.ShowDialog();	
		}

		private void btnParticipantes_Click(object sender, System.EventArgs e)
		{
			using (frmConteoPersonal Personal = new frmConteoPersonal((int)this.txtidCierre.Value, 1,dtFecha, Bodega,(int)this.txtEstado.Value))
			{				
				if (DialogResult.OK == Personal.ShowDialog())
				{
					
				}
			}
		}

		private void btnCalificar_Click(object sender, System.EventArgs e)
		{			
			using (frmReporteAuditoriaLocal1 AuditoriaL = new frmReporteAuditoriaLocal1(Bodega, dtFecha, (int)this.txtEstado.Value))
			{
				if (DialogResult.OK == AuditoriaL.ShowDialog())
				{
					
				}
			}
		}

		private void uGridInforme_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}
	}
}

