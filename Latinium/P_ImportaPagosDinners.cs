using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using DevExpress.XtraEditors;

using System.IO;
using System.Globalization;
using System.Data.SqlClient;
using System.Data;
using C1.C1Excel;


using CrystalDecisions.Shared;
using System.Data.OleDb;

namespace Latinium
{
	/// <summary>
	/// Summary description for P_ImportaPagosDinners.
	/// </summary>
	public class P_ImportaPagosDinners : DevExpress.XtraEditors.XtraForm
	{
		private Infragistics.Win.Misc.UltraButton btLeer;
		private Infragistics.Win.Misc.UltraButton btSubirR;
		private Infragistics.Win.Misc.UltraButton btArchivo;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtPago;
		private C1.C1Excel.C1XLBook c1XLBook;
		private System.Windows.Forms.OpenFileDialog openFileDialog;
		private System.Windows.Forms.Label label1;
		private System.Windows.Forms.Label label2;
		private System.Windows.Forms.Label label3;
		private Infragistics.Win.Misc.UltraButton btAgregar;
		private System.Windows.Forms.ImageList imageList;
		private Infragistics.Win.UltraWinGrid.UltraGrid ultraGrid2;
		private Infragistics.Win.UltraWinGrid.UltraGrid ultraGrid1;
		private Infragistics.Win.UltraWinGrid.UltraGrid ultraGrid3;
		private C1.Data.C1DataSet cdsCompra;
		private System.Windows.Forms.ErrorProvider errorProvider;
		private C1.Data.C1DataSet cdsPagoLote;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource udsTmpArt;
		private Infragistics.Win.Misc.UltraButton btLeerE;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbBanco;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbCuentaCorr;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbFormaPago;
		private DevExpress.XtraEditors.SplitContainerControl splitContainerControl1;
		private System.Windows.Forms.Label label10;
		private System.Windows.Forms.Label label4;
		private System.Windows.Forms.Label label5;
		private System.Windows.Forms.Label label6;
		private C1.Data.C1DataSet cdsCuentaCorr;
		private C1.Data.C1DataSet cdsFormaPago;
		private System.ComponentModel.IContainer components;

		public P_ImportaPagosDinners()
		{
			//
			// Required for Windows Form Designer support
			//
			InitializeComponent();

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
			this.components = new System.ComponentModel.Container();
			Infragistics.Win.Appearance appearance1 = new Infragistics.Win.Appearance();
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(P_ImportaPagosDinners));
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCompra");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Total");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Numero");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cheque");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Concepto");
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCompra");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Total");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn3 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Numero");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn4 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Cheque");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn5 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Concepto");
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance11 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance12 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance13 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance14 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance15 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance16 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance17 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Banco", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idBanco");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Pais");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Borrar");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn11 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Hoja");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn12 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ProcesoOk");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand3 = new Infragistics.Win.UltraWinGrid.UltraGridBand("CuentaCorriente", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn13 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCuentaCorriente");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn14 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idSucursal");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn15 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idBanco");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn16 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCuenta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn17 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("NumCuenta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn18 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("NumCheque");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn19 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Agencia");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn20 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Reporte");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn21 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Automatico");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn22 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Borrar");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn23 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Asiento");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn24 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCuentaTipo");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand4 = new Infragistics.Win.UltraWinGrid.UltraGridBand("CompraFormaPago", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn25 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idFormaPago");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn26 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPrecio");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn27 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FormaPago");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn28 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCuenta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn29 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Borrar");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn30 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Descuento");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn31 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Ruc");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn32 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idTipo");
			this.btLeer = new Infragistics.Win.Misc.UltraButton();
			this.btSubirR = new Infragistics.Win.Misc.UltraButton();
			this.btArchivo = new Infragistics.Win.Misc.UltraButton();
			this.txtPago = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.c1XLBook = new C1.C1Excel.C1XLBook();
			this.openFileDialog = new System.Windows.Forms.OpenFileDialog();
			this.label1 = new System.Windows.Forms.Label();
			this.label2 = new System.Windows.Forms.Label();
			this.label3 = new System.Windows.Forms.Label();
			this.btAgregar = new Infragistics.Win.Misc.UltraButton();
			this.imageList = new System.Windows.Forms.ImageList(this.components);
			this.ultraGrid2 = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.udsTmpArt = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.ultraGrid1 = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraGrid3 = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.cdsCompra = new C1.Data.C1DataSet();
			this.errorProvider = new System.Windows.Forms.ErrorProvider();
			this.btLeerE = new Infragistics.Win.Misc.UltraButton();
			this.cdsPagoLote = new C1.Data.C1DataSet();
			this.cmbBanco = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.cdsCuentaCorr = new C1.Data.C1DataSet();
			this.cmbCuentaCorr = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.cmbFormaPago = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.cdsFormaPago = new C1.Data.C1DataSet();
			this.splitContainerControl1 = new DevExpress.XtraEditors.SplitContainerControl();
			this.label10 = new System.Windows.Forms.Label();
			this.label4 = new System.Windows.Forms.Label();
			this.label5 = new System.Windows.Forms.Label();
			this.label6 = new System.Windows.Forms.Label();
			((System.ComponentModel.ISupportInitialize)(this.txtPago)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid2)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.udsTmpArt)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid3)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsCompra)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsPagoLote)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbBanco)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsCuentaCorr)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCuentaCorr)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbFormaPago)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsFormaPago)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.splitContainerControl1)).BeginInit();
			this.splitContainerControl1.SuspendLayout();
			this.SuspendLayout();
			// 
			// btLeer
			// 
			this.btLeer.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.btLeer.Location = new System.Drawing.Point(192, 8);
			this.btLeer.Name = "btLeer";
			this.btLeer.Size = new System.Drawing.Size(96, 24);
			this.btLeer.TabIndex = 33;
			this.btLeer.Text = "Leer Archivo";
			this.btLeer.Click += new System.EventHandler(this.btLeer_Click);
			// 
			// btSubirR
			// 
			this.btSubirR.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.btSubirR.Location = new System.Drawing.Point(168, 16);
			this.btSubirR.Name = "btSubirR";
			this.btSubirR.Size = new System.Drawing.Size(120, 24);
			this.btSubirR.TabIndex = 32;
			this.btSubirR.Text = "Generar Pagos";
			this.btSubirR.Click += new System.EventHandler(this.btSubirR_Click);
			// 
			// btArchivo
			// 
			this.btArchivo.Location = new System.Drawing.Point(160, 8);
			this.btArchivo.Name = "btArchivo";
			this.btArchivo.Size = new System.Drawing.Size(24, 26);
			this.btArchivo.TabIndex = 35;
			this.btArchivo.Text = "...";
			this.btArchivo.Click += new System.EventHandler(this.btArchivo_Click);
			// 
			// txtPago
			// 
			this.txtPago.Location = new System.Drawing.Point(16, 8);
			this.txtPago.Name = "txtPago";
			this.txtPago.Size = new System.Drawing.Size(144, 22);
			this.txtPago.TabIndex = 34;
			this.txtPago.Text = "c:\\Latinium\\";
			// 
			// label1
			// 
			this.label1.Location = new System.Drawing.Point(8, 48);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(176, 23);
			this.label1.TabIndex = 38;
			this.label1.Text = "Facturas de Archivo";
			// 
			// label2
			// 
			this.label2.Location = new System.Drawing.Point(8, 272);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(232, 23);
			this.label2.TabIndex = 39;
			this.label2.Text = "Facturas Anteriores No constan en Archivo";
			// 
			// label3
			// 
			this.label3.Location = new System.Drawing.Point(-136, 16);
			this.label3.Name = "label3";
			this.label3.Size = new System.Drawing.Size(136, 23);
			this.label3.TabIndex = 40;
			this.label3.Text = "Facturas  a Realizar Pago";
			// 
			// btAgregar
			// 
			appearance1.Image = 2;
			this.btAgregar.Appearance = appearance1;
			this.btAgregar.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.btAgregar.ImageList = this.imageList;
			this.btAgregar.Location = new System.Drawing.Point(0, 312);
			this.btAgregar.Name = "btAgregar";
			this.btAgregar.Size = new System.Drawing.Size(24, 24);
			this.btAgregar.TabIndex = 41;
			this.btAgregar.Click += new System.EventHandler(this.btAgregar_Click);
			// 
			// imageList
			// 
			this.imageList.ColorDepth = System.Windows.Forms.ColorDepth.Depth16Bit;
			this.imageList.ImageSize = new System.Drawing.Size(18, 18);
			this.imageList.ImageStream = ((System.Windows.Forms.ImageListStreamer)(resources.GetObject("imageList.ImageStream")));
			this.imageList.TransparentColor = System.Drawing.Color.Transparent;
			// 
			// ultraGrid2
			// 
			this.ultraGrid2.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.ultraGrid2.Cursor = System.Windows.Forms.Cursors.Default;
			this.ultraGrid2.DataMember = "Band 0";
			this.ultraGrid2.DataSource = this.udsTmpArt;
			appearance2.BackColor = System.Drawing.Color.White;
			this.ultraGrid2.DisplayLayout.Appearance = appearance2;
			this.ultraGrid2.DisplayLayout.AutoFitColumns = true;
			ultraGridBand1.AddButtonCaption = "PagoLote";
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn2.Header.Caption = "Pago";
			ultraGridColumn2.Header.VisiblePosition = 2;
			ultraGridColumn2.Width = 80;
			ultraGridColumn3.Header.Caption = "Factura";
			ultraGridColumn3.Header.VisiblePosition = 1;
			ultraGridColumn3.Width = 74;
			ultraGridColumn4.Header.VisiblePosition = 3;
			ultraGridColumn4.Width = 84;
			ultraGridColumn5.Header.VisiblePosition = 4;
			ultraGridColumn5.Width = 161;
			ultraGridBand1.Columns.AddRange(new object[] {
																										 ultraGridColumn1,
																										 ultraGridColumn2,
																										 ultraGridColumn3,
																										 ultraGridColumn4,
																										 ultraGridColumn5});
			this.ultraGrid2.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.ultraGrid2.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.True;
			this.ultraGrid2.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance3.BackColor = System.Drawing.Color.Transparent;
			this.ultraGrid2.DisplayLayout.Override.CardAreaAppearance = appearance3;
			appearance4.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance4.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance4.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance4.FontData.BoldAsString = "True";
			appearance4.FontData.Name = "Arial";
			appearance4.FontData.SizeInPoints = 10F;
			appearance4.ForeColor = System.Drawing.Color.White;
			appearance4.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.ultraGrid2.DisplayLayout.Override.HeaderAppearance = appearance4;
			this.ultraGrid2.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortSingle;
			appearance5.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance5.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance5.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid2.DisplayLayout.Override.RowSelectorAppearance = appearance5;
			appearance6.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance6.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance6.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid2.DisplayLayout.Override.SelectedRowAppearance = appearance6;
			this.ultraGrid2.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.ultraGrid2.Location = new System.Drawing.Point(32, 136);
			this.ultraGrid2.Name = "ultraGrid2";
			this.ultraGrid2.Size = new System.Drawing.Size(420, 256);
			this.ultraGrid2.TabIndex = 42;
			// 
			// udsTmpArt
			// 
			ultraDataColumn1.DataType = typeof(int);
			ultraDataColumn1.DefaultValue = 0;
			ultraDataColumn2.DataType = typeof(System.Double);
			ultraDataColumn2.DefaultValue = 0;
			this.udsTmpArt.Band.Columns.AddRange(new object[] {
																													ultraDataColumn1,
																													ultraDataColumn2,
																													ultraDataColumn3,
																													ultraDataColumn4,
																													ultraDataColumn5});
			// 
			// ultraGrid1
			// 
			this.ultraGrid1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.ultraGrid1.Cursor = System.Windows.Forms.Cursors.Default;
			appearance7.BackColor = System.Drawing.Color.White;
			this.ultraGrid1.DisplayLayout.Appearance = appearance7;
			this.ultraGrid1.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
			this.ultraGrid1.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.ultraGrid1.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.False;
			appearance8.BackColor = System.Drawing.Color.Transparent;
			this.ultraGrid1.DisplayLayout.Override.CardAreaAppearance = appearance8;
			this.ultraGrid1.DisplayLayout.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.RowSelect;
			this.ultraGrid1.DisplayLayout.Override.ColumnAutoSizeMode = Infragistics.Win.UltraWinGrid.ColumnAutoSizeMode.AllRowsInBand;
			appearance9.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance9.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance9.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance9.FontData.BoldAsString = "True";
			appearance9.FontData.Name = "Arial";
			appearance9.FontData.SizeInPoints = 10F;
			appearance9.ForeColor = System.Drawing.Color.White;
			appearance9.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.ultraGrid1.DisplayLayout.Override.HeaderAppearance = appearance9;
			this.ultraGrid1.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti;
			appearance10.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance10.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance10.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.RowSelectorAppearance = appearance10;
			appearance11.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance11.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance11.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.SelectedRowAppearance = appearance11;
			this.ultraGrid1.DisplayLayout.Override.SelectTypeCell = Infragistics.Win.UltraWinGrid.SelectType.None;
			this.ultraGrid1.DisplayLayout.Override.SelectTypeCol = Infragistics.Win.UltraWinGrid.SelectType.None;
			this.ultraGrid1.DisplayLayout.Override.SelectTypeRow = Infragistics.Win.UltraWinGrid.SelectType.Extended;
			this.ultraGrid1.DisplayLayout.ScrollBounds = Infragistics.Win.UltraWinGrid.ScrollBounds.ScrollToFill;
			this.ultraGrid1.DisplayLayout.ScrollStyle = Infragistics.Win.UltraWinGrid.ScrollStyle.Immediate;
			this.ultraGrid1.DisplayLayout.TabNavigation = Infragistics.Win.UltraWinGrid.TabNavigation.NextControl;
			this.ultraGrid1.DisplayLayout.ViewStyle = Infragistics.Win.UltraWinGrid.ViewStyle.SingleBand;
			this.ultraGrid1.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.ultraGrid1.Location = new System.Drawing.Point(8, 64);
			this.ultraGrid1.Name = "ultraGrid1";
			this.ultraGrid1.Size = new System.Drawing.Size(389, 192);
			this.ultraGrid1.TabIndex = 43;
			// 
			// ultraGrid3
			// 
			this.ultraGrid3.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.ultraGrid3.Cursor = System.Windows.Forms.Cursors.Default;
			appearance12.BackColor = System.Drawing.Color.White;
			this.ultraGrid3.DisplayLayout.Appearance = appearance12;
			this.ultraGrid3.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
			this.ultraGrid3.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.ultraGrid3.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.False;
			appearance13.BackColor = System.Drawing.Color.Transparent;
			this.ultraGrid3.DisplayLayout.Override.CardAreaAppearance = appearance13;
			this.ultraGrid3.DisplayLayout.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.RowSelect;
			this.ultraGrid3.DisplayLayout.Override.ColumnAutoSizeMode = Infragistics.Win.UltraWinGrid.ColumnAutoSizeMode.AllRowsInBand;
			appearance14.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance14.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance14.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance14.FontData.BoldAsString = "True";
			appearance14.FontData.Name = "Arial";
			appearance14.FontData.SizeInPoints = 10F;
			appearance14.ForeColor = System.Drawing.Color.White;
			appearance14.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.ultraGrid3.DisplayLayout.Override.HeaderAppearance = appearance14;
			this.ultraGrid3.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti;
			appearance15.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance15.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance15.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid3.DisplayLayout.Override.RowSelectorAppearance = appearance15;
			appearance16.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance16.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance16.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid3.DisplayLayout.Override.SelectedRowAppearance = appearance16;
			this.ultraGrid3.DisplayLayout.Override.SelectTypeCell = Infragistics.Win.UltraWinGrid.SelectType.None;
			this.ultraGrid3.DisplayLayout.Override.SelectTypeCol = Infragistics.Win.UltraWinGrid.SelectType.None;
			this.ultraGrid3.DisplayLayout.Override.SelectTypeRow = Infragistics.Win.UltraWinGrid.SelectType.Extended;
			this.ultraGrid3.DisplayLayout.ScrollBounds = Infragistics.Win.UltraWinGrid.ScrollBounds.ScrollToFill;
			this.ultraGrid3.DisplayLayout.ScrollStyle = Infragistics.Win.UltraWinGrid.ScrollStyle.Immediate;
			this.ultraGrid3.DisplayLayout.TabNavigation = Infragistics.Win.UltraWinGrid.TabNavigation.NextControl;
			this.ultraGrid3.DisplayLayout.ViewStyle = Infragistics.Win.UltraWinGrid.ViewStyle.SingleBand;
			this.ultraGrid3.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.ultraGrid3.Location = new System.Drawing.Point(8, 288);
			this.ultraGrid3.Name = "ultraGrid3";
			this.ultraGrid3.Size = new System.Drawing.Size(389, 116);
			this.ultraGrid3.TabIndex = 44;
			// 
			// cdsCompra
			// 
			this.cdsCompra.BindingContextCtrl = this;
			this.cdsCompra.DataLibrary = "LibFacturacion";
			this.cdsCompra.DataLibraryUrl = "";
			this.cdsCompra.DataSetDef = "dsCompra";
			this.cdsCompra.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsCompra.Name = "cdsCompra";
			this.cdsCompra.SchemaClassName = "LibFacturacion.DataClass";
			this.cdsCompra.SchemaDef = null;
			this.cdsCompra.BeforeFill += new C1.Data.FillEventHandler(this.cdsCompra_BeforeFill);
			// 
			// errorProvider
			// 
			this.errorProvider.ContainerControl = this;
			// 
			// btLeerE
			// 
			appearance17.Image = 2;
			this.btLeerE.Appearance = appearance17;
			this.btLeerE.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.btLeerE.ImageList = this.imageList;
			this.btLeerE.Location = new System.Drawing.Point(0, 144);
			this.btLeerE.Name = "btLeerE";
			this.btLeerE.Size = new System.Drawing.Size(24, 24);
			this.btLeerE.TabIndex = 45;
			this.btLeerE.Click += new System.EventHandler(this.btLeerE_Click);
			// 
			// cdsPagoLote
			// 
			this.cdsPagoLote.BindingContextCtrl = this;
			this.cdsPagoLote.DataLibrary = "LibFacturacion";
			this.cdsPagoLote.DataLibraryUrl = "";
			this.cdsPagoLote.DataSetDef = "dsPagoLote";
			this.cdsPagoLote.FillOnRequest = false;
			this.cdsPagoLote.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsPagoLote.Name = "cdsPagoLote";
			this.cdsPagoLote.SchemaClassName = "LibFacturacion.DataClass";
			this.cdsPagoLote.SchemaDef = null;
			this.cdsPagoLote.BeforeFill += new C1.Data.FillEventHandler(this.cdsPagoLote_BeforeFill);
			// 
			// cmbBanco
			// 
			this.cmbBanco.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbBanco.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbBanco.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsPagoLote, "_PagoLote.idBanco"));
			this.cmbBanco.DataMember = "Banco";
			this.cmbBanco.DataSource = this.cdsCuentaCorr;
			this.cmbBanco.DisplayLayout.AutoFitColumns = true;
			ultraGridColumn6.Header.VisiblePosition = 0;
			ultraGridColumn6.Hidden = true;
			ultraGridColumn6.Width = 15;
			ultraGridColumn7.Header.VisiblePosition = 1;
			ultraGridColumn7.Width = 157;
			ultraGridColumn8.Header.VisiblePosition = 2;
			ultraGridColumn8.Hidden = true;
			ultraGridColumn8.Width = 25;
			ultraGridColumn9.Header.VisiblePosition = 3;
			ultraGridColumn9.Hidden = true;
			ultraGridColumn9.Width = 25;
			ultraGridColumn10.Header.VisiblePosition = 4;
			ultraGridColumn10.Hidden = true;
			ultraGridColumn10.Width = 17;
			ultraGridColumn11.Header.VisiblePosition = 5;
			ultraGridColumn11.Hidden = true;
			ultraGridColumn11.Width = 25;
			ultraGridColumn12.Header.VisiblePosition = 6;
			ultraGridColumn12.Hidden = true;
			ultraGridColumn12.Width = 25;
			ultraGridBand2.Columns.AddRange(new object[] {
																										 ultraGridColumn6,
																										 ultraGridColumn7,
																										 ultraGridColumn8,
																										 ultraGridColumn9,
																										 ultraGridColumn10,
																										 ultraGridColumn11,
																										 ultraGridColumn12});
			this.cmbBanco.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			this.cmbBanco.DisplayMember = "Nombre";
			this.cmbBanco.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbBanco.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbBanco.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbBanco.Location = new System.Drawing.Point(112, 104);
			this.cmbBanco.Name = "cmbBanco";
			this.cmbBanco.Size = new System.Drawing.Size(176, 21);
			this.cmbBanco.TabIndex = 48;
			this.cmbBanco.ValueMember = "idBanco";
			// 
			// cdsCuentaCorr
			// 
			this.cdsCuentaCorr.BindingContextCtrl = this;
			this.cdsCuentaCorr.DataLibrary = "LibContabilidad";
			this.cdsCuentaCorr.DataLibraryUrl = "";
			this.cdsCuentaCorr.DataSetDef = "dsCuentaCorriente";
			this.cdsCuentaCorr.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsCuentaCorr.Name = "cdsCuentaCorr";
			this.cdsCuentaCorr.SchemaClassName = "LibContabilidad.DataClass";
			this.cdsCuentaCorr.SchemaDef = null;
			this.cdsCuentaCorr.BeforeFill += new C1.Data.FillEventHandler(this.cdsCuentaCorr_BeforeFill);
			// 
			// cmbCuentaCorr
			// 
			this.cmbCuentaCorr.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbCuentaCorr.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbCuentaCorr.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsPagoLote, "_PagoLote.idCuentaCorriente"));
			this.cmbCuentaCorr.DataMember = "CuentaCorriente";
			this.cmbCuentaCorr.DataSource = this.cdsCuentaCorr;
			ultraGridColumn13.Header.VisiblePosition = 0;
			ultraGridColumn13.Hidden = true;
			ultraGridColumn14.Header.VisiblePosition = 1;
			ultraGridColumn14.Hidden = true;
			ultraGridColumn15.Header.VisiblePosition = 2;
			ultraGridColumn15.Hidden = true;
			ultraGridColumn16.Header.VisiblePosition = 3;
			ultraGridColumn16.Hidden = true;
			ultraGridColumn17.Header.VisiblePosition = 4;
			ultraGridColumn18.Header.VisiblePosition = 5;
			ultraGridColumn18.Hidden = true;
			ultraGridColumn19.Header.VisiblePosition = 6;
			ultraGridColumn19.Hidden = true;
			ultraGridColumn20.Header.VisiblePosition = 7;
			ultraGridColumn20.Hidden = true;
			ultraGridColumn21.Header.VisiblePosition = 8;
			ultraGridColumn21.Hidden = true;
			ultraGridColumn22.Header.VisiblePosition = 9;
			ultraGridColumn22.Hidden = true;
			ultraGridColumn23.Header.VisiblePosition = 10;
			ultraGridColumn23.Hidden = true;
			ultraGridColumn24.Header.VisiblePosition = 11;
			ultraGridColumn24.Hidden = true;
			ultraGridBand3.Columns.AddRange(new object[] {
																										 ultraGridColumn13,
																										 ultraGridColumn14,
																										 ultraGridColumn15,
																										 ultraGridColumn16,
																										 ultraGridColumn17,
																										 ultraGridColumn18,
																										 ultraGridColumn19,
																										 ultraGridColumn20,
																										 ultraGridColumn21,
																										 ultraGridColumn22,
																										 ultraGridColumn23,
																										 ultraGridColumn24});
			this.cmbCuentaCorr.DisplayLayout.BandsSerializer.Add(ultraGridBand3);
			this.cmbCuentaCorr.DisplayMember = "NumCuenta";
			this.cmbCuentaCorr.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbCuentaCorr.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbCuentaCorr.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbCuentaCorr.Location = new System.Drawing.Point(112, 56);
			this.cmbCuentaCorr.Name = "cmbCuentaCorr";
			this.cmbCuentaCorr.Size = new System.Drawing.Size(176, 22);
			this.cmbCuentaCorr.TabIndex = 47;
			this.cmbCuentaCorr.ValueMember = "idCuentaCorriente";
			// 
			// cmbFormaPago
			// 
			this.cmbFormaPago.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbFormaPago.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbFormaPago.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsPagoLote, "_PagoLote.idFormaPago"));
			this.cmbFormaPago.DataMember = "CompraFormaPago";
			this.cmbFormaPago.DataSource = this.cdsFormaPago;
			this.cmbFormaPago.DisplayLayout.AutoFitColumns = true;
			ultraGridColumn25.Header.VisiblePosition = 0;
			ultraGridColumn25.Hidden = true;
			ultraGridColumn25.Width = 21;
			ultraGridColumn26.Header.VisiblePosition = 1;
			ultraGridColumn26.Hidden = true;
			ultraGridColumn26.Width = 15;
			ultraGridColumn27.Header.VisiblePosition = 2;
			ultraGridColumn27.Width = 157;
			ultraGridColumn28.Header.VisiblePosition = 3;
			ultraGridColumn28.Hidden = true;
			ultraGridColumn28.Width = 16;
			ultraGridColumn29.Header.VisiblePosition = 4;
			ultraGridColumn29.Hidden = true;
			ultraGridColumn29.Width = 17;
			ultraGridColumn30.Header.VisiblePosition = 5;
			ultraGridColumn30.Hidden = true;
			ultraGridColumn30.Width = 21;
			ultraGridColumn31.Header.VisiblePosition = 6;
			ultraGridColumn31.Hidden = true;
			ultraGridColumn31.Width = 26;
			ultraGridColumn32.Header.VisiblePosition = 7;
			ultraGridColumn32.Hidden = true;
			ultraGridColumn32.Width = 15;
			ultraGridBand4.Columns.AddRange(new object[] {
																										 ultraGridColumn25,
																										 ultraGridColumn26,
																										 ultraGridColumn27,
																										 ultraGridColumn28,
																										 ultraGridColumn29,
																										 ultraGridColumn30,
																										 ultraGridColumn31,
																										 ultraGridColumn32});
			this.cmbFormaPago.DisplayLayout.BandsSerializer.Add(ultraGridBand4);
			this.cmbFormaPago.DisplayMember = "FormaPago";
			this.cmbFormaPago.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbFormaPago.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbFormaPago.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbFormaPago.Location = new System.Drawing.Point(112, 80);
			this.cmbFormaPago.Name = "cmbFormaPago";
			this.cmbFormaPago.Size = new System.Drawing.Size(176, 21);
			this.cmbFormaPago.TabIndex = 46;
			this.cmbFormaPago.ValueMember = "idFormaPago";
			// 
			// cdsFormaPago
			// 
			this.cdsFormaPago.BindingContextCtrl = this;
			this.cdsFormaPago.DataLibrary = "LibFacturacion";
			this.cdsFormaPago.DataLibraryUrl = "";
			this.cdsFormaPago.DataSetDef = "dsCompraTabla";
			this.cdsFormaPago.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsFormaPago.Name = "cdsFormaPago";
			this.cdsFormaPago.SchemaClassName = "LibFacturacion.DataClass";
			this.cdsFormaPago.SchemaDef = null;
			this.cdsFormaPago.BeforeFill += new C1.Data.FillEventHandler(this.cdsFormaPago_BeforeFill);
			// 
			// splitContainerControl1
			// 
			this.splitContainerControl1.Dock = System.Windows.Forms.DockStyle.Fill;
			this.splitContainerControl1.Location = new System.Drawing.Point(0, 0);
			this.splitContainerControl1.Name = "splitContainerControl1";
			this.splitContainerControl1.Panel1.Controls.Add(this.ultraGrid3);
			this.splitContainerControl1.Panel1.Controls.Add(this.ultraGrid1);
			this.splitContainerControl1.Panel1.Controls.Add(this.label2);
			this.splitContainerControl1.Panel1.Controls.Add(this.label1);
			this.splitContainerControl1.Panel1.Controls.Add(this.txtPago);
			this.splitContainerControl1.Panel1.Controls.Add(this.btLeer);
			this.splitContainerControl1.Panel1.Controls.Add(this.btArchivo);
			this.splitContainerControl1.Panel1.ShowCaption = true;
			this.splitContainerControl1.Panel1.Text = "Cargar Datos";
			this.splitContainerControl1.Panel2.Controls.Add(this.label10);
			this.splitContainerControl1.Panel2.Controls.Add(this.label4);
			this.splitContainerControl1.Panel2.Controls.Add(this.label5);
			this.splitContainerControl1.Panel2.Controls.Add(this.btLeerE);
			this.splitContainerControl1.Panel2.Controls.Add(this.btAgregar);
			this.splitContainerControl1.Panel2.Controls.Add(this.cmbCuentaCorr);
			this.splitContainerControl1.Panel2.Controls.Add(this.cmbFormaPago);
			this.splitContainerControl1.Panel2.Controls.Add(this.ultraGrid2);
			this.splitContainerControl1.Panel2.Controls.Add(this.label3);
			this.splitContainerControl1.Panel2.Controls.Add(this.btSubirR);
			this.splitContainerControl1.Panel2.Controls.Add(this.cmbBanco);
			this.splitContainerControl1.Panel2.Controls.Add(this.label6);
			this.splitContainerControl1.Panel2.ShowCaption = true;
			this.splitContainerControl1.Panel2.Text = "Facturas a Generar Pagos";
			this.splitContainerControl1.Size = new System.Drawing.Size(880, 422);
			this.splitContainerControl1.SplitterPosition = 405;
			this.splitContainerControl1.TabIndex = 52;
			this.splitContainerControl1.Text = "splitContainerControl1";
			// 
			// label10
			// 
			this.label10.AutoSize = true;
			this.label10.Location = new System.Drawing.Point(32, 104);
			this.label10.Name = "label10";
			this.label10.Size = new System.Drawing.Size(34, 16);
			this.label10.TabIndex = 54;
			this.label10.Text = "Banco";
			// 
			// label4
			// 
			this.label4.AutoSize = true;
			this.label4.Location = new System.Drawing.Point(32, 56);
			this.label4.Name = "label4";
			this.label4.Size = new System.Drawing.Size(67, 16);
			this.label4.TabIndex = 53;
			this.label4.Text = "Cuenta Corr.";
			// 
			// label5
			// 
			this.label5.AutoSize = true;
			this.label5.Location = new System.Drawing.Point(32, 80);
			this.label5.Name = "label5";
			this.label5.Size = new System.Drawing.Size(80, 16);
			this.label5.TabIndex = 52;
			this.label5.Text = "Forma de Pago";
			// 
			// label6
			// 
			this.label6.Location = new System.Drawing.Point(32, 16);
			this.label6.Name = "label6";
			this.label6.Size = new System.Drawing.Size(104, 23);
			this.label6.TabIndex = 45;
			this.label6.Text = "Facturas a Pagar";
			// 
			// P_ImportaPagosDinners
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 14);
			this.ClientSize = new System.Drawing.Size(880, 422);
			this.Controls.Add(this.splitContainerControl1);
			this.Name = "P_ImportaPagosDinners";
			this.Text = "P_ImportaPagosDinners";
			this.Load += new System.EventHandler(this.P_ImportaPagosDinners_Load);
			((System.ComponentModel.ISupportInitialize)(this.txtPago)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid2)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.udsTmpArt)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid3)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsCompra)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsPagoLote)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbBanco)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsCuentaCorr)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCuentaCorr)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbFormaPago)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsFormaPago)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.splitContainerControl1)).EndInit();
			this.splitContainerControl1.ResumeLayout(false);
			this.ResumeLayout(false);

		}
		#endregion		
		#region Variables
		string stArchivo = "";
		DataSet miDataSet = new DataSet();
		DateTime FechaProceso = new DateTime(1900,01,01);
		#endregion Variables
		#region Funcionalidad Botones 
		private void btArchivo_Click(object sender, System.EventArgs e)
		{
			this.openFileDialog.Filter = "xls (*.xls)|*.xls";
			this.openFileDialog.InitialDirectory = @"c:\Latinium\";
			if (DialogResult.OK == this.openFileDialog.ShowDialog())
			{
				stArchivo = this.openFileDialog.FileName.ToString();
				txtPago.Text = stArchivo;
			}
		}

		private void btLeer_Click(object sender, System.EventArgs e)
		{
		if (stArchivo.Length <1)return;
		bool i = LeerDatosExcel();
			if (i == false)return;
		cargaFactAnteriores();
		}

		private void btLeerE_Click(object sender, System.EventArgs e)
		{
			BorraErrores();
			int IdC =0;

			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in ultraGrid1.Selected.Rows)
			{
				if (!Verifica1()) return;		
				if (!verificaCuentas()) return;		
				string stNumero = "Select count(*) from compra where idtipoFactura = 1 and otro = '"+this.ultraGrid1.ActiveRow.Cells["# VALE"].Value+"'";			
				if(Funcion.iEscalarSQL(cdsCompra,stNumero)<=0)
				{					
					this.errorProvider.SetError(this.ultraGrid1, "Numero no Existe ");										
				}
				else
				{
					stNumero = "Select idCompra from compra where idtipoFactura = 1 and otro = '"+dr.Cells["# VALE"].Value+"'";				
					IdC = (int) Funcion.iEscalarSQL(cdsCompra,stNumero);
				}
				string stSelect = "Select idCompra,numero, total from compra where idcompra = "+IdC;
				this.cdsCompra.Schema.Connections[0].Open();
				SqlDataReader miReader;
				C1.Data.SchemaObjects.Connection miConeccion = cdsCompra.Schema.Connections[0];
				SqlCommand cmdCant = new SqlCommand (
					stSelect, (SqlConnection) miConeccion.DbConnection);
				miReader = cmdCant.ExecuteReader();
				try
				{
					while (miReader.Read())
					{
						ultraGrid2.DisplayLayout.Bands[0].AddNew();
						int IdArt = miReader.GetInt32(0);	
						ultraGrid2.ActiveRow.Cells["idCompra"].Value = IdC;															
						ultraGrid2.ActiveRow.Cells["Total"].Value = miReader.GetDouble(2);						
						ultraGrid2.ActiveRow.Cells["Numero"].Value = miReader.GetString(1);
					}
				}
				catch(System.Exception ex)
				{
					MessageBox.Show(ex.Message, "Pasar Lineas");
				}
				miReader.Close();
			}
			this.cdsCompra.Schema.Connections[0].Close();			
		}

		private void btSubirR_Click(object sender, System.EventArgs e)
		{
			errorProvider.SetError(btSubirR,"");
			if(ultraGrid2.Rows.Count < 1)
			{
				errorProvider.SetError(this.btSubirR,"No Hay datos para Generar....");
				return;
			}
			if (DialogResult.No == MessageBox.Show("Desea Generar Pagos de estas Facturas", 
				"Información", MessageBoxButtons.YesNo, MessageBoxIcon.Question)) return;	
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in ultraGrid2.Rows.All)
			{																
				Cursor = Cursors.WaitCursor;
				string stProced = string.Format("Exec P_PagoLoteDinners {0},{1},{2},{3},'{4}','{5}',{6},'{7}'",
												dr.Cells["idCompra"].Value,cmbCuentaCorr.Value,cmbFormaPago.Value,cmbBanco.Value,
												dr.Cells["Cheque"].Value,dr.Cells["Concepto"].Value,dr.Cells["Total"].Value,FechaProceso.ToString("yyyyMMdd"));
				Funcion.EjecutaSQL(cdsCompra,stProced);
					
				Cursor = Cursors.Default;
			}
			

		}

		private void btAgregar_Click(object sender, System.EventArgs e)
		{			
			BorraErrores();

			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in ultraGrid3.Selected.Rows)
			{
				if (!Verifica()) return;	
				if (!verificaCuentas()) return;	

				int IdC = (int) dr.Cells["idCompra"].Value;
				string stSelect = "Select idCompra,numero, total from compra where idcompra = "+IdC;
				//				ultraGrid2.DataSource = Funcion.dvProcedimiento(cdsCompra,"Select idCompra,numero, total from compra where idcompra = "+IdC);

				this.cdsCompra.Schema.Connections[0].Open();
				SqlDataReader miReader;
				C1.Data.SchemaObjects.Connection miConeccion = cdsCompra.Schema.Connections[0];
				SqlCommand cmdCant = new SqlCommand (
					stSelect, (SqlConnection) miConeccion.DbConnection);
				miReader = cmdCant.ExecuteReader();
				try
				{
					while (miReader.Read())
					{
						ultraGrid2.DisplayLayout.Bands[0].AddNew();
						int IdArt = miReader.GetInt32(0);	
						ultraGrid2.ActiveRow.Cells["idCompra"].Value = IdC;															
						ultraGrid2.ActiveRow.Cells["Total"].Value = miReader.GetDouble(2);						
						ultraGrid2.ActiveRow.Cells["Numero"].Value = miReader.GetString(1);
					}
				}
				catch(System.Exception ex)
				{
					MessageBox.Show(ex.Message, "Pasar Lineas");
				}
				miReader.Close();
			}
			this.cdsCompra.Schema.Connections[0].Close();					
		}

		private void P_ImportaPagosDinners_Load(object sender, System.EventArgs e)
		{		
		}
		
		#endregion Funcionalidad Botones 		
		#region Funciones 
		private bool LeerDatosExcel()
		{
			//boton de facturas para subir del excel a la grilla
			string strConn ="Provider=Microsoft.Jet.OLEDB.4.0;" +
				"Data Source=" + stArchivo + ";" +
				"Extended Properties='Excel 8.0;'";		
			string stHoja = "select * from [Hoja1$]";
			OleDbDataAdapter myData = new OleDbDataAdapter(stHoja, strConn);
			myData.TableMappings.Add("Tabla", "ExcelSube");
			miDataSet.Clear();
			try
			{
				myData.Fill(miDataSet);
			}
			catch(Exception ex)
			{
				MessageBox.Show("Debe tener cerrado el archivo " + stArchivo + " en Excel",
					ex.Message);
				btSubirR.Enabled=false;
				return false;
			}
			ultraGrid1.DataSource = miDataSet.Tables[0].DefaultView;
			ultraGrid1.DataBind();
			return true;
		}

		private void cargaFactAnteriores()
		{
			FechaProceso = new DateTime(1900,01,01);
			if(ultraGrid1.Rows.Count < 1)return;
			C1XLBook book1 = new C1XLBook();
			try
			{
				book1.Load(this.stArchivo);
			}
			catch (Exception exception1)
			{
				MessageBox.Show(exception1.Message);
				return;
			}

			#region Verificacion de Titulos
			XLSheet sheet1 = book1.Sheets[0];			
			string[] stCampo = new string[200];
			XLCell cell1 = sheet1.GetCell(0, 0);						
			cell1 = sheet1.GetCell(0, 0);

			if ((cell1 == null) || (cell1.Value == null)) return;
			if ( cell1.Value.ToString() == "FECHA PROCESO")
			{			
				cell1 = sheet1.GetCell(1, 0);								
				if ((cell1 == null) || (cell1.Value == null)) return;				
				FechaProceso = DateTime.Parse( cell1.Value.ToString());
			}									
		
		#endregion Verificacion de Titulos

			if (FechaProceso.ToString("yyyyMMdd") == "19000101") return;
				string Vale="";
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in ultraGrid1.Rows.All)
			{
				if (dr.Cells["# VALE"].Value.ToString() != "")
				Vale = Vale + dr.Cells["# VALE"].Value+",";
			}
			Vale = Vale.Substring(0,Vale.Length -1);


			string stSelect = " Exec P_DinnersLlenaFacturas '"+FechaProceso.ToString("yyyyMMdd")+"','"+Vale+"'";
			ultraGrid3.DataSource = Funcion.dvProcedimiento(cdsCompra, stSelect);
			ultraGrid3.DisplayLayout.Bands[0].Columns["idCompra"].Hidden = true;
			
		}
	
		private void BorraErrores()
		{		
			this.errorProvider.SetError(this.ultraGrid3, "");		
			this.errorProvider.SetError(this.ultraGrid2, "");
			this.errorProvider.SetError(this.ultraGrid1, "");
			this.errorProvider.SetError(this.cmbCuentaCorr, "");
			this.errorProvider.SetError(this.cmbFormaPago, "");
			this.errorProvider.SetError(this.cmbBanco, "");
		}

		private bool Verifica()
		{
			// Verificacion de Factura no repetida
			bool bOk = true;
			int IdC = (int) this.ultraGrid3.ActiveRow.Cells["idCompra"].Value;		
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.ultraGrid2.Rows.All)
			{
				if (dr.Cells["idCompra"].Value == System.DBNull.Value)
				{
					this.errorProvider.SetError(this.ultraGrid3, "IdCompra no valido");
					bOk = false;
				}
				else
				{
					if (IdC == (int) dr.Cells["idCompra"].Value)
					{
						this.errorProvider.SetError(this.ultraGrid2, "Factura ya existe en lista de pagos");
						bOk = false;
					}
				}
			}
			return bOk;
		}
		private bool Verifica1()
		{
			// Verificacion de Factura no repetida
			bool bOk = true;	int IdC=0;
			string stNumero = "Select count(*) from compra where idtipoFactura = 1 and otro = '"+this.ultraGrid1.ActiveRow.Cells["# VALE"].Value+"'";			
			if(Funcion.iEscalarSQL(cdsCompra,stNumero)<=0)
			{
				bOk = false;
				this.errorProvider.SetError(this.ultraGrid1, "Numero no Existe ");										
			}
			else
			{
				stNumero = "Select idCompra from compra where idtipoFactura = 1 and otro = '"+this.ultraGrid1.ActiveRow.Cells["# VALE"].Value+"'";				
				IdC = (int) Funcion.iEscalarSQL(cdsCompra,stNumero);
			}
	
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.ultraGrid2.Rows.All)
			{
				if (dr.Cells["idCompra"].Value == System.DBNull.Value)
				{
					bOk = false;
					this.errorProvider.SetError(this.ultraGrid2, "IdCompra no valido");					
				}
				else
				{
					if (IdC == (int) dr.Cells["idCompra"].Value)
					{
						bOk = false;
						this.errorProvider.SetError(this.ultraGrid2, "Factura ya existe en lista de pagos");						
					}
				}
			}
			return bOk;
		}
		
		private bool verificaCuentas()
		{
				bool bOk = true;		
				
					if ((this.cmbFormaPago.Value == System.DBNull.Value)||((int) this.cmbFormaPago.Value == 0 ))
					{
						this.errorProvider.SetError(this.cmbFormaPago, "Ingrese Forma de Pago");
						bOk = false;
					}
			if ((this.cmbCuentaCorr.Value == System.DBNull.Value)||((int) this.cmbCuentaCorr.Value == 0))
			{
				this.errorProvider.SetError(this.cmbCuentaCorr, "Ingrese Cuenta corriente");
				bOk = false;
			}			
				if((cmbBanco.Value == System.DBNull.Value) || ((int) this.cmbBanco.Value == 0))
				{
					this.errorProvider.SetError(this.cmbBanco, "Ingrese Banco");
					bOk = false;
				}
				return bOk;
		}

		#endregion Funciones 

		private void cdsCompra_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsCompra.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);
		}
		
		private void cdsPagoLote_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			this.cdsPagoLote.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);
		}

		private void cdsFormaPago_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			this.cdsFormaPago.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);
		}

		private void cdsCuentaCorr_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			this.cdsCuentaCorr.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirConta);
		}

		private void btCarga_Click(object sender, System.EventArgs e)
		{
			
		}

		

	



	}
}

