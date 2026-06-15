using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using DevExpress.XtraEditors;
using System.Globalization;
using Infragistics.Win.UltraWinGrid;
using System.Data;
using System.Data.SqlClient;

namespace Latinium
{
	/// <summary>
	/// Summary description for frmDetalleGastosTarjetas.
	/// </summary>
	public class frmDetalleGastosTarjetas : DevExpress.XtraEditors.XtraForm
	{
		public Infragistics.Win.UltraWinGrid.UltraGrid uGridGastos;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource1;
		private System.Windows.Forms.Button btnAceptar;
		private System.Windows.Forms.Button btnCancelar;
		private C1.Data.C1DataSet cdsSeteoF;
		private System.Windows.Forms.Label lblSubTotal;
		private System.Windows.Forms.Label lblIva;
		private System.Windows.Forms.Label label1;
		private System.Windows.Forms.Label label2;
		private System.Windows.Forms.GroupBox groupBox1;
		public Infragistics.Win.UltraWinEditors.UltraNumericEditor txtIVA;
		public Infragistics.Win.UltraWinEditors.UltraNumericEditor txtSubtotalIVA0;
		public Infragistics.Win.UltraWinEditors.UltraNumericEditor txtSubtotalIVA;
		public Infragistics.Win.UltraWinEditors.UltraNumericEditor txtTotal;
		/// <summary>
		/// Required designer variable.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public frmDetalleGastosTarjetas()
		{
			//
			// Required for Windows Form Designer support
			//
			InitializeComponent();

			//
			// TODO: Add any constructor code after InitializeComponent call
			//
		}

		#region Variables
		bool bEliminaAlValidar = false;

		int idArticulo1 = 0;
		double dIVA1 = 0;
		double dSubtotal1 = 0;

		int idArticulo2 = 0;
		double dIVA2 = 0;
		double dSubtotal2 = 0;

		int idArticulo3 = 0;
		double dIVA3 = 0;
		double dSubtotal3 = 0;

		int idArticulo4 = 0;
		double dIVA4 = 0;
		double dSubtotal4 = 0;

		int idArticulo5 = 0;
		double dIVA5 = 0;
		double dSubtotal5 = 0;

		int idArticulo6 = 0;
		double dIVA6 = 0;
		double dSubtotal6 = 0;

		int idArticulo7 = 0;
		double dIVA7 = 0;
		double dSubtotal7 = 0;

		int idArticulo8 = 0;
		double dIVA8 = 0;
		double dSubtotal8 = 0;

		int idArticulo9 = 0;
		double dIVA9 = 0;
		double dSubtotal9 = 0;

		int idArticulo10 = 0;
		double dIVA10 = 0;
		double dSubtotal10 = 0;

		int idArticulo11 = 0;
		double dIVA11 = 0;
		double dSubtotal11 = 0;

		int idArticulo12 = 0;
		double dIVA12 = 0;
		double dSubtotal12 = 0;

		int idArticulo13 = 0;
		double dIVA13 = 0;
		double dSubtotal13 = 0;

		int idArticulo14 = 0;
		double dIVA14 = 0;
		double dSubtotal14 = 0;

		int idArticulo15 = 0;
		double dIVA15 = 0;
		double dSubtotal15 = 0;

		bool BSinFactura = false;
		#endregion Variables

		public frmDetalleGastosTarjetas(int IdArticulo1, double DIVA1, double DSubtotal1, int IdArticulo2, double DIVA2, double DSubtotal2, int IdArticulo3, double DIVA3, double DSubtotal3, int IdArticulo4, double DIVA4, double DSubtotal4, int IdArticulo5, double DIVA5, double DSubtotal5, int IdArticulo6, double DIVA6, double DSubtotal6, int IdArticulo7, double DIVA7, double DSubtotal7, int IdArticulo8, double DIVA8, double DSubtotal8, int IdArticulo9, double DIVA9, double DSubtotal9, int IdArticulo10, double DIVA10, double DSubtotal10, int IdArticulo11, double DIVA11, double DSubtotal11, int IdArticulo12, double DIVA12, double DSubtotal12, int IdArticulo13, double DIVA13, double DSubtotal13, int IdArticulo14, double DIVA14, double DSubtotal14, int IdArticulo15, double DIVA15, double DSubtotal15, bool bSinFactura)
		{
			//
			// Required for Windows Form Designer support
			//
			InitializeComponent();

			#region Asigna Valores
			idArticulo1 = IdArticulo1;
			dIVA1 = DIVA1;
			dSubtotal1 = DSubtotal1;

			idArticulo2 = IdArticulo2;
			dIVA2 = DIVA2;
			dSubtotal2 = DSubtotal2;

			idArticulo3 = IdArticulo3;
			dIVA3 = DIVA3;
			dSubtotal3 = DSubtotal3;

			idArticulo4 = IdArticulo4;
			dIVA4 = DIVA4;
			dSubtotal4 = DSubtotal4;

			idArticulo5 = IdArticulo5;
			dIVA5 = DIVA5;
			dSubtotal5 = DSubtotal5;

			idArticulo6 = IdArticulo6;
			dIVA6 = DIVA6;
			dSubtotal6 = DSubtotal6;

			idArticulo7 = IdArticulo7;
			dIVA7 = DIVA7;
			dSubtotal7 = DSubtotal7;

			idArticulo8 = IdArticulo8;
			dIVA8 = DIVA8;
			dSubtotal8 = DSubtotal8;

			idArticulo9 = IdArticulo9;
			dIVA9 = DIVA9;
			dSubtotal9 = DSubtotal9;

			idArticulo10 = IdArticulo10;
			dIVA10 = DIVA10;
			dSubtotal10 = DSubtotal10;

			idArticulo11 = IdArticulo11;
			dIVA11 = DIVA11;
			dSubtotal11 = DSubtotal11;

			idArticulo12 = IdArticulo12;
			dIVA12 = DIVA12;
			dSubtotal12 = DSubtotal12;

			idArticulo13 = IdArticulo13;
			dIVA13 = DIVA13;
			dSubtotal13 = DSubtotal13;

			idArticulo14 = IdArticulo14;
			dIVA14 = DIVA14;
			dSubtotal14 = DSubtotal14;

			idArticulo15 = IdArticulo15;
			dIVA15 = DIVA15;
			dSubtotal15 = DSubtotal15;

			BSinFactura = bSinFactura;
			#endregion Asigna Valores
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
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idDetalleGastosTarjetas");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Posicion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idArticulo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Articulo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("IVA");
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Subtotal");
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idDetalleGastosTarjetas");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Posicion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn3 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idArticulo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn4 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Codigo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn5 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Articulo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn6 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("IVA");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn7 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Subtotal");
			Infragistics.Win.Appearance appearance11 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance12 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance13 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance14 = new Infragistics.Win.Appearance();
			this.uGridGastos = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource1 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.btnAceptar = new System.Windows.Forms.Button();
			this.btnCancelar = new System.Windows.Forms.Button();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.lblSubTotal = new System.Windows.Forms.Label();
			this.lblIva = new System.Windows.Forms.Label();
			this.txtIVA = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtSubtotalIVA0 = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtSubtotalIVA = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label1 = new System.Windows.Forms.Label();
			this.label2 = new System.Windows.Forms.Label();
			this.txtTotal = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.groupBox1 = new System.Windows.Forms.GroupBox();
			((System.ComponentModel.ISupportInitialize)(this.uGridGastos)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIVA)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtSubtotalIVA0)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtSubtotalIVA)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtTotal)).BeginInit();
			this.SuspendLayout();
			// 
			// uGridGastos
			// 
			this.uGridGastos.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.uGridGastos.Cursor = System.Windows.Forms.Cursors.Default;
			this.uGridGastos.DataSource = this.ultraDataSource1;
			appearance1.BackColor = System.Drawing.Color.White;
			this.uGridGastos.DisplayLayout.Appearance = appearance1;
			ultraGridColumn1.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn1.Width = 180;
			ultraGridColumn2.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn2.Header.Caption = "N.";
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridColumn2.Width = 25;
			ultraGridColumn3.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn3.Header.VisiblePosition = 2;
			ultraGridColumn3.Hidden = true;
			ultraGridColumn3.Width = 81;
			ultraGridColumn4.Header.VisiblePosition = 3;
			ultraGridColumn4.Width = 186;
			ultraGridColumn5.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn5.Header.VisiblePosition = 4;
			ultraGridColumn5.Width = 300;
			appearance2.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn6.CellAppearance = appearance2;
			ultraGridColumn6.Format = "#,##0.00";
			ultraGridColumn6.Header.VisiblePosition = 5;
			ultraGridColumn6.Width = 60;
			appearance3.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn7.CellAppearance = appearance3;
			ultraGridColumn7.Format = "#,##0.00";
			ultraGridColumn7.Header.Caption = "Valor";
			ultraGridColumn7.Header.VisiblePosition = 6;
			ultraGridColumn7.Width = 90;
			ultraGridBand1.Columns.AddRange(new object[] {
																										 ultraGridColumn1,
																										 ultraGridColumn2,
																										 ultraGridColumn3,
																										 ultraGridColumn4,
																										 ultraGridColumn5,
																										 ultraGridColumn6,
																										 ultraGridColumn7});
			ultraGridBand1.Header.Caption = "";
			ultraGridBand1.SummaryFooterCaption = "";
			this.uGridGastos.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			appearance4.ForeColor = System.Drawing.Color.Black;
			appearance4.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridGastos.DisplayLayout.Override.ActiveRowAppearance = appearance4;
			this.uGridGastos.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.TemplateOnBottom;
			this.uGridGastos.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.True;
			this.uGridGastos.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance5.BackColor = System.Drawing.Color.Transparent;
			this.uGridGastos.DisplayLayout.Override.CardAreaAppearance = appearance5;
			appearance6.ForeColor = System.Drawing.Color.Black;
			appearance6.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridGastos.DisplayLayout.Override.CellAppearance = appearance6;
			appearance7.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance7.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance7.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance7.FontData.BoldAsString = "True";
			appearance7.FontData.Name = "Arial";
			appearance7.FontData.SizeInPoints = 10F;
			appearance7.ForeColor = System.Drawing.Color.White;
			appearance7.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridGastos.DisplayLayout.Override.HeaderAppearance = appearance7;
			appearance8.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance8.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance8.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridGastos.DisplayLayout.Override.RowAlternateAppearance = appearance8;
			appearance9.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance9.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance9.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridGastos.DisplayLayout.Override.RowSelectorAppearance = appearance9;
			appearance10.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance10.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance10.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridGastos.DisplayLayout.Override.SelectedRowAppearance = appearance10;
			this.uGridGastos.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridGastos.Location = new System.Drawing.Point(8, 8);
			this.uGridGastos.Name = "uGridGastos";
			this.uGridGastos.Size = new System.Drawing.Size(682, 360);
			this.uGridGastos.TabIndex = 608;
			this.uGridGastos.AfterRowsDeleted += new System.EventHandler(this.uGridGastos_AfterRowsDeleted);
			this.uGridGastos.AfterRowInsert += new Infragistics.Win.UltraWinGrid.RowEventHandler(this.uGridGastos_AfterRowInsert);
			this.uGridGastos.BeforeCellUpdate += new Infragistics.Win.UltraWinGrid.BeforeCellUpdateEventHandler(this.uGridGastos_BeforeCellUpdate);
			this.uGridGastos.BeforeRowsDeleted += new Infragistics.Win.UltraWinGrid.BeforeRowsDeletedEventHandler(this.uGridGastos_BeforeRowsDeleted);
			this.uGridGastos.BeforeRowInsert += new Infragistics.Win.UltraWinGrid.BeforeRowInsertEventHandler(this.uGridGastos_BeforeRowInsert);
			this.uGridGastos.KeyDown += new System.Windows.Forms.KeyEventHandler(this.uGridGastos_KeyDown);
			this.uGridGastos.AfterCellUpdate += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.uGridGastos_AfterCellUpdate);
			this.uGridGastos.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.uGridGastos_InitializeLayout);
			// 
			// ultraDataSource1
			// 
			ultraDataColumn1.DataType = typeof(int);
			ultraDataColumn2.DataType = typeof(int);
			ultraDataColumn3.DataType = typeof(int);
			ultraDataColumn6.DataType = typeof(System.Double);
			ultraDataColumn7.DataType = typeof(System.Double);
			this.ultraDataSource1.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn1,
																																 ultraDataColumn2,
																																 ultraDataColumn3,
																																 ultraDataColumn4,
																																 ultraDataColumn5,
																																 ultraDataColumn6,
																																 ultraDataColumn7});
			// 
			// btnAceptar
			// 
			this.btnAceptar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnAceptar.CausesValidation = false;
			this.btnAceptar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnAceptar.ImeMode = System.Windows.Forms.ImeMode.NoControl;
			this.btnAceptar.Location = new System.Drawing.Point(8, 458);
			this.btnAceptar.Name = "btnAceptar";
			this.btnAceptar.Size = new System.Drawing.Size(88, 23);
			this.btnAceptar.TabIndex = 610;
			this.btnAceptar.Text = "&Aceptar";
			this.btnAceptar.Click += new System.EventHandler(this.btnAceptar_Click);
			// 
			// btnCancelar
			// 
			this.btnCancelar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnCancelar.CausesValidation = false;
			this.btnCancelar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnCancelar.Location = new System.Drawing.Point(104, 458);
			this.btnCancelar.Name = "btnCancelar";
			this.btnCancelar.Size = new System.Drawing.Size(88, 23);
			this.btnCancelar.TabIndex = 609;
			this.btnCancelar.Text = "&Cancelar";
			this.btnCancelar.Click += new System.EventHandler(this.btnCancelar_Click);
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
			// lblSubTotal
			// 
			this.lblSubTotal.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.lblSubTotal.AutoSize = true;
			this.lblSubTotal.Location = new System.Drawing.Point(456, 395);
			this.lblSubTotal.Name = "lblSubTotal";
			this.lblSubTotal.Size = new System.Drawing.Size(98, 16);
			this.lblSubTotal.TabIndex = 613;
			this.lblSubTotal.Text = "Subtotal I.V.A. > 0";
			this.lblSubTotal.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// lblIva
			// 
			this.lblIva.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.lblIva.AutoSize = true;
			this.lblIva.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblIva.Location = new System.Drawing.Point(456, 443);
			this.lblIva.Name = "lblIva";
			this.lblIva.Size = new System.Drawing.Size(22, 16);
			this.lblIva.TabIndex = 612;
			this.lblIva.Text = "IVA";
			this.lblIva.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtIVA
			// 
			this.txtIVA.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			appearance11.BackColorDisabled = System.Drawing.SystemColors.Control;
			appearance11.BackColorDisabled2 = System.Drawing.SystemColors.Control;
			appearance11.BorderColor = System.Drawing.Color.Black;
			appearance11.BorderColor3DBase = System.Drawing.Color.Black;
			appearance11.ForeColor = System.Drawing.Color.Blue;
			appearance11.ForeColorDisabled = System.Drawing.Color.Blue;
			this.txtIVA.Appearance = appearance11;
			this.txtIVA.BorderStyle = Infragistics.Win.UIElementBorderStyle.None;
			this.txtIVA.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtIVA.Enabled = false;
			this.txtIVA.FormatString = "#,##0.00";
			this.txtIVA.Location = new System.Drawing.Point(592, 440);
			this.txtIVA.Name = "txtIVA";
			this.txtIVA.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtIVA.PromptChar = ' ';
			this.txtIVA.Size = new System.Drawing.Size(96, 18);
			this.txtIVA.TabIndex = 616;
			// 
			// txtSubtotalIVA0
			// 
			this.txtSubtotalIVA0.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			appearance12.BackColorDisabled = System.Drawing.SystemColors.Control;
			appearance12.BackColorDisabled2 = System.Drawing.SystemColors.Control;
			appearance12.BorderColor = System.Drawing.Color.Black;
			appearance12.BorderColor3DBase = System.Drawing.Color.Black;
			appearance12.ForeColor = System.Drawing.Color.Blue;
			appearance12.ForeColorDisabled = System.Drawing.Color.Blue;
			this.txtSubtotalIVA0.Appearance = appearance12;
			this.txtSubtotalIVA0.BorderStyle = Infragistics.Win.UIElementBorderStyle.None;
			this.txtSubtotalIVA0.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtSubtotalIVA0.Enabled = false;
			this.txtSubtotalIVA0.FormatString = "#,##0.00";
			this.txtSubtotalIVA0.Location = new System.Drawing.Point(592, 416);
			this.txtSubtotalIVA0.Name = "txtSubtotalIVA0";
			this.txtSubtotalIVA0.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtSubtotalIVA0.PromptChar = ' ';
			this.txtSubtotalIVA0.Size = new System.Drawing.Size(96, 18);
			this.txtSubtotalIVA0.TabIndex = 615;
			// 
			// txtSubtotalIVA
			// 
			this.txtSubtotalIVA.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			appearance13.BackColorDisabled = System.Drawing.SystemColors.Control;
			appearance13.BackColorDisabled2 = System.Drawing.SystemColors.Control;
			appearance13.BorderColor = System.Drawing.Color.Black;
			appearance13.BorderColor3DBase = System.Drawing.Color.Black;
			appearance13.ForeColor = System.Drawing.Color.Blue;
			appearance13.ForeColorDisabled = System.Drawing.Color.Blue;
			this.txtSubtotalIVA.Appearance = appearance13;
			this.txtSubtotalIVA.BorderStyle = Infragistics.Win.UIElementBorderStyle.None;
			this.txtSubtotalIVA.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtSubtotalIVA.Enabled = false;
			this.txtSubtotalIVA.FormatString = "#,##0.00";
			this.txtSubtotalIVA.Location = new System.Drawing.Point(592, 392);
			this.txtSubtotalIVA.Name = "txtSubtotalIVA";
			this.txtSubtotalIVA.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtSubtotalIVA.PromptChar = ' ';
			this.txtSubtotalIVA.Size = new System.Drawing.Size(96, 18);
			this.txtSubtotalIVA.TabIndex = 614;
			// 
			// label1
			// 
			this.label1.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.label1.AutoSize = true;
			this.label1.Location = new System.Drawing.Point(456, 419);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(87, 16);
			this.label1.TabIndex = 617;
			this.label1.Text = "Subtotal I.V.A. 0";
			this.label1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label2
			// 
			this.label2.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.label2.AutoSize = true;
			this.label2.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label2.Location = new System.Drawing.Point(456, 467);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(29, 16);
			this.label2.TabIndex = 618;
			this.label2.Text = "Total";
			this.label2.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtTotal
			// 
			this.txtTotal.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			appearance14.BackColorDisabled = System.Drawing.SystemColors.Control;
			appearance14.BackColorDisabled2 = System.Drawing.SystemColors.Control;
			appearance14.BorderColor = System.Drawing.Color.Black;
			appearance14.BorderColor3DBase = System.Drawing.Color.Black;
			appearance14.ForeColor = System.Drawing.Color.Blue;
			appearance14.ForeColorDisabled = System.Drawing.Color.Blue;
			this.txtTotal.Appearance = appearance14;
			this.txtTotal.BorderStyle = Infragistics.Win.UIElementBorderStyle.None;
			this.txtTotal.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtTotal.Enabled = false;
			this.txtTotal.FormatString = "#,##0.00";
			this.txtTotal.Location = new System.Drawing.Point(592, 464);
			this.txtTotal.Name = "txtTotal";
			this.txtTotal.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtTotal.PromptChar = ' ';
			this.txtTotal.Size = new System.Drawing.Size(96, 18);
			this.txtTotal.TabIndex = 619;
			// 
			// groupBox1
			// 
			this.groupBox1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.groupBox1.Location = new System.Drawing.Point(0, 376);
			this.groupBox1.Name = "groupBox1";
			this.groupBox1.Size = new System.Drawing.Size(704, 8);
			this.groupBox1.TabIndex = 620;
			this.groupBox1.TabStop = false;
			// 
			// frmDetalleGastosTarjetas
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 14);
			this.CausesValidation = false;
			this.ClientSize = new System.Drawing.Size(698, 488);
			this.Controls.Add(this.groupBox1);
			this.Controls.Add(this.label2);
			this.Controls.Add(this.txtTotal);
			this.Controls.Add(this.label1);
			this.Controls.Add(this.lblSubTotal);
			this.Controls.Add(this.lblIva);
			this.Controls.Add(this.txtIVA);
			this.Controls.Add(this.txtSubtotalIVA0);
			this.Controls.Add(this.txtSubtotalIVA);
			this.Controls.Add(this.btnAceptar);
			this.Controls.Add(this.btnCancelar);
			this.Controls.Add(this.uGridGastos);
			this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
			this.KeyPreview = true;
			this.MaximizeBox = false;
			this.MinimizeBox = false;
			this.Name = "frmDetalleGastosTarjetas";
			this.ShowInTaskbar = false;
			this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
			this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.frmDetalleGastosTarjetas_KeyDown);
			this.Load += new System.EventHandler(this.frmDetalleGastosTarjetas_Load);
			((System.ComponentModel.ISupportInitialize)(this.uGridGastos)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIVA)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtSubtotalIVA0)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtSubtotalIVA)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtTotal)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void btnCancelar_Click(object sender, System.EventArgs e)
		{
			this.DialogResult = DialogResult.No;
		}

		private void frmDetalleGastosTarjetas_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Escape) this.DialogResult = DialogResult.No;
		}

		private void frmDetalleGastosTarjetas_Load(object sender, System.EventArgs e)
		{
			#region Asigna Valores al Grid
			if (idArticulo1 > 0)
			{
				this.uGridGastos.Rows.Band.AddNew();
				this.uGridGastos.ActiveRow.Cells["Posicion"].Value = 1;
				this.uGridGastos.ActiveRow.Cells["idArticulo"].Value = idArticulo1;
				this.uGridGastos.ActiveRow.Cells["Codigo"].Value = Funcion.sEscalarSQL(cdsSeteoF, string.Format("Select Codigo From Articulo Where idArticulo = {0}", idArticulo1));
				this.uGridGastos.ActiveRow.Cells["Articulo"].Value = Funcion.sEscalarSQL(cdsSeteoF, string.Format("Select Articulo From Articulo Where idArticulo = {0}", idArticulo1));
				this.uGridGastos.ActiveRow.Cells["IVA"].Value = dIVA1;
				this.uGridGastos.ActiveRow.Cells["Subtotal"].Value = dSubtotal1;								
			}

			if (idArticulo2 > 0)
			{
				this.uGridGastos.Rows.Band.AddNew();
				this.uGridGastos.ActiveRow.Cells["Posicion"].Value = 2;
				this.uGridGastos.ActiveRow.Cells["idArticulo"].Value = idArticulo2;
				this.uGridGastos.ActiveRow.Cells["Codigo"].Value = Funcion.sEscalarSQL(cdsSeteoF, string.Format("Select Codigo From Articulo Where idArticulo = {0}", idArticulo2));
				this.uGridGastos.ActiveRow.Cells["Articulo"].Value = Funcion.sEscalarSQL(cdsSeteoF, string.Format("Select Articulo From Articulo Where idArticulo = {0}", idArticulo2));
				this.uGridGastos.ActiveRow.Cells["IVA"].Value = dIVA2;
				this.uGridGastos.ActiveRow.Cells["Subtotal"].Value = dSubtotal2;				
			}

			if (idArticulo3 > 0)
			{
				this.uGridGastos.Rows.Band.AddNew();
				this.uGridGastos.ActiveRow.Cells["Posicion"].Value = 3;
				this.uGridGastos.ActiveRow.Cells["idArticulo"].Value = idArticulo3;
				this.uGridGastos.ActiveRow.Cells["Codigo"].Value = Funcion.sEscalarSQL(cdsSeteoF, string.Format("Select Codigo From Articulo Where idArticulo = {0}", idArticulo3));
				this.uGridGastos.ActiveRow.Cells["Articulo"].Value = Funcion.sEscalarSQL(cdsSeteoF, string.Format("Select Articulo From Articulo Where idArticulo = {0}", idArticulo3));
				this.uGridGastos.ActiveRow.Cells["IVA"].Value = dIVA3;
				this.uGridGastos.ActiveRow.Cells["Subtotal"].Value = dSubtotal3;				
			}

			if (idArticulo4 > 0)
			{
				this.uGridGastos.Rows.Band.AddNew();
				this.uGridGastos.ActiveRow.Cells["Posicion"].Value = 4;
				this.uGridGastos.ActiveRow.Cells["idArticulo"].Value = idArticulo4;
				this.uGridGastos.ActiveRow.Cells["Codigo"].Value = Funcion.sEscalarSQL(cdsSeteoF, string.Format("Select Codigo From Articulo Where idArticulo = {0}", idArticulo4));
				this.uGridGastos.ActiveRow.Cells["Articulo"].Value = Funcion.sEscalarSQL(cdsSeteoF, string.Format("Select Articulo From Articulo Where idArticulo = {0}", idArticulo4));
				this.uGridGastos.ActiveRow.Cells["IVA"].Value = dIVA4;
				this.uGridGastos.ActiveRow.Cells["Subtotal"].Value = dSubtotal4;								
			}

			if (idArticulo5 > 0)
			{
				this.uGridGastos.Rows.Band.AddNew();
				this.uGridGastos.ActiveRow.Cells["Posicion"].Value = 5;
				this.uGridGastos.ActiveRow.Cells["idArticulo"].Value = idArticulo5;
				this.uGridGastos.ActiveRow.Cells["Codigo"].Value = Funcion.sEscalarSQL(cdsSeteoF, string.Format("Select Codigo From Articulo Where idArticulo = {0}", idArticulo5));
				this.uGridGastos.ActiveRow.Cells["Articulo"].Value = Funcion.sEscalarSQL(cdsSeteoF, string.Format("Select Articulo From Articulo Where idArticulo = {0}", idArticulo5));
				this.uGridGastos.ActiveRow.Cells["IVA"].Value = dIVA5;
				this.uGridGastos.ActiveRow.Cells["Subtotal"].Value = dSubtotal5;								
			}

			if (idArticulo6 > 0)
			{
				this.uGridGastos.Rows.Band.AddNew();
				this.uGridGastos.ActiveRow.Cells["Posicion"].Value = 6;
				this.uGridGastos.ActiveRow.Cells["idArticulo"].Value = idArticulo6;
				this.uGridGastos.ActiveRow.Cells["Codigo"].Value = Funcion.sEscalarSQL(cdsSeteoF, string.Format("Select Codigo From Articulo Where idArticulo = {0}", idArticulo6));
				this.uGridGastos.ActiveRow.Cells["Articulo"].Value = Funcion.sEscalarSQL(cdsSeteoF, string.Format("Select Articulo From Articulo Where idArticulo = {0}", idArticulo6));
				this.uGridGastos.ActiveRow.Cells["IVA"].Value = dIVA6;
				this.uGridGastos.ActiveRow.Cells["Subtotal"].Value = dSubtotal6;								
			}

			if (idArticulo7 > 0)
			{
				this.uGridGastos.Rows.Band.AddNew();
				this.uGridGastos.ActiveRow.Cells["Posicion"].Value = 7;
				this.uGridGastos.ActiveRow.Cells["idArticulo"].Value = idArticulo7;
				this.uGridGastos.ActiveRow.Cells["Codigo"].Value = Funcion.sEscalarSQL(cdsSeteoF, string.Format("Select Codigo From Articulo Where idArticulo = {0}", idArticulo7));
				this.uGridGastos.ActiveRow.Cells["Articulo"].Value = Funcion.sEscalarSQL(cdsSeteoF, string.Format("Select Articulo From Articulo Where idArticulo = {0}", idArticulo7));
				this.uGridGastos.ActiveRow.Cells["IVA"].Value = dIVA7;
				this.uGridGastos.ActiveRow.Cells["Subtotal"].Value = dSubtotal7;								
			}

			if (idArticulo8 > 0)
			{
				this.uGridGastos.Rows.Band.AddNew();
				this.uGridGastos.ActiveRow.Cells["Posicion"].Value = 8;
				this.uGridGastos.ActiveRow.Cells["idArticulo"].Value = idArticulo8;
				this.uGridGastos.ActiveRow.Cells["Codigo"].Value = Funcion.sEscalarSQL(cdsSeteoF, string.Format("Select Codigo From Articulo Where idArticulo = {0}", idArticulo8));
				this.uGridGastos.ActiveRow.Cells["Articulo"].Value = Funcion.sEscalarSQL(cdsSeteoF, string.Format("Select Articulo From Articulo Where idArticulo = {0}", idArticulo8));
				this.uGridGastos.ActiveRow.Cells["IVA"].Value = dIVA8;
				this.uGridGastos.ActiveRow.Cells["Subtotal"].Value = dSubtotal8;								
			}

			if (idArticulo9 > 0)
			{
				this.uGridGastos.Rows.Band.AddNew();
				this.uGridGastos.ActiveRow.Cells["Posicion"].Value = 9;
				this.uGridGastos.ActiveRow.Cells["idArticulo"].Value = idArticulo9;
				this.uGridGastos.ActiveRow.Cells["Codigo"].Value = Funcion.sEscalarSQL(cdsSeteoF, string.Format("Select Codigo From Articulo Where idArticulo = {0}", idArticulo9));
				this.uGridGastos.ActiveRow.Cells["Articulo"].Value = Funcion.sEscalarSQL(cdsSeteoF, string.Format("Select Articulo From Articulo Where idArticulo = {0}", idArticulo9));
				this.uGridGastos.ActiveRow.Cells["IVA"].Value = dIVA9;
				this.uGridGastos.ActiveRow.Cells["Subtotal"].Value = dSubtotal9;								
			}

			if (idArticulo10 > 0)
			{
				this.uGridGastos.Rows.Band.AddNew();
				this.uGridGastos.ActiveRow.Cells["Posicion"].Value = 10;
				this.uGridGastos.ActiveRow.Cells["idArticulo"].Value = idArticulo10;
				this.uGridGastos.ActiveRow.Cells["Codigo"].Value = Funcion.sEscalarSQL(cdsSeteoF, string.Format("Select Codigo From Articulo Where idArticulo = {0}", idArticulo10));
				this.uGridGastos.ActiveRow.Cells["Articulo"].Value = Funcion.sEscalarSQL(cdsSeteoF, string.Format("Select Articulo From Articulo Where idArticulo = {0}", idArticulo10));
				this.uGridGastos.ActiveRow.Cells["IVA"].Value = dIVA10;
				this.uGridGastos.ActiveRow.Cells["Subtotal"].Value = dSubtotal10;								
			}

			if (idArticulo11 > 0)
			{
				this.uGridGastos.Rows.Band.AddNew();
				this.uGridGastos.ActiveRow.Cells["Posicion"].Value = 11;
				this.uGridGastos.ActiveRow.Cells["idArticulo"].Value = idArticulo11;
				this.uGridGastos.ActiveRow.Cells["Codigo"].Value = Funcion.sEscalarSQL(cdsSeteoF, string.Format("Select Codigo From Articulo Where idArticulo = {0}", idArticulo11));
				this.uGridGastos.ActiveRow.Cells["Articulo"].Value = Funcion.sEscalarSQL(cdsSeteoF, string.Format("Select Articulo From Articulo Where idArticulo = {0}", idArticulo11));
				this.uGridGastos.ActiveRow.Cells["IVA"].Value = dIVA11;
				this.uGridGastos.ActiveRow.Cells["Subtotal"].Value = dSubtotal11;								
			}

			if (idArticulo12 > 0)
			{
				this.uGridGastos.Rows.Band.AddNew();
				this.uGridGastos.ActiveRow.Cells["Posicion"].Value = 12;
				this.uGridGastos.ActiveRow.Cells["idArticulo"].Value = idArticulo12;
				this.uGridGastos.ActiveRow.Cells["Codigo"].Value = Funcion.sEscalarSQL(cdsSeteoF, string.Format("Select Codigo From Articulo Where idArticulo = {0}", idArticulo12));
				this.uGridGastos.ActiveRow.Cells["Articulo"].Value = Funcion.sEscalarSQL(cdsSeteoF, string.Format("Select Articulo From Articulo Where idArticulo = {0}", idArticulo12));
				this.uGridGastos.ActiveRow.Cells["IVA"].Value = dIVA12;
				this.uGridGastos.ActiveRow.Cells["Subtotal"].Value = dSubtotal12;								
			}

			if (idArticulo13 > 0)
			{
				this.uGridGastos.Rows.Band.AddNew();
				this.uGridGastos.ActiveRow.Cells["Posicion"].Value = 13;
				this.uGridGastos.ActiveRow.Cells["idArticulo"].Value = idArticulo13;
				this.uGridGastos.ActiveRow.Cells["Codigo"].Value = Funcion.sEscalarSQL(cdsSeteoF, string.Format("Select Codigo From Articulo Where idArticulo = {0}", idArticulo13));
				this.uGridGastos.ActiveRow.Cells["Articulo"].Value = Funcion.sEscalarSQL(cdsSeteoF, string.Format("Select Articulo From Articulo Where idArticulo = {0}", idArticulo13));
				this.uGridGastos.ActiveRow.Cells["IVA"].Value = dIVA13;
				this.uGridGastos.ActiveRow.Cells["Subtotal"].Value = dSubtotal13;								
			}

			if (idArticulo14 > 0)
			{
				this.uGridGastos.Rows.Band.AddNew();
				this.uGridGastos.ActiveRow.Cells["Posicion"].Value = 14;
				this.uGridGastos.ActiveRow.Cells["idArticulo"].Value = idArticulo14;
				this.uGridGastos.ActiveRow.Cells["Codigo"].Value = Funcion.sEscalarSQL(cdsSeteoF, string.Format("Select Codigo From Articulo Where idArticulo = {0}", idArticulo14));
				this.uGridGastos.ActiveRow.Cells["Articulo"].Value = Funcion.sEscalarSQL(cdsSeteoF, string.Format("Select Articulo From Articulo Where idArticulo = {0}", idArticulo14));
				this.uGridGastos.ActiveRow.Cells["IVA"].Value = dIVA14;
				this.uGridGastos.ActiveRow.Cells["Subtotal"].Value = dSubtotal14;								
			}

			if (idArticulo15 > 0)
			{
				this.uGridGastos.Rows.Band.AddNew();
				this.uGridGastos.ActiveRow.Cells["Posicion"].Value = 15;
				this.uGridGastos.ActiveRow.Cells["idArticulo"].Value = idArticulo15;
				this.uGridGastos.ActiveRow.Cells["Codigo"].Value = Funcion.sEscalarSQL(cdsSeteoF, string.Format("Select Codigo From Articulo Where idArticulo = {0}", idArticulo15));
				this.uGridGastos.ActiveRow.Cells["Articulo"].Value = Funcion.sEscalarSQL(cdsSeteoF, string.Format("Select Articulo From Articulo Where idArticulo = {0}", idArticulo15));
				this.uGridGastos.ActiveRow.Cells["IVA"].Value = dIVA15;
				this.uGridGastos.ActiveRow.Cells["Subtotal"].Value = dSubtotal15;								
			}
			this.uGridGastos.Rows.Band.AddNew();
			#endregion Asigna Valores al Grid

			Totales();
		}

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);	
		}

		private void Totales()
		{
			double dSubtotalIVA0 = 0;
			double dSubtotalIVA = 0;
			double dIVA = 0;
			double dTotal = 0;

			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridGastos.Rows.All)
			{
				if (dr.Cells["IVA"].Value != System.DBNull.Value)
				{
					if (dr.Cells["Subtotal"].Value != System.DBNull.Value)
					{
						if ((double)dr.Cells["IVA"].Value > 0)
						{
							dSubtotalIVA = dSubtotalIVA + (double) dr.Cells["Subtotal"].Value;					
							dIVA = dIVA + ((double) dr.Cells["Subtotal"].Value * (double) dr.Cells["IVA"].Value)/100;
						}

						if ((double)dr.Cells["IVA"].Value == 0)
						{
							dSubtotalIVA0 = dSubtotalIVA0 + (double) dr.Cells["Subtotal"].Value;
						}
					}
				}
			}

			//dIVA = (dSubtotalIVA * 1.14) / 100;
			dTotal = dSubtotalIVA + dSubtotalIVA0 + dIVA;

			this.txtSubtotalIVA.Value = dSubtotalIVA;	
			this.txtSubtotalIVA0.Value = dSubtotalIVA0;			
			this.txtIVA.Value = dIVA;
			this.txtTotal.Value = dTotal;
		}

		private void uGridGastos_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
			#region Formato numerico
			CultureInfo cultureENUS = CultureInfo.CreateSpecificCulture("ja-jp");
			string stMoneda = "#,##0.";
			string stMascara = "nnn,nnn,nnn.";
			string stCeros = "0000000000";
			string stNumero = "nnnnnnnnnn";
			string stFormato = "";
			string stInput = "";
			#endregion Formato numerico

			#region Decimales para el total
			stFormato = stMoneda + stCeros.Substring(0, 2);
			stInput = stMascara + stNumero.Substring(0, 2);
			e.Layout.Bands[0].Columns["IVA"].Format = stFormato; 
			e.Layout.Bands[0].Columns["IVA"].MaskInput = stInput;
			e.Layout.Bands[0].Columns["Subtotal"].Format = stFormato; 
			e.Layout.Bands[0].Columns["Subtotal"].MaskInput = stInput;			
			#endregion Decimales para el total
		}

		private void uGridGastos_AfterRowInsert(object sender, Infragistics.Win.UltraWinGrid.RowEventArgs e)
		{
			e.Row.Cells["Posicion"].Value = this.uGridGastos.Rows.Count;
			e.Row.Cells["IVA"].Value = 0;
			e.Row.Cells["Subtotal"].Value = 0;
		}

		private void uGridGastos_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (this.uGridGastos.ActiveCell == null) return;
			try
			{
				switch(e.KeyValue)
				{
						#region F3
					case (int) Keys.F3:
						string sArticulo = this.uGridGastos.ActiveRow.Cells["Codigo"].Value.ToString();						
													
						#region Servicios
						if ((this.uGridGastos.ActiveCell.Column.Key.ToString().ToUpper() == "CODIGO" || this.uGridGastos.ActiveCell.Column.Key.ToString().ToUpper() == "ARTICULO"))
						{							
							using (BuscaBienesServicios miBusqueda = new BuscaBienesServicios(sArticulo, 4))
							{
								if (DialogResult.OK == miBusqueda.ShowDialog())
								{									
									this.uGridGastos.ActiveRow.Cells["Codigo"].Value = miBusqueda.ultraGrid1.ActiveRow.Cells["Codigo"].Value;
								}
							}
						}									
						#endregion Servicios						
											
						break;
						#endregion F3						
						#region Flechas
					case (int) Keys.Up:
						if (uGridGastos.ActiveCell.DroppedDown) return;

						uGridGastos.PerformAction(UltraGridAction.ExitEditMode, false, false);
						uGridGastos.PerformAction(UltraGridAction.AboveCell, false, false);
						e.Handled = true;
						uGridGastos.PerformAction(UltraGridAction.EnterEditMode, false, false);
						break;
					case (int) Keys.Left:
						uGridGastos.PerformAction(UltraGridAction.ExitEditMode, false, false);
						uGridGastos.PerformAction(UltraGridAction.PrevCellByTab, false, false);
						e.Handled = true;
						uGridGastos.PerformAction(UltraGridAction.EnterEditMode, false, false);
						break;
					case (int) Keys.Down:
						if (uGridGastos.ActiveCell.DroppedDown) return;
			
						uGridGastos.PerformAction(UltraGridAction.ExitEditMode, false, false);
						uGridGastos.PerformAction(UltraGridAction.BelowCell, false, false);
						e.Handled = true;
						uGridGastos.PerformAction(UltraGridAction.EnterEditMode, false, false);
						break;
						#endregion Flechas
						#region Enter
					case (int) Keys.Enter:
						uGridGastos.PerformAction(UltraGridAction.ExitEditMode, false, false);
						uGridGastos.PerformAction(UltraGridAction.NextCellByTab, false, false);
						e.Handled = true;
						uGridGastos.PerformAction(UltraGridAction.EnterEditMode, false, false);
						if(uGridGastos.ActiveCell.DroppedDown == false)
							uGridGastos.ActiveCell.DroppedDown = true;
						//						uGridGastos.PerformAction(UltraGridAction.EnterEditModeAndDropdown, false, false);
						break;
						#endregion Enter
				}
			}
			catch(System.Exception ex)
			{
				MessageBox.Show(ex.Message, "Error en el grid.");
			}
		}

		private void uGridGastos_AfterCellUpdate(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			if (e.Cell.Column.ToString() == "Codigo")
			{
				if (Funcion.iEscalarSQL(cdsSeteoF, string.Format("Select Count(*) From Articulo Where (Codigo = '{0}' Or Articulo = '{0}')", e.Cell.Row.Cells["Codigo"].Value.ToString())) == 1)
				{
					string sSQL = string.Format("Select idArticulo, Codigo, Articulo, Iva From Articulo Where (Codigo = '{0}' Or Articulo = '{0}')", e.Cell.Row.Cells["Codigo"].Value.ToString());
					SqlDataReader dr = Funcion.rEscalarSQL(cdsSeteoF, sSQL, true);
					dr.Read();
					if (dr.HasRows)
					{
						e.Cell.Row.Cells["idArticulo"].Value = dr.GetInt32(0);						
						e.Cell.Row.Cells["Articulo"].Value = dr.GetString(2);
						//if (BSinFactura) e.Cell.Row.Cells["IVA"].Value = dr.GetDouble(3);
					}
					dr.Close();
				}
				else
				{													
					#region Servicios
					using (BuscaBienesServicios miBusqueda = new BuscaBienesServicios(this.uGridGastos.ActiveRow.Cells["Codigo"].Value.ToString(), 4))
					{
						if (DialogResult.OK == miBusqueda.ShowDialog())
						{									
							this.uGridGastos.ActiveRow.Cells["Codigo"].Value = miBusqueda.ultraGrid1.ActiveRow.Cells["Codigo"].Value;
						}
					}								
					#endregion Servicios
				}
			}

			if (e.Cell.Column.ToString() == "IVA")
			{
				Totales();
			}

			if (e.Cell.Column.ToString() == "Subtotal")
			{
				Totales();
			}
		}

		private void btnAceptar_Click(object sender, System.EventArgs e)
		{
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridGastos.Rows.All)
			{
				if (dr.Cells["idArticulo"].Value == System.DBNull.Value || dr.Cells["Codigo"].Value == System.DBNull.Value || dr.Cells["Articulo"].Value == System.DBNull.Value)
				{
					bEliminaAlValidar = true;
					dr.Delete();
				}
			}

			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridGastos.Rows.All)
			{
				if ((double)dr.Cells["Subtotal"].Value == 0)
				{
					MessageBox.Show("Escriba el Valor", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					this.uGridGastos.ActiveRow = dr;
					this.uGridGastos.ActiveRow.Selected = true;
					return;
				}
			}

			this.DialogResult = DialogResult.OK;
		}

		private void uGridGastos_BeforeRowsDeleted(object sender, Infragistics.Win.UltraWinGrid.BeforeRowsDeletedEventArgs e)
		{
			if (!bEliminaAlValidar)
			{
				for (int i = 0; i < e.Rows.Length; i++)
				{	
					if (DialogResult.No ==	MessageBox.Show(string.Format("¿Esta seguro de Borrar el Item '{0}' - '{1}'?", e.Rows[i].Cells["Codigo"].Value.ToString(), e.Rows[i].Cells["Articulo"].Value.ToString()), "Point Technology", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2)) e.Cancel = true;				
					else
					{
						e.DisplayPromptMsg = false;					
					}
				}				
			}
			else e.DisplayPromptMsg = false;
		}

		private void uGridGastos_AfterRowsDeleted(object sender, System.EventArgs e)
		{
			int iCont = 1;	
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridGastos.Rows.All)
			{
				dr.Cells["Posicion"].Value = iCont;
				iCont++;
			}

			Totales();
		}

		private void uGridGastos_BeforeRowInsert(object sender, Infragistics.Win.UltraWinGrid.BeforeRowInsertEventArgs e)
		{
			int iContItems = this.uGridGastos.Rows.Count;
			
			if (iContItems > 14) 
			{
				e.Cancel = true;
				MessageBox.Show("No Puede Ingresar mas de 15 Items", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop); 
			}
		}

		private void uGridGastos_BeforeCellUpdate(object sender, Infragistics.Win.UltraWinGrid.BeforeCellUpdateEventArgs e)
		{
			if (e.Cell.Column.ToString() == "IVA")
			{
				if (BSinFactura)
				{
					e.Cancel = true;
				}
			}
		}
	}
}

