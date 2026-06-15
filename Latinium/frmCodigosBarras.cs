using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using DevExpress.XtraEditors;
using Infragistics.Win.UltraWinGrid;

namespace Latinium
{
	/// <summary>
	/// Summary description for frmCodigosBarras.
	/// </summary>
	public class frmCodigosBarras : DevExpress.XtraEditors.XtraForm
	{
		public Infragistics.Win.UltraWinGrid.UltraGrid ultraGrid1;
		private C1.Data.C1DataSet cdsSeteoF;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource1;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource2;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource3;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource4;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource5;
		private Infragistics.Win.UltraWinGrid.UltraDropDown cmbCliente5;
		private Infragistics.Win.UltraWinGrid.UltraDropDown cmbCliente4;
		private Infragistics.Win.UltraWinGrid.UltraDropDown cmbCliente3;
		private Infragistics.Win.UltraWinGrid.UltraDropDown cmbCliente2;
		private Infragistics.Win.UltraWinGrid.UltraDropDown cmbCliente1;
		private C1.Data.C1DataSet cdsCodigosBarras;
		private DevExpress.XtraEditors.SimpleButton btCancelar;
		private C1.Data.C1DataView cdvArticulo;
		private C1.Data.C1DataView cdvCodArt;
		private C1.Data.C1DataSet cdsArticulo;
		private System.Windows.Forms.Label lblVendedor;
		private DevExpress.XtraEditors.SimpleButton btBuscar;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtArticulo;
		private System.Windows.Forms.Label label1;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource6;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource7;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbSubGrupo;
		private DevExpress.XtraEditors.SimpleButton btAsignacion;
		private DevExpress.XtraEditors.SimpleButton btGuardar;
		private System.Windows.Forms.GroupBox groupBox1;
		/// <summary>
		/// Required designer variable.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public frmCodigosBarras()
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
			Infragistics.Win.Appearance appearance1 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCodigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idArticulo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Articulo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo1");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo2");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo3");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCliente1", -1, "cmbCliente1");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("NCodigo1");
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCliente2", -1, "cmbCliente2");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn11 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("NCodigo2");
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn12 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCliente3", -1, "cmbCliente3");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn13 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("NCodigo3");
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn14 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCliente4", -1, "cmbCliente4");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn15 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("NCodigo4");
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn16 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCliente5", -1, "cmbCliente5");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn17 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("NCodigo5");
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance11 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance12 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance13 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCodigo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idArticulo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn3 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Codigo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn4 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Articulo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn5 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Codigo1");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn6 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Codigo2");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn7 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Codigo3");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn8 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCliente1");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn9 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("NCodigo1");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn10 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCliente2");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn11 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("NCodigo2");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn12 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCliente3");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn13 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("NCodigo3");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn14 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCliente4");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn15 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("NCodigo4");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn16 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCliente5");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn17 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("NCodigo5");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn18 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCliente");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn19 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn18 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCliente");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn19 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand3 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn20 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCliente");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn21 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn20 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCliente");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn21 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand4 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn22 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCliente");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn23 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn22 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCliente");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn23 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand5 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn24 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCliente");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn25 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn24 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCliente");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn25 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand6 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn26 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCliente");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn27 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn26 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCliente");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn27 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.Appearance appearance14 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn28 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idSubGrupo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn29 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("SubGrupo");
			Infragistics.Win.Appearance appearance15 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand7 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn28 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idSubGrupo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn29 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("SubGrupo");
			this.ultraGrid1 = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource6 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.cdsCodigosBarras = new C1.Data.C1DataSet();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.cmbCliente5 = new Infragistics.Win.UltraWinGrid.UltraDropDown();
			this.ultraDataSource5 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.cmbCliente4 = new Infragistics.Win.UltraWinGrid.UltraDropDown();
			this.ultraDataSource4 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.cmbCliente3 = new Infragistics.Win.UltraWinGrid.UltraDropDown();
			this.ultraDataSource3 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.cmbCliente2 = new Infragistics.Win.UltraWinGrid.UltraDropDown();
			this.ultraDataSource2 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.cmbCliente1 = new Infragistics.Win.UltraWinGrid.UltraDropDown();
			this.ultraDataSource1 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.btCancelar = new DevExpress.XtraEditors.SimpleButton();
			this.cdvCodArt = new C1.Data.C1DataView();
			this.cdsArticulo = new C1.Data.C1DataSet();
			this.cdvArticulo = new C1.Data.C1DataView();
			this.lblVendedor = new System.Windows.Forms.Label();
			this.btBuscar = new DevExpress.XtraEditors.SimpleButton();
			this.txtArticulo = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label1 = new System.Windows.Forms.Label();
			this.ultraDataSource7 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.cmbSubGrupo = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.btAsignacion = new DevExpress.XtraEditors.SimpleButton();
			this.btGuardar = new DevExpress.XtraEditors.SimpleButton();
			this.groupBox1 = new System.Windows.Forms.GroupBox();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource6)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsCodigosBarras)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCliente5)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource5)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCliente4)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource4)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCliente3)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource3)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCliente2)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCliente1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvCodArt)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsArticulo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvArticulo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtArticulo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource7)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbSubGrupo)).BeginInit();
			this.SuspendLayout();
			// 
			// ultraGrid1
			// 
			this.ultraGrid1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.ultraGrid1.Cursor = System.Windows.Forms.Cursors.Default;
			this.ultraGrid1.DataSource = this.ultraDataSource6;
			appearance1.BackColor = System.Drawing.Color.White;
			this.ultraGrid1.DisplayLayout.Appearance = appearance1;
			ultraGridBand1.AddButtonCaption = "CodigosDeBarras";
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn1.Width = 36;
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridColumn2.Hidden = true;
			ultraGridColumn2.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.Edit;
			ultraGridColumn2.Width = 35;
			ultraGridColumn3.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn3.Header.VisiblePosition = 2;
			ultraGridColumn3.Width = 180;
			ultraGridColumn4.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn4.Header.VisiblePosition = 3;
			ultraGridColumn4.Width = 220;
			ultraGridColumn5.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn5.Header.Caption = "Código 1";
			ultraGridColumn5.Header.VisiblePosition = 4;
			ultraGridColumn5.Width = 120;
			ultraGridColumn6.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn6.Header.Caption = "Código 2";
			ultraGridColumn6.Header.VisiblePosition = 5;
			ultraGridColumn6.Width = 120;
			ultraGridColumn7.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn7.Header.Caption = "Código 3";
			ultraGridColumn7.Header.VisiblePosition = 6;
			ultraGridColumn7.Width = 120;
			ultraGridColumn8.Header.Caption = "Cliente 1";
			ultraGridColumn8.Header.VisiblePosition = 7;
			ultraGridColumn8.Width = 150;
			appearance2.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn9.CellAppearance = appearance2;
			ultraGridColumn9.Header.Caption = "...";
			ultraGridColumn9.Header.VisiblePosition = 8;
			ultraGridColumn9.PromptChar = ' ';
			ultraGridColumn9.Width = 30;
			ultraGridColumn10.Header.Caption = "Cliente 2";
			ultraGridColumn10.Header.VisiblePosition = 9;
			ultraGridColumn10.Width = 150;
			appearance3.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn11.CellAppearance = appearance3;
			ultraGridColumn11.Header.Caption = "...";
			ultraGridColumn11.Header.VisiblePosition = 10;
			ultraGridColumn11.PromptChar = ' ';
			ultraGridColumn11.Width = 30;
			ultraGridColumn12.Header.Caption = "Cliente 3";
			ultraGridColumn12.Header.VisiblePosition = 11;
			ultraGridColumn12.Width = 150;
			appearance4.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn13.CellAppearance = appearance4;
			ultraGridColumn13.Header.Caption = "...";
			ultraGridColumn13.Header.VisiblePosition = 12;
			ultraGridColumn13.PromptChar = ' ';
			ultraGridColumn13.Width = 30;
			ultraGridColumn14.Header.Caption = "Cliente 4";
			ultraGridColumn14.Header.VisiblePosition = 13;
			ultraGridColumn14.Width = 150;
			appearance5.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn15.CellAppearance = appearance5;
			ultraGridColumn15.Header.Caption = "...";
			ultraGridColumn15.Header.VisiblePosition = 14;
			ultraGridColumn15.PromptChar = ' ';
			ultraGridColumn15.Width = 30;
			ultraGridColumn16.Header.Caption = "Cliente 5";
			ultraGridColumn16.Header.VisiblePosition = 15;
			ultraGridColumn16.Width = 150;
			appearance6.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn17.CellAppearance = appearance6;
			ultraGridColumn17.Header.Caption = "...";
			ultraGridColumn17.Header.VisiblePosition = 16;
			ultraGridColumn17.PromptChar = ' ';
			ultraGridColumn17.Width = 30;
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
																										 ultraGridColumn17});
			appearance7.TextHAlign = Infragistics.Win.HAlign.Left;
			ultraGridBand1.Header.Appearance = appearance7;
			ultraGridBand1.Header.Caption = "";
			ultraGridBand1.HeaderVisible = true;
			this.ultraGrid1.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			appearance8.ForeColor = System.Drawing.Color.Black;
			appearance8.ForeColorDisabled = System.Drawing.Color.Black;
			this.ultraGrid1.DisplayLayout.Override.ActiveRowAppearance = appearance8;
			this.ultraGrid1.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.TemplateOnBottom;
			this.ultraGrid1.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			appearance9.BackColor = System.Drawing.Color.Transparent;
			this.ultraGrid1.DisplayLayout.Override.CardAreaAppearance = appearance9;
			appearance10.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance10.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance10.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance10.FontData.BoldAsString = "True";
			appearance10.FontData.Name = "Arial";
			appearance10.FontData.SizeInPoints = 8F;
			appearance10.ForeColor = System.Drawing.Color.White;
			appearance10.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.ultraGrid1.DisplayLayout.Override.HeaderAppearance = appearance10;
			this.ultraGrid1.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortSingle;
			appearance11.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance11.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance11.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.RowAlternateAppearance = appearance11;
			appearance12.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance12.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance12.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.RowSelectorAppearance = appearance12;
			appearance13.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance13.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance13.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.SelectedRowAppearance = appearance13;
			this.ultraGrid1.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.ultraGrid1.Location = new System.Drawing.Point(8, 56);
			this.ultraGrid1.Name = "ultraGrid1";
			this.ultraGrid1.Size = new System.Drawing.Size(1144, 328);
			this.ultraGrid1.TabIndex = 154;
			this.ultraGrid1.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.ultraGrid1_KeyPress);
			this.ultraGrid1.KeyDown += new System.Windows.Forms.KeyEventHandler(this.ultraGrid1_KeyDown);
			this.ultraGrid1.AfterCellUpdate += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.ultraGrid1_AfterCellUpdate);
			this.ultraGrid1.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.ultraGrid1_InitializeLayout);
			// 
			// ultraDataSource6
			// 
			ultraDataColumn1.DataType = typeof(int);
			ultraDataColumn2.DataType = typeof(int);
			ultraDataColumn8.DataType = typeof(int);
			ultraDataColumn9.DataType = typeof(int);
			ultraDataColumn10.DataType = typeof(int);
			ultraDataColumn11.DataType = typeof(int);
			ultraDataColumn12.DataType = typeof(int);
			ultraDataColumn13.DataType = typeof(int);
			ultraDataColumn14.DataType = typeof(int);
			ultraDataColumn15.DataType = typeof(int);
			ultraDataColumn16.DataType = typeof(int);
			ultraDataColumn17.DataType = typeof(int);
			this.ultraDataSource6.Band.Columns.AddRange(new object[] {
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
			// cdsCodigosBarras
			// 
			this.cdsCodigosBarras.BindingContextCtrl = this;
			this.cdsCodigosBarras.DataLibrary = "LibFacturacion";
			this.cdsCodigosBarras.DataLibraryUrl = "";
			this.cdsCodigosBarras.DataSetDef = "dsCodigoBarras";
			this.cdsCodigosBarras.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsCodigosBarras.Name = "cdsCodigosBarras";
			this.cdsCodigosBarras.SchemaClassName = "LibFacturacion.DataClass";
			this.cdsCodigosBarras.SchemaDef = null;
			this.cdsCodigosBarras.BeforeFill += new C1.Data.FillEventHandler(this.cdsCodigosBarras_BeforeFill);
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
			// cmbCliente5
			// 
			this.cmbCliente5.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbCliente5.DataSource = this.ultraDataSource5;
			ultraGridColumn18.Header.VisiblePosition = 0;
			ultraGridColumn18.Hidden = true;
			ultraGridColumn19.Header.VisiblePosition = 1;
			ultraGridColumn19.Width = 300;
			ultraGridBand2.Columns.AddRange(new object[] {
																										 ultraGridColumn18,
																										 ultraGridColumn19});
			this.cmbCliente5.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			this.cmbCliente5.DisplayMember = "Nombre";
			this.cmbCliente5.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbCliente5.Location = new System.Drawing.Point(680, 224);
			this.cmbCliente5.Name = "cmbCliente5";
			this.cmbCliente5.Size = new System.Drawing.Size(320, 88);
			this.cmbCliente5.TabIndex = 163;
			this.cmbCliente5.ValueMember = "idCliente";
			this.cmbCliente5.Visible = false;
			// 
			// ultraDataSource5
			// 
			ultraDataColumn18.DataType = typeof(int);
			this.ultraDataSource5.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn18,
																																 ultraDataColumn19});
			// 
			// cmbCliente4
			// 
			this.cmbCliente4.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbCliente4.DataSource = this.ultraDataSource4;
			ultraGridColumn20.Header.VisiblePosition = 0;
			ultraGridColumn20.Hidden = true;
			ultraGridColumn21.Header.VisiblePosition = 1;
			ultraGridColumn21.Width = 300;
			ultraGridBand3.Columns.AddRange(new object[] {
																										 ultraGridColumn20,
																										 ultraGridColumn21});
			this.cmbCliente4.DisplayLayout.BandsSerializer.Add(ultraGridBand3);
			this.cmbCliente4.DisplayMember = "Nombre";
			this.cmbCliente4.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbCliente4.Location = new System.Drawing.Point(352, 224);
			this.cmbCliente4.Name = "cmbCliente4";
			this.cmbCliente4.Size = new System.Drawing.Size(320, 88);
			this.cmbCliente4.TabIndex = 162;
			this.cmbCliente4.ValueMember = "idCliente";
			this.cmbCliente4.Visible = false;
			// 
			// ultraDataSource4
			// 
			ultraDataColumn20.DataType = typeof(int);
			this.ultraDataSource4.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn20,
																																 ultraDataColumn21});
			// 
			// cmbCliente3
			// 
			this.cmbCliente3.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbCliente3.DataSource = this.ultraDataSource3;
			ultraGridColumn22.Header.VisiblePosition = 0;
			ultraGridColumn22.Hidden = true;
			ultraGridColumn23.Header.VisiblePosition = 1;
			ultraGridColumn23.Width = 300;
			ultraGridBand4.Columns.AddRange(new object[] {
																										 ultraGridColumn22,
																										 ultraGridColumn23});
			this.cmbCliente3.DisplayLayout.BandsSerializer.Add(ultraGridBand4);
			this.cmbCliente3.DisplayMember = "Nombre";
			this.cmbCliente3.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbCliente3.Location = new System.Drawing.Point(24, 224);
			this.cmbCliente3.Name = "cmbCliente3";
			this.cmbCliente3.Size = new System.Drawing.Size(320, 88);
			this.cmbCliente3.TabIndex = 161;
			this.cmbCliente3.ValueMember = "idCliente";
			this.cmbCliente3.Visible = false;
			// 
			// ultraDataSource3
			// 
			ultraDataColumn22.DataType = typeof(int);
			this.ultraDataSource3.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn22,
																																 ultraDataColumn23});
			// 
			// cmbCliente2
			// 
			this.cmbCliente2.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbCliente2.DataSource = this.ultraDataSource2;
			ultraGridColumn24.Header.VisiblePosition = 0;
			ultraGridColumn24.Hidden = true;
			ultraGridColumn25.Header.VisiblePosition = 1;
			ultraGridColumn25.Width = 300;
			ultraGridBand5.Columns.AddRange(new object[] {
																										 ultraGridColumn24,
																										 ultraGridColumn25});
			this.cmbCliente2.DisplayLayout.BandsSerializer.Add(ultraGridBand5);
			this.cmbCliente2.DisplayMember = "Nombre";
			this.cmbCliente2.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbCliente2.Location = new System.Drawing.Point(352, 128);
			this.cmbCliente2.Name = "cmbCliente2";
			this.cmbCliente2.Size = new System.Drawing.Size(320, 88);
			this.cmbCliente2.TabIndex = 160;
			this.cmbCliente2.ValueMember = "idCliente";
			this.cmbCliente2.Visible = false;
			this.cmbCliente2.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.cmbArtAlterno3_InitializeLayout);
			// 
			// ultraDataSource2
			// 
			ultraDataColumn24.DataType = typeof(int);
			this.ultraDataSource2.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn24,
																																 ultraDataColumn25});
			// 
			// cmbCliente1
			// 
			this.cmbCliente1.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbCliente1.DataSource = this.ultraDataSource1;
			ultraGridBand6.AddButtonCaption = "SeteoF";
			ultraGridColumn26.Header.VisiblePosition = 1;
			ultraGridColumn26.Hidden = true;
			ultraGridColumn26.Width = 250;
			ultraGridColumn27.Header.VisiblePosition = 0;
			ultraGridColumn27.Width = 300;
			ultraGridBand6.Columns.AddRange(new object[] {
																										 ultraGridColumn26,
																										 ultraGridColumn27});
			this.cmbCliente1.DisplayLayout.BandsSerializer.Add(ultraGridBand6);
			this.cmbCliente1.DisplayMember = "Nombre";
			this.cmbCliente1.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbCliente1.Location = new System.Drawing.Point(24, 128);
			this.cmbCliente1.Name = "cmbCliente1";
			this.cmbCliente1.Size = new System.Drawing.Size(320, 88);
			this.cmbCliente1.TabIndex = 159;
			this.cmbCliente1.ValueMember = "idCliente";
			this.cmbCliente1.Visible = false;
			// 
			// ultraDataSource1
			// 
			ultraDataColumn26.DataType = typeof(int);
			this.ultraDataSource1.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn26,
																																 ultraDataColumn27});
			// 
			// btCancelar
			// 
			this.btCancelar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btCancelar.DialogResult = System.Windows.Forms.DialogResult.Cancel;
			this.btCancelar.Location = new System.Drawing.Point(104, 392);
			this.btCancelar.Name = "btCancelar";
			this.btCancelar.Size = new System.Drawing.Size(88, 24);
			this.btCancelar.TabIndex = 165;
			this.btCancelar.Text = "&Salir";
			this.btCancelar.Click += new System.EventHandler(this.btCancelar_Click);
			// 
			// cdvCodArt
			// 
			this.cdvCodArt.BindingContextCtrl = this;
			this.cdvCodArt.DataSet = this.cdsArticulo;
			this.cdvCodArt.Sort = "Codigo";
			this.cdvCodArt.TableName = "";
			this.cdvCodArt.TableViewName = "Articulo";
			// 
			// cdsArticulo
			// 
			this.cdsArticulo.BindingContextCtrl = this;
			this.cdsArticulo.CaseSensitive = false;
			this.cdsArticulo.DataLibrary = "LibFacturacion";
			this.cdsArticulo.DataLibraryUrl = "";
			this.cdsArticulo.DataSetDef = "dsArticulos";
			this.cdsArticulo.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsArticulo.Name = "cdsArticulo";
			this.cdsArticulo.SchemaClassName = "LibFacturacion.DataClass";
			this.cdsArticulo.SchemaDef = null;
			this.cdsArticulo.BeforeFill += new C1.Data.FillEventHandler(this.cdsArticulo_BeforeFill);
			// 
			// cdvArticulo
			// 
			this.cdvArticulo.BindingContextCtrl = this;
			this.cdvArticulo.DataSet = this.cdsArticulo;
			this.cdvArticulo.Sort = "Articulo";
			this.cdvArticulo.TableName = "";
			this.cdvArticulo.TableViewName = "Articulo";
			// 
			// lblVendedor
			// 
			this.lblVendedor.AutoSize = true;
			this.lblVendedor.Location = new System.Drawing.Point(344, 11);
			this.lblVendedor.Name = "lblVendedor";
			this.lblVendedor.Size = new System.Drawing.Size(57, 16);
			this.lblVendedor.TabIndex = 206;
			this.lblVendedor.Text = "Subgrupos";
			this.lblVendedor.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// btBuscar
			// 
			this.btBuscar.Location = new System.Drawing.Point(760, 8);
			this.btBuscar.Name = "btBuscar";
			this.btBuscar.Size = new System.Drawing.Size(80, 24);
			this.btBuscar.TabIndex = 208;
			this.btBuscar.Text = "&Buscar";
			this.btBuscar.Click += new System.EventHandler(this.btBuscar_Click);
			// 
			// txtArticulo
			// 
			appearance14.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtArticulo.Appearance = appearance14;
			this.txtArticulo.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtArticulo.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtArticulo.Location = new System.Drawing.Point(72, 8);
			this.txtArticulo.Name = "txtArticulo";
			this.txtArticulo.Size = new System.Drawing.Size(232, 22);
			this.txtArticulo.TabIndex = 207;
			this.txtArticulo.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtArticulo_KeyPress);
			// 
			// label1
			// 
			this.label1.AutoSize = true;
			this.label1.Location = new System.Drawing.Point(8, 11);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(42, 16);
			this.label1.TabIndex = 209;
			this.label1.Text = "Artículo";
			this.label1.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// ultraDataSource7
			// 
			ultraDataColumn28.DataType = typeof(int);
			this.ultraDataSource7.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn28,
																																 ultraDataColumn29});
			// 
			// cmbSubGrupo
			// 
			appearance15.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbSubGrupo.Appearance = appearance15;
			this.cmbSubGrupo.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbSubGrupo.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbSubGrupo.DataSource = this.ultraDataSource7;
			ultraGridColumn28.Header.VisiblePosition = 0;
			ultraGridColumn28.Hidden = true;
			ultraGridColumn29.Header.VisiblePosition = 1;
			ultraGridColumn29.Width = 320;
			ultraGridBand7.Columns.AddRange(new object[] {
																										 ultraGridColumn28,
																										 ultraGridColumn29});
			this.cmbSubGrupo.DisplayLayout.BandsSerializer.Add(ultraGridBand7);
			this.cmbSubGrupo.DisplayMember = "SubGrupo";
			this.cmbSubGrupo.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbSubGrupo.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbSubGrupo.Location = new System.Drawing.Point(424, 8);
			this.cmbSubGrupo.Name = "cmbSubGrupo";
			this.cmbSubGrupo.Size = new System.Drawing.Size(312, 21);
			this.cmbSubGrupo.TabIndex = 210;
			this.cmbSubGrupo.ValueMember = "idSubGrupo";
			// 
			// btAsignacion
			// 
			this.btAsignacion.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btAsignacion.Location = new System.Drawing.Point(200, 392);
			this.btAsignacion.Name = "btAsignacion";
			this.btAsignacion.Size = new System.Drawing.Size(88, 24);
			this.btAsignacion.TabIndex = 212;
			this.btAsignacion.Text = "&Borrar Clientes";
			this.btAsignacion.Click += new System.EventHandler(this.btAsignacion_Click);
			// 
			// btGuardar
			// 
			this.btGuardar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btGuardar.Location = new System.Drawing.Point(8, 392);
			this.btGuardar.Name = "btGuardar";
			this.btGuardar.Size = new System.Drawing.Size(88, 24);
			this.btGuardar.TabIndex = 213;
			this.btGuardar.Text = "&Guardar";
			this.btGuardar.Click += new System.EventHandler(this.btGuardar_Click);
			// 
			// groupBox1
			// 
			this.groupBox1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.groupBox1.Location = new System.Drawing.Point(0, 40);
			this.groupBox1.Name = "groupBox1";
			this.groupBox1.Size = new System.Drawing.Size(1264, 8);
			this.groupBox1.TabIndex = 214;
			this.groupBox1.TabStop = false;
			// 
			// frmCodigosBarras
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 14);
			this.ClientSize = new System.Drawing.Size(1160, 422);
			this.Controls.Add(this.groupBox1);
			this.Controls.Add(this.btGuardar);
			this.Controls.Add(this.btAsignacion);
			this.Controls.Add(this.cmbSubGrupo);
			this.Controls.Add(this.label1);
			this.Controls.Add(this.btBuscar);
			this.Controls.Add(this.txtArticulo);
			this.Controls.Add(this.lblVendedor);
			this.Controls.Add(this.btCancelar);
			this.Controls.Add(this.cmbCliente5);
			this.Controls.Add(this.cmbCliente4);
			this.Controls.Add(this.cmbCliente3);
			this.Controls.Add(this.cmbCliente2);
			this.Controls.Add(this.cmbCliente1);
			this.Controls.Add(this.ultraGrid1);
			this.KeyPreview = true;
			this.Name = "frmCodigosBarras";
			this.Text = "Códigos de Barras";
			this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.frmCodigosBarras_KeyDown);
			this.Load += new System.EventHandler(this.frmCodigosBarras_Load);
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource6)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsCodigosBarras)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCliente5)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource5)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCliente4)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource4)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCliente3)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource3)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCliente2)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCliente1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvCodArt)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsArticulo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvArticulo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtArticulo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource7)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbSubGrupo)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void cmbArtAlterno3_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);	
		}

		private void cdsCodigosBarras_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsCodigosBarras.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);
		}

		private void btCancelar_Click(object sender, System.EventArgs e)
		{
			this.Close();
		}

		private void Consulta()
		{
			int IdSubGrupo = 0;
			if (this.cmbSubGrupo.Text.ToString().Length > 0) IdSubGrupo = (int) this.cmbSubGrupo.Value;

			string sSQL = string.Format("Exec CodigosDeBarrasConsulta {0}, '{1}'", IdSubGrupo, this.txtArticulo.Text.ToString());
			this.ultraGrid1.DataSource = Funcion.dvProcedimiento(cdsSeteoF, sSQL);

			this.ultraGrid1.DisplayLayout.Bands[0].Header.Caption = this.ultraGrid1.Rows.Count.ToString() + " registros encontrados.";
		}

		private void ultraGrid1_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
//			int ICodigo = 0;
//			if (this.ultraGrid1.ActiveRow.Cells["Codigo1"].Value.ToString().Length > 0 && this.ultraGrid1.ActiveRow.Cells["Codigo2"].Value.ToString().Length > 0 && this.ultraGrid1.ActiveRow.Cells["Codigo3"].Value.ToString().Length > 0) ICodigo = 3;
//			if (this.ultraGrid1.ActiveRow.Cells["Codigo1"].Value.ToString().Length > 0 && this.ultraGrid1.ActiveRow.Cells["Codigo2"].Value.ToString().Length > 0 && this.ultraGrid1.ActiveRow.Cells["Codigo3"].Value.ToString().Length == 0) ICodigo = 2;
//			if (this.ultraGrid1.ActiveRow.Cells["Codigo1"].Value.ToString().Length > 0 && this.ultraGrid1.ActiveRow.Cells["Codigo2"].Value.ToString().Length == 0 && this.ultraGrid1.ActiveRow.Cells["Codigo3"].Value.ToString().Length == 0) ICodigo = 1;
//			if (this.ultraGrid1.ActiveRow.Cells["Codigo1"].Value.ToString().Length == 0 && this.ultraGrid1.ActiveRow.Cells["Codigo2"].Value.ToString().Length == 0 && this.ultraGrid1.ActiveRow.Cells["Codigo3"].Value.ToString().Length == 0) ICodigo = 0;
//				
//      if ((this.ultraGrid1.ActiveCell.Column.Key.ToString().ToUpper() == "NCODIGO1" && (int) this.ultraGrid1.ActiveRow.Cells["NCodigo1"].Value > ICodigo) 
//				|| (this.ultraGrid1.ActiveCell.Column.Key.ToString().ToUpper() == "NCODIGO2" && (int) this.ultraGrid1.ActiveRow.Cells["NCodigo2"].Value > ICodigo)
//				|| (this.ultraGrid1.ActiveCell.Column.Key.ToString().ToUpper() == "NCODIGO3" && (int) this.ultraGrid1.ActiveRow.Cells["NCodigo3"].Value > ICodigo)
//				|| (this.ultraGrid1.ActiveCell.Column.Key.ToString().ToUpper() == "NCODIGO4" && (int) this.ultraGrid1.ActiveRow.Cells["NCodigo4"].Value > ICodigo)
//				|| (this.ultraGrid1.ActiveCell.Column.Key.ToString().ToUpper() == "NCODIGO5" && (int) this.ultraGrid1.ActiveRow.Cells["NCodigo5"].Value > ICodigo)) e.Handled = true;

			if ((this.ultraGrid1.ActiveCell.Column.Key.ToString().ToUpper() == "NCODIGO1" && this.ultraGrid1.ActiveRow.Cells["idCliente1"].Value == System.DBNull.Value)
				|| (this.ultraGrid1.ActiveCell.Column.Key.ToString().ToUpper() == "NCODIGO2" && this.ultraGrid1.ActiveRow.Cells["idCliente2"].Value == System.DBNull.Value)
				|| (this.ultraGrid1.ActiveCell.Column.Key.ToString().ToUpper() == "NCODIGO3" && this.ultraGrid1.ActiveRow.Cells["idCliente3"].Value == System.DBNull.Value)
				|| (this.ultraGrid1.ActiveCell.Column.Key.ToString().ToUpper() == "NCODIGO4" && this.ultraGrid1.ActiveRow.Cells["idCliente4"].Value == System.DBNull.Value)
				|| (this.ultraGrid1.ActiveCell.Column.Key.ToString().ToUpper() == "NCODIGO5" && this.ultraGrid1.ActiveRow.Cells["idCliente5"].Value == System.DBNull.Value)) e.Handled = true;
		}

		private void ultraGrid1_AfterCellUpdate(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			if (e.Cell.IsInEditMode)
			{
				if (e.Cell.Column.ToString() == "idArticulo")
					if (e.Cell.OriginalValue != e.Cell.Value)
					{
						e.Cell.Row.Cells["idArticulo1"].Value = e.Cell.Row.Cells["idArticulo"].Value;
					}
			}
		}

		private void ultraGrid1_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (ultraGrid1.ActiveCell == null) return;
			switch(e.KeyValue)
			{
				#region F3
				case (int) Keys.F3:
											
					string sArticulo = this.ultraGrid1.ActiveRow.Cells["idArticulo"].Value.ToString();
					if (this.ultraGrid1.ActiveCell.Column.Key.ToString().ToUpper() == "IDARTICULO")
					{
						using (BuscaArticulos miBusqueda = new BuscaArticulos(sArticulo, 1, 7))
						{
							if (DialogResult.OK == miBusqueda.ShowDialog())
							{
								ultraGrid1.ActiveRow.Cells["idArticulo"].Value = miBusqueda.idArticulo;
							}
						}
					}
					
					break;
					#endregion F3
				#region Flechas
				case (int) Keys.Up:
					if (ultraGrid1.ActiveCell.DroppedDown) return;

					ultraGrid1.PerformAction(UltraGridAction.ExitEditMode, false, false);
					ultraGrid1.PerformAction(UltraGridAction.AboveCell, false, false);
					e.Handled = true;
					ultraGrid1.PerformAction(UltraGridAction.EnterEditMode, false, false);
					break;
				case (int) Keys.Left:
					ultraGrid1.PerformAction(UltraGridAction.ExitEditMode, false, false);
					ultraGrid1.PerformAction(UltraGridAction.PrevCellByTab, false, false);
					e.Handled = true;
					ultraGrid1.PerformAction(UltraGridAction.EnterEditMode, false, false);
					break;
				case (int) Keys.Down:
					if (ultraGrid1.ActiveCell.DroppedDown) return;
			
					ultraGrid1.PerformAction(UltraGridAction.ExitEditMode, false, false);
					ultraGrid1.PerformAction(UltraGridAction.BelowCell, false, false);
					e.Handled = true;
					ultraGrid1.PerformAction(UltraGridAction.EnterEditMode, false, false);
					break;
					#endregion Flechas
				#region Enter
				case (int) Keys.Enter:
						
					ultraGrid1.PerformAction(UltraGridAction.ExitEditMode, false, false);
					ultraGrid1.PerformAction(UltraGridAction.NextCellByTab, false, false);
					e.Handled = true;
					ultraGrid1.PerformAction(UltraGridAction.EnterEditMode, false, false);
					if(ultraGrid1.ActiveCell.DroppedDown == false) ultraGrid1.ActiveCell.DroppedDown = true;
					
					break;
					#endregion Enter
			}
		}

		private void frmCodigosBarras_Load(object sender, System.EventArgs e)
		{
			this.cmbSubGrupo.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Exec CodigosDeBarrasSubGrupos");
			string sSQL = string.Format("Select idCliente, Nombre From Cliente Where Proveedor = 0 And CodigoBarras = 1 Order By Nombre Asc");
			this.cmbCliente1.DataSource = Funcion.dvProcedimiento(cdsSeteoF, sSQL);

			this.cmbCliente2.DataSource = this.cmbCliente1.DataSource;
			this.cmbCliente3.DataSource = this.cmbCliente1.DataSource;
			this.cmbCliente4.DataSource = this.cmbCliente1.DataSource;
			this.cmbCliente5.DataSource = this.cmbCliente1.DataSource;

			Consulta();

			this.cmbSubGrupo.Select();
		}

		private void ultraGrid1_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}

		private void frmCodigosBarras_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Escape) this.Close();
		}

		private void cdsArticulo_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			this.cdsArticulo.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);
		}

		private void btBuscar_Click(object sender, System.EventArgs e)
		{
			Consulta();
		}

		private void cmbSubGrupo_ValueChanged(object sender, System.EventArgs e)
		{
			Consulta();
		}

		private void txtArticulo_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (e.KeyChar == 13) Consulta();
		}

		private void btAsignacion_Click(object sender, System.EventArgs e)
		{
			if (this.ultraGrid1.ActiveRow == null)
			{
				MessageBox.Show("Seleccione una fila.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
				return;
			}

			this.ultraGrid1.ActiveRow.Cells["idCliente1"].Value = System.DBNull.Value;
			this.ultraGrid1.ActiveRow.Cells["idCliente2"].Value = System.DBNull.Value;
			this.ultraGrid1.ActiveRow.Cells["idCliente3"].Value = System.DBNull.Value;
			this.ultraGrid1.ActiveRow.Cells["idCliente4"].Value = System.DBNull.Value;
			this.ultraGrid1.ActiveRow.Cells["idCliente5"].Value = System.DBNull.Value;

			this.ultraGrid1.ActiveRow.Cells["NCodigo1"].Value = System.DBNull.Value;
			this.ultraGrid1.ActiveRow.Cells["NCodigo2"].Value = System.DBNull.Value;
			this.ultraGrid1.ActiveRow.Cells["NCodigo3"].Value = System.DBNull.Value;
			this.ultraGrid1.ActiveRow.Cells["NCodigo4"].Value = System.DBNull.Value;
			this.ultraGrid1.ActiveRow.Cells["NCodigo5"].Value = System.DBNull.Value;
		}

		private void btGuardar_Click(object sender, System.EventArgs e)
		{
			try
			{
				foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in ultraGrid1.Rows.All)
				{
					string sSQL = "";
					int idCodigo = 0;
					int idCliente1 = 0;
					int idCliente2 = 0;
					int idCliente3 = 0;
					int idCliente4 = 0;
					int idCliente5 = 0;
					int NCodigo1 = 0;
					int NCodigo2 = 0;
					int NCodigo3 = 0;
					int NCodigo4 = 0;
					int NCodigo5 = 0;

					if ((int)dr.Cells["idCodigo"].Value > 0) idCodigo = (int)dr.Cells["idCodigo"].Value;
					if (dr.Cells["idCliente1"].Value.ToString().Length > 0) idCliente1 = (int)dr.Cells["idCliente1"].Value;
					if (dr.Cells["NCodigo1"].Value.ToString().Length > 0) NCodigo1 = int.Parse(dr.Cells["NCodigo1"].Value.ToString());
					if (dr.Cells["idCliente2"].Value.ToString().Length > 0) idCliente2 = (int)dr.Cells["idCliente2"].Value;
					if (dr.Cells["NCodigo2"].Value.ToString().Length > 0) NCodigo2 = int.Parse(dr.Cells["NCodigo2"].Value.ToString());
					if (dr.Cells["idCliente3"].Value.ToString().Length > 0) idCliente3 = (int)dr.Cells["idCliente3"].Value;
					if (dr.Cells["NCodigo3"].Value.ToString().Length > 0) NCodigo3 = int.Parse(dr.Cells["NCodigo3"].Value.ToString());
					if (dr.Cells["idCliente4"].Value.ToString().Length > 0) idCliente4 = (int)dr.Cells["idCliente4"].Value;
					if (dr.Cells["NCodigo4"].Value.ToString().Length > 0) NCodigo4 = int.Parse(dr.Cells["NCodigo4"].Value.ToString());
					if (dr.Cells["idCliente5"].Value.ToString().Length > 0) idCliente5 = (int)dr.Cells["idCliente5"].Value;
					if (dr.Cells["NCodigo5"].Value.ToString().Length > 0) NCodigo5 = int.Parse(dr.Cells["NCodigo5"].Value.ToString());					

					sSQL = string.Format("Exec CodigosDeBarrrasGuarda {0}, {1}, '{2}', '{3}', '{4}', {5}, {6}, {7}, {8}, {9}, {10}, {11}, {12}, {13}, {14}", 
						idCodigo, (int)dr.Cells["idArticulo"].Value, 
						dr.Cells["Codigo1"].Value.ToString(), dr.Cells["Codigo2"].Value.ToString(), dr.Cells["Codigo3"].Value.ToString(),	
						idCliente1, NCodigo1,
						idCliente2, NCodigo2,
						idCliente3, NCodigo3,
						idCliente4, NCodigo4,
						idCliente5, NCodigo5);

					Funcion.EjecutaSQL(cdsSeteoF, sSQL, true);					
				}

				Consulta();
			}
			catch(System.Exception ex)
			{
				MessageBox.Show(ex.Message, "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
			}
		}
	}
}

