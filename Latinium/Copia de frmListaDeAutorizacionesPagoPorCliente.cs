using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using DevExpress.XtraEditors;
using Infragistics.Win.UltraWinGrid;
using System.Globalization;

namespace Latinium
{
	/// <summary>
	/// Summary description for frmListaDeAutorizacionesPagoPorCliente.
	/// </summary>
	public class frmListaDeAutorizacionesPagoPorCliente : DevExpress.XtraEditors.XtraForm
	{
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource1;
		private System.Windows.Forms.Label label3;
		private DevExpress.XtraEditors.SimpleButton btBuscar;
		private Infragistics.Win.UltraWinGrid.UltraDropDown cmbFormaPago;
		private System.Windows.Forms.Label label4;
		public Infragistics.Win.UltraWinGrid.UltraDropDown cmbCuenta;
		public Infragistics.Win.UltraWinGrid.UltraCombo cmbCuentas;
		private DevExpress.XtraEditors.SimpleButton btAsignar;
		private DevExpress.XtraEditors.SimpleButton btnAsignar;
		private System.Windows.Forms.Label lblCobro;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtPago;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtBusqueda;
		public Infragistics.Win.UltraWinGrid.UltraGrid ultraGrid1;
		private C1.Data.C1DataSet cdsSeteoF;
		private DevExpress.XtraEditors.SimpleButton btnAceptar;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkSeleccionar;
		private DevExpress.XtraEditors.SimpleButton btnCancelar;
		/// <summary>
		/// Required designer variable.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public frmListaDeAutorizacionesPagoPorCliente()
		{
			//
			// Required for Windows Form Designer support
			//
			InitializeComponent();

			//
			// TODO: Add any constructor code after InitializeComponent call
			//
		}

		int IdTipo = 0;
		int IdCliente = 0;
		DateTime DTDesde = DateTime.Today;
		DateTime DTHasta = DateTime.Today;

		public frmListaDeAutorizacionesPagoPorCliente(int idTipo, int idCliente, DateTime dtDesde, DateTime dtHasta)
		{
			//
			// Required for Windows Form Designer support
			//
			InitializeComponent();
			
			IdTipo = idTipo;
			IdCliente = idCliente;
			DTDesde = dtDesde;
			DTHasta = dtHasta;

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
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idAutorizacionPago");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idTipo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idOrigen");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Numero");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Concepto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Notas");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaPago");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idFormaPago", -1, "cmbFormaPago");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCuenta", -1, "cmbCuenta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Autorizado");
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn11 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Valor");
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn12 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("NuevoPago");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn13 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("NumPago");
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn14 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Sel");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn15 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Aux");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn16 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Saldo");
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings1 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Autorizado", 9, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Autorizado", 9, true);
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings2 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Valor", 10, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Valor", 10, true);
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings3 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Saldo", 15, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Saldo", 15, true);
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance11 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance12 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance13 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance14 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idAutorizacionPago");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idTipo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn3 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idOrigen");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn4 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Numero");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn5 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Concepto");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn6 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Notas");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn7 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("FechaPago");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn8 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idFormaPago");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn9 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCuenta");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn10 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Autorizado");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn11 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Saldo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn12 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Valor");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn13 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("NuevoPago");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn14 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("NumPago");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn15 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Sel");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn16 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Aux");
			Infragistics.Win.Appearance appearance15 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance16 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance17 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance18 = new Infragistics.Win.Appearance();
			this.ultraGrid1 = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource1 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.label3 = new System.Windows.Forms.Label();
			this.txtBusqueda = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.btBuscar = new DevExpress.XtraEditors.SimpleButton();
			this.cmbFormaPago = new Infragistics.Win.UltraWinGrid.UltraDropDown();
			this.label4 = new System.Windows.Forms.Label();
			this.cmbCuenta = new Infragistics.Win.UltraWinGrid.UltraDropDown();
			this.cmbCuentas = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.btAsignar = new DevExpress.XtraEditors.SimpleButton();
			this.btnAsignar = new DevExpress.XtraEditors.SimpleButton();
			this.lblCobro = new System.Windows.Forms.Label();
			this.txtPago = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.chkSeleccionar = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.btnAceptar = new DevExpress.XtraEditors.SimpleButton();
			this.btnCancelar = new DevExpress.XtraEditors.SimpleButton();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtBusqueda)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbFormaPago)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCuenta)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCuentas)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtPago)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			this.SuspendLayout();
			// 
			// ultraGrid1
			// 
			this.ultraGrid1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.ultraGrid1.Cursor = System.Windows.Forms.Cursors.Default;
			this.ultraGrid1.DataSource = this.ultraDataSource1;
			appearance1.BackColor = System.Drawing.Color.White;
			this.ultraGrid1.DisplayLayout.Appearance = appearance1;
			this.ultraGrid1.DisplayLayout.AutoFitColumns = true;
			ultraGridColumn1.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn1.Width = 88;
			ultraGridColumn2.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridColumn2.Hidden = true;
			ultraGridColumn2.Width = 90;
			ultraGridColumn3.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn3.Header.VisiblePosition = 2;
			ultraGridColumn3.Hidden = true;
			ultraGridColumn3.Width = 95;
			ultraGridColumn4.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn4.Header.Caption = "Número";
			ultraGridColumn4.Header.VisiblePosition = 3;
			ultraGridColumn4.Width = 95;
			ultraGridColumn5.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn5.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn5.Header.VisiblePosition = 4;
			ultraGridColumn5.Width = 187;
			ultraGridColumn6.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn6.Header.VisiblePosition = 5;
			ultraGridColumn6.Width = 280;
			ultraGridColumn7.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn7.Header.Caption = "Pago";
			ultraGridColumn7.Header.VisiblePosition = 6;
			ultraGridColumn7.Width = 70;
			ultraGridColumn8.Header.Caption = "Forma de Pago";
			ultraGridColumn8.Header.VisiblePosition = 7;
			ultraGridColumn8.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownList;
			ultraGridColumn8.Width = 140;
			ultraGridColumn9.Header.Caption = "Cuenta";
			ultraGridColumn9.Header.VisiblePosition = 8;
			ultraGridColumn9.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownList;
			ultraGridColumn9.Width = 140;
			appearance2.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn10.CellAppearance = appearance2;
			ultraGridColumn10.Format = "#,##0.00";
			ultraGridColumn10.Header.VisiblePosition = 9;
			ultraGridColumn10.MaskInput = "{LOC}nn,nnn,nnn.nn";
			ultraGridColumn10.Width = 93;
			ultraGridColumn11.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance3.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn11.CellAppearance = appearance3;
			ultraGridColumn11.Format = "#,##0.00";
			ultraGridColumn11.Header.Caption = "Total";
			ultraGridColumn11.Header.VisiblePosition = 11;
			ultraGridColumn11.MaskInput = "{LOC}nn,nnn,nnn.nn";
			ultraGridColumn11.Width = 93;
			ultraGridColumn12.Header.Caption = "Pago Saldo";
			ultraGridColumn12.Header.VisiblePosition = 12;
			ultraGridColumn12.Width = 70;
			ultraGridColumn13.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance4.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn13.CellAppearance = appearance4;
			ultraGridColumn13.Header.Caption = "Número";
			ultraGridColumn13.Header.VisiblePosition = 13;
			ultraGridColumn13.Width = 56;
			ultraGridColumn14.Header.VisiblePosition = 14;
			ultraGridColumn14.Width = 32;
			ultraGridColumn15.Header.VisiblePosition = 15;
			ultraGridColumn15.Hidden = true;
			ultraGridColumn16.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance5.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn16.CellAppearance = appearance5;
			ultraGridColumn16.Format = "#,##0.00";
			ultraGridColumn16.Header.VisiblePosition = 10;
			ultraGridColumn16.Width = 91;
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
																										 ultraGridColumn16});
			ultraGridBand1.GroupHeaderLines = 4;
			ultraGridBand1.Header.Caption = "";
			ultraGridBand1.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			ultraGridBand1.Override.AllowGroupBy = Infragistics.Win.DefaultableBoolean.True;
			appearance6.FontData.BoldAsString = "True";
			appearance6.FontData.SizeInPoints = 10F;
			appearance6.ForeColor = System.Drawing.Color.Black;
			appearance6.ForeColorDisabled = System.Drawing.Color.Black;
			ultraGridBand1.Override.GroupByColumnHeaderAppearance = appearance6;
			appearance7.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings1.Appearance = appearance7;
			summarySettings1.DisplayFormat = "{0: #,##0.00}";
			summarySettings1.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance8.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings2.Appearance = appearance8;
			summarySettings2.DisplayFormat = "{0: #,##0.00}";
			summarySettings2.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance9.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings3.Appearance = appearance9;
			summarySettings3.DisplayFormat = "{0: #,##0.00}";
			summarySettings3.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			ultraGridBand1.Summaries.AddRange(new Infragistics.Win.UltraWinGrid.SummarySettings[] {
																																															summarySettings1,
																																															summarySettings2,
																																															summarySettings3});
			ultraGridBand1.SummaryFooterCaption = "";
			this.ultraGrid1.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.ultraGrid1.DisplayLayout.GroupByBox.Hidden = true;
			appearance10.ForeColor = System.Drawing.Color.Black;
			appearance10.ForeColorDisabled = System.Drawing.Color.Black;
			this.ultraGrid1.DisplayLayout.Override.ActiveRowAppearance = appearance10;
			this.ultraGrid1.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.ultraGrid1.DisplayLayout.Override.AllowGroupBy = Infragistics.Win.DefaultableBoolean.True;
			appearance11.BackColor = System.Drawing.Color.Transparent;
			this.ultraGrid1.DisplayLayout.Override.CardAreaAppearance = appearance11;
			appearance12.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance12.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance12.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance12.FontData.BoldAsString = "True";
			appearance12.FontData.Name = "Arial";
			appearance12.FontData.SizeInPoints = 8F;
			appearance12.ForeColor = System.Drawing.Color.White;
			appearance12.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.ultraGrid1.DisplayLayout.Override.HeaderAppearance = appearance12;
			this.ultraGrid1.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortSingle;
			appearance13.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance13.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance13.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.RowSelectorAppearance = appearance13;
			appearance14.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance14.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance14.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.SelectedRowAppearance = appearance14;
			this.ultraGrid1.DisplayLayout.ViewStyleBand = Infragistics.Win.UltraWinGrid.ViewStyleBand.OutlookGroupBy;
			this.ultraGrid1.Font = new System.Drawing.Font("Tahoma", 7.75F);
			this.ultraGrid1.Location = new System.Drawing.Point(8, 80);
			this.ultraGrid1.Name = "ultraGrid1";
			this.ultraGrid1.Size = new System.Drawing.Size(1368, 376);
			this.ultraGrid1.TabIndex = 7;
			this.ultraGrid1.CellChange += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.ultraGrid1_CellChange);
			this.ultraGrid1.KeyDown += new System.Windows.Forms.KeyEventHandler(this.ultraGrid1_KeyDown);
			this.ultraGrid1.AfterCellUpdate += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.ultraGrid1_AfterCellUpdate);
			this.ultraGrid1.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.ultraGrid1_InitializeLayout);
			// 
			// ultraDataSource1
			// 
			ultraDataColumn1.DataType = typeof(int);
			ultraDataColumn2.DataType = typeof(int);
			ultraDataColumn3.DataType = typeof(int);
			ultraDataColumn7.DataType = typeof(System.DateTime);
			ultraDataColumn7.DefaultValue = new System.DateTime(((long)(0)));
			ultraDataColumn8.DataType = typeof(int);
			ultraDataColumn9.DataType = typeof(int);
			ultraDataColumn10.DataType = typeof(System.Double);
			ultraDataColumn10.DefaultValue = 0;
			ultraDataColumn11.DataType = typeof(System.Double);
			ultraDataColumn12.DataType = typeof(System.Double);
			ultraDataColumn12.DefaultValue = 0;
			ultraDataColumn13.DataType = typeof(System.DateTime);
			ultraDataColumn15.DataType = typeof(bool);
			ultraDataColumn15.DefaultValue = false;
			ultraDataColumn16.DataType = typeof(bool);
			ultraDataColumn16.DefaultValue = false;
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
																																 ultraDataColumn11,
																																 ultraDataColumn12,
																																 ultraDataColumn13,
																																 ultraDataColumn14,
																																 ultraDataColumn15,
																																 ultraDataColumn16});
			// 
			// label3
			// 
			this.label3.AutoSize = true;
			this.label3.Location = new System.Drawing.Point(8, 11);
			this.label3.Name = "label3";
			this.label3.Size = new System.Drawing.Size(55, 16);
			this.label3.TabIndex = 287;
			this.label3.Text = "Proveedor";
			this.label3.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtBusqueda
			// 
			appearance15.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtBusqueda.Appearance = appearance15;
			this.txtBusqueda.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtBusqueda.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtBusqueda.Location = new System.Drawing.Point(80, 8);
			this.txtBusqueda.Name = "txtBusqueda";
			this.txtBusqueda.Size = new System.Drawing.Size(272, 22);
			this.txtBusqueda.TabIndex = 0;
			this.txtBusqueda.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtBusqueda_KeyPress);
			this.txtBusqueda.ValueChanged += new System.EventHandler(this.txtBusqueda_ValueChanged);
			// 
			// btBuscar
			// 
			this.btBuscar.Location = new System.Drawing.Point(376, 8);
			this.btBuscar.Name = "btBuscar";
			this.btBuscar.Size = new System.Drawing.Size(96, 22);
			this.btBuscar.TabIndex = 1;
			this.btBuscar.Text = "&Buscar";
			this.btBuscar.Click += new System.EventHandler(this.btBuscar_Click);
			// 
			// cmbFormaPago
			// 
			this.cmbFormaPago.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbFormaPago.DisplayMember = "FormaPago";
			this.cmbFormaPago.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbFormaPago.Location = new System.Drawing.Point(704, 272);
			this.cmbFormaPago.Name = "cmbFormaPago";
			this.cmbFormaPago.Size = new System.Drawing.Size(176, 86);
			this.cmbFormaPago.TabIndex = 289;
			this.cmbFormaPago.ValueMember = "idFormaPago";
			this.cmbFormaPago.Visible = false;
			// 
			// label4
			// 
			this.label4.AutoSize = true;
			this.label4.Location = new System.Drawing.Point(8, 43);
			this.label4.Name = "label4";
			this.label4.Size = new System.Drawing.Size(39, 16);
			this.label4.TabIndex = 291;
			this.label4.Text = "Cuenta";
			this.label4.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// cmbCuenta
			// 
			this.cmbCuenta.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbCuenta.DisplayMember = "Descripcion";
			this.cmbCuenta.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbCuenta.Location = new System.Drawing.Point(512, 272);
			this.cmbCuenta.Name = "cmbCuenta";
			this.cmbCuenta.Size = new System.Drawing.Size(176, 86);
			this.cmbCuenta.TabIndex = 290;
			this.cmbCuenta.ValueMember = "idCuenta";
			this.cmbCuenta.Visible = false;
			// 
			// cmbCuentas
			// 
			appearance16.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbCuentas.Appearance = appearance16;
			this.cmbCuentas.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbCuentas.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbCuentas.DisplayMember = "Descripcion";
			this.cmbCuentas.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbCuentas.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbCuentas.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbCuentas.Location = new System.Drawing.Point(80, 40);
			this.cmbCuentas.Name = "cmbCuentas";
			this.cmbCuentas.Size = new System.Drawing.Size(272, 21);
			this.cmbCuentas.TabIndex = 2;
			this.cmbCuentas.ValueMember = "idCuenta";
			this.cmbCuentas.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.cmbCuentas_InitializeLayout);
			// 
			// btAsignar
			// 
			this.btAsignar.Location = new System.Drawing.Point(376, 40);
			this.btAsignar.Name = "btAsignar";
			this.btAsignar.Size = new System.Drawing.Size(96, 22);
			this.btAsignar.TabIndex = 3;
			this.btAsignar.Text = "&Asignar Cuenta ";
			this.btAsignar.ToolTip = "Asigna la cuenta a los pagos seleccionados";
			this.btAsignar.Click += new System.EventHandler(this.btAsignar_Click);
			// 
			// btnAsignar
			// 
			this.btnAsignar.Location = new System.Drawing.Point(704, 40);
			this.btnAsignar.Name = "btnAsignar";
			this.btnAsignar.Size = new System.Drawing.Size(112, 22);
			this.btnAsignar.TabIndex = 5;
			this.btnAsignar.Text = "&Asignar Valor";
			this.btnAsignar.Click += new System.EventHandler(this.btnAsignar_Click);
			// 
			// lblCobro
			// 
			this.lblCobro.AutoSize = true;
			this.lblCobro.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblCobro.Location = new System.Drawing.Point(512, 43);
			this.lblCobro.Name = "lblCobro";
			this.lblCobro.Size = new System.Drawing.Size(31, 16);
			this.lblCobro.TabIndex = 636;
			this.lblCobro.Text = "Pago";
			this.lblCobro.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtPago
			// 
			appearance17.ForeColorDisabled = System.Drawing.Color.Red;
			this.txtPago.Appearance = appearance17;
			this.txtPago.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtPago.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.txtPago.FormatString = "#,##0.00";
			this.txtPago.Location = new System.Drawing.Point(568, 40);
			this.txtPago.Name = "txtPago";
			this.txtPago.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtPago.PromptChar = ' ';
			this.txtPago.Size = new System.Drawing.Size(112, 22);
			this.txtPago.TabIndex = 4;
			this.txtPago.ValueChanged += new System.EventHandler(this.txtPago_ValueChanged);
			// 
			// chkSeleccionar
			// 
			this.chkSeleccionar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			appearance18.ForeColorDisabled = System.Drawing.Color.Black;
			this.chkSeleccionar.Appearance = appearance18;
			this.chkSeleccionar.BackColor = System.Drawing.Color.Transparent;
			this.chkSeleccionar.CheckAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.chkSeleccionar.Location = new System.Drawing.Point(1280, 48);
			this.chkSeleccionar.Name = "chkSeleccionar";
			this.chkSeleccionar.Size = new System.Drawing.Size(80, 22);
			this.chkSeleccionar.TabIndex = 6;
			this.chkSeleccionar.Text = "Seleccionar";
			this.chkSeleccionar.CheckedChanged += new System.EventHandler(this.chkSeleccionar_CheckedChanged);
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
			// btnAceptar
			// 
			this.btnAceptar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnAceptar.Location = new System.Drawing.Point(8, 464);
			this.btnAceptar.Name = "btnAceptar";
			this.btnAceptar.Size = new System.Drawing.Size(72, 22);
			this.btnAceptar.TabIndex = 637;
			this.btnAceptar.Text = "&Aceptar";
			this.btnAceptar.Click += new System.EventHandler(this.btnAceptar_Click);
			// 
			// btnCancelar
			// 
			this.btnCancelar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnCancelar.Location = new System.Drawing.Point(88, 464);
			this.btnCancelar.Name = "btnCancelar";
			this.btnCancelar.Size = new System.Drawing.Size(72, 22);
			this.btnCancelar.TabIndex = 638;
			this.btnCancelar.Text = "&Cancelar";
			this.btnCancelar.Click += new System.EventHandler(this.btnCancelar_Click);
			// 
			// frmListaDeAutorizacionesPagoPorCliente
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 14);
			this.ClientSize = new System.Drawing.Size(1384, 494);
			this.Controls.Add(this.btnCancelar);
			this.Controls.Add(this.btnAceptar);
			this.Controls.Add(this.chkSeleccionar);
			this.Controls.Add(this.btnAsignar);
			this.Controls.Add(this.lblCobro);
			this.Controls.Add(this.txtPago);
			this.Controls.Add(this.btAsignar);
			this.Controls.Add(this.label4);
			this.Controls.Add(this.cmbCuenta);
			this.Controls.Add(this.cmbCuentas);
			this.Controls.Add(this.cmbFormaPago);
			this.Controls.Add(this.label3);
			this.Controls.Add(this.txtBusqueda);
			this.Controls.Add(this.btBuscar);
			this.Controls.Add(this.ultraGrid1);
			this.KeyPreview = true;
			this.Name = "frmListaDeAutorizacionesPagoPorCliente";
			this.ShowInTaskbar = false;
			this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
			this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.frmListaDeAutorizacionesPagoPorCliente_KeyDown);
			this.Load += new System.EventHandler(this.frmListaDeAutorizacionesPagoPorCliente_Load);
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtBusqueda)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbFormaPago)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCuenta)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCuentas)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtPago)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void Consulta()
		{
			string sSQL = string.Format("Exec AutorizacionesDePagoListaDocumentos {0}, {1}, '{2}', '{3}', '{4}'", IdTipo, IdCliente, DTDesde.ToString("yyyyMMdd"), DTHasta.ToString("yyyyMMdd"), this.txtBusqueda.Text.ToString());
			this.ultraGrid1.DataSource = Funcion.dvProcedimiento(cdsSeteoF, sSQL);
		}

		private void frmListaDeAutorizacionesPagoPorCliente_Load(object sender, System.EventArgs e)
		{
			#region Listas
			this.cmbCuentas.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Exec ListaCuentasCierreCaja");
			this.cmbCuentas.DisplayLayout.Bands[0].Columns["Descripcion"].Width = 250;
			this.cmbCuentas.DisplayLayout.Bands[0].Columns["idCuenta"].Hidden = true;

			this.cmbFormaPago.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Select idFormaPago, FormaPago From CompraFormaPago Where idFormaPago In (2, 4, 7)");
			this.cmbFormaPago.DisplayLayout.Bands[0].Columns["FormaPago"].Width = 200;
			this.cmbFormaPago.DisplayLayout.Bands[0].Columns["idFormaPago"].Hidden = true;

			this.cmbCuenta.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Exec ListaCuentasCierreCaja");
			this.cmbCuenta.DisplayLayout.Bands[0].Columns["Descripcion"].Width = 200;
			this.cmbCuenta.DisplayLayout.Bands[0].Columns["idCuenta"].Hidden = true;
			#endregion Listas

			Consulta();
		}

		private void btBuscar_Click(object sender, System.EventArgs e)
		{
			Consulta();
		}

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);		
		}

		private void ultraGrid1_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
			CultureInfo cultureENUS = CultureInfo.CreateSpecificCulture("ja-jp");
			string stMoneda = "#,##0.";
			string stMascara = "nnn,nnn,nnn.";
			string stCeros = "0000000000";
			string stNumero = "nnnnnnnnnn";
			string stFormato = "";
			string stInput = "";
			stFormato = stMoneda + stCeros.Substring(1, 2);
			stInput = stMascara + stNumero.Substring(1, 2);
			e.Layout.Bands[0].Columns["Autorizado"].Format = stFormato; 
			e.Layout.Bands[0].Columns["Autorizado"].MaskInput = stInput;
		}

		private void btAsignar_Click(object sender, System.EventArgs e)
		{
			if (this.ultraGrid1.Rows.Count == 0)
			{
				MessageBox.Show("No existen pagos para asignarles una cuenta.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
				return;
			}

			if (this.cmbCuentas.ActiveRow == null)
			{
				MessageBox.Show("Seleccione una cuenta.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
				return;
			}

			int i = 0;

			if (this.ultraGrid1.ActiveRow == null)
			{
				MessageBox.Show("Seleccione una fila", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
				return;
			}	

			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in ultraGrid1.Rows.All)
			{
				if ((bool)dr.Cells["Sel"].Value == true) i++;
			}

			if (i == 0)
			{
				MessageBox.Show(string.Format("Seleccione el o los pagos para asignarles la cuenta {0}.", this.cmbCuentas.Text.ToString()), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
				return;
			}

			//string sRUC = this.ultraGrid1.ActiveRow.Cells["RUC"].Value.ToString();

			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.ultraGrid1.Rows.All)
			{
				//string RUC = dr.Cells["RUC"].Value.ToString();

				//if (RUC == sRUC) 
					if ((bool)dr.Cells["Sel"].Value == true)	dr.Cells["idCuenta"].Value = (int)this.cmbCuentas.Value;
			}		
		}

		private void chkSeleccionar_CheckedChanged(object sender, System.EventArgs e)
		{
			if (this.chkSeleccionar.Checked)
			{
				foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.ultraGrid1.Rows.All)
				{
					dr.Cells["Sel"].Value = true;
				}
			}
			else
			{
				foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.ultraGrid1.Rows.All)
				{
					dr.Cells["Sel"].Value = false;
				}
			}
		}

		private void btnAsignar_Click(object sender, System.EventArgs e)
		{
			if (this.ultraGrid1.Rows.Count == 0)
			{
				MessageBox.Show("No existen filas para asignar pagos", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
				return;
			}	
			
			if ((double)this.txtPago.Value == 0)
			{
				MessageBox.Show("Ingrese un valor", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
				this.txtPago.Focus();
				return;
			}			

//			if (this.ultraGrid1.ActiveRow == null)
//			{
//				MessageBox.Show("Seleccione una fila", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
//				return;
//			}	
			
			//string sRUC = this.ultraGrid1.ActiveRow.Cells["RUC"].Value.ToString();
			double dValor = (double)this.txtPago.Value;

			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in ultraGrid1.Rows.All)
			{
				//string RUC = dr.Cells["RUC"].Value.ToString();

//				if (RUC == sRUC)
//				{
				if ((bool)dr.Cells["Sel"].Value)
				{
					if (dValor > (double)dr.Cells["Saldo"].Value)
					{
						dr.Cells["Autorizado"].Value = (double)dr.Cells["Saldo"].Value;
//						dr.Cells["Sel"].Value = true;
						dValor = dValor - (double)dr.Cells["Saldo"].Value;
					}
					else if (dValor < (double)dr.Cells["Saldo"].Value && dValor > 0)
					{
						dr.Cells["Autorizado"].Value = dValor;
//						dr.Cells["Sel"].Value = true;
						dValor = 0;
					}
					else if (dValor == (double)dr.Cells["Saldo"].Value)
					{
						dr.Cells["Autorizado"].Value = dValor;
//						dr.Cells["Sel"].Value = true;
						dValor = 0;
					}				
				}
//				}				
			}

			this.txtPago.Value = 0;
		}

		private void btnCancelar_Click(object sender, System.EventArgs e)
		{
			this.DialogResult = DialogResult.No;
		}

		private void frmListaDeAutorizacionesPagoPorCliente_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Escape) this.DialogResult = DialogResult.No;
		}

		private void btnAceptar_Click(object sender, System.EventArgs e)
		{
			#region Valida si hay registros en el grid
			if (this.ultraGrid1.Rows.Count == 0)
			{
				MessageBox.Show("No existen pagos para autorizar.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
				return;
			}
			#endregion Valida si hay registros en el grid

			#region Valida si se ha seleccionado un pago
			int i = 0;

			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in ultraGrid1.Rows.All)
			{
				if ((bool)dr.Cells["Sel"].Value == true) i++;
			}

			if (i == 0)
			{
				MessageBox.Show("Seleccione el o los pagos que va a autorizar", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
				return;
			}
			#endregion Valida si se ha seleccionado un pago

			#region Valida Forma de pago - Cuenta - Valor - Fecha Nuevo Pago
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in ultraGrid1.Rows.All)
			{
				if ((bool)dr.Cells["Sel"].Value)
				{
					if (dr.Cells["idFormaPago"].Value == DBNull.Value)
					{
						MessageBox.Show(string.Format("Seleccione la forma de pago para la autorización del documento Número '{0}'.", dr.Cells["Numero"].Value.ToString()), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
						this.ultraGrid1.ActiveRow = dr;
						this.ultraGrid1.ActiveCell = dr.Cells["idFormaPago"];
						if(ultraGrid1.ActiveCell.DroppedDown == false) ultraGrid1.ActiveCell.DroppedDown = true;						
						return;
					}
					if ((int)dr.Cells["idFormaPago"].Value != 1)
					{
						if (dr.Cells["idCuenta"].Value == DBNull.Value)
						{
							MessageBox.Show(string.Format("Seleccione la Cuenta para la autorización del Documento Número '{0}'.", dr.Cells["Numero"].Value.ToString()), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
							this.ultraGrid1.ActiveRow = dr;
							this.ultraGrid1.ActiveCell = dr.Cells["idCuenta"];
							if(ultraGrid1.ActiveCell.DroppedDown == false) ultraGrid1.ActiveCell.DroppedDown = true;							
							return;
						}
					}
					if (dr.Cells["Autorizado"].Value == DBNull.Value || (double)dr.Cells["Autorizado"].Value < 0.001)
					{
						MessageBox.Show(string.Format("Ingrese el valor de la autorización del Documento Número '{1}'.", dr.Cells["Numero"].Value.ToString()), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
						this.ultraGrid1.ActiveRow = dr;
						this.ultraGrid1.ActiveCell = dr.Cells["Autorizado"];						
						return;
					}
					if ((double)dr.Cells["Autorizado"].Value < (double)dr.Cells["Valor"].Value)
					{
						if (dr.Cells["NuevoPago"].Value.ToString().Length == 0)
						{
							MessageBox.Show(string.Format("Ingrese la fecha para el pago del saldo de la autorización del Documento Número '{1}'.", dr.Cells["Numero"].Value.ToString()), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
							this.ultraGrid1.ActiveRow = dr;
							this.ultraGrid1.ActiveCell = dr.Cells["NuevoPago"];
							if(ultraGrid1.ActiveCell.DroppedDown == false) ultraGrid1.ActiveCell.DroppedDown = true;
							return;
						}
					}
				}
			}
			#endregion Valida Forma de pago - Cuenta - Valor - Fecha Nuevo Pago

			this.DialogResult = DialogResult.OK;
		}

		private void ultraGrid1_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			switch(e.KeyValue)
			{
					#region Page Down Abre lista desplegable
				case (int) Keys.PageDown:
					if(ultraGrid1.ActiveCell.DroppedDown == false)
						ultraGrid1.ActiveCell.DroppedDown = true;
					break;
					#endregion Page Down Abre lista desplegable
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
					if(ultraGrid1.ActiveCell.DroppedDown == false)
						ultraGrid1.ActiveCell.DroppedDown = true;
					break;
					#endregion Enter
			}
		}

		private void ultraGrid1_AfterCellUpdate(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			if (e.Cell.Column.ToString() == "Notas")
			{					
				Funcion.EjecutaSQL(cdsSeteoF, string.Format("Update AutorizacionDePagos Set Notas = '{0}' Where idAutorizacionPago = {1}", e.Cell.Row.Cells["Notas"].Value.ToString(), (int)e.Cell.Row.Cells["idAutorizacionPago"].Value), true);
			}

			if (e.Cell.Column.ToString() == "idFormaPago")
			{				
				int idFormaPago = 0;

				if (e.Cell.Row.Cells["idFormaPago"].Value != System.DBNull.Value) idFormaPago = (int)e.Cell.Row.Cells["idFormaPago"].Value;
				
				Funcion.EjecutaSQL(cdsSeteoF, string.Format("Update AutorizacionDePagos Set idFormaPago = {0} Where idAutorizacionPago = {1}", idFormaPago, (int)e.Cell.Row.Cells["idAutorizacionPago"].Value), true);
			}

			if (e.Cell.Column.ToString() == "idCuenta")
			{				
				int idCuenta = 0;

				if (e.Cell.Row.Cells["idCuenta"].Value != System.DBNull.Value) idCuenta = (int)e.Cell.Row.Cells["idCuenta"].Value;
				
				Funcion.EjecutaSQL(cdsSeteoF, string.Format("Update AutorizacionDePagos Set idCuenta = {0} Where idAutorizacionPago = {1}", idCuenta, (int)e.Cell.Row.Cells["idAutorizacionPago"].Value), true);
			}


			if (e.Cell.Column.ToString() == "Autorizado")
			{ 
				double Saldo = (double)e.Cell.Row.Cells["Saldo"].Value;
				double Autorizado = (double)e.Cell.Row.Cells["Autorizado"].Value;

				if (Autorizado > Saldo)
				{
					e.Cell.Row.Cells["Autorizado"].Value = (double)e.Cell.OriginalValue;
					MessageBox.Show(string.Format("El valor máximo a autorizar no puede ser mayor a {0}.", Saldo), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
					return;
				}

				if (Math.Round(Autorizado, 2) < Math.Round(Saldo, 2))
				{
					DateTime FPago = (DateTime)e.Cell.Row.Cells["FechaPago"].Value;
					e.Cell.Row.Cells["NuevoPago"].Value = FPago.AddDays(8);
				}				

				Funcion.EjecutaSQL(cdsSeteoF, string.Format("Update AutorizacionDePagos Set Autorizado = {0} Where idAutorizacionPago = {1}", (double)e.Cell.Row.Cells["Autorizado"].Value, (int)e.Cell.Row.Cells["idAutorizacionPago"].Value), true);
			}
	
			if (e.Cell.Column.ToString() == "NuevoPago")
			{ 
				if (e.Cell.Row.Cells["NuevoPago"].Value != System.DBNull.Value)
				{
					DateTime FPagoAnt = (DateTime)e.Cell.Row.Cells["FechaPago"].Value;
					DateTime FPagoNue = (DateTime)e.Cell.Row.Cells["NuevoPago"].Value;
					if (FPagoAnt > FPagoNue)
					{
						e.Cell.Row.Cells["NuevoPago"].Value = (DateTime)e.Cell.OriginalValue;
						MessageBox.Show(string.Format("La fecha para el pago del saldo no puede ser antes de '{0}'.", FPagoAnt), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
					}

					if (e.Cell.Row.Cells["NuevoPago"].Value != System.DBNull.Value) 
					{
						Funcion.EjecutaSQL(cdsSeteoF, string.Format("Update AutorizacionDePagos Set NuevoPago = '{0}' Where idAutorizacionPago = {1}", 
							FPagoNue.ToString("yyyyMMdd"), (int)e.Cell.Row.Cells["idAutorizacionPago"].Value), true);
					}
					else
					{
						Funcion.EjecutaSQL(cdsSeteoF, string.Format("Update AutorizacionDePagos Set NuevoPago = NULL Where idAutorizacionPago = {0}", 
							(int)e.Cell.Row.Cells["idAutorizacionPago"].Value), true);
					}
				}
			}

			if (e.Cell.Column.ToString() == "Prioridad")
			{				
				int bTemp = 0;

				if ((bool)e.Cell.Row.Cells["Prioridad"].Value) bTemp = 1;
				
				Funcion.EjecutaSQL(cdsSeteoF, string.Format("Update AutorizacionDePagos Set Prioridad = {0}, Notas = '{1}' Where idAutorizacionPago = {2}", bTemp, e.Cell.Row.Cells["Notas"].Value.ToString(), (int)e.Cell.Row.Cells["idAutorizacionPago"].Value), true);
			}

			if (e.Cell.Column.ToString() == "Sel")
			{				
				int bTemp = 0;

				if ((bool)e.Cell.Row.Cells["Sel"].Value) bTemp = 1;
				
				Funcion.EjecutaSQL(cdsSeteoF, string.Format("Update AutorizacionDePagos Set Sel = {0} Where idAutorizacionPago = {1}", bTemp, (int)e.Cell.Row.Cells["idAutorizacionPago"].Value), true);

				if (bTemp == 1)
				{ 
					if ((double)e.Cell.Row.Cells["Autorizado"].Value == 0) e.Cell.Row.Cells["Autorizado"].Value = (double)e.Cell.Row.Cells["Saldo"].Value;
				}
				else 
				{
					e.Cell.Row.Cells["Autorizado"].Value = 0;									
				}
			}
		}

		private void txtPago_ValueChanged(object sender, System.EventArgs e)
		{
		
		}

		private void cmbCuentas_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}

		private void txtBusqueda_ValueChanged(object sender, System.EventArgs e)
		{
		
		}

		private void txtBusqueda_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (e.KeyChar == 13) this.btBuscar_Click(sender, e);
		}

		private void ultraGrid1_CellChange(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			if (e.Cell.Column.ToString() == "Sel")
			{			
				this.ultraGrid1.UpdateData();

				int bTemp = 0;

				if ((bool)e.Cell.Row.Cells["Sel"].Value) bTemp = 1;
				
				Funcion.EjecutaSQL(cdsSeteoF, string.Format("Update AutorizacionDePagos Set Sel = {0} Where idAutorizacionPago = {1}", bTemp, (int)e.Cell.Row.Cells["idAutorizacionPago"].Value), true);

				if (bTemp == 1)
				{ 
					if ((double)e.Cell.Row.Cells["Autorizado"].Value == 0) e.Cell.Row.Cells["Autorizado"].Value = (double)e.Cell.Row.Cells["Saldo"].Value;
				}
				else 
				{
					e.Cell.Row.Cells["Autorizado"].Value = 0;									
				}
			}
		}
	}
}

