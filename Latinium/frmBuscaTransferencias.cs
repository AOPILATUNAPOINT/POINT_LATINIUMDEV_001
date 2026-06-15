using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using DevExpress.XtraEditors;

namespace Latinium
{
	/// <summary>
	/// Summary description for frmBuscaTransferencias.
	/// </summary>
	public class frmBuscaTransferencias : DevExpress.XtraEditors.XtraForm
	{
		private Infragistics.Win.Misc.UltraLabel ultraLabel4;
		public Infragistics.Win.UltraWinGrid.UltraGrid ultraGrid1;
		private System.Windows.Forms.Label label2;
		private System.Windows.Forms.Label label8;
		private System.Windows.Forms.Label label10;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbBodDestino;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbBodOrigen;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtDesde;
		private System.Windows.Forms.Label label1;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtHasta;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtNumero;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource1;
		private C1.Data.C1DataSet cdsSeteoF;
		private System.Windows.Forms.Label lblContador;
		private System.Windows.Forms.GroupBox groupBox1;
		private System.Windows.Forms.Button btnAceptar;
		private System.Windows.Forms.Button btnCancelar;
		private System.Windows.Forms.Button btnBuscar;
		/// <summary>
		/// Required designer variable.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public frmBuscaTransferencias()
		{
			//
			// Required for Windows Form Designer support
			//
			InitializeComponent();

			//
			// TODO: Add any constructor code after InitializeComponent call
			//
		}

		int idBodega = 0;

		public frmBuscaTransferencias(int IdBodega)
		{
			//
			// Required for Windows Form Designer support
			//
			InitializeComponent();
			idBodega = IdBodega; 
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
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCompraTransf");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Fecha");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bodega");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Origen");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bodega1");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Destino");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("EstadoM");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Estado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nota");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn11 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ObserEnvio");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn12 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idTipoTransporte");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn13 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idTransporte");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn14 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("NumGuia");
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCompraTransf");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Codigo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn3 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Fecha");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn4 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Bodega");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn5 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Origen");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn6 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Bodega1");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn7 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Destino");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn8 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("EstadoM");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn9 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Estado");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn10 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nota");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn11 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("ObserEnvio");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn12 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idTipoTransporte");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn13 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idTransporte");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn14 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("NumGuia");
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton1 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmBuscaTransferencias));
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn15 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn16 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bodega");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn17 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand3 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn18 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn19 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bodega");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn20 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton2 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			this.ultraLabel4 = new Infragistics.Win.Misc.UltraLabel();
			this.ultraGrid1 = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource1 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.txtNumero = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label2 = new System.Windows.Forms.Label();
			this.dtDesde = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.label8 = new System.Windows.Forms.Label();
			this.label10 = new System.Windows.Forms.Label();
			this.cmbBodDestino = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.cmbBodOrigen = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.label1 = new System.Windows.Forms.Label();
			this.dtHasta = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.lblContador = new System.Windows.Forms.Label();
			this.groupBox1 = new System.Windows.Forms.GroupBox();
			this.btnAceptar = new System.Windows.Forms.Button();
			this.btnCancelar = new System.Windows.Forms.Button();
			this.btnBuscar = new System.Windows.Forms.Button();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNumero)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtDesde)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbBodDestino)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbBodOrigen)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtHasta)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			this.SuspendLayout();
			// 
			// ultraLabel4
			// 
			this.ultraLabel4.AutoSize = true;
			this.ultraLabel4.BackColor = System.Drawing.Color.Transparent;
			this.ultraLabel4.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.ultraLabel4.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(50)), ((System.Byte)(50)), ((System.Byte)(50)));
			this.ultraLabel4.Location = new System.Drawing.Point(8, 108);
			this.ultraLabel4.Name = "ultraLabel4";
			this.ultraLabel4.Size = new System.Drawing.Size(44, 15);
			this.ultraLabel4.TabIndex = 266;
			this.ultraLabel4.Text = "Número";
			// 
			// ultraGrid1
			// 
			this.ultraGrid1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.ultraGrid1.Cursor = System.Windows.Forms.Cursors.Default;
			this.ultraGrid1.DataSource = this.ultraDataSource1;
			appearance1.BackColor = System.Drawing.Color.White;
			appearance1.ForeColor = System.Drawing.Color.Black;
			appearance1.ForeColorDisabled = System.Drawing.Color.Black;
			this.ultraGrid1.DisplayLayout.Appearance = appearance1;
			this.ultraGrid1.DisplayLayout.AutoFitColumns = true;
			ultraGridColumn1.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn1.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn2.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn2.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn2.Header.Caption = "Número";
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridColumn2.Width = 89;
			ultraGridColumn3.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn3.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn3.Header.VisiblePosition = 2;
			ultraGridColumn3.Width = 68;
			ultraGridColumn4.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn4.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn4.Header.VisiblePosition = 12;
			ultraGridColumn4.Hidden = true;
			ultraGridColumn4.Width = 74;
			ultraGridColumn5.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn5.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn5.Header.VisiblePosition = 3;
			ultraGridColumn5.Width = 185;
			ultraGridColumn6.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn6.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn6.Header.VisiblePosition = 13;
			ultraGridColumn6.Hidden = true;
			ultraGridColumn6.Width = 94;
			ultraGridColumn7.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn7.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn7.Header.VisiblePosition = 4;
			ultraGridColumn7.Width = 185;
			ultraGridColumn8.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn8.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn8.Header.Caption = "Estado";
			ultraGridColumn8.Header.VisiblePosition = 5;
			ultraGridColumn8.Width = 143;
			ultraGridColumn9.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn9.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn9.Header.VisiblePosition = 6;
			ultraGridColumn9.Hidden = true;
			ultraGridColumn9.Width = 38;
			ultraGridColumn10.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn10.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn10.Header.VisiblePosition = 7;
			ultraGridColumn10.Hidden = true;
			ultraGridColumn10.Width = 44;
			ultraGridColumn11.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn11.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn11.Header.VisiblePosition = 8;
			ultraGridColumn11.Hidden = true;
			ultraGridColumn11.Width = 48;
			ultraGridColumn12.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn12.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn12.Header.VisiblePosition = 9;
			ultraGridColumn12.Hidden = true;
			ultraGridColumn12.Width = 48;
			ultraGridColumn13.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn13.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn13.Header.VisiblePosition = 10;
			ultraGridColumn13.Hidden = true;
			ultraGridColumn13.Width = 51;
			ultraGridColumn14.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn14.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn14.Header.VisiblePosition = 11;
			ultraGridColumn14.Hidden = true;
			ultraGridColumn14.Width = 68;
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
																										 ultraGridColumn14});
			this.ultraGrid1.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			appearance2.ForeColor = System.Drawing.Color.Black;
			this.ultraGrid1.DisplayLayout.CaptionAppearance = appearance2;
			this.ultraGrid1.DisplayLayout.GroupByBox.Prompt = "Arrastre columna a Agrupar";
			appearance3.ForeColor = System.Drawing.Color.Black;
			appearance3.ForeColorDisabled = System.Drawing.Color.Black;
			this.ultraGrid1.DisplayLayout.Override.ActiveRowAppearance = appearance3;
			this.ultraGrid1.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
			this.ultraGrid1.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.ultraGrid1.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.False;
			appearance4.BackColor = System.Drawing.Color.Transparent;
			this.ultraGrid1.DisplayLayout.Override.CardAreaAppearance = appearance4;
			this.ultraGrid1.DisplayLayout.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.RowSelect;
			this.ultraGrid1.DisplayLayout.Override.ColumnAutoSizeMode = Infragistics.Win.UltraWinGrid.ColumnAutoSizeMode.AllRowsInBand;
			appearance5.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance5.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance5.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance5.FontData.BoldAsString = "True";
			appearance5.FontData.Name = "Arial";
			appearance5.FontData.SizeInPoints = 10F;
			appearance5.ForeColor = System.Drawing.Color.White;
			appearance5.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.ultraGrid1.DisplayLayout.Override.HeaderAppearance = appearance5;
			this.ultraGrid1.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti;
			appearance6.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance6.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance6.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.RowSelectorAppearance = appearance6;
			appearance7.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance7.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance7.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.SelectedRowAppearance = appearance7;
			this.ultraGrid1.DisplayLayout.Override.SelectTypeCell = Infragistics.Win.UltraWinGrid.SelectType.None;
			this.ultraGrid1.DisplayLayout.Override.SelectTypeCol = Infragistics.Win.UltraWinGrid.SelectType.None;
			this.ultraGrid1.DisplayLayout.Override.SelectTypeRow = Infragistics.Win.UltraWinGrid.SelectType.Single;
			this.ultraGrid1.DisplayLayout.ScrollBounds = Infragistics.Win.UltraWinGrid.ScrollBounds.ScrollToFill;
			this.ultraGrid1.DisplayLayout.ScrollStyle = Infragistics.Win.UltraWinGrid.ScrollStyle.Immediate;
			this.ultraGrid1.DisplayLayout.TabNavigation = Infragistics.Win.UltraWinGrid.TabNavigation.NextControl;
			this.ultraGrid1.DisplayLayout.ViewStyle = Infragistics.Win.UltraWinGrid.ViewStyle.SingleBand;
			this.ultraGrid1.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.ultraGrid1.Location = new System.Drawing.Point(9, 152);
			this.ultraGrid1.Name = "ultraGrid1";
			this.ultraGrid1.Size = new System.Drawing.Size(691, 168);
			this.ultraGrid1.TabIndex = 6;
			this.ultraGrid1.DoubleClick += new System.EventHandler(this.ultraGrid1_DoubleClick);
			this.ultraGrid1.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.ultraGrid1_KeyPress);
			this.ultraGrid1.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.ultraGrid1_InitializeLayout);
			// 
			// ultraDataSource1
			// 
			ultraDataColumn1.DataType = typeof(int);
			ultraDataColumn3.DataType = typeof(System.DateTime);
			ultraDataColumn3.DefaultValue = new System.DateTime(((long)(0)));
			ultraDataColumn4.DataType = typeof(int);
			ultraDataColumn6.DataType = typeof(int);
			ultraDataColumn9.DataType = typeof(int);
			ultraDataColumn12.DataType = typeof(int);
			ultraDataColumn13.DataType = typeof(int);
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
																																 ultraDataColumn14});
			// 
			// txtNumero
			// 
			appearance8.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtNumero.Appearance = appearance8;
			this.txtNumero.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtNumero.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNumero.Location = new System.Drawing.Point(64, 104);
			this.txtNumero.Name = "txtNumero";
			this.txtNumero.Size = new System.Drawing.Size(136, 22);
			this.txtNumero.TabIndex = 0;
			this.txtNumero.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtNumero_KeyPress);
			this.txtNumero.ValueChanged += new System.EventHandler(this.txtNumero_ValueChanged);
			// 
			// label2
			// 
			this.label2.AutoSize = true;
			this.label2.Location = new System.Drawing.Point(8, 10);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(35, 16);
			this.label2.TabIndex = 268;
			this.label2.Text = "Desde";
			this.label2.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.label2.Click += new System.EventHandler(this.label2_Click);
			// 
			// dtDesde
			// 
			this.dtDesde.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton1.Caption = "Today";
			this.dtDesde.DateButtons.Add(dateButton1);
			this.dtDesde.Enabled = false;
			this.dtDesde.Format = "dd/MM/yyyy";
			this.dtDesde.Location = new System.Drawing.Point(64, 8);
			this.dtDesde.Name = "dtDesde";
			this.dtDesde.NonAutoSizeHeight = 23;
			this.dtDesde.Size = new System.Drawing.Size(112, 21);
			this.dtDesde.SpinButtonsVisible = true;
			this.dtDesde.TabIndex = 1;
			this.dtDesde.Value = ((object)(resources.GetObject("dtDesde.Value")));
			this.dtDesde.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.dtDesde_KeyPress);
			// 
			// label8
			// 
			this.label8.AutoSize = true;
			this.label8.BackColor = System.Drawing.Color.Transparent;
			this.label8.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label8.Location = new System.Drawing.Point(8, 74);
			this.label8.Name = "label8";
			this.label8.Size = new System.Drawing.Size(42, 17);
			this.label8.TabIndex = 272;
			this.label8.Text = "Destino";
			// 
			// label10
			// 
			this.label10.AutoSize = true;
			this.label10.BackColor = System.Drawing.Color.Transparent;
			this.label10.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label10.Location = new System.Drawing.Point(8, 42);
			this.label10.Name = "label10";
			this.label10.Size = new System.Drawing.Size(37, 17);
			this.label10.TabIndex = 271;
			this.label10.Text = "Origen";
			// 
			// cmbBodDestino
			// 
			this.cmbBodDestino.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbBodDestino.Cursor = System.Windows.Forms.Cursors.Default;
			ultraGridColumn15.Header.VisiblePosition = 0;
			ultraGridColumn15.Width = 312;
			ultraGridColumn16.Header.VisiblePosition = 1;
			ultraGridColumn16.Hidden = true;
			ultraGridColumn16.Width = 50;
			ultraGridColumn17.Header.VisiblePosition = 2;
			ultraGridColumn17.Width = 80;
			ultraGridBand2.Columns.AddRange(new object[] {
																										 ultraGridColumn15,
																										 ultraGridColumn16,
																										 ultraGridColumn17});
			this.cmbBodDestino.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			this.cmbBodDestino.DisplayMember = "Nombre";
			this.cmbBodDestino.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbBodDestino.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbBodDestino.Location = new System.Drawing.Point(64, 72);
			this.cmbBodDestino.MaxDropDownItems = 30;
			this.cmbBodDestino.Name = "cmbBodDestino";
			this.cmbBodDestino.Size = new System.Drawing.Size(296, 21);
			this.cmbBodDestino.TabIndex = 4;
			this.cmbBodDestino.ValueMember = "Bodega";
			this.cmbBodDestino.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.cmbBodDestino_KeyPress);
			// 
			// cmbBodOrigen
			// 
			this.cmbBodOrigen.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbBodOrigen.Cursor = System.Windows.Forms.Cursors.Default;
			ultraGridColumn18.Header.VisiblePosition = 0;
			ultraGridColumn18.Width = 312;
			ultraGridColumn19.Header.VisiblePosition = 1;
			ultraGridColumn19.Hidden = true;
			ultraGridColumn19.Width = 50;
			ultraGridColumn20.Header.VisiblePosition = 2;
			ultraGridColumn20.Width = 80;
			ultraGridBand3.Columns.AddRange(new object[] {
																										 ultraGridColumn18,
																										 ultraGridColumn19,
																										 ultraGridColumn20});
			this.cmbBodOrigen.DisplayLayout.BandsSerializer.Add(ultraGridBand3);
			this.cmbBodOrigen.DisplayMember = "Nombre";
			this.cmbBodOrigen.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbBodOrigen.Enabled = false;
			this.cmbBodOrigen.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbBodOrigen.Location = new System.Drawing.Point(64, 40);
			this.cmbBodOrigen.MaxDropDownItems = 30;
			this.cmbBodOrigen.Name = "cmbBodOrigen";
			this.cmbBodOrigen.Size = new System.Drawing.Size(296, 21);
			this.cmbBodOrigen.TabIndex = 3;
			this.cmbBodOrigen.ValueMember = "Bodega";
			this.cmbBodOrigen.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.cmbBodOrigen_KeyPress);
			this.cmbBodOrigen.ValueChanged += new System.EventHandler(this.cmbBodOrigen_ValueChanged);
			// 
			// label1
			// 
			this.label1.AutoSize = true;
			this.label1.Location = new System.Drawing.Point(200, 10);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(33, 16);
			this.label1.TabIndex = 274;
			this.label1.Text = "Hasta";
			this.label1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// dtHasta
			// 
			this.dtHasta.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton2.Caption = "Today";
			this.dtHasta.DateButtons.Add(dateButton2);
			this.dtHasta.Enabled = false;
			this.dtHasta.Format = "dd/MM/yyyy";
			this.dtHasta.Location = new System.Drawing.Point(248, 8);
			this.dtHasta.Name = "dtHasta";
			this.dtHasta.NonAutoSizeHeight = 23;
			this.dtHasta.Size = new System.Drawing.Size(112, 21);
			this.dtHasta.SpinButtonsVisible = true;
			this.dtHasta.TabIndex = 2;
			this.dtHasta.Value = ((object)(resources.GetObject("dtHasta.Value")));
			this.dtHasta.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.dtHasta_KeyPress);
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
			// lblContador
			// 
			this.lblContador.AutoSize = true;
			this.lblContador.Font = new System.Drawing.Font("Tahoma", 9F);
			this.lblContador.ForeColor = System.Drawing.Color.Firebrick;
			this.lblContador.Location = new System.Drawing.Point(472, 105);
			this.lblContador.Name = "lblContador";
			this.lblContador.Size = new System.Drawing.Size(0, 18);
			this.lblContador.TabIndex = 275;
			this.lblContador.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// groupBox1
			// 
			this.groupBox1.Location = new System.Drawing.Point(0, 136);
			this.groupBox1.Name = "groupBox1";
			this.groupBox1.Size = new System.Drawing.Size(712, 8);
			this.groupBox1.TabIndex = 276;
			this.groupBox1.TabStop = false;
			// 
			// btnAceptar
			// 
			this.btnAceptar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnAceptar.CausesValidation = false;
			this.btnAceptar.Location = new System.Drawing.Point(8, 328);
			this.btnAceptar.Name = "btnAceptar";
			this.btnAceptar.TabIndex = 278;
			this.btnAceptar.Text = "&Aceptar";
			this.btnAceptar.Click += new System.EventHandler(this.btnAceptar_Click);
			// 
			// btnCancelar
			// 
			this.btnCancelar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnCancelar.CausesValidation = false;
			this.btnCancelar.Location = new System.Drawing.Point(88, 328);
			this.btnCancelar.Name = "btnCancelar";
			this.btnCancelar.TabIndex = 279;
			this.btnCancelar.Text = "&Cancelar";
			this.btnCancelar.Click += new System.EventHandler(this.btnCancelar_Click);
			// 
			// btnBuscar
			// 
			this.btnBuscar.CausesValidation = false;
			this.btnBuscar.Location = new System.Drawing.Point(280, 104);
			this.btnBuscar.Name = "btnBuscar";
			this.btnBuscar.TabIndex = 277;
			this.btnBuscar.Text = "&Buscar";
			this.btnBuscar.Click += new System.EventHandler(this.btnBuscar_Click);
			// 
			// frmBuscaTransferencias
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 14);
			this.ClientSize = new System.Drawing.Size(710, 360);
			this.ControlBox = false;
			this.Controls.Add(this.btnAceptar);
			this.Controls.Add(this.btnCancelar);
			this.Controls.Add(this.btnBuscar);
			this.Controls.Add(this.groupBox1);
			this.Controls.Add(this.lblContador);
			this.Controls.Add(this.label1);
			this.Controls.Add(this.dtHasta);
			this.Controls.Add(this.label8);
			this.Controls.Add(this.label10);
			this.Controls.Add(this.cmbBodDestino);
			this.Controls.Add(this.cmbBodOrigen);
			this.Controls.Add(this.label2);
			this.Controls.Add(this.dtDesde);
			this.Controls.Add(this.ultraLabel4);
			this.Controls.Add(this.ultraGrid1);
			this.Controls.Add(this.txtNumero);
			this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
			this.KeyPreview = true;
			this.MaximizeBox = false;
			this.MinimizeBox = false;
			this.Name = "frmBuscaTransferencias";
			this.ShowInTaskbar = false;
			this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
			this.Text = "Buscar";
			this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.frmBuscaTransferencias_KeyDown);
			this.Load += new System.EventHandler(this.frmBuscaTransferencias_Load);
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNumero)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtDesde)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbBodDestino)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbBodOrigen)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtHasta)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void label2_Click(object sender, System.EventArgs e)
		{
			
		}

		private void Aceptar()
		{
			if (this.ultraGrid1.Rows.Count == 0)
			{
				MessageBox.Show("No hay filas, Realice una busqueda.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.txtNumero.Select();
				return;
			}

			if (this.ultraGrid1.ActiveRow == null || this.ultraGrid1.ActiveRow.Cells["idCompraTransf"].Value == System.DBNull.Value)
			{
				MessageBox.Show("Seleccione una fila.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.ultraGrid1.Select();
				return;
			}

			this.DialogResult = DialogResult.OK;
		}

		private void frmBuscaTransferencias_Load(object sender, System.EventArgs e)
		{	
			this.cmbBodOrigen.DataSource = Funcion.dvProcedimiento(cdsSeteoF, string.Format("Exec ListaBodegas {0}, 42", MenuLatinium.IdUsuario));

			if (idBodega == 0) 
			{
				this.cmbBodOrigen.Enabled = true; 
				this.dtDesde.Enabled = true;
				this.dtHasta.Enabled = true;
			}
			else 
			{
				this.cmbBodOrigen.Value = idBodega;				
				this.dtDesde.Value = "01/" + DateTime.Now.Month.ToString() + "/" + DateTime.Now.Year.ToString();
				this.dtHasta.Value = DateTime.DaysInMonth(DateTime.Now.Year, DateTime.Now.Month).ToString() + "/" + DateTime.Now.Month.ToString() + "/" + DateTime.Now.Year.ToString();
			}
						
			this.txtNumero.Focus();
		}

		private void txtNumero_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (e.KeyChar == 13) this.btnBuscar.Focus();
		}

		private void dtDesde_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (e.KeyChar == 13) this.dtHasta.Focus();
		}

		private void dtHasta_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (e.KeyChar == 13) this.cmbBodOrigen.Focus();
		}

		private void cmbBodOrigen_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (e.KeyChar == 13) this.cmbBodDestino.Focus();
		}

		private void cmbBodDestino_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (e.KeyChar == 13) this.txtNumero.Focus();
		}

		private void frmBuscaTransferencias_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Escape) this.DialogResult = DialogResult.No;
		}

		private void ultraGrid1_DoubleClick(object sender, System.EventArgs e)
		{
			Aceptar();
		}

		private void ultraGrid1_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (e.KeyChar == 13) Aceptar();
		}

		private void txtNumero_ValueChanged(object sender, System.EventArgs e)
		{
		
		}

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);		
		}

		private void ultraGrid1_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}

		private void cmbBodOrigen_ValueChanged(object sender, System.EventArgs e)
		{
			DateTime dtFecha = DateTime.Today;

			if (this.cmbBodOrigen.ActiveRow != null)
			{
				bool bConsignacion = Funcion.bEscalarSQL(cdsSeteoF, string.Format("Select Consignacion From Bodega Where Bodega = {0}", (int)this.cmbBodOrigen.Value), true);

				int iBodegaOrigen = (int)this.cmbBodOrigen.Value;
				if (bConsignacion)
					iBodegaOrigen = Funcion.iEscalarSQL(cdsSeteoF, string.Format("select Local from AsignaBodegasLocales Where Bodega = {0} And Consignacion = 1", (int)this.cmbBodOrigen.Value));

				this.cmbBodDestino.DataSource = Funcion.dvProcedimiento(cdsSeteoF, string.Format("Exec ListaDeBodegasActivas '{0}', {1}, 0", dtFecha.ToString("yyyyMMdd"), (int)this.cmbBodOrigen.Value));
			}
		}

		private void btnBuscar_Click(object sender, System.EventArgs e)
		{
			bool bFechas = false;

			DateTime dtDesde = DateTime.Today;
			DateTime dtHasta = DateTime.Today;
			
			int IBodega = 0;
			int IBodega1 = 0;

			if (this.dtDesde.Value != System.DBNull.Value || this.dtHasta.Value != System.DBNull.Value)
			{
				bFechas = true;

				dtDesde = (DateTime)this.dtDesde.Value;
				dtHasta = (DateTime)this.dtHasta.Value;
			}

			if (this.cmbBodOrigen.ActiveRow != null) IBodega = (int)this.cmbBodOrigen.Value;
			if (this.cmbBodDestino.ActiveRow != null) IBodega1 = (int)this.cmbBodDestino.Value;

			string sSQL = string.Format("Exec ConsultaTransferencias '{0}', '{1}', '{2}', {3}, {4}, {5}", this.txtNumero.Text.ToString(), dtDesde.ToString("yyyyMMdd"), dtHasta.ToString("yyyyMMdd"), bFechas, IBodega, IBodega1);
			this.ultraGrid1.DataSource = Funcion.dvProcedimiento(cdsSeteoF, sSQL);

			this.lblContador.Text = ultraGrid1.Rows.Count + " REGISTROS ENCONTRADOS";

			if (this.ultraGrid1.Rows.Count > 0) this.ultraGrid1.Focus();
			else this.txtNumero.Focus();
		}

		private void btnCancelar_Click(object sender, System.EventArgs e)
		{
			this.DialogResult = DialogResult.No;
		}

		private void btnAceptar_Click(object sender, System.EventArgs e)
		{
			Aceptar();
		}
	}
}

