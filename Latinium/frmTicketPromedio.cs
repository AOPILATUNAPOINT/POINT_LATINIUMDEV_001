using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using DevExpress.XtraEditors;
using CrystalDecisions.Shared;
using Infragistics.Win.UltraWinGrid;
using System.Data;
using System.Data.SqlClient;
using System.Globalization;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de frmTicketPromedio.
	/// </summary>
	public class frmTicketPromedio : System.Windows.Forms.Form
	{
		private System.Windows.Forms.Button btnVer;
		private System.Windows.Forms.Label lblBodega;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbBodega;
		private System.Windows.Forms.Label label1;
		private System.Windows.Forms.Label label2;
		private System.Windows.Forms.Label label53;
		private System.Windows.Forms.Label label3;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtDesde;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtHasta;
		private System.Windows.Forms.Label label4;
		private System.Windows.Forms.Button button1;
		private System.Windows.Forms.GroupBox groupBox1;
		private System.Windows.Forms.Label label5;
		private Infragistics.Win.UltraWinGrid.UltraGrid uGridGastos;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtdTicketLocal;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtdSillon;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtTicket;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource1;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor dTicketSillon;
		private System.Windows.Forms.Label label6;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtAnio;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtidTicketPromedio;
		private C1.Data.C1DataSet cdsSeteoF;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public frmTicketPromedio()
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
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmTicketPromedio));
			Infragistics.Win.Appearance appearance1 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bodega");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo", 0);
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton1 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton2 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idDetalleTicketPromedio");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idTicketPromedio");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ENERO");
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FEBRERO");
			Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance11 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("MARZO");
			Infragistics.Win.Appearance appearance12 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance13 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ABRIL");
			Infragistics.Win.Appearance appearance14 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance15 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn11 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("MAYO");
			Infragistics.Win.Appearance appearance16 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance17 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn12 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("JUNIO");
			Infragistics.Win.Appearance appearance18 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance19 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn13 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("JULIO");
			Infragistics.Win.Appearance appearance20 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance21 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn14 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("AGOSTO");
			Infragistics.Win.Appearance appearance22 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance23 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn15 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("SEPTIEMBRE");
			Infragistics.Win.Appearance appearance24 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance25 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn16 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("OCTUBRE");
			Infragistics.Win.Appearance appearance26 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance27 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn17 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("NOVIEMBRE");
			Infragistics.Win.Appearance appearance28 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance29 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn18 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("DICIEMBRE");
			Infragistics.Win.Appearance appearance30 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance31 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn19 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Total");
			Infragistics.Win.Appearance appearance32 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance33 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance34 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance35 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance36 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance37 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance38 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance39 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idDetalleTicketPromedio");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idTicketPromedio");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn3 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Codigo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn4 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("ENERO");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn5 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("FEBRERO");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn6 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("MARZO");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn7 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("ABRIL");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn8 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("MAYO");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn9 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("JUNIO");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn10 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("JULIO");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn11 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("AGOSTO");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn12 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("SEPTIEMBRE");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn13 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("OCTUBRE");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn14 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("NOVIEMBRE");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn15 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("DICIEMBRE");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn16 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Total");
			Infragistics.Win.Appearance appearance40 = new Infragistics.Win.Appearance();
			this.txtAnio = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.btnVer = new System.Windows.Forms.Button();
			this.lblBodega = new System.Windows.Forms.Label();
			this.cmbBodega = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.label1 = new System.Windows.Forms.Label();
			this.label2 = new System.Windows.Forms.Label();
			this.txtdTicketLocal = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtdSillon = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label53 = new System.Windows.Forms.Label();
			this.label3 = new System.Windows.Forms.Label();
			this.dtDesde = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.dtHasta = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.txtTicket = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label4 = new System.Windows.Forms.Label();
			this.button1 = new System.Windows.Forms.Button();
			this.groupBox1 = new System.Windows.Forms.GroupBox();
			this.label5 = new System.Windows.Forms.Label();
			this.uGridGastos = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource1 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.dTicketSillon = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label6 = new System.Windows.Forms.Label();
			this.txtidTicketPromedio = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			((System.ComponentModel.ISupportInitialize)(this.txtAnio)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbBodega)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtdTicketLocal)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtdSillon)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtDesde)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtHasta)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtTicket)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridGastos)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dTicketSillon)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtidTicketPromedio)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			this.SuspendLayout();
			// 
			// txtAnio
			// 
			this.txtAnio.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtAnio.Location = new System.Drawing.Point(56, 6);
			this.txtAnio.MinValue = 2017;
			this.txtAnio.Name = "txtAnio";
			this.txtAnio.PromptChar = ' ';
			this.txtAnio.Size = new System.Drawing.Size(88, 21);
			this.txtAnio.SpinButtonDisplayStyle = Infragistics.Win.ButtonDisplayStyle.Always;
			this.txtAnio.TabIndex = 741;
			// 
			// btnVer
			// 
			this.btnVer.Image = ((System.Drawing.Image)(resources.GetObject("btnVer.Image")));
			this.btnVer.Location = new System.Drawing.Point(168, 5);
			this.btnVer.Name = "btnVer";
			this.btnVer.Size = new System.Drawing.Size(30, 23);
			this.btnVer.TabIndex = 742;
			this.btnVer.Click += new System.EventHandler(this.btnVer_Click);
			// 
			// lblBodega
			// 
			this.lblBodega.AutoSize = true;
			this.lblBodega.Location = new System.Drawing.Point(8, 43);
			this.lblBodega.Name = "lblBodega";
			this.lblBodega.Size = new System.Drawing.Size(31, 16);
			this.lblBodega.TabIndex = 744;
			this.lblBodega.Text = "Local";
			this.lblBodega.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// cmbBodega
			// 
			appearance1.FontData.Name = "Tahoma";
			this.cmbBodega.Appearance = appearance1;
			this.cmbBodega.CausesValidation = false;
			this.cmbBodega.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbBodega.Cursor = System.Windows.Forms.Cursors.Default;
			ultraGridColumn1.Header.VisiblePosition = 2;
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridColumn2.Width = 180;
			ultraGridColumn3.Header.VisiblePosition = 0;
			ultraGridBand1.Columns.AddRange(new object[] {
																										 ultraGridColumn1,
																										 ultraGridColumn2,
																										 ultraGridColumn3});
			this.cmbBodega.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.cmbBodega.DisplayMember = "Nombre";
			this.cmbBodega.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbBodega.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbBodega.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbBodega.Location = new System.Drawing.Point(56, 40);
			this.cmbBodega.MaxDropDownItems = 30;
			this.cmbBodega.Name = "cmbBodega";
			this.cmbBodega.Size = new System.Drawing.Size(180, 22);
			this.cmbBodega.TabIndex = 743;
			this.cmbBodega.ValueMember = "Bodega";
			this.cmbBodega.ValueChanged += new System.EventHandler(this.cmbBodega_ValueChanged);
			// 
			// label1
			// 
			this.label1.AutoSize = true;
			this.label1.Location = new System.Drawing.Point(8, 72);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(65, 16);
			this.label1.TabIndex = 745;
			this.label1.Text = "Sillones No:";
			this.label1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label2
			// 
			this.label2.AutoSize = true;
			this.label2.Location = new System.Drawing.Point(8, 98);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(140, 16);
			this.label2.TabIndex = 746;
			this.label2.Text = "Ticket Promedio Por Local:";
			this.label2.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtdTicketLocal
			// 
			appearance2.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtdTicketLocal.Appearance = appearance2;
			this.txtdTicketLocal.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtdTicketLocal.FormatString = "#,##0.00";
			this.txtdTicketLocal.Location = new System.Drawing.Point(160, 96);
			this.txtdTicketLocal.Name = "txtdTicketLocal";
			this.txtdTicketLocal.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtdTicketLocal.PromptChar = ' ';
			this.txtdTicketLocal.Size = new System.Drawing.Size(80, 21);
			this.txtdTicketLocal.TabIndex = 748;
			// 
			// txtdSillon
			// 
			appearance3.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtdSillon.Appearance = appearance3;
			this.txtdSillon.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtdSillon.FormatString = "#,##0.00";
			this.txtdSillon.Location = new System.Drawing.Point(160, 72);
			this.txtdSillon.Name = "txtdSillon";
			this.txtdSillon.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtdSillon.PromptChar = ' ';
			this.txtdSillon.Size = new System.Drawing.Size(80, 21);
			this.txtdSillon.TabIndex = 747;
			// 
			// label53
			// 
			this.label53.AutoSize = true;
			this.label53.BackColor = System.Drawing.Color.Transparent;
			this.label53.Location = new System.Drawing.Point(424, 40);
			this.label53.Name = "label53";
			this.label53.Size = new System.Drawing.Size(34, 16);
			this.label53.TabIndex = 752;
			this.label53.Text = "Hasta";
			this.label53.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label3
			// 
			this.label3.AutoSize = true;
			this.label3.BackColor = System.Drawing.Color.Transparent;
			this.label3.Location = new System.Drawing.Point(424, 10);
			this.label3.Name = "label3";
			this.label3.Size = new System.Drawing.Size(37, 16);
			this.label3.TabIndex = 751;
			this.label3.Text = "Desde";
			this.label3.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// dtDesde
			// 
			appearance4.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtDesde.Appearance = appearance4;
			this.dtDesde.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton1.Caption = "Today";
			this.dtDesde.DateButtons.Add(dateButton1);
			this.dtDesde.Format = "dd/MM/yyyy";
			this.dtDesde.Location = new System.Drawing.Point(472, 8);
			this.dtDesde.Name = "dtDesde";
			this.dtDesde.NonAutoSizeHeight = 23;
			this.dtDesde.Size = new System.Drawing.Size(112, 21);
			this.dtDesde.SpinButtonsVisible = true;
			this.dtDesde.TabIndex = 749;
			this.dtDesde.Value = ((object)(resources.GetObject("dtDesde.Value")));
			// 
			// dtHasta
			// 
			appearance5.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtHasta.Appearance = appearance5;
			this.dtHasta.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton2.Caption = "Today";
			this.dtHasta.DateButtons.Add(dateButton2);
			this.dtHasta.Format = "dd/MM/yyyy";
			this.dtHasta.Location = new System.Drawing.Point(472, 38);
			this.dtHasta.Name = "dtHasta";
			this.dtHasta.NonAutoSizeHeight = 23;
			this.dtHasta.Size = new System.Drawing.Size(112, 21);
			this.dtHasta.SpinButtonsVisible = true;
			this.dtHasta.TabIndex = 750;
			this.dtHasta.Value = ((object)(resources.GetObject("dtHasta.Value")));
			this.dtHasta.BeforeDropDown += new System.ComponentModel.CancelEventHandler(this.dtHasta_BeforeDropDown);
			// 
			// txtTicket
			// 
			appearance6.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtTicket.Appearance = appearance6;
			this.txtTicket.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtTicket.Enabled = false;
			this.txtTicket.FormatString = "#,##0.00";
			this.txtTicket.Location = new System.Drawing.Point(472, 70);
			this.txtTicket.Name = "txtTicket";
			this.txtTicket.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtTicket.PromptChar = ' ';
			this.txtTicket.Size = new System.Drawing.Size(80, 21);
			this.txtTicket.TabIndex = 754;
			// 
			// label4
			// 
			this.label4.AutoSize = true;
			this.label4.Location = new System.Drawing.Point(384, 72);
			this.label4.Name = "label4";
			this.label4.Size = new System.Drawing.Size(86, 16);
			this.label4.TabIndex = 753;
			this.label4.Text = "Ticket Promedio";
			this.label4.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// button1
			// 
			this.button1.Image = ((System.Drawing.Image)(resources.GetObject("button1.Image")));
			this.button1.Location = new System.Drawing.Point(560, 69);
			this.button1.Name = "button1";
			this.button1.Size = new System.Drawing.Size(30, 23);
			this.button1.TabIndex = 755;
			this.button1.Click += new System.EventHandler(this.button1_Click);
			// 
			// groupBox1
			// 
			this.groupBox1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.groupBox1.Location = new System.Drawing.Point(0, 144);
			this.groupBox1.Name = "groupBox1";
			this.groupBox1.Size = new System.Drawing.Size(1184, 8);
			this.groupBox1.TabIndex = 756;
			this.groupBox1.TabStop = false;
			// 
			// label5
			// 
			this.label5.AutoSize = true;
			this.label5.Location = new System.Drawing.Point(8, 8);
			this.label5.Name = "label5";
			this.label5.Size = new System.Drawing.Size(43, 16);
			this.label5.TabIndex = 757;
			this.label5.Text = "Periodo";
			this.label5.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// uGridGastos
			// 
			this.uGridGastos.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left)));
			this.uGridGastos.Cursor = System.Windows.Forms.Cursors.Default;
			this.uGridGastos.DataSource = this.ultraDataSource1;
			appearance7.BackColor = System.Drawing.Color.White;
			this.uGridGastos.DisplayLayout.Appearance = appearance7;
			ultraGridColumn4.Header.VisiblePosition = 0;
			ultraGridColumn4.Hidden = true;
			ultraGridColumn5.Header.VisiblePosition = 1;
			ultraGridColumn5.Hidden = true;
			ultraGridColumn6.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn6.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn6.Header.Caption = "Descripción";
			ultraGridColumn6.Header.VisiblePosition = 2;
			ultraGridColumn6.Width = 110;
			appearance8.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn7.CellAppearance = appearance8;
			ultraGridColumn7.Format = "#,##0.00";
			appearance9.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn7.FormulaErrorAppearance = appearance9;
			ultraGridColumn7.Header.VisiblePosition = 3;
			appearance10.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn8.CellAppearance = appearance10;
			ultraGridColumn8.Format = "#,##0.00";
			appearance11.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn8.FormulaErrorAppearance = appearance11;
			ultraGridColumn8.Header.VisiblePosition = 4;
			appearance12.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn9.CellAppearance = appearance12;
			ultraGridColumn9.Format = "#,##0.00";
			appearance13.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn9.FormulaErrorAppearance = appearance13;
			ultraGridColumn9.Header.VisiblePosition = 5;
			appearance14.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn10.CellAppearance = appearance14;
			ultraGridColumn10.Format = "#,##0.00";
			appearance15.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn10.FormulaErrorAppearance = appearance15;
			ultraGridColumn10.Header.VisiblePosition = 6;
			appearance16.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn11.CellAppearance = appearance16;
			ultraGridColumn11.Format = "#,##0.00";
			appearance17.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn11.FormulaErrorAppearance = appearance17;
			ultraGridColumn11.Header.VisiblePosition = 7;
			appearance18.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn12.CellAppearance = appearance18;
			ultraGridColumn12.Format = "#,##0.00";
			appearance19.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn12.FormulaErrorAppearance = appearance19;
			ultraGridColumn12.Header.VisiblePosition = 8;
			appearance20.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn13.CellAppearance = appearance20;
			ultraGridColumn13.Format = "#,##0.00";
			appearance21.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn13.FormulaErrorAppearance = appearance21;
			ultraGridColumn13.Header.VisiblePosition = 9;
			appearance22.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn14.CellAppearance = appearance22;
			ultraGridColumn14.Format = "#,##0.00";
			appearance23.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn14.FormulaErrorAppearance = appearance23;
			ultraGridColumn14.Header.VisiblePosition = 10;
			appearance24.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn15.CellAppearance = appearance24;
			ultraGridColumn15.Format = "#,##0.00";
			appearance25.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn15.FormulaErrorAppearance = appearance25;
			ultraGridColumn15.Header.VisiblePosition = 11;
			appearance26.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn16.CellAppearance = appearance26;
			ultraGridColumn16.Format = "#,##0.00";
			appearance27.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn16.FormulaErrorAppearance = appearance27;
			ultraGridColumn16.Header.VisiblePosition = 12;
			appearance28.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn17.CellAppearance = appearance28;
			ultraGridColumn17.Format = "#,##0.00";
			appearance29.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn17.FormulaErrorAppearance = appearance29;
			ultraGridColumn17.Header.VisiblePosition = 13;
			appearance30.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn18.CellAppearance = appearance30;
			ultraGridColumn18.Format = "#,##0.00";
			appearance31.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn18.FormulaErrorAppearance = appearance31;
			ultraGridColumn18.Header.VisiblePosition = 14;
			ultraGridColumn19.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance32.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn19.CellAppearance = appearance32;
			ultraGridColumn19.Format = "#,##0.00";
			appearance33.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn19.FormulaErrorAppearance = appearance33;
			ultraGridColumn19.Header.VisiblePosition = 15;
			ultraGridBand2.Columns.AddRange(new object[] {
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
																										 ultraGridColumn18,
																										 ultraGridColumn19});
			ultraGridBand2.Header.Caption = "Ingresos";
			ultraGridBand2.HeaderVisible = true;
			this.uGridGastos.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			appearance34.ForeColor = System.Drawing.Color.Black;
			appearance34.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridGastos.DisplayLayout.Override.ActiveRowAppearance = appearance34;
			this.uGridGastos.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
			this.uGridGastos.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			appearance35.BackColor = System.Drawing.Color.Transparent;
			this.uGridGastos.DisplayLayout.Override.CardAreaAppearance = appearance35;
			appearance36.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance36.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance36.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance36.FontData.BoldAsString = "True";
			appearance36.FontData.Name = "Arial";
			appearance36.FontData.SizeInPoints = 8F;
			appearance36.ForeColor = System.Drawing.Color.White;
			appearance36.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridGastos.DisplayLayout.Override.HeaderAppearance = appearance36;
			this.uGridGastos.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortSingle;
			appearance37.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance37.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance37.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridGastos.DisplayLayout.Override.RowAlternateAppearance = appearance37;
			appearance38.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance38.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance38.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridGastos.DisplayLayout.Override.RowSelectorAppearance = appearance38;
			appearance39.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance39.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance39.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridGastos.DisplayLayout.Override.SelectedRowAppearance = appearance39;
			this.uGridGastos.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridGastos.Location = new System.Drawing.Point(8, 176);
			this.uGridGastos.Name = "uGridGastos";
			this.uGridGastos.Size = new System.Drawing.Size(1136, 240);
			this.uGridGastos.TabIndex = 758;
			this.uGridGastos.AfterCellUpdate += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.uGridGastos_AfterCellUpdate);
			// 
			// ultraDataSource1
			// 
			ultraDataColumn1.DataType = typeof(int);
			ultraDataColumn2.DataType = typeof(int);
			ultraDataColumn4.DataType = typeof(System.Decimal);
			ultraDataColumn5.DataType = typeof(System.Decimal);
			ultraDataColumn6.DataType = typeof(System.Decimal);
			ultraDataColumn7.DataType = typeof(System.Decimal);
			ultraDataColumn8.DataType = typeof(System.Decimal);
			ultraDataColumn9.DataType = typeof(System.Decimal);
			ultraDataColumn10.DataType = typeof(System.Decimal);
			ultraDataColumn11.DataType = typeof(System.Decimal);
			ultraDataColumn12.DataType = typeof(System.Decimal);
			ultraDataColumn13.DataType = typeof(System.Decimal);
			ultraDataColumn14.DataType = typeof(System.Decimal);
			ultraDataColumn15.DataType = typeof(System.Decimal);
			ultraDataColumn16.DataType = typeof(System.Decimal);
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
			// dTicketSillon
			// 
			appearance40.ForeColorDisabled = System.Drawing.Color.Black;
			this.dTicketSillon.Appearance = appearance40;
			this.dTicketSillon.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.dTicketSillon.FormatString = "#,##0.00";
			this.dTicketSillon.Location = new System.Drawing.Point(160, 120);
			this.dTicketSillon.Name = "dTicketSillon";
			this.dTicketSillon.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.dTicketSillon.PromptChar = ' ';
			this.dTicketSillon.Size = new System.Drawing.Size(80, 21);
			this.dTicketSillon.TabIndex = 760;
			// 
			// label6
			// 
			this.label6.AutoSize = true;
			this.label6.Location = new System.Drawing.Point(8, 120);
			this.label6.Name = "label6";
			this.label6.Size = new System.Drawing.Size(141, 16);
			this.label6.TabIndex = 759;
			this.label6.Text = "Ticket Promedio Por Sillon:";
			this.label6.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtidTicketPromedio
			// 
			this.txtidTicketPromedio.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.txtidTicketPromedio.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtidTicketPromedio.Enabled = false;
			this.txtidTicketPromedio.Location = new System.Drawing.Point(208, 8);
			this.txtidTicketPromedio.Name = "txtidTicketPromedio";
			this.txtidTicketPromedio.PromptChar = ' ';
			this.txtidTicketPromedio.Size = new System.Drawing.Size(104, 21);
			this.txtidTicketPromedio.TabIndex = 761;
			this.txtidTicketPromedio.Visible = false;
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
			// frmTicketPromedio
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(1152, 422);
			this.Controls.Add(this.txtidTicketPromedio);
			this.Controls.Add(this.dTicketSillon);
			this.Controls.Add(this.label6);
			this.Controls.Add(this.uGridGastos);
			this.Controls.Add(this.label5);
			this.Controls.Add(this.groupBox1);
			this.Controls.Add(this.button1);
			this.Controls.Add(this.txtTicket);
			this.Controls.Add(this.label4);
			this.Controls.Add(this.label53);
			this.Controls.Add(this.label3);
			this.Controls.Add(this.dtDesde);
			this.Controls.Add(this.dtHasta);
			this.Controls.Add(this.txtdTicketLocal);
			this.Controls.Add(this.txtdSillon);
			this.Controls.Add(this.label2);
			this.Controls.Add(this.label1);
			this.Controls.Add(this.lblBodega);
			this.Controls.Add(this.cmbBodega);
			this.Controls.Add(this.txtAnio);
			this.Controls.Add(this.btnVer);
			this.Name = "frmTicketPromedio";
			this.Text = "Ticket Promedio";
			this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.frmTicketPromedio_KeyDown);
			this.Load += new System.EventHandler(this.frmTicketPromedio_Load);
			((System.ComponentModel.ISupportInitialize)(this.txtAnio)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbBodega)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtdTicketLocal)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtdSillon)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtDesde)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtHasta)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtTicket)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridGastos)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dTicketSillon)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtidTicketPromedio)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void frmTicketPromedio_Load(object sender, System.EventArgs e)
		{

				FuncionesProcedimientos.ListaBodegas(this.cmbBodega, MenuLatinium.IdUsuario, 42, DateTime.Today, false, cdsSeteoF);	
			this.txtAnio.MinValue = 2021;
			this.txtAnio.MaxValue = DateTime.Today.Year + 1;
			this.txtAnio.Value = DateTime.Today.Year;

		}

		private void dtHasta_BeforeDropDown(object sender, System.ComponentModel.CancelEventArgs e)
		{
		
		}

		private void button1_Click(object sender, System.EventArgs e)
		{
			#region Valida Fechas
			if (this.dtDesde.Value == System.DBNull.Value)
			{
				MessageBox.Show("Seleccione la Fecha Inicial", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.dtDesde.Focus();
				return;
			}

			if (this.dtHasta.Value == System.DBNull.Value)
			{
				MessageBox.Show("Seleccione la Fecha Final", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.dtHasta.Focus();
				return;
			}

			if (Convert.ToDateTime(this.dtDesde.Value) > Convert.ToDateTime(this.dtHasta.Value))
			{
				MessageBox.Show("La Fecha Inicial no puede ser mayor a la Fecha Final", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.dtHasta.Focus();
				return;
			}
			#endregion Valida Fechas

			if (!Validacion.vbComboVacio(this.cmbBodega, "Seleccione el Almacen")) return;

			this.txtTicket.Value = Funcion.decEscalarSQL(cdsSeteoF, string.Format(" Exec [CalculaVentaclinica] '{0}', '{1}', {2}",
				 Convert.ToDateTime(this.dtDesde.Value).ToString("yyyyMMdd"),  Convert.ToDateTime(this.dtHasta.Value).ToString("yyyyMMdd"), (int) this.cmbBodega.Value));		
     if ((decimal) this.txtTicket.Value > 0) this.txtdTicketLocal.Value = this.txtTicket.Value;

			if (Convert.ToDecimal (this.txtdSillon.Value) > 0.00m)
			{
				this.dTicketSillon.Value = Convert.ToDecimal (this.txtdTicketLocal.Value) /Convert.ToDecimal (this.txtdSillon.Value);
			}

			this.dTicketSillon.Value = Funcion.decEscalarSQL(cdsSeteoF, string.Format(" Exec [CalculaVentaclinicaSillones] '{0}', '{1}', {2}, {3}",
				Convert.ToDateTime(this.dtDesde.Value).ToString("yyyyMMdd"),  Convert.ToDateTime(this.dtHasta.Value).ToString("yyyyMMdd"), (int) this.cmbBodega.Value, Convert.ToDecimal(this.txtdSillon.Value)));	
		}

		private void btnVer_Click(object sender, System.EventArgs e)
		{
			if (!Validacion.vbCampoEnteroVacio(this.txtAnio, "Ingrese el periodo", 2020, DateTime.Today.Year + 1)) return;
			if (!Validacion.vbComboVacio(this.cmbBodega, "Seleccione el Almacen")) return;

			int iCount = 0;
			iCount = Funcion.iEscalarSQL(cdsSeteoF, string.Format(" Exec [ValidaTicket] {0}, {1}, {2}, {3}",
				(int) this.txtAnio.Value,  (int) this.cmbBodega.Value, 0, (int) this.txtidTicketPromedio.Value));		
     
			if (iCount > 0)
			{
				string sSQL = string.Format("Exec ValidaTicket {0}, {1}, {2}, {3}", (int) this.txtAnio.Value,  (int) this.cmbBodega.Value, 1, iCount);
				SqlDataReader drTransferencia = Funcion.rEscalarSQL(cdsSeteoF, sSQL, true);
				drTransferencia.Read();
				if (drTransferencia.HasRows)
				{					
					this.txtidTicketPromedio.Value = (int)drTransferencia.GetValue(0);
					this.txtAnio.Value = (int)drTransferencia.GetValue(1);
					this.cmbBodega.Value = (int)drTransferencia.GetValue(2);
					this.txtdSillon.Value = (int)drTransferencia.GetDecimal(3);
					this.txtdTicketLocal.Value = (int)drTransferencia.GetDecimal(4);
					this.dTicketSillon.Value = (int)drTransferencia.GetDecimal(5);
				}
				drTransferencia.Close();

				sSQL = string.Format("Exec ValidaTicket {0}, {1}, {2}, {3}", (int) this.txtAnio.Value,  (int) this.cmbBodega.Value, 2, (int) this.txtidTicketPromedio.Value);
				FuncionesProcedimientos.dsGeneral(sSQL, this.uGridGastos);
			}
			else
			{
				if (DialogResult.Yes == MessageBox.Show("¿Desea Generar un Periodo Nuevo?", "Point Technology", MessageBoxButtons.YesNo, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2))
				{
					string sSQL = string.Format("Exec [GeneraTicket] {0},{1}",
						(int) this.txtAnio.Value, (int) this.cmbBodega.Value );
					this.txtidTicketPromedio.Value = Funcion.iEscalarSQL(cdsSeteoF,sSQL);
					this.btnVer_Click (sender,e);
				}
			}

		}

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);		
	
		}

		private void cmbBodega_ValueChanged(object sender, System.EventArgs e)
		{
			 this.txtidTicketPromedio.Value = 0;
				this.btnVer_Click (sender,e);
		}

		private void uGridGastos_AfterCellUpdate(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			#region Actualiza Valores
			if (e.Cell.Column.ToString() == "ENERO")
			{
				ActualizaValores ( Convert.ToInt32( e.Cell.Row.Cells["idTicketPromedio"].Value) ,
					e.Cell.Row.Cells["Codigo"].Value.ToString(), "ENERO", Convert.ToDecimal(e.Cell.Value));
				Sumatorias(e);	

			}

			if (e.Cell.Column.ToString() == "FEBRERO")
			{
				ActualizaValores ( Convert.ToInt32( e.Cell.Row.Cells["idTicketPromedio"].Value) ,
					e.Cell.Row.Cells["Codigo"].Value.ToString(), "FEBRERO", Convert.ToDecimal(e.Cell.Value));
						Sumatorias(e);	
			}

			if (e.Cell.Column.ToString() == "MARZO")
			{
				ActualizaValores ( Convert.ToInt32( e.Cell.Row.Cells["idTicketPromedio"].Value) ,
					e.Cell.Row.Cells["Codigo"].Value.ToString(), "MARZO", Convert.ToDecimal(e.Cell.Value));
					Sumatorias(e);
			}

			if (e.Cell.Column.ToString() == "ABRIL")
			{
				ActualizaValores ( Convert.ToInt32( e.Cell.Row.Cells["idTicketPromedio"].Value) ,
					e.Cell.Row.Cells["Codigo"].Value.ToString(), "ABRIL", Convert.ToDecimal(e.Cell.Value));
					Sumatorias(e);
			}

			if (e.Cell.Column.ToString() == "MAYO")
			{
				ActualizaValores ( Convert.ToInt32( e.Cell.Row.Cells["idTicketPromedio"].Value) ,
					e.Cell.Row.Cells["Codigo"].Value.ToString(), "MAYO", Convert.ToDecimal(e.Cell.Value));
					Sumatorias(e);
			}

			if (e.Cell.Column.ToString() == "JUNIO")
			{
				ActualizaValores ( Convert.ToInt32( e.Cell.Row.Cells["idTicketPromedio"].Value) ,
					e.Cell.Row.Cells["Codigo"].Value.ToString(), "JUNIO", Convert.ToDecimal(e.Cell.Value));
					Sumatorias(e);
			}

			if (e.Cell.Column.ToString() == "JULIO")
			{
				ActualizaValores ( Convert.ToInt32( e.Cell.Row.Cells["idTicketPromedio"].Value) ,
					e.Cell.Row.Cells["Codigo"].Value.ToString(), "JULIO", Convert.ToDecimal(e.Cell.Value));
					Sumatorias(e);
			}

			if (e.Cell.Column.ToString() == "AGOSTO")
			{
				ActualizaValores ( Convert.ToInt32( e.Cell.Row.Cells["idTicketPromedio"].Value) ,
					e.Cell.Row.Cells["Codigo"].Value.ToString(), "AGOSTO", Convert.ToDecimal(e.Cell.Value));
					Sumatorias(e);
			}

			if (e.Cell.Column.ToString() == "SEPTIEMBRE")
			{
				ActualizaValores ( Convert.ToInt32( e.Cell.Row.Cells["idTicketPromedio"].Value) ,
					e.Cell.Row.Cells["Codigo"].Value.ToString(), "SEPTIEMBRE", Convert.ToDecimal(e.Cell.Value));
					Sumatorias(e);
			}

			if (e.Cell.Column.ToString() == "OCTUBRE")
			{
				ActualizaValores ( Convert.ToInt32( e.Cell.Row.Cells["idTicketPromedio"].Value) ,
					e.Cell.Row.Cells["Codigo"].Value.ToString(), "OCTUBRE", Convert.ToDecimal(e.Cell.Value));
					Sumatorias(e);
			}

			if (e.Cell.Column.ToString() == "NOVIEMBRE")
			{
				ActualizaValores ( Convert.ToInt32( e.Cell.Row.Cells["idTicketPromedio"].Value) ,
					e.Cell.Row.Cells["Codigo"].Value.ToString(), "NOVIEMBRE", Convert.ToDecimal(e.Cell.Value));
					Sumatorias(e);
			}

			if (e.Cell.Column.ToString() == "DICIEMBRE")
			{
				ActualizaValores ( Convert.ToInt32( e.Cell.Row.Cells["idTicketPromedio"].Value) ,
					e.Cell.Row.Cells["Codigo"].Value.ToString(), "DICIEMBRE", Convert.ToDecimal(e.Cell.Value));
					Sumatorias(e);
			}
			#endregion Actualiza Valores

			string sSQL = string.Format("Exec ValidaTicket {0}, {1}, {2}, {3}", (int) this.txtAnio.Value,  (int) this.cmbBodega.Value, 2, (int) this.txtidTicketPromedio.Value);
			FuncionesProcedimientos.dsGeneral(sSQL, this.uGridGastos);

		}

		private void ActualizaValores (int idTicketPromedio ,	string Codigo , string Campo , decimal Valor )
		{
			string sSQL = string.Format("Exec [ActualizaTicket] {0}, '{1}', '{2}', {3}, {4}, {5}, {6}", 
				Convert.ToInt32(this.txtidTicketPromedio.Value), Codigo, Campo, Valor, Convert.ToDecimal(this.txtdSillon.Value), Convert.ToDecimal(this.txtdTicketLocal.Value), Convert.ToDecimal( this.dTicketSillon.Value));
			Funcion.EjecutaSQL(cdsSeteoF, sSQL);
			
		}

		private void Sumatorias(Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			e.Cell.Row.Cells["Total"].Value = 
				Convert.ToDecimal(e.Cell.Row.Cells["ENERO"].Value) + 
				Convert.ToDecimal(e.Cell.Row.Cells["FEBRERO"].Value) +
				Convert.ToDecimal(e.Cell.Row.Cells["MARZO"].Value) +
				Convert.ToDecimal(e.Cell.Row.Cells["ABRIL"].Value) + 
				Convert.ToDecimal(e.Cell.Row.Cells["MAYO"].Value) +
				Convert.ToDecimal(e.Cell.Row.Cells["JUNIO"].Value) +
				Convert.ToDecimal(e.Cell.Row.Cells["JULIO"].Value) + 
				Convert.ToDecimal(e.Cell.Row.Cells["AGOSTO"].Value) +
				Convert.ToDecimal(e.Cell.Row.Cells["SEPTIEMBRE"].Value) +
				Convert.ToDecimal(e.Cell.Row.Cells["OCTUBRE"].Value) + 
				Convert.ToDecimal(e.Cell.Row.Cells["NOVIEMBRE"].Value) +
				Convert.ToDecimal(e.Cell.Row.Cells["DICIEMBRE"].Value);			
		}

		private void frmTicketPromedio_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
		if (e.KeyCode == Keys.Escape) this.Close();
		}

	}
}
