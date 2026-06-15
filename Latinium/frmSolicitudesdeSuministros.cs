using System;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de frmSolicitudesdeSuministros.
	/// </summary>
	public class frmSolicitudesdeSuministros : System.Windows.Forms.Form
	{
		private System.Windows.Forms.Label label3;
		private System.Windows.Forms.Label label1;
		private System.Windows.Forms.Button btnVer;
		private Infragistics.Win.UltraWinEditors.UltraOptionSet optEstado;
		private System.Windows.Forms.Label lblEstado;
		private System.Windows.Forms.Button btnExportar;
		private C1.Data.C1DataSet cdsSeteoF;
		private System.Windows.Forms.Label lblComprobante;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		
		int IdCompra = 0;	
		bool solicitud;
		DateTime DtpFecha = DateTime.Today;
		int idBloqueaTransacciones = 0;

		private Infragistics.Win.UltraWinGrid.UltraCombo cmbBodega;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtDesde;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtHasta;
		public Infragistics.Win.UltraWinEditors.UltraNumericEditor txtEstado;
		public Infragistics.Win.UltraWinEditors.UltraNumericEditor txtidCompra;
		public Infragistics.Win.UltraWinEditors.UltraNumericEditor txtidPersonal;
		public Infragistics.Win.UltraWinGrid.UltraGrid ultraGrid1;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource1;
		private System.Windows.Forms.GroupBox groupBox1;
		bool bLimiteFechas = true;
		private Acceso miAcceso;

		public frmSolicitudesdeSuministros()
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
			Infragistics.Win.Appearance appearance1 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton1 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmSolicitudesdeSuministros));
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton2 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.ValueListItem valueListItem1 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem2 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem3 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem4 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem5 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem6 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idProyecto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idRequerimientoSuministros");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idProyecto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCompra");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Estado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Fecha");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Numero");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Solicitante");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CentroCosto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn11 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bodega");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn12 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaAprobacion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn13 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("UsuarioAProbo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn14 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("NotasAprobacion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn15 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Aceptar", 0);
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn16 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Ir", 1);
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance11 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance12 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance13 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance14 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance15 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idRequerimientoSuministros");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idProyecto");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn3 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCompra");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn4 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Estado");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn5 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Fecha");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn6 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Numero");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn7 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Solicitante");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn8 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("CentroCosto");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn9 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Bodega");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn10 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("FechaAprobacion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn11 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("UsuarioAProbo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn12 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("NotasAprobacion");
			this.label3 = new System.Windows.Forms.Label();
			this.dtDesde = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.dtHasta = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.label1 = new System.Windows.Forms.Label();
			this.btnVer = new System.Windows.Forms.Button();
			this.optEstado = new Infragistics.Win.UltraWinEditors.UltraOptionSet();
			this.lblEstado = new System.Windows.Forms.Label();
			this.btnExportar = new System.Windows.Forms.Button();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.lblComprobante = new System.Windows.Forms.Label();
			this.cmbBodega = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.txtEstado = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtidCompra = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtidPersonal = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.ultraGrid1 = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource1 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.groupBox1 = new System.Windows.Forms.GroupBox();
			((System.ComponentModel.ISupportInitialize)(this.dtDesde)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtHasta)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.optEstado)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbBodega)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtEstado)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtidCompra)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtidPersonal)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).BeginInit();
			this.SuspendLayout();
			// 
			// label3
			// 
			this.label3.AutoSize = true;
			this.label3.Location = new System.Drawing.Point(8, 10);
			this.label3.Name = "label3";
			this.label3.Size = new System.Drawing.Size(35, 17);
			this.label3.TabIndex = 759;
			this.label3.Text = "Desde";
			this.label3.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// dtDesde
			// 
			appearance1.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtDesde.Appearance = appearance1;
			this.dtDesde.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton1.Caption = "Today";
			this.dtDesde.DateButtons.Add(dateButton1);
			this.dtDesde.Format = "dd/MM/yyyy";
			this.dtDesde.Location = new System.Drawing.Point(96, 8);
			this.dtDesde.MonthOrientation = Infragistics.Win.UltraWinSchedule.MonthOrientation.DownThenAcross;
			this.dtDesde.Name = "dtDesde";
			this.dtDesde.NonAutoSizeHeight = 23;
			this.dtDesde.Size = new System.Drawing.Size(112, 21);
			this.dtDesde.SpinButtonsVisible = true;
			this.dtDesde.TabIndex = 760;
			this.dtDesde.Value = ((object)(resources.GetObject("dtDesde.Value")));
			// 
			// dtHasta
			// 
			appearance2.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtHasta.Appearance = appearance2;
			this.dtHasta.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton2.Caption = "Today";
			this.dtHasta.DateButtons.Add(dateButton2);
			this.dtHasta.Format = "dd/MM/yyyy";
			this.dtHasta.Location = new System.Drawing.Point(288, 8);
			this.dtHasta.MonthOrientation = Infragistics.Win.UltraWinSchedule.MonthOrientation.DownThenAcross;
			this.dtHasta.Name = "dtHasta";
			this.dtHasta.NonAutoSizeHeight = 23;
			this.dtHasta.Size = new System.Drawing.Size(112, 21);
			this.dtHasta.SpinButtonsVisible = true;
			this.dtHasta.TabIndex = 758;
			this.dtHasta.Value = ((object)(resources.GetObject("dtHasta.Value")));
			// 
			// label1
			// 
			this.label1.AutoSize = true;
			this.label1.Location = new System.Drawing.Point(232, 10);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(33, 17);
			this.label1.TabIndex = 762;
			this.label1.Text = "Hasta";
			this.label1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// btnVer
			// 
			this.btnVer.CausesValidation = false;
			this.btnVer.Image = ((System.Drawing.Image)(resources.GetObject("btnVer.Image")));
			this.btnVer.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnVer.Location = new System.Drawing.Point(528, 73);
			this.btnVer.Name = "btnVer";
			this.btnVer.Size = new System.Drawing.Size(72, 24);
			this.btnVer.TabIndex = 763;
			this.btnVer.Text = "Ver";
			this.btnVer.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnVer.Click += new System.EventHandler(this.btnVer_Click);
			// 
			// optEstado
			// 
			appearance3.ForeColorDisabled = System.Drawing.Color.Black;
			this.optEstado.Appearance = appearance3;
			this.optEstado.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.optEstado.CheckedIndex = 1;
			this.optEstado.ItemAppearance = appearance4;
			this.optEstado.ItemOrigin = new System.Drawing.Point(10, 0);
			valueListItem1.DataValue = 0;
			valueListItem1.DisplayText = "Todos";
			valueListItem2.DataValue = 1;
			valueListItem2.DisplayText = "Pendiente";
			valueListItem3.DataValue = 5;
			valueListItem3.DisplayText = "Aprobado";
			valueListItem4.DataValue = 3;
			valueListItem4.DisplayText = "Facturado";
			valueListItem5.DataValue = 4;
			valueListItem5.DisplayText = "Rechazado";
			valueListItem6.DataValue = 2;
			valueListItem6.DisplayText = "Anulado";
			this.optEstado.Items.Add(valueListItem1);
			this.optEstado.Items.Add(valueListItem2);
			this.optEstado.Items.Add(valueListItem3);
			this.optEstado.Items.Add(valueListItem4);
			this.optEstado.Items.Add(valueListItem5);
			this.optEstado.Items.Add(valueListItem6);
			this.optEstado.ItemSpacingVertical = 10;
			this.optEstado.Location = new System.Drawing.Point(8, 72);
			this.optEstado.Name = "optEstado";
			this.optEstado.Size = new System.Drawing.Size(464, 26);
			this.optEstado.TabIndex = 762;
			this.optEstado.Text = "Pendiente";
			this.optEstado.ValueChanged += new System.EventHandler(this.optEstado_ValueChanged);
			// 
			// lblEstado
			// 
			this.lblEstado.AutoSize = true;
			this.lblEstado.Font = new System.Drawing.Font("Tahoma", 11F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblEstado.ForeColor = System.Drawing.Color.Firebrick;
			this.lblEstado.Location = new System.Drawing.Point(656, 75);
			this.lblEstado.Name = "lblEstado";
			this.lblEstado.Size = new System.Drawing.Size(0, 21);
			this.lblEstado.TabIndex = 764;
			this.lblEstado.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// btnExportar
			// 
			this.btnExportar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btnExportar.CausesValidation = false;
			this.btnExportar.Image = ((System.Drawing.Image)(resources.GetObject("btnExportar.Image")));
			this.btnExportar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnExportar.Location = new System.Drawing.Point(976, 74);
			this.btnExportar.Name = "btnExportar";
			this.btnExportar.Size = new System.Drawing.Size(88, 23);
			this.btnExportar.TabIndex = 765;
			this.btnExportar.Text = "Exportar";
			this.btnExportar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnExportar.Click += new System.EventHandler(this.btnExportar_Click);
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
			// lblComprobante
			// 
			this.lblComprobante.AutoSize = true;
			this.lblComprobante.BackColor = System.Drawing.Color.Transparent;
			this.lblComprobante.ImeMode = System.Windows.Forms.ImeMode.NoControl;
			this.lblComprobante.Location = new System.Drawing.Point(8, 42);
			this.lblComprobante.Name = "lblComprobante";
			this.lblComprobante.Size = new System.Drawing.Size(85, 17);
			this.lblComprobante.TabIndex = 767;
			this.lblComprobante.Text = "Centro de Costo";
			// 
			// cmbBodega
			// 
			appearance5.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbBodega.Appearance = appearance5;
			this.cmbBodega.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbBodega.Cursor = System.Windows.Forms.Cursors.Default;
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridColumn2.Width = 320;
			ultraGridBand1.Columns.AddRange(new object[] {
																										 ultraGridColumn1,
																										 ultraGridColumn2});
			ultraGridBand1.SummaryFooterCaption = "Grand Summaries";
			this.cmbBodega.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.cmbBodega.DisplayMember = "Nombre";
			this.cmbBodega.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbBodega.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbBodega.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F);
			this.cmbBodega.Location = new System.Drawing.Point(112, 40);
			this.cmbBodega.Name = "cmbBodega";
			this.cmbBodega.Size = new System.Drawing.Size(288, 21);
			this.cmbBodega.TabIndex = 768;
			this.cmbBodega.ValueMember = "idProyecto";
			// 
			// txtEstado
			// 
			this.txtEstado.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtEstado.Enabled = false;
			this.txtEstado.Location = new System.Drawing.Point(448, 7);
			this.txtEstado.Name = "txtEstado";
			this.txtEstado.PromptChar = ' ';
			this.txtEstado.Size = new System.Drawing.Size(40, 22);
			this.txtEstado.SpinButtonAlignment = Infragistics.Win.ButtonAlignment.Left;
			this.txtEstado.TabIndex = 812;
			// 
			// txtidCompra
			// 
			this.txtidCompra.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtidCompra.Enabled = false;
			this.txtidCompra.Location = new System.Drawing.Point(496, 7);
			this.txtidCompra.Name = "txtidCompra";
			this.txtidCompra.PromptChar = ' ';
			this.txtidCompra.Size = new System.Drawing.Size(40, 22);
			this.txtidCompra.SpinButtonAlignment = Infragistics.Win.ButtonAlignment.Left;
			this.txtidCompra.TabIndex = 813;
			// 
			// txtidPersonal
			// 
			this.txtidPersonal.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtidPersonal.Enabled = false;
			this.txtidPersonal.Location = new System.Drawing.Point(544, 7);
			this.txtidPersonal.Name = "txtidPersonal";
			this.txtidPersonal.PromptChar = ' ';
			this.txtidPersonal.Size = new System.Drawing.Size(40, 22);
			this.txtidPersonal.SpinButtonAlignment = Infragistics.Win.ButtonAlignment.Left;
			this.txtidPersonal.TabIndex = 815;
			// 
			// ultraGrid1
			// 
			this.ultraGrid1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.ultraGrid1.Cursor = System.Windows.Forms.Cursors.Default;
			this.ultraGrid1.DataSource = this.ultraDataSource1;
			appearance6.BackColor = System.Drawing.Color.White;
			this.ultraGrid1.DisplayLayout.Appearance = appearance6;
			ultraGridColumn3.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn3.Header.VisiblePosition = 0;
			ultraGridColumn3.Hidden = true;
			ultraGridColumn3.Width = 167;
			ultraGridColumn4.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn4.Header.VisiblePosition = 1;
			ultraGridColumn4.Hidden = true;
			ultraGridColumn5.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn5.Header.VisiblePosition = 2;
			ultraGridColumn5.Hidden = true;
			ultraGridColumn6.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn6.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn6.Header.VisiblePosition = 5;
			ultraGridColumn6.Width = 100;
			ultraGridColumn7.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn7.Format = "dd/MM/yyyy HH:mm";
			ultraGridColumn7.Header.Caption = "Fecha Solicitud";
			ultraGridColumn7.Header.VisiblePosition = 6;
			ultraGridColumn7.Width = 120;
			ultraGridColumn8.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn8.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn8.Header.Caption = "Número";
			ultraGridColumn8.Header.VisiblePosition = 9;
			ultraGridColumn8.Width = 100;
			ultraGridColumn9.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn9.Header.VisiblePosition = 7;
			ultraGridColumn10.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn10.Header.Caption = "Centro De Costo";
			ultraGridColumn10.Header.VisiblePosition = 8;
			ultraGridColumn10.Width = 140;
			ultraGridColumn11.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn11.Header.VisiblePosition = 10;
			ultraGridColumn11.Hidden = true;
			ultraGridColumn11.Width = 119;
			ultraGridColumn12.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn12.Format = "dd/MM/yyyy HH:mm";
			ultraGridColumn12.Header.Caption = "Fecha Aprobación";
			ultraGridColumn12.Header.VisiblePosition = 11;
			ultraGridColumn12.Width = 120;
			ultraGridColumn13.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn13.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn13.Header.Caption = "Aprobado Por";
			ultraGridColumn13.Header.VisiblePosition = 12;
			ultraGridColumn13.Width = 120;
			ultraGridColumn14.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn14.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn14.Header.Caption = "Notas Aprobación";
			ultraGridColumn14.Header.VisiblePosition = 13;
			ultraGridColumn14.Width = 300;
			ultraGridColumn15.ButtonDisplayStyle = Infragistics.Win.UltraWinGrid.ButtonDisplayStyle.Always;
			appearance7.BackColorAlpha = Infragistics.Win.Alpha.Opaque;
			appearance7.Image = ((object)(resources.GetObject("appearance7.Image")));
			appearance7.ImageHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn15.CellButtonAppearance = appearance7;
			ultraGridColumn15.Header.Caption = "...";
			ultraGridColumn15.Header.VisiblePosition = 4;
			ultraGridColumn15.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.Button;
			ultraGridColumn15.Width = 23;
			ultraGridColumn16.ButtonDisplayStyle = Infragistics.Win.UltraWinGrid.ButtonDisplayStyle.Always;
			appearance8.BackColorAlpha = Infragistics.Win.Alpha.Opaque;
			appearance8.Image = ((object)(resources.GetObject("appearance8.Image")));
			appearance8.ImageHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn16.CellButtonAppearance = appearance8;
			ultraGridColumn16.Header.Caption = "Ver";
			ultraGridColumn16.Header.VisiblePosition = 3;
			ultraGridColumn16.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.Button;
			ultraGridColumn16.Width = 24;
			ultraGridBand2.Columns.AddRange(new object[] {
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
			this.ultraGrid1.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			appearance9.ForeColor = System.Drawing.Color.Black;
			appearance9.ForeColorDisabled = System.Drawing.Color.Black;
			this.ultraGrid1.DisplayLayout.Override.ActiveRowAppearance = appearance9;
			this.ultraGrid1.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
			this.ultraGrid1.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.ultraGrid1.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.False;
			appearance10.BackColor = System.Drawing.Color.Transparent;
			this.ultraGrid1.DisplayLayout.Override.CardAreaAppearance = appearance10;
			appearance11.ForeColor = System.Drawing.Color.Black;
			appearance11.ForeColorDisabled = System.Drawing.Color.Black;
			this.ultraGrid1.DisplayLayout.Override.CellAppearance = appearance11;
			this.ultraGrid1.DisplayLayout.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.RowSelect;
			appearance12.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance12.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance12.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance12.FontData.BoldAsString = "True";
			appearance12.FontData.Name = "Arial";
			appearance12.FontData.SizeInPoints = 9F;
			appearance12.ForeColor = System.Drawing.Color.White;
			appearance12.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.ultraGrid1.DisplayLayout.Override.HeaderAppearance = appearance12;
			appearance13.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance13.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance13.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.RowAlternateAppearance = appearance13;
			appearance14.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance14.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance14.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.RowSelectorAppearance = appearance14;
			appearance15.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance15.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance15.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.SelectedRowAppearance = appearance15;
			this.ultraGrid1.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.ultraGrid1.Location = new System.Drawing.Point(8, 128);
			this.ultraGrid1.Name = "ultraGrid1";
			this.ultraGrid1.Size = new System.Drawing.Size(1056, 232);
			this.ultraGrid1.TabIndex = 819;
			this.ultraGrid1.ClickCellButton += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.ultraGrid1_ClickCellButton);
			this.ultraGrid1.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.ultraGrid1_InitializeLayout);
			// 
			// ultraDataSource1
			// 
			ultraDataColumn1.DataType = typeof(int);
			ultraDataColumn2.DataType = typeof(int);
			ultraDataColumn3.DataType = typeof(int);
			ultraDataColumn5.DataType = typeof(System.DateTime);
			ultraDataColumn9.DataType = typeof(int);
			ultraDataColumn10.DataType = typeof(System.DateTime);
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
																																 ultraDataColumn12});
			// 
			// groupBox1
			// 
			this.groupBox1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.groupBox1.Location = new System.Drawing.Point(0, 112);
			this.groupBox1.Name = "groupBox1";
			this.groupBox1.Size = new System.Drawing.Size(1072, 8);
			this.groupBox1.TabIndex = 820;
			this.groupBox1.TabStop = false;
			// 
			// frmSolicitudesdeSuministros
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 14);
			this.ClientSize = new System.Drawing.Size(1072, 366);
			this.Controls.Add(this.groupBox1);
			this.Controls.Add(this.ultraGrid1);
			this.Controls.Add(this.txtidPersonal);
			this.Controls.Add(this.txtidCompra);
			this.Controls.Add(this.txtEstado);
			this.Controls.Add(this.lblComprobante);
			this.Controls.Add(this.label1);
			this.Controls.Add(this.label3);
			this.Controls.Add(this.cmbBodega);
			this.Controls.Add(this.btnExportar);
			this.Controls.Add(this.lblEstado);
			this.Controls.Add(this.dtDesde);
			this.Controls.Add(this.dtHasta);
			this.Controls.Add(this.optEstado);
			this.Controls.Add(this.btnVer);
			this.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.KeyPreview = true;
			this.Name = "frmSolicitudesdeSuministros";
			this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
			this.Text = "Aprobación De Suministros";
			this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.frmSolicitudesdeSuministros_KeyDown);
			this.Load += new System.EventHandler(this.frmSolicitudesdeSuministros_Load);
			((System.ComponentModel.ISupportInitialize)(this.dtDesde)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtHasta)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.optEstado)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbBodega)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtEstado)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtidCompra)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtidPersonal)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void UnloadMe()
		{
			this.Close();
		}

		private void frmSolicitudesdeSuministros_Load(object sender, System.EventArgs e)
		{	
			miAcceso = new Acceso(cdsSeteoF, "0223");

			if (!Funcion.Permiso("1278", cdsSeteoF))
			{				
				MessageBox.Show("No puede ingresar a Aprobacion de Suministros", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				BeginInvoke(new MethodInvoker(UnloadMe));
				return;
			}

			this.dtDesde.Value = DateTime.Parse("01/01/2015");			
			this.dtHasta.Value = DateTime.Today;
			this.cmbBodega.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Select idProyecto, Nombre From Proyecto Where Visible = 1 Order by Nombre");
			CargaLista();		
		}

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);	
		}

		private void btnVer_Click(object sender, System.EventArgs e)
		{
			CargaLista();
		}

		private  void CargaLista()
		{
			int iBodega = 0;			

      if (!Validacion.vbComparaFechas(Convert.ToDateTime(this.dtDesde.Value), Convert.ToDateTime(this.dtHasta.Value), this.dtDesde, "La fecha inicial no puede ser posterior a la fecha final", false)) return;

			Estados();

			if (this.cmbBodega.ActiveRow != null) iBodega = (int)this.cmbBodega.Value;

			string sSQL = string.Format("Exec ListaSolicitudesSuministros '{0}', '{1}', {2}, {3}", 
				Convert.ToDateTime(this.dtDesde.Value).ToString("yyyyMMdd"), 
				Convert.ToDateTime(this.dtHasta.Value).ToString("yyyyMMdd"),
				iBodega, 
				(int)this.optEstado.Value);
			this.ultraGrid1.DataSource = Funcion.dvProcedimiento(cdsSeteoF, sSQL);

			if (this.ultraGrid1.Rows.Count > 0) this.ultraGrid1.Focus();
			else this.optEstado.Focus();

			this.lblEstado.Text = this.ultraGrid1.Rows.Count +" "+ " Registros Encontrados";		
		}

		private void Estados(){
			if((int)this.optEstado.Value == 0){
				this.ultraGrid1.DisplayLayout.Bands[0].Columns["Aceptar"].Hidden = true;
				this.ultraGrid1.DisplayLayout.Bands[0].Columns["Ir"].Hidden = true;
			}	
			if((int)this.optEstado.Value == 1)
			{
				this.ultraGrid1.DisplayLayout.Bands[0].Columns["Aceptar"].Hidden = false;
				this.ultraGrid1.DisplayLayout.Bands[0].Columns["Ir"].Hidden = false;
			}	
			if((int)this.optEstado.Value == 2)
			{
				this.ultraGrid1.DisplayLayout.Bands[0].Columns["Aceptar"].Hidden = true;
				this.ultraGrid1.DisplayLayout.Bands[0].Columns["Ir"].Hidden = false;
			}	
			if((int)this.optEstado.Value ==4){
				this.ultraGrid1.DisplayLayout.Bands[0].Columns["Aceptar"].Hidden = true;
				this.ultraGrid1.DisplayLayout.Bands[0].Columns["Ir"].Hidden = false;
			}
			if((int)this.optEstado.Value == 5)
			{
				this.ultraGrid1.DisplayLayout.Bands[0].Columns["Aceptar"].Hidden = true;
				this.ultraGrid1.DisplayLayout.Bands[0].Columns["Ir"].Hidden = false;
			}
		}

		private void optEstado_ValueChanged(object sender, System.EventArgs e)
		{
			this.btnVer_Click (sender,  e);
			CargaLista();
			Estados ();
		}

		private void ultraGrid1_ClickCellButton(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			if (e.Cell.Column.ToString() == "Aceptar")
			{	
				//if (Funcion.ValidaBloqueo((int)(int)e.Cell.Row.Cells["idRequerimientoSuministros"].Value, 1, cdsSeteoF)) return;

//				string sSQLBloqueo = string.Format("Exec RegistraBloqueoTransacciones 0, 1, {0}, 1", (int)e.Cell.Row.Cells["idRequerimientoSuministros"].Value);
//				idBloqueaTransacciones = Funcion.iEscalarSQL(cdsSeteoF, sSQLBloqueo, true);

				using (frmNAprobacion AP  = new frmNAprobacion((int)e.Cell.Row.Cells["idRequerimientoSuministros"].Value))
				{
					if (DialogResult.OK == AP.ShowDialog())
					{						
						string sSQL = string.Format("Update RequerimientoSuministros Set NotasAprobacion = NotasAprobacion + ' ' + '{0}', Estado = {1} , UsuarioAprobo = SUSER_SNAME(), FechaAprobacion = '{2}' Where idRequerimientoSuministros = {3}" ,
							AP.txtNotas.Text, (int)AP.optEstado.Value, Convert.ToDateTime(DtpFecha).ToString("yyyyMMdd") ,(int)e.Cell.Row.Cells["idRequerimientoSuministros"].Value);
						Funcion.EjecutaSQL(cdsSeteoF, sSQL);
												
						this.btnVer_Click(sender, e);
					}
				}
			}

			if (e.Cell.Column.ToString() == "Ir")
			{	
				using (frmRequerimientosDeSuministros1 SS  = new frmRequerimientosDeSuministros1((int)e.Cell.Row.Cells["idRequerimientoSuministros"].Value))
				{
					if (DialogResult.OK == SS.ShowDialog())
					{
						
					}
				}
			}
		}

		private void frmSolicitudesdeSuministros_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Escape) this.Close();	
		}

		private void btnExportar_Click(object sender, System.EventArgs e)
		{
			FuncionesProcedimientos.ExportaExcel(this.ultraGrid1);
		}

		private void ultraGrid1_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}
	}
}
