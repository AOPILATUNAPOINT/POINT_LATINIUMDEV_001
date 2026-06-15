using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de frmReporteAprobacionODC.
	/// </summary>
	public class frmReporteAprobacionODC : System.Windows.Forms.Form
	{
		private System.Windows.Forms.GroupBox groupBox1;
		private Infragistics.Win.UltraWinEditors.UltraOptionSet optOrigen;
		private System.Windows.Forms.Label label8;
		private System.Windows.Forms.Label label6;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtHasta;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtDesde;
		private Infragistics.Win.Misc.UltraLabel lblContador;
		private System.Windows.Forms.Button btnBuscar;
		private Infragistics.Win.Misc.UltraLabel ultraLabel4;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtBusqueda;
		public Infragistics.Win.UltraWinGrid.UltraGrid uGridOrdenesDeCompra;
		private C1.Data.C1DataSet cdsSeteoF;
		private Infragistics.Win.UltraWinEditors.UltraOptionSet optEstados;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource1;
		private System.Windows.Forms.Button btnExcel;
		private Infragistics.Win.UltraWinGrid.UltraDropDown cmbCuenta;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource2;
		private Infragistics.Win.UltraWinEditors.UltraOptionSet optEstadoDentalis;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public frmReporteAprobacionODC()
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
    int idBloqueaTransacciones = 0;
		#region Código generado por el Diseñador de Windows Forms
		/// <summary>
		/// Método necesario para admitir el Diseñador. No se puede modificar
		/// el contenido del método con el editor de código.
		/// </summary>
		private void InitializeComponent()
		{
			Infragistics.Win.Appearance appearance1 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.ValueListItem valueListItem1 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem2 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem3 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton1 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmReporteAprobacionODC));
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton2 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCompra");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Fecha");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaTentativaLlegada");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("HoraLlegada");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Numero");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Ruc");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Proveedor");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Total");
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Estado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn11 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCuenta", -1, "cmbCuenta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn12 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Periodo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn13 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Presupuesto");
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn14 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Ir", 0);
			Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn15 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Aceptar", 1);
			Infragistics.Win.Appearance appearance11 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance12 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance13 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance14 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance15 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance16 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance17 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCompra");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Fecha");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn3 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("FechaTentativaLlegada");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn4 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("HoraLlegada");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn5 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Numero");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn6 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn7 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Ruc");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn8 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Proveedor");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn9 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Total");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn10 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Estado");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn11 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCuenta");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn12 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Periodo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn13 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Presupuesto");
			Infragistics.Win.Appearance appearance18 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance19 = new Infragistics.Win.Appearance();
			Infragistics.Win.ValueListItem valueListItem4 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem5 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem6 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem7 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem8 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem9 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn16 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCuenta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn17 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cuenta");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn14 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCuenta");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn15 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Cuenta");
			Infragistics.Win.Appearance appearance20 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance21 = new Infragistics.Win.Appearance();
			Infragistics.Win.ValueListItem valueListItem10 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem11 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem12 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem13 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem14 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem15 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem16 = new Infragistics.Win.ValueListItem();
			this.groupBox1 = new System.Windows.Forms.GroupBox();
			this.optOrigen = new Infragistics.Win.UltraWinEditors.UltraOptionSet();
			this.label8 = new System.Windows.Forms.Label();
			this.label6 = new System.Windows.Forms.Label();
			this.dtHasta = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.dtDesde = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.lblContador = new Infragistics.Win.Misc.UltraLabel();
			this.btnBuscar = new System.Windows.Forms.Button();
			this.ultraLabel4 = new Infragistics.Win.Misc.UltraLabel();
			this.txtBusqueda = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.uGridOrdenesDeCompra = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource1 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.optEstados = new Infragistics.Win.UltraWinEditors.UltraOptionSet();
			this.btnExcel = new System.Windows.Forms.Button();
			this.cmbCuenta = new Infragistics.Win.UltraWinGrid.UltraDropDown();
			this.ultraDataSource2 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.optEstadoDentalis = new Infragistics.Win.UltraWinEditors.UltraOptionSet();
			((System.ComponentModel.ISupportInitialize)(this.optOrigen)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtHasta)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtDesde)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtBusqueda)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridOrdenesDeCompra)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.optEstados)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCuenta)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.optEstadoDentalis)).BeginInit();
			this.SuspendLayout();
			// 
			// groupBox1
			// 
			this.groupBox1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.groupBox1.Location = new System.Drawing.Point(0, 96);
			this.groupBox1.Name = "groupBox1";
			this.groupBox1.Size = new System.Drawing.Size(1024, 8);
			this.groupBox1.TabIndex = 289;
			this.groupBox1.TabStop = false;
			// 
			// optOrigen
			// 
			appearance1.ForeColorDisabled = System.Drawing.Color.Black;
			this.optOrigen.Appearance = appearance1;
			this.optOrigen.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.optOrigen.CheckedIndex = 0;
			this.optOrigen.ItemAppearance = appearance2;
			this.optOrigen.ItemOrigin = new System.Drawing.Point(10, 0);
			valueListItem1.DataValue = 0;
			valueListItem1.DisplayText = "(F7) Todos";
			valueListItem2.DataValue = 1;
			valueListItem2.DisplayText = "(F8) Locales";
			valueListItem3.DataValue = 2;
			valueListItem3.DisplayText = "(F9) Importaciones";
			this.optOrigen.Items.Add(valueListItem1);
			this.optOrigen.Items.Add(valueListItem2);
			this.optOrigen.Items.Add(valueListItem3);
			this.optOrigen.ItemSpacingVertical = 10;
			this.optOrigen.Location = new System.Drawing.Point(584, 40);
			this.optOrigen.Name = "optOrigen";
			this.optOrigen.Size = new System.Drawing.Size(312, 26);
			this.optOrigen.TabIndex = 278;
			this.optOrigen.Text = "(F7) Todos";
			// 
			// label8
			// 
			this.label8.AutoSize = true;
			this.label8.Location = new System.Drawing.Point(192, 10);
			this.label8.Name = "label8";
			this.label8.Size = new System.Drawing.Size(34, 16);
			this.label8.TabIndex = 287;
			this.label8.Text = "Hasta";
			// 
			// label6
			// 
			this.label6.AutoSize = true;
			this.label6.Location = new System.Drawing.Point(8, 10);
			this.label6.Name = "label6";
			this.label6.Size = new System.Drawing.Size(37, 16);
			this.label6.TabIndex = 286;
			this.label6.Text = "Desde";
			// 
			// dtHasta
			// 
			appearance3.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtHasta.Appearance = appearance3;
			this.dtHasta.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton1.Caption = "Today";
			this.dtHasta.DateButtons.Add(dateButton1);
			this.dtHasta.Format = "dd/MM/yyyy";
			this.dtHasta.Location = new System.Drawing.Point(248, 8);
			this.dtHasta.Name = "dtHasta";
			this.dtHasta.NonAutoSizeHeight = 23;
			this.dtHasta.Size = new System.Drawing.Size(104, 21);
			this.dtHasta.SpinButtonsVisible = true;
			this.dtHasta.TabIndex = 277;
			this.dtHasta.Value = ((object)(resources.GetObject("dtHasta.Value")));
			// 
			// dtDesde
			// 
			appearance4.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtDesde.Appearance = appearance4;
			this.dtDesde.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton2.Caption = "Today";
			this.dtDesde.DateButtons.Add(dateButton2);
			this.dtDesde.Format = "dd/MM/yyyy";
			this.dtDesde.Location = new System.Drawing.Point(64, 8);
			this.dtDesde.Name = "dtDesde";
			this.dtDesde.NonAutoSizeHeight = 23;
			this.dtDesde.Size = new System.Drawing.Size(104, 21);
			this.dtDesde.SpinButtonsVisible = true;
			this.dtDesde.TabIndex = 276;
			this.dtDesde.Value = ((object)(resources.GetObject("dtDesde.Value")));
			// 
			// lblContador
			// 
			appearance5.ForeColor = System.Drawing.Color.Firebrick;
			this.lblContador.Appearance = appearance5;
			this.lblContador.AutoSize = true;
			this.lblContador.BackColor = System.Drawing.Color.Transparent;
			this.lblContador.Font = new System.Drawing.Font("Tahoma", 10F);
			this.lblContador.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(50)), ((System.Byte)(50)), ((System.Byte)(50)));
			this.lblContador.Location = new System.Drawing.Point(432, 80);
			this.lblContador.Name = "lblContador";
			this.lblContador.Size = new System.Drawing.Size(0, 0);
			this.lblContador.TabIndex = 285;
			// 
			// btnBuscar
			// 
			this.btnBuscar.Location = new System.Drawing.Point(288, 71);
			this.btnBuscar.Name = "btnBuscar";
			this.btnBuscar.TabIndex = 280;
			this.btnBuscar.Text = "&Buscar";
			this.btnBuscar.Click += new System.EventHandler(this.btnBuscar_Click);
			// 
			// ultraLabel4
			// 
			this.ultraLabel4.AutoSize = true;
			this.ultraLabel4.BackColor = System.Drawing.Color.Transparent;
			this.ultraLabel4.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.ultraLabel4.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(50)), ((System.Byte)(50)), ((System.Byte)(50)));
			this.ultraLabel4.Location = new System.Drawing.Point(8, 75);
			this.ultraLabel4.Name = "ultraLabel4";
			this.ultraLabel4.Size = new System.Drawing.Size(37, 15);
			this.ultraLabel4.TabIndex = 284;
			this.ultraLabel4.Text = "Buscar";
			// 
			// txtBusqueda
			// 
			appearance6.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtBusqueda.Appearance = appearance6;
			this.txtBusqueda.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtBusqueda.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtBusqueda.Location = new System.Drawing.Point(64, 72);
			this.txtBusqueda.Name = "txtBusqueda";
			this.txtBusqueda.Size = new System.Drawing.Size(200, 21);
			this.txtBusqueda.TabIndex = 279;
			// 
			// uGridOrdenesDeCompra
			// 
			this.uGridOrdenesDeCompra.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.uGridOrdenesDeCompra.Cursor = System.Windows.Forms.Cursors.Default;
			this.uGridOrdenesDeCompra.DataSource = this.ultraDataSource1;
			appearance7.BackColor = System.Drawing.Color.White;
			appearance7.ForeColor = System.Drawing.Color.Black;
			appearance7.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridOrdenesDeCompra.DisplayLayout.Appearance = appearance7;
			ultraGridColumn1.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn1.Width = 95;
			ultraGridColumn2.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn2.Header.VisiblePosition = 3;
			ultraGridColumn2.Width = 59;
			ultraGridColumn3.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn3.Header.Caption = "Llegada";
			ultraGridColumn3.Header.VisiblePosition = 4;
			ultraGridColumn3.Width = 59;
			ultraGridColumn4.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn4.Format = "HH:mm";
			ultraGridColumn4.Header.Caption = "Hora";
			ultraGridColumn4.Header.VisiblePosition = 5;
			ultraGridColumn4.Width = 52;
			ultraGridColumn5.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn5.Header.Caption = "Número";
			ultraGridColumn5.Header.VisiblePosition = 6;
			ultraGridColumn5.Width = 78;
			ultraGridColumn6.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn6.Header.Caption = "Importación";
			ultraGridColumn6.Header.VisiblePosition = 7;
			ultraGridColumn6.Width = 158;
			ultraGridColumn7.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn7.Header.Caption = "RUC";
			ultraGridColumn7.Header.VisiblePosition = 8;
			ultraGridColumn7.Width = 92;
			ultraGridColumn8.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn8.Header.VisiblePosition = 9;
			ultraGridColumn8.Width = 251;
			ultraGridColumn9.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance8.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn9.CellAppearance = appearance8;
			ultraGridColumn9.Format = "#,##0.00";
			ultraGridColumn9.Header.VisiblePosition = 10;
			ultraGridColumn9.Width = 72;
			ultraGridColumn10.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn10.Header.VisiblePosition = 11;
			ultraGridColumn10.Width = 101;
			ultraGridColumn11.Header.Caption = "Cuenta";
			ultraGridColumn11.Header.VisiblePosition = 12;
			ultraGridColumn11.Hidden = true;
			ultraGridColumn11.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDown;
			ultraGridColumn11.Width = 120;
			ultraGridColumn12.Header.VisiblePosition = 13;
			ultraGridColumn12.Hidden = true;
			ultraGridColumn12.Width = 100;
			appearance9.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn13.CellAppearance = appearance9;
			ultraGridColumn13.Format = "#,##0.00";
			ultraGridColumn13.Header.VisiblePosition = 14;
			ultraGridColumn13.Hidden = true;
			ultraGridColumn13.Width = 120;
			ultraGridColumn14.ButtonDisplayStyle = Infragistics.Win.UltraWinGrid.ButtonDisplayStyle.Always;
			ultraGridColumn14.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			appearance10.Image = ((object)(resources.GetObject("appearance10.Image")));
			appearance10.ImageHAlign = Infragistics.Win.HAlign.Center;
			appearance10.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn14.CellButtonAppearance = appearance10;
			ultraGridColumn14.Header.VisiblePosition = 1;
			ultraGridColumn14.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.Button;
			ultraGridColumn14.Width = 33;
			ultraGridColumn15.ButtonDisplayStyle = Infragistics.Win.UltraWinGrid.ButtonDisplayStyle.Always;
			ultraGridColumn15.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			appearance11.Image = ((object)(resources.GetObject("appearance11.Image")));
			appearance11.ImageHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn15.CellButtonAppearance = appearance11;
			ultraGridColumn15.Header.Caption = "Ok";
			ultraGridColumn15.Header.VisiblePosition = 2;
			ultraGridColumn15.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.Button;
			ultraGridColumn15.Width = 32;
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
																										 ultraGridColumn15});
			ultraGridBand1.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
			ultraGridBand1.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			ultraGridBand1.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			this.uGridOrdenesDeCompra.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			appearance12.ForeColor = System.Drawing.Color.Black;
			appearance12.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridOrdenesDeCompra.DisplayLayout.Override.ActiveRowAppearance = appearance12;
			this.uGridOrdenesDeCompra.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.TemplateOnBottom;
			this.uGridOrdenesDeCompra.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.True;
			this.uGridOrdenesDeCompra.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance13.BackColor = System.Drawing.Color.Transparent;
			this.uGridOrdenesDeCompra.DisplayLayout.Override.CardAreaAppearance = appearance13;
			appearance14.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance14.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance14.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance14.FontData.BoldAsString = "True";
			appearance14.FontData.Name = "Arial";
			appearance14.FontData.SizeInPoints = 8F;
			appearance14.ForeColor = System.Drawing.Color.White;
			appearance14.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridOrdenesDeCompra.DisplayLayout.Override.HeaderAppearance = appearance14;
			this.uGridOrdenesDeCompra.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortSingle;
			appearance15.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance15.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance15.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridOrdenesDeCompra.DisplayLayout.Override.RowAlternateAppearance = appearance15;
			appearance16.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance16.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance16.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridOrdenesDeCompra.DisplayLayout.Override.RowSelectorAppearance = appearance16;
			appearance17.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance17.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance17.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridOrdenesDeCompra.DisplayLayout.Override.SelectedRowAppearance = appearance17;
			this.uGridOrdenesDeCompra.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridOrdenesDeCompra.Location = new System.Drawing.Point(8, 112);
			this.uGridOrdenesDeCompra.Name = "uGridOrdenesDeCompra";
			this.uGridOrdenesDeCompra.Size = new System.Drawing.Size(1008, 248);
			this.uGridOrdenesDeCompra.TabIndex = 281;
			this.uGridOrdenesDeCompra.ClickCellButton += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.uGridOrdenesDeCompra_ClickCellButton);
			this.uGridOrdenesDeCompra.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.uGridOrdenesDeCompra_InitializeLayout);
			// 
			// ultraDataSource1
			// 
			ultraDataColumn1.DataType = typeof(int);
			ultraDataColumn2.DataType = typeof(System.DateTime);
			ultraDataColumn2.DefaultValue = new System.DateTime(((long)(0)));
			ultraDataColumn3.DataType = typeof(System.DateTime);
			ultraDataColumn3.DefaultValue = new System.DateTime(((long)(0)));
			ultraDataColumn4.DataType = typeof(System.DateTime);
			ultraDataColumn4.DefaultValue = new System.DateTime(((long)(0)));
			ultraDataColumn9.DataType = typeof(System.Decimal);
			ultraDataColumn9.DefaultValue = new System.Decimal(new int[] {
																																		 0,
																																		 0,
																																		 0,
																																		 0});
			ultraDataColumn11.DataType = typeof(int);
			ultraDataColumn12.DataType = typeof(System.DateTime);
			ultraDataColumn13.DataType = typeof(System.Decimal);
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
																																 ultraDataColumn13});
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
			// optEstados
			// 
			appearance18.ForeColorDisabled = System.Drawing.Color.Black;
			this.optEstados.Appearance = appearance18;
			this.optEstados.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.optEstados.CheckedIndex = 1;
			this.optEstados.ItemAppearance = appearance19;
			this.optEstados.ItemOrigin = new System.Drawing.Point(10, 0);
			valueListItem4.DataValue = 0;
			valueListItem4.DisplayText = "Todos";
			valueListItem5.DataValue = 1;
			valueListItem5.DisplayText = "Pendientes";
			valueListItem6.DataValue = 3;
			valueListItem6.DisplayText = "Procesado";
			valueListItem7.DataValue = 9;
			valueListItem7.DisplayText = "Facturado";
			valueListItem8.DataValue = 5;
			valueListItem8.DisplayText = "Aprobado";
			valueListItem9.DataValue = 4;
			valueListItem9.DisplayText = "Rechazado";
			this.optEstados.Items.Add(valueListItem4);
			this.optEstados.Items.Add(valueListItem5);
			this.optEstados.Items.Add(valueListItem6);
			this.optEstados.Items.Add(valueListItem7);
			this.optEstados.Items.Add(valueListItem8);
			this.optEstados.Items.Add(valueListItem9);
			this.optEstados.ItemSpacingVertical = 10;
			this.optEstados.Location = new System.Drawing.Point(8, 40);
			this.optEstados.Name = "optEstados";
			this.optEstados.Size = new System.Drawing.Size(496, 26);
			this.optEstados.TabIndex = 311;
			this.optEstados.Text = "Pendientes";
			this.optEstados.ValueChanged += new System.EventHandler(this.optEstados_ValueChanged);
			// 
			// btnExcel
			// 
			this.btnExcel.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btnExcel.Image = ((System.Drawing.Image)(resources.GetObject("btnExcel.Image")));
			this.btnExcel.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnExcel.Location = new System.Drawing.Point(944, 6);
			this.btnExcel.Name = "btnExcel";
			this.btnExcel.Size = new System.Drawing.Size(75, 24);
			this.btnExcel.TabIndex = 737;
			this.btnExcel.Text = "Excel";
			this.btnExcel.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnExcel.Visible = false;
			this.btnExcel.Click += new System.EventHandler(this.btnExcel_Click);
			// 
			// cmbCuenta
			// 
			this.cmbCuenta.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbCuenta.DataSource = this.ultraDataSource2;
			this.cmbCuenta.DisplayLayout.AutoFitColumns = true;
			ultraGridBand2.AddButtonCaption = "ImportacionTipo";
			ultraGridColumn16.Header.VisiblePosition = 0;
			ultraGridColumn16.Hidden = true;
			ultraGridColumn16.Width = 144;
			ultraGridColumn17.Header.VisiblePosition = 1;
			ultraGridColumn17.Width = 374;
			ultraGridBand2.Columns.AddRange(new object[] {
																										 ultraGridColumn16,
																										 ultraGridColumn17});
			this.cmbCuenta.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			this.cmbCuenta.DisplayMember = "Cuenta";
			this.cmbCuenta.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbCuenta.Location = new System.Drawing.Point(416, 192);
			this.cmbCuenta.Name = "cmbCuenta";
			this.cmbCuenta.Size = new System.Drawing.Size(376, 112);
			this.cmbCuenta.TabIndex = 738;
			this.cmbCuenta.ValueMember = "idCuenta";
			this.cmbCuenta.Visible = false;
			// 
			// ultraDataSource2
			// 
			ultraDataColumn14.DataType = typeof(int);
			this.ultraDataSource2.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn14,
																																 ultraDataColumn15});
			// 
			// optEstadoDentalis
			// 
			appearance20.ForeColorDisabled = System.Drawing.Color.Black;
			this.optEstadoDentalis.Appearance = appearance20;
			this.optEstadoDentalis.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.optEstadoDentalis.CheckedIndex = 1;
			this.optEstadoDentalis.ItemAppearance = appearance21;
			this.optEstadoDentalis.ItemOrigin = new System.Drawing.Point(10, 0);
			valueListItem10.DataValue = 0;
			valueListItem10.DisplayText = "Todos";
			valueListItem11.DataValue = 1;
			valueListItem11.DisplayText = "Pendientes";
			valueListItem12.DataValue = 6;
			valueListItem12.DisplayText = "Revisión";
			valueListItem13.DataValue = 3;
			valueListItem13.DisplayText = "Procesado";
			valueListItem14.DataValue = 9;
			valueListItem14.DisplayText = "Facturado";
			valueListItem15.DataValue = 5;
			valueListItem15.DisplayText = "Aprobado";
			valueListItem16.DataValue = 4;
			valueListItem16.DisplayText = "Rechazado";
			this.optEstadoDentalis.Items.Add(valueListItem10);
			this.optEstadoDentalis.Items.Add(valueListItem11);
			this.optEstadoDentalis.Items.Add(valueListItem12);
			this.optEstadoDentalis.Items.Add(valueListItem13);
			this.optEstadoDentalis.Items.Add(valueListItem14);
			this.optEstadoDentalis.Items.Add(valueListItem15);
			this.optEstadoDentalis.Items.Add(valueListItem16);
			this.optEstadoDentalis.ItemSpacingVertical = 10;
			this.optEstadoDentalis.Location = new System.Drawing.Point(8, 40);
			this.optEstadoDentalis.Name = "optEstadoDentalis";
			this.optEstadoDentalis.Size = new System.Drawing.Size(568, 26);
			this.optEstadoDentalis.TabIndex = 739;
			this.optEstadoDentalis.Text = "Pendientes";
			this.optEstadoDentalis.ValueChanged += new System.EventHandler(this.optEstadoDentalis_ValueChanged);
			// 
			// frmReporteAprobacionODC
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(1024, 366);
			this.Controls.Add(this.optEstadoDentalis);
			this.Controls.Add(this.cmbCuenta);
			this.Controls.Add(this.btnExcel);
			this.Controls.Add(this.optEstados);
			this.Controls.Add(this.groupBox1);
			this.Controls.Add(this.optOrigen);
			this.Controls.Add(this.label8);
			this.Controls.Add(this.label6);
			this.Controls.Add(this.dtHasta);
			this.Controls.Add(this.dtDesde);
			this.Controls.Add(this.lblContador);
			this.Controls.Add(this.btnBuscar);
			this.Controls.Add(this.ultraLabel4);
			this.Controls.Add(this.txtBusqueda);
			this.Controls.Add(this.uGridOrdenesDeCompra);
			this.KeyPreview = true;
			this.Name = "frmReporteAprobacionODC";
			this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
			this.Text = "Aprobación Ordenes De Compra";
			this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.frmReporteAprobacionODC_KeyDown);
			this.Load += new System.EventHandler(this.frmReporteAprobacionODC_Load);
			((System.ComponentModel.ISupportInitialize)(this.optOrigen)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtHasta)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtDesde)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtBusqueda)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridOrdenesDeCompra)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.optEstados)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCuenta)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.optEstadoDentalis)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion
		private Acceso miAcceso;
     bool ODCDiaCredito = false;
		private void frmReporteAprobacionODC_Load(object sender, System.EventArgs e)
		{

			miAcceso = new Acceso(cdsSeteoF, "0628");

			if (miAcceso.BExportar)
			{
				this.btnExcel.Visible =  true;

			}
			
			
				
			this.dtDesde.CalendarInfo.MinDate = DateTime.Parse("01/01/2019");
			this.dtHasta.CalendarInfo.MinDate = DateTime.Parse("01/01/2019");				
			
			this.dtDesde.CalendarInfo.MaxDate = DateTime.Today;
			this.dtHasta.CalendarInfo.MaxDate = DateTime.Today;

			this.dtDesde.Value = "01/" + DateTime.Now.Month.ToString() + "/" + DateTime.Now.Year.ToString();			
			this.dtHasta.Value = DateTime.Today;

			this.cmbCuenta.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "select idCuenta, c.Codigo +'  '+ c.Descripcion + ' / ' + g.Grupo as Cuenta  from Cuenta c Inner Join CuentaGrupo g on c.idCuentaAgrupa = g.idCuentaAgrupa where GASTO > 0");

			ODCDiaCredito = Funcion.bEscalarSQL(cdsSeteoF, string.Format("select isnull (ODCDiaCredito, 0) from Seteo"), true);
//			if (ODCDiaCredito)
//			{
//
//				this.uGridOrdenesDeCompra.DisplayLayout.Bands[0].Columns["idCuenta"].Hidden = false;
//				this.uGridOrdenesDeCompra.DisplayLayout.Bands[0].Columns["Periodo"].Hidden = false;
//				this.uGridOrdenesDeCompra.DisplayLayout.Bands[0].Columns["Presupuesto"].Hidden = false;
//			}

			if (ODCDiaCredito)
			{
				this.optEstadoDentalis.Visible = true;
				this.optEstados.Visible = false;
			}
			else
			{
				this.optEstadoDentalis.Visible = false;
				this.optEstados.Visible = true;
			}
		}

		private void btnBuscar_Click(object sender, System.EventArgs e)
		{
			int IEstado = 0;
			bool IBorrar = false;

			bool bFechas = false;

			DateTime dtDesde = DateTime.Today;
			DateTime dtHasta = DateTime.Today;

			if (this.dtDesde.Value != System.DBNull.Value || this.dtHasta.Value != System.DBNull.Value)
			{
				bFechas = true;

				dtDesde = (DateTime)this.dtDesde.Value;
				dtHasta = (DateTime)this.dtHasta.Value;
			}

			if (!Validacion.vbComparaFechas(Convert.ToDateTime(this.dtDesde.Value), Convert.ToDateTime(this.dtHasta.Value), this.dtDesde, "La fecha inicial no puede ser posterior a la fecha final", false)) return;
      int iEstado = 0;

			if (!ODCDiaCredito) 
			{
				iEstado = (int) this.optEstados.Value;
			}
			else
			{
       iEstado = (int) this.optEstadoDentalis.Value;
				if (miAcceso.Imprimir) this.optEstadoDentalis.Value = 1;
			}
			string sSQL = string.Format("Exec ConsultaAprobacionOrdenDeCompra '{0}', {1}, {2}, {3}, {4}, '{5}', '{6}'", 
				this.txtBusqueda.Text.ToString(), iEstado, IBorrar, (int)this.optOrigen.Value, bFechas, dtDesde.ToString("yyyyMMdd"), dtHasta.ToString("yyyyMMdd"));
			this.uGridOrdenesDeCompra.DataSource = Funcion.dvProcedimiento(cdsSeteoF, sSQL);

			if (miAcceso.Editar)
			{
				this.uGridOrdenesDeCompra.DisplayLayout.Bands[0].Columns["Aceptar"].Hidden = false;
			}
			else
			{
				this.uGridOrdenesDeCompra.DisplayLayout.Bands[0].Columns["Aceptar"].Hidden = true;
			}
			if (!ODCDiaCredito)
			{
				Estados ();
			}
			else
			{
				EstadosDentalis();
			}

			if (this.uGridOrdenesDeCompra.Rows.Count > 0) this.uGridOrdenesDeCompra.Focus();
			else this.txtBusqueda.Focus();

			this.lblContador.Text = this.uGridOrdenesDeCompra.Rows.Count + " Registros Encontrados";
		}

		private void frmReporteAprobacionODC_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Escape) this.Close();
			if (e.KeyCode == Keys.F7) this.optOrigen.Value = 0;
			if (e.KeyCode == Keys.F8) this.optOrigen.Value = 1;
			if (e.KeyCode == Keys.F9) this.optOrigen.Value = 2;
		}

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);	
		}

		private void optEstados_ValueChanged(object sender, System.EventArgs e)
		{
			
      Estados ();
			this.btnBuscar_Click (sender,  e);
		}

		private void Estados ()
		{
			if((int) this.optEstados.Value == 0)
			{

				this.uGridOrdenesDeCompra.DisplayLayout.Bands[0].Columns["Aceptar"].Hidden = true;
				this.uGridOrdenesDeCompra.DisplayLayout.Bands[0].Columns["Ir"].Hidden = true;
			}

			if((int) this.optEstados.Value == 1)
			{
				if (miAcceso.Editar)
				{
					this.uGridOrdenesDeCompra.DisplayLayout.Bands[0].Columns["Aceptar"].Hidden = false;
				}
				else
				{
					this.uGridOrdenesDeCompra.DisplayLayout.Bands[0].Columns["Aceptar"].Hidden = true;
				}
		
				this.uGridOrdenesDeCompra.DisplayLayout.Bands[0].Columns["Ir"].Hidden = false;
			}
			if((int) this.optEstados.Value == 3)
			{

				this.uGridOrdenesDeCompra.DisplayLayout.Bands[0].Columns["Aceptar"].Hidden = true;
				this.uGridOrdenesDeCompra.DisplayLayout.Bands[0].Columns["Ir"].Hidden = false;
			}
			
			if((int) this.optEstados.Value == 9)
			{

				this.uGridOrdenesDeCompra.DisplayLayout.Bands[0].Columns["Aceptar"].Hidden = true;
				this.uGridOrdenesDeCompra.DisplayLayout.Bands[0].Columns["Ir"].Hidden = false;
			}
			if((int) this.optEstados.Value == 5)
			{

				this.uGridOrdenesDeCompra.DisplayLayout.Bands[0].Columns["Aceptar"].Hidden = true;
				this.uGridOrdenesDeCompra.DisplayLayout.Bands[0].Columns["Ir"].Hidden = false;
			}
			if((int) this.optEstados.Value == 4)
			{

				this.uGridOrdenesDeCompra.DisplayLayout.Bands[0].Columns["Aceptar"].Hidden = true;
				this.uGridOrdenesDeCompra.DisplayLayout.Bands[0].Columns["Ir"].Hidden = false;
			}
			if((int) this.optEstados.Value == 6)
			{
				this.uGridOrdenesDeCompra.DisplayLayout.Bands[0].Columns["Ir"].Hidden = true;
				this.uGridOrdenesDeCompra.DisplayLayout.Bands[0].Columns["Aceptar"].Hidden = true;
			}
		}


		private void EstadosDentalis ()
		{
//							this.uGridOrdenesDeCompra.DisplayLayout.Bands[0].Columns["idCuenta"].Hidden = false;
//							this.uGridOrdenesDeCompra.DisplayLayout.Bands[0].Columns["Periodo"].Hidden = false;
//							this.uGridOrdenesDeCompra.DisplayLayout.Bands[0].Columns["Presupuesto"].Hidden = false;

			if((int) this.optEstadoDentalis.Value == 0)
			{

				this.uGridOrdenesDeCompra.DisplayLayout.Bands[0].Columns["Aceptar"].Hidden = true;
				this.uGridOrdenesDeCompra.DisplayLayout.Bands[0].Columns["Ir"].Hidden = true;
			}

			if((int) this.optEstadoDentalis.Value == 1)
			{
				if (miAcceso.Editar)
				{
					this.uGridOrdenesDeCompra.DisplayLayout.Bands[0].Columns["Aceptar"].Hidden = false;
				}
				else
				{
					this.uGridOrdenesDeCompra.DisplayLayout.Bands[0].Columns["Aceptar"].Hidden = true;
				}
		
				this.uGridOrdenesDeCompra.DisplayLayout.Bands[0].Columns["Ir"].Hidden = false;
			}
			if((int) this.optEstadoDentalis.Value == 3)
			{

				this.uGridOrdenesDeCompra.DisplayLayout.Bands[0].Columns["Aceptar"].Hidden = true;
				this.uGridOrdenesDeCompra.DisplayLayout.Bands[0].Columns["Ir"].Hidden = false;
			}
			
			if((int) this.optEstadoDentalis.Value == 9)
			{

				this.uGridOrdenesDeCompra.DisplayLayout.Bands[0].Columns["Aceptar"].Hidden = true;
				this.uGridOrdenesDeCompra.DisplayLayout.Bands[0].Columns["Ir"].Hidden = false;
			}
			if((int) this.optEstadoDentalis.Value == 5)
			{


				this.uGridOrdenesDeCompra.DisplayLayout.Bands[0].Columns["Aceptar"].Hidden = true;
				this.uGridOrdenesDeCompra.DisplayLayout.Bands[0].Columns["Ir"].Hidden = false;
			}
			if((int) this.optEstadoDentalis.Value == 4)
			{

				this.uGridOrdenesDeCompra.DisplayLayout.Bands[0].Columns["Aceptar"].Hidden = true;
				this.uGridOrdenesDeCompra.DisplayLayout.Bands[0].Columns["Ir"].Hidden = false;
			}
			if((int) this.optEstadoDentalis.Value == 6)
			{
				this.uGridOrdenesDeCompra.DisplayLayout.Bands[0].Columns["idCuenta"].Hidden = false;
				this.uGridOrdenesDeCompra.DisplayLayout.Bands[0].Columns["Periodo"].Hidden = false;
				this.uGridOrdenesDeCompra.DisplayLayout.Bands[0].Columns["Presupuesto"].Hidden = false;
				this.uGridOrdenesDeCompra.DisplayLayout.Bands[0].Columns["Ir"].Hidden = false;
				this.uGridOrdenesDeCompra.DisplayLayout.Bands[0].Columns["Aceptar"].Hidden = false;
			}
		}

		private void uGridOrdenesDeCompra_ClickCellButton(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			if (e.Cell.Column.ToString() == "Aceptar")
			{	
				if (Funcion.ValidaBloqueo((int)(int)e.Cell.Row.Cells["idCompra"].Value, 1, cdsSeteoF)) return;

				#region Dentalis
				if (ODCDiaCredito)
				{
					if ((string)e.Cell.Row.Cells["Estado"].Value =="REVISION")
					{
						if (e.Cell.Row.Cells["idCuenta"].Value == System.DBNull.Value)
						{
							MessageBox.Show("Seleccione la cuenta del Gasto", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);


							return;
						}	
						if (e.Cell.Row.Cells["Periodo"].Value == System.DBNull.Value)
						{
							MessageBox.Show("Seleccione el Periodo afectado para el Gasto", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);

							return;
						}	
						if ((int) e.Cell.Row.Cells["idCuenta"].Value == 0)
						{
							MessageBox.Show("Seleccione la cuenta del Gasto", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);

							return;
						}	

						decimal dPresupuesto = Funcion.decEscalarSQL(cdsSeteoF, string.Format(" Exec [ValidaPResupuestoGasto] {0}, '{1}', {2}",
							(int) e.Cell.Row.Cells["idCuenta"].Value, Convert.ToDateTime(e.Cell.Row.Cells["Periodo"].Value).ToString("yyyyMMdd") , 1));		

						decimal dValorCuenta = Funcion.decEscalarSQL(cdsSeteoF, string.Format(" Exec [ValidaPResupuestoGasto] {0}, '{1}', {2}",
							(int) e.Cell.Row.Cells["idCuenta"].Value, Convert.ToDateTime(e.Cell.Row.Cells["Periodo"].Value).ToString("yyyyMMdd") , 2));		

						string sCuenta = Funcion.sEscalarSQL(cdsSeteoF, string.Format(" Exec [ValidaPResupuestoGasto] {0}, '{1}', {2}",
							(int) e.Cell.Row.Cells["idCuenta"].Value, Convert.ToDateTime(e.Cell.Row.Cells["Periodo"].Value).ToString("yyyyMMdd") , 3));	

						decimal dODC = (decimal) e.Cell.Row.Cells["Total"].Value;
						decimal dDisponible = dPresupuesto - (dValorCuenta + dODC);
						if (dPresupuesto > 0)  e.Cell.Row.Cells["Presupuesto"].Value = dPresupuesto;

						if	(DialogResult.Yes == MessageBox.Show("Cuenta: "+ sCuenta +"\n"+ "Presupuesto : "+ dPresupuesto.ToString("#,##0.00") + "\n"+" Utilizado Cuenta : "+ dValorCuenta.ToString("#,##0.00") + "\n"+" Monto ODC : "+  dODC.ToString("#,##0.00")+ "\n\n\t Disponible " + dDisponible.ToString("#,##0.00") + "\n\n ¿Esta seguro de Aprobar?.", "Point Technology", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2))
						{
							this.DialogResult = DialogResult.OK;
//							this.Close();
						}
						else
						{
							return;
						}

					}
				}
				#endregion Dentalis

			
//				idBloqueaTransacciones = Funcion.iEscalarSQL(cdsSeteoF, string.Format(" Exec RegistraBloqueoTransacciones {0} , {1} , {2}, {3} ", 0, 1, Convert.ToInt32( e.Cell.Row.Cells["idCompra"].Value),1));	
//

				int iOptEstado = 0;
        string sSQLNumero="";
        string sMensaje ="Aprobación";

				if (ODCDiaCredito)
				{
					if ((int) this.optEstadoDentalis.Value == 1)sMensaje ="Aprobación Gerencia";
					if ((int) this.optEstadoDentalis.Value == 6)sMensaje ="Aprobación Presupuestaria";
				}


				using (AprobarOrdenDeCompra ODC  = new AprobarOrdenDeCompra(sMensaje))
				{
					if (DialogResult.OK == ODC.ShowDialog())
					{		
						iOptEstado = (int)ODC.optEstados.Value;
						if(ODCDiaCredito)
						{
							if((int) this.optEstadoDentalis.Value == 1 && (int)ODC.optEstados.Value == 5 )
							{
								iOptEstado = 6;
							}
							if((int) this.optEstadoDentalis.Value == 6 && (int)ODC.optEstados.Value == 5 )
							{
								iOptEstado = 5;

								string ssSQL = string.Format("Exec NumeracionLocales 76, 0, 0");
								sSQLNumero = Funcion.sEscalarSQL(cdsSeteoF, ssSQL);	


							}
						}
						if (!ODCDiaCredito)
						{
							e.Cell.Row.Cells["idCuenta"].Value = 0;
							e.Cell.Row.Cells["Periodo"].Value = "02/02/2015";
						}

             int iOrigenPresupuesto = 0;
              iOrigenPresupuesto = Funcion.iEscalarSQL(cdsSeteoF, string.Format("Exec [PresupuestoValidaOrigenODC] {0}", (int)e.Cell.Row.Cells["idCompra"].Value));	
				   
						if ( iOptEstado == 5)
						{
							if( iOrigenPresupuesto > 0 ) 
							{
								iOptEstado = 3;
							}
						}
						
							string sSQL = string.Format("update Compra set Notas = Notas + ' '+'{0}', Estado = {1}, idCuenta = {3}, Periodo = '{4}', Numero2='{5}' where  idCompra = {2}" ,ODC.txtNotas.Text, iOptEstado, (int)e.Cell.Row.Cells["idCompra"].Value, (int)e.Cell.Row.Cells["idCuenta"].Value, Convert.ToDateTime(e.Cell.Row.Cells["Periodo"].Value).ToString("yyyyMMdd"), sSQLNumero );
							Funcion.EjecutaSQL(cdsSeteoF, sSQL);

						if ( iOrigenPresupuesto > 0 )
						{
							if ( iOptEstado == 3)
							{
								string sSQLa = string.Format("update Compra set Notas = Notas + ' '+'{0}', Estado = {1} where  idCompra = {2}" ,ODC.txtNotas.Text, 5, iOrigenPresupuesto  );
								Funcion.EjecutaSQL(cdsSeteoF, sSQLa);
							}
							if ( iOptEstado == 4)
							{
								string sSQLb = string.Format("update Compra set Notas = Notas + ' '+'{0}', Estado = {1} where  idCompra = {2}" ,ODC.txtNotas.Text, 6, iOrigenPresupuesto  );
								Funcion.EjecutaSQL(cdsSeteoF, sSQLb);
							}

						}


						this.btnBuscar_Click(sender, e);

						//						string sSQLBloque = string.Format("Exec RegistraBloqueoTransacciones {0}, 0, {1}, 1", idBloqueaTransacciones, Convert.ToInt32( e.Cell.Row.Cells["idCompra"].Value));
						//						Funcion.EjecutaSQL(cdsSeteoF, sSQLBloque, true);

					}
					
				}
			}

			if (e.Cell.Column.ToString() == "Ir")
			{	
				using (frmOrdenesDeCompra ORDC  = new frmOrdenesDeCompra((int)e.Cell.Row.Cells["idCompra"].Value))
				{
					if (DialogResult.OK == ORDC.ShowDialog())
					{
						

					}
				}
			}
		}

		private void btnExcel_Click(object sender, System.EventArgs e)
		{
			FuncionesProcedimientos.ExportaExcel(this.uGridOrdenesDeCompra);
		}

		private void uGridOrdenesDeCompra_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}

		private void optEstadoDentalis_ValueChanged(object sender, System.EventArgs e)
		{
			if (miAcceso.Imprimir) this.optEstadoDentalis.Value = 1;
      EstadosDentalis();
			this.btnBuscar_Click (sender, e);
		}
	}
}
