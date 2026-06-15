using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;
using System.Data.SqlClient;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de frmImportaRetencionesBancarias.
	/// </summary>
	public class frmImportaRetencionesBancarias : System.Windows.Forms.Form
	{
		private System.Windows.Forms.Label label7;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtBusqueda;
		private System.Windows.Forms.Button btnVer;
		private System.Windows.Forms.Button btnSubirArchivos;
		private System.Windows.Forms.Button btnExcel;
		private System.Windows.Forms.GroupBox groupBox1;
		private System.Windows.Forms.Button btnImportar;
		private C1.Data.C1DataSet cdsSeteoF;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtRuc;
		private System.Windows.Forms.Label label4;
		private System.Windows.Forms.Label label5;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtHasta;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtDesde;
		public Infragistics.Win.UltraWinGrid.UltraGrid uGridRetenciones;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource1;
		private System.Windows.Forms.Label lblContador;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public frmImportaRetencionesBancarias()
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
		private void UnloadMe()
		{
			this.Close();
		}

		#region Código generado por el Diseñador de Windows Forms
		/// <summary>
		/// Método necesario para admitir el Diseñador. No se puede modificar
		/// el contenido del método con el editor de código.
		/// </summary>
		private void InitializeComponent()
		{
			Infragistics.Win.Appearance appearance1 = new Infragistics.Win.Appearance();
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmImportaRetencionesBancarias));
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idRetencionesBancarias");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("RUC");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cliente");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Serie");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Secuencial");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Autorizacion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Caducidad");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Fecha");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("TotalIVA");
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("TotalRenta");
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn11 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Total");
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn12 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("IVA");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn13 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CodigoIVA");
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn14 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("PorcentajeIVA");
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn15 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("BaseIVA");
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn16 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ValorIVA");
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn17 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Renta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn18 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CodigoRenta");
			Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn19 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("PorcentajeRenta");
			Infragistics.Win.Appearance appearance11 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn20 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("BaseRenta");
			Infragistics.Win.Appearance appearance12 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn21 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ValorRenta");
			Infragistics.Win.Appearance appearance13 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn22 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaSistema");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn23 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Usuario");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn24 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Estacion");
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings1 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "BaseIVA", 14, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "BaseIVA", 14, true);
			Infragistics.Win.Appearance appearance14 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings2 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "ValorIVA", 15, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "ValorIVA", 15, true);
			Infragistics.Win.Appearance appearance15 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings3 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "BaseRenta", 19, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "BaseRenta", 19, true);
			Infragistics.Win.Appearance appearance16 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings4 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "ValorRenta", 20, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "ValorRenta", 20, true);
			Infragistics.Win.Appearance appearance17 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance18 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance19 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance20 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance21 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance22 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance23 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idRetencionesBancarias");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("RUC");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn3 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Cliente");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn4 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Serie");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn5 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Secuencial");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn6 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Autorizacion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn7 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Caducidad");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn8 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Fecha");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn9 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("TotalIVA");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn10 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("TotalRenta");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn11 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Total");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn12 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("IVA");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn13 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("CodigoIVA");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn14 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("PorcentajeIVA");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn15 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("BaseIVA");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn16 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("ValorIVA");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn17 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Renta");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn18 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("CodigoRenta");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn19 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("PorcentajeRenta");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn20 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("BaseRenta");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn21 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("ValorRenta");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn22 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("FechaSistema");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn23 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Usuario");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn24 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Estacion");
			Infragistics.Win.Appearance appearance24 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance25 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton1 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.Appearance appearance26 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton2 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			this.label7 = new System.Windows.Forms.Label();
			this.txtBusqueda = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.btnVer = new System.Windows.Forms.Button();
			this.btnSubirArchivos = new System.Windows.Forms.Button();
			this.uGridRetenciones = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource1 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.btnImportar = new System.Windows.Forms.Button();
			this.btnExcel = new System.Windows.Forms.Button();
			this.groupBox1 = new System.Windows.Forms.GroupBox();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.txtRuc = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label4 = new System.Windows.Forms.Label();
			this.label5 = new System.Windows.Forms.Label();
			this.dtHasta = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.dtDesde = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.lblContador = new System.Windows.Forms.Label();
			((System.ComponentModel.ISupportInitialize)(this.txtBusqueda)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridRetenciones)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtRuc)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtHasta)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtDesde)).BeginInit();
			this.SuspendLayout();
			// 
			// label7
			// 
			this.label7.AutoSize = true;
			this.label7.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label7.Location = new System.Drawing.Point(8, 36);
			this.label7.Name = "label7";
			this.label7.Size = new System.Drawing.Size(79, 16);
			this.label7.TabIndex = 743;
			this.label7.Text = "Cedula/Cliente";
			this.label7.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtBusqueda
			// 
			appearance1.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtBusqueda.Appearance = appearance1;
			this.txtBusqueda.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtBusqueda.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtBusqueda.Location = new System.Drawing.Point(104, 34);
			this.txtBusqueda.Name = "txtBusqueda";
			this.txtBusqueda.Size = new System.Drawing.Size(248, 21);
			this.txtBusqueda.TabIndex = 739;
			// 
			// btnVer
			// 
			this.btnVer.Image = ((System.Drawing.Image)(resources.GetObject("btnVer.Image")));
			this.btnVer.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnVer.Location = new System.Drawing.Point(368, 32);
			this.btnVer.Name = "btnVer";
			this.btnVer.Size = new System.Drawing.Size(75, 24);
			this.btnVer.TabIndex = 740;
			this.btnVer.Text = "&Ver";
			this.btnVer.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnVer.Click += new System.EventHandler(this.btnVer_Click);
			// 
			// btnSubirArchivos
			// 
			this.btnSubirArchivos.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btnSubirArchivos.CausesValidation = false;
			this.btnSubirArchivos.Image = ((System.Drawing.Image)(resources.GetObject("btnSubirArchivos.Image")));
			this.btnSubirArchivos.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnSubirArchivos.Location = new System.Drawing.Point(1256, 0);
			this.btnSubirArchivos.Name = "btnSubirArchivos";
			this.btnSubirArchivos.Size = new System.Drawing.Size(80, 23);
			this.btnSubirArchivos.TabIndex = 741;
			this.btnSubirArchivos.Text = "Importar";
			this.btnSubirArchivos.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// uGridRetenciones
			// 
			this.uGridRetenciones.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.uGridRetenciones.Cursor = System.Windows.Forms.Cursors.Default;
			this.uGridRetenciones.DataSource = this.ultraDataSource1;
			appearance2.BackColor = System.Drawing.Color.White;
			appearance2.ForeColor = System.Drawing.Color.Black;
			appearance2.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridRetenciones.DisplayLayout.Appearance = appearance2;
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn2.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn2.Header.Caption = "R.U.C.";
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridColumn2.Width = 100;
			ultraGridColumn3.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn3.Header.VisiblePosition = 2;
			ultraGridColumn3.Width = 200;
			ultraGridColumn4.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn4.Header.VisiblePosition = 3;
			ultraGridColumn4.Width = 60;
			ultraGridColumn5.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn5.Header.VisiblePosition = 4;
			ultraGridColumn5.Width = 90;
			ultraGridColumn6.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn6.Header.VisiblePosition = 5;
			ultraGridColumn6.Width = 120;
			ultraGridColumn7.Header.VisiblePosition = 6;
			ultraGridColumn7.Width = 80;
			ultraGridColumn8.Header.VisiblePosition = 7;
			ultraGridColumn8.Width = 80;
			ultraGridColumn9.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance3.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn9.CellAppearance = appearance3;
			ultraGridColumn9.Format = "#,##0.00";
			ultraGridColumn9.Header.Caption = "Total IVA";
			ultraGridColumn9.Header.VisiblePosition = 8;
			ultraGridColumn9.Hidden = true;
			ultraGridColumn9.Width = 90;
			ultraGridColumn10.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance4.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn10.CellAppearance = appearance4;
			ultraGridColumn10.Format = "#,##0.00";
			ultraGridColumn10.Header.Caption = "Total Renta";
			ultraGridColumn10.Header.VisiblePosition = 9;
			ultraGridColumn10.Hidden = true;
			ultraGridColumn10.Width = 90;
			ultraGridColumn11.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance5.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn11.CellAppearance = appearance5;
			ultraGridColumn11.Format = "#,##0.00";
			ultraGridColumn11.Header.VisiblePosition = 10;
			ultraGridColumn11.Hidden = true;
			ultraGridColumn11.Width = 90;
			ultraGridColumn12.Header.VisiblePosition = 11;
			ultraGridColumn12.Hidden = true;
			ultraGridColumn13.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance6.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn13.CellAppearance = appearance6;
			ultraGridColumn13.Header.Caption = "Codigo IVA";
			ultraGridColumn13.Header.VisiblePosition = 12;
			ultraGridColumn13.Width = 80;
			ultraGridColumn14.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance7.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn14.CellAppearance = appearance7;
			ultraGridColumn14.Format = "#,##0.00";
			ultraGridColumn14.Header.Caption = "% IVA";
			ultraGridColumn14.Header.VisiblePosition = 13;
			ultraGridColumn14.Width = 50;
			ultraGridColumn15.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance8.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn15.CellAppearance = appearance8;
			ultraGridColumn15.Format = "#,##0.00";
			ultraGridColumn15.Header.Caption = "Base IVA";
			ultraGridColumn15.Header.VisiblePosition = 14;
			ultraGridColumn15.Width = 90;
			ultraGridColumn16.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance9.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn16.CellAppearance = appearance9;
			ultraGridColumn16.Format = "#,##0.00";
			ultraGridColumn16.Header.Caption = "Valor IVA";
			ultraGridColumn16.Header.VisiblePosition = 15;
			ultraGridColumn16.Width = 90;
			ultraGridColumn17.Header.VisiblePosition = 16;
			ultraGridColumn17.Hidden = true;
			ultraGridColumn18.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance10.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn18.CellAppearance = appearance10;
			ultraGridColumn18.Header.Caption = "Codigo IR";
			ultraGridColumn18.Header.VisiblePosition = 17;
			ultraGridColumn18.Width = 80;
			ultraGridColumn19.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance11.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn19.CellAppearance = appearance11;
			ultraGridColumn19.Format = "#,##0.00";
			ultraGridColumn19.Header.Caption = "% IR";
			ultraGridColumn19.Header.VisiblePosition = 18;
			ultraGridColumn19.Width = 50;
			ultraGridColumn20.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance12.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn20.CellAppearance = appearance12;
			ultraGridColumn20.Format = "#,##0.00";
			ultraGridColumn20.Header.Caption = "Base IR";
			ultraGridColumn20.Header.VisiblePosition = 19;
			ultraGridColumn20.Width = 90;
			ultraGridColumn21.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance13.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn21.CellAppearance = appearance13;
			ultraGridColumn21.Format = "#,##0.00";
			ultraGridColumn21.Header.Caption = "Valor IR";
			ultraGridColumn21.Header.VisiblePosition = 20;
			ultraGridColumn21.Width = 90;
			ultraGridColumn22.Format = "dd/MM/yyyy HH:mm";
			ultraGridColumn22.Header.Caption = "Fecha Subida";
			ultraGridColumn22.Header.VisiblePosition = 21;
			ultraGridColumn22.Width = 120;
			ultraGridColumn23.Header.VisiblePosition = 22;
			ultraGridColumn23.Width = 100;
			ultraGridColumn24.Header.VisiblePosition = 23;
			ultraGridColumn24.Width = 100;
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
																										 ultraGridColumn17,
																										 ultraGridColumn18,
																										 ultraGridColumn19,
																										 ultraGridColumn20,
																										 ultraGridColumn21,
																										 ultraGridColumn22,
																										 ultraGridColumn23,
																										 ultraGridColumn24});
			ultraGridBand1.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
			ultraGridBand1.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.True;
			ultraGridBand1.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.False;
			appearance14.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings1.Appearance = appearance14;
			summarySettings1.DisplayFormat = "{0: #,##0.00}";
			summarySettings1.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance15.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings2.Appearance = appearance15;
			summarySettings2.DisplayFormat = "{0: #,##0.00}";
			summarySettings2.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance16.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings3.Appearance = appearance16;
			summarySettings3.DisplayFormat = "{0: #,##0.00}";
			summarySettings3.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance17.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings4.Appearance = appearance17;
			summarySettings4.DisplayFormat = "{0: #,##0.00}";
			summarySettings4.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			ultraGridBand1.Summaries.AddRange(new Infragistics.Win.UltraWinGrid.SummarySettings[] {
																																															summarySettings1,
																																															summarySettings2,
																																															summarySettings3,
																																															summarySettings4});
			ultraGridBand1.SummaryFooterCaption = "";
			this.uGridRetenciones.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			appearance18.ForeColor = System.Drawing.Color.Black;
			appearance18.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridRetenciones.DisplayLayout.Override.ActiveRowAppearance = appearance18;
			this.uGridRetenciones.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
			this.uGridRetenciones.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.True;
			this.uGridRetenciones.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.False;
			appearance19.BackColor = System.Drawing.Color.Transparent;
			this.uGridRetenciones.DisplayLayout.Override.CardAreaAppearance = appearance19;
			appearance20.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance20.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance20.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance20.FontData.BoldAsString = "True";
			appearance20.FontData.Name = "Arial";
			appearance20.FontData.SizeInPoints = 8F;
			appearance20.ForeColor = System.Drawing.Color.White;
			appearance20.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridRetenciones.DisplayLayout.Override.HeaderAppearance = appearance20;
			this.uGridRetenciones.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti;
			appearance21.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance21.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance21.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridRetenciones.DisplayLayout.Override.RowAlternateAppearance = appearance21;
			appearance22.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance22.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance22.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridRetenciones.DisplayLayout.Override.RowSelectorAppearance = appearance22;
			appearance23.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance23.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance23.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridRetenciones.DisplayLayout.Override.SelectedRowAppearance = appearance23;
			this.uGridRetenciones.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridRetenciones.Location = new System.Drawing.Point(6, 80);
			this.uGridRetenciones.Name = "uGridRetenciones";
			this.uGridRetenciones.Size = new System.Drawing.Size(1116, 272);
			this.uGridRetenciones.TabIndex = 742;
			this.uGridRetenciones.BeforeRowsDeleted += new Infragistics.Win.UltraWinGrid.BeforeRowsDeletedEventHandler(this.uGridRetenciones_BeforeRowsDeleted);
			// 
			// ultraDataSource1
			// 
			ultraDataColumn1.DataType = typeof(int);
			ultraDataColumn7.DataType = typeof(System.DateTime);
			ultraDataColumn7.DefaultValue = new System.DateTime(((long)(0)));
			ultraDataColumn8.DataType = typeof(System.DateTime);
			ultraDataColumn8.DefaultValue = new System.DateTime(((long)(0)));
			ultraDataColumn9.DataType = typeof(System.Decimal);
			ultraDataColumn9.DefaultValue = new System.Decimal(new int[] {
																																		 0,
																																		 0,
																																		 0,
																																		 0});
			ultraDataColumn10.DataType = typeof(System.Decimal);
			ultraDataColumn10.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
			ultraDataColumn11.DataType = typeof(System.Decimal);
			ultraDataColumn11.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
			ultraDataColumn12.DataType = typeof(bool);
			ultraDataColumn12.DefaultValue = false;
			ultraDataColumn14.DataType = typeof(System.Decimal);
			ultraDataColumn14.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
			ultraDataColumn15.DataType = typeof(System.Decimal);
			ultraDataColumn15.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
			ultraDataColumn16.DataType = typeof(System.Decimal);
			ultraDataColumn16.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
			ultraDataColumn17.DataType = typeof(bool);
			ultraDataColumn17.DefaultValue = false;
			ultraDataColumn19.DataType = typeof(System.Decimal);
			ultraDataColumn19.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
			ultraDataColumn20.DataType = typeof(System.Decimal);
			ultraDataColumn20.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
			ultraDataColumn21.DataType = typeof(System.Decimal);
			ultraDataColumn21.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
			ultraDataColumn22.DataType = typeof(System.DateTime);
			ultraDataColumn22.DefaultValue = new System.DateTime(((long)(0)));
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
																																 ultraDataColumn16,
																																 ultraDataColumn17,
																																 ultraDataColumn18,
																																 ultraDataColumn19,
																																 ultraDataColumn20,
																																 ultraDataColumn21,
																																 ultraDataColumn22,
																																 ultraDataColumn23,
																																 ultraDataColumn24});
			// 
			// btnImportar
			// 
			this.btnImportar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btnImportar.CausesValidation = false;
			this.btnImportar.Image = ((System.Drawing.Image)(resources.GetObject("btnImportar.Image")));
			this.btnImportar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnImportar.Location = new System.Drawing.Point(1040, 7);
			this.btnImportar.Name = "btnImportar";
			this.btnImportar.Size = new System.Drawing.Size(80, 23);
			this.btnImportar.TabIndex = 745;
			this.btnImportar.Text = "Importar";
			this.btnImportar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnImportar.Click += new System.EventHandler(this.btnImportar_Click);
			// 
			// btnExcel
			// 
			this.btnExcel.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btnExcel.Image = ((System.Drawing.Image)(resources.GetObject("btnExcel.Image")));
			this.btnExcel.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnExcel.Location = new System.Drawing.Point(1040, 32);
			this.btnExcel.Name = "btnExcel";
			this.btnExcel.Size = new System.Drawing.Size(80, 24);
			this.btnExcel.TabIndex = 746;
			this.btnExcel.Text = "Exportar";
			this.btnExcel.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnExcel.Click += new System.EventHandler(this.btnExcel_Click);
			// 
			// groupBox1
			// 
			this.groupBox1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.groupBox1.Location = new System.Drawing.Point(0, 64);
			this.groupBox1.Name = "groupBox1";
			this.groupBox1.Size = new System.Drawing.Size(1128, 8);
			this.groupBox1.TabIndex = 747;
			this.groupBox1.TabStop = false;
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
			// txtRuc
			// 
			this.txtRuc.AcceptsTab = true;
			this.txtRuc.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			appearance24.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtRuc.Appearance = appearance24;
			this.txtRuc.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtRuc.Enabled = false;
			this.txtRuc.Location = new System.Drawing.Point(904, 34);
			this.txtRuc.Name = "txtRuc";
			this.txtRuc.Size = new System.Drawing.Size(128, 21);
			this.txtRuc.TabIndex = 748;
			this.txtRuc.Visible = false;
			// 
			// label4
			// 
			this.label4.AutoSize = true;
			this.label4.BackColor = System.Drawing.Color.Transparent;
			this.label4.Location = new System.Drawing.Point(192, 10);
			this.label4.Name = "label4";
			this.label4.Size = new System.Drawing.Size(34, 16);
			this.label4.TabIndex = 752;
			this.label4.Text = "Hasta";
			this.label4.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label5
			// 
			this.label5.AutoSize = true;
			this.label5.BackColor = System.Drawing.Color.Transparent;
			this.label5.Location = new System.Drawing.Point(8, 10);
			this.label5.Name = "label5";
			this.label5.Size = new System.Drawing.Size(37, 16);
			this.label5.TabIndex = 751;
			this.label5.Text = "Desde";
			this.label5.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// dtHasta
			// 
			appearance25.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtHasta.Appearance = appearance25;
			this.dtHasta.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton1.Action = Infragistics.Win.UltraWinSchedule.DateButtonAction.AddToSelection;
			dateButton1.Caption = "Today";
			dateButton1.Type = Infragistics.Win.UltraWinSchedule.DateButtonType.Custom;
			this.dtHasta.DateButtons.Add(dateButton1);
			this.dtHasta.Format = "dd/MM/yyyy";
			this.dtHasta.Location = new System.Drawing.Point(248, 8);
			this.dtHasta.Name = "dtHasta";
			this.dtHasta.NonAutoSizeHeight = 23;
			this.dtHasta.Size = new System.Drawing.Size(104, 21);
			this.dtHasta.SpinButtonsVisible = true;
			this.dtHasta.TabIndex = 750;
			this.dtHasta.Value = ((object)(resources.GetObject("dtHasta.Value")));
			// 
			// dtDesde
			// 
			appearance26.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtDesde.Appearance = appearance26;
			this.dtDesde.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton2.Action = Infragistics.Win.UltraWinSchedule.DateButtonAction.AddToSelection;
			dateButton2.Caption = "Today";
			dateButton2.Type = Infragistics.Win.UltraWinSchedule.DateButtonType.Custom;
			this.dtDesde.DateButtons.Add(dateButton2);
			this.dtDesde.Format = "dd/MM/yyyy";
			this.dtDesde.Location = new System.Drawing.Point(64, 8);
			this.dtDesde.Name = "dtDesde";
			this.dtDesde.NonAutoSizeHeight = 23;
			this.dtDesde.Size = new System.Drawing.Size(104, 21);
			this.dtDesde.SpinButtonsVisible = true;
			this.dtDesde.TabIndex = 749;
			this.dtDesde.Value = ((object)(resources.GetObject("dtDesde.Value")));
			// 
			// lblContador
			// 
			this.lblContador.AutoSize = true;
			this.lblContador.Font = new System.Drawing.Font("Microsoft Sans Serif", 11F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblContador.ForeColor = System.Drawing.Color.Firebrick;
			this.lblContador.Location = new System.Drawing.Point(456, 34);
			this.lblContador.Name = "lblContador";
			this.lblContador.Size = new System.Drawing.Size(0, 20);
			this.lblContador.TabIndex = 753;
			this.lblContador.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// frmImportaRetencionesBancarias
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(1128, 358);
			this.Controls.Add(this.lblContador);
			this.Controls.Add(this.label4);
			this.Controls.Add(this.label5);
			this.Controls.Add(this.label7);
			this.Controls.Add(this.dtHasta);
			this.Controls.Add(this.dtDesde);
			this.Controls.Add(this.txtRuc);
			this.Controls.Add(this.groupBox1);
			this.Controls.Add(this.btnExcel);
			this.Controls.Add(this.btnImportar);
			this.Controls.Add(this.txtBusqueda);
			this.Controls.Add(this.btnVer);
			this.Controls.Add(this.btnSubirArchivos);
			this.Controls.Add(this.uGridRetenciones);
			this.Name = "frmImportaRetencionesBancarias";
			this.Text = "Importa Retenciones Bancarias";
			this.Load += new System.EventHandler(this.frmImportaRetencionesBancarias_Load);
			((System.ComponentModel.ISupportInitialize)(this.txtBusqueda)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridRetenciones)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtRuc)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtHasta)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtDesde)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion


		private void frmImportaRetencionesBancarias_Load(object sender, System.EventArgs e)
		{
			if (!Funcion.Permiso("912", cdsSeteoF))
			{				
				MessageBox.Show("No puede ingresar a Flujo efectivo ", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				BeginInvoke(new MethodInvoker(UnloadMe));
				return;
			}
			
			
			DateTime dtFecha = DateTime.Parse("01/06/2016");

			this.dtDesde.CalendarInfo.MinDate = dtFecha;
			this.dtHasta.CalendarInfo.MinDate = dtFecha;

			this.dtDesde.Value = "01/" + DateTime.Now.Month.ToString() + "/" + DateTime.Now.Year.ToString();
			this.dtHasta.Value = DateTime.Today;

			this.dtDesde.CalendarInfo.MaxDate = Convert.ToDateTime(this.dtHasta.Value);
			this.dtHasta.CalendarInfo.MaxDate = Convert.ToDateTime(this.dtHasta.Value);
		}

		private void btnImportar_Click(object sender, System.EventArgs e)
		{
			int iCont = 0;

			try
			{
				OpenFileDialog sArchivo = new OpenFileDialog();

				sArchivo.Filter = "Microsoft Excel (*.xls)|*.xls";

				if (sArchivo.ShowDialog() == System.Windows.Forms.DialogResult.OK && sArchivo.FileName.Length > 0)
				{
					this.Cursor = Cursors.WaitCursor;                                  
										
					string sSecuencial = "";

					#region Valida Archivo
					string sSQLVD = "";

					foreach (DataRow row in FuncionesProcedimientos.ImportaDeExcel(sArchivo.FileName, "SELECT * FROM [Retenciones$]", "Retenciones").Tables[0].Rows)
					{							
						sSQLVD = "";

						#region Identificacion
						this.txtRuc.Text = row["RUC"].ToString();

						if (this.txtRuc.Text.ToString().Length < 13) this.txtRuc.Text = "0" + this.txtRuc.Text.ToString().Trim();

						if (!Validacion.vbIdentificacion(this.txtRuc, 4, cdsSeteoF)) 
						{
							sSQLVD = string.Format("Exec ValidaDuplicadosRetencionesBancarias '{0}', '{1}', '{2}', {3}, 'IDENTIFICACION NO VALIDA'",
								this.txtRuc.Text.ToString(), row["Serie"].ToString(), row["Secuencial"].ToString(), iCont);
							Funcion.EjecutaSQL(cdsSeteoF, sSQLVD, true);
						}

						sSQLVD = "";

						if (!Funcion.bEscalarSQL(cdsSeteoF, string.Format("Exec FacturacionValidaClienteExiste '{0}', {1}", this.txtRuc.Text.ToString(), false), true)) 
						{
							sSQLVD = string.Format("Exec ValidaDuplicadosRetencionesBancarias '{0}', '{1}', '{2}', {3}, 'CLIENTE NO EXISTE'",
								this.txtRuc.Text.ToString(), row["Serie"].ToString(), row["Secuencial"].ToString(), iCont);
							Funcion.EjecutaSQL(cdsSeteoF, sSQLVD, true);
						}
						#endregion Identificacion

						sSQLVD = "";

						#region Serie
						if (!Validacion.bNumeros(row["Serie"].ToString()))
						{
							sSQLVD = string.Format("Exec ValidaDuplicadosRetencionesBancarias '{0}', '{1}', '{2}', {3}, 'EL NUMERO DE SERIE DEBE TENER SOLO DIGITOS'",
								this.txtRuc.Text.ToString(), row["Serie"].ToString(), row["Secuencial"].ToString(), iCont);											
							Funcion.EjecutaSQL(cdsSeteoF, sSQLVD, true);
						}

						sSQLVD = "";

						if (row["Serie"].ToString().Length != 6)
						{
							sSQLVD = string.Format("Exec ValidaDuplicadosRetencionesBancarias '{0}', '{1}', '{2}', {3}, 'EL NUMERO DE SERIE DEBE TENER 6 DIGITOS'",
								this.txtRuc.Text.ToString(), row["Serie"].ToString(), row["Secuencial"].ToString(), iCont);											
							Funcion.EjecutaSQL(cdsSeteoF, sSQLVD, true);
						}

						sSQLVD = "";
						#endregion Serie

						#region Numero
						if (row["Secuencial"].ToString().Length == 0)
						{
							sSQLVD = string.Format("Exec ValidaDuplicadosRetencionesBancarias '{0}', '{1}', '{2}', {3}, 'INGRESE EL NUMERO DEL DOCUMENTO'",
								this.txtRuc.Text.ToString(), row["Serie"].ToString(), row["Secuencial"].ToString(), iCont);
							Funcion.EjecutaSQL(cdsSeteoF, sSQLVD, true);
						}
						
						sSQLVD = "";

						if (!Validacion.bNumeros(row["Secuencial"].ToString()))
						{
							sSQLVD = string.Format("Exec ValidaDuplicadosRetencionesBancarias '{0}', '{1}', '{2}', {3}, 'EL SECUENCIAL DEBE TENER SOLO DIGITOS'",
								this.txtRuc.Text.ToString(), row["Serie"].ToString(), row["Secuencial"].ToString(), iCont);
							Funcion.EjecutaSQL(cdsSeteoF, sSQLVD, true);
						}

						sSQLVD = "";

						char Pad = '0';
						sSecuencial = sSecuencial.ToString().PadLeft(9, Pad);

						if (sSecuencial.ToString().Length > 9)
						{
							sSQLVD = string.Format("Exec ValidaDuplicadosRetencionesBancarias '{0}', '{1}', '{2}', {3}, 'EL NUMERO DEL DOCUMENTO NO PUEDE TENER MAS DE 9 DIGITOS'",
								this.txtRuc.Text.ToString(), row["Serie"].ToString(), sSecuencial.ToString(), iCont);
							Funcion.EjecutaSQL(cdsSeteoF, sSQLVD, true);
						}

						if (sSecuencial.ToString().Length < 9)
						{
							sSQLVD = string.Format("Exec ValidaDuplicadosRetencionesBancarias '{0}', '{1}', '{2}', {3}, 'EL NUMERO DEL DOCUMENTO NO PUEDE TENER MENOS DE 9 DIGITOS'",
								this.txtRuc.Text.ToString(), row["Serie"].ToString(), sSecuencial.ToString(), iCont);
							Funcion.EjecutaSQL(cdsSeteoF, sSQLVD, true);
						}

						sSQLVD = "";
						#endregion Numero

						#region Fecha
						if (row["Fecha"].ToString().Length == 0)
						{
							sSQLVD = string.Format("Exec ValidaDuplicadosRetencionesBancarias '{0}', '{1}', '{2}', {3}, 'INGRESE LA FECHA DEL DOCUMENTO'",
								this.txtRuc.Text.ToString(), row["Serie"].ToString(), sSecuencial.ToString(), iCont);
							Funcion.EjecutaSQL(cdsSeteoF, sSQLVD, true);
						}
						
						sSQLVD = "";

						if (DateTime.Parse(row["Fecha"].ToString()) < DateTime.Parse("01/03/2018"))
						{						
							sSQLVD = string.Format("Exec ValidaDuplicadosRetencionesBancarias '{0}', '{1}', '{2}', {3}, 'NO PUEDE INGRESAR DOCUMENTOS CON FECHAS ANTERIORES AL 01/03/2018'",
								this.txtRuc.Text.ToString(), row["Serie"].ToString(), sSecuencial.ToString(), iCont);
							Funcion.EjecutaSQL(cdsSeteoF, sSQLVD, true);
						}	

						sSQLVD = "";

						if (DateTime.Parse(row["Fecha"].ToString()) > DateTime.Today)
						{						
							sSQLVD = string.Format("Exec ValidaDuplicadosRetencionesBancarias '{0}', '{1}', '{2}', {3}, 'NO PUEDE INGRESAR DOCUMENTOS CON FECHAS POSTERIORES AL DIA ACTUAL'",
								this.txtRuc.Text.ToString(), row["Serie"].ToString(), sSecuencial.ToString(), iCont);
							Funcion.EjecutaSQL(cdsSeteoF, sSQLVD, true);
						}

						sSQLVD = "";

						if (Funcion.ValidaPeriodo(Convert.ToDateTime(row["Fecha"]), cdsSeteoF, "VENTA"))
						{						
							sSQLVD = string.Format("Exec ValidaDuplicadosRetencionesBancarias '{0}', '{1}', '{2}', {3}, 'NO PUEDE INGRESAR DOCUMENTOS EN UN PERIODO CERRADO'",
								this.txtRuc.Text.ToString(), row["Serie"].ToString(), sSecuencial.ToString(), iCont);
							Funcion.EjecutaSQL(cdsSeteoF, sSQLVD, true);
						}						
						#endregion Fecha

						#region IVA
//						if (Decimal.Parse(row["PorcentajeIVA"].ToString()))
//						{
//
//						}
						#endregion IVA

						sSQLVD = "";

						#region Duplicados
						if (Funcion.iEscalarSQL(cdsSeteoF, string.Format("Select COUNT(*) From RetencionesBancarias Where RUC = '{0}' And Serie = '{1}' And Secuencial = '{2}'",
							this.txtRuc.Text.ToString(), row["Serie"].ToString(), sSecuencial.ToString())) > 0)
						{
							sSQLVD = string.Format("Exec ValidaDuplicadosRetencionesBancarias '{0}', '{1}', '{2}', {3}, 'REGISTRO DUPLICADO'",
								this.txtRuc.Text.ToString(), row["Serie"].ToString(), sSecuencial.ToString(), iCont);							
							Funcion.EjecutaSQL(cdsSeteoF, sSQLVD, true);
						}
						#endregion Duplicados
						
						this.txtRuc.Text = "";

						iCont++;
					}
					#endregion Valida Archivo							
					
					if (Funcion.iEscalarSQL(cdsSeteoF, "Select COUNT(*) From ValidacionRetencionesBancarias") > 0)
					{
						frmValidacionRetencionesBancarias VRB = new frmValidacionRetencionesBancarias();
						VRB.ShowDialog();

						Funcion.EjecutaSQL(cdsSeteoF, "Delete From ValidacionRetencionesBancarias");

						this.Cursor = Cursors.Default;

						return;
					}
					
					iCont = 0;

					#region Sube archivo
					foreach (DataRow row in FuncionesProcedimientos.ImportaDeExcel(sArchivo.FileName, "SELECT * FROM [Retenciones$]", "Retenciones").Tables[0].Rows)
					{						
						#region Guarda Informacion
						/*
						@RUC Varchar(15), @Serie Varchar(10), @Secuencial Varchar(20), @Autorizacion Varchar(60), @Caducidad Datetime, @Fecha Datetime, 5
//						@TotalIVA Decimal(18, 2), @TotalRenta Decimal(18, 2), @Total Decimal(18, 2), @IVA Bit, 
						@CodigoIVA Varchar(10), @PorcentajeIVA Decimal(18, 2), @BaseIVA Decimal(18, 2), @ValorIVA  Decimal(18, 2), 9
						//@Renta Bit, 
						@CodigoRenta Varchar(10), @PorcentajeRenta  Decimal(18, 2), @BaseRenta Decimal(18, 2), @ValorRenta Decimal(18, 2)  
						*/
						/*RUC	Serie	Secuencial	Autorizacion	Caducidad	Fecha	CodigoIVA	PorcentajeIVA	BaseIVA	ValorIVA	CodigoRenta	PorcentajeRenta	BaseRenta	ValorRenta
*/
						DateTime dtCaducidad = DateTime.Parse("01/01/2000");
						if (row["Caducidad"] != System.DBNull.Value) dtCaducidad = DateTime.Parse(row["Caducidad"].ToString());

						this.txtRuc.Text = row["RUC"].ToString();
						if (this.txtRuc.Text.ToString().Length < 13) this.txtRuc.Text = "0" + this.txtRuc.Text.ToString().Trim();

						sSecuencial = row["Secuencial"].ToString();
						char Pad = '0';
						sSecuencial = sSecuencial.ToString().PadLeft(9, Pad);
						//RUC	Serie	Secuencial	AUTORIZACION	Fecha	CodigoIVA	PorcentajeIVA	 BaseIVA 	 ValorIVA 	CodigoRenta	PorcentajeRenta	 BaseRenta 	 ValorRenta 

						string sSQLArchivo = string.Format("Exec ImportaRetencionesBancarias '{0}', '{1}', '{2}', '{3}', '{4}', '{5}', '{6}', {7}, {8}, {9}, '{10}', {11}, {12}, {13}",
							this.txtRuc.Text.ToString(), row["Serie"].ToString(), sSecuencial.ToString(), row["Autorizacion"].ToString(), dtCaducidad.ToString("yyyyMMdd"), DateTime.Parse(row["Fecha"].ToString()).ToString("yyyyMMdd"),
							//decimal.Parse(row["TotalIVA"].ToString()), decimal.Parse(row["TotalRenta"].ToString()), decimal.Parse(row["Total"].ToString()), int.Parse(row["IVA"].ToString()), 
							row["CodigoIVA"].ToString(), decimal.Parse(row["PorcentajeIVA"].ToString()), decimal.Parse(row["BaseIVA"].ToString()), decimal.Parse(row["ValorIVA"].ToString()),
							//int.Parse(row["Renta"].ToString()), 
							row["CodigoRenta"].ToString(), decimal.Parse(row["PorcentajeRenta"].ToString()), decimal.Parse(row["BaseRenta"].ToString()), decimal.Parse(row["ValorRenta"].ToString()));
						Funcion.EjecutaSQL(cdsSeteoF, sSQLArchivo);
						#endregion Guarda Informacion

						iCont++;
					}				
					#endregion Sube archivo

					MessageBox.Show(string.Format("Proceso Finalizado Correctamente, {0} Retenciones Importadas", iCont), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
		
					this.btnVer_Click(sender, e);
					this.Cursor = Cursors.Default;
				}
			}
			catch (Exception ex)
			{
				MessageBox.Show("Fila : " + iCont.ToString());
				MessageBox.Show(string.Format("Commit Exception Type : {0} {1}", ex.GetType(), ex.Message), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
				this.Cursor = Cursors.Default;
				return;
			}
		}

		private void btnVer_Click(object sender, System.EventArgs e)
		{
			string sSQL = string.Format("Exec ConsultaRetencionesBancarias '{0}', '{1}', '{2}'",
				Convert.ToDateTime(this.dtDesde.Value).ToString("yyyyMMdd"), Convert.ToDateTime(this.dtHasta.Value).ToString("yyyyMMdd"), this.txtBusqueda.Text.ToString());
			FuncionesProcedimientos.dsGeneral(sSQL, this.uGridRetenciones);

			this.lblContador.Text = this.uGridRetenciones.Rows.Count + " Registros Encontrados";
		}

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);	
		}

		private void btnExcel_Click(object sender, System.EventArgs e)
		{
			FuncionesProcedimientos.ExportaExcel(this.uGridRetenciones);
		}

		private void uGridRetenciones_BeforeRowsDeleted(object sender, Infragistics.Win.UltraWinGrid.BeforeRowsDeletedEventArgs e)
		{
			e.DisplayPromptMsg = false;

			if (DialogResult.Yes ==	MessageBox.Show("¿Esta seguro de borrar las lineas seleccionadas?", "Confirmación", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2)) 
			{
				for (int i = 0; i < e.Rows.Length; i++)
				{
					string sSQL = string.Format("Delete From RetencionesBancarias Where idRetencionesBancarias = {0}", 
						(int)e.Rows[i].Cells["idRetencionesBancarias"].Value);
					Funcion.EjecutaSQL(cdsSeteoF, sSQL, true);
				}
			}
			else e.Cancel = true;
		}
	}
}
