using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using Infragistics.Win.UltraWinGrid;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de frmFlujoEfectivo.
	/// </summary>
	public class frmFlujoEfectivo : System.Windows.Forms.Form
	{
		private System.Windows.Forms.Label label4;
		private System.Windows.Forms.DateTimePicker dtpPeriodo;
		private System.Windows.Forms.Button btnVer;
		private Infragistics.Win.UltraWinGrid.UltraGrid uGridSaldo;
		private Infragistics.Win.UltraWinGrid.UltraGrid uGridIngresos;
		private Infragistics.Win.UltraWinGrid.UltraGrid uGridSalidas;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource1;
		private C1.Data.C1DataSet cdsSeteoF;
		private System.Windows.Forms.Button btnExcel;
		private System.Windows.Forms.ContextMenu mnuExportar;
		private System.Windows.Forms.MenuItem mnuSI;
		private System.Windows.Forms.MenuItem mnuIngresos;
		private System.Windows.Forms.MenuItem mnuEgresos;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public frmFlujoEfectivo()
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
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmFlujoEfectivo));
			Infragistics.Win.Appearance appearance1 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("", -1);
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("", -1);
			Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance11 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance12 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance13 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance14 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance15 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance16 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance17 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand3 = new Infragistics.Win.UltraWinGrid.UltraGridBand("", -1);
			Infragistics.Win.Appearance appearance18 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance19 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance20 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance21 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance22 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance23 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance24 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Column 0");
			this.label4 = new System.Windows.Forms.Label();
			this.dtpPeriodo = new System.Windows.Forms.DateTimePicker();
			this.btnVer = new System.Windows.Forms.Button();
			this.uGridSaldo = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.uGridIngresos = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.uGridSalidas = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource1 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.btnExcel = new System.Windows.Forms.Button();
			this.mnuExportar = new System.Windows.Forms.ContextMenu();
			this.mnuSI = new System.Windows.Forms.MenuItem();
			this.mnuIngresos = new System.Windows.Forms.MenuItem();
			this.mnuEgresos = new System.Windows.Forms.MenuItem();
			((System.ComponentModel.ISupportInitialize)(this.uGridSaldo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridIngresos)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridSalidas)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			this.SuspendLayout();
			// 
			// label4
			// 
			this.label4.AutoSize = true;
			this.label4.BackColor = System.Drawing.Color.Transparent;
			this.label4.Location = new System.Drawing.Point(8, 10);
			this.label4.Name = "label4";
			this.label4.Size = new System.Drawing.Size(42, 17);
			this.label4.TabIndex = 734;
			this.label4.Text = "Periodo";
			this.label4.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// dtpPeriodo
			// 
			this.dtpPeriodo.CustomFormat = "MM/yyyy";
			this.dtpPeriodo.Format = System.Windows.Forms.DateTimePickerFormat.Custom;
			this.dtpPeriodo.Location = new System.Drawing.Point(64, 8);
			this.dtpPeriodo.MinDate = new System.DateTime(2015, 1, 1, 0, 0, 0, 0);
			this.dtpPeriodo.Name = "dtpPeriodo";
			this.dtpPeriodo.Size = new System.Drawing.Size(96, 21);
			this.dtpPeriodo.TabIndex = 735;
			this.dtpPeriodo.Value = new System.DateTime(2022, 9, 1, 0, 0, 0, 0);
			this.dtpPeriodo.ValueChanged += new System.EventHandler(this.dtpPeriodo_ValueChanged);
			// 
			// btnVer
			// 
			this.btnVer.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(230)), ((System.Byte)(230)), ((System.Byte)(230)));
			this.btnVer.CausesValidation = false;
			this.btnVer.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.btnVer.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(50)), ((System.Byte)(50)), ((System.Byte)(50)));
			this.btnVer.Image = ((System.Drawing.Image)(resources.GetObject("btnVer.Image")));
			this.btnVer.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnVer.Location = new System.Drawing.Point(184, 7);
			this.btnVer.Name = "btnVer";
			this.btnVer.Size = new System.Drawing.Size(80, 23);
			this.btnVer.TabIndex = 736;
			this.btnVer.Text = "Ver";
			this.btnVer.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnVer.Click += new System.EventHandler(this.btnVer_Click);
			// 
			// uGridSaldo
			// 
			this.uGridSaldo.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.uGridSaldo.Cursor = System.Windows.Forms.Cursors.Default;
			appearance1.BackColor = System.Drawing.Color.White;
			this.uGridSaldo.DisplayLayout.Appearance = appearance1;
			ultraGridBand1.SummaryFooterCaption = "";
			this.uGridSaldo.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.uGridSaldo.DisplayLayout.GroupByBox.Hidden = true;
			appearance2.ForeColor = System.Drawing.Color.Black;
			appearance2.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridSaldo.DisplayLayout.Override.ActiveRowAppearance = appearance2;
			this.uGridSaldo.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
			this.uGridSaldo.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.uGridSaldo.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance3.BackColor = System.Drawing.Color.Transparent;
			this.uGridSaldo.DisplayLayout.Override.CardAreaAppearance = appearance3;
			appearance4.ForeColor = System.Drawing.Color.Black;
			appearance4.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridSaldo.DisplayLayout.Override.CellAppearance = appearance4;
			this.uGridSaldo.DisplayLayout.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.Edit;
			appearance5.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance5.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance5.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance5.FontData.BoldAsString = "True";
			appearance5.FontData.Name = "Arial";
			appearance5.FontData.SizeInPoints = 8F;
			appearance5.ForeColor = System.Drawing.Color.White;
			appearance5.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridSaldo.DisplayLayout.Override.HeaderAppearance = appearance5;
			appearance6.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance6.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance6.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridSaldo.DisplayLayout.Override.RowAlternateAppearance = appearance6;
			appearance7.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance7.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance7.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridSaldo.DisplayLayout.Override.RowSelectorAppearance = appearance7;
			appearance8.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance8.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance8.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridSaldo.DisplayLayout.Override.SelectedRowAppearance = appearance8;
			this.uGridSaldo.DisplayLayout.ViewStyleBand = Infragistics.Win.UltraWinGrid.ViewStyleBand.OutlookGroupBy;
			this.uGridSaldo.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridSaldo.Location = new System.Drawing.Point(8, 40);
			this.uGridSaldo.Name = "uGridSaldo";
			this.uGridSaldo.Size = new System.Drawing.Size(1348, 80);
			this.uGridSaldo.TabIndex = 737;
			this.uGridSaldo.AfterCellUpdate += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.uGridSaldo_AfterCellUpdate);
			this.uGridSaldo.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.uGridSaldo_InitializeLayout);
			// 
			// uGridIngresos
			// 
			this.uGridIngresos.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.uGridIngresos.Cursor = System.Windows.Forms.Cursors.Default;
			appearance9.BackColor = System.Drawing.Color.White;
			this.uGridIngresos.DisplayLayout.Appearance = appearance9;
			ultraGridBand2.SummaryFooterCaption = "";
			this.uGridIngresos.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			this.uGridIngresos.DisplayLayout.GroupByBox.Hidden = true;
			appearance10.ForeColor = System.Drawing.Color.Black;
			appearance10.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridIngresos.DisplayLayout.Override.ActiveRowAppearance = appearance10;
			this.uGridIngresos.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
			this.uGridIngresos.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.uGridIngresos.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance11.BackColor = System.Drawing.Color.Transparent;
			this.uGridIngresos.DisplayLayout.Override.CardAreaAppearance = appearance11;
			appearance12.ForeColor = System.Drawing.Color.Black;
			appearance12.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridIngresos.DisplayLayout.Override.CellAppearance = appearance12;
			this.uGridIngresos.DisplayLayout.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.Edit;
			this.uGridIngresos.DisplayLayout.Override.GroupByColumnsHidden = Infragistics.Win.DefaultableBoolean.True;
			appearance13.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance13.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance13.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance13.FontData.BoldAsString = "True";
			appearance13.FontData.Name = "Arial";
			appearance13.FontData.SizeInPoints = 8F;
			appearance13.ForeColor = System.Drawing.Color.White;
			appearance13.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridIngresos.DisplayLayout.Override.HeaderAppearance = appearance13;
			this.uGridIngresos.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortSingle;
			appearance14.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance14.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance14.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridIngresos.DisplayLayout.Override.RowAlternateAppearance = appearance14;
			appearance15.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance15.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance15.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridIngresos.DisplayLayout.Override.RowSelectorAppearance = appearance15;
			appearance16.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance16.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance16.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridIngresos.DisplayLayout.Override.SelectedRowAppearance = appearance16;
			this.uGridIngresos.DisplayLayout.ViewStyleBand = Infragistics.Win.UltraWinGrid.ViewStyleBand.OutlookGroupBy;
			this.uGridIngresos.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridIngresos.Location = new System.Drawing.Point(8, 136);
			this.uGridIngresos.Name = "uGridIngresos";
			this.uGridIngresos.Size = new System.Drawing.Size(1348, 216);
			this.uGridIngresos.TabIndex = 738;
			this.uGridIngresos.AfterCellUpdate += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.uGridIngresos_AfterCellUpdate);
			this.uGridIngresos.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.uGridIngresos_InitializeLayout);
			// 
			// uGridSalidas
			// 
			this.uGridSalidas.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.uGridSalidas.Cursor = System.Windows.Forms.Cursors.Default;
			appearance17.BackColor = System.Drawing.Color.White;
			this.uGridSalidas.DisplayLayout.Appearance = appearance17;
			ultraGridBand3.SummaryFooterCaption = "";
			this.uGridSalidas.DisplayLayout.BandsSerializer.Add(ultraGridBand3);
			this.uGridSalidas.DisplayLayout.GroupByBox.Hidden = true;
			appearance18.ForeColor = System.Drawing.Color.Black;
			appearance18.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridSalidas.DisplayLayout.Override.ActiveRowAppearance = appearance18;
			this.uGridSalidas.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
			this.uGridSalidas.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.uGridSalidas.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance19.BackColor = System.Drawing.Color.Transparent;
			this.uGridSalidas.DisplayLayout.Override.CardAreaAppearance = appearance19;
			appearance20.ForeColor = System.Drawing.Color.Black;
			appearance20.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridSalidas.DisplayLayout.Override.CellAppearance = appearance20;
			this.uGridSalidas.DisplayLayout.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.Edit;
			appearance21.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance21.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance21.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance21.FontData.BoldAsString = "True";
			appearance21.FontData.Name = "Arial";
			appearance21.FontData.SizeInPoints = 8F;
			appearance21.ForeColor = System.Drawing.Color.White;
			appearance21.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridSalidas.DisplayLayout.Override.HeaderAppearance = appearance21;
			this.uGridSalidas.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortSingle;
			appearance22.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance22.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance22.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridSalidas.DisplayLayout.Override.RowAlternateAppearance = appearance22;
			appearance23.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance23.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance23.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridSalidas.DisplayLayout.Override.RowSelectorAppearance = appearance23;
			appearance24.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance24.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance24.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridSalidas.DisplayLayout.Override.SelectedRowAppearance = appearance24;
			this.uGridSalidas.DisplayLayout.ViewStyleBand = Infragistics.Win.UltraWinGrid.ViewStyleBand.OutlookGroupBy;
			this.uGridSalidas.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridSalidas.Location = new System.Drawing.Point(8, 368);
			this.uGridSalidas.Name = "uGridSalidas";
			this.uGridSalidas.Size = new System.Drawing.Size(1348, 216);
			this.uGridSalidas.TabIndex = 739;
			this.uGridSalidas.AfterCellUpdate += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.uGridSalidas_AfterCellUpdate);
			this.uGridSalidas.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.uGridSalidas_InitializeLayout);
			// 
			// ultraDataSource1
			// 
			this.ultraDataSource1.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn1});
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
			// btnExcel
			// 
			this.btnExcel.Image = ((System.Drawing.Image)(resources.GetObject("btnExcel.Image")));
			this.btnExcel.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnExcel.Location = new System.Drawing.Point(1280, 8);
			this.btnExcel.Name = "btnExcel";
			this.btnExcel.Size = new System.Drawing.Size(75, 24);
			this.btnExcel.TabIndex = 957;
			this.btnExcel.Text = "&Excel";
			this.btnExcel.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnExcel.Click += new System.EventHandler(this.btnExcel_Click);
			// 
			// mnuExportar
			// 
			this.mnuExportar.MenuItems.AddRange(new System.Windows.Forms.MenuItem[] {
																																								this.mnuSI,
																																								this.mnuIngresos,
																																								this.mnuEgresos});
			// 
			// mnuSI
			// 
			this.mnuSI.Index = 0;
			this.mnuSI.Text = "Saldos Iniciales";
			this.mnuSI.Click += new System.EventHandler(this.mnuSI_Click);
			// 
			// mnuIngresos
			// 
			this.mnuIngresos.Index = 1;
			this.mnuIngresos.Text = "Ingresos";
			this.mnuIngresos.Click += new System.EventHandler(this.mnuIngresos_Click);
			// 
			// mnuEgresos
			// 
			this.mnuEgresos.Index = 2;
			this.mnuEgresos.Text = "Egresos";
			this.mnuEgresos.Click += new System.EventHandler(this.mnuEgresos_Click);
			// 
			// frmFlujoEfectivo
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 14);
			this.ClientSize = new System.Drawing.Size(1364, 612);
			this.Controls.Add(this.btnExcel);
			this.Controls.Add(this.uGridSalidas);
			this.Controls.Add(this.uGridIngresos);
			this.Controls.Add(this.uGridSaldo);
			this.Controls.Add(this.btnVer);
			this.Controls.Add(this.label4);
			this.Controls.Add(this.dtpPeriodo);
			this.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.KeyPreview = true;
			this.Name = "frmFlujoEfectivo";
			this.Text = "Flujo De Efectivo";
			this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.frmFlujoEfectivo_KeyDown);
			this.Load += new System.EventHandler(this.frmFlujoEfectivo_Load);
			((System.ComponentModel.ISupportInitialize)(this.uGridSaldo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridIngresos)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridSalidas)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void frmFlujoEfectivo_Load(object sender, System.EventArgs e)
		{
			if (!Funcion.Permiso("1804", cdsSeteoF))
			{				
				MessageBox.Show("No puede ingresar a Flujo efectivo ", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				BeginInvoke(new MethodInvoker(UnloadMe));
				return;
			}

			this.dtpPeriodo.Value = DateTime.Parse(DateTime.DaysInMonth(DateTime.Now.Year, DateTime.Now.Month).ToString() + "/" + DateTime.Now.Month.ToString() + "/" + DateTime.Now.Year.ToString());
		}

		private void btnVer_Click(object sender, System.EventArgs e)
		{
			this.uGridSaldo.DataSource = null;
			this.uGridIngresos.DataSource = null;
			this.uGridSalidas.DataSource = null;

			FuncionesProcedimientos.dsGeneralModulo(string.Format("Exec Flujo_Consulta '{0}', 1", this.dtpPeriodo.Text.ToString()), this.uGridSaldo);
			DiseñoGrid(this.uGridSaldo, false, false);
			FuncionesProcedimientos.dsGeneralModulo(string.Format("Exec Flujo_Consulta '{0}', 2", this.dtpPeriodo.Text.ToString()), this.uGridIngresos);
			DiseñoGrid(this.uGridIngresos, true, true);			
			FuncionesProcedimientos.dsGeneralModulo(string.Format("Exec Flujo_Consulta '{0}', 3", this.dtpPeriodo.Text.ToString()), this.uGridSalidas);
			DiseñoGrid(this.uGridSalidas, true, true);
		}

		private void DiseñoGrid(Infragistics.Win.UltraWinGrid.UltraGrid ultraGrid1, bool bSumatoria, bool bAgrupado)
		{
			if (ultraGrid1.Rows.Count > 0)
			{				
				if (bAgrupado) ultraGrid1.DisplayLayout.Bands[0].SortedColumns.Add("Tipo", false, false);	

				ultraGrid1.DisplayLayout.Bands[0].Columns["PosIS"].Hidden = true;
				ultraGrid1.DisplayLayout.Bands[0].Columns["PosMov"].Hidden = true;
								
				ultraGrid1.DisplayLayout.Bands[0].Columns["Tipo"].Width = 200;
				ultraGrid1.DisplayLayout.Bands[0].Columns["Movimiento"].Width = 200;
				ultraGrid1.DisplayLayout.Bands[0].Columns["Total"].Width = 80;

				ultraGrid1.DisplayLayout.Bands[0].Columns["Tipo"].Header.Caption = "Tipo";				
				ultraGrid1.DisplayLayout.Bands[0].Columns["Movimiento"].Header.Caption = "Movimiento";
				ultraGrid1.DisplayLayout.Bands[0].Columns["Total"].Header.Caption = "Total";
				
				ultraGrid1.DisplayLayout.Bands[0].Columns["Tipo"].CellActivation = Activation.ActivateOnly;
				ultraGrid1.DisplayLayout.Bands[0].Columns["Movimiento"].CellActivation = Activation.ActivateOnly;
				ultraGrid1.DisplayLayout.Bands[0].Columns["Total"].CellActivation = Activation.Disabled;

				ultraGrid1.DisplayLayout.Bands[0].Columns["Tipo"].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Left;
				ultraGrid1.DisplayLayout.Bands[0].Columns["Movimiento"].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Left;
				ultraGrid1.DisplayLayout.Bands[0].Columns["Total"].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right;

				ultraGrid1.DisplayLayout.Bands[0].Columns["Total"].Format = "#,##0.00";
				
				ultraGrid1.DisplayLayout.Bands[0].Columns["Total"].CellAppearance.BackColor = Color.LightYellow;
				ultraGrid1.DisplayLayout.Bands[0].Columns["Total"].CellAppearance.BackColor2 = Color.LightYellow;

				if (bSumatoria)
				{
					ultraGrid1.DisplayLayout.Bands[0].Summaries.Add("Total", Infragistics.Win.UltraWinGrid.SummaryType.Sum, ultraGrid1.DisplayLayout.Bands[0].Columns["Total"], Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn);
					ultraGrid1.DisplayLayout.Bands[0].Summaries["Total"].Appearance.TextHAlign = Infragistics.Win.HAlign.Right;
					ultraGrid1.DisplayLayout.Bands[0].Summaries["Total"].DisplayFormat = "{0: #,##0.00}";
				}

				int iContColumnas = (int)ultraGrid1.DisplayLayout.Bands[0].Columns.Count;

				int iContador = 5;

				string nombrCol = "";
				
				for (int j = iContador; j < iContColumnas; j++)
				{
					nombrCol = ultraGrid1.DisplayLayout.Bands[0].Columns[j].Header.Caption;

					ultraGrid1.DisplayLayout.Bands[0].Columns[nombrCol].CellActivation = Activation.AllowEdit;
					ultraGrid1.DisplayLayout.Bands[0].Columns[nombrCol].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right;
					ultraGrid1.DisplayLayout.Bands[0].Columns[nombrCol].Width = 75;
					ultraGrid1.DisplayLayout.Bands[0].Columns[nombrCol].Format = "#,##0.00";
					
					if (j%2 == 0) 
					{
						ultraGrid1.DisplayLayout.Bands[0].Columns[nombrCol].CellAppearance.BackColor = Color.LightBlue;
						ultraGrid1.DisplayLayout.Bands[0].Columns[nombrCol].CellAppearance.BackColor2 = Color.LightBlue;
					}

					if (bSumatoria)
					{
						ultraGrid1.DisplayLayout.Bands[0].Summaries.Add(nombrCol, Infragistics.Win.UltraWinGrid.SummaryType.Sum, ultraGrid1.DisplayLayout.Bands[0].Columns[nombrCol], Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn);
						ultraGrid1.DisplayLayout.Bands[0].Summaries[nombrCol].Appearance.TextHAlign = Infragistics.Win.HAlign.Right;
						ultraGrid1.DisplayLayout.Bands[0].Summaries[nombrCol].DisplayFormat = "{0: #,##0.00}";
					}						
				}

				foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in ultraGrid1.Rows.All)
				{
					iContador = 5;

					for (int j = iContador; j < iContColumnas; j++)
					{
						nombrCol = ultraGrid1.DisplayLayout.Bands[0].Columns[j].Header.Caption;

						if (dr.Cells[nombrCol].Value != System.DBNull.Value)
						{
							dr.Cells["Total"].Value = Convert.ToDecimal(dr.Cells["Total"].Value) + 
								Convert.ToDecimal(dr.Cells[nombrCol].Value);
						}
					}
				}

//				if (bAgrupado)
//				{
//					ultraGrid1.DisplayLayout.Bands[0].SortedColumns.Add("Tipo", false, true);			
//					ultraGrid1.Rows.ExpandAll(true);
//				}				
			}
		}

		private void ActualizaValores (DateTime dtFecha, decimal dValor, string sMovimiento)
		{
			string sSQL = string.Format("Exec FlujoActualizaValoresMovimientos '{0}', {1}, '{2}'", dtFecha.ToString("yyyyMMdd"), dValor, sMovimiento);
			Funcion.EjecutaSQL(cdsSeteoF, sSQL);
		}

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);		
		}

		private void Sumatorias(Infragistics.Win.UltraWinGrid.CellEventArgs e, string sColumna)
		{			
			e.Cell.Row.Cells["Total"].Value = 
				Math.Round(Convert.ToDecimal(e.Cell.Row.Cells["Total"].Value), 2) +
				Math.Round(Convert.ToDecimal(e.Cell.Row.Cells[sColumna].Value), 2);
		}

		private void uGridSaldo_AfterCellUpdate(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			if (e.Cell.Column.ToString() != "Total")
			{
				string sColumna = e.Cell.Column.Header.Caption;

				if (e.Cell.Column.ToString() == sColumna)
				{
					Sumatorias(e, sColumna);
					ActualizaValores (DateTime.Parse(sColumna), Convert.ToDecimal(e.Cell.Row.Cells[sColumna].Value), e.Cell.Row.Cells["Movimiento"].Value.ToString());
				}
			}
		}

		private void uGridIngresos_AfterCellUpdate(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			if (e.Cell.Column.ToString() != "Total")
			{
				string sColumna = e.Cell.Column.Header.Caption;

				if (e.Cell.Column.ToString() == sColumna)
				{
					Sumatorias(e, sColumna);
					ActualizaValores (DateTime.Parse(sColumna), Convert.ToDecimal(e.Cell.Row.Cells[sColumna].Value), e.Cell.Row.Cells["Movimiento"].Value.ToString());
				}
			}
		}

		private void uGridSalidas_AfterCellUpdate(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			if (e.Cell.Column.ToString() != "Total")
			{
				string sColumna = e.Cell.Column.Header.Caption;

				if (e.Cell.Column.ToString() == sColumna)
				{
					Sumatorias(e, sColumna);
					ActualizaValores (DateTime.Parse(sColumna), Convert.ToDecimal(e.Cell.Row.Cells[sColumna].Value), e.Cell.Row.Cells["Movimiento"].Value.ToString());
				}
			}
		}

		private void uGridSaldo_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
//			string sColumna = e.Layout.Bands[0].Columns[].Header.Caption;
//			FuncionesProcedimientos.FormatoGrid(e, sColumna, 2);
		}

		private void uGridIngresos_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
			//string sColumna = e.Cell.Column.Header.Caption;
//			FuncionesProcedimientos.FormatoGrid(e, sColumna, 2);
		}

		private void uGridSalidas_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
//			string sColumna = e.Cell.Column.Header.Caption;
//			FuncionesProcedimientos.FormatoGrid(e, sColumna, 2);
		}

		private void frmFlujoEfectivo_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Escape) this.Close();		
		}

		private void dtpPeriodo_ValueChanged(object sender, System.EventArgs e)
		{
		
		}

		private void mnuSI_Click(object sender, System.EventArgs e)
		{
			FuncionesProcedimientos.ExportaExcel(this.uGridSaldo);
		}

		private void mnuIngresos_Click(object sender, System.EventArgs e)
		{
			FuncionesProcedimientos.ExportaExcel(this.uGridIngresos);
		}

		private void mnuEgresos_Click(object sender, System.EventArgs e)
		{
			FuncionesProcedimientos.ExportaExcel(this.uGridSalidas);
		}

		private void btnExcel_Click(object sender, System.EventArgs e)
		{
			this.mnuExportar.Show(this.btnExcel, new Point(1, 25));
		}
	}
}
