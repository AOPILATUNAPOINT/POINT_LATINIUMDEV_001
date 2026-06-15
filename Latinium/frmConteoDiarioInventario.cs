using System;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using CrystalDecisions;
using CrystalDecisions.CrystalReports.Engine;
using CrystalDecisions.ReportSource;
using CrystalDecisions.Shared;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de frmConteoDiarioInventario.
	/// </summary>
	public class frmConteoDiarioInventario : System.Windows.Forms.Form
	{
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource1;
		private System.Windows.Forms.Label label7;
		private System.Windows.Forms.Label label55;
		private System.Windows.Forms.Label label53;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtArticulo;
		private System.Windows.Forms.Button btnVer;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbBodega;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtFecha;
		private System.Windows.Forms.Label lblEstado;
		private System.Windows.Forms.Button btnGenerar;
		private System.Windows.Forms.Button btnImprimir;
		private System.Windows.Forms.Button btnCerrar;
		private System.Windows.Forms.Button btnSalir;
		private System.Windows.Forms.GroupBox groupBox1;
		private C1.Data.C1DataSet cdsSeteoF;
		private Infragistics.Win.UltraWinGrid.UltraGrid uGridConteo;
		private System.Windows.Forms.Label label12;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource2;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbSubGrupo;
		private System.Windows.Forms.Label lblContador;
		private System.Windows.Forms.Label label1;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtNotas;
		private System.Windows.Forms.Label lblVendedor;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbVendedor;
		private System.Windows.Forms.ContextMenu mnuImprimir;
		private System.Windows.Forms.MenuItem mnuImprimirInventario;
		private System.Windows.Forms.MenuItem mnuImprimirInforme;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtDia;
		private System.Windows.Forms.Button btnExcel;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public frmConteoDiarioInventario()
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
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idDetalleConteoDiarioInventario");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idArticulo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Articulo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Estado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("EstadoM");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Existencia", -1, null, 0, Infragistics.Win.UltraWinGrid.SortIndicator.Ascending, false);
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Exibicion");
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bodega");
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Conteo");
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn11 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Diferencias");
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn12 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Notas");
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings1 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Existencia", 6, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Existencia", 6, true);
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings2 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Conteo", 9, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Conteo", 9, true);
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings3 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Exibicion", 7, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Exibicion", 7, true);
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings4 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Bodega", 8, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Bodega", 8, true);
			Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings5 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Diferencias", 10, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Diferencias", 10, true);
			Infragistics.Win.Appearance appearance11 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance12 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance13 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance14 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance15 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance16 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance17 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance18 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance19 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idDetalleConteoDiarioInventario");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idArticulo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn3 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Codigo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn4 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Articulo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn5 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Estado");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn6 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("EstadoM");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn7 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Existencia");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn8 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Exibicion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn9 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Bodega");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn10 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Conteo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn11 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Diferencias");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn12 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Notas");
			Infragistics.Win.Appearance appearance20 = new Infragistics.Win.Appearance();
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmConteoDiarioInventario));
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn13 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bodega");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn14 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn15 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo", 0);
			Infragistics.Win.Appearance appearance21 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton1 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn13 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idSubgrupo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn14 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("SubGrupo");
			Infragistics.Win.Appearance appearance22 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand3 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn16 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idSubgrupo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn17 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("SubGrupo");
			Infragistics.Win.Appearance appearance23 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance24 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand4 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn18 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPersonal");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn19 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn20 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.Appearance appearance25 = new Infragistics.Win.Appearance();
			this.uGridConteo = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource1 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.label7 = new System.Windows.Forms.Label();
			this.label55 = new System.Windows.Forms.Label();
			this.label53 = new System.Windows.Forms.Label();
			this.txtArticulo = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.btnVer = new System.Windows.Forms.Button();
			this.cmbBodega = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.dtFecha = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.lblEstado = new System.Windows.Forms.Label();
			this.btnGenerar = new System.Windows.Forms.Button();
			this.btnImprimir = new System.Windows.Forms.Button();
			this.btnCerrar = new System.Windows.Forms.Button();
			this.btnSalir = new System.Windows.Forms.Button();
			this.groupBox1 = new System.Windows.Forms.GroupBox();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.label12 = new System.Windows.Forms.Label();
			this.ultraDataSource2 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.cmbSubGrupo = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.lblContador = new System.Windows.Forms.Label();
			this.txtNotas = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label1 = new System.Windows.Forms.Label();
			this.lblVendedor = new System.Windows.Forms.Label();
			this.cmbVendedor = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.mnuImprimir = new System.Windows.Forms.ContextMenu();
			this.mnuImprimirInventario = new System.Windows.Forms.MenuItem();
			this.mnuImprimirInforme = new System.Windows.Forms.MenuItem();
			this.txtDia = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.btnExcel = new System.Windows.Forms.Button();
			((System.ComponentModel.ISupportInitialize)(this.uGridConteo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtArticulo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbBodega)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFecha)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbSubGrupo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNotas)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbVendedor)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDia)).BeginInit();
			this.SuspendLayout();
			// 
			// uGridConteo
			// 
			this.uGridConteo.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.uGridConteo.Cursor = System.Windows.Forms.Cursors.Default;
			this.uGridConteo.DataSource = this.ultraDataSource1;
			appearance1.BackColor = System.Drawing.Color.White;
			this.uGridConteo.DisplayLayout.Appearance = appearance1;
			ultraGridColumn1.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn2.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridColumn2.Hidden = true;
			ultraGridColumn3.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn3.Header.VisiblePosition = 2;
			ultraGridColumn3.Width = 180;
			ultraGridColumn4.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn4.Header.Caption = "Artículo";
			ultraGridColumn4.Header.VisiblePosition = 3;
			ultraGridColumn4.Width = 250;
			ultraGridColumn5.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn5.Header.VisiblePosition = 7;
			ultraGridColumn5.Hidden = true;
			ultraGridColumn5.Width = 180;
			ultraGridColumn6.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn6.Header.Caption = "Estado";
			ultraGridColumn6.Header.VisiblePosition = 4;
			ultraGridColumn6.Width = 100;
			ultraGridColumn7.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance2.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn7.CellAppearance = appearance2;
			ultraGridColumn7.Format = "#,##0";
			ultraGridColumn7.Header.VisiblePosition = 5;
			ultraGridColumn7.Width = 75;
			ultraGridColumn8.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance3.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn8.CellAppearance = appearance3;
			ultraGridColumn8.Format = "#,##0";
			ultraGridColumn8.Header.Caption = "Exhibicion";
			ultraGridColumn8.Header.VisiblePosition = 8;
			ultraGridColumn8.Width = 75;
			ultraGridColumn9.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance4.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn9.CellAppearance = appearance4;
			ultraGridColumn9.Format = "#,##0";
			ultraGridColumn9.Header.VisiblePosition = 9;
			ultraGridColumn9.Width = 75;
			ultraGridColumn10.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance5.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn10.CellAppearance = appearance5;
			ultraGridColumn10.Format = "#,##0";
			ultraGridColumn10.Header.VisiblePosition = 6;
			ultraGridColumn10.Width = 75;
			ultraGridColumn11.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance6.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn11.CellAppearance = appearance6;
			ultraGridColumn11.Header.VisiblePosition = 10;
			ultraGridColumn12.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn12.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn12.Header.VisiblePosition = 11;
			ultraGridColumn12.Width = 400;
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
																										 ultraGridColumn12});
			appearance7.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings1.Appearance = appearance7;
			summarySettings1.DisplayFormat = "{0: #,##0}";
			summarySettings1.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance8.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings2.Appearance = appearance8;
			summarySettings2.DisplayFormat = "{0: #,##0}";
			summarySettings2.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance9.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings3.Appearance = appearance9;
			summarySettings3.DisplayFormat = "{0: #,##0}";
			summarySettings3.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance10.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings4.Appearance = appearance10;
			summarySettings4.DisplayFormat = "{0: #,##0}";
			summarySettings4.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance11.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings5.Appearance = appearance11;
			summarySettings5.DisplayFormat = "{0: #,##0}";
			summarySettings5.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			ultraGridBand1.Summaries.AddRange(new Infragistics.Win.UltraWinGrid.SummarySettings[] {
																																															summarySettings1,
																																															summarySettings2,
																																															summarySettings3,
																																															summarySettings4,
																																															summarySettings5});
			ultraGridBand1.SummaryFooterCaption = "";
			this.uGridConteo.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			appearance12.ForeColor = System.Drawing.Color.Black;
			appearance12.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridConteo.DisplayLayout.Override.ActiveRowAppearance = appearance12;
			this.uGridConteo.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
			this.uGridConteo.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.uGridConteo.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance13.BackColor = System.Drawing.Color.Transparent;
			this.uGridConteo.DisplayLayout.Override.CardAreaAppearance = appearance13;
			appearance14.ForeColor = System.Drawing.Color.Black;
			appearance14.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridConteo.DisplayLayout.Override.CellAppearance = appearance14;
			appearance15.ForeColor = System.Drawing.Color.Black;
			appearance15.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridConteo.DisplayLayout.Override.GroupByRowAppearance = appearance15;
			appearance16.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance16.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance16.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance16.FontData.BoldAsString = "True";
			appearance16.FontData.Name = "Arial";
			appearance16.FontData.SizeInPoints = 8.5F;
			appearance16.ForeColor = System.Drawing.Color.White;
			appearance16.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridConteo.DisplayLayout.Override.HeaderAppearance = appearance16;
			this.uGridConteo.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortSingle;
			appearance17.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance17.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance17.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridConteo.DisplayLayout.Override.RowAlternateAppearance = appearance17;
			appearance18.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance18.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance18.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridConteo.DisplayLayout.Override.RowSelectorAppearance = appearance18;
			appearance19.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance19.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance19.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridConteo.DisplayLayout.Override.SelectedRowAppearance = appearance19;
			this.uGridConteo.Font = new System.Drawing.Font("Tahoma", 8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridConteo.Location = new System.Drawing.Point(8, 184);
			this.uGridConteo.Name = "uGridConteo";
			this.uGridConteo.Size = new System.Drawing.Size(1080, 264);
			this.uGridConteo.TabIndex = 3;
			this.uGridConteo.KeyDown += new System.Windows.Forms.KeyEventHandler(this.uGridConteo_KeyDown);
			this.uGridConteo.AfterCellUpdate += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.uGridConteo_AfterCellUpdate);
			this.uGridConteo.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.uGridConteo_InitializeLayout);
			// 
			// ultraDataSource1
			// 
			ultraDataColumn1.DataType = typeof(int);
			ultraDataColumn2.DataType = typeof(int);
			ultraDataColumn5.DataType = typeof(int);
			ultraDataColumn7.DataType = typeof(int);
			ultraDataColumn7.DefaultValue = 0;
			ultraDataColumn8.DataType = typeof(int);
			ultraDataColumn8.DefaultValue = 0;
			ultraDataColumn9.DataType = typeof(int);
			ultraDataColumn9.DefaultValue = 0;
			ultraDataColumn10.DataType = typeof(int);
			ultraDataColumn10.DefaultValue = 0;
			ultraDataColumn11.DataType = typeof(int);
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
			// label7
			// 
			this.label7.AutoSize = true;
			this.label7.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label7.Location = new System.Drawing.Point(8, 107);
			this.label7.Name = "label7";
			this.label7.Size = new System.Drawing.Size(42, 17);
			this.label7.TabIndex = 376;
			this.label7.Text = "Artículo";
			this.label7.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label55
			// 
			this.label55.AutoSize = true;
			this.label55.BackColor = System.Drawing.Color.Transparent;
			this.label55.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label55.Location = new System.Drawing.Point(8, 43);
			this.label55.Name = "label55";
			this.label55.Size = new System.Drawing.Size(41, 17);
			this.label55.TabIndex = 375;
			this.label55.Text = "Bodega";
			this.label55.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label53
			// 
			this.label53.AutoSize = true;
			this.label53.BackColor = System.Drawing.Color.Transparent;
			this.label53.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label53.Location = new System.Drawing.Point(8, 10);
			this.label53.Name = "label53";
			this.label53.Size = new System.Drawing.Size(34, 17);
			this.label53.TabIndex = 374;
			this.label53.Text = "Fecha";
			this.label53.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtArticulo
			// 
			appearance20.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtArticulo.Appearance = appearance20;
			this.txtArticulo.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtArticulo.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtArticulo.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtArticulo.Location = new System.Drawing.Point(80, 104);
			this.txtArticulo.Name = "txtArticulo";
			this.txtArticulo.Size = new System.Drawing.Size(264, 22);
			this.txtArticulo.TabIndex = 372;
			this.txtArticulo.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtArticulo_KeyDown);
			// 
			// btnVer
			// 
			this.btnVer.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.btnVer.Image = ((System.Drawing.Image)(resources.GetObject("btnVer.Image")));
			this.btnVer.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnVer.Location = new System.Drawing.Point(376, 104);
			this.btnVer.Name = "btnVer";
			this.btnVer.Size = new System.Drawing.Size(75, 24);
			this.btnVer.TabIndex = 373;
			this.btnVer.Text = "Ver";
			this.btnVer.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnVer.Click += new System.EventHandler(this.btnVer_Click);
			// 
			// cmbBodega
			// 
			this.cmbBodega.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbBodega.Cursor = System.Windows.Forms.Cursors.Default;
			ultraGridColumn13.Header.VisiblePosition = 2;
			ultraGridColumn14.Header.VisiblePosition = 1;
			ultraGridColumn14.Width = 180;
			ultraGridColumn15.Header.VisiblePosition = 0;
			ultraGridBand2.Columns.AddRange(new object[] {
																										 ultraGridColumn13,
																										 ultraGridColumn14,
																										 ultraGridColumn15});
			this.cmbBodega.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			this.cmbBodega.DisplayMember = "Nombre";
			this.cmbBodega.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbBodega.Enabled = false;
			this.cmbBodega.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbBodega.Location = new System.Drawing.Point(80, 40);
			this.cmbBodega.MaxDropDownItems = 30;
			this.cmbBodega.Name = "cmbBodega";
			this.cmbBodega.Size = new System.Drawing.Size(264, 22);
			this.cmbBodega.TabIndex = 370;
			this.cmbBodega.ValueMember = "Bodega";
			this.cmbBodega.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cmbBodega_KeyDown);
			this.cmbBodega.ValueChanged += new System.EventHandler(this.cmbBodega_ValueChanged);
			this.cmbBodega.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.cmbBodega_InitializeLayout);
			// 
			// dtFecha
			// 
			appearance21.FontData.Name = "Tahoma";
			appearance21.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtFecha.Appearance = appearance21;
			this.dtFecha.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton1.Caption = "Today";
			this.dtFecha.DateButtons.Add(dateButton1);
			this.dtFecha.Format = "dd/MM/yyyy";
			this.dtFecha.Location = new System.Drawing.Point(80, 8);
			this.dtFecha.Name = "dtFecha";
			this.dtFecha.NonAutoSizeHeight = 23;
			this.dtFecha.Size = new System.Drawing.Size(112, 21);
			this.dtFecha.SpinButtonsVisible = true;
			this.dtFecha.TabIndex = 371;
			this.dtFecha.Value = ((object)(resources.GetObject("dtFecha.Value")));
			this.dtFecha.KeyDown += new System.Windows.Forms.KeyEventHandler(this.dtFecha_KeyDown);
			this.dtFecha.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.dtFecha_KeyPress);
			this.dtFecha.BeforeDropDown += new System.ComponentModel.CancelEventHandler(this.dtFecha_BeforeDropDown);
			this.dtFecha.ValueChanged += new System.EventHandler(this.dtFecha_ValueChanged);
			// 
			// lblEstado
			// 
			this.lblEstado.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.lblEstado.AutoSize = true;
			this.lblEstado.Font = new System.Drawing.Font("Tahoma", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblEstado.ForeColor = System.Drawing.Color.Firebrick;
			this.lblEstado.ImeMode = System.Windows.Forms.ImeMode.NoControl;
			this.lblEstado.Location = new System.Drawing.Point(464, 40);
			this.lblEstado.Name = "lblEstado";
			this.lblEstado.Size = new System.Drawing.Size(0, 23);
			this.lblEstado.TabIndex = 377;
			this.lblEstado.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// btnGenerar
			// 
			this.btnGenerar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btnGenerar.CausesValidation = false;
			this.btnGenerar.Enabled = false;
			this.btnGenerar.Image = ((System.Drawing.Image)(resources.GetObject("btnGenerar.Image")));
			this.btnGenerar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnGenerar.Location = new System.Drawing.Point(1008, 104);
			this.btnGenerar.Name = "btnGenerar";
			this.btnGenerar.Size = new System.Drawing.Size(77, 23);
			this.btnGenerar.TabIndex = 378;
			this.btnGenerar.Text = "Generar";
			this.btnGenerar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnGenerar.Click += new System.EventHandler(this.btnGenerar_Click);
			// 
			// btnImprimir
			// 
			this.btnImprimir.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btnImprimir.CausesValidation = false;
			this.btnImprimir.Enabled = false;
			this.btnImprimir.Image = ((System.Drawing.Image)(resources.GetObject("btnImprimir.Image")));
			this.btnImprimir.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnImprimir.Location = new System.Drawing.Point(1008, 40);
			this.btnImprimir.Name = "btnImprimir";
			this.btnImprimir.Size = new System.Drawing.Size(77, 23);
			this.btnImprimir.TabIndex = 379;
			this.btnImprimir.Text = "&Imprimir";
			this.btnImprimir.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnImprimir.Click += new System.EventHandler(this.btnImprimir_Click);
			// 
			// btnCerrar
			// 
			this.btnCerrar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btnCerrar.CausesValidation = false;
			this.btnCerrar.Enabled = false;
			this.btnCerrar.Image = ((System.Drawing.Image)(resources.GetObject("btnCerrar.Image")));
			this.btnCerrar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnCerrar.Location = new System.Drawing.Point(1008, 8);
			this.btnCerrar.Name = "btnCerrar";
			this.btnCerrar.Size = new System.Drawing.Size(77, 23);
			this.btnCerrar.TabIndex = 381;
			this.btnCerrar.Text = "&Cerrar";
			this.btnCerrar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnCerrar.Click += new System.EventHandler(this.btnCerrar_Click);
			// 
			// btnSalir
			// 
			this.btnSalir.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btnSalir.CausesValidation = false;
			this.btnSalir.Image = ((System.Drawing.Image)(resources.GetObject("btnSalir.Image")));
			this.btnSalir.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnSalir.Location = new System.Drawing.Point(1008, 72);
			this.btnSalir.Name = "btnSalir";
			this.btnSalir.Size = new System.Drawing.Size(77, 23);
			this.btnSalir.TabIndex = 382;
			this.btnSalir.Text = "&Salir";
			this.btnSalir.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnSalir.Click += new System.EventHandler(this.btnSalir_Click);
			// 
			// groupBox1
			// 
			this.groupBox1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.groupBox1.Location = new System.Drawing.Point(8, 168);
			this.groupBox1.Name = "groupBox1";
			this.groupBox1.Size = new System.Drawing.Size(1096, 8);
			this.groupBox1.TabIndex = 383;
			this.groupBox1.TabStop = false;
			this.groupBox1.Enter += new System.EventHandler(this.groupBox1_Enter);
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
			// label12
			// 
			this.label12.AutoSize = true;
			this.label12.BackColor = System.Drawing.Color.Transparent;
			this.label12.Location = new System.Drawing.Point(8, 75);
			this.label12.Name = "label12";
			this.label12.Size = new System.Drawing.Size(56, 16);
			this.label12.TabIndex = 385;
			this.label12.Text = "SubGrupo";
			this.label12.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// ultraDataSource2
			// 
			ultraDataColumn13.DataType = typeof(int);
			this.ultraDataSource2.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn13,
																																 ultraDataColumn14});
			// 
			// cmbSubGrupo
			// 
			appearance22.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbSubGrupo.Appearance = appearance22;
			this.cmbSubGrupo.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbSubGrupo.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbSubGrupo.DataSource = this.ultraDataSource2;
			ultraGridColumn16.Header.VisiblePosition = 0;
			ultraGridColumn16.Hidden = true;
			ultraGridColumn17.Header.Caption = "Subgrupo";
			ultraGridColumn17.Header.VisiblePosition = 1;
			ultraGridColumn17.Width = 300;
			ultraGridBand3.Columns.AddRange(new object[] {
																										 ultraGridColumn16,
																										 ultraGridColumn17});
			this.cmbSubGrupo.DisplayLayout.BandsSerializer.Add(ultraGridBand3);
			this.cmbSubGrupo.DisplayMember = "Subgrupo";
			this.cmbSubGrupo.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbSubGrupo.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbSubGrupo.Location = new System.Drawing.Point(80, 72);
			this.cmbSubGrupo.Name = "cmbSubGrupo";
			this.cmbSubGrupo.Size = new System.Drawing.Size(264, 22);
			this.cmbSubGrupo.TabIndex = 384;
			this.cmbSubGrupo.ValueMember = "idSubgrupo";
			this.cmbSubGrupo.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cmbSubGrupo_KeyDown);
			this.cmbSubGrupo.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.cmbSubGrupo_InitializeLayout);
			// 
			// lblContador
			// 
			this.lblContador.AutoSize = true;
			this.lblContador.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblContador.ForeColor = System.Drawing.Color.Firebrick;
			this.lblContador.ImeMode = System.Windows.Forms.ImeMode.NoControl;
			this.lblContador.Location = new System.Drawing.Point(472, 107);
			this.lblContador.Name = "lblContador";
			this.lblContador.Size = new System.Drawing.Size(0, 17);
			this.lblContador.TabIndex = 386;
			this.lblContador.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// txtNotas
			// 
			this.txtNotas.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			appearance23.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtNotas.Appearance = appearance23;
			this.txtNotas.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtNotas.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNotas.Enabled = false;
			this.txtNotas.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtNotas.Location = new System.Drawing.Point(736, 8);
			this.txtNotas.Multiline = true;
			this.txtNotas.Name = "txtNotas";
			this.txtNotas.Size = new System.Drawing.Size(264, 66);
			this.txtNotas.TabIndex = 387;
			this.txtNotas.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtNotas_KeyDown);
			// 
			// label1
			// 
			this.label1.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.label1.AutoSize = true;
			this.label1.BackColor = System.Drawing.Color.Transparent;
			this.label1.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label1.Location = new System.Drawing.Point(672, 33);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(33, 17);
			this.label1.TabIndex = 388;
			this.label1.Text = "Notas";
			this.label1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// lblVendedor
			// 
			this.lblVendedor.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.lblVendedor.AutoSize = true;
			this.lblVendedor.Location = new System.Drawing.Point(672, 83);
			this.lblVendedor.Name = "lblVendedor";
			this.lblVendedor.Size = new System.Drawing.Size(56, 16);
			this.lblVendedor.TabIndex = 390;
			this.lblVendedor.Text = "Jefe Local";
			// 
			// cmbVendedor
			// 
			this.cmbVendedor.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			appearance24.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbVendedor.Appearance = appearance24;
			this.cmbVendedor.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbVendedor.Cursor = System.Windows.Forms.Cursors.Default;
			ultraGridColumn18.Header.VisiblePosition = 0;
			ultraGridColumn18.Hidden = true;
			ultraGridColumn19.Header.VisiblePosition = 1;
			ultraGridColumn19.Width = 335;
			ultraGridColumn20.Header.VisiblePosition = 2;
			ultraGridBand4.Columns.AddRange(new object[] {
																										 ultraGridColumn18,
																										 ultraGridColumn19,
																										 ultraGridColumn20});
			this.cmbVendedor.DisplayLayout.BandsSerializer.Add(ultraGridBand4);
			this.cmbVendedor.DisplayMember = "Nombre";
			this.cmbVendedor.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbVendedor.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbVendedor.Enabled = false;
			this.cmbVendedor.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbVendedor.Location = new System.Drawing.Point(736, 80);
			this.cmbVendedor.Name = "cmbVendedor";
			this.cmbVendedor.Size = new System.Drawing.Size(264, 22);
			this.cmbVendedor.TabIndex = 389;
			this.cmbVendedor.ValueMember = "idPersonal";
			// 
			// mnuImprimir
			// 
			this.mnuImprimir.MenuItems.AddRange(new System.Windows.Forms.MenuItem[] {
																																								this.mnuImprimirInventario,
																																								this.mnuImprimirInforme});
			// 
			// mnuImprimirInventario
			// 
			this.mnuImprimirInventario.Enabled = false;
			this.mnuImprimirInventario.Index = 0;
			this.mnuImprimirInventario.Text = "Inventario";
			this.mnuImprimirInventario.Click += new System.EventHandler(this.mnuImprimirInventario_Click);
			// 
			// mnuImprimirInforme
			// 
			this.mnuImprimirInforme.Enabled = false;
			this.mnuImprimirInforme.Index = 1;
			this.mnuImprimirInforme.Text = "Informe";
			this.mnuImprimirInforme.Click += new System.EventHandler(this.mnuImprimirInforme_Click);
			// 
			// txtDia
			// 
			appearance25.BackColorDisabled = System.Drawing.SystemColors.Control;
			appearance25.BackColorDisabled2 = System.Drawing.Color.Transparent;
			appearance25.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance25.BackHatchStyle = Infragistics.Win.BackHatchStyle.Horizontal;
			appearance25.BorderAlpha = Infragistics.Win.Alpha.Transparent;
			appearance25.BorderColor = System.Drawing.Color.Transparent;
			appearance25.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtDia.Appearance = appearance25;
			this.txtDia.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtDia.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtDia.Enabled = false;
			this.txtDia.Font = new System.Drawing.Font("Tahoma", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtDia.Location = new System.Drawing.Point(224, 7);
			this.txtDia.Name = "txtDia";
			this.txtDia.Size = new System.Drawing.Size(120, 23);
			this.txtDia.TabIndex = 391;
			// 
			// btnExcel
			// 
			this.btnExcel.Enabled = false;
			this.btnExcel.Image = ((System.Drawing.Image)(resources.GetObject("btnExcel.Image")));
			this.btnExcel.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnExcel.Location = new System.Drawing.Point(1008, 136);
			this.btnExcel.Name = "btnExcel";
			this.btnExcel.Size = new System.Drawing.Size(75, 24);
			this.btnExcel.TabIndex = 979;
			this.btnExcel.Text = "Excel";
			this.btnExcel.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnExcel.Click += new System.EventHandler(this.btnExcel_Click);
			// 
			// frmConteoDiarioInventario
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(1096, 454);
			this.Controls.Add(this.btnExcel);
			this.Controls.Add(this.txtDia);
			this.Controls.Add(this.lblVendedor);
			this.Controls.Add(this.label1);
			this.Controls.Add(this.label7);
			this.Controls.Add(this.label55);
			this.Controls.Add(this.label53);
			this.Controls.Add(this.label12);
			this.Controls.Add(this.cmbVendedor);
			this.Controls.Add(this.txtNotas);
			this.Controls.Add(this.lblContador);
			this.Controls.Add(this.cmbSubGrupo);
			this.Controls.Add(this.groupBox1);
			this.Controls.Add(this.btnGenerar);
			this.Controls.Add(this.btnImprimir);
			this.Controls.Add(this.btnCerrar);
			this.Controls.Add(this.btnSalir);
			this.Controls.Add(this.lblEstado);
			this.Controls.Add(this.txtArticulo);
			this.Controls.Add(this.btnVer);
			this.Controls.Add(this.cmbBodega);
			this.Controls.Add(this.dtFecha);
			this.Controls.Add(this.uGridConteo);
			this.KeyPreview = true;
			this.Name = "frmConteoDiarioInventario";
			this.Text = "Conteo Diario de Inventario";
			this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.frmConteoDiarioInventario_KeyDown);
			this.Load += new System.EventHandler(this.frmConteoDiarioInventario_Load);
			((System.ComponentModel.ISupportInitialize)(this.uGridConteo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtArticulo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbBodega)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFecha)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbSubGrupo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNotas)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbVendedor)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDia)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		int iBodegaPredef = 0;
		int idConteo = 0;
		int iEstado = 0;
		System.DayOfWeek  sDia = new DayOfWeek(); 
		private Acceso miAcceso;
		private void UnloadMe()
		{
			this.Close();
		}

		private void frmConteoDiarioInventario_Load(object sender, System.EventArgs e)
		{	
			miAcceso = new Acceso(cdsSeteoF, "0327");

			if (!Funcion.Permiso("1944", cdsSeteoF))
			{				
				MessageBox.Show("No puede ingresar a conteo diario de Inventario.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				BeginInvoke(new MethodInvoker(UnloadMe));
				return;
			}

			string sScripP ;

			this.dtFecha.CalendarInfo.MinDate = DateTime.Parse("19/07/2017");
			this.dtFecha.CalendarInfo.MaxDate = DateTime.Today;			
			this.dtFecha.Value = DateTime.Today;
			sDia = Convert.ToDateTime(dtFecha.Value).DayOfWeek ;  
			this.txtDia.Value =  Convert.ToString(sDia); 

			string sSQL = string.Format("Exec ListaBodegas {0}, 47", MenuLatinium.IdUsuario);
			this.cmbBodega.DataSource = Funcion.dvProcedimiento(cdsSeteoF, sSQL);

			sSQL = string.Format("Exec BodegaAsignadaPorUsuario {0}", MenuLatinium.IdUsuario);		
			iBodegaPredef = Funcion.iEscalarSQL(cdsSeteoF, sSQL);

			if(Convert.ToString(this.txtDia.Value) != Convert.ToString("MONDAY"))	sScripP = "Select idSubgrupo, SubGrupo From ArticuloSubGrupo Where idSubGrupo In (23, 105, 126) Order By SubGrupo";
			else sScripP = "Select idSubgrupo, SubGrupo From ArticuloSubGrupo  Order By SubGrupo"; 
			this.cmbSubGrupo.DataSource = Funcion.dvProcedimiento(cdsSeteoF, sScripP);

			if (iBodegaPredef > 0) this.cmbBodega.Value = iBodegaPredef;
			else this.cmbBodega.Enabled = true;		
		}

		private void ListaConteo()
		{
			int idSubgrupo = 0;
			if (this.cmbSubGrupo.ActiveRow != null) idSubgrupo = (int)this.cmbSubGrupo.Value;

			string sSQL = string.Format("Exec ListaConteoDiarioInventario {0}, '{1}', {2}", idConteo, this.txtArticulo.Text.ToString(), idSubgrupo);
			FuncionesProcedimientos.dsGeneral(sSQL, this.uGridConteo);

			this.lblContador.Text = this.uGridConteo.Rows.Count + " REGISTROS ENCONTRADOS";
			
			foreach (Infragistics.Win.UltraWinGrid.UltraGridRow row in uGridConteo.Rows)
			{
				if (row.Cells.Exists("Conteo") && row.Cells.Exists("Existencia") && row.Cells.Exists("Diferencias"))
				{
					if (row.Cells["Conteo"].Value != System.DBNull.Value &&
						row.Cells["Existencia"].Value != System.DBNull.Value)
					{
						int conteo = Convert.ToInt32(row.Cells["Conteo"].Value);
						int existencia = Convert.ToInt32(row.Cells["Existencia"].Value);
						int diferencia = conteo - existencia;

						string resultado = diferencia > 0 ? "+" + diferencia.ToString() : diferencia.ToString();
						row.Cells["Diferencias"].Value = resultado;
					}
				}
			}
		}

		private void btnGenerar_Click(object sender, System.EventArgs e)
		{			
			if (DialogResult.Yes == MessageBox.Show("¿Esta seguro de generar el conteo?\n\nEste proceso no puede ser revertido", "Point Technology", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2)) 
			{
				string sSQL = string.Format("Exec GenerarConteoDiarioInventario {0}, {1}, '{2}', 1, ''", idConteo, (int)this.cmbBodega.Value, Convert.ToDateTime(this.dtFecha.Value).ToString("yyyyMMdd"));
				idConteo = Funcion.iEscalarSQL(cdsSeteoF, sSQL);	

				MessageBox.Show("Conteo de inventario generado correctamente", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);				

				this.Consulta();

				foreach (Infragistics.Win.UltraWinGrid.UltraGridRow row in uGridConteo.Rows)
				{
					if (row.Cells.Exists("Conteo") && row.Cells.Exists("Existencia") && row.Cells.Exists("Diferencias"))
					{
						if (row.Cells["Conteo"].Value != System.DBNull.Value &&
							row.Cells["Existencia"].Value != System.DBNull.Value)
						{
							int conteo = Convert.ToInt32(row.Cells["Conteo"].Value);
							int existencia = Convert.ToInt32(row.Cells["Existencia"].Value);
							int diferencia = conteo - existencia;

							string resultado = diferencia > 0 ? "+" + diferencia.ToString() : diferencia.ToString();
							row.Cells["Diferencias"].Value = resultado;
						}
					}
				}
			}
		}

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);
		}

		private void btnVer_Click(object sender, System.EventArgs e)
		{
			this.ListaConteo();
		}

		private void btnCerrar_Click(object sender, System.EventArgs e)
		{
			#region Cierre Conteo
			if (!Validacion.vbComboVacio(this.cmbVendedor, "Seleccione un vendedor")) return;

			if (DialogResult.Yes == MessageBox.Show("¿Esta seguro de cerrar el conteo?\n\nEste proceso no puede ser revertido", "Point Technology", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2)) 
			{							
				string sSQL = string.Format("Exec CierraConteoDiarioInventario {0}, '{1}', {2}", idConteo, this.txtNotas.Text.ToString().Trim(), (int)this.cmbVendedor.Value);
				Funcion.EjecutaSQL(cdsSeteoF, sSQL, true);									
							
				MessageBox.Show("Conteo de inventario cerrado correctamente", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);				

				this.Consulta();
			}
			#endregion Cierre Conteo			
		}

		private void Consulta()
		{
			if (this.cmbBodega.ActiveRow == null) return;

			idConteo = 0;
			iEstado = 0;
			this.cmbVendedor.Value = System.DBNull.Value;
			this.txtNotas.Text = "";

			this.cmbVendedor.Enabled = false;
			this.txtNotas.Enabled = false;			
			this.btnGenerar.Enabled = false;
			this.btnCerrar.Enabled = false;
			this.btnImprimir.Enabled = false;
			this.mnuImprimirInventario.Enabled = false;
			this.mnuImprimirInforme.Enabled = false;

			this.uGridConteo.DisplayLayout.Bands[0].Columns["Exibicion"].CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			this.uGridConteo.DisplayLayout.Bands[0].Columns["Bodega"].CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			this.uGridConteo.DisplayLayout.Bands[0].Columns["Notas"].CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			
			//this.cmbVendedor.DataSource = FuncionesProcedimientos.Lista_Vendedores(Convert.ToDateTime(this.dtFecha.Value), (int)this.cmbBodega.Value, cdsSeteoF);

			string sSQLRemates =string.Format("Exec CIValidaBodegaRemates {0}", (int)this.cmbBodega.Value);
			int iLocal = Funcion.iEscalarSQL(cdsSeteoF, sSQLRemates, true);

			this.cmbVendedor.DataSource = Funcion.dvProcedimiento(cdsSeteoF, string.Format("Exec ConsultaJefeLocal '{0}', {1}",  Convert.ToDateTime(this.dtFecha.Value).ToString("MM/yyyy"), iLocal));
			
			string sSQL = string.Format("Exec ConsultaConteoDiarioInventario {0}, '{1}'", (int)this.cmbBodega.Value, Convert.ToDateTime(this.dtFecha.Value).ToString("yyyyMMdd"));
			SqlDataReader dr = Funcion.rEscalarSQL(cdsSeteoF, sSQL, true);
			dr.Read();
			if (dr.HasRows)
			{
				idConteo = dr.GetInt32(0);
				iEstado = dr.GetInt32(1);
				this.cmbVendedor.Value = dr.GetInt32(2);
				this.lblEstado.Text = dr.GetString(3);
				this.txtNotas.Text = dr.GetString(4);
			}
			dr.Close();	

			this.ListaConteo();

			if (iEstado == 0) 
			{
				this.btnGenerar.Enabled = true;
				this.lblEstado.Text = "PENDIENTE";
			}
			if (iEstado == 1)	
			{
				if(Convert.ToDateTime(this.dtFecha.Value).ToString("yyyyMMdd") == Convert.ToDateTime(DateTime.Today).ToString("yyyyMMdd")  )
				{
					this.uGridConteo.DisplayLayout.Bands[0].Columns["Exibicion"].CellActivation = Infragistics.Win.UltraWinGrid.Activation.AllowEdit;
					this.uGridConteo.DisplayLayout.Bands[0].Columns["Bodega"].CellActivation = Infragistics.Win.UltraWinGrid.Activation.AllowEdit;
					this.uGridConteo.DisplayLayout.Bands[0].Columns["Notas"].CellActivation = Infragistics.Win.UltraWinGrid.Activation.AllowEdit;

					this.cmbVendedor.Enabled = true;
					this.txtNotas.Enabled = true;			
					this.btnCerrar.Enabled = true;
					this.btnImprimir.Enabled = true;
					this.mnuImprimirInventario.Enabled = true;
					this.btnExcel.Enabled = true;
				}
			}
			if (iEstado == 2)	
			{
				this.btnImprimir.Enabled = true;
				this.mnuImprimirInforme.Enabled = true;
			}
		}

		private void dtFecha_ValueChanged(object sender, System.EventArgs e)
		{
			this.Consulta();
		}

		private void btnSalir_Click(object sender, System.EventArgs e)
		{
			this.Close();
		}

		private void frmConteoDiarioInventario_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Escape) this.Close();
		}

		private void SumaConteo(Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			int exibicion = 0;
			int bodega = 0;

			if (e.Cell.Row.Cells["Exibicion"].Value != System.DBNull.Value)
				exibicion = Convert.ToInt32(e.Cell.Row.Cells["Exibicion"].Value);

			if (e.Cell.Row.Cells["Bodega"].Value != System.DBNull.Value)
				bodega = Convert.ToInt32(e.Cell.Row.Cells["Bodega"].Value);

			e.Cell.Row.Cells["Conteo"].Value = exibicion + bodega;
		}

		private void uGridConteo_AfterCellUpdate(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			if (e.Cell.Column.ToString() == "Exibicion")
			{
				if (e.Cell.Row.Cells["Exibicion"].Value != System.DBNull.Value)
				{
					string sSQL = string.Format("Exec ActualizaConteoInventarioDiario {0}, {1}, '', 1",
						(int)e.Cell.Row.Cells["idDetalleConteoDiarioInventario"].Value,
						(int)e.Cell.Row.Cells["Exibicion"].Value);
					Funcion.EjecutaSQL(cdsSeteoF, sSQL);

					SumaConteo(e);

					// Calcular diferencia visual
					if (e.Cell.Row.Cells["Conteo"].Value != System.DBNull.Value &&
						e.Cell.Row.Cells["Existencia"].Value != System.DBNull.Value &&
						e.Cell.Row.Cells.Exists("Diferencias"))
					{
						int conteo = Convert.ToInt32(e.Cell.Row.Cells["Conteo"].Value);
						int existencia = Convert.ToInt32(e.Cell.Row.Cells["Existencia"].Value);
						int diferencia = conteo - existencia;

						string resultado = diferencia > 0 ? "+" + diferencia.ToString() : diferencia.ToString();
						e.Cell.Row.Cells["Diferencias"].Value = resultado;
					}
				}
			}

			if (e.Cell.Column.ToString() == "Bodega")
			{
				if (e.Cell.Row.Cells["Bodega"].Value != System.DBNull.Value)
				{
					string sSQL = string.Format("Exec ActualizaConteoInventarioDiario {0}, {1}, '', 2",
						(int)e.Cell.Row.Cells["idDetalleConteoDiarioInventario"].Value,
						(int)e.Cell.Row.Cells["Bodega"].Value);
					Funcion.EjecutaSQL(cdsSeteoF, sSQL);

					SumaConteo(e);

					if (e.Cell.Row.Cells["Conteo"].Value != System.DBNull.Value &&
						e.Cell.Row.Cells["Existencia"].Value != System.DBNull.Value &&
						e.Cell.Row.Cells.Exists("Diferencias"))
					{
						int conteo = Convert.ToInt32(e.Cell.Row.Cells["Conteo"].Value);
						int existencia = Convert.ToInt32(e.Cell.Row.Cells["Existencia"].Value);
						int diferencia = conteo - existencia;

						string resultado = diferencia > 0 ? "+" + diferencia.ToString() : diferencia.ToString();
						e.Cell.Row.Cells["Diferencias"].Value = resultado;
					}
				}
			}

			if (e.Cell.Column.ToString() == "Conteo")
			{
				string sSQL = string.Format("Exec ActualizaConteoInventarioDiario {0}, {1}, '', 3",
					(int)e.Cell.Row.Cells["idDetalleConteoDiarioInventario"].Value,
					(int)e.Cell.Row.Cells["Conteo"].Value);
				Funcion.EjecutaSQL(cdsSeteoF, sSQL);

				if (e.Cell.Row.Cells["Conteo"].Value != System.DBNull.Value &&
					e.Cell.Row.Cells["Existencia"].Value != System.DBNull.Value &&
					e.Cell.Row.Cells.Exists("Diferencias"))
				{
					int conteo = Convert.ToInt32(e.Cell.Row.Cells["Conteo"].Value);
					int existencia = Convert.ToInt32(e.Cell.Row.Cells["Existencia"].Value);
					int diferencia = conteo - existencia;

					string resultado = diferencia > 0 ? "+" + diferencia.ToString() : diferencia.ToString();
					e.Cell.Row.Cells["Diferencias"].Value = resultado;
				}
			}

			if (e.Cell.Column.ToString() == "Notas")
			{
				string sSQL = string.Format("Exec ActualizaConteoInventarioDiario {0}, 0, '{1}', 5",
					(int)e.Cell.Row.Cells["idDetalleConteoDiarioInventario"].Value,
					e.Cell.Row.Cells["Notas"].Value.ToString());
				Funcion.EjecutaSQL(cdsSeteoF, sSQL);
			}
		}


		private void cmbBodega_ValueChanged(object sender, System.EventArgs e)
		{
			this.Consulta();
		}

		private void uGridConteo_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			FuncionesProcedimientos.DesplazamientoGrids(sender, e, this.uGridConteo);
		}

		private void cmbBodega_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}

		private void dtFecha_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Delete) e.Handled = true;
			if (e.KeyCode == Keys.Enter) 
				if (iBodegaPredef == 0) this.cmbBodega.Focus(); else this.cmbSubGrupo.Focus();
		}

		private void dtFecha_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (e.KeyChar == 8) e.Handled = true;
		}

		private void txtArticulo_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.btnVer_Click(sender, e);
		}

		private void txtNotas_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.cmbVendedor.Focus();
		}

		private void btnImprimir_Click(object sender, System.EventArgs e)
		{
			this.mnuImprimir.Show(this.btnImprimir, new Point(1, 25));			
		}

		private void cmbBodega_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.cmbSubGrupo.Focus();
		}

		private void cmbSubGrupo_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.txtArticulo.Focus();
		}

		private void uGridConteo_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}

		private void mnuImprimirInforme_Click(object sender, System.EventArgs e)
		{
			#region Imprime el reporte
			ParameterFields paramFields = new ParameterFields ();
				
			ParameterField pfRet = new ParameterField ();
			ParameterDiscreteValue id = new ParameterDiscreteValue ();
			pfRet.ParameterFieldName = "@idConteoDiarioInventario";
			id.Value = idConteo;
			pfRet.CurrentValues.Add (id);
			paramFields.Add (pfRet);
				
			Reporte miRepor = new Reporte("InformeCDI.rpt", "");
			miRepor.MdiParent = this.MdiParent;			
			miRepor.crVista.ParameterFieldInfo = paramFields;			
			miRepor.Show();
			#endregion Imprime el reporte
		}

		private void mnuImprimirInventario_Click(object sender, System.EventArgs e)
		{
			#region Imprime el reporte
			ParameterFields paramFields = new ParameterFields ();
				
			ParameterField pfRet = new ParameterField ();
			ParameterDiscreteValue id = new ParameterDiscreteValue ();
			pfRet.ParameterFieldName = "@idConteoDiarioInventario";
			id.Value = idConteo;
			pfRet.CurrentValues.Add (id);
			paramFields.Add (pfRet);
				
			Reporte miRepor = new Reporte("ListaInventarioCDI.rpt", "");
			miRepor.MdiParent = this.MdiParent;			
			miRepor.crVista.ParameterFieldInfo = paramFields;			
			miRepor.Show();
			#endregion Imprime el reporte
		}

		private void dtFecha_BeforeDropDown(object sender, System.ComponentModel.CancelEventArgs e)
		{
		
		}

		private void cmbSubGrupo_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}

		private void groupBox1_Enter(object sender, System.EventArgs e)
		{
		
		}

		private void btnExcel_Click(object sender, System.EventArgs e)
		{
			if (this.uGridConteo.Rows.Count == 0)
			{
				MessageBox.Show("No hay Filas para Exportar", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return;
			}
			FuncionesProcedimientos.ExportaExcel(this.uGridConteo);
		}
	}
}
