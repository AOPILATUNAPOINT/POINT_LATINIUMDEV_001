using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data.SqlClient;
using System.Data;
using CrystalDecisions.Shared;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de RepCompInv.
	/// </summary>
	public class RepCompInv : DevExpress.XtraEditors.XtraForm
	{
		private Infragistics.Win.Misc.UltraButton btInventario;
		private Infragistics.Win.UltraWinEditors.UltraDateTimeEditor cmbDesde;
		private Infragistics.Win.Misc.UltraLabel ultraLabel1;
		private Infragistics.Win.UltraWinEditors.UltraDateTimeEditor cmbHasta;
		private Infragistics.Win.UltraWinTabControl.UltraTabSharedControlsPage ultraTabSharedControlsPage1;
		private Infragistics.Win.UltraWinTabControl.UltraTabPageControl ultraTabPageControl1;
		private Infragistics.Win.UltraWinTabControl.UltraTabPageControl ultraTabPageControl2;
		private Infragistics.Win.UltraWinGrid.UltraGrid ultraGrid1;
		private Infragistics.Win.UltraWinGrid.UltraGrid ultraGrid2;
		private Infragistics.Win.Misc.UltraButton btKardex;
		private Infragistics.Win.UltraWinTabControl.UltraTabControl tcGrillas;
		private Infragistics.Win.UltraWinTabControl.UltraTabPageControl ultraTabPageControl4;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkSerie;
		private Infragistics.Win.Misc.UltraButton btFactura;
		private C1.Data.C1DataSet cdsGrupo;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbGrupoArt;
		private Infragistics.Win.Misc.UltraButton btFiltroGrupo;
		private Infragistics.Win.Misc.UltraButton btSubGrupo;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbSubGrupo;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtFiltro;
		private Infragistics.Win.UltraWinEditors.UltraComboEditor cmbAndOr;
		private Infragistics.Win.Misc.UltraButton btMarca;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbMarca;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor spnBodega;
		private Infragistics.Win.Misc.UltraLabel ultraLabel2;
		private System.Windows.Forms.ToolTip toolTip1;
		private Infragistics.Win.Misc.UltraButton btTipoGrupo;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbTipo;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbArtSri;
		private Infragistics.Win.Misc.UltraButton btArtSri;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbReporte;
		private C1.Data.C1DataSet cdsReporte;
		private C1.Data.C1DataView cdvReporte;
		private Infragistics.Win.Misc.UltraButton btReporte;
		private Infragistics.Win.Misc.UltraLabel ultraLabel3;
		private Infragistics.Win.UltraWinEditors.UltraOptionSet optBusca;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtSerie;
		private Infragistics.Win.UltraWinTabControl.UltraTabPageControl pgFiltro;
		private Infragistics.Win.Misc.UltraButton btBuscaKardex;
		private Infragistics.Win.Misc.UltraLabel ultraLabel4;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtArticulo;
		private Infragistics.Win.Misc.UltraButton btBuscaInventario;
		private Infragistics.Win.UltraWinGrid.UltraGrid ultraGrid3;
		private Infragistics.Win.Misc.UltraButton btBuscaSerie;
		private Infragistics.Win.Misc.UltraButton btInconcistencia;
		private Infragistics.Win.Misc.UltraButton btFacturaS;
		private Infragistics.Win.Misc.UltraButton btDisponible;
		private Infragistics.Win.Misc.UltraButton btVendido;
		private Infragistics.Win.Misc.UltraButton btSerie;
		private System.ComponentModel.IContainer components;

		public RepCompInv()
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
			this.components = new System.ComponentModel.Container();
			Infragistics.Win.Appearance appearance1 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Reporte", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Detalle");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idGrupo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idReporte");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Reporte");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Tipo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Filtro");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("NoFiltro");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("ArticuloSRI", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idArticuloSRI");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Detalle");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Porcentaje");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn11 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCuentaI");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn12 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCuentaE");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn13 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand3 = new Infragistics.Win.UltraWinGrid.UltraGridBand("ArticuloTipo", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn14 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idTipoGrupo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn15 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Tipo");
			Infragistics.Win.ValueListItem valueListItem1 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem2 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand4 = new Infragistics.Win.UltraWinGrid.UltraGridBand("ArticuloMarca", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn16 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idMarca");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn17 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Marca");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn18 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.Appearance appearance11 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand5 = new Infragistics.Win.UltraWinGrid.UltraGridBand("ArticuloSubGrupo", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn19 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idSubGrupo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn20 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("SubGrupo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn21 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.Appearance appearance12 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance13 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance14 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance15 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance16 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand6 = new Infragistics.Win.UltraWinGrid.UltraGridBand("ArticuloGrupo", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn22 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idGrupoArticulo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn23 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idTipoGrupo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn24 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Grupo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn25 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCuentas2");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn26 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCuentas3");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn27 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCuentas4");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn28 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Borrar");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn29 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCuentas5");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn30 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCuentas6");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn31 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idArticuloSRI");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn32 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn33 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCuentas7");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn34 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCuentas8");
			Infragistics.Win.Appearance appearance17 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance18 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance19 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance20 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance21 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance22 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance23 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance24 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance25 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance26 = new Infragistics.Win.Appearance();
			Infragistics.Win.ValueListItem valueListItem3 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem4 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem5 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.UltraWinTabControl.UltraTab ultraTab1 = new Infragistics.Win.UltraWinTabControl.UltraTab();
			Infragistics.Win.UltraWinTabControl.UltraTab ultraTab2 = new Infragistics.Win.UltraWinTabControl.UltraTab();
			Infragistics.Win.UltraWinTabControl.UltraTab ultraTab3 = new Infragistics.Win.UltraWinTabControl.UltraTab();
			Infragistics.Win.UltraWinTabControl.UltraTab ultraTab4 = new Infragistics.Win.UltraWinTabControl.UltraTab();
			this.ultraTabPageControl1 = new Infragistics.Win.UltraWinTabControl.UltraTabPageControl();
			this.ultraGrid1 = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraTabPageControl2 = new Infragistics.Win.UltraWinTabControl.UltraTabPageControl();
			this.ultraGrid2 = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraTabPageControl4 = new Infragistics.Win.UltraWinTabControl.UltraTabPageControl();
			this.btReporte = new Infragistics.Win.Misc.UltraButton();
			this.cmbReporte = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.cdvReporte = new C1.Data.C1DataView();
			this.cdsReporte = new C1.Data.C1DataSet();
			this.btArtSri = new Infragistics.Win.Misc.UltraButton();
			this.cmbArtSri = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.cdsGrupo = new C1.Data.C1DataSet();
			this.btTipoGrupo = new Infragistics.Win.Misc.UltraButton();
			this.cmbTipo = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.ultraLabel2 = new Infragistics.Win.Misc.UltraLabel();
			this.spnBodega = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.cmbAndOr = new Infragistics.Win.UltraWinEditors.UltraComboEditor();
			this.btMarca = new Infragistics.Win.Misc.UltraButton();
			this.cmbMarca = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.btSubGrupo = new Infragistics.Win.Misc.UltraButton();
			this.cmbSubGrupo = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.btFiltroGrupo = new Infragistics.Win.Misc.UltraButton();
			this.txtFiltro = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.cmbGrupoArt = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.chkSerie = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.pgFiltro = new Infragistics.Win.UltraWinTabControl.UltraTabPageControl();
			this.btSerie = new Infragistics.Win.Misc.UltraButton();
			this.btVendido = new Infragistics.Win.Misc.UltraButton();
			this.btDisponible = new Infragistics.Win.Misc.UltraButton();
			this.btFacturaS = new Infragistics.Win.Misc.UltraButton();
			this.btInconcistencia = new Infragistics.Win.Misc.UltraButton();
			this.btBuscaSerie = new Infragistics.Win.Misc.UltraButton();
			this.ultraGrid3 = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraLabel4 = new Infragistics.Win.Misc.UltraLabel();
			this.btBuscaInventario = new Infragistics.Win.Misc.UltraButton();
			this.txtArticulo = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.optBusca = new Infragistics.Win.UltraWinEditors.UltraOptionSet();
			this.ultraLabel3 = new Infragistics.Win.Misc.UltraLabel();
			this.btBuscaKardex = new Infragistics.Win.Misc.UltraButton();
			this.txtSerie = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.btInventario = new Infragistics.Win.Misc.UltraButton();
			this.cmbDesde = new Infragistics.Win.UltraWinEditors.UltraDateTimeEditor();
			this.ultraLabel1 = new Infragistics.Win.Misc.UltraLabel();
			this.cmbHasta = new Infragistics.Win.UltraWinEditors.UltraDateTimeEditor();
			this.tcGrillas = new Infragistics.Win.UltraWinTabControl.UltraTabControl();
			this.ultraTabSharedControlsPage1 = new Infragistics.Win.UltraWinTabControl.UltraTabSharedControlsPage();
			this.btKardex = new Infragistics.Win.Misc.UltraButton();
			this.btFactura = new Infragistics.Win.Misc.UltraButton();
			this.toolTip1 = new System.Windows.Forms.ToolTip(this.components);
			this.ultraTabPageControl1.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).BeginInit();
			this.ultraTabPageControl2.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid2)).BeginInit();
			this.ultraTabPageControl4.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.cmbReporte)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvReporte)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsReporte)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbArtSri)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsGrupo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbTipo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.spnBodega)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbAndOr)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbMarca)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbSubGrupo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtFiltro)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbGrupoArt)).BeginInit();
			this.pgFiltro.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid3)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtArticulo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.optBusca)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtSerie)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbDesde)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbHasta)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.tcGrillas)).BeginInit();
			this.tcGrillas.SuspendLayout();
			this.SuspendLayout();
			// 
			// ultraTabPageControl1
			// 
			this.ultraTabPageControl1.Controls.Add(this.ultraGrid1);
			this.ultraTabPageControl1.Location = new System.Drawing.Point(1, 23);
			this.ultraTabPageControl1.Name = "ultraTabPageControl1";
			this.ultraTabPageControl1.Size = new System.Drawing.Size(620, 275);
			// 
			// ultraGrid1
			// 
			this.ultraGrid1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.ultraGrid1.Cursor = System.Windows.Forms.Cursors.Default;
			appearance1.BackColor = System.Drawing.Color.White;
			this.ultraGrid1.DisplayLayout.Appearance = appearance1;
			appearance2.BackColor = System.Drawing.Color.Transparent;
			this.ultraGrid1.DisplayLayout.Override.CardAreaAppearance = appearance2;
			appearance3.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance3.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance3.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance3.FontData.BoldAsString = "True";
			appearance3.FontData.Name = "Arial";
			appearance3.FontData.SizeInPoints = 10F;
			appearance3.ForeColor = System.Drawing.Color.White;
			appearance3.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.ultraGrid1.DisplayLayout.Override.HeaderAppearance = appearance3;
			this.ultraGrid1.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti;
			appearance4.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance4.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance4.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.RowSelectorAppearance = appearance4;
			appearance5.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance5.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance5.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.SelectedRowAppearance = appearance5;
			this.ultraGrid1.DisplayLayout.ViewStyleBand = Infragistics.Win.UltraWinGrid.ViewStyleBand.OutlookGroupBy;
			this.ultraGrid1.Location = new System.Drawing.Point(16, 17);
			this.ultraGrid1.Name = "ultraGrid1";
			this.ultraGrid1.Size = new System.Drawing.Size(592, 243);
			this.ultraGrid1.TabIndex = 0;
			// 
			// ultraTabPageControl2
			// 
			this.ultraTabPageControl2.Controls.Add(this.ultraGrid2);
			this.ultraTabPageControl2.Location = new System.Drawing.Point(-10000, -10000);
			this.ultraTabPageControl2.Name = "ultraTabPageControl2";
			this.ultraTabPageControl2.Size = new System.Drawing.Size(620, 275);
			// 
			// ultraGrid2
			// 
			this.ultraGrid2.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.ultraGrid2.Cursor = System.Windows.Forms.Cursors.Default;
			appearance6.BackColor = System.Drawing.Color.White;
			this.ultraGrid2.DisplayLayout.Appearance = appearance6;
			appearance7.BackColor = System.Drawing.Color.Transparent;
			this.ultraGrid2.DisplayLayout.Override.CardAreaAppearance = appearance7;
			appearance8.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance8.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance8.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance8.FontData.BoldAsString = "True";
			appearance8.FontData.Name = "Arial";
			appearance8.FontData.SizeInPoints = 10F;
			appearance8.ForeColor = System.Drawing.Color.White;
			appearance8.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.ultraGrid2.DisplayLayout.Override.HeaderAppearance = appearance8;
			appearance9.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance9.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance9.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid2.DisplayLayout.Override.RowSelectorAppearance = appearance9;
			appearance10.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance10.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance10.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid2.DisplayLayout.Override.SelectedRowAppearance = appearance10;
			this.ultraGrid2.DisplayLayout.ViewStyleBand = Infragistics.Win.UltraWinGrid.ViewStyleBand.OutlookGroupBy;
			this.ultraGrid2.Location = new System.Drawing.Point(16, 17);
			this.ultraGrid2.Name = "ultraGrid2";
			this.ultraGrid2.Size = new System.Drawing.Size(591, 242);
			this.ultraGrid2.TabIndex = 1;
			// 
			// ultraTabPageControl4
			// 
			this.ultraTabPageControl4.Controls.Add(this.btReporte);
			this.ultraTabPageControl4.Controls.Add(this.cmbReporte);
			this.ultraTabPageControl4.Controls.Add(this.btArtSri);
			this.ultraTabPageControl4.Controls.Add(this.cmbArtSri);
			this.ultraTabPageControl4.Controls.Add(this.btTipoGrupo);
			this.ultraTabPageControl4.Controls.Add(this.cmbTipo);
			this.ultraTabPageControl4.Controls.Add(this.ultraLabel2);
			this.ultraTabPageControl4.Controls.Add(this.spnBodega);
			this.ultraTabPageControl4.Controls.Add(this.cmbAndOr);
			this.ultraTabPageControl4.Controls.Add(this.btMarca);
			this.ultraTabPageControl4.Controls.Add(this.cmbMarca);
			this.ultraTabPageControl4.Controls.Add(this.btSubGrupo);
			this.ultraTabPageControl4.Controls.Add(this.cmbSubGrupo);
			this.ultraTabPageControl4.Controls.Add(this.btFiltroGrupo);
			this.ultraTabPageControl4.Controls.Add(this.txtFiltro);
			this.ultraTabPageControl4.Controls.Add(this.cmbGrupoArt);
			this.ultraTabPageControl4.Controls.Add(this.chkSerie);
			this.ultraTabPageControl4.Location = new System.Drawing.Point(-10000, -10000);
			this.ultraTabPageControl4.Name = "ultraTabPageControl4";
			this.ultraTabPageControl4.Size = new System.Drawing.Size(620, 275);
			// 
			// btReporte
			// 
			this.btReporte.Location = new System.Drawing.Point(528, 26);
			this.btReporte.Name = "btReporte";
			this.btReporte.Size = new System.Drawing.Size(24, 26);
			this.btReporte.TabIndex = 45;
			this.btReporte.Text = "...";
			this.btReporte.Click += new System.EventHandler(this.btReporte_Click);
			// 
			// cmbReporte
			// 
			this.cmbReporte.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbReporte.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbReporte.DataSource = this.cdvReporte;
			this.cmbReporte.DisplayLayout.AutoFitColumns = true;
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn1.Width = 117;
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridColumn2.Hidden = true;
			ultraGridColumn3.Header.VisiblePosition = 2;
			ultraGridColumn3.Hidden = true;
			ultraGridColumn4.Header.VisiblePosition = 3;
			ultraGridColumn4.Hidden = true;
			ultraGridColumn5.Header.VisiblePosition = 4;
			ultraGridColumn5.Hidden = true;
			ultraGridColumn6.Header.VisiblePosition = 5;
			ultraGridColumn6.Hidden = true;
			ultraGridColumn6.Width = 54;
			ultraGridColumn7.Header.VisiblePosition = 6;
			ultraGridColumn7.Hidden = true;
			ultraGridColumn7.Width = 54;
			ultraGridBand1.Columns.AddRange(new object[] {
																										 ultraGridColumn1,
																										 ultraGridColumn2,
																										 ultraGridColumn3,
																										 ultraGridColumn4,
																										 ultraGridColumn5,
																										 ultraGridColumn6,
																										 ultraGridColumn7});
			this.cmbReporte.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.cmbReporte.DisplayMember = "Reporte";
			this.cmbReporte.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbReporte.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbReporte.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbReporte.Location = new System.Drawing.Point(384, 26);
			this.cmbReporte.Name = "cmbReporte";
			this.cmbReporte.Size = new System.Drawing.Size(136, 21);
			this.cmbReporte.TabIndex = 44;
			this.toolTip1.SetToolTip(this.cmbReporte, "Grupo 50 Tipo Avanzado");
			this.cmbReporte.ValueMember = "idReporte";
			// 
			// cdvReporte
			// 
			this.cdvReporte.BindingContextCtrl = this;
			this.cdvReporte.DataSet = this.cdsReporte;
			this.cdvReporte.RowFilter = "Tipo = 10 And idGrupo = 50";
			this.cdvReporte.TableName = "";
			this.cdvReporte.TableViewName = "Reporte";
			// 
			// cdsReporte
			// 
			this.cdsReporte.BindingContextCtrl = this;
			this.cdsReporte.DataLibrary = "LibEmpresa";
			this.cdsReporte.DataLibraryUrl = "";
			this.cdsReporte.DataSetDef = "dsReporte";
			this.cdsReporte.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsReporte.Name = "cdsReporte";
			this.cdsReporte.SchemaClassName = "LibEmpresa.DataClass";
			this.cdsReporte.SchemaDef = null;
			// 
			// btArtSri
			// 
			this.btArtSri.Location = new System.Drawing.Point(176, 198);
			this.btArtSri.Name = "btArtSri";
			this.btArtSri.Size = new System.Drawing.Size(32, 26);
			this.btArtSri.TabIndex = 15;
			this.btArtSri.Text = "->";
			this.btArtSri.Click += new System.EventHandler(this.btArtSri_Click);
			// 
			// cmbArtSri
			// 
			this.cmbArtSri.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbArtSri.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbArtSri.DataMember = "ArticuloSRI";
			this.cmbArtSri.DataSource = this.cdsGrupo;
			ultraGridBand2.AddButtonCaption = "ArticuloGrupo";
			ultraGridColumn8.Header.VisiblePosition = 2;
			ultraGridColumn8.Hidden = true;
			ultraGridColumn9.Header.VisiblePosition = 0;
			ultraGridColumn9.Width = 239;
			ultraGridColumn10.Header.VisiblePosition = 3;
			ultraGridColumn10.Hidden = true;
			ultraGridColumn11.Header.VisiblePosition = 4;
			ultraGridColumn11.Hidden = true;
			ultraGridColumn12.Header.VisiblePosition = 5;
			ultraGridColumn12.Hidden = true;
			ultraGridColumn13.Header.VisiblePosition = 1;
			ultraGridColumn13.Width = 76;
			ultraGridBand2.Columns.AddRange(new object[] {
																										 ultraGridColumn8,
																										 ultraGridColumn9,
																										 ultraGridColumn10,
																										 ultraGridColumn11,
																										 ultraGridColumn12,
																										 ultraGridColumn13});
			this.cmbArtSri.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			this.cmbArtSri.DisplayMember = "Detalle";
			this.cmbArtSri.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbArtSri.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbArtSri.Location = new System.Drawing.Point(40, 198);
			this.cmbArtSri.Name = "cmbArtSri";
			this.cmbArtSri.Size = new System.Drawing.Size(136, 21);
			this.cmbArtSri.TabIndex = 14;
			this.cmbArtSri.ValueMember = "idArticuloSRI";
			// 
			// cdsGrupo
			// 
			this.cdsGrupo.BindingContextCtrl = this;
			this.cdsGrupo.DataLibrary = "LibFacturacion";
			this.cdsGrupo.DataLibraryUrl = "";
			this.cdsGrupo.DataSetDef = "dsArticuloTabla";
			this.cdsGrupo.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsGrupo.Name = "cdsGrupo";
			this.cdsGrupo.SchemaClassName = "LibFacturacion.DataClass";
			this.cdsGrupo.SchemaDef = null;
			this.cdsGrupo.BeforeFill += new C1.Data.FillEventHandler(this.cdsGrupo_BeforeFill);
			// 
			// btTipoGrupo
			// 
			this.btTipoGrupo.Location = new System.Drawing.Point(176, 164);
			this.btTipoGrupo.Name = "btTipoGrupo";
			this.btTipoGrupo.Size = new System.Drawing.Size(32, 26);
			this.btTipoGrupo.TabIndex = 13;
			this.btTipoGrupo.Text = "->";
			this.btTipoGrupo.Click += new System.EventHandler(this.btTipoGrupo_Click);
			// 
			// cmbTipo
			// 
			this.cmbTipo.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbTipo.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbTipo.DataMember = "ArticuloTipo";
			this.cmbTipo.DataSource = this.cdsGrupo;
			ultraGridBand3.AddButtonCaption = "ArticuloGrupo";
			ultraGridColumn14.Header.VisiblePosition = 1;
			ultraGridColumn14.Hidden = true;
			ultraGridColumn15.Header.VisiblePosition = 0;
			ultraGridBand3.Columns.AddRange(new object[] {
																										 ultraGridColumn14,
																										 ultraGridColumn15});
			this.cmbTipo.DisplayLayout.BandsSerializer.Add(ultraGridBand3);
			this.cmbTipo.DisplayMember = "Tipo";
			this.cmbTipo.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbTipo.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbTipo.Location = new System.Drawing.Point(40, 164);
			this.cmbTipo.Name = "cmbTipo";
			this.cmbTipo.Size = new System.Drawing.Size(136, 21);
			this.cmbTipo.TabIndex = 12;
			this.cmbTipo.ValueMember = "idTipoGrupo";
			// 
			// ultraLabel2
			// 
			this.ultraLabel2.AutoSize = true;
			this.ultraLabel2.BackColor = System.Drawing.Color.Transparent;
			this.ultraLabel2.Location = new System.Drawing.Point(232, 26);
			this.ultraLabel2.Name = "ultraLabel2";
			this.ultraLabel2.Size = new System.Drawing.Size(41, 15);
			this.ultraLabel2.TabIndex = 11;
			this.ultraLabel2.Text = "Bodega";
			// 
			// spnBodega
			// 
			this.spnBodega.Location = new System.Drawing.Point(288, 26);
			this.spnBodega.MaskInput = "nnn";
			this.spnBodega.MinValue = 0;
			this.spnBodega.Name = "spnBodega";
			this.spnBodega.PromptChar = ' ';
			this.spnBodega.Size = new System.Drawing.Size(72, 22);
			this.spnBodega.SpinButtonDisplayStyle = Infragistics.Win.ButtonDisplayStyle.Always;
			this.spnBodega.TabIndex = 10;
			this.toolTip1.SetToolTip(this.spnBodega, "0 para Todas las bodegas");
			// 
			// cmbAndOr
			// 
			this.cmbAndOr.DropDownStyle = Infragistics.Win.DropDownStyle.DropDownList;
			valueListItem1.DataValue = "And";
			valueListItem1.DisplayText = "And";
			valueListItem2.DataValue = "Or";
			valueListItem2.DisplayText = "Or";
			this.cmbAndOr.Items.Add(valueListItem1);
			this.cmbAndOr.Items.Add(valueListItem2);
			this.cmbAndOr.Location = new System.Drawing.Point(152, 26);
			this.cmbAndOr.Name = "cmbAndOr";
			this.cmbAndOr.Size = new System.Drawing.Size(48, 22);
			this.cmbAndOr.TabIndex = 9;
			this.cmbAndOr.Text = "And";
			// 
			// btMarca
			// 
			this.btMarca.Location = new System.Drawing.Point(176, 129);
			this.btMarca.Name = "btMarca";
			this.btMarca.Size = new System.Drawing.Size(32, 26);
			this.btMarca.TabIndex = 8;
			this.btMarca.Text = "->";
			this.btMarca.Click += new System.EventHandler(this.btMarca_Click);
			// 
			// cmbMarca
			// 
			this.cmbMarca.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbMarca.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbMarca.DataMember = "ArticuloMarca";
			this.cmbMarca.DataSource = this.cdsGrupo;
			ultraGridBand4.AddButtonCaption = "ArticuloGrupo";
			ultraGridColumn16.Header.VisiblePosition = 1;
			ultraGridColumn16.Hidden = true;
			ultraGridColumn17.Header.VisiblePosition = 0;
			ultraGridColumn18.Header.VisiblePosition = 2;
			ultraGridColumn18.Hidden = true;
			ultraGridBand4.Columns.AddRange(new object[] {
																										 ultraGridColumn16,
																										 ultraGridColumn17,
																										 ultraGridColumn18});
			this.cmbMarca.DisplayLayout.BandsSerializer.Add(ultraGridBand4);
			this.cmbMarca.DisplayMember = "Marca";
			this.cmbMarca.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbMarca.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbMarca.Location = new System.Drawing.Point(40, 129);
			this.cmbMarca.Name = "cmbMarca";
			this.cmbMarca.Size = new System.Drawing.Size(136, 21);
			this.cmbMarca.TabIndex = 7;
			this.cmbMarca.ValueMember = "idMarca";
			// 
			// btSubGrupo
			// 
			this.btSubGrupo.Location = new System.Drawing.Point(176, 95);
			this.btSubGrupo.Name = "btSubGrupo";
			this.btSubGrupo.Size = new System.Drawing.Size(32, 26);
			this.btSubGrupo.TabIndex = 6;
			this.btSubGrupo.Text = "->";
			this.btSubGrupo.Click += new System.EventHandler(this.btSubGrupo_Click);
			// 
			// cmbSubGrupo
			// 
			this.cmbSubGrupo.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbSubGrupo.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbSubGrupo.DataMember = "ArticuloSubGrupo";
			this.cmbSubGrupo.DataSource = this.cdsGrupo;
			appearance11.BackColor = System.Drawing.Color.White;
			this.cmbSubGrupo.DisplayLayout.Appearance = appearance11;
			ultraGridBand5.AddButtonCaption = "ArticuloGrupo";
			ultraGridColumn19.Header.VisiblePosition = 2;
			ultraGridColumn19.Hidden = true;
			ultraGridColumn20.Header.VisiblePosition = 0;
			ultraGridColumn21.Header.VisiblePosition = 1;
			ultraGridBand5.Columns.AddRange(new object[] {
																										 ultraGridColumn19,
																										 ultraGridColumn20,
																										 ultraGridColumn21});
			this.cmbSubGrupo.DisplayLayout.BandsSerializer.Add(ultraGridBand5);
			appearance12.BackColor = System.Drawing.Color.Transparent;
			this.cmbSubGrupo.DisplayLayout.Override.CardAreaAppearance = appearance12;
			appearance13.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance13.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance13.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance13.FontData.BoldAsString = "True";
			appearance13.FontData.Name = "Arial";
			appearance13.FontData.SizeInPoints = 10F;
			appearance13.ForeColor = System.Drawing.Color.White;
			appearance13.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.cmbSubGrupo.DisplayLayout.Override.HeaderAppearance = appearance13;
			appearance14.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance14.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance14.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.cmbSubGrupo.DisplayLayout.Override.RowSelectorAppearance = appearance14;
			appearance15.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance15.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance15.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.cmbSubGrupo.DisplayLayout.Override.SelectedRowAppearance = appearance15;
			this.cmbSubGrupo.DisplayMember = "SubGrupo";
			this.cmbSubGrupo.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbSubGrupo.Location = new System.Drawing.Point(40, 95);
			this.cmbSubGrupo.Name = "cmbSubGrupo";
			this.cmbSubGrupo.Size = new System.Drawing.Size(136, 22);
			this.cmbSubGrupo.TabIndex = 5;
			this.cmbSubGrupo.ValueMember = "idSubGrupo";
			// 
			// btFiltroGrupo
			// 
			this.btFiltroGrupo.Location = new System.Drawing.Point(176, 60);
			this.btFiltroGrupo.Name = "btFiltroGrupo";
			this.btFiltroGrupo.Size = new System.Drawing.Size(32, 26);
			this.btFiltroGrupo.TabIndex = 4;
			this.btFiltroGrupo.Text = "->";
			this.btFiltroGrupo.Click += new System.EventHandler(this.btFiltroGrupo_Click);
			// 
			// txtFiltro
			// 
			this.txtFiltro.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.txtFiltro.Location = new System.Drawing.Point(232, 60);
			this.txtFiltro.Multiline = true;
			this.txtFiltro.Name = "txtFiltro";
			this.txtFiltro.Size = new System.Drawing.Size(367, 199);
			this.txtFiltro.TabIndex = 3;
			// 
			// cmbGrupoArt
			// 
			this.cmbGrupoArt.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbGrupoArt.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbGrupoArt.DataMember = "ArticuloGrupo";
			this.cmbGrupoArt.DataSource = this.cdsGrupo;
			appearance16.BackColor = System.Drawing.Color.White;
			this.cmbGrupoArt.DisplayLayout.Appearance = appearance16;
			ultraGridColumn22.Header.VisiblePosition = 2;
			ultraGridColumn22.Hidden = true;
			ultraGridColumn23.Header.VisiblePosition = 3;
			ultraGridColumn23.Hidden = true;
			ultraGridColumn24.Header.VisiblePosition = 0;
			ultraGridColumn25.Header.VisiblePosition = 4;
			ultraGridColumn25.Hidden = true;
			ultraGridColumn26.Header.VisiblePosition = 5;
			ultraGridColumn26.Hidden = true;
			ultraGridColumn27.Header.VisiblePosition = 6;
			ultraGridColumn27.Hidden = true;
			ultraGridColumn28.Header.VisiblePosition = 7;
			ultraGridColumn28.Hidden = true;
			ultraGridColumn29.Header.VisiblePosition = 8;
			ultraGridColumn29.Hidden = true;
			ultraGridColumn30.Header.VisiblePosition = 9;
			ultraGridColumn30.Hidden = true;
			ultraGridColumn31.Header.VisiblePosition = 10;
			ultraGridColumn31.Hidden = true;
			ultraGridColumn32.Header.VisiblePosition = 1;
			ultraGridColumn33.Header.VisiblePosition = 11;
			ultraGridColumn33.Hidden = true;
			ultraGridColumn34.Header.VisiblePosition = 12;
			ultraGridBand6.Columns.AddRange(new object[] {
																										 ultraGridColumn22,
																										 ultraGridColumn23,
																										 ultraGridColumn24,
																										 ultraGridColumn25,
																										 ultraGridColumn26,
																										 ultraGridColumn27,
																										 ultraGridColumn28,
																										 ultraGridColumn29,
																										 ultraGridColumn30,
																										 ultraGridColumn31,
																										 ultraGridColumn32,
																										 ultraGridColumn33,
																										 ultraGridColumn34});
			this.cmbGrupoArt.DisplayLayout.BandsSerializer.Add(ultraGridBand6);
			appearance17.BackColor = System.Drawing.Color.Transparent;
			this.cmbGrupoArt.DisplayLayout.Override.CardAreaAppearance = appearance17;
			appearance18.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance18.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance18.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance18.FontData.BoldAsString = "True";
			appearance18.FontData.Name = "Arial";
			appearance18.FontData.SizeInPoints = 10F;
			appearance18.ForeColor = System.Drawing.Color.White;
			appearance18.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.cmbGrupoArt.DisplayLayout.Override.HeaderAppearance = appearance18;
			appearance19.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance19.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance19.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.cmbGrupoArt.DisplayLayout.Override.RowSelectorAppearance = appearance19;
			appearance20.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance20.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance20.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.cmbGrupoArt.DisplayLayout.Override.SelectedRowAppearance = appearance20;
			this.cmbGrupoArt.DisplayMember = "Grupo";
			this.cmbGrupoArt.Location = new System.Drawing.Point(40, 60);
			this.cmbGrupoArt.Name = "cmbGrupoArt";
			this.cmbGrupoArt.Size = new System.Drawing.Size(136, 22);
			this.cmbGrupoArt.TabIndex = 1;
			this.cmbGrupoArt.ValueMember = "idGrupoArticulo";
			// 
			// chkSerie
			// 
			this.chkSerie.BackColor = System.Drawing.Color.Transparent;
			this.chkSerie.Location = new System.Drawing.Point(40, 26);
			this.chkSerie.Name = "chkSerie";
			this.chkSerie.Size = new System.Drawing.Size(56, 17);
			this.chkSerie.TabIndex = 0;
			this.chkSerie.Text = "Serie";
			// 
			// pgFiltro
			// 
			this.pgFiltro.Controls.Add(this.btSerie);
			this.pgFiltro.Controls.Add(this.btVendido);
			this.pgFiltro.Controls.Add(this.btDisponible);
			this.pgFiltro.Controls.Add(this.btFacturaS);
			this.pgFiltro.Controls.Add(this.btInconcistencia);
			this.pgFiltro.Controls.Add(this.btBuscaSerie);
			this.pgFiltro.Controls.Add(this.ultraGrid3);
			this.pgFiltro.Controls.Add(this.ultraLabel4);
			this.pgFiltro.Controls.Add(this.btBuscaInventario);
			this.pgFiltro.Controls.Add(this.txtArticulo);
			this.pgFiltro.Controls.Add(this.optBusca);
			this.pgFiltro.Controls.Add(this.ultraLabel3);
			this.pgFiltro.Controls.Add(this.btBuscaKardex);
			this.pgFiltro.Controls.Add(this.txtSerie);
			this.pgFiltro.Location = new System.Drawing.Point(-10000, -10000);
			this.pgFiltro.Name = "pgFiltro";
			this.pgFiltro.Size = new System.Drawing.Size(620, 275);
			// 
			// btSerie
			// 
			this.btSerie.Location = new System.Drawing.Point(592, 52);
			this.btSerie.Name = "btSerie";
			this.btSerie.Size = new System.Drawing.Size(24, 26);
			this.btSerie.TabIndex = 13;
			this.btSerie.Text = "S";
			this.toolTip1.SetToolTip(this.btSerie, "Facturas en series");
			this.btSerie.Click += new System.EventHandler(this.btSerie_Click);
			// 
			// btVendido
			// 
			this.btVendido.Location = new System.Drawing.Point(560, 52);
			this.btVendido.Name = "btVendido";
			this.btVendido.Size = new System.Drawing.Size(24, 26);
			this.btVendido.TabIndex = 12;
			this.btVendido.Text = "V";
			this.toolTip1.SetToolTip(this.btVendido, "Serie NO Disponible");
			this.btVendido.Click += new System.EventHandler(this.btVendido_Click);
			// 
			// btDisponible
			// 
			this.btDisponible.Location = new System.Drawing.Point(560, 17);
			this.btDisponible.Name = "btDisponible";
			this.btDisponible.Size = new System.Drawing.Size(24, 26);
			this.btDisponible.TabIndex = 11;
			this.btDisponible.Text = "D";
			this.toolTip1.SetToolTip(this.btDisponible, "Serie Disponible");
			this.btDisponible.Click += new System.EventHandler(this.btDisponible_Click);
			// 
			// btFacturaS
			// 
			this.btFacturaS.Location = new System.Drawing.Point(592, 17);
			this.btFacturaS.Name = "btFacturaS";
			this.btFacturaS.Size = new System.Drawing.Size(24, 26);
			this.btFacturaS.TabIndex = 10;
			this.btFacturaS.Text = "F";
			this.toolTip1.SetToolTip(this.btFacturaS, "Facturas en series");
			this.btFacturaS.Click += new System.EventHandler(this.btFacturaS_Click);
			// 
			// btInconcistencia
			// 
			this.btInconcistencia.Location = new System.Drawing.Point(224, 52);
			this.btInconcistencia.Name = "btInconcistencia";
			this.btInconcistencia.Size = new System.Drawing.Size(24, 26);
			this.btInconcistencia.TabIndex = 9;
			this.btInconcistencia.Text = "I";
			this.toolTip1.SetToolTip(this.btInconcistencia, "Inconcistencias en series");
			this.btInconcistencia.Click += new System.EventHandler(this.btInconcistencia_Click);
			// 
			// btBuscaSerie
			// 
			this.btBuscaSerie.Location = new System.Drawing.Point(224, 17);
			this.btBuscaSerie.Name = "btBuscaSerie";
			this.btBuscaSerie.Size = new System.Drawing.Size(24, 26);
			this.btBuscaSerie.TabIndex = 8;
			this.btBuscaSerie.Text = ">";
			this.btBuscaSerie.Click += new System.EventHandler(this.btBuscaSerie_Click);
			// 
			// ultraGrid3
			// 
			this.ultraGrid3.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.ultraGrid3.Cursor = System.Windows.Forms.Cursors.Default;
			appearance21.BackColor = System.Drawing.Color.White;
			this.ultraGrid3.DisplayLayout.Appearance = appearance21;
			this.ultraGrid3.DisplayLayout.AutoFitColumns = true;
			appearance22.BackColor = System.Drawing.Color.Transparent;
			this.ultraGrid3.DisplayLayout.Override.CardAreaAppearance = appearance22;
			appearance23.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance23.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance23.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance23.FontData.BoldAsString = "True";
			appearance23.FontData.Name = "Arial";
			appearance23.FontData.SizeInPoints = 10F;
			appearance23.ForeColor = System.Drawing.Color.White;
			appearance23.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.ultraGrid3.DisplayLayout.Override.HeaderAppearance = appearance23;
			this.ultraGrid3.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti;
			appearance24.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance24.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance24.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid3.DisplayLayout.Override.RowSelectorAppearance = appearance24;
			appearance25.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance25.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance25.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid3.DisplayLayout.Override.SelectedRowAppearance = appearance25;
			this.ultraGrid3.DisplayLayout.ViewStyleBand = Infragistics.Win.UltraWinGrid.ViewStyleBand.OutlookGroupBy;
			this.ultraGrid3.Location = new System.Drawing.Point(16, 95);
			this.ultraGrid3.Name = "ultraGrid3";
			this.ultraGrid3.Size = new System.Drawing.Size(591, 173);
			this.ultraGrid3.TabIndex = 7;
			// 
			// ultraLabel4
			// 
			this.ultraLabel4.BackColor = System.Drawing.Color.Transparent;
			this.ultraLabel4.Location = new System.Drawing.Point(24, 52);
			this.ultraLabel4.Name = "ultraLabel4";
			this.ultraLabel4.Size = new System.Drawing.Size(48, 17);
			this.ultraLabel4.TabIndex = 6;
			this.ultraLabel4.Text = "Artículo";
			// 
			// btBuscaInventario
			// 
			this.btBuscaInventario.Location = new System.Drawing.Point(264, 52);
			this.btBuscaInventario.Name = "btBuscaInventario";
			this.btBuscaInventario.Size = new System.Drawing.Size(144, 26);
			this.btBuscaInventario.TabIndex = 5;
			this.btBuscaInventario.Text = "Búsqueda en Inventario";
			this.btBuscaInventario.Click += new System.EventHandler(this.btBuscaInventario_Click);
			// 
			// txtArticulo
			// 
			this.txtArticulo.Location = new System.Drawing.Point(80, 52);
			this.txtArticulo.Name = "txtArticulo";
			this.txtArticulo.Size = new System.Drawing.Size(128, 22);
			this.txtArticulo.TabIndex = 4;
			// 
			// optBusca
			// 
			this.optBusca.CheckedIndex = 2;
			this.optBusca.ItemAppearance = appearance26;
			valueListItem3.DataValue = 0;
			valueListItem3.DisplayText = "Empieza con";
			valueListItem4.DataValue = 1;
			valueListItem4.DisplayText = "Exacta";
			valueListItem5.DataValue = 2;
			valueListItem5.DisplayText = "Cualquier parte";
			this.optBusca.Items.Add(valueListItem3);
			this.optBusca.Items.Add(valueListItem4);
			this.optBusca.Items.Add(valueListItem5);
			this.optBusca.ItemSpacingHorizontal = 6;
			this.optBusca.ItemSpacingVertical = 5;
			this.optBusca.Location = new System.Drawing.Point(432, 17);
			this.optBusca.Name = "optBusca";
			this.optBusca.Size = new System.Drawing.Size(112, 69);
			this.optBusca.TabIndex = 3;
			this.optBusca.Text = "Cualquier parte";
			// 
			// ultraLabel3
			// 
			this.ultraLabel3.BackColor = System.Drawing.Color.Transparent;
			this.ultraLabel3.Location = new System.Drawing.Point(24, 17);
			this.ultraLabel3.Name = "ultraLabel3";
			this.ultraLabel3.Size = new System.Drawing.Size(40, 17);
			this.ultraLabel3.TabIndex = 2;
			this.ultraLabel3.Text = "Serie";
			// 
			// btBuscaKardex
			// 
			this.btBuscaKardex.Location = new System.Drawing.Point(264, 17);
			this.btBuscaKardex.Name = "btBuscaKardex";
			this.btBuscaKardex.Size = new System.Drawing.Size(144, 26);
			this.btBuscaKardex.TabIndex = 1;
			this.btBuscaKardex.Text = "Búsqueda en Kardex";
			this.btBuscaKardex.Click += new System.EventHandler(this.btBuscaKardex_Click);
			// 
			// txtSerie
			// 
			this.txtSerie.Location = new System.Drawing.Point(80, 17);
			this.txtSerie.Name = "txtSerie";
			this.txtSerie.Size = new System.Drawing.Size(128, 22);
			this.txtSerie.TabIndex = 0;
			// 
			// btInventario
			// 
			this.btInventario.Location = new System.Drawing.Point(352, 17);
			this.btInventario.Name = "btInventario";
			this.btInventario.Size = new System.Drawing.Size(64, 26);
			this.btInventario.TabIndex = 0;
			this.btInventario.Text = "Inventario";
			this.btInventario.Click += new System.EventHandler(this.btInventario_Click);
			// 
			// cmbDesde
			// 
			this.cmbDesde.DateTime = new System.DateTime(2009, 3, 1, 0, 0, 0, 0);
			this.cmbDesde.FormatString = "dd/MMM/yyyy";
			this.cmbDesde.Location = new System.Drawing.Point(80, 17);
			this.cmbDesde.Name = "cmbDesde";
			this.cmbDesde.Size = new System.Drawing.Size(112, 22);
			this.cmbDesde.SpinButtonDisplayStyle = Infragistics.Win.ButtonDisplayStyle.Always;
			this.cmbDesde.TabIndex = 2;
			this.cmbDesde.Value = new System.DateTime(2009, 3, 1, 0, 0, 0, 0);
			// 
			// ultraLabel1
			// 
			this.ultraLabel1.AutoSize = true;
			this.ultraLabel1.Location = new System.Drawing.Point(24, 17);
			this.ultraLabel1.Name = "ultraLabel1";
			this.ultraLabel1.Size = new System.Drawing.Size(42, 15);
			this.ultraLabel1.TabIndex = 3;
			this.ultraLabel1.Text = "Período";
			// 
			// cmbHasta
			// 
			this.cmbHasta.DateTime = new System.DateTime(2009, 3, 1, 0, 0, 0, 0);
			this.cmbHasta.FormatString = "dd/MMM/yyyy";
			this.cmbHasta.Location = new System.Drawing.Point(200, 17);
			this.cmbHasta.Name = "cmbHasta";
			this.cmbHasta.Size = new System.Drawing.Size(112, 22);
			this.cmbHasta.SpinButtonDisplayStyle = Infragistics.Win.ButtonDisplayStyle.Always;
			this.cmbHasta.TabIndex = 4;
			this.cmbHasta.Value = new System.DateTime(2009, 3, 1, 0, 0, 0, 0);
			// 
			// tcGrillas
			// 
			this.tcGrillas.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.tcGrillas.Controls.Add(this.ultraTabSharedControlsPage1);
			this.tcGrillas.Controls.Add(this.ultraTabPageControl1);
			this.tcGrillas.Controls.Add(this.ultraTabPageControl2);
			this.tcGrillas.Controls.Add(this.ultraTabPageControl4);
			this.tcGrillas.Controls.Add(this.pgFiltro);
			this.tcGrillas.Location = new System.Drawing.Point(24, 52);
			this.tcGrillas.Name = "tcGrillas";
			this.tcGrillas.SharedControlsPage = this.ultraTabSharedControlsPage1;
			this.tcGrillas.Size = new System.Drawing.Size(624, 301);
			this.tcGrillas.TabIndex = 5;
			ultraTab1.TabPage = this.ultraTabPageControl1;
			ultraTab1.Text = "Principal";
			ultraTab2.TabPage = this.ultraTabPageControl2;
			ultraTab2.Text = "Detalle";
			ultraTab3.TabPage = this.ultraTabPageControl4;
			ultraTab3.Text = "Opciones";
			ultraTab4.TabPage = this.pgFiltro;
			ultraTab4.Text = "Filtros";
			this.tcGrillas.Tabs.AddRange(new Infragistics.Win.UltraWinTabControl.UltraTab[] {
																																												ultraTab1,
																																												ultraTab2,
																																												ultraTab3,
																																												ultraTab4});
			// 
			// ultraTabSharedControlsPage1
			// 
			this.ultraTabSharedControlsPage1.Location = new System.Drawing.Point(-10000, -10000);
			this.ultraTabSharedControlsPage1.Name = "ultraTabSharedControlsPage1";
			this.ultraTabSharedControlsPage1.Size = new System.Drawing.Size(620, 275);
			// 
			// btKardex
			// 
			this.btKardex.Location = new System.Drawing.Point(424, 17);
			this.btKardex.Name = "btKardex";
			this.btKardex.Size = new System.Drawing.Size(64, 26);
			this.btKardex.TabIndex = 6;
			this.btKardex.Text = "Kardex";
			this.btKardex.Click += new System.EventHandler(this.btKardex_Click);
			// 
			// btFactura
			// 
			this.btFactura.Location = new System.Drawing.Point(496, 17);
			this.btFactura.Name = "btFactura";
			this.btFactura.Size = new System.Drawing.Size(64, 26);
			this.btFactura.TabIndex = 7;
			this.btFactura.Text = "Factura";
			this.btFactura.Click += new System.EventHandler(this.btFactura_Click);
			// 
			// RepCompInv
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 14);
			this.ClientSize = new System.Drawing.Size(664, 368);
			this.Controls.Add(this.btFactura);
			this.Controls.Add(this.btKardex);
			this.Controls.Add(this.tcGrillas);
			this.Controls.Add(this.cmbHasta);
			this.Controls.Add(this.ultraLabel1);
			this.Controls.Add(this.cmbDesde);
			this.Controls.Add(this.btInventario);
			this.Name = "RepCompInv";
			this.Text = "Comparación  Inventarios";
			this.Load += new System.EventHandler(this.RepCompInv_Load);
			this.ultraTabPageControl1.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).EndInit();
			this.ultraTabPageControl2.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid2)).EndInit();
			this.ultraTabPageControl4.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.cmbReporte)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvReporte)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsReporte)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbArtSri)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsGrupo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbTipo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.spnBodega)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbAndOr)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbMarca)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbSubGrupo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtFiltro)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbGrupoArt)).EndInit();
			this.pgFiltro.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid3)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtArticulo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.optBusca)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtSerie)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbDesde)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbHasta)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.tcGrillas)).EndInit();
			this.tcGrillas.ResumeLayout(false);
			this.ResumeLayout(false);

		}
		#endregion
		DataSet miDataSet = new DataSet();
		DataSet dsDetalle = new DataSet();
		DataView dv = new DataView();
		DataView dvK = new DataView();

		private void btInventario_Click(object sender, System.EventArgs e)
		{
			DateTime dtDesde = (DateTime) cmbDesde.Value;
			DateTime dtHasta = (DateTime) cmbHasta.Value;
			Funcion miFuncion = new Funcion();
			string strConn = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);	
			string stStore = "InventarioGenera";
			if (chkSerie.Checked) stStore = "InventarioSerie";
			string stProcedimiento = "EXEC " + stStore + " '" + 
				dtDesde.ToString("yyyyMMdd") + "', '" + dtHasta.ToString("yyyyMMdd") + "', ";
			stProcedimiento += "'" + txtFiltro.Text.ToString().Trim() + "', "
				+ spnBodega.Value.ToString();
			SqlDataAdapter myData = new SqlDataAdapter(stProcedimiento, strConn);

			myData.TableMappings.Add("Tabla", "ExcelSube");
			
			miDataSet.Reset();
			miDataSet.Clear();
			try
			{
				myData.Fill(miDataSet);
			}
			catch(Exception ex)
			{
				MessageBox.Show(ex.Message, "Error al abrir la Hoja A del Archivo: " );
				return;
			}
			tcGrillas.Tabs[0].Selected = true;

			dv = miDataSet.Tables[0].DefaultView;
			ultraGrid1.DataSource = dv;
			ultraGrid1.DisplayLayout.Bands[0].Columns["idArticulo"].Hidden = true;
			ultraGrid1.DisplayLayout.Bands[0].Columns["Saldo"].Format = "#,##0.00";
			ultraGrid1.DataBind();
		}

		private void RepCompInv_Load(object sender, System.EventArgs e)
		{
			cmbDesde.Value = DateTime.Today;
			cmbHasta.Value = DateTime.Today;
		}

		private void btKardex_Click(object sender, System.EventArgs e)
		{
			if (ultraGrid1.ActiveRow == null)
			{
				MessageBox.Show("Escoja Articulo" ,"Información");
				return;
			}
			DateTime dtDesde = (DateTime) cmbDesde.Value;
			DateTime dtHasta = (DateTime) cmbHasta.Value;
			Funcion miFuncion = new Funcion();
			string strConn = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);	
			string stStore = "KardexGenera";
			if (chkSerie.Checked) stStore = "KardexSerie";
			string stArticulo = "', " + ultraGrid1.ActiveRow.Cells["idArticulo"].Value.ToString();
			string stProcedimiento = "EXEC " + stStore + " '" + 
				dtDesde.ToString("yyyyMMdd") + "', '" + dtHasta.ToString("yyyyMMdd") + stArticulo;
			SqlDataAdapter myData = new SqlDataAdapter(stProcedimiento, strConn);

			myData.TableMappings.Add("Tabla", "ExcelSube");
			dsDetalle.Reset();
			dsDetalle.Clear();
			try
			{
				myData.Fill(dsDetalle);
			}
			catch(Exception ex)
			{
				MessageBox.Show(ex.Message, "Error al abrir la Hoja A del Archivo: " );
				return;
			}
			tcGrillas.Tabs[1].Selected = true;
			dvK = dsDetalle.Tables[0].DefaultView;
			ultraGrid2.DataSource = dvK;
			ultraGrid2.DisplayLayout.Bands[0].Columns["idArticulo"].Hidden = true;
			ultraGrid2.DisplayLayout.Bands[0].Columns["idDetCompra"].Hidden = true;
			ultraGrid2.DisplayLayout.Bands[0].Columns["idTipoFactura"].Hidden = true;
			ultraGrid2.DisplayLayout.Bands[0].Columns["Fecha"].Format = "dd/MMM/yyyy";
			ultraGrid2.DisplayLayout.Bands[0].Columns["Precio"].Format = "#,##0.00";
			ultraGrid2.DataBind();	
		}

		private void btFactura_Click(object sender, System.EventArgs e)
		{
			if (ultraGrid2.ActiveRow == null)
			{
				MessageBox.Show("Escoja Linea de kardex" ,"Información");
				return;
			}
			int IdDetCompra = int.Parse(ultraGrid2.ActiveRow.Cells["idDetCompra"].Text);
			if (IdDetCompra == 0)
			{
				MessageBox.Show("Linea de kardex no tiene una factura" ,"Información");
				return;
			}
			string stSelect = "Select idCompra From DetCompra where idDetCompra = "
				+ IdDetCompra.ToString();
			int IdCompra = Funcion.iEscalarSQL(cdsGrupo, stSelect, true);
			stSelect = "Select idTipoFactura From Compra where idCompra = "
				+ IdCompra.ToString();
			int IdTipoFactura = Funcion.iEscalarSQL(cdsGrupo, stSelect, true);
//			Compra miCompra = new Compra(IdTipoFactura, IdCompra);
//			miCompra.MdiParent = this.MdiParent;
//			miCompra.Show();
		}

		private void cdsGrupo_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			this.cdsGrupo.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);		
		}

		private void btFiltroGrupo_Click(object sender, System.EventArgs e)
		{
			txtFiltro.Text += "  " + cmbAndOr.Text + " a.idGrupoArticulo = " + cmbGrupoArt.Value.ToString();
		}

		private void btSubGrupo_Click(object sender, System.EventArgs e)
		{
			txtFiltro.Text += "  " + cmbAndOr.Text + " a.idSubGrupo = " + cmbSubGrupo.Value.ToString();
		}

		private void btMarca_Click(object sender, System.EventArgs e)
		{
			txtFiltro.Text += "  " + cmbAndOr.Text + " a.idMarca = " + cmbMarca.Value.ToString();
		}

		private void btTipoGrupo_Click(object sender, System.EventArgs e)
		{
			txtFiltro.Text += "  " + cmbAndOr.Text + " a.idTipoGrupo = " + cmbTipo.Value.ToString();
		}

		private void btArtSri_Click(object sender, System.EventArgs e)
		{
			txtFiltro.Text += "  " + cmbAndOr.Text + " a.idArticuloSRI = " + cmbArtSri.Value.ToString();
		}

		private void btReporte_Click(object sender, System.EventArgs e)
		{
			if (ultraGrid1.ActiveRow == null)
			{
				MessageBox.Show("Escoja Articulo" ,"Información");
				return;
			}
			Cursor = Cursors.WaitCursor;
			DateTime dtDesde = (DateTime) cmbDesde.Value;
			DateTime dtHasta = (DateTime) cmbHasta.Value;
			string stTitulo = "Desde " + dtDesde.ToString("dd/MMM/yyyy")
				+ " hasta " + dtHasta.ToString("dd/MMM/yyyy");
			ParameterFields paramFields = new ParameterFields ();

			#region Parametro Fecha Desde Hasta
			ParameterField paramField1 = new ParameterField ();
			ParameterDiscreteValue discreteVal1 = new ParameterDiscreteValue ();
			paramField1.ParameterFieldName = "@FechaDesde";
			discreteVal1.Value = dtDesde.ToString("yyyy-MM-dd HH:mm:ss");
			paramField1.CurrentValues.Add (discreteVal1);
			paramFields.Add (paramField1);
		
			ParameterField paramField2 = new ParameterField ();
			ParameterDiscreteValue discreteVal2 = new ParameterDiscreteValue ();
			paramField2.ParameterFieldName = "@FechaHasta";
			discreteVal2.Value = dtHasta.ToString("yyyy-MM-dd HH:mm:ss");
			paramField2.CurrentValues.Add (discreteVal2);
			paramFields.Add (paramField2);
			#endregion Parametro Fecha Desde

			string stReporte = cmbReporte.Text;
			ParameterField paramField3 = new ParameterField ();
			ParameterDiscreteValue discreteVal3 = new ParameterDiscreteValue ();
			if (stReporte.ToUpper().StartsWith("KARDEX"))
			{
				paramField3.ParameterFieldName = "@idArticulo";
				discreteVal3.Value = int.Parse(ultraGrid1.ActiveRow.Cells["idArticulo"].Value.ToString());
				paramField3.CurrentValues.Add (discreteVal3);
			}
			else
			{
				paramField3.ParameterFieldName = "@Filtro";
				discreteVal3.Value = txtFiltro.Text;
				paramField3.CurrentValues.Add (discreteVal3);
			}
			paramFields.Add (paramField3);

			ParameterField paramField4 = new ParameterField ();
			ParameterDiscreteValue discreteVal4 = new ParameterDiscreteValue ();
			paramField4.ParameterFieldName = "@Bodega";
			discreteVal4.Value = spnBodega.Value;
			paramField4.CurrentValues.Add (discreteVal4);
			paramFields.Add (paramField4);

			Reporte miRep = new Reporte(stReporte);
			miRep.MdiParent = this.MdiParent;
			miRep.crVista.ParameterFieldInfo = paramFields;
			miRep.Titulo(stTitulo);
			Cursor = Cursors.Default;
			miRep.Show();
		}

		private string FiltroSerie()
		{
			string stFiltroAntes = "%";
			string stFiltroDespues = "%";
			int iTipo = (int) optBusca.Value;
			if (iTipo == 0) stFiltroAntes = "";
			if (iTipo == 1) 
			{
				stFiltroAntes = "";
				stFiltroDespues = "";
			}
			string stSerie = txtSerie.Text.ToString().Trim();
			return "'" + stFiltroAntes + stSerie + stFiltroDespues + "'";		
		}

		private void btBuscaKardex_Click(object sender, System.EventArgs e)
		{
			dvK.RowFilter = "Serie Like " + FiltroSerie();
		}

		private void btBuscaInventario_Click(object sender, System.EventArgs e)
		{
			string stFiltroAntes = "%";
			string stFiltroDespues = "%";
			int iTipo = (int) optBusca.Value;
			if (iTipo == 0) stFiltroAntes = "";
			if (iTipo == 1) 
			{
				stFiltroAntes = "";
				stFiltroDespues = "";
			}
			string stSerie = txtArticulo.Text.ToString();
			dv.RowFilter = "Articulo Like '" + stFiltroAntes + stSerie 
				+ stFiltroDespues + "'";		
		}

		private void btBuscaSerie_Click(object sender, System.EventArgs e)
		{
			Funcion miFuncion = new Funcion();
			string strConn = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);	
			string stProcedimiento = "EXEC SerieBuscar " + FiltroSerie() + ", "
				+ spnBodega.Value.ToString() + ", '" + txtArticulo.Text.Trim() + "'";
			SqlDataAdapter myData = new SqlDataAdapter(stProcedimiento, strConn);

			myData.TableMappings.Add("Tabla", "ExcelSube");
			
			miDataSet.Reset();
			miDataSet.Clear();
			try
			{
				myData.Fill(miDataSet);
			}
			catch(Exception ex)
			{
				MessageBox.Show(ex.Message, "Error al abrir procedimiento SerieBusca" );
				return;
			}

			dv = miDataSet.Tables[0].DefaultView;
			ultraGrid3.DataSource = dv;
			ultraGrid3.DisplayLayout.Bands[0].Columns["idCompra"].Hidden = true;
			ultraGrid3.DisplayLayout.Bands[0].Columns["Fecha"].Format = "dd/MMM/yyyy";
			ultraGrid3.DisplayLayout.Bands[0].Columns["Precio"].Format = "#,##0.00";
			ultraGrid3.DataBind();
		}

		private void btInconcistencia_Click(object sender, System.EventArgs e)
		{
			Cursor = Cursors.WaitCursor;
			Funcion miFuncion = new Funcion();
			string strConn = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);	
			string stProcedimiento = "EXEC SerieInconcistencia ";
			//+ FiltroSerie() + ", "+ spnBodega.Value.ToString();
			SqlDataAdapter myData = new SqlDataAdapter(stProcedimiento, strConn);

			myData.TableMappings.Add("Tabla", "ExcelSube");
			
			miDataSet.Reset();
			miDataSet.Clear();
			try
			{
				myData.Fill(miDataSet);
			}
			catch(Exception ex)
			{
				MessageBox.Show(ex.Message, "Error al abrir procedimiento SerieBusca" );
				return;
			}

			dv = miDataSet.Tables[0].DefaultView;
			ultraGrid3.DataSource = dv;
			ultraGrid3.DisplayLayout.Bands[0].Columns["idCompra"].Hidden = true;
			ultraGrid3.DisplayLayout.Bands[0].Columns["Fecha"].Format = "dd/MMM/yyyy";
			ultraGrid3.DisplayLayout.Bands[0].Columns["Precio"].Format = "#,##0.00";
			ultraGrid3.DataBind();
			Cursor = Cursors.Default;
		}

		private void btFacturaS_Click(object sender, System.EventArgs e)
		{
			int IdCompra = int.Parse(ultraGrid3.ActiveRow.Cells["idCompra"].Text);
			if (IdCompra == 0)
			{
				MessageBox.Show("Linea de serie no tiene una factura" ,"Información");
				return;
			}

			int IdTipoFactura = int.Parse(ultraGrid3.ActiveRow.Cells["idTipoFactura"].Text);
			if (IdTipoFactura == 0)
			{
				MessageBox.Show("Linea de serie no tiene tipo de factura" ,"Información");
				return;
			}

//			Compra miCompra = new Compra(IdTipoFactura, IdCompra);
//			miCompra.MdiParent = this.MdiParent;
//			miCompra.Show();
		}

		private void btDisponible_Click(object sender, System.EventArgs e)
		{
			if (DialogResult.No == MessageBox.Show("Desea cambiar serie a Disponible",
				"Información", MessageBoxButtons.YesNo, MessageBoxIcon.Question)) return;
			int IdSerie = int.Parse(ultraGrid3.ActiveRow.Cells["idSerie"].Text);
			if (IdSerie == 0)
			{
				MessageBox.Show("Linea de serie no tiene Serie" ,"Información");
				return;
			}
			string stSelect = "Update Serie Set Disponible = 1 Where idSerie = "
				+ IdSerie.ToString();
      Funcion.EjecutaSQL(cdsGrupo, stSelect, true);
			MessageBox.Show("Serie Cambiada a Disponible", "Información");
		}

		private void btVendido_Click(object sender, System.EventArgs e)
		{
			if (DialogResult.No == MessageBox.Show("Desea cambiar serie a NO Disponible",
				"Información", MessageBoxButtons.YesNo, MessageBoxIcon.Question)) return;
			int IdSerie = int.Parse(ultraGrid3.ActiveRow.Cells["idSerie"].Text);
			if (IdSerie == 0)
			{
				MessageBox.Show("Linea de serie no tiene Serie" ,"Información");
				return;
			}
			string stSelect = "Update Serie Set Disponible = 0 Where idSerie = "
				+ IdSerie.ToString();
			Funcion.EjecutaSQL(cdsGrupo, stSelect, true);
			MessageBox.Show("Serie Cambiada a NO Disponible", "Información");
		}

		private void btSerie_Click(object sender, System.EventArgs e)
		{
			int idArticulo = int.Parse(ultraGrid3.ActiveRow.Cells["idArticulo"].Text);
			if (idArticulo == 0)
			{
				MessageBox.Show("Linea de serie no tiene Articulo" ,"Información");
				return;
			}
			int idCompra = int.Parse(ultraGrid3.ActiveRow.Cells["idCompra"].Text);
			if (idCompra == 0)
			{
				MessageBox.Show("Linea de serie no tiene Compra" ,"Información");
				return;
			}
			int idDetCompra = int.Parse(ultraGrid3.ActiveRow.Cells["idDetCompra"].Text);
			if (idDetCompra == 0)
			{
				MessageBox.Show("Linea de serie no tiene DetCompra" ,"Información");
				return;
			}
//			CompraSerie miSerie = new CompraSerie(0, idArticulo, idCompra, "", 0, 0, idDetCompra, false, 0, false);
//			miSerie.ShowDialog();
		}
	}
}
