using System;
using System.Data.SqlClient;
using C1.Data;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using CrystalDecisions.Shared;
using Infragistics.Win.UltraWinGrid;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de PlanCuentas.
	/// </summary>
	public class AsignaCuentaArbol : DevExpress.XtraEditors.XtraForm
	{
		private DevExpress.XtraTreeList.TreeList tlsCuentas;
		private DevExpress.XtraEditors.SimpleButton btNoAsignar;
		private DevExpress.XtraEditors.SimpleButton btCancelar;
		private DevExpress.XtraEditors.SimpleButton btAceptar;
		private C1.Data.C1DataView cdvCuentaCod;
		private C1.Data.C1DataSet cdsCuenta;
		private C1.Data.C1DataSet cdsAsignaCuenta;
		private C1.Data.C1DataView cdvCuenta;
		private DevExpress.XtraEditors.Repository.RepositoryItemLookUpEdit repositoryItemLookUpEdit1;
		private DevExpress.XtraEditors.Repository.RepositoryItemLookUpEdit repositoryItemLookUpEdit2;
		private DevExpress.XtraTreeList.Columns.TreeListColumn colAsignacion;
		private DevExpress.XtraTreeList.Columns.TreeListColumn colidCuenta;
		private DevExpress.XtraTreeList.Columns.TreeListColumn colCuenta;
		private System.Windows.Forms.TextBox txtIdAsignacion;
		private DevExpress.XtraEditors.SimpleButton btRetencionFuente;
		private DevExpress.XtraEditors.SimpleButton btRetIva;
		private DevExpress.XtraEditors.SimpleButton btGrupoCl;
		private DevExpress.XtraEditors.SimpleButton btGrupoItems;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbCuenta;
		private System.Windows.Forms.Label label2;
		private DevExpress.XtraEditors.SimpleButton btCreaGrupoClient;
		private C1.Data.C1DataView cdvCuentaTodo;
		private System.Windows.Forms.Label label4;
		private DevExpress.XtraEditors.SimpleButton btCreaGrupoItem;
		private System.Windows.Forms.Label label3;
		private System.Windows.Forms.Label label1;
		private System.Windows.Forms.Panel panel1;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbArticuloSRI;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbTipo;
		private System.Windows.Forms.ErrorProvider errorProvider1;
		private DevExpress.XtraTab.XtraTabControl xtraTabControl1;
		private DevExpress.XtraTab.XtraTabPage xtraTabPage1;
		private DevExpress.XtraTab.XtraTabPage xtraTabPage2;
		private DevExpress.XtraTab.XtraTabPage xtraTabPage3;
		private DevExpress.XtraTab.XtraTabPage xtraTabPage4;
		private System.Windows.Forms.Label label7;
		private System.Windows.Forms.Label label8;
		private DevExpress.XtraEditors.SimpleButton btGrupo;
		private DevExpress.XtraEditors.CheckEdit chkCrearItem;
		private DevExpress.XtraEditors.SimpleButton btIntercambio;
		private System.Windows.Forms.Label lblCosto;
		private System.Windows.Forms.Label lblVenta;
		private DevExpress.XtraEditors.LookUpEdit cmbCuentaCompra;
		private DevExpress.XtraEditors.LookUpEdit cmbCuentaVenta;
		private DevExpress.XtraEditors.LookUpEdit cmbCosto;
		private DevExpress.XtraEditors.LookUpEdit cmbDescuento;
		private DevExpress.XtraEditors.LookUpEdit cmbEgreso;
		private DevExpress.XtraEditors.GroupControl groupControl1;
		private DevExpress.XtraEditors.GroupControl groupControl2;
		private Infragistics.Win.UltraWinGrid.UltraGrid ultraGrid1;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource udsTipo;
		private Infragistics.Win.UltraWinGrid.UltraDropDown cmbCuentaFP;
		private C1.Data.C1DataSet cdsFormaPago;
		private Infragistics.Win.UltraWinGrid.UltraDropDown cmbTipoFP;
		private System.ComponentModel.IContainer components;

		public AsignaCuentaArbol()
		{
			//
			// Necesario para admitir el Diseñador de Windows Forms
			//
			InitializeComponent();

			//
			// TODO: Agregar código de constructor después de llamar a InitializeComponent
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

		#region Windows Form Designer generated code
		/// <summary>
		/// Método necesario para admitir el Diseñador, no se puede modificar
		/// el contenido del método con el editor de código.
		/// </summary>
		private void InitializeComponent()
		{
			Infragistics.Win.Appearance appearance1 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Cuenta", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCuenta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Grupo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CodRapido");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Descripcion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CuentaOrigen");
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("ArticuloTipo", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idTipoGrupo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Tipo");
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand3 = new Infragistics.Win.UltraWinGrid.UltraGridBand("ArticuloSRI", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idArticuloSRI");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Detalle");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn11 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Porcentaje");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn12 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand4 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn13 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idTipo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn14 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Tipo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idTipo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Tipo");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand5 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Cuenta", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn15 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn16 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCuenta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn17 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Grupo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn18 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CodRapido");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn19 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Descripcion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn20 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CuentaOrigen");
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand6 = new Infragistics.Win.UltraWinGrid.UltraGridBand("CompraFormaPago", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn21 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idFormaPago");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn22 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPrecio");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn23 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FormaPago");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn24 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCuenta", -1, "cmbCuentaFP");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn25 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Borrar");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn26 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Descuento");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn27 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Ruc");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn28 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idTipo", -1, "cmbTipoFP");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn29 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance11 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance12 = new Infragistics.Win.Appearance();
			this.tlsCuentas = new DevExpress.XtraTreeList.TreeList();
			this.colAsignacion = new DevExpress.XtraTreeList.Columns.TreeListColumn();
			this.colidCuenta = new DevExpress.XtraTreeList.Columns.TreeListColumn();
			this.repositoryItemLookUpEdit1 = new DevExpress.XtraEditors.Repository.RepositoryItemLookUpEdit();
			this.cdvCuentaCod = new C1.Data.C1DataView();
			this.cdsCuenta = new C1.Data.C1DataSet();
			this.colCuenta = new DevExpress.XtraTreeList.Columns.TreeListColumn();
			this.repositoryItemLookUpEdit2 = new DevExpress.XtraEditors.Repository.RepositoryItemLookUpEdit();
			this.cdsAsignaCuenta = new C1.Data.C1DataSet();
			this.btNoAsignar = new DevExpress.XtraEditors.SimpleButton();
			this.btCancelar = new DevExpress.XtraEditors.SimpleButton();
			this.btAceptar = new DevExpress.XtraEditors.SimpleButton();
			this.cdvCuenta = new C1.Data.C1DataView();
			this.txtIdAsignacion = new System.Windows.Forms.TextBox();
			this.btRetencionFuente = new DevExpress.XtraEditors.SimpleButton();
			this.btRetIva = new DevExpress.XtraEditors.SimpleButton();
			this.btGrupoCl = new DevExpress.XtraEditors.SimpleButton();
			this.btGrupoItems = new DevExpress.XtraEditors.SimpleButton();
			this.cmbCuenta = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.cdvCuentaTodo = new C1.Data.C1DataView();
			this.label2 = new System.Windows.Forms.Label();
			this.btCreaGrupoClient = new DevExpress.XtraEditors.SimpleButton();
			this.panel1 = new System.Windows.Forms.Panel();
			this.cmbEgreso = new DevExpress.XtraEditors.LookUpEdit();
			this.cmbDescuento = new DevExpress.XtraEditors.LookUpEdit();
			this.cmbCosto = new DevExpress.XtraEditors.LookUpEdit();
			this.cmbCuentaVenta = new DevExpress.XtraEditors.LookUpEdit();
			this.cmbCuentaCompra = new DevExpress.XtraEditors.LookUpEdit();
			this.btIntercambio = new DevExpress.XtraEditors.SimpleButton();
			this.chkCrearItem = new DevExpress.XtraEditors.CheckEdit();
			this.btGrupo = new DevExpress.XtraEditors.SimpleButton();
			this.label8 = new System.Windows.Forms.Label();
			this.label7 = new System.Windows.Forms.Label();
			this.lblCosto = new System.Windows.Forms.Label();
			this.lblVenta = new System.Windows.Forms.Label();
			this.label3 = new System.Windows.Forms.Label();
			this.label4 = new System.Windows.Forms.Label();
			this.label1 = new System.Windows.Forms.Label();
			this.btCreaGrupoItem = new DevExpress.XtraEditors.SimpleButton();
			this.cmbTipo = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.cmbArticuloSRI = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.errorProvider1 = new System.Windows.Forms.ErrorProvider();
			this.xtraTabControl1 = new DevExpress.XtraTab.XtraTabControl();
			this.xtraTabPage1 = new DevExpress.XtraTab.XtraTabPage();
			this.xtraTabPage2 = new DevExpress.XtraTab.XtraTabPage();
			this.groupControl2 = new DevExpress.XtraEditors.GroupControl();
			this.cmbTipoFP = new Infragistics.Win.UltraWinGrid.UltraDropDown();
			this.udsTipo = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.cmbCuentaFP = new Infragistics.Win.UltraWinGrid.UltraDropDown();
			this.ultraGrid1 = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.cdsFormaPago = new C1.Data.C1DataSet();
			this.groupControl1 = new DevExpress.XtraEditors.GroupControl();
			this.xtraTabPage3 = new DevExpress.XtraTab.XtraTabPage();
			this.xtraTabPage4 = new DevExpress.XtraTab.XtraTabPage();
			((System.ComponentModel.ISupportInitialize)(this.tlsCuentas)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.repositoryItemLookUpEdit1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvCuentaCod)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsCuenta)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.repositoryItemLookUpEdit2)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsAsignaCuenta)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvCuenta)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCuenta)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvCuentaTodo)).BeginInit();
			this.panel1.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.cmbEgreso.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbDescuento.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCosto.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCuentaVenta.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCuentaCompra.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.chkCrearItem.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbTipo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbArticuloSRI)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.xtraTabControl1)).BeginInit();
			this.xtraTabControl1.SuspendLayout();
			this.xtraTabPage1.SuspendLayout();
			this.xtraTabPage2.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.groupControl2)).BeginInit();
			this.groupControl2.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.cmbTipoFP)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.udsTipo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCuentaFP)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsFormaPago)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.groupControl1)).BeginInit();
			this.groupControl1.SuspendLayout();
			this.xtraTabPage3.SuspendLayout();
			this.xtraTabPage4.SuspendLayout();
			this.SuspendLayout();
			// 
			// tlsCuentas
			// 
			this.tlsCuentas.Appearance.Empty.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(243)), ((System.Byte)(243)), ((System.Byte)(243)));
			this.tlsCuentas.Appearance.Empty.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(255)), ((System.Byte)(255)), ((System.Byte)(255)));
			this.tlsCuentas.Appearance.Empty.Options.UseBackColor = true;
			this.tlsCuentas.Appearance.Empty.Options.UseForeColor = true;
			this.tlsCuentas.Appearance.EvenRow.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(223)), ((System.Byte)(223)), ((System.Byte)(223)));
			this.tlsCuentas.Appearance.EvenRow.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F);
			this.tlsCuentas.Appearance.EvenRow.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(0)), ((System.Byte)(0)), ((System.Byte)(0)));
			this.tlsCuentas.Appearance.EvenRow.Options.UseBackColor = true;
			this.tlsCuentas.Appearance.EvenRow.Options.UseFont = true;
			this.tlsCuentas.Appearance.EvenRow.Options.UseForeColor = true;
			this.tlsCuentas.Appearance.FocusedCell.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(255)), ((System.Byte)(255)), ((System.Byte)(225)));
			this.tlsCuentas.Appearance.FocusedCell.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F);
			this.tlsCuentas.Appearance.FocusedCell.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(0)), ((System.Byte)(0)), ((System.Byte)(0)));
			this.tlsCuentas.Appearance.FocusedCell.Options.UseBackColor = true;
			this.tlsCuentas.Appearance.FocusedCell.Options.UseFont = true;
			this.tlsCuentas.Appearance.FocusedCell.Options.UseForeColor = true;
			this.tlsCuentas.Appearance.FocusedRow.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(0)), ((System.Byte)(0)), ((System.Byte)(128)));
			this.tlsCuentas.Appearance.FocusedRow.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F);
			this.tlsCuentas.Appearance.FocusedRow.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(255)), ((System.Byte)(255)), ((System.Byte)(255)));
			this.tlsCuentas.Appearance.FocusedRow.Options.UseBackColor = true;
			this.tlsCuentas.Appearance.FocusedRow.Options.UseFont = true;
			this.tlsCuentas.Appearance.FocusedRow.Options.UseForeColor = true;
			this.tlsCuentas.Appearance.FooterPanel.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(192)), ((System.Byte)(192)), ((System.Byte)(192)));
			this.tlsCuentas.Appearance.FooterPanel.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F);
			this.tlsCuentas.Appearance.FooterPanel.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(0)), ((System.Byte)(0)), ((System.Byte)(0)));
			this.tlsCuentas.Appearance.FooterPanel.Options.UseBackColor = true;
			this.tlsCuentas.Appearance.FooterPanel.Options.UseFont = true;
			this.tlsCuentas.Appearance.FooterPanel.Options.UseForeColor = true;
			this.tlsCuentas.Appearance.GroupButton.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(255)), ((System.Byte)(255)), ((System.Byte)(255)));
			this.tlsCuentas.Appearance.GroupButton.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F);
			this.tlsCuentas.Appearance.GroupButton.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(0)), ((System.Byte)(0)), ((System.Byte)(0)));
			this.tlsCuentas.Appearance.GroupButton.Options.UseBackColor = true;
			this.tlsCuentas.Appearance.GroupButton.Options.UseFont = true;
			this.tlsCuentas.Appearance.GroupButton.Options.UseForeColor = true;
			this.tlsCuentas.Appearance.GroupFooter.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(202)), ((System.Byte)(202)), ((System.Byte)(202)));
			this.tlsCuentas.Appearance.GroupFooter.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(0)), ((System.Byte)(0)), ((System.Byte)(0)));
			this.tlsCuentas.Appearance.GroupFooter.Options.UseBackColor = true;
			this.tlsCuentas.Appearance.GroupFooter.Options.UseForeColor = true;
			this.tlsCuentas.Appearance.HeaderPanel.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(192)), ((System.Byte)(192)), ((System.Byte)(192)));
			this.tlsCuentas.Appearance.HeaderPanel.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Bold);
			this.tlsCuentas.Appearance.HeaderPanel.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(0)), ((System.Byte)(0)), ((System.Byte)(0)));
			this.tlsCuentas.Appearance.HeaderPanel.Options.UseBackColor = true;
			this.tlsCuentas.Appearance.HeaderPanel.Options.UseFont = true;
			this.tlsCuentas.Appearance.HeaderPanel.Options.UseForeColor = true;
			this.tlsCuentas.Appearance.HideSelectionRow.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(128)), ((System.Byte)(128)), ((System.Byte)(128)));
			this.tlsCuentas.Appearance.HideSelectionRow.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F);
			this.tlsCuentas.Appearance.HideSelectionRow.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(212)), ((System.Byte)(208)), ((System.Byte)(200)));
			this.tlsCuentas.Appearance.HideSelectionRow.Options.UseBackColor = true;
			this.tlsCuentas.Appearance.HideSelectionRow.Options.UseFont = true;
			this.tlsCuentas.Appearance.HideSelectionRow.Options.UseForeColor = true;
			this.tlsCuentas.Appearance.HorzLine.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(192)), ((System.Byte)(192)), ((System.Byte)(192)));
			this.tlsCuentas.Appearance.HorzLine.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F);
			this.tlsCuentas.Appearance.HorzLine.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(128)), ((System.Byte)(128)), ((System.Byte)(128)));
			this.tlsCuentas.Appearance.HorzLine.Options.UseBackColor = true;
			this.tlsCuentas.Appearance.HorzLine.Options.UseFont = true;
			this.tlsCuentas.Appearance.HorzLine.Options.UseForeColor = true;
			this.tlsCuentas.Appearance.OddRow.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(255)), ((System.Byte)(255)), ((System.Byte)(255)));
			this.tlsCuentas.Appearance.OddRow.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F);
			this.tlsCuentas.Appearance.OddRow.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(0)), ((System.Byte)(0)), ((System.Byte)(0)));
			this.tlsCuentas.Appearance.OddRow.Options.UseBackColor = true;
			this.tlsCuentas.Appearance.OddRow.Options.UseFont = true;
			this.tlsCuentas.Appearance.OddRow.Options.UseForeColor = true;
			this.tlsCuentas.Appearance.Preview.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(255)), ((System.Byte)(255)), ((System.Byte)(255)));
			this.tlsCuentas.Appearance.Preview.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F);
			this.tlsCuentas.Appearance.Preview.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(0)), ((System.Byte)(0)), ((System.Byte)(128)));
			this.tlsCuentas.Appearance.Preview.Options.UseBackColor = true;
			this.tlsCuentas.Appearance.Preview.Options.UseFont = true;
			this.tlsCuentas.Appearance.Preview.Options.UseForeColor = true;
			this.tlsCuentas.Appearance.Row.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(255)), ((System.Byte)(255)), ((System.Byte)(255)));
			this.tlsCuentas.Appearance.Row.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F);
			this.tlsCuentas.Appearance.Row.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(0)), ((System.Byte)(0)), ((System.Byte)(0)));
			this.tlsCuentas.Appearance.Row.Options.UseBackColor = true;
			this.tlsCuentas.Appearance.Row.Options.UseFont = true;
			this.tlsCuentas.Appearance.Row.Options.UseForeColor = true;
			this.tlsCuentas.Appearance.SelectedRow.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(10)), ((System.Byte)(10)), ((System.Byte)(138)));
			this.tlsCuentas.Appearance.SelectedRow.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F);
			this.tlsCuentas.Appearance.SelectedRow.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(255)), ((System.Byte)(255)), ((System.Byte)(255)));
			this.tlsCuentas.Appearance.SelectedRow.Options.UseBackColor = true;
			this.tlsCuentas.Appearance.SelectedRow.Options.UseFont = true;
			this.tlsCuentas.Appearance.SelectedRow.Options.UseForeColor = true;
			this.tlsCuentas.Appearance.TreeLine.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(255)), ((System.Byte)(255)), ((System.Byte)(255)));
			this.tlsCuentas.Appearance.TreeLine.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(128)), ((System.Byte)(128)), ((System.Byte)(128)));
			this.tlsCuentas.Appearance.TreeLine.Options.UseBackColor = true;
			this.tlsCuentas.Appearance.TreeLine.Options.UseForeColor = true;
			this.tlsCuentas.Appearance.VertLine.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(192)), ((System.Byte)(192)), ((System.Byte)(192)));
			this.tlsCuentas.Appearance.VertLine.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F);
			this.tlsCuentas.Appearance.VertLine.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(128)), ((System.Byte)(128)), ((System.Byte)(128)));
			this.tlsCuentas.Appearance.VertLine.Options.UseBackColor = true;
			this.tlsCuentas.Appearance.VertLine.Options.UseFont = true;
			this.tlsCuentas.Appearance.VertLine.Options.UseForeColor = true;
			this.tlsCuentas.Appearance.VertLine.Options.UseTextOptions = true;
			this.tlsCuentas.Appearance.VertLine.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Near;
			this.tlsCuentas.Columns.AddRange(new DevExpress.XtraTreeList.Columns.TreeListColumn[] {
																																															this.colAsignacion,
																																															this.colidCuenta,
																																															this.colCuenta});
			this.tlsCuentas.DataMember = "AsignaCuenta";
			this.tlsCuentas.DataSource = this.cdsAsignaCuenta;
			this.tlsCuentas.Dock = System.Windows.Forms.DockStyle.Fill;
			this.tlsCuentas.ImageIndexFieldName = "";
			this.tlsCuentas.KeyFieldName = "idAsignacion";
			this.tlsCuentas.Location = new System.Drawing.Point(0, 0);
			this.tlsCuentas.Name = "tlsCuentas";
			this.tlsCuentas.ParentFieldName = "idPadre";
			this.tlsCuentas.RepositoryItems.AddRange(new DevExpress.XtraEditors.Repository.RepositoryItem[] {
																																																				this.repositoryItemLookUpEdit1,
																																																				this.repositoryItemLookUpEdit2});
			this.tlsCuentas.RootValue = "0";
			this.tlsCuentas.Size = new System.Drawing.Size(647, 346);
			this.tlsCuentas.Styles.AddReplace("PressedColumn", new DevExpress.Utils.ViewStyle("PressedColumn", "TreeList", new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Bold), "HeaderPanel", ((DevExpress.Utils.StyleOptions)(((DevExpress.Utils.StyleOptions.StyleEnabled | DevExpress.Utils.StyleOptions.UseBackColor) 
				| DevExpress.Utils.StyleOptions.UseForeColor))), true, false, false, DevExpress.Utils.HorzAlignment.Near, DevExpress.Utils.VertAlignment.Center, null, System.Drawing.Color.FromArgb(((System.Byte)(0)), ((System.Byte)(0)), ((System.Byte)(128))), System.Drawing.Color.FromArgb(((System.Byte)(255)), ((System.Byte)(255)), ((System.Byte)(255)))));
			this.tlsCuentas.TabIndex = 12;
			this.tlsCuentas.FocusedNodeChanged += new DevExpress.XtraTreeList.FocusedNodeChangedEventHandler(this.tlsCuentas_FocusedNodeChanged);
			// 
			// colAsignacion
			// 
			this.colAsignacion.Caption = "Asignación";
			this.colAsignacion.FieldName = "Asignacion";
			this.colAsignacion.Name = "colAsignacion";
			this.colAsignacion.OptionsColumn.AllowEdit = false;
			this.colAsignacion.SortOrder = System.Windows.Forms.SortOrder.Ascending;
			this.colAsignacion.VisibleIndex = 0;
			this.colAsignacion.Width = 118;
			// 
			// colidCuenta
			// 
			this.colidCuenta.Caption = "Código";
			this.colidCuenta.ColumnEdit = this.repositoryItemLookUpEdit1;
			this.colidCuenta.FieldName = "idCuenta";
			this.colidCuenta.Format.FormatType = DevExpress.Utils.FormatType.Numeric;
			this.colidCuenta.Name = "colidCuenta";
			this.colidCuenta.VisibleIndex = 1;
			this.colidCuenta.Width = 119;
			// 
			// repositoryItemLookUpEdit1
			// 
			this.repositoryItemLookUpEdit1.AutoHeight = false;
			this.repositoryItemLookUpEdit1.Buttons.AddRange(new DevExpress.XtraEditors.Controls.EditorButton[] {
																																																					 new DevExpress.XtraEditors.Controls.EditorButton(DevExpress.XtraEditors.Controls.ButtonPredefines.Combo)});
			this.repositoryItemLookUpEdit1.Columns.AddRange(new DevExpress.XtraEditors.Controls.LookUpColumnInfo[] {
																																																							 new DevExpress.XtraEditors.Controls.LookUpColumnInfo("CodRapido", "CodRapido", 66, DevExpress.Utils.FormatType.None, "", true, DevExpress.Utils.HorzAlignment.Near, DevExpress.Data.ColumnSortOrder.Ascending),
																																																							 new DevExpress.XtraEditors.Controls.LookUpColumnInfo("Descripcion", "Descripcion", 200, DevExpress.Utils.FormatType.None, "", true, DevExpress.Utils.HorzAlignment.Near)});
			this.repositoryItemLookUpEdit1.DataSource = this.cdvCuentaCod;
			this.repositoryItemLookUpEdit1.DisplayMember = "CodRapido";
			this.repositoryItemLookUpEdit1.Name = "repositoryItemLookUpEdit1";
			this.repositoryItemLookUpEdit1.NullText = "";
			this.repositoryItemLookUpEdit1.PopupFormWidth = 250;
			this.repositoryItemLookUpEdit1.PopupWidth = 250;
			this.repositoryItemLookUpEdit1.ValueMember = "idCuenta";
			// 
			// cdvCuentaCod
			// 
			this.cdvCuentaCod.BindingContextCtrl = this;
			this.cdvCuentaCod.DataSet = this.cdsCuenta;
			this.cdvCuentaCod.RowFilter = "Grupo = 0";
			this.cdvCuentaCod.Sort = "Descripcion";
			this.cdvCuentaCod.TableName = "";
			this.cdvCuentaCod.TableViewName = "Cuenta";
			// 
			// cdsCuenta
			// 
			this.cdsCuenta.BindingContextCtrl = this;
			this.cdsCuenta.CaseSensitive = false;
			this.cdsCuenta.DataLibrary = "LibContabilidad";
			this.cdsCuenta.DataLibraryUrl = "";
			this.cdsCuenta.DataSetDef = "dsCuenta2";
			this.cdsCuenta.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsCuenta.Name = "cdsCuenta";
			this.cdsCuenta.SchemaClassName = "LibContabilidad.DataClass";
			this.cdsCuenta.SchemaDef = null;
			this.cdsCuenta.BeforeFill += new C1.Data.FillEventHandler(this.cdsCuenta_BeforeFill);
			// 
			// colCuenta
			// 
			this.colCuenta.Caption = "Cuenta";
			this.colCuenta.ColumnEdit = this.repositoryItemLookUpEdit2;
			this.colCuenta.FieldName = "idCuenta";
			this.colCuenta.Format.FormatType = DevExpress.Utils.FormatType.Numeric;
			this.colCuenta.Name = "colCuenta";
			this.colCuenta.VisibleIndex = 2;
			this.colCuenta.Width = 224;
			// 
			// repositoryItemLookUpEdit2
			// 
			this.repositoryItemLookUpEdit2.AutoHeight = false;
			this.repositoryItemLookUpEdit2.Buttons.AddRange(new DevExpress.XtraEditors.Controls.EditorButton[] {
																																																					 new DevExpress.XtraEditors.Controls.EditorButton(DevExpress.XtraEditors.Controls.ButtonPredefines.Combo)});
			this.repositoryItemLookUpEdit2.Columns.AddRange(new DevExpress.XtraEditors.Controls.LookUpColumnInfo[] {
																																																							 new DevExpress.XtraEditors.Controls.LookUpColumnInfo("Descripcion", "Descripcion", 200, DevExpress.Utils.FormatType.None, "", true, DevExpress.Utils.HorzAlignment.Near, DevExpress.Data.ColumnSortOrder.Ascending),
																																																							 new DevExpress.XtraEditors.Controls.LookUpColumnInfo("Codigo", "Codigo", 80, DevExpress.Utils.FormatType.None, "", true, DevExpress.Utils.HorzAlignment.Near)});
			this.repositoryItemLookUpEdit2.DataSource = this.cdvCuentaCod;
			this.repositoryItemLookUpEdit2.DisplayMember = "Descripcion";
			this.repositoryItemLookUpEdit2.Name = "repositoryItemLookUpEdit2";
			this.repositoryItemLookUpEdit2.NullText = "";
			this.repositoryItemLookUpEdit2.PopupFormWidth = 250;
			this.repositoryItemLookUpEdit2.PopupWidth = 250;
			this.repositoryItemLookUpEdit2.ValueMember = "idCuenta";
			// 
			// cdsAsignaCuenta
			// 
			this.cdsAsignaCuenta.BindingContextCtrl = this;
			this.cdsAsignaCuenta.DataLibrary = "LibContabilidad";
			this.cdsAsignaCuenta.DataLibraryUrl = "";
			this.cdsAsignaCuenta.DataSetDef = "dsAsignaCuenta";
			this.cdsAsignaCuenta.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsAsignaCuenta.Name = "cdsAsignaCuenta";
			this.cdsAsignaCuenta.SchemaClassName = "LibContabilidad.DataClass";
			this.cdsAsignaCuenta.SchemaDef = null;
			// 
			// btNoAsignar
			// 
			this.btNoAsignar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btNoAsignar.Location = new System.Drawing.Point(128, 405);
			this.btNoAsignar.Name = "btNoAsignar";
			this.btNoAsignar.Size = new System.Drawing.Size(75, 25);
			this.btNoAsignar.TabIndex = 40;
			this.btNoAsignar.Text = "&No Asignar";
			this.btNoAsignar.Click += new System.EventHandler(this.btNoAsignar_Click);
			// 
			// btCancelar
			// 
			this.btCancelar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.btCancelar.DialogResult = System.Windows.Forms.DialogResult.Cancel;
			this.btCancelar.Location = new System.Drawing.Point(592, 405);
			this.btCancelar.Name = "btCancelar";
			this.btCancelar.Size = new System.Drawing.Size(75, 25);
			this.btCancelar.TabIndex = 39;
			this.btCancelar.Text = "&Cancelar";
			this.btCancelar.Click += new System.EventHandler(this.btCancelar_Click);
			// 
			// btAceptar
			// 
			this.btAceptar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btAceptar.Location = new System.Drawing.Point(32, 405);
			this.btAceptar.Name = "btAceptar";
			this.btAceptar.Size = new System.Drawing.Size(75, 25);
			this.btAceptar.TabIndex = 38;
			this.btAceptar.Text = "&Aceptar";
			this.btAceptar.Click += new System.EventHandler(this.btAceptar_Click);
			// 
			// cdvCuenta
			// 
			this.cdvCuenta.BindingContextCtrl = this;
			this.cdvCuenta.DataSet = this.cdsCuenta;
			this.cdvCuenta.RowFilter = "Grupo = 0";
			this.cdvCuenta.Sort = "Codigo";
			this.cdvCuenta.TableName = "";
			this.cdvCuenta.TableViewName = "Cuenta";
			// 
			// txtIdAsignacion
			// 
			this.txtIdAsignacion.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.txtIdAsignacion.DataBindings.Add(new System.Windows.Forms.Binding("Text", this.cdsAsignaCuenta, "AsignaCuenta.idAsignacion"));
			this.txtIdAsignacion.Location = new System.Drawing.Point(448, 414);
			this.txtIdAsignacion.Name = "txtIdAsignacion";
			this.txtIdAsignacion.TabIndex = 41;
			this.txtIdAsignacion.Text = "idAsignacion";
			// 
			// btRetencionFuente
			// 
			this.btRetencionFuente.Location = new System.Drawing.Point(136, 48);
			this.btRetencionFuente.Name = "btRetencionFuente";
			this.btRetencionFuente.Size = new System.Drawing.Size(120, 25);
			this.btRetencionFuente.TabIndex = 42;
			this.btRetencionFuente.Text = "Retención en &Fuente";
			this.btRetencionFuente.Click += new System.EventHandler(this.btRetencionFuente_Click);
			// 
			// btRetIva
			// 
			this.btRetIva.Location = new System.Drawing.Point(312, 48);
			this.btRetIva.Name = "btRetIva";
			this.btRetIva.Size = new System.Drawing.Size(120, 25);
			this.btRetIva.TabIndex = 43;
			this.btRetIva.Text = "&Retención en IVA";
			this.btRetIva.Click += new System.EventHandler(this.btRetIva_Click);
			// 
			// btGrupoCl
			// 
			this.btGrupoCl.Location = new System.Drawing.Point(328, 72);
			this.btGrupoCl.Name = "btGrupoCl";
			this.btGrupoCl.Size = new System.Drawing.Size(192, 25);
			this.btGrupoCl.TabIndex = 44;
			this.btGrupoCl.Text = "Mostrar &Grupos de Clientes";
			this.btGrupoCl.Click += new System.EventHandler(this.btGrupoCl_Click);
			// 
			// btGrupoItems
			// 
			this.btGrupoItems.Location = new System.Drawing.Point(376, 136);
			this.btGrupoItems.Name = "btGrupoItems";
			this.btGrupoItems.Size = new System.Drawing.Size(184, 25);
			this.btGrupoItems.TabIndex = 46;
			this.btGrupoItems.Text = "Mostrar Grupos de &Artículos";
			this.btGrupoItems.Click += new System.EventHandler(this.btGrupoItems_Click);
			// 
			// cmbCuenta
			// 
			this.cmbCuenta.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbCuenta.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbCuenta.DataSource = this.cdvCuentaTodo;
			appearance1.BackColor = System.Drawing.Color.White;
			this.cmbCuenta.DisplayLayout.Appearance = appearance1;
			ultraGridColumn1.Header.VisiblePosition = 1;
			ultraGridColumn1.Width = 133;
			ultraGridColumn2.Header.VisiblePosition = 4;
			ultraGridColumn2.Hidden = true;
			ultraGridColumn3.Header.VisiblePosition = 3;
			ultraGridColumn3.Hidden = true;
			ultraGridColumn4.Header.Caption = "Codigo";
			ultraGridColumn4.Header.VisiblePosition = 0;
			ultraGridColumn4.Hidden = true;
			ultraGridColumn5.Header.Caption = "Descripción";
			ultraGridColumn5.Header.VisiblePosition = 2;
			ultraGridColumn5.Width = 319;
			ultraGridColumn6.Header.VisiblePosition = 5;
			ultraGridBand1.Columns.AddRange(new object[] {
																										 ultraGridColumn1,
																										 ultraGridColumn2,
																										 ultraGridColumn3,
																										 ultraGridColumn4,
																										 ultraGridColumn5,
																										 ultraGridColumn6});
			this.cmbCuenta.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			appearance2.BackColor = System.Drawing.Color.Transparent;
			this.cmbCuenta.DisplayLayout.Override.CardAreaAppearance = appearance2;
			appearance3.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance3.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance3.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance3.FontData.BoldAsString = "True";
			appearance3.FontData.Name = "Arial";
			appearance3.FontData.SizeInPoints = 10F;
			appearance3.ForeColor = System.Drawing.Color.White;
			appearance3.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.cmbCuenta.DisplayLayout.Override.HeaderAppearance = appearance3;
			appearance4.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance4.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance4.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.cmbCuenta.DisplayLayout.Override.RowSelectorAppearance = appearance4;
			appearance5.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance5.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance5.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.cmbCuenta.DisplayLayout.Override.SelectedRowAppearance = appearance5;
			this.cmbCuenta.DisplayMember = "Codigo";
			this.cmbCuenta.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.OfficeXP;
			this.cmbCuenta.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbCuenta.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbCuenta.Location = new System.Drawing.Point(64, 32);
			this.cmbCuenta.Name = "cmbCuenta";
			this.cmbCuenta.Size = new System.Drawing.Size(192, 21);
			this.cmbCuenta.TabIndex = 48;
			this.cmbCuenta.ValueMember = "idCuenta";
			// 
			// cdvCuentaTodo
			// 
			this.cdvCuentaTodo.BindingContextCtrl = this;
			this.cdvCuentaTodo.DataSet = this.cdsCuenta;
			this.cdvCuentaTodo.Sort = "Codigo";
			this.cdvCuentaTodo.TableName = "";
			this.cdvCuentaTodo.TableViewName = "Cuenta";
			// 
			// label2
			// 
			this.label2.AutoSize = true;
			this.label2.Location = new System.Drawing.Point(-72, 16);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(39, 16);
			this.label2.TabIndex = 49;
			this.label2.Text = "Cuenta";
			// 
			// btCreaGrupoClient
			// 
			this.btCreaGrupoClient.Location = new System.Drawing.Point(328, 32);
			this.btCreaGrupoClient.Name = "btCreaGrupoClient";
			this.btCreaGrupoClient.Size = new System.Drawing.Size(192, 25);
			this.btCreaGrupoClient.TabIndex = 50;
			this.btCreaGrupoClient.Text = "Creación de Grupos de Clientes";
			this.btCreaGrupoClient.Click += new System.EventHandler(this.btCreaGrupoClient_Click);
			// 
			// panel1
			// 
			this.panel1.Controls.Add(this.cmbEgreso);
			this.panel1.Controls.Add(this.cmbDescuento);
			this.panel1.Controls.Add(this.cmbCosto);
			this.panel1.Controls.Add(this.cmbCuentaVenta);
			this.panel1.Controls.Add(this.cmbCuentaCompra);
			this.panel1.Controls.Add(this.btIntercambio);
			this.panel1.Controls.Add(this.chkCrearItem);
			this.panel1.Controls.Add(this.btGrupo);
			this.panel1.Controls.Add(this.label8);
			this.panel1.Controls.Add(this.label7);
			this.panel1.Controls.Add(this.lblCosto);
			this.panel1.Controls.Add(this.lblVenta);
			this.panel1.Controls.Add(this.label3);
			this.panel1.Controls.Add(this.label4);
			this.panel1.Controls.Add(this.label1);
			this.panel1.Controls.Add(this.btCreaGrupoItem);
			this.panel1.Controls.Add(this.cmbTipo);
			this.panel1.Controls.Add(this.cmbArticuloSRI);
			this.panel1.Controls.Add(this.btGrupoItems);
			this.panel1.Location = new System.Drawing.Point(16, 16);
			this.panel1.Name = "panel1";
			this.panel1.Size = new System.Drawing.Size(616, 232);
			this.panel1.TabIndex = 58;
			// 
			// cmbEgreso
			// 
			this.cmbEgreso.Location = new System.Drawing.Point(88, 160);
			this.cmbEgreso.Name = "cmbEgreso";
			// 
			// cmbEgreso.Properties
			// 
			this.cmbEgreso.Properties.Buttons.AddRange(new DevExpress.XtraEditors.Controls.EditorButton[] {
																																																			new DevExpress.XtraEditors.Controls.EditorButton(DevExpress.XtraEditors.Controls.ButtonPredefines.Combo)});
			this.cmbEgreso.Properties.Columns.AddRange(new DevExpress.XtraEditors.Controls.LookUpColumnInfo[] {
																																																					new DevExpress.XtraEditors.Controls.LookUpColumnInfo("Codigo", "Codigo", 60, DevExpress.Utils.FormatType.None, "", true, DevExpress.Utils.HorzAlignment.Near, DevExpress.Data.ColumnSortOrder.Ascending),
																																																					new DevExpress.XtraEditors.Controls.LookUpColumnInfo("Descripcion", "Descripcion", 150, DevExpress.Utils.FormatType.None, "", true, DevExpress.Utils.HorzAlignment.Near)});
			this.cmbEgreso.Properties.DataSource = this.cdvCuentaTodo;
			this.cmbEgreso.Properties.DisplayMember = "Codigo";
			this.cmbEgreso.Properties.NullText = "";
			this.cmbEgreso.Properties.PopupWidth = 400;
			this.cmbEgreso.Properties.ValueMember = "idCuenta";
			this.cmbEgreso.Size = new System.Drawing.Size(184, 20);
			this.cmbEgreso.TabIndex = 81;
			// 
			// cmbDescuento
			// 
			this.cmbDescuento.Location = new System.Drawing.Point(88, 136);
			this.cmbDescuento.Name = "cmbDescuento";
			// 
			// cmbDescuento.Properties
			// 
			this.cmbDescuento.Properties.Buttons.AddRange(new DevExpress.XtraEditors.Controls.EditorButton[] {
																																																				 new DevExpress.XtraEditors.Controls.EditorButton(DevExpress.XtraEditors.Controls.ButtonPredefines.Combo)});
			this.cmbDescuento.Properties.Columns.AddRange(new DevExpress.XtraEditors.Controls.LookUpColumnInfo[] {
																																																						 new DevExpress.XtraEditors.Controls.LookUpColumnInfo("Codigo", "Codigo", 60, DevExpress.Utils.FormatType.None, "", true, DevExpress.Utils.HorzAlignment.Near, DevExpress.Data.ColumnSortOrder.Ascending),
																																																						 new DevExpress.XtraEditors.Controls.LookUpColumnInfo("Descripcion", "Descripcion", 150, DevExpress.Utils.FormatType.None, "", true, DevExpress.Utils.HorzAlignment.Near)});
			this.cmbDescuento.Properties.DataSource = this.cdvCuentaTodo;
			this.cmbDescuento.Properties.DisplayMember = "Codigo";
			this.cmbDescuento.Properties.NullText = "";
			this.cmbDescuento.Properties.PopupWidth = 400;
			this.cmbDescuento.Properties.ValueMember = "idCuenta";
			this.cmbDescuento.Size = new System.Drawing.Size(184, 20);
			this.cmbDescuento.TabIndex = 80;
			// 
			// cmbCosto
			// 
			this.cmbCosto.Location = new System.Drawing.Point(88, 112);
			this.cmbCosto.Name = "cmbCosto";
			// 
			// cmbCosto.Properties
			// 
			this.cmbCosto.Properties.Buttons.AddRange(new DevExpress.XtraEditors.Controls.EditorButton[] {
																																																		 new DevExpress.XtraEditors.Controls.EditorButton(DevExpress.XtraEditors.Controls.ButtonPredefines.Combo)});
			this.cmbCosto.Properties.Columns.AddRange(new DevExpress.XtraEditors.Controls.LookUpColumnInfo[] {
																																																				 new DevExpress.XtraEditors.Controls.LookUpColumnInfo("Codigo", "Codigo", 60, DevExpress.Utils.FormatType.None, "", true, DevExpress.Utils.HorzAlignment.Near, DevExpress.Data.ColumnSortOrder.Ascending),
																																																				 new DevExpress.XtraEditors.Controls.LookUpColumnInfo("Descripcion", "Descripcion", 150, DevExpress.Utils.FormatType.None, "", true, DevExpress.Utils.HorzAlignment.Near)});
			this.cmbCosto.Properties.DataSource = this.cdvCuentaTodo;
			this.cmbCosto.Properties.DisplayMember = "Codigo";
			this.cmbCosto.Properties.NullText = "";
			this.cmbCosto.Properties.PopupWidth = 400;
			this.cmbCosto.Properties.ValueMember = "idCuenta";
			this.cmbCosto.Size = new System.Drawing.Size(184, 20);
			this.cmbCosto.TabIndex = 79;
			// 
			// cmbCuentaVenta
			// 
			this.cmbCuentaVenta.Location = new System.Drawing.Point(88, 88);
			this.cmbCuentaVenta.Name = "cmbCuentaVenta";
			// 
			// cmbCuentaVenta.Properties
			// 
			this.cmbCuentaVenta.Properties.Buttons.AddRange(new DevExpress.XtraEditors.Controls.EditorButton[] {
																																																					 new DevExpress.XtraEditors.Controls.EditorButton(DevExpress.XtraEditors.Controls.ButtonPredefines.Combo)});
			this.cmbCuentaVenta.Properties.Columns.AddRange(new DevExpress.XtraEditors.Controls.LookUpColumnInfo[] {
																																																							 new DevExpress.XtraEditors.Controls.LookUpColumnInfo("Codigo", "Codigo", 60, DevExpress.Utils.FormatType.None, "", true, DevExpress.Utils.HorzAlignment.Near, DevExpress.Data.ColumnSortOrder.Ascending),
																																																							 new DevExpress.XtraEditors.Controls.LookUpColumnInfo("Descripcion", "Descripcion", 150, DevExpress.Utils.FormatType.None, "", true, DevExpress.Utils.HorzAlignment.Near)});
			this.cmbCuentaVenta.Properties.DataSource = this.cdvCuentaTodo;
			this.cmbCuentaVenta.Properties.DisplayMember = "Codigo";
			this.cmbCuentaVenta.Properties.NullText = "";
			this.cmbCuentaVenta.Properties.PopupWidth = 400;
			this.cmbCuentaVenta.Properties.ValueMember = "idCuenta";
			this.cmbCuentaVenta.Size = new System.Drawing.Size(184, 20);
			this.cmbCuentaVenta.TabIndex = 78;
			// 
			// cmbCuentaCompra
			// 
			this.cmbCuentaCompra.Location = new System.Drawing.Point(88, 64);
			this.cmbCuentaCompra.Name = "cmbCuentaCompra";
			// 
			// cmbCuentaCompra.Properties
			// 
			this.cmbCuentaCompra.Properties.Buttons.AddRange(new DevExpress.XtraEditors.Controls.EditorButton[] {
																																																						new DevExpress.XtraEditors.Controls.EditorButton(DevExpress.XtraEditors.Controls.ButtonPredefines.Combo)});
			this.cmbCuentaCompra.Properties.Columns.AddRange(new DevExpress.XtraEditors.Controls.LookUpColumnInfo[] {
																																																								new DevExpress.XtraEditors.Controls.LookUpColumnInfo("Codigo", "Codigo", 60, DevExpress.Utils.FormatType.None, "", true, DevExpress.Utils.HorzAlignment.Near, DevExpress.Data.ColumnSortOrder.Ascending),
																																																								new DevExpress.XtraEditors.Controls.LookUpColumnInfo("Descripcion", "Descripcion", 150, DevExpress.Utils.FormatType.None, "", true, DevExpress.Utils.HorzAlignment.Near)});
			this.cmbCuentaCompra.Properties.DataSource = this.cdvCuentaTodo;
			this.cmbCuentaCompra.Properties.DisplayMember = "Codigo";
			this.cmbCuentaCompra.Properties.NullText = "";
			this.cmbCuentaCompra.Properties.PopupWidth = 400;
			this.cmbCuentaCompra.Properties.ValueMember = "idCuenta";
			this.cmbCuentaCompra.Size = new System.Drawing.Size(184, 20);
			this.cmbCuentaCompra.TabIndex = 77;
			// 
			// btIntercambio
			// 
			this.btIntercambio.Location = new System.Drawing.Point(376, 184);
			this.btIntercambio.Name = "btIntercambio";
			this.btIntercambio.Size = new System.Drawing.Size(184, 23);
			this.btIntercambio.TabIndex = 76;
			this.btIntercambio.Text = "Ordenar Listas por Descripción";
			this.btIntercambio.Click += new System.EventHandler(this.btIntercambio_Click);
			// 
			// chkCrearItem
			// 
			this.chkCrearItem.Location = new System.Drawing.Point(88, 200);
			this.chkCrearItem.Name = "chkCrearItem";
			// 
			// chkCrearItem.Properties
			// 
			this.chkCrearItem.Properties.Caption = "Crear Artículo";
			this.chkCrearItem.Size = new System.Drawing.Size(176, 18);
			this.chkCrearItem.TabIndex = 75;
			// 
			// btGrupo
			// 
			this.btGrupo.Location = new System.Drawing.Point(376, 104);
			this.btGrupo.Name = "btGrupo";
			this.btGrupo.Size = new System.Drawing.Size(184, 25);
			this.btGrupo.TabIndex = 69;
			this.btGrupo.Text = "Creación de Grupo Unico";
			this.btGrupo.Click += new System.EventHandler(this.btGrupo_Click);
			// 
			// label8
			// 
			this.label8.AutoSize = true;
			this.label8.Location = new System.Drawing.Point(24, 160);
			this.label8.Name = "label8";
			this.label8.Size = new System.Drawing.Size(38, 16);
			this.label8.TabIndex = 68;
			this.label8.Text = "Egreso";
			// 
			// label7
			// 
			this.label7.AutoSize = true;
			this.label7.Location = new System.Drawing.Point(24, 136);
			this.label7.Name = "label7";
			this.label7.Size = new System.Drawing.Size(57, 16);
			this.label7.TabIndex = 66;
			this.label7.Text = "Descuento";
			// 
			// lblCosto
			// 
			this.lblCosto.AutoSize = true;
			this.lblCosto.Location = new System.Drawing.Point(24, 112);
			this.lblCosto.Name = "lblCosto";
			this.lblCosto.Size = new System.Drawing.Size(32, 16);
			this.lblCosto.TabIndex = 64;
			this.lblCosto.Text = "Costo";
			// 
			// lblVenta
			// 
			this.lblVenta.AutoSize = true;
			this.lblVenta.Location = new System.Drawing.Point(24, 88);
			this.lblVenta.Name = "lblVenta";
			this.lblVenta.Size = new System.Drawing.Size(33, 16);
			this.lblVenta.TabIndex = 62;
			this.lblVenta.Text = "Venta";
			// 
			// label3
			// 
			this.label3.AutoSize = true;
			this.label3.Location = new System.Drawing.Point(24, 16);
			this.label3.Name = "label3";
			this.label3.Size = new System.Drawing.Size(26, 16);
			this.label3.TabIndex = 53;
			this.label3.Text = "Tipo";
			// 
			// label4
			// 
			this.label4.AutoSize = true;
			this.label4.Location = new System.Drawing.Point(336, 17);
			this.label4.Name = "label4";
			this.label4.Size = new System.Drawing.Size(22, 16);
			this.label4.TabIndex = 55;
			this.label4.Text = "SRI";
			// 
			// label1
			// 
			this.label1.AutoSize = true;
			this.label1.Location = new System.Drawing.Point(24, 64);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(43, 16);
			this.label1.TabIndex = 57;
			this.label1.Text = "Compra";
			// 
			// btCreaGrupoItem
			// 
			this.btCreaGrupoItem.Location = new System.Drawing.Point(376, 72);
			this.btCreaGrupoItem.Name = "btCreaGrupoItem";
			this.btCreaGrupoItem.Size = new System.Drawing.Size(184, 25);
			this.btCreaGrupoItem.TabIndex = 51;
			this.btCreaGrupoItem.Text = "Creación de Grupos de Artículos";
			this.btCreaGrupoItem.Click += new System.EventHandler(this.btCreaGrupoItem_Click);
			// 
			// cmbTipo
			// 
			appearance6.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbTipo.Appearance = appearance6;
			this.cmbTipo.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbTipo.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbTipo.DisplayLayout.AutoFitColumns = true;
			ultraGridColumn7.Header.VisiblePosition = 1;
			ultraGridColumn7.Hidden = true;
			ultraGridColumn7.Width = 254;
			ultraGridColumn8.Header.VisiblePosition = 0;
			ultraGridColumn8.Width = 165;
			ultraGridBand2.Columns.AddRange(new object[] {
																										 ultraGridColumn7,
																										 ultraGridColumn8});
			this.cmbTipo.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			this.cmbTipo.DisplayMember = "Tipo";
			this.cmbTipo.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbTipo.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbTipo.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbTipo.Location = new System.Drawing.Point(88, 17);
			this.cmbTipo.Name = "cmbTipo";
			this.cmbTipo.Size = new System.Drawing.Size(184, 21);
			this.cmbTipo.TabIndex = 59;
			this.cmbTipo.ValueMember = "idTipoGrupo";
			this.cmbTipo.ValueChanged += new System.EventHandler(this.cmbTipo_ValueChanged);
			// 
			// cmbArticuloSRI
			// 
			appearance7.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbArticuloSRI.Appearance = appearance7;
			this.cmbArticuloSRI.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbArticuloSRI.Cursor = System.Windows.Forms.Cursors.Default;
			ultraGridColumn9.Header.VisiblePosition = 0;
			ultraGridColumn9.Hidden = true;
			ultraGridColumn10.Header.VisiblePosition = 1;
			ultraGridColumn10.Width = 396;
			ultraGridColumn11.Header.VisiblePosition = 2;
			ultraGridColumn11.Width = 69;
			ultraGridColumn12.Header.VisiblePosition = 3;
			ultraGridColumn12.Width = 68;
			ultraGridBand3.Columns.AddRange(new object[] {
																										 ultraGridColumn9,
																										 ultraGridColumn10,
																										 ultraGridColumn11,
																										 ultraGridColumn12});
			this.cmbArticuloSRI.DisplayLayout.BandsSerializer.Add(ultraGridBand3);
			this.cmbArticuloSRI.DisplayMember = "Detalle";
			this.cmbArticuloSRI.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbArticuloSRI.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbArticuloSRI.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbArticuloSRI.Location = new System.Drawing.Point(376, 17);
			this.cmbArticuloSRI.Name = "cmbArticuloSRI";
			this.cmbArticuloSRI.Size = new System.Drawing.Size(184, 21);
			this.cmbArticuloSRI.TabIndex = 60;
			this.cmbArticuloSRI.ValueMember = "idArticuloSRI";
			// 
			// errorProvider1
			// 
			this.errorProvider1.ContainerControl = this;
			// 
			// xtraTabControl1
			// 
			this.xtraTabControl1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.xtraTabControl1.Location = new System.Drawing.Point(16, 16);
			this.xtraTabControl1.Name = "xtraTabControl1";
			this.xtraTabControl1.SelectedTabPage = this.xtraTabPage1;
			this.xtraTabControl1.Size = new System.Drawing.Size(656, 376);
			this.xtraTabControl1.TabIndex = 57;
			this.xtraTabControl1.TabPages.AddRange(new DevExpress.XtraTab.XtraTabPage[] {
																																										this.xtraTabPage1,
																																										this.xtraTabPage2,
																																										this.xtraTabPage3,
																																										this.xtraTabPage4});
			this.xtraTabControl1.Text = "xtraTabControl1";
			// 
			// xtraTabPage1
			// 
			this.xtraTabPage1.Controls.Add(this.tlsCuentas);
			this.xtraTabPage1.Name = "xtraTabPage1";
			this.xtraTabPage1.Size = new System.Drawing.Size(647, 346);
			this.xtraTabPage1.Text = "Asignaciones Generales";
			// 
			// xtraTabPage2
			// 
			this.xtraTabPage2.Controls.Add(this.groupControl2);
			this.xtraTabPage2.Controls.Add(this.groupControl1);
			this.xtraTabPage2.Name = "xtraTabPage2";
			this.xtraTabPage2.Size = new System.Drawing.Size(647, 346);
			this.xtraTabPage2.Text = "Cobros y Pagos";
			// 
			// groupControl2
			// 
			this.groupControl2.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.groupControl2.Controls.Add(this.cmbTipoFP);
			this.groupControl2.Controls.Add(this.cmbCuentaFP);
			this.groupControl2.Controls.Add(this.ultraGrid1);
			this.groupControl2.Location = new System.Drawing.Point(8, 136);
			this.groupControl2.Name = "groupControl2";
			this.groupControl2.Size = new System.Drawing.Size(632, 200);
			this.groupControl2.TabIndex = 52;
			this.groupControl2.Text = "Contado";
			// 
			// cmbTipoFP
			// 
			this.cmbTipoFP.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbTipoFP.DataMember = "Band 0";
			this.cmbTipoFP.DataSource = this.udsTipo;
			ultraGridBand4.AddButtonCaption = "ClientePrecio";
			ultraGridColumn13.Header.VisiblePosition = 0;
			ultraGridColumn13.Hidden = true;
			ultraGridColumn14.Header.VisiblePosition = 1;
			ultraGridColumn14.Width = 143;
			ultraGridBand4.Columns.AddRange(new object[] {
																										 ultraGridColumn13,
																										 ultraGridColumn14});
			this.cmbTipoFP.DisplayLayout.BandsSerializer.Add(ultraGridBand4);
			this.cmbTipoFP.DisplayMember = "Tipo";
			this.cmbTipoFP.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbTipoFP.Location = new System.Drawing.Point(368, 73);
			this.cmbTipoFP.Name = "cmbTipoFP";
			this.cmbTipoFP.Size = new System.Drawing.Size(168, 86);
			this.cmbTipoFP.TabIndex = 10;
			this.cmbTipoFP.Text = "Tipo";
			this.cmbTipoFP.ValueMember = "idTipo";
			this.cmbTipoFP.Visible = false;
			// 
			// udsTipo
			// 
			ultraDataColumn1.DataType = typeof(int);
			ultraDataColumn1.DefaultValue = 0;
			ultraDataColumn2.DefaultValue = "";
			this.udsTipo.Band.Columns.AddRange(new object[] {
																												ultraDataColumn1,
																												ultraDataColumn2});
			this.udsTipo.Rows.AddRange(new object[] {
																								new Infragistics.Win.UltraWinDataSource.UltraDataRow(new object[] {
																																																										"idTipo",
																																																										1,
																																																										"Tipo",
																																																										"Compra"}, new object[0]),
																								new Infragistics.Win.UltraWinDataSource.UltraDataRow(new object[] {
																																																										"idTipo",
																																																										2,
																																																										"Tipo",
																																																										"Venta"}, new object[0]),
																								new Infragistics.Win.UltraWinDataSource.UltraDataRow(new object[] {
																																																										"idTipo",
																																																										3,
																																																										"Tipo",
																																																										"Todo"}, new object[0])});
			// 
			// cmbCuentaFP
			// 
			this.cmbCuentaFP.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbCuentaFP.DataSource = this.cdvCuenta;
			ultraGridColumn15.Header.VisiblePosition = 0;
			ultraGridColumn16.Header.VisiblePosition = 2;
			ultraGridColumn16.Hidden = true;
			ultraGridColumn17.Header.VisiblePosition = 3;
			ultraGridColumn17.Hidden = true;
			ultraGridColumn18.Header.VisiblePosition = 4;
			ultraGridColumn18.Hidden = true;
			ultraGridColumn19.Header.VisiblePosition = 1;
			ultraGridColumn19.Width = 161;
			ultraGridColumn20.Header.VisiblePosition = 5;
			ultraGridBand5.Columns.AddRange(new object[] {
																										 ultraGridColumn15,
																										 ultraGridColumn16,
																										 ultraGridColumn17,
																										 ultraGridColumn18,
																										 ultraGridColumn19,
																										 ultraGridColumn20});
			this.cmbCuentaFP.DisplayLayout.BandsSerializer.Add(ultraGridBand5);
			this.cmbCuentaFP.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.TemplateOnBottom;
			this.cmbCuentaFP.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.True;
			this.cmbCuentaFP.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			this.cmbCuentaFP.DisplayMember = "Codigo";
			this.cmbCuentaFP.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbCuentaFP.Location = new System.Drawing.Point(80, 72);
			this.cmbCuentaFP.Name = "cmbCuentaFP";
			this.cmbCuentaFP.Size = new System.Drawing.Size(280, 87);
			this.cmbCuentaFP.TabIndex = 8;
			this.cmbCuentaFP.Text = "ultraDropDown2";
			this.cmbCuentaFP.ValueMember = "idCuenta";
			this.cmbCuentaFP.Visible = false;
			// 
			// ultraGrid1
			// 
			this.ultraGrid1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.ultraGrid1.Cursor = System.Windows.Forms.Cursors.Default;
			this.ultraGrid1.DataMember = "CompraFormaPago";
			this.ultraGrid1.DataSource = this.cdsFormaPago;
			appearance8.BackColor = System.Drawing.Color.White;
			this.ultraGrid1.DisplayLayout.Appearance = appearance8;
			this.ultraGrid1.DisplayLayout.AutoFitColumns = true;
			ultraGridColumn21.Header.VisiblePosition = 0;
			ultraGridColumn21.Hidden = true;
			ultraGridColumn21.Width = 68;
			ultraGridColumn22.Header.VisiblePosition = 1;
			ultraGridColumn22.Hidden = true;
			ultraGridColumn22.Width = 81;
			ultraGridColumn23.Header.VisiblePosition = 2;
			ultraGridColumn23.Width = 155;
			ultraGridColumn24.Header.Caption = "Cuenta";
			ultraGridColumn24.Header.VisiblePosition = 3;
			ultraGridColumn24.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownValidate;
			ultraGridColumn24.Width = 157;
			ultraGridColumn25.Header.VisiblePosition = 4;
			ultraGridColumn25.Hidden = true;
			ultraGridColumn25.Width = 90;
			ultraGridColumn26.Header.VisiblePosition = 5;
			ultraGridColumn26.Hidden = true;
			ultraGridColumn26.Width = 107;
			ultraGridColumn27.Header.VisiblePosition = 6;
			ultraGridColumn27.Hidden = true;
			ultraGridColumn27.Width = 135;
			ultraGridColumn28.Header.Caption = "Tipo";
			ultraGridColumn28.Header.VisiblePosition = 7;
			ultraGridColumn28.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownList;
			ultraGridColumn28.Width = 181;
			ultraGridColumn29.Header.VisiblePosition = 8;
			ultraGridColumn29.Width = 86;
			ultraGridBand6.Columns.AddRange(new object[] {
																										 ultraGridColumn21,
																										 ultraGridColumn22,
																										 ultraGridColumn23,
																										 ultraGridColumn24,
																										 ultraGridColumn25,
																										 ultraGridColumn26,
																										 ultraGridColumn27,
																										 ultraGridColumn28,
																										 ultraGridColumn29});
			this.ultraGrid1.DisplayLayout.BandsSerializer.Add(ultraGridBand6);
			this.ultraGrid1.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.TemplateOnBottom;
			appearance9.BackColor = System.Drawing.Color.Transparent;
			this.ultraGrid1.DisplayLayout.Override.CardAreaAppearance = appearance9;
			appearance10.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance10.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance10.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance10.FontData.BoldAsString = "True";
			appearance10.FontData.Name = "Arial";
			appearance10.FontData.SizeInPoints = 10F;
			appearance10.ForeColor = System.Drawing.Color.White;
			appearance10.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.ultraGrid1.DisplayLayout.Override.HeaderAppearance = appearance10;
			appearance11.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance11.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance11.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.RowSelectorAppearance = appearance11;
			appearance12.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance12.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance12.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.SelectedRowAppearance = appearance12;
			this.ultraGrid1.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.ultraGrid1.Location = new System.Drawing.Point(16, 32);
			this.ultraGrid1.Name = "ultraGrid1";
			this.ultraGrid1.Size = new System.Drawing.Size(600, 160);
			this.ultraGrid1.TabIndex = 7;
			// 
			// cdsFormaPago
			// 
			this.cdsFormaPago.BindingContextCtrl = this;
			this.cdsFormaPago.DataLibrary = "LibFacturacion";
			this.cdsFormaPago.DataLibraryUrl = "";
			this.cdsFormaPago.DataSetDef = "dsCompraTabla";
			this.cdsFormaPago.Locale = new System.Globalization.CultureInfo("es-MX");
			this.cdsFormaPago.Name = "cdsFormaPago";
			this.cdsFormaPago.SchemaClassName = "LibFacturacion.DataClass";
			this.cdsFormaPago.SchemaDef = null;
			this.cdsFormaPago.BeforeFill += new C1.Data.FillEventHandler(this.cdsFormaPago_BeforeFill);
			// 
			// groupControl1
			// 
			this.groupControl1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.groupControl1.Controls.Add(this.cmbCuenta);
			this.groupControl1.Controls.Add(this.label2);
			this.groupControl1.Controls.Add(this.btCreaGrupoClient);
			this.groupControl1.Controls.Add(this.btGrupoCl);
			this.groupControl1.Location = new System.Drawing.Point(8, 8);
			this.groupControl1.Name = "groupControl1";
			this.groupControl1.Size = new System.Drawing.Size(632, 120);
			this.groupControl1.TabIndex = 51;
			this.groupControl1.Text = "Crédito";
			// 
			// xtraTabPage3
			// 
			this.xtraTabPage3.Controls.Add(this.panel1);
			this.xtraTabPage3.Name = "xtraTabPage3";
			this.xtraTabPage3.Size = new System.Drawing.Size(647, 346);
			this.xtraTabPage3.Text = "Inventarios A.F. Servicios";
			// 
			// xtraTabPage4
			// 
			this.xtraTabPage4.Controls.Add(this.btRetencionFuente);
			this.xtraTabPage4.Controls.Add(this.btRetIva);
			this.xtraTabPage4.Name = "xtraTabPage4";
			this.xtraTabPage4.Size = new System.Drawing.Size(647, 346);
			this.xtraTabPage4.Text = "Retenciones";
			// 
			// AsignaCuentaArbol
			// 
			this.AcceptButton = this.btAceptar;
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 14);
			this.CancelButton = this.btCancelar;
			this.ClientSize = new System.Drawing.Size(688, 446);
			this.Controls.Add(this.xtraTabControl1);
			this.Controls.Add(this.txtIdAsignacion);
			this.Controls.Add(this.btNoAsignar);
			this.Controls.Add(this.btCancelar);
			this.Controls.Add(this.btAceptar);
			this.Name = "AsignaCuentaArbol";
			this.Text = "Asignación de Cuentas";
			this.Load += new System.EventHandler(this.AsignaCuentaArbol_Load);
			((System.ComponentModel.ISupportInitialize)(this.tlsCuentas)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.repositoryItemLookUpEdit1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvCuentaCod)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsCuenta)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.repositoryItemLookUpEdit2)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsAsignaCuenta)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvCuenta)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCuenta)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvCuentaTodo)).EndInit();
			this.panel1.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.cmbEgreso.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbDescuento.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCosto.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCuentaVenta.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCuentaCompra.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.chkCrearItem.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbTipo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbArticuloSRI)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.xtraTabControl1)).EndInit();
			this.xtraTabControl1.ResumeLayout(false);
			this.xtraTabPage1.ResumeLayout(false);
			this.xtraTabPage2.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.groupControl2)).EndInit();
			this.groupControl2.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.cmbTipoFP)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.udsTipo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCuentaFP)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsFormaPago)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.groupControl1)).EndInit();
			this.groupControl1.ResumeLayout(false);
			this.xtraTabPage3.ResumeLayout(false);
			this.xtraTabPage4.ResumeLayout(false);
			this.ResumeLayout(false);

		}
		#endregion


		private void btCancelar_Click(object sender, System.EventArgs e)
		{
			Close();
		}

		private void AsignaCuentaArbol_Load(object sender, System.EventArgs e)
		{
			txtIdAsignacion.Width = 0;
			DevExpress.XtraTreeList.Nodes.TreeListNode nd = this.tlsCuentas.FindNodeByKeyID(80);
			this.tlsCuentas.FocusedNode = nd;
			cmbTipo.DataSource = Funcion.dvProcedimiento(cdsAsignaCuenta, "Select idTipoGrupo, Tipo From ArticuloTipo");
			cmbArticuloSRI.DataSource = Funcion.dvProcedimiento(cdsAsignaCuenta, 
				"Select idArticuloSRI, Detalle, Codigo, Porcentaje From ArticuloSRI Order by Codigo");
		}

		private void btAceptar_Click(object sender, System.EventArgs e)
		{
			this.cdsAsignaCuenta.Update();
			cdsFormaPago.Update();
			this.Close();
		}

		private void btNoAsignar_Click(object sender, System.EventArgs e)
		{
			this.cdsAsignaCuenta.Update();
			string stUpdate = "Update AsignaCuenta Set idCuenta = 0 Where idAsignacion = " + txtIdAsignacion.Text;
			Funcion.EjecutaSQL(cdsAsignaCuenta, stUpdate);
			cdsAsignaCuenta.Clear();
			cdsAsignaCuenta.Fill();
		}

		private void cdsCuenta_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			this.cdsCuenta.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);
		}

		private void btRetencionFuente_Click(object sender, System.EventArgs e)
		{
			using(ArticuloSRI miArt = new ArticuloSRI())
			{
				miArt.ShowDialog();
			}
		}

		private void btGrupoCl_Click(object sender, System.EventArgs e)
		{
			using(ClienteGrupo miGrupoC = new ClienteGrupo(false))
			{
				miGrupoC.ShowDialog();
			}
		}

		private void btGrupoItems_Click(object sender, System.EventArgs e)
		{
			using(ArtGrupo miGrupo = new ArtGrupo())
			{
				miGrupo.ShowDialog();
			}
		}

		private void btRetIva_Click(object sender, System.EventArgs e)
		{
			using(RetIVA miRetIva = new RetIVA())
			{
				miRetIva.ShowDialog();
			}
		}

		private void btCreaGrupoClient_Click(object sender, System.EventArgs e)
		{
			if (cmbCuenta.Text.Trim().Length == 0)
			{
				errorProvider1.SetError(cmbCuenta, "Seleccione una cuenta cuyas subcuentas desea que sean grupos de clientes");
				return;
			}

			string stCuenta = cmbCuenta.Text.Trim();
			if (DialogResult.No == MessageBox.Show("Desea crear grupos de clientes con las subcuentas de " +stCuenta, 
				"Confirmacion", MessageBoxButtons.YesNo, MessageBoxIcon.Question)) return;
			string stSelect = string.Format("Exec AsignacionCliente '{0}'", stCuenta);
			string stMensaje = Funcion.sEscalarSQL(cdsAsignaCuenta, stSelect);
			MessageBox.Show(stMensaje, "Información");
		}

		private void btCreaGrupoItem_Click(object sender, System.EventArgs e)
		{
			errorProvider1.SetError(cmbCuentaCompra, "");
			errorProvider1.SetError(cmbTipo, "");
			errorProvider1.SetError(cmbArticuloSRI, "");
			bool bOk = true;
			if (cmbCuentaCompra.Text.Trim().Length == 0)
			{
				errorProvider1.SetError(cmbCuentaCompra, "Seleccione una cuenta cuyas sucuentas desea que sean grupos de Artículos");
				bOk = false;
			}
			if (cmbCuentaVenta.Text.Trim().Length == 0)
			{
				errorProvider1.SetError(cmbCuentaVenta, "Seleccione una cuenta cuyas sucuentas desea que sean grupos de Artículos");
				bOk = false;
			}
			if (cmbCuentaCompra.Text.Trim().Length == 0)
			{
				errorProvider1.SetError(cmbCuentaCompra, "Seleccione una cuenta cuyas sucuentas desea que sean grupos de Artículos");
				bOk = false;
			}
			if (cmbCuentaCompra.Text.Trim().Length == 0)
			{
				errorProvider1.SetError(cmbCuentaCompra, "Seleccione una cuenta cuyas sucuentas desea que sean grupos de Artículos");
				bOk = false;
			}
			if (cmbTipo.Text.Trim().Length == 0)
			{
				errorProvider1.SetError(cmbTipo, "Seleccione un tipo de Artículo");
				bOk = false;
			}
			if (cmbArticuloSRI.Text.Trim().Length == 0)
			{
				errorProvider1.SetError(cmbArticuloSRI, "Seleccione el código de retención que aplique");
				bOk = false;
			}
			if (!bOk) return;
			
			string stCuenta = cmbCuentaCompra.Text.Trim();
			int iTipo = (int)cmbTipo.Value;
			int idArtSRI = (int)cmbArticuloSRI.Value;
			int iCreaItem = 0;
			if (chkCrearItem.Checked) iCreaItem = 1;

			if (DialogResult.No == MessageBox.Show("Desea crear Grupos de Artículos con las sucuentas de " + stCuenta, 
				"Confirmacion", MessageBoxButtons.YesNo, MessageBoxIcon.Question)) return;
			string stSelect = string.Format("Exec AsignacionArticulo '{0}', {1}, {2}, {3}", 
				stCuenta, iTipo, idArtSRI, iCreaItem);
			string stMensaje = Funcion.sEscalarSQL(cdsAsignaCuenta, stSelect);
			MessageBox.Show(stMensaje, "Información");
		}

		private void btGrupo_Click(object sender, System.EventArgs e)
		{
			#region Limpia Errores
			errorProvider1.SetError(cmbCuentaCompra, "");
			errorProvider1.SetError(cmbCuentaVenta, "");
			errorProvider1.SetError(cmbDescuento, "");
			errorProvider1.SetError(cmbCosto, "");
			errorProvider1.SetError(cmbEgreso, "");
			errorProvider1.SetError(cmbTipo, "");
			errorProvider1.SetError(cmbArticuloSRI, "");
			bool bOk = true;
			#endregion Limpia Errores
			#region Verifica Combos
			if (cmbCuentaCompra.Text.Trim().EndsWith("."))
			{
				errorProvider1.SetError(cmbCuentaCompra, "Seleccione una cuenta de movimiento");
				bOk = false;
			}
			if (cmbCuentaVenta.Text.Trim().EndsWith("."))
			{
				errorProvider1.SetError(cmbCuentaVenta, "Seleccione una cuenta de movimiento");
				bOk = false;
			}
			if (cmbDescuento.Text.Trim().EndsWith("."))
			{
				errorProvider1.SetError(cmbDescuento, "Seleccione una cuenta de movimiento");
				bOk = false;
			}
			if (cmbEgreso.Text.Trim().EndsWith("."))
			{
				errorProvider1.SetError(cmbEgreso, "Seleccione una cuenta de movimiento");
				bOk = false;
			}			
			
			if (cmbCosto.Text.Trim().EndsWith("."))
			{
				errorProvider1.SetError(cmbCosto, "Seleccione una cuenta de movimiento");
				bOk = false;
			}
			if (cmbTipo.Text.Trim().Length == 0)
			{
				errorProvider1.SetError(cmbTipo, "Seleccione un tipo de Artículo");
				bOk = false;
			}
			if (cmbArticuloSRI.Text.Trim().Length == 0)
			{
				errorProvider1.SetError(cmbArticuloSRI, "Seleccione el código de retención que aplique");
				bOk = false;
			}
			if (!bOk) return;
			#endregion Verifica Combos			
			#region Variables
			int iCuenta2 = 0;
			int iCuenta3 = 0;
			int iCuenta4 = 0;
			int iCuenta5 = 0;
			int iCuenta6 = 0;

			if (cmbCuentaCompra.EditValue != null && cmbCuentaCompra.Text.Length > 0) 
				iCuenta2 = (int) cmbCuentaCompra.EditValue;
			if (cmbCuentaVenta.EditValue != null && cmbCuentaVenta.Text.Length > 0) 
				iCuenta3 = (int) cmbCuentaVenta.EditValue;
			if (cmbCosto.EditValue != null && cmbCosto.Text.Length > 0) 
				iCuenta4 = (int) cmbCosto.EditValue;
			if (cmbDescuento.EditValue != null && cmbDescuento.Text.Length > 0) 
				iCuenta5 = (int) cmbDescuento.EditValue;
			if (cmbEgreso.EditValue != null && cmbEgreso.Text.Length > 0) 
				iCuenta6 = (int) cmbEgreso.EditValue;
			int iTipo = (int)cmbTipo.Value;
			int idArtSRI = (int)cmbArticuloSRI.Value;
			int iCreaItem = 0;
			if (chkCrearItem.Checked) iCreaItem = 1;
			#endregion Variables

			if (DialogResult.No == MessageBox.Show("Desea crear Grupo de Artículos con cuenta seleccionada?", 
				"Confirmacion", MessageBoxButtons.YesNo, MessageBoxIcon.Question)) return;
			string stSelect = string.Format("Exec AsignacionArticuloCuentas {0}, {1}, {2}, {3}, {4}, {5}, {6}, {7}", 
				iCuenta2, iCuenta3, iCuenta4, iCuenta5, iCuenta6, iTipo, idArtSRI, iCreaItem);
			string stMensaje = Funcion.sEscalarSQL(cdsAsignaCuenta, stSelect);
			MessageBox.Show(stMensaje, "Información");
		}

		private void cmbCuentaCompra_MouseUp(object sender, System.Windows.Forms.MouseEventArgs e)
		{
			if (e.Button != MouseButtons.Right) return;
			cmbCuentaCompra.EditValue = 0;
		}

		private void cmbCuentaVenta_MouseUp(object sender, System.Windows.Forms.MouseEventArgs e)
		{
			if (e.Button != MouseButtons.Right) return;
			cmbCuentaVenta.EditValue = 0;
		}

		private void cmbCosto_MouseUp(object sender, System.Windows.Forms.MouseEventArgs e)
		{
			if (e.Button != MouseButtons.Right) return;
			cmbCosto.EditValue = 0;
		}

		private void cmbDescuento_MouseUp(object sender, System.Windows.Forms.MouseEventArgs e)
		{
			if (e.Button != MouseButtons.Right) return;
			cmbDescuento.EditValue = 0;
		}

		private void cmbEgreso_MouseUp(object sender, System.Windows.Forms.MouseEventArgs e)
		{
			if (e.Button != MouseButtons.Right) return;
			cmbEgreso.EditValue = 0;
		}

		private void cmbTipo_ValueChanged(object sender, System.EventArgs e)
		{
			if (cmbTipo.Value == null) return;
			if ((int) cmbTipo.Value == 5)
			{
				lblCosto.Text = "Depreciac.";
				lblVenta.Text = "Dep. Acum.";
			}
			else
			{
				lblCosto.Text = "Costo";
				lblVenta.Text = "Venta";
			}
			
			cmbCuentaCompra.Enabled = true;
			cmbCosto.Enabled = true;
			cmbCuentaVenta.Enabled = true;
			cmbDescuento.Enabled = true;
			cmbEgreso.Enabled = true;

			switch ((int) cmbTipo.Value)
			{
				case 1:
					break;
				case 2: // Gasto
					cmbCosto.Enabled = false;
					cmbCuentaVenta.Enabled = false;
					cmbDescuento.Enabled = false;
					cmbEgreso.Enabled = false;
					break;
				case 3: // Servicio
					cmbCosto.Enabled = false;
					break;
				case 4: // Saldo Incial
					cmbCuentaCompra.Enabled = false;
					cmbCosto.Enabled = false;
					cmbCuentaVenta.Enabled = false;
					cmbDescuento.Enabled = false;
					cmbEgreso.Enabled = false;
					break;
			}
		}

		private void btIntercambio_Click(object sender, System.EventArgs e)
		{
			string stColumna = "";
			string stColumna1 = "";
			int iAncho = 60;
			int iAncho1 = 150;
			cmbCuentaCompra.Properties.Columns.Clear();
			cmbCuentaVenta.Properties.Columns.Clear();
			cmbCosto.Properties.Columns.Clear();
			cmbDescuento.Properties.Columns.Clear();
			cmbEgreso.Properties.Columns.Clear();
			if (cmbCuentaCompra.Properties.DisplayMember == "Codigo")
			{
				stColumna = "Descripcion";
				stColumna1 = "Codigo";
				iAncho = 150;
				iAncho1 = 60;
				btIntercambio.Text = "Ordenar Listas por Descripción";
			}
			else
			{
				stColumna = "Codigo";
				stColumna1 = "Descripcion";
				btIntercambio.Text = "Ordenar Listas por Código";
			}
			cmbCuentaCompra.Properties.Columns.AddRange(new DevExpress.XtraEditors.Controls.LookUpColumnInfo[] {
																																																					 new DevExpress.XtraEditors.Controls.LookUpColumnInfo(stColumna, stColumna, iAncho, DevExpress.Utils.FormatType.None, "", true, DevExpress.Utils.HorzAlignment.Near),
																																																					 new DevExpress.XtraEditors.Controls.LookUpColumnInfo(stColumna1, stColumna1, iAncho1, DevExpress.Utils.FormatType.None, "", true, DevExpress.Utils.HorzAlignment.Near)});
			cmbCuentaVenta.Properties.Columns.AddRange(new DevExpress.XtraEditors.Controls.LookUpColumnInfo[] {
																																																					new DevExpress.XtraEditors.Controls.LookUpColumnInfo(stColumna, stColumna, iAncho, DevExpress.Utils.FormatType.None, "", true, DevExpress.Utils.HorzAlignment.Near),
																																																					new DevExpress.XtraEditors.Controls.LookUpColumnInfo(stColumna1, stColumna1, iAncho1, DevExpress.Utils.FormatType.None, "", true, DevExpress.Utils.HorzAlignment.Near)});
			cmbCosto.Properties.Columns.AddRange(new DevExpress.XtraEditors.Controls.LookUpColumnInfo[] {
																																																		new DevExpress.XtraEditors.Controls.LookUpColumnInfo(stColumna, stColumna, iAncho, DevExpress.Utils.FormatType.None, "", true, DevExpress.Utils.HorzAlignment.Near),
																																																		new DevExpress.XtraEditors.Controls.LookUpColumnInfo(stColumna1, stColumna1, iAncho1, DevExpress.Utils.FormatType.None, "", true, DevExpress.Utils.HorzAlignment.Near)});
			cmbDescuento.Properties.Columns.AddRange(new DevExpress.XtraEditors.Controls.LookUpColumnInfo[] {
																																																				new DevExpress.XtraEditors.Controls.LookUpColumnInfo(stColumna, stColumna, iAncho, DevExpress.Utils.FormatType.None, "", true, DevExpress.Utils.HorzAlignment.Near),
																																																				new DevExpress.XtraEditors.Controls.LookUpColumnInfo(stColumna1, stColumna1, iAncho1, DevExpress.Utils.FormatType.None, "", true, DevExpress.Utils.HorzAlignment.Near)});
			cmbEgreso.Properties.Columns.AddRange(new DevExpress.XtraEditors.Controls.LookUpColumnInfo[] {
																																																		 new DevExpress.XtraEditors.Controls.LookUpColumnInfo(stColumna, stColumna, iAncho, DevExpress.Utils.FormatType.None, "", true, DevExpress.Utils.HorzAlignment.Near),
																																																		 new DevExpress.XtraEditors.Controls.LookUpColumnInfo(stColumna1, stColumna1, iAncho1, DevExpress.Utils.FormatType.None, "", true, DevExpress.Utils.HorzAlignment.Near)});
			
			cmbCuentaCompra.Properties.DisplayMember = stColumna;
			cmbCuentaVenta.Properties.DisplayMember = stColumna;
			cmbCosto.Properties.DisplayMember = stColumna;
			cmbDescuento.Properties.DisplayMember = stColumna;
			cmbEgreso.Properties.DisplayMember = stColumna;

			cmbCuentaCompra.Properties.Columns[0].SortOrder = DevExpress.Data.ColumnSortOrder.Ascending;
			cmbCuentaVenta.Properties.Columns[0].SortOrder = DevExpress.Data.ColumnSortOrder.Ascending;
			cmbCosto.Properties.Columns[0].SortOrder = DevExpress.Data.ColumnSortOrder.Ascending;
			cmbDescuento.Properties.Columns[0].SortOrder = DevExpress.Data.ColumnSortOrder.Ascending;
			cmbEgreso.Properties.Columns[0].SortOrder = DevExpress.Data.ColumnSortOrder.Ascending;
		}

		private void cdsFormaPago_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			this.cdsFormaPago.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);
		}

		private void xtraTabPage5_Paint(object sender, System.Windows.Forms.PaintEventArgs e)
		{
		
		}

		private void tlsCuentas_FocusedNodeChanged(object sender, DevExpress.XtraTreeList.FocusedNodeChangedEventArgs e)
		{
		
		}
	}
}
