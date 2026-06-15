using System;
using C1.Data;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de Compra2.
	/// </summary>
	public class Compra1 : System.Windows.Forms.Form
	{
		private DevExpress.XtraGrid.GridControl gridControl1;
		private DevExpress.XtraGrid.Views.Grid.GridView gridView1;
		private DevExpress.XtraEditors.TextEdit txtIdCompra;
		private BarraDatos.BarraDato barraDato1;
		private C1.Data.C1DataSet cdsCompra;
		private C1.Data.C1DataView cdvDetCompra;
		private C1.Data.C1DataSet cdsCliente;
		private C1.Data.C1DataView cdvCliente;
		private DevExpress.XtraEditors.LookUpEdit cmbCliente;
		private DevExpress.XtraGrid.Columns.GridColumn colCantidad;
		private C1.Data.C1DataSet cdsComprobante;
		private C1.Data.C1DataSet cdsArticulo;
		private C1.Data.C1DataView cdvArtCodigo;
		private C1.Data.C1DataView cdvNumero;
		private C1.Data.C1DataView cdvArticulo;
		private C1.Data.C1DataSet cdsNumero;
		private C1.Data.C1DataView cdvComprobante;
		private DevExpress.XtraGrid.Columns.GridColumn colidArticulo;
		private DevExpress.XtraGrid.Columns.GridColumn colidArticulo1;
		private DevExpress.XtraGrid.Columns.GridColumn colPrecio;
		private DevExpress.XtraGrid.Columns.GridColumn colSubTotal;
		private DevExpress.XtraGrid.Columns.GridColumn colDescuentoArt;
		private DevExpress.XtraEditors.Repository.RepositoryItemLookUpEdit repositoryItemLookUpEdit1;
		private DevExpress.XtraEditors.Repository.RepositoryItemLookUpEdit repositoryItemLookUpEdit2;
		private DevExpress.XtraGrid.Columns.GridColumn colImpuesto;
		private DevExpress.XtraGrid.Columns.GridColumn colDescuentoPorc;
		private DevExpress.XtraEditors.LookUpEdit cmbComprobante;
		private DevExpress.XtraEditors.LookUpEdit cmbTipoCompra;
		private DevExpress.XtraEditors.TextEdit txtNumero;
		private DevExpress.XtraEditors.DateEdit dtFecha;
		private System.Windows.Forms.TabControl tabControl1;
		private System.Windows.Forms.TabPage tabPage1;
		private System.Windows.Forms.TabPage tabPage2;
		private System.Windows.Forms.TabPage tabPage3;
		private System.Windows.Forms.TabPage tabPage4;
		private DevExpress.XtraEditors.DateEdit dateEdit1;
		private DevExpress.XtraEditors.SpinEdit spnDias;
		private DevExpress.XtraEditors.TextEdit textEdit1;
		private DevExpress.XtraEditors.TextEdit textEdit2;
		private System.Windows.Forms.Label label1;
		private System.Windows.Forms.Label label2;
		private System.Windows.Forms.Label label3;
		private System.Windows.Forms.Label label4;
		private System.Windows.Forms.ErrorProvider errorProvider;
		private System.Windows.Forms.ImageList imageList;
		private DevExpress.XtraEditors.SimpleButton simpleButton2;
		private DevExpress.XtraEditors.TextEdit textEdit3;
		private DevExpress.XtraEditors.TextEdit textEdit4;
		private DevExpress.XtraEditors.TextEdit textEdit5;
		private System.Windows.Forms.TabPage tabPage5;
		private System.Windows.Forms.TabPage tabPage6;
		private System.Windows.Forms.TabPage tabPage7;
		private System.Windows.Forms.TabPage tabPage8;
		private DevExpress.XtraEditors.MemoEdit memoEdit1;
		private System.Windows.Forms.Label label5;
		private DevExpress.XtraEditors.TextEdit textEdit6;
		private System.Windows.Forms.Label label6;
		private System.Windows.Forms.Label label7;
		private System.Windows.Forms.Label label8;
		private System.Windows.Forms.Label label9;
		private DevExpress.XtraEditors.SimpleButton btCliente;
		private System.Windows.Forms.Label label11;
		private System.Windows.Forms.Label label10;
		private DevExpress.XtraEditors.TextEdit textEdit7;
		private System.Windows.Forms.Label label12;
		private DevExpress.XtraEditors.TextEdit textEdit8;
		private DevExpress.XtraGrid.Columns.GridColumn colBodega;
		private DevExpress.XtraEditors.SpinEdit spnBodega;
		private C1.Data.C1DataSet cdsSeguridad;
		private System.Windows.Forms.Label lblAnulado;
		private DevExpress.XtraEditors.TextEdit txtBorrar;
		private DevExpress.XtraEditors.SimpleButton btSaldo;
		private DevExpress.XtraEditors.SimpleButton btNotaEntrega;
		private DevExpress.XtraEditors.SimpleButton btGuiaRemision;
		private DevExpress.XtraEditors.SimpleButton btCreaAsiento;
		private DevExpress.XtraEditors.SimpleButton btRetencion;
		private DevExpress.XtraEditors.SimpleButton btTransforma;
		private DevExpress.XtraEditors.SimpleButton btBuscaArticulo;
		private DevExpress.XtraEditors.RadioGroup rdgContadiCred;
		private DevExpress.XtraEditors.SimpleButton btPagos;
		private C1.Data.C1DataSet cdsProyecto;
		private C1.Data.C1DataView cdvProyecto;
		private C1.Data.C1DataView cdvSubProyecto;
		private DevExpress.XtraEditors.LookUpEdit cmbProyecto;
		private DevExpress.XtraEditors.LookUpEdit cmbSubProyecto;
		private System.Windows.Forms.Label label13;
		private System.Windows.Forms.Label label14;
		private System.Windows.Forms.Label label15;
		private DevExpress.XtraEditors.TextEdit textEdit9;
		private C1.Data.C1DataSet cdsFormPago;
		private C1.Data.C1DataView cdvFormaPago;
		private DevExpress.XtraEditors.LookUpEdit cmbFormaPago;
		private System.Windows.Forms.Label label16;
		private DevExpress.XtraEditors.TextEdit txtDescIva;
		private DevExpress.XtraEditors.TextEdit txtDescPorcIva;
		private DevExpress.XtraEditors.TextEdit textEdit10;
		private System.Windows.Forms.Label label17;
		private DevExpress.XtraEditors.TextEdit txtDescIva0;
		private System.ComponentModel.IContainer components;

		public Compra1()
		{
			InitializeComponent();
		}

		int IdTipoFactura = 1;
		public Compra1(int idTransaccion)
		{
			InitializeComponent();
			IdTipoFactura = idTransaccion;
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
			this.components = new System.ComponentModel.Container();
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(Compra1));
			this.gridControl1 = new DevExpress.XtraGrid.GridControl();
			this.cdvDetCompra = new C1.Data.C1DataView();
			this.cdsCompra = new C1.Data.C1DataSet();
			this.gridView1 = new DevExpress.XtraGrid.Views.Grid.GridView();
			this.colCantidad = new DevExpress.XtraGrid.Columns.GridColumn();
			this.colidArticulo = new DevExpress.XtraGrid.Columns.GridColumn();
			this.repositoryItemLookUpEdit1 = new DevExpress.XtraEditors.Repository.RepositoryItemLookUpEdit();
			this.cdvArticulo = new C1.Data.C1DataView();
			this.cdsArticulo = new C1.Data.C1DataSet();
			this.colidArticulo1 = new DevExpress.XtraGrid.Columns.GridColumn();
			this.repositoryItemLookUpEdit2 = new DevExpress.XtraEditors.Repository.RepositoryItemLookUpEdit();
			this.colPrecio = new DevExpress.XtraGrid.Columns.GridColumn();
			this.colImpuesto = new DevExpress.XtraGrid.Columns.GridColumn();
			this.colSubTotal = new DevExpress.XtraGrid.Columns.GridColumn();
			this.colDescuentoArt = new DevExpress.XtraGrid.Columns.GridColumn();
			this.colDescuentoPorc = new DevExpress.XtraGrid.Columns.GridColumn();
			this.colBodega = new DevExpress.XtraGrid.Columns.GridColumn();
			this.txtIdCompra = new DevExpress.XtraEditors.TextEdit();
			this.barraDato1 = new BarraDatos.BarraDato();
			this.cdsCliente = new C1.Data.C1DataSet();
			this.cdvCliente = new C1.Data.C1DataView();
			this.cmbCliente = new DevExpress.XtraEditors.LookUpEdit();
			this.cdsComprobante = new C1.Data.C1DataSet();
			this.cdvArtCodigo = new C1.Data.C1DataView();
			this.cdvNumero = new C1.Data.C1DataView();
			this.cdsNumero = new C1.Data.C1DataSet();
			this.cdvComprobante = new C1.Data.C1DataView();
			this.cmbComprobante = new DevExpress.XtraEditors.LookUpEdit();
			this.cmbTipoCompra = new DevExpress.XtraEditors.LookUpEdit();
			this.txtNumero = new DevExpress.XtraEditors.TextEdit();
			this.dtFecha = new DevExpress.XtraEditors.DateEdit();
			this.spnDias = new DevExpress.XtraEditors.SpinEdit();
			this.tabControl1 = new System.Windows.Forms.TabControl();
			this.tabPage1 = new System.Windows.Forms.TabPage();
			this.label15 = new System.Windows.Forms.Label();
			this.textEdit9 = new DevExpress.XtraEditors.TextEdit();
			this.label14 = new System.Windows.Forms.Label();
			this.label13 = new System.Windows.Forms.Label();
			this.cmbSubProyecto = new DevExpress.XtraEditors.LookUpEdit();
			this.cdvSubProyecto = new C1.Data.C1DataView();
			this.cdsProyecto = new C1.Data.C1DataSet();
			this.memoEdit1 = new DevExpress.XtraEditors.MemoEdit();
			this.label5 = new System.Windows.Forms.Label();
			this.textEdit6 = new DevExpress.XtraEditors.TextEdit();
			this.cmbProyecto = new DevExpress.XtraEditors.LookUpEdit();
			this.cdvProyecto = new C1.Data.C1DataView();
			this.tabPage2 = new System.Windows.Forms.TabPage();
			this.cmbFormaPago = new DevExpress.XtraEditors.LookUpEdit();
			this.cdvFormaPago = new C1.Data.C1DataView();
			this.cdsFormPago = new C1.Data.C1DataSet();
			this.btPagos = new DevExpress.XtraEditors.SimpleButton();
			this.rdgContadiCred = new DevExpress.XtraEditors.RadioGroup();
			this.tabPage3 = new System.Windows.Forms.TabPage();
			this.textEdit10 = new DevExpress.XtraEditors.TextEdit();
			this.label17 = new System.Windows.Forms.Label();
			this.txtDescIva0 = new DevExpress.XtraEditors.TextEdit();
			this.txtDescPorcIva = new DevExpress.XtraEditors.TextEdit();
			this.label16 = new System.Windows.Forms.Label();
			this.txtDescIva = new DevExpress.XtraEditors.TextEdit();
			this.tabPage4 = new System.Windows.Forms.TabPage();
			this.tabPage5 = new System.Windows.Forms.TabPage();
			this.tabPage6 = new System.Windows.Forms.TabPage();
			this.tabPage7 = new System.Windows.Forms.TabPage();
			this.tabPage8 = new System.Windows.Forms.TabPage();
			this.dateEdit1 = new DevExpress.XtraEditors.DateEdit();
			this.spnBodega = new DevExpress.XtraEditors.SpinEdit();
			this.textEdit1 = new DevExpress.XtraEditors.TextEdit();
			this.textEdit2 = new DevExpress.XtraEditors.TextEdit();
			this.label1 = new System.Windows.Forms.Label();
			this.label2 = new System.Windows.Forms.Label();
			this.label3 = new System.Windows.Forms.Label();
			this.label4 = new System.Windows.Forms.Label();
			this.btSaldo = new DevExpress.XtraEditors.SimpleButton();
			this.imageList = new System.Windows.Forms.ImageList(this.components);
			this.errorProvider = new System.Windows.Forms.ErrorProvider();
			this.simpleButton2 = new DevExpress.XtraEditors.SimpleButton();
			this.btNotaEntrega = new DevExpress.XtraEditors.SimpleButton();
			this.btGuiaRemision = new DevExpress.XtraEditors.SimpleButton();
			this.btCreaAsiento = new DevExpress.XtraEditors.SimpleButton();
			this.btRetencion = new DevExpress.XtraEditors.SimpleButton();
			this.btTransforma = new DevExpress.XtraEditors.SimpleButton();
			this.btBuscaArticulo = new DevExpress.XtraEditors.SimpleButton();
			this.textEdit3 = new DevExpress.XtraEditors.TextEdit();
			this.textEdit4 = new DevExpress.XtraEditors.TextEdit();
			this.textEdit5 = new DevExpress.XtraEditors.TextEdit();
			this.label6 = new System.Windows.Forms.Label();
			this.label7 = new System.Windows.Forms.Label();
			this.label8 = new System.Windows.Forms.Label();
			this.label9 = new System.Windows.Forms.Label();
			this.btCliente = new DevExpress.XtraEditors.SimpleButton();
			this.label11 = new System.Windows.Forms.Label();
			this.label10 = new System.Windows.Forms.Label();
			this.textEdit7 = new DevExpress.XtraEditors.TextEdit();
			this.label12 = new System.Windows.Forms.Label();
			this.textEdit8 = new DevExpress.XtraEditors.TextEdit();
			this.cdsSeguridad = new C1.Data.C1DataSet();
			this.lblAnulado = new System.Windows.Forms.Label();
			this.txtBorrar = new DevExpress.XtraEditors.TextEdit();
			((System.ComponentModel.ISupportInitialize)(this.gridControl1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvDetCompra)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsCompra)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.gridView1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.repositoryItemLookUpEdit1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvArticulo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsArticulo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.repositoryItemLookUpEdit2)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdCompra.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsCliente)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvCliente)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCliente.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsComprobante)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvArtCodigo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvNumero)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsNumero)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvComprobante)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbComprobante.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbTipoCompra.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNumero.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFecha.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.spnDias.Properties)).BeginInit();
			this.tabControl1.SuspendLayout();
			this.tabPage1.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.textEdit9.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbSubProyecto.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvSubProyecto)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsProyecto)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.memoEdit1.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.textEdit6.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbProyecto.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvProyecto)).BeginInit();
			this.tabPage2.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.cmbFormaPago.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvFormaPago)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsFormPago)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.rdgContadiCred.Properties)).BeginInit();
			this.tabPage3.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.textEdit10.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDescIva0.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDescPorcIva.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDescIva.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dateEdit1.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.spnBodega.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.textEdit1.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.textEdit2.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.textEdit3.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.textEdit4.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.textEdit5.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.textEdit7.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.textEdit8.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeguridad)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtBorrar.Properties)).BeginInit();
			this.SuspendLayout();
			// 
			// gridControl1
			// 
			this.gridControl1.Anchor = (((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right);
			this.gridControl1.DataSource = this.cdvDetCompra;
			this.gridControl1.Location = new System.Drawing.Point(16, 64);
			this.gridControl1.MainView = this.gridView1;
			this.gridControl1.Name = "gridControl1";
			this.gridControl1.RepositoryItems.AddRange(new DevExpress.XtraEditors.Repository.RepositoryItem[] {
																																																					this.repositoryItemLookUpEdit1,
																																																					this.repositoryItemLookUpEdit2});
			this.gridControl1.Size = new System.Drawing.Size(656, 176);
			this.gridControl1.Styles.AddReplace("Preview", new DevExpress.Utils.ViewStyleEx("Preview", "Grid", new System.Drawing.Font("Tahoma", 10F), "", true, true, false, DevExpress.Utils.HorzAlignment.Near, DevExpress.Utils.VertAlignment.Top, null, System.Drawing.Color.White, System.Drawing.Color.FromArgb(((System.Byte)(10)), ((System.Byte)(36)), ((System.Byte)(106))), System.Drawing.Color.White, System.Drawing.Drawing2D.LinearGradientMode.Horizontal));
			this.gridControl1.Styles.AddReplace("FooterPanel", new DevExpress.Utils.ViewStyleEx("FooterPanel", "Grid", new System.Drawing.Font("Tahoma", 10F), "", true, false, false, DevExpress.Utils.HorzAlignment.Far, DevExpress.Utils.VertAlignment.Center, null, System.Drawing.Color.FromArgb(((System.Byte)(212)), ((System.Byte)(208)), ((System.Byte)(200))), System.Drawing.Color.Black, System.Drawing.Color.Empty, System.Drawing.Drawing2D.LinearGradientMode.Horizontal));
			this.gridControl1.Styles.AddReplace("GroupButton", new DevExpress.Utils.ViewStyleEx("GroupButton", "Grid", new System.Drawing.Font("Tahoma", 10F), System.Drawing.Color.FromArgb(((System.Byte)(212)), ((System.Byte)(208)), ((System.Byte)(200))), System.Drawing.Color.Black, System.Drawing.Color.Empty, System.Drawing.Drawing2D.LinearGradientMode.Horizontal));
			this.gridControl1.Styles.AddReplace("FilterCloseButton", new DevExpress.Utils.ViewStyleEx("FilterCloseButton", "Grid", new System.Drawing.Font("Arial", 8F), System.Drawing.Color.FromArgb(((System.Byte)(212)), ((System.Byte)(208)), ((System.Byte)(200))), System.Drawing.Color.Black, System.Drawing.Color.FromArgb(((System.Byte)(118)), ((System.Byte)(170)), ((System.Byte)(225))), System.Drawing.Drawing2D.LinearGradientMode.ForwardDiagonal));
			this.gridControl1.Styles.AddReplace("EvenRow", new DevExpress.Utils.ViewStyleEx("EvenRow", "Grid", new System.Drawing.Font("Tahoma", 10F), (DevExpress.Utils.StyleOptions.StyleEnabled | DevExpress.Utils.StyleOptions.UseBackColor), System.Drawing.Color.FromArgb(((System.Byte)(212)), ((System.Byte)(208)), ((System.Byte)(200))), System.Drawing.Color.Black, System.Drawing.Color.GhostWhite, System.Drawing.Drawing2D.LinearGradientMode.ForwardDiagonal));
			this.gridControl1.Styles.AddReplace("HideSelectionRow", new DevExpress.Utils.ViewStyleEx("HideSelectionRow", "Grid", new System.Drawing.Font("Tahoma", 10F), (((((DevExpress.Utils.StyleOptions.StyleEnabled | DevExpress.Utils.StyleOptions.UseBackColor) 
				| DevExpress.Utils.StyleOptions.UseDrawFocusRect) 
				| DevExpress.Utils.StyleOptions.UseFont) 
				| DevExpress.Utils.StyleOptions.UseForeColor) 
				| DevExpress.Utils.StyleOptions.UseImage), System.Drawing.Color.Gray, System.Drawing.Color.FromArgb(((System.Byte)(212)), ((System.Byte)(208)), ((System.Byte)(200))), System.Drawing.Color.Empty, System.Drawing.Drawing2D.LinearGradientMode.Horizontal));
			this.gridControl1.Styles.AddReplace("FixedLine", new DevExpress.Utils.ViewStyleEx("FixedLine", "Grid", new System.Drawing.Font("Arial", 8F), System.Drawing.Color.FromArgb(((System.Byte)(58)), ((System.Byte)(58)), ((System.Byte)(58))), System.Drawing.Color.Black, System.Drawing.Color.Empty, System.Drawing.Drawing2D.LinearGradientMode.Horizontal));
			this.gridControl1.Styles.AddReplace("HeaderPanel", new DevExpress.Utils.ViewStyleEx("HeaderPanel", "Grid", new System.Drawing.Font("Tahoma", 10F, System.Drawing.FontStyle.Bold), "", true, false, false, DevExpress.Utils.HorzAlignment.Near, DevExpress.Utils.VertAlignment.Center, null, System.Drawing.Color.FromArgb(((System.Byte)(212)), ((System.Byte)(208)), ((System.Byte)(200))), System.Drawing.Color.Black, System.Drawing.Color.Empty, System.Drawing.Drawing2D.LinearGradientMode.Horizontal));
			this.gridControl1.Styles.AddReplace("GroupPanel", new DevExpress.Utils.ViewStyleEx("GroupPanel", "Grid", new System.Drawing.Font("Tahoma", 10F, System.Drawing.FontStyle.Bold), "", true, false, false, DevExpress.Utils.HorzAlignment.Near, DevExpress.Utils.VertAlignment.Center, null, System.Drawing.Color.FromArgb(((System.Byte)(58)), ((System.Byte)(110)), ((System.Byte)(165))), System.Drawing.Color.White, System.Drawing.Color.White, System.Drawing.Drawing2D.LinearGradientMode.Horizontal));
			this.gridControl1.Styles.AddReplace("Empty", new DevExpress.Utils.ViewStyleEx("Empty", "Grid", new System.Drawing.Font("Arial", 8F), System.Drawing.Color.FromArgb(((System.Byte)(232)), ((System.Byte)(228)), ((System.Byte)(220))), System.Drawing.Color.White, System.Drawing.Color.Empty, System.Drawing.Drawing2D.LinearGradientMode.Horizontal));
			this.gridControl1.Styles.AddReplace("GroupFooter", new DevExpress.Utils.ViewStyleEx("GroupFooter", "Grid", new System.Drawing.Font("Tahoma", 8F), "", true, false, false, DevExpress.Utils.HorzAlignment.Far, DevExpress.Utils.VertAlignment.Center, null, System.Drawing.Color.FromArgb(((System.Byte)(222)), ((System.Byte)(218)), ((System.Byte)(210))), System.Drawing.Color.Black, System.Drawing.Color.Empty, System.Drawing.Drawing2D.LinearGradientMode.Horizontal));
			this.gridControl1.Styles.AddReplace("GroupRow", new DevExpress.Utils.ViewStyleEx("GroupRow", "Grid", new System.Drawing.Font("Tahoma", 10F), "", true, false, false, DevExpress.Utils.HorzAlignment.Near, DevExpress.Utils.VertAlignment.Center, null, System.Drawing.Color.Gray, System.Drawing.Color.FromArgb(((System.Byte)(212)), ((System.Byte)(208)), ((System.Byte)(200))), System.Drawing.Color.Empty, System.Drawing.Drawing2D.LinearGradientMode.Horizontal));
			this.gridControl1.Styles.AddReplace("HorzLine", new DevExpress.Utils.ViewStyleEx("HorzLine", "Grid", new System.Drawing.Font("Tahoma", 10F), System.Drawing.Color.FromArgb(((System.Byte)(212)), ((System.Byte)(208)), ((System.Byte)(200))), System.Drawing.Color.Gray, System.Drawing.Color.Empty, System.Drawing.Drawing2D.LinearGradientMode.Horizontal));
			this.gridControl1.Styles.AddReplace("ColumnFilterButton", new DevExpress.Utils.ViewStyleEx("ColumnFilterButton", "Grid", new System.Drawing.Font("Arial", 8F), System.Drawing.Color.FromArgb(((System.Byte)(212)), ((System.Byte)(208)), ((System.Byte)(200))), System.Drawing.Color.Gray, System.Drawing.Color.FromArgb(((System.Byte)(232)), ((System.Byte)(228)), ((System.Byte)(220))), System.Drawing.Drawing2D.LinearGradientMode.Horizontal));
			this.gridControl1.Styles.AddReplace("FocusedRow", new DevExpress.Utils.ViewStyleEx("FocusedRow", "Grid", new System.Drawing.Font("Tahoma", 10F), (((((DevExpress.Utils.StyleOptions.StyleEnabled | DevExpress.Utils.StyleOptions.UseBackColor) 
				| DevExpress.Utils.StyleOptions.UseDrawFocusRect) 
				| DevExpress.Utils.StyleOptions.UseFont) 
				| DevExpress.Utils.StyleOptions.UseForeColor) 
				| DevExpress.Utils.StyleOptions.UseImage), System.Drawing.Color.FromArgb(((System.Byte)(10)), ((System.Byte)(36)), ((System.Byte)(106))), System.Drawing.Color.White, System.Drawing.Color.FromArgb(((System.Byte)(60)), ((System.Byte)(86)), ((System.Byte)(156))), System.Drawing.Drawing2D.LinearGradientMode.Horizontal));
			this.gridControl1.Styles.AddReplace("VertLine", new DevExpress.Utils.ViewStyleEx("VertLine", "Grid", new System.Drawing.Font("Tahoma", 10F), "", true, false, false, DevExpress.Utils.HorzAlignment.Near, DevExpress.Utils.VertAlignment.Center, null, System.Drawing.Color.FromArgb(((System.Byte)(212)), ((System.Byte)(208)), ((System.Byte)(200))), System.Drawing.Color.Gray, System.Drawing.Color.Empty, System.Drawing.Drawing2D.LinearGradientMode.Horizontal));
			this.gridControl1.Styles.AddReplace("FocusedCell", new DevExpress.Utils.ViewStyleEx("FocusedCell", "Grid", new System.Drawing.Font("Tahoma", 10F), (((((DevExpress.Utils.StyleOptions.StyleEnabled | DevExpress.Utils.StyleOptions.UseBackColor) 
				| DevExpress.Utils.StyleOptions.UseDrawFocusRect) 
				| DevExpress.Utils.StyleOptions.UseFont) 
				| DevExpress.Utils.StyleOptions.UseForeColor) 
				| DevExpress.Utils.StyleOptions.UseImage), System.Drawing.Color.FromArgb(((System.Byte)(255)), ((System.Byte)(255)), ((System.Byte)(225))), System.Drawing.Color.Black, System.Drawing.Color.Empty, System.Drawing.Drawing2D.LinearGradientMode.Horizontal));
			this.gridControl1.Styles.AddReplace("OddRow", new DevExpress.Utils.ViewStyleEx("OddRow", "Grid", new System.Drawing.Font("Tahoma", 10F), (DevExpress.Utils.StyleOptions.StyleEnabled | DevExpress.Utils.StyleOptions.UseBackColor), System.Drawing.Color.White, System.Drawing.Color.Black, System.Drawing.Color.White, System.Drawing.Drawing2D.LinearGradientMode.BackwardDiagonal));
			this.gridControl1.Styles.AddReplace("SelectedRow", new DevExpress.Utils.ViewStyleEx("SelectedRow", "Grid", new System.Drawing.Font("Tahoma", 10F), (((((DevExpress.Utils.StyleOptions.StyleEnabled | DevExpress.Utils.StyleOptions.UseBackColor) 
				| DevExpress.Utils.StyleOptions.UseDrawFocusRect) 
				| DevExpress.Utils.StyleOptions.UseFont) 
				| DevExpress.Utils.StyleOptions.UseForeColor) 
				| DevExpress.Utils.StyleOptions.UseImage), System.Drawing.Color.FromArgb(((System.Byte)(20)), ((System.Byte)(46)), ((System.Byte)(116))), System.Drawing.Color.White, System.Drawing.Color.Empty, System.Drawing.Drawing2D.LinearGradientMode.Horizontal));
			this.gridControl1.Styles.AddReplace("Row", new DevExpress.Utils.ViewStyleEx("Row", "Grid", new System.Drawing.Font("Tahoma", 10F), DevExpress.Utils.StyleOptions.StyleEnabled, System.Drawing.Color.White, System.Drawing.Color.Black, System.Drawing.Color.Empty, System.Drawing.Drawing2D.LinearGradientMode.Horizontal));
			this.gridControl1.Styles.AddReplace("FilterPanel", new DevExpress.Utils.ViewStyleEx("FilterPanel", "Grid", new System.Drawing.Font("Arial", 8F), "", true, false, false, DevExpress.Utils.HorzAlignment.Near, DevExpress.Utils.VertAlignment.Center, null, System.Drawing.Color.FromArgb(((System.Byte)(28)), ((System.Byte)(80)), ((System.Byte)(135))), System.Drawing.Color.White, System.Drawing.Color.FromArgb(((System.Byte)(212)), ((System.Byte)(208)), ((System.Byte)(200))), System.Drawing.Drawing2D.LinearGradientMode.ForwardDiagonal));
			this.gridControl1.Styles.AddReplace("RowSeparator", new DevExpress.Utils.ViewStyleEx("RowSeparator", "Grid", new System.Drawing.Font("Arial", 8F), DevExpress.Utils.StyleOptions.StyleEnabled, System.Drawing.Color.White, System.Drawing.Color.Gray, System.Drawing.Color.FromArgb(((System.Byte)(232)), ((System.Byte)(228)), ((System.Byte)(220))), System.Drawing.Drawing2D.LinearGradientMode.Horizontal));
			this.gridControl1.Styles.AddReplace("ColumnFilterButtonActive", new DevExpress.Utils.ViewStyleEx("ColumnFilterButtonActive", "Grid", new System.Drawing.Font("Arial", 8F), System.Drawing.Color.LightGray, System.Drawing.Color.Blue, System.Drawing.Color.WhiteSmoke, System.Drawing.Drawing2D.LinearGradientMode.Horizontal));
			this.gridControl1.Styles.AddReplace("DetailTip", new DevExpress.Utils.ViewStyleEx("DetailTip", "Grid", new System.Drawing.Font("Arial", 8F), System.Drawing.Color.FromArgb(((System.Byte)(255)), ((System.Byte)(255)), ((System.Byte)(225))), System.Drawing.Color.Black, System.Drawing.Color.Empty, System.Drawing.Drawing2D.LinearGradientMode.Horizontal));
			this.gridControl1.TabIndex = 0;
			this.gridControl1.Text = "gridControl1";
			// 
			// cdvDetCompra
			// 
			this.cdvDetCompra.BindingContextCtrl = this;
			this.cdvDetCompra.DataSet = this.cdsCompra;
			this.cdvDetCompra.TableName = "";
			this.cdvDetCompra.TableViewName = "DetCompra";
			this.cdvDetCompra.PositionChanged += new C1.Data.PositionChangeEventHandler(this.cdvDetCompra_PositionChanged);
			// 
			// cdsCompra
			// 
			this.cdsCompra.BindingContextCtrl = this;
			this.cdsCompra.DataLibrary = "LibFacturacion";
			this.cdsCompra.DataLibraryUrl = "";
			this.cdsCompra.DataSetDef = "dsCompra";
			this.cdsCompra.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsCompra.Name = "cdsCompra";
			this.cdsCompra.SchemaClassName = "LibFacturacion.DataClass";
			this.cdsCompra.SchemaDef = null;
			this.cdsCompra.PositionChanged += new C1.Data.PositionChangeEventHandler(this.cdsCompra_PositionChanged);
			this.cdsCompra.BeforeFill += new C1.Data.FillEventHandler(this.cdsCompra_BeforeFill);
			// 
			// gridView1
			// 
			this.gridView1.Columns.AddRange(new DevExpress.XtraGrid.Columns.GridColumn[] {
																																										 this.colCantidad,
																																										 this.colidArticulo,
																																										 this.colidArticulo1,
																																										 this.colPrecio,
																																										 this.colImpuesto,
																																										 this.colSubTotal,
																																										 this.colDescuentoArt,
																																										 this.colDescuentoPorc,
																																										 this.colBodega});
			this.gridView1.Name = "gridView1";
			this.gridView1.OptionsView.ColumnAutoWidth = false;
			this.gridView1.OptionsView.ShowGroupPanel = false;
			this.gridView1.OptionsView.ShowNewItemRow = true;
			// 
			// colCantidad
			// 
			this.colCantidad.Caption = "Cantidad";
			this.colCantidad.FieldName = "Cantidad";
			this.colCantidad.Fixed = DevExpress.XtraGrid.Columns.FixedStyle.Left;
			this.colCantidad.Name = "colCantidad";
			this.colCantidad.VisibleIndex = 0;
			// 
			// colidArticulo
			// 
			this.colidArticulo.Caption = "Descripción";
			this.colidArticulo.ColumnEdit = this.repositoryItemLookUpEdit1;
			this.colidArticulo.FieldName = "idArticulo";
			this.colidArticulo.Fixed = DevExpress.XtraGrid.Columns.FixedStyle.Left;
			this.colidArticulo.Name = "colidArticulo";
			this.colidArticulo.VisibleIndex = 1;
			this.colidArticulo.Width = 129;
			// 
			// repositoryItemLookUpEdit1
			// 
			this.repositoryItemLookUpEdit1.AutoHeight = false;
			this.repositoryItemLookUpEdit1.Columns.AddRange(new DevExpress.XtraEditors.Controls.LookUpColumnInfo[] {
																																																							 new DevExpress.XtraEditors.Controls.LookUpColumnInfo("Articulo", "Articulo", 47, DevExpress.Utils.FormatType.None, "", true, DevExpress.Utils.HorzAlignment.Near),
																																																							 new DevExpress.XtraEditors.Controls.LookUpColumnInfo("Codigo", "Codigo", 45, DevExpress.Utils.FormatType.None, "", true, DevExpress.Utils.HorzAlignment.Near)});
			this.repositoryItemLookUpEdit1.DataSource = this.cdvArticulo;
			this.repositoryItemLookUpEdit1.DisplayMember = "Articulo";
			this.repositoryItemLookUpEdit1.Name = "repositoryItemLookUpEdit1";
			this.repositoryItemLookUpEdit1.NullText = "";
			this.repositoryItemLookUpEdit1.ValueMember = "idArticulo";
			// 
			// cdvArticulo
			// 
			this.cdvArticulo.BindingContextCtrl = this;
			this.cdvArticulo.DataSet = this.cdsArticulo;
			this.cdvArticulo.Sort = "Articulo";
			this.cdvArticulo.TableName = "";
			this.cdvArticulo.TableViewName = "Articulo";
			// 
			// cdsArticulo
			// 
			this.cdsArticulo.BindingContextCtrl = this;
			this.cdsArticulo.DataLibrary = "LibFacturacion";
			this.cdsArticulo.DataLibraryUrl = "";
			this.cdsArticulo.DataSetDef = "dsArticulo";
			this.cdsArticulo.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsArticulo.Name = "cdsArticulo";
			this.cdsArticulo.SchemaClassName = "LibFacturacion.DataClass";
			this.cdsArticulo.SchemaDef = null;
			// 
			// colidArticulo1
			// 
			this.colidArticulo1.Caption = "Código";
			this.colidArticulo1.ColumnEdit = this.repositoryItemLookUpEdit2;
			this.colidArticulo1.FieldName = "idArticulo";
			this.colidArticulo1.Name = "colidArticulo1";
			this.colidArticulo1.VisibleIndex = 2;
			// 
			// repositoryItemLookUpEdit2
			// 
			this.repositoryItemLookUpEdit2.AutoHeight = false;
			this.repositoryItemLookUpEdit2.Buttons.AddRange(new DevExpress.XtraEditors.Controls.EditorButton[] {
																																																					 new DevExpress.XtraEditors.Controls.EditorButton(DevExpress.XtraEditors.Controls.ButtonPredefines.Combo)});
			this.repositoryItemLookUpEdit2.Columns.AddRange(new DevExpress.XtraEditors.Controls.LookUpColumnInfo[] {
																																																							 new DevExpress.XtraEditors.Controls.LookUpColumnInfo("Codigo", "Codigo", 45, DevExpress.Utils.FormatType.None, "", true, DevExpress.Utils.HorzAlignment.Near),
																																																							 new DevExpress.XtraEditors.Controls.LookUpColumnInfo("Articulo", "Articulo", 47, DevExpress.Utils.FormatType.None, "", true, DevExpress.Utils.HorzAlignment.Near)});
			this.repositoryItemLookUpEdit2.DataSource = this.cdvArticulo;
			this.repositoryItemLookUpEdit2.DisplayMember = "Codigo";
			this.repositoryItemLookUpEdit2.Name = "repositoryItemLookUpEdit2";
			this.repositoryItemLookUpEdit2.NullText = "";
			this.repositoryItemLookUpEdit2.ValueMember = "idArticulo";
			// 
			// colPrecio
			// 
			this.colPrecio.Caption = "Precio";
			this.colPrecio.FieldName = "Precio";
			this.colPrecio.Name = "colPrecio";
			this.colPrecio.VisibleIndex = 4;
			// 
			// colImpuesto
			// 
			this.colImpuesto.Caption = "Iva";
			this.colImpuesto.FieldName = "Impuesto";
			this.colImpuesto.Name = "colImpuesto";
			this.colImpuesto.VisibleIndex = 3;
			this.colImpuesto.Width = 29;
			// 
			// colSubTotal
			// 
			this.colSubTotal.Caption = "SubTotal";
			this.colSubTotal.FieldName = "SubTotal";
			this.colSubTotal.Fixed = DevExpress.XtraGrid.Columns.FixedStyle.Right;
			this.colSubTotal.Name = "colSubTotal";
			this.colSubTotal.VisibleIndex = 8;
			// 
			// colDescuentoArt
			// 
			this.colDescuentoArt.Caption = "Desc.";
			this.colDescuentoArt.FieldName = "DescuentoArt";
			this.colDescuentoArt.Name = "colDescuentoArt";
			this.colDescuentoArt.VisibleIndex = 5;
			// 
			// colDescuentoPorc
			// 
			this.colDescuentoPorc.Caption = "% Desc.";
			this.colDescuentoPorc.FieldName = "DescuentoPorc";
			this.colDescuentoPorc.Name = "colDescuentoPorc";
			this.colDescuentoPorc.VisibleIndex = 6;
			// 
			// colBodega
			// 
			this.colBodega.Caption = "Bodega";
			this.colBodega.FieldName = "Bodega";
			this.colBodega.Name = "colBodega";
			this.colBodega.VisibleIndex = 7;
			// 
			// txtIdCompra
			// 
			this.txtIdCompra.DataBindings.Add(new System.Windows.Forms.Binding("EditValue", this.cdsCompra, "Compra.idCompra"));
			this.txtIdCompra.EditValue = "txtIdCompra";
			this.txtIdCompra.Location = new System.Drawing.Point(608, 224);
			this.txtIdCompra.Name = "txtIdCompra";
			this.txtIdCompra.Size = new System.Drawing.Size(75, 19);
			this.txtIdCompra.TabIndex = 1;
			this.txtIdCompra.EditValueChanged += new System.EventHandler(this.txtIdCompra_EditValueChanged);
			// 
			// barraDato1
			// 
			this.barraDato1.Anchor = (System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left);
			this.barraDato1.BarraMovimiento = true;
			this.barraDato1.DataMember = "Compra";
			this.barraDato1.DataSource = this.cdsCompra;
			this.barraDato1.Location = new System.Drawing.Point(16, 400);
			this.barraDato1.Name = "barraDato1";
			this.barraDato1.Size = new System.Drawing.Size(304, 24);
			this.barraDato1.TabIndex = 2;
			this.barraDato1.VisibleBorrar = true;
			this.barraDato1.VisibleBuscar = true;
			this.barraDato1.VisibleEditar = true;
			this.barraDato1.VisibleImprimir = true;
			this.barraDato1.VisibleNuevo = true;
			this.barraDato1.Buscar += new System.EventHandler(this.barraDato1_Buscar);
			this.barraDato1.Borrar += new System.EventHandler(this.barraDato1_Borrar);
			this.barraDato1.Imprimir += new System.EventHandler(this.barraDato1_Imprimir);
			this.barraDato1.Crear += new System.EventHandler(this.barraDato1_Crear);
			this.barraDato1.Deshacer += new System.EventHandler(this.barraDato1_Deshacer);
			this.barraDato1.Editar += new System.EventHandler(this.barraDato1_Editar);
			this.barraDato1.Grabar += new System.EventHandler(this.barraDato1_Grabar);
			// 
			// cdsCliente
			// 
			this.cdsCliente.BindingContextCtrl = this;
			this.cdsCliente.DataLibrary = "LibFacturacion";
			this.cdsCliente.DataLibraryUrl = "";
			this.cdsCliente.DataSetDef = "dsCliente";
			this.cdsCliente.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsCliente.Name = "cdsCliente";
			this.cdsCliente.SchemaClassName = "LibFacturacion.DataClass";
			this.cdsCliente.SchemaDef = null;
			this.cdsCliente.BeforeFill += new C1.Data.FillEventHandler(this.cdsCliente_BeforeFill);
			// 
			// cdvCliente
			// 
			this.cdvCliente.BindingContextCtrl = this;
			this.cdvCliente.DataSet = this.cdsCliente;
			this.cdvCliente.TableName = "";
			this.cdvCliente.TableViewName = "Cliente";
			// 
			// cmbCliente
			// 
			this.cmbCliente.DataBindings.Add(new System.Windows.Forms.Binding("EditValue", this.cdsCompra, "Compra.idCliente"));
			this.cmbCliente.Location = new System.Drawing.Point(280, 8);
			this.cmbCliente.Name = "cmbCliente";
			// 
			// cmbCliente.Properties
			// 
			this.cmbCliente.Properties.Buttons.AddRange(new DevExpress.XtraEditors.Controls.EditorButton[] {
																																																			 new DevExpress.XtraEditors.Controls.EditorButton(DevExpress.XtraEditors.Controls.ButtonPredefines.Combo)});
			this.cmbCliente.Properties.Columns.AddRange(new DevExpress.XtraEditors.Controls.LookUpColumnInfo[] {
																																																					 new DevExpress.XtraEditors.Controls.LookUpColumnInfo("Nombre", "Nombre", 49, DevExpress.Utils.FormatType.None, "", true, DevExpress.Utils.HorzAlignment.Near)});
			this.cmbCliente.Properties.DataSource = this.cdvCliente;
			this.cmbCliente.Properties.DisplayMember = "Nombre";
			this.cmbCliente.Properties.NullText = "";
			this.cmbCliente.Properties.ShowFooter = false;
			this.cmbCliente.Properties.ShowHeader = false;
			this.cmbCliente.Properties.ValueMember = "idCliente";
			this.cmbCliente.Size = new System.Drawing.Size(208, 19);
			this.cmbCliente.TabIndex = 3;
			// 
			// cdsComprobante
			// 
			this.cdsComprobante.BindingContextCtrl = this;
			this.cdsComprobante.DataLibrary = "LibFacturacion";
			this.cdsComprobante.DataLibraryUrl = "";
			this.cdsComprobante.DataSetDef = "dsCompraComprob";
			this.cdsComprobante.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsComprobante.Name = "cdsComprobante";
			this.cdsComprobante.SchemaClassName = "LibFacturacion.DataClass";
			this.cdsComprobante.SchemaDef = null;
			// 
			// cdvArtCodigo
			// 
			this.cdvArtCodigo.BindingContextCtrl = this;
			this.cdvArtCodigo.DataSet = this.cdsArticulo;
			this.cdvArtCodigo.Sort = "Codigo";
			this.cdvArtCodigo.TableName = "";
			this.cdvArtCodigo.TableViewName = "Articulo";
			// 
			// cdvNumero
			// 
			this.cdvNumero.BindingContextCtrl = this;
			this.cdvNumero.DataSet = this.cdsNumero;
			this.cdvNumero.TableName = "";
			this.cdvNumero.TableViewName = "CompraNumero";
			// 
			// cdsNumero
			// 
			this.cdsNumero.BindingContextCtrl = this;
			this.cdsNumero.DataLibrary = "LibFacturacion";
			this.cdsNumero.DataLibraryUrl = "";
			this.cdsNumero.DataSetDef = "dsCompraNumero";
			this.cdsNumero.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsNumero.Name = "cdsNumero";
			this.cdsNumero.SchemaClassName = "LibFacturacion.DataClass";
			this.cdsNumero.SchemaDef = null;
			// 
			// cdvComprobante
			// 
			this.cdvComprobante.BindingContextCtrl = this;
			this.cdvComprobante.DataSet = this.cdsComprobante;
			this.cdvComprobante.TableName = "";
			this.cdvComprobante.TableViewName = "CompraComprobante";
			// 
			// cmbComprobante
			// 
			this.cmbComprobante.DataBindings.Add(new System.Windows.Forms.Binding("EditValue", this.cdsCompra, "Compra.idComprobante"));
			this.cmbComprobante.Location = new System.Drawing.Point(16, 32);
			this.cmbComprobante.Name = "cmbComprobante";
			// 
			// cmbComprobante.Properties
			// 
			this.cmbComprobante.Properties.Buttons.AddRange(new DevExpress.XtraEditors.Controls.EditorButton[] {
																																																					 new DevExpress.XtraEditors.Controls.EditorButton(DevExpress.XtraEditors.Controls.ButtonPredefines.Combo)});
			this.cmbComprobante.Properties.Columns.AddRange(new DevExpress.XtraEditors.Controls.LookUpColumnInfo[] {
																																																							 new DevExpress.XtraEditors.Controls.LookUpColumnInfo("Comprobante", "Comprobante", 78, DevExpress.Utils.FormatType.None, "", true, DevExpress.Utils.HorzAlignment.Near)});
			this.cmbComprobante.Properties.DataSource = this.cdvComprobante;
			this.cmbComprobante.Properties.DisplayMember = "Comprobante";
			this.cmbComprobante.Properties.NullText = "";
			this.cmbComprobante.Properties.ShowFooter = false;
			this.cmbComprobante.Properties.ShowHeader = false;
			this.cmbComprobante.Properties.ValueMember = "idComprobante";
			this.cmbComprobante.Size = new System.Drawing.Size(208, 19);
			this.cmbComprobante.TabIndex = 4;
			// 
			// cmbTipoCompra
			// 
			this.cmbTipoCompra.DataBindings.Add(new System.Windows.Forms.Binding("EditValue", this.cdsCompra, "Compra.idTipoFactura"));
			this.cmbTipoCompra.Location = new System.Drawing.Point(16, 8);
			this.cmbTipoCompra.Name = "cmbTipoCompra";
			// 
			// cmbTipoCompra.Properties
			// 
			this.cmbTipoCompra.Properties.BorderStyle = DevExpress.XtraEditors.Controls.BorderStyles.NoBorder;
			this.cmbTipoCompra.Properties.Buttons.AddRange(new DevExpress.XtraEditors.Controls.EditorButton[] {
																																																					new DevExpress.XtraEditors.Controls.EditorButton(DevExpress.XtraEditors.Controls.ButtonPredefines.Combo, "", -1, true, false, false, DevExpress.Utils.HorzAlignment.Center, null)});
			this.cmbTipoCompra.Properties.Columns.AddRange(new DevExpress.XtraEditors.Controls.LookUpColumnInfo[] {
																																																							new DevExpress.XtraEditors.Controls.LookUpColumnInfo("Nombre", "Nombre", 49, DevExpress.Utils.FormatType.None, "", true, DevExpress.Utils.HorzAlignment.Near)});
			this.cmbTipoCompra.Properties.DataSource = this.cdvNumero;
			this.cmbTipoCompra.Properties.DisplayMember = "Nombre";
			this.cmbTipoCompra.Properties.NullText = "";
			this.cmbTipoCompra.Properties.ShowFooter = false;
			this.cmbTipoCompra.Properties.ShowHeader = false;
			this.cmbTipoCompra.Properties.Style = new DevExpress.Utils.ViewStyle("ControlStyle", null, new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0))), "", DevExpress.Utils.StyleOptions.StyleEnabled, true, false, false, DevExpress.Utils.HorzAlignment.Default, DevExpress.Utils.VertAlignment.Center, null, System.Drawing.SystemColors.Control, System.Drawing.SystemColors.WindowText);
			this.cmbTipoCompra.Properties.ValueMember = "idTipoFactura";
			this.cmbTipoCompra.Size = new System.Drawing.Size(104, 20);
			this.cmbTipoCompra.TabIndex = 5;
			// 
			// txtNumero
			// 
			this.txtNumero.DataBindings.Add(new System.Windows.Forms.Binding("EditValue", this.cdsCompra, "Compra.Numero"));
			this.txtNumero.EditValue = "";
			this.txtNumero.Location = new System.Drawing.Point(128, 8);
			this.txtNumero.Name = "txtNumero";
			this.txtNumero.Size = new System.Drawing.Size(96, 19);
			this.txtNumero.TabIndex = 6;
			// 
			// dtFecha
			// 
			this.dtFecha.DataBindings.Add(new System.Windows.Forms.Binding("EditValue", this.cdsCompra, "Compra.Fecha"));
			this.dtFecha.EditValue = new System.DateTime(2004, 9, 26, 0, 0, 0, 0);
			this.dtFecha.Location = new System.Drawing.Point(280, 32);
			this.dtFecha.Name = "dtFecha";
			// 
			// dtFecha.Properties
			// 
			this.dtFecha.Properties.Buttons.AddRange(new DevExpress.XtraEditors.Controls.EditorButton[] {
																																																		new DevExpress.XtraEditors.Controls.EditorButton(DevExpress.XtraEditors.Controls.ButtonPredefines.Combo)});
			this.dtFecha.Properties.DisplayFormat.FormatString = "dd/MMM/yyyy";
			this.dtFecha.Properties.DisplayFormat.FormatType = DevExpress.Utils.FormatType.Custom;
			this.dtFecha.Properties.EditFormat.FormatString = "dd/MM/yyyy";
			this.dtFecha.Properties.EditFormat.FormatType = DevExpress.Utils.FormatType.DateTime;
			this.dtFecha.Size = new System.Drawing.Size(96, 19);
			this.dtFecha.TabIndex = 7;
			// 
			// spnDias
			// 
			this.spnDias.DataBindings.Add(new System.Windows.Forms.Binding("EditValue", this.cdsCompra, "Compra.DiasCredito"));
			this.spnDias.EditValue = 0;
			this.spnDias.Location = new System.Drawing.Point(432, 32);
			this.spnDias.Name = "spnDias";
			// 
			// spnDias.Properties
			// 
			this.spnDias.Properties.Buttons.AddRange(new DevExpress.XtraEditors.Controls.EditorButton[] {
																																																		new DevExpress.XtraEditors.Controls.EditorButton()});
			this.spnDias.Properties.UseCtrlIncrement = false;
			this.spnDias.Size = new System.Drawing.Size(56, 19);
			this.spnDias.TabIndex = 8;
			this.spnDias.ToolTip = "Días de Crédito";
			// 
			// tabControl1
			// 
			this.tabControl1.Anchor = (System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left);
			this.tabControl1.Controls.AddRange(new System.Windows.Forms.Control[] {
																																							this.tabPage1,
																																							this.tabPage2,
																																							this.tabPage3,
																																							this.tabPage4,
																																							this.tabPage5,
																																							this.tabPage6,
																																							this.tabPage7,
																																							this.tabPage8});
			this.tabControl1.Location = new System.Drawing.Point(16, 248);
			this.tabControl1.Name = "tabControl1";
			this.tabControl1.SelectedIndex = 0;
			this.tabControl1.Size = new System.Drawing.Size(432, 144);
			this.tabControl1.TabIndex = 9;
			// 
			// tabPage1
			// 
			this.tabPage1.Controls.AddRange(new System.Windows.Forms.Control[] {
																																					 this.label15,
																																					 this.textEdit9,
																																					 this.label14,
																																					 this.label13,
																																					 this.cmbSubProyecto,
																																					 this.memoEdit1,
																																					 this.label5,
																																					 this.textEdit6,
																																					 this.cmbProyecto});
			this.tabPage1.Location = new System.Drawing.Point(4, 22);
			this.tabPage1.Name = "tabPage1";
			this.tabPage1.Size = new System.Drawing.Size(424, 118);
			this.tabPage1.TabIndex = 0;
			this.tabPage1.Text = "Notas";
			// 
			// label15
			// 
			this.label15.AutoSize = true;
			this.label15.Location = new System.Drawing.Point(208, 64);
			this.label15.Name = "label15";
			this.label15.Size = new System.Drawing.Size(73, 13);
			this.label15.TabIndex = 48;
			this.label15.Text = "Comprobante";
			// 
			// textEdit9
			// 
			this.textEdit9.DataBindings.Add(new System.Windows.Forms.Binding("EditValue", this.cdsCompra, "Compra.Comprobante"));
			this.textEdit9.EditValue = "";
			this.textEdit9.Location = new System.Drawing.Point(288, 64);
			this.textEdit9.Name = "textEdit9";
			this.textEdit9.Size = new System.Drawing.Size(120, 19);
			this.textEdit9.TabIndex = 47;
			// 
			// label14
			// 
			this.label14.AutoSize = true;
			this.label14.Location = new System.Drawing.Point(208, 88);
			this.label14.Name = "label14";
			this.label14.Size = new System.Drawing.Size(58, 13);
			this.label14.TabIndex = 46;
			this.label14.Text = "SubCentro";
			// 
			// label13
			// 
			this.label13.AutoSize = true;
			this.label13.Location = new System.Drawing.Point(8, 88);
			this.label13.Name = "label13";
			this.label13.Size = new System.Drawing.Size(38, 13);
			this.label13.TabIndex = 45;
			this.label13.Text = "Centro";
			// 
			// cmbSubProyecto
			// 
			this.cmbSubProyecto.DataBindings.Add(new System.Windows.Forms.Binding("EditValue", this.cdsCompra, "Compra.idSubProyecto"));
			this.cmbSubProyecto.Location = new System.Drawing.Point(288, 88);
			this.cmbSubProyecto.Name = "cmbSubProyecto";
			// 
			// cmbSubProyecto.Properties
			// 
			this.cmbSubProyecto.Properties.Buttons.AddRange(new DevExpress.XtraEditors.Controls.EditorButton[] {
																																																					 new DevExpress.XtraEditors.Controls.EditorButton(DevExpress.XtraEditors.Controls.ButtonPredefines.Combo)});
			this.cmbSubProyecto.Properties.Columns.AddRange(new DevExpress.XtraEditors.Controls.LookUpColumnInfo[] {
																																																							 new DevExpress.XtraEditors.Controls.LookUpColumnInfo("Nombre", "Nombre", 49, DevExpress.Utils.FormatType.None, "", true, DevExpress.Utils.HorzAlignment.Near)});
			this.cmbSubProyecto.Properties.DataSource = this.cdvSubProyecto;
			this.cmbSubProyecto.Properties.DisplayMember = "Nombre";
			this.cmbSubProyecto.Properties.NullText = "";
			this.cmbSubProyecto.Properties.ShowFooter = false;
			this.cmbSubProyecto.Properties.ShowHeader = false;
			this.cmbSubProyecto.Properties.ValueMember = "idSubProyecto";
			this.cmbSubProyecto.Size = new System.Drawing.Size(120, 19);
			this.cmbSubProyecto.TabIndex = 44;
			// 
			// cdvSubProyecto
			// 
			this.cdvSubProyecto.BindingContextCtrl = this;
			this.cdvSubProyecto.DataSet = this.cdsProyecto;
			this.cdvSubProyecto.TableName = "";
			this.cdvSubProyecto.TableViewName = "SubProyecto";
			// 
			// cdsProyecto
			// 
			this.cdsProyecto.BindingContextCtrl = this;
			this.cdsProyecto.DataLibrary = "LibContabilidad";
			this.cdsProyecto.DataLibraryUrl = "";
			this.cdsProyecto.DataSetDef = "dsProyecto";
			this.cdsProyecto.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsProyecto.Name = "cdsProyecto";
			this.cdsProyecto.SchemaClassName = "LibContabilidad.DataClass";
			this.cdsProyecto.SchemaDef = null;
			// 
			// memoEdit1
			// 
			this.memoEdit1.DataBindings.Add(new System.Windows.Forms.Binding("EditValue", this.cdsCompra, "Compra.Notas"));
			this.memoEdit1.EditValue = "";
			this.memoEdit1.Location = new System.Drawing.Point(8, 8);
			this.memoEdit1.Name = "memoEdit1";
			this.memoEdit1.Size = new System.Drawing.Size(400, 48);
			this.memoEdit1.TabIndex = 0;
			// 
			// label5
			// 
			this.label5.AutoSize = true;
			this.label5.Location = new System.Drawing.Point(8, 64);
			this.label5.Name = "label5";
			this.label5.Size = new System.Drawing.Size(53, 13);
			this.label5.TabIndex = 30;
			this.label5.Text = "Vendedor";
			// 
			// textEdit6
			// 
			this.textEdit6.DataBindings.Add(new System.Windows.Forms.Binding("EditValue", this.cdsCompra, "Compra.Otro"));
			this.textEdit6.EditValue = "";
			this.textEdit6.Location = new System.Drawing.Point(72, 64);
			this.textEdit6.Name = "textEdit6";
			this.textEdit6.Size = new System.Drawing.Size(120, 19);
			this.textEdit6.TabIndex = 30;
			// 
			// cmbProyecto
			// 
			this.cmbProyecto.DataBindings.Add(new System.Windows.Forms.Binding("EditValue", this.cdsCompra, "Compra.idProyecto"));
			this.cmbProyecto.Location = new System.Drawing.Point(72, 88);
			this.cmbProyecto.Name = "cmbProyecto";
			// 
			// cmbProyecto.Properties
			// 
			this.cmbProyecto.Properties.Buttons.AddRange(new DevExpress.XtraEditors.Controls.EditorButton[] {
																																																				new DevExpress.XtraEditors.Controls.EditorButton(DevExpress.XtraEditors.Controls.ButtonPredefines.Combo)});
			this.cmbProyecto.Properties.Columns.AddRange(new DevExpress.XtraEditors.Controls.LookUpColumnInfo[] {
																																																						new DevExpress.XtraEditors.Controls.LookUpColumnInfo("Nombre", "Nombre", 49, DevExpress.Utils.FormatType.None, "", true, DevExpress.Utils.HorzAlignment.Near)});
			this.cmbProyecto.Properties.DataSource = this.cdvProyecto;
			this.cmbProyecto.Properties.DisplayMember = "Nombre";
			this.cmbProyecto.Properties.NullText = "";
			this.cmbProyecto.Properties.ShowFooter = false;
			this.cmbProyecto.Properties.ShowHeader = false;
			this.cmbProyecto.Properties.ValueMember = "idProyecto";
			this.cmbProyecto.Size = new System.Drawing.Size(120, 19);
			this.cmbProyecto.TabIndex = 43;
			// 
			// cdvProyecto
			// 
			this.cdvProyecto.BindingContextCtrl = this;
			this.cdvProyecto.DataSet = this.cdsProyecto;
			this.cdvProyecto.TableName = "";
			this.cdvProyecto.TableViewName = "Proyecto";
			// 
			// tabPage2
			// 
			this.tabPage2.Controls.AddRange(new System.Windows.Forms.Control[] {
																																					 this.cmbFormaPago,
																																					 this.btPagos,
																																					 this.rdgContadiCred});
			this.tabPage2.Location = new System.Drawing.Point(4, 22);
			this.tabPage2.Name = "tabPage2";
			this.tabPage2.Size = new System.Drawing.Size(424, 118);
			this.tabPage2.TabIndex = 1;
			this.tabPage2.Text = "Pago";
			// 
			// cmbFormaPago
			// 
			this.cmbFormaPago.DataBindings.Add(new System.Windows.Forms.Binding("EditValue", this.cdsCompra, "Compra.idProyecto"));
			this.cmbFormaPago.Location = new System.Drawing.Point(24, 50);
			this.cmbFormaPago.Name = "cmbFormaPago";
			// 
			// cmbFormaPago.Properties
			// 
			this.cmbFormaPago.Properties.Buttons.AddRange(new DevExpress.XtraEditors.Controls.EditorButton[] {
																																																				 new DevExpress.XtraEditors.Controls.EditorButton(DevExpress.XtraEditors.Controls.ButtonPredefines.Combo)});
			this.cmbFormaPago.Properties.Columns.AddRange(new DevExpress.XtraEditors.Controls.LookUpColumnInfo[] {
																																																						 new DevExpress.XtraEditors.Controls.LookUpColumnInfo("FormaPago", "FormaPago", 68, DevExpress.Utils.FormatType.None, "", true, DevExpress.Utils.HorzAlignment.Near)});
			this.cmbFormaPago.Properties.DataSource = this.cdvFormaPago;
			this.cmbFormaPago.Properties.DisplayMember = "FormaPago";
			this.cmbFormaPago.Properties.NullText = "";
			this.cmbFormaPago.Properties.ShowFooter = false;
			this.cmbFormaPago.Properties.ShowHeader = false;
			this.cmbFormaPago.Properties.ValueMember = "idFormaPago";
			this.cmbFormaPago.Size = new System.Drawing.Size(152, 19);
			this.cmbFormaPago.TabIndex = 44;
			// 
			// cdvFormaPago
			// 
			this.cdvFormaPago.BindingContextCtrl = this;
			this.cdvFormaPago.DataSet = this.cdsFormPago;
			this.cdvFormaPago.TableName = "";
			this.cdvFormaPago.TableViewName = "CompraFormaPago";
			// 
			// cdsFormPago
			// 
			this.cdsFormPago.BindingContextCtrl = this;
			this.cdsFormPago.DataLibrary = "LibFacturacion";
			this.cdsFormPago.DataLibraryUrl = "";
			this.cdsFormPago.DataSetDef = "dsCompraFormaPago";
			this.cdsFormPago.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsFormPago.Name = "cdsFormPago";
			this.cdsFormPago.SchemaClassName = "LibFacturacion.DataClass";
			this.cdsFormPago.SchemaDef = null;
			// 
			// btPagos
			// 
			this.btPagos.Location = new System.Drawing.Point(200, 8);
			this.btPagos.Name = "btPagos";
			this.btPagos.TabIndex = 1;
			this.btPagos.Text = "Pagos";
			// 
			// rdgContadiCred
			// 
			this.rdgContadiCred.DataBindings.Add(new System.Windows.Forms.Binding("EditValue", this.cdsCompra, "Compra.ContadoCredito"));
			this.rdgContadiCred.Location = new System.Drawing.Point(24, 8);
			this.rdgContadiCred.Name = "rdgContadiCred";
			// 
			// rdgContadiCred.Properties
			// 
			this.rdgContadiCred.Properties.Items.AddRange(new DevExpress.XtraEditors.Controls.RadioGroupItem[] {
																																																					 new DevExpress.XtraEditors.Controls.RadioGroupItem(1, "Contado"),
																																																					 new DevExpress.XtraEditors.Controls.RadioGroupItem(2, "Crédito")});
			this.rdgContadiCred.Size = new System.Drawing.Size(152, 32);
			this.rdgContadiCred.TabIndex = 0;
			// 
			// tabPage3
			// 
			this.tabPage3.Controls.AddRange(new System.Windows.Forms.Control[] {
																																					 this.textEdit10,
																																					 this.label17,
																																					 this.txtDescIva0,
																																					 this.txtDescPorcIva,
																																					 this.label16,
																																					 this.txtDescIva});
			this.tabPage3.Location = new System.Drawing.Point(4, 22);
			this.tabPage3.Name = "tabPage3";
			this.tabPage3.Size = new System.Drawing.Size(424, 118);
			this.tabPage3.TabIndex = 2;
			this.tabPage3.Text = "Descuento";
			// 
			// textEdit10
			// 
			this.textEdit10.DataBindings.Add(new System.Windows.Forms.Binding("EditValue", this.cdsCompra, "Compra.DescPorcIva0"));
			this.textEdit10.EditValue = "";
			this.textEdit10.Location = new System.Drawing.Point(240, 40);
			this.textEdit10.Name = "textEdit10";
			this.textEdit10.Size = new System.Drawing.Size(64, 19);
			this.textEdit10.TabIndex = 36;
			// 
			// label17
			// 
			this.label17.AutoSize = true;
			this.label17.Location = new System.Drawing.Point(16, 40);
			this.label17.Name = "label17";
			this.label17.Size = new System.Drawing.Size(89, 13);
			this.label17.TabIndex = 35;
			this.label17.Text = "Descuento IVA 0";
			// 
			// txtDescIva0
			// 
			this.txtDescIva0.DataBindings.Add(new System.Windows.Forms.Binding("EditValue", this.cdsCompra, "Compra.Desc2"));
			this.txtDescIva0.EditValue = "";
			this.txtDescIva0.Location = new System.Drawing.Point(120, 40);
			this.txtDescIva0.Name = "txtDescIva0";
			this.txtDescIva0.Size = new System.Drawing.Size(104, 19);
			this.txtDescIva0.TabIndex = 34;
			// 
			// txtDescPorcIva
			// 
			this.txtDescPorcIva.DataBindings.Add(new System.Windows.Forms.Binding("EditValue", this.cdsCompra, "Compra.DescPorcIva"));
			this.txtDescPorcIva.EditValue = "";
			this.txtDescPorcIva.Location = new System.Drawing.Point(240, 16);
			this.txtDescPorcIva.Name = "txtDescPorcIva";
			this.txtDescPorcIva.Size = new System.Drawing.Size(64, 19);
			this.txtDescPorcIva.TabIndex = 33;
			// 
			// label16
			// 
			this.label16.AutoSize = true;
			this.label16.Location = new System.Drawing.Point(16, 16);
			this.label16.Name = "label16";
			this.label16.Size = new System.Drawing.Size(80, 13);
			this.label16.TabIndex = 32;
			this.label16.Text = "Descuento IVA";
			// 
			// txtDescIva
			// 
			this.txtDescIva.DataBindings.Add(new System.Windows.Forms.Binding("EditValue", this.cdsCompra, "Compra.Desc1"));
			this.txtDescIva.EditValue = "";
			this.txtDescIva.Location = new System.Drawing.Point(120, 16);
			this.txtDescIva.Name = "txtDescIva";
			this.txtDescIva.Size = new System.Drawing.Size(104, 19);
			this.txtDescIva.TabIndex = 31;
			// 
			// tabPage4
			// 
			this.tabPage4.Location = new System.Drawing.Point(4, 22);
			this.tabPage4.Name = "tabPage4";
			this.tabPage4.Size = new System.Drawing.Size(424, 118);
			this.tabPage4.TabIndex = 3;
			this.tabPage4.Text = "Cuota";
			// 
			// tabPage5
			// 
			this.tabPage5.Location = new System.Drawing.Point(4, 22);
			this.tabPage5.Name = "tabPage5";
			this.tabPage5.Size = new System.Drawing.Size(424, 118);
			this.tabPage5.TabIndex = 4;
			this.tabPage5.Text = "Autoriz.";
			// 
			// tabPage6
			// 
			this.tabPage6.Location = new System.Drawing.Point(4, 22);
			this.tabPage6.Name = "tabPage6";
			this.tabPage6.Size = new System.Drawing.Size(424, 118);
			this.tabPage6.TabIndex = 5;
			this.tabPage6.Text = "COA";
			// 
			// tabPage7
			// 
			this.tabPage7.Location = new System.Drawing.Point(4, 22);
			this.tabPage7.Name = "tabPage7";
			this.tabPage7.Size = new System.Drawing.Size(424, 118);
			this.tabPage7.TabIndex = 6;
			this.tabPage7.Text = "Packing";
			// 
			// tabPage8
			// 
			this.tabPage8.Location = new System.Drawing.Point(4, 22);
			this.tabPage8.Name = "tabPage8";
			this.tabPage8.Size = new System.Drawing.Size(424, 118);
			this.tabPage8.TabIndex = 7;
			this.tabPage8.Text = "Otros";
			// 
			// dateEdit1
			// 
			this.dateEdit1.DataBindings.Add(new System.Windows.Forms.Binding("EditValue", this.cdsCompra, "Compra.FechaVencimiento"));
			this.dateEdit1.EditValue = new System.DateTime(2004, 9, 26, 0, 0, 0, 0);
			this.dateEdit1.Location = new System.Drawing.Point(552, 32);
			this.dateEdit1.Name = "dateEdit1";
			// 
			// dateEdit1.Properties
			// 
			this.dateEdit1.Properties.Buttons.AddRange(new DevExpress.XtraEditors.Controls.EditorButton[] {
																																																			new DevExpress.XtraEditors.Controls.EditorButton(DevExpress.XtraEditors.Controls.ButtonPredefines.Combo)});
			this.dateEdit1.Properties.DisplayFormat.FormatString = "dd/MMM/yyyy";
			this.dateEdit1.Properties.DisplayFormat.FormatType = DevExpress.Utils.FormatType.Custom;
			this.dateEdit1.Properties.EditFormat.FormatString = "dd/MM/yyyy";
			this.dateEdit1.Properties.EditFormat.FormatType = DevExpress.Utils.FormatType.DateTime;
			this.dateEdit1.Size = new System.Drawing.Size(96, 19);
			this.dateEdit1.TabIndex = 10;
			// 
			// spnBodega
			// 
			this.spnBodega.DataBindings.Add(new System.Windows.Forms.Binding("EditValue", this.cdsCompra, "Compra.Bodega"));
			this.spnBodega.EditValue = 0;
			this.spnBodega.Location = new System.Drawing.Point(600, 8);
			this.spnBodega.Name = "spnBodega";
			// 
			// spnBodega.Properties
			// 
			this.spnBodega.Properties.Buttons.AddRange(new DevExpress.XtraEditors.Controls.EditorButton[] {
																																																			new DevExpress.XtraEditors.Controls.EditorButton()});
			this.spnBodega.Properties.UseCtrlIncrement = false;
			this.spnBodega.Size = new System.Drawing.Size(48, 19);
			this.spnBodega.TabIndex = 11;
			// 
			// textEdit1
			// 
			this.textEdit1.Anchor = (System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right);
			this.textEdit1.DataBindings.Add(new System.Windows.Forms.Binding("EditValue", this.cdsCompra, "Compra.SubtotalExcento"));
			this.textEdit1.EditValue = "";
			this.textEdit1.Location = new System.Drawing.Point(528, 264);
			this.textEdit1.Name = "textEdit1";
			// 
			// textEdit1.Properties
			// 
			this.textEdit1.Properties.DisplayFormat.FormatString = "#,##0.00";
			this.textEdit1.Properties.DisplayFormat.FormatType = DevExpress.Utils.FormatType.Custom;
			this.textEdit1.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
			this.textEdit1.Size = new System.Drawing.Size(64, 19);
			this.textEdit1.TabIndex = 12;
			// 
			// textEdit2
			// 
			this.textEdit2.Anchor = (System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right);
			this.textEdit2.DataBindings.Add(new System.Windows.Forms.Binding("EditValue", this.cdsCompra, "Compra.SubtotalIva"));
			this.textEdit2.EditValue = "";
			this.textEdit2.Location = new System.Drawing.Point(600, 264);
			this.textEdit2.Name = "textEdit2";
			// 
			// textEdit2.Properties
			// 
			this.textEdit2.Properties.DisplayFormat.FormatString = "#,##0.00";
			this.textEdit2.Properties.DisplayFormat.FormatType = DevExpress.Utils.FormatType.Custom;
			this.textEdit2.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
			this.textEdit2.Size = new System.Drawing.Size(72, 19);
			this.textEdit2.TabIndex = 13;
			// 
			// label1
			// 
			this.label1.Anchor = (System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right);
			this.label1.AutoSize = true;
			this.label1.Location = new System.Drawing.Point(552, 248);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(32, 13);
			this.label1.TabIndex = 14;
			this.label1.Text = "IVA 0";
			// 
			// label2
			// 
			this.label2.Anchor = (System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right);
			this.label2.AutoSize = true;
			this.label2.Location = new System.Drawing.Point(632, 248);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(22, 13);
			this.label2.TabIndex = 15;
			this.label2.Text = "IVA";
			// 
			// label3
			// 
			this.label3.Anchor = (System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right);
			this.label3.AutoSize = true;
			this.label3.Location = new System.Drawing.Point(464, 312);
			this.label3.Name = "label3";
			this.label3.Size = new System.Drawing.Size(22, 13);
			this.label3.TabIndex = 17;
			this.label3.Text = "IVA";
			// 
			// label4
			// 
			this.label4.Anchor = (System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right);
			this.label4.AutoSize = true;
			this.label4.Location = new System.Drawing.Point(464, 288);
			this.label4.Name = "label4";
			this.label4.Size = new System.Drawing.Size(58, 13);
			this.label4.TabIndex = 16;
			this.label4.Text = "Descuento";
			// 
			// btSaldo
			// 
			this.btSaldo.Anchor = (System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left);
			this.btSaldo.ImageIndex = 5;
			this.btSaldo.ImageList = this.imageList;
			this.btSaldo.Location = new System.Drawing.Point(336, 400);
			this.btSaldo.Name = "btSaldo";
			this.btSaldo.Size = new System.Drawing.Size(24, 23);
			this.btSaldo.TabIndex = 18;
			this.btSaldo.Text = "simpleButton1";
			this.btSaldo.ToolTip = "Saldo de Artículo seleccionado";
			// 
			// imageList
			// 
			this.imageList.ColorDepth = System.Windows.Forms.ColorDepth.Depth8Bit;
			this.imageList.ImageSize = new System.Drawing.Size(16, 16);
			this.imageList.ImageStream = ((System.Windows.Forms.ImageListStreamer)(resources.GetObject("imageList.ImageStream")));
			this.imageList.TransparentColor = System.Drawing.Color.Transparent;
			// 
			// errorProvider
			// 
			this.errorProvider.DataMember = null;
			// 
			// simpleButton2
			// 
			this.simpleButton2.Anchor = (System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right);
			this.simpleButton2.Location = new System.Drawing.Point(464, 336);
			this.simpleButton2.Name = "simpleButton2";
			this.simpleButton2.Size = new System.Drawing.Size(40, 23);
			this.simpleButton2.TabIndex = 20;
			this.simpleButton2.Text = "Total";
			// 
			// btNotaEntrega
			// 
			this.btNotaEntrega.Anchor = (System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left);
			this.btNotaEntrega.ImageIndex = 7;
			this.btNotaEntrega.ImageList = this.imageList;
			this.btNotaEntrega.Location = new System.Drawing.Point(480, 400);
			this.btNotaEntrega.Name = "btNotaEntrega";
			this.btNotaEntrega.Size = new System.Drawing.Size(24, 23);
			this.btNotaEntrega.TabIndex = 21;
			this.btNotaEntrega.Text = "simpleButton3";
			this.btNotaEntrega.ToolTip = "Impresion de nota de entrega";
			// 
			// btGuiaRemision
			// 
			this.btGuiaRemision.Anchor = (System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left);
			this.btGuiaRemision.ImageIndex = 8;
			this.btGuiaRemision.ImageList = this.imageList;
			this.btGuiaRemision.Location = new System.Drawing.Point(456, 400);
			this.btGuiaRemision.Name = "btGuiaRemision";
			this.btGuiaRemision.Size = new System.Drawing.Size(24, 23);
			this.btGuiaRemision.TabIndex = 22;
			this.btGuiaRemision.Text = "simpleButton4";
			this.btGuiaRemision.ToolTip = "Guia de Remisión";
			// 
			// btCreaAsiento
			// 
			this.btCreaAsiento.Anchor = (System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left);
			this.btCreaAsiento.ImageIndex = 9;
			this.btCreaAsiento.ImageList = this.imageList;
			this.btCreaAsiento.Location = new System.Drawing.Point(432, 400);
			this.btCreaAsiento.Name = "btCreaAsiento";
			this.btCreaAsiento.Size = new System.Drawing.Size(24, 23);
			this.btCreaAsiento.TabIndex = 23;
			this.btCreaAsiento.Text = "simpleButton5";
			this.btCreaAsiento.ToolTip = "Creación de asiento";
			// 
			// btRetencion
			// 
			this.btRetencion.Anchor = (System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left);
			this.btRetencion.ImageIndex = 0;
			this.btRetencion.ImageList = this.imageList;
			this.btRetencion.Location = new System.Drawing.Point(408, 400);
			this.btRetencion.Name = "btRetencion";
			this.btRetencion.Size = new System.Drawing.Size(24, 23);
			this.btRetencion.TabIndex = 24;
			this.btRetencion.Text = "simpleButton6";
			this.btRetencion.ToolTip = "Retención en la fuente";
			// 
			// btTransforma
			// 
			this.btTransforma.Anchor = (System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left);
			this.btTransforma.ImageIndex = 3;
			this.btTransforma.ImageList = this.imageList;
			this.btTransforma.Location = new System.Drawing.Point(384, 400);
			this.btTransforma.Name = "btTransforma";
			this.btTransforma.Size = new System.Drawing.Size(24, 23);
			this.btTransforma.TabIndex = 25;
			this.btTransforma.Text = "simpleButton7";
			this.btTransforma.ToolTip = "Transformacion de Factura";
			// 
			// btBuscaArticulo
			// 
			this.btBuscaArticulo.Anchor = (System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left);
			this.btBuscaArticulo.ImageIndex = 6;
			this.btBuscaArticulo.ImageList = this.imageList;
			this.btBuscaArticulo.Location = new System.Drawing.Point(360, 400);
			this.btBuscaArticulo.Name = "btBuscaArticulo";
			this.btBuscaArticulo.Size = new System.Drawing.Size(24, 23);
			this.btBuscaArticulo.TabIndex = 26;
			this.btBuscaArticulo.Text = "simpleButton8";
			this.btBuscaArticulo.ToolTip = "Busqueda de articulo seleccionado";
			// 
			// textEdit3
			// 
			this.textEdit3.Anchor = (System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right);
			this.textEdit3.DataBindings.Add(new System.Windows.Forms.Binding("EditValue", this.cdsCompra, "Compra.DescPorcIva"));
			this.textEdit3.EditValue = "";
			this.textEdit3.Location = new System.Drawing.Point(600, 288);
			this.textEdit3.Name = "textEdit3";
			// 
			// textEdit3.Properties
			// 
			this.textEdit3.Properties.DisplayFormat.FormatString = "#,##0.00";
			this.textEdit3.Properties.DisplayFormat.FormatType = DevExpress.Utils.FormatType.Custom;
			this.textEdit3.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
			this.textEdit3.Size = new System.Drawing.Size(72, 19);
			this.textEdit3.TabIndex = 27;
			// 
			// textEdit4
			// 
			this.textEdit4.Anchor = (System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right);
			this.textEdit4.DataBindings.Add(new System.Windows.Forms.Binding("EditValue", this.cdsCompra, "Compra.Iva"));
			this.textEdit4.EditValue = "";
			this.textEdit4.Location = new System.Drawing.Point(512, 312);
			this.textEdit4.Name = "textEdit4";
			// 
			// textEdit4.Properties
			// 
			this.textEdit4.Properties.DisplayFormat.FormatString = "#,##0.00";
			this.textEdit4.Properties.DisplayFormat.FormatType = DevExpress.Utils.FormatType.Custom;
			this.textEdit4.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
			this.textEdit4.Size = new System.Drawing.Size(56, 19);
			this.textEdit4.TabIndex = 28;
			// 
			// textEdit5
			// 
			this.textEdit5.Anchor = (System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right);
			this.textEdit5.DataBindings.Add(new System.Windows.Forms.Binding("EditValue", this.cdsCompra, "Compra.Total"));
			this.textEdit5.EditValue = "";
			this.textEdit5.Location = new System.Drawing.Point(512, 336);
			this.textEdit5.Name = "textEdit5";
			// 
			// textEdit5.Properties
			// 
			this.textEdit5.Properties.DisplayFormat.FormatString = "#,##0.00";
			this.textEdit5.Properties.DisplayFormat.FormatType = DevExpress.Utils.FormatType.Custom;
			this.textEdit5.Properties.EditFormat.FormatString = "#,##0.00";
			this.textEdit5.Properties.EditFormat.FormatType = DevExpress.Utils.FormatType.Custom;
			this.textEdit5.Properties.Style = new DevExpress.Utils.ViewStyle("ControlStyle", null, new System.Drawing.Font("Microsoft Sans Serif", 21.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0))), "", DevExpress.Utils.StyleOptions.StyleEnabled, true, false, false, DevExpress.Utils.HorzAlignment.Default, DevExpress.Utils.VertAlignment.Center, null, System.Drawing.SystemColors.Window, System.Drawing.SystemColors.WindowText);
			this.textEdit5.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
			this.textEdit5.Size = new System.Drawing.Size(160, 42);
			this.textEdit5.TabIndex = 29;
			// 
			// label6
			// 
			this.label6.AutoSize = true;
			this.label6.Location = new System.Drawing.Point(232, 32);
			this.label6.Name = "label6";
			this.label6.Size = new System.Drawing.Size(36, 13);
			this.label6.TabIndex = 30;
			this.label6.Text = "Fecha";
			// 
			// label7
			// 
			this.label7.AutoSize = true;
			this.label7.Location = new System.Drawing.Point(384, 32);
			this.label7.Name = "label7";
			this.label7.Size = new System.Drawing.Size(27, 13);
			this.label7.TabIndex = 31;
			this.label7.Text = "Días";
			// 
			// label8
			// 
			this.label8.AutoSize = true;
			this.label8.Location = new System.Drawing.Point(496, 32);
			this.label8.Name = "label8";
			this.label8.Size = new System.Drawing.Size(39, 13);
			this.label8.TabIndex = 32;
			this.label8.Text = "Vence:";
			// 
			// label9
			// 
			this.label9.AutoSize = true;
			this.label9.Location = new System.Drawing.Point(544, 8);
			this.label9.Name = "label9";
			this.label9.Size = new System.Drawing.Size(43, 13);
			this.label9.TabIndex = 33;
			this.label9.Text = "Bodega";
			// 
			// btCliente
			// 
			this.btCliente.Location = new System.Drawing.Point(496, 8);
			this.btCliente.Name = "btCliente";
			this.btCliente.Size = new System.Drawing.Size(24, 23);
			this.btCliente.TabIndex = 34;
			this.btCliente.Text = "...";
			// 
			// label11
			// 
			this.label11.AutoSize = true;
			this.label11.Location = new System.Drawing.Point(232, 8);
			this.label11.Name = "label11";
			this.label11.Size = new System.Drawing.Size(43, 13);
			this.label11.TabIndex = 36;
			this.label11.Text = "Cliente:";
			// 
			// label10
			// 
			this.label10.Anchor = (System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right);
			this.label10.AutoSize = true;
			this.label10.Location = new System.Drawing.Point(464, 264);
			this.label10.Name = "label10";
			this.label10.Size = new System.Drawing.Size(50, 13);
			this.label10.TabIndex = 37;
			this.label10.Text = "SubTotal";
			// 
			// textEdit7
			// 
			this.textEdit7.Anchor = (System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right);
			this.textEdit7.DataBindings.Add(new System.Windows.Forms.Binding("EditValue", this.cdsCompra, "Compra.Ice"));
			this.textEdit7.EditValue = "";
			this.textEdit7.Location = new System.Drawing.Point(616, 312);
			this.textEdit7.Name = "textEdit7";
			// 
			// textEdit7.Properties
			// 
			this.textEdit7.Properties.DisplayFormat.FormatString = "#,##0.00";
			this.textEdit7.Properties.DisplayFormat.FormatType = DevExpress.Utils.FormatType.Custom;
			this.textEdit7.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
			this.textEdit7.Size = new System.Drawing.Size(56, 19);
			this.textEdit7.TabIndex = 39;
			// 
			// label12
			// 
			this.label12.Anchor = (System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right);
			this.label12.AutoSize = true;
			this.label12.Location = new System.Drawing.Point(584, 312);
			this.label12.Name = "label12";
			this.label12.Size = new System.Drawing.Size(23, 13);
			this.label12.TabIndex = 38;
			this.label12.Text = "ICE";
			// 
			// textEdit8
			// 
			this.textEdit8.Anchor = (System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right);
			this.textEdit8.DataBindings.Add(new System.Windows.Forms.Binding("EditValue", this.cdsCompra, "Compra.DescPorcIva0"));
			this.textEdit8.EditValue = "";
			this.textEdit8.Location = new System.Drawing.Point(528, 288);
			this.textEdit8.Name = "textEdit8";
			// 
			// textEdit8.Properties
			// 
			this.textEdit8.Properties.DisplayFormat.FormatString = "#,##0.00";
			this.textEdit8.Properties.DisplayFormat.FormatType = DevExpress.Utils.FormatType.Custom;
			this.textEdit8.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
			this.textEdit8.Size = new System.Drawing.Size(64, 19);
			this.textEdit8.TabIndex = 40;
			// 
			// cdsSeguridad
			// 
			this.cdsSeguridad.BindingContextCtrl = this;
			this.cdsSeguridad.DataLibrary = "LibContabilidad";
			this.cdsSeguridad.DataLibraryUrl = "";
			this.cdsSeguridad.DataSetDef = "dsUsuario";
			this.cdsSeguridad.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsSeguridad.Name = "cdsSeguridad";
			this.cdsSeguridad.SchemaClassName = "LibContabilidad.DataClass";
			this.cdsSeguridad.SchemaDef = null;
			// 
			// lblAnulado
			// 
			this.lblAnulado.Anchor = ((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right);
			this.lblAnulado.DataBindings.Add(new System.Windows.Forms.Binding("Visible", this.cdsCompra, "Compra.Borrar"));
			this.lblAnulado.Font = new System.Drawing.Font("Microsoft Sans Serif", 15F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblAnulado.ForeColor = System.Drawing.Color.Firebrick;
			this.lblAnulado.Location = new System.Drawing.Point(544, 400);
			this.lblAnulado.Name = "lblAnulado";
			this.lblAnulado.Size = new System.Drawing.Size(120, 23);
			this.lblAnulado.TabIndex = 41;
			this.lblAnulado.Text = "ANULADO";
			// 
			// txtBorrar
			// 
			this.txtBorrar.DataBindings.Add(new System.Windows.Forms.Binding("EditValue", this.cdsCompra, "Compra.Borrar"));
			this.txtBorrar.EditValue = "Borrar";
			this.txtBorrar.Location = new System.Drawing.Point(608, 376);
			this.txtBorrar.Name = "txtBorrar";
			this.txtBorrar.Size = new System.Drawing.Size(75, 19);
			this.txtBorrar.TabIndex = 42;
			// 
			// Compra1
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(688, 429);
			this.Controls.AddRange(new System.Windows.Forms.Control[] {
																																	this.txtBorrar,
																																	this.lblAnulado,
																																	this.textEdit8,
																																	this.textEdit7,
																																	this.label12,
																																	this.label10,
																																	this.label11,
																																	this.label9,
																																	this.label8,
																																	this.label7,
																																	this.label6,
																																	this.label3,
																																	this.label4,
																																	this.label2,
																																	this.label1,
																																	this.btCliente,
																																	this.textEdit5,
																																	this.textEdit4,
																																	this.textEdit3,
																																	this.btBuscaArticulo,
																																	this.btTransforma,
																																	this.btRetencion,
																																	this.btCreaAsiento,
																																	this.btGuiaRemision,
																																	this.btNotaEntrega,
																																	this.simpleButton2,
																																	this.btSaldo,
																																	this.textEdit2,
																																	this.textEdit1,
																																	this.spnBodega,
																																	this.dateEdit1,
																																	this.tabControl1,
																																	this.spnDias,
																																	this.dtFecha,
																																	this.txtNumero,
																																	this.cmbTipoCompra,
																																	this.cmbComprobante,
																																	this.cmbCliente,
																																	this.barraDato1,
																																	this.txtIdCompra,
																																	this.gridControl1});
			this.Name = "Compra1";
			this.Text = "Compra";
			this.Load += new System.EventHandler(this.Compra2_Load);
			((System.ComponentModel.ISupportInitialize)(this.gridControl1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvDetCompra)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsCompra)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.gridView1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.repositoryItemLookUpEdit1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvArticulo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsArticulo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.repositoryItemLookUpEdit2)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdCompra.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsCliente)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvCliente)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCliente.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsComprobante)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvArtCodigo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvNumero)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsNumero)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvComprobante)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbComprobante.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbTipoCompra.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNumero.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFecha.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.spnDias.Properties)).EndInit();
			this.tabControl1.ResumeLayout(false);
			this.tabPage1.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.textEdit9.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbSubProyecto.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvSubProyecto)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsProyecto)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.memoEdit1.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.textEdit6.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbProyecto.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvProyecto)).EndInit();
			this.tabPage2.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.cmbFormaPago.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvFormaPago)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsFormPago)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.rdgContadiCred.Properties)).EndInit();
			this.tabPage3.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.textEdit10.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDescIva0.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDescPorcIva.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDescIva.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dateEdit1.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.spnBodega.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.textEdit1.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.textEdit2.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.textEdit3.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.textEdit4.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.textEdit5.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.textEdit7.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.textEdit8.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeguridad)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtBorrar.Properties)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void cdvDetCompra_PositionChanged(object sender, C1.Data.PositionChangeEventArgs e)
		{
			if (e.Row == null) return;
			if ((int) e.Row["idDetCompra"] < 0)
			{
				e.Row["Signo"] = IdSigno;
				e.Row["Bodega"] = (int) this.spnBodega.EditValue;
				e.Row["idCompra"] = (int) this.txtIdCompra.EditValue;
			}		
		}

		private void barraDato1_Crear(object sender, System.EventArgs e)
		{
			this.barraDato1.CrearRegistro(false);
			C1DataRow dr = this.cdsCompra.TableViews["Compra"].AddNew();
			dr["idCliente"] = 1;
			BindingContext[cdsCompra,"Compra"].Position = 
				(this.cdsCompra.Tables["Compra"].Rows.Count-1);
			dr = this.cdsCompra.TableViews["DetCompra"].AddNew();
			dr["idCompra"] = (int) this.txtIdCompra.EditValue;
		}

		private void barraDato1_Grabar(object sender, System.EventArgs e)
		{
			this.barraDato1.Focus();
			C1DataRow dr = this.cdsCompra.TableViews["DetCompra"].AddNew();
			dr.Delete();
			this.barraDato1.GrabaRegistro();
			this.cdvDetCompra.RowFilter = "idCompra = " + txtIdCompra.EditValue.ToString()
				+ " Or idDetCompra <= 0";
		}


		private void txtIdCompra_EditValueChanged(object sender, System.EventArgs e)
		{
			return;
		}

		private void barraDato1_Editar(object sender, System.EventArgs e)
		{
			this.barraDato1.EditarRegistro();
		}

		private void barraDato1_Deshacer(object sender, System.EventArgs e)
		{
			this.barraDato1.DeshacerRegistro();
		}

		private void barraDato1_Borrar(object sender, System.EventArgs e)
		{
			if (!lblAnulado.Visible)
			{
				if (!miAcceso.Anular)
				{
					MessageBox.Show("No tiene acceso a Anular Facturas");
					return;
				}
				if (DialogResult.No == MessageBox.Show(
					"¿Está seguro que desea Anular la Factura?", 
					"Confirmación", MessageBoxButtons.YesNo, MessageBoxIcon.Warning, 
					MessageBoxDefaultButton.Button2)) return;

				this.txtBorrar.EditValue = true;
				this.cdsCompra.Update();
			}
			else // Si esta anulado se elimina
			{
				if (!miAcceso.Eliminar)
				{
					MessageBox.Show("No tiene acceso a Eliminar Registros");
					return;
				}
				if (DialogResult.No == MessageBox.Show(
					"¿Está seguro que desea borrar la Factura?"
					+ ".\nEste proceso no se puede reversar", 
					"Confirmación", MessageBoxButtons.YesNo, MessageBoxIcon.Warning, 
					MessageBoxDefaultButton.Button2)) return;
				this.barraDato1.BorraRegistro();		
			}
		}

		private void barraDato1_Buscar(object sender, System.EventArgs e)
		{
			BuscaCompra miBusqueda = new BuscaCompra();
			miBusqueda.ShowDialog();
		}

		private void barraDato1_Imprimir(object sender, System.EventArgs e)
		{
			if (!miAcceso.Imprimir)
			{
				MessageBox.Show("No tiene acceso a impresión de Facturas");
				return;
			}

			string stReporte = "Compra.rpt";
			string stFiltro = "{Compra.idCompra} = " + this.txtIdCompra.EditValue.ToString();
			switch((int) this.cmbTipoCompra.EditValue)
			{
				case 0: // Compra
					stReporte = "Compra.Rpt";
					break;
			}
			Reporte miReporte = new Reporte(stReporte, stFiltro);
			miReporte.MdiParent = this.MdiParent;
			miReporte.Show();
		}

		private void cdsCompra_PositionChanged(object sender, C1.Data.PositionChangeEventArgs e)
		{
			this.cdvDetCompra.RowFilter = "idCompra = " + txtIdCompra.EditValue.ToString()
				+ " Or idDetCompra <= 0";		
		}

		int IdSigno = 2;
		private Acceso miAcceso;
		private void Compra2_Load(object sender, System.EventArgs e)
		{
			this.txtIdCompra.Width=0;
			this.txtBorrar.Width = 0;
			this.barraDato1.PosUltima();
			C1DataRow dr = cdsNumero.TableViews[0].Rows[IdTipoFactura-1];
			IdSigno = (int) dr["Signo"];
			this.Text = this.cmbTipoCompra.Text;
			miAcceso =new Acceso("0403", this.cdsSeguridad.StorageDataSet.Tables[1]);
		}

		private void cdsCliente_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			C1.Data.SchemaObjects.DataSetDef dataSetDef = e.DataSet.Schema.DataSetDefs["dsCliente"];

			e.Filter.Add(new C1.Data.FilterCondition(dataSetDef.TableViews["Cliente"],
				"[Ambos] = True OR [Proveedor] = False"));
		}

		private void cdsCompra_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			C1.Data.SchemaObjects.DataSetDef dataSetDef = e.DataSet.Schema.DataSetDefs["dsCompra"];

			e.Filter.Add(new C1.Data.FilterCondition(dataSetDef.TableViews["Compra"],
				"[idTipoFactura] = " + this.IdTipoFactura.ToString()));


			int FiltroSigno = 0;
			switch(IdTipoFactura)
			{
				case 1:
				case 6:
				case 8:
				case 10:
					FiltroSigno = -1;
					break;
				case 4:
				case 5:
				case 9:
				case 12:
					FiltroSigno = 1;
					break;
				case 2:
				case 3:
				case 13:
				case 14:
				case 15:
				case 16:
					FiltroSigno = 0;
					break;
			}

			e.Filter.Add(new C1.Data.FilterCondition(dataSetDef.TableViews["DetCompra"],
				"[Signo] = " + FiltroSigno.ToString()));
		}
	}
}
