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
	/// Descripción breve de CuentaAuxiliar.
	/// </summary>
	public class CuentaAuxiliar : DevExpress.XtraEditors.XtraForm
	{
		private DevExpress.XtraEditors.SimpleButton btTablaArbol;
		private DevExpress.XtraEditors.SimpleButton btContraer;
		private DevExpress.XtraEditors.SimpleButton btExpandir;
		private C1.Data.C1DataSet cdsCuenta;
		private C1.Data.C1DataView cdvCuenta;
		private System.Windows.Forms.ImageList imageList1;
		private DevExpress.XtraTreeList.TreeList tlsCuentas;
		private DevExpress.XtraTreeList.Columns.TreeListColumn colCodigo;
		private DevExpress.XtraTreeList.Columns.TreeListColumn colDescripcion;
		private DevExpress.XtraTreeList.Columns.TreeListColumn colPresupuesto;
		private DevExpress.XtraEditors.TextEdit txtIdCuenta;
		private C1.Data.C1DataSet cdsSeguridad;
		private DevExpress.XtraEditors.SimpleButton btMayor;
		private DevExpress.XtraEditors.SimpleButton btBalance;
		private DevExpress.XtraEditors.SimpleButton simpleButton1;
		private System.Windows.Forms.ErrorProvider errorProvider;
		private Infragistics.Win.UltraWinGrid.UltraGrid ultraGrid1;
		private Infragistics.Win.UltraWinEditors.UltraOptionSet optTipoAPC;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtCodigo;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtDescripcion;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtCodPresup;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbCuenta;
		private Infragistics.Win.UltraWinGrid.UltraDropDown cmbFlujo;
		private C1.Data.C1DataSet cdsCuentaGob;
		private Infragistics.Win.UltraWinGrid.UltraDropDown cmbEjecucion;
		private Infragistics.Win.UltraWinGrid.UltraDropDown cmbCuentaBalanceG;
		private Infragistics.Win.Misc.UltraButton btAuditoria;
		private BarraDatoSQL.BarraDatoSQL barraDatoSQL1;
		private System.Windows.Forms.Button btnExcel;
		private Infragistics.Win.UltraWinGrid.ExcelExport.UltraGridExcelExporter ultraGridExcelExporter1;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor ultraCheckEditor1;
		private System.ComponentModel.IContainer components;

		public CuentaAuxiliar()
		{
			InitializeComponent();
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
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(CuentaAuxiliar));
			Infragistics.Win.Appearance appearance1 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Cuenta", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCuenta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idSucursal");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Padre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CodPresup");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("DescIngles");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Borrar");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Grupo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Flujo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Tipo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn11 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaIngreso");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn12 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Editar");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn13 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nuevo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn14 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaEditado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn15 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCuentaFlujo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn16 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCuentaEjecucionP", -1, null, 0, Infragistics.Win.UltraWinGrid.SortIndicator.Ascending, false);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn17 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCuentaBalance");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn18 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CodRapido");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn19 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Descripcion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn20 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("NumPuntos");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn21 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CuentaOrigen");
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.ValueListItem valueListItem1 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem2 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem3 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem4 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem5 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem6 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Cuenta", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn22 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCuenta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn23 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idSucursal");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn24 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn25 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Padre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn26 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CodPresup");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn27 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("DescIngles");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn28 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Borrar");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn29 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Grupo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn30 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Flujo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn31 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Tipo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn32 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaIngreso");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn33 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Editar");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn34 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nuevo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn35 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaEditado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn36 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCuentaFlujo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn37 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCuentaEjecucionP");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn38 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCuentaBalance");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn39 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CodRapido");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn40 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Descripcion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn41 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("NumPuntos");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn42 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CuentaOrigen");
			Infragistics.Win.Appearance appearance11 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance12 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance13 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance14 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand3 = new Infragistics.Win.UltraWinGrid.UltraGridBand("CuentaFlujo", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn43 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCuentaFlujo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn44 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn45 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Flujo");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand4 = new Infragistics.Win.UltraWinGrid.UltraGridBand("CuentaEjecucionP", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn46 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCuentaEjecucionP");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn47 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn48 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Ejecucion");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand5 = new Infragistics.Win.UltraWinGrid.UltraGridBand("CuentaBalance", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn49 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCuentaBalance");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn50 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn51 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Balance");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn52 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Tipo");
			Infragistics.Win.Appearance appearance15 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance16 = new Infragistics.Win.Appearance();
			this.btTablaArbol = new DevExpress.XtraEditors.SimpleButton();
			this.btContraer = new DevExpress.XtraEditors.SimpleButton();
			this.imageList1 = new System.Windows.Forms.ImageList(this.components);
			this.btExpandir = new DevExpress.XtraEditors.SimpleButton();
			this.cdsCuenta = new C1.Data.C1DataSet();
			this.cdvCuenta = new C1.Data.C1DataView();
			this.tlsCuentas = new DevExpress.XtraTreeList.TreeList();
			this.colCodigo = new DevExpress.XtraTreeList.Columns.TreeListColumn();
			this.colDescripcion = new DevExpress.XtraTreeList.Columns.TreeListColumn();
			this.colPresupuesto = new DevExpress.XtraTreeList.Columns.TreeListColumn();
			this.txtIdCuenta = new DevExpress.XtraEditors.TextEdit();
			this.cdsSeguridad = new C1.Data.C1DataSet();
			this.btMayor = new DevExpress.XtraEditors.SimpleButton();
			this.btBalance = new DevExpress.XtraEditors.SimpleButton();
			this.simpleButton1 = new DevExpress.XtraEditors.SimpleButton();
			this.errorProvider = new System.Windows.Forms.ErrorProvider();
			this.ultraGrid1 = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.optTipoAPC = new Infragistics.Win.UltraWinEditors.UltraOptionSet();
			this.txtCodigo = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtDescripcion = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtCodPresup = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.cmbCuenta = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.cmbFlujo = new Infragistics.Win.UltraWinGrid.UltraDropDown();
			this.cdsCuentaGob = new C1.Data.C1DataSet();
			this.cmbEjecucion = new Infragistics.Win.UltraWinGrid.UltraDropDown();
			this.cmbCuentaBalanceG = new Infragistics.Win.UltraWinGrid.UltraDropDown();
			this.btAuditoria = new Infragistics.Win.Misc.UltraButton();
			this.barraDatoSQL1 = new BarraDatoSQL.BarraDatoSQL();
			this.btnExcel = new System.Windows.Forms.Button();
			this.ultraGridExcelExporter1 = new Infragistics.Win.UltraWinGrid.ExcelExport.UltraGridExcelExporter();
			this.ultraCheckEditor1 = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			((System.ComponentModel.ISupportInitialize)(this.cdsCuenta)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvCuenta)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.tlsCuentas)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdCuenta.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeguridad)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.optTipoAPC)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCodigo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDescripcion)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCodPresup)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCuenta)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbFlujo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsCuentaGob)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbEjecucion)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCuentaBalanceG)).BeginInit();
			this.SuspendLayout();
			// 
			// btTablaArbol
			// 
			this.btTablaArbol.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btTablaArbol.Image = ((System.Drawing.Image)(resources.GetObject("btTablaArbol.Image")));
			this.btTablaArbol.Location = new System.Drawing.Point(264, 365);
			this.btTablaArbol.Name = "btTablaArbol";
			this.btTablaArbol.Size = new System.Drawing.Size(24, 25);
			this.btTablaArbol.TabIndex = 6;
			this.btTablaArbol.Text = "Tabla";
			this.btTablaArbol.ToolTip = "Mostrar en forma de tabla";
			this.btTablaArbol.Click += new System.EventHandler(this.btTablaArbol_Click);
			// 
			// btContraer
			// 
			this.btContraer.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btContraer.ImageIndex = 2;
			this.btContraer.ImageList = this.imageList1;
			this.btContraer.Location = new System.Drawing.Point(312, 365);
			this.btContraer.Name = "btContraer";
			this.btContraer.Size = new System.Drawing.Size(24, 25);
			this.btContraer.TabIndex = 8;
			this.btContraer.Text = "-";
			this.btContraer.ToolTip = "Contraer plan de cuentas";
			this.btContraer.Click += new System.EventHandler(this.btContraer_Click);
			// 
			// imageList1
			// 
			this.imageList1.ImageSize = new System.Drawing.Size(16, 16);
			this.imageList1.ImageStream = ((System.Windows.Forms.ImageListStreamer)(resources.GetObject("imageList1.ImageStream")));
			this.imageList1.TransparentColor = System.Drawing.Color.Transparent;
			// 
			// btExpandir
			// 
			this.btExpandir.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btExpandir.ImageIndex = 3;
			this.btExpandir.ImageList = this.imageList1;
			this.btExpandir.Location = new System.Drawing.Point(288, 365);
			this.btExpandir.Name = "btExpandir";
			this.btExpandir.Size = new System.Drawing.Size(24, 25);
			this.btExpandir.TabIndex = 7;
			this.btExpandir.Text = "+";
			this.btExpandir.ToolTip = "Expandir plan de cuentas";
			this.btExpandir.Click += new System.EventHandler(this.btExpandir_Click);
			// 
			// cdsCuenta
			// 
			this.cdsCuenta.BindingContextCtrl = this;
			this.cdsCuenta.DataLibrary = "LibContabilidad";
			this.cdsCuenta.DataLibraryUrl = "";
			this.cdsCuenta.DataSetDef = "dsCuentaContable";
			this.cdsCuenta.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsCuenta.Name = "cdsCuenta";
			this.cdsCuenta.SchemaClassName = "LibContabilidad.DataClass";
			this.cdsCuenta.SchemaDef = null;
			this.cdsCuenta.UpdateError += new C1.Data.DataSetEventHandler(this.cdsCuenta_UpdateError);
			this.cdsCuenta.PositionChanged += new C1.Data.PositionChangeEventHandler(this.cdsCuenta_PositionChanged);
			this.cdsCuenta.BeforeFill += new C1.Data.FillEventHandler(this.cdsCuenta_BeforeFill);
			// 
			// cdvCuenta
			// 
			this.cdvCuenta.BindingContextCtrl = this;
			this.cdvCuenta.DataSet = this.cdsCuenta;
			this.cdvCuenta.Sort = "Codigo";
			this.cdvCuenta.TableName = "";
			this.cdvCuenta.TableViewName = "Cuenta";
			this.cdvCuenta.PositionChanged += new C1.Data.PositionChangeEventHandler(this.cdvCuenta_PositionChanged);
			// 
			// tlsCuentas
			// 
			this.tlsCuentas.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
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
			this.tlsCuentas.Appearance.TreeLine.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Bold);
			this.tlsCuentas.Appearance.TreeLine.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(128)), ((System.Byte)(128)), ((System.Byte)(128)));
			this.tlsCuentas.Appearance.TreeLine.Options.UseBackColor = true;
			this.tlsCuentas.Appearance.TreeLine.Options.UseFont = true;
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
																																															this.colCodigo,
																																															this.colDescripcion,
																																															this.colPresupuesto});
			this.tlsCuentas.CustomizationFormBounds = new System.Drawing.Rectangle(553, 375, 208, 191);
			this.tlsCuentas.DataSource = this.cdvCuenta;
			this.tlsCuentas.ImageIndexFieldName = "";
			this.tlsCuentas.KeyFieldName = "idCuenta";
			this.tlsCuentas.Location = new System.Drawing.Point(8, 8);
			this.tlsCuentas.Name = "tlsCuentas";
			this.tlsCuentas.ParentFieldName = "Padre";
			this.tlsCuentas.RootValue = "0";
			this.tlsCuentas.SelectImageList = this.imageList1;
			this.tlsCuentas.Size = new System.Drawing.Size(1248, 309);
			this.tlsCuentas.Styles.AddReplace("PressedColumn", new DevExpress.Utils.ViewStyle("PressedColumn", "TreeList", new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Bold), "HeaderPanel", ((DevExpress.Utils.StyleOptions)(((DevExpress.Utils.StyleOptions.StyleEnabled | DevExpress.Utils.StyleOptions.UseBackColor) 
				| DevExpress.Utils.StyleOptions.UseForeColor))), true, false, false, DevExpress.Utils.HorzAlignment.Near, DevExpress.Utils.VertAlignment.Center, null, System.Drawing.Color.FromArgb(((System.Byte)(0)), ((System.Byte)(0)), ((System.Byte)(128))), System.Drawing.Color.FromArgb(((System.Byte)(255)), ((System.Byte)(255)), ((System.Byte)(255)))));
			this.tlsCuentas.TabIndex = 12;
			this.tlsCuentas.FocusedNodeChanged += new DevExpress.XtraTreeList.FocusedNodeChangedEventHandler(this.tlsCuentas_FocusedNodeChanged);
			// 
			// colCodigo
			// 
			this.colCodigo.Caption = "Código";
			this.colCodigo.FieldName = "Codigo";
			this.colCodigo.MinWidth = 27;
			this.colCodigo.Name = "colCodigo";
			this.colCodigo.OptionsColumn.ReadOnly = true;
			this.colCodigo.VisibleIndex = 0;
			this.colCodigo.Width = 92;
			// 
			// colDescripcion
			// 
			this.colDescripcion.Caption = "Descripción";
			this.colDescripcion.FieldName = "Descripcion";
			this.colDescripcion.Name = "colDescripcion";
			this.colDescripcion.VisibleIndex = 1;
			this.colDescripcion.Width = 329;
			// 
			// colPresupuesto
			// 
			this.colPresupuesto.Caption = "Presupuesto";
			this.colPresupuesto.FieldName = "CodPresup";
			this.colPresupuesto.MinWidth = 27;
			this.colPresupuesto.Name = "colPresupuesto";
			this.colPresupuesto.VisibleIndex = 2;
			this.colPresupuesto.Width = 1210;
			// 
			// txtIdCuenta
			// 
			this.txtIdCuenta.DataBindings.Add(new System.Windows.Forms.Binding("EditValue", this.cdvCuenta, "idCuenta"));
			this.txtIdCuenta.EditValue = "idCuenta";
			this.txtIdCuenta.Location = new System.Drawing.Point(296, 336);
			this.txtIdCuenta.Name = "txtIdCuenta";
			this.txtIdCuenta.Size = new System.Drawing.Size(75, 20);
			this.txtIdCuenta.TabIndex = 25;
			this.txtIdCuenta.TabStop = false;
			this.txtIdCuenta.EditValueChanged += new System.EventHandler(this.txtIdCuenta_EditValueChanged);
			// 
			// cdsSeguridad
			// 
			this.cdsSeguridad.BindingContextCtrl = this;
			this.cdsSeguridad.DataLibrary = "LibContabilidad";
			this.cdsSeguridad.DataLibraryUrl = "";
			this.cdsSeguridad.DataSetDef = "dsSeguridad";
			this.cdsSeguridad.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsSeguridad.Name = "cdsSeguridad";
			this.cdsSeguridad.SchemaClassName = "LibContabilidad.DataClass";
			this.cdsSeguridad.SchemaDef = null;
			this.cdsSeguridad.BeforeFill += new C1.Data.FillEventHandler(this.cdsSeguridad_BeforeFill);
			// 
			// btMayor
			// 
			this.btMayor.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btMayor.Image = ((System.Drawing.Image)(resources.GetObject("btMayor.Image")));
			this.btMayor.ImageList = this.imageList1;
			this.btMayor.Location = new System.Drawing.Point(360, 365);
			this.btMayor.Name = "btMayor";
			this.btMayor.Size = new System.Drawing.Size(24, 25);
			this.btMayor.TabIndex = 9;
			this.btMayor.Text = "-";
			this.btMayor.ToolTip = "Mayor de cuenta seleccionada";
			this.btMayor.Visible = false;
			this.btMayor.Click += new System.EventHandler(this.btMayor_Click);
			// 
			// btBalance
			// 
			this.btBalance.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btBalance.Image = ((System.Drawing.Image)(resources.GetObject("btBalance.Image")));
			this.btBalance.ImageList = this.imageList1;
			this.btBalance.Location = new System.Drawing.Point(384, 365);
			this.btBalance.Name = "btBalance";
			this.btBalance.Size = new System.Drawing.Size(24, 25);
			this.btBalance.TabIndex = 10;
			this.btBalance.Text = "-";
			this.btBalance.ToolTip = "Balance de comprobación";
			this.btBalance.Visible = false;
			this.btBalance.Click += new System.EventHandler(this.btBalance_Click);
			// 
			// simpleButton1
			// 
			this.simpleButton1.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.simpleButton1.ImageIndex = 4;
			this.simpleButton1.ImageList = this.imageList1;
			this.simpleButton1.Location = new System.Drawing.Point(336, 365);
			this.simpleButton1.Name = "simpleButton1";
			this.simpleButton1.Size = new System.Drawing.Size(24, 25);
			this.simpleButton1.TabIndex = 11;
			this.simpleButton1.Text = "-";
			this.simpleButton1.ToolTip = "Genera Arbol";
			this.simpleButton1.Click += new System.EventHandler(this.simpleButton1_Click);
			// 
			// errorProvider
			// 
			this.errorProvider.ContainerControl = this;
			this.errorProvider.DataMember = "";
			// 
			// ultraGrid1
			// 
			this.ultraGrid1.Cursor = System.Windows.Forms.Cursors.Default;
			this.ultraGrid1.DataSource = this.cdvCuenta;
			appearance1.BackColor = System.Drawing.Color.White;
			this.ultraGrid1.DisplayLayout.Appearance = appearance1;
			this.ultraGrid1.DisplayLayout.AutoFitColumns = true;
			ultraGridColumn1.Header.VisiblePosition = 11;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn2.Header.VisiblePosition = 15;
			ultraGridColumn2.Hidden = true;
			ultraGridColumn2.Width = 60;
			ultraGridColumn3.Header.Caption = "Código";
			ultraGridColumn3.Header.VisiblePosition = 0;
			ultraGridColumn3.Width = 20;
			ultraGridColumn4.Header.VisiblePosition = 12;
			ultraGridColumn4.Hidden = true;
			ultraGridColumn5.Header.Caption = "Presupuesto";
			ultraGridColumn5.Header.VisiblePosition = 2;
			ultraGridColumn5.Width = 20;
			ultraGridColumn6.Header.VisiblePosition = 9;
			ultraGridColumn6.Width = 20;
			ultraGridColumn7.Header.VisiblePosition = 7;
			ultraGridColumn7.Hidden = true;
			ultraGridColumn8.Header.VisiblePosition = 10;
			ultraGridColumn8.Hidden = true;
			ultraGridColumn9.Header.VisiblePosition = 6;
			ultraGridColumn9.Hidden = true;
			ultraGridColumn9.Width = 20;
			ultraGridColumn10.Header.VisiblePosition = 13;
			ultraGridColumn10.Hidden = true;
			ultraGridColumn11.Header.VisiblePosition = 14;
			ultraGridColumn11.Hidden = true;
			ultraGridColumn12.Header.VisiblePosition = 16;
			ultraGridColumn12.Hidden = true;
			ultraGridColumn12.Width = 29;
			ultraGridColumn13.Header.VisiblePosition = 17;
			ultraGridColumn13.Hidden = true;
			ultraGridColumn13.Width = 33;
			ultraGridColumn14.Header.VisiblePosition = 18;
			ultraGridColumn14.Hidden = true;
			ultraGridColumn14.Width = 73;
			ultraGridColumn15.Header.Caption = "Flujo";
			ultraGridColumn15.Header.VisiblePosition = 3;
			ultraGridColumn15.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownValidate;
			ultraGridColumn15.Width = 43;
			ultraGridColumn16.Header.Caption = "Ejec. Pres.";
			ultraGridColumn16.Header.VisiblePosition = 4;
			ultraGridColumn16.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownValidate;
			ultraGridColumn16.Width = 43;
			ultraGridColumn17.Header.Caption = "Balance Gob";
			ultraGridColumn17.Header.VisiblePosition = 5;
			ultraGridColumn17.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownValidate;
			ultraGridColumn17.Width = 43;
			ultraGridColumn18.Header.VisiblePosition = 8;
			ultraGridColumn18.Hidden = true;
			ultraGridColumn19.Header.Caption = "Descripción";
			ultraGridColumn19.Header.VisiblePosition = 1;
			ultraGridColumn19.Width = 26;
			ultraGridColumn20.Header.VisiblePosition = 19;
			ultraGridColumn20.Width = 25;
			ultraGridColumn21.Header.VisiblePosition = 20;
			ultraGridColumn21.Width = 74;
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
																										 ultraGridColumn21});
			this.ultraGrid1.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
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
			this.ultraGrid1.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortSingle;
			appearance4.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance4.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance4.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.RowSelectorAppearance = appearance4;
			appearance5.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance5.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance5.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.SelectedRowAppearance = appearance5;
			this.ultraGrid1.DisplayLayout.ViewStyleBand = Infragistics.Win.UltraWinGrid.ViewStyleBand.OutlookGroupBy;
			this.ultraGrid1.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.ultraGrid1.Location = new System.Drawing.Point(320, 190);
			this.ultraGrid1.Name = "ultraGrid1";
			this.ultraGrid1.Size = new System.Drawing.Size(352, 120);
			this.ultraGrid1.TabIndex = 26;
			this.ultraGrid1.Visible = false;
			this.ultraGrid1.KeyDown += new System.Windows.Forms.KeyEventHandler(this.ultraGrid1_KeyDown);
			this.ultraGrid1.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.ultraGrid1_InitializeLayout);
			// 
			// optTipoAPC
			// 
			this.optTipoAPC.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.optTipoAPC.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdvCuenta, "Tipo"));
			this.optTipoAPC.ItemAppearance = appearance6;
			valueListItem1.DataValue = 1;
			valueListItem1.DisplayText = "Activo";
			valueListItem2.DataValue = 2;
			valueListItem2.DisplayText = "Pasivo";
			valueListItem3.DataValue = 3;
			valueListItem3.DisplayText = "Capital";
			valueListItem4.DataValue = 4;
			valueListItem4.DisplayText = "Ingreso";
			valueListItem5.DataValue = 5;
			valueListItem5.DisplayText = "Egreso";
			valueListItem6.DataValue = 6;
			valueListItem6.DisplayText = "Orden";
			this.optTipoAPC.Items.Add(valueListItem1);
			this.optTipoAPC.Items.Add(valueListItem2);
			this.optTipoAPC.Items.Add(valueListItem3);
			this.optTipoAPC.Items.Add(valueListItem4);
			this.optTipoAPC.Items.Add(valueListItem5);
			this.optTipoAPC.Items.Add(valueListItem6);
			this.optTipoAPC.Location = new System.Drawing.Point(1080, 357);
			this.optTipoAPC.Name = "optTipoAPC";
			this.optTipoAPC.Size = new System.Drawing.Size(176, 35);
			this.optTipoAPC.TabIndex = 27;
			this.optTipoAPC.ValueChanged += new System.EventHandler(this.optTipoAPC_ValueChanged);
			// 
			// txtCodigo
			// 
			this.txtCodigo.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			appearance7.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtCodigo.Appearance = appearance7;
			this.txtCodigo.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdvCuenta, "Codigo"));
			this.txtCodigo.Location = new System.Drawing.Point(8, 325);
			this.txtCodigo.Name = "txtCodigo";
			this.txtCodigo.Size = new System.Drawing.Size(216, 22);
			this.txtCodigo.TabIndex = 28;
			this.txtCodigo.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtCodigo_KeyDown);
			this.txtCodigo.ValueChanged += new System.EventHandler(this.txtCodigo_ValueChanged);
			// 
			// txtDescripcion
			// 
			this.txtDescripcion.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			appearance8.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtDescripcion.Appearance = appearance8;
			this.txtDescripcion.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtDescripcion.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdvCuenta, "Descripcion"));
			this.txtDescripcion.Location = new System.Drawing.Point(232, 325);
			this.txtDescripcion.Name = "txtDescripcion";
			this.txtDescripcion.Size = new System.Drawing.Size(336, 22);
			this.txtDescripcion.TabIndex = 29;
			this.txtDescripcion.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtDescripcion_KeyDown);
			this.txtDescripcion.ValueChanged += new System.EventHandler(this.txtDescripcion_ValueChanged);
			// 
			// txtCodPresup
			// 
			this.txtCodPresup.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			appearance9.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtCodPresup.Appearance = appearance9;
			this.txtCodPresup.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdvCuenta, "CodPresup"));
			this.txtCodPresup.Location = new System.Drawing.Point(576, 325);
			this.txtCodPresup.Name = "txtCodPresup";
			this.txtCodPresup.Size = new System.Drawing.Size(104, 22);
			this.txtCodPresup.TabIndex = 30;
			this.txtCodPresup.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtCodPresup_KeyDown);
			this.txtCodPresup.ValueChanged += new System.EventHandler(this.txtCodPresup_ValueChanged);
			// 
			// cmbCuenta
			// 
			this.cmbCuenta.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.cmbCuenta.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbCuenta.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbCuenta.DataSource = this.cdvCuenta;
			appearance10.BackColor = System.Drawing.Color.White;
			this.cmbCuenta.DisplayLayout.Appearance = appearance10;
			ultraGridColumn22.Header.VisiblePosition = 8;
			ultraGridColumn22.Hidden = true;
			ultraGridColumn23.Header.VisiblePosition = 12;
			ultraGridColumn23.Hidden = true;
			ultraGridColumn24.Header.VisiblePosition = 3;
			ultraGridColumn24.Hidden = true;
			ultraGridColumn25.Header.VisiblePosition = 9;
			ultraGridColumn25.Hidden = true;
			ultraGridColumn26.Header.VisiblePosition = 4;
			ultraGridColumn26.Hidden = true;
			ultraGridColumn27.Header.VisiblePosition = 5;
			ultraGridColumn27.Hidden = true;
			ultraGridColumn28.Header.VisiblePosition = 2;
			ultraGridColumn28.Hidden = true;
			ultraGridColumn29.Header.VisiblePosition = 7;
			ultraGridColumn29.Hidden = true;
			ultraGridColumn30.Header.VisiblePosition = 6;
			ultraGridColumn30.Hidden = true;
			ultraGridColumn31.Header.VisiblePosition = 10;
			ultraGridColumn31.Hidden = true;
			ultraGridColumn32.Header.VisiblePosition = 11;
			ultraGridColumn32.Hidden = true;
			ultraGridColumn33.Header.VisiblePosition = 13;
			ultraGridColumn33.Hidden = true;
			ultraGridColumn34.Header.VisiblePosition = 14;
			ultraGridColumn34.Hidden = true;
			ultraGridColumn35.Header.VisiblePosition = 15;
			ultraGridColumn35.Hidden = true;
			ultraGridColumn36.Header.VisiblePosition = 16;
			ultraGridColumn36.Hidden = true;
			ultraGridColumn37.Header.VisiblePosition = 17;
			ultraGridColumn37.Hidden = true;
			ultraGridColumn38.Header.VisiblePosition = 18;
			ultraGridColumn38.Hidden = true;
			ultraGridColumn39.Header.Caption = "Codigo";
			ultraGridColumn39.Header.VisiblePosition = 0;
			ultraGridColumn40.Header.VisiblePosition = 1;
			ultraGridColumn41.Header.VisiblePosition = 19;
			ultraGridColumn41.Hidden = true;
			ultraGridColumn42.Header.VisiblePosition = 20;
			ultraGridBand2.Columns.AddRange(new object[] {
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
																										 ultraGridColumn34,
																										 ultraGridColumn35,
																										 ultraGridColumn36,
																										 ultraGridColumn37,
																										 ultraGridColumn38,
																										 ultraGridColumn39,
																										 ultraGridColumn40,
																										 ultraGridColumn41,
																										 ultraGridColumn42});
			this.cmbCuenta.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			appearance11.BackColor = System.Drawing.Color.Transparent;
			this.cmbCuenta.DisplayLayout.Override.CardAreaAppearance = appearance11;
			appearance12.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance12.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance12.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance12.FontData.BoldAsString = "True";
			appearance12.FontData.Name = "Arial";
			appearance12.FontData.SizeInPoints = 10F;
			appearance12.ForeColor = System.Drawing.Color.White;
			appearance12.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.cmbCuenta.DisplayLayout.Override.HeaderAppearance = appearance12;
			appearance13.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance13.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance13.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.cmbCuenta.DisplayLayout.Override.RowSelectorAppearance = appearance13;
			appearance14.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance14.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance14.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.cmbCuenta.DisplayLayout.Override.SelectedRowAppearance = appearance14;
			this.cmbCuenta.DisplayMember = "CodRapido";
			this.cmbCuenta.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.OfficeXP;
			this.cmbCuenta.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbCuenta.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbCuenta.Location = new System.Drawing.Point(144, 365);
			this.cmbCuenta.Name = "cmbCuenta";
			this.cmbCuenta.Size = new System.Drawing.Size(112, 21);
			this.cmbCuenta.TabIndex = 32;
			this.cmbCuenta.ValueMember = "idCuenta";
			this.cmbCuenta.Visible = false;
			this.cmbCuenta.MouseDown += new System.Windows.Forms.MouseEventHandler(this.cmbCuenta_MouseDown);
			this.cmbCuenta.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.cmbCuenta_InitializeLayout);
			// 
			// cmbFlujo
			// 
			this.cmbFlujo.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbFlujo.DataMember = "CuentaFlujo";
			this.cmbFlujo.DataSource = this.cdsCuentaGob;
			ultraGridColumn43.Header.VisiblePosition = 2;
			ultraGridColumn43.Hidden = true;
			ultraGridColumn44.Header.VisiblePosition = 1;
			ultraGridColumn44.Width = 74;
			ultraGridColumn45.Header.VisiblePosition = 0;
			ultraGridColumn45.Width = 192;
			ultraGridBand3.Columns.AddRange(new object[] {
																										 ultraGridColumn43,
																										 ultraGridColumn44,
																										 ultraGridColumn45});
			this.cmbFlujo.DisplayLayout.BandsSerializer.Add(ultraGridBand3);
			this.cmbFlujo.DisplayMember = "Flujo";
			this.cmbFlujo.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbFlujo.Location = new System.Drawing.Point(528, 198);
			this.cmbFlujo.Name = "cmbFlujo";
			this.cmbFlujo.Size = new System.Drawing.Size(128, 69);
			this.cmbFlujo.TabIndex = 34;
			this.cmbFlujo.ValueMember = "idCuentaFlujo";
			this.cmbFlujo.Visible = false;
			this.cmbFlujo.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.cmbFlujo_InitializeLayout);
			// 
			// cdsCuentaGob
			// 
			this.cdsCuentaGob.BindingContextCtrl = this;
			this.cdsCuentaGob.DataLibrary = "LibContabilidad";
			this.cdsCuentaGob.DataLibraryUrl = "";
			this.cdsCuentaGob.DataSetDef = "dsCuentaGov";
			this.cdsCuentaGob.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsCuentaGob.Name = "cdsCuentaGob";
			this.cdsCuentaGob.SchemaClassName = "LibContabilidad.DataClass";
			this.cdsCuentaGob.SchemaDef = null;
			// 
			// cmbEjecucion
			// 
			this.cmbEjecucion.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbEjecucion.DataMember = "CuentaEjecucionP";
			this.cmbEjecucion.DataSource = this.cdsCuentaGob;
			ultraGridColumn46.Header.VisiblePosition = 2;
			ultraGridColumn46.Hidden = true;
			ultraGridColumn47.Header.VisiblePosition = 1;
			ultraGridColumn47.Width = 77;
			ultraGridColumn48.Header.VisiblePosition = 0;
			ultraGridColumn48.Width = 240;
			ultraGridBand4.Columns.AddRange(new object[] {
																										 ultraGridColumn46,
																										 ultraGridColumn47,
																										 ultraGridColumn48});
			this.cmbEjecucion.DisplayLayout.BandsSerializer.Add(ultraGridBand4);
			this.cmbEjecucion.DisplayMember = "Ejecucion";
			this.cmbEjecucion.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbEjecucion.Location = new System.Drawing.Point(320, 190);
			this.cmbEjecucion.Name = "cmbEjecucion";
			this.cmbEjecucion.Size = new System.Drawing.Size(136, 60);
			this.cmbEjecucion.TabIndex = 36;
			this.cmbEjecucion.ValueMember = "idCuentaEjecucionP";
			this.cmbEjecucion.Visible = false;
			this.cmbEjecucion.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.cmbEjecucion_InitializeLayout);
			// 
			// cmbCuentaBalanceG
			// 
			this.cmbCuentaBalanceG.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbCuentaBalanceG.DataMember = "CuentaBalance";
			this.cmbCuentaBalanceG.DataSource = this.cdsCuentaGob;
			ultraGridColumn49.Header.VisiblePosition = 2;
			ultraGridColumn49.Hidden = true;
			ultraGridColumn50.Header.VisiblePosition = 1;
			ultraGridColumn50.Width = 66;
			ultraGridColumn51.Header.VisiblePosition = 0;
			ultraGridColumn51.Width = 305;
			ultraGridColumn52.Header.VisiblePosition = 3;
			ultraGridColumn52.Hidden = true;
			ultraGridBand5.Columns.AddRange(new object[] {
																										 ultraGridColumn49,
																										 ultraGridColumn50,
																										 ultraGridColumn51,
																										 ultraGridColumn52});
			this.cmbCuentaBalanceG.DisplayLayout.BandsSerializer.Add(ultraGridBand5);
			this.cmbCuentaBalanceG.DisplayMember = "Balance";
			this.cmbCuentaBalanceG.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbCuentaBalanceG.Location = new System.Drawing.Point(432, 250);
			this.cmbCuentaBalanceG.Name = "cmbCuentaBalanceG";
			this.cmbCuentaBalanceG.Size = new System.Drawing.Size(104, 60);
			this.cmbCuentaBalanceG.TabIndex = 37;
			this.cmbCuentaBalanceG.ValueMember = "idCuentaBalance";
			this.cmbCuentaBalanceG.Visible = false;
			this.cmbCuentaBalanceG.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.cmbCuentaBalanceG_InitializeLayout);
			// 
			// btAuditoria
			// 
			this.btAuditoria.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			appearance15.Image = ((object)(resources.GetObject("appearance15.Image")));
			this.btAuditoria.Appearance = appearance15;
			this.btAuditoria.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.btAuditoria.Location = new System.Drawing.Point(408, 365);
			this.btAuditoria.Name = "btAuditoria";
			this.btAuditoria.Size = new System.Drawing.Size(24, 26);
			this.btAuditoria.TabIndex = 122;
			this.btAuditoria.Click += new System.EventHandler(this.btAuditoria_Click);
			this.btAuditoria.MouseUp += new System.Windows.Forms.MouseEventHandler(this.btAuditoria_MouseUp);
			// 
			// barraDatoSQL1
			// 
			this.barraDatoSQL1.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.barraDatoSQL1.BarraMovimiento = false;
			this.barraDatoSQL1.DataMember = "Cuenta";
			this.barraDatoSQL1.DataNombreId = "idCuenta";
			this.barraDatoSQL1.DataOrden = "";
			this.barraDatoSQL1.DataSource = this.cdsCuenta;
			this.barraDatoSQL1.DataTabla = "Cuenta";
			this.barraDatoSQL1.DataTablaHija = "";
			this.barraDatoSQL1.Location = new System.Drawing.Point(16, 365);
			this.barraDatoSQL1.Name = "barraDatoSQL1";
			this.barraDatoSQL1.Size = new System.Drawing.Size(128, 26);
			this.barraDatoSQL1.TabIndex = 123;
			this.barraDatoSQL1.VisibleBorrar = true;
			this.barraDatoSQL1.VisibleBuscar = true;
			this.barraDatoSQL1.VisibleEditar = true;
			this.barraDatoSQL1.VisibleImprimir = true;
			this.barraDatoSQL1.VisibleNuevo = true;
			this.barraDatoSQL1.Buscar += new System.EventHandler(this.barraDato1_Buscar);
			this.barraDatoSQL1.Load += new System.EventHandler(this.barraDatoSQL1_Load);
			this.barraDatoSQL1.Borrar += new System.EventHandler(this.barraDato1_Borrar);
			this.barraDatoSQL1.Imprimir += new System.EventHandler(this.barraDato1_Imprimir);
			this.barraDatoSQL1.Restaurar += new System.EventHandler(this.barraDato1_Restaurar);
			this.barraDatoSQL1.Refresca += new System.EventHandler(this.barraDatoSQL1_Refresca);
			this.barraDatoSQL1.Crear += new System.EventHandler(this.barraDato1_Crear);
			this.barraDatoSQL1.Deshacer += new System.EventHandler(this.barraDato1_Deshacer);
			this.barraDatoSQL1.Copiar += new System.EventHandler(this.barraDato1_Copiar);
			this.barraDatoSQL1.Editar += new System.EventHandler(this.barraDato1_Editar);
			this.barraDatoSQL1.Grabar += new System.EventHandler(this.barraDato1_Grabar);
			// 
			// btnExcel
			// 
			this.btnExcel.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnExcel.Location = new System.Drawing.Point(440, 365);
			this.btnExcel.Name = "btnExcel";
			this.btnExcel.Size = new System.Drawing.Size(64, 24);
			this.btnExcel.TabIndex = 124;
			this.btnExcel.Text = "&Exportar";
			this.btnExcel.Click += new System.EventHandler(this.btnExcel_Click);
			// 
			// ultraGridExcelExporter1
			// 
			this.ultraGridExcelExporter1.BeginExport += new Infragistics.Win.UltraWinGrid.ExcelExport.BeginExportEventHandler(this.ultraGridExcelExporter1_BeginExport);
			// 
			// ultraCheckEditor1
			// 
			appearance16.ForeColorDisabled = System.Drawing.Color.Black;
			this.ultraCheckEditor1.Appearance = appearance16;
			this.ultraCheckEditor1.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.ultraCheckEditor1.DataBindings.Add(new System.Windows.Forms.Binding("CheckedValue", this.cdvCuenta, "CuentaOrigen"));
			this.ultraCheckEditor1.Location = new System.Drawing.Point(144, 368);
			this.ultraCheckEditor1.Name = "ultraCheckEditor1";
			this.ultraCheckEditor1.Size = new System.Drawing.Size(128, 21);
			this.ultraCheckEditor1.TabIndex = 125;
			this.ultraCheckEditor1.Text = "Cuenta De Diario";
			this.ultraCheckEditor1.CheckedChanged += new System.EventHandler(this.ultraCheckEditor1_CheckedChanged);
			// 
			// CuentaAuxiliar
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 14);
			this.ClientSize = new System.Drawing.Size(720, 398);
			this.Controls.Add(this.ultraCheckEditor1);
			this.Controls.Add(this.btnExcel);
			this.Controls.Add(this.barraDatoSQL1);
			this.Controls.Add(this.btAuditoria);
			this.Controls.Add(this.cmbCuentaBalanceG);
			this.Controls.Add(this.cmbEjecucion);
			this.Controls.Add(this.cmbFlujo);
			this.Controls.Add(this.cmbCuenta);
			this.Controls.Add(this.txtCodPresup);
			this.Controls.Add(this.txtDescripcion);
			this.Controls.Add(this.txtCodigo);
			this.Controls.Add(this.optTipoAPC);
			this.Controls.Add(this.ultraGrid1);
			this.Controls.Add(this.simpleButton1);
			this.Controls.Add(this.btBalance);
			this.Controls.Add(this.btMayor);
			this.Controls.Add(this.txtIdCuenta);
			this.Controls.Add(this.tlsCuentas);
			this.Controls.Add(this.btTablaArbol);
			this.Controls.Add(this.btContraer);
			this.Controls.Add(this.btExpandir);
			this.Name = "CuentaAuxiliar";
			this.Text = "Plan de Cuentas";
			this.Closing += new System.ComponentModel.CancelEventHandler(this.CuentaAuxiliar_Closing);
			this.Load += new System.EventHandler(this.CuentaAuxiliar_Load);
			((System.ComponentModel.ISupportInitialize)(this.cdsCuenta)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvCuenta)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.tlsCuentas)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdCuenta.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeguridad)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.optTipoAPC)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCodigo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDescripcion)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCodPresup)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCuenta)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbFlujo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsCuentaGob)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbEjecucion)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCuentaBalanceG)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void btExpandir_Click(object sender, System.EventArgs e)
		{
			this.tlsCuentas.ExpandAll();
		}

		private void btContraer_Click(object sender, System.EventArgs e)
		{
			this.tlsCuentas.CollapseAll();
		}

		private void btTablaArbol_Click(object sender, System.EventArgs e)
		{
			if (this.tlsCuentas.KeyFieldName == "")
			{
				this.tlsCuentas.KeyFieldName = "idCuenta";
				this.tlsCuentas.ParentFieldName = "Padre";
				this.btTablaArbol.ToolTip = "Mostrar en forma de Tabla";
			}
			else
			{
				this.tlsCuentas.KeyFieldName = "";
				this.tlsCuentas.ParentFieldName = "";
				this.btTablaArbol.ToolTip = "Mostrar en forma de Arbol";
			}		
		}

		private void barraDato1_Borrar(object sender, System.EventArgs e)
		{
			if (!miAcceso.Eliminar)
			{
				MessageBox.Show("No tiene acceso a Eliminar Cuentas", 
					"Información", MessageBoxButtons.OK, MessageBoxIcon.Information);
				return;
			}
			if (Contabilidad.ExisteSubCuenta(this.txtCodigo.Text.ToString(),
				cdsCuenta.StorageDataSet.Tables["Cuenta"]))
			{
				MessageBox.Show("No puede Borrar la Cuenta: " + 
					this.txtCodigo.Text + " porque tiene SubCuentas",
					"Información", MessageBoxButtons.OK, MessageBoxIcon.Information);
				return;
			}

			int iRepite = Contabilidad.LineasAsiento(this.cdsCuenta, (int) this.txtIdCuenta.EditValue);
			if (iRepite > 0)
			{
				MessageBox.Show("No puede Borrar Cuenta: " +
					this.txtCodigo.Text + " porque tiene " + 
					iRepite.ToString() + " movimientos",
					"Información", MessageBoxButtons.OK, MessageBoxIcon.Information);
				return;
			}
			if (DialogResult.No == MessageBox.Show("¿Desea Borrar cuenta: " 
				+ this.txtCodigo.Text + "?", "Confirmación", 
				MessageBoxButtons.YesNo, MessageBoxIcon.Warning, 
				MessageBoxDefaultButton.Button2))
				return;

			tlsCuentas.SuspendLayout();
			string stBorrar = "Delete Cuenta Where idCuenta = " + txtIdCuenta.Text;
			Funcion.EjecutaSQL(cdsCuenta, stBorrar);
			cdsCuenta.Clear();
			cdsCuenta.Fill();
			tlsCuentas.ResumeLayout();
			MessageBox.Show("Cuenta Borrada", "Información",
				MessageBoxButtons.OK, MessageBoxIcon.Information);
		}

		int IdCuenta = 0;
		private void barraDato1_Crear(object sender, System.EventArgs e)
		{
			if (this.tlsCuentas.KeyFieldName == "")
			{
				MessageBox.Show("No tiene Acceso a Crear Nuevas Cuentas cuando esta en grilla", 
					"Información", MessageBoxButtons.OK, MessageBoxIcon.Information);
				return;
			}

			if (!miAcceso.Nuevo)
			{
				MessageBox.Show("No tiene Acceso a Crear Nuevas Cuentas", 
					"Información", MessageBoxButtons.OK, MessageBoxIcon.Information);
				return;
			}
			string stCod = this.txtCodigo.Text.ToString().Trim();
			this.cdvCuenta.Sort = "";
			barraDatoSQL1.CrearRegistro();
			stCod = Contabilidad.CuentaResumen(stCod);
			drCuenta["Codigo"] = stCod; // + "X";
			IdCuenta = (int) drCuenta["idCuenta"];
			DevExpress.XtraTreeList.Nodes.TreeListNode nd = this.tlsCuentas.FindNodeByKeyID(IdCuenta);
			this.tlsCuentas.FocusedNode = nd;
			this.tlsCuentas.SuspendLayout();
			this.cmbCuenta.Enabled = false;
			tlsCuentas.Enabled = false;
			this.txtCodigo.Select();
		}

		private void barraDato1_Deshacer(object sender, System.EventArgs e)
		{
			errorProvider.SetError(this.txtCodigo,"");
			barraDatoSQL1.DeshacerRegistro();
			cdsCuenta.Clear();
			cdsCuenta.Fill();
			cmbCuenta.Enabled = true;
			tlsCuentas.Enabled = true;
		}


		private void barraDato1_Editar(object sender, System.EventArgs e)
		{
			if (!miAcceso.Editar)
			{
				MessageBox.Show("No tiene Acceso a Edición de Plan de Cuentas", "Información", MessageBoxButtons.OK, MessageBoxIcon.Information);
				return;
			}

			barraDatoSQL1.EditarRegistro();
			cmbCuenta.Enabled = false;

			// Busca numero de veces que existe la cuenta en los asientos o es subcuenta
//			if (Contabilidad.LineasAsiento(this.cdsCuenta, (int) this.txtIdCuenta.EditValue) > 0 || 
//							Contabilidad.ExisteSubCuenta(this.txtCodigo.Text.ToString(),
//				cdsCuenta.StorageDataSet.Tables["Cuenta"]))
//
//			{
//				this.txtCodigo.Enabled = false;
//				this.txtDescripcion.Enabled = false;
//			}
//
//			if (miAcceso.EditarCuentaCodigo)
//			{
//				//				HabilitaCodigo();
//				this.txtCodigo.Enabled = true;
//			}
//			if (miAcceso.EditarCuentaNombre)
//			{
//				//				HabilitaNombre();
//				this.txtDescripcion.Enabled = true;
//			}
//			tlsCuentas.Enabled = false;
//			//			this.cdvCuenta.Sort = "";
		}

		private bool Verifica()
		{
			bool bOk = true;
			// Verificacion de codigo de cuenta
			errorProvider.SetError(txtCodigo,"");
			errorProvider.SetError(txtDescripcion,"");
			if (txtCodigo.Value == null || txtCodigo.Value.ToString().Trim().Length == 0)
			{
				errorProvider.SetError(txtCodigo,"Ingrese código de Cuenta");
				bOk = false;
			}
			if (txtDescripcion.Value == null || txtDescripcion.Value.ToString().Trim().Length == 0)
			{
				errorProvider.SetError(txtDescripcion,"Ingrese Descripción de Cuenta");
				bOk = false;
			}
			this.cdsCuenta.Schema.Connections[0].Open();
			C1.Data.SchemaObjects.Connection miConeccion = this.cdsCuenta.Schema.Connections[0];
			string stSelect = "SELECT COUNT(*) FROM Cuenta WHERE Codigo = '" +
				this.txtCodigo.Text.ToString().Trim() + "'";

			if (!barraDatoSQL1.bNuevo)
				stSelect += " And idCuenta <> " +
					this.txtIdCuenta.EditValue.ToString().Trim();

			SqlCommand cmdCuenta = new SqlCommand (
				stSelect, (SqlConnection) miConeccion.DbConnection);
			int iCuenta = (int) cmdCuenta.ExecuteScalar();
			if (iCuenta > 0) 
			{
				this.errorProvider.SetError(this.txtCodigo,"Código de cuenta ya existe");
				bOk = false;
			}
			this.cdsCuenta.Schema.Connections[0].Close();
			// Fin de verificacion de Código repetido
			return bOk;
		}

		private bool CreaDetalle()
		{
			// Busqueda de codigo padre
			drCuenta["codigo"] = drCuenta["codigo"].ToString().Trim();
			string stCodActual = drCuenta["codigo"].ToString();
			string stCodActual1 = stCodActual;
			int iLong = stCodActual.Length - 1;
			if (stCodActual.ToString().EndsWith(".")) 
				stCodActual1 = stCodActual.Remove(iLong, 1);
			string stCodPadre = Contabilidad.CuentaResumen(stCodActual1);

			// Verificacion de codigo de cuenta
			this.cdsCuenta.Schema.Connections[0].Open();
			C1.Data.SchemaObjects.Connection miConeccion = this.cdsCuenta.Schema.Connections[0];
			string stSelect = "SELECT idCuenta FROM Cuenta WHERE Codigo = '" +
				stCodPadre + "'";

			SqlCommand cmdCuenta = new SqlCommand (
				stSelect, (SqlConnection) miConeccion.DbConnection);
			int iCuenta = 0;
			try
			{
				iCuenta = (int) cmdCuenta.ExecuteScalar();
			}
			catch
			{
				if (stCodPadre.Length > 0)
				{
					this.errorProvider.SetError(this.txtCodigo, "No existe Código: " + stCodPadre);
					return false;
				}
			}

			// Asigna Cuentas que son de Grupo las que terminan con "."
			if (drCuenta["Codigo"].ToString().EndsWith(".")) 
				drCuenta["Grupo"] = true;
			else
				drCuenta["Grupo"] = false;

			// Genera CodRapido
			drCuenta["CodRapido"] = drCuenta["Codigo"].ToString().Replace(".","");
			int iNumPuntos = 0;
			for (int i= 0; i < drCuenta["Codigo"].ToString().Length ; i++)
			{
				if (drCuenta["Codigo"].ToString().Substring(i, 1) == ".")
					iNumPuntos ++;
			}
			drCuenta["NumPuntos"] = iNumPuntos;
			drCuenta["Tipo"] = Funcion.iEscalarSQL(cdsCuenta, 
				string.Format("Exec CuentaAsignaTipo '{0}'", drCuenta["Codigo"].ToString().Substring(0, 1)));

			if (iCuenta > 0) drCuenta["Padre"] = iCuenta;
			else drCuenta["Padre"] = 0;
			return true;
		}

		private void barraDato1_Grabar(object sender, System.EventArgs e)
		{
			if (!Verifica()) return;
			txtCodPresup.Focus();
			txtCodigo.Focus();
			if (!CreaDetalle()) return;
			tlsCuentas.ResumeLayout();
			barraDatoSQL1.GrabaRegistro();
			string stSelect = "Select idCuenta From Cuenta Where Codigo = '" + txtCodigo.Text.Trim() + "'";
			int iCuenta = Funcion.iEscalarSQL(cdsCuenta, stSelect);
			cdsCuenta.Clear();
			cdsCuenta.Fill();
			cdvCuenta.Sort = "Codigo";
			DevExpress.XtraTreeList.Nodes.TreeListNode nd = tlsCuentas.FindNodeByKeyID(iCuenta);
			tlsCuentas.FocusedNode = nd;
			cmbCuenta.Enabled = true;
			tlsCuentas.Enabled = true;
		}

		private void barraDato1_Imprimir(object sender, System.EventArgs e)
		{
			if (!miAcceso.Imprimir)
			{
				MessageBox.Show("No tiene acceso a impresión de plan de cuentas", "Información", MessageBoxButtons.OK, MessageBoxIcon.Information);
				return;
			}

			// Generacion de Titulo
			string strTitulo="";
			this.Cursor = Cursors.WaitCursor;
			strTitulo = "Plan de Cuentas";

			Reporte miReporte;
			miReporte = new Reporte("Cuenta.rpt");
//			miReporte.MdiParent = this.MdiParent;
			miReporte.Titulo(strTitulo);
			miReporte.Show();
			this.Cursor = Cursors.Arrow;
		}

		private Acceso miAcceso;
		private void UnloadMe()
		{
			this.Close();
		}
		private void CuentaAuxiliar_Load(object sender, System.EventArgs e)
		{
			if (!Funcion.Permiso("43",cdsSeguridad))
			{
				MessageBox.Show("No puede ingresar al Plan de Cuentas", "Informacion");
				BeginInvoke(new MethodInvoker(UnloadMe));
				return;
			}

			string stAudita = "Exec AuditaCrear 53, 6, 'Plan Cuentas'";
			Funcion.EjecutaSQL(cdsCuenta, stAudita, true);
			Cursor = Cursors.Default;
			miAcceso =new Acceso("0401", cdsSeguridad.StorageDataSet.Tables["Seguridad"]);
			barraDatoSQL1.HabilitaControles(false);
			this.txtIdCuenta.Width = 0;
			this.HabilitaArbol(false);
			if (miAcceso.EditarTabla)
			{
				this.cdvCuenta.Sort = "";
				this.ultraGrid1.Visible = true;
				this.ultraGrid1.Top = this.tlsCuentas.Top;
				this.ultraGrid1.Left = this.tlsCuentas.Left;
				this.ultraGrid1.Width = this.tlsCuentas.Width;
				this.ultraGrid1.Height = this.tlsCuentas.Height + 30;
				this.ultraGrid1.Anchor = this.tlsCuentas.Anchor;
				this.txtCodigo.Width = 0;
				this.txtCodPresup.Width = 0;
				this.txtDescripcion.Width = 0;
				this.tlsCuentas.Visible = false;
			}
			else optTipoAPC.Width = 0;

			this.cmbCuenta.Enabled = true;
		}

		private void HabilitaArbol(bool Habilita)
		{
			this.colCodigo.OptionsColumn.ReadOnly = !Habilita;
			this.colDescripcion.OptionsColumn.ReadOnly = !Habilita;
			this.colPresupuesto.OptionsColumn.ReadOnly = !Habilita;
		}
		
		private void HabilitaCodigo()
		{
			this.colCodigo.OptionsColumn.ReadOnly = false;
			this.colPresupuesto.OptionsColumn.ReadOnly =false;
		}

		private void HabilitaNombre()
		{
			this.colDescripcion.OptionsColumn.ReadOnly = false;
		}

		private void barraDato1_Buscar(object sender, System.EventArgs e)
		{
			/*if (this.cmbCuenta.Value == null)
			{
				this.cmbCuenta.Focus();
				MessageBox.Show("Escoja cuenta a buscar");
				return;
			}*/
			using (BuscaCuentas miBusqueda = new BuscaCuentas("", 1))
			{
				if (DialogResult.OK == miBusqueda.ShowDialog())
				{	
					int IdCuenta = (int) miBusqueda.ultraGrid1.ActiveRow.Cells["idCuenta"].Value;// this.cmbCuenta.Value;
					DevExpress.XtraTreeList.Nodes.TreeListNode nd = this.tlsCuentas.FindNodeByKeyID(IdCuenta);
					//			MessageBox.Show(nd.GetDisplayText(0).ToString() + " Pos: " + IdCuenta.ToString());
					//			this.tlsCuentas.FullExpandNode(nd);
					this.tlsCuentas.FocusedNode = nd;
					//			this.tlsCuentas.Refresh();
				}
			}
		}

		private void btMayor_Click(object sender, System.EventArgs e)
		{
			this.Cursor = Cursors.WaitCursor;
			// Definicion de Parametros
			ParameterFields paramFields = new ParameterFields ();
			ParameterField paramField = new ParameterField ();

			ParameterDiscreteValue discreteVal = new ParameterDiscreteValue ();
			paramField.ParameterFieldName = "FechaIni";
			discreteVal.Value = DateTime.Today.ToString("yyyy-MM-dd");
			paramField.CurrentValues.Add (discreteVal);
			paramFields.Add (paramField);
			// Fin de parametros

			int IdCuenta = (int) this.txtIdCuenta.EditValue;
			string stFiltro = "{Cuenta.idCuenta} = " + IdCuenta.ToString();
			Reporte miReporte;
			miReporte = new Reporte("Mayor.rpt", stFiltro);
			miReporte.MdiParent = this.MdiParent;
			miReporte.crVista.ParameterFieldInfo = paramFields;
			miReporte.Show();		
			this.Cursor = Cursors.Arrow;
		}

		private void btBalance_Click(object sender, System.EventArgs e)
		{
			this.Cursor = Cursors.WaitCursor;
			// Definicion de Parametros
			ParameterFields paramFields = new ParameterFields ();
			ParameterField paramField = new ParameterField ();

			ParameterDiscreteValue discreteVal = new ParameterDiscreteValue ();
			paramField.ParameterFieldName = "FechaIni";
			discreteVal.Value = DateTime.Today.AddDays(-10000).ToString("yyyy-MM-dd");
			paramField.CurrentValues.Add (discreteVal);
			paramFields.Add (paramField);

			ParameterField paramField1 = new ParameterField ();
			ParameterDiscreteValue discreteVal1 = new ParameterDiscreteValue ();
			paramField1.ParameterFieldName = "pcTipo";
			discreteVal1.Value = "Hasta el " + DateTime.Today.ToString("dd/MMM/yyyy");
			paramField1.CurrentValues.Add (discreteVal1);
			paramFields.Add (paramField1);
			// Fin de parametros

			Reporte miReporte;
			miReporte = new Reporte("Balance.rpt");
//			miReporte.MdiParent = this.MdiParent;
			miReporte.crVista.ParameterFieldInfo = paramFields;
			miReporte.Titulo("Balance de Comprobación");
			miReporte.Show();
			this.Cursor = Cursors.Arrow;
		}

		private void simpleButton1_Click(object sender, System.EventArgs e)
		{
			Cursor = Cursors.WaitCursor;
			string stSelect = "Update Cuenta Set Tipo = 0";
			Funcion.EjecutaSQL(cdsCuenta, stSelect, true);
			Funcion.EjecutaSQL(cdsCuenta, "Exec CuentaArbol", true);

//			Contabilidad.GeneraArbol(this.cdsCuenta);
			cdsCuenta.Clear();
			cdsCuenta.Fill();
			Cursor = Cursors.Default;
			MessageBox.Show("Arbol de Cuentas Generado", "Informaciòn",
				MessageBoxButtons.OK, MessageBoxIcon.Information);
		}

		private void cdsCuenta_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			this.cdsCuenta.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirConta);
		}

		private void cdsSeguridad_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			this.cdsSeguridad.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirConta);		
		}

		C1.Data.C1DataRow drCuenta;
		private void cdsCuenta_PositionChanged(object sender, C1.Data.PositionChangeEventArgs e)
		{
			drCuenta = e.Row;		
		}

		private void cdsCuenta_UpdateError(object sender, C1.Data.DataSetEventArgs e)
		{
			MessageBox.Show(e.Error.Message.ToString(), "Error al actualizar");
		}

		private void cdvCuenta_PositionChanged(object sender, C1.Data.PositionChangeEventArgs e)
		{
			drCuenta = e.Row;		
		}

		private void barraDato1_Copiar(object sender, System.EventArgs e)
		{
			MessageBox.Show("No se puede Copiar Cuentas", "Información", 
				MessageBoxButtons.OK, MessageBoxIcon.Information);
		}

		private void barraDato1_Restaurar(object sender, System.EventArgs e)
		{
			MessageBox.Show("No se Recuperar Cuentas Borradas", "Información", 
				MessageBoxButtons.OK, MessageBoxIcon.Information);		
		}


		private void CuentaAuxiliar_Closing(object sender, System.ComponentModel.CancelEventArgs e)
		{
			string stAudita = "Exec AuditaCrear 53, 7, 'Plan Cuentas'";
			Funcion.EjecutaSQL(cdsCuenta, stAudita, true);
		}

		private void btFlujo_MouseUp(object sender, System.Windows.Forms.MouseEventArgs e)
		{
			if (e.Button != MouseButtons.Right) return;
			ultraGrid1.ActiveRow.Cells["idCuentaFlujo"].Value = DBNull.Value;
			ultraGrid1.ActiveRow.Cells["idCuentaEjecucionP"].Value = DBNull.Value;
		}

		private void cmbCuenta_MouseDown(object sender, System.Windows.Forms.MouseEventArgs e)
		{
			if (e.Button == MouseButtons.Right)
			{
				string stPrimero = "Descripcion";
				if (cmbCuenta.DisplayLayout.Bands[0].Columns[stPrimero].Header.VisiblePosition == 0)
					stPrimero = "CodRapido";
				cdvCuenta.Sort = stPrimero;
				cmbCuenta.DisplayLayout.Bands[0].Columns[stPrimero].Header.VisiblePosition = 0;
				cmbCuenta.DisplayMember = stPrimero;
			}
		}

		private void ultraGrid1_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (ultraGrid1.ActiveCell == null) return;
			try
			{
				if (ultraGrid1.ActiveCell.Column.Key.ToString().ToUpper() == "NOTAS") return;
				switch(e.KeyValue)
				{
					case (int) Keys.PageDown:
						if(ultraGrid1.ActiveCell.DroppedDown == false)
							ultraGrid1.ActiveCell.DroppedDown = true;
						break;

					case (int) Keys.Up:
						if (ultraGrid1.ActiveCell.DroppedDown) return;

						ultraGrid1.PerformAction(UltraGridAction.ExitEditMode, false, false);
						ultraGrid1.PerformAction(UltraGridAction.AboveCell, false, false);
						e.Handled = true;
						ultraGrid1.PerformAction(UltraGridAction.EnterEditMode, false, false);
						break;
					case (int) Keys.Left:
						ultraGrid1.PerformAction(UltraGridAction.ExitEditMode, false, false);
						ultraGrid1.PerformAction(UltraGridAction.PrevCellByTab, false, false);
						e.Handled = true;
						ultraGrid1.PerformAction(UltraGridAction.EnterEditMode, false, false);
						break;
					case (int) Keys.Down:
						if (ultraGrid1.ActiveCell.DroppedDown) return;
			
						ultraGrid1.PerformAction(UltraGridAction.ExitEditMode, false, false);
						ultraGrid1.PerformAction(UltraGridAction.BelowCell, false, false);
						e.Handled = true;
						ultraGrid1.PerformAction(UltraGridAction.EnterEditMode, false, false);
						break;
					case (int) Keys.Enter:
						ultraGrid1.PerformAction(UltraGridAction.ExitEditMode, false, false);
						ultraGrid1.PerformAction(UltraGridAction.NextCellByTab, false, false);
						e.Handled = true;
						ultraGrid1.PerformAction(UltraGridAction.EnterEditMode, false, false);
						if(ultraGrid1.ActiveCell.DroppedDown == false)
							ultraGrid1.ActiveCell.DroppedDown = true;
						//						ultraGrid1.PerformAction(UltraGridAction.EnterEditModeAndDropdown, false, false);
						break;
				}
			}
			catch(System.Exception ex)
			{
				Console.WriteLine(ex.Message);
				//				MessageBox.Show(ex.Message, "Tecla");
			}

		}

		private void btAuditoria_Click(object sender, System.EventArgs e)
		{
			Cursor = Cursors.WaitCursor;
			Funcion.AuditoriaImprime(27, (int) txtIdCuenta.EditValue, true);
			Cursor = Cursors.Default;
		}

		private void btAuditoria_MouseUp(object sender, System.Windows.Forms.MouseEventArgs e)
		{
			if (e.Button != MouseButtons.Right) return;
			Cursor = Cursors.WaitCursor;
			Funcion.AuditoriaImprime(27, (int) txtIdCuenta.EditValue, false);
			Cursor = Cursors.Default;
		}

		private void barraDatoSQL1_Refresca(object sender, System.EventArgs e)
		{
		
		}

		private void barraDatoSQL1_Load(object sender, System.EventArgs e)
		{
		
		}

		private void tlsCuentas_FocusedNodeChanged(object sender, DevExpress.XtraTreeList.FocusedNodeChangedEventArgs e)
		{
		
		}

		private void btnExcel_Click(object sender, System.EventArgs e)
		{
			if (this.ultraGrid1.Rows.Count == 0)
			{
				MessageBox.Show("No Hay datos para exportar.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
				return;
			}

			string stDir = MenuLatinium.stDirInicio + "\\"; 
			ultraGridExcelExporter1.Export(this.ultraGrid1, "PlanDeCuentas.xls");
			System.Diagnostics.Process.Start(stDir + "PlanDeCuentas.xls");
		}

		private void txtCodigo_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.txtDescripcion.Focus();
		}

		private void txtDescripcion_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.txtCodPresup.Focus();
		}

		private void txtCodPresup_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.barraDatoSQL1.Focus();
		}

		private void ultraGrid1_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}

		private void txtIdCuenta_EditValueChanged(object sender, System.EventArgs e)
		{
		
		}

		private void txtCodigo_ValueChanged(object sender, System.EventArgs e)
		{
		
		}

		private void txtDescripcion_ValueChanged(object sender, System.EventArgs e)
		{
		
		}

		private void txtCodPresup_ValueChanged(object sender, System.EventArgs e)
		{
		
		}

		private void cmbCuenta_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}

		private void cmbFlujo_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}

		private void cmbCuentaBalanceG_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}

		private void ultraGridExcelExporter1_BeginExport(object sender, Infragistics.Win.UltraWinGrid.ExcelExport.BeginExportEventArgs e)
		{
		
		}

		private void ultraCheckEditor1_CheckedChanged(object sender, System.EventArgs e)
		{
		
		}

		private void cmbEjecucion_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}

		private void optTipoAPC_ValueChanged(object sender, System.EventArgs e)
		{
		
		}
	}
}
