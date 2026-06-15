using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de CuentaCorriente.
	/// </summary>
	public class CuentaCorriente : System.Windows.Forms.Form
	{
		private C1.Data.C1DataSet cdsCuentaCorr;
		private C1.Data.C1DataView cdvCuentaCorr;
		private DevExpress.XtraGrid.GridControl gridControl1;
		private DevExpress.XtraGrid.Views.Grid.GridView gridView1;
		private DevExpress.XtraGrid.Columns.GridColumn colAutomatico;
		private DevExpress.XtraGrid.Columns.GridColumn colidCuenta;
		private DevExpress.XtraGrid.Columns.GridColumn colNumCheque;
		private DevExpress.XtraGrid.Columns.GridColumn colNumCuenta;
		private DevExpress.XtraGrid.Columns.GridColumn colReporte;
		private DevExpress.XtraGrid.Columns.GridColumn colidCuenta1;
		private DevExpress.XtraEditors.Repository.RepositoryItemLookUpEdit repositoryItemLookUpEdit1;
		private C1.Data.C1DataSet cdsCuentas;
		private C1.Data.C1DataView cdvCuenta;
		private C1.Data.C1DataView cdvCodCuenta;
		private DevExpress.XtraEditors.Repository.RepositoryItemLookUpEdit repositoryItemLookUpEdit2;
		private DevExpress.XtraEditors.SimpleButton btAceptar;
		private DevExpress.XtraEditors.SimpleButton btCancelar;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public CuentaCorriente()
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
			this.cdsCuentaCorr = new C1.Data.C1DataSet();
			this.cdvCuentaCorr = new C1.Data.C1DataView();
			this.gridControl1 = new DevExpress.XtraGrid.GridControl();
			this.gridView1 = new DevExpress.XtraGrid.Views.Grid.GridView();
			this.colNumCuenta = new DevExpress.XtraGrid.Columns.GridColumn();
			this.colNumCheque = new DevExpress.XtraGrid.Columns.GridColumn();
			this.colidCuenta = new DevExpress.XtraGrid.Columns.GridColumn();
			this.repositoryItemLookUpEdit2 = new DevExpress.XtraEditors.Repository.RepositoryItemLookUpEdit();
			this.cdvCodCuenta = new C1.Data.C1DataView();
			this.cdsCuentas = new C1.Data.C1DataSet();
			this.colidCuenta1 = new DevExpress.XtraGrid.Columns.GridColumn();
			this.repositoryItemLookUpEdit1 = new DevExpress.XtraEditors.Repository.RepositoryItemLookUpEdit();
			this.cdvCuenta = new C1.Data.C1DataView();
			this.colAutomatico = new DevExpress.XtraGrid.Columns.GridColumn();
			this.colReporte = new DevExpress.XtraGrid.Columns.GridColumn();
			this.btAceptar = new DevExpress.XtraEditors.SimpleButton();
			this.btCancelar = new DevExpress.XtraEditors.SimpleButton();
			((System.ComponentModel.ISupportInitialize)(this.cdsCuentaCorr)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvCuentaCorr)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.gridControl1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.gridView1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.repositoryItemLookUpEdit2)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvCodCuenta)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsCuentas)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.repositoryItemLookUpEdit1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvCuenta)).BeginInit();
			this.SuspendLayout();
			// 
			// cdsCuentaCorr
			// 
			this.cdsCuentaCorr.BindingContextCtrl = this;
			this.cdsCuentaCorr.DataLibrary = "LibContabilidad";
			this.cdsCuentaCorr.DataLibraryUrl = "";
			this.cdsCuentaCorr.DataSetDef = "dsCuentaCorriente";
			this.cdsCuentaCorr.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsCuentaCorr.Name = "cdsCuentaCorr";
			this.cdsCuentaCorr.SchemaClassName = "LibContabilidad.DataClass";
			this.cdsCuentaCorr.SchemaDef = null;
			this.cdsCuentaCorr.BeforeFill += new C1.Data.FillEventHandler(this.cdsCuentaCorr_BeforeFill);
			// 
			// cdvCuentaCorr
			// 
			this.cdvCuentaCorr.BindingContextCtrl = this;
			this.cdvCuentaCorr.DataSet = this.cdsCuentaCorr;
			this.cdvCuentaCorr.TableName = "";
			this.cdvCuentaCorr.TableViewName = "CuentaCorriente";
			// 
			// gridControl1
			// 
			this.gridControl1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.gridControl1.DataSource = this.cdvCuentaCorr;
			// 
			// gridControl1.EmbeddedNavigator
			// 
			this.gridControl1.EmbeddedNavigator.Name = "";
			this.gridControl1.Location = new System.Drawing.Point(16, 16);
			this.gridControl1.MainView = this.gridView1;
			this.gridControl1.Name = "gridControl1";
			this.gridControl1.RepositoryItems.AddRange(new DevExpress.XtraEditors.Repository.RepositoryItem[] {
																																																					this.repositoryItemLookUpEdit1,
																																																					this.repositoryItemLookUpEdit2});
			this.gridControl1.Size = new System.Drawing.Size(440, 224);
			this.gridControl1.Styles.AddReplace("Preview", new DevExpress.Utils.ViewStyleEx("Preview", "Grid", new System.Drawing.Font("Tahoma", 10F), "", true, true, false, DevExpress.Utils.HorzAlignment.Near, DevExpress.Utils.VertAlignment.Top, null, System.Drawing.Color.White, System.Drawing.Color.FromArgb(((System.Byte)(10)), ((System.Byte)(36)), ((System.Byte)(106))), System.Drawing.Color.White, System.Drawing.Drawing2D.LinearGradientMode.Horizontal));
			this.gridControl1.Styles.AddReplace("FooterPanel", new DevExpress.Utils.ViewStyleEx("FooterPanel", "Grid", new System.Drawing.Font("Tahoma", 10F), "", true, false, false, DevExpress.Utils.HorzAlignment.Far, DevExpress.Utils.VertAlignment.Center, null, System.Drawing.Color.FromArgb(((System.Byte)(212)), ((System.Byte)(208)), ((System.Byte)(200))), System.Drawing.Color.Black, System.Drawing.Color.Empty, System.Drawing.Drawing2D.LinearGradientMode.Horizontal));
			this.gridControl1.Styles.AddReplace("GroupButton", new DevExpress.Utils.ViewStyleEx("GroupButton", "Grid", new System.Drawing.Font("Tahoma", 10F), System.Drawing.Color.FromArgb(((System.Byte)(212)), ((System.Byte)(208)), ((System.Byte)(200))), System.Drawing.Color.Black, System.Drawing.Color.Empty, System.Drawing.Drawing2D.LinearGradientMode.Horizontal));
			this.gridControl1.Styles.AddReplace("FilterCloseButton", new DevExpress.Utils.ViewStyleEx("FilterCloseButton", "Grid", new System.Drawing.Font("Arial", 8F), System.Drawing.Color.FromArgb(((System.Byte)(212)), ((System.Byte)(208)), ((System.Byte)(200))), System.Drawing.Color.Black, System.Drawing.Color.FromArgb(((System.Byte)(118)), ((System.Byte)(170)), ((System.Byte)(225))), System.Drawing.Drawing2D.LinearGradientMode.ForwardDiagonal));
			this.gridControl1.Styles.AddReplace("EvenRow", new DevExpress.Utils.ViewStyleEx("EvenRow", "Grid", new System.Drawing.Font("Tahoma", 10F), ((DevExpress.Utils.StyleOptions)((DevExpress.Utils.StyleOptions.StyleEnabled | DevExpress.Utils.StyleOptions.UseBackColor))), System.Drawing.Color.FromArgb(((System.Byte)(212)), ((System.Byte)(208)), ((System.Byte)(200))), System.Drawing.Color.Black, System.Drawing.Color.GhostWhite, System.Drawing.Drawing2D.LinearGradientMode.ForwardDiagonal));
			this.gridControl1.Styles.AddReplace("HideSelectionRow", new DevExpress.Utils.ViewStyleEx("HideSelectionRow", "Grid", new System.Drawing.Font("Tahoma", 10F), ((DevExpress.Utils.StyleOptions)((((((DevExpress.Utils.StyleOptions.StyleEnabled | DevExpress.Utils.StyleOptions.UseBackColor) 
				| DevExpress.Utils.StyleOptions.UseDrawFocusRect) 
				| DevExpress.Utils.StyleOptions.UseFont) 
				| DevExpress.Utils.StyleOptions.UseForeColor) 
				| DevExpress.Utils.StyleOptions.UseImage))), System.Drawing.Color.Gray, System.Drawing.Color.FromArgb(((System.Byte)(212)), ((System.Byte)(208)), ((System.Byte)(200))), System.Drawing.Color.Empty, System.Drawing.Drawing2D.LinearGradientMode.Horizontal));
			this.gridControl1.Styles.AddReplace("FixedLine", new DevExpress.Utils.ViewStyleEx("FixedLine", "Grid", new System.Drawing.Font("Arial", 8F), System.Drawing.Color.FromArgb(((System.Byte)(58)), ((System.Byte)(58)), ((System.Byte)(58))), System.Drawing.Color.Black, System.Drawing.Color.Empty, System.Drawing.Drawing2D.LinearGradientMode.Horizontal));
			this.gridControl1.Styles.AddReplace("HeaderPanel", new DevExpress.Utils.ViewStyleEx("HeaderPanel", "Grid", new System.Drawing.Font("Tahoma", 10F, System.Drawing.FontStyle.Bold), "", true, false, false, DevExpress.Utils.HorzAlignment.Near, DevExpress.Utils.VertAlignment.Center, null, System.Drawing.Color.FromArgb(((System.Byte)(212)), ((System.Byte)(208)), ((System.Byte)(200))), System.Drawing.Color.Black, System.Drawing.Color.Empty, System.Drawing.Drawing2D.LinearGradientMode.Horizontal));
			this.gridControl1.Styles.AddReplace("GroupPanel", new DevExpress.Utils.ViewStyleEx("GroupPanel", "Grid", new System.Drawing.Font("Tahoma", 10F, System.Drawing.FontStyle.Bold), "", true, false, false, DevExpress.Utils.HorzAlignment.Near, DevExpress.Utils.VertAlignment.Center, null, System.Drawing.Color.FromArgb(((System.Byte)(58)), ((System.Byte)(110)), ((System.Byte)(165))), System.Drawing.Color.White, System.Drawing.Color.White, System.Drawing.Drawing2D.LinearGradientMode.Horizontal));
			this.gridControl1.Styles.AddReplace("Empty", new DevExpress.Utils.ViewStyleEx("Empty", "Grid", new System.Drawing.Font("Arial", 8F), System.Drawing.Color.FromArgb(((System.Byte)(232)), ((System.Byte)(228)), ((System.Byte)(220))), System.Drawing.Color.White, System.Drawing.Color.Empty, System.Drawing.Drawing2D.LinearGradientMode.Horizontal));
			this.gridControl1.Styles.AddReplace("GroupFooter", new DevExpress.Utils.ViewStyleEx("GroupFooter", "Grid", new System.Drawing.Font("Tahoma", 8F), "", true, false, false, DevExpress.Utils.HorzAlignment.Far, DevExpress.Utils.VertAlignment.Center, null, System.Drawing.Color.FromArgb(((System.Byte)(222)), ((System.Byte)(218)), ((System.Byte)(210))), System.Drawing.Color.Black, System.Drawing.Color.Empty, System.Drawing.Drawing2D.LinearGradientMode.Horizontal));
			this.gridControl1.Styles.AddReplace("GroupRow", new DevExpress.Utils.ViewStyleEx("GroupRow", "Grid", new System.Drawing.Font("Tahoma", 10F), "", true, false, false, DevExpress.Utils.HorzAlignment.Near, DevExpress.Utils.VertAlignment.Center, null, System.Drawing.Color.Gray, System.Drawing.Color.FromArgb(((System.Byte)(212)), ((System.Byte)(208)), ((System.Byte)(200))), System.Drawing.Color.Empty, System.Drawing.Drawing2D.LinearGradientMode.Horizontal));
			this.gridControl1.Styles.AddReplace("HorzLine", new DevExpress.Utils.ViewStyleEx("HorzLine", "Grid", new System.Drawing.Font("Tahoma", 10F), System.Drawing.Color.FromArgb(((System.Byte)(212)), ((System.Byte)(208)), ((System.Byte)(200))), System.Drawing.Color.Gray, System.Drawing.Color.Empty, System.Drawing.Drawing2D.LinearGradientMode.Horizontal));
			this.gridControl1.Styles.AddReplace("ColumnFilterButton", new DevExpress.Utils.ViewStyleEx("ColumnFilterButton", "Grid", new System.Drawing.Font("Arial", 8F), System.Drawing.Color.FromArgb(((System.Byte)(212)), ((System.Byte)(208)), ((System.Byte)(200))), System.Drawing.Color.Gray, System.Drawing.Color.FromArgb(((System.Byte)(232)), ((System.Byte)(228)), ((System.Byte)(220))), System.Drawing.Drawing2D.LinearGradientMode.Horizontal));
			this.gridControl1.Styles.AddReplace("FocusedRow", new DevExpress.Utils.ViewStyleEx("FocusedRow", "Grid", new System.Drawing.Font("Tahoma", 10F), ((DevExpress.Utils.StyleOptions)((((((DevExpress.Utils.StyleOptions.StyleEnabled | DevExpress.Utils.StyleOptions.UseBackColor) 
				| DevExpress.Utils.StyleOptions.UseDrawFocusRect) 
				| DevExpress.Utils.StyleOptions.UseFont) 
				| DevExpress.Utils.StyleOptions.UseForeColor) 
				| DevExpress.Utils.StyleOptions.UseImage))), System.Drawing.Color.FromArgb(((System.Byte)(10)), ((System.Byte)(36)), ((System.Byte)(106))), System.Drawing.Color.White, System.Drawing.Color.FromArgb(((System.Byte)(60)), ((System.Byte)(86)), ((System.Byte)(156))), System.Drawing.Drawing2D.LinearGradientMode.Horizontal));
			this.gridControl1.Styles.AddReplace("VertLine", new DevExpress.Utils.ViewStyleEx("VertLine", "Grid", new System.Drawing.Font("Tahoma", 10F), "", true, false, false, DevExpress.Utils.HorzAlignment.Near, DevExpress.Utils.VertAlignment.Center, null, System.Drawing.Color.FromArgb(((System.Byte)(212)), ((System.Byte)(208)), ((System.Byte)(200))), System.Drawing.Color.Gray, System.Drawing.Color.Empty, System.Drawing.Drawing2D.LinearGradientMode.Horizontal));
			this.gridControl1.Styles.AddReplace("FocusedCell", new DevExpress.Utils.ViewStyleEx("FocusedCell", "Grid", new System.Drawing.Font("Tahoma", 10F), ((DevExpress.Utils.StyleOptions)((((((DevExpress.Utils.StyleOptions.StyleEnabled | DevExpress.Utils.StyleOptions.UseBackColor) 
				| DevExpress.Utils.StyleOptions.UseDrawFocusRect) 
				| DevExpress.Utils.StyleOptions.UseFont) 
				| DevExpress.Utils.StyleOptions.UseForeColor) 
				| DevExpress.Utils.StyleOptions.UseImage))), System.Drawing.Color.FromArgb(((System.Byte)(255)), ((System.Byte)(255)), ((System.Byte)(225))), System.Drawing.Color.Black, System.Drawing.Color.Empty, System.Drawing.Drawing2D.LinearGradientMode.Horizontal));
			this.gridControl1.Styles.AddReplace("OddRow", new DevExpress.Utils.ViewStyleEx("OddRow", "Grid", new System.Drawing.Font("Tahoma", 10F), ((DevExpress.Utils.StyleOptions)((DevExpress.Utils.StyleOptions.StyleEnabled | DevExpress.Utils.StyleOptions.UseBackColor))), System.Drawing.Color.White, System.Drawing.Color.Black, System.Drawing.Color.White, System.Drawing.Drawing2D.LinearGradientMode.BackwardDiagonal));
			this.gridControl1.Styles.AddReplace("SelectedRow", new DevExpress.Utils.ViewStyleEx("SelectedRow", "Grid", new System.Drawing.Font("Tahoma", 10F), ((DevExpress.Utils.StyleOptions)((((((DevExpress.Utils.StyleOptions.StyleEnabled | DevExpress.Utils.StyleOptions.UseBackColor) 
				| DevExpress.Utils.StyleOptions.UseDrawFocusRect) 
				| DevExpress.Utils.StyleOptions.UseFont) 
				| DevExpress.Utils.StyleOptions.UseForeColor) 
				| DevExpress.Utils.StyleOptions.UseImage))), System.Drawing.Color.FromArgb(((System.Byte)(20)), ((System.Byte)(46)), ((System.Byte)(116))), System.Drawing.Color.White, System.Drawing.Color.Empty, System.Drawing.Drawing2D.LinearGradientMode.Horizontal));
			this.gridControl1.Styles.AddReplace("Row", new DevExpress.Utils.ViewStyleEx("Row", "Grid", new System.Drawing.Font("Tahoma", 10F), DevExpress.Utils.StyleOptions.StyleEnabled, System.Drawing.Color.White, System.Drawing.Color.Black, System.Drawing.Color.Empty, System.Drawing.Drawing2D.LinearGradientMode.Horizontal));
			this.gridControl1.Styles.AddReplace("FilterPanel", new DevExpress.Utils.ViewStyleEx("FilterPanel", "Grid", new System.Drawing.Font("Arial", 8F), "", true, false, false, DevExpress.Utils.HorzAlignment.Near, DevExpress.Utils.VertAlignment.Center, null, System.Drawing.Color.FromArgb(((System.Byte)(28)), ((System.Byte)(80)), ((System.Byte)(135))), System.Drawing.Color.White, System.Drawing.Color.FromArgb(((System.Byte)(212)), ((System.Byte)(208)), ((System.Byte)(200))), System.Drawing.Drawing2D.LinearGradientMode.ForwardDiagonal));
			this.gridControl1.Styles.AddReplace("RowSeparator", new DevExpress.Utils.ViewStyleEx("RowSeparator", "Grid", new System.Drawing.Font("Arial", 8F), DevExpress.Utils.StyleOptions.StyleEnabled, System.Drawing.Color.White, System.Drawing.Color.Gray, System.Drawing.Color.FromArgb(((System.Byte)(232)), ((System.Byte)(228)), ((System.Byte)(220))), System.Drawing.Drawing2D.LinearGradientMode.Horizontal));
			this.gridControl1.Styles.AddReplace("ColumnFilterButtonActive", new DevExpress.Utils.ViewStyleEx("ColumnFilterButtonActive", "Grid", new System.Drawing.Font("Arial", 8F), System.Drawing.Color.LightGray, System.Drawing.Color.Blue, System.Drawing.Color.WhiteSmoke, System.Drawing.Drawing2D.LinearGradientMode.Horizontal));
			this.gridControl1.Styles.AddReplace("DetailTip", new DevExpress.Utils.ViewStyleEx("DetailTip", "Grid", new System.Drawing.Font("Arial", 8F), System.Drawing.Color.FromArgb(((System.Byte)(255)), ((System.Byte)(255)), ((System.Byte)(225))), System.Drawing.Color.Black, System.Drawing.Color.Empty, System.Drawing.Drawing2D.LinearGradientMode.Horizontal));
			this.gridControl1.TabIndex = 0;
			this.gridControl1.Text = "gridControl1";
			// 
			// gridView1
			// 
			this.gridView1.Columns.AddRange(new DevExpress.XtraGrid.Columns.GridColumn[] {
																																										 this.colNumCuenta,
																																										 this.colNumCheque,
																																										 this.colidCuenta,
																																										 this.colidCuenta1,
																																										 this.colAutomatico,
																																										 this.colReporte});
			this.gridView1.Name = "gridView1";
			this.gridView1.OptionsView.ShowGroupPanel = false;
			this.gridView1.OptionsView.ShowNewItemRow = true;
			// 
			// colNumCuenta
			// 
			this.colNumCuenta.Caption = "Nº Cuenta";
			this.colNumCuenta.FieldName = "NumCuenta";
			this.colNumCuenta.Name = "colNumCuenta";
			this.colNumCuenta.VisibleIndex = 0;
			this.colNumCuenta.Width = 90;
			// 
			// colNumCheque
			// 
			this.colNumCheque.Caption = "Nº Cheque";
			this.colNumCheque.FieldName = "NumCheque";
			this.colNumCheque.Name = "colNumCheque";
			this.colNumCheque.VisibleIndex = 1;
			this.colNumCheque.Width = 82;
			// 
			// colidCuenta
			// 
			this.colidCuenta.Caption = "Cuenta";
			this.colidCuenta.ColumnEdit = this.repositoryItemLookUpEdit2;
			this.colidCuenta.FieldName = "idCuenta";
			this.colidCuenta.Name = "colidCuenta";
			this.colidCuenta.VisibleIndex = 2;
			this.colidCuenta.Width = 88;
			// 
			// repositoryItemLookUpEdit2
			// 
			this.repositoryItemLookUpEdit2.AutoHeight = false;
			this.repositoryItemLookUpEdit2.Buttons.AddRange(new DevExpress.XtraEditors.Controls.EditorButton[] {
																																																					 new DevExpress.XtraEditors.Controls.EditorButton(DevExpress.XtraEditors.Controls.ButtonPredefines.Combo)});
			this.repositoryItemLookUpEdit2.Columns.AddRange(new DevExpress.XtraEditors.Controls.LookUpColumnInfo[] {
																																																							 new DevExpress.XtraEditors.Controls.LookUpColumnInfo("Descripcion", "Descripcion", 68, DevExpress.Utils.FormatType.None, "", true, DevExpress.Utils.HorzAlignment.Near),
																																																							 new DevExpress.XtraEditors.Controls.LookUpColumnInfo("Codigo", "Codigo", 45, DevExpress.Utils.FormatType.None, "", true, DevExpress.Utils.HorzAlignment.Near)});
			this.repositoryItemLookUpEdit2.DataSource = this.cdvCodCuenta;
			this.repositoryItemLookUpEdit2.DisplayMember = "Descripcion";
			this.repositoryItemLookUpEdit2.Name = "repositoryItemLookUpEdit2";
			this.repositoryItemLookUpEdit2.NullText = "";
			this.repositoryItemLookUpEdit2.ValueMember = "idCuenta";
			// 
			// cdvCodCuenta
			// 
			this.cdvCodCuenta.BindingContextCtrl = this;
			this.cdvCodCuenta.DataSet = this.cdsCuentas;
			this.cdvCodCuenta.Sort = "Codigo";
			this.cdvCodCuenta.TableName = "";
			this.cdvCodCuenta.TableViewName = "Cuenta";
			// 
			// cdsCuentas
			// 
			this.cdsCuentas.BindingContextCtrl = this;
			this.cdsCuentas.DataLibrary = "LibContabilidad";
			this.cdsCuentas.DataLibraryUrl = "";
			this.cdsCuentas.DataSetDef = "dsCuenta";
			this.cdsCuentas.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsCuentas.Name = "cdsCuentas";
			this.cdsCuentas.SchemaClassName = "LibContabilidad.DataClass";
			this.cdsCuentas.SchemaDef = null;
			this.cdsCuentas.BeforeFill += new C1.Data.FillEventHandler(this.cdsCuentas_BeforeFill);
			// 
			// colidCuenta1
			// 
			this.colidCuenta1.Caption = "Código";
			this.colidCuenta1.ColumnEdit = this.repositoryItemLookUpEdit1;
			this.colidCuenta1.FieldName = "idCuenta";
			this.colidCuenta1.Name = "colidCuenta1";
			this.colidCuenta1.VisibleIndex = 3;
			this.colidCuenta1.Width = 83;
			// 
			// repositoryItemLookUpEdit1
			// 
			this.repositoryItemLookUpEdit1.AutoHeight = false;
			this.repositoryItemLookUpEdit1.Buttons.AddRange(new DevExpress.XtraEditors.Controls.EditorButton[] {
																																																					 new DevExpress.XtraEditors.Controls.EditorButton(DevExpress.XtraEditors.Controls.ButtonPredefines.Combo)});
			this.repositoryItemLookUpEdit1.Columns.AddRange(new DevExpress.XtraEditors.Controls.LookUpColumnInfo[] {
																																																							 new DevExpress.XtraEditors.Controls.LookUpColumnInfo("CodRapido", "CodRapido", 66, DevExpress.Utils.FormatType.None, "", true, DevExpress.Utils.HorzAlignment.Near),
																																																							 new DevExpress.XtraEditors.Controls.LookUpColumnInfo("Descripcion", "Descripcion", 68, DevExpress.Utils.FormatType.None, "", true, DevExpress.Utils.HorzAlignment.Near)});
			this.repositoryItemLookUpEdit1.DataSource = this.cdvCuenta;
			this.repositoryItemLookUpEdit1.DisplayMember = "Codigo";
			this.repositoryItemLookUpEdit1.Name = "repositoryItemLookUpEdit1";
			this.repositoryItemLookUpEdit1.NullText = "";
			this.repositoryItemLookUpEdit1.ValueMember = "idCuenta";
			// 
			// cdvCuenta
			// 
			this.cdvCuenta.BindingContextCtrl = this;
			this.cdvCuenta.DataSet = this.cdsCuentas;
			this.cdvCuenta.Sort = "Descripcion";
			this.cdvCuenta.TableName = "";
			this.cdvCuenta.TableViewName = "Cuenta";
			// 
			// colAutomatico
			// 
			this.colAutomatico.Caption = "Automatico";
			this.colAutomatico.FieldName = "Automatico";
			this.colAutomatico.Name = "colAutomatico";
			this.colAutomatico.VisibleIndex = 4;
			this.colAutomatico.Width = 76;
			// 
			// colReporte
			// 
			this.colReporte.Caption = "Reporte";
			this.colReporte.FieldName = "Reporte";
			this.colReporte.Name = "colReporte";
			this.colReporte.VisibleIndex = 5;
			this.colReporte.Width = 20;
			// 
			// btAceptar
			// 
			this.btAceptar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btAceptar.Location = new System.Drawing.Point(64, 256);
			this.btAceptar.Name = "btAceptar";
			this.btAceptar.TabIndex = 1;
			this.btAceptar.Text = "&Aceptar";
			this.btAceptar.Click += new System.EventHandler(this.btAceptar_Click);
			// 
			// btCancelar
			// 
			this.btCancelar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.btCancelar.Location = new System.Drawing.Point(304, 256);
			this.btCancelar.Name = "btCancelar";
			this.btCancelar.TabIndex = 2;
			this.btCancelar.Text = "&Cancelar";
			this.btCancelar.Click += new System.EventHandler(this.btCancelar_Click);
			// 
			// CuentaCorriente
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(472, 293);
			this.Controls.Add(this.btCancelar);
			this.Controls.Add(this.btAceptar);
			this.Controls.Add(this.gridControl1);
			this.Name = "CuentaCorriente";
			this.Text = "Cuentas Corrientes";
			this.Load += new System.EventHandler(this.CuentaCorriente_Load);
			((System.ComponentModel.ISupportInitialize)(this.cdsCuentaCorr)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvCuentaCorr)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.gridControl1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.gridView1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.repositoryItemLookUpEdit2)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvCodCuenta)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsCuentas)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.repositoryItemLookUpEdit1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvCuenta)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void btCancelar_Click(object sender, System.EventArgs e)
		{
			this.Close();
		}

		private void btAceptar_Click(object sender, System.EventArgs e)
		{
			try
			{
				this.cdsCuentaCorr.Update();
				this.Close();
			}
			catch(System.Exception ex)
			{
				MessageBox.Show(ex.Message);
			}
		}

		private void cdsCuentaCorr_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			this.cdsCuentaCorr.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirConta);
		}

		private void cdsCuentas_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			this.cdsCuentas.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirConta);
		}

		private void CuentaCorriente_Load(object sender, System.EventArgs e)
		{
			Cursor = Cursors.Default;
		}
	}
}
