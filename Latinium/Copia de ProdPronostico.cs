using System;
using C1.Data;
using System.Data;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de ProdPronostico.
	/// </summary>
	public class ProdPronostico : DevExpress.XtraEditors.XtraForm
	{
		private DevExpress.XtraGrid.GridControl gridControl1;
		private DevExpress.XtraGrid.Views.Grid.GridView gridView1;
		private DevExpress.XtraEditors.SimpleButton btAceptar;
		private DevExpress.XtraEditors.SimpleButton btCancelar;
		private C1.Data.C1DataSet cdsProdPronostico;
		private C1.Data.C1DataView cdvProdPronostico;
		private DevExpress.XtraGrid.Columns.GridColumn colidCliente;
		private DevExpress.XtraGrid.Columns.GridColumn colMes1;
		private DevExpress.XtraGrid.Columns.GridColumn colMes10;
		private DevExpress.XtraGrid.Columns.GridColumn colMes11;
		private DevExpress.XtraGrid.Columns.GridColumn colMes12;
		private DevExpress.XtraGrid.Columns.GridColumn colMes2;
		private DevExpress.XtraGrid.Columns.GridColumn colMes3;
		private DevExpress.XtraGrid.Columns.GridColumn colMes4;
		private DevExpress.XtraGrid.Columns.GridColumn colMes5;
		private DevExpress.XtraGrid.Columns.GridColumn colMes6;
		private DevExpress.XtraGrid.Columns.GridColumn colMes7;
		private DevExpress.XtraGrid.Columns.GridColumn colMes8;
		private DevExpress.XtraGrid.Columns.GridColumn colMes9;
		private C1.Data.C1DataSet cdsCliente;
		private C1.Data.C1DataView cdvCliente;
		private DevExpress.XtraEditors.Repository.RepositoryItemLookUpEdit repositoryItemLookUpEdit1;
		private C1.Data.C1DataSet cdsArticulo;
		private C1.Data.C1DataView cdvArticulo;
		private DevExpress.XtraGrid.Columns.GridColumn colidArticulo;
		private DevExpress.XtraEditors.Repository.RepositoryItemLookUpEdit repositoryItemLookUpEdit2;
		private DevExpress.XtraGrid.Columns.GridColumn colTableField;
		private DevExpress.XtraGrid.Columns.GridColumn colTableField_1;
		private DevExpress.XtraEditors.SimpleButton btImprimir;
		private C1.Data.C1DataSet cdsDetCompra;
		private C1.Data.C1DataView cdvDetCompra;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public ProdPronostico()
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
			DevExpress.XtraGrid.StyleFormatCondition styleFormatCondition1 = new DevExpress.XtraGrid.StyleFormatCondition();
			this.gridControl1 = new DevExpress.XtraGrid.GridControl();
			this.cdvProdPronostico = new C1.Data.C1DataView();
			this.cdsProdPronostico = new C1.Data.C1DataSet();
			this.gridView1 = new DevExpress.XtraGrid.Views.Grid.GridView();
			this.colidCliente = new DevExpress.XtraGrid.Columns.GridColumn();
			this.repositoryItemLookUpEdit1 = new DevExpress.XtraEditors.Repository.RepositoryItemLookUpEdit();
			this.cdvCliente = new C1.Data.C1DataView();
			this.cdsCliente = new C1.Data.C1DataSet();
			this.colidArticulo = new DevExpress.XtraGrid.Columns.GridColumn();
			this.repositoryItemLookUpEdit2 = new DevExpress.XtraEditors.Repository.RepositoryItemLookUpEdit();
			this.cdvArticulo = new C1.Data.C1DataView();
			this.cdsArticulo = new C1.Data.C1DataSet();
			this.colMes1 = new DevExpress.XtraGrid.Columns.GridColumn();
			this.colMes2 = new DevExpress.XtraGrid.Columns.GridColumn();
			this.colMes3 = new DevExpress.XtraGrid.Columns.GridColumn();
			this.colMes4 = new DevExpress.XtraGrid.Columns.GridColumn();
			this.colMes5 = new DevExpress.XtraGrid.Columns.GridColumn();
			this.colMes6 = new DevExpress.XtraGrid.Columns.GridColumn();
			this.colMes7 = new DevExpress.XtraGrid.Columns.GridColumn();
			this.colMes8 = new DevExpress.XtraGrid.Columns.GridColumn();
			this.colMes9 = new DevExpress.XtraGrid.Columns.GridColumn();
			this.colMes10 = new DevExpress.XtraGrid.Columns.GridColumn();
			this.colMes11 = new DevExpress.XtraGrid.Columns.GridColumn();
			this.colMes12 = new DevExpress.XtraGrid.Columns.GridColumn();
			this.colTableField = new DevExpress.XtraGrid.Columns.GridColumn();
			this.colTableField_1 = new DevExpress.XtraGrid.Columns.GridColumn();
			this.btAceptar = new DevExpress.XtraEditors.SimpleButton();
			this.btCancelar = new DevExpress.XtraEditors.SimpleButton();
			this.btImprimir = new DevExpress.XtraEditors.SimpleButton();
			this.cdsDetCompra = new C1.Data.C1DataSet();
			this.cdvDetCompra = new C1.Data.C1DataView();
			((System.ComponentModel.ISupportInitialize)(this.gridControl1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvProdPronostico)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsProdPronostico)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.gridView1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.repositoryItemLookUpEdit1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvCliente)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsCliente)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.repositoryItemLookUpEdit2)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvArticulo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsArticulo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsDetCompra)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvDetCompra)).BeginInit();
			this.SuspendLayout();
			// 
			// gridControl1
			// 
			this.gridControl1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.gridControl1.DataSource = this.cdvProdPronostico;
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
			this.gridControl1.Size = new System.Drawing.Size(656, 272);
			this.gridControl1.Styles.AddReplace("Style1", new DevExpress.Utils.ViewStyleEx("Style1", "", System.Drawing.Color.LightCoral, System.Drawing.SystemColors.HotTrack, System.Drawing.Color.Aqua, System.Drawing.Drawing2D.LinearGradientMode.Horizontal));
			this.gridControl1.TabIndex = 0;
			this.gridControl1.ViewCollection.AddRange(new DevExpress.XtraGrid.Views.Base.BaseView[] {
																																																this.gridView1});
			// 
			// cdvProdPronostico
			// 
			this.cdvProdPronostico.BindingContextCtrl = this;
			this.cdvProdPronostico.DataSet = this.cdsProdPronostico;
			this.cdvProdPronostico.TableName = "";
			this.cdvProdPronostico.TableViewName = "ProdPronostico";
			// 
			// cdsProdPronostico
			// 
			this.cdsProdPronostico.BindingContextCtrl = this;
			this.cdsProdPronostico.DataLibrary = "LibFacturacion";
			this.cdsProdPronostico.DataLibraryUrl = "";
			this.cdsProdPronostico.DataSetDef = "dsProdPronostico";
			this.cdsProdPronostico.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsProdPronostico.Name = "cdsProdPronostico";
			this.cdsProdPronostico.SchemaClassName = "LibFacturacion.DataClass";
			this.cdsProdPronostico.SchemaDef = null;
			this.cdsProdPronostico.BeforeFill += new C1.Data.FillEventHandler(this.cdsProdPronostico_BeforeFill);
			// 
			// gridView1
			// 
			this.gridView1.Appearance.ColumnFilterButton.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(212)), ((System.Byte)(208)), ((System.Byte)(200)));
			this.gridView1.Appearance.ColumnFilterButton.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(232)), ((System.Byte)(228)), ((System.Byte)(220)));
			this.gridView1.Appearance.ColumnFilterButton.BorderColor = System.Drawing.Color.FromArgb(((System.Byte)(212)), ((System.Byte)(208)), ((System.Byte)(200)));
			this.gridView1.Appearance.ColumnFilterButton.Font = new System.Drawing.Font("Arial", 8F);
			this.gridView1.Appearance.ColumnFilterButton.ForeColor = System.Drawing.Color.Gray;
			this.gridView1.Appearance.ColumnFilterButton.Options.UseBackColor = true;
			this.gridView1.Appearance.ColumnFilterButton.Options.UseBorderColor = true;
			this.gridView1.Appearance.ColumnFilterButton.Options.UseFont = true;
			this.gridView1.Appearance.ColumnFilterButton.Options.UseForeColor = true;
			this.gridView1.Appearance.ColumnFilterButtonActive.BackColor = System.Drawing.Color.LightGray;
			this.gridView1.Appearance.ColumnFilterButtonActive.BackColor2 = System.Drawing.Color.WhiteSmoke;
			this.gridView1.Appearance.ColumnFilterButtonActive.BorderColor = System.Drawing.Color.LightGray;
			this.gridView1.Appearance.ColumnFilterButtonActive.Font = new System.Drawing.Font("Arial", 8F);
			this.gridView1.Appearance.ColumnFilterButtonActive.Options.UseBackColor = true;
			this.gridView1.Appearance.ColumnFilterButtonActive.Options.UseBorderColor = true;
			this.gridView1.Appearance.ColumnFilterButtonActive.Options.UseFont = true;
			this.gridView1.Appearance.DetailTip.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(255)), ((System.Byte)(255)), ((System.Byte)(225)));
			this.gridView1.Appearance.DetailTip.Font = new System.Drawing.Font("Arial", 8F);
			this.gridView1.Appearance.DetailTip.ForeColor = System.Drawing.Color.Black;
			this.gridView1.Appearance.DetailTip.Options.UseBackColor = true;
			this.gridView1.Appearance.DetailTip.Options.UseFont = true;
			this.gridView1.Appearance.DetailTip.Options.UseForeColor = true;
			this.gridView1.Appearance.Empty.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(232)), ((System.Byte)(228)), ((System.Byte)(220)));
			this.gridView1.Appearance.Empty.Font = new System.Drawing.Font("Arial", 8F);
			this.gridView1.Appearance.Empty.ForeColor = System.Drawing.Color.White;
			this.gridView1.Appearance.Empty.Options.UseBackColor = true;
			this.gridView1.Appearance.Empty.Options.UseFont = true;
			this.gridView1.Appearance.Empty.Options.UseForeColor = true;
			this.gridView1.Appearance.EvenRow.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(212)), ((System.Byte)(208)), ((System.Byte)(200)));
			this.gridView1.Appearance.EvenRow.BackColor2 = System.Drawing.Color.GhostWhite;
			this.gridView1.Appearance.EvenRow.Font = new System.Drawing.Font("Tahoma", 10F);
			this.gridView1.Appearance.EvenRow.ForeColor = System.Drawing.Color.Black;
			this.gridView1.Appearance.EvenRow.GradientMode = System.Drawing.Drawing2D.LinearGradientMode.ForwardDiagonal;
			this.gridView1.Appearance.EvenRow.Options.UseBackColor = true;
			this.gridView1.Appearance.EvenRow.Options.UseFont = true;
			this.gridView1.Appearance.EvenRow.Options.UseForeColor = true;
			this.gridView1.Appearance.FilterCloseButton.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(212)), ((System.Byte)(208)), ((System.Byte)(200)));
			this.gridView1.Appearance.FilterCloseButton.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(118)), ((System.Byte)(170)), ((System.Byte)(225)));
			this.gridView1.Appearance.FilterCloseButton.BorderColor = System.Drawing.Color.FromArgb(((System.Byte)(212)), ((System.Byte)(208)), ((System.Byte)(200)));
			this.gridView1.Appearance.FilterCloseButton.Font = new System.Drawing.Font("Arial", 8F);
			this.gridView1.Appearance.FilterCloseButton.ForeColor = System.Drawing.Color.Black;
			this.gridView1.Appearance.FilterCloseButton.GradientMode = System.Drawing.Drawing2D.LinearGradientMode.ForwardDiagonal;
			this.gridView1.Appearance.FilterCloseButton.Options.UseBackColor = true;
			this.gridView1.Appearance.FilterCloseButton.Options.UseBorderColor = true;
			this.gridView1.Appearance.FilterCloseButton.Options.UseFont = true;
			this.gridView1.Appearance.FilterCloseButton.Options.UseForeColor = true;
			this.gridView1.Appearance.FilterPanel.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(28)), ((System.Byte)(80)), ((System.Byte)(135)));
			this.gridView1.Appearance.FilterPanel.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(212)), ((System.Byte)(208)), ((System.Byte)(200)));
			this.gridView1.Appearance.FilterPanel.Font = new System.Drawing.Font("Arial", 8F);
			this.gridView1.Appearance.FilterPanel.ForeColor = System.Drawing.Color.White;
			this.gridView1.Appearance.FilterPanel.Options.UseBackColor = true;
			this.gridView1.Appearance.FilterPanel.Options.UseFont = true;
			this.gridView1.Appearance.FilterPanel.Options.UseForeColor = true;
			this.gridView1.Appearance.FixedLine.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(58)), ((System.Byte)(58)), ((System.Byte)(58)));
			this.gridView1.Appearance.FixedLine.Font = new System.Drawing.Font("Arial", 8F);
			this.gridView1.Appearance.FixedLine.ForeColor = System.Drawing.Color.Black;
			this.gridView1.Appearance.FixedLine.Options.UseBackColor = true;
			this.gridView1.Appearance.FixedLine.Options.UseFont = true;
			this.gridView1.Appearance.FixedLine.Options.UseForeColor = true;
			this.gridView1.Appearance.FocusedCell.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(255)), ((System.Byte)(255)), ((System.Byte)(225)));
			this.gridView1.Appearance.FocusedCell.Font = new System.Drawing.Font("Tahoma", 10F);
			this.gridView1.Appearance.FocusedCell.ForeColor = System.Drawing.Color.Black;
			this.gridView1.Appearance.FocusedCell.Options.UseBackColor = true;
			this.gridView1.Appearance.FocusedCell.Options.UseFont = true;
			this.gridView1.Appearance.FocusedCell.Options.UseForeColor = true;
			this.gridView1.Appearance.FocusedRow.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(10)), ((System.Byte)(36)), ((System.Byte)(106)));
			this.gridView1.Appearance.FocusedRow.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(60)), ((System.Byte)(86)), ((System.Byte)(156)));
			this.gridView1.Appearance.FocusedRow.Font = new System.Drawing.Font("Tahoma", 10F);
			this.gridView1.Appearance.FocusedRow.ForeColor = System.Drawing.Color.White;
			this.gridView1.Appearance.FocusedRow.Options.UseBackColor = true;
			this.gridView1.Appearance.FocusedRow.Options.UseFont = true;
			this.gridView1.Appearance.FocusedRow.Options.UseForeColor = true;
			this.gridView1.Appearance.FooterPanel.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(212)), ((System.Byte)(208)), ((System.Byte)(200)));
			this.gridView1.Appearance.FooterPanel.BorderColor = System.Drawing.Color.FromArgb(((System.Byte)(212)), ((System.Byte)(208)), ((System.Byte)(200)));
			this.gridView1.Appearance.FooterPanel.Font = new System.Drawing.Font("Tahoma", 10F);
			this.gridView1.Appearance.FooterPanel.ForeColor = System.Drawing.Color.Black;
			this.gridView1.Appearance.FooterPanel.Options.UseBackColor = true;
			this.gridView1.Appearance.FooterPanel.Options.UseBorderColor = true;
			this.gridView1.Appearance.FooterPanel.Options.UseFont = true;
			this.gridView1.Appearance.FooterPanel.Options.UseForeColor = true;
			this.gridView1.Appearance.GroupButton.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(212)), ((System.Byte)(208)), ((System.Byte)(200)));
			this.gridView1.Appearance.GroupButton.BorderColor = System.Drawing.Color.FromArgb(((System.Byte)(212)), ((System.Byte)(208)), ((System.Byte)(200)));
			this.gridView1.Appearance.GroupButton.Font = new System.Drawing.Font("Tahoma", 10F);
			this.gridView1.Appearance.GroupButton.ForeColor = System.Drawing.Color.Black;
			this.gridView1.Appearance.GroupButton.Options.UseBackColor = true;
			this.gridView1.Appearance.GroupButton.Options.UseBorderColor = true;
			this.gridView1.Appearance.GroupButton.Options.UseFont = true;
			this.gridView1.Appearance.GroupButton.Options.UseForeColor = true;
			this.gridView1.Appearance.GroupFooter.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(222)), ((System.Byte)(218)), ((System.Byte)(210)));
			this.gridView1.Appearance.GroupFooter.BorderColor = System.Drawing.Color.FromArgb(((System.Byte)(222)), ((System.Byte)(218)), ((System.Byte)(210)));
			this.gridView1.Appearance.GroupFooter.Font = new System.Drawing.Font("Tahoma", 8F);
			this.gridView1.Appearance.GroupFooter.ForeColor = System.Drawing.Color.Black;
			this.gridView1.Appearance.GroupFooter.Options.UseBackColor = true;
			this.gridView1.Appearance.GroupFooter.Options.UseBorderColor = true;
			this.gridView1.Appearance.GroupFooter.Options.UseFont = true;
			this.gridView1.Appearance.GroupFooter.Options.UseForeColor = true;
			this.gridView1.Appearance.GroupPanel.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(58)), ((System.Byte)(110)), ((System.Byte)(165)));
			this.gridView1.Appearance.GroupPanel.BackColor2 = System.Drawing.Color.White;
			this.gridView1.Appearance.GroupPanel.Font = new System.Drawing.Font("Tahoma", 10F, System.Drawing.FontStyle.Bold);
			this.gridView1.Appearance.GroupPanel.ForeColor = System.Drawing.Color.White;
			this.gridView1.Appearance.GroupPanel.Options.UseBackColor = true;
			this.gridView1.Appearance.GroupPanel.Options.UseFont = true;
			this.gridView1.Appearance.GroupPanel.Options.UseForeColor = true;
			this.gridView1.Appearance.GroupRow.BackColor = System.Drawing.Color.Gray;
			this.gridView1.Appearance.GroupRow.Font = new System.Drawing.Font("Tahoma", 10F);
			this.gridView1.Appearance.GroupRow.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(212)), ((System.Byte)(208)), ((System.Byte)(200)));
			this.gridView1.Appearance.GroupRow.Options.UseBackColor = true;
			this.gridView1.Appearance.GroupRow.Options.UseFont = true;
			this.gridView1.Appearance.GroupRow.Options.UseForeColor = true;
			this.gridView1.Appearance.HeaderPanel.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(212)), ((System.Byte)(208)), ((System.Byte)(200)));
			this.gridView1.Appearance.HeaderPanel.BorderColor = System.Drawing.Color.FromArgb(((System.Byte)(212)), ((System.Byte)(208)), ((System.Byte)(200)));
			this.gridView1.Appearance.HeaderPanel.Font = new System.Drawing.Font("Tahoma", 10F, System.Drawing.FontStyle.Bold);
			this.gridView1.Appearance.HeaderPanel.ForeColor = System.Drawing.Color.Black;
			this.gridView1.Appearance.HeaderPanel.Options.UseBackColor = true;
			this.gridView1.Appearance.HeaderPanel.Options.UseBorderColor = true;
			this.gridView1.Appearance.HeaderPanel.Options.UseFont = true;
			this.gridView1.Appearance.HeaderPanel.Options.UseForeColor = true;
			this.gridView1.Appearance.HideSelectionRow.BackColor = System.Drawing.Color.Gray;
			this.gridView1.Appearance.HideSelectionRow.Font = new System.Drawing.Font("Tahoma", 10F);
			this.gridView1.Appearance.HideSelectionRow.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(212)), ((System.Byte)(208)), ((System.Byte)(200)));
			this.gridView1.Appearance.HideSelectionRow.Options.UseBackColor = true;
			this.gridView1.Appearance.HideSelectionRow.Options.UseFont = true;
			this.gridView1.Appearance.HideSelectionRow.Options.UseForeColor = true;
			this.gridView1.Appearance.HorzLine.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(212)), ((System.Byte)(208)), ((System.Byte)(200)));
			this.gridView1.Appearance.HorzLine.Font = new System.Drawing.Font("Tahoma", 10F);
			this.gridView1.Appearance.HorzLine.ForeColor = System.Drawing.Color.Gray;
			this.gridView1.Appearance.HorzLine.Options.UseBackColor = true;
			this.gridView1.Appearance.HorzLine.Options.UseFont = true;
			this.gridView1.Appearance.HorzLine.Options.UseForeColor = true;
			this.gridView1.Appearance.OddRow.BackColor = System.Drawing.Color.White;
			this.gridView1.Appearance.OddRow.BackColor2 = System.Drawing.Color.White;
			this.gridView1.Appearance.OddRow.Font = new System.Drawing.Font("Tahoma", 10F);
			this.gridView1.Appearance.OddRow.ForeColor = System.Drawing.Color.Black;
			this.gridView1.Appearance.OddRow.GradientMode = System.Drawing.Drawing2D.LinearGradientMode.BackwardDiagonal;
			this.gridView1.Appearance.OddRow.Options.UseBackColor = true;
			this.gridView1.Appearance.OddRow.Options.UseFont = true;
			this.gridView1.Appearance.OddRow.Options.UseForeColor = true;
			this.gridView1.Appearance.Preview.BackColor = System.Drawing.Color.White;
			this.gridView1.Appearance.Preview.BackColor2 = System.Drawing.Color.White;
			this.gridView1.Appearance.Preview.Font = new System.Drawing.Font("Tahoma", 10F);
			this.gridView1.Appearance.Preview.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(10)), ((System.Byte)(36)), ((System.Byte)(106)));
			this.gridView1.Appearance.Preview.Options.UseBackColor = true;
			this.gridView1.Appearance.Preview.Options.UseFont = true;
			this.gridView1.Appearance.Preview.Options.UseForeColor = true;
			this.gridView1.Appearance.Row.BackColor = System.Drawing.Color.White;
			this.gridView1.Appearance.Row.Font = new System.Drawing.Font("Tahoma", 10F);
			this.gridView1.Appearance.Row.ForeColor = System.Drawing.Color.Black;
			this.gridView1.Appearance.Row.Options.UseBackColor = true;
			this.gridView1.Appearance.Row.Options.UseFont = true;
			this.gridView1.Appearance.Row.Options.UseForeColor = true;
			this.gridView1.Appearance.RowSeparator.BackColor = System.Drawing.Color.White;
			this.gridView1.Appearance.RowSeparator.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(232)), ((System.Byte)(228)), ((System.Byte)(220)));
			this.gridView1.Appearance.RowSeparator.Font = new System.Drawing.Font("Arial", 8F);
			this.gridView1.Appearance.RowSeparator.ForeColor = System.Drawing.Color.Gray;
			this.gridView1.Appearance.RowSeparator.Options.UseBackColor = true;
			this.gridView1.Appearance.RowSeparator.Options.UseFont = true;
			this.gridView1.Appearance.RowSeparator.Options.UseForeColor = true;
			this.gridView1.Appearance.SelectedRow.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(20)), ((System.Byte)(46)), ((System.Byte)(116)));
			this.gridView1.Appearance.SelectedRow.Font = new System.Drawing.Font("Tahoma", 10F);
			this.gridView1.Appearance.SelectedRow.ForeColor = System.Drawing.Color.White;
			this.gridView1.Appearance.SelectedRow.Options.UseBackColor = true;
			this.gridView1.Appearance.SelectedRow.Options.UseFont = true;
			this.gridView1.Appearance.SelectedRow.Options.UseForeColor = true;
			this.gridView1.Appearance.VertLine.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(212)), ((System.Byte)(208)), ((System.Byte)(200)));
			this.gridView1.Appearance.VertLine.Font = new System.Drawing.Font("Tahoma", 10F);
			this.gridView1.Appearance.VertLine.ForeColor = System.Drawing.Color.Gray;
			this.gridView1.Appearance.VertLine.Options.UseBackColor = true;
			this.gridView1.Appearance.VertLine.Options.UseFont = true;
			this.gridView1.Appearance.VertLine.Options.UseForeColor = true;
			this.gridView1.Columns.AddRange(new DevExpress.XtraGrid.Columns.GridColumn[] {
																																										 this.colidCliente,
																																										 this.colidArticulo,
																																										 this.colMes1,
																																										 this.colMes2,
																																										 this.colMes3,
																																										 this.colMes4,
																																										 this.colMes5,
																																										 this.colMes6,
																																										 this.colMes7,
																																										 this.colMes8,
																																										 this.colMes9,
																																										 this.colMes10,
																																										 this.colMes11,
																																										 this.colMes12,
																																										 this.colTableField,
																																										 this.colTableField_1});
			styleFormatCondition1.Appearance.BackColor = System.Drawing.Color.LightCoral;
			styleFormatCondition1.Appearance.BackColor2 = System.Drawing.Color.Aqua;
			styleFormatCondition1.Appearance.ForeColor = System.Drawing.SystemColors.HotTrack;
			styleFormatCondition1.Appearance.Options.UseBackColor = true;
			styleFormatCondition1.Appearance.Options.UseForeColor = true;
			styleFormatCondition1.Column = this.colMes1;
			styleFormatCondition1.Condition = DevExpress.XtraGrid.FormatConditionEnum.LessOrEqual;
			styleFormatCondition1.Value1 = 1;
			this.gridView1.FormatConditions.AddRange(new DevExpress.XtraGrid.StyleFormatCondition[] {
																																																styleFormatCondition1});
			this.gridView1.GridControl = this.gridControl1;
			this.gridView1.GroupPanelText = "Arrastre columna a agrupar";
			this.gridView1.GroupSummary.AddRange(new DevExpress.XtraGrid.GridSummaryItem[] {
																																											 new DevExpress.XtraGrid.GridGroupSummaryItem(DevExpress.Data.SummaryItemType.Sum, "Mes1", this.colMes1, "{0:n2}"),
																																											 new DevExpress.XtraGrid.GridGroupSummaryItem(DevExpress.Data.SummaryItemType.Sum, "Mes2", this.colMes2, "{0:n2}"),
																																											 new DevExpress.XtraGrid.GridGroupSummaryItem(DevExpress.Data.SummaryItemType.Sum, "Mes3", this.colMes3, "{0:n2}"),
																																											 new DevExpress.XtraGrid.GridGroupSummaryItem(DevExpress.Data.SummaryItemType.Sum, "Mes4", this.colMes4, "{0:n2}"),
																																											 new DevExpress.XtraGrid.GridGroupSummaryItem(DevExpress.Data.SummaryItemType.Sum, "Mes5", this.colMes5, "{0:n2}"),
																																											 new DevExpress.XtraGrid.GridGroupSummaryItem(DevExpress.Data.SummaryItemType.Sum, "Mes6", this.colMes6, "{0:n2}"),
																																											 new DevExpress.XtraGrid.GridGroupSummaryItem(DevExpress.Data.SummaryItemType.Sum, "Mes7", this.colMes7, "{0:n2}"),
																																											 new DevExpress.XtraGrid.GridGroupSummaryItem(DevExpress.Data.SummaryItemType.Sum, "Mes8", this.colMes8, "{0:n2}"),
																																											 new DevExpress.XtraGrid.GridGroupSummaryItem(DevExpress.Data.SummaryItemType.Sum, "Mes9", this.colMes9, "{0:n2}"),
																																											 new DevExpress.XtraGrid.GridGroupSummaryItem(DevExpress.Data.SummaryItemType.Sum, "Mes10", this.colMes10, "{0:n2}"),
																																											 new DevExpress.XtraGrid.GridGroupSummaryItem(DevExpress.Data.SummaryItemType.Sum, "Mes11", this.colMes11, "{0:n2}"),
																																											 new DevExpress.XtraGrid.GridGroupSummaryItem(DevExpress.Data.SummaryItemType.Sum, "Mes12", this.colMes12, "{0:n2}"),
																																											 new DevExpress.XtraGrid.GridGroupSummaryItem(DevExpress.Data.SummaryItemType.Sum, "TableField", this.colTableField, "{0:n2}"),
																																											 new DevExpress.XtraGrid.GridGroupSummaryItem(DevExpress.Data.SummaryItemType.Average, "TableField_1", this.colTableField_1, "{0:n2}"),
																																											 new DevExpress.XtraGrid.GridGroupSummaryItem(DevExpress.Data.SummaryItemType.Count, "idArticulo", this.colidArticulo, "")});
			this.gridView1.Name = "gridView1";
			this.gridView1.OptionsView.ColumnAutoWidth = false;
			this.gridView1.OptionsView.NewItemRowPosition = DevExpress.XtraGrid.Views.Grid.NewItemRowPosition.Bottom;
			this.gridView1.OptionsView.ShowFooter = true;
			// 
			// colidCliente
			// 
			this.colidCliente.Caption = "Cliente";
			this.colidCliente.ColumnEdit = this.repositoryItemLookUpEdit1;
			this.colidCliente.FieldName = "idCliente";
			this.colidCliente.Fixed = DevExpress.XtraGrid.Columns.FixedStyle.Left;
			this.colidCliente.Name = "colidCliente";
			this.colidCliente.Visible = true;
			this.colidCliente.VisibleIndex = 0;
			this.colidCliente.Width = 163;
			// 
			// repositoryItemLookUpEdit1
			// 
			this.repositoryItemLookUpEdit1.AutoHeight = false;
			this.repositoryItemLookUpEdit1.Buttons.AddRange(new DevExpress.XtraEditors.Controls.EditorButton[] {
																																																					 new DevExpress.XtraEditors.Controls.EditorButton(DevExpress.XtraEditors.Controls.ButtonPredefines.Combo)});
			this.repositoryItemLookUpEdit1.Columns.AddRange(new DevExpress.XtraEditors.Controls.LookUpColumnInfo[] {
																																																							 new DevExpress.XtraEditors.Controls.LookUpColumnInfo("Nombre", "Nombre", 49, DevExpress.Utils.FormatType.None, "", true, DevExpress.Utils.HorzAlignment.Near),
																																																							 new DevExpress.XtraEditors.Controls.LookUpColumnInfo("Codigo", "Codigo", 45, DevExpress.Utils.FormatType.None, "", true, DevExpress.Utils.HorzAlignment.Near)});
			this.repositoryItemLookUpEdit1.DataSource = this.cdvCliente;
			this.repositoryItemLookUpEdit1.DisplayMember = "Nombre";
			this.repositoryItemLookUpEdit1.Name = "repositoryItemLookUpEdit1";
			this.repositoryItemLookUpEdit1.NullText = "";
			this.repositoryItemLookUpEdit1.ValueMember = "idCliente";
			// 
			// cdvCliente
			// 
			this.cdvCliente.BindingContextCtrl = this;
			this.cdvCliente.DataSet = this.cdsCliente;
			this.cdvCliente.Sort = "Nombre";
			this.cdvCliente.TableName = "";
			this.cdvCliente.TableViewName = "Cliente";
			// 
			// cdsCliente
			// 
			this.cdsCliente.BindingContextCtrl = this;
			this.cdsCliente.DataLibrary = "LibFacturacion";
			this.cdsCliente.DataLibraryUrl = "";
			this.cdsCliente.DataSetDef = "dsClientes";
			this.cdsCliente.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsCliente.Name = "cdsCliente";
			this.cdsCliente.SchemaClassName = "LibFacturacion.DataClass";
			this.cdsCliente.SchemaDef = null;
			this.cdsCliente.BeforeFill += new C1.Data.FillEventHandler(this.cdsCliente_BeforeFill);
			// 
			// colidArticulo
			// 
			this.colidArticulo.Caption = "Producto";
			this.colidArticulo.ColumnEdit = this.repositoryItemLookUpEdit2;
			this.colidArticulo.FieldName = "idArticulo";
			this.colidArticulo.Fixed = DevExpress.XtraGrid.Columns.FixedStyle.Left;
			this.colidArticulo.Name = "colidArticulo";
			this.colidArticulo.Visible = true;
			this.colidArticulo.VisibleIndex = 1;
			// 
			// repositoryItemLookUpEdit2
			// 
			this.repositoryItemLookUpEdit2.AutoHeight = false;
			this.repositoryItemLookUpEdit2.Buttons.AddRange(new DevExpress.XtraEditors.Controls.EditorButton[] {
																																																					 new DevExpress.XtraEditors.Controls.EditorButton(DevExpress.XtraEditors.Controls.ButtonPredefines.Combo)});
			this.repositoryItemLookUpEdit2.Columns.AddRange(new DevExpress.XtraEditors.Controls.LookUpColumnInfo[] {
																																																							 new DevExpress.XtraEditors.Controls.LookUpColumnInfo("Articulo", "Articulo", 47, DevExpress.Utils.FormatType.None, "", true, DevExpress.Utils.HorzAlignment.Near),
																																																							 new DevExpress.XtraEditors.Controls.LookUpColumnInfo("Codigo", "Codigo", 45, DevExpress.Utils.FormatType.None, "", true, DevExpress.Utils.HorzAlignment.Near)});
			this.repositoryItemLookUpEdit2.DataSource = this.cdvArticulo;
			this.repositoryItemLookUpEdit2.DisplayMember = "Articulo";
			this.repositoryItemLookUpEdit2.Name = "repositoryItemLookUpEdit2";
			this.repositoryItemLookUpEdit2.NullText = "";
			this.repositoryItemLookUpEdit2.ValueMember = "idArticulo";
			// 
			// cdvArticulo
			// 
			this.cdvArticulo.BindingContextCtrl = this;
			this.cdvArticulo.DataSet = this.cdsArticulo;
			this.cdvArticulo.RowFilter = "idTipoGrupo = 6";
			this.cdvArticulo.TableName = "";
			this.cdvArticulo.TableViewName = "Articulo";
			// 
			// cdsArticulo
			// 
			this.cdsArticulo.BindingContextCtrl = this;
			this.cdsArticulo.DataLibrary = "LibFacturacion";
			this.cdsArticulo.DataLibraryUrl = "";
			this.cdsArticulo.DataSetDef = "dsArticulos";
			this.cdsArticulo.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsArticulo.Name = "cdsArticulo";
			this.cdsArticulo.SchemaClassName = "LibFacturacion.DataClass";
			this.cdsArticulo.SchemaDef = null;
			this.cdsArticulo.BeforeFill += new C1.Data.FillEventHandler(this.cdsArticulo_BeforeFill);
			// 
			// colMes1
			// 
			this.colMes1.Caption = "Enero";
			this.colMes1.DisplayFormat.FormatString = "{0:n2}";
			this.colMes1.DisplayFormat.FormatType = DevExpress.Utils.FormatType.Numeric;
			this.colMes1.FieldName = "Mes1";
			this.colMes1.Name = "colMes1";
			this.colMes1.SummaryItem.DisplayFormat = "{0:n2}";
			this.colMes1.SummaryItem.SummaryType = DevExpress.Data.SummaryItemType.Sum;
			this.colMes1.Visible = true;
			this.colMes1.VisibleIndex = 2;
			// 
			// colMes2
			// 
			this.colMes2.Caption = "Febrero";
			this.colMes2.DisplayFormat.FormatString = "{0:n2}";
			this.colMes2.DisplayFormat.FormatType = DevExpress.Utils.FormatType.Numeric;
			this.colMes2.FieldName = "Mes2";
			this.colMes2.Name = "colMes2";
			this.colMes2.SummaryItem.DisplayFormat = "{0:n2}";
			this.colMes2.SummaryItem.SummaryType = DevExpress.Data.SummaryItemType.Sum;
			this.colMes2.Visible = true;
			this.colMes2.VisibleIndex = 3;
			// 
			// colMes3
			// 
			this.colMes3.Caption = "Marzo";
			this.colMes3.DisplayFormat.FormatString = "{0:n2}";
			this.colMes3.DisplayFormat.FormatType = DevExpress.Utils.FormatType.Numeric;
			this.colMes3.FieldName = "Mes3";
			this.colMes3.Name = "colMes3";
			this.colMes3.SummaryItem.DisplayFormat = "{0:n2}";
			this.colMes3.SummaryItem.SummaryType = DevExpress.Data.SummaryItemType.Sum;
			this.colMes3.Visible = true;
			this.colMes3.VisibleIndex = 4;
			// 
			// colMes4
			// 
			this.colMes4.Caption = "Abril";
			this.colMes4.DisplayFormat.FormatString = "{0:n2}";
			this.colMes4.DisplayFormat.FormatType = DevExpress.Utils.FormatType.Numeric;
			this.colMes4.FieldName = "Mes4";
			this.colMes4.Name = "colMes4";
			this.colMes4.SummaryItem.DisplayFormat = "{0:n2}";
			this.colMes4.SummaryItem.SummaryType = DevExpress.Data.SummaryItemType.Sum;
			this.colMes4.Visible = true;
			this.colMes4.VisibleIndex = 5;
			// 
			// colMes5
			// 
			this.colMes5.Caption = "Mayo";
			this.colMes5.DisplayFormat.FormatString = "{0:n2}";
			this.colMes5.DisplayFormat.FormatType = DevExpress.Utils.FormatType.Numeric;
			this.colMes5.FieldName = "Mes5";
			this.colMes5.Name = "colMes5";
			this.colMes5.SummaryItem.DisplayFormat = "{0:n2}";
			this.colMes5.SummaryItem.SummaryType = DevExpress.Data.SummaryItemType.Sum;
			this.colMes5.Visible = true;
			this.colMes5.VisibleIndex = 6;
			// 
			// colMes6
			// 
			this.colMes6.Caption = "Junio";
			this.colMes6.DisplayFormat.FormatString = "{0:n2}";
			this.colMes6.DisplayFormat.FormatType = DevExpress.Utils.FormatType.Numeric;
			this.colMes6.FieldName = "Mes6";
			this.colMes6.Name = "colMes6";
			this.colMes6.SummaryItem.DisplayFormat = "{0:n2}";
			this.colMes6.SummaryItem.SummaryType = DevExpress.Data.SummaryItemType.Sum;
			this.colMes6.Visible = true;
			this.colMes6.VisibleIndex = 7;
			// 
			// colMes7
			// 
			this.colMes7.Caption = "Julio";
			this.colMes7.DisplayFormat.FormatString = "{0:n2}";
			this.colMes7.DisplayFormat.FormatType = DevExpress.Utils.FormatType.Numeric;
			this.colMes7.FieldName = "Mes7";
			this.colMes7.Name = "colMes7";
			this.colMes7.SummaryItem.DisplayFormat = "{0:n2}";
			this.colMes7.SummaryItem.SummaryType = DevExpress.Data.SummaryItemType.Sum;
			this.colMes7.Visible = true;
			this.colMes7.VisibleIndex = 8;
			// 
			// colMes8
			// 
			this.colMes8.Caption = "Agosto";
			this.colMes8.DisplayFormat.FormatString = "{0:n2}";
			this.colMes8.DisplayFormat.FormatType = DevExpress.Utils.FormatType.Numeric;
			this.colMes8.FieldName = "Mes8";
			this.colMes8.Name = "colMes8";
			this.colMes8.SummaryItem.DisplayFormat = "{0:n2}";
			this.colMes8.SummaryItem.SummaryType = DevExpress.Data.SummaryItemType.Sum;
			this.colMes8.Visible = true;
			this.colMes8.VisibleIndex = 9;
			// 
			// colMes9
			// 
			this.colMes9.Caption = "Septiembre";
			this.colMes9.DisplayFormat.FormatString = "{0:n2}";
			this.colMes9.DisplayFormat.FormatType = DevExpress.Utils.FormatType.Numeric;
			this.colMes9.FieldName = "Mes9";
			this.colMes9.Name = "colMes9";
			this.colMes9.SummaryItem.DisplayFormat = "{0:n2}";
			this.colMes9.SummaryItem.SummaryType = DevExpress.Data.SummaryItemType.Sum;
			this.colMes9.Visible = true;
			this.colMes9.VisibleIndex = 10;
			// 
			// colMes10
			// 
			this.colMes10.Caption = "Octubre";
			this.colMes10.DisplayFormat.FormatString = "{0:n2}";
			this.colMes10.DisplayFormat.FormatType = DevExpress.Utils.FormatType.Numeric;
			this.colMes10.FieldName = "Mes10";
			this.colMes10.Name = "colMes10";
			this.colMes10.SummaryItem.DisplayFormat = "{0:n2}";
			this.colMes10.SummaryItem.SummaryType = DevExpress.Data.SummaryItemType.Sum;
			this.colMes10.Visible = true;
			this.colMes10.VisibleIndex = 11;
			// 
			// colMes11
			// 
			this.colMes11.Caption = "Noviembre";
			this.colMes11.DisplayFormat.FormatString = "{0:n2}";
			this.colMes11.DisplayFormat.FormatType = DevExpress.Utils.FormatType.Numeric;
			this.colMes11.FieldName = "Mes11";
			this.colMes11.Name = "colMes11";
			this.colMes11.SummaryItem.DisplayFormat = "{0:n2}";
			this.colMes11.SummaryItem.SummaryType = DevExpress.Data.SummaryItemType.Sum;
			this.colMes11.Visible = true;
			this.colMes11.VisibleIndex = 12;
			// 
			// colMes12
			// 
			this.colMes12.Caption = "Diciembre";
			this.colMes12.DisplayFormat.FormatString = "{0:n2}";
			this.colMes12.DisplayFormat.FormatType = DevExpress.Utils.FormatType.Numeric;
			this.colMes12.FieldName = "Mes12";
			this.colMes12.Name = "colMes12";
			this.colMes12.SummaryItem.DisplayFormat = "{0:n2}";
			this.colMes12.SummaryItem.SummaryType = DevExpress.Data.SummaryItemType.Sum;
			this.colMes12.Visible = true;
			this.colMes12.VisibleIndex = 13;
			// 
			// colTableField
			// 
			this.colTableField.Caption = "Total";
			this.colTableField.DisplayFormat.FormatString = "{0:n2}";
			this.colTableField.DisplayFormat.FormatType = DevExpress.Utils.FormatType.Numeric;
			this.colTableField.FieldName = "TableField";
			this.colTableField.Fixed = DevExpress.XtraGrid.Columns.FixedStyle.Right;
			this.colTableField.Name = "colTableField";
			this.colTableField.SummaryItem.DisplayFormat = "{0:n2}";
			this.colTableField.SummaryItem.SummaryType = DevExpress.Data.SummaryItemType.Sum;
			this.colTableField.Visible = true;
			this.colTableField.VisibleIndex = 14;
			// 
			// colTableField_1
			// 
			this.colTableField_1.Caption = "Promedio";
			this.colTableField_1.DisplayFormat.FormatString = "{0:n2}";
			this.colTableField_1.DisplayFormat.FormatType = DevExpress.Utils.FormatType.Numeric;
			this.colTableField_1.FieldName = "TableField_1";
			this.colTableField_1.Fixed = DevExpress.XtraGrid.Columns.FixedStyle.Right;
			this.colTableField_1.Name = "colTableField_1";
			this.colTableField_1.SummaryItem.DisplayFormat = "{0:n2}";
			this.colTableField_1.SummaryItem.SummaryType = DevExpress.Data.SummaryItemType.Average;
			this.colTableField_1.Visible = true;
			this.colTableField_1.VisibleIndex = 15;
			// 
			// btAceptar
			// 
			this.btAceptar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btAceptar.Location = new System.Drawing.Point(64, 304);
			this.btAceptar.Name = "btAceptar";
			this.btAceptar.TabIndex = 1;
			this.btAceptar.Text = "&Aceptar";
			this.btAceptar.Click += new System.EventHandler(this.btAceptar_Click);
			// 
			// btCancelar
			// 
			this.btCancelar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.btCancelar.Location = new System.Drawing.Point(512, 304);
			this.btCancelar.Name = "btCancelar";
			this.btCancelar.TabIndex = 2;
			this.btCancelar.Text = "&Cancelar";
			this.btCancelar.Click += new System.EventHandler(this.btCancelar_Click);
			// 
			// btImprimir
			// 
			this.btImprimir.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btImprimir.Location = new System.Drawing.Point(152, 304);
			this.btImprimir.Name = "btImprimir";
			this.btImprimir.TabIndex = 3;
			this.btImprimir.Text = "&Imprimir";
			this.btImprimir.Click += new System.EventHandler(this.btImprimir_Click);
			// 
			// cdsDetCompra
			// 
			this.cdsDetCompra.BindingContextCtrl = this;
			this.cdsDetCompra.DataLibrary = "LibFacturacion";
			this.cdsDetCompra.DataLibraryUrl = "";
			this.cdsDetCompra.DataSetDef = "dsCompra";
			this.cdsDetCompra.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsDetCompra.Name = "cdsDetCompra";
			this.cdsDetCompra.SchemaClassName = "LibFacturacion.DataClass";
			this.cdsDetCompra.SchemaDef = null;
			this.cdsDetCompra.BeforeFill += new C1.Data.FillEventHandler(this.cdsDetCompra_BeforeFill);
			// 
			// cdvDetCompra
			// 
			this.cdvDetCompra.BindingContextCtrl = this;
			this.cdvDetCompra.DataSet = this.cdsDetCompra;
			this.cdvDetCompra.TableName = "";
			this.cdvDetCompra.TableViewName = "DetCompra";
			// 
			// ProdPronostico
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(688, 349);
			this.Controls.Add(this.btImprimir);
			this.Controls.Add(this.btCancelar);
			this.Controls.Add(this.btAceptar);
			this.Controls.Add(this.gridControl1);
			this.Name = "ProdPronostico";
			this.Text = "Presupuesto mensual de clientes";
			this.Closing += new System.ComponentModel.CancelEventHandler(this.ProdPronostico_Closing);
			this.Load += new System.EventHandler(this.ProdPronostico_Load);
			((System.ComponentModel.ISupportInitialize)(this.gridControl1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvProdPronostico)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsProdPronostico)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.gridView1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.repositoryItemLookUpEdit1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvCliente)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsCliente)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.repositoryItemLookUpEdit2)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvArticulo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsArticulo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsDetCompra)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvDetCompra)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void btCancelar_Click(object sender, System.EventArgs e)
		{
			this.Close();
		}

		private void btAceptar_Click(object sender, System.EventArgs e)
		{
			this.cdsProdPronostico.Update();
			this.Close();
		}

		private void btImprimir_Click(object sender, System.EventArgs e)
		{
//			RepProduccion miReporte = new RepProduccion();
			Cursor = Cursors.WaitCursor;
			Reporte miReporte = new Reporte("ProdPronostico.rpt", "");
			miReporte.MdiParent = this.MdiParent;
			miReporte.Show();
			Cursor = Cursors.Default;
		}

		private float SaldoArt(DataTable dt, string idArt)
		{
			string stFiltro = "idArticulo = " + idArt.ToString();
			DataRow [] drDetC = dt.Select(stFiltro);
			float fSaldo = 0;
			foreach(DataRow dr in drDetC)
			{
				if ((int) dr["Signo"] == 1)
					fSaldo += (float) dr["Cantidad"];
				if ((int) dr["Signo"] == -1)
					fSaldo -= (float) dr["Cantidad"];
			}
			return fSaldo;
		}

		private void ProdPronostico_Load(object sender, System.EventArgs e)
		{
			string stAudita = "Exec AuditaCrear 53, 6, 'ProdPronostico'";
			Funcion.EjecutaSQL(cdsArticulo, stAudita, true);

			Cursor = Cursors.Default;
		}

		private void cdsProdPronostico_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			this.cdsProdPronostico.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);		
		}

		private void cdsCliente_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			this.cdsCliente.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);		
		}

		private void cdsArticulo_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			this.cdsArticulo.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);		
		}

		private void cdsDetCompra_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			this.cdsDetCompra.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);	
		}

		private void ProdPronostico_Closing(object sender, System.ComponentModel.CancelEventArgs e)
		{
			string stAudita = "Exec AuditaCrear 53, 7, 'ProdPromostico'";
			Funcion.EjecutaSQL(cdsArticulo, stAudita, true);
		}

	}
}
