using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de ArticuloSeleccion.
	/// </summary>
	public class ArticuloSeleccion : System.Windows.Forms.Form
	{
		private C1.Data.C1DataSet cdsArticulo;
		private C1.Data.C1DataView cdvArticulo;
		private C1.Data.C1DataSet cdsSubGrupo;
		private C1.Data.C1DataView cdvSubGrupo;
		private DevExpress.XtraGrid.GridControl gridControl1;
		private DevExpress.XtraGrid.Views.Grid.GridView gridView1;
		private DevExpress.XtraGrid.Columns.GridColumn colSubGrupo;
		private DevExpress.XtraGrid.GridControl gridControl2;
		private DevExpress.XtraGrid.Views.Grid.GridView gridView2;
		private DevExpress.XtraGrid.Columns.GridColumn colNombre;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public ArticuloSeleccion()
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
			this.cdsArticulo = new C1.Data.C1DataSet();
			this.cdvArticulo = new C1.Data.C1DataView();
			this.cdsSubGrupo = new C1.Data.C1DataSet();
			this.cdvSubGrupo = new C1.Data.C1DataView();
			this.gridControl1 = new DevExpress.XtraGrid.GridControl();
			this.gridView1 = new DevExpress.XtraGrid.Views.Grid.GridView();
			this.colSubGrupo = new DevExpress.XtraGrid.Columns.GridColumn();
			this.gridControl2 = new DevExpress.XtraGrid.GridControl();
			this.gridView2 = new DevExpress.XtraGrid.Views.Grid.GridView();
			this.colNombre = new DevExpress.XtraGrid.Columns.GridColumn();
			((System.ComponentModel.ISupportInitialize)(this.cdsArticulo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvArticulo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSubGrupo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvSubGrupo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.gridControl1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.gridView1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.gridControl2)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.gridView2)).BeginInit();
			this.SuspendLayout();
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
			// cdvArticulo
			// 
			this.cdvArticulo.BindingContextCtrl = this;
			this.cdvArticulo.DataSet = this.cdsArticulo;
			this.cdvArticulo.TableName = "";
			this.cdvArticulo.TableViewName = "Articulo";
			// 
			// cdsSubGrupo
			// 
			this.cdsSubGrupo.BindingContextCtrl = this;
			this.cdsSubGrupo.DataLibrary = "LibFacturacion";
			this.cdsSubGrupo.DataLibraryUrl = "";
			this.cdsSubGrupo.DataSetDef = "dsArticuloSubGrupo";
			this.cdsSubGrupo.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsSubGrupo.Name = "cdsSubGrupo";
			this.cdsSubGrupo.SchemaClassName = "LibFacturacion.DataClass";
			this.cdsSubGrupo.SchemaDef = null;
			// 
			// cdvSubGrupo
			// 
			this.cdvSubGrupo.BindingContextCtrl = this;
			this.cdvSubGrupo.DataSet = this.cdsSubGrupo;
			this.cdvSubGrupo.TableName = "";
			this.cdvSubGrupo.TableViewName = "ArticuloSubGrupo";
			this.cdvSubGrupo.PositionChanged += new C1.Data.PositionChangeEventHandler(this.cdvSubGrupo_PositionChanged);
			// 
			// gridControl1
			// 
			this.gridControl1.DataSource = this.cdvSubGrupo;
			this.gridControl1.Location = new System.Drawing.Point(16, 16);
			this.gridControl1.MainView = this.gridView1;
			this.gridControl1.Name = "gridControl1";
			this.gridControl1.Size = new System.Drawing.Size(328, 200);
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
			// gridView1
			// 
			this.gridView1.Columns.AddRange(new DevExpress.XtraGrid.Columns.GridColumn[] {
																																										 this.colSubGrupo});
			this.gridView1.Name = "gridView1";
			this.gridView1.OptionsView.ShowGroupPanel = false;
			// 
			// colSubGrupo
			// 
			this.colSubGrupo.Caption = "SubGrupo";
			this.colSubGrupo.FieldName = "SubGrupo";
			this.colSubGrupo.Name = "colSubGrupo";
			this.colSubGrupo.VisibleIndex = 0;
			// 
			// gridControl2
			// 
			this.gridControl2.DataSource = this.cdvArticulo;
			this.gridControl2.Location = new System.Drawing.Point(360, 16);
			this.gridControl2.MainView = this.gridView2;
			this.gridControl2.Name = "gridControl2";
			this.gridControl2.Size = new System.Drawing.Size(328, 200);
			this.gridControl2.Styles.AddReplace("ColumnFilterButtonActive", new DevExpress.Utils.ViewStyleEx("ColumnFilterButtonActive", "Grid", System.Drawing.SystemColors.Control, System.Drawing.Color.Blue, System.Drawing.SystemColors.ControlLightLight, System.Drawing.Drawing2D.LinearGradientMode.Horizontal));
			this.gridControl2.TabIndex = 3;
			this.gridControl2.Text = "gridControl2";
			this.gridControl2.Click += new System.EventHandler(this.gridControl2_Click);
			// 
			// gridView2
			// 
			this.gridView2.Columns.AddRange(new DevExpress.XtraGrid.Columns.GridColumn[] {
																																										 this.colNombre});
			this.gridView2.Name = "gridView2";
			// 
			// colNombre
			// 
			this.colNombre.Caption = "Nombre";
			this.colNombre.FieldName = "Nombre";
			this.colNombre.Name = "colNombre";
			this.colNombre.VisibleIndex = 0;
			// 
			// ArticuloSeleccion
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(704, 357);
			this.Controls.AddRange(new System.Windows.Forms.Control[] {
																																	this.gridControl2,
																																	this.gridControl1});
			this.Name = "ArticuloSeleccion";
			this.Text = "ArticuloSeleccion";
			((System.ComponentModel.ISupportInitialize)(this.cdsArticulo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvArticulo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSubGrupo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvSubGrupo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.gridControl1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.gridView1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.gridControl2)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.gridView2)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void gridControl2_Click(object sender, System.EventArgs e)
		{
		
		}

		private void cdvSubGrupo_PositionChanged(object sender, C1.Data.PositionChangeEventArgs e)
		{
			int iIdSubG = (int) e.Row["idSubGrupo"];
			this.cdvArticulo.RowFilter = "idSubGrupo = " + iIdSubG.ToString();
		}
	}
}
