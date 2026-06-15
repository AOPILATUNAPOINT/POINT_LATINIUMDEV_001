using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de SeteoG.
	/// </summary>
	public class SeteoG : DevExpress.XtraEditors.XtraForm
	{
		private Infragistics.Win.Misc.UltraButton btAceptar;
		private Infragistics.Win.Misc.UltraButton btCancelar;
		private C1.Data.C1DataSet cdsSeteoG;
		private DevExpress.XtraTreeList.TreeList treeList1;
		private C1.Data.C1DataView cdvSeteoG;
		private Infragistics.Win.Misc.UltraButton btArbol;
		private DevExpress.XtraTreeList.Columns.TreeListColumn colNombre;
		private DevExpress.XtraTreeList.Columns.TreeListColumn colActivado;
		private DevExpress.XtraTreeList.Columns.TreeListColumn colValor;
		private DevExpress.XtraTreeList.Columns.TreeListColumn colCodigo;
		private DevExpress.XtraGrid.GridControl gridControl1;
		private DevExpress.XtraGrid.Views.Grid.GridView gridView1;
		private DevExpress.XtraGrid.Views.Grid.GridView gridView2;
		private DevExpress.XtraGrid.Columns.GridColumn colidSeteoG;
		private DevExpress.XtraGrid.Columns.GridColumn colNombre1;
		private DevExpress.XtraGrid.Columns.GridColumn colActivado1;
		private DevExpress.XtraGrid.Columns.GridColumn colValor1;
		private DevExpress.XtraGrid.Columns.GridColumn colidPadre;
		private DevExpress.XtraGrid.Columns.GridColumn colCodigo1;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public SeteoG()
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
			this.cdsSeteoG = new C1.Data.C1DataSet();
			this.btAceptar = new Infragistics.Win.Misc.UltraButton();
			this.btCancelar = new Infragistics.Win.Misc.UltraButton();
			this.treeList1 = new DevExpress.XtraTreeList.TreeList();
			this.colNombre = new DevExpress.XtraTreeList.Columns.TreeListColumn();
			this.colActivado = new DevExpress.XtraTreeList.Columns.TreeListColumn();
			this.colValor = new DevExpress.XtraTreeList.Columns.TreeListColumn();
			this.colCodigo = new DevExpress.XtraTreeList.Columns.TreeListColumn();
			this.cdvSeteoG = new C1.Data.C1DataView();
			this.btArbol = new Infragistics.Win.Misc.UltraButton();
			this.gridControl1 = new DevExpress.XtraGrid.GridControl();
			this.gridView1 = new DevExpress.XtraGrid.Views.Grid.GridView();
			this.gridView2 = new DevExpress.XtraGrid.Views.Grid.GridView();
			this.colidSeteoG = new DevExpress.XtraGrid.Columns.GridColumn();
			this.colNombre1 = new DevExpress.XtraGrid.Columns.GridColumn();
			this.colActivado1 = new DevExpress.XtraGrid.Columns.GridColumn();
			this.colValor1 = new DevExpress.XtraGrid.Columns.GridColumn();
			this.colidPadre = new DevExpress.XtraGrid.Columns.GridColumn();
			this.colCodigo1 = new DevExpress.XtraGrid.Columns.GridColumn();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoG)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.treeList1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvSeteoG)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.gridControl1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.gridView1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.gridView2)).BeginInit();
			this.SuspendLayout();
			// 
			// cdsSeteoG
			// 
			this.cdsSeteoG.BindingContextCtrl = this;
			this.cdsSeteoG.DataLibrary = "LibContabilidad";
			this.cdsSeteoG.DataLibraryUrl = "";
			this.cdsSeteoG.DataSetDef = "dsSeteoG";
			this.cdsSeteoG.EnforceConstraints = false;
			this.cdsSeteoG.Locale = new System.Globalization.CultureInfo("");
			this.cdsSeteoG.Name = "cdsSeteoG";
			this.cdsSeteoG.SchemaClassName = "LibContabilidad.DataClass";
			this.cdsSeteoG.SchemaDef = null;
			this.cdsSeteoG.BeforeFill += new C1.Data.FillEventHandler(this.cdsSeteoG_BeforeFill);
			// 
			// btAceptar
			// 
			this.btAceptar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btAceptar.Location = new System.Drawing.Point(32, 362);
			this.btAceptar.Name = "btAceptar";
			this.btAceptar.Size = new System.Drawing.Size(80, 26);
			this.btAceptar.TabIndex = 1;
			this.btAceptar.Text = "&Aceptar";
			this.btAceptar.Click += new System.EventHandler(this.btAceptar_Click);
			// 
			// btCancelar
			// 
			this.btCancelar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.btCancelar.DialogResult = System.Windows.Forms.DialogResult.Cancel;
			this.btCancelar.Location = new System.Drawing.Point(912, 362);
			this.btCancelar.Name = "btCancelar";
			this.btCancelar.Size = new System.Drawing.Size(80, 26);
			this.btCancelar.TabIndex = 2;
			this.btCancelar.Text = "&Cancelar";
			this.btCancelar.Click += new System.EventHandler(this.btCancelar_Click);
			// 
			// treeList1
			// 
			this.treeList1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.treeList1.Appearance.Empty.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(255)), ((System.Byte)(255)), ((System.Byte)(255)));
			this.treeList1.Appearance.Empty.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(255)), ((System.Byte)(255)), ((System.Byte)(255)));
			this.treeList1.Appearance.Empty.Options.UseBackColor = true;
			this.treeList1.Appearance.Empty.Options.UseForeColor = true;
			this.treeList1.Appearance.EvenRow.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(248)), ((System.Byte)(248)), ((System.Byte)(255)));
			this.treeList1.Appearance.EvenRow.Font = new System.Drawing.Font("Tahoma", 8F);
			this.treeList1.Appearance.EvenRow.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(0)), ((System.Byte)(0)), ((System.Byte)(0)));
			this.treeList1.Appearance.EvenRow.Options.UseBackColor = true;
			this.treeList1.Appearance.EvenRow.Options.UseFont = true;
			this.treeList1.Appearance.EvenRow.Options.UseForeColor = true;
			this.treeList1.Appearance.FocusedCell.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(255)), ((System.Byte)(228)), ((System.Byte)(225)));
			this.treeList1.Appearance.FocusedCell.Font = new System.Drawing.Font("Tahoma", 8F);
			this.treeList1.Appearance.FocusedCell.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(0)), ((System.Byte)(0)), ((System.Byte)(0)));
			this.treeList1.Appearance.FocusedCell.Options.UseBackColor = true;
			this.treeList1.Appearance.FocusedCell.Options.UseFont = true;
			this.treeList1.Appearance.FocusedCell.Options.UseForeColor = true;
			this.treeList1.Appearance.FocusedRow.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(255)), ((System.Byte)(0)), ((System.Byte)(0)));
			this.treeList1.Appearance.FocusedRow.Font = new System.Drawing.Font("Tahoma", 8F);
			this.treeList1.Appearance.FocusedRow.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(255)), ((System.Byte)(255)), ((System.Byte)(255)));
			this.treeList1.Appearance.FocusedRow.Options.UseBackColor = true;
			this.treeList1.Appearance.FocusedRow.Options.UseFont = true;
			this.treeList1.Appearance.FocusedRow.Options.UseForeColor = true;
			this.treeList1.Appearance.FooterPanel.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(0)), ((System.Byte)(206)), ((System.Byte)(209)));
			this.treeList1.Appearance.FooterPanel.Font = new System.Drawing.Font("Tahoma", 8F);
			this.treeList1.Appearance.FooterPanel.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(0)), ((System.Byte)(0)), ((System.Byte)(0)));
			this.treeList1.Appearance.FooterPanel.Options.UseBackColor = true;
			this.treeList1.Appearance.FooterPanel.Options.UseFont = true;
			this.treeList1.Appearance.FooterPanel.Options.UseForeColor = true;
			this.treeList1.Appearance.GroupButton.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(255)), ((System.Byte)(255)), ((System.Byte)(255)));
			this.treeList1.Appearance.GroupButton.Font = new System.Drawing.Font("Tahoma", 8F);
			this.treeList1.Appearance.GroupButton.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(0)), ((System.Byte)(0)), ((System.Byte)(0)));
			this.treeList1.Appearance.GroupButton.Options.UseBackColor = true;
			this.treeList1.Appearance.GroupButton.Options.UseFont = true;
			this.treeList1.Appearance.GroupButton.Options.UseForeColor = true;
			this.treeList1.Appearance.GroupFooter.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(10)), ((System.Byte)(216)), ((System.Byte)(219)));
			this.treeList1.Appearance.GroupFooter.Font = new System.Drawing.Font("Tahoma", 8F);
			this.treeList1.Appearance.GroupFooter.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(0)), ((System.Byte)(0)), ((System.Byte)(0)));
			this.treeList1.Appearance.GroupFooter.Options.UseBackColor = true;
			this.treeList1.Appearance.GroupFooter.Options.UseFont = true;
			this.treeList1.Appearance.GroupFooter.Options.UseForeColor = true;
			this.treeList1.Appearance.HeaderPanel.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(0)), ((System.Byte)(206)), ((System.Byte)(209)));
			this.treeList1.Appearance.HeaderPanel.Font = new System.Drawing.Font("Arial", 8F, System.Drawing.FontStyle.Bold);
			this.treeList1.Appearance.HeaderPanel.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(0)), ((System.Byte)(0)), ((System.Byte)(0)));
			this.treeList1.Appearance.HeaderPanel.Options.UseBackColor = true;
			this.treeList1.Appearance.HeaderPanel.Options.UseFont = true;
			this.treeList1.Appearance.HeaderPanel.Options.UseForeColor = true;
			this.treeList1.Appearance.HideSelectionRow.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(128)), ((System.Byte)(128)), ((System.Byte)(128)));
			this.treeList1.Appearance.HideSelectionRow.Font = new System.Drawing.Font("Tahoma", 8F);
			this.treeList1.Appearance.HideSelectionRow.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(212)), ((System.Byte)(208)), ((System.Byte)(200)));
			this.treeList1.Appearance.HideSelectionRow.Options.UseBackColor = true;
			this.treeList1.Appearance.HideSelectionRow.Options.UseFont = true;
			this.treeList1.Appearance.HideSelectionRow.Options.UseForeColor = true;
			this.treeList1.Appearance.HorzLine.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(128)), ((System.Byte)(128)), ((System.Byte)(128)));
			this.treeList1.Appearance.HorzLine.Font = new System.Drawing.Font("Tahoma", 8F);
			this.treeList1.Appearance.HorzLine.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(128)), ((System.Byte)(128)), ((System.Byte)(128)));
			this.treeList1.Appearance.HorzLine.Options.UseBackColor = true;
			this.treeList1.Appearance.HorzLine.Options.UseFont = true;
			this.treeList1.Appearance.HorzLine.Options.UseForeColor = true;
			this.treeList1.Appearance.OddRow.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(230)), ((System.Byte)(230)), ((System.Byte)(250)));
			this.treeList1.Appearance.OddRow.Font = new System.Drawing.Font("Arial Black", 7F);
			this.treeList1.Appearance.OddRow.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(0)), ((System.Byte)(0)), ((System.Byte)(0)));
			this.treeList1.Appearance.OddRow.Options.UseBackColor = true;
			this.treeList1.Appearance.OddRow.Options.UseFont = true;
			this.treeList1.Appearance.OddRow.Options.UseForeColor = true;
			this.treeList1.Appearance.Preview.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(255)), ((System.Byte)(255)), ((System.Byte)(255)));
			this.treeList1.Appearance.Preview.Font = new System.Drawing.Font("Tahoma", 8F);
			this.treeList1.Appearance.Preview.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(0)), ((System.Byte)(0)), ((System.Byte)(255)));
			this.treeList1.Appearance.Preview.Options.UseBackColor = true;
			this.treeList1.Appearance.Preview.Options.UseFont = true;
			this.treeList1.Appearance.Preview.Options.UseForeColor = true;
			this.treeList1.Appearance.Row.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(255)), ((System.Byte)(255)), ((System.Byte)(255)));
			this.treeList1.Appearance.Row.Font = new System.Drawing.Font("Tahoma", 8F);
			this.treeList1.Appearance.Row.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(0)), ((System.Byte)(0)), ((System.Byte)(0)));
			this.treeList1.Appearance.Row.Options.UseBackColor = true;
			this.treeList1.Appearance.Row.Options.UseFont = true;
			this.treeList1.Appearance.Row.Options.UseForeColor = true;
			this.treeList1.Appearance.SelectedRow.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(255)), ((System.Byte)(10)), ((System.Byte)(10)));
			this.treeList1.Appearance.SelectedRow.Font = new System.Drawing.Font("Tahoma", 8F);
			this.treeList1.Appearance.SelectedRow.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(255)), ((System.Byte)(255)), ((System.Byte)(255)));
			this.treeList1.Appearance.SelectedRow.Options.UseBackColor = true;
			this.treeList1.Appearance.SelectedRow.Options.UseFont = true;
			this.treeList1.Appearance.SelectedRow.Options.UseForeColor = true;
			this.treeList1.Appearance.TreeLine.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(255)), ((System.Byte)(255)), ((System.Byte)(255)));
			this.treeList1.Appearance.TreeLine.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(128)), ((System.Byte)(128)), ((System.Byte)(128)));
			this.treeList1.Appearance.TreeLine.Options.UseBackColor = true;
			this.treeList1.Appearance.TreeLine.Options.UseForeColor = true;
			this.treeList1.Appearance.VertLine.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(128)), ((System.Byte)(128)), ((System.Byte)(128)));
			this.treeList1.Appearance.VertLine.Font = new System.Drawing.Font("Tahoma", 8F);
			this.treeList1.Appearance.VertLine.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(128)), ((System.Byte)(128)), ((System.Byte)(128)));
			this.treeList1.Appearance.VertLine.Options.UseBackColor = true;
			this.treeList1.Appearance.VertLine.Options.UseFont = true;
			this.treeList1.Appearance.VertLine.Options.UseForeColor = true;
			this.treeList1.Appearance.VertLine.Options.UseTextOptions = true;
			this.treeList1.Appearance.VertLine.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Near;
			this.treeList1.Columns.AddRange(new DevExpress.XtraTreeList.Columns.TreeListColumn[] {
																																														 this.colNombre,
																																														 this.colActivado,
																																														 this.colValor,
																																														 this.colCodigo});
			this.treeList1.DataSource = this.cdvSeteoG;
			this.treeList1.KeyFieldName = "idSeteoG";
			this.treeList1.Location = new System.Drawing.Point(16, 9);
			this.treeList1.Name = "treeList1";
			this.treeList1.ParentFieldName = "idPadre";
			this.treeList1.RootValue = "0";
			this.treeList1.Size = new System.Drawing.Size(1000, 163);
			this.treeList1.Styles.AddReplace("PressedColumn", new DevExpress.Utils.ViewStyle("PressedColumn", "TreeList", new System.Drawing.Font("Arial", 8F, System.Drawing.FontStyle.Bold), "HeaderPanel", ((DevExpress.Utils.StyleOptions)(((DevExpress.Utils.StyleOptions.StyleEnabled | DevExpress.Utils.StyleOptions.UseBackColor) 
				| DevExpress.Utils.StyleOptions.UseForeColor))), true, false, false, DevExpress.Utils.HorzAlignment.Near, DevExpress.Utils.VertAlignment.Center, null, System.Drawing.Color.FromArgb(((System.Byte)(128)), ((System.Byte)(128)), ((System.Byte)(128))), System.Drawing.Color.FromArgb(((System.Byte)(255)), ((System.Byte)(255)), ((System.Byte)(255)))));
			this.treeList1.TabIndex = 3;
			// 
			// colNombre
			// 
			this.colNombre.Caption = "Nombre";
			this.colNombre.FieldName = "Nombre";
			this.colNombre.Name = "colNombre";
			this.colNombre.VisibleIndex = 0;
			this.colNombre.Width = 258;
			// 
			// colActivado
			// 
			this.colActivado.Caption = "Activado";
			this.colActivado.FieldName = "Activado";
			this.colActivado.Name = "colActivado";
			this.colActivado.VisibleIndex = 1;
			this.colActivado.Width = 55;
			// 
			// colValor
			// 
			this.colValor.Caption = "Valor";
			this.colValor.FieldName = "Valor";
			this.colValor.Format.FormatType = DevExpress.Utils.FormatType.Numeric;
			this.colValor.Name = "colValor";
			this.colValor.VisibleIndex = 2;
			this.colValor.Width = 76;
			// 
			// colCodigo
			// 
			this.colCodigo.Caption = "Código";
			this.colCodigo.FieldName = "Codigo";
			this.colCodigo.Name = "colCodigo";
			this.colCodigo.VisibleIndex = 3;
			this.colCodigo.Width = 74;
			// 
			// cdvSeteoG
			// 
			this.cdvSeteoG.BindingContextCtrl = this;
			this.cdvSeteoG.DataSet = this.cdsSeteoG;
			this.cdvSeteoG.Sort = "Codigo";
			this.cdvSeteoG.TableName = "";
			this.cdvSeteoG.TableViewName = "SeteoG";
			// 
			// btArbol
			// 
			this.btArbol.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btArbol.Location = new System.Drawing.Point(144, 362);
			this.btArbol.Name = "btArbol";
			this.btArbol.Size = new System.Drawing.Size(80, 26);
			this.btArbol.TabIndex = 4;
			this.btArbol.Text = "A&rbol";
			this.btArbol.Click += new System.EventHandler(this.btArbol_Click);
			// 
			// gridControl1
			// 
			this.gridControl1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.gridControl1.DataMember = "SeteoG";
			this.gridControl1.DataSource = this.cdsSeteoG;
			// 
			// gridControl1.EmbeddedNavigator
			// 
			this.gridControl1.EmbeddedNavigator.Name = "";
			this.gridControl1.Location = new System.Drawing.Point(16, 184);
			this.gridControl1.MainView = this.gridView2;
			this.gridControl1.Name = "gridControl1";
			this.gridControl1.Size = new System.Drawing.Size(1000, 168);
			this.gridControl1.TabIndex = 5;
			this.gridControl1.ViewCollection.AddRange(new DevExpress.XtraGrid.Views.Base.BaseView[] {
																																																this.gridView2,
																																																this.gridView1});
			// 
			// gridView1
			// 
			this.gridView1.GridControl = this.gridControl1;
			this.gridView1.Name = "gridView1";
			// 
			// gridView2
			// 
			this.gridView2.Columns.AddRange(new DevExpress.XtraGrid.Columns.GridColumn[] {
																																										 this.colidSeteoG,
																																										 this.colNombre1,
																																										 this.colActivado1,
																																										 this.colValor1,
																																										 this.colidPadre,
																																										 this.colCodigo1});
			this.gridView2.GridControl = this.gridControl1;
			this.gridView2.Name = "gridView2";
			this.gridView2.OptionsView.NewItemRowPosition = DevExpress.XtraGrid.Views.Grid.NewItemRowPosition.Bottom;
			this.gridView2.OptionsView.ShowAutoFilterRow = true;
			this.gridView2.OptionsView.ShowGroupPanel = false;
			// 
			// colidSeteoG
			// 
			this.colidSeteoG.Caption = "idSeteoG";
			this.colidSeteoG.FieldName = "idSeteoG";
			this.colidSeteoG.Name = "colidSeteoG";
			// 
			// colNombre1
			// 
			this.colNombre1.Caption = "Nombre";
			this.colNombre1.FieldName = "Nombre";
			this.colNombre1.Name = "colNombre1";
			this.colNombre1.Visible = true;
			this.colNombre1.VisibleIndex = 0;
			this.colNombre1.Width = 554;
			// 
			// colActivado1
			// 
			this.colActivado1.Caption = "Activado";
			this.colActivado1.FieldName = "Activado";
			this.colActivado1.Name = "colActivado1";
			this.colActivado1.Visible = true;
			this.colActivado1.VisibleIndex = 1;
			this.colActivado1.Width = 107;
			// 
			// colValor1
			// 
			this.colValor1.Caption = "Valor";
			this.colValor1.FieldName = "Valor";
			this.colValor1.Name = "colValor1";
			this.colValor1.Visible = true;
			this.colValor1.VisibleIndex = 2;
			this.colValor1.Width = 144;
			// 
			// colidPadre
			// 
			this.colidPadre.Caption = "idPadre";
			this.colidPadre.FieldName = "idPadre";
			this.colidPadre.Name = "colidPadre";
			// 
			// colCodigo1
			// 
			this.colCodigo1.Caption = "Codigo";
			this.colCodigo1.FieldName = "Codigo";
			this.colCodigo1.Name = "colCodigo1";
			this.colCodigo1.Visible = true;
			this.colCodigo1.VisibleIndex = 3;
			this.colCodigo1.Width = 149;
			// 
			// SeteoG
			// 
			this.AcceptButton = this.btAceptar;
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 14);
			this.CancelButton = this.btCancelar;
			this.ClientSize = new System.Drawing.Size(1028, 402);
			this.Controls.Add(this.gridControl1);
			this.Controls.Add(this.btArbol);
			this.Controls.Add(this.treeList1);
			this.Controls.Add(this.btCancelar);
			this.Controls.Add(this.btAceptar);
			this.Name = "SeteoG";
			this.Text = "Seteo Global";
			this.Load += new System.EventHandler(this.SeteoG_Load);
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoG)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.treeList1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvSeteoG)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.gridControl1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.gridView1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.gridView2)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void cdsSeteoG_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			this.cdsSeteoG.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirConta);
		}

		private void btAceptar_Click(object sender, System.EventArgs e)
		{
			try
			{
				cdsSeteoG.Update();
			}
			catch(System.Exception ex)
			{
				MessageBox.Show("No se puede grabar: " + ex.Message, "Información");
				return;
			}
			Close();
		}

		private void btCancelar_Click(object sender, System.EventArgs e)
		{
			Close();
		}

		private void btArbol_Click(object sender, System.EventArgs e)
		{
			Cursor = Cursors.WaitCursor;
			try
			{
				cdsSeteoG.Update();
			}
			catch(System.Exception ex)
			{
				Cursor = Cursors.Default;
				MessageBox.Show("No se puede grabar: " + ex.Message, "Informaciòn");
				return;
			}
			string stExec = "Exec SeteoGArbol";
      Funcion.EjecutaSQL(cdsSeteoG, stExec, true);
			cdsSeteoG.Clear();
			cdsSeteoG.Fill();
			Cursor = Cursors.Default;
			MessageBox.Show("Arbol Generado", "Información");
		}

		private void SeteoG_Load(object sender, System.EventArgs e)
		{
		
		}
	}
}
