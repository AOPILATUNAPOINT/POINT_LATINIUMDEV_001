using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using DevExpress.XtraEditors;

namespace Latinium
{
	/// <summary>
	/// Summary description for CompraReporte.
	/// </summary>
	public class CompraReporte : DevExpress.XtraEditors.XtraForm
	{
		private DevExpress.XtraGrid.GridControl gridControl1;
		private DevExpress.XtraGrid.Views.Grid.GridView gridView1;
		private DevExpress.XtraEditors.SimpleButton btAceptar;
		private C1.Data.C1DataSet cdsCompraTabla;
		private DevExpress.XtraGrid.Columns.GridColumn colidCompraReporte;
		private DevExpress.XtraGrid.Columns.GridColumn colReporte;
		private DevExpress.XtraGrid.Columns.GridColumn colidTipoFactura;
		private DevExpress.XtraGrid.Columns.GridColumn colRuta;
		private DevExpress.XtraEditors.SimpleButton btCancelar;
		private DevExpress.XtraEditors.Repository.RepositoryItemLookUpEdit repositoryItemLookUpEdit1;
		private C1.Data.C1DataView cdvCompraNumero;
		private DevExpress.XtraGrid.GridControl gridControl2;
		private DevExpress.XtraEditors.Repository.RepositoryItemLookUpEdit repositoryItemLookUpEdit2;
		private DevExpress.XtraGrid.Columns.GridColumn colidCompraRepDet;
		private DevExpress.XtraGrid.Columns.GridColumn colCampo;
		private DevExpress.XtraGrid.Columns.GridColumn colTabla;
		private DevExpress.XtraGrid.Columns.GridColumn colUbicacion;
		private DevExpress.XtraGrid.Columns.GridColumn colidCompraReporte1;
		private DevExpress.XtraGrid.Columns.GridColumn colExplicacion;
		private DevExpress.XtraGrid.Columns.GridColumn colObligatorio;
		private System.Windows.Forms.TextBox txtIdCompraRep;
		private DevExpress.XtraGrid.Views.Grid.GridView grvDetalle;
		private DevExpress.XtraEditors.Repository.RepositoryItemMemoEdit repositoryItemMemoEdit1;
		private DevExpress.XtraGrid.Columns.GridColumn colDatos;
		/// <summary>
		/// Required designer variable.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public CompraReporte()
		{
			//
			// Required for Windows Form Designer support
			//
			InitializeComponent();

			//
			// TODO: Add any constructor code after InitializeComponent call
			//
		}

		/// <summary>
		/// Clean up any resources being used.
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
		/// Required method for Designer support - do not modify
		/// the contents of this method with the code editor.
		/// </summary>
		private void InitializeComponent()
		{
			this.gridControl1 = new DevExpress.XtraGrid.GridControl();
			this.cdsCompraTabla = new C1.Data.C1DataSet();
			this.gridView1 = new DevExpress.XtraGrid.Views.Grid.GridView();
			this.colidCompraReporte = new DevExpress.XtraGrid.Columns.GridColumn();
			this.colReporte = new DevExpress.XtraGrid.Columns.GridColumn();
			this.colidTipoFactura = new DevExpress.XtraGrid.Columns.GridColumn();
			this.repositoryItemLookUpEdit1 = new DevExpress.XtraEditors.Repository.RepositoryItemLookUpEdit();
			this.cdvCompraNumero = new C1.Data.C1DataView();
			this.colRuta = new DevExpress.XtraGrid.Columns.GridColumn();
			this.btAceptar = new DevExpress.XtraEditors.SimpleButton();
			this.btCancelar = new DevExpress.XtraEditors.SimpleButton();
			this.gridControl2 = new DevExpress.XtraGrid.GridControl();
			this.grvDetalle = new DevExpress.XtraGrid.Views.Grid.GridView();
			this.colidCompraRepDet = new DevExpress.XtraGrid.Columns.GridColumn();
			this.colTabla = new DevExpress.XtraGrid.Columns.GridColumn();
			this.colCampo = new DevExpress.XtraGrid.Columns.GridColumn();
			this.colUbicacion = new DevExpress.XtraGrid.Columns.GridColumn();
			this.repositoryItemMemoEdit1 = new DevExpress.XtraEditors.Repository.RepositoryItemMemoEdit();
			this.colidCompraReporte1 = new DevExpress.XtraGrid.Columns.GridColumn();
			this.colExplicacion = new DevExpress.XtraGrid.Columns.GridColumn();
			this.colObligatorio = new DevExpress.XtraGrid.Columns.GridColumn();
			this.colDatos = new DevExpress.XtraGrid.Columns.GridColumn();
			this.repositoryItemLookUpEdit2 = new DevExpress.XtraEditors.Repository.RepositoryItemLookUpEdit();
			this.txtIdCompraRep = new System.Windows.Forms.TextBox();
			((System.ComponentModel.ISupportInitialize)(this.gridControl1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsCompraTabla)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.gridView1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.repositoryItemLookUpEdit1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvCompraNumero)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.gridControl2)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.grvDetalle)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.repositoryItemMemoEdit1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.repositoryItemLookUpEdit2)).BeginInit();
			this.SuspendLayout();
			// 
			// gridControl1
			// 
			this.gridControl1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.gridControl1.DataMember = "CompraReporte";
			this.gridControl1.DataSource = this.cdsCompraTabla;
			// 
			// gridControl1.EmbeddedNavigator
			// 
			this.gridControl1.EmbeddedNavigator.Name = "";
			this.gridControl1.Location = new System.Drawing.Point(16, 16);
			this.gridControl1.MainView = this.gridView1;
			this.gridControl1.Name = "gridControl1";
			this.gridControl1.RepositoryItems.AddRange(new DevExpress.XtraEditors.Repository.RepositoryItem[] {
																																																					this.repositoryItemLookUpEdit1});
			this.gridControl1.Size = new System.Drawing.Size(760, 224);
			this.gridControl1.TabIndex = 0;
			this.gridControl1.ViewCollection.AddRange(new DevExpress.XtraGrid.Views.Base.BaseView[] {
																																																this.gridView1});
			// 
			// cdsCompraTabla
			// 
			this.cdsCompraTabla.BindingContextCtrl = this;
			this.cdsCompraTabla.DataLibrary = "LibFacturacion";
			this.cdsCompraTabla.DataLibraryUrl = "";
			this.cdsCompraTabla.DataSetDef = "dsCompraTabla";
			this.cdsCompraTabla.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsCompraTabla.Name = "cdsCompraTabla";
			this.cdsCompraTabla.SchemaClassName = "LibFacturacion.DataClass";
			this.cdsCompraTabla.SchemaDef = null;
			this.cdsCompraTabla.BeforeFill += new C1.Data.FillEventHandler(this.cdsCompraTabla_BeforeFill);
			// 
			// gridView1
			// 
			this.gridView1.Columns.AddRange(new DevExpress.XtraGrid.Columns.GridColumn[] {
																																										 this.colidCompraReporte,
																																										 this.colReporte,
																																										 this.colidTipoFactura,
																																										 this.colRuta});
			this.gridView1.GridControl = this.gridControl1;
			this.gridView1.Name = "gridView1";
			this.gridView1.OptionsView.NewItemRowPosition = DevExpress.XtraGrid.Views.Grid.NewItemRowPosition.Bottom;
			// 
			// colidCompraReporte
			// 
			this.colidCompraReporte.Caption = "idCompraReporte";
			this.colidCompraReporte.FieldName = "idCompraReporte";
			this.colidCompraReporte.Name = "colidCompraReporte";
			// 
			// colReporte
			// 
			this.colReporte.Caption = "Reporte";
			this.colReporte.FieldName = "Reporte";
			this.colReporte.Name = "colReporte";
			this.colReporte.Visible = true;
			this.colReporte.VisibleIndex = 0;
			// 
			// colidTipoFactura
			// 
			this.colidTipoFactura.Caption = "Tipo de Factura";
			this.colidTipoFactura.ColumnEdit = this.repositoryItemLookUpEdit1;
			this.colidTipoFactura.FieldName = "idTipoFactura";
			this.colidTipoFactura.Name = "colidTipoFactura";
			this.colidTipoFactura.Visible = true;
			this.colidTipoFactura.VisibleIndex = 1;
			// 
			// repositoryItemLookUpEdit1
			// 
			this.repositoryItemLookUpEdit1.AutoHeight = false;
			this.repositoryItemLookUpEdit1.Buttons.AddRange(new DevExpress.XtraEditors.Controls.EditorButton[] {
																																																					 new DevExpress.XtraEditors.Controls.EditorButton(DevExpress.XtraEditors.Controls.ButtonPredefines.Combo)});
			this.repositoryItemLookUpEdit1.Columns.AddRange(new DevExpress.XtraEditors.Controls.LookUpColumnInfo[] {
																																																							 new DevExpress.XtraEditors.Controls.LookUpColumnInfo("Nombre", "Nombre", 43, DevExpress.Utils.FormatType.None, "", true, DevExpress.Utils.HorzAlignment.Near, DevExpress.Data.ColumnSortOrder.Ascending)});
			this.repositoryItemLookUpEdit1.DataSource = this.cdvCompraNumero;
			this.repositoryItemLookUpEdit1.DisplayMember = "Nombre";
			this.repositoryItemLookUpEdit1.Name = "repositoryItemLookUpEdit1";
			this.repositoryItemLookUpEdit1.NullText = "";
			this.repositoryItemLookUpEdit1.ValueMember = "idTipoFactura";
			// 
			// cdvCompraNumero
			// 
			this.cdvCompraNumero.BindingContextCtrl = this;
			this.cdvCompraNumero.DataSet = this.cdsCompraTabla;
			this.cdvCompraNumero.TableName = "";
			this.cdvCompraNumero.TableViewName = "CompraNumero";
			// 
			// colRuta
			// 
			this.colRuta.Caption = "Ruta";
			this.colRuta.FieldName = "Ruta";
			this.colRuta.Name = "colRuta";
			this.colRuta.Visible = true;
			this.colRuta.VisibleIndex = 2;
			// 
			// btAceptar
			// 
			this.btAceptar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btAceptar.Location = new System.Drawing.Point(104, 368);
			this.btAceptar.Name = "btAceptar";
			this.btAceptar.TabIndex = 1;
			this.btAceptar.Text = "&Grabar";
			this.btAceptar.Click += new System.EventHandler(this.btAceptar_Click);
			// 
			// btCancelar
			// 
			this.btCancelar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.btCancelar.DialogResult = System.Windows.Forms.DialogResult.Cancel;
			this.btCancelar.Location = new System.Drawing.Point(584, 368);
			this.btCancelar.Name = "btCancelar";
			this.btCancelar.TabIndex = 2;
			this.btCancelar.Text = "&Salir";
			this.btCancelar.Click += new System.EventHandler(this.btCancelar_Click);
			// 
			// gridControl2
			// 
			this.gridControl2.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.gridControl2.DataMember = "CompraRepDetalle";
			this.gridControl2.DataSource = this.cdsCompraTabla;
			// 
			// gridControl2.EmbeddedNavigator
			// 
			this.gridControl2.EmbeddedNavigator.Name = "";
			this.gridControl2.Location = new System.Drawing.Point(16, 248);
			this.gridControl2.MainView = this.grvDetalle;
			this.gridControl2.Name = "gridControl2";
			this.gridControl2.RepositoryItems.AddRange(new DevExpress.XtraEditors.Repository.RepositoryItem[] {
																																																					this.repositoryItemLookUpEdit2,
																																																					this.repositoryItemMemoEdit1});
			this.gridControl2.Size = new System.Drawing.Size(760, 112);
			this.gridControl2.TabIndex = 3;
			this.gridControl2.ViewCollection.AddRange(new DevExpress.XtraGrid.Views.Base.BaseView[] {
																																																this.grvDetalle});
			// 
			// grvDetalle
			// 
			this.grvDetalle.Columns.AddRange(new DevExpress.XtraGrid.Columns.GridColumn[] {
																																											this.colidCompraRepDet,
																																											this.colTabla,
																																											this.colCampo,
																																											this.colUbicacion,
																																											this.colidCompraReporte1,
																																											this.colExplicacion,
																																											this.colObligatorio,
																																											this.colDatos});
			this.grvDetalle.GridControl = this.gridControl2;
			this.grvDetalle.Name = "grvDetalle";
			this.grvDetalle.OptionsView.NewItemRowPosition = DevExpress.XtraGrid.Views.Grid.NewItemRowPosition.Bottom;
			this.grvDetalle.OptionsView.RowAutoHeight = true;
			this.grvDetalle.OptionsView.ShowGroupPanel = false;
			this.grvDetalle.InitNewRow += new DevExpress.XtraGrid.Views.Grid.InitNewRowEventHandler(this.gridView2_InitNewRow);
			// 
			// colidCompraRepDet
			// 
			this.colidCompraRepDet.Caption = "idCompraRepDet";
			this.colidCompraRepDet.FieldName = "idCompraRepDet";
			this.colidCompraRepDet.Name = "colidCompraRepDet";
			// 
			// colTabla
			// 
			this.colTabla.Caption = "Tabla";
			this.colTabla.FieldName = "Tabla";
			this.colTabla.Name = "colTabla";
			this.colTabla.Visible = true;
			this.colTabla.VisibleIndex = 3;
			this.colTabla.Width = 99;
			// 
			// colCampo
			// 
			this.colCampo.Caption = "Campo";
			this.colCampo.FieldName = "Campo";
			this.colCampo.Name = "colCampo";
			this.colCampo.Visible = true;
			this.colCampo.VisibleIndex = 4;
			this.colCampo.Width = 96;
			// 
			// colUbicacion
			// 
			this.colUbicacion.Caption = "Ubicación";
			this.colUbicacion.ColumnEdit = this.repositoryItemMemoEdit1;
			this.colUbicacion.FieldName = "Ubicacion";
			this.colUbicacion.Name = "colUbicacion";
			this.colUbicacion.Visible = true;
			this.colUbicacion.VisibleIndex = 2;
			this.colUbicacion.Width = 209;
			// 
			// repositoryItemMemoEdit1
			// 
			this.repositoryItemMemoEdit1.Name = "repositoryItemMemoEdit1";
			// 
			// colidCompraReporte1
			// 
			this.colidCompraReporte1.Caption = "idCompraReporte";
			this.colidCompraReporte1.FieldName = "idCompraReporte";
			this.colidCompraReporte1.Name = "colidCompraReporte1";
			// 
			// colExplicacion
			// 
			this.colExplicacion.Caption = "Explicación";
			this.colExplicacion.ColumnEdit = this.repositoryItemMemoEdit1;
			this.colExplicacion.FieldName = "Explicacion";
			this.colExplicacion.Name = "colExplicacion";
			this.colExplicacion.Visible = true;
			this.colExplicacion.VisibleIndex = 1;
			this.colExplicacion.Width = 121;
			// 
			// colObligatorio
			// 
			this.colObligatorio.Caption = "Obligatorio";
			this.colObligatorio.FieldName = "Obligatorio";
			this.colObligatorio.Name = "colObligatorio";
			this.colObligatorio.Visible = true;
			this.colObligatorio.VisibleIndex = 5;
			this.colObligatorio.Width = 85;
			// 
			// colDatos
			// 
			this.colDatos.Caption = "Datos";
			this.colDatos.FieldName = "Datos";
			this.colDatos.Name = "colDatos";
			this.colDatos.Visible = true;
			this.colDatos.VisibleIndex = 0;
			this.colDatos.Width = 121;
			// 
			// repositoryItemLookUpEdit2
			// 
			this.repositoryItemLookUpEdit2.AutoHeight = false;
			this.repositoryItemLookUpEdit2.Buttons.AddRange(new DevExpress.XtraEditors.Controls.EditorButton[] {
																																																					 new DevExpress.XtraEditors.Controls.EditorButton(DevExpress.XtraEditors.Controls.ButtonPredefines.Combo)});
			this.repositoryItemLookUpEdit2.Columns.AddRange(new DevExpress.XtraEditors.Controls.LookUpColumnInfo[] {
																																																							 new DevExpress.XtraEditors.Controls.LookUpColumnInfo("Nombre", "Nombre", 43, DevExpress.Utils.FormatType.None, "", true, DevExpress.Utils.HorzAlignment.Near, DevExpress.Data.ColumnSortOrder.Ascending)});
			this.repositoryItemLookUpEdit2.DataSource = this.cdvCompraNumero;
			this.repositoryItemLookUpEdit2.DisplayMember = "Nombre";
			this.repositoryItemLookUpEdit2.Name = "repositoryItemLookUpEdit2";
			this.repositoryItemLookUpEdit2.NullText = "";
			this.repositoryItemLookUpEdit2.ValueMember = "idTipoFactura";
			// 
			// txtIdCompraRep
			// 
			this.txtIdCompraRep.DataBindings.Add(new System.Windows.Forms.Binding("Text", this.cdsCompraTabla, "CompraReporte.idCompraReporte"));
			this.txtIdCompraRep.Location = new System.Drawing.Point(392, 368);
			this.txtIdCompraRep.Name = "txtIdCompraRep";
			this.txtIdCompraRep.Size = new System.Drawing.Size(120, 20);
			this.txtIdCompraRep.TabIndex = 4;
			this.txtIdCompraRep.Text = "idRep";
			this.txtIdCompraRep.TextChanged += new System.EventHandler(this.txtIdCompraRep_TextChanged);
			// 
			// CompraReporte
			// 
			this.AcceptButton = this.btAceptar;
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 14);
			this.CancelButton = this.btCancelar;
			this.ClientSize = new System.Drawing.Size(792, 406);
			this.Controls.Add(this.txtIdCompraRep);
			this.Controls.Add(this.gridControl2);
			this.Controls.Add(this.btCancelar);
			this.Controls.Add(this.btAceptar);
			this.Controls.Add(this.gridControl1);
			this.Name = "CompraReporte";
			this.Text = "Reportes Especiales";
			this.Load += new System.EventHandler(this.CompraReporte_Load);
			((System.ComponentModel.ISupportInitialize)(this.gridControl1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsCompraTabla)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.gridView1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.repositoryItemLookUpEdit1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvCompraNumero)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.gridControl2)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.grvDetalle)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.repositoryItemMemoEdit1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.repositoryItemLookUpEdit2)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void cdsCompraTabla_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsCompraTabla.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirNumero);
		}

		private void CompraReporte_Load(object sender, System.EventArgs e)
		{
			txtIdCompraRep.Width = 0;
		}

		private void btAceptar_Click(object sender, System.EventArgs e)
		{
			cdsCompraTabla.Update();
//			Close();
		}

		private void btCancelar_Click(object sender, System.EventArgs e)
		{
			Close();
		}

		private void txtIdCompraRep_TextChanged(object sender, System.EventArgs e)
		{
			if (cdsCompraTabla.HasChanges())
			{
				if (DialogResult.Yes == MessageBox.Show("¿Desea Grabar los cambios?", "Confirmación",
					MessageBoxButtons.YesNo, MessageBoxIcon.Question)) 
				{
					cdsCompraTabla.AcceptChanges();
					cdsCompraTabla.Update();
				}
			}
			C1.Data.FilterCondition fcEstudios = new 
				C1.Data.FilterCondition(cdsCompraTabla, "CompraRepDetalle", "idCompraReporte = " + txtIdCompraRep.Text);
			C1.Data.FilterConditions fcTotal = new C1.Data.FilterConditions();

			fcTotal.Add(fcEstudios);
			cdsCompraTabla.Fill(fcTotal, false);
		}

		private void gridView2_InitNewRow(object sender, DevExpress.XtraGrid.Views.Grid.InitNewRowEventArgs e)
		{
			grvDetalle.SetRowCellValue(e.RowHandle, colidCompraReporte1, int.Parse(txtIdCompraRep.Text));
		}

	}
}

