using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de Moneda.
	/// </summary>
	public class Moneda : DevExpress.XtraEditors.XtraForm
	{
		private DevExpress.XtraGrid.GridControl gridControl1;
		private DevExpress.XtraGrid.Views.Grid.GridView gridView1;
		private DevExpress.XtraEditors.SimpleButton btAceptar;
		private DevExpress.XtraEditors.SimpleButton btCancelar;
		private C1.Data.C1DataSet cdsMoneda;
		private C1.Data.C1DataView cdvMoneda;
		private DevExpress.XtraGrid.Columns.GridColumn colNombre;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public Moneda()
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
			this.gridControl1 = new DevExpress.XtraGrid.GridControl();
			this.cdvMoneda = new C1.Data.C1DataView();
			this.cdsMoneda = new C1.Data.C1DataSet();
			this.gridView1 = new DevExpress.XtraGrid.Views.Grid.GridView();
			this.colNombre = new DevExpress.XtraGrid.Columns.GridColumn();
			this.btAceptar = new DevExpress.XtraEditors.SimpleButton();
			this.btCancelar = new DevExpress.XtraEditors.SimpleButton();
			((System.ComponentModel.ISupportInitialize)(this.gridControl1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvMoneda)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsMoneda)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.gridView1)).BeginInit();
			this.SuspendLayout();
			// 
			// gridControl1
			// 
			this.gridControl1.DataSource = this.cdvMoneda;
			// 
			// gridControl1.EmbeddedNavigator
			// 
			this.gridControl1.EmbeddedNavigator.Name = "";
			this.gridControl1.Location = new System.Drawing.Point(16, 16);
			this.gridControl1.MainView = this.gridView1;
			this.gridControl1.Name = "gridControl1";
			this.gridControl1.Size = new System.Drawing.Size(400, 208);
			this.gridControl1.Styles.AddReplace("ColumnFilterButtonActive", new DevExpress.Utils.ViewStyleEx("ColumnFilterButtonActive", "Grid", System.Drawing.SystemColors.Control, System.Drawing.Color.Blue, System.Drawing.SystemColors.ControlLightLight, System.Drawing.Drawing2D.LinearGradientMode.Horizontal));
			this.gridControl1.TabIndex = 0;
			this.gridControl1.Text = "gridControl1";
			// 
			// cdvMoneda
			// 
			this.cdvMoneda.BindingContextCtrl = this;
			this.cdvMoneda.DataSet = this.cdsMoneda;
			this.cdvMoneda.TableName = "";
			this.cdvMoneda.TableViewName = "MonedaTipo";
			// 
			// cdsMoneda
			// 
			this.cdsMoneda.BindingContextCtrl = this;
			this.cdsMoneda.DataLibrary = "LibContabilidad";
			this.cdsMoneda.DataLibraryUrl = "";
			this.cdsMoneda.DataSetDef = "dsMoneda";
			this.cdsMoneda.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsMoneda.Name = "cdsMoneda";
			this.cdsMoneda.SchemaClassName = "LibContabilidad.DataClass";
			this.cdsMoneda.SchemaDef = null;
			this.cdsMoneda.BeforeFill += new C1.Data.FillEventHandler(this.cdsMoneda_BeforeFill);
			// 
			// gridView1
			// 
			this.gridView1.Columns.AddRange(new DevExpress.XtraGrid.Columns.GridColumn[] {
																																										 this.colNombre});
			this.gridView1.Name = "gridView1";
			this.gridView1.OptionsView.ShowGroupPanel = false;
			// 
			// colNombre
			// 
			this.colNombre.Caption = "Nombre";
			this.colNombre.FieldName = "Nombre";
			this.colNombre.Name = "colNombre";
			this.colNombre.VisibleIndex = 0;
			// 
			// btAceptar
			// 
			this.btAceptar.Location = new System.Drawing.Point(64, 240);
			this.btAceptar.Name = "btAceptar";
			this.btAceptar.TabIndex = 1;
			this.btAceptar.Text = "&Aceptar";
			this.btAceptar.Click += new System.EventHandler(this.btAceptar_Click);
			// 
			// btCancelar
			// 
			this.btCancelar.Location = new System.Drawing.Point(264, 240);
			this.btCancelar.Name = "btCancelar";
			this.btCancelar.TabIndex = 2;
			this.btCancelar.Text = "&Cancelar";
			this.btCancelar.Click += new System.EventHandler(this.btCancelar_Click);
			// 
			// Moneda
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(424, 285);
			this.Controls.Add(this.btCancelar);
			this.Controls.Add(this.btAceptar);
			this.Controls.Add(this.gridControl1);
			this.Name = "Moneda";
			this.Text = "Moneda";
			this.Closing += new System.ComponentModel.CancelEventHandler(this.Moneda_Closing);
			this.Load += new System.EventHandler(this.Moneda_Load);
			((System.ComponentModel.ISupportInitialize)(this.gridControl1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvMoneda)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsMoneda)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.gridView1)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void btAceptar_Click(object sender, System.EventArgs e)
		{
			this.cdsMoneda.Update();
			this.Close();
		}

		private void btCancelar_Click(object sender, System.EventArgs e)
		{
			this.Close();
		}

		private void cdsMoneda_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			this.cdsMoneda.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);
		}

		private void Moneda_Load(object sender, System.EventArgs e)
		{
			string stAudita = "Exec AuditaCrear 53, 6, 'Moneda'";
			Funcion.EjecutaSQL(cdsMoneda, stAudita, true);
			Cursor = Cursors.Default;
		}

		private void Moneda_Closing(object sender, System.ComponentModel.CancelEventArgs e)
		{
			string stAudita = "Exec AuditaCrear 53, 7, 'Moneda'";
			Funcion.EjecutaSQL(cdsMoneda, stAudita, true);
		}
	}
}
