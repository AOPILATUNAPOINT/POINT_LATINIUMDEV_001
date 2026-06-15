using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de FlujoCuenta.
	/// </summary>
	public class FlujoCuenta : DevExpress.XtraEditors.XtraForm
	{
		private Infragistics.Win.Misc.UltraButton btSalir;
		private Infragistics.Win.Misc.UltraButton btGrabar;
		private Infragistics.Win.UltraWinGrid.UltraGrid ultraGrid1;
		private C1.Data.C1DataSet cdsFlujoManual;
		private C1.Data.C1DataSet cdsCuenta;
		private C1.Data.C1DataView cdvCodCuenta;
		private Infragistics.Win.UltraWinGrid.UltraDropDown cmbCodCuenta;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public FlujoCuenta()
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
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("FlujoCuenta", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idFlujoCuenta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCuenta", -1, "cmbCodCuenta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Cuenta", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CodRapido");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Descripcion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCuenta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Grupo");
			this.btSalir = new Infragistics.Win.Misc.UltraButton();
			this.btGrabar = new Infragistics.Win.Misc.UltraButton();
			this.ultraGrid1 = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.cdsFlujoManual = new C1.Data.C1DataSet();
			this.cdsCuenta = new C1.Data.C1DataSet();
			this.cmbCodCuenta = new Infragistics.Win.UltraWinGrid.UltraDropDown();
			this.cdvCodCuenta = new C1.Data.C1DataView();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsFlujoManual)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsCuenta)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCodCuenta)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvCodCuenta)).BeginInit();
			this.SuspendLayout();
			// 
			// btSalir
			// 
			this.btSalir.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.btSalir.DialogResult = System.Windows.Forms.DialogResult.Cancel;
			this.btSalir.Location = new System.Drawing.Point(370, 232);
			this.btSalir.Name = "btSalir";
			this.btSalir.Size = new System.Drawing.Size(64, 24);
			this.btSalir.TabIndex = 5;
			this.btSalir.Text = "&Salir";
			this.btSalir.Click += new System.EventHandler(this.btSalir_Click);
			// 
			// btGrabar
			// 
			this.btGrabar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btGrabar.Location = new System.Drawing.Point(30, 232);
			this.btGrabar.Name = "btGrabar";
			this.btGrabar.Size = new System.Drawing.Size(64, 24);
			this.btGrabar.TabIndex = 4;
			this.btGrabar.Text = "&Grabar";
			this.btGrabar.Click += new System.EventHandler(this.btGrabar_Click);
			// 
			// ultraGrid1
			// 
			this.ultraGrid1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.ultraGrid1.Cursor = System.Windows.Forms.Cursors.Default;
			this.ultraGrid1.DataMember = "FlujoCuenta";
			this.ultraGrid1.DataSource = this.cdsFlujoManual;
			this.ultraGrid1.DisplayLayout.AutoFitColumns = true;
			ultraGridBand1.AddButtonCaption = "FlujoManual";
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn2.Header.Caption = "Cuenta";
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridColumn2.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownValidate;
			ultraGridColumn2.Width = 235;
			ultraGridColumn3.Header.VisiblePosition = 2;
			ultraGridColumn3.Width = 176;
			ultraGridBand1.Columns.AddRange(new object[] {
																										 ultraGridColumn1,
																										 ultraGridColumn2,
																										 ultraGridColumn3});
			this.ultraGrid1.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.ultraGrid1.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.TemplateOnBottom;
			this.ultraGrid1.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.True;
			this.ultraGrid1.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			this.ultraGrid1.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti;
			this.ultraGrid1.DisplayLayout.ViewStyleBand = Infragistics.Win.UltraWinGrid.ViewStyleBand.OutlookGroupBy;
			this.ultraGrid1.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.ultraGrid1.Location = new System.Drawing.Point(16, 16);
			this.ultraGrid1.Name = "ultraGrid1";
			this.ultraGrid1.Size = new System.Drawing.Size(432, 200);
			this.ultraGrid1.TabIndex = 3;
			// 
			// cdsFlujoManual
			// 
			this.cdsFlujoManual.BindingContextCtrl = this;
			this.cdsFlujoManual.DataLibrary = "LibFacturacion";
			this.cdsFlujoManual.DataLibraryUrl = "";
			this.cdsFlujoManual.DataSetDef = "dsFlujo";
			this.cdsFlujoManual.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsFlujoManual.Name = "cdsFlujoManual";
			this.cdsFlujoManual.SchemaClassName = "LibFacturacion.DataClass";
			this.cdsFlujoManual.SchemaDef = null;
			this.cdsFlujoManual.BeforeFill += new C1.Data.FillEventHandler(this.cdsFlujoManual_BeforeFill);
			// 
			// cdsCuenta
			// 
			this.cdsCuenta.BindingContextCtrl = this;
			this.cdsCuenta.DataLibrary = "LibContabilidad";
			this.cdsCuenta.DataLibraryUrl = "";
			this.cdsCuenta.DataSetDef = "dsCuenta2";
			this.cdsCuenta.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsCuenta.Name = "cdsCuenta";
			this.cdsCuenta.SchemaClassName = "LibContabilidad.DataClass";
			this.cdsCuenta.SchemaDef = null;
			// 
			// cmbCodCuenta
			// 
			this.cmbCodCuenta.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbCodCuenta.DataSource = this.cdvCodCuenta;
			ultraGridColumn4.Header.VisiblePosition = 0;
			ultraGridColumn4.Hidden = true;
			ultraGridColumn5.Header.VisiblePosition = 1;
			ultraGridColumn6.Header.VisiblePosition = 2;
			ultraGridColumn7.Header.VisiblePosition = 3;
			ultraGridColumn7.Hidden = true;
			ultraGridColumn8.Header.VisiblePosition = 4;
			ultraGridColumn8.Hidden = true;
			ultraGridBand2.Columns.AddRange(new object[] {
																										 ultraGridColumn4,
																										 ultraGridColumn5,
																										 ultraGridColumn6,
																										 ultraGridColumn7,
																										 ultraGridColumn8});
			this.cmbCodCuenta.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			this.cmbCodCuenta.DisplayMember = "CodRapido";
			this.cmbCodCuenta.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbCodCuenta.Location = new System.Drawing.Point(208, 104);
			this.cmbCodCuenta.Name = "cmbCodCuenta";
			this.cmbCodCuenta.Size = new System.Drawing.Size(248, 80);
			this.cmbCodCuenta.TabIndex = 6;
			this.cmbCodCuenta.Text = "ultraDropDown1";
			this.cmbCodCuenta.ValueMember = "idCuenta";
			this.cmbCodCuenta.Visible = false;
			// 
			// cdvCodCuenta
			// 
			this.cdvCodCuenta.BindingContextCtrl = this;
			this.cdvCodCuenta.DataSet = this.cdsCuenta;
			this.cdvCodCuenta.Sort = "Codigo";
			this.cdvCodCuenta.TableName = "";
			this.cdvCodCuenta.TableViewName = "Cuenta";
			// 
			// FlujoCuenta
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(472, 273);
			this.Controls.Add(this.cmbCodCuenta);
			this.Controls.Add(this.btSalir);
			this.Controls.Add(this.btGrabar);
			this.Controls.Add(this.ultraGrid1);
			this.Name = "FlujoCuenta";
			this.Text = "Cuentas del Flujo";
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsFlujoManual)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsCuenta)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCodCuenta)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvCodCuenta)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void btGrabar_Click(object sender, System.EventArgs e)
		{
			cdsFlujoManual.Update();
			MessageBox.Show("Informacion grabada", "Información");
		}

		private void btSalir_Click(object sender, System.EventArgs e)
		{
			Close();
		}

		private void cdsFlujoManual_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			this.cdsFlujoManual.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);		
		}
	}
}
