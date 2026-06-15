using System;
using System.Data;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using Infragistics.Win.UltraWinGrid;
using C1.Data;
using CrystalDecisions.Shared;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de RecetaArticulo.
	/// </summary>
	public class RecetaArticulo : DevExpress.XtraEditors.XtraForm
	{
		private System.Windows.Forms.Label label1;
		private Infragistics.Win.Misc.UltraButton btImprime;
		private C1.Data.C1DataSet cdsArticulo;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbMaPri;
		private C1.Data.C1DataView cdvArticulo;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public RecetaArticulo()
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
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Articulo", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Descontinuado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idArticulo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("dCompra");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("dVenta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idTipoGrupo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Articulo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idGrupoArticulo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ExigeLote");
			this.cmbMaPri = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.cdvArticulo = new C1.Data.C1DataView();
			this.cdsArticulo = new C1.Data.C1DataSet();
			this.label1 = new System.Windows.Forms.Label();
			this.btImprime = new Infragistics.Win.Misc.UltraButton();
			((System.ComponentModel.ISupportInitialize)(this.cmbMaPri)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvArticulo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsArticulo)).BeginInit();
			this.SuspendLayout();
			// 
			// cmbMaPri
			// 
			this.cmbMaPri.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbMaPri.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbMaPri.DataSource = this.cdvArticulo;
			ultraGridBand1.AddButtonCaption = "Producto";
			ultraGridColumn1.Header.VisiblePosition = 3;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn2.Header.VisiblePosition = 2;
			ultraGridColumn2.Hidden = true;
			ultraGridColumn3.Header.VisiblePosition = 4;
			ultraGridColumn3.Hidden = true;
			ultraGridColumn4.Header.VisiblePosition = 5;
			ultraGridColumn4.Hidden = true;
			ultraGridColumn5.Header.VisiblePosition = 6;
			ultraGridColumn5.Hidden = true;
			ultraGridColumn6.Header.VisiblePosition = 1;
			ultraGridColumn6.Width = 217;
			ultraGridColumn7.Header.VisiblePosition = 0;
			ultraGridColumn7.Width = 126;
			ultraGridColumn8.Header.VisiblePosition = 7;
			ultraGridColumn8.Hidden = true;
			ultraGridColumn9.Header.VisiblePosition = 8;
			ultraGridBand1.Columns.AddRange(new object[] {
																										 ultraGridColumn1,
																										 ultraGridColumn2,
																										 ultraGridColumn3,
																										 ultraGridColumn4,
																										 ultraGridColumn5,
																										 ultraGridColumn6,
																										 ultraGridColumn7,
																										 ultraGridColumn8,
																										 ultraGridColumn9});
			this.cmbMaPri.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.cmbMaPri.DisplayMember = "Articulo";
			this.cmbMaPri.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbMaPri.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbMaPri.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbMaPri.Location = new System.Drawing.Point(128, 16);
			this.cmbMaPri.Name = "cmbMaPri";
			this.cmbMaPri.Size = new System.Drawing.Size(208, 21);
			this.cmbMaPri.TabIndex = 15;
			this.cmbMaPri.ValueMember = "idArticulo";
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
			this.cdsArticulo.CaseSensitive = false;
			this.cdsArticulo.DataLibrary = "LibFacturacion";
			this.cdsArticulo.DataLibraryUrl = "";
			this.cdsArticulo.DataSetDef = "dsArticulos";
			this.cdsArticulo.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsArticulo.Name = "cdsArticulo";
			this.cdsArticulo.SchemaClassName = "LibFacturacion.DataClass";
			this.cdsArticulo.SchemaDef = null;
			// 
			// label1
			// 
			this.label1.AutoSize = true;
			this.label1.Location = new System.Drawing.Point(16, 16);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(78, 16);
			this.label1.TabIndex = 16;
			this.label1.Text = "Materia Prima:";
			// 
			// btImprime
			// 
			this.btImprime.Location = new System.Drawing.Point(160, 64);
			this.btImprime.Name = "btImprime";
			this.btImprime.Size = new System.Drawing.Size(72, 24);
			this.btImprime.TabIndex = 23;
			this.btImprime.Text = "Imprime";
			this.btImprime.Click += new System.EventHandler(this.btImprime_Click);
			// 
			// RecetaArticulo
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(400, 117);
			this.Controls.Add(this.btImprime);
			this.Controls.Add(this.label1);
			this.Controls.Add(this.cmbMaPri);
			this.Name = "RecetaArticulo";
			this.Text = "Receta Artículo";
			((System.ComponentModel.ISupportInitialize)(this.cmbMaPri)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvArticulo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsArticulo)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void btImprime_Click(object sender, System.EventArgs e)
		{
			string stFiltro = "";
			stFiltro="{ProductoDetalle.IdArticulo} = " + cmbMaPri.Value;

			Cursor = Cursors.WaitCursor;
			Reporte miReporte = new Reporte("RecetaDetalle.rpt",stFiltro);
			miReporte.MdiParent = this.MdiParent;
			miReporte.Show();
			Cursor = Cursors.Default;
		}

		
	}
}
