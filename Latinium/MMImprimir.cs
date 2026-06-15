using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de MMImprimir.
	/// </summary>
	public class MMImprimir : DevExpress.XtraEditors.XtraForm
	{
		private C1.Data.C1DataSet cdsMarceloMaestro;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbMarceloMaestro;
		private Infragistics.Win.Misc.UltraButton btImprimir;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public MMImprimir()
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
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("MarceloMaestro", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idMarceloMaestro");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Fecha");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Descripcion");
			this.cdsMarceloMaestro = new C1.Data.C1DataSet();
			this.cmbMarceloMaestro = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.btImprimir = new Infragistics.Win.Misc.UltraButton();
			((System.ComponentModel.ISupportInitialize)(this.cdsMarceloMaestro)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbMarceloMaestro)).BeginInit();
			this.SuspendLayout();
			// 
			// cdsMarceloMaestro
			// 
			this.cdsMarceloMaestro.BindingContextCtrl = this;
			this.cdsMarceloMaestro.DataLibrary = "LibFacturacion";
			this.cdsMarceloMaestro.DataLibraryUrl = "";
			this.cdsMarceloMaestro.DataSetDef = "MarceloMaestro";
			this.cdsMarceloMaestro.Locale = new System.Globalization.CultureInfo("es-MX");
			this.cdsMarceloMaestro.Name = "cdsMarceloMaestro";
			this.cdsMarceloMaestro.SchemaClassName = "LibFacturacion.DataClass";
			this.cdsMarceloMaestro.SchemaDef = null;
			// 
			// cmbMarceloMaestro
			// 
			this.cmbMarceloMaestro.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbMarceloMaestro.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbMarceloMaestro.DataSource = this.cdsMarceloMaestro;
			this.cmbMarceloMaestro.DisplayLayout.AutoFitColumns = true;
			ultraGridBand1.AddButtonCaption = "MarceloCodigo";
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn1.Width = 14;
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridColumn2.Width = 76;
			ultraGridColumn3.Header.VisiblePosition = 2;
			ultraGridColumn3.Width = 365;
			ultraGridBand1.Columns.AddRange(new object[] {
																										 ultraGridColumn1,
																										 ultraGridColumn2,
																										 ultraGridColumn3});
			this.cmbMarceloMaestro.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.cmbMarceloMaestro.DisplayMember = "Descripcion";
			this.cmbMarceloMaestro.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbMarceloMaestro.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbMarceloMaestro.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbMarceloMaestro.Location = new System.Drawing.Point(16, 16);
			this.cmbMarceloMaestro.Name = "cmbMarceloMaestro";
			this.cmbMarceloMaestro.Size = new System.Drawing.Size(472, 21);
			this.cmbMarceloMaestro.TabIndex = 6;
			this.cmbMarceloMaestro.ValueMember = "idMarceloMaestro";
			// 
			// btImprimir
			// 
			this.btImprimir.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btImprimir.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.btImprimir.Font = new System.Drawing.Font("Microsoft Sans Serif", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.btImprimir.Location = new System.Drawing.Point(384, 88);
			this.btImprimir.Name = "btImprimir";
			this.btImprimir.Size = new System.Drawing.Size(104, 23);
			this.btImprimir.TabIndex = 144;
			this.btImprimir.Text = "Imprimir";
			this.btImprimir.Click += new System.EventHandler(this.btImprimir_Click);
			// 
			// MMImprimir
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(504, 125);
			this.Controls.Add(this.btImprimir);
			this.Controls.Add(this.cmbMarceloMaestro);
			this.Name = "MMImprimir";
			this.Text = "MMImprimir";
			((System.ComponentModel.ISupportInitialize)(this.cdsMarceloMaestro)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbMarceloMaestro)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void btImprimir_Click(object sender, System.EventArgs e)
		{
			Cursor = Cursors.WaitCursor;
			string stFiltro="";
			stFiltro += " {MarceloCodigo.idMarceloMaestro}=" + this.cmbMarceloMaestro.Value.ToString().Trim();
			Reporte miReporte = new Reporte("MMReparte.rpt", stFiltro);
			miReporte.MdiParent = this.MdiParent;
			//miReporte.Titulo(stTitulo);
			miReporte.Show();
			Cursor = Cursors.Default;
		}
	}
}
