using System;
using System.Data.SqlClient;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de CompraImporImp.
	/// </summary>
	public class CompraImporImp : DevExpress.XtraEditors.XtraForm
	{
		private Infragistics.Win.Misc.UltraButton btAceptar;
		private Infragistics.Win.Misc.UltraButton btCancelar;
		private System.Windows.Forms.Label label1;
		private System.Windows.Forms.Label label2;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtDesde;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtHasta;
		private Infragistics.Win.Misc.UltraButton btGenerar;
		private C1.Data.C1DataSet cdsSeteoF;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo cmbFecha;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public CompraImporImp()
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
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton1 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.txtDesde = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtHasta = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.btAceptar = new Infragistics.Win.Misc.UltraButton();
			this.btCancelar = new Infragistics.Win.Misc.UltraButton();
			this.label1 = new System.Windows.Forms.Label();
			this.label2 = new System.Windows.Forms.Label();
			this.btGenerar = new Infragistics.Win.Misc.UltraButton();
			this.cmbFecha = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDesde)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtHasta)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbFecha)).BeginInit();
			this.SuspendLayout();
			// 
			// cdsSeteoF
			// 
			this.cdsSeteoF.BindingContextCtrl = this;
			this.cdsSeteoF.DataLibrary = "LibFacturacion";
			this.cdsSeteoF.DataLibraryUrl = "";
			this.cdsSeteoF.DataSetDef = "dsSeteoF";
			this.cdsSeteoF.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsSeteoF.Name = "cdsSeteoF";
			this.cdsSeteoF.SchemaClassName = "LibFacturacion.DataClass";
			this.cdsSeteoF.SchemaDef = null;
			// 
			// txtDesde
			// 
			this.txtDesde.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtDesde.Location = new System.Drawing.Point(136, 48);
			this.txtDesde.Name = "txtDesde";
			this.txtDesde.Size = new System.Drawing.Size(128, 21);
			this.txtDesde.TabIndex = 0;
			// 
			// txtHasta
			// 
			this.txtHasta.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtHasta.Location = new System.Drawing.Point(136, 80);
			this.txtHasta.Name = "txtHasta";
			this.txtHasta.Size = new System.Drawing.Size(128, 21);
			this.txtHasta.TabIndex = 1;
			// 
			// btAceptar
			// 
			this.btAceptar.Location = new System.Drawing.Point(48, 192);
			this.btAceptar.Name = "btAceptar";
			this.btAceptar.TabIndex = 2;
			this.btAceptar.Text = "&Aceptar";
			this.btAceptar.Click += new System.EventHandler(this.btAceptar_Click);
			// 
			// btCancelar
			// 
			this.btCancelar.Location = new System.Drawing.Point(184, 192);
			this.btCancelar.Name = "btCancelar";
			this.btCancelar.TabIndex = 3;
			this.btCancelar.Text = "&Cancelar";
			this.btCancelar.Click += new System.EventHandler(this.btCancelar_Click);
			// 
			// label1
			// 
			this.label1.AutoSize = true;
			this.label1.Location = new System.Drawing.Point(48, 48);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(40, 16);
			this.label1.TabIndex = 4;
			this.label1.Text = "Desde:";
			// 
			// label2
			// 
			this.label2.AutoSize = true;
			this.label2.Location = new System.Drawing.Point(48, 80);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(37, 16);
			this.label2.TabIndex = 5;
			this.label2.Text = "Hasta:";
			// 
			// btGenerar
			// 
			this.btGenerar.Location = new System.Drawing.Point(192, 120);
			this.btGenerar.Name = "btGenerar";
			this.btGenerar.TabIndex = 6;
			this.btGenerar.Text = "&Generar";
			this.btGenerar.Click += new System.EventHandler(this.btGenerar_Click);
			// 
			// cmbFecha
			// 
			dateButton1.Caption = "Today";
			this.cmbFecha.DateButtons.Add(dateButton1);
			this.cmbFecha.Format = "dd/MMM/yyyy";
			this.cmbFecha.Location = new System.Drawing.Point(48, 120);
			this.cmbFecha.Name = "cmbFecha";
			this.cmbFecha.NonAutoSizeHeight = 23;
			this.cmbFecha.Size = new System.Drawing.Size(120, 21);
			this.cmbFecha.SpinButtonIncrement = Infragistics.Win.UltraWinSchedule.SpinIncrementUnit.Years;
			this.cmbFecha.SpinButtonsVisible = true;
			this.cmbFecha.TabIndex = 7;
			this.cmbFecha.Value = new System.DateTime(2007, 5, 31, 0, 0, 0, 0);
			// 
			// CompraImporImp
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(296, 245);
			this.Controls.Add(this.cmbFecha);
			this.Controls.Add(this.btGenerar);
			this.Controls.Add(this.label2);
			this.Controls.Add(this.label1);
			this.Controls.Add(this.btCancelar);
			this.Controls.Add(this.btAceptar);
			this.Controls.Add(this.txtHasta);
			this.Controls.Add(this.txtDesde);
			this.Name = "CompraImporImp";
			this.Text = "Impresion de Plan de Compras";
			this.Load += new System.EventHandler(this.CompraImporImp_Load);
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDesde)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtHasta)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbFecha)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void btAceptar_Click(object sender, System.EventArgs e)
		{
			Cursor = Cursors.WaitCursor;
			string stFiltro = "";
			string stSigno = "= '";
			if (txtHasta.Text.Trim().Length > 0)
				stSigno = ">= '";
			if (txtDesde.Text.Trim().Length > 0)
				stFiltro += "{Articulo.Codigo} " + stSigno + txtDesde.Text.Trim() + "'";
			if (txtHasta.Text.Trim().Length > 0)
				stFiltro += " And {Articulo.Codigo} <= '" + txtHasta.Text.Trim() + "'";
			Reporte miRep = new Reporte("PlanCompra", stFiltro);
			miRep.ShowDialog();
			Cursor = Cursors.Default;
		}

		private void btGenerar_Click(object sender, System.EventArgs e)
		{
			DateTime dtFecha = (DateTime) cmbFecha.Value;
			string stCod1 = txtDesde.Text.Trim();
			string stCod2 = txtHasta.Text.Trim();
			if (stCod2.Length == 0) stCod2 = stCod1;
			string stSelect = "SELECT idArticulo FROM Articulo WHERE Codigo >= '" + stCod1
				+ "' AND Codigo <= '" + stCod2 + "' ORDER BY Codigo";
			SqlDataReader miReader = Funcion.rEscalarSQL(cdsSeteoF, stSelect, true);
			int [] IdArt = new Int32[1000];
			int i = 0;
			while (miReader.Read())
			{
				IdArt[i] = miReader.GetInt32(0);
				i++;
			}
			miReader.Close();

			if (DialogResult.No == MessageBox.Show("¿Desea Generar " + i.ToString() + " Planes de compra?",
				"Confirmación", MessageBoxButtons.YesNo, MessageBoxIcon.Question)) return;
			Cursor = Cursors.WaitCursor;
			stSelect = "Delete TmpPlanCompra";
			Funcion.EjecutaSQL(cdsSeteoF, stSelect, true);
			for (int i1=0; i1< i; i1++)
			{
				stSelect = "Exec PlanCompras " + IdArt[i1].ToString() + ",'" + dtFecha.ToString("yyyyMMdd") + "'";
				Funcion.EjecutaSQL(cdsSeteoF, stSelect, true);
			}
			Cursor = Cursors.Default;
			MessageBox.Show("Plan de Compras Generado", "Información");
		}

		private void btCancelar_Click(object sender, System.EventArgs e)
		{
			Close();
		}

		private void CompraImporImp_Load(object sender, System.EventArgs e)
		{
			cmbFecha.Value = (DateTime) DateTime.Today;
		}
	}
}
