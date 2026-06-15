using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de CompraArt.
	/// </summary>
	public class CompraArt : DevExpress.XtraEditors.XtraForm
	{
		private Infragistics.Win.Misc.UltraLabel ultraLabel1;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor ultraNumericEditor1;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor ultraNumericEditor2;
		private Infragistics.Win.Misc.UltraLabel ultraLabel2;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor ultraNumericEditor3;
		private Infragistics.Win.Misc.UltraLabel ultraLabel3;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor ultraNumericEditor4;
		private Infragistics.Win.Misc.UltraLabel ultraLabel4;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor ultraNumericEditor5;
		private Infragistics.Win.Misc.UltraLabel ultraLabel5;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor ultraNumericEditor6;
		private Infragistics.Win.Misc.UltraLabel ultraLabel6;
		private Infragistics.Win.Misc.UltraButton btAceptar;
		private BarraDatoSQL.BarraDatoSQL barraDatoSQL1;
		private C1.Data.C1DataSet cdsCompra;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor ultraNumericEditor7;
		private Infragistics.Win.Misc.UltraLabel ultraLabel7;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public CompraArt()
		{
			InitializeComponent();
		}
		int IdArticulo = 0;
		public CompraArt(int idArticulo)
		{
			InitializeComponent();
			IdArticulo = idArticulo;
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
			this.ultraLabel1 = new Infragistics.Win.Misc.UltraLabel();
			this.ultraNumericEditor1 = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.cdsCompra = new C1.Data.C1DataSet();
			this.ultraNumericEditor2 = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.ultraLabel2 = new Infragistics.Win.Misc.UltraLabel();
			this.ultraNumericEditor3 = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.ultraLabel3 = new Infragistics.Win.Misc.UltraLabel();
			this.ultraNumericEditor4 = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.ultraLabel4 = new Infragistics.Win.Misc.UltraLabel();
			this.ultraNumericEditor5 = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.ultraLabel5 = new Infragistics.Win.Misc.UltraLabel();
			this.ultraNumericEditor6 = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.ultraLabel6 = new Infragistics.Win.Misc.UltraLabel();
			this.btAceptar = new Infragistics.Win.Misc.UltraButton();
			this.barraDatoSQL1 = new BarraDatoSQL.BarraDatoSQL();
			this.ultraNumericEditor7 = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.ultraLabel7 = new Infragistics.Win.Misc.UltraLabel();
			((System.ComponentModel.ISupportInitialize)(this.ultraNumericEditor1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsCompra)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraNumericEditor2)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraNumericEditor3)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraNumericEditor4)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraNumericEditor5)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraNumericEditor6)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraNumericEditor7)).BeginInit();
			this.SuspendLayout();
			// 
			// ultraLabel1
			// 
			this.ultraLabel1.AutoSize = true;
			this.ultraLabel1.Location = new System.Drawing.Point(32, 26);
			this.ultraLabel1.Name = "ultraLabel1";
			this.ultraLabel1.Size = new System.Drawing.Size(44, 15);
			this.ultraLabel1.TabIndex = 3;
			this.ultraLabel1.Text = "Precio 1";
			// 
			// ultraNumericEditor1
			// 
			this.ultraNumericEditor1.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsCompra, "Articulo.Precio1"));
			this.ultraNumericEditor1.FormatString = "#,##0.00";
			this.ultraNumericEditor1.Location = new System.Drawing.Point(120, 26);
			this.ultraNumericEditor1.Name = "ultraNumericEditor1";
			this.ultraNumericEditor1.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.ultraNumericEditor1.PromptChar = ' ';
			this.ultraNumericEditor1.Size = new System.Drawing.Size(100, 22);
			this.ultraNumericEditor1.TabIndex = 4;
			// 
			// cdsCompra
			// 
			this.cdsCompra.BindingContextCtrl = this;
			this.cdsCompra.DataLibrary = "LibFacturacion";
			this.cdsCompra.DataLibraryUrl = "";
			this.cdsCompra.DataSetDef = "dsArticulo";
			this.cdsCompra.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsCompra.Name = "cdsCompra";
			this.cdsCompra.SchemaClassName = "LibFacturacion.DataClass";
			this.cdsCompra.SchemaDef = null;
			// 
			// ultraNumericEditor2
			// 
			this.ultraNumericEditor2.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsCompra, "Articulo.Precio2"));
			this.ultraNumericEditor2.FormatString = "#,##0.00";
			this.ultraNumericEditor2.Location = new System.Drawing.Point(120, 60);
			this.ultraNumericEditor2.Name = "ultraNumericEditor2";
			this.ultraNumericEditor2.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.ultraNumericEditor2.PromptChar = ' ';
			this.ultraNumericEditor2.Size = new System.Drawing.Size(100, 22);
			this.ultraNumericEditor2.TabIndex = 6;
			// 
			// ultraLabel2
			// 
			this.ultraLabel2.AutoSize = true;
			this.ultraLabel2.Location = new System.Drawing.Point(32, 60);
			this.ultraLabel2.Name = "ultraLabel2";
			this.ultraLabel2.Size = new System.Drawing.Size(44, 15);
			this.ultraLabel2.TabIndex = 5;
			this.ultraLabel2.Text = "Precio 2";
			// 
			// ultraNumericEditor3
			// 
			this.ultraNumericEditor3.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsCompra, "Articulo.Precio3"));
			this.ultraNumericEditor3.FormatString = "#,##0.00";
			this.ultraNumericEditor3.Location = new System.Drawing.Point(120, 95);
			this.ultraNumericEditor3.Name = "ultraNumericEditor3";
			this.ultraNumericEditor3.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.ultraNumericEditor3.PromptChar = ' ';
			this.ultraNumericEditor3.Size = new System.Drawing.Size(100, 22);
			this.ultraNumericEditor3.TabIndex = 8;
			// 
			// ultraLabel3
			// 
			this.ultraLabel3.AutoSize = true;
			this.ultraLabel3.Location = new System.Drawing.Point(32, 95);
			this.ultraLabel3.Name = "ultraLabel3";
			this.ultraLabel3.Size = new System.Drawing.Size(44, 15);
			this.ultraLabel3.TabIndex = 7;
			this.ultraLabel3.Text = "Precio 3";
			// 
			// ultraNumericEditor4
			// 
			this.ultraNumericEditor4.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsCompra, "Articulo.Precio4"));
			this.ultraNumericEditor4.FormatString = "#,##0.00";
			this.ultraNumericEditor4.Location = new System.Drawing.Point(120, 129);
			this.ultraNumericEditor4.Name = "ultraNumericEditor4";
			this.ultraNumericEditor4.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.ultraNumericEditor4.PromptChar = ' ';
			this.ultraNumericEditor4.Size = new System.Drawing.Size(100, 22);
			this.ultraNumericEditor4.TabIndex = 10;
			// 
			// ultraLabel4
			// 
			this.ultraLabel4.AutoSize = true;
			this.ultraLabel4.Location = new System.Drawing.Point(32, 129);
			this.ultraLabel4.Name = "ultraLabel4";
			this.ultraLabel4.Size = new System.Drawing.Size(44, 15);
			this.ultraLabel4.TabIndex = 9;
			this.ultraLabel4.Text = "Precio 4";
			// 
			// ultraNumericEditor5
			// 
			this.ultraNumericEditor5.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsCompra, "Articulo.Precio5"));
			this.ultraNumericEditor5.FormatString = "#,##0.00";
			this.ultraNumericEditor5.Location = new System.Drawing.Point(120, 164);
			this.ultraNumericEditor5.Name = "ultraNumericEditor5";
			this.ultraNumericEditor5.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.ultraNumericEditor5.PromptChar = ' ';
			this.ultraNumericEditor5.Size = new System.Drawing.Size(100, 22);
			this.ultraNumericEditor5.TabIndex = 12;
			// 
			// ultraLabel5
			// 
			this.ultraLabel5.AutoSize = true;
			this.ultraLabel5.Location = new System.Drawing.Point(32, 164);
			this.ultraLabel5.Name = "ultraLabel5";
			this.ultraLabel5.Size = new System.Drawing.Size(44, 15);
			this.ultraLabel5.TabIndex = 11;
			this.ultraLabel5.Text = "Precio 5";
			// 
			// ultraNumericEditor6
			// 
			this.ultraNumericEditor6.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsCompra, "Articulo.CostoPromedio"));
			this.ultraNumericEditor6.FormatString = "#,##0.00";
			this.ultraNumericEditor6.Location = new System.Drawing.Point(120, 198);
			this.ultraNumericEditor6.Name = "ultraNumericEditor6";
			this.ultraNumericEditor6.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.ultraNumericEditor6.PromptChar = ' ';
			this.ultraNumericEditor6.Size = new System.Drawing.Size(100, 22);
			this.ultraNumericEditor6.TabIndex = 14;
			// 
			// ultraLabel6
			// 
			this.ultraLabel6.AutoSize = true;
			this.ultraLabel6.Location = new System.Drawing.Point(32, 233);
			this.ultraLabel6.Name = "ultraLabel6";
			this.ultraLabel6.Size = new System.Drawing.Size(68, 15);
			this.ultraLabel6.TabIndex = 13;
			this.ultraLabel6.Text = "Costo Ultimo";
			// 
			// btAceptar
			// 
			this.btAceptar.DialogResult = System.Windows.Forms.DialogResult.Cancel;
			this.btAceptar.Location = new System.Drawing.Point(80, 267);
			this.btAceptar.Name = "btAceptar";
			this.btAceptar.Size = new System.Drawing.Size(75, 25);
			this.btAceptar.TabIndex = 15;
			this.btAceptar.Text = "&Aceptar";
			this.btAceptar.Click += new System.EventHandler(this.btAceptar_Click);
			// 
			// barraDatoSQL1
			// 
			this.barraDatoSQL1.BarraMovimiento = true;
			this.barraDatoSQL1.DataMember = "idArticulo";
			this.barraDatoSQL1.DataNombreId = "idArticulo";
			this.barraDatoSQL1.DataOrden = "";
			this.barraDatoSQL1.DataSource = this.cdsCompra;
			this.barraDatoSQL1.DataTabla = "Articulo";
			this.barraDatoSQL1.DataTablaHija = "";
			this.barraDatoSQL1.Location = new System.Drawing.Point(8, 284);
			this.barraDatoSQL1.Name = "barraDatoSQL1";
			this.barraDatoSQL1.Size = new System.Drawing.Size(304, 26);
			this.barraDatoSQL1.TabIndex = 16;
			this.barraDatoSQL1.VisibleBorrar = true;
			this.barraDatoSQL1.VisibleBuscar = true;
			this.barraDatoSQL1.VisibleEditar = true;
			this.barraDatoSQL1.VisibleImprimir = true;
			this.barraDatoSQL1.VisibleNuevo = true;
			this.barraDatoSQL1.Refresca += new System.EventHandler(this.barraDatoSQL1_Refresca);
			// 
			// ultraNumericEditor7
			// 
			this.ultraNumericEditor7.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsCompra, "Articulo.CostoUltimo"));
			this.ultraNumericEditor7.FormatString = "#,##0.00";
			this.ultraNumericEditor7.Location = new System.Drawing.Point(122, 233);
			this.ultraNumericEditor7.Name = "ultraNumericEditor7";
			this.ultraNumericEditor7.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.ultraNumericEditor7.PromptChar = ' ';
			this.ultraNumericEditor7.Size = new System.Drawing.Size(100, 22);
			this.ultraNumericEditor7.TabIndex = 18;
			// 
			// ultraLabel7
			// 
			this.ultraLabel7.AutoSize = true;
			this.ultraLabel7.Location = new System.Drawing.Point(34, 198);
			this.ultraLabel7.Name = "ultraLabel7";
			this.ultraLabel7.Size = new System.Drawing.Size(65, 15);
			this.ultraLabel7.TabIndex = 17;
			this.ultraLabel7.Text = "Costo Prom.";
			// 
			// CompraArt
			// 
			this.AcceptButton = this.btAceptar;
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 14);
			this.ClientSize = new System.Drawing.Size(328, 350);
			this.Controls.Add(this.ultraNumericEditor7);
			this.Controls.Add(this.ultraLabel7);
			this.Controls.Add(this.barraDatoSQL1);
			this.Controls.Add(this.btAceptar);
			this.Controls.Add(this.ultraNumericEditor6);
			this.Controls.Add(this.ultraLabel6);
			this.Controls.Add(this.ultraNumericEditor5);
			this.Controls.Add(this.ultraLabel5);
			this.Controls.Add(this.ultraNumericEditor4);
			this.Controls.Add(this.ultraLabel4);
			this.Controls.Add(this.ultraNumericEditor3);
			this.Controls.Add(this.ultraLabel3);
			this.Controls.Add(this.ultraNumericEditor2);
			this.Controls.Add(this.ultraLabel2);
			this.Controls.Add(this.ultraNumericEditor1);
			this.Controls.Add(this.ultraLabel1);
			this.MaximizeBox = false;
			this.MinimizeBox = false;
			this.Name = "CompraArt";
			this.Text = "Precio Item";
			this.Load += new System.EventHandler(this.CompraArt_Load);
			((System.ComponentModel.ISupportInitialize)(this.ultraNumericEditor1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsCompra)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraNumericEditor2)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraNumericEditor3)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraNumericEditor4)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraNumericEditor5)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraNumericEditor6)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraNumericEditor7)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void barraDatoSQL1_Refresca(object sender, System.EventArgs e)
		{
		
		}

		private void CompraArt_Load(object sender, System.EventArgs e)
		{
			barraDatoSQL1.Width = 0;
			if (IdArticulo > 0)
			{
				barraDatoSQL1.IdRegistro = IdArticulo;
				barraDatoSQL1.ProximoId(5);
			}
			else
			{
				barraDatoSQL1.ProximoId(4);
			}		
		}

		private void btAceptar_Click(object sender, System.EventArgs e)
		{
			Close();
		}
	}
}
