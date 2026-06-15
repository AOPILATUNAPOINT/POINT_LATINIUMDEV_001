using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de CambiarPrecio.
	/// </summary>
	public class CambiarPrecio : DevExpress.XtraEditors.XtraForm
	{
		private System.Windows.Forms.Label lblDescripcion;
		private Infragistics.Win.Misc.UltraButton btAceptar;
		private Infragistics.Win.Misc.UltraButton btCancelar;
		public Infragistics.Win.UltraWinEditors.UltraNumericEditor txtPrecioPer;
		private C1.Data.C1DataSet cdsArticulo;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtPrecio1;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtPrecio2;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtPrecio3;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtPrecio4;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtPrecio5;
		private Infragistics.Win.UltraWinEditors.UltraOptionSet optPrecio;
		public Infragistics.Win.UltraWinEditors.UltraNumericEditor txtPrecio;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public CambiarPrecio()
		{
			InitializeComponent();
		}
		private string stArticulo = "";
		public CambiarPrecio(string stArt)
		{
			InitializeComponent();
			stArticulo = stArt;
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
			Infragistics.Win.Appearance appearance1 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.ValueListItem valueListItem1 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem2 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem3 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem4 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem5 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem6 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			this.txtPrecioPer = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.lblDescripcion = new System.Windows.Forms.Label();
			this.btAceptar = new Infragistics.Win.Misc.UltraButton();
			this.btCancelar = new Infragistics.Win.Misc.UltraButton();
			this.cdsArticulo = new C1.Data.C1DataSet();
			this.txtPrecio1 = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtPrecio2 = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtPrecio3 = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtPrecio4 = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtPrecio5 = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.optPrecio = new Infragistics.Win.UltraWinEditors.UltraOptionSet();
			this.txtPrecio = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			((System.ComponentModel.ISupportInitialize)(this.txtPrecioPer)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsArticulo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtPrecio1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtPrecio2)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtPrecio3)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtPrecio4)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtPrecio5)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.optPrecio)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtPrecio)).BeginInit();
			this.SuspendLayout();
			// 
			// txtPrecioPer
			// 
			appearance1.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtPrecioPer.Appearance = appearance1;
			this.txtPrecioPer.FormatString = "#,##0.00";
			this.txtPrecioPer.Location = new System.Drawing.Point(160, 192);
			this.txtPrecioPer.Name = "txtPrecioPer";
			this.txtPrecioPer.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtPrecioPer.PromptChar = ' ';
			this.txtPrecioPer.Size = new System.Drawing.Size(100, 21);
			this.txtPrecioPer.TabIndex = 0;
			// 
			// lblDescripcion
			// 
			this.lblDescripcion.Location = new System.Drawing.Point(24, 24);
			this.lblDescripcion.Name = "lblDescripcion";
			this.lblDescripcion.Size = new System.Drawing.Size(240, 32);
			this.lblDescripcion.TabIndex = 1;
			this.lblDescripcion.Text = "Descripcion";
			// 
			// btAceptar
			// 
			this.btAceptar.Location = new System.Drawing.Point(40, 232);
			this.btAceptar.Name = "btAceptar";
			this.btAceptar.TabIndex = 4;
			this.btAceptar.Text = "&Aceptar";
			this.btAceptar.Click += new System.EventHandler(this.btAceptar_Click);
			// 
			// btCancelar
			// 
			this.btCancelar.DialogResult = System.Windows.Forms.DialogResult.Cancel;
			this.btCancelar.Location = new System.Drawing.Point(160, 232);
			this.btCancelar.Name = "btCancelar";
			this.btCancelar.TabIndex = 5;
			this.btCancelar.Text = "&Cancelar";
			// 
			// cdsArticulo
			// 
			this.cdsArticulo.BindingContextCtrl = this;
			this.cdsArticulo.DataLibrary = "LibFacturacion";
			this.cdsArticulo.DataLibraryUrl = "";
			this.cdsArticulo.DataSetDef = "dsArticulo";
			this.cdsArticulo.FillOnRequest = false;
			this.cdsArticulo.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsArticulo.Name = "cdsArticulo";
			this.cdsArticulo.SchemaClassName = "LibFacturacion.DataClass";
			this.cdsArticulo.SchemaDef = null;
			// 
			// txtPrecio1
			// 
			appearance2.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtPrecio1.Appearance = appearance2;
			this.txtPrecio1.FormatString = "#,##0.00";
			this.txtPrecio1.Location = new System.Drawing.Point(160, 72);
			this.txtPrecio1.Name = "txtPrecio1";
			this.txtPrecio1.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtPrecio1.PromptChar = ' ';
			this.txtPrecio1.Size = new System.Drawing.Size(100, 21);
			this.txtPrecio1.TabIndex = 6;
			// 
			// txtPrecio2
			// 
			appearance3.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtPrecio2.Appearance = appearance3;
			this.txtPrecio2.FormatString = "#,##0.00";
			this.txtPrecio2.Location = new System.Drawing.Point(160, 96);
			this.txtPrecio2.Name = "txtPrecio2";
			this.txtPrecio2.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtPrecio2.PromptChar = ' ';
			this.txtPrecio2.Size = new System.Drawing.Size(100, 21);
			this.txtPrecio2.TabIndex = 7;
			// 
			// txtPrecio3
			// 
			appearance4.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtPrecio3.Appearance = appearance4;
			this.txtPrecio3.FormatString = "#,##0.00";
			this.txtPrecio3.Location = new System.Drawing.Point(160, 120);
			this.txtPrecio3.Name = "txtPrecio3";
			this.txtPrecio3.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtPrecio3.PromptChar = ' ';
			this.txtPrecio3.Size = new System.Drawing.Size(100, 21);
			this.txtPrecio3.TabIndex = 8;
			// 
			// txtPrecio4
			// 
			appearance5.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtPrecio4.Appearance = appearance5;
			this.txtPrecio4.FormatString = "#,##0.00";
			this.txtPrecio4.Location = new System.Drawing.Point(160, 144);
			this.txtPrecio4.Name = "txtPrecio4";
			this.txtPrecio4.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtPrecio4.PromptChar = ' ';
			this.txtPrecio4.Size = new System.Drawing.Size(100, 21);
			this.txtPrecio4.TabIndex = 9;
			// 
			// txtPrecio5
			// 
			appearance6.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtPrecio5.Appearance = appearance6;
			this.txtPrecio5.FormatString = "#,##0.00";
			this.txtPrecio5.Location = new System.Drawing.Point(160, 168);
			this.txtPrecio5.Name = "txtPrecio5";
			this.txtPrecio5.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtPrecio5.PromptChar = ' ';
			this.txtPrecio5.Size = new System.Drawing.Size(100, 21);
			this.txtPrecio5.TabIndex = 10;
			// 
			// optPrecio
			// 
			appearance7.ForeColorDisabled = System.Drawing.Color.Black;
			this.optPrecio.Appearance = appearance7;
			this.optPrecio.ItemAppearance = appearance8;
			this.optPrecio.ItemOrigin = new System.Drawing.Point(6, 6);
			valueListItem1.DataValue = 0;
			valueListItem1.DisplayText = "Precio 1";
			valueListItem2.DataValue = 1;
			valueListItem2.DisplayText = "Precio 2";
			valueListItem3.DataValue = 2;
			valueListItem3.DisplayText = "Precio 3";
			valueListItem4.DataValue = 3;
			valueListItem4.DisplayText = "Precio 4";
			valueListItem5.DataValue = 4;
			valueListItem5.DisplayText = "Precio 5";
			valueListItem6.DataValue = 5;
			valueListItem6.DisplayText = "Personalizado";
			this.optPrecio.Items.Add(valueListItem1);
			this.optPrecio.Items.Add(valueListItem2);
			this.optPrecio.Items.Add(valueListItem3);
			this.optPrecio.Items.Add(valueListItem4);
			this.optPrecio.Items.Add(valueListItem5);
			this.optPrecio.Items.Add(valueListItem6);
			this.optPrecio.ItemSpacingVertical = 8;
			this.optPrecio.Location = new System.Drawing.Point(24, 72);
			this.optPrecio.Name = "optPrecio";
			this.optPrecio.Size = new System.Drawing.Size(112, 144);
			this.optPrecio.TabIndex = 11;
			this.optPrecio.ValueChanged += new System.EventHandler(this.optPrecio_ValueChanged);
			// 
			// txtPrecio
			// 
			appearance9.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtPrecio.Appearance = appearance9;
			this.txtPrecio.FormatString = "#,##0.00";
			this.txtPrecio.Location = new System.Drawing.Point(128, 0);
			this.txtPrecio.Name = "txtPrecio";
			this.txtPrecio.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtPrecio.PromptChar = ' ';
			this.txtPrecio.Size = new System.Drawing.Size(100, 21);
			this.txtPrecio.TabIndex = 12;
			// 
			// CambiarPrecio
			// 
			this.AcceptButton = this.btAceptar;
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.CancelButton = this.btCancelar;
			this.ClientSize = new System.Drawing.Size(280, 270);
			this.Controls.Add(this.txtPrecio);
			this.Controls.Add(this.optPrecio);
			this.Controls.Add(this.txtPrecio5);
			this.Controls.Add(this.txtPrecio4);
			this.Controls.Add(this.txtPrecio3);
			this.Controls.Add(this.txtPrecio2);
			this.Controls.Add(this.txtPrecio1);
			this.Controls.Add(this.btCancelar);
			this.Controls.Add(this.btAceptar);
			this.Controls.Add(this.lblDescripcion);
			this.Controls.Add(this.txtPrecioPer);
			this.Name = "CambiarPrecio";
			this.Text = "Cambio de Precios";
			this.Load += new System.EventHandler(this.CambiarPrecio_Load);
			((System.ComponentModel.ISupportInitialize)(this.txtPrecioPer)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsArticulo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtPrecio1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtPrecio2)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtPrecio3)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtPrecio4)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtPrecio5)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.optPrecio)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtPrecio)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void CambiarPrecio_Load(object sender, System.EventArgs e)
		{
			txtPrecio.Width= 0;
			lblDescripcion.Text = stArticulo;
			optPrecio.Value = 5;
		}

		private void btAceptar_Click(object sender, System.EventArgs e)
		{
			switch ((int) optPrecio.Value + 1)
			{
				case 1:
					txtPrecio.Value = txtPrecio1.Value;
					break;
				case 2:
					txtPrecio.Value = txtPrecio2.Value;
					break;
				case 3:
					txtPrecio.Value = txtPrecio3.Value;
					break;
				case 4:
					txtPrecio.Value = txtPrecio4.Value;
					break;
				case 5:
					txtPrecio.Value = txtPrecio5.Value;
					break;
				case 6:
					txtPrecio.Value = txtPrecioPer.Value;
					break;
			}

			DialogResult = DialogResult.OK;
		}

		private void optPrecio_ValueChanged(object sender, System.EventArgs e)
		{
			txtPrecio1.Enabled = false;
			txtPrecio2.Enabled = false;
			txtPrecio3.Enabled = false;
			txtPrecio4.Enabled = false;
			txtPrecio5.Enabled = false;
			txtPrecioPer.Enabled = false;
			switch ((int) optPrecio.Value + 1)
			{
				case 1:
					txtPrecio1.Enabled = true;
					txtPrecio1.Select();
					break;
				case 2:
					txtPrecio2.Enabled = true;
					txtPrecio2.Select();
					break;
				case 3:
					txtPrecio3.Enabled = true;
					txtPrecio3.Select();
					break;
				case 4:
					txtPrecio4.Enabled = true;
					txtPrecio4.Select();
					break;
				case 5:
					txtPrecio5.Enabled = true;
					txtPrecio5.Select();
					break;
				case 6:
					txtPrecioPer.Enabled = true;
					txtPrecioPer.Select();
					break;
			}
		}
	}
}
