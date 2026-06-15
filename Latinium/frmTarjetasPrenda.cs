using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de frmTarjetasPrenda.
	/// </summary>
	public class frmTarjetasPrenda : System.Windows.Forms.Form
	{
		private System.Windows.Forms.Label lblCuenta;
		private System.Windows.Forms.Label lblBanco;
		public Infragistics.Win.UltraWinGrid.UltraCombo cmbTarjeta;
		public Infragistics.Win.UltraWinGrid.UltraCombo cmbBanco;
		private System.Windows.Forms.Button btCancelar;
		private System.Windows.Forms.Button btGrabar;
		private C1.Data.C1DataSet cdsSeteoF;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public frmTarjetasPrenda()
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
			Infragistics.Win.Appearance appearance1 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idBanco");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmTarjetasPrenda));
			this.lblCuenta = new System.Windows.Forms.Label();
			this.lblBanco = new System.Windows.Forms.Label();
			this.cmbTarjeta = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.cmbBanco = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.btCancelar = new System.Windows.Forms.Button();
			this.btGrabar = new System.Windows.Forms.Button();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			((System.ComponentModel.ISupportInitialize)(this.cmbTarjeta)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbBanco)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			this.SuspendLayout();
			// 
			// lblCuenta
			// 
			this.lblCuenta.AutoSize = true;
			this.lblCuenta.BackColor = System.Drawing.Color.Transparent;
			this.lblCuenta.Location = new System.Drawing.Point(8, 32);
			this.lblCuenta.Name = "lblCuenta";
			this.lblCuenta.Size = new System.Drawing.Size(39, 16);
			this.lblCuenta.TabIndex = 173;
			this.lblCuenta.Text = "Tarjeta";
			// 
			// lblBanco
			// 
			this.lblBanco.AutoSize = true;
			this.lblBanco.BackColor = System.Drawing.Color.Transparent;
			this.lblBanco.Location = new System.Drawing.Point(8, 8);
			this.lblBanco.Name = "lblBanco";
			this.lblBanco.Size = new System.Drawing.Size(36, 16);
			this.lblBanco.TabIndex = 172;
			this.lblBanco.Text = "Banco";
			// 
			// cmbTarjeta
			// 
			appearance1.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbTarjeta.Appearance = appearance1;
			this.cmbTarjeta.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbTarjeta.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbTarjeta.DisplayMember = "Nombre";
			this.cmbTarjeta.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbTarjeta.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbTarjeta.Location = new System.Drawing.Point(64, 32);
			this.cmbTarjeta.Name = "cmbTarjeta";
			this.cmbTarjeta.Size = new System.Drawing.Size(224, 21);
			this.cmbTarjeta.TabIndex = 170;
			this.cmbTarjeta.ValueMember = "idTarjeta";
			// 
			// cmbBanco
			// 
			appearance2.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbBanco.Appearance = appearance2;
			this.cmbBanco.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.cmbBanco.Cursor = System.Windows.Forms.Cursors.Default;
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridColumn2.Width = 224;
			ultraGridBand1.Columns.AddRange(new object[] {
																										 ultraGridColumn1,
																										 ultraGridColumn2});
			this.cmbBanco.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.cmbBanco.DisplayMember = "Nombre";
			this.cmbBanco.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbBanco.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbBanco.Location = new System.Drawing.Point(64, 8);
			this.cmbBanco.Name = "cmbBanco";
			this.cmbBanco.Size = new System.Drawing.Size(224, 21);
			this.cmbBanco.TabIndex = 169;
			this.cmbBanco.ValueMember = "idBanco";
			this.cmbBanco.ValueChanged += new System.EventHandler(this.cmbBanco_ValueChanged);
			// 
			// btCancelar
			// 
			this.btCancelar.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.btCancelar.Image = ((System.Drawing.Image)(resources.GetObject("btCancelar.Image")));
			this.btCancelar.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btCancelar.Location = new System.Drawing.Point(104, 66);
			this.btCancelar.Name = "btCancelar";
			this.btCancelar.Size = new System.Drawing.Size(85, 23);
			this.btCancelar.TabIndex = 177;
			this.btCancelar.Text = "&Cancelar";
			this.btCancelar.Click += new System.EventHandler(this.btCancelar_Click);
			// 
			// btGrabar
			// 
			this.btGrabar.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.btGrabar.Image = ((System.Drawing.Image)(resources.GetObject("btGrabar.Image")));
			this.btGrabar.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btGrabar.Location = new System.Drawing.Point(8, 65);
			this.btGrabar.Name = "btGrabar";
			this.btGrabar.Size = new System.Drawing.Size(85, 24);
			this.btGrabar.TabIndex = 176;
			this.btGrabar.Text = "&Grabar";
			this.btGrabar.Click += new System.EventHandler(this.btGrabar_Click);
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
			this.cdsSeteoF.BeforeFill += new C1.Data.FillEventHandler(this.cdsSeteoF_BeforeFill);
			// 
			// frmTarjetasPrenda
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(298, 96);
			this.Controls.Add(this.btCancelar);
			this.Controls.Add(this.btGrabar);
			this.Controls.Add(this.lblCuenta);
			this.Controls.Add(this.lblBanco);
			this.Controls.Add(this.cmbTarjeta);
			this.Controls.Add(this.cmbBanco);
			this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
			this.MaximizeBox = false;
			this.MinimizeBox = false;
			this.Name = "frmTarjetasPrenda";
			this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
			this.Text = "Tarjeta";
			this.Load += new System.EventHandler(this.frmTarjetasPrenda_Load);
			((System.ComponentModel.ISupportInitialize)(this.cmbTarjeta)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbBanco)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);
		}

		private void btCancelar_Click(object sender, System.EventArgs e)
		{
			this.DialogResult = DialogResult.No;
		}

		private void frmTarjetasPrenda_Load(object sender, System.EventArgs e)
		{
			this.cmbBanco.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Select idBanco, Nombre From Banco Where Caja = 1 Order By Nombre");
		}

		private void cmbBanco_ValueChanged(object sender, System.EventArgs e)
		{
			if (this.cmbBanco.ActiveRow != null)
			{
				this.cmbTarjeta.Value = System.DBNull.Value;
				
				string sSQL = string.Format("Exec BTPP {0}, 0, 0, 0", (int)this.cmbBanco.Value);
				this.cmbTarjeta.DataSource = Funcion.dvProcedimiento(cdsSeteoF, sSQL);
				this.cmbTarjeta.DisplayLayout.Bands[0].Columns["idTarjeta"].Hidden = true;
				this.cmbTarjeta.DisplayLayout.Bands[0].Columns["Nombre"].Width = 224;
			}
		}

		private void btGrabar_Click(object sender, System.EventArgs e)
		{
			#region Validacion
			if (this.cmbBanco.ActiveRow == null)
			{
				MessageBox.Show("Seleccione el Banco", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.cmbBanco.Focus();
				return;										
			}
			else if (this.cmbTarjeta.ActiveRow == null)
			{
				MessageBox.Show("Seleccione la Tarjeta", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.cmbTarjeta.Focus();
				return;										
			}
			else
			{
				this.DialogResult = DialogResult.OK;
			}
			#endregion Validacion
		}
	}
}
