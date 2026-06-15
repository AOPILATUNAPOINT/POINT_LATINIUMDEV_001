using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;
using System.Data.SqlClient;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de frmDTCreaProductoDeRemate.
	/// </summary>
	public class frmDTCreaProductoDeRemate : System.Windows.Forms.Form
	{
		private C1.Data.C1DataSet cdsSeteoF;
		private System.Windows.Forms.Button btnBuscar;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtCodigo;
		private System.Windows.Forms.Label label1;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtSerial;
		private System.Windows.Forms.Label label72;
		private System.Windows.Forms.Label lblBodega;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbBodega;
		private System.Windows.Forms.Label label2;
		private System.Windows.Forms.Button btnSalir;
		private System.Windows.Forms.Button btnGuardar;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtIdSerial;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtIdArticulo;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtCodigoRemates;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtIdArticuloRemates;
		private System.Windows.Forms.GroupBox groupBox1;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public frmDTCreaProductoDeRemate()
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
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bodega");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo", 0);
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.btnBuscar = new System.Windows.Forms.Button();
			this.txtCodigo = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label1 = new System.Windows.Forms.Label();
			this.txtSerial = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label72 = new System.Windows.Forms.Label();
			this.lblBodega = new System.Windows.Forms.Label();
			this.cmbBodega = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.txtCodigoRemates = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label2 = new System.Windows.Forms.Label();
			this.btnSalir = new System.Windows.Forms.Button();
			this.btnGuardar = new System.Windows.Forms.Button();
			this.txtIdSerial = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtIdArticulo = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtIdArticuloRemates = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.groupBox1 = new System.Windows.Forms.GroupBox();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCodigo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtSerial)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbBodega)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCodigoRemates)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdSerial)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdArticulo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdArticuloRemates)).BeginInit();
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
			this.cdsSeteoF.BeforeFill += new C1.Data.FillEventHandler(this.cdsSeteoF_BeforeFill);
			// 
			// btnBuscar
			// 
			this.btnBuscar.Location = new System.Drawing.Point(360, 8);
			this.btnBuscar.Name = "btnBuscar";
			this.btnBuscar.TabIndex = 169;
			this.btnBuscar.Text = "&Buscar";
			this.btnBuscar.Click += new System.EventHandler(this.btnBuscar_Click);
			// 
			// txtCodigo
			// 
			this.txtCodigo.AcceptsTab = true;
			appearance1.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtCodigo.Appearance = appearance1;
			this.txtCodigo.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtCodigo.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtCodigo.Enabled = false;
			this.txtCodigo.Location = new System.Drawing.Point(96, 72);
			this.txtCodigo.Name = "txtCodigo";
			this.txtCodigo.Size = new System.Drawing.Size(336, 22);
			this.txtCodigo.TabIndex = 170;
			// 
			// label1
			// 
			this.label1.AutoSize = true;
			this.label1.BackColor = System.Drawing.Color.Transparent;
			this.label1.Location = new System.Drawing.Point(8, 72);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(73, 17);
			this.label1.TabIndex = 173;
			this.label1.Text = "Codigo Actual";
			// 
			// txtSerial
			// 
			this.txtSerial.AcceptsTab = true;
			appearance2.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtSerial.Appearance = appearance2;
			this.txtSerial.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtSerial.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtSerial.Location = new System.Drawing.Point(96, 8);
			this.txtSerial.Name = "txtSerial";
			this.txtSerial.Size = new System.Drawing.Size(256, 22);
			this.txtSerial.TabIndex = 168;
			this.txtSerial.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtSerial_KeyDown);
			// 
			// label72
			// 
			this.label72.AutoSize = true;
			this.label72.BackColor = System.Drawing.Color.Transparent;
			this.label72.Location = new System.Drawing.Point(8, 8);
			this.label72.Name = "label72";
			this.label72.Size = new System.Drawing.Size(32, 17);
			this.label72.TabIndex = 172;
			this.label72.Text = "Serial";
			// 
			// lblBodega
			// 
			this.lblBodega.AutoSize = true;
			this.lblBodega.Location = new System.Drawing.Point(8, 40);
			this.lblBodega.Name = "lblBodega";
			this.lblBodega.Size = new System.Drawing.Size(41, 17);
			this.lblBodega.TabIndex = 171;
			this.lblBodega.Text = "Bodega";
			this.lblBodega.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// cmbBodega
			// 
			this.cmbBodega.CausesValidation = false;
			this.cmbBodega.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbBodega.Cursor = System.Windows.Forms.Cursors.Default;
			ultraGridColumn1.Header.VisiblePosition = 2;
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridColumn2.Width = 180;
			ultraGridColumn3.Header.VisiblePosition = 0;
			ultraGridBand1.Columns.AddRange(new object[] {
																										 ultraGridColumn1,
																										 ultraGridColumn2,
																										 ultraGridColumn3});
			this.cmbBodega.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.cmbBodega.DisplayMember = "Nombre";
			this.cmbBodega.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbBodega.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbBodega.Enabled = false;
			this.cmbBodega.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbBodega.Location = new System.Drawing.Point(96, 40);
			this.cmbBodega.MaxDropDownItems = 30;
			this.cmbBodega.Name = "cmbBodega";
			this.cmbBodega.Size = new System.Drawing.Size(256, 21);
			this.cmbBodega.TabIndex = 167;
			this.cmbBodega.ValueMember = "Bodega";
			// 
			// txtCodigoRemates
			// 
			this.txtCodigoRemates.AcceptsTab = true;
			appearance3.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtCodigoRemates.Appearance = appearance3;
			this.txtCodigoRemates.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtCodigoRemates.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtCodigoRemates.Enabled = false;
			this.txtCodigoRemates.Location = new System.Drawing.Point(96, 104);
			this.txtCodigoRemates.Name = "txtCodigoRemates";
			this.txtCodigoRemates.Size = new System.Drawing.Size(336, 22);
			this.txtCodigoRemates.TabIndex = 174;
			// 
			// label2
			// 
			this.label2.AutoSize = true;
			this.label2.BackColor = System.Drawing.Color.Transparent;
			this.label2.Location = new System.Drawing.Point(8, 104);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(74, 17);
			this.label2.TabIndex = 175;
			this.label2.Text = "Codigo Nuevo";
			// 
			// btnSalir
			// 
			this.btnSalir.Location = new System.Drawing.Point(88, 144);
			this.btnSalir.Name = "btnSalir";
			this.btnSalir.TabIndex = 177;
			this.btnSalir.Text = "&Salir";
			this.btnSalir.Click += new System.EventHandler(this.btnSalir_Click);
			// 
			// btnGuardar
			// 
			this.btnGuardar.Enabled = false;
			this.btnGuardar.Location = new System.Drawing.Point(8, 144);
			this.btnGuardar.Name = "btnGuardar";
			this.btnGuardar.TabIndex = 176;
			this.btnGuardar.Text = "&Modificar";
			this.btnGuardar.Click += new System.EventHandler(this.btnGuardar_Click);
			// 
			// txtIdSerial
			// 
			this.txtIdSerial.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.txtIdSerial.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtIdSerial.Enabled = false;
			this.txtIdSerial.FormatString = "#,##0.00";
			this.txtIdSerial.Location = new System.Drawing.Point(384, 144);
			this.txtIdSerial.Name = "txtIdSerial";
			this.txtIdSerial.PromptChar = ' ';
			this.txtIdSerial.Size = new System.Drawing.Size(16, 22);
			this.txtIdSerial.TabIndex = 304;
			this.txtIdSerial.Visible = false;
			// 
			// txtIdArticulo
			// 
			this.txtIdArticulo.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.txtIdArticulo.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtIdArticulo.Enabled = false;
			this.txtIdArticulo.FormatString = "#,##0.00";
			this.txtIdArticulo.Location = new System.Drawing.Point(400, 144);
			this.txtIdArticulo.Name = "txtIdArticulo";
			this.txtIdArticulo.PromptChar = ' ';
			this.txtIdArticulo.Size = new System.Drawing.Size(16, 22);
			this.txtIdArticulo.TabIndex = 303;
			this.txtIdArticulo.Visible = false;
			// 
			// txtIdArticuloRemates
			// 
			this.txtIdArticuloRemates.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.txtIdArticuloRemates.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtIdArticuloRemates.Enabled = false;
			this.txtIdArticuloRemates.FormatString = "#,##0.00";
			this.txtIdArticuloRemates.Location = new System.Drawing.Point(416, 144);
			this.txtIdArticuloRemates.Name = "txtIdArticuloRemates";
			this.txtIdArticuloRemates.PromptChar = ' ';
			this.txtIdArticuloRemates.Size = new System.Drawing.Size(16, 22);
			this.txtIdArticuloRemates.TabIndex = 305;
			this.txtIdArticuloRemates.Visible = false;
			// 
			// groupBox1
			// 
			this.groupBox1.Location = new System.Drawing.Point(0, 128);
			this.groupBox1.Name = "groupBox1";
			this.groupBox1.Size = new System.Drawing.Size(448, 8);
			this.groupBox1.TabIndex = 306;
			this.groupBox1.TabStop = false;
			// 
			// frmDTCreaProductoDeRemate
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 14);
			this.ClientSize = new System.Drawing.Size(448, 174);
			this.Controls.Add(this.groupBox1);
			this.Controls.Add(this.txtIdArticuloRemates);
			this.Controls.Add(this.txtIdSerial);
			this.Controls.Add(this.txtIdArticulo);
			this.Controls.Add(this.btnSalir);
			this.Controls.Add(this.btnGuardar);
			this.Controls.Add(this.txtCodigoRemates);
			this.Controls.Add(this.label2);
			this.Controls.Add(this.btnBuscar);
			this.Controls.Add(this.txtCodigo);
			this.Controls.Add(this.label1);
			this.Controls.Add(this.txtSerial);
			this.Controls.Add(this.label72);
			this.Controls.Add(this.lblBodega);
			this.Controls.Add(this.cmbBodega);
			this.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
			this.KeyPreview = true;
			this.MaximizeBox = false;
			this.Name = "frmDTCreaProductoDeRemate";
			this.Text = "Ingreso de Productos de Remate";
			this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.frmDTCreaProductoDeRemate_KeyDown);
			this.Load += new System.EventHandler(this.frmDTCreaProductoDeRemate_Load);
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCodigo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtSerial)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbBodega)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCodigoRemates)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdSerial)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdArticulo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdArticuloRemates)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void frmDTCreaProductoDeRemate_Load(object sender, System.EventArgs e)
		{
			string sSQL = string.Format("Exec ListaBodegas {0}, 47", MenuLatinium.IdUsuario);
			this.cmbBodega.DataSource = Funcion.dvProcedimiento(cdsSeteoF, sSQL);
		}

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);
		}

		private void btnBuscar_Click(object sender, System.EventArgs e)
		{
			try
			{
				Limpiar(false);

				if (this.txtSerial.Text.ToString().Length == 0)
				{
					MessageBox.Show("Ingrese una Serial", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					this.txtSerial.Focus();
					return;
				}
			
				SqlDataReader dr = Funcion.rEscalarSQL(cdsSeteoF, string.Format("Exec BuscaSerialRemates '{0}'", this.txtSerial.Text.ToString()), true);
				dr.Read();
				if (dr.HasRows)
				{
					this.txtIdSerial.Value = dr.GetInt32(0);
					this.txtIdArticulo.Value = dr.GetInt32(1);
					this.txtCodigo.Text = dr.GetString(2);
					this.cmbBodega.Value = dr.GetInt32(4);

					dr.Close();

					this.txtCodigoRemates.Text = "REM-" + this.txtCodigo.Text.ToString().Trim();

					this.txtIdArticuloRemates.Value = Funcion.iEscalarSQL(cdsSeteoF, string.Format("Exec ValidaExisteArticuloRemates '{0}'", this.txtCodigoRemates.Text.ToString().Trim()));

					if (this.txtCodigo.Text.ToString().Substring(0, 3) == "REM")
					{
						MessageBox.Show(string.Format("La serial '{0}' ya pertenece a un producto de remate '{1}'", this.txtSerial.Text.ToString(), this.txtCodigo.Text.ToString()), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
						this.Limpiar(true);
						return;
					}

					if (this.txtCodigo.Text.ToString().Substring(0, 4) == "RECU")
					{
						MessageBox.Show(string.Format("La serial '{0}' pertenece a un producto de recuperación de crédito '{1}'", this.txtSerial.Text.ToString(), this.txtCodigo.Text.ToString()), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
						this.Limpiar(true);
						return;
					}

					this.btnGuardar.Enabled = true;
					this.btnGuardar.Focus();
				}
				else
				{
					dr.Close();
					MessageBox.Show(string.Format("La serial '{0}' no existe", this.txtSerial.Text.ToString()), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				}	
			}
			catch (Exception ex)
			{
				MessageBox.Show(string.Format("Error del Sistema : {0} {1}", ex.GetType(), ex.Message), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);				
			}			
		}

		private void Limpiar(bool bGuarda)
		{
			this.txtIdSerial.Value = 0;
			this.txtIdArticulo.Value = 0;
			this.txtIdArticuloRemates.Value = 0;
			this.cmbBodega.Value = System.DBNull.Value;
			if (bGuarda) this.txtSerial.Text = "";
			this.txtCodigo.Text = "";
			this.txtCodigoRemates.Text = "";
			this.btnGuardar.Enabled = false;
		}

		private void btnGuardar_Click(object sender, System.EventArgs e)
		{
			if (DialogResult.Yes == MessageBox.Show("¿Esta seguro de crear codigo de remates de este articulo?", "Point Technology", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2)) 
			{
				try
				{
					if ((int)this.txtIdArticuloRemates.Value == 0)
					{
						string sSQL = string.Format("Exec CreaArticuloRemates {0}, '{1}'", (int)this.txtIdArticulo.Value, this.txtCodigoRemates.Text.ToString());
						this.txtIdArticuloRemates.Value = Funcion.iEscalarSQL(cdsSeteoF, sSQL);			
					}

					decimal dCosto = Funcion.decEscalarSQL(cdsSeteoF, string.Format("Select dbo.CalculaUltimoCosto({0}, '{1}')", (int)this.txtIdArticulo.Value, DateTime.Today.ToString("yyyyMMdd")));

					string sSQLEB = string.Format("Exec CreaIEBodegaRemates {0}, {1}, {2}, 8, {3}", (int)this.txtIdArticulo.Value, (int)this.txtIdSerial.Value, dCosto, (int)this.cmbBodega.Value);
					int idCompraEB = Funcion.iEscalarSQL(cdsSeteoF, sSQLEB);
											
					#region Imprime el reporte de Egreso de Bodega
					using (Reporte miReporte = new Reporte("EgresosDeBodega.Rpt", "{Compra.idCompra} = " + idCompraEB.ToString()))
					{						
						miReporte.ShowDialog();				
					}
					#endregion Imprime el reporte de Egreso de Bodega

					string sSQLIB = string.Format("Exec CreaIEBodegaRemates {0}, {1}, {2}, 9, {3}", (int)this.txtIdArticuloRemates.Value, (int)this.txtIdSerial.Value, dCosto, (int)this.cmbBodega.Value);
					int idCompraIB = Funcion.iEscalarSQL(cdsSeteoF, sSQLIB);
				
					#region Imprime el Reporte de Ingreso de Bodega
					using (Reporte miReporte = new Reporte("IngresoBodega.Rpt", "{Compra.idCompra} = " + idCompraIB.ToString()))
					{						
						miReporte.ShowDialog();				
					}
					#endregion Imprime el Reporte de Ingreso de Bodega				

					MessageBox.Show("Articulo Modificado a Producto de Remate", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);

					Limpiar(true);

					this.txtSerial.Focus();
				}
				catch (Exception ex)
				{
					MessageBox.Show(string.Format("Error del Sistema : {0} {1}", ex.GetType(), ex.Message), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);				
				}
			}
		}

		private void txtSerial_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.btnBuscar_Click(sender, e);
		}

		private void btnSalir_Click(object sender, System.EventArgs e)
		{
			this.Close();
		}

		private void frmDTCreaProductoDeRemate_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Escape) this.Close();
		}
	}
}
