using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de frmInfoTarjetasDeCreditoDB.
	/// </summary>
	public class frmInfoTarjetasDeCreditoDB : System.Windows.Forms.Form
	{
		public System.Windows.Forms.Label lblfecha;
		public Infragistics.Win.UltraWinGrid.UltraCombo cmbPlanes;
		private System.Windows.Forms.Label lblPlanes;
		private System.Windows.Forms.Label lblTarjeta;
		private System.Windows.Forms.Label lblPlazo;
		private System.Windows.Forms.Label lblBanco;
		public Infragistics.Win.UltraWinGrid.UltraCombo cmbPlazo;
		public Infragistics.Win.UltraWinGrid.UltraCombo cmbTarjeta;
		public Infragistics.Win.UltraWinGrid.UltraCombo cmbBanco;
		private System.Windows.Forms.Button btCancelar;
		private System.Windows.Forms.Button btGrabar;
		private C1.Data.C1DataSet cdsSeteoF;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource1;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public frmInfoTarjetasDeCreditoDB()
		{
			//
			// Necesario para admitir el Diseñador de Windows Forms
			//
			InitializeComponent();

			//
			// TODO: agregar código de constructor después de llamar a InitializeComponent
			//
		}

		int iBodega = 0;
		DateTime dtfecha;
		int idFormaPago = 0;

		public frmInfoTarjetasDeCreditoDB(int IBodega, DateTime dtFecha, int IdFormaPago)
		{
			//
			// Necesario para admitir el Diseñador de Windows Forms
			//
			InitializeComponent();
			iBodega = IBodega;
			dtfecha = dtFecha;
			idFormaPago = IdFormaPago;

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
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idBanco");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idBanco");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			this.lblfecha = new System.Windows.Forms.Label();
			this.cmbPlanes = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.lblPlanes = new System.Windows.Forms.Label();
			this.lblTarjeta = new System.Windows.Forms.Label();
			this.lblPlazo = new System.Windows.Forms.Label();
			this.lblBanco = new System.Windows.Forms.Label();
			this.cmbPlazo = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.cmbTarjeta = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.cmbBanco = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.btCancelar = new System.Windows.Forms.Button();
			this.btGrabar = new System.Windows.Forms.Button();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.ultraDataSource1 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			((System.ComponentModel.ISupportInitialize)(this.cmbPlanes)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbPlazo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbTarjeta)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbBanco)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).BeginInit();
			this.SuspendLayout();
			// 
			// lblfecha
			// 
			this.lblfecha.AutoSize = true;
			this.lblfecha.BackColor = System.Drawing.Color.Transparent;
			this.lblfecha.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
			this.lblfecha.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Bold);
			this.lblfecha.ForeColor = System.Drawing.Color.Firebrick;
			this.lblfecha.Location = new System.Drawing.Point(64, 176);
			this.lblfecha.Name = "lblfecha";
			this.lblfecha.Size = new System.Drawing.Size(0, 20);
			this.lblfecha.TabIndex = 706;
			this.lblfecha.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			this.lblfecha.Visible = false;
			// 
			// cmbPlanes
			// 
			appearance1.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbPlanes.Appearance = appearance1;
			this.cmbPlanes.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbPlanes.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbPlanes.DisplayMember = "Descripcion";
			this.cmbPlanes.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbPlanes.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbPlanes.Location = new System.Drawing.Point(80, 72);
			this.cmbPlanes.Name = "cmbPlanes";
			this.cmbPlanes.Size = new System.Drawing.Size(224, 22);
			this.cmbPlanes.TabIndex = 700;
			this.cmbPlanes.ValueMember = "idPlan";
			this.cmbPlanes.ValueChanged += new System.EventHandler(this.cmbPlanes_ValueChanged);
			// 
			// lblPlanes
			// 
			this.lblPlanes.AutoSize = true;
			this.lblPlanes.BackColor = System.Drawing.Color.Transparent;
			this.lblPlanes.Location = new System.Drawing.Point(8, 72);
			this.lblPlanes.Name = "lblPlanes";
			this.lblPlanes.Size = new System.Drawing.Size(39, 16);
			this.lblPlanes.TabIndex = 705;
			this.lblPlanes.Text = "Planes";
			this.lblPlanes.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// lblTarjeta
			// 
			this.lblTarjeta.AutoSize = true;
			this.lblTarjeta.BackColor = System.Drawing.Color.Transparent;
			this.lblTarjeta.Location = new System.Drawing.Point(8, 40);
			this.lblTarjeta.Name = "lblTarjeta";
			this.lblTarjeta.Size = new System.Drawing.Size(39, 16);
			this.lblTarjeta.TabIndex = 704;
			this.lblTarjeta.Text = "Tarjeta";
			this.lblTarjeta.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// lblPlazo
			// 
			this.lblPlazo.AutoSize = true;
			this.lblPlazo.BackColor = System.Drawing.Color.Transparent;
			this.lblPlazo.Location = new System.Drawing.Point(8, 104);
			this.lblPlazo.Name = "lblPlazo";
			this.lblPlazo.Size = new System.Drawing.Size(33, 16);
			this.lblPlazo.TabIndex = 703;
			this.lblPlazo.Text = "Plazo";
			this.lblPlazo.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// lblBanco
			// 
			this.lblBanco.AutoSize = true;
			this.lblBanco.BackColor = System.Drawing.Color.Transparent;
			this.lblBanco.Location = new System.Drawing.Point(8, 8);
			this.lblBanco.Name = "lblBanco";
			this.lblBanco.Size = new System.Drawing.Size(36, 16);
			this.lblBanco.TabIndex = 702;
			this.lblBanco.Text = "Banco";
			this.lblBanco.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// cmbPlazo
			// 
			appearance2.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbPlazo.Appearance = appearance2;
			this.cmbPlazo.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbPlazo.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbPlazo.DisplayMember = "Descripcion";
			this.cmbPlazo.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbPlazo.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbPlazo.Location = new System.Drawing.Point(80, 104);
			this.cmbPlazo.Name = "cmbPlazo";
			this.cmbPlazo.Size = new System.Drawing.Size(224, 22);
			this.cmbPlazo.TabIndex = 701;
			this.cmbPlazo.ValueMember = "idPlazo";
			this.cmbPlazo.ValueChanged += new System.EventHandler(this.cmbPlazo_ValueChanged);
			// 
			// cmbTarjeta
			// 
			appearance3.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbTarjeta.Appearance = appearance3;
			this.cmbTarjeta.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbTarjeta.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbTarjeta.DisplayMember = "Nombre";
			this.cmbTarjeta.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbTarjeta.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbTarjeta.Location = new System.Drawing.Point(80, 40);
			this.cmbTarjeta.Name = "cmbTarjeta";
			this.cmbTarjeta.Size = new System.Drawing.Size(224, 22);
			this.cmbTarjeta.TabIndex = 699;
			this.cmbTarjeta.ValueMember = "idTarjeta";
			this.cmbTarjeta.ValueChanged += new System.EventHandler(this.cmbTarjeta_ValueChanged);
			// 
			// cmbBanco
			// 
			appearance4.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbBanco.Appearance = appearance4;
			this.cmbBanco.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.cmbBanco.Cursor = System.Windows.Forms.Cursors.Default;
			ultraGridBand1.AddButtonCaption = "Compra";
			ultraGridColumn1.Header.VisiblePosition = 1;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn2.Header.VisiblePosition = 0;
			ultraGridColumn2.Width = 224;
			ultraGridBand1.Columns.AddRange(new object[] {
																										 ultraGridColumn1,
																										 ultraGridColumn2});
			this.cmbBanco.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.cmbBanco.DisplayMember = "Nombre";
			this.cmbBanco.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbBanco.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbBanco.Location = new System.Drawing.Point(80, 8);
			this.cmbBanco.Name = "cmbBanco";
			this.cmbBanco.Size = new System.Drawing.Size(224, 22);
			this.cmbBanco.TabIndex = 698;
			this.cmbBanco.ValueMember = "idBanco";
			this.cmbBanco.ValueChanged += new System.EventHandler(this.cmbBanco_ValueChanged);
			// 
			// btCancelar
			// 
			this.btCancelar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btCancelar.Location = new System.Drawing.Point(112, 144);
			this.btCancelar.Name = "btCancelar";
			this.btCancelar.Size = new System.Drawing.Size(99, 23);
			this.btCancelar.TabIndex = 708;
			this.btCancelar.Text = "&Cancelar";
			this.btCancelar.Click += new System.EventHandler(this.btCancelar_Click);
			// 
			// btGrabar
			// 
			this.btGrabar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btGrabar.Location = new System.Drawing.Point(8, 144);
			this.btGrabar.Name = "btGrabar";
			this.btGrabar.Size = new System.Drawing.Size(99, 23);
			this.btGrabar.TabIndex = 707;
			this.btGrabar.Text = "&Continuar";
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
			// ultraDataSource1
			// 
			ultraDataColumn1.DataType = typeof(int);
			this.ultraDataSource1.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn1,
																																 ultraDataColumn2});
			// 
			// frmInfoTarjetasDeCreditoDB
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(312, 174);
			this.ControlBox = false;
			this.Controls.Add(this.btCancelar);
			this.Controls.Add(this.btGrabar);
			this.Controls.Add(this.lblfecha);
			this.Controls.Add(this.cmbPlanes);
			this.Controls.Add(this.lblPlanes);
			this.Controls.Add(this.lblTarjeta);
			this.Controls.Add(this.lblPlazo);
			this.Controls.Add(this.lblBanco);
			this.Controls.Add(this.cmbPlazo);
			this.Controls.Add(this.cmbTarjeta);
			this.Controls.Add(this.cmbBanco);
			this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
			this.KeyPreview = true;
			this.MaximizeBox = false;
			this.MinimizeBox = false;
			this.Name = "frmInfoTarjetasDeCreditoDB";
			this.ShowInTaskbar = false;
			this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
			this.Text = "Información de Tarjetas De Crédito/Débito";
			this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.frmInfoTarjetasDeCreditoDB_KeyDown);
			this.Load += new System.EventHandler(this.frmInfoTarjetasDeCreditoDB_Load);
			((System.ComponentModel.ISupportInitialize)(this.cmbPlanes)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbPlazo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbTarjeta)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbBanco)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		int iRegion = 0;

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);
		}

		private void btCancelar_Click(object sender, System.EventArgs e)
		{
			this.DialogResult = DialogResult.No;
		}

		private void frmInfoTarjetasDeCreditoDB_Load(object sender, System.EventArgs e)
		{
			iRegion = Funcion.iEscalarSQL(cdsSeteoF, string.Format("Select Region From Bodega Where Bodega = {0}", iBodega));

			string sSQL = string.Format("Exec BTPP 0, 0, 0, 3, {0}, '{1}'", iRegion, Convert.ToDateTime(dtfecha).ToString("yyyyMMdd"));
			this.cmbBanco.DataSource = Funcion.dvProcedimiento(cdsSeteoF, sSQL);
		}

		private void cmbBanco_ValueChanged(object sender, System.EventArgs e)
		{
			if (this.cmbBanco.ActiveRow != null)
			{
				this.cmbTarjeta.Value = System.DBNull.Value;
				this.cmbPlanes.Value = System.DBNull.Value;
				this.cmbPlazo.Value = System.DBNull.Value;
				//@idBanco Int, @idTarjeta Int, @idPlan Int, @Tipo Int, @Bodega Int = 0
				string sSQL = string.Format("Exec BTPP {0}, 0, 0, 0, {1}, '{2}', {3}", 
					(int)this.cmbBanco.Value, iRegion, Convert.ToDateTime(dtfecha).ToString("yyyyMMdd"), idFormaPago);
				this.cmbTarjeta.DataSource = Funcion.dvProcedimiento(cdsSeteoF, sSQL);

				this.cmbTarjeta.DisplayLayout.Bands[0].Columns["idTarjeta"].Hidden = true;
				this.cmbTarjeta.DisplayLayout.Bands[0].Columns["Nombre"].Width = 224;
			}
		}

		private void cmbTarjeta_ValueChanged(object sender, System.EventArgs e)
		{
			if (this.cmbBanco.ActiveRow != null)
			{
				if (this.cmbTarjeta.ActiveRow != null) 
				{
					this.cmbPlanes.Value = System.DBNull.Value;
					this.cmbPlazo.Value = System.DBNull.Value;

					string sSQL = string.Format("Exec BTPP {0}, {1}, 0, 1, {2}, '{3}', {4},{5}", 
						(int)this.cmbBanco.Value, (int)this.cmbTarjeta.Value, iRegion, Convert.ToDateTime(dtfecha).ToString("yyyyMMdd"), idFormaPago,MenuLatinium.iNivel);
					this.cmbPlanes.DataSource = Funcion.dvProcedimiento(cdsSeteoF, sSQL);

					this.cmbPlanes.DisplayLayout.Bands[0].Columns["idPlan"].Hidden = true;
					this.cmbPlanes.DisplayLayout.Bands[0].Columns["Descripcion"].Width = 224;
				}
			}
		}

		private void cmbPlanes_ValueChanged(object sender, System.EventArgs e)
		{
			if (this.cmbBanco.ActiveRow != null)
			{
				if (this.cmbTarjeta.ActiveRow != null)
				{
					if (this.cmbPlanes.ActiveRow != null) 
					{
						this.cmbPlazo.Value = System.DBNull.Value;

						string sSQL = string.Format("Exec BTPP {0}, {1}, {2}, 2, {3}, '{4}', {5}", 
							(int)this.cmbBanco.Value, (int)this.cmbTarjeta.Value, (int)this.cmbPlanes.Value, iRegion, Convert.ToDateTime(dtfecha).ToString("yyyyMMdd"), idFormaPago);
						this.cmbPlazo.DataSource = Funcion.dvProcedimiento(cdsSeteoF, sSQL);

						this.cmbPlazo.DisplayLayout.Bands[0].Columns["idPlazo"].Hidden = true;
						this.cmbPlazo.DisplayLayout.Bands[0].Columns["Descripcion"].Width = 224;

						if (this.cmbPlazo.Rows.Count == 1)
						{
							int idPlazo = 0;
							idPlazo = Funcion.iEscalarSQL(cdsSeteoF, string.Format("Select idPlazo From PlanesPlazos Where idPlan = {0} And idTarjeta = {1} And idBanco = {2}", 
								(int)this.cmbPlanes.Value, (int)this.cmbTarjeta.Value, (int)this.cmbBanco.Value)); 
							if (idPlazo > 0) this.cmbPlazo.Value = idPlazo;
						}
					}
				}
			}
		}


		private void cmbPlazo_ValueChanged(object sender, System.EventArgs e)
		{
			if (this.cmbBanco.ActiveRow != null)
			{
				if (this.cmbTarjeta.ActiveRow != null)
				{
					if (this.cmbPlanes.ActiveRow != null) 
					{
						if(this.cmbPlazo.ActiveRow != null)
						{
//							this.cmbPorcentaje.Value = System.DBNull.Value;
//
//							string sSQL = string.Format("select isnull(dbo.PorcComisionTarjeta({0}, {1}, {2}, {3}), 0)",
//								(int)this.cmbBanco.Value, (int)this.cmbTarjeta.Value, (int)this.cmbPlanes.Value, (int)this.cmbPlazo.Value);
//
//							this.cmbPorcentaje.DataSource = Funcion.dvProcedimiento(cdsSeteoF, sSQL);
//
//							this.cmbPorcentaje.DisplayLayout.Bands[0].Columns["idPorcentaje"].Hidden = true;
//							this.cmbPorcentaje.DisplayLayout.Bands[0].Columns["Descripcion"].Width = 224;
						}
					}
				}
			}
		}

		private void btGrabar_Click(object sender, System.EventArgs e)
		{
			if (!Validacion.vbComboVacio(this.cmbBanco, "Seleccione el Banco")) return;
			if (!Validacion.vbComboVacio(this.cmbTarjeta, "Seleccione la Tarjeta de Crédito")) return;
			if (!Validacion.vbComboVacio(this.cmbPlanes, "Seleccione el Plan")) return;
			if (!Validacion.vbComboVacio(this.cmbPlazo, "Seleccione el Plazo")) return;

			this.DialogResult = DialogResult.OK;
		}

		private void frmInfoTarjetasDeCreditoDB_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Escape) this.DialogResult = DialogResult.No;
		}
	}
}
