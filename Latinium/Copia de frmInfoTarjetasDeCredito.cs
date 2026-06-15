using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using DevExpress.XtraEditors;

namespace Latinium
{
	/// <summary>
	/// Summary description for frmInfoTarjetasDeCredito.
	/// </summary>
	public class frmInfoTarjetasDeCredito : DevExpress.XtraEditors.XtraForm
	{
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
		public Infragistics.Win.UltraWinEditors.UltraNumericEditor spnTasa;
		public System.Windows.Forms.Label lblfecha;
		/// <summary>
		/// Required designer variable.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public frmInfoTarjetasDeCredito()
		{
			//
			// Required for Windows Form Designer support
			//
			InitializeComponent();

			//
			// TODO: Add any constructor code after InitializeComponent call
			//
		}

		int iBodega = 0;
		DateTime dtfecha;
		int idFormaPago = 0;
		
		public frmInfoTarjetasDeCredito(int IBodega, DateTime dtFecha, int IdFormaPago)
		{
			//
			// Required for Windows Form Designer support
			//
			InitializeComponent();

			iBodega = IBodega;
			dtfecha = dtFecha;
			idFormaPago = IdFormaPago;

			//
			// TODO: Add any constructor code after InitializeComponent call
			//
		}

		/// <summary>
		/// Clean up any resources being used.
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
		/// Required method for Designer support - do not modify
		/// the contents of this method with the code editor.
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
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			this.cmbPlanes = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.lblPlanes = new System.Windows.Forms.Label();
			this.lblTarjeta = new System.Windows.Forms.Label();
			this.lblPlazo = new System.Windows.Forms.Label();
			this.lblBanco = new System.Windows.Forms.Label();
			this.cmbPlazo = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.cmbTarjeta = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.cmbBanco = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.ultraDataSource1 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.btCancelar = new System.Windows.Forms.Button();
			this.btGrabar = new System.Windows.Forms.Button();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.spnTasa = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.lblfecha = new System.Windows.Forms.Label();
			((System.ComponentModel.ISupportInitialize)(this.cmbPlanes)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbPlazo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbTarjeta)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbBanco)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.spnTasa)).BeginInit();
			this.SuspendLayout();
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
			this.cmbPlanes.Location = new System.Drawing.Point(64, 72);
			this.cmbPlanes.Name = "cmbPlanes";
			this.cmbPlanes.Size = new System.Drawing.Size(224, 22);
			this.cmbPlanes.TabIndex = 2;
			this.cmbPlanes.ValueMember = "idPlan";
			this.cmbPlanes.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cmbPlanes_KeyDown);
			this.cmbPlanes.ValueChanged += new System.EventHandler(this.cmbPlanes_ValueChanged);
			this.cmbPlanes.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.cmbPlanes_InitializeLayout);
			// 
			// lblPlanes
			// 
			this.lblPlanes.AutoSize = true;
			this.lblPlanes.BackColor = System.Drawing.Color.Transparent;
			this.lblPlanes.Location = new System.Drawing.Point(8, 75);
			this.lblPlanes.Name = "lblPlanes";
			this.lblPlanes.Size = new System.Drawing.Size(36, 16);
			this.lblPlanes.TabIndex = 184;
			this.lblPlanes.Text = "Planes";
			this.lblPlanes.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// lblTarjeta
			// 
			this.lblTarjeta.AutoSize = true;
			this.lblTarjeta.BackColor = System.Drawing.Color.Transparent;
			this.lblTarjeta.Location = new System.Drawing.Point(8, 43);
			this.lblTarjeta.Name = "lblTarjeta";
			this.lblTarjeta.Size = new System.Drawing.Size(40, 16);
			this.lblTarjeta.TabIndex = 183;
			this.lblTarjeta.Text = "Tarjeta";
			this.lblTarjeta.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// lblPlazo
			// 
			this.lblPlazo.AutoSize = true;
			this.lblPlazo.BackColor = System.Drawing.Color.Transparent;
			this.lblPlazo.Location = new System.Drawing.Point(8, 107);
			this.lblPlazo.Name = "lblPlazo";
			this.lblPlazo.Size = new System.Drawing.Size(30, 16);
			this.lblPlazo.TabIndex = 182;
			this.lblPlazo.Text = "Plazo";
			this.lblPlazo.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// lblBanco
			// 
			this.lblBanco.AutoSize = true;
			this.lblBanco.BackColor = System.Drawing.Color.Transparent;
			this.lblBanco.Location = new System.Drawing.Point(8, 11);
			this.lblBanco.Name = "lblBanco";
			this.lblBanco.Size = new System.Drawing.Size(34, 16);
			this.lblBanco.TabIndex = 181;
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
			this.cmbPlazo.Location = new System.Drawing.Point(64, 104);
			this.cmbPlazo.Name = "cmbPlazo";
			this.cmbPlazo.Size = new System.Drawing.Size(224, 22);
			this.cmbPlazo.TabIndex = 3;
			this.cmbPlazo.ValueMember = "idPlazo";
			this.cmbPlazo.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cmbPlazo_KeyDown);
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
			this.cmbTarjeta.Location = new System.Drawing.Point(64, 40);
			this.cmbTarjeta.Name = "cmbTarjeta";
			this.cmbTarjeta.Size = new System.Drawing.Size(224, 22);
			this.cmbTarjeta.TabIndex = 1;
			this.cmbTarjeta.ValueMember = "idTarjeta";
			this.cmbTarjeta.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cmbTarjeta_KeyDown);
			this.cmbTarjeta.ValueChanged += new System.EventHandler(this.cmbTarjeta_ValueChanged);
			this.cmbTarjeta.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.cmbTarjeta_InitializeLayout);
			// 
			// cmbBanco
			// 
			appearance4.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbBanco.Appearance = appearance4;
			this.cmbBanco.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.cmbBanco.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbBanco.DataSource = this.ultraDataSource1;
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
			this.cmbBanco.Location = new System.Drawing.Point(64, 8);
			this.cmbBanco.Name = "cmbBanco";
			this.cmbBanco.Size = new System.Drawing.Size(224, 22);
			this.cmbBanco.TabIndex = 0;
			this.cmbBanco.ValueMember = "idBanco";
			this.cmbBanco.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cmbBanco_KeyDown);
			this.cmbBanco.ValueChanged += new System.EventHandler(this.cmbBanco_ValueChanged);
			this.cmbBanco.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.cmbBanco_InitializeLayout);
			// 
			// ultraDataSource1
			// 
			ultraDataColumn1.DataType = typeof(int);
			this.ultraDataSource1.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn1,
																																 ultraDataColumn2});
			// 
			// btCancelar
			// 
			this.btCancelar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btCancelar.Location = new System.Drawing.Point(112, 170);
			this.btCancelar.Name = "btCancelar";
			this.btCancelar.Size = new System.Drawing.Size(99, 23);
			this.btCancelar.TabIndex = 5;
			this.btCancelar.Text = "&Cancelar";
			this.btCancelar.Click += new System.EventHandler(this.btCancelar_Click);
			// 
			// btGrabar
			// 
			this.btGrabar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btGrabar.Location = new System.Drawing.Point(8, 170);
			this.btGrabar.Name = "btGrabar";
			this.btGrabar.Size = new System.Drawing.Size(99, 23);
			this.btGrabar.TabIndex = 4;
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
			// spnTasa
			// 
			appearance5.ForeColor = System.Drawing.Color.Firebrick;
			appearance5.ForeColorDisabled = System.Drawing.Color.Firebrick;
			this.spnTasa.Appearance = appearance5;
			this.spnTasa.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.spnTasa.Enabled = false;
			this.spnTasa.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Bold);
			this.spnTasa.FormatString = "#,##0.00";
			this.spnTasa.Location = new System.Drawing.Point(416, 136);
			this.spnTasa.Name = "spnTasa";
			this.spnTasa.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.spnTasa.PromptChar = ' ';
			this.spnTasa.Size = new System.Drawing.Size(8, 22);
			this.spnTasa.TabIndex = 185;
			this.spnTasa.Visible = false;
			// 
			// lblfecha
			// 
			this.lblfecha.AutoSize = true;
			this.lblfecha.BackColor = System.Drawing.Color.Transparent;
			this.lblfecha.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
			this.lblfecha.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Bold);
			this.lblfecha.ForeColor = System.Drawing.Color.Firebrick;
			this.lblfecha.Location = new System.Drawing.Point(64, 136);
			this.lblfecha.Name = "lblfecha";
			this.lblfecha.Size = new System.Drawing.Size(0, 20);
			this.lblfecha.TabIndex = 697;
			this.lblfecha.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			this.lblfecha.Visible = false;
			// 
			// frmInfoTarjetasDeCredito
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 14);
			this.ClientSize = new System.Drawing.Size(298, 200);
			this.ControlBox = false;
			this.Controls.Add(this.lblfecha);
			this.Controls.Add(this.spnTasa);
			this.Controls.Add(this.btCancelar);
			this.Controls.Add(this.btGrabar);
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
			this.Name = "frmInfoTarjetasDeCredito";
			this.ShowInTaskbar = false;
			this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
			this.Text = "Información de Tarjetas De Crédito/Débito";
			this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.frmInfoTarjetasDeCredito_KeyDown);
			this.Load += new System.EventHandler(this.frmInfoTarjetasDeCredito_Load);
			((System.ComponentModel.ISupportInitialize)(this.cmbPlanes)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbPlazo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbTarjeta)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbBanco)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.spnTasa)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		int iRegion = 0;
		private void frmInfoTarjetasDeCredito_Load(object sender, System.EventArgs e)
		{
			//this.cmbBanco.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Select idBanco, Nombre From Banco Where Caja = 1 Order By Nombre");

			iRegion = Funcion.iEscalarSQL(cdsSeteoF, string.Format("Select Region From Bodega Where Bodega = {0}", iBodega));

			string sSQL = string.Format("Exec BTPP 0, 0, 0, 3, {0}, '{1}'", iRegion, Convert.ToDateTime(dtfecha).ToString("yyyyMMdd"));
			this.cmbBanco.DataSource = Funcion.dvProcedimiento(cdsSeteoF, sSQL);
		}

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);
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

					string sSQL = string.Format("Exec BTPP {0}, {1}, 0, 1, {2}, '{3}', {4}", 
						(int)this.cmbBanco.Value, (int)this.cmbTarjeta.Value, iRegion, Convert.ToDateTime(dtfecha).ToString("yyyyMMdd"), idFormaPago);
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

		private void btCancelar_Click(object sender, System.EventArgs e)
		{
			this.DialogResult = DialogResult.No;
		}

		private void frmInfoTarjetasDeCredito_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Escape) this.DialogResult = DialogResult.No;
		}

		private void btGrabar_Click(object sender, System.EventArgs e)
		{
			if (!Validacion.vbComboVacio(this.cmbBanco, "Seleccione el Banco")) return;
			if (!Validacion.vbComboVacio(this.cmbTarjeta, "Seleccione la Tarjeta de Crédito")) return;
			if (!Validacion.vbComboVacio(this.cmbPlanes, "Seleccione el Plan")) return;
			if (!Validacion.vbComboVacio(this.cmbPlazo, "Seleccione el Plazo")) return;
			
			Decimal dPorcInte ;
			dPorcInte = Funcion.decEscalarSQL(cdsSeteoF, string.Format("Select Convert (Decimal(18,2),PorcInteres) as PorcInteres From Plazos p Inner Join PlanesPlazos pp On pp.idPlazo = p.idPlazo Where pp.idPlan = {0} And pp.idTarjeta = {1} And pp.idBanco = {2} And pp.idRegion = {3}", 
				(int)this.cmbPlanes.Value, (int)this.cmbTarjeta.Value, (int)this.cmbBanco.Value, iRegion)); 

			string sFecha = Funcion.sEscalarSQL (cdsSeteoF, String.Format ("select Convert ( varchar (20), day(hasta))+'/'+ Convert ( varchar (20), month(hasta)) + '/' + Convert ( varchar (20), year(hasta)) as fecha  From Plazos p Inner Join PlanesPlazos pp On pp.idPlazo = p.idPlazo Where pp.idPlan = {0} And pp.idTarjeta = {1} And pp.idBanco = {2} And pp.idRegion = {3}", 
				(int)this.cmbPlanes.Value, (int)this.cmbTarjeta.Value, (int)this.cmbBanco.Value, iRegion)); 
			this.spnTasa.Value = Convert.ToDecimal( dPorcInte);
			this.lblfecha.Text = sFecha;			

			this.DialogResult = DialogResult.OK;
		}

		private void cmbBanco_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.cmbTarjeta.Focus();
		}

		private void cmbTarjeta_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.cmbPlanes.Focus();
		}

		private void cmbPlanes_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.cmbPlazo.Focus();
		}

		private void cmbPlazo_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.btGrabar.Focus();
		}

		private void cmbBanco_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}

		private void cmbTarjeta_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}

		private void cmbPlanes_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}	
	}
}
