using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de frmFacturaDividida.
	/// </summary>
	public class frmFacturaDividida : System.Windows.Forms.Form
	{
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;
		private System.Windows.Forms.Button btnGenerar;
		private C1.Data.C1DataSet cdsSeteoF;
		public Infragistics.Win.UltraWinGrid.UltraGrid grdDetalle;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource1;
		private System.Windows.Forms.Button btnCancelar;

		public frmFacturaDividida()
		{
			//
			// Necesario para admitir el Diseñador de Windows Forms
			//
			InitializeComponent();

			//
			// TODO: agregar código de constructor después de llamar a InitializeComponent
			//
		}
		     
		public int idCompra = 0;      
		public frmFacturaDividida(int IdCompra)
		{
			//
			// Necesario para admitir el Diseñador de Windows Forms
			//
			InitializeComponent();
			 idCompra = IdCompra;			
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
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCompra");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idDetCompra");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idArticulo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Articulo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cantidad");
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("SubTotal");
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bloqueado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Sel");
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCompra");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idDetCompra");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn3 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idArticulo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn4 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Articulo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn5 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Cantidad");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn6 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("SubTotal");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn7 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Bloqueado");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn8 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Sel");
			this.btnGenerar = new System.Windows.Forms.Button();
			this.btnCancelar = new System.Windows.Forms.Button();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.grdDetalle = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource1 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.grdDetalle)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).BeginInit();
			this.SuspendLayout();
			// 
			// btnGenerar
			// 
			this.btnGenerar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.btnGenerar.Location = new System.Drawing.Point(616, 397);
			this.btnGenerar.Name = "btnGenerar";
			this.btnGenerar.Size = new System.Drawing.Size(75, 30);
			this.btnGenerar.TabIndex = 1;
			this.btnGenerar.Text = "Generar";
			this.btnGenerar.Click += new System.EventHandler(this.btnGenerar_Click);
			// 
			// btnCancelar
			// 
			this.btnCancelar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.btnCancelar.Location = new System.Drawing.Point(697, 397);
			this.btnCancelar.Name = "btnCancelar";
			this.btnCancelar.Size = new System.Drawing.Size(75, 30);
			this.btnCancelar.TabIndex = 2;
			this.btnCancelar.Text = "Cancelar";
			this.btnCancelar.Click += new System.EventHandler(this.btnCancelar_Click);
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
			// grdDetalle
			// 
			this.grdDetalle.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.grdDetalle.Cursor = System.Windows.Forms.Cursors.Default;
			this.grdDetalle.DataSource = this.ultraDataSource1;
			appearance1.BackColor = System.Drawing.Color.White;
			this.grdDetalle.DisplayLayout.Appearance = appearance1;
			this.grdDetalle.DisplayLayout.AutoFitColumns = true;
			ultraGridColumn1.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn1.Width = 43;
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridColumn2.Hidden = true;
			ultraGridColumn2.Width = 68;
			ultraGridColumn3.Header.VisiblePosition = 2;
			ultraGridColumn3.Hidden = true;
			ultraGridColumn3.Width = 75;
			ultraGridColumn4.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn4.Header.VisiblePosition = 3;
			ultraGridColumn4.Width = 466;
			ultraGridColumn5.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance2.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn5.CellAppearance = appearance2;
			ultraGridColumn5.Header.VisiblePosition = 4;
			ultraGridColumn5.Width = 70;
			ultraGridColumn6.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance3.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn6.CellAppearance = appearance3;
			ultraGridColumn6.Header.VisiblePosition = 5;
			ultraGridColumn6.Width = 119;
			ultraGridColumn7.Header.VisiblePosition = 6;
			ultraGridColumn7.Hidden = true;
			ultraGridColumn7.Width = 91;
			ultraGridColumn8.Header.Caption = "Seleccionar";
			ultraGridColumn8.Header.VisiblePosition = 7;
			ultraGridColumn8.Width = 100;
			ultraGridBand1.Columns.AddRange(new object[] {
																										 ultraGridColumn1,
																										 ultraGridColumn2,
																										 ultraGridColumn3,
																										 ultraGridColumn4,
																										 ultraGridColumn5,
																										 ultraGridColumn6,
																										 ultraGridColumn7,
																										 ultraGridColumn8});
			this.grdDetalle.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			appearance4.ForeColor = System.Drawing.Color.Black;
			appearance4.ForeColorDisabled = System.Drawing.Color.Black;
			this.grdDetalle.DisplayLayout.Override.ActiveRowAppearance = appearance4;
			this.grdDetalle.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
			this.grdDetalle.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.True;
			this.grdDetalle.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance5.BackColor = System.Drawing.Color.Transparent;
			this.grdDetalle.DisplayLayout.Override.CardAreaAppearance = appearance5;
			appearance6.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance6.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance6.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance6.FontData.BoldAsString = "True";
			appearance6.FontData.Name = "Arial";
			appearance6.FontData.SizeInPoints = 10F;
			appearance6.ForeColor = System.Drawing.Color.White;
			appearance6.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.grdDetalle.DisplayLayout.Override.HeaderAppearance = appearance6;
			appearance7.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance7.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance7.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.grdDetalle.DisplayLayout.Override.RowAlternateAppearance = appearance7;
			appearance8.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance8.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance8.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.grdDetalle.DisplayLayout.Override.RowSelectorAppearance = appearance8;
			appearance9.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance9.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance9.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.grdDetalle.DisplayLayout.Override.SelectedRowAppearance = appearance9;
			this.grdDetalle.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.grdDetalle.Location = new System.Drawing.Point(0, 32);
			this.grdDetalle.Name = "grdDetalle";
			this.grdDetalle.Size = new System.Drawing.Size(776, 352);
			this.grdDetalle.TabIndex = 4;
			this.grdDetalle.CellChange += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.grdDetalle_CellChange);
			this.grdDetalle.AfterCellUpdate += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.grdDetalle_AfterCellUpdate);
			// 
			// ultraDataSource1
			// 
			ultraDataColumn1.DataType = typeof(int);
			ultraDataColumn2.DataType = typeof(int);
			ultraDataColumn3.DataType = typeof(int);
			ultraDataColumn5.DataType = typeof(int);
			ultraDataColumn6.DataType = typeof(System.Decimal);
			ultraDataColumn7.DataType = typeof(int);
			ultraDataColumn8.DataType = typeof(bool);
			this.ultraDataSource1.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn1,
																																 ultraDataColumn2,
																																 ultraDataColumn3,
																																 ultraDataColumn4,
																																 ultraDataColumn5,
																																 ultraDataColumn6,
																																 ultraDataColumn7,
																																 ultraDataColumn8});
			// 
			// frmFacturaDividida
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(784, 438);
			this.ControlBox = false;
			this.Controls.Add(this.grdDetalle);
			this.Controls.Add(this.btnCancelar);
			this.Controls.Add(this.btnGenerar);
			this.Name = "frmFacturaDividida";
			this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
			this.Text = "Dividir Factura de Compra";
			this.Load += new System.EventHandler(this.frmFacturaDividida_Load);
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.grdDetalle)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void frmFacturaDividida_Load(object sender, System.EventArgs e)
		{
			try
			{
				this.grdDetalle.DataSource = Funcion.dvProcedimiento(
					cdsSeteoF, string.Format("Exec CargaInformacionDividirFacturas {0}", idCompra));

				if (this.grdDetalle.Rows.Count > 0)
				{
					if (!this.grdDetalle.DisplayLayout.Bands[0].Columns.Exists("Sel"))
						this.grdDetalle.DisplayLayout.Bands[0].Columns.Add("Sel", "Seleccionar").DataType = typeof(bool);

					int i;
					for (i = 0; i < this.grdDetalle.Rows.Count; i++)
					{
						Infragistics.Win.UltraWinGrid.UltraGridRow row = this.grdDetalle.Rows[i];

						bool sel = false;
						if (row.Cells.Exists("Sel") && row.Cells["Sel"].Value != System.DBNull.Value)
							sel = System.Convert.ToBoolean(row.Cells["Sel"].Value);

						bool bloqueado = false;
						if (row.Cells.Exists("Bloqueado") && row.Cells["Bloqueado"].Value != System.DBNull.Value)
							bloqueado = System.Convert.ToBoolean(row.Cells["Bloqueado"].Value);

						if (sel || bloqueado)
						{
							row.Cells["Sel"].Value = true;
							row.Cells["Sel"].Activation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
						}
						else
						{
							row.Cells["Sel"].Activation = Infragistics.Win.UltraWinGrid.Activation.AllowEdit;
							row.Cells["Sel"].Value = false;
						}					
					}

					this.grdDetalle.Focus();
				}
			}
			catch (System.Exception ex)
			{
				System.Windows.Forms.MessageBox.Show("Error al cargar: " + ex.Message,
					"Point Technology", System.Windows.Forms.MessageBoxButtons.OK,
					System.Windows.Forms.MessageBoxIcon.Error);
			}
		}

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);	
		}

		private void grdDetalle_AfterCellUpdate(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			if (e.Cell.Column.Key == "Sel")
			{
				bool bloqueado = false;
				if (e.Cell.Row.Cells["Bloqueado"].Value != System.DBNull.Value)
					bloqueado = System.Convert.ToBoolean(e.Cell.Row.Cells["Bloqueado"].Value);
				if (bloqueado)
				{
					e.Cell.Value = false;
					return;
				}
				int idDetCompra = (int)e.Cell.Row.Cells["idDetCompra"].Value;				
				bool sel = false;
				if (e.Cell.Value != System.DBNull.Value) sel = System.Convert.ToBoolean(e.Cell.Value);

				if (!sel)
				{
					string sSQLDelete = string.Format(
						"DELETE FROM FacturaDivididaDetalle WHERE idCompra = {0} AND idDetCompra = {1} AND Factura = 1",
						idCompra, idDetCompra);
					Funcion.EjecutaSQL(cdsSeteoF, sSQLDelete);
				}
				else
				{
					string sSQLInsert = string.Format(
						"IF NOT EXISTS(SELECT 1 FROM FacturaDivididaDetalle WHERE idCompra = {0} AND idDetCompra = {1}) " +
						"INSERT INTO FacturaDivididaDetalle(idCompra, idDetCompra, Factura, Procesado) VALUES({0}, {1}, 1, 0) " +
						"ELSE UPDATE FacturaDivididaDetalle SET Factura = 1, Procesado = 0 WHERE idCompra = {0} AND idDetCompra = {1}",
						idCompra, idDetCompra);
					Funcion.EjecutaSQL(cdsSeteoF, sSQLInsert);
				}
			}
		}

		private void grdDetalle_CellChange(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			if (e.Cell.Column.Key == "Factura1")
			{
				bool valor = Convert.ToBoolean(e.Cell.Text);
				if (valor)
				{
					e.Cell.Row.Cells["Factura2"].Value = false;
					e.Cell.Row.Cells["Factura2"].Activation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
				}
				else
				{
					e.Cell.Row.Cells["Factura2"].Activation = Infragistics.Win.UltraWinGrid.Activation.AllowEdit;
				}
			}
			else if (e.Cell.Column.Key == "Factura2")
			{
				bool valor = Convert.ToBoolean(e.Cell.Text);
				if (valor)
				{
					e.Cell.Row.Cells["Factura1"].Value = false;
					e.Cell.Row.Cells["Factura1"].Activation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
				}
				else
				{
					e.Cell.Row.Cells["Factura1"].Activation = Infragistics.Win.UltraWinGrid.Activation.AllowEdit;
				}
			}
		}

		private void btnGenerar_Click(object sender, System.EventArgs e)
		{
			try
			{
				int countSel = 0;
				int i;
				for (i = 0; i < this.grdDetalle.Rows.Count; i++)
				{
					Infragistics.Win.UltraWinGrid.UltraGridRow row = this.grdDetalle.Rows[i];
					if (row == null) continue;

					bool bloqueado = false;
					if (row.Cells["Bloqueado"].Value != System.DBNull.Value)
						bloqueado = System.Convert.ToBoolean(row.Cells["Bloqueado"].Value);

					bool sel = false;
					if (row.Cells["Sel"].Value != System.DBNull.Value)
						sel = System.Convert.ToBoolean(row.Cells["Sel"].Value);

					if (!bloqueado && sel) countSel++;
				}

				if (countSel == 0)
				{
					System.Windows.Forms.MessageBox.Show(
						"Seleccione al menos un producto para facturar.",
						"Point Technology",
						System.Windows.Forms.MessageBoxButtons.OK,
						System.Windows.Forms.MessageBoxIcon.Warning);
					return;
				}
				this.DialogResult  = System.Windows.Forms.DialogResult.OK;
				this.Close();
			}
			catch (System.Exception ex)
			{
				System.Windows.Forms.MessageBox.Show(
					"Error al validar: " + ex.Message,
					"Point Technology",
					System.Windows.Forms.MessageBoxButtons.OK,
					System.Windows.Forms.MessageBoxIcon.Error);
			}
		}

		private void btnCancelar_Click(object sender, System.EventArgs e)
		{
			try
			{
				string sql = string.Format(
					"DELETE FROM FacturaDivididaDetalle WHERE idCompra = {0} AND Procesado = 0",this.idCompra);
				Funcion.EjecutaSQL(cdsSeteoF, sql);

				this.DialogResult = System.Windows.Forms.DialogResult.No;
				this.Close();
			}
			catch (System.Exception ex)
			{
				System.Windows.Forms.MessageBox.Show("Error al cancelar: " + ex.Message,"Point Technology",System.Windows.Forms.MessageBoxButtons.OK,System.Windows.Forms.MessageBoxIcon.Error);
			}
		}
	}
}
