using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;
using Infragistics.Win.UltraWinGrid;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de frmLiquidarVentaVerde.
	/// </summary>
	public class frmLiquidarVentaVerde : System.Windows.Forms.Form
	{
		private int idProveedor = 0;
		public DataTable ProductosSeleccionados = null;
		private Infragistics.Win.UltraWinGrid.UltraGrid uGridInformacion;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource1;
		private System.Windows.Forms.Button btnAceptar;
		private System.Windows.Forms.Button btnCancelar;
		private System.Windows.Forms.Label label8;
		private System.Windows.Forms.Label label6;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtHasta;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtDesde;
		private System.Windows.Forms.Button btnBuscar;

		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public frmLiquidarVentaVerde()
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
		/// 


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
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Seleccionar");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCompraOrigenVV");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idDetCompraOrigenVV");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idArticulo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Articulo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Proveedor");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cantidad");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Costo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaCompra", -1, null, 0, Infragistics.Win.UltraWinGrid.SortIndicator.Ascending, false);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn11 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Estado");
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Seleccionar");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCompraOrigenVV");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn3 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idDetCompraOrigenVV");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn4 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idArticulo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn5 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Articulo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn6 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Codigo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn7 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Proveedor");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn8 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Cantidad");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn9 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Costo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn10 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("FechaCompra");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn11 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Estado");
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton1 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmLiquidarVentaVerde));
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton2 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			this.uGridInformacion = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource1 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.btnAceptar = new System.Windows.Forms.Button();
			this.btnCancelar = new System.Windows.Forms.Button();
			this.label8 = new System.Windows.Forms.Label();
			this.label6 = new System.Windows.Forms.Label();
			this.dtHasta = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.dtDesde = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.btnBuscar = new System.Windows.Forms.Button();
			((System.ComponentModel.ISupportInitialize)(this.uGridInformacion)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtHasta)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtDesde)).BeginInit();
			this.SuspendLayout();
			// 
			// uGridInformacion
			// 
			this.uGridInformacion.Cursor = System.Windows.Forms.Cursors.Default;
			this.uGridInformacion.DataSource = this.ultraDataSource1;
			appearance1.BackColor = System.Drawing.Color.White;
			this.uGridInformacion.DisplayLayout.Appearance = appearance1;
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridColumn2.Hidden = true;
			ultraGridColumn3.Header.VisiblePosition = 3;
			ultraGridColumn3.Hidden = true;
			ultraGridColumn4.Header.VisiblePosition = 5;
			ultraGridColumn4.Hidden = true;
			ultraGridColumn5.Header.VisiblePosition = 2;
			ultraGridColumn6.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn6.Header.VisiblePosition = 4;
			ultraGridColumn6.Width = 186;
			ultraGridColumn7.Header.VisiblePosition = 6;
			ultraGridColumn8.Header.VisiblePosition = 7;
			ultraGridColumn9.Header.VisiblePosition = 9;
			ultraGridColumn10.Header.VisiblePosition = 8;
			ultraGridColumn11.Header.VisiblePosition = 10;
			ultraGridBand1.Columns.AddRange(new object[] {
																										 ultraGridColumn1,
																										 ultraGridColumn2,
																										 ultraGridColumn3,
																										 ultraGridColumn4,
																										 ultraGridColumn5,
																										 ultraGridColumn6,
																										 ultraGridColumn7,
																										 ultraGridColumn8,
																										 ultraGridColumn9,
																										 ultraGridColumn10,
																										 ultraGridColumn11});
			ultraGridBand1.SummaryFooterCaption = "";
			this.uGridInformacion.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.uGridInformacion.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
			this.uGridInformacion.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.uGridInformacion.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance2.BackColor = System.Drawing.Color.Transparent;
			this.uGridInformacion.DisplayLayout.Override.CardAreaAppearance = appearance2;
			appearance3.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance3.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance3.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance3.FontData.BoldAsString = "True";
			appearance3.FontData.Name = "Arial";
			appearance3.FontData.SizeInPoints = 8F;
			appearance3.ForeColor = System.Drawing.Color.White;
			appearance3.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridInformacion.DisplayLayout.Override.HeaderAppearance = appearance3;
			this.uGridInformacion.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti;
			appearance4.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance4.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance4.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridInformacion.DisplayLayout.Override.RowSelectorAppearance = appearance4;
			appearance5.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance5.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance5.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridInformacion.DisplayLayout.Override.SelectedRowAppearance = appearance5;
			this.uGridInformacion.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridInformacion.Location = new System.Drawing.Point(0, 48);
			this.uGridInformacion.Name = "uGridInformacion";
			this.uGridInformacion.Size = new System.Drawing.Size(792, 144);
			this.uGridInformacion.TabIndex = 383;
			
			// 
			// ultraDataSource1
			// 
			ultraDataColumn1.DataType = typeof(bool);
			ultraDataColumn2.ReadOnly = Infragistics.Win.DefaultableBoolean.False;
			ultraDataColumn3.DataType = typeof(int);
			ultraDataColumn3.ReadOnly = Infragistics.Win.DefaultableBoolean.False;
			ultraDataColumn4.ReadOnly = Infragistics.Win.DefaultableBoolean.False;
			ultraDataColumn8.DataType = typeof(int);
			ultraDataColumn9.DataType = typeof(System.Decimal);
			ultraDataColumn10.DataType = typeof(System.DateTime);
			this.ultraDataSource1.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn1,
																																 ultraDataColumn2,
																																 ultraDataColumn3,
																																 ultraDataColumn4,
																																 ultraDataColumn5,
																																 ultraDataColumn6,
																																 ultraDataColumn7,
																																 ultraDataColumn8,
																																 ultraDataColumn9,
																																 ultraDataColumn10,
																																 ultraDataColumn11});
			// 
			// btnAceptar
			// 
			this.btnAceptar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnAceptar.Location = new System.Drawing.Point(24, 208);
			this.btnAceptar.Name = "btnAceptar";
			this.btnAceptar.Size = new System.Drawing.Size(88, 23);
			this.btnAceptar.TabIndex = 384;
			this.btnAceptar.Text = "&Aceptar";
			this.btnAceptar.Click += new System.EventHandler(this.btnAceptar_Click);
			// 
			// btnCancelar
			// 
			this.btnCancelar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnCancelar.Location = new System.Drawing.Point(136, 208);
			this.btnCancelar.Name = "btnCancelar";
			this.btnCancelar.Size = new System.Drawing.Size(88, 23);
			this.btnCancelar.TabIndex = 385;
			this.btnCancelar.Text = "&Cancelar";
			this.btnCancelar.Click += new System.EventHandler(this.btnCancelar_Click);
			// 
			// label8
			// 
			this.label8.AutoSize = true;
			this.label8.Location = new System.Drawing.Point(192, 16);
			this.label8.Name = "label8";
			this.label8.Size = new System.Drawing.Size(34, 16);
			this.label8.TabIndex = 389;
			this.label8.Text = "Hasta";
			// 
			// label6
			// 
			this.label6.AutoSize = true;
			this.label6.Location = new System.Drawing.Point(8, 16);
			this.label6.Name = "label6";
			this.label6.Size = new System.Drawing.Size(37, 16);
			this.label6.TabIndex = 388;
			this.label6.Text = "Desde";
			// 
			// dtHasta
			// 
			appearance6.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtHasta.Appearance = appearance6;
			this.dtHasta.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton1.Caption = "Today";
			this.dtHasta.DateButtons.Add(dateButton1);
			this.dtHasta.Format = "dd/MM/yyyy";
			this.dtHasta.Location = new System.Drawing.Point(248, 8);
			this.dtHasta.Name = "dtHasta";
			this.dtHasta.NonAutoSizeHeight = 23;
			this.dtHasta.Size = new System.Drawing.Size(104, 21);
			this.dtHasta.SpinButtonsVisible = true;
			this.dtHasta.TabIndex = 387;
			this.dtHasta.Value = ((object)(resources.GetObject("dtHasta.Value")));
			// 
			// dtDesde
			// 
			appearance7.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtDesde.Appearance = appearance7;
			this.dtDesde.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton2.Caption = "Today";
			this.dtDesde.DateButtons.Add(dateButton2);
			this.dtDesde.Format = "dd/MM/yyyy";
			this.dtDesde.Location = new System.Drawing.Point(56, 8);
			this.dtDesde.Name = "dtDesde";
			this.dtDesde.NonAutoSizeHeight = 23;
			this.dtDesde.Size = new System.Drawing.Size(104, 21);
			this.dtDesde.SpinButtonsVisible = true;
			this.dtDesde.TabIndex = 386;
			this.dtDesde.Value = ((object)(resources.GetObject("dtDesde.Value")));
			// 
			// btnBuscar
			// 
			this.btnBuscar.Location = new System.Drawing.Point(384, 8);
			this.btnBuscar.Name = "btnBuscar";
			this.btnBuscar.TabIndex = 392;
			this.btnBuscar.Text = "&Buscar";
			this.btnBuscar.Click += new System.EventHandler(this.btnBuscar_Click);
			// 
			// frmLiquidarVentaVerde
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(792, 237);
			this.Controls.Add(this.btnBuscar);
			this.Controls.Add(this.label8);
			this.Controls.Add(this.label6);
			this.Controls.Add(this.dtHasta);
			this.Controls.Add(this.dtDesde);
			this.Controls.Add(this.btnAceptar);
			this.Controls.Add(this.btnCancelar);
			this.Controls.Add(this.uGridInformacion);
			this.Name = "frmLiquidarVentaVerde";
			this.Text = "Productos por liquidar VV";
			this.Click += new System.EventHandler(this.s);
			this.Load += new System.EventHandler(this.frmLiquidarVentaVerde_Load);
			((System.ComponentModel.ISupportInitialize)(this.uGridInformacion)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtHasta)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtDesde)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion
	
		
		
		public frmLiquidarVentaVerde(int idProveedor)
		{
			InitializeComponent();
			this.idProveedor = idProveedor;
		}
		
		private void frmLiquidarVentaVerde_Load(object sender, System.EventArgs e)
		{
			this.dtDesde.Value = DateTime.Today.AddYears(-1);
			this.dtHasta.Value = DateTime.Today;

			CargarProductos();
		}
		private void CargarProductos()
		{
			if (this.dtDesde.Value == null || this.dtDesde.Value == System.DBNull.Value)
			{
				MessageBox.Show(
					"Debe ingresar la fecha desde.",
					"Venta Verde",
					MessageBoxButtons.OK,
					MessageBoxIcon.Warning
					);
				return;
			}

			if (this.dtHasta.Value == null || this.dtHasta.Value == System.DBNull.Value)
			{
				MessageBox.Show(
					"Debe ingresar la fecha hasta.",
					"Venta Verde",
					MessageBoxButtons.OK,
					MessageBoxIcon.Warning
					);
				return;
			}

			DateTime fechaDesde = Convert.ToDateTime(this.dtDesde.Value);
			DateTime fechaHasta = Convert.ToDateTime(this.dtHasta.Value);

			if (fechaDesde > fechaHasta)
			{
				MessageBox.Show(
					"La fecha inicial no puede ser mayor a la fecha final.",
					"Venta Verde",
					MessageBoxButtons.OK,
					MessageBoxIcon.Warning
					);
				return;
			}

			string sSQL = string.Format(
				"Exec Web_ProductosVentaVerdePorLiquidar {0}, '{1}', '{2}'",
				this.idProveedor,
				fechaDesde.ToString("yyyyMMdd"),
				fechaHasta.ToString("yyyyMMdd")
				);

			FuncionesProcedimientos.dsGeneralModulo(sSQL, this.uGridInformacion);
		}
		

		private void uGridInformacion_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
			if (e.Layout.Bands[0].Columns.Exists("Seleccionar"))
			{
				e.Layout.Bands[0].Columns["Seleccionar"].Style =
					Infragistics.Win.UltraWinGrid.ColumnStyle.CheckBox;

				e.Layout.Bands[0].Columns["Seleccionar"].Header.Caption = "";
				e.Layout.Bands[0].Columns["Seleccionar"].Width = 40;
				e.Layout.Bands[0].Columns["Seleccionar"].CellActivation =
					Infragistics.Win.UltraWinGrid.Activation.AllowEdit;
			}

			if (e.Layout.Bands[0].Columns.Exists("idCompraOrigenVV"))
				e.Layout.Bands[0].Columns["idCompraOrigenVV"].Hidden = true;

			if (e.Layout.Bands[0].Columns.Exists("idDetCompraOrigenVV"))
				e.Layout.Bands[0].Columns["idDetCompraOrigenVV"].Hidden = true;

			if (e.Layout.Bands[0].Columns.Exists("idArticulo"))
				e.Layout.Bands[0].Columns["idArticulo"].Hidden = true;

			e.Layout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
			e.Layout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			e.Layout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
		}

		private void s(object sender, System.EventArgs e)
		{
		
		}

		private void btnAceptar_Click(object sender, System.EventArgs e)
		{
			this.uGridInformacion.UpdateData();

			DataTable dt = new DataTable();

			dt.Columns.Add("idCompraOrigenVV", typeof(int));
			dt.Columns.Add("idDetCompraOrigenVV", typeof(int));
			dt.Columns.Add("idArticulo", typeof(int));
			dt.Columns.Add("Articulo", typeof(string));
			dt.Columns.Add("Codigo", typeof(string));
			dt.Columns.Add("Proveedor", typeof(string));
			dt.Columns.Add("Cantidad", typeof(decimal));
			dt.Columns.Add("Costo", typeof(decimal));
			dt.Columns.Add("FechaCompra", typeof(DateTime));
			dt.Columns.Add("Estado", typeof(string));

			foreach (UltraGridRow row in this.uGridInformacion.Rows)
			{
				bool seleccionado = false;

				if (row.Cells["Seleccionar"].Value != null &&
					row.Cells["Seleccionar"].Value != System.DBNull.Value)
				{
					seleccionado = Convert.ToBoolean(row.Cells["Seleccionar"].Value);
				}

				if (seleccionado)
				{
					DataRow dr = dt.NewRow();

					dr["idCompraOrigenVV"] = Convert.ToInt32(row.Cells["idCompraOrigenVV"].Value);
					dr["idDetCompraOrigenVV"] = Convert.ToInt32(row.Cells["idDetCompraOrigenVV"].Value);
					dr["idArticulo"] = Convert.ToInt32(row.Cells["idArticulo"].Value);
					dr["Articulo"] = row.Cells["Articulo"].Value.ToString();
					dr["Codigo"] = row.Cells["Codigo"].Value.ToString();
					dr["Proveedor"] = row.Cells["Proveedor"].Value.ToString();
					dr["Cantidad"] = Convert.ToDecimal(row.Cells["Cantidad"].Value);
					dr["Costo"] = Convert.ToDecimal(row.Cells["Costo"].Value);
					dr["FechaCompra"] = Convert.ToDateTime(row.Cells["FechaCompra"].Value);
					dr["Estado"] = row.Cells["Estado"].Value.ToString();

					dt.Rows.Add(dr);
				}
			}

			if (dt.Rows.Count == 0)
			{
				MessageBox.Show(
					"Debe seleccionar al menos un producto.",
					"Venta Verde",
					MessageBoxButtons.OK,
					MessageBoxIcon.Warning
					);
				return;
			}

			this.ProductosSeleccionados = dt;
			this.DialogResult = DialogResult.OK;
			this.Close();
		}
		private void btnBuscar_Click(object sender, System.EventArgs e)
		{
			CargarProductos();

		}

		private void btnCancelar_Click(object sender, System.EventArgs e)
		{
			this.ProductosSeleccionados = null;
			this.DialogResult = DialogResult.Cancel;
			this.Close();
		}

	}
}
