using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Data;
using System.Windows.Forms;
using Infragistics.Win.UltraWinGrid;
using Latinium.Models.OrdenDeCompra.Bodega;

namespace Latinium
{
	public class frmOrdenDeCompraBodega : System.Windows.Forms.Form
	{
		private System.ComponentModel.Container components = null;
		private System.Windows.Forms.Label lblInfo;

		private UltraGrid ultraGridBodegas;
		private Button btnAceptar;
		private Button btnCancelar;

		private DataTable dtBodegas;
		public DataTable DataBodegas;
		public int CantidadTotal = 0;
		public string Articulo = "";
		private Label lblTotalRequerido;
		private Label lblTotalIngresado;
		private Label lblDiferencia;
		public bool SoloLectura = true;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource1;

		// resultado que devolverá al form principal
		public ArrayList Resultado;
		private System.Windows.Forms.Label lblBuscar;
		private System.Windows.Forms.TextBox txtBuscar;

		// dataset del ERP (lo recibimos desde afuera)
		public C1.Data.C1DataSet cdsSeteoF;

		public frmOrdenDeCompraBodega()
		{
			InitializeComponent();
		}

		protected override void Dispose(bool disposing)
		{
			if (disposing)
			{
				if (components != null)
					components.Dispose();
			}
			base.Dispose(disposing);
		}

		#region Código generado por el Diseñador
		private void InitializeComponent()
		{
			Infragistics.Win.Appearance appearance1 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bodega");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cantidad");
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			this.lblTotalRequerido = new System.Windows.Forms.Label();
			this.lblTotalIngresado = new System.Windows.Forms.Label();
			this.lblDiferencia = new System.Windows.Forms.Label();
			this.lblInfo = new System.Windows.Forms.Label();
			this.ultraGridBodegas = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.btnAceptar = new System.Windows.Forms.Button();
			this.btnCancelar = new System.Windows.Forms.Button();
			this.ultraDataSource1 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.lblBuscar = new System.Windows.Forms.Label();
			this.txtBuscar = new System.Windows.Forms.TextBox();
			((System.ComponentModel.ISupportInitialize)(this.ultraGridBodegas)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).BeginInit();
			this.SuspendLayout();
			// 
			// lblTotalRequerido
			// 
			this.lblTotalRequerido.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.lblTotalRequerido.Location = new System.Drawing.Point(8, 544);
			this.lblTotalRequerido.Name = "lblTotalRequerido";
			this.lblTotalRequerido.Size = new System.Drawing.Size(200, 16);
			this.lblTotalRequerido.TabIndex = 0;
			// 
			// lblTotalIngresado
			// 
			this.lblTotalIngresado.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.lblTotalIngresado.Location = new System.Drawing.Point(8, 567);
			this.lblTotalIngresado.Name = "lblTotalIngresado";
			this.lblTotalIngresado.Size = new System.Drawing.Size(200, 16);
			this.lblTotalIngresado.TabIndex = 1;
			// 
			// lblDiferencia
			// 
			this.lblDiferencia.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.lblDiferencia.Location = new System.Drawing.Point(8, 592);
			this.lblDiferencia.Name = "lblDiferencia";
			this.lblDiferencia.Size = new System.Drawing.Size(300, 16);
			this.lblDiferencia.TabIndex = 2;
			// 
			// lblInfo
			// 
			this.lblInfo.Font = new System.Drawing.Font("Arial", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblInfo.ForeColor = System.Drawing.Color.SlateGray;
			this.lblInfo.Location = new System.Drawing.Point(8, 8);
			this.lblInfo.Name = "lblInfo";
			this.lblInfo.Size = new System.Drawing.Size(432, 16);
			this.lblInfo.TabIndex = 0;
			this.lblInfo.Text = "Ingrese la cantidad por bodega. Solo se guardarán las que tengan valores mayores " +
				"a 0.";
			this.lblInfo.Click += new System.EventHandler(this.lblInfo_Click);
			// 
			// ultraGridBodegas
			// 
			this.ultraGridBodegas.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.ultraGridBodegas.Cursor = System.Windows.Forms.Cursors.Default;
			appearance1.BackColor = System.Drawing.Color.White;
			this.ultraGridBodegas.DisplayLayout.Appearance = appearance1;
			this.ultraGridBodegas.DisplayLayout.AutoFitColumns = true;
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn2.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn2.Header.Caption = "Bodega";
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridColumn2.Width = 343;
			ultraGridColumn3.Format = "#,##0";
			ultraGridColumn3.Header.VisiblePosition = 2;
			ultraGridColumn3.MaskInput = "";
			ultraGridColumn3.NullText = "0";
			ultraGridColumn3.PromptChar = ' ';
			ultraGridColumn3.Width = 103;
			ultraGridBand1.Columns.AddRange(new object[] {
																										 ultraGridColumn1,
																										 ultraGridColumn2,
																										 ultraGridColumn3});
			this.ultraGridBodegas.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.ultraGridBodegas.DisplayLayout.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			appearance2.ForeColor = System.Drawing.Color.Black;
			appearance2.ForeColorDisabled = System.Drawing.Color.Black;
			this.ultraGridBodegas.DisplayLayout.Override.ActiveRowAppearance = appearance2;
			this.ultraGridBodegas.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.ultraGridBodegas.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			this.ultraGridBodegas.DisplayLayout.Override.BorderStyleCell = Infragistics.Win.UIElementBorderStyle.Dotted;
			this.ultraGridBodegas.DisplayLayout.Override.BorderStyleRow = Infragistics.Win.UIElementBorderStyle.Dotted;
			appearance3.BackColor = System.Drawing.Color.Transparent;
			this.ultraGridBodegas.DisplayLayout.Override.CardAreaAppearance = appearance3;
			appearance4.ForeColor = System.Drawing.Color.Black;
			appearance4.ForeColorDisabled = System.Drawing.Color.Black;
			this.ultraGridBodegas.DisplayLayout.Override.CellAppearance = appearance4;
			appearance5.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance5.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance5.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance5.FontData.BoldAsString = "True";
			appearance5.FontData.Name = "Arial";
			appearance5.FontData.SizeInPoints = 8.5F;
			appearance5.ForeColor = System.Drawing.Color.White;
			appearance5.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.ultraGridBodegas.DisplayLayout.Override.HeaderAppearance = appearance5;
			appearance6.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance6.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance6.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGridBodegas.DisplayLayout.Override.RowAlternateAppearance = appearance6;
			this.ultraGridBodegas.DisplayLayout.Override.RowSelectors = Infragistics.Win.DefaultableBoolean.False;
			appearance7.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance7.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance7.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGridBodegas.DisplayLayout.Override.SelectedRowAppearance = appearance7;
			this.ultraGridBodegas.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.ultraGridBodegas.Location = new System.Drawing.Point(0, 64);
			this.ultraGridBodegas.Name = "ultraGridBodegas";
			this.ultraGridBodegas.Size = new System.Drawing.Size(448, 464);
			this.ultraGridBodegas.TabIndex = 1;
			this.ultraGridBodegas.BeforeCellUpdate += new Infragistics.Win.UltraWinGrid.BeforeCellUpdateEventHandler(this.ultraGridBodegas_BeforeCellUpdate);
			this.ultraGridBodegas.KeyDown += new System.Windows.Forms.KeyEventHandler(this.ultraGridBodegas_KeyDown);
			this.ultraGridBodegas.AfterCellUpdate += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.ultraGridBodegas_AfterCellUpdate);
			this.ultraGridBodegas.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.ultraGridBodegas_InitializeLayout);
			// 
			// btnAceptar
			// 
			this.btnAceptar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.btnAceptar.Location = new System.Drawing.Point(244, 608);
			this.btnAceptar.Name = "btnAceptar";
			this.btnAceptar.Size = new System.Drawing.Size(90, 25);
			this.btnAceptar.TabIndex = 2;
			this.btnAceptar.Text = "Aceptar";
			this.btnAceptar.Click += new System.EventHandler(this.btnAceptar_Click);
			// 
			// btnCancelar
			// 
			this.btnCancelar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.btnCancelar.Location = new System.Drawing.Point(348, 608);
			this.btnCancelar.Name = "btnCancelar";
			this.btnCancelar.Size = new System.Drawing.Size(90, 25);
			this.btnCancelar.TabIndex = 3;
			this.btnCancelar.Text = "Cancelar";
			this.btnCancelar.Click += new System.EventHandler(this.btnCancelar_Click);
			// 
			// lblBuscar
			// 
			this.lblBuscar.Location = new System.Drawing.Point(8, 32);
			this.lblBuscar.Name = "lblBuscar";
			this.lblBuscar.Size = new System.Drawing.Size(64, 20);
			this.lblBuscar.TabIndex = 8;
			this.lblBuscar.Text = "Bodega:";
			// 
			// txtBuscar
			// 
			this.txtBuscar.Location = new System.Drawing.Point(88, 32);
			this.txtBuscar.Name = "txtBuscar";
			this.txtBuscar.Size = new System.Drawing.Size(352, 20);
			this.txtBuscar.TabIndex = 9;
			this.txtBuscar.Text = "";
			this.txtBuscar.TextChanged += new System.EventHandler(this.txtBuscar_TextChanged);
			// 
			// frmOrdenDeCompraBodega
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(452, 646);
			this.Controls.Add(this.lblBuscar);
			this.Controls.Add(this.txtBuscar);
			this.Controls.Add(this.lblTotalRequerido);
			this.Controls.Add(this.lblTotalIngresado);
			this.Controls.Add(this.lblDiferencia);
			this.Controls.Add(this.ultraGridBodegas);
			this.Controls.Add(this.lblInfo);
			this.Controls.Add(this.btnAceptar);
			this.Controls.Add(this.btnCancelar);
			this.MinimumSize = new System.Drawing.Size(460, 350);
			this.Name = "frmOrdenDeCompraBodega";
			this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
			this.Text = "Distribución por Bodega";
			this.Load += new System.EventHandler(this.frmOrdenDeCompraBodega_Load);
			((System.ComponentModel.ISupportInitialize)(this.ultraGridBodegas)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void frmOrdenDeCompraBodega_Load(object sender, System.EventArgs e)
		{
			if (DataBodegas == null)
			{
				MessageBox.Show("No hay bodegas cargadas");
				this.Close();
				return;
			}

			dtBodegas = DataBodegas.Copy();

			if (!dtBodegas.Columns.Contains("Cantidad"))
				dtBodegas.Columns.Add("Cantidad", typeof(int));

			foreach (DataRow r in dtBodegas.Rows)
			{
				if (r["Cantidad"] == DBNull.Value)
					r["Cantidad"] = 0;
			}

			ultraGridBodegas.DataSource = dtBodegas.DefaultView;
			dtBodegas.DefaultView.RowFilter = "Cantidad > 0";

			ConfigurarGrid();

			// MODO SOLO LECTURA
			if (SoloLectura)
			{
				ultraGridBodegas.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.False;
				ultraGridBodegas.DisplayLayout.Override.CellClickAction = CellClickAction.RowSelect;

				btnAceptar.Visible = false;

				this.Text += " (Consulta)";
				lblInfo.Text = "Modo consulta: solo visualización";

				this.BackColor = Color.WhiteSmoke; 
				ultraGridBodegas.DisplayLayout.Appearance.BackColor = Color.WhiteSmoke;
			}

			foreach (UltraGridRow row in ultraGridBodegas.Rows)
			{
				PintarFila(row);
			}

			ActualizarResumen();
		}

//		private void ConfigurarGrid()
//		{
//			if (ultraGridBodegas.DisplayLayout.Bands.Count == 0)
//				return;
//
//			Infragistics.Win.Appearance appearanceHeader =
//				new Infragistics.Win.Appearance();
//
//			Infragistics.Win.Appearance appearanceRowAlt =
//				new Infragistics.Win.Appearance();
//
//			Infragistics.Win.Appearance appearanceSelected =
//				new Infragistics.Win.Appearance();
//
//			Infragistics.Win.Appearance appearanceCell =
//				new Infragistics.Win.Appearance();
//
//			Infragistics.Win.Appearance appearanceActiveRow =
//				new Infragistics.Win.Appearance();
//
//			Infragistics.Win.Appearance appearanceCantidad =
//				new Infragistics.Win.Appearance();
//
//			Infragistics.Win.UltraWinGrid.UltraGridBand band =
//				ultraGridBodegas.DisplayLayout.Bands[0];
//
//			// =========================
//			// APARIENCIA GENERAL
//			// =========================
//
//			ultraGridBodegas.Font =
//				new Font(
//				"Tahoma",
//				8.25F,
//				FontStyle.Regular,
//				GraphicsUnit.Point,
//				((System.Byte)(0)));
//
//			ultraGridBodegas.DisplayLayout.Appearance.BackColor =
//				Color.White;
//
//			ultraGridBodegas.DisplayLayout.BorderStyle =
//				Infragistics.Win.UIElementBorderStyle.Solid;
//
//			// =========================
//			// HEADER
//			// =========================
//
//			appearanceHeader.BackColor =
//				Color.FromArgb(89, 135, 214);
//
//			appearanceHeader.BackColor2 =
//				Color.FromArgb(7, 59, 150);
//
//			appearanceHeader.BackGradientStyle =
//				Infragistics.Win.GradientStyle.Vertical;
//
//			appearanceHeader.ForeColor =
//				Color.White;
//
//			appearanceHeader.FontData.Bold =
//				Infragistics.Win.DefaultableBoolean.True;
//
//			appearanceHeader.FontData.Name = "Arial";
//
//			appearanceHeader.FontData.SizeInPoints = 8.5F;
//
//			appearanceHeader.ThemedElementAlpha =
//				Infragistics.Win.Alpha.Transparent;
//
//			ultraGridBodegas.DisplayLayout.Override.HeaderAppearance =
//				appearanceHeader;
//
//			// =========================
//			// FILAS
//			// =========================
//
//			appearanceRowAlt.BackColor =
//				SystemColors.ControlLight;
//
//			appearanceRowAlt.BackColor2 =
//				Color.Gainsboro;
//
//			appearanceRowAlt.BackGradientStyle =
//				Infragistics.Win.GradientStyle.Vertical;
//
//			ultraGridBodegas.DisplayLayout.Override.RowAlternateAppearance =
//				appearanceRowAlt;
//
//			appearanceSelected.BackColor =
//				Color.FromArgb(251, 230, 148);
//
//			appearanceSelected.BackColor2 =
//				Color.FromArgb(238, 149, 21);
//
//			appearanceSelected.BackGradientStyle =
//				Infragistics.Win.GradientStyle.Vertical;
//
//			ultraGridBodegas.DisplayLayout.Override.SelectedRowAppearance =
//				appearanceSelected;
//
//			appearanceCell.ForeColor = Color.Black;
//			appearanceCell.ForeColorDisabled = Color.Black;
//
//			ultraGridBodegas.DisplayLayout.Override.CellAppearance =
//				appearanceCell;
//
//			appearanceActiveRow.ForeColor = Color.Black;
//			appearanceActiveRow.ForeColorDisabled = Color.Black;
//
//			ultraGridBodegas.DisplayLayout.Override.ActiveRowAppearance =
//				appearanceActiveRow;
//
//			// =========================
//			// CONFIGURACIÓN GRID
//			// =========================
//
//			ultraGridBodegas.DisplayLayout.Override.AllowDelete =
//				Infragistics.Win.DefaultableBoolean.False;
//
//			ultraGridBodegas.DisplayLayout.Override.AllowUpdate =
//				Infragistics.Win.DefaultableBoolean.True;
//
//			ultraGridBodegas.DisplayLayout.Override.CellClickAction =
//				Infragistics.Win.UltraWinGrid.CellClickAction.Edit;
//
//			ultraGridBodegas.DisplayLayout.Override.RowSelectors =
//				Infragistics.Win.DefaultableBoolean.False;
//
//			ultraGridBodegas.DisplayLayout.Override.BorderStyleCell =
//				Infragistics.Win.UIElementBorderStyle.Dotted;
//
//			ultraGridBodegas.DisplayLayout.Override.BorderStyleRow =
//				Infragistics.Win.UIElementBorderStyle.Dotted;
//
//			ultraGridBodegas.DisplayLayout.Override.CellPadding = 3;
//
//			ultraGridBodegas.DisplayLayout.Override.HeaderClickAction =
//				Infragistics.Win.UltraWinGrid.HeaderClickAction.SortSingle;
//
//			ultraGridBodegas.DisplayLayout.ScrollBounds =
//				Infragistics.Win.UltraWinGrid.ScrollBounds.ScrollToFill;
//
//			ultraGridBodegas.DisplayLayout.ScrollStyle =
//				Infragistics.Win.UltraWinGrid.ScrollStyle.Immediate;
//
//			// =========================
//			// COLUMNAS
//			// =========================
//
//			band.Columns["Bodega"].Hidden = true;
//
//			// NOMBRE
//			band.Columns["Nombre"].Header.Caption = "Bodega";
//			band.Columns["Nombre"].Width = 260;
//			band.Columns["Nombre"].CellActivation =
//				Infragistics.Win.UltraWinGrid.Activation.NoEdit;
//
//			band.Columns["Nombre"].CellAppearance.BackColor =
//				Color.White;
//
//			// CANTIDAD
//			band.Columns["Cantidad"].Header.Caption = "Cantidad";
//
//			band.Columns["Cantidad"].Width = 100;
//
//			band.Columns["Cantidad"].Format = "#,##0";
//
//			band.Columns["Cantidad"].MaskInput =
//				"{double:9.0}";
//
//			band.Columns["Cantidad"].PromptChar = ' ';
//
//			appearanceCantidad.TextHAlign =
//				Infragistics.Win.HAlign.Center;
//
//			band.Columns["Cantidad"].CellAppearance =
//				appearanceCantidad;
//
//			// POSICIONES
//			band.Columns["Nombre"].Header.VisiblePosition = 0;
//			band.Columns["Cantidad"].Header.VisiblePosition = 1;
//
//			// =========================
//			// SOLO LECTURA
//			// =========================
//
//			if (SoloLectura)
//			{
//				ultraGridBodegas.DisplayLayout.Override.AllowUpdate =
//					Infragistics.Win.DefaultableBoolean.False;
//
//				ultraGridBodegas.DisplayLayout.Override.CellClickAction =
//					Infragistics.Win.UltraWinGrid.CellClickAction.RowSelect;
//
//				band.Columns["Cantidad"].CellActivation =
//					Infragistics.Win.UltraWinGrid.Activation.NoEdit;
//			}
//
//			// =========================
//			// FILTRO VISUAL
//			// =========================
//
//			ultraGridBodegas.DisplayLayout.Override.FixedHeaderIndicator =
//				Infragistics.Win.UltraWinGrid.FixedHeaderIndicator.None;
//
//			ultraGridBodegas.DisplayLayout.GroupByBox.Hidden = true;
//
//			// =========================
//			// AUTO AJUSTE
//			// =========================
//
//			try
//			{
//				ultraGridBodegas.DisplayLayout.AutoFitColumns = true;
//			}
//			catch
//			{
//				// Compatible .NET 1.1
//			}
//
//			AjustarColumnas();
//		}


		private void ConfigurarGrid()
		{
			if (this.ultraGridBodegas.DisplayLayout.Bands.Count == 0)
				return;

			Infragistics.Win.UltraWinGrid.UltraGridBand band =
				this.ultraGridBodegas.DisplayLayout.Bands[0];

			// ENTEROS
			band.Columns["Cantidad"].Format =
				"#,##0";

			band.Columns["Cantidad"].MaskInput =
				"";

			band.Columns["Cantidad"].NullText =
				"0";

			band.Columns["Cantidad"].PromptChar =
				' ';

			// SOLO LECTURA
			if (SoloLectura)
			{
				this.ultraGridBodegas.DisplayLayout.Override.AllowUpdate =
					Infragistics.Win.DefaultableBoolean.False;

				this.ultraGridBodegas.DisplayLayout.Override.CellClickAction =
					Infragistics.Win.UltraWinGrid.CellClickAction.RowSelect;

				band.Columns["Cantidad"].CellActivation =
					Infragistics.Win.UltraWinGrid.Activation.NoEdit;
			}
			else
			{
				this.ultraGridBodegas.DisplayLayout.Override.AllowUpdate =
					Infragistics.Win.DefaultableBoolean.True;

				this.ultraGridBodegas.DisplayLayout.Override.CellClickAction =
					Infragistics.Win.UltraWinGrid.CellClickAction.Edit;
			}

			this.ultraGridBodegas.Refresh();
		}

		private void ultraGridBodegas_BeforeCellUpdate(
			object sender,
			Infragistics.Win.UltraWinGrid.BeforeCellUpdateEventArgs e)
		{
			if (e.Cell.Column.Key == "Cantidad")
			{
				try
				{
					int valor =
						Convert.ToInt32(e.NewValue);

					if (valor < 0)
					{
						e.Cancel = true;
					}
				}
				catch
				{
					e.Cancel = true;
				}
			}
		}

		private void AjustarColumnas()
		{
			UltraGridBand band = ultraGridBodegas.DisplayLayout.Bands[0];

			band.Columns["Nombre"].Width = 220;
			band.Columns["Cantidad"].Width = 100;
		}

		private void ultraGridBodegas_InitializeLayout(
			object sender,
			Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
			ConfigurarGrid();
		}

		private void ultraGridBodegas_KeyDown(object sender, KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Delete)
			{
				e.Handled = true;
			}
		}

		// ACEPTAR
		private void btnAceptar_Click(object sender, System.EventArgs e)
		{

			if (SoloLectura)
			{
				this.DialogResult = DialogResult.Cancel;
				this.Close();
				return;
			}

			// Forzar commit de edición
			this.ultraGridBodegas.UpdateData();
			this.ultraGridBodegas.PerformAction(Infragistics.Win.UltraWinGrid.UltraGridAction.ExitEditMode);

			Resultado = new ArrayList();

			int totalBodegas = 0;

			foreach (UltraGridRow row in ultraGridBodegas.Rows)
			{
				int cantidad = 0;

				// Validación segura de cantidad
				if (row.Cells["Cantidad"].Value != DBNull.Value)
				{
					try
					{
						cantidad = Convert.ToInt32(row.Cells["Cantidad"].Value);
					}
					catch
					{
						cantidad = 0;
					}
				}

				if (cantidad > 0)
				{
					// Validar Bodega antes de usarla
					if (row.Cells["Bodega"].Value == DBNull.Value)
						continue;

					DetalleBodegaModel item = new DetalleBodegaModel();

					item.Bodega = Convert.ToInt32(row.Cells["Bodega"].Value);
					item.Cantidad = cantidad;

					Resultado.Add(item);

					totalBodegas += cantidad;
				}
			}

			// Validación 1: mínimo
			if (Resultado.Count == 0)
			{
				MessageBox.Show("Debe ingresar al menos una bodega con cantidad mayor a 0",
					"Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
				return;
			}

			this.DialogResult = DialogResult.OK;
			this.Close();
		}

		// CANCELAR
		private void btnCancelar_Click(object sender, System.EventArgs e)
		{
			this.DialogResult = DialogResult.Cancel;
			this.Close();
		}

		private void lblInfo_Click(object sender, System.EventArgs e)
		{
		
		}

		private void ultraGridBodegas_AfterCellUpdate(object sender, CellEventArgs e)
		{
			if (SoloLectura) return;

			if (e.Cell.Column.Key == "Cantidad")
			{
				PintarFila(e.Cell.Row);
				ActualizarResumen();
			}
		}

		private void PintarFila(UltraGridRow row)
		{
			int cantidad = 0;

			if (row.Cells["Cantidad"].Value != DBNull.Value)
			{
				try
				{
					cantidad = Convert.ToInt32(row.Cells["Cantidad"].Value);
				}
				catch
				{
					cantidad = 0;
				}
			}

			if (cantidad > 0)
			{
				// COLOR CUANDO TIENE VALOR
				row.Appearance.BackColor = Color.Honeydew;
				row.Appearance.ForeColor = Color.Black;
			}
			else
			{
				// RESTAURAR NORMAL
				row.Appearance.ResetBackColor();
				row.Appearance.ResetForeColor();
			}
		}

		private void ActualizarResumen()
		{
			int totalDistribuido = 0;
			int bodegasActivas = 0;

			foreach (UltraGridRow row in ultraGridBodegas.Rows)
			{
				int cantidad = 0;

				if (row.Cells["Cantidad"].Value != DBNull.Value)
				{
					try
					{
						cantidad = Convert.ToInt32(row.Cells["Cantidad"].Value);
					}
					catch
					{
						cantidad = 0;
					}
				}

				if (cantidad > 0)
				{
					totalDistribuido += cantidad;
					bodegasActivas++;
				}
			}

			// Total
			lblTotalRequerido.Text = "Total distribuido: " + totalDistribuido;

			// Cantidad de bodegas
			lblTotalIngresado.Text = "Bodegas activas: " + bodegasActivas;

			// Estado
			if (totalDistribuido == 0)
			{
				lblDiferencia.Text = "Sin distribución";
				lblDiferencia.ForeColor = Color.Gray;
			}
			else
			{
				lblDiferencia.Text = "Distribución en proceso";
				lblDiferencia.ForeColor = Color.Blue;
			}
		}

		private void txtBuscar_TextChanged(object sender, System.EventArgs e)
		{
			string filtro =
				this.txtBuscar.Text.Trim().Replace("'", "''");

			DataView dv =
				this.dtBodegas.DefaultView;

			if (filtro == "")
			{
				dv.RowFilter = "Cantidad > 0";
			}
			else
			{
				dv.RowFilter =
					"(Convert(Bodega, 'System.String') LIKE '%" + filtro + "%' " +
					"OR Nombre LIKE '%" + filtro + "%') " +
					"AND Cantidad > 0";
			}

			this.ultraGridBodegas.Refresh();

			foreach (UltraGridRow row in this.ultraGridBodegas.Rows)
			{
				PintarFila(row);
			}
		}
	}
}