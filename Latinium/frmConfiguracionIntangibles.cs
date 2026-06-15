using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de frmConfiguracionIntangibles.
	/// </summary>
	public class frmConfiguracionIntangibles : System.Windows.Forms.Form
	{
		private C1.Data.C1DataSet cdsSeteoF;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource4;
		public Infragistics.Win.UltraWinGrid.UltraGrid uGridIntangibles;
		private Infragistics.Win.UltraWinGrid.UltraDropDown cmbCargos;
		private Infragistics.Win.UltraWinGrid.UltraDropDown cmbFormasPago;
		private System.Windows.Forms.Button btnExportar;
		private System.Windows.Forms.Button btnGuardar;
		private System.Windows.Forms.Button btnSalir;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public frmConfiguracionIntangibles()
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
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idConfiguracionIntangibles");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCargo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn3 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Desde");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn4 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Hasta");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn5 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idArticulo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn6 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Codigo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn7 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idTipoPrecio");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn8 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Valor");
			Infragistics.Win.Appearance appearance1 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idConfiguracionIntangibles");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCargo", -1, "cmbCargos");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Desde");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Hasta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idArticulo");
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idTipoPrecio", -1, "cmbFormasPago");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Valor");
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings1 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Valor", 7, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Valor", 7, true);
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCargo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cargo");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand3 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn11 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idFormaPago");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn12 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FormaPago");
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmConfiguracionIntangibles));
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.ultraDataSource4 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.uGridIntangibles = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.cmbCargos = new Infragistics.Win.UltraWinGrid.UltraDropDown();
			this.cmbFormasPago = new Infragistics.Win.UltraWinGrid.UltraDropDown();
			this.btnExportar = new System.Windows.Forms.Button();
			this.btnGuardar = new System.Windows.Forms.Button();
			this.btnSalir = new System.Windows.Forms.Button();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource4)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridIntangibles)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCargos)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbFormasPago)).BeginInit();
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
			// ultraDataSource4
			// 
			ultraDataColumn1.DataType = typeof(int);
			ultraDataColumn2.DataType = typeof(int);
			ultraDataColumn3.DataType = typeof(System.DateTime);
			ultraDataColumn4.DataType = typeof(System.DateTime);
			ultraDataColumn5.DataType = typeof(int);
			ultraDataColumn7.DataType = typeof(int);
			ultraDataColumn8.DataType = typeof(System.Decimal);
			this.ultraDataSource4.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn1,
																																 ultraDataColumn2,
																																 ultraDataColumn3,
																																 ultraDataColumn4,
																																 ultraDataColumn5,
																																 ultraDataColumn6,
																																 ultraDataColumn7,
																																 ultraDataColumn8});
			// 
			// uGridIntangibles
			// 
			this.uGridIntangibles.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.uGridIntangibles.Cursor = System.Windows.Forms.Cursors.Default;
			this.uGridIntangibles.DataSource = this.ultraDataSource4;
			appearance1.BackColor = System.Drawing.Color.White;
			this.uGridIntangibles.DisplayLayout.Appearance = appearance1;
			this.uGridIntangibles.DisplayLayout.AutoFitColumns = true;
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn1.Width = 127;
			ultraGridColumn2.ButtonDisplayStyle = Infragistics.Win.UltraWinGrid.ButtonDisplayStyle.Always;
			ultraGridColumn2.Header.Caption = "Cargo";
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridColumn2.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownList;
			ultraGridColumn2.Width = 193;
			ultraGridColumn3.Header.VisiblePosition = 2;
			ultraGridColumn3.Width = 120;
			ultraGridColumn4.Header.VisiblePosition = 3;
			ultraGridColumn4.Width = 124;
			appearance2.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn5.CellAppearance = appearance2;
			ultraGridColumn5.Format = "#,##0.00";
			ultraGridColumn5.Header.VisiblePosition = 4;
			ultraGridColumn5.Hidden = true;
			ultraGridColumn5.Width = 102;
			appearance3.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn6.CellAppearance = appearance3;
			ultraGridColumn6.Format = "#,##0.00";
			ultraGridColumn6.Header.VisiblePosition = 5;
			ultraGridColumn6.Width = 118;
			ultraGridColumn7.ButtonDisplayStyle = Infragistics.Win.UltraWinGrid.ButtonDisplayStyle.Always;
			ultraGridColumn7.Header.Caption = "Tipo Precio";
			ultraGridColumn7.Header.VisiblePosition = 6;
			ultraGridColumn7.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownList;
			ultraGridColumn7.Width = 195;
			appearance4.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn8.CellAppearance = appearance4;
			ultraGridColumn8.Format = "#,##0.00";
			ultraGridColumn8.Header.VisiblePosition = 7;
			ultraGridColumn8.Width = 77;
			ultraGridBand1.Columns.AddRange(new object[] {
																										 ultraGridColumn1,
																										 ultraGridColumn2,
																										 ultraGridColumn3,
																										 ultraGridColumn4,
																										 ultraGridColumn5,
																										 ultraGridColumn6,
																										 ultraGridColumn7,
																										 ultraGridColumn8});
			ultraGridBand1.Header.Caption = "Intangibles";
			ultraGridBand1.HeaderVisible = true;
			ultraGridBand1.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.Edit;
			appearance5.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings1.Appearance = appearance5;
			summarySettings1.DisplayFormat = "{0:#,##0.00}";
			summarySettings1.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			ultraGridBand1.Summaries.AddRange(new Infragistics.Win.UltraWinGrid.SummarySettings[] {
																																															summarySettings1});
			ultraGridBand1.SummaryFooterCaption = "Total";
			this.uGridIntangibles.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.uGridIntangibles.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.TemplateOnBottom;
			appearance6.BackColor = System.Drawing.Color.Transparent;
			this.uGridIntangibles.DisplayLayout.Override.CardAreaAppearance = appearance6;
			appearance7.ForeColor = System.Drawing.Color.Black;
			appearance7.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridIntangibles.DisplayLayout.Override.CellAppearance = appearance7;
			this.uGridIntangibles.DisplayLayout.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.Edit;
			appearance8.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance8.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance8.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance8.FontData.BoldAsString = "True";
			appearance8.FontData.Name = "Arial";
			appearance8.FontData.SizeInPoints = 8F;
			appearance8.ForeColor = System.Drawing.Color.White;
			appearance8.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridIntangibles.DisplayLayout.Override.HeaderAppearance = appearance8;
			appearance9.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance9.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance9.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridIntangibles.DisplayLayout.Override.RowSelectorAppearance = appearance9;
			appearance10.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance10.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance10.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridIntangibles.DisplayLayout.Override.SelectedRowAppearance = appearance10;
			this.uGridIntangibles.Font = new System.Drawing.Font("Tahoma", 8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridIntangibles.Location = new System.Drawing.Point(8, 48);
			this.uGridIntangibles.Name = "uGridIntangibles";
			this.uGridIntangibles.Size = new System.Drawing.Size(848, 360);
			this.uGridIntangibles.TabIndex = 253;
			this.uGridIntangibles.AfterRowInsert += new Infragistics.Win.UltraWinGrid.RowEventHandler(this.uGridIntangibles_AfterRowInsert);
			this.uGridIntangibles.BeforeRowsDeleted += new Infragistics.Win.UltraWinGrid.BeforeRowsDeletedEventHandler(this.uGridIntangibles_BeforeRowsDeleted);
			this.uGridIntangibles.KeyDown += new System.Windows.Forms.KeyEventHandler(this.uGridIntangibles_KeyDown);
			this.uGridIntangibles.AfterCellUpdate += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.uGridIntangibles_AfterCellUpdate);
			// 
			// cmbCargos
			// 
			this.cmbCargos.Cursor = System.Windows.Forms.Cursors.Default;
			ultraGridColumn9.Header.VisiblePosition = 0;
			ultraGridColumn9.Hidden = true;
			ultraGridColumn10.Header.VisiblePosition = 1;
			ultraGridColumn10.Width = 200;
			ultraGridBand2.Columns.AddRange(new object[] {
																										 ultraGridColumn9,
																										 ultraGridColumn10});
			this.cmbCargos.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			this.cmbCargos.DisplayMember = "Cargo";
			this.cmbCargos.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F);
			this.cmbCargos.Location = new System.Drawing.Point(120, 232);
			this.cmbCargos.Name = "cmbCargos";
			this.cmbCargos.Size = new System.Drawing.Size(200, 112);
			this.cmbCargos.TabIndex = 769;
			this.cmbCargos.ValueMember = "idCargo";
			this.cmbCargos.Visible = false;
			// 
			// cmbFormasPago
			// 
			this.cmbFormasPago.Cursor = System.Windows.Forms.Cursors.Default;
			ultraGridColumn11.Header.VisiblePosition = 0;
			ultraGridColumn11.Hidden = true;
			ultraGridColumn12.Header.VisiblePosition = 1;
			ultraGridColumn12.Width = 200;
			ultraGridBand3.Columns.AddRange(new object[] {
																										 ultraGridColumn11,
																										 ultraGridColumn12});
			this.cmbFormasPago.DisplayLayout.BandsSerializer.Add(ultraGridBand3);
			this.cmbFormasPago.DisplayMember = "FormaPago";
			this.cmbFormasPago.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F);
			this.cmbFormasPago.Location = new System.Drawing.Point(496, 232);
			this.cmbFormasPago.Name = "cmbFormasPago";
			this.cmbFormasPago.Size = new System.Drawing.Size(200, 112);
			this.cmbFormasPago.TabIndex = 776;
			this.cmbFormasPago.ValueMember = "idFormaPago";
			this.cmbFormasPago.Visible = false;
			// 
			// btnExportar
			// 
			this.btnExportar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btnExportar.Image = ((System.Drawing.Image)(resources.GetObject("btnExportar.Image")));
			this.btnExportar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnExportar.Location = new System.Drawing.Point(624, 8);
			this.btnExportar.Name = "btnExportar";
			this.btnExportar.Size = new System.Drawing.Size(72, 24);
			this.btnExportar.TabIndex = 779;
			this.btnExportar.Text = "Exportar";
			this.btnExportar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnExportar.Visible = false;
			this.btnExportar.Click += new System.EventHandler(this.btnExportar_Click);
			// 
			// btnGuardar
			// 
			this.btnGuardar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btnGuardar.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(230)), ((System.Byte)(230)), ((System.Byte)(230)));
			this.btnGuardar.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.btnGuardar.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(50)), ((System.Byte)(50)), ((System.Byte)(50)));
			this.btnGuardar.Image = ((System.Drawing.Image)(resources.GetObject("btnGuardar.Image")));
			this.btnGuardar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnGuardar.Location = new System.Drawing.Point(704, 8);
			this.btnGuardar.Name = "btnGuardar";
			this.btnGuardar.Size = new System.Drawing.Size(76, 23);
			this.btnGuardar.TabIndex = 777;
			this.btnGuardar.Text = "Guardar";
			this.btnGuardar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnGuardar.Click += new System.EventHandler(this.btnGuardar_Click);
			// 
			// btnSalir
			// 
			this.btnSalir.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btnSalir.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(230)), ((System.Byte)(230)), ((System.Byte)(230)));
			this.btnSalir.CausesValidation = false;
			this.btnSalir.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.btnSalir.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(50)), ((System.Byte)(50)), ((System.Byte)(50)));
			this.btnSalir.Image = ((System.Drawing.Image)(resources.GetObject("btnSalir.Image")));
			this.btnSalir.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnSalir.Location = new System.Drawing.Point(784, 8);
			this.btnSalir.Name = "btnSalir";
			this.btnSalir.TabIndex = 778;
			this.btnSalir.Text = "Salir";
			this.btnSalir.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnSalir.Click += new System.EventHandler(this.btnSalir_Click_1);
			// 
			// frmConfiguracionIntangibles
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(864, 414);
			this.Controls.Add(this.btnExportar);
			this.Controls.Add(this.btnGuardar);
			this.Controls.Add(this.btnSalir);
			this.Controls.Add(this.cmbFormasPago);
			this.Controls.Add(this.cmbCargos);
			this.Controls.Add(this.uGridIntangibles);
			this.Name = "frmConfiguracionIntangibles";
			this.Text = "Configuración Intangibles";
			this.Load += new System.EventHandler(this.frmConfiguracionIntangibles_Load);
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource4)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridIntangibles)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCargos)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbFormasPago)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private Acceso miAcceso;

		private void UnloadMe()
		{
			this.Close();
		}

		private void frmConfiguracionIntangibles_Load(object sender, System.EventArgs e)
		{
			miAcceso = new Acceso(cdsSeteoF, "052116");

			if (!Funcion.Permiso("2174", cdsSeteoF))
			{				
				MessageBox.Show("No puede ingresar a Configuración de Intangible", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				BeginInvoke(new MethodInvoker(UnloadMe));
				return;
			}

			if (miAcceso.BExportar) this.btnExportar.Visible = true;
			

			this.cmbCargos.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "	 Select idCargo, Cargo From dbo.Com_CargosDeVentas Where Comisiona = 1 and idCargo in (1,3) Order By Orden");
			this.cmbFormasPago.DataSource = Funcion.dvProcedimiento(cdsSeteoF, 
				@"Select idFormaPago, FormaPago From dbo.CompraFormaPago Where Activo = 1 and idFormaPago in (1,3,9) 
				union
				Select 0 as idFormaPago, 'TODOS' as FormaPago
				Order By idFormaPago");
			Consulta ();
		}
		private void Consulta()
		{
			this.uGridIntangibles.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Exec ConsultaConfiguracionIntangibles");
		}

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);		
	
		}

		private void btnExportar_Click(object sender, System.EventArgs e)
		{
			FuncionesProcedimientos.ExportaExcel(this.uGridIntangibles);		
		}

		private void btnSalir_Click_1(object sender, System.EventArgs e)
		{
			this.Close();
		}

		private void uGridIntangibles_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.F3)
			{
				if (this.uGridIntangibles.ActiveCell.Column.Key.ToString().ToUpper() == "CODIGO" )
				{
					using (BuscaArticulos busqueda = new BuscaArticulos("", 1, 99))
					{
						if (busqueda.ShowDialog() == DialogResult.OK)
						{
							// Obtener los valores que necesitas
							int idArticulo = busqueda.idArticulo;
							string codigo = busqueda.sCodigo;

							// Usarlos como desees, por ejemplo:
							uGridIntangibles.ActiveRow.Cells["idArticulo"].Value = idArticulo;							
							uGridIntangibles.ActiveRow.Cells["Codigo"].Value = codigo;
						}
					}

				}
			}
		}

		private void btnGuardar_Click(object sender, System.EventArgs e)
		{
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridIntangibles.Rows.All)
			{
				#region Cargo
				if (dr.Cells["idCargo"].Value == System.DBNull.Value)
				{
					MessageBox.Show("Seleccione un Cargo", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);

					this.uGridIntangibles.ActiveRow = dr;
					this.uGridIntangibles.ActiveRow.Selected = true;

					return;
				}

				if ((int)dr.Cells["idCargo"].Value == 0)
				{
					MessageBox.Show("Seleccione un Cargo", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);

					this.uGridIntangibles.ActiveRow = dr;
					this.uGridIntangibles.ActiveRow.Selected = true;

					return;
				}
				#endregion Cargo

				#region Tipo Precio
				if (dr.Cells["idTipoPrecio"].Value == System.DBNull.Value)
				{
					MessageBox.Show("Seleccione el Tipo de Precio", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);

					this.uGridIntangibles.ActiveRow = dr;
					this.uGridIntangibles.ActiveRow.Selected = true;

					return;
				}

				#endregion Tipo Precio

				#region Articulo
				if (dr.Cells["idArticulo"].Value == System.DBNull.Value)
				{
					MessageBox.Show("Ingrese un Articulo", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);

					this.uGridIntangibles.ActiveRow = dr;
					this.uGridIntangibles.ActiveRow.Selected = true;

					return;
				}

				if ((int)dr.Cells["idArticulo"].Value == 0)
				{
					MessageBox.Show("Ingrese un Articulo", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);

					this.uGridIntangibles.ActiveRow = dr;
					this.uGridIntangibles.ActiveRow.Selected = true;

					return;
				}
				#endregion Articulo

				
				#region Validacion Fecha

				if (dr.Cells["Desde"].Value == System.DBNull.Value)
				{
					MessageBox.Show("Seleccione un la Fecha inicial", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);

					this.uGridIntangibles.ActiveRow = dr;
					this.uGridIntangibles.ActiveRow.Selected = true;

					return;
				}

				if (dr.Cells["Hasta"].Value == System.DBNull.Value)
				{
					MessageBox.Show("Seleccione un la Fecha Final", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);

					this.uGridIntangibles.ActiveRow = dr;
					this.uGridIntangibles.ActiveRow.Selected = true;

					return;
				}

				if (Convert.ToDateTime(dr.Cells["Desde"].Value) > Convert.ToDateTime(dr.Cells["Hasta"].Value))
				{
					MessageBox.Show("La fecha inicial no puede ser mayor a la fecha final.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);

					this.uGridIntangibles.ActiveRow = dr;
					this.uGridIntangibles.ActiveRow.Selected = true;

					return;
				}
				#endregion Validacion Fecha

			}

			if (DialogResult.Yes == MessageBox.Show("Esta seguro de Guardar los cambios", "Point Technology", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2)) 
			{
				foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridIntangibles.Rows.All)
				{
					string sSQLDetalle = string.Format("Exec [GuardaConfiguracionIntangibles] {0}, {1}, '{2}', '{3}', {4}, {5}, {6}",
						(int)dr.Cells["idConfiguracionIntangibles"].Value, 
						(int)dr.Cells["idCargo"].Value,
						Convert.ToDateTime(dr.Cells["Desde"].Value).ToString("yyyyMMdd"), 
						Convert.ToDateTime(dr.Cells["Hasta"].Value).ToString("yyyyMMdd"),
						(int)dr.Cells["idArticulo"].Value,
						(int)dr.Cells["idTipoPrecio"].Value,
						(decimal)dr.Cells["Valor"].Value);

					Funcion.EjecutaSQL(cdsSeteoF, sSQLDetalle);
				}

				Consulta();
			}
		}

		private void uGridIntangibles_AfterRowInsert(object sender, Infragistics.Win.UltraWinGrid.RowEventArgs e)
		{
			e.Row.Cells["idConfiguracionIntangibles"].Value = 0;
			e.Row.Cells["idCargo"].Value = 0;
			e.Row.Cells["idArticulo"].Value = 0;			
			e.Row.Cells["idTipoPrecio"].Value = 0.00m;
			e.Row.Cells["Valor"].Value = 0.00m;
		}
	
		bool bActualiza = false;
		private void uGridIntangibles_AfterCellUpdate(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			if (bActualiza) return;

			if (e.Cell.Column.ToString() == "Valor")
			{	
				bActualiza = true;

				if (e.Cell.Row.Cells["Valor"].Value == System.DBNull.Value)
				{					
					e.Cell.Row.Cells["Valor"].Value = (decimal)e.Cell.OriginalValue;	
				}	
				if ((decimal ) e.Cell.Row.Cells["Valor"].Value < 0 )
				{					
					e.Cell.Row.Cells["Valor"].Value = (decimal)e.Cell.OriginalValue;	
				}
				bActualiza = false;
			}
		}

		private void uGridIntangibles_BeforeRowsDeleted(object sender, Infragistics.Win.UltraWinGrid.BeforeRowsDeletedEventArgs e)
		{
			
			#region Elimina

			if (miAcceso.Eliminar)
			{
				if (DialogResult.Yes ==	MessageBox.Show("¿Esta seguro de Eliminar el registro?", "Point Technology", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2))
				{				
					for (int i=0; i < e.Rows.Length; i++)
					{	
						if ((int)e.Rows[i].Cells["idConfiguracionIntangibles"].Value > 0)
						{
							string sSQL = string.Format("Delete From ConfiguracionIntangibles where idConfiguracionIntangibles = {0} ", (int)e.Rows[i].Cells["idConfiguracionIntangibles"].Value);
							Funcion.EjecutaSQL(cdsSeteoF, sSQL, true);

							e.DisplayPromptMsg = false;
						}
						else e.DisplayPromptMsg = false;
					}					
				}
				else e.Cancel = true;
				#endregion Elimina
			}
			else
			{
				e.DisplayPromptMsg = false;
				e.Cancel = true;
			}
		}

		
	}
}
