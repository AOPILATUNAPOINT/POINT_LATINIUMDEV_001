using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de frmPorcentajeCobradoresProyectado.
	/// </summary>
	public class frmPorcentajeCobradoresProyectado : System.Windows.Forms.Form
	{
		private System.Windows.Forms.DateTimePicker dtpPeriodo;
		private System.Windows.Forms.Label label4;
		private System.Windows.Forms.GroupBox groupBox1;
		private Infragistics.Win.UltraWinGrid.UltraGrid uGridAsignacion;
		private C1.Data.C1DataSet cdsSeteoF;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource1;
		private System.Windows.Forms.Button btnGuardar;
		private System.Windows.Forms.Button btnVer;
		private System.Windows.Forms.Button btnComandos;
		private System.Windows.Forms.ContextMenu mnuAcciones;
		private System.Windows.Forms.MenuItem mnuAccionesCopiarUltimoMes;
		private System.Windows.Forms.MenuItem mnuAccionesAgregarNuevoLocal;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public frmPorcentajeCobradoresProyectado()
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
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPorcentajeCobradoresProyectado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("DiaDesde");
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("DiaHasta");
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Porcentaje");
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance11 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idPorcentajeCobradoresProyectado");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("DiaDesde");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn3 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("DiaHasta");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn4 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Porcentaje");
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmPorcentajeCobradoresProyectado));
			this.dtpPeriodo = new System.Windows.Forms.DateTimePicker();
			this.label4 = new System.Windows.Forms.Label();
			this.groupBox1 = new System.Windows.Forms.GroupBox();
			this.uGridAsignacion = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource1 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.btnGuardar = new System.Windows.Forms.Button();
			this.btnVer = new System.Windows.Forms.Button();
			this.btnComandos = new System.Windows.Forms.Button();
			this.mnuAcciones = new System.Windows.Forms.ContextMenu();
			this.mnuAccionesCopiarUltimoMes = new System.Windows.Forms.MenuItem();
			this.mnuAccionesAgregarNuevoLocal = new System.Windows.Forms.MenuItem();
			((System.ComponentModel.ISupportInitialize)(this.uGridAsignacion)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			this.SuspendLayout();
			// 
			// dtpPeriodo
			// 
			this.dtpPeriodo.CustomFormat = "MM/yyyy";
			this.dtpPeriodo.Format = System.Windows.Forms.DateTimePickerFormat.Custom;
			this.dtpPeriodo.Location = new System.Drawing.Point(72, 16);
			this.dtpPeriodo.MinDate = new System.DateTime(2016, 5, 1, 0, 0, 0, 0);
			this.dtpPeriodo.Name = "dtpPeriodo";
			this.dtpPeriodo.ShowUpDown = true;
			this.dtpPeriodo.Size = new System.Drawing.Size(96, 20);
			this.dtpPeriodo.TabIndex = 746;
			this.dtpPeriodo.Value = new System.DateTime(2016, 5, 1, 0, 0, 0, 0);
			this.dtpPeriodo.ValueChanged += new System.EventHandler(this.dtpPeriodo_ValueChanged);
			// 
			// label4
			// 
			this.label4.AutoSize = true;
			this.label4.BackColor = System.Drawing.Color.Transparent;
			this.label4.Location = new System.Drawing.Point(8, 16);
			this.label4.Name = "label4";
			this.label4.Size = new System.Drawing.Size(43, 16);
			this.label4.TabIndex = 745;
			this.label4.Text = "Periodo";
			this.label4.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// groupBox1
			// 
			this.groupBox1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.groupBox1.Location = new System.Drawing.Point(0, 48);
			this.groupBox1.Name = "groupBox1";
			this.groupBox1.Size = new System.Drawing.Size(872, 8);
			this.groupBox1.TabIndex = 748;
			this.groupBox1.TabStop = false;
			// 
			// uGridAsignacion
			// 
			this.uGridAsignacion.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.uGridAsignacion.Cursor = System.Windows.Forms.Cursors.Default;
			this.uGridAsignacion.DataSource = this.ultraDataSource1;
			appearance1.BackColor = System.Drawing.Color.White;
			this.uGridAsignacion.DisplayLayout.Appearance = appearance1;
			this.uGridAsignacion.DisplayLayout.AutoFitColumns = true;
			ultraGridBand1.AddButtonCaption = "BodegaPromocion";
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn1.Width = 299;
			appearance2.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn2.CellAppearance = appearance2;
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridColumn2.Width = 189;
			appearance3.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn3.CellAppearance = appearance3;
			ultraGridColumn3.Header.VisiblePosition = 2;
			ultraGridColumn3.Width = 184;
			appearance4.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn4.CellAppearance = appearance4;
			ultraGridColumn4.Format = "#,##0.00";
			ultraGridColumn4.Header.VisiblePosition = 3;
			ultraGridColumn4.Width = 238;
			ultraGridBand1.Columns.AddRange(new object[] {
																										 ultraGridColumn1,
																										 ultraGridColumn2,
																										 ultraGridColumn3,
																										 ultraGridColumn4});
			ultraGridBand1.Header.Caption = "";
			ultraGridBand1.SummaryFooterCaption = "";
			this.uGridAsignacion.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			appearance5.ForeColor = System.Drawing.Color.Black;
			appearance5.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridAsignacion.DisplayLayout.Override.ActiveRowAppearance = appearance5;
			this.uGridAsignacion.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.TemplateOnBottom;
			this.uGridAsignacion.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			appearance6.BackColor = System.Drawing.Color.Transparent;
			this.uGridAsignacion.DisplayLayout.Override.CardAreaAppearance = appearance6;
			appearance7.ForeColor = System.Drawing.Color.Black;
			appearance7.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridAsignacion.DisplayLayout.Override.CellAppearance = appearance7;
			appearance8.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance8.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance8.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance8.FontData.BoldAsString = "True";
			appearance8.FontData.Name = "Arial";
			appearance8.FontData.SizeInPoints = 8.5F;
			appearance8.ForeColor = System.Drawing.Color.White;
			appearance8.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridAsignacion.DisplayLayout.Override.HeaderAppearance = appearance8;
			appearance9.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance9.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance9.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridAsignacion.DisplayLayout.Override.RowAlternateAppearance = appearance9;
			appearance10.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance10.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance10.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridAsignacion.DisplayLayout.Override.RowSelectorAppearance = appearance10;
			appearance11.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance11.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance11.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridAsignacion.DisplayLayout.Override.SelectedRowAppearance = appearance11;
			this.uGridAsignacion.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridAsignacion.Location = new System.Drawing.Point(8, 64);
			this.uGridAsignacion.Name = "uGridAsignacion";
			this.uGridAsignacion.Size = new System.Drawing.Size(632, 400);
			this.uGridAsignacion.TabIndex = 747;
			this.uGridAsignacion.AfterRowInsert += new Infragistics.Win.UltraWinGrid.RowEventHandler(this.uGridAsignacion_AfterRowInsert);
			this.uGridAsignacion.BeforeRowsDeleted += new Infragistics.Win.UltraWinGrid.BeforeRowsDeletedEventHandler(this.uGridAsignacion_BeforeRowsDeleted);
			this.uGridAsignacion.AfterCellUpdate += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.uGridAsignacion_AfterCellUpdate);
			// 
			// ultraDataSource1
			// 
			ultraDataColumn1.DataType = typeof(int);
			ultraDataColumn2.DataType = typeof(int);
			ultraDataColumn3.DataType = typeof(int);
			ultraDataColumn4.DataType = typeof(System.Decimal);
			this.ultraDataSource1.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn1,
																																 ultraDataColumn2,
																																 ultraDataColumn3,
																																 ultraDataColumn4});
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
			// btnGuardar
			// 
			this.btnGuardar.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(230)), ((System.Byte)(230)), ((System.Byte)(230)));
			this.btnGuardar.CausesValidation = false;
			this.btnGuardar.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.btnGuardar.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(50)), ((System.Byte)(50)), ((System.Byte)(50)));
			this.btnGuardar.Image = ((System.Drawing.Image)(resources.GetObject("btnGuardar.Image")));
			this.btnGuardar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnGuardar.Location = new System.Drawing.Point(264, 16);
			this.btnGuardar.Name = "btnGuardar";
			this.btnGuardar.Size = new System.Drawing.Size(72, 23);
			this.btnGuardar.TabIndex = 750;
			this.btnGuardar.Text = "Guardar";
			this.btnGuardar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnGuardar.Visible = false;
			this.btnGuardar.Click += new System.EventHandler(this.btnGuardar_Click);
			// 
			// btnVer
			// 
			this.btnVer.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(230)), ((System.Byte)(230)), ((System.Byte)(230)));
			this.btnVer.CausesValidation = false;
			this.btnVer.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.btnVer.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(50)), ((System.Byte)(50)), ((System.Byte)(50)));
			this.btnVer.Image = ((System.Drawing.Image)(resources.GetObject("btnVer.Image")));
			this.btnVer.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnVer.Location = new System.Drawing.Point(184, 16);
			this.btnVer.Name = "btnVer";
			this.btnVer.Size = new System.Drawing.Size(72, 23);
			this.btnVer.TabIndex = 749;
			this.btnVer.Text = "Ver";
			this.btnVer.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnVer.Click += new System.EventHandler(this.btnVer_Click);
			// 
			// btnComandos
			// 
			this.btnComandos.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btnComandos.Image = ((System.Drawing.Image)(resources.GetObject("btnComandos.Image")));
			this.btnComandos.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnComandos.Location = new System.Drawing.Point(568, 16);
			this.btnComandos.Name = "btnComandos";
			this.btnComandos.TabIndex = 977;
			this.btnComandos.Text = "Acciones";
			this.btnComandos.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnComandos.Visible = false;
			this.btnComandos.Click += new System.EventHandler(this.btnComandos_Click);
			// 
			// mnuAcciones
			// 
			this.mnuAcciones.MenuItems.AddRange(new System.Windows.Forms.MenuItem[] {
																																								this.mnuAccionesCopiarUltimoMes,
																																								this.mnuAccionesAgregarNuevoLocal});
			// 
			// mnuAccionesCopiarUltimoMes
			// 
			this.mnuAccionesCopiarUltimoMes.Index = 0;
			this.mnuAccionesCopiarUltimoMes.Text = "Copiar Ultimo Mes";
			this.mnuAccionesCopiarUltimoMes.Click += new System.EventHandler(this.mnuAccionesCopiarUltimoMes_Click);
			// 
			// mnuAccionesAgregarNuevoLocal
			// 
			this.mnuAccionesAgregarNuevoLocal.Index = 1;
			this.mnuAccionesAgregarNuevoLocal.Text = "Exportar";
			this.mnuAccionesAgregarNuevoLocal.Click += new System.EventHandler(this.mnuAccionesAgregarNuevoLocal_Click);
			// 
			// frmPorcentajeCobradoresProyectado
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(656, 470);
			this.Controls.Add(this.btnComandos);
			this.Controls.Add(this.btnGuardar);
			this.Controls.Add(this.btnVer);
			this.Controls.Add(this.groupBox1);
			this.Controls.Add(this.uGridAsignacion);
			this.Controls.Add(this.dtpPeriodo);
			this.Controls.Add(this.label4);
			this.KeyPreview = true;
			this.Name = "frmPorcentajeCobradoresProyectado";
			this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
			this.Text = "Porcentaje Cobranza";
			this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.frmPorcentajeCobradoresProyectado_KeyDown);
			this.Load += new System.EventHandler(this.frmPorcentajeCobradoresProyectado_Load);
			((System.ComponentModel.ISupportInitialize)(this.uGridAsignacion)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);
		
		}

		private Acceso miAcceso;
		
		private void UnloadMe()
		{
			this.Close();
		}
		private void frmPorcentajeCobradoresProyectado_Load(object sender, System.EventArgs e)
		{
			miAcceso = new Acceso(cdsSeteoF, "083611");

			if (!Funcion.Permiso("1847", cdsSeteoF))
			{		
				MessageBox.Show("No tiene Acceso a Porcentaje Cobranza", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				BeginInvoke(new MethodInvoker(UnloadMe));
				return;
			}

			if(miAcceso.BExportar)this.btnComandos.Visible = true;
			if(miAcceso.Nuevo)this.btnGuardar.Visible = true;

			this.dtpPeriodo.MaxDate = DateTime.Parse(DateTime.DaysInMonth(DateTime.Now.Year, DateTime.Now.Month).ToString() + "/" + DateTime.Now.Month.ToString() + "/" + DateTime.Now.Year.ToString());
			this.dtpPeriodo.Value = DateTime.Parse("01/" + DateTime.Now.Month.ToString() + "/" + DateTime.Now.Year.ToString());
		
		}

		private void uGridAsignacion_AfterRowInsert(object sender, Infragistics.Win.UltraWinGrid.RowEventArgs e)
		{
			e.Row.Cells["idPorcentajeCobradoresProyectado"].Value = 0;
			e.Row.Cells["DiaDesde"].Value = 0;
			e.Row.Cells["DiaHasta"].Value = 0;
			e.Row.Cells["Porcentaje"].Value = 0.00m; 	

		}
    bool bActualiza = false;
		private void uGridAsignacion_AfterCellUpdate(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			if (bActualiza) return;

			if (e.Cell.Column.ToString() == "DiaDesde")
			{	
				bActualiza = true;

				if ( e.Cell.Row.Cells["DiaDesde"].Value == System.DBNull.Value)
				{											
					e.Cell.Row.Cells["DiaDesde"].Value = (int)e.Cell.OriginalValue;						
				}	
				bActualiza = false;
			}

			if (e.Cell.Column.ToString() == "DiaHasta")
			{	
				bActualiza = true;

				if ( e.Cell.Row.Cells["DiaHasta"].Value == System.DBNull.Value)
				{											
					e.Cell.Row.Cells["DiaHasta"].Value = (int)e.Cell.OriginalValue;						
				}	
				bActualiza = false;
			}

			if (e.Cell.Column.ToString() == "Porcentaje")
			{	
				bActualiza = true;

				if ( e.Cell.Row.Cells["Porcentaje"].Value == System.DBNull.Value)
				{											
					e.Cell.Row.Cells["Porcentaje"].Value = (decimal)e.Cell.OriginalValue;						
				}	
				bActualiza = false;
			}

		}

		private void dtpPeriodo_ValueChanged(object sender, System.EventArgs e)
		{
			if (this.dtpPeriodo.Text.ToString().Trim().Length == 0)
			{
				MessageBox.Show("Seleccione un periodo", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.dtpPeriodo.Focus();
				return;
			}
			Consultar ();

		}
		
		private void Consultar ()
		{

      DateTime dtDesde = DateTime.Parse("01/" +  "/" + this.dtpPeriodo.Value.Month.ToString()+  "/" + this.dtpPeriodo.Value.Year.ToString());
			string sSQL = string.Format("Exec [ConsultaGeneralPorcentajeCobradoresProyectado] '{0}'", Convert.ToDateTime(dtDesde).ToString("yyyyMMdd"));
			this.uGridAsignacion.DataSource = FuncionesProcedimientos.dtGeneral(sSQL);

		}

		private void btnVer_Click(object sender, System.EventArgs e)
		{
			if (this.dtpPeriodo.Text.ToString().Trim().Length == 0)
			{
				MessageBox.Show("Seleccione un periodo", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.dtpPeriodo.Focus();
				return;
			}
			Consultar ();
		}

		private void frmPorcentajeCobradoresProyectado_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Escape) this.Close();
		}

		private void btnGuardar_Click(object sender, System.EventArgs e)
		{
			if (this.dtpPeriodo.Text.ToString().Trim().Length == 0)
			{
				MessageBox.Show("Seleccione un periodo", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.dtpPeriodo.Focus();
				return;
			}
			if ((int) this.uGridAsignacion.Rows.Count == 0)
			{
				MessageBox.Show("Ingrese los valores.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.uGridAsignacion.Focus ();
				return;
			}

			DateTime dtDesde = DateTime.Parse("01/" +  "/" + this.dtpPeriodo.Value.Month.ToString()+  "/" + this.dtpPeriodo.Value.Year.ToString());
			//DateTime dtHasta = DateTime.Parse(DateTime.DaysInMonth(DateTime.Now.Year, DateTime.Now.Month).ToString() + "/" + DateTime.Now.Month.ToString() + "/" + DateTime.Now.Year.ToString());
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridAsignacion.Rows.All)
			{
				if (dr.Cells["Porcentaje"].Value == System.DBNull.Value)
				{
					MessageBox.Show("Ingrese el porcentaje ", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					this.uGridAsignacion.ActiveRow = dr;
					this.uGridAsignacion.ActiveRow.Selected = true;
					return;
				}
				if (dr.Cells["DiaDesde"].Value == System.DBNull.Value)
				{
					MessageBox.Show("Ingrese el dia Inicial ", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					this.uGridAsignacion.ActiveRow = dr;
					this.uGridAsignacion.ActiveRow.Selected = true;
					return;
				}
				if (dr.Cells["DiaHasta"].Value == System.DBNull.Value)
				{
					MessageBox.Show("Ingrese el dia Final ", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					this.uGridAsignacion.ActiveRow = dr;
					this.uGridAsignacion.ActiveRow.Selected = true;
					return;
				}
				

				if ((int)dr.Cells["DiaDesde"].Value > (int) dr.Cells["DiaHasta"].Value  )
				{
					MessageBox.Show("El día Inicial no puede ser mayor al Final ", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					this.uGridAsignacion.ActiveRow = dr;
					this.uGridAsignacion.ActiveRow.Selected = true;
					return;
				}
				
			}
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridAsignacion.Rows.All)
			{
				string sSQL = string.Format("Exec [Com_GuardaPorcentajeCobradoresProyectado] {0}, '{1}', {2}, {3}, {4}", 
					(int)dr.Cells["idPorcentajeCobradoresProyectado"].Value, 
					Convert.ToDateTime(dtDesde).ToString("yyyyMMdd"),
					(int)dr.Cells["DiaDesde"].Value,
					(int)dr.Cells["DiaHasta"].Value,
					Convert.ToDecimal(dr.Cells["Porcentaje"].Value)
					);
				Funcion.EjecutaSQL(cdsSeteoF, sSQL);
			}

			this.btnVer_Click(sender, e);
			
			MessageBox.Show("Asignacion Guardada Correctamente", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);

		}

		private void uGridAsignacion_BeforeRowsDeleted(object sender, Infragistics.Win.UltraWinGrid.BeforeRowsDeletedEventArgs e)
		{
			if (!miAcceso.BEliminaSeriales) 
			{
				e.Cancel = true;
			}
			else
			{
				if (DialogResult.Yes ==	MessageBox.Show("¿Esta seguro de Borrar las Líneas Seleccionadas?", "Point Technology", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2))				
				{
					for (int i = 0; i < e.Rows.Length; i++)
					{
						if (e.Rows[i].Cells["idPorcentajeCobradoresProyectado"].Value != System.DBNull.Value)
						{									
							if ((int)e.Rows[i].Cells["idPorcentajeCobradoresProyectado"].Value > 0)
							{
								string sSQL = string.Format("Delete From PorcentajeCobradoresProyectado Where idPorcentajeCobradoresProyectado = {0}", (int)e.Rows[i].Cells["idPorcentajeCobradoresProyectado"].Value);
								Funcion.EjecutaSQL(cdsSeteoF, sSQL, true);
							}
						}			
					}
					e.DisplayPromptMsg = false;
				}
				else
				{
					e.Cancel = true;
				}
			
			}
		}

		private void btnComandos_Click(object sender, System.EventArgs e)
		{
		this.mnuAcciones.Show(this.btnComandos, new Point(1, 25));
		}

		private void mnuAccionesCopiarUltimoMes_Click(object sender, System.EventArgs e)
		{
			int iRetorna = 0;
			if (this.dtpPeriodo.Text.ToString().Trim().Length == 0)
			{
				MessageBox.Show("Seleccione un periodo", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.dtpPeriodo.Focus();
				return;
			}

			string sPeriodo = Funcion.sEscalarSQL(cdsSeteoF, "Select top 1 Periodo from dbo.PorcentajeCobradoresProyectado order by  idPorcentajeCobradoresProyectado desc");

			if (DialogResult.Yes == MessageBox.Show("¿Esta Seguro De Copiar La Información del anterior Mes?", "Point Technology", MessageBoxButtons.YesNo, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2))
			{	

				iRetorna = Funcion.iEscalarSQL(cdsSeteoF, string.Format("Exec CopiaAsignacioncobradoresPorcentajeCobranza"));
	

				MessageBox.Show(string.Format("Se copio la Información del Periodo {0} correctamente", sPeriodo), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
				//this.btnVer_Click(sender, e);
				
			}
		}

		private void mnuAccionesAgregarNuevoLocal_Click(object sender, System.EventArgs e)
		{
		FuncionesProcedimientos.ExportaExcel(this.uGridAsignacion);
		}
	}
}
