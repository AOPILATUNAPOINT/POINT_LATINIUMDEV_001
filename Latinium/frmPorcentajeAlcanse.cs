using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de frmPorcentajeAlcanse.
	/// </summary>
	public class frmPorcentajeAlcanse : System.Windows.Forms.Form
	{
		private System.Windows.Forms.GroupBox groupBox1;
		private Infragistics.Win.UltraWinGrid.UltraGrid uGridCarpa;
		private C1.Data.C1DataSet cdsSeteoF;
		private System.Windows.Forms.Label label4;
		public Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtFecha;
		private System.Windows.Forms.Button btnVer;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource1;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public frmPorcentajeAlcanse()
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
		private Acceso miAcceso;

		private void UnloadMe()
		{
			this.Close();
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
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Descripcion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Meta");
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Venta");
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
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Descripcion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Meta");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn3 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Venta");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn4 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Porcentaje");
			Infragistics.Win.Appearance appearance12 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton1 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmPorcentajeAlcanse));
			this.groupBox1 = new System.Windows.Forms.GroupBox();
			this.uGridCarpa = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource1 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.label4 = new System.Windows.Forms.Label();
			this.dtFecha = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.btnVer = new System.Windows.Forms.Button();
			((System.ComponentModel.ISupportInitialize)(this.uGridCarpa)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFecha)).BeginInit();
			this.SuspendLayout();
			// 
			// groupBox1
			// 
			this.groupBox1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.groupBox1.Location = new System.Drawing.Point(-256, 32);
			this.groupBox1.Name = "groupBox1";
			this.groupBox1.Size = new System.Drawing.Size(1360, 8);
			this.groupBox1.TabIndex = 736;
			this.groupBox1.TabStop = false;
			// 
			// uGridCarpa
			// 
			this.uGridCarpa.Cursor = System.Windows.Forms.Cursors.Default;
			this.uGridCarpa.DataSource = this.ultraDataSource1;
			appearance1.BackColor = System.Drawing.Color.White;
			this.uGridCarpa.DisplayLayout.Appearance = appearance1;
			this.uGridCarpa.DisplayLayout.AutoFitColumns = true;
			ultraGridColumn1.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn1.Header.Caption = "Descripción";
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn1.Width = 292;
			ultraGridColumn2.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance2.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn2.CellAppearance = appearance2;
			ultraGridColumn2.Format = "#,##0.00";
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridColumn2.Hidden = true;
			ultraGridColumn2.Width = 62;
			ultraGridColumn3.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance3.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn3.CellAppearance = appearance3;
			ultraGridColumn3.Format = "#,##0.00";
			ultraGridColumn3.Header.VisiblePosition = 2;
			ultraGridColumn3.Hidden = true;
			ultraGridColumn3.Width = 62;
			ultraGridColumn4.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance4.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn4.CellAppearance = appearance4;
			ultraGridColumn4.Format = "#,##0.00";
			ultraGridColumn4.Header.VisiblePosition = 3;
			ultraGridColumn4.Width = 239;
			ultraGridBand1.Columns.AddRange(new object[] {
																										 ultraGridColumn1,
																										 ultraGridColumn2,
																										 ultraGridColumn3,
																										 ultraGridColumn4});
			ultraGridBand1.Header.Caption = "Ventas";
			ultraGridBand1.HeaderVisible = true;
			ultraGridBand1.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.Edit;
			ultraGridBand1.SummaryFooterCaption = "TOTALES";
			this.uGridCarpa.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			appearance5.ForeColor = System.Drawing.Color.Black;
			appearance5.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridCarpa.DisplayLayout.Override.ActiveRowAppearance = appearance5;
			this.uGridCarpa.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
			this.uGridCarpa.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.uGridCarpa.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance6.BackColor = System.Drawing.Color.Transparent;
			this.uGridCarpa.DisplayLayout.Override.CardAreaAppearance = appearance6;
			appearance7.ForeColor = System.Drawing.Color.Black;
			appearance7.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridCarpa.DisplayLayout.Override.CellAppearance = appearance7;
			this.uGridCarpa.DisplayLayout.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.Edit;
			appearance8.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance8.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance8.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance8.FontData.BoldAsString = "True";
			appearance8.FontData.Name = "Arial";
			appearance8.FontData.SizeInPoints = 8F;
			appearance8.ForeColor = System.Drawing.Color.White;
			appearance8.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridCarpa.DisplayLayout.Override.HeaderAppearance = appearance8;
			appearance9.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance9.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance9.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridCarpa.DisplayLayout.Override.RowAlternateAppearance = appearance9;
			appearance10.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance10.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance10.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridCarpa.DisplayLayout.Override.RowSelectorAppearance = appearance10;
			appearance11.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance11.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance11.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridCarpa.DisplayLayout.Override.SelectedRowAppearance = appearance11;
			this.uGridCarpa.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridCarpa.Location = new System.Drawing.Point(8, 48);
			this.uGridCarpa.Name = "uGridCarpa";
			this.uGridCarpa.Size = new System.Drawing.Size(552, 144);
			this.uGridCarpa.TabIndex = 740;
			this.uGridCarpa.InitializeRow += new Infragistics.Win.UltraWinGrid.InitializeRowEventHandler(this.uGridCarpa_InitializeRow);
			// 
			// ultraDataSource1
			// 
			ultraDataColumn2.DataType = typeof(System.Decimal);
			ultraDataColumn3.DataType = typeof(System.Decimal);
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
			// label4
			// 
			this.label4.AutoSize = true;
			this.label4.Location = new System.Drawing.Point(8, 8);
			this.label4.Name = "label4";
			this.label4.Size = new System.Drawing.Size(43, 16);
			this.label4.TabIndex = 865;
			this.label4.Text = "Periodo";
			this.label4.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// dtFecha
			// 
			appearance12.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtFecha.Appearance = appearance12;
			this.dtFecha.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton1.Caption = "Today";
			this.dtFecha.DateButtons.Add(dateButton1);
			this.dtFecha.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.dtFecha.Format = "MM/yyyy ";
			this.dtFecha.Location = new System.Drawing.Point(64, 8);
			this.dtFecha.Name = "dtFecha";
			this.dtFecha.NonAutoSizeHeight = 23;
			this.dtFecha.Size = new System.Drawing.Size(104, 21);
			this.dtFecha.SpinButtonsVisible = true;
			this.dtFecha.TabIndex = 864;
			this.dtFecha.Value = ((object)(resources.GetObject("dtFecha.Value")));
			// 
			// btnVer
			// 
			this.btnVer.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(230)), ((System.Byte)(230)), ((System.Byte)(230)));
			this.btnVer.CausesValidation = false;
			this.btnVer.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.btnVer.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(50)), ((System.Byte)(50)), ((System.Byte)(50)));
			this.btnVer.Image = ((System.Drawing.Image)(resources.GetObject("btnVer.Image")));
			this.btnVer.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnVer.Location = new System.Drawing.Point(240, 8);
			this.btnVer.Name = "btnVer";
			this.btnVer.Size = new System.Drawing.Size(72, 23);
			this.btnVer.TabIndex = 866;
			this.btnVer.Text = "Ver";
			this.btnVer.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnVer.Click += new System.EventHandler(this.btnVer_Click);
			// 
			// frmPorcentajeAlcanse
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(568, 198);
			this.Controls.Add(this.btnVer);
			this.Controls.Add(this.label4);
			this.Controls.Add(this.dtFecha);
			this.Controls.Add(this.uGridCarpa);
			this.Controls.Add(this.groupBox1);
			this.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.KeyPreview = true;
			this.MaximizeBox = false;
			this.MinimizeBox = false;
			this.Name = "frmPorcentajeAlcanse";
			this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
			this.Text = "Avance  General - Credito";
			this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.frmPorcentajeAlcanse_KeyDown);
			this.Load += new System.EventHandler(this.frmPorcentajeAlcanse_Load);
			((System.ComponentModel.ISupportInitialize)(this.uGridCarpa)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFecha)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void frmPorcentajeAlcanse_Load(object sender, System.EventArgs e)
		{
			miAcceso = new Acceso(cdsSeteoF, "0226");

			if (!Funcion.Permiso("1412", cdsSeteoF))
			{				
				MessageBox.Show("No tiene Acceso a Ventas Generales", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				BeginInvoke(new MethodInvoker(UnloadMe));
				return;
			}

			if (miAcceso.BExportar)
			{
				this.uGridCarpa.DisplayLayout.Bands[0].Columns["Meta"].Hidden = false;
				this.uGridCarpa.DisplayLayout.Bands[0].Columns["Venta"].Hidden = false;
			}
			else
			{
				this.uGridCarpa.DisplayLayout.Bands[0].Columns["Meta"].Hidden = true;
				this.uGridCarpa.DisplayLayout.Bands[0].Columns["Venta"].Hidden = true;

			}

			this.dtFecha.CalendarInfo.MinDate = DateTime.Parse("01/01/2021");
			this.dtFecha.CalendarInfo.MaxDate = DateTime.Today;		
			this.dtFecha.Value = DateTime.Today;

			int	iMesA = Convert.ToInt32( ( Convert.ToDateTime(this.dtFecha.Value).Month));
			int iYearA = Convert.ToInt32( ( Convert.ToDateTime(this.dtFecha.Value).Year));
			DateTime dtInicio = DateTime.Parse("01/" + iMesA + "/" + iYearA);

			DateTime dtFinalPeriodo = DateTime.Parse(DateTime.DaysInMonth(dtInicio.Date.Year, dtInicio.Month).ToString() + "/" + dtInicio.Month.ToString() + "/" + dtInicio.Year.ToString());

			this.btnVer_Click (sender, e);
		}

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);
		}

		private void btnVer_Click(object sender, System.EventArgs e)
		{
			if (!Validacion.vbFechaEnDocumentos(this.dtFecha, "Seleccione la Fecha De Salida", false, "", cdsSeteoF)) return;

			int iMesM = Convert.ToInt32( ( Convert.ToDateTime(this.dtFecha.Value).Month));
			int iYearM = Convert.ToInt32( ( Convert.ToDateTime(this.dtFecha.Value).Year));

			DateTime dtFinalPeriodo = DateTime.Parse(DateTime.DaysInMonth(( Convert.ToDateTime(this.dtFecha.Value).Year), ( Convert.ToDateTime(this.dtFecha.Value).Month)).ToString() + "/" + iMesM + "/" + iYearM );
			 
			string sSQLPa = string.Format("Exec [frmPorcentajeAlcanse] '{0}'", 
				 Convert.ToDateTime(dtFinalPeriodo).ToString("yyyyMMdd"));

			FuncionesProcedimientos.dsGeneralModulo(sSQLPa, this.uGridCarpa);
		
		}

		private void frmPorcentajeAlcanse_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
				if (e.KeyCode == Keys.Escape) this.Close();	
		}


		private void uGridCarpa_InitializeRow(object sender, Infragistics.Win.UltraWinGrid.InitializeRowEventArgs e)
		{
			if (Convert.ToDecimal(e.Row.Cells["Porcentaje"].Value) >= 100.00m) e.Row.Appearance.BackColor = Color.LimeGreen;
			else if (Convert.ToDecimal(e.Row.Cells["Porcentaje"].Value) >= 85.00m && Convert.ToDecimal(e.Row.Cells["Porcentaje"].Value) < 100.00m) e.Row.Appearance.BackColor = Color.Yellow;
			else if (Convert.ToDecimal(e.Row.Cells["Porcentaje"].Value) >= 70.00m && Convert.ToDecimal(e.Row.Cells["Porcentaje"].Value) < 85.00m) e.Row.Appearance.BackColor = Color.Orange;			
			else if (Convert.ToDecimal(e.Row.Cells["Porcentaje"].Value) < 70.00m) e.Row.Appearance.BackColor = Color.Red;
		}

	}
}
