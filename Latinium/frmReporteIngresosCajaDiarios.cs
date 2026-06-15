using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using Infragistics.Win.UltraWinGrid;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de frmReporteIngresosCajaDiarios.
	/// </summary>
	public class frmReporteIngresosCajaDiarios : System.Windows.Forms.Form
	{
		private System.Windows.Forms.Button btnBuscar;
		private Infragistics.Win.UltraWinGrid.UltraGrid uGridInformacion;
		private System.Windows.Forms.Label label4;
		private System.Windows.Forms.Label label1;
		private System.Windows.Forms.GroupBox groupBox1;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtDesde;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtHasta;
		private Infragistics.Win.UltraWinEditors.UltraOptionSet optOpciones;
		private System.Windows.Forms.Button btnExcel;
		private System.Windows.Forms.Label lblContador;
		private C1.Data.C1DataSet cdsSeteoF;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public frmReporteIngresosCajaDiarios()
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
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmReporteIngresosCajaDiarios));
			Infragistics.Win.Appearance appearance1 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton1 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton2 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.ValueListItem valueListItem1 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem2 = new Infragistics.Win.ValueListItem();
			this.btnBuscar = new System.Windows.Forms.Button();
			this.uGridInformacion = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.label4 = new System.Windows.Forms.Label();
			this.dtDesde = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.label1 = new System.Windows.Forms.Label();
			this.dtHasta = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.groupBox1 = new System.Windows.Forms.GroupBox();
			this.optOpciones = new Infragistics.Win.UltraWinEditors.UltraOptionSet();
			this.btnExcel = new System.Windows.Forms.Button();
			this.lblContador = new System.Windows.Forms.Label();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			((System.ComponentModel.ISupportInitialize)(this.uGridInformacion)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtDesde)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtHasta)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.optOpciones)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			this.SuspendLayout();
			// 
			// btnBuscar
			// 
			this.btnBuscar.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.btnBuscar.Image = ((System.Drawing.Image)(resources.GetObject("btnBuscar.Image")));
			this.btnBuscar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnBuscar.ImeMode = System.Windows.Forms.ImeMode.NoControl;
			this.btnBuscar.Location = new System.Drawing.Point(392, 7);
			this.btnBuscar.Name = "btnBuscar";
			this.btnBuscar.Size = new System.Drawing.Size(77, 23);
			this.btnBuscar.TabIndex = 2;
			this.btnBuscar.Text = "&Buscar";
			this.btnBuscar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnBuscar.Click += new System.EventHandler(this.btnBuscar_Click);
			// 
			// uGridInformacion
			// 
			this.uGridInformacion.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.uGridInformacion.Cursor = System.Windows.Forms.Cursors.Default;
			appearance1.BackColor = System.Drawing.Color.White;
			this.uGridInformacion.DisplayLayout.Appearance = appearance1;
			this.uGridInformacion.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
			this.uGridInformacion.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.uGridInformacion.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.False;
			appearance2.BackColor = System.Drawing.Color.Transparent;
			this.uGridInformacion.DisplayLayout.Override.CardAreaAppearance = appearance2;
			appearance3.ForeColor = System.Drawing.Color.Black;
			appearance3.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridInformacion.DisplayLayout.Override.CellAppearance = appearance3;
			appearance4.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance4.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance4.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance4.FontData.BoldAsString = "True";
			appearance4.FontData.Name = "Arial";
			appearance4.FontData.SizeInPoints = 8.5F;
			appearance4.ForeColor = System.Drawing.Color.White;
			appearance4.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridInformacion.DisplayLayout.Override.HeaderAppearance = appearance4;
			this.uGridInformacion.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti;
			appearance5.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance5.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance5.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridInformacion.DisplayLayout.Override.RowSelectorAppearance = appearance5;
			appearance6.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance6.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance6.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridInformacion.DisplayLayout.Override.SelectedRowAppearance = appearance6;
			this.uGridInformacion.Font = new System.Drawing.Font("Tahoma", 8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridInformacion.Location = new System.Drawing.Point(8, 56);
			this.uGridInformacion.Name = "uGridInformacion";
			this.uGridInformacion.Size = new System.Drawing.Size(1168, 264);
			this.uGridInformacion.TabIndex = 665;
			// 
			// label4
			// 
			this.label4.AutoSize = true;
			this.label4.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label4.Location = new System.Drawing.Point(8, 10);
			this.label4.Name = "label4";
			this.label4.Size = new System.Drawing.Size(35, 17);
			this.label4.TabIndex = 671;
			this.label4.Text = "Desde";
			this.label4.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// dtDesde
			// 
			this.dtDesde.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton1.Caption = "Today";
			this.dtDesde.DateButtons.Add(dateButton1);
			this.dtDesde.Format = "dd/MM/yyyy";
			this.dtDesde.Location = new System.Drawing.Point(64, 8);
			this.dtDesde.Name = "dtDesde";
			this.dtDesde.NonAutoSizeHeight = 23;
			this.dtDesde.Size = new System.Drawing.Size(112, 21);
			this.dtDesde.SpinButtonsVisible = true;
			this.dtDesde.TabIndex = 670;
			this.dtDesde.Value = ((object)(resources.GetObject("dtDesde.Value")));
			// 
			// label1
			// 
			this.label1.AutoSize = true;
			this.label1.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label1.Location = new System.Drawing.Point(200, 10);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(33, 17);
			this.label1.TabIndex = 673;
			this.label1.Text = "Hasta";
			this.label1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// dtHasta
			// 
			this.dtHasta.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton2.Caption = "Today";
			this.dtHasta.DateButtons.Add(dateButton2);
			this.dtHasta.Format = "dd/MM/yyyy";
			this.dtHasta.Location = new System.Drawing.Point(256, 8);
			this.dtHasta.Name = "dtHasta";
			this.dtHasta.NonAutoSizeHeight = 23;
			this.dtHasta.Size = new System.Drawing.Size(112, 21);
			this.dtHasta.SpinButtonsVisible = true;
			this.dtHasta.TabIndex = 672;
			this.dtHasta.Value = ((object)(resources.GetObject("dtHasta.Value")));
			// 
			// groupBox1
			// 
			this.groupBox1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.groupBox1.Location = new System.Drawing.Point(0, 40);
			this.groupBox1.Name = "groupBox1";
			this.groupBox1.Size = new System.Drawing.Size(1192, 8);
			this.groupBox1.TabIndex = 674;
			this.groupBox1.TabStop = false;
			// 
			// optOpciones
			// 
			appearance7.ForeColorDisabled = System.Drawing.Color.Black;
			this.optOpciones.Appearance = appearance7;
			this.optOpciones.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.optOpciones.CheckedIndex = 0;
			this.optOpciones.ItemAppearance = appearance8;
			this.optOpciones.ItemOrigin = new System.Drawing.Point(10, 0);
			valueListItem1.DataValue = 1;
			valueListItem1.DisplayText = "Bodega";
			valueListItem2.DataValue = 2;
			valueListItem2.DisplayText = "Fecha";
			this.optOpciones.Items.Add(valueListItem1);
			this.optOpciones.Items.Add(valueListItem2);
			this.optOpciones.ItemSpacingVertical = 10;
			this.optOpciones.Location = new System.Drawing.Point(488, 6);
			this.optOpciones.Name = "optOpciones";
			this.optOpciones.Size = new System.Drawing.Size(144, 24);
			this.optOpciones.TabIndex = 675;
			this.optOpciones.Text = "Bodega";
			this.optOpciones.ValueChanged += new System.EventHandler(this.optOpciones_ValueChanged);
			// 
			// btnExcel
			// 
			this.btnExcel.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btnExcel.Image = ((System.Drawing.Image)(resources.GetObject("btnExcel.Image")));
			this.btnExcel.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnExcel.Location = new System.Drawing.Point(1088, 6);
			this.btnExcel.Name = "btnExcel";
			this.btnExcel.Size = new System.Drawing.Size(88, 24);
			this.btnExcel.TabIndex = 677;
			this.btnExcel.Text = "Excel";
			this.btnExcel.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnExcel.Click += new System.EventHandler(this.btnExcel_Click);
			// 
			// lblContador
			// 
			this.lblContador.AutoSize = true;
			this.lblContador.Font = new System.Drawing.Font("Tahoma", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblContador.ForeColor = System.Drawing.Color.Firebrick;
			this.lblContador.Location = new System.Drawing.Point(664, 10);
			this.lblContador.Name = "lblContador";
			this.lblContador.Size = new System.Drawing.Size(0, 20);
			this.lblContador.TabIndex = 678;
			this.lblContador.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
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
			// frmReporteIngresosCajaDiarios
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(1184, 326);
			this.Controls.Add(this.lblContador);
			this.Controls.Add(this.btnExcel);
			this.Controls.Add(this.optOpciones);
			this.Controls.Add(this.groupBox1);
			this.Controls.Add(this.label1);
			this.Controls.Add(this.label4);
			this.Controls.Add(this.dtHasta);
			this.Controls.Add(this.dtDesde);
			this.Controls.Add(this.uGridInformacion);
			this.Controls.Add(this.btnBuscar);
			this.KeyPreview = true;
			this.Name = "frmReporteIngresosCajaDiarios";
			this.Text = "Ingresos Diarios de Caja";
			this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.frmReporteIngresosCajaDiarios_KeyDown);
			this.Load += new System.EventHandler(this.frmReporteIngresosCajaDiarios_Load);
			((System.ComponentModel.ISupportInitialize)(this.uGridInformacion)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtDesde)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtHasta)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.optOpciones)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void Consulta()
		{
			string sSQL = string.Format("Exec ResumenIngresoDiarioCajas '{0}', '{1}'", Convert.ToDateTime(this.dtDesde.Value).ToString("yyyyMMdd"), Convert.ToDateTime(this.dtHasta.Value).ToString("yyyyMMdd"));

			if ((int)this.optOpciones.Value == 2) 
				sSQL = string.Format("Exec ResumenDiarioCajasJulio '{0}', '{1}'", Convert.ToDateTime(this.dtDesde.Value).ToString("yyyyMMdd"), Convert.ToDateTime(this.dtHasta.Value).ToString("yyyyMMdd"));

			this.uGridInformacion.DataSource = null;

			FuncionesProcedimientos.dsGeneral(sSQL, this.uGridInformacion);

			this.lblContador.Text = this.uGridInformacion.Rows.Count + " REGISTROS ENCONTRADOS";

			if ((int)this.optOpciones.Value == 1) DiseñoGrid();	else DiseñoGrid1();
		}

		private void btnBuscar_Click(object sender, System.EventArgs e)
		{
			if (!Validacion.vbComparaFechas(Convert.ToDateTime(this.dtDesde.Value), Convert.ToDateTime(this.dtHasta.Value), this.dtDesde, "Fechas erroneas", false)) return;
			
			Consulta();
		}

		private void frmReporteIngresosCajaDiarios_Load(object sender, System.EventArgs e)
		{
			miAcceso = new Acceso(cdsSeteoF, "044402");

			if (!Funcion.Permiso("924", cdsSeteoF))
			{				
				MessageBox.Show("No puede ingresar a Ingresos Caja Diarios ", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);			
				BeginInvoke(new MethodInvoker(UnloadMe));
				return;
			}
			this.dtDesde.Value = "01/" + DateTime.Now.Month.ToString() + "/" + DateTime.Now.Year.ToString();
			this.dtHasta.Value = DateTime.Today;			
		}

		private void DiseñoGrid1()
		{
			if (this.uGridInformacion.Rows.Count > 0)
			{				
				this.uGridInformacion.DisplayLayout.Bands[0].Columns["Caja"].Width = 120;								
				this.uGridInformacion.DisplayLayout.Bands[0].Columns["Caja"].CellActivation = Activation.ActivateOnly;	
		
				
				this.uGridInformacion.DisplayLayout.Bands[0].Columns["Total"].Width = 80;								
				this.uGridInformacion.DisplayLayout.Bands[0].Columns["Total"].CellActivation = Activation.Disabled;
				this.uGridInformacion.DisplayLayout.Bands[0].Columns["Total"].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right;
				this.uGridInformacion.DisplayLayout.Bands[0].Columns["Total"].CellAppearance.BackColor=Color.Silver;
				this.uGridInformacion.DisplayLayout.Bands[0].Summaries.Add("Total", Infragistics.Win.UltraWinGrid.SummaryType.Sum, this.uGridInformacion.DisplayLayout.Bands[0].Columns["Total"], Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn);
				this.uGridInformacion.DisplayLayout.Bands[0].Summaries["Total"].Appearance.TextHAlign = Infragistics.Win.HAlign.Right;
				this.uGridInformacion.DisplayLayout.Bands[0].Summaries["Total"].DisplayFormat = "{0: #,##0.00}";

				this.uGridInformacion.DisplayLayout.Bands[0].Columns["Promedio"].Width = 80;								
				this.uGridInformacion.DisplayLayout.Bands[0].Columns["Promedio"].CellActivation = Activation.Disabled;				
				this.uGridInformacion.DisplayLayout.Bands[0].Columns["Promedio"].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right;
				this.uGridInformacion.DisplayLayout.Bands[0].Columns["Promedio"].CellAppearance.BackColor=Color.Silver;
				this.uGridInformacion.DisplayLayout.Bands[0].Summaries.Add("Promedio", Infragistics.Win.UltraWinGrid.SummaryType.Sum, this.uGridInformacion.DisplayLayout.Bands[0].Columns["Promedio"], Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn);
				this.uGridInformacion.DisplayLayout.Bands[0].Summaries["Promedio"].Appearance.TextHAlign = Infragistics.Win.HAlign.Right;
				this.uGridInformacion.DisplayLayout.Bands[0].Summaries["Promedio"].DisplayFormat = "{0: #,##0.00}";

				int iContColumnas = (int)this.uGridInformacion.DisplayLayout.Bands[0].Columns.Count;

				int iContador = 3;

				string nombrCol = "";

				for (int j = iContador; j < iContColumnas; j++)
				{				
					nombrCol = this.uGridInformacion.DisplayLayout.Bands[0].Columns[j].Header.Caption;
					
					
					this.uGridInformacion.DisplayLayout.Bands[0].Columns[j].Width = 70;
					this.uGridInformacion.DisplayLayout.Bands[0].Columns[j].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right;
					this.uGridInformacion.DisplayLayout.Bands[0].Columns[j].CellActivation = Activation.Disabled;

					this.uGridInformacion.DisplayLayout.Bands[0].Summaries.Add(nombrCol, Infragistics.Win.UltraWinGrid.SummaryType.Sum, this.uGridInformacion.DisplayLayout.Bands[0].Columns[nombrCol], Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn);
					this.uGridInformacion.DisplayLayout.Bands[0].Summaries[nombrCol].Appearance.TextHAlign = Infragistics.Win.HAlign.Right;
					this.uGridInformacion.DisplayLayout.Bands[0].Summaries[nombrCol].DisplayFormat = "{0: #,##0.00}";

				}
				string sSuma = "";
				foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridInformacion.Rows.All)
				{
					int iCol = 0;

					for (int j = 3; j < iContColumnas; j++)
					{
						sSuma = this.uGridInformacion.DisplayLayout.Bands[0].Columns[j].Header.Caption;
						if (dr.Cells[sSuma].Value != System.DBNull.Value) 
							dr.Cells["Total"].Value = Convert.ToDecimal(dr.Cells["Total"].Value) + Convert.ToDecimal(dr.Cells[sSuma].Value);

						iCol++;
					}

					dr.Cells["Promedio"].Value = decimal.Round(Convert.ToDecimal(dr.Cells["Total"].Value) / iCol, 2);
				}
			}
		}

		private void DiseñoGrid()
		{
			if (this.uGridInformacion.Rows.Count > 0)
			{				
				this.uGridInformacion.DisplayLayout.Bands[0].Columns["Fecha"].Width = 70;								
				this.uGridInformacion.DisplayLayout.Bands[0].Columns["Fecha"].CellActivation = Activation.ActivateOnly;				

				this.uGridInformacion.DisplayLayout.Bands[0].Columns["Total"].Width = 80;								
				this.uGridInformacion.DisplayLayout.Bands[0].Columns["Total"].CellActivation = Activation.Disabled;
				this.uGridInformacion.DisplayLayout.Bands[0].Columns["Total"].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right;
				this.uGridInformacion.DisplayLayout.Bands[0].Columns["Total"].CellAppearance.BackColor=Color.Silver;
				this.uGridInformacion.DisplayLayout.Bands[0].Summaries.Add("Total", Infragistics.Win.UltraWinGrid.SummaryType.Sum, this.uGridInformacion.DisplayLayout.Bands[0].Columns["Total"], Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn);
				this.uGridInformacion.DisplayLayout.Bands[0].Summaries["Total"].Appearance.TextHAlign = Infragistics.Win.HAlign.Right;
				this.uGridInformacion.DisplayLayout.Bands[0].Summaries["Total"].DisplayFormat = "{0: #,##0.00}";

				this.uGridInformacion.DisplayLayout.Bands[0].Columns["Promedio"].Width = 80;								
				this.uGridInformacion.DisplayLayout.Bands[0].Columns["Promedio"].CellActivation = Activation.Disabled;				
				this.uGridInformacion.DisplayLayout.Bands[0].Columns["Promedio"].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right;
				this.uGridInformacion.DisplayLayout.Bands[0].Columns["Promedio"].CellAppearance.BackColor=Color.Silver;
				this.uGridInformacion.DisplayLayout.Bands[0].Summaries.Add("Promedio", Infragistics.Win.UltraWinGrid.SummaryType.Sum, this.uGridInformacion.DisplayLayout.Bands[0].Columns["Promedio"], Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn);
				this.uGridInformacion.DisplayLayout.Bands[0].Summaries["Promedio"].Appearance.TextHAlign = Infragistics.Win.HAlign.Right;
				this.uGridInformacion.DisplayLayout.Bands[0].Summaries["Promedio"].DisplayFormat = "{0: #,##0.00}";

				int iContColumnas = (int)this.uGridInformacion.DisplayLayout.Bands[0].Columns.Count;

				int iContador = 3;

				string nombrCol = "";

				for (int j = iContador; j < iContColumnas; j++)
				{				
					nombrCol = this.uGridInformacion.DisplayLayout.Bands[0].Columns[j].Header.Caption;
					
					
					this.uGridInformacion.DisplayLayout.Bands[0].Columns[j].Width = 70;
					this.uGridInformacion.DisplayLayout.Bands[0].Columns[j].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right;
					this.uGridInformacion.DisplayLayout.Bands[0].Columns[j].CellActivation = Activation.Disabled;
					

					this.uGridInformacion.DisplayLayout.Bands[0].Summaries.Add(nombrCol, Infragistics.Win.UltraWinGrid.SummaryType.Sum, this.uGridInformacion.DisplayLayout.Bands[0].Columns[nombrCol], Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn);
					this.uGridInformacion.DisplayLayout.Bands[0].Summaries[nombrCol].Appearance.TextHAlign = Infragistics.Win.HAlign.Right;
					this.uGridInformacion.DisplayLayout.Bands[0].Summaries[nombrCol].DisplayFormat = "{0: #,##0.00}";

				}

				string sSuma = "";
				foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridInformacion.Rows.All)
				{
					int iCol = 0;

					for (int j = 3; j < iContColumnas; j++)
					{
						sSuma = this.uGridInformacion.DisplayLayout.Bands[0].Columns[j].Header.Caption;
						if (dr.Cells[sSuma].Value != System.DBNull.Value) 
							dr.Cells["Total"].Value = Convert.ToDecimal(dr.Cells["Total"].Value) + Convert.ToDecimal(dr.Cells[sSuma].Value);

						iCol++;
					}

					dr.Cells["Promedio"].Value = decimal.Round(Convert.ToDecimal(dr.Cells["Total"].Value) / iCol, 2);
				}
			}
		}

		private void optOpciones_ValueChanged(object sender, System.EventArgs e)
		{		
			Consulta();
		}

		private void btnExcel_Click(object sender, System.EventArgs e)
		{
			FuncionesProcedimientos.ExportaExcel(this.uGridInformacion);
		}
		
		private void frmReporteIngresosCajaDiarios_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Escape) this.Close();
		}

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);	
		}	
	}
}

