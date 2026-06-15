using System;
using System.Globalization;
using System.Data;
using System.Data.SqlClient;
using C1.Data;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using Infragistics.Win.UltraWinGrid;
using System.Web.Mail;
using CrystalDecisions.Shared;
using System.Runtime.InteropServices;
using System.Diagnostics;
using System.IO;
using CrystalDecisions;
using CrystalDecisions.CrystalReports.Engine;
using CrystalDecisions.ReportSource;
using System.Net;
using System.Web;
using System.Xml;
using System.Xml.Schema;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de frmConfiTransporte.
	/// </summary>
	public class frmConfiTransporte : System.Windows.Forms.Form
	{
		private Infragistics.Win.UltraWinGrid.UltraDropDown cmbBodega;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource1;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtFecha;
		private C1.Data.C1DataSet cdsSeteoF;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource2;
		public System.Windows.Forms.Button btnGuardar;
		private System.Windows.Forms.Button btnSalir;
		private Infragistics.Win.UltraWinGrid.UltraGrid uGridDetalle;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public frmConfiTransporte()
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
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idTransporteValores");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Bodega");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn3 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Dias");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn4 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Semanas");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn5 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Salidas");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn6 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Valor");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bodega");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn7 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Bodega");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn8 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn9 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Codigo");
			Infragistics.Win.Appearance appearance1 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton1 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmConfiTransporte));
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idTransporteValores");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bodega", -1, "cmbBodega");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Dias");
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Semanas");
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Salidas");
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Valor");
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings1 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Valor", 5, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Valor", 5, true);
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance11 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance12 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance13 = new Infragistics.Win.Appearance();
			this.ultraDataSource2 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.cmbBodega = new Infragistics.Win.UltraWinGrid.UltraDropDown();
			this.ultraDataSource1 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.dtFecha = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.btnGuardar = new System.Windows.Forms.Button();
			this.btnSalir = new System.Windows.Forms.Button();
			this.uGridDetalle = new Infragistics.Win.UltraWinGrid.UltraGrid();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbBodega)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFecha)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridDetalle)).BeginInit();
			this.SuspendLayout();
			// 
			// ultraDataSource2
			// 
			ultraDataColumn1.DataType = typeof(int);
			ultraDataColumn1.DefaultValue = 0;
			ultraDataColumn2.DataType = typeof(int);
			ultraDataColumn2.DefaultValue = 0;
			ultraDataColumn3.DataType = typeof(int);
			ultraDataColumn3.DefaultValue = 0;
			ultraDataColumn4.DataType = typeof(int);
			ultraDataColumn4.DefaultValue = 0;
			ultraDataColumn5.DataType = typeof(int);
			ultraDataColumn5.DefaultValue = 0;
			ultraDataColumn6.DataType = typeof(System.Decimal);
			this.ultraDataSource2.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn1,
																																 ultraDataColumn2,
																																 ultraDataColumn3,
																																 ultraDataColumn4,
																																 ultraDataColumn5,
																																 ultraDataColumn6});
			this.ultraDataSource2.CellDataRequested += new Infragistics.Win.UltraWinDataSource.CellDataRequestedEventHandler(this.ultraDataSource2_CellDataRequested);
			// 
			// cmbBodega
			// 
			this.cmbBodega.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbBodega.DataSource = this.ultraDataSource1;
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridColumn2.Width = 150;
			ultraGridColumn3.Header.VisiblePosition = 2;
			ultraGridBand1.Columns.AddRange(new object[] {
																										 ultraGridColumn1,
																										 ultraGridColumn2,
																										 ultraGridColumn3});
			this.cmbBodega.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.cmbBodega.DisplayMember = "Nombre";
			this.cmbBodega.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbBodega.Location = new System.Drawing.Point(224, 168);
			this.cmbBodega.Name = "cmbBodega";
			this.cmbBodega.Size = new System.Drawing.Size(232, 96);
			this.cmbBodega.TabIndex = 986;
			this.cmbBodega.ValueMember = "Bodega";
			this.cmbBodega.Visible = false;
			// 
			// ultraDataSource1
			// 
			ultraDataColumn7.DataType = typeof(int);
			this.ultraDataSource1.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn7,
																																 ultraDataColumn8,
																																 ultraDataColumn9});
			// 
			// dtFecha
			// 
			appearance1.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtFecha.Appearance = appearance1;
			this.dtFecha.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton1.Caption = "Today";
			this.dtFecha.DateButtons.Add(dateButton1);
			this.dtFecha.Enabled = false;
			this.dtFecha.Format = "dd/MM/yyyy";
			this.dtFecha.Location = new System.Drawing.Point(440, 408);
			this.dtFecha.Name = "dtFecha";
			this.dtFecha.NonAutoSizeHeight = 23;
			this.dtFecha.Size = new System.Drawing.Size(112, 21);
			this.dtFecha.SpinButtonsVisible = true;
			this.dtFecha.TabIndex = 987;
			this.dtFecha.Value = ((object)(resources.GetObject("dtFecha.Value")));
			this.dtFecha.Visible = false;
			this.dtFecha.ValueChanged += new System.EventHandler(this.dtFecha_ValueChanged);
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
			this.btnGuardar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnGuardar.Image = ((System.Drawing.Image)(resources.GetObject("btnGuardar.Image")));
			this.btnGuardar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnGuardar.Location = new System.Drawing.Point(8, 400);
			this.btnGuardar.Name = "btnGuardar";
			this.btnGuardar.Size = new System.Drawing.Size(76, 24);
			this.btnGuardar.TabIndex = 988;
			this.btnGuardar.Text = "&Guardar";
			this.btnGuardar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnGuardar.Click += new System.EventHandler(this.btnGuardar_Click);
			// 
			// btnSalir
			// 
			this.btnSalir.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnSalir.CausesValidation = false;
			this.btnSalir.Image = ((System.Drawing.Image)(resources.GetObject("btnSalir.Image")));
			this.btnSalir.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnSalir.Location = new System.Drawing.Point(96, 400);
			this.btnSalir.Name = "btnSalir";
			this.btnSalir.TabIndex = 989;
			this.btnSalir.Text = "&Salir";
			this.btnSalir.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnSalir.Click += new System.EventHandler(this.btnSalir_Click);
			// 
			// uGridDetalle
			// 
			this.uGridDetalle.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.uGridDetalle.Cursor = System.Windows.Forms.Cursors.Default;
			this.uGridDetalle.DataSource = this.ultraDataSource2;
			appearance2.BackColor = System.Drawing.Color.White;
			this.uGridDetalle.DisplayLayout.Appearance = appearance2;
			ultraGridColumn4.Header.VisiblePosition = 0;
			ultraGridColumn4.Hidden = true;
			ultraGridColumn5.Header.VisiblePosition = 1;
			ultraGridColumn5.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDown;
			ultraGridColumn5.Width = 250;
			appearance3.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn6.CellAppearance = appearance3;
			ultraGridColumn6.Header.VisiblePosition = 2;
			appearance4.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn7.CellAppearance = appearance4;
			ultraGridColumn7.Header.VisiblePosition = 3;
			ultraGridColumn8.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance5.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn8.CellAppearance = appearance5;
			ultraGridColumn8.Header.VisiblePosition = 4;
			appearance6.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn9.CellAppearance = appearance6;
			ultraGridColumn9.Format = "#,##0.00";
			ultraGridColumn9.Header.VisiblePosition = 5;
			ultraGridColumn9.Width = 120;
			ultraGridBand2.Columns.AddRange(new object[] {
																										 ultraGridColumn4,
																										 ultraGridColumn5,
																										 ultraGridColumn6,
																										 ultraGridColumn7,
																										 ultraGridColumn8,
																										 ultraGridColumn9});
			ultraGridBand2.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.Edit;
			appearance7.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings1.Appearance = appearance7;
			summarySettings1.DisplayFormat = "{0: #,##0.00}";
			summarySettings1.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			ultraGridBand2.Summaries.AddRange(new Infragistics.Win.UltraWinGrid.SummarySettings[] {
																																															summarySettings1});
			ultraGridBand2.SummaryFooterCaption = "Totales";
			this.uGridDetalle.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			appearance8.ForeColor = System.Drawing.Color.Black;
			appearance8.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridDetalle.DisplayLayout.Override.ActiveRowAppearance = appearance8;
			this.uGridDetalle.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.TemplateOnBottom;
			this.uGridDetalle.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.True;
			this.uGridDetalle.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance9.BackColor = System.Drawing.Color.Transparent;
			this.uGridDetalle.DisplayLayout.Override.CardAreaAppearance = appearance9;
			this.uGridDetalle.DisplayLayout.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.Edit;
			appearance10.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance10.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance10.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance10.FontData.BoldAsString = "True";
			appearance10.FontData.Name = "Arial";
			appearance10.FontData.SizeInPoints = 8.5F;
			appearance10.ForeColor = System.Drawing.Color.White;
			appearance10.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridDetalle.DisplayLayout.Override.HeaderAppearance = appearance10;
			appearance11.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance11.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance11.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridDetalle.DisplayLayout.Override.RowAlternateAppearance = appearance11;
			appearance12.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance12.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance12.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridDetalle.DisplayLayout.Override.RowSelectorAppearance = appearance12;
			appearance13.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance13.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance13.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridDetalle.DisplayLayout.Override.SelectedRowAppearance = appearance13;
			this.uGridDetalle.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridDetalle.Location = new System.Drawing.Point(8, 8);
			this.uGridDetalle.Name = "uGridDetalle";
			this.uGridDetalle.Size = new System.Drawing.Size(584, 384);
			this.uGridDetalle.TabIndex = 859;
			this.uGridDetalle.AfterRowInsert += new Infragistics.Win.UltraWinGrid.RowEventHandler(this.uGridDetalle_AfterRowInsert);
			this.uGridDetalle.BeforeRowsDeleted += new Infragistics.Win.UltraWinGrid.BeforeRowsDeletedEventHandler(this.uGridDetalle_BeforeRowsDeleted);
			this.uGridDetalle.AfterCellUpdate += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.uGridDetalle_AfterCellUpdate);
			// 
			// frmConfiTransporte
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(608, 438);
			this.Controls.Add(this.btnSalir);
			this.Controls.Add(this.btnGuardar);
			this.Controls.Add(this.cmbBodega);
			this.Controls.Add(this.uGridDetalle);
			this.Controls.Add(this.dtFecha);
			this.KeyPreview = true;
			this.MaximizeBox = false;
			this.MinimizeBox = false;
			this.Name = "frmConfiTransporte";
			this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
			this.Text = "Configuración Voucher Transporte";
			this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.frmConfiTransporte_KeyDown);
			this.Load += new System.EventHandler(this.frmConfiTransporte_Load);
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbBodega)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFecha)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridDetalle)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion
		private Acceso miAcceso;
		private void UnloadMe()
		{
			this.Close();
		}

			bool bEliminaAlValidar = false;
		private void frmConfiTransporte_Load(object sender, System.EventArgs e)
		{

			miAcceso = new Acceso(cdsSeteoF, "201101");

			if (!Funcion.Permiso("1815", cdsSeteoF))
			{				
				MessageBox.Show("No puede ingresar a Configuración Voucher Transporte", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				BeginInvoke(new MethodInvoker(UnloadMe));
				return;
			}

			FuncionesProcedimientos.RetornaFechaServidor(this.dtFecha, cdsSeteoF, false);
			
			string sSQLPa = string.Format(" EXEC [ConsultaTransporteValores]");
			FuncionesProcedimientos.dsGeneral(sSQLPa, this.uGridDetalle);

			if (this.dtFecha.Value != System.DBNull.Value)
			{
				this.cmbBodega.DataSource = Funcion.dvProcedimiento(cdsSeteoF, string.Format("EXEC FacturacionListaBodegas {0}, {1}, '{2}', {3} ", 
					MenuLatinium.IdUsuario, 1, Convert.ToDateTime(this.dtFecha.Value).ToString("yyyyMMdd"), false));
			}
		}

		private void dtFecha_ValueChanged(object sender, System.EventArgs e)
		{
			
		}

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);		
		
		}

		private void ultraDataSource2_CellDataRequested(object sender, Infragistics.Win.UltraWinDataSource.CellDataRequestedEventArgs e)
		{
		
		}

		private void uGridDetalle_AfterRowInsert(object sender, Infragistics.Win.UltraWinGrid.RowEventArgs e)
		{
			e.Row.Cells["idTransporteValores"].Value = 0;
			e.Row.Cells["Bodega"].Value = 0;
			e.Row.Cells["Dias"].Value = 1;
			e.Row.Cells["Valor"].Value = 0.00 ;	
			e.Row.Cells["Semanas"].Value = 4;
			e.Row.Cells["Salidas"].Value = ((int) e.Row.Cells["Semanas"].Value * (int) e.Row.Cells["Dias"].Value);
			
		
		}
    bool bActualiza = false;
		private void btnGuardar_Click(object sender, System.EventArgs e)
		{
			if ((int) this.uGridDetalle.Rows.Count == 0)
			{
				MessageBox.Show("Ingrese los valores por Bodega", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.uGridDetalle.Focus ();
				return;
			}

			#region bodega duplicada
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridDetalle.Rows.All)
			{
				if (dr.Cells["Bodega"].Value != System.DBNull.Value)
				{
					if ((int) dr.Cells["Bodega"].Value > 0)
					{
						int icont = 0;
						foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dra in this.uGridDetalle.Rows.All)
						{
							if (dra.Cells["Bodega"].Value != System.DBNull.Value)
							{
								if ((int) dra.Cells["Bodega"].Value > 0)
								{
									if ((int) dra.Cells["Bodega"].Value == (int) dr.Cells["Bodega"].Value )
									{
										icont ++;
									}
									if (icont > 1)
									{
										MessageBox.Show("La bodega seleccionado ya fue ingresada.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
										this.uGridDetalle.ActiveRow = dra;
										this.uGridDetalle.ActiveRow.Selected = true;
										return;
									}
								}
							}
						}
					}
				}
			}
			#endregion bodega duplicada

			#region Valores
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow drw in this.uGridDetalle.Rows.All)
			{
				if (drw.Cells["Valor"].Value == System.DBNull.Value)
				{ 
					MessageBox.Show("Ingrese el Valor.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					this.uGridDetalle.ActiveRow = drw;
					this.uGridDetalle.ActiveRow.Selected = true;
					return;
				}
				if ((decimal) drw.Cells["Valor"].Value == 0)
				{ 
					MessageBox.Show("El valor ingresado no puede ser 0.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					this.uGridDetalle.ActiveRow = drw;
					this.uGridDetalle.ActiveRow.Selected = true;
					return;
				}
			}
			#endregion Valores

			#region Elimina Filas Vacias
			bEliminaAlValidar = true;
			vEliminaFilasVacias(this.uGridDetalle, "Bodega");
			bEliminaAlValidar = false;
			#endregion Elimina Filas Vacias
      

			#region Guardar

			//			DateTime dtInicia = DateTime.Today;
			//			DateTime dtRetorna = DateTime.Today;
			//				dtInicia = (DateTime)this.dtInicia.Value;
			//				dtRetorna = (DateTime)this.dtRetorna.Value;

			using (SqlConnection oConexion = new SqlConnection("Password=Bl45o6$9;Persist Security Info=True;User ID= " + MenuLatinium.stUsuario + ";Initial Catalog=" + MenuLatinium.stDirFacturacion + ";Data Source=" + MenuLatinium.stDirServidor + "; max pool size=1024"))
			{
				oConexion.Open();

				SqlTransaction oTransaction = null;
				SqlCommand oCmdActualiza = oConexion.CreateCommand();
				oCmdActualiza.Connection = oConexion;
				oCmdActualiza.CommandType = CommandType.Text;
				oCmdActualiza.CommandTimeout = 120;

				try
				{
					oTransaction = oConexion.BeginTransaction();
					oCmdActualiza.Transaction = oTransaction;

					#region Graba Maestro
					foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridDetalle.Rows.All)
					{
						string  sSQLDet = string.Format("Exec GuardaTransporteValores {0}, {1}, '{2}', {3}, '{4}'" ,
							Convert.ToInt32 ( dr.Cells["idTransporteValores"].Value), 
							Convert.ToInt32 ( dr.Cells["Bodega"].Value),
							Convert.ToInt32 ( dr.Cells["Dias"].Value),
							Convert.ToInt32 ( dr.Cells["Semanas"].Value),
							Convert.ToDecimal ( dr.Cells["Valor"].Value)
							);
						oCmdActualiza.CommandText = sSQLDet;
						oCmdActualiza.ExecuteNonQuery();	
					}
					#endregion Graba Maestro

					oTransaction.Commit();

					Cursor = Cursors.Default;	

					MessageBox.Show("Registro Guardado Correctamente", "Información", MessageBoxButtons.OK, MessageBoxIcon.Information);

					#region Controles
				
					string sSQLPa = string.Format(" EXEC [ConsultaTransporteValores]");
					FuncionesProcedimientos.dsGeneral(sSQLPa, this.uGridDetalle);

					#endregion Controles			
				}
				catch(Exception ex)
				{
					try
					{ 
						oTransaction.Rollback();

						MessageBox.Show(string.Format("Commit Exception Type : {0} {1}", ex.GetType(), ex.Message), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
					}
					catch (Exception ex2)
					{
						MessageBox.Show(string.Format("Rollback Exception Type: {0} {1}", ex2.GetType(), ex2.Message), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);							
					}
				}
				finally
				{
					oConexion.Close();
				}
			}
				
			#endregion Guardar


		}

	  public void vEliminaFilasVacias(Infragistics.Win.UltraWinGrid.UltraGrid uGrid, string sColumna)
		{	
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in uGrid.Rows.All)
			{
				if (dr.Cells[sColumna].Value == System.DBNull.Value) dr.Delete();
				if ((int)dr.Cells[sColumna].Value == 0) dr.Delete();
			}
		}

		private void uGridDetalle_AfterCellUpdate(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			if (bActualiza) return;

			#region bodega
			if (e.Cell.Column.ToString() == "Bodega")
			{	
				bActualiza = true;

				if ( uGridDetalle.ActiveRow.Cells["Bodega"].Value != System.DBNull.Value )
				{								
					if ((int) uGridDetalle.ActiveRow.Cells["idTransporteValores"].Value > 0 )
					{				
						int icont = Funcion.iEscalarSQL(cdsSeteoF, string.Format("SELECT count(1) from dbo.RegistroTransporte where Bodega = {0}",  (int)e.Cell.OriginalValue), true);
						if( icont > 0)
						{
							e.Cell.Row.Cells["Bodega"].Value = (int)e.Cell.OriginalValue;	
						}
					}
				}
				bActualiza = false;
			}
      #endregion bodega

			#region dias
			if (e.Cell.Column.ToString() == "Dias")
			{	
				bActualiza = true;

				if ( uGridDetalle.ActiveRow.Cells["Dias"].Value == System.DBNull.Value )
				{								
					e.Cell.Row.Cells["Dias"].Value = (int)e.Cell.OriginalValue;				
				}	
				if ((int) uGridDetalle.ActiveRow.Cells["Dias"].Value == 0 )
				{								
					e.Cell.Row.Cells["Dias"].Value = (int)e.Cell.OriginalValue;				
				}
				if ((int) uGridDetalle.ActiveRow.Cells["Dias"].Value > 7 )
				{								
					e.Cell.Row.Cells["Dias"].Value = (int)e.Cell.OriginalValue;				
				}
			
				bActualiza = false;
			}
      #endregion dias

			#region semanas
			if (e.Cell.Column.ToString() == "Semanas")
			{	
				bActualiza = true;

				if ( uGridDetalle.ActiveRow.Cells["Semanas"].Value == System.DBNull.Value )
				{								
					e.Cell.Row.Cells["Semanas"].Value = (int)e.Cell.OriginalValue;				
				}	
				if ((int) uGridDetalle.ActiveRow.Cells["Semanas"].Value == 0 )
				{								
					e.Cell.Row.Cells["Semanas"].Value = (int)e.Cell.OriginalValue;				
				}
				if ((int) uGridDetalle.ActiveRow.Cells["Semanas"].Value > 5 )
				{								
					e.Cell.Row.Cells["Semanas"].Value = (int)e.Cell.OriginalValue;				
				}
			
				bActualiza = false;
			}
			#endregion semanas

			#region sumas
			if (e.Cell.Column.ToString() == "Semanas" || e.Cell.Column.ToString() == "Dias")
			{
				if ( uGridDetalle.ActiveRow.Cells["Dias"].Value != System.DBNull.Value && uGridDetalle.ActiveRow.Cells["Semanas"].Value != System.DBNull.Value )
				{
					if ( (int) uGridDetalle.ActiveRow.Cells["Dias"].Value > 0  && (int) uGridDetalle.ActiveRow.Cells["Semanas"].Value > 0 )
					{
						uGridDetalle.ActiveRow.Cells["Salidas"].Value = (int) uGridDetalle.ActiveRow.Cells["Dias"].Value * (int) uGridDetalle.ActiveRow.Cells["Semanas"].Value;
					}
				}
			}
     #endregion sumas
      
			#region Valor
			if (e.Cell.Column.ToString() == "Valor")
			{	
				bActualiza = true;

				if ( uGridDetalle.ActiveRow.Cells["Valor"].Value == System.DBNull.Value )
				{								
					e.Cell.Row.Cells["Valor"].Value = (decimal)e.Cell.OriginalValue;				
				}	
				if ((decimal) uGridDetalle.ActiveRow.Cells["Valor"].Value < 0 )
				{								
					e.Cell.Row.Cells["Valor"].Value = (decimal)e.Cell.OriginalValue;				
				}
				bActualiza = false;
			}
     #endregion Valor
		}

		private void frmConfiTransporte_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Escape) this.Close();
		}

		private void btnSalir_Click(object sender, System.EventArgs e)
		{
		 this.Close();
		}

		private void uGridDetalle_BeforeRowsDeleted(object sender, Infragistics.Win.UltraWinGrid.BeforeRowsDeletedEventArgs e)
		{
			if (!miAcceso.Eliminar) 
			{
				e.Cancel = true;
			}
			else
			{
				#region Elimina
//				if (DialogResult.Yes ==	MessageBox.Show("¿Esta seguro de borrar el registro?", "Point Technology", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2))
//				{				
				for (int i=0; i < e.Rows.Length; i++)
				{	
					int icont = Funcion.iEscalarSQL(cdsSeteoF, string.Format("SELECT count(1) from RegistroTransporte where Bodega = {0}", (int)e.Rows[i].Cells["Bodega"].Value), true);
					if( icont == 0)
					{
						if (DialogResult.Yes ==	MessageBox.Show("¿Esta seguro de borrar el registro?", "Point Technology", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2))
						{
							if ((int)e.Rows[i].Cells["idTransporteValores"].Value > 0)
							{
								string sSQL = string.Format("Delete From TransporteValores where idTransporteValores = {0} ", (int)e.Rows[i].Cells["idTransporteValores"].Value);
								Funcion.EjecutaSQL(cdsSeteoF, sSQL, true);

								e.DisplayPromptMsg = false;
							}
							else e.Cancel = true;
						}	
						else e.Cancel = true;
					}
					else e.Cancel = true;	
				}
				#endregion Elimina
			}
		}
	}
}
