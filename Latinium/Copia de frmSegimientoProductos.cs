using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de frmSegimientoProductos.
	/// </summary>
	public class frmSegimientoProductos : System.Windows.Forms.Form
	{
		private Infragistics.Win.UltraWinGrid.UltraGrid ugDTProductos;
		private System.Windows.Forms.Label label5;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtFechaHasta;
		private System.Windows.Forms.Label label3;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtFechaDesde;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtBusca;
		private System.Windows.Forms.Label label4;
		private System.Windows.Forms.Button btnActualizar;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource1;
		private System.Windows.Forms.Button button1;
		private System.Windows.Forms.ToolTip toolTip1;
		private C1.Data.C1DataSet cdsSeteoF;
		private System.ComponentModel.IContainer components;

		public frmSegimientoProductos()
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
			this.components = new System.ComponentModel.Container();
			Infragistics.Win.Appearance appearance1 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idProcesoDT");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Fecha");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Usuario");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Ingreso");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Salida");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Proceso");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("MaquinaProceso");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Observaciones");
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idProcesoDT");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Fecha");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn3 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Usuario");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn4 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Ingreso");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn5 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Salida");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn6 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Proceso");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn7 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("MaquinaProceso");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn8 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Observaciones");
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton1 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmSegimientoProductos));
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton2 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			this.ugDTProductos = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource1 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.label5 = new System.Windows.Forms.Label();
			this.dtFechaHasta = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.label3 = new System.Windows.Forms.Label();
			this.dtFechaDesde = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.txtBusca = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label4 = new System.Windows.Forms.Label();
			this.btnActualizar = new System.Windows.Forms.Button();
			this.button1 = new System.Windows.Forms.Button();
			this.toolTip1 = new System.Windows.Forms.ToolTip(this.components);
			this.cdsSeteoF = new C1.Data.C1DataSet();
			((System.ComponentModel.ISupportInitialize)(this.ugDTProductos)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFechaHasta)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFechaDesde)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtBusca)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			this.SuspendLayout();
			// 
			// ugDTProductos
			// 
			this.ugDTProductos.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.ugDTProductos.Cursor = System.Windows.Forms.Cursors.Default;
			this.ugDTProductos.DataSource = this.ultraDataSource1;
			appearance1.BackColor = System.Drawing.Color.White;
			this.ugDTProductos.DisplayLayout.Appearance = appearance1;
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridColumn3.Header.VisiblePosition = 2;
			ultraGridColumn4.Header.Caption = "# Ingreso";
			ultraGridColumn4.Header.VisiblePosition = 3;
			ultraGridColumn5.Header.Caption = "# Salida";
			ultraGridColumn5.Header.VisiblePosition = 4;
			ultraGridColumn6.Header.VisiblePosition = 5;
			ultraGridColumn7.Header.Caption = "Máquina";
			ultraGridColumn7.Header.VisiblePosition = 6;
			ultraGridColumn8.Header.VisiblePosition = 7;
			ultraGridBand1.Columns.AddRange(new object[] {
																										 ultraGridColumn1,
																										 ultraGridColumn2,
																										 ultraGridColumn3,
																										 ultraGridColumn4,
																										 ultraGridColumn5,
																										 ultraGridColumn6,
																										 ultraGridColumn7,
																										 ultraGridColumn8});
			this.ugDTProductos.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			appearance2.ForeColor = System.Drawing.Color.Black;
			appearance2.ForeColorDisabled = System.Drawing.Color.Black;
			this.ugDTProductos.DisplayLayout.Override.ActiveRowAppearance = appearance2;
			this.ugDTProductos.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.ugDTProductos.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance3.BackColor = System.Drawing.Color.Transparent;
			this.ugDTProductos.DisplayLayout.Override.CardAreaAppearance = appearance3;
			appearance4.ForeColor = System.Drawing.Color.Black;
			appearance4.ForeColorDisabled = System.Drawing.Color.Black;
			this.ugDTProductos.DisplayLayout.Override.GroupByRowAppearance = appearance4;
			appearance5.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance5.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance5.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance5.FontData.BoldAsString = "True";
			appearance5.FontData.Name = "Arial";
			appearance5.FontData.SizeInPoints = 10F;
			appearance5.ForeColor = System.Drawing.Color.White;
			appearance5.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.ugDTProductos.DisplayLayout.Override.HeaderAppearance = appearance5;
			this.ugDTProductos.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortSingle;
			appearance6.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance6.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance6.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ugDTProductos.DisplayLayout.Override.RowAlternateAppearance = appearance6;
			appearance7.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance7.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance7.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ugDTProductos.DisplayLayout.Override.RowSelectorAppearance = appearance7;
			appearance8.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance8.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance8.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ugDTProductos.DisplayLayout.Override.SelectedRowAppearance = appearance8;
			this.ugDTProductos.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.ugDTProductos.Location = new System.Drawing.Point(8, 56);
			this.ugDTProductos.Name = "ugDTProductos";
			this.ugDTProductos.Size = new System.Drawing.Size(976, 488);
			this.ugDTProductos.TabIndex = 695;
			// 
			// ultraDataSource1
			// 
			ultraDataColumn1.DataType = typeof(int);
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
			// label5
			// 
			this.label5.Font = new System.Drawing.Font("Tahoma", 8F, System.Drawing.FontStyle.Bold);
			this.label5.Location = new System.Drawing.Point(368, 8);
			this.label5.Name = "label5";
			this.label5.Size = new System.Drawing.Size(64, 16);
			this.label5.TabIndex = 694;
			this.label5.Text = "Hasta";
			this.label5.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.label5.Visible = false;
			// 
			// dtFechaHasta
			// 
			this.dtFechaHasta.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton1.Caption = "Today";
			this.dtFechaHasta.DateButtons.Add(dateButton1);
			this.dtFechaHasta.Format = "dd/MM/yyyy";
			this.dtFechaHasta.Location = new System.Drawing.Point(368, 24);
			this.dtFechaHasta.Name = "dtFechaHasta";
			this.dtFechaHasta.NonAutoSizeHeight = 23;
			this.dtFechaHasta.Size = new System.Drawing.Size(112, 21);
			this.dtFechaHasta.SpinButtonsVisible = true;
			this.dtFechaHasta.TabIndex = 693;
			this.dtFechaHasta.Value = ((object)(resources.GetObject("dtFechaHasta.Value")));
			this.dtFechaHasta.Visible = false;
			// 
			// label3
			// 
			this.label3.Font = new System.Drawing.Font("Tahoma", 8F, System.Drawing.FontStyle.Bold);
			this.label3.Location = new System.Drawing.Point(256, 8);
			this.label3.Name = "label3";
			this.label3.Size = new System.Drawing.Size(64, 16);
			this.label3.TabIndex = 692;
			this.label3.Text = "Desde";
			this.label3.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.label3.Visible = false;
			// 
			// dtFechaDesde
			// 
			this.dtFechaDesde.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton2.Caption = "Today";
			this.dtFechaDesde.DateButtons.Add(dateButton2);
			this.dtFechaDesde.Format = "dd/MM/yyyy";
			this.dtFechaDesde.Location = new System.Drawing.Point(256, 24);
			this.dtFechaDesde.Name = "dtFechaDesde";
			this.dtFechaDesde.NonAutoSizeHeight = 23;
			this.dtFechaDesde.Size = new System.Drawing.Size(112, 21);
			this.dtFechaDesde.SpinButtonsVisible = true;
			this.dtFechaDesde.TabIndex = 691;
			this.dtFechaDesde.Value = ((object)(resources.GetObject("dtFechaDesde.Value")));
			this.dtFechaDesde.Visible = false;
			// 
			// txtBusca
			// 
			appearance9.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtBusca.Appearance = appearance9;
			this.txtBusca.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtBusca.Location = new System.Drawing.Point(8, 24);
			this.txtBusca.Name = "txtBusca";
			this.txtBusca.Size = new System.Drawing.Size(240, 21);
			this.txtBusca.TabIndex = 690;
			// 
			// label4
			// 
			this.label4.Font = new System.Drawing.Font("Tahoma", 8F, System.Drawing.FontStyle.Bold);
			this.label4.Location = new System.Drawing.Point(8, 8);
			this.label4.Name = "label4";
			this.label4.Size = new System.Drawing.Size(208, 16);
			this.label4.TabIndex = 689;
			this.label4.Text = "# Ingreso / Egreso / Nombre Clte.";
			this.label4.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// btnActualizar
			// 
			this.btnActualizar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnActualizar.CausesValidation = false;
			this.btnActualizar.Image = ((System.Drawing.Image)(resources.GetObject("btnActualizar.Image")));
			this.btnActualizar.Location = new System.Drawing.Point(536, 8);
			this.btnActualizar.Name = "btnActualizar";
			this.btnActualizar.Size = new System.Drawing.Size(40, 40);
			this.btnActualizar.TabIndex = 688;
			this.btnActualizar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.toolTip1.SetToolTip(this.btnActualizar, "Buscar información...");
			this.btnActualizar.Click += new System.EventHandler(this.btnActualizar_Click);
			// 
			// button1
			// 
			this.button1.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.button1.CausesValidation = false;
			this.button1.Image = ((System.Drawing.Image)(resources.GetObject("button1.Image")));
			this.button1.Location = new System.Drawing.Point(576, 8);
			this.button1.Name = "button1";
			this.button1.Size = new System.Drawing.Size(40, 40);
			this.button1.TabIndex = 696;
			this.button1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.toolTip1.SetToolTip(this.button1, "Exportar a EXCEL");
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
			// frmSegimientoProductos
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(992, 550);
			this.Controls.Add(this.button1);
			this.Controls.Add(this.ugDTProductos);
			this.Controls.Add(this.label5);
			this.Controls.Add(this.dtFechaHasta);
			this.Controls.Add(this.label3);
			this.Controls.Add(this.dtFechaDesde);
			this.Controls.Add(this.txtBusca);
			this.Controls.Add(this.label4);
			this.Controls.Add(this.btnActualizar);
			this.Name = "frmSegimientoProductos";
			this.Text = "Segimiento de Productos DT";
			this.Load += new System.EventHandler(this.frmSegimientoProductos_Load);
			((System.ComponentModel.ISupportInitialize)(this.ugDTProductos)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFechaHasta)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFechaDesde)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtBusca)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);
		}

		private void frmSegimientoProductos_Load(object sender, System.EventArgs e)
		{
		
		}

		private void btnActualizar_Click(object sender, System.EventArgs e)
		{
		
		}
	}
}
