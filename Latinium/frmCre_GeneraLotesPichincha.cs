using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.IO;
using System.Text;
using System.Threading;
using System.Data;
using System.Data.SqlClient;
using CrystalDecisions;
using CrystalDecisions.CrystalReports.Engine;
using CrystalDecisions.ReportSource;
using CrystalDecisions.Shared;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de frmCre_GeneraLotesPichincha.
	/// </summary>
	public class frmCre_GeneraLotesPichincha : System.Windows.Forms.Form
	{
		private Infragistics.Win.UltraWinEditors.UltraOptionSet optEstado;
		private System.Windows.Forms.GroupBox groupBox1;
		private System.Windows.Forms.Button btnGenerar;
		private System.Windows.Forms.Label label7;
		private System.Windows.Forms.Label label6;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtBuscar;
		private System.Windows.Forms.Button btnVer;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtpFecha;
		private System.Windows.Forms.Button btnExportar;
		private Infragistics.Win.UltraWinGrid.UltraGrid uGridOperaciones;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtLote;
		private System.Windows.Forms.Label label3;
		private System.Windows.Forms.Label label1;
		private System.Windows.Forms.Label lblContador;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource1;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtUltimoLote;
		private Infragistics.Win.UltraWinEditors.UltraOptionSet optCapturados;
		private C1.Data.C1DataSet cdsSeteoF;
		private System.Windows.Forms.ContextMenu mnuAcreditacion;
		private System.Windows.Forms.MenuItem menuItem1;
		private System.Windows.Forms.MenuItem menuItem2;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource2;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public frmCre_GeneraLotesPichincha()
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
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.ValueListItem valueListItem1 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem2 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem3 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.ValueListItem valueListItem4 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem5 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem6 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem7 = new Infragistics.Win.ValueListItem();
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmCre_GeneraLotesPichincha));
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton1 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCompra");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Lote");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaTransitoLote");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("NumeroLote");
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Numero");
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Fecha");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Ruc");
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cuotas_Por_Pagar");
			Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaPrimerPago");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn11 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Capital");
			Infragistics.Win.Appearance appearance11 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn12 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Fecha_Efectiva");
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings1 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Capital", 10, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Capital", 10, true);
			Infragistics.Win.Appearance appearance12 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance13 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance14 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance15 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance16 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance17 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCompra");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Lote");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn3 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("FechaTransitoLote");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn4 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("NumeroLote");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn5 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Numero");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn6 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Fecha");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn7 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Ruc");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn8 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn9 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Cuotas_Por_Pagar");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn10 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("FechaPrimerPago");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn11 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Capital");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn12 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Fecha_Efectiva");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn13 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCompra");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn14 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("CapturaLote");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn15 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("NumeroLote");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn16 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Numero");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn17 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Fecha");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn18 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Ruc");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn19 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn20 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Cuotas_Por_Pagar");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn21 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("FechaPrimerPago");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn22 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Capital");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn23 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Fecha_Efectiva");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn24 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Vence");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn25 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Total_Credito");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn26 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Saldo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn27 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Capital_Credito");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn28 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Interes_Credito");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn29 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Estado_Solicitud");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn30 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("EstadoCredito");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn31 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("EstadoEPC");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn32 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Notas");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn33 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("FechaTransitoLote");
			Infragistics.Win.Appearance appearance18 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance19 = new Infragistics.Win.Appearance();
			this.optCapturados = new Infragistics.Win.UltraWinEditors.UltraOptionSet();
			this.optEstado = new Infragistics.Win.UltraWinEditors.UltraOptionSet();
			this.groupBox1 = new System.Windows.Forms.GroupBox();
			this.btnGenerar = new System.Windows.Forms.Button();
			this.label7 = new System.Windows.Forms.Label();
			this.label6 = new System.Windows.Forms.Label();
			this.txtBuscar = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.btnVer = new System.Windows.Forms.Button();
			this.dtpFecha = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.btnExportar = new System.Windows.Forms.Button();
			this.uGridOperaciones = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource2 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.ultraDataSource1 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.txtLote = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label3 = new System.Windows.Forms.Label();
			this.txtUltimoLote = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label1 = new System.Windows.Forms.Label();
			this.lblContador = new System.Windows.Forms.Label();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.mnuAcreditacion = new System.Windows.Forms.ContextMenu();
			this.menuItem1 = new System.Windows.Forms.MenuItem();
			this.menuItem2 = new System.Windows.Forms.MenuItem();
			((System.ComponentModel.ISupportInitialize)(this.optCapturados)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.optEstado)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtBuscar)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtpFecha)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridOperaciones)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtLote)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtUltimoLote)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			this.SuspendLayout();
			// 
			// optCapturados
			// 
			appearance1.ForeColorDisabled = System.Drawing.Color.Black;
			this.optCapturados.Appearance = appearance1;
			this.optCapturados.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.optCapturados.CheckedIndex = 0;
			this.optCapturados.Enabled = false;
			this.optCapturados.ItemAppearance = appearance2;
			this.optCapturados.ItemOrigin = new System.Drawing.Point(10, 0);
			valueListItem1.DataValue = 0;
			valueListItem1.DisplayText = "Todos";
			valueListItem2.DataValue = 1;
			valueListItem2.DisplayText = "Automaticos";
			valueListItem3.DataValue = 2;
			valueListItem3.DisplayText = "Manuales";
			this.optCapturados.Items.Add(valueListItem1);
			this.optCapturados.Items.Add(valueListItem2);
			this.optCapturados.Items.Add(valueListItem3);
			this.optCapturados.ItemSpacingVertical = 10;
			this.optCapturados.Location = new System.Drawing.Point(424, 37);
			this.optCapturados.Name = "optCapturados";
			this.optCapturados.Size = new System.Drawing.Size(240, 26);
			this.optCapturados.TabIndex = 761;
			this.optCapturados.Text = "Todos";
			this.optCapturados.Visible = false;
			// 
			// optEstado
			// 
			appearance3.ForeColorDisabled = System.Drawing.Color.Black;
			this.optEstado.Appearance = appearance3;
			this.optEstado.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.optEstado.CheckedIndex = 1;
			this.optEstado.ItemAppearance = appearance4;
			this.optEstado.ItemOrigin = new System.Drawing.Point(10, 0);
			valueListItem4.DataValue = 0;
			valueListItem4.DisplayText = "Todos";
			valueListItem5.DataValue = 1;
			valueListItem5.DisplayText = "Pendiente";
			valueListItem6.DataValue = 2;
			valueListItem6.DisplayText = "Enviado";
			valueListItem7.DataValue = 3;
			valueListItem7.DisplayText = "Liquidado";
			this.optEstado.Items.Add(valueListItem4);
			this.optEstado.Items.Add(valueListItem5);
			this.optEstado.Items.Add(valueListItem6);
			this.optEstado.Items.Add(valueListItem7);
			this.optEstado.ItemSpacingVertical = 10;
			this.optEstado.Location = new System.Drawing.Point(424, 5);
			this.optEstado.Name = "optEstado";
			this.optEstado.Size = new System.Drawing.Size(304, 26);
			this.optEstado.TabIndex = 760;
			this.optEstado.Text = "Pendiente";
			// 
			// groupBox1
			// 
			this.groupBox1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.groupBox1.Location = new System.Drawing.Point(0, 104);
			this.groupBox1.Name = "groupBox1";
			this.groupBox1.Size = new System.Drawing.Size(1071, 8);
			this.groupBox1.TabIndex = 759;
			this.groupBox1.TabStop = false;
			// 
			// btnGenerar
			// 
			this.btnGenerar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btnGenerar.CausesValidation = false;
			this.btnGenerar.Enabled = false;
			this.btnGenerar.Image = ((System.Drawing.Image)(resources.GetObject("btnGenerar.Image")));
			this.btnGenerar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnGenerar.Location = new System.Drawing.Point(879, 8);
			this.btnGenerar.Name = "btnGenerar";
			this.btnGenerar.Size = new System.Drawing.Size(88, 23);
			this.btnGenerar.TabIndex = 758;
			this.btnGenerar.Text = "Generar TXT";
			this.btnGenerar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnGenerar.Click += new System.EventHandler(this.btnGenerar_Click);
			// 
			// label7
			// 
			this.label7.AutoSize = true;
			this.label7.Location = new System.Drawing.Point(8, 72);
			this.label7.Name = "label7";
			this.label7.Size = new System.Drawing.Size(39, 16);
			this.label7.TabIndex = 757;
			this.label7.Text = "Buscar";
			this.label7.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label6
			// 
			this.label6.AutoSize = true;
			this.label6.Location = new System.Drawing.Point(8, 10);
			this.label6.Name = "label6";
			this.label6.Size = new System.Drawing.Size(84, 16);
			this.label6.TabIndex = 753;
			this.label6.Text = "Fecha de Venta";
			this.label6.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtBuscar
			// 
			this.txtBuscar.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.txtBuscar.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtBuscar.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtBuscar.Location = new System.Drawing.Point(104, 71);
			this.txtBuscar.Name = "txtBuscar";
			this.txtBuscar.Size = new System.Drawing.Size(216, 19);
			this.txtBuscar.TabIndex = 756;
			this.txtBuscar.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtBuscar_KeyDown);
			this.txtBuscar.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtBuscar_KeyPress);
			this.txtBuscar.Validated += new System.EventHandler(this.txtBuscar_Validated);
			// 
			// btnVer
			// 
			this.btnVer.CausesValidation = false;
			this.btnVer.Image = ((System.Drawing.Image)(resources.GetObject("btnVer.Image")));
			this.btnVer.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnVer.Location = new System.Drawing.Point(336, 69);
			this.btnVer.Name = "btnVer";
			this.btnVer.TabIndex = 755;
			this.btnVer.Text = "Ver";
			this.btnVer.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnVer.Click += new System.EventHandler(this.btnVer_Click);
			// 
			// dtpFecha
			// 
			appearance5.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtpFecha.Appearance = appearance5;
			this.dtpFecha.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton1.Caption = "Today";
			this.dtpFecha.DateButtons.Add(dateButton1);
			this.dtpFecha.Format = "dd/MM/yyyy";
			this.dtpFecha.Location = new System.Drawing.Point(104, 8);
			this.dtpFecha.MonthOrientation = Infragistics.Win.UltraWinSchedule.MonthOrientation.DownThenAcross;
			this.dtpFecha.Name = "dtpFecha";
			this.dtpFecha.NonAutoSizeHeight = 23;
			this.dtpFecha.Size = new System.Drawing.Size(112, 21);
			this.dtpFecha.SpinButtonIncrement = Infragistics.Win.UltraWinSchedule.SpinIncrementUnit.Months;
			this.dtpFecha.SpinButtonsVisible = true;
			this.dtpFecha.TabIndex = 754;
			this.dtpFecha.Value = ((object)(resources.GetObject("dtpFecha.Value")));
			this.dtpFecha.KeyDown += new System.Windows.Forms.KeyEventHandler(this.dtpFecha_KeyDown);
			// 
			// btnExportar
			// 
			this.btnExportar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btnExportar.CausesValidation = false;
			this.btnExportar.Enabled = false;
			this.btnExportar.Image = ((System.Drawing.Image)(resources.GetObject("btnExportar.Image")));
			this.btnExportar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnExportar.Location = new System.Drawing.Point(975, 8);
			this.btnExportar.Name = "btnExportar";
			this.btnExportar.Size = new System.Drawing.Size(88, 23);
			this.btnExportar.TabIndex = 752;
			this.btnExportar.Text = "Exportar";
			this.btnExportar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnExportar.Click += new System.EventHandler(this.btnExportar_Click);
			// 
			// uGridOperaciones
			// 
			this.uGridOperaciones.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.uGridOperaciones.Cursor = System.Windows.Forms.Cursors.Default;
			this.uGridOperaciones.DataSource = this.ultraDataSource2;
			appearance6.BackColor = System.Drawing.Color.White;
			this.uGridOperaciones.DisplayLayout.Appearance = appearance6;
			ultraGridColumn1.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn2.Header.Caption = "...";
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridColumn3.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn3.Format = "dd/MM/yyyy HH:mm";
			ultraGridColumn3.Header.Caption = "Fecha Asignado";
			ultraGridColumn3.Header.VisiblePosition = 11;
			ultraGridColumn3.Width = 100;
			ultraGridColumn4.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance7.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn4.CellAppearance = appearance7;
			ultraGridColumn4.Header.Caption = "Lote";
			ultraGridColumn4.Header.VisiblePosition = 2;
			ultraGridColumn4.Width = 80;
			ultraGridColumn5.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance8.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn5.CellAppearance = appearance8;
			ultraGridColumn5.Header.Caption = "N. Documento";
			ultraGridColumn5.Header.VisiblePosition = 3;
			ultraGridColumn5.Width = 90;
			ultraGridColumn6.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn6.Header.VisiblePosition = 4;
			ultraGridColumn6.Width = 70;
			ultraGridColumn7.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance9.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn7.CellAppearance = appearance9;
			ultraGridColumn7.Header.VisiblePosition = 6;
			ultraGridColumn7.Width = 120;
			ultraGridColumn8.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn8.Header.Caption = "Cliente";
			ultraGridColumn8.Header.VisiblePosition = 7;
			ultraGridColumn8.Width = 250;
			ultraGridColumn9.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance10.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn9.CellAppearance = appearance10;
			ultraGridColumn9.Header.VisiblePosition = 8;
			ultraGridColumn10.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn10.Header.Caption = "Primer Pago";
			ultraGridColumn10.Header.VisiblePosition = 5;
			ultraGridColumn10.Hidden = true;
			ultraGridColumn11.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance11.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn11.CellAppearance = appearance11;
			ultraGridColumn11.Format = "#,##0.00";
			ultraGridColumn11.Header.VisiblePosition = 10;
			ultraGridColumn11.Width = 70;
			ultraGridColumn12.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn12.Header.Caption = "Fec Efectiva";
			ultraGridColumn12.Header.VisiblePosition = 9;
			ultraGridColumn12.Width = 80;
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
																										 ultraGridColumn11,
																										 ultraGridColumn12});
			ultraGridBand1.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
			ultraGridBand1.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			ultraGridBand1.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.False;
			appearance12.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings1.Appearance = appearance12;
			summarySettings1.DisplayFormat = "{0: #,##0.00}";
			summarySettings1.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			ultraGridBand1.Summaries.AddRange(new Infragistics.Win.UltraWinGrid.SummarySettings[] {
																																															summarySettings1});
			ultraGridBand1.SummaryFooterCaption = "Totales";
			this.uGridOperaciones.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.uGridOperaciones.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
			this.uGridOperaciones.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.uGridOperaciones.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance13.BackColor = System.Drawing.Color.Transparent;
			this.uGridOperaciones.DisplayLayout.Override.CardAreaAppearance = appearance13;
			appearance14.ForeColor = System.Drawing.Color.Black;
			appearance14.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridOperaciones.DisplayLayout.Override.CellAppearance = appearance14;
			appearance15.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance15.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance15.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance15.FontData.BoldAsString = "True";
			appearance15.FontData.Name = "Arial";
			appearance15.FontData.SizeInPoints = 8.5F;
			appearance15.ForeColor = System.Drawing.Color.White;
			appearance15.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridOperaciones.DisplayLayout.Override.HeaderAppearance = appearance15;
			this.uGridOperaciones.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti;
			appearance16.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance16.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance16.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridOperaciones.DisplayLayout.Override.RowSelectorAppearance = appearance16;
			appearance17.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance17.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance17.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridOperaciones.DisplayLayout.Override.SelectedRowAppearance = appearance17;
			this.uGridOperaciones.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridOperaciones.Location = new System.Drawing.Point(8, 120);
			this.uGridOperaciones.Name = "uGridOperaciones";
			this.uGridOperaciones.Size = new System.Drawing.Size(1064, 224);
			this.uGridOperaciones.TabIndex = 751;
			this.uGridOperaciones.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.uGridOperaciones_InitializeLayout);
			// 
			// ultraDataSource2
			// 
			ultraDataColumn1.DataType = typeof(int);
			ultraDataColumn2.DataType = typeof(bool);
			ultraDataColumn2.DefaultValue = false;
			ultraDataColumn3.DataType = typeof(System.DateTime);
			ultraDataColumn4.DataType = typeof(int);
			ultraDataColumn6.DataType = typeof(System.DateTime);
			ultraDataColumn9.DataType = typeof(int);
			ultraDataColumn10.DataType = typeof(System.DateTime);
			ultraDataColumn11.DataType = typeof(System.Decimal);
			ultraDataColumn12.DataType = typeof(System.DateTime);
			this.ultraDataSource2.Band.Columns.AddRange(new object[] {
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
																																 ultraDataColumn11,
																																 ultraDataColumn12});
			// 
			// ultraDataSource1
			// 
			ultraDataColumn13.DataType = typeof(int);
			ultraDataColumn14.DataType = typeof(bool);
			ultraDataColumn15.DataType = typeof(int);
			ultraDataColumn17.DataType = typeof(System.DateTime);
			ultraDataColumn20.DataType = typeof(int);
			ultraDataColumn21.DataType = typeof(System.DateTime);
			ultraDataColumn22.DataType = typeof(System.Decimal);
			ultraDataColumn23.DataType = typeof(System.DateTime);
			ultraDataColumn24.DataType = typeof(System.DateTime);
			ultraDataColumn25.DataType = typeof(System.Decimal);
			ultraDataColumn26.DataType = typeof(System.Decimal);
			ultraDataColumn27.DataType = typeof(System.Decimal);
			ultraDataColumn28.DataType = typeof(System.Decimal);
			ultraDataColumn30.DataType = typeof(int);
			ultraDataColumn31.DataType = typeof(int);
			ultraDataColumn33.DataType = typeof(System.DateTime);
			this.ultraDataSource1.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn13,
																																 ultraDataColumn14,
																																 ultraDataColumn15,
																																 ultraDataColumn16,
																																 ultraDataColumn17,
																																 ultraDataColumn18,
																																 ultraDataColumn19,
																																 ultraDataColumn20,
																																 ultraDataColumn21,
																																 ultraDataColumn22,
																																 ultraDataColumn23,
																																 ultraDataColumn24,
																																 ultraDataColumn25,
																																 ultraDataColumn26,
																																 ultraDataColumn27,
																																 ultraDataColumn28,
																																 ultraDataColumn29,
																																 ultraDataColumn30,
																																 ultraDataColumn31,
																																 ultraDataColumn32,
																																 ultraDataColumn33});
			// 
			// txtLote
			// 
			appearance18.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtLote.Appearance = appearance18;
			this.txtLote.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.txtLote.FormatString = "";
			this.txtLote.Location = new System.Drawing.Point(104, 40);
			this.txtLote.Name = "txtLote";
			this.txtLote.PromptChar = ' ';
			this.txtLote.Size = new System.Drawing.Size(112, 21);
			this.txtLote.SpinButtonDisplayStyle = Infragistics.Win.ButtonDisplayStyle.Always;
			this.txtLote.TabIndex = 762;
			this.txtLote.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtLote_KeyDown);
			this.txtLote.Click += new System.EventHandler(this.txtLote_Click);
			// 
			// label3
			// 
			this.label3.AutoSize = true;
			this.label3.BackColor = System.Drawing.Color.Transparent;
			this.label3.Location = new System.Drawing.Point(8, 42);
			this.label3.Name = "label3";
			this.label3.Size = new System.Drawing.Size(85, 16);
			this.label3.TabIndex = 763;
			this.label3.Text = "Numero de Lote";
			this.label3.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtUltimoLote
			// 
			appearance19.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtUltimoLote.Appearance = appearance19;
			this.txtUltimoLote.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.txtUltimoLote.Enabled = false;
			this.txtUltimoLote.FormatString = "";
			this.txtUltimoLote.Location = new System.Drawing.Point(296, 40);
			this.txtUltimoLote.Name = "txtUltimoLote";
			this.txtUltimoLote.PromptChar = ' ';
			this.txtUltimoLote.Size = new System.Drawing.Size(88, 21);
			this.txtUltimoLote.SpinButtonDisplayStyle = Infragistics.Win.ButtonDisplayStyle.Always;
			this.txtUltimoLote.TabIndex = 764;
			// 
			// label1
			// 
			this.label1.AutoSize = true;
			this.label1.BackColor = System.Drawing.Color.Transparent;
			this.label1.Location = new System.Drawing.Point(224, 42);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(61, 16);
			this.label1.TabIndex = 765;
			this.label1.Text = "Ultimo Lote";
			this.label1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// lblContador
			// 
			this.lblContador.AutoSize = true;
			this.lblContador.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblContador.ForeColor = System.Drawing.Color.Firebrick;
			this.lblContador.Location = new System.Drawing.Point(432, 72);
			this.lblContador.Name = "lblContador";
			this.lblContador.Size = new System.Drawing.Size(0, 17);
			this.lblContador.TabIndex = 766;
			this.lblContador.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
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
			// mnuAcreditacion
			// 
			this.mnuAcreditacion.MenuItems.AddRange(new System.Windows.Forms.MenuItem[] {
																																										this.menuItem1,
																																										this.menuItem2});
			// 
			// menuItem1
			// 
			this.menuItem1.Index = 0;
			this.menuItem1.Text = "Trama De Clientes";
			this.menuItem1.Click += new System.EventHandler(this.menuItem1_Click);
			// 
			// menuItem2
			// 
			this.menuItem2.Index = 1;
			this.menuItem2.Text = "Trama De Operaciones";
			this.menuItem2.Click += new System.EventHandler(this.menuItem2_Click);
			// 
			// frmCre_GeneraLotesPichincha
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(1072, 346);
			this.Controls.Add(this.lblContador);
			this.Controls.Add(this.txtUltimoLote);
			this.Controls.Add(this.label1);
			this.Controls.Add(this.label3);
			this.Controls.Add(this.label7);
			this.Controls.Add(this.label6);
			this.Controls.Add(this.txtLote);
			this.Controls.Add(this.optCapturados);
			this.Controls.Add(this.optEstado);
			this.Controls.Add(this.groupBox1);
			this.Controls.Add(this.btnGenerar);
			this.Controls.Add(this.txtBuscar);
			this.Controls.Add(this.btnVer);
			this.Controls.Add(this.dtpFecha);
			this.Controls.Add(this.btnExportar);
			this.Controls.Add(this.uGridOperaciones);
			this.KeyPreview = true;
			this.Name = "frmCre_GeneraLotesPichincha";
			this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
			this.Text = "Generación Lote Banco Pichincha";
			this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.frmCre_GeneraLotesPichincha_KeyDown);
			this.Load += new System.EventHandler(this.frmCre_GeneraLotesPichincha_Load);
			((System.ComponentModel.ISupportInitialize)(this.optCapturados)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.optEstado)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtBuscar)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtpFecha)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridOperaciones)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtLote)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtUltimoLote)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion
		private Acceso miAcceso;
		private void UnloadMe()
		{
			this.Close();
		}

		private void frmCre_GeneraLotesPichincha_Load(object sender, System.EventArgs e)
		{
			miAcceso = new Acceso(cdsSeteoF, "08320204");
			if (!Funcion.Permiso("1184", cdsSeteoF))
			{				
				MessageBox.Show("No puede ingresar a Seguro DFI", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				BeginInvoke(new MethodInvoker(UnloadMe));
				return;
			}
			if (miAcceso.BExportar)this.btnExportar.Enabled = true;
			if (miAcceso.BGenerar)this.btnGenerar.Enabled = true;

			this.dtpFecha.CalendarInfo.MinDate = DateTime.Parse("01/01/2015");
	
			this.dtpFecha.Value = DateTime.Today;

			string ssqlDo= string.Format ("Exec Cre_RetornaUltimoLote {0}", 4); 
			SqlDataReader drDO = Funcion.rEscalarSQL(cdsSeteoF,ssqlDo,true);
			drDO.Read();
			if(drDO.HasRows)
			{
				this.txtUltimoLote.Value = drDO.GetInt32 (0);
			}
			drDO.Close();

		}

		private void btnExportar_Click(object sender, System.EventArgs e)
		{
			FuncionesProcedimientos.ExportaExcel(this.uGridOperaciones);
		}

		private void btnGenerar_Click(object sender, System.EventArgs e)
		{
			this.mnuAcreditacion.Show(this.btnGenerar, new Point(1, 25));
		}

		private void btnVer_Click(object sender, System.EventArgs e)
		{
			int iCuotas = 0;
			decimal dCapital = 0.00m;

			if (this.dtpFecha.Value == System.DBNull.Value)
			{
				MessageBox.Show("Ingrese la Fecha", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.dtpFecha.Focus();
				return;
			}

			if (!Validacion.vbFechaEnDocumentos(this.dtpFecha, "Ingrese la Fecha ", true, "Bancos", cdsSeteoF)) return;
			if (!Validacion.vbCampoEnteroVacio(this.txtLote,"Ingrese el Numero de Lote",1,1000000))
			{
				this.txtLote.Focus ();
				return;
			}

			if (this.txtBuscar.Text.ToString().Length > 0)
			{    
				int iCont = Funcion.iEscalarSQL(cdsSeteoF, string.Format("SELECT COUNT (*) FROM Compra c Inner join Cliente cl on cl.idCliente = c.idCliente WHERE idTipoFactura = 1 AND c.Borrar = 0 AND c.Estado = 10 And c.idFormaPago = 9 and Confirmado = 0 And idEntidadFinanciera = 4 And cl.Ruc = '{0}' And ISNULL(SaldoAnticipo, 0) > 0.05",this.txtBuscar.Text.ToString()));	
				if (iCont > 0)
				{
					int idCompra = Funcion.iEscalarSQL(cdsSeteoF, string.Format("SELECT idCompra  FROM Compra c Inner join Cliente cl on cl.idCliente = c.idCliente WHERE idTipoFactura = 1 AND c.Borrar = 0 AND c.Estado = 10 and c.idFormaPago = 9 and Confirmado = 0  And idEntidadFinanciera = 4 And  cl.Ruc = '{0}' And ISNULL(SaldoAnticipo, 0) > 0.05",this.txtBuscar.Text.ToString()));	
				
					string sMensaje = Funcion.sEscalarSQL(cdsSeteoF, string.Format("Exec Cre_ValidaAsignacionLotePichincha {0},'{1}'",
						idCompra, Convert.ToDateTime(this.dtpFecha.Value).ToString("yyyyMMdd")));	

					if (sMensaje.Length > 0)
					{
						MessageBox.Show(string.Format("Cliente Presenta la siguiente Novedad {0} ", sMensaje), "Validacion", MessageBoxButtons.OK, MessageBoxIcon.Stop);
						return;        
					}
					
					#region consulta Cuota Capital
					string sSQLCuota= string.Format("Exec Cre_ConsultaLotePichincha {0}", idCompra);
					SqlDataReader drC = Funcion.rEscalarSQL(cdsSeteoF,sSQLCuota,true);
					drC.Read();
					if (drC.GetValue(0) != System.DBNull.Value)iCuotas = drC.GetInt32 (0);
					if (drC.GetValue(1) != System.DBNull.Value)dCapital = drC.GetDecimal(1);
					drC.Close();
					#endregion consulta Cuota Capital

					string sSQLActualiza = string.Format("Exec Cre_LoteBancoPichincha {0}, {1}, {2}, {3}, {4}, {5}, '{6}'",
						1, idCompra, dCapital, iCuotas, (int) this.txtLote.Value, 4 ,"");
					Funcion.EjecutaSQL(cdsSeteoF, sSQLActualiza);					
				}				
			}

			string sSQL = string.Format("Exec Cre_ConsultaGeneralLotePichincha '{0}', {1}, '{2}'", 
				Convert.ToDateTime(this.dtpFecha.Value).ToString("yyyyMMdd"), (int)this.txtLote.Value, "");
			FuncionesProcedimientos.dsGeneral(sSQL, this.uGridOperaciones);
			
			this.lblContador.Text = this.uGridOperaciones.Rows.Count + " Registros Encontrados";

			this.txtBuscar.Text = "";

			this.txtBuscar.Focus();			
		}

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);	
		}

		private void txtBuscar_Validated(object sender, System.EventArgs e)
		{
          
		}

		private void txtBuscar_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter) this.btnVer_Click(sender , e);
		}

		private void frmCre_GeneraLotesPichincha_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Escape) this.Close();
		}

		private void dtpFecha_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter) this.txtLote.Focus();
		}

		private void txtLote_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter) 
			{
				string sSQL = string.Format("Exec Cre_ConsultaGeneralLotePichincha '{0}', {1}, '{2}'", 
					Convert.ToDateTime(this.dtpFecha.Value).ToString("yyyyMMdd"), (int)this.txtLote.Value, "");
				FuncionesProcedimientos.dsGeneral(sSQL, this.uGridOperaciones);
				this.txtBuscar.Text = "";

				this.txtBuscar.Focus ();

			}
			 
		}

		private void txtLote_Click(object sender, System.EventArgs e)
		{
			this.txtLote.SelectAll ();
		}

		private void menuItem2_Click(object sender, System.EventArgs e)
		{
			#region valida Fecha

				if (this.dtpFecha.Value == System.DBNull.Value)
				{
					MessageBox.Show("Ingrese la Fecha", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					this.dtpFecha.Focus();
					return;
				}
				if (!Validacion.vbFechaEnDocumentos(this.dtpFecha, "Ingrese la Fecha ", true, "Cobros", cdsSeteoF)) return;

      #endregion valida Fecha

			if (!Validacion.vbCampoEnteroVacio(this.txtLote,"Ingrese el Numero de Lote",1,1000000))
			{
				this.txtLote.Focus ();
				return;
			}

			int iContLote = Funcion.iEscalarSQL(cdsSeteoF, string.Format(" EXEC Cre_ValidaLotes 4,{0}",(int)this.txtLote.Value));	
			if( iContLote < 1)
			{
				MessageBox.Show(string.Format("No Existen Creditos Seleccionados para generar el Lote N. {0}", Convert.ToInt32(this.txtLote.Text)), "Validacion", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.txtLote.Focus();
			}
			else
			{
				try
				{

					#region TXT
					SaveFileDialog sArchivo = new SaveFileDialog();

					sArchivo.Filter = "Documentos de texto (*.txt)|*.txt";
					sArchivo.FileName = "GAFPNT.txt";
					sArchivo.OverwritePrompt = true;
					sArchivo.Title = "Guardar Trama de Operaciones Banco Pichincha";
				                                
					if (sArchivo.ShowDialog() == DialogResult.OK)
					{					
						#region Pagos
						using (StreamWriter file = new StreamWriter(sArchivo.FileName, true))
						{
							#region Normales
							int i = 0;
							int j = 0;

							foreach (DataRow dr in FuncionesProcedimientos.dtGeneral(string.Format("Exec Cre_TramasOperaciones {0}, '{1}'",(int)this.txtLote.Value,
								Convert.ToDateTime(this.dtpFecha.Value).ToString("yyyyMMdd"))).Rows)
							{
								if ((Convert.ToDecimal(dr["TotalCuotas"]) - (Convert.ToDecimal(dr["TotalCapital"]) + Convert.ToDecimal(dr["TotalInteres"]))) != 0.00m)
								{
									MessageBox.Show(string.Format("En el Cliente Con Cedula '{0}' Existe Diferencia de {1} Origen '{2}'",
										dr["Datos_Cliente"].ToString(), Convert.ToDecimal(dr["TotalCuotas"]) - (Convert.ToDecimal(dr["TotalCapital"]) + Convert.ToDecimal(dr["TotalInteres"])), 2));
								}
								string sGarante = "";// nCre_LotesBancoPichincha.Cre_TramaDeOperacionesGarantes(Convert.ToString(dr["Ruc"]));

								if (sGarante.Length > 0)
								{
									sGarante = sGarante + ";";
									j++;
								}
								string sTabla = Funcion.sEscalarSQL(cdsSeteoF, string.Format(" EXEC Cre_CuotasLotesBancoPichincha {0},{1}",Convert.ToInt32(dr["idCompra"]),Convert.ToInt32(dr["Cuotas"]) ));	
								//							string sTabla = FuncionesProcedimientos.dtGeneral(string.Format("Exec Cre_CuotasLotesBancoPichincha {0}, {1}",Convert.ToInt32(dr["idCompra"]), Convert.ToInt32(dr["Cuotas"]))).Rows;

								string iLinea = dr["Fecha"].ToString() + "\t" + dr["NumeroLote"].ToString() + "\t" + dr["Peticion"].ToString() + "\t" + dr["Producto"].ToString() + "\t" + dr["Referencia"].ToString() + "\t" + dr["AgenciaIngreso"].ToString() + "\t" + dr["Usuario"].ToString() + "\t" + dr["FechaIngreso"].ToString() + "\t" + dr["HoraIngreso"].ToString() + "\t" + dr["Sistemaentrada"].ToString() + "\t"
									+ dr["Datos_Cliente"].ToString() + ";" + dr["Valor_Prestamo"].ToString() + ";" + dr["Fecha_Efectiva"].ToString() + ";" + dr["Fecha_Vencimiento"].ToString() + ";" + dr["Plazo"].ToString() + ";"
									+ dr["PeriodoPlazo"].ToString() + ";" + dr["PeriodoPago"].ToString() + ";" + dr["TasaInteres"].ToString() + ";" + dr["FormaVencDividendo"].ToString() + ";" + dr["CriterioVencDividendo"].ToString() + ";"
									+ dr["TipoCredito"].ToString() + ";" + dr["DatAgencia"].ToString() + ";" + dr["Datos_Imp_Com"].ToString() + ";" + sGarante + sTabla + dr["FechaNego"].ToString() + ";"
									+ dr["ValDesembolso"].ToString() + "|";
							
								file.WriteLine(iLinea);

								i++;
							}
							MessageBox.Show(string.Format("{0} Tramas de Operaciones con {1} Generadas Correctamente", i, j), "Informacion", MessageBoxButtons.OK, MessageBoxIcon.Information);
                       
							#endregion Normales
						}
						#endregion Pagos
					}
					#endregion TXT
				}
				catch (Exception exc)
				{
					MessageBox.Show(exc.Message);
				}
			}
		}

		private void menuItem1_Click(object sender, System.EventArgs e)
		{
			#region valida Fecha

			if (this.dtpFecha.Value == System.DBNull.Value)
			{
				MessageBox.Show("Ingrese la Fecha", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.dtpFecha.Focus();
				return;
			}
			if (!Validacion.vbFechaEnDocumentos(this.dtpFecha, "Ingrese la Fecha ", true, "Cobros", cdsSeteoF)) return;

			#endregion valida Fecha
      
			#region Valida Lotte
			if (!Validacion.vbCampoEnteroVacio(this.txtLote,"Ingrese el Numero de Lote",1,1000000))
			{
				this.txtLote.Focus ();
				return;
			}

			int iContLote = Funcion.iEscalarSQL(cdsSeteoF, string.Format(" EXEC Cre_ValidaLotes 4,{0}",(int)this.txtLote.Value));	
			#endregion Valida Lotte

			if( iContLote < 1)
			{
				MessageBox.Show(string.Format("No Existen Creditos Seleccionados para generar el Lote N. {0}", Convert.ToInt32(this.txtLote.Text)), "Validacion", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.txtLote.Focus();
			}
			else
			{
				#region Validacion
				foreach (DataRow dr in FuncionesProcedimientos.dtGeneral(string.Format("Exec Cre_TramasDeClientes {0}",(int)this.txtLote.Value)).Rows)
				{
					DataTable table = new DataTable();
					DataColumn column;
					DataRow row;
					DataView view;

					column = new DataColumn();
					column.DataType = System.Type.GetType("System.String");
					column.ColumnName = "Cedula";
					table.Columns.Add(column);

					column = new DataColumn();
					column.DataType = Type.GetType("System.String");
					column.ColumnName = "Nombre";
					table.Columns.Add(column);

					column = new DataColumn();
					column.DataType = Type.GetType("System.String");
					column.ColumnName = "Campo";
					table.Columns.Add(column);

					column = new DataColumn();
					column.DataType = Type.GetType("System.String");
					column.ColumnName = "Usuario";
					table.Columns.Add(column);

					column = new DataColumn();
					column.DataType = Type.GetType("System.String");
					column.ColumnName = "Mensaje";
					table.Columns.Add(column);

					row = table.NewRow();
					row["Cedula"] = dr["COD_NUM_DOC_OFICIAL"].ToString();
					row["Nombre"] = dr["NOM_APELLIDO1"].ToString().Trim() + dr["NOM_APELLIDO2"].ToString().Trim() + dr["NOM_PERSONA1"].ToString().Trim() + dr["NOM_PERSONA2"].ToString().Trim();
					row["Campo"] = "CEDULA";
					row["Mensaje"] = "CEDULA DEL CLIENTE DEBE DE SER DE 10 CARACTERES";
					table.Rows.Add(row);
					view = new DataView(table);
				}
				#endregion Validacion
			
			}
			try
			{
				SaveFileDialog sArchivo = new SaveFileDialog();

				sArchivo.Filter = "Documentos de texto (*.txt)|*.txt";
				sArchivo.FileName = "CLIPNT.txt";
				sArchivo.OverwritePrompt = true;
				sArchivo.Title = "Guardar Trama de Clientes Banco Pichincha";


				if (sArchivo.ShowDialog() == DialogResult.OK)
				{
					#region Clientes
					using (StreamWriter file = new StreamWriter(sArchivo.FileName, true))
					{

						int i = 0;
						foreach (DataRow dr in FuncionesProcedimientos.dtGeneral(string.Format("Exec Cre_TramasDeClientes {0}",(int)this.txtLote.Value)).Rows)
						{
							string sLinea = dr["IND_ACCION_IDENT"].ToString() + dr["COD_EMPRESA"].ToString() + dr["COD_AGENCIA"].ToString()
								+ dr["COD_DOC_OFICIAL"].ToString() + dr["COD_NUM_DOC_OFICIAL"].ToString() + dr["NOM_APELLIDO1"].ToString() + dr["NOM_APELLIDO2"].ToString() + dr["NOM_PERSONA1"].ToString() + dr["NOM_PERSONA2"].ToString()
								+ dr["COD_NACIONALIDAD"].ToString() + dr["FEC_EXPIRA_DOC"].ToString() + dr["FECHA_INICIO_RESIDENCIA_EN_EL_PAIS"].ToString()
								+ dr["PAIS_DE_NACIMIENTO"].ToString() + dr["COD_PROVINCIA"].ToString() + dr["COD_MUNICIPIO"].ToString()
								+ dr["COD_SEXO"].ToString() + dr["COD_ESTADO_CIVIL"].ToString() + dr["COD_NIVEL_ESTUDIOS"].ToString()
								+ dr["FEC_NACIMIENTO"].ToString() + dr["IND_VIVIENDA"].ToString() + dr["NUM_CARGAS"].ToString()
								+ dr["COD_SIT_LABORAL"].ToString() + dr["COD_CNAE"].ToString() + dr["NOM_DENOM_SOCIAL_A"].ToString() + dr["FEC_INGRESO"].ToString() + dr["COD_SECTOR_BANCO_E"].ToString()
								+ dr["INGRESOS_MENSUALES"].ToString() + dr["EGRESOS_MENSUALES"].ToString() + dr["TOTAL_DE_ACTIVOS"].ToString() + dr["TOTAL_DE_PASIVOS"].ToString()
								+ dr["COD_TIPO_DIREC"].ToString() + dr["COD_PAIS_D"].ToString() + dr["COD_PROVINCIA_DOM"].ToString() + dr["COD_MUNICIPIO_DOM"].ToString()
								+ dr["NOM_VIA"].ToString() + dr["MAS_DETALLE"].ToString() + dr["NOM_AMPL_DOMICILIO"].ToString() + dr["PREFIJO_INT_TF"].ToString()
								+ dr["PREFIJO_TF"].ToString() + dr["NUMERO_DE_TELEFONO"].ToString() + dr["NUMERO_DE_CELULAR"].ToString() + dr["E_MAIL"].ToString()
								+ dr["COD_REFERENCIA_1"].ToString() + dr["NUM_REFERENCIA_1"].ToString() + dr["NOM_REF_PERSONAL_1"].ToString() + dr["NUM_TFNO_REF_1"].ToString() + dr["COD_PROVINCIA_1"].ToString() + dr["COD_MUNICIPIO_1"].ToString()
								+ dr["COD_REFERENCIA_2"].ToString() + dr["NUM_REFERENCIA_2"].ToString() + dr["NOM_REF_PERSONAL_2"].ToString() + dr["NUM_TFNO_REF_2"].ToString() + dr["COD_PROVINCIA_2"].ToString() + dr["COD_MUNICIPIO_2"].ToString()
								+ dr["COD_CANTON"].ToString() + dr["COD_PARROQUIA"].ToString() + dr["TIPO_EMPRESA"].ToString() + dr["PRO_RELACION_COMERCIAL"].ToString();

							file.WriteLine(sLinea);

							i++;
						}

						int j = 0;
					

						MessageBox.Show(string.Format("{0} Tramas de Clientes {1} Tramas de Garantes Generadas Correctamente", i, j), "Informacion", MessageBoxButtons.OK, MessageBoxIcon.Information);
					}
					#endregion Clientes
				}
			}
			catch (Exception ex)
			{
				

				MessageBox.Show(string.Format("Commit Exception Type : {0} {1}", ex.GetType(), ex.Message), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);	
				return;
			}
		}

		private void txtBuscar_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
		
		}

		private void uGridOperaciones_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}
	}
}
