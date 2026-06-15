using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using DevExpress.XtraEditors;
using System.Data;
using System.Data.SqlClient;
using CrystalDecisions.Shared;
using Infragistics.Win.UltraWinGrid;
using System.Globalization;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de frmReporteFiniquitos.
	/// </summary>
	public class frmReporteFiniquitos : System.Windows.Forms.Form
	{
		private Infragistics.Win.UltraWinEditors.UltraOptionSet optEstados;
		private System.Windows.Forms.Label lblContador;
		private System.Windows.Forms.GroupBox groupBox1;
		private System.Windows.Forms.Button btnExportar;
		private System.Windows.Forms.Button btnVer;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtBuscar;
		private System.Windows.Forms.Label label2;
		private Infragistics.Win.UltraWinGrid.UltraGrid uGridDetalle;
		private C1.Data.C1DataSet cdsSeteoF;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource1;
		private System.Windows.Forms.Label label8;
		private System.Windows.Forms.Label label1;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtHasta;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtDesde;
		private System.Windows.Forms.Label lblSumatoria;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtVerde;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtNaranja;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtRojo;
		private System.Windows.Forms.Label label5;
		private System.Windows.Forms.Label label6;
		private System.Windows.Forms.Label label7;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource2;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource3;
		private System.Windows.Forms.Label label9;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbMotivosDeSalida;
		private System.Windows.Forms.Label lblInicia;
		private System.Windows.Forms.Label lblSalida;
		private System.Windows.Forms.Label lblMedio;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public frmReporteFiniquitos()
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

		int diasInicia = 10;
		int ddiasMedioI = 11;
		int dDiasMedioF= 15;
		int dDiasFinal = 16;

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
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmReporteFiniquitos));
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombres");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cedula");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cargo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Departamento");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Ciudad");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaEntrada");
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaSalida");
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Dias");
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaFiniquito");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("TotalPagar");
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn11 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaEgreso");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn12 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ComprobanteEgreso");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn13 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("PagoRealizado");
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn14 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Motivo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn15 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("MT");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn16 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idRolPrestamo");
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings1 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "TotalPagar", 9, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "TotalPagar", 9, true);
			Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance11 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance12 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance13 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance14 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance15 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance16 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance17 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombres");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Cedula");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn3 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Cargo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn4 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Departamento");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn5 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Ciudad");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn6 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("FechaEntrada");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn7 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("FechaSalida");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn8 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Dias");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn9 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("FechaFiniquito");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn10 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("TotalPagar");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn11 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("FechaEgreso");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn12 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("ComprobanteEgreso");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn13 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("PagoRealizado");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn14 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Motivo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn15 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("MT");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn16 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idRolPrestamo");
			Infragistics.Win.Appearance appearance18 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton1 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.Appearance appearance19 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton2 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.Appearance appearance20 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance21 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance22 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn17 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idNominaMotivosDeSalida");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn18 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Motivo");
			Infragistics.Win.Appearance appearance23 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn17 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idNominaMotivosDeSalida");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn18 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Motivo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn19 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idNominaMotivosDeSalida");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn20 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Motivo");
			this.optEstados = new Infragistics.Win.UltraWinEditors.UltraOptionSet();
			this.lblContador = new System.Windows.Forms.Label();
			this.groupBox1 = new System.Windows.Forms.GroupBox();
			this.btnExportar = new System.Windows.Forms.Button();
			this.btnVer = new System.Windows.Forms.Button();
			this.txtBuscar = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label2 = new System.Windows.Forms.Label();
			this.uGridDetalle = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource1 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.label8 = new System.Windows.Forms.Label();
			this.label1 = new System.Windows.Forms.Label();
			this.dtHasta = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.dtDesde = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.lblInicia = new System.Windows.Forms.Label();
			this.lblSalida = new System.Windows.Forms.Label();
			this.lblMedio = new System.Windows.Forms.Label();
			this.lblSumatoria = new System.Windows.Forms.Label();
			this.txtVerde = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtNaranja = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtRojo = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label5 = new System.Windows.Forms.Label();
			this.label6 = new System.Windows.Forms.Label();
			this.label7 = new System.Windows.Forms.Label();
			this.ultraDataSource2 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.cmbMotivosDeSalida = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.ultraDataSource3 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.label9 = new System.Windows.Forms.Label();
			((System.ComponentModel.ISupportInitialize)(this.optEstados)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtBuscar)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridDetalle)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtHasta)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtDesde)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtVerde)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNaranja)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtRojo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbMotivosDeSalida)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource3)).BeginInit();
			this.SuspendLayout();
			// 
			// optEstados
			// 
			appearance1.ForeColorDisabled = System.Drawing.Color.Black;
			this.optEstados.Appearance = appearance1;
			this.optEstados.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.optEstados.CheckedIndex = 0;
			this.optEstados.ItemAppearance = appearance2;
			this.optEstados.ItemOrigin = new System.Drawing.Point(10, 0);
			valueListItem1.DataValue = 0;
			valueListItem1.DisplayText = "Pendiente";
			valueListItem2.DataValue = 2;
			valueListItem2.DisplayText = "Pagado";
			this.optEstados.Items.Add(valueListItem1);
			this.optEstados.Items.Add(valueListItem2);
			this.optEstados.ItemSpacingVertical = 10;
			this.optEstados.Location = new System.Drawing.Point(440, 5);
			this.optEstados.Name = "optEstados";
			this.optEstados.Size = new System.Drawing.Size(160, 24);
			this.optEstados.TabIndex = 714;
			this.optEstados.Text = "Pendiente";
			// 
			// lblContador
			// 
			this.lblContador.AutoSize = true;
			this.lblContador.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblContador.ForeColor = System.Drawing.Color.Firebrick;
			this.lblContador.Location = new System.Drawing.Point(872, 56);
			this.lblContador.Name = "lblContador";
			this.lblContador.Size = new System.Drawing.Size(0, 17);
			this.lblContador.TabIndex = 713;
			// 
			// groupBox1
			// 
			this.groupBox1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.groupBox1.Location = new System.Drawing.Point(0, 104);
			this.groupBox1.Name = "groupBox1";
			this.groupBox1.Size = new System.Drawing.Size(1463, 8);
			this.groupBox1.TabIndex = 712;
			this.groupBox1.TabStop = false;
			// 
			// btnExportar
			// 
			this.btnExportar.CausesValidation = false;
			this.btnExportar.Enabled = false;
			this.btnExportar.Image = ((System.Drawing.Image)(resources.GetObject("btnExportar.Image")));
			this.btnExportar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnExportar.Location = new System.Drawing.Point(440, 40);
			this.btnExportar.Name = "btnExportar";
			this.btnExportar.Size = new System.Drawing.Size(72, 23);
			this.btnExportar.TabIndex = 709;
			this.btnExportar.Text = "Exportar";
			this.btnExportar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnExportar.Click += new System.EventHandler(this.btnExportar_Click);
			// 
			// btnVer
			// 
			this.btnVer.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(230)), ((System.Byte)(230)), ((System.Byte)(230)));
			this.btnVer.CausesValidation = false;
			this.btnVer.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.btnVer.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(50)), ((System.Byte)(50)), ((System.Byte)(50)));
			this.btnVer.Image = ((System.Drawing.Image)(resources.GetObject("btnVer.Image")));
			this.btnVer.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnVer.Location = new System.Drawing.Point(440, 72);
			this.btnVer.Name = "btnVer";
			this.btnVer.Size = new System.Drawing.Size(72, 23);
			this.btnVer.TabIndex = 708;
			this.btnVer.Text = "Ver";
			this.btnVer.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnVer.Click += new System.EventHandler(this.btnVer_Click);
			// 
			// txtBuscar
			// 
			appearance3.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtBuscar.Appearance = appearance3;
			this.txtBuscar.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtBuscar.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtBuscar.Location = new System.Drawing.Point(104, 73);
			this.txtBuscar.MaxLength = 300;
			this.txtBuscar.Name = "txtBuscar";
			this.txtBuscar.Size = new System.Drawing.Size(296, 21);
			this.txtBuscar.TabIndex = 707;
			// 
			// label2
			// 
			this.label2.AutoSize = true;
			this.label2.Location = new System.Drawing.Point(8, 75);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(84, 16);
			this.label2.TabIndex = 711;
			this.label2.Text = "Cedula/Nombre";
			this.label2.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// uGridDetalle
			// 
			this.uGridDetalle.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.uGridDetalle.Cursor = System.Windows.Forms.Cursors.Default;
			this.uGridDetalle.DataSource = this.ultraDataSource1;
			appearance4.BackColor = System.Drawing.Color.White;
			this.uGridDetalle.DisplayLayout.Appearance = appearance4;
			ultraGridColumn1.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn1.Width = 200;
			ultraGridColumn2.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridColumn2.Width = 97;
			ultraGridColumn3.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn3.Header.VisiblePosition = 6;
			ultraGridColumn3.Width = 143;
			ultraGridColumn4.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn4.Header.VisiblePosition = 2;
			ultraGridColumn5.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn5.Header.VisiblePosition = 3;
			ultraGridColumn6.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance5.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn6.CellAppearance = appearance5;
			ultraGridColumn6.Header.Caption = "Fecha de Entrada";
			ultraGridColumn6.Header.VisiblePosition = 4;
			ultraGridColumn6.Width = 125;
			ultraGridColumn7.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance6.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn7.CellAppearance = appearance6;
			ultraGridColumn7.Header.Caption = "Fecha de Salida";
			ultraGridColumn7.Header.VisiblePosition = 5;
			ultraGridColumn7.Width = 115;
			ultraGridColumn8.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance7.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn8.CellAppearance = appearance7;
			ultraGridColumn8.Header.VisiblePosition = 8;
			ultraGridColumn9.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn9.Format = "#,##0.00";
			ultraGridColumn9.Header.Caption = "Fecha de Finiquito";
			ultraGridColumn9.Header.VisiblePosition = 9;
			ultraGridColumn9.Width = 130;
			ultraGridColumn10.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance8.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn10.CellAppearance = appearance8;
			ultraGridColumn10.Format = "#,##0.00";
			ultraGridColumn10.Header.Caption = "Total a Pagar ";
			ultraGridColumn10.Header.VisiblePosition = 10;
			ultraGridColumn10.Width = 100;
			ultraGridColumn11.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn11.Header.Caption = "Fecha de Egreso";
			ultraGridColumn11.Header.VisiblePosition = 11;
			ultraGridColumn11.Width = 120;
			ultraGridColumn12.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn12.Header.Caption = "Comprobante de Egreso";
			ultraGridColumn12.Header.VisiblePosition = 12;
			ultraGridColumn12.Width = 105;
			ultraGridColumn13.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance9.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn13.CellAppearance = appearance9;
			ultraGridColumn13.Header.Caption = "Pago Realizado";
			ultraGridColumn13.Header.VisiblePosition = 13;
			ultraGridColumn14.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn14.Header.VisiblePosition = 7;
			ultraGridColumn15.Header.VisiblePosition = 14;
			ultraGridColumn16.Header.VisiblePosition = 15;
			ultraGridColumn16.Hidden = true;
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
																										 ultraGridColumn12,
																										 ultraGridColumn13,
																										 ultraGridColumn14,
																										 ultraGridColumn15,
																										 ultraGridColumn16});
			ultraGridBand1.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.Edit;
			appearance10.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings1.Appearance = appearance10;
			summarySettings1.DisplayFormat = "{0: #,##0.00}";
			summarySettings1.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			ultraGridBand1.Summaries.AddRange(new Infragistics.Win.UltraWinGrid.SummarySettings[] {
																																															summarySettings1});
			this.uGridDetalle.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.uGridDetalle.DisplayLayout.GroupByBox.Hidden = true;
			appearance11.ForeColor = System.Drawing.Color.Black;
			appearance11.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridDetalle.DisplayLayout.Override.ActiveRowAppearance = appearance11;
			this.uGridDetalle.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
			this.uGridDetalle.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.uGridDetalle.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance12.BackColor = System.Drawing.Color.Transparent;
			this.uGridDetalle.DisplayLayout.Override.CardAreaAppearance = appearance12;
			appearance13.ForeColor = System.Drawing.Color.Black;
			appearance13.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridDetalle.DisplayLayout.Override.CellAppearance = appearance13;
			this.uGridDetalle.DisplayLayout.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.Edit;
			appearance14.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance14.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance14.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance14.FontData.BoldAsString = "True";
			appearance14.FontData.Name = "Arial";
			appearance14.FontData.SizeInPoints = 10F;
			appearance14.ForeColor = System.Drawing.Color.White;
			appearance14.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridDetalle.DisplayLayout.Override.HeaderAppearance = appearance14;
			appearance15.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance15.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance15.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridDetalle.DisplayLayout.Override.RowAlternateAppearance = appearance15;
			appearance16.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance16.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance16.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridDetalle.DisplayLayout.Override.RowSelectorAppearance = appearance16;
			appearance17.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance17.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance17.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridDetalle.DisplayLayout.Override.SelectedRowAppearance = appearance17;
			this.uGridDetalle.DisplayLayout.ViewStyleBand = Infragistics.Win.UltraWinGrid.ViewStyleBand.OutlookGroupBy;
			this.uGridDetalle.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridDetalle.Location = new System.Drawing.Point(8, 120);
			this.uGridDetalle.Name = "uGridDetalle";
			this.uGridDetalle.Size = new System.Drawing.Size(1229, 240);
			this.uGridDetalle.TabIndex = 710;
			this.uGridDetalle.CellChange += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.uGridDetalle_CellChange);
			// 
			// ultraDataSource1
			// 
			ultraDataColumn6.DataType = typeof(System.DateTime);
			ultraDataColumn7.DataType = typeof(System.DateTime);
			ultraDataColumn8.DataType = typeof(int);
			ultraDataColumn9.DataType = typeof(System.DateTime);
			ultraDataColumn10.DataType = typeof(System.Decimal);
			ultraDataColumn11.DataType = typeof(System.DateTime);
			ultraDataColumn13.DataType = typeof(bool);
			ultraDataColumn15.DataType = typeof(bool);
			ultraDataColumn16.DataType = typeof(int);
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
																																 ultraDataColumn11,
																																 ultraDataColumn12,
																																 ultraDataColumn13,
																																 ultraDataColumn14,
																																 ultraDataColumn15,
																																 ultraDataColumn16});
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
			// label8
			// 
			this.label8.AutoSize = true;
			this.label8.Location = new System.Drawing.Point(248, 13);
			this.label8.Name = "label8";
			this.label8.Size = new System.Drawing.Size(34, 16);
			this.label8.TabIndex = 756;
			this.label8.Text = "Hasta";
			// 
			// label1
			// 
			this.label1.AutoSize = true;
			this.label1.Location = new System.Drawing.Point(16, 13);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(37, 16);
			this.label1.TabIndex = 755;
			this.label1.Text = "Desde";
			// 
			// dtHasta
			// 
			appearance18.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtHasta.Appearance = appearance18;
			this.dtHasta.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton1.Caption = "Today";
			this.dtHasta.DateButtons.Add(dateButton1);
			this.dtHasta.Format = "dd/MM/yyyy";
			this.dtHasta.Location = new System.Drawing.Point(296, 8);
			this.dtHasta.Name = "dtHasta";
			this.dtHasta.NonAutoSizeHeight = 23;
			this.dtHasta.Size = new System.Drawing.Size(104, 21);
			this.dtHasta.SpinButtonsVisible = true;
			this.dtHasta.TabIndex = 754;
			this.dtHasta.Value = ((object)(resources.GetObject("dtHasta.Value")));
			// 
			// dtDesde
			// 
			appearance19.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtDesde.Appearance = appearance19;
			this.dtDesde.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton2.Caption = "Today";
			this.dtDesde.DateButtons.Add(dateButton2);
			this.dtDesde.Format = "dd/MM/yyyy";
			this.dtDesde.Location = new System.Drawing.Point(104, 8);
			this.dtDesde.Name = "dtDesde";
			this.dtDesde.NonAutoSizeHeight = 23;
			this.dtDesde.Size = new System.Drawing.Size(104, 21);
			this.dtDesde.SpinButtonsVisible = true;
			this.dtDesde.TabIndex = 753;
			this.dtDesde.Value = ((object)(resources.GetObject("dtDesde.Value")));
			// 
			// lblInicia
			// 
			this.lblInicia.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.lblInicia.BackColor = System.Drawing.Color.Green;
			this.lblInicia.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
			this.lblInicia.Font = new System.Drawing.Font("Tahoma", 10F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblInicia.ForeColor = System.Drawing.Color.Black;
			this.lblInicia.Location = new System.Drawing.Point(864, 8);
			this.lblInicia.Name = "lblInicia";
			this.lblInicia.Size = new System.Drawing.Size(120, 40);
			this.lblInicia.TabIndex = 788;
			this.lblInicia.Text = "Dia Salida hasta 10 días";
			this.lblInicia.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			this.lblInicia.Click += new System.EventHandler(this.label3_Click);
			// 
			// lblSalida
			// 
			this.lblSalida.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.lblSalida.BackColor = System.Drawing.Color.Red;
			this.lblSalida.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
			this.lblSalida.Font = new System.Drawing.Font("Tahoma", 10F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblSalida.ForeColor = System.Drawing.Color.Black;
			this.lblSalida.Location = new System.Drawing.Point(1120, 8);
			this.lblSalida.Name = "lblSalida";
			this.lblSalida.Size = new System.Drawing.Size(120, 40);
			this.lblSalida.TabIndex = 787;
			this.lblSalida.Text = "16 días en adelante ";
			this.lblSalida.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			this.lblSalida.Click += new System.EventHandler(this.label4_Click);
			// 
			// lblMedio
			// 
			this.lblMedio.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.lblMedio.BackColor = System.Drawing.Color.DarkOrange;
			this.lblMedio.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
			this.lblMedio.Font = new System.Drawing.Font("Tahoma", 10F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblMedio.ForeColor = System.Drawing.Color.Black;
			this.lblMedio.Location = new System.Drawing.Point(992, 8);
			this.lblMedio.Name = "lblMedio";
			this.lblMedio.Size = new System.Drawing.Size(120, 40);
			this.lblMedio.TabIndex = 786;
			this.lblMedio.Text = "De 11 a 15 Días ";
			this.lblMedio.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			this.lblMedio.Click += new System.EventHandler(this.lblUsuario_Click);
			// 
			// lblSumatoria
			// 
			this.lblSumatoria.AutoSize = true;
			this.lblSumatoria.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblSumatoria.ForeColor = System.Drawing.Color.Firebrick;
			this.lblSumatoria.Location = new System.Drawing.Point(624, 40);
			this.lblSumatoria.Name = "lblSumatoria";
			this.lblSumatoria.Size = new System.Drawing.Size(0, 17);
			this.lblSumatoria.TabIndex = 789;
			// 
			// txtVerde
			// 
			appearance20.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtVerde.Appearance = appearance20;
			this.txtVerde.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtVerde.Enabled = false;
			this.txtVerde.FormatString = "#,##0.00";
			this.txtVerde.Location = new System.Drawing.Point(760, 8);
			this.txtVerde.Name = "txtVerde";
			this.txtVerde.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtVerde.PromptChar = ' ';
			this.txtVerde.Size = new System.Drawing.Size(90, 21);
			this.txtVerde.TabIndex = 790;
			// 
			// txtNaranja
			// 
			appearance21.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtNaranja.Appearance = appearance21;
			this.txtNaranja.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNaranja.Enabled = false;
			this.txtNaranja.FormatString = "#,##0.00";
			this.txtNaranja.Location = new System.Drawing.Point(760, 35);
			this.txtNaranja.Name = "txtNaranja";
			this.txtNaranja.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtNaranja.PromptChar = ' ';
			this.txtNaranja.Size = new System.Drawing.Size(90, 21);
			this.txtNaranja.TabIndex = 791;
			// 
			// txtRojo
			// 
			appearance22.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtRojo.Appearance = appearance22;
			this.txtRojo.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtRojo.Enabled = false;
			this.txtRojo.FormatString = "#,##0.00";
			this.txtRojo.Location = new System.Drawing.Point(760, 62);
			this.txtRojo.Name = "txtRojo";
			this.txtRojo.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtRojo.PromptChar = ' ';
			this.txtRojo.Size = new System.Drawing.Size(90, 21);
			this.txtRojo.TabIndex = 792;
			// 
			// label5
			// 
			this.label5.AutoSize = true;
			this.label5.Location = new System.Drawing.Point(648, 8);
			this.label5.Name = "label5";
			this.label5.Size = new System.Drawing.Size(106, 16);
			this.label5.TabIndex = 793;
			this.label5.Text = "Total  de 1 a 10 días";
			// 
			// label6
			// 
			this.label6.AutoSize = true;
			this.label6.Location = new System.Drawing.Point(648, 36);
			this.label6.Name = "label6";
			this.label6.Size = new System.Drawing.Size(109, 16);
			this.label6.TabIndex = 794;
			this.label6.Text = "Total de 11 a 15 días";
			this.label6.Click += new System.EventHandler(this.label6_Click);
			// 
			// label7
			// 
			this.label7.AutoSize = true;
			this.label7.Location = new System.Drawing.Point(648, 64);
			this.label7.Name = "label7";
			this.label7.Size = new System.Drawing.Size(112, 16);
			this.label7.TabIndex = 795;
			this.label7.Text = "Total  más de 15 días";
			// 
			// ultraDataSource2
			// 
			ultraDataColumn17.DataType = typeof(int);
			this.ultraDataSource2.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn17,
																																 ultraDataColumn18});
			// 
			// cmbMotivosDeSalida
			// 
			appearance23.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbMotivosDeSalida.Appearance = appearance23;
			this.cmbMotivosDeSalida.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbMotivosDeSalida.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbMotivosDeSalida.DataSource = this.ultraDataSource3;
			ultraGridColumn17.Header.VisiblePosition = 0;
			ultraGridColumn17.Hidden = true;
			ultraGridColumn18.Header.VisiblePosition = 1;
			ultraGridColumn18.Width = 295;
			ultraGridBand2.Columns.AddRange(new object[] {
																										 ultraGridColumn17,
																										 ultraGridColumn18});
			this.cmbMotivosDeSalida.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			this.cmbMotivosDeSalida.DisplayMember = "Motivo";
			this.cmbMotivosDeSalida.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbMotivosDeSalida.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbMotivosDeSalida.Location = new System.Drawing.Point(104, 40);
			this.cmbMotivosDeSalida.Name = "cmbMotivosDeSalida";
			this.cmbMotivosDeSalida.Size = new System.Drawing.Size(296, 21);
			this.cmbMotivosDeSalida.TabIndex = 840;
			this.cmbMotivosDeSalida.ValueMember = "idNominaMotivosDeSalida";
			this.cmbMotivosDeSalida.ValueChanged += new System.EventHandler(this.cmbMotivosDeSalida_ValueChanged);
			// 
			// ultraDataSource3
			// 
			ultraDataColumn19.DataType = typeof(int);
			this.ultraDataSource3.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn19,
																																 ultraDataColumn20});
			// 
			// label9
			// 
			this.label9.AutoSize = true;
			this.label9.Location = new System.Drawing.Point(16, 40);
			this.label9.Name = "label9";
			this.label9.Size = new System.Drawing.Size(38, 16);
			this.label9.TabIndex = 841;
			this.label9.Text = "Motivo";
			this.label9.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// frmReporteFiniquitos
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(1245, 366);
			this.Controls.Add(this.label9);
			this.Controls.Add(this.cmbMotivosDeSalida);
			this.Controls.Add(this.label7);
			this.Controls.Add(this.label6);
			this.Controls.Add(this.label5);
			this.Controls.Add(this.txtRojo);
			this.Controls.Add(this.txtNaranja);
			this.Controls.Add(this.txtVerde);
			this.Controls.Add(this.lblSumatoria);
			this.Controls.Add(this.lblInicia);
			this.Controls.Add(this.lblSalida);
			this.Controls.Add(this.lblMedio);
			this.Controls.Add(this.label8);
			this.Controls.Add(this.label1);
			this.Controls.Add(this.dtHasta);
			this.Controls.Add(this.dtDesde);
			this.Controls.Add(this.optEstados);
			this.Controls.Add(this.lblContador);
			this.Controls.Add(this.groupBox1);
			this.Controls.Add(this.btnExportar);
			this.Controls.Add(this.btnVer);
			this.Controls.Add(this.txtBuscar);
			this.Controls.Add(this.label2);
			this.Controls.Add(this.uGridDetalle);
			this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
			this.KeyPreview = true;
			this.Name = "frmReporteFiniquitos";
			this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
			this.Text = "Detalle de Actas De Finiquito";
			this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.frmReporteFiniquitos_KeyDown);
			this.Load += new System.EventHandler(this.frmReporteFiniquitos_Load);
			((System.ComponentModel.ISupportInitialize)(this.optEstados)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtBuscar)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridDetalle)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtHasta)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtDesde)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtVerde)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNaranja)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtRojo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbMotivosDeSalida)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource3)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);	
		}

		private void frmReporteFiniquitos_Load(object sender, System.EventArgs e)
		{
			miAcceso = new Acceso(cdsSeteoF, "200313");

			if (!Funcion.Permiso("2048", cdsSeteoF))
			{				
				MessageBox.Show("No puede ingresar  ", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				BeginInvoke(new MethodInvoker(UnloadMe));
				return;
			}
			if (miAcceso.BExportar) this.btnExportar.Enabled = true;
       if (!miAcceso.Editar) this.uGridDetalle.DisplayLayout.Bands[0].Columns["MT"].Hidden = true;
			this.cmbMotivosDeSalida.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Exec Listas 53, 0");
			this.dtDesde.CalendarInfo.MinDate = DateTime.Parse("01/01/2018");
			this.dtHasta.CalendarInfo.MinDate = DateTime.Parse("01/01/2018");

			this.dtDesde.CalendarInfo.MaxDate = DateTime.Today;
			this.dtHasta.CalendarInfo.MaxDate = DateTime.Today;

			this.dtDesde.Value = DateTime.Today;
			this.dtHasta.Value = DateTime.Today;
		}

		private void btnVer_Click(object sender, System.EventArgs e)
		{
			#region Valida Fechas
			if (this.dtDesde.Value == System.DBNull.Value)
			{
				MessageBox.Show("Seleccione la Fecha Inicial", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.dtDesde.Focus();
				return;
			}

			if (this.dtHasta.Value == System.DBNull.Value)
			{
				MessageBox.Show("Seleccione la Fecha Final", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.dtHasta.Focus();
				return;
			}

			if (Convert.ToDateTime(this.dtDesde.Value) > Convert.ToDateTime(this.dtHasta.Value))
			{
				MessageBox.Show("La Fecha Inicial no puede ser mayor a la Fecha Final", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.dtHasta.Focus();
				return;
			}
			#endregion Valida Fechas

			int idMotivo = 0;				
						
			if (this.cmbMotivosDeSalida.ActiveRow != null) idMotivo = (int) this.cmbMotivosDeSalida.Value;

			string sSQL = string.Format("Exec ReporteDeFiniquitos '{0}',{1},'{2}','{3}',{4}", this.txtBuscar.Text.ToString(),(int)this.optEstados.Value,Convert.ToDateTime(this.dtDesde.Value).ToString("yyyyMMdd"), 
				Convert.ToDateTime(this.dtHasta.Value).ToString("yyyyMMdd"), idMotivo);
			this.uGridDetalle.DataSource = FuncionesProcedimientos.dtGeneral(sSQL);

			this.lblContador.Text = this.uGridDetalle.Rows.Count + " REGISTROS ENCONTRADOS";

			this.DiseñoGridColor(this.uGridDetalle);
			Total();
		}
		private void DiseñoGridColor(Infragistics.Win.UltraWinGrid.UltraGrid uGrid)
		{
			string sColumna = "";
			int iContColumnas = (int)this.uGridDetalle.DisplayLayout.Bands[0].Columns.Count;	
					
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridDetalle.Rows.All)
			{
				for (int i = 2; i < iContColumnas; i++)
				{

					sColumna = this.uGridDetalle.DisplayLayout.Bands[0].Columns[i].Header.Caption;			

					if (dr.Cells["Dias"].Value.ToString().Trim().Length == 0)  dr.Cells["Dias"].Value = 0;

					#region Rango Dias
					if( sColumna == "Dias")
					{
						if (dr.Cells[sColumna].Value.ToString().Trim().Length == 0) 
						{
							dr.Cells[sColumna].Appearance.BackColor = Color.White;
							dr.Cells[sColumna].Appearance.BackColor2 = Color.White;								
						}
						else if (dr.Cells[sColumna].Value.ToString().Trim().Length > 0) 
						{
							int iDias = Convert.ToInt32(dr.Cells["Dias"].Value);

							if(iDias >= 0 &&  iDias <= diasInicia)
							{
					
								dr.Cells[sColumna].Appearance.BackColor = Color.Green;
								dr.Cells[sColumna].Appearance.BackColor2 = Color.Green;
							}
							if(iDias >= ddiasMedioI &&  iDias <= dDiasMedioF)
							{
										
								dr.Cells[sColumna].Appearance.BackColor = Color.Orange;
								dr.Cells[sColumna].Appearance.BackColor2 = Color.Orange;
							}

							if(iDias >= dDiasFinal )
							{
										
								dr.Cells[sColumna].Appearance.BackColor = Color.Red;
								dr.Cells[sColumna].Appearance.BackColor2 = Color.Red;
							}
												
						}
					}
					#endregion Dias

					#region PagoRealizado
					if( sColumna == "Pago Realizado")
					{
						if(dr.Cells["PagoRealizado"].Value.ToString() == "SI")
						{
							dr.Cells["Dias"].Appearance.BackColor = Color.White;
							dr.Cells["Dias"].Appearance.BackColor2 = Color.White;
						}										
					}
					#endregion PagoRealizado

				}
			}			
		}

		private void lblUsuario_Click(object sender, System.EventArgs e)
		{
		
		}

		private void label3_Click(object sender, System.EventArgs e)
		{
		
		}

		private void label4_Click(object sender, System.EventArgs e)
		{
		}
		private void Total()
		{
			decimal dVerde = 0;
			decimal dNaranja = 0;
			decimal dRojo = 0;

			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in uGridDetalle.Rows.All)
			{
				if (dr.Cells["Dias"].Value != System.DBNull.Value) 
				{
					int iDias = Convert.ToInt32(dr.Cells["Dias"].Value);
					if(iDias >= 0 &&  iDias <= diasInicia)
					{
						dVerde += (decimal) dr.Cells["TotalPagar"].Value;
					}
					if(iDias >= ddiasMedioI &&  iDias <= dDiasMedioF)
					{
						dNaranja += (decimal) dr.Cells["TotalPagar"].Value;
					}
					if(iDias >= dDiasFinal )
					{
						dRojo += (decimal) dr.Cells["TotalPagar"].Value;
					}
				
				}
			}
			
			txtVerde.Value = Math.Round(dVerde, 2);
			txtNaranja.Value = Math.Round(dNaranja, 2);
			txtRojo.Value = Math.Round(dRojo, 2);

		}

		private void label6_Click(object sender, System.EventArgs e)
		{
		
		}

		private void btnExportar_Click(object sender, System.EventArgs e)
		{
			if (this.uGridDetalle.Rows.Count == 0)
			{
				MessageBox.Show("No hay Filas para Exportar", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);				
			}
			else
			{
				FuncionesProcedimientos.ExportaExcel(this.uGridDetalle);
			}
		}

		bool bActualiza = false;
		private void uGridDetalle_CellChange(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			if (bActualiza) return;
			if (e.Cell.Column.ToString() == "MT")
			{
				if (e.Cell.Row.Cells["PagoRealizado"].Value != System.DBNull.Value && (bool) e.Cell.Row.Cells["PagoRealizado"].Value == true && (bool)e.Cell.Row.Cells["MT"].Value == false )
				{
					bActualiza = true;
					this.uGridDetalle.UpdateData();
					bool bDisponible = false;		
					if ((bool)e.Cell.Row.Cells["MT"].Value) 
					{
						bDisponible = true;
						int iMt = 0;
						if( bDisponible ) iMt = 1;
						Funcion.EjecutaSQL(cdsSeteoF, string.Format("update RolPrestamo set MinisterioTrabajo = {0} where idRolPrestamo = {1}", iMt, (int)e.Cell.Row.Cells["idRolPrestamo"].Value));
							
					}
					else 
					{
						bDisponible = false;
					}


				}
				else
				{
					bActualiza = true;
					this.uGridDetalle.UpdateData();
					e.Cell.Row.Cells["MT"].Value = (bool)e.Cell.OriginalValue;
				}
				bActualiza = false;
			}
		}

		private void frmReporteFiniquitos_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Escape) this.Close();		
		}

		private void cmbMotivosDeSalida_ValueChanged(object sender, System.EventArgs e)
		{
			if (this.cmbMotivosDeSalida.ActiveRow != null)
			{
				int idMotivo = 0;				
						
				if (this.cmbMotivosDeSalida.ActiveRow != null) idMotivo = (int) this.cmbMotivosDeSalida.Value;
				if(idMotivo == 0 || idMotivo == 1 )
				{
					diasInicia = 10;
					ddiasMedioI = 11;
					dDiasMedioF= 15;
					dDiasFinal = 16;
					this.lblInicia.Text ="Dia Salida hasta 10 días";
					this.lblMedio.Text = "De 11 a 15 Días ";
					this.lblSalida.Text ="16 días en adelante ";

					this.label5.Text ="Total  de 1 a 10 días";
					this.label6.Text = "Total de 11 a 15 días";
					this.label7.Text ="Total  más de 15 días";
				}
				else
				{
					diasInicia = 20;
					ddiasMedioI = 21;
					dDiasMedioF= 29;
					dDiasFinal = 30;
					this.lblInicia.Text ="Dia Salida hasta 20 días";
					this.lblMedio.Text = "De 21 a 29 Días ";
					this.lblSalida.Text ="30 días en adelante ";

					this.label5.Text ="Total  de 1 a 20 días";
					this.label6.Text = "Total de 21 a 29 días";
					this.label7.Text ="Total  más de 30 días";
				} 
			}
			else
			{
				diasInicia = 10;
				ddiasMedioI = 11;
				dDiasMedioF= 15;
				dDiasFinal = 16;
				this.lblInicia.Text ="Dia Salida hasta 10 días";
				this.lblMedio.Text = "De 11 a 15 Días ";
				this.lblSalida.Text ="16 días en adelante ";
				this.label5.Text ="Total  de 1 a 10 días";
				this.label6.Text = "Total de 11 a 15 días";
				this.label7.Text ="Total  más de 15 días";

			}
		}

	}
}
