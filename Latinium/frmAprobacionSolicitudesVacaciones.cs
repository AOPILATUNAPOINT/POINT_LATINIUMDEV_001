using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;
using System.Data.SqlClient;
using DevExpress.XtraEditors;
using CrystalDecisions;
using CrystalDecisions.CrystalReports.Engine;
using CrystalDecisions.ReportSource;
using CrystalDecisions.Shared;


namespace Latinium
{
	/// <summary>
	/// Descripción breve de frmAprobacionSolicitudesVacaciones.
	/// </summary>
	public class frmAprobacionSolicitudesVacaciones : System.Windows.Forms.Form
	{
		private Infragistics.Win.UltraWinGrid.UltraGrid ultraGrid1;
		private System.Windows.Forms.GroupBox groupBox1;
		private System.Windows.Forms.Label lblContador;
		private System.Windows.Forms.Button btnBuscar;
		private System.Windows.Forms.Label lblCliente;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtBusqueda;
		private System.Windows.Forms.Label label8;
		private System.Windows.Forms.Label label6;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtHasta;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtDesde;
		private C1.Data.C1DataSet cdsSeteoF;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource4;
		private System.Windows.Forms.Button btnExcel;
		private Infragistics.Win.UltraWinEditors.UltraOptionSet optEstados;
		bool bEdicion = false;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public frmAprobacionSolicitudesVacaciones()
		{
			//
			// Necesario para admitir el Diseñador de Windows Forms
			//
			InitializeComponent();

			//
			// TODO: agregar código de constructor después de llamar a InitializeComponent
			//
		}
		DateTime dtUltimoDiaMes = DateTime.Today;
		int idPersonal;
		int idCom_AsignacionDeVendedores;
		int iBodega;
		string sPeriodo;
		int idCom_Rango = 0;
		int idCargo = 0;
		int idCom_Estado;


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
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idsolicitudVacaciones");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Fecha");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Personal_Solicitante");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPersonal");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaSalida");
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaRetorna");
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("N_dias", -1, null, 0, Infragistics.Win.UltraWinGrid.SortIndicator.Ascending, false);
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Observacion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Estado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Motivo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn11 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("observacionAprobacion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn12 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("observacionRechazo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn13 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Aceptar", 0);
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmAprobacionSolicitudesVacaciones));
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn14 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Anular", 1);
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn15 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Imprimir", 2);
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn16 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Supervisor", 3);
			Infragistics.Win.Appearance appearance11 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance12 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance13 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance14 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance15 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance16 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance17 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance18 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idsolicitudVacaciones");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Fecha");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn3 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Personal_Solicitante");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn4 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idPersonal");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn5 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("FechaSalida");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn6 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("FechaRetorna");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn7 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("N_dias");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn8 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Observacion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn9 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Estado");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn10 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Motivo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn11 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("observacionAprobacion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn12 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("observacionRechazo");
			Infragistics.Win.Appearance appearance19 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance20 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton1 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.Appearance appearance21 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton2 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.Appearance appearance22 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance23 = new Infragistics.Win.Appearance();
			Infragistics.Win.ValueListItem valueListItem1 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem2 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem3 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem4 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem5 = new Infragistics.Win.ValueListItem();
			this.ultraGrid1 = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource4 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.groupBox1 = new System.Windows.Forms.GroupBox();
			this.lblContador = new System.Windows.Forms.Label();
			this.btnBuscar = new System.Windows.Forms.Button();
			this.lblCliente = new System.Windows.Forms.Label();
			this.txtBusqueda = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label8 = new System.Windows.Forms.Label();
			this.label6 = new System.Windows.Forms.Label();
			this.dtHasta = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.dtDesde = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.btnExcel = new System.Windows.Forms.Button();
			this.optEstados = new Infragistics.Win.UltraWinEditors.UltraOptionSet();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource4)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtBusqueda)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtHasta)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtDesde)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.optEstados)).BeginInit();
			this.SuspendLayout();
			// 
			// ultraGrid1
			// 
			this.ultraGrid1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.ultraGrid1.Cursor = System.Windows.Forms.Cursors.Default;
			this.ultraGrid1.DataSource = this.ultraDataSource4;
			appearance1.BackColor = System.Drawing.Color.White;
			appearance1.ForeColor = System.Drawing.Color.Black;
			appearance1.ForeColorDisabled = System.Drawing.Color.Black;
			this.ultraGrid1.DisplayLayout.Appearance = appearance1;
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn2.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
			ultraGridColumn2.Header.VisiblePosition = 7;
			ultraGridColumn2.Width = 70;
			ultraGridColumn3.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn3.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
			ultraGridColumn3.Header.Caption = "Personal Solicitante";
			ultraGridColumn3.Header.VisiblePosition = 8;
			ultraGridColumn3.Width = 230;
			ultraGridColumn4.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
			ultraGridColumn4.Header.VisiblePosition = 1;
			ultraGridColumn4.Hidden = true;
			ultraGridColumn5.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
			appearance2.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn5.CellAppearance = appearance2;
			ultraGridColumn5.Header.Caption = "Desde";
			ultraGridColumn5.Header.VisiblePosition = 13;
			ultraGridColumn5.Width = 95;
			ultraGridColumn6.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
			appearance3.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn6.CellAppearance = appearance3;
			ultraGridColumn6.Header.Caption = "Hasta";
			ultraGridColumn6.Header.VisiblePosition = 14;
			ultraGridColumn6.Width = 95;
			ultraGridColumn7.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
			appearance4.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn7.CellAppearance = appearance4;
			ultraGridColumn7.Header.Caption = "D Solicitados";
			ultraGridColumn7.Header.VisiblePosition = 10;
			ultraGridColumn8.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn8.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn8.Header.VisiblePosition = 15;
			ultraGridColumn8.Width = 190;
			ultraGridColumn9.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
			ultraGridColumn9.Header.VisiblePosition = 6;
			ultraGridColumn10.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn10.Header.VisiblePosition = 9;
			ultraGridColumn10.Hidden = true;
			ultraGridColumn11.Header.VisiblePosition = 11;
			ultraGridColumn11.Hidden = true;
			ultraGridColumn12.Header.VisiblePosition = 12;
			ultraGridColumn12.Hidden = true;
			ultraGridColumn13.ButtonDisplayStyle = Infragistics.Win.UltraWinGrid.ButtonDisplayStyle.Always;
			appearance5.ImageHAlign = Infragistics.Win.HAlign.Center;
			appearance5.ImageVAlign = Infragistics.Win.VAlign.Middle;
			ultraGridColumn13.CellAppearance = appearance5;
			appearance6.Image = ((object)(resources.GetObject("appearance6.Image")));
			appearance6.ImageHAlign = Infragistics.Win.HAlign.Center;
			appearance6.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn13.CellButtonAppearance = appearance6;
			ultraGridColumn13.Header.Caption = "Validar";
			ultraGridColumn13.Header.VisiblePosition = 3;
			ultraGridColumn13.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.Button;
			ultraGridColumn13.Width = 60;
			ultraGridColumn14.ButtonDisplayStyle = Infragistics.Win.UltraWinGrid.ButtonDisplayStyle.Always;
			appearance7.ImageHAlign = Infragistics.Win.HAlign.Center;
			appearance7.ImageVAlign = Infragistics.Win.VAlign.Middle;
			ultraGridColumn14.CellAppearance = appearance7;
			appearance8.ForeColorDisabled = System.Drawing.Color.Black;
			appearance8.Image = ((object)(resources.GetObject("appearance8.Image")));
			appearance8.ImageHAlign = Infragistics.Win.HAlign.Center;
			appearance8.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn14.CellButtonAppearance = appearance8;
			ultraGridColumn14.Header.VisiblePosition = 4;
			ultraGridColumn14.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.Button;
			ultraGridColumn14.Width = 60;
			ultraGridColumn15.ButtonDisplayStyle = Infragistics.Win.UltraWinGrid.ButtonDisplayStyle.Always;
			appearance9.Image = ((object)(resources.GetObject("appearance9.Image")));
			appearance9.ImageHAlign = Infragistics.Win.HAlign.Center;
			appearance9.ImageVAlign = Infragistics.Win.VAlign.Middle;
			ultraGridColumn15.CellAppearance = appearance9;
			appearance10.Image = ((object)(resources.GetObject("appearance10.Image")));
			appearance10.ImageHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn15.CellButtonAppearance = appearance10;
			ultraGridColumn15.Header.VisiblePosition = 5;
			ultraGridColumn15.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.Button;
			ultraGridColumn15.Width = 60;
			ultraGridColumn16.ButtonDisplayStyle = Infragistics.Win.UltraWinGrid.ButtonDisplayStyle.Always;
			appearance11.ImageHAlign = Infragistics.Win.HAlign.Center;
			appearance11.ImageVAlign = Infragistics.Win.VAlign.Middle;
			ultraGridColumn16.CellAppearance = appearance11;
			appearance12.Image = ((object)(resources.GetObject("appearance12.Image")));
			appearance12.ImageHAlign = Infragistics.Win.HAlign.Center;
			appearance12.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn16.CellButtonAppearance = appearance12;
			ultraGridColumn16.Header.Caption = "Jefe de área ";
			ultraGridColumn16.Header.VisiblePosition = 2;
			ultraGridColumn16.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.Button;
			ultraGridColumn16.Width = 60;
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
			ultraGridBand1.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			ultraGridBand1.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			this.ultraGrid1.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			appearance13.ForeColor = System.Drawing.Color.Black;
			this.ultraGrid1.DisplayLayout.CaptionAppearance = appearance13;
			this.ultraGrid1.DisplayLayout.GroupByBox.Prompt = "Arrastre columna a Agrupar";
			appearance14.ForeColor = System.Drawing.Color.Black;
			appearance14.ForeColorDisabled = System.Drawing.Color.Black;
			this.ultraGrid1.DisplayLayout.Override.ActiveRowAppearance = appearance14;
			this.ultraGrid1.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
			this.ultraGrid1.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.ultraGrid1.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.False;
			appearance15.BackColor = System.Drawing.Color.Transparent;
			this.ultraGrid1.DisplayLayout.Override.CardAreaAppearance = appearance15;
			appearance16.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance16.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance16.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance16.FontData.BoldAsString = "True";
			appearance16.FontData.Name = "Arial";
			appearance16.FontData.SizeInPoints = 8F;
			appearance16.ForeColor = System.Drawing.Color.White;
			appearance16.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.ultraGrid1.DisplayLayout.Override.HeaderAppearance = appearance16;
			this.ultraGrid1.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti;
			appearance17.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance17.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance17.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.RowSelectorAppearance = appearance17;
			appearance18.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance18.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance18.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.SelectedRowAppearance = appearance18;
			this.ultraGrid1.DisplayLayout.Override.SelectTypeCell = Infragistics.Win.UltraWinGrid.SelectType.None;
			this.ultraGrid1.DisplayLayout.Override.SelectTypeCol = Infragistics.Win.UltraWinGrid.SelectType.None;
			this.ultraGrid1.DisplayLayout.Override.SelectTypeRow = Infragistics.Win.UltraWinGrid.SelectType.Single;
			this.ultraGrid1.DisplayLayout.ScrollBounds = Infragistics.Win.UltraWinGrid.ScrollBounds.ScrollToFill;
			this.ultraGrid1.DisplayLayout.ScrollStyle = Infragistics.Win.UltraWinGrid.ScrollStyle.Immediate;
			this.ultraGrid1.DisplayLayout.TabNavigation = Infragistics.Win.UltraWinGrid.TabNavigation.NextControl;
			this.ultraGrid1.DisplayLayout.ViewStyle = Infragistics.Win.UltraWinGrid.ViewStyle.SingleBand;
			this.ultraGrid1.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.ultraGrid1.Location = new System.Drawing.Point(8, 88);
			this.ultraGrid1.Name = "ultraGrid1";
			this.ultraGrid1.Size = new System.Drawing.Size(1048, 232);
			this.ultraGrid1.TabIndex = 1020;
			this.ultraGrid1.ClickCellButton += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.ultraGrid1_ClickCellButton);
			this.ultraGrid1.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.ultraGrid1_InitializeLayout);
			// 
			// ultraDataSource4
			// 
			ultraDataColumn1.DataType = typeof(int);
			ultraDataColumn2.DataType = typeof(System.DateTime);
			ultraDataColumn4.DataType = typeof(int);
			ultraDataColumn5.DataType = typeof(System.DateTime);
			ultraDataColumn6.DataType = typeof(System.DateTime);
			ultraDataColumn7.DataType = typeof(int);
			this.ultraDataSource4.Band.Columns.AddRange(new object[] {
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
			// groupBox1
			// 
			this.groupBox1.Location = new System.Drawing.Point(8, 120);
			this.groupBox1.Name = "groupBox1";
			this.groupBox1.Size = new System.Drawing.Size(1184, 8);
			this.groupBox1.TabIndex = 1019;
			this.groupBox1.TabStop = false;
			// 
			// lblContador
			// 
			this.lblContador.AutoSize = true;
			this.lblContador.Font = new System.Drawing.Font("Tahoma", 12F, System.Drawing.FontStyle.Bold);
			this.lblContador.ForeColor = System.Drawing.Color.Firebrick;
			this.lblContador.Location = new System.Drawing.Point(544, 56);
			this.lblContador.Name = "lblContador";
			this.lblContador.Size = new System.Drawing.Size(0, 23);
			this.lblContador.TabIndex = 1018;
			this.lblContador.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// btnBuscar
			// 
			this.btnBuscar.CausesValidation = false;
			this.btnBuscar.Location = new System.Drawing.Point(424, 55);
			this.btnBuscar.Name = "btnBuscar";
			this.btnBuscar.Size = new System.Drawing.Size(75, 24);
			this.btnBuscar.TabIndex = 1016;
			this.btnBuscar.Text = "Buscar";
			this.btnBuscar.Click += new System.EventHandler(this.btnBuscar_Click);
			// 
			// lblCliente
			// 
			this.lblCliente.AutoSize = true;
			this.lblCliente.Location = new System.Drawing.Point(24, 59);
			this.lblCliente.Name = "lblCliente";
			this.lblCliente.Size = new System.Drawing.Size(39, 16);
			this.lblCliente.TabIndex = 1017;
			this.lblCliente.Text = "Buscar";
			// 
			// txtBusqueda
			// 
			appearance19.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtBusqueda.Appearance = appearance19;
			this.txtBusqueda.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtBusqueda.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtBusqueda.Location = new System.Drawing.Point(88, 57);
			this.txtBusqueda.Name = "txtBusqueda";
			this.txtBusqueda.Size = new System.Drawing.Size(304, 21);
			this.txtBusqueda.TabIndex = 1015;
			// 
			// label8
			// 
			this.label8.AutoSize = true;
			this.label8.Location = new System.Drawing.Point(224, 24);
			this.label8.Name = "label8";
			this.label8.Size = new System.Drawing.Size(34, 16);
			this.label8.TabIndex = 1013;
			this.label8.Text = "Hasta";
			// 
			// label6
			// 
			this.label6.AutoSize = true;
			this.label6.Location = new System.Drawing.Point(24, 24);
			this.label6.Name = "label6";
			this.label6.Size = new System.Drawing.Size(37, 16);
			this.label6.TabIndex = 1012;
			this.label6.Text = "Desde";
			// 
			// dtHasta
			// 
			appearance20.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtHasta.Appearance = appearance20;
			this.dtHasta.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton1.Caption = "Today";
			this.dtHasta.DateButtons.Add(dateButton1);
			this.dtHasta.Format = "dd/MM/yyyy";
			this.dtHasta.Location = new System.Drawing.Point(288, 24);
			this.dtHasta.Name = "dtHasta";
			this.dtHasta.NonAutoSizeHeight = 23;
			this.dtHasta.Size = new System.Drawing.Size(104, 21);
			this.dtHasta.SpinButtonsVisible = true;
			this.dtHasta.TabIndex = 1011;
			this.dtHasta.Value = ((object)(resources.GetObject("dtHasta.Value")));
			// 
			// dtDesde
			// 
			appearance21.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtDesde.Appearance = appearance21;
			this.dtDesde.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton2.Caption = "Today";
			this.dtDesde.DateButtons.Add(dateButton2);
			this.dtDesde.Format = "dd/MM/yyyy";
			this.dtDesde.Location = new System.Drawing.Point(88, 24);
			this.dtDesde.Name = "dtDesde";
			this.dtDesde.NonAutoSizeHeight = 23;
			this.dtDesde.Size = new System.Drawing.Size(104, 21);
			this.dtDesde.SpinButtonsVisible = true;
			this.dtDesde.TabIndex = 1010;
			this.dtDesde.Value = ((object)(resources.GetObject("dtDesde.Value")));
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
			// btnExcel
			// 
			this.btnExcel.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btnExcel.CausesValidation = false;
			this.btnExcel.Image = ((System.Drawing.Image)(resources.GetObject("btnExcel.Image")));
			this.btnExcel.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnExcel.Location = new System.Drawing.Point(816, 56);
			this.btnExcel.Name = "btnExcel";
			this.btnExcel.Size = new System.Drawing.Size(72, 23);
			this.btnExcel.TabIndex = 1021;
			this.btnExcel.Text = "&Exportar ";
			this.btnExcel.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnExcel.Visible = false;
			this.btnExcel.Click += new System.EventHandler(this.btnExcel_Click);
			// 
			// optEstados
			// 
			appearance22.ForeColorDisabled = System.Drawing.Color.Black;
			appearance22.Image = ((object)(resources.GetObject("appearance22.Image")));
			appearance22.ImageHAlign = Infragistics.Win.HAlign.Center;
			appearance22.TextHAlign = Infragistics.Win.HAlign.Center;
			this.optEstados.Appearance = appearance22;
			this.optEstados.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.optEstados.CheckedIndex = 0;
			appearance23.Image = ((object)(resources.GetObject("appearance23.Image")));
			appearance23.ImageHAlign = Infragistics.Win.HAlign.Right;
			this.optEstados.ItemAppearance = appearance23;
			this.optEstados.ItemOrigin = new System.Drawing.Point(10, 0);
			valueListItem1.DataValue = 0;
			valueListItem1.DisplayText = "Todos";
			valueListItem2.DataValue = 1;
			valueListItem2.DisplayText = "Pendientes";
			valueListItem3.DataValue = 3;
			valueListItem3.DisplayText = "Anulados";
			valueListItem4.DataValue = 4;
			valueListItem4.DisplayText = "Validar";
			valueListItem5.DataValue = 2;
			valueListItem5.DisplayText = "Aprobado";
			this.optEstados.Items.Add(valueListItem1);
			this.optEstados.Items.Add(valueListItem2);
			this.optEstados.Items.Add(valueListItem3);
			this.optEstados.Items.Add(valueListItem4);
			this.optEstados.Items.Add(valueListItem5);
			this.optEstados.ItemSpacingVertical = 10;
			this.optEstados.Location = new System.Drawing.Point(424, 16);
			this.optEstados.Name = "optEstados";
			this.optEstados.Size = new System.Drawing.Size(464, 24);
			this.optEstados.TabIndex = 1022;
			this.optEstados.Text = "Todos";
			this.optEstados.ValueChanged += new System.EventHandler(this.optEstados_ValueChanged);
			// 
			// frmAprobacionSolicitudesVacaciones
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(1056, 326);
			this.Controls.Add(this.optEstados);
			this.Controls.Add(this.btnExcel);
			this.Controls.Add(this.ultraGrid1);
			this.Controls.Add(this.groupBox1);
			this.Controls.Add(this.lblContador);
			this.Controls.Add(this.btnBuscar);
			this.Controls.Add(this.lblCliente);
			this.Controls.Add(this.label8);
			this.Controls.Add(this.label6);
			this.Controls.Add(this.txtBusqueda);
			this.Controls.Add(this.dtHasta);
			this.Controls.Add(this.dtDesde);
			this.Name = "frmAprobacionSolicitudesVacaciones";
			this.Text = "Aprobacion Solicitudes de Vacaciones";
			this.Load += new System.EventHandler(this.frmAprobacionSolicitudesVacaciones_Load);
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource4)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtBusqueda)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtHasta)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtDesde)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.optEstados)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);
		}

		private void btnBuscar_Click(object sender, System.EventArgs e)
		{
			if (!Validacion.vbFechaEnDocumentos(this.dtDesde, "Seleccione la fecha inicial", false, "", cdsSeteoF)) return;
			if (!Validacion.vbFechaEnDocumentos(this.dtHasta, "Seleccione la fecha final", false, "", cdsSeteoF)) return;
			if (!Validacion.vbComparaFechas(Convert.ToDateTime(this.dtDesde.Value), Convert.ToDateTime(this.dtHasta.Value), this.dtDesde, "La fecha inicial no puede ser posterior a la fecha final", false)) return;
			this.lblContador.Text = "";
			string sSQL = string.Format("Exec ConsultaAprobacionVacaciones '{0}', '{1}', {2}", Convert.ToDateTime(this.dtDesde.Value).ToString("yyyyMMdd"),
				Convert.ToDateTime(this.dtHasta.Value).ToString("yyyyMMdd"),(int)this.optEstados.Value);
			FuncionesProcedimientos.dsGeneral(sSQL, this.ultraGrid1);

			this.lblContador.Text = this.ultraGrid1.Rows.Count + " Registros Encontrados";
			this.Seguridad ();
		}

		private void optEstados_ValueChanged(object sender, System.EventArgs e)
		{
			this.btnBuscar_Click (sender, e );
		}

		private void frmAprobacionSolicitudesVacaciones_Load(object sender, System.EventArgs e)
		{
			miAcceso = new Acceso(cdsSeteoF, "0227");

			if (!Funcion.Permiso("1417", cdsSeteoF))
			{				
				MessageBox.Show("No puede ingresar  ", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				BeginInvoke(new MethodInvoker(UnloadMe));
				return;
			}
			this.dtDesde.Value = DateTime.Today.AddDays(-30);
			this.dtHasta.Value = DateTime.Today;

			this.Seguridad ();

			if (miAcceso.BExportarExcel) this.btnExcel.Enabled = true;
			this.btnBuscar_Click (sender, e );
		
		
		}

		private void ultraGrid1_ClickCellButton(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			int iUltDiames = FuncionesProcedimientos.iUltimoDiaDelMes(dtUltimoDiaMes.Month, dtUltimoDiaMes.Year);
			#region Deposito
//			if (e.Cell.Column.ToString() == "btnMotivo")
//			{	
//				using (frmObservacionesVacaciones Observaciones = new frmObservacionesFlujo(1, e.Cell.Row.Cells["observacionAprobacion"].Value.ToString().Trim(), bEdicion))
//				{				
//					if (DialogResult.OK == Observaciones.ShowDialog())
//					{
//						e.Cell.Row.Cells["observacionAprobacion"].Value = Observaciones.txtObservaciones.Text.ToString().Trim();
//					}
//				}
//			}
			#endregion Deposito

			#region Anular
			if (e.Cell.Column.ToString() == "Anular")
			{	
				if (!miAcceso.BAnularAprobado) 
				{
					MessageBox.Show("No tiene acceso Anular ", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					return;
				}
				#region Abre observacion
				bEdicion=true;
				using (frmObservacionesVacaciones Observaciones = new frmObservacionesVacaciones(1, e.Cell.Row.Cells["observacionRechazo"].Value.ToString().Trim(), bEdicion))
				{				
					if (DialogResult.OK == Observaciones.ShowDialog())
					{
						e.Cell.Row.Cells["observacionRechazo"].Value = Observaciones.txtObservaciones.Text.ToString().Trim();
					}
				}
				#endregion Abre observacion

				#region Valida Observacion

				if (e.Cell.Row.Cells["observacionRechazo"].Value.ToString().Trim().Length == 0 )
				{
					MessageBox.Show("Ingrese un Observación", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					this.ultraGrid1.ActiveRow.Selected = true;
					return;
				}
				if (e.Cell.Row.Cells["observacionRechazo"].Value.ToString().Trim().Length < 10 )
				{
					MessageBox.Show("La Observación debe tener minimo 10 caracteres ", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					this.ultraGrid1.ActiveRow.Selected = true;
					return;
				}

				if (e.Cell.Row.Cells["observacionRechazo"].Value.ToString().Trim().Length > 100 )
				{
					MessageBox.Show("La Observación debe tener maximo 100 caracteres ", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					this.ultraGrid1.ActiveRow.Selected = true;
					return;
				}

				#endregion Valida Observacion
			
				#region Guardar      
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


						if ((string)e.Cell.Row.Cells["Estado"].Value=="AUTORIZADO")
						{

							string sSQLAutEs = string.Format("Update solicitudvacaciones set Estado = 3  where idsolicitudvacaciones = {0}",
								(int) e.Cell.Row.Cells["idsolicitudvacaciones"].Value);
							oCmdActualiza.CommandText = sSQLAutEs;
							oCmdActualiza.ExecuteNonQuery();		

							string sSQLAut = string.Format("Exec AnulaVacacaiones {0}",(int) e.Cell.Row.Cells["idsolicitudvacaciones"].Value);			
							oCmdActualiza.CommandText = sSQLAut;
							oCmdActualiza.ExecuteNonQuery();	
						}
						else {
							string sSQLAutEs = string.Format("Update solicitudvacaciones set Estado = 3  where idsolicitudvacaciones = {0}",
								(int) e.Cell.Row.Cells["idsolicitudvacaciones"].Value);
							oCmdActualiza.CommandText = sSQLAutEs;
							oCmdActualiza.ExecuteNonQuery();					
						}
		
						oTransaction.Commit();
						Cursor = Cursors.Default;	

						MessageBox.Show("Registro Guardado Correctamente", "Información", MessageBoxButtons.OK, MessageBoxIcon.Information);
						this.btnBuscar_Click (sender, e);
				
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
			#endregion Anular

			#region Supervisor
			if (e.Cell.Column.ToString() == "Supervisor")
			{	
				if (!miAcceso.Autorizacion) 
				{
					MessageBox.Show("No tiene acceso AUTORIZAR ", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					return;
				}
				#region Abre observacion
				bEdicion=true;
				using (frmObservacionesVacaciones Observaciones = new frmObservacionesVacaciones(1, e.Cell.Row.Cells["observacionAprobacion"].Value.ToString().Trim(), bEdicion))
				{				
					if (DialogResult.OK == Observaciones.ShowDialog())
					{
						e.Cell.Row.Cells["observacionAprobacion"].Value = Observaciones.txtObservaciones.Text.ToString().Trim();
					}
				}
				#endregion Abre observacion

				#region Valida Observacion

				if (e.Cell.Row.Cells["observacionAprobacion"].Value.ToString().Trim().Length == 0 )
				{
					MessageBox.Show("Ingrese un Observación", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					this.ultraGrid1.ActiveRow.Selected = true;
					return;
				}
				if (e.Cell.Row.Cells["observacionAprobacion"].Value.ToString().Trim().Length < 10 )
				{
					MessageBox.Show("La Observación debe tener minimo 10 caracteres ", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					this.ultraGrid1.ActiveRow.Selected = true;
					return;
				}

				if (e.Cell.Row.Cells["observacionAprobacion"].Value.ToString().Trim().Length > 100 )
				{
					MessageBox.Show("La Observación debe tener maximo 100 caracteres ", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					this.ultraGrid1.ActiveRow.Selected = true;
					return;
				}

				#endregion Valida Observacion

						
					#region Guardar      
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

				
					      
							string sSQLAutEs = string.Format("Update solicitudvacaciones set Estado = 4  where idsolicitudvacaciones = {0}",
								(int) e.Cell.Row.Cells["idsolicitudvacaciones"].Value);
			

							oCmdActualiza.CommandText = sSQLAutEs;
							oCmdActualiza.ExecuteNonQuery();		

							oTransaction.Commit();

							Cursor = Cursors.Default;	

				

							MessageBox.Show("Registro Guardado Correctamente", "Información", MessageBoxButtons.OK, MessageBoxIcon.Information);
							this.btnBuscar_Click (sender, e);
				
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
			#endregion supervisor 

			#region Aprobacion
			if (e.Cell.Column.ToString() == "Aceptar")
			{	
				if (!miAcceso.BGenerarCobro) 
				{
					MessageBox.Show("No tiene acceso AUTORIZAR ", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					return;
				}

				#region Guardar      
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
						int iVendedor = Funcion.iEscalarSQL(cdsSeteoF, string.Format("SELECT  count(*) FROM  Com_AsignacionDeVendedores WHERE   idPersonal = {0}",e.Cell.Row.Cells["idPersonal"].Value));
						if(iVendedor > 0)
						{
							#region Asignacion Vendedores 
							string sSQL = string.Format("Exec ConsultaSolicitaIndividualVacaciones {0}",Convert.ToInt32(e.Cell.Row.Cells["idPersonal"].Value));
							SqlDataReader dr = Funcion.rEscalarSQL(cdsSeteoF, sSQL, true);
							dr.Read();
							if (dr.HasRows)
							{
								if (dr.GetValue(0) != System.DBNull.Value) idPersonal = dr.GetInt32(0);
								if (dr.GetValue(1) != System.DBNull.Value) idCom_AsignacionDeVendedores = dr.GetInt32(1);
								if (dr.GetValue(2) != System.DBNull.Value) iBodega = dr.GetInt32(2); 
								if (dr.GetValue(3) != System.DBNull.Value) sPeriodo = dr.GetString(3);
								if (dr.GetValue(4) != System.DBNull.Value) idCargo = dr.GetInt32(4); 
								if (dr.GetValue(5) != System.DBNull.Value) idCom_Rango= dr.GetInt32(5);
								if (dr.GetValue(6) != System.DBNull.Value) idCom_Estado= dr.GetInt32(6);
							}
							dr.Close();

							#region Permisos						
							using (frmNominaPermisosFaltas NPF = new frmNominaPermisosFaltas(idPersonal,idCom_AsignacionDeVendedores,iBodega,Convert.ToString(sPeriodo),2,Convert.ToDateTime(e.Cell.Row.Cells["FechaSalida"].Value),Convert.ToDateTime(e.Cell.Row.Cells["FechaRetorna"].Value), dtUltimoDiaMes,idCargo, idCom_Rango,idCom_Estado))
							{
								if (DialogResult.OK == NPF.ShowDialog())
								{	
									string sSQLAutEs = string.Format("Update solicitudvacaciones set Estado = 2  where idsolicitudvacaciones = {0}",
										(int) e.Cell.Row.Cells["idsolicitudvacaciones"].Value);
									oCmdActualiza.CommandText = sSQLAutEs;
									oCmdActualiza.ExecuteNonQuery();		
									this.btnBuscar_Click (sender, e );
								}								
							}
														
							#endregion Permisos

							#endregion Asignacion Vendedores 
							
						}
						else
						{
							#region Abre observacion
							bEdicion=true;
							using (frmObservacionesVacaciones Observaciones = new frmObservacionesVacaciones(1, e.Cell.Row.Cells["Motivo"].Value.ToString().Trim(), bEdicion))
							{				
								if (DialogResult.OK == Observaciones.ShowDialog())
								{
									e.Cell.Row.Cells["Motivo"].Value = Observaciones.txtObservaciones.Text.ToString().Trim();
								}
							}
							#endregion Abre observacion

							#region Valida Motivo

							if (e.Cell.Row.Cells["Motivo"].Value.ToString().Trim().Length == 0 )
							{
								MessageBox.Show("Ingrese un Motivo", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
								this.ultraGrid1.ActiveRow.Selected = true;
								return;
							}
							if (e.Cell.Row.Cells["Motivo"].Value.ToString().Trim().Length < 10 )
							{
								MessageBox.Show("La Motivo debe tener minimo 10 caracteres ", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
								this.ultraGrid1.ActiveRow.Selected = true;
								return;
							}

							if (e.Cell.Row.Cells["Motivo"].Value.ToString().Trim().Length > 100 )
							{
								MessageBox.Show("La Motivo debe tener maximo 100 caracteres ", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
								this.ultraGrid1.ActiveRow.Selected = true;
								return;
							}

							#endregion Valida Motivo					

							#region Administrativos
							oTransaction = oConexion.BeginTransaction();
							oCmdActualiza.Transaction = oTransaction;
							string sSQLAutEs = string.Format("Update solicitudvacaciones set Estado = 2  where idsolicitudvacaciones = {0}",
								(int) e.Cell.Row.Cells["idsolicitudvacaciones"].Value);
							oCmdActualiza.CommandText = sSQLAutEs;
							oCmdActualiza.ExecuteNonQuery();		

							string sSQLAut = string.Format("Exec AutorizacionVacaciones {0},{1}, {2}, '{3}', '{4}', '{5}','{6}'", 
								(int)e.Cell.Row.Cells["idPersonal"].Value,3,0,Convert.ToDateTime(e.Cell.Row.Cells["FechaSalida"].Value).ToString("yyyyMMdd"),Convert.ToDateTime(e.Cell.Row.Cells["FechaRetorna"].Value).ToString("yyyyMMdd"),e.Cell.Row.Cells["Motivo"].Value,e.Cell.Row.Cells["Observacion"].Value);			
							oCmdActualiza.CommandText = sSQLAut;
							oCmdActualiza.ExecuteNonQuery();	

							oTransaction.Commit();
							Cursor = Cursors.Default;	

							this.btnBuscar_Click (sender, e);
							#endregion Administrativos
						}
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
			#endregion Aprobacion

			#region impresiones

			#region impresion 
								
			if (e.Cell.Column.ToString() == "Imprimir")
			{	
				if (!miAcceso.Imprimir) 
				{
					MessageBox.Show("No tiene acceso IMPRIMIR ", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					return;
				}
				else 
				{


					ParameterFields paramFields = new ParameterFields();
					ParameterField paramFieldCedula = new ParameterField ();
					ParameterDiscreteValue ValCedula = new ParameterDiscreteValue ();
					paramFieldCedula.ParameterFieldName = "@idsolicitudVacaciones";
					ValCedula.Value = (int) e.Cell.Row.Cells["idsolicitudVacaciones"].Value;
					paramFieldCedula.CurrentValues.Add (ValCedula);
					paramFields.Add (paramFieldCedula);
					Reporte miRep = new Reporte("SolicitudVacaciones.rpt", "");
					miRep.MdiParent = this.MdiParent;
					miRep.Titulo("Solicitud Vacaciones");
					miRep.crVista.ParameterFieldInfo = paramFields;
					miRep.Show();
				}
			}
			#endregion impresion 

			#endregion impresiones


		}

		private void Seguridad ()
		{
			if ( (int) this.optEstados.Value == 0)
			{
				this.ultraGrid1.DisplayLayout.Bands[0].Columns["Aceptar"].Hidden = true;
				this.ultraGrid1.DisplayLayout.Bands[0].Columns["Supervisor"].Hidden = true;
				this.ultraGrid1.DisplayLayout.Bands[0].Columns["Anular"].Hidden = true;
				this.ultraGrid1.DisplayLayout.Bands[0].Columns["Imprimir"].Hidden = true;
			}
			if ( (int) this.optEstados.Value == 1)
			{
				if (miAcceso.BGenerarCobro) this.ultraGrid1.DisplayLayout.Bands[0].Columns["Aceptar"].Hidden = true;
				if (miAcceso.BAnularAprobado)this.ultraGrid1.DisplayLayout.Bands[0].Columns["Anular"].Hidden = false;
				if (miAcceso.Autorizacion)this.ultraGrid1.DisplayLayout.Bands[0].Columns["Supervisor"].Hidden = false;
				this.ultraGrid1.DisplayLayout.Bands[0].Columns["Imprimir"].Hidden = true;
			}
			if ( (int) this.optEstados.Value == 2)
			{
				if (miAcceso.BGenerarCobro) this.ultraGrid1.DisplayLayout.Bands[0].Columns["Aceptar"].Hidden = true;
				if (miAcceso.BAnularAprobado)this.ultraGrid1.DisplayLayout.Bands[0].Columns["Anular"].Hidden = false;
				this.ultraGrid1.DisplayLayout.Bands[0].Columns["Imprimir"].Hidden = false;
				if (miAcceso.Imprimir)this.ultraGrid1.DisplayLayout.Bands[0].Columns["Supervisor"].Hidden = true;

			}
			if ( (int) this.optEstados.Value == 3)
			{
				this.ultraGrid1.DisplayLayout.Bands[0].Columns["Aceptar"].Hidden = true;
				this.ultraGrid1.DisplayLayout.Bands[0].Columns["Anular"].Hidden = true;
				this.ultraGrid1.DisplayLayout.Bands[0].Columns["Imprimir"].Hidden = true;
				this.ultraGrid1.DisplayLayout.Bands[0].Columns["Supervisor"].Hidden = true;
			}
			if ( (int) this.optEstados.Value == 4)
			{
				if (miAcceso.BGenerarCobro) this.ultraGrid1.DisplayLayout.Bands[0].Columns["Aceptar"].Hidden = false;
				if (miAcceso.BAnularAprobado)this.ultraGrid1.DisplayLayout.Bands[0].Columns["Anular"].Hidden = false;
				if (miAcceso.Autorizacion)this.ultraGrid1.DisplayLayout.Bands[0].Columns["Supervisor"].Hidden = true;
				this.ultraGrid1.DisplayLayout.Bands[0].Columns["Imprimir"].Hidden = true;
			}
		}

		private Acceso miAcceso;
		private void UnloadMe()
		{
			this.Close();
		}

		private void btnExcel_Click(object sender, System.EventArgs e)
		{
			FuncionesProcedimientos.ExportaExcel(this.ultraGrid1);
		}

		private void optEstados_ValueChanged_1(object sender, System.EventArgs e)
		{
			this.btnBuscar_Click (sender, e );
		}

		private void ultraGrid1_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}
	}
}
