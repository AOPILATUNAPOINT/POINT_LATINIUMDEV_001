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
	/// Descripción breve de frmConusltaGeneralAumentoPresupuesto.
	/// </summary>
	public class frmConusltaGeneralAumentoPresupuesto : System.Windows.Forms.Form
	{
		private System.Windows.Forms.Label label8;
		private System.Windows.Forms.Label label6;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtHasta;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtDesde;
		private System.Windows.Forms.Label lblComprobante;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbProyecto;
		private System.Windows.Forms.Button btnAceptar;
		private System.Windows.Forms.Button btnCancelar;
		private Infragistics.Win.UltraWinEditors.UltraOptionSet optEstado;
		private System.Windows.Forms.Label lblContador;
		private System.Windows.Forms.Button btnBuscar;
		public Infragistics.Win.UltraWinGrid.UltraGrid uGridCreditos;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource1;
		private C1.Data.C1DataSet cdsSeteoF;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;
    bool bVistas = false;
		public frmConusltaGeneralAumentoPresupuesto(bool BVistas)
		{
			//
			// Necesario para admitir el Diseñador de Windows Forms
			//
		 bVistas =  BVistas;
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
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton1 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmConusltaGeneralAumentoPresupuesto));
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton2 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bodega");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.ValueListItem valueListItem1 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem2 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem3 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem4 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Fecha");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("periodo");
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Mes");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Descripcion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Estado");
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Valor_Solicitado");
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn11 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Usuario");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn12 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idAumentoPresupuesto");
			Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance11 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance12 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance13 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance14 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance15 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Fecha");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn3 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("periodo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn4 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Mes");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn5 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Codigo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn6 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Descripcion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn7 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Estado");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn8 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Valor_Solicitado");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn9 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Usuario");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn10 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idAumentoPresupuesto");
			this.label8 = new System.Windows.Forms.Label();
			this.label6 = new System.Windows.Forms.Label();
			this.dtHasta = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.dtDesde = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.lblComprobante = new System.Windows.Forms.Label();
			this.cmbProyecto = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.btnAceptar = new System.Windows.Forms.Button();
			this.btnCancelar = new System.Windows.Forms.Button();
			this.optEstado = new Infragistics.Win.UltraWinEditors.UltraOptionSet();
			this.lblContador = new System.Windows.Forms.Label();
			this.btnBuscar = new System.Windows.Forms.Button();
			this.uGridCreditos = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource1 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			((System.ComponentModel.ISupportInitialize)(this.dtHasta)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtDesde)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbProyecto)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.optEstado)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridCreditos)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			this.SuspendLayout();
			// 
			// label8
			// 
			this.label8.AutoSize = true;
			this.label8.Location = new System.Drawing.Point(200, 10);
			this.label8.Name = "label8";
			this.label8.Size = new System.Drawing.Size(34, 16);
			this.label8.TabIndex = 268;
			this.label8.Text = "Hasta";
			// 
			// label6
			// 
			this.label6.AutoSize = true;
			this.label6.Location = new System.Drawing.Point(16, 10);
			this.label6.Name = "label6";
			this.label6.Size = new System.Drawing.Size(37, 16);
			this.label6.TabIndex = 267;
			this.label6.Text = "Desde";
			// 
			// dtHasta
			// 
			appearance1.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtHasta.Appearance = appearance1;
			this.dtHasta.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton1.Caption = "Today";
			this.dtHasta.DateButtons.Add(dateButton1);
			this.dtHasta.Format = "dd/MM/yyyy";
			this.dtHasta.Location = new System.Drawing.Point(248, 8);
			this.dtHasta.Name = "dtHasta";
			this.dtHasta.NonAutoSizeHeight = 23;
			this.dtHasta.Size = new System.Drawing.Size(104, 21);
			this.dtHasta.SpinButtonsVisible = true;
			this.dtHasta.TabIndex = 266;
			this.dtHasta.Value = ((object)(resources.GetObject("dtHasta.Value")));
			this.dtHasta.ValueChanged += new System.EventHandler(this.dtHasta_ValueChanged);
			// 
			// dtDesde
			// 
			appearance2.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtDesde.Appearance = appearance2;
			this.dtDesde.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton2.Caption = "Today";
			this.dtDesde.DateButtons.Add(dateButton2);
			this.dtDesde.Format = "dd/MM/yyyy";
			this.dtDesde.Location = new System.Drawing.Point(72, 8);
			this.dtDesde.Name = "dtDesde";
			this.dtDesde.NonAutoSizeHeight = 23;
			this.dtDesde.Size = new System.Drawing.Size(104, 21);
			this.dtDesde.SpinButtonsVisible = true;
			this.dtDesde.TabIndex = 265;
			this.dtDesde.Value = ((object)(resources.GetObject("dtDesde.Value")));
			// 
			// lblComprobante
			// 
			this.lblComprobante.AutoSize = true;
			this.lblComprobante.BackColor = System.Drawing.Color.Transparent;
			this.lblComprobante.ImeMode = System.Windows.Forms.ImeMode.NoControl;
			this.lblComprobante.Location = new System.Drawing.Point(16, 89);
			this.lblComprobante.Name = "lblComprobante";
			this.lblComprobante.Size = new System.Drawing.Size(49, 16);
			this.lblComprobante.TabIndex = 1008;
			this.lblComprobante.Text = "Proyecto";
			// 
			// cmbProyecto
			// 
			appearance3.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbProyecto.Appearance = appearance3;
			this.cmbProyecto.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbProyecto.Cursor = System.Windows.Forms.Cursors.Default;
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridColumn2.Width = 230;
			ultraGridBand1.Columns.AddRange(new object[] {
																										 ultraGridColumn1,
																										 ultraGridColumn2});
			ultraGridBand1.SummaryFooterCaption = "Grand Summaries";
			this.cmbProyecto.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.cmbProyecto.DisplayMember = "Nombre";
			this.cmbProyecto.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbProyecto.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbProyecto.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F);
			this.cmbProyecto.Location = new System.Drawing.Point(72, 87);
			this.cmbProyecto.Name = "cmbProyecto";
			this.cmbProyecto.Size = new System.Drawing.Size(232, 21);
			this.cmbProyecto.TabIndex = 1009;
			this.cmbProyecto.ValueMember = "Bodega";
			// 
			// btnAceptar
			// 
			this.btnAceptar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnAceptar.CausesValidation = false;
			this.btnAceptar.Location = new System.Drawing.Point(8, 336);
			this.btnAceptar.Name = "btnAceptar";
			this.btnAceptar.TabIndex = 1010;
			this.btnAceptar.Text = "&Aceptar";
			this.btnAceptar.Click += new System.EventHandler(this.btnAceptar_Click);
			// 
			// btnCancelar
			// 
			this.btnCancelar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnCancelar.CausesValidation = false;
			this.btnCancelar.Location = new System.Drawing.Point(96, 336);
			this.btnCancelar.Name = "btnCancelar";
			this.btnCancelar.TabIndex = 1011;
			this.btnCancelar.Text = "&Cancelar";
			this.btnCancelar.Click += new System.EventHandler(this.btnCancelar_Click);
			// 
			// optEstado
			// 
			appearance4.ForeColorDisabled = System.Drawing.Color.Black;
			this.optEstado.Appearance = appearance4;
			this.optEstado.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.optEstado.CheckedIndex = 0;
			this.optEstado.ItemAppearance = appearance5;
			this.optEstado.ItemOrigin = new System.Drawing.Point(10, 0);
			valueListItem1.DataValue = 0;
			valueListItem1.DisplayText = "Todos";
			valueListItem2.DataValue = 1;
			valueListItem2.DisplayText = "Pendiente";
			valueListItem3.DataValue = 2;
			valueListItem3.DisplayText = "Aprobado";
			valueListItem4.DataValue = 3;
			valueListItem4.DisplayText = "Rechazado";
			this.optEstado.Items.Add(valueListItem1);
			this.optEstado.Items.Add(valueListItem2);
			this.optEstado.Items.Add(valueListItem3);
			this.optEstado.Items.Add(valueListItem4);
			this.optEstado.ItemSpacingVertical = 10;
			this.optEstado.Location = new System.Drawing.Point(72, 40);
			this.optEstado.Name = "optEstado";
			this.optEstado.Size = new System.Drawing.Size(320, 32);
			this.optEstado.TabIndex = 1012;
			this.optEstado.Text = "Todos";
			// 
			// lblContador
			// 
			this.lblContador.AutoSize = true;
			this.lblContador.Font = new System.Drawing.Font("Tahoma", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblContador.ForeColor = System.Drawing.Color.Firebrick;
			this.lblContador.Location = new System.Drawing.Point(400, 88);
			this.lblContador.Name = "lblContador";
			this.lblContador.Size = new System.Drawing.Size(0, 19);
			this.lblContador.TabIndex = 1013;
			this.lblContador.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// btnBuscar
			// 
			this.btnBuscar.CausesValidation = false;
			this.btnBuscar.Location = new System.Drawing.Point(312, 86);
			this.btnBuscar.Name = "btnBuscar";
			this.btnBuscar.TabIndex = 1014;
			this.btnBuscar.Text = "&Buscar";
			this.btnBuscar.Click += new System.EventHandler(this.btnBuscar_Click);
			// 
			// uGridCreditos
			// 
			this.uGridCreditos.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.uGridCreditos.Cursor = System.Windows.Forms.Cursors.Default;
			this.uGridCreditos.DataSource = this.ultraDataSource1;
			appearance6.BackColor = System.Drawing.Color.White;
			this.uGridCreditos.DisplayLayout.Appearance = appearance6;
			ultraGridBand2.AddButtonCaption = "Serie";
			ultraGridColumn3.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn3.Header.VisiblePosition = 0;
			ultraGridColumn3.Width = 85;
			ultraGridColumn4.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn4.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn4.Header.Caption = "Proyecto";
			ultraGridColumn4.Header.VisiblePosition = 1;
			ultraGridColumn5.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance7.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn5.CellAppearance = appearance7;
			ultraGridColumn5.Header.Caption = "Periodo";
			ultraGridColumn5.Header.VisiblePosition = 2;
			ultraGridColumn6.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn6.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn6.Header.VisiblePosition = 3;
			ultraGridColumn7.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn7.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn7.Header.VisiblePosition = 4;
			ultraGridColumn8.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn8.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn8.Header.VisiblePosition = 5;
			ultraGridColumn8.Width = 150;
			ultraGridColumn9.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn9.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance8.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn9.CellAppearance = appearance8;
			ultraGridColumn9.Header.VisiblePosition = 6;
			ultraGridColumn9.Width = 150;
			ultraGridColumn10.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance9.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn10.CellAppearance = appearance9;
			ultraGridColumn10.Format = "#,##0.00";
			ultraGridColumn10.Header.Caption = "Solicitado";
			ultraGridColumn10.Header.VisiblePosition = 7;
			ultraGridColumn11.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn11.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn11.Header.VisiblePosition = 8;
			ultraGridColumn12.Header.VisiblePosition = 9;
			ultraGridColumn12.Hidden = true;
			ultraGridBand2.Columns.AddRange(new object[] {
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
			ultraGridBand2.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			ultraGridBand2.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			ultraGridBand2.SummaryFooterCaption = "";
			this.uGridCreditos.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			appearance10.ForeColor = System.Drawing.Color.Black;
			appearance10.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridCreditos.DisplayLayout.Override.ActiveRowAppearance = appearance10;
			this.uGridCreditos.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.uGridCreditos.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance11.BackColor = System.Drawing.Color.Transparent;
			this.uGridCreditos.DisplayLayout.Override.CardAreaAppearance = appearance11;
			appearance12.ForeColor = System.Drawing.Color.Black;
			appearance12.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridCreditos.DisplayLayout.Override.CellAppearance = appearance12;
			this.uGridCreditos.DisplayLayout.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.Edit;
			appearance13.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance13.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance13.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance13.FontData.BoldAsString = "True";
			appearance13.FontData.Name = "Arial";
			appearance13.FontData.SizeInPoints = 8.5F;
			appearance13.ForeColor = System.Drawing.Color.White;
			appearance13.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridCreditos.DisplayLayout.Override.HeaderAppearance = appearance13;
			this.uGridCreditos.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti;
			appearance14.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance14.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance14.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridCreditos.DisplayLayout.Override.RowSelectorAppearance = appearance14;
			appearance15.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance15.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance15.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridCreditos.DisplayLayout.Override.SelectedRowAppearance = appearance15;
			this.uGridCreditos.DisplayLayout.Override.SelectTypeRow = Infragistics.Win.UltraWinGrid.SelectType.Extended;
			this.uGridCreditos.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridCreditos.Location = new System.Drawing.Point(8, 120);
			this.uGridCreditos.Name = "uGridCreditos";
			this.uGridCreditos.Size = new System.Drawing.Size(920, 208);
			this.uGridCreditos.TabIndex = 1015;
			// 
			// ultraDataSource1
			// 
			ultraDataColumn1.DataType = typeof(System.DateTime);
			ultraDataColumn3.DataType = typeof(int);
			ultraDataColumn8.DataType = typeof(System.Decimal);
			ultraDataColumn10.DataType = typeof(int);
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
																																 ultraDataColumn10});
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
			// frmConusltaGeneralAumentoPresupuesto
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(936, 366);
			this.ControlBox = false;
			this.Controls.Add(this.uGridCreditos);
			this.Controls.Add(this.btnBuscar);
			this.Controls.Add(this.lblContador);
			this.Controls.Add(this.optEstado);
			this.Controls.Add(this.btnAceptar);
			this.Controls.Add(this.btnCancelar);
			this.Controls.Add(this.lblComprobante);
			this.Controls.Add(this.cmbProyecto);
			this.Controls.Add(this.label8);
			this.Controls.Add(this.label6);
			this.Controls.Add(this.dtHasta);
			this.Controls.Add(this.dtDesde);
			this.KeyPreview = true;
			this.Name = "frmConusltaGeneralAumentoPresupuesto";
			this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
			this.Text = "Buscar";
			this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.frmConusltaGeneralAumentoPresupuesto_KeyDown);
			this.Load += new System.EventHandler(this.frmConusltaGeneralAumentoPresupuesto_Load);
			((System.ComponentModel.ISupportInitialize)(this.dtHasta)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtDesde)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbProyecto)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.optEstado)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridCreditos)).EndInit();
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

		private void frmConusltaGeneralAumentoPresupuesto_Load(object sender, System.EventArgs e)
		{
			this.dtDesde.CalendarInfo.MinDate = DateTime.Parse("20/06/2019");
			this.dtHasta.CalendarInfo.MinDate = DateTime.Parse("20/06/2019");			

			this.dtDesde.Value = DateTime.Today;
			this.dtHasta.Value = DateTime.Today;
		}

		private void ListaBodegas (Infragistics.Win.UltraWinGrid.UltraCombo uCombo, int idUsuario,  DateTime dtFecha,  C1.Data.C1DataSet ds)
		{
			if (bVistas)
			{
				uCombo.DataSource = Funcion.dvProcedimiento(ds, string.Format("Exec [VerCentroDecostoSolicitud] {0}", 1));
			}
			else
			{
				uCombo.DataSource = Funcion.dvProcedimiento(ds, string.Format("Exec PresupuestoListaCC {0}, '{1}'", idUsuario,  dtFecha.ToString("yyyyMMdd")));
			}
		}

		private void dtHasta_ValueChanged(object sender, System.EventArgs e)
		{
			ListaBodegas(this.cmbProyecto, MenuLatinium.IdUsuario,  (DateTime)this.dtHasta.Value,  cdsSeteoF);
		}

		private void btnBuscar_Click(object sender, System.EventArgs e)
		{
			#region Variables
			if (!Validacion.vbFechaEnDocumentos(this.dtDesde, "Seleccione la fecha inicial", false, "", cdsSeteoF)) return;
			if (!Validacion.vbFechaEnDocumentos(this.dtHasta, "Seleccione la fecha final", false, "", cdsSeteoF)) return;
			if (!Validacion.vbComparaFechas(Convert.ToDateTime(this.dtDesde.Value), Convert.ToDateTime(this.dtHasta.Value), this.dtDesde, "La fecha inicial no puede ser posterior a la fecha final", false)) return;
			if (!bVistas) if (!Validacion.vbComboVacio(this.cmbProyecto, "Seleccione el proyecto")) return;


			#endregion Variables

			int Idsector = 0;
			if ( this.cmbProyecto.ActiveRow != null) Idsector  = ( int) this.cmbProyecto.Value;

			#region Consulta
			this.lblContador.Text = "";
			this.uGridCreditos.DataSource = ultraDataSource1;

			string sSQL = string.Format("Exec [ConusltaGeneralAumentoPresupuesto] '{0}', '{1}', {2}, {3}", 
				Convert.ToDateTime(this.dtDesde.Value).ToString("yyyyMMdd"), Convert.ToDateTime(this.dtHasta.Value).ToString("yyyyMMdd"), 
				Idsector,
				(int) this.optEstado.Value);
			
			
			this.uGridCreditos.DataSource = Funcion.dvProcedimiento(cdsSeteoF, sSQL);

			if (this.uGridCreditos.Rows.Count > 0) this.uGridCreditos.Focus();

			this.lblContador.Text = this.uGridCreditos.Rows.Count + " REGISTROS ENCONTRADOS";
			#endregion Consulta

			
		}

		private void btnCancelar_Click(object sender, System.EventArgs e)
		{
			this.DialogResult = DialogResult.No;
			
		}

		private void frmConusltaGeneralAumentoPresupuesto_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Escape) this.DialogResult = DialogResult.No;
		}

		private void btnAceptar_Click(object sender, System.EventArgs e)
		{
			Aceptar();
		}
		private void Aceptar()
		{
			if (this.uGridCreditos.Rows.Count == 0)
			{
				MessageBox.Show("No hay documentos encontrados.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
			//	this.txtBusqueda.Focus();			
			}
			else if (this.uGridCreditos.ActiveRow == null)
			{
				MessageBox.Show("Seleccione un documento.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);				
			}
			else if (this.uGridCreditos.ActiveRow.Cells["idAumentoPresupuesto"].Value == System.DBNull.Value)
			{
				MessageBox.Show("Seleccione un documento.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);				
			}
			else 
			{
				this.DialogResult = DialogResult.OK;
			}
		}


	}
}
