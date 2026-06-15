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
	/// Descripción breve de frmAprobarPresupuestoPedidoCXC.
	/// </summary>
	public class frmAprobarPresupuestoPedidoCXC : System.Windows.Forms.Form
	{
		private System.Windows.Forms.Label label3;
		public Infragistics.Win.UltraWinGrid.UltraCombo cmbEstado;
		private System.Windows.Forms.Label label1;
		public Infragistics.Win.UltraWinGrid.UltraCombo cmbProyecto;
		public Infragistics.Win.UltraWinGrid.UltraGrid uGridCreditos;
		private System.Windows.Forms.Label lblContador;
		private System.Windows.Forms.GroupBox groupBox1;
		private Infragistics.Win.Misc.UltraLabel ultraLabel1;
		private System.Windows.Forms.Button btnBuscar;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtBusqueda;
		private System.Windows.Forms.Label label8;
		private System.Windows.Forms.Label label6;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtHasta;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtDesde;
		private C1.Data.C1DataSet cdsSeteoF;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource3;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public frmAprobarPresupuestoPedidoCXC()
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
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Estado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombres");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bodega");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.Appearance appearance1 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand3 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCompra");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Periodo");
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Fecha");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombres");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Solicitante");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Proyecto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn11 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Numero");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn12 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("btnAprobar", 0);
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmAprobarPresupuestoPedidoCXC));
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn13 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("btnRechazar", 1);
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance11 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance12 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton1 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.Appearance appearance13 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton2 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCompra");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Periodo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn3 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Fecha");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn4 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombres");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn5 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Solicitante");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn6 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Proyecto");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn7 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Numero");
			this.label3 = new System.Windows.Forms.Label();
			this.cmbEstado = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.label1 = new System.Windows.Forms.Label();
			this.cmbProyecto = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.uGridCreditos = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.lblContador = new System.Windows.Forms.Label();
			this.groupBox1 = new System.Windows.Forms.GroupBox();
			this.ultraLabel1 = new Infragistics.Win.Misc.UltraLabel();
			this.btnBuscar = new System.Windows.Forms.Button();
			this.txtBusqueda = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label8 = new System.Windows.Forms.Label();
			this.label6 = new System.Windows.Forms.Label();
			this.dtHasta = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.dtDesde = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.ultraDataSource3 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			((System.ComponentModel.ISupportInitialize)(this.cmbEstado)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbProyecto)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridCreditos)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtBusqueda)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtHasta)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtDesde)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource3)).BeginInit();
			this.SuspendLayout();
			// 
			// label3
			// 
			this.label3.AutoSize = true;
			this.label3.Location = new System.Drawing.Point(8, 40);
			this.label3.Name = "label3";
			this.label3.Size = new System.Drawing.Size(39, 16);
			this.label3.TabIndex = 1008;
			this.label3.Text = "Estado";
			// 
			// cmbEstado
			// 
			this.cmbEstado.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbEstado.Cursor = System.Windows.Forms.Cursors.Default;
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridColumn2.Width = 260;
			ultraGridBand1.Columns.AddRange(new object[] {
																										 ultraGridColumn1,
																										 ultraGridColumn2});
			this.cmbEstado.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.cmbEstado.DisplayMember = "Nombres";
			this.cmbEstado.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbEstado.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbEstado.Location = new System.Drawing.Point(80, 40);
			this.cmbEstado.MaxDropDownItems = 30;
			this.cmbEstado.Name = "cmbEstado";
			this.cmbEstado.Size = new System.Drawing.Size(264, 21);
			this.cmbEstado.TabIndex = 1007;
			this.cmbEstado.ValueMember = "Estado";
			// 
			// label1
			// 
			this.label1.AutoSize = true;
			this.label1.Location = new System.Drawing.Point(8, 72);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(49, 16);
			this.label1.TabIndex = 1006;
			this.label1.Text = "Proyecto";
			// 
			// cmbProyecto
			// 
			this.cmbProyecto.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbProyecto.Cursor = System.Windows.Forms.Cursors.Default;
			ultraGridColumn3.Header.VisiblePosition = 0;
			ultraGridColumn3.Hidden = true;
			ultraGridColumn4.Header.VisiblePosition = 1;
			ultraGridColumn4.Width = 270;
			ultraGridBand2.Columns.AddRange(new object[] {
																										 ultraGridColumn3,
																										 ultraGridColumn4});
			this.cmbProyecto.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			this.cmbProyecto.DisplayMember = "Nombre";
			this.cmbProyecto.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbProyecto.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbProyecto.Location = new System.Drawing.Point(64, 72);
			this.cmbProyecto.MaxDropDownItems = 30;
			this.cmbProyecto.Name = "cmbProyecto";
			this.cmbProyecto.Size = new System.Drawing.Size(280, 21);
			this.cmbProyecto.TabIndex = 1005;
			this.cmbProyecto.ValueMember = "Bodega";
			// 
			// uGridCreditos
			// 
			this.uGridCreditos.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.uGridCreditos.Cursor = System.Windows.Forms.Cursors.Default;
			this.uGridCreditos.DataSource = this.ultraDataSource3;
			appearance1.BackColor = System.Drawing.Color.White;
			this.uGridCreditos.DisplayLayout.Appearance = appearance1;
			ultraGridBand3.AddButtonCaption = "Serie";
			ultraGridColumn5.Header.VisiblePosition = 0;
			ultraGridColumn5.Hidden = true;
			ultraGridColumn6.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance2.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn6.CellAppearance = appearance2;
			ultraGridColumn6.Header.VisiblePosition = 3;
			ultraGridColumn7.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn7.Header.VisiblePosition = 4;
			ultraGridColumn7.Width = 85;
			ultraGridColumn8.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn8.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn8.Header.Caption = "Estado";
			ultraGridColumn8.Header.VisiblePosition = 8;
			ultraGridColumn8.Width = 120;
			ultraGridColumn9.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn9.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn9.Header.VisiblePosition = 6;
			ultraGridColumn9.Width = 250;
			ultraGridColumn10.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn10.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn10.Header.VisiblePosition = 5;
			ultraGridColumn10.Width = 150;
			ultraGridColumn11.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn11.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn11.Header.VisiblePosition = 7;
			ultraGridColumn11.Width = 110;
			ultraGridColumn12.ButtonDisplayStyle = Infragistics.Win.UltraWinGrid.ButtonDisplayStyle.Always;
			appearance3.Image = ((object)(resources.GetObject("appearance3.Image")));
			appearance3.ImageHAlign = Infragistics.Win.HAlign.Center;
			appearance3.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn12.CellButtonAppearance = appearance3;
			ultraGridColumn12.Header.Caption = "Aprobar";
			ultraGridColumn12.Header.VisiblePosition = 1;
			ultraGridColumn12.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.Button;
			ultraGridColumn12.Width = 40;
			ultraGridColumn13.ButtonDisplayStyle = Infragistics.Win.UltraWinGrid.ButtonDisplayStyle.Always;
			appearance4.Image = ((object)(resources.GetObject("appearance4.Image")));
			appearance4.ImageHAlign = Infragistics.Win.HAlign.Center;
			appearance4.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn13.CellButtonAppearance = appearance4;
			ultraGridColumn13.Header.Caption = "Rechazar";
			ultraGridColumn13.Header.VisiblePosition = 2;
			ultraGridColumn13.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.Button;
			ultraGridColumn13.Width = 65;
			ultraGridBand3.Columns.AddRange(new object[] {
																										 ultraGridColumn5,
																										 ultraGridColumn6,
																										 ultraGridColumn7,
																										 ultraGridColumn8,
																										 ultraGridColumn9,
																										 ultraGridColumn10,
																										 ultraGridColumn11,
																										 ultraGridColumn12,
																										 ultraGridColumn13});
			ultraGridBand3.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			ultraGridBand3.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			ultraGridBand3.SummaryFooterCaption = "";
			this.uGridCreditos.DisplayLayout.BandsSerializer.Add(ultraGridBand3);
			appearance5.ForeColor = System.Drawing.Color.Black;
			appearance5.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridCreditos.DisplayLayout.Override.ActiveRowAppearance = appearance5;
			this.uGridCreditos.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.uGridCreditos.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance6.BackColor = System.Drawing.Color.Transparent;
			this.uGridCreditos.DisplayLayout.Override.CardAreaAppearance = appearance6;
			appearance7.ForeColor = System.Drawing.Color.Black;
			appearance7.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridCreditos.DisplayLayout.Override.CellAppearance = appearance7;
			this.uGridCreditos.DisplayLayout.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.Edit;
			appearance8.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance8.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance8.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance8.FontData.BoldAsString = "True";
			appearance8.FontData.Name = "Arial";
			appearance8.FontData.SizeInPoints = 8.5F;
			appearance8.ForeColor = System.Drawing.Color.White;
			appearance8.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridCreditos.DisplayLayout.Override.HeaderAppearance = appearance8;
			this.uGridCreditos.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti;
			appearance9.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance9.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance9.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridCreditos.DisplayLayout.Override.RowSelectorAppearance = appearance9;
			appearance10.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance10.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance10.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridCreditos.DisplayLayout.Override.SelectedRowAppearance = appearance10;
			this.uGridCreditos.DisplayLayout.Override.SelectTypeRow = Infragistics.Win.UltraWinGrid.SelectType.Extended;
			this.uGridCreditos.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridCreditos.Location = new System.Drawing.Point(8, 160);
			this.uGridCreditos.Name = "uGridCreditos";
			this.uGridCreditos.Size = new System.Drawing.Size(912, 224);
			this.uGridCreditos.TabIndex = 1004;
			this.uGridCreditos.ClickCellButton += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.uGridCreditos_ClickCellButton);
			// 
			// lblContador
			// 
			this.lblContador.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.lblContador.AutoSize = true;
			this.lblContador.Font = new System.Drawing.Font("Tahoma", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblContador.ForeColor = System.Drawing.Color.Firebrick;
			this.lblContador.Location = new System.Drawing.Point(360, 64);
			this.lblContador.Name = "lblContador";
			this.lblContador.Size = new System.Drawing.Size(0, 19);
			this.lblContador.TabIndex = 1003;
			this.lblContador.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// groupBox1
			// 
			this.groupBox1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.groupBox1.Location = new System.Drawing.Point(-40, 136);
			this.groupBox1.Name = "groupBox1";
			this.groupBox1.Size = new System.Drawing.Size(1128, 8);
			this.groupBox1.TabIndex = 1002;
			this.groupBox1.TabStop = false;
			// 
			// ultraLabel1
			// 
			this.ultraLabel1.AutoSize = true;
			this.ultraLabel1.BackColor = System.Drawing.Color.Transparent;
			this.ultraLabel1.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.ultraLabel1.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(50)), ((System.Byte)(50)), ((System.Byte)(50)));
			this.ultraLabel1.Location = new System.Drawing.Point(8, 104);
			this.ultraLabel1.Name = "ultraLabel1";
			this.ultraLabel1.Size = new System.Drawing.Size(44, 15);
			this.ultraLabel1.TabIndex = 999;
			this.ultraLabel1.Text = "Número";
			// 
			// btnBuscar
			// 
			this.btnBuscar.CausesValidation = false;
			this.btnBuscar.Location = new System.Drawing.Point(272, 104);
			this.btnBuscar.Name = "btnBuscar";
			this.btnBuscar.TabIndex = 998;
			this.btnBuscar.Text = "&Buscar";
			this.btnBuscar.Click += new System.EventHandler(this.btnBuscar_Click);
			// 
			// txtBusqueda
			// 
			appearance11.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtBusqueda.Appearance = appearance11;
			this.txtBusqueda.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtBusqueda.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtBusqueda.Location = new System.Drawing.Point(64, 104);
			this.txtBusqueda.Name = "txtBusqueda";
			this.txtBusqueda.Size = new System.Drawing.Size(192, 21);
			this.txtBusqueda.TabIndex = 997;
			// 
			// label8
			// 
			this.label8.AutoSize = true;
			this.label8.Location = new System.Drawing.Point(192, 8);
			this.label8.Name = "label8";
			this.label8.Size = new System.Drawing.Size(34, 16);
			this.label8.TabIndex = 996;
			this.label8.Text = "Hasta";
			// 
			// label6
			// 
			this.label6.AutoSize = true;
			this.label6.Location = new System.Drawing.Point(8, 8);
			this.label6.Name = "label6";
			this.label6.Size = new System.Drawing.Size(37, 16);
			this.label6.TabIndex = 995;
			this.label6.Text = "Desde";
			// 
			// dtHasta
			// 
			appearance12.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtHasta.Appearance = appearance12;
			this.dtHasta.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton1.Caption = "Today";
			this.dtHasta.DateButtons.Add(dateButton1);
			this.dtHasta.Format = "dd/MM/yyyy";
			this.dtHasta.Location = new System.Drawing.Point(240, 8);
			this.dtHasta.Name = "dtHasta";
			this.dtHasta.NonAutoSizeHeight = 23;
			this.dtHasta.Size = new System.Drawing.Size(104, 21);
			this.dtHasta.SpinButtonsVisible = true;
			this.dtHasta.TabIndex = 994;
			this.dtHasta.Value = ((object)(resources.GetObject("dtHasta.Value")));
			// 
			// dtDesde
			// 
			appearance13.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtDesde.Appearance = appearance13;
			this.dtDesde.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton2.Caption = "Today";
			this.dtDesde.DateButtons.Add(dateButton2);
			this.dtDesde.Format = "dd/MM/yyyy";
			this.dtDesde.Location = new System.Drawing.Point(64, 8);
			this.dtDesde.Name = "dtDesde";
			this.dtDesde.NonAutoSizeHeight = 23;
			this.dtDesde.Size = new System.Drawing.Size(104, 21);
			this.dtDesde.SpinButtonsVisible = true;
			this.dtDesde.TabIndex = 993;
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
			// ultraDataSource3
			// 
			ultraDataColumn1.DataType = typeof(int);
			ultraDataColumn2.DataType = typeof(int);
			ultraDataColumn3.DataType = typeof(System.DateTime);
			this.ultraDataSource3.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn1,
																																 ultraDataColumn2,
																																 ultraDataColumn3,
																																 ultraDataColumn4,
																																 ultraDataColumn5,
																																 ultraDataColumn6,
																																 ultraDataColumn7});
			// 
			// frmAprobarPresupuestoPedidoCXC
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(936, 390);
			this.Controls.Add(this.label3);
			this.Controls.Add(this.cmbEstado);
			this.Controls.Add(this.label1);
			this.Controls.Add(this.cmbProyecto);
			this.Controls.Add(this.uGridCreditos);
			this.Controls.Add(this.lblContador);
			this.Controls.Add(this.groupBox1);
			this.Controls.Add(this.ultraLabel1);
			this.Controls.Add(this.btnBuscar);
			this.Controls.Add(this.txtBusqueda);
			this.Controls.Add(this.label8);
			this.Controls.Add(this.label6);
			this.Controls.Add(this.dtHasta);
			this.Controls.Add(this.dtDesde);
			this.KeyPreview = true;
			this.Name = "frmAprobarPresupuestoPedidoCXC";
			this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
			this.Text = "Aprobar pedidos \"Presupuesto\"";
			this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.frmAprobarPresupuestoPedidoCXC_KeyDown);
			this.Load += new System.EventHandler(this.frmAprobarPresupuestoPedidoCXC_Load);
			((System.ComponentModel.ISupportInitialize)(this.cmbEstado)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbProyecto)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridCreditos)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtBusqueda)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtHasta)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtDesde)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource3)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private Acceso miAcceso;
		private void UnloadMe()
		{
			this.Close();
		}
		bool bAprobar = false;
		bool bAnular = false;
		bool bVistaGeneral = false;
		private void frmAprobarPresupuestoPedidoCXC_Load(object sender, System.EventArgs e)
		{
			miAcceso = new Acceso(cdsSeteoF, "2609");

			if (!Funcion.Permiso("1743", cdsSeteoF))
			{				
				MessageBox.Show("No puede ingresar Aprobar pedidos 'Presupuesto'", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				BeginInvoke(new MethodInvoker(UnloadMe));
				return;
			}

			if (miAcceso.BAPAutorizar) bAprobar = true;
			if (miAcceso.Anular) bAnular = true;
      if (miAcceso.BVistas) bVistaGeneral = true;
		



			this.dtDesde.CalendarInfo.MinDate = DateTime.Parse("20/06/2019");
			this.dtHasta.CalendarInfo.MinDate = DateTime.Parse("20/06/2019");			

			this.dtDesde.Value = DateTime.Today;
			this.dtHasta.Value = DateTime.Today;

			
			int idUsuario = Funcion.iEscalarSQL(cdsSeteoF, string.Format("exec [PresupuestoListaCCBusqueda] {0}, 1", MenuLatinium.IdUsuario));
			int idLider = Funcion.iEscalarSQL(cdsSeteoF, string.Format("exec [ValidaLideres] {0}", MenuLatinium.IdUsuario));
			this.cmbProyecto.Value = idUsuario;	
			if (idLider == 1 ) this.cmbProyecto.Enabled = false;

			if ( bVistaGeneral)
			{
				ListaBodegasTodas(this.cmbProyecto, MenuLatinium.IdUsuario,   cdsSeteoF);
				idLider = 0;
			}
			else
			{
				ListaBodegas(this.cmbProyecto, MenuLatinium.IdUsuario,   cdsSeteoF);
			}
			Listaestado(this.cmbEstado,   cdsSeteoF);

		}

		int idLider = 0 ; 
		private void ListaBodegas (Infragistics.Win.UltraWinGrid.UltraCombo uCombo, int idUsuario,    C1.Data.C1DataSet ds)
		{
			uCombo.DataSource = Funcion.dvProcedimiento(ds, string.Format("Exec [PresupuestoListaCCBusqueda] {0}, {1}", idUsuario,  0));
		}
		private void ListaBodegasTodas (Infragistics.Win.UltraWinGrid.UltraCombo uCombo, int idUsuario,    C1.Data.C1DataSet ds)
		{
			uCombo.DataSource = Funcion.dvProcedimiento(ds, string.Format("Select p.idProyecto as Bodega , pp.Nombre as Nombre   from dbo.PresupuestoCCxPersonal p Inner Join dbo.Proyecto pp on pp.idProyecto = p.idProyecto group by  p.idProyecto  , pp.Nombre order by Nombre"));
		}
		private void Listaestado (Infragistics.Win.UltraWinGrid.UltraCombo uCombo,   C1.Data.C1DataSet ds)
		{
			uCombo.DataSource = Funcion.dvProcedimiento(ds, string.Format("  SELECT Estado, Nombres FROM ESTADOPEDIDOPRESUPUESTO order by nombres" ));
		}

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);
		
		}

		private void frmAprobarPresupuestoPedidoCXC_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
				if (e.KeyCode == Keys.Escape) this.Close();	
		}

		private void btnBuscar_Click(object sender, System.EventArgs e)
		{
			#region Variables
			if (!Validacion.vbFechaEnDocumentos(this.dtDesde, "Seleccione la fecha inicial", false, "", cdsSeteoF)) return;
			if (!Validacion.vbFechaEnDocumentos(this.dtHasta, "Seleccione la fecha final", false, "", cdsSeteoF)) return;
			if (!Validacion.vbComparaFechas(Convert.ToDateTime(this.dtDesde.Value), Convert.ToDateTime(this.dtHasta.Value), this.dtDesde, "La fecha inicial no puede ser posterior a la fecha final", false)) return;
			
			#endregion Variables

			int iEstado = 0;
			int iProyecto = 0;


			if (this.cmbEstado.ActiveRow != null) iEstado = (int)this.cmbEstado.Value;
			if (this.cmbProyecto.ActiveRow != null) iProyecto = (int)this.cmbProyecto.Value;
			


			string sSQL = string.Format("Exec [ConsultaMaestroPResupuestoCXC] {0}, '{1}', '{2}', {3}, {4}, '{5}', {6}", 
				idLider, Convert.ToDateTime(this.dtDesde.Value).ToString("yyyyMMdd"),  Convert.ToDateTime(this.dtHasta.Value).ToString("yyyyMMdd"),
				iEstado, iProyecto, (string)this.txtBusqueda.Text, MenuLatinium.IdUsuario );

			this.uGridCreditos.DataSource = Funcion.dvProcedimiento(cdsSeteoF, sSQL);

			if (this.uGridCreditos.Rows.Count > 0) this.uGridCreditos.Focus();


			this.lblContador.Text = this.uGridCreditos.Rows.Count + " Registros Encontrados";
		}

		private void uGridCreditos_ClickCellButton(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{

			#region Aprobacion
			if (e.Cell.Column.ToString() == "btnAprobar")
			{	
				if (!bAprobar) 
				{
					MessageBox.Show("No tiene acceso APROBAR ", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					return;
				}
				if ((string)e.Cell.Row.Cells["Nombres"].Value == "ANULADO DESDE ORDENES DE COMPRA")
				{
					return;
				}
				if ((string)e.Cell.Row.Cells["Nombres"].Value == "PROCESADO DESDE ODC")
				{
					return;
				}
				if ((string)e.Cell.Row.Cells["Nombres"].Value == "RECHAZADO DESDE ODC")
				{
					return;
				}
			


				if ((string)e.Cell.Row.Cells["Nombres"].Value == "RECHAZADO")
				{
					MessageBox.Show("El Documento ya esta Rechazado. ", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					return;
				}
				if ((string)e.Cell.Row.Cells["Nombres"].Value == "APROBADO")
				{
					MessageBox.Show("El Documento ya esta Aprobado. ", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					return;
				}

				using (frmNotas Nota = new frmNotas())
				{
					if (DialogResult.OK == Nota.ShowDialog())
					{	
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

				
					      
								string sSQLAutEs = string.Format("Update compra set Notas = Notas + '{0}', Estado = 2 where idcompra = {1}", Nota.txtNotas.Text,  (int) e.Cell.Row.Cells["idCompra"].Value);			
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
				}
			}
			#endregion Aprobacion

			#region Rechazar
			if (e.Cell.Column.ToString() == "btnRechazar")
			{	
				if (!bAnular) 
				{
					MessageBox.Show("No tiene acceso a RECHAZAR ", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					return;
				}
				if ((string)e.Cell.Row.Cells["Nombres"].Value == "ANULADO DESDE ORDENES DE COMPRA")
				{
					return;
				}
				if ((string)e.Cell.Row.Cells["Nombres"].Value == "PROCESADO DESDE ODC")
				{
					return;
				}
				if ((string)e.Cell.Row.Cells["Nombres"].Value == "RECHAZADO DESDE ODC")
				{
					return;
				}

				if ((string)e.Cell.Row.Cells["Nombres"].Value == "RECHAZADO")
				{
					MessageBox.Show("El Documento ya esta Rechazado. ", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					return;
				}

				using (frmNotas Nota = new frmNotas())
				{
					if (DialogResult.OK == Nota.ShowDialog())
					{	
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

				
					      
								string sSQLAutEs = string.Format("Update compra set Notas = Notas + '{0}', Estado = 3 where idcompra = {1}", Nota.txtNotas.Text,  (int) e.Cell.Row.Cells["idCompra"].Value);			
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
				}
			}
			#endregion Rechazar

		}
	}
}
