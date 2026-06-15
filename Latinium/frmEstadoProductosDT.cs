using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de frmEstadoProductosDT.
	/// </summary>
	public class frmEstadoProductosDT : System.Windows.Forms.Form
	{
		private System.Windows.Forms.Label label5;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtFechaHasta;
		private System.Windows.Forms.Label label3;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtFechaDesde;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtBusca;
		private System.Windows.Forms.Label label4;
		private System.Windows.Forms.Button btnActualizar;
		private System.Windows.Forms.Label label8;
		private System.Windows.Forms.Label label1;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtInfoSalida;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtInfoTecnico;
		private System.Windows.Forms.Label label2;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtDaño;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource1;
		private System.Windows.Forms.ToolTip toolTip1;
		private C1.Data.C1DataSet cdsSeteoF;
		private System.Windows.Forms.Label label6;
		private Infragistics.Win.UltraWinGrid.UltraGrid ugDTProductos;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtInfoRMA;
		private System.Windows.Forms.Label label7;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtAFacturar;
		private System.Windows.Forms.Label lblSubTotal;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtRepuesto;
		private System.Windows.Forms.Label label9;
		private System.Windows.Forms.Label label11;
		private System.Windows.Forms.Label label12;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtInfoAutorizacion;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtInfoSolicitud;
		private System.ComponentModel.IContainer components;

		public frmEstadoProductosDT()
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
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton1 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmEstadoProductosDT));
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton2 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.Appearance appearance1 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idIngreso");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Egreso");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPersonal");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Tecnico");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Ingreso");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Fecha");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CodArt");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Articulo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Serial");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn11 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Daño");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn12 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Estdo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn13 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn14 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("InfoTecnko");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn15 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("InformeSalida");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn16 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("InformeRMA");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn17 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ValorRepuesto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn18 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ValorFactura");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn19 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("NotasProducto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn20 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("InformeNC");
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idIngreso");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Egreso");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn3 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idPersonal");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn4 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Tecnico");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn5 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Ingreso");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn6 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Codigo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn7 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Fecha");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn8 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("CodArt");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn9 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Articulo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn10 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Serial");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn11 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Daño");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn12 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Estdo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn13 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn14 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("InfoTecnko");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn15 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("InformeSalida");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn16 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("InformeRMA");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn17 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("ValorRepuesto");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn18 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("ValorFactura");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn19 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("NotasProducto");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn20 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("InformeNC");
			Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance11 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance12 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance13 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance14 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance15 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance16 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance17 = new Infragistics.Win.Appearance();
			this.label5 = new System.Windows.Forms.Label();
			this.dtFechaHasta = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.label3 = new System.Windows.Forms.Label();
			this.dtFechaDesde = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.txtBusca = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label4 = new System.Windows.Forms.Label();
			this.btnActualizar = new System.Windows.Forms.Button();
			this.ugDTProductos = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource1 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.label8 = new System.Windows.Forms.Label();
			this.label1 = new System.Windows.Forms.Label();
			this.txtInfoSalida = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtInfoTecnico = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label2 = new System.Windows.Forms.Label();
			this.txtDaño = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.toolTip1 = new System.Windows.Forms.ToolTip(this.components);
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.label6 = new System.Windows.Forms.Label();
			this.txtInfoRMA = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label7 = new System.Windows.Forms.Label();
			this.txtAFacturar = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.lblSubTotal = new System.Windows.Forms.Label();
			this.txtRepuesto = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label9 = new System.Windows.Forms.Label();
			this.label11 = new System.Windows.Forms.Label();
			this.label12 = new System.Windows.Forms.Label();
			this.txtInfoAutorizacion = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtInfoSolicitud = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			((System.ComponentModel.ISupportInitialize)(this.dtFechaHasta)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFechaDesde)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtBusca)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ugDTProductos)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtInfoSalida)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtInfoTecnico)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDaño)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtInfoRMA)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtAFacturar)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtRepuesto)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtInfoAutorizacion)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtInfoSolicitud)).BeginInit();
			this.SuspendLayout();
			// 
			// label5
			// 
			this.label5.Font = new System.Drawing.Font("Tahoma", 8F, System.Drawing.FontStyle.Bold);
			this.label5.Location = new System.Drawing.Point(368, 8);
			this.label5.Name = "label5";
			this.label5.Size = new System.Drawing.Size(64, 16);
			this.label5.TabIndex = 686;
			this.label5.Text = "Hasta";
			this.label5.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
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
			this.dtFechaHasta.TabIndex = 685;
			this.dtFechaHasta.Value = ((object)(resources.GetObject("dtFechaHasta.Value")));
			// 
			// label3
			// 
			this.label3.Font = new System.Drawing.Font("Tahoma", 8F, System.Drawing.FontStyle.Bold);
			this.label3.Location = new System.Drawing.Point(256, 8);
			this.label3.Name = "label3";
			this.label3.Size = new System.Drawing.Size(64, 16);
			this.label3.TabIndex = 684;
			this.label3.Text = "Desde";
			this.label3.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
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
			this.dtFechaDesde.TabIndex = 683;
			this.dtFechaDesde.Value = ((object)(resources.GetObject("dtFechaDesde.Value")));
			// 
			// txtBusca
			// 
			appearance1.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtBusca.Appearance = appearance1;
			this.txtBusca.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtBusca.Location = new System.Drawing.Point(8, 24);
			this.txtBusca.Name = "txtBusca";
			this.txtBusca.Size = new System.Drawing.Size(240, 21);
			this.txtBusca.TabIndex = 682;
			this.txtBusca.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtBusca_KeyPress);
			// 
			// label4
			// 
			this.label4.Font = new System.Drawing.Font("Tahoma", 8F, System.Drawing.FontStyle.Bold);
			this.label4.Location = new System.Drawing.Point(8, 8);
			this.label4.Name = "label4";
			this.label4.Size = new System.Drawing.Size(208, 16);
			this.label4.TabIndex = 681;
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
			this.btnActualizar.TabIndex = 680;
			this.btnActualizar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnActualizar.Click += new System.EventHandler(this.btnActualizar_Click);
			// 
			// ugDTProductos
			// 
			this.ugDTProductos.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.ugDTProductos.Cursor = System.Windows.Forms.Cursors.Default;
			this.ugDTProductos.DataSource = this.ultraDataSource1;
			appearance2.BackColor = System.Drawing.Color.White;
			this.ugDTProductos.DisplayLayout.Appearance = appearance2;
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn2.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn2.Header.Caption = "# Egreso";
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridColumn3.Header.VisiblePosition = 2;
			ultraGridColumn3.Hidden = true;
			ultraGridColumn4.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn4.Header.VisiblePosition = 3;
			ultraGridColumn4.Width = 200;
			ultraGridColumn5.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn5.Header.Caption = "# Ingreso";
			ultraGridColumn5.Header.VisiblePosition = 4;
			ultraGridColumn6.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn6.Header.Caption = "Bod";
			ultraGridColumn6.Header.VisiblePosition = 5;
			ultraGridColumn6.Width = 35;
			ultraGridColumn7.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn7.Header.VisiblePosition = 6;
			ultraGridColumn8.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn8.Header.Caption = "Código";
			ultraGridColumn8.Header.VisiblePosition = 7;
			ultraGridColumn9.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn9.Header.VisiblePosition = 8;
			ultraGridColumn9.Width = 200;
			ultraGridColumn10.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn10.Header.Caption = "# Serie";
			ultraGridColumn10.Header.VisiblePosition = 9;
			ultraGridColumn10.Width = 120;
			ultraGridColumn11.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn11.Header.VisiblePosition = 10;
			ultraGridColumn11.Width = 300;
			ultraGridColumn12.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn12.Header.Caption = "Estado";
			ultraGridColumn12.Header.VisiblePosition = 11;
			ultraGridColumn12.Width = 100;
			ultraGridColumn13.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn13.Header.Caption = "Cliente";
			ultraGridColumn13.Header.VisiblePosition = 12;
			ultraGridColumn13.Width = 120;
			ultraGridColumn14.Header.VisiblePosition = 13;
			ultraGridColumn14.Hidden = true;
			ultraGridColumn15.Header.VisiblePosition = 14;
			ultraGridColumn15.Hidden = true;
			ultraGridColumn16.Header.VisiblePosition = 15;
			ultraGridColumn16.Hidden = true;
			ultraGridColumn17.Header.VisiblePosition = 16;
			ultraGridColumn18.Header.VisiblePosition = 17;
			ultraGridColumn19.Header.VisiblePosition = 18;
			ultraGridColumn19.Hidden = true;
			ultraGridColumn20.Header.VisiblePosition = 19;
			ultraGridColumn20.Hidden = true;
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
																										 ultraGridColumn16,
																										 ultraGridColumn17,
																										 ultraGridColumn18,
																										 ultraGridColumn19,
																										 ultraGridColumn20});
			this.ugDTProductos.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			appearance3.ForeColor = System.Drawing.Color.Black;
			appearance3.ForeColorDisabled = System.Drawing.Color.Black;
			this.ugDTProductos.DisplayLayout.Override.ActiveRowAppearance = appearance3;
			this.ugDTProductos.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.ugDTProductos.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance4.BackColor = System.Drawing.Color.Transparent;
			this.ugDTProductos.DisplayLayout.Override.CardAreaAppearance = appearance4;
			appearance5.ForeColor = System.Drawing.Color.Black;
			appearance5.ForeColorDisabled = System.Drawing.Color.Black;
			this.ugDTProductos.DisplayLayout.Override.GroupByRowAppearance = appearance5;
			appearance6.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance6.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance6.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance6.FontData.BoldAsString = "True";
			appearance6.FontData.Name = "Arial";
			appearance6.FontData.SizeInPoints = 10F;
			appearance6.ForeColor = System.Drawing.Color.White;
			appearance6.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.ugDTProductos.DisplayLayout.Override.HeaderAppearance = appearance6;
			this.ugDTProductos.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortSingle;
			appearance7.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance7.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance7.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ugDTProductos.DisplayLayout.Override.RowAlternateAppearance = appearance7;
			appearance8.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance8.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance8.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ugDTProductos.DisplayLayout.Override.RowSelectorAppearance = appearance8;
			appearance9.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance9.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance9.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ugDTProductos.DisplayLayout.Override.SelectedRowAppearance = appearance9;
			this.ugDTProductos.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.ugDTProductos.Location = new System.Drawing.Point(8, 56);
			this.ugDTProductos.Name = "ugDTProductos";
			this.ugDTProductos.Size = new System.Drawing.Size(1064, 168);
			this.ugDTProductos.TabIndex = 687;
			this.ugDTProductos.Click += new System.EventHandler(this.ugDTProductos_Click);
			// 
			// ultraDataSource1
			// 
			ultraDataColumn17.DataType = typeof(System.Decimal);
			ultraDataColumn18.DataType = typeof(System.Decimal);
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
																																 ultraDataColumn16,
																																 ultraDataColumn17,
																																 ultraDataColumn18,
																																 ultraDataColumn19,
																																 ultraDataColumn20});
			// 
			// label8
			// 
			this.label8.BackColor = System.Drawing.Color.Transparent;
			this.label8.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Bold);
			this.label8.Location = new System.Drawing.Point(408, 324);
			this.label8.Name = "label8";
			this.label8.Size = new System.Drawing.Size(56, 32);
			this.label8.TabIndex = 693;
			this.label8.Text = "Informe SALIDA:";
			this.label8.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
			// 
			// label1
			// 
			this.label1.BackColor = System.Drawing.Color.Transparent;
			this.label1.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Bold);
			this.label1.Location = new System.Drawing.Point(8, 324);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(64, 32);
			this.label1.TabIndex = 691;
			this.label1.Text = "Informe TECNICO:";
			this.label1.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
			// 
			// txtInfoSalida
			// 
			appearance10.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtInfoSalida.Appearance = appearance10;
			this.txtInfoSalida.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtInfoSalida.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtInfoSalida.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.txtInfoSalida.Location = new System.Drawing.Point(472, 324);
			this.txtInfoSalida.Multiline = true;
			this.txtInfoSalida.Name = "txtInfoSalida";
			this.txtInfoSalida.ReadOnly = true;
			this.txtInfoSalida.Scrollbars = System.Windows.Forms.ScrollBars.Horizontal;
			this.txtInfoSalida.Size = new System.Drawing.Size(328, 80);
			this.txtInfoSalida.TabIndex = 692;
			// 
			// txtInfoTecnico
			// 
			appearance11.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtInfoTecnico.Appearance = appearance11;
			this.txtInfoTecnico.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtInfoTecnico.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtInfoTecnico.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.txtInfoTecnico.Location = new System.Drawing.Point(80, 324);
			this.txtInfoTecnico.Multiline = true;
			this.txtInfoTecnico.Name = "txtInfoTecnico";
			this.txtInfoTecnico.ReadOnly = true;
			this.txtInfoTecnico.Scrollbars = System.Windows.Forms.ScrollBars.Horizontal;
			this.txtInfoTecnico.Size = new System.Drawing.Size(312, 80);
			this.txtInfoTecnico.TabIndex = 690;
			// 
			// label2
			// 
			this.label2.BackColor = System.Drawing.Color.Transparent;
			this.label2.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Bold);
			this.label2.Location = new System.Drawing.Point(8, 232);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(64, 32);
			this.label2.TabIndex = 689;
			this.label2.Text = "DAÑO Producto:";
			this.label2.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
			// 
			// txtDaño
			// 
			appearance12.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtDaño.Appearance = appearance12;
			this.txtDaño.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtDaño.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtDaño.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.txtDaño.Location = new System.Drawing.Point(80, 232);
			this.txtDaño.Multiline = true;
			this.txtDaño.Name = "txtDaño";
			this.txtDaño.ReadOnly = true;
			this.txtDaño.Scrollbars = System.Windows.Forms.ScrollBars.Horizontal;
			this.txtDaño.Size = new System.Drawing.Size(312, 88);
			this.txtDaño.TabIndex = 688;
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
			// label6
			// 
			this.label6.BackColor = System.Drawing.Color.Transparent;
			this.label6.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Bold);
			this.label6.Location = new System.Drawing.Point(408, 232);
			this.label6.Name = "label6";
			this.label6.Size = new System.Drawing.Size(56, 32);
			this.label6.TabIndex = 695;
			this.label6.Text = "Informe RMA:";
			this.label6.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
			// 
			// txtInfoRMA
			// 
			appearance13.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtInfoRMA.Appearance = appearance13;
			this.txtInfoRMA.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtInfoRMA.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtInfoRMA.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.txtInfoRMA.Location = new System.Drawing.Point(472, 232);
			this.txtInfoRMA.Multiline = true;
			this.txtInfoRMA.Name = "txtInfoRMA";
			this.txtInfoRMA.ReadOnly = true;
			this.txtInfoRMA.Scrollbars = System.Windows.Forms.ScrollBars.Horizontal;
			this.txtInfoRMA.Size = new System.Drawing.Size(328, 88);
			this.txtInfoRMA.TabIndex = 694;
			// 
			// label7
			// 
			this.label7.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.label7.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label7.Location = new System.Drawing.Point(904, 299);
			this.label7.Name = "label7";
			this.label7.Size = new System.Drawing.Size(48, 24);
			this.label7.TabIndex = 698;
			this.label7.Text = "$ Val. A Facturar";
			this.label7.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
			// 
			// txtAFacturar
			// 
			this.txtAFacturar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			appearance14.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtAFacturar.Appearance = appearance14;
			this.txtAFacturar.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtAFacturar.Enabled = false;
			this.txtAFacturar.FormatString = "#,##0.00";
			this.txtAFacturar.Location = new System.Drawing.Point(952, 299);
			this.txtAFacturar.Name = "txtAFacturar";
			this.txtAFacturar.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtAFacturar.PromptChar = ' ';
			this.txtAFacturar.Size = new System.Drawing.Size(64, 21);
			this.txtAFacturar.TabIndex = 699;
			// 
			// lblSubTotal
			// 
			this.lblSubTotal.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.lblSubTotal.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblSubTotal.Location = new System.Drawing.Point(888, 275);
			this.lblSubTotal.Name = "lblSubTotal";
			this.lblSubTotal.Size = new System.Drawing.Size(64, 24);
			this.lblSubTotal.TabIndex = 696;
			this.lblSubTotal.Text = "$ Val. Respuesto";
			this.lblSubTotal.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
			// 
			// txtRepuesto
			// 
			this.txtRepuesto.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			appearance15.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtRepuesto.Appearance = appearance15;
			this.txtRepuesto.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtRepuesto.Enabled = false;
			this.txtRepuesto.FormatString = "#,##0.00";
			this.txtRepuesto.Location = new System.Drawing.Point(952, 275);
			this.txtRepuesto.Name = "txtRepuesto";
			this.txtRepuesto.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtRepuesto.PromptChar = ' ';
			this.txtRepuesto.Size = new System.Drawing.Size(64, 21);
			this.txtRepuesto.TabIndex = 697;
			// 
			// label9
			// 
			this.label9.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.label9.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label9.ForeColor = System.Drawing.Color.Maroon;
			this.label9.Location = new System.Drawing.Point(824, 347);
			this.label9.Name = "label9";
			this.label9.Size = new System.Drawing.Size(240, 29);
			this.label9.TabIndex = 700;
			this.label9.Text = "Estado PRODUCTO:";
			this.label9.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label11
			// 
			this.label11.BackColor = System.Drawing.Color.Transparent;
			this.label11.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Bold);
			this.label11.Location = new System.Drawing.Point(400, 408);
			this.label11.Name = "label11";
			this.label11.Size = new System.Drawing.Size(64, 32);
			this.label11.TabIndex = 706;
			this.label11.Text = "Informe AUTORIZ:";
			this.label11.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
			// 
			// label12
			// 
			this.label12.BackColor = System.Drawing.Color.Transparent;
			this.label12.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Bold);
			this.label12.Location = new System.Drawing.Point(0, 408);
			this.label12.Name = "label12";
			this.label12.Size = new System.Drawing.Size(72, 24);
			this.label12.TabIndex = 704;
			this.label12.Text = "Informe SOLICITUD:";
			this.label12.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
			// 
			// txtInfoAutorizacion
			// 
			appearance16.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtInfoAutorizacion.Appearance = appearance16;
			this.txtInfoAutorizacion.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtInfoAutorizacion.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtInfoAutorizacion.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.txtInfoAutorizacion.Location = new System.Drawing.Point(472, 408);
			this.txtInfoAutorizacion.Multiline = true;
			this.txtInfoAutorizacion.Name = "txtInfoAutorizacion";
			this.txtInfoAutorizacion.ReadOnly = true;
			this.txtInfoAutorizacion.Scrollbars = System.Windows.Forms.ScrollBars.Horizontal;
			this.txtInfoAutorizacion.Size = new System.Drawing.Size(328, 80);
			this.txtInfoAutorizacion.TabIndex = 705;
			// 
			// txtInfoSolicitud
			// 
			appearance17.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtInfoSolicitud.Appearance = appearance17;
			this.txtInfoSolicitud.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtInfoSolicitud.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtInfoSolicitud.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.txtInfoSolicitud.Location = new System.Drawing.Point(80, 408);
			this.txtInfoSolicitud.Multiline = true;
			this.txtInfoSolicitud.Name = "txtInfoSolicitud";
			this.txtInfoSolicitud.ReadOnly = true;
			this.txtInfoSolicitud.Scrollbars = System.Windows.Forms.ScrollBars.Horizontal;
			this.txtInfoSolicitud.Size = new System.Drawing.Size(312, 80);
			this.txtInfoSolicitud.TabIndex = 703;
			// 
			// frmEstadoProductosDT
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(1080, 494);
			this.Controls.Add(this.label11);
			this.Controls.Add(this.label12);
			this.Controls.Add(this.txtInfoAutorizacion);
			this.Controls.Add(this.txtInfoSolicitud);
			this.Controls.Add(this.label9);
			this.Controls.Add(this.label7);
			this.Controls.Add(this.txtAFacturar);
			this.Controls.Add(this.lblSubTotal);
			this.Controls.Add(this.txtRepuesto);
			this.Controls.Add(this.label6);
			this.Controls.Add(this.txtInfoRMA);
			this.Controls.Add(this.label8);
			this.Controls.Add(this.label1);
			this.Controls.Add(this.txtInfoSalida);
			this.Controls.Add(this.txtInfoTecnico);
			this.Controls.Add(this.label2);
			this.Controls.Add(this.txtDaño);
			this.Controls.Add(this.ugDTProductos);
			this.Controls.Add(this.label5);
			this.Controls.Add(this.dtFechaHasta);
			this.Controls.Add(this.label3);
			this.Controls.Add(this.dtFechaDesde);
			this.Controls.Add(this.txtBusca);
			this.Controls.Add(this.label4);
			this.Controls.Add(this.btnActualizar);
			this.Name = "frmEstadoProductosDT";
			this.Text = "Estado de Productos DEPARTAMENTO TECNICO";
			this.Load += new System.EventHandler(this.frmEstadoProductosDT_Load);
			((System.ComponentModel.ISupportInitialize)(this.dtFechaHasta)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFechaDesde)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtBusca)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ugDTProductos)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtInfoSalida)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtInfoTecnico)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDaño)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtInfoRMA)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtAFacturar)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtRepuesto)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtInfoAutorizacion)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtInfoSolicitud)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);
		}

		private void frmEstadoProductosDT_Load(object sender, System.EventArgs e)
		{
			this.dtFechaDesde.Value = "01/" + DateTime.Now.Month.ToString() + "/" + DateTime.Now.Year.ToString();
			this.dtFechaHasta.Value = DateTime.Today;
		}

		private void btnActualizar_Click(object sender, System.EventArgs e)
		{
			DateTime dtFechaI = (DateTime)this.dtFechaDesde.Value;
			DateTime dtFechaF = (DateTime)this.dtFechaHasta.Value;

			string sSQL = string.Format("Exec DT_EstadoProductos_DT '{0}', '{1}', '{2}'", 
				dtFechaI.ToString("yyyyMMdd"), dtFechaF.ToString("yyyyMMdd"), this.txtBusca.Text.ToString().Trim());
			this.ugDTProductos.DataSource = Funcion.dvProcedimiento(cdsSeteoF, sSQL);
		}

		private void ugDTProductos_Click(object sender, System.EventArgs e)
		{
			/* Visualizar el daño del producto */
			this.txtDaño.Value = this.ugDTProductos.ActiveRow.Cells["Daño"].Value;
			this.txtInfoTecnico.Value = this.ugDTProductos.ActiveRow.Cells["InfoTecnko"].Value;
			this.txtInfoSalida.Value = this.ugDTProductos.ActiveRow.Cells["InformeSalida"].Value;
			this.txtInfoRMA.Value = this.ugDTProductos.ActiveRow.Cells["InformeRMA"].Value;
			this.txtAFacturar.Value = this.ugDTProductos.ActiveRow.Cells["ValorFactura"].Value;
			this.txtRepuesto.Value = this.ugDTProductos.ActiveRow.Cells["ValorRepuesto"].Value;
			this.label9.Text = "Estado PRODUCTO:"+this.ugDTProductos.ActiveRow.Cells["Estdo"].Value;
			this.txtInfoSolicitud.Value = this.ugDTProductos.ActiveRow.Cells["NotasProducto"].Value;
			this.txtInfoAutorizacion.Value = this.ugDTProductos.ActiveRow.Cells["InformeNC"].Value;
		}

		private void txtBusca_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (e.KeyChar == 13) btnActualizar.PerformClick();
		}
	}
}
