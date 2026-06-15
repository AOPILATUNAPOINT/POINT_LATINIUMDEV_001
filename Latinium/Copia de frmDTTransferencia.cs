using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using CrystalDecisions;
using CrystalDecisions.CrystalReports.Engine;
using CrystalDecisions.ReportSource;
using CrystalDecisions.Shared;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de frmDTTransferencia.
	/// </summary>
	public class frmDTTransferencia : System.Windows.Forms.Form
	{
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource1;
		private System.Windows.Forms.Label label1;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbBodega;
		private System.Windows.Forms.Button btnBuscar;
		private System.Windows.Forms.Label lblContador;
		private System.Windows.Forms.Button btnTransferir;
		private C1.Data.C1DataSet cdsSeteoF;
		private Infragistics.Win.UltraWinGrid.UltraGrid uGridIngresos;
		private System.Windows.Forms.Label lblNumero;
		private System.Windows.Forms.Label label2;
		public Infragistics.Win.UltraWinEditors.UltraTextEditor txtNumero;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbTransportista;
		private System.Windows.Forms.GroupBox groupBox1;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource2;
		private Infragistics.Win.UltraWinGrid.UltraDropDown cmbBodegaDestino;
		private System.Windows.Forms.Label label3;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbBodegaOrigen;
		private System.Windows.Forms.Label label4;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public frmDTTransferencia()
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
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Sel");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idIngreso");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("BodegaDestino", -1, "cmbBodegaDestino");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Estado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Ingreso");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaIngreso");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaReparado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idArticulo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn11 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idSerial");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn12 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Serial");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn13 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Tecnico");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn14 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("NotasReparacion");
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Sel");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idIngreso");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn3 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("BodegaDestino");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn4 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Estado");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn5 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Ingreso");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn6 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("FechaIngreso");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn7 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("FechaReparado");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn8 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn9 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idArticulo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn10 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Codigo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn11 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idSerial");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn12 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Serial");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn13 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Tecnico");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn14 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("NotasReparacion");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn15 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bodega");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn16 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn17 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo", 0);
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmDTTransferencia));
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand3 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn18 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPersonal");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn19 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand4 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn20 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bodega");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn21 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn22 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn15 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Bodega");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn16 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Codigo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn17 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand5 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn23 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bodega");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn24 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn25 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo", 0);
			this.uGridIngresos = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource1 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.label1 = new System.Windows.Forms.Label();
			this.cmbBodega = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.btnBuscar = new System.Windows.Forms.Button();
			this.lblContador = new System.Windows.Forms.Label();
			this.btnTransferir = new System.Windows.Forms.Button();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.lblNumero = new System.Windows.Forms.Label();
			this.label2 = new System.Windows.Forms.Label();
			this.txtNumero = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.cmbTransportista = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.groupBox1 = new System.Windows.Forms.GroupBox();
			this.cmbBodegaDestino = new Infragistics.Win.UltraWinGrid.UltraDropDown();
			this.ultraDataSource2 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.cmbBodegaOrigen = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.label3 = new System.Windows.Forms.Label();
			this.label4 = new System.Windows.Forms.Label();
			((System.ComponentModel.ISupportInitialize)(this.uGridIngresos)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbBodega)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNumero)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbTransportista)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbBodegaDestino)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbBodegaOrigen)).BeginInit();
			this.SuspendLayout();
			// 
			// uGridIngresos
			// 
			this.uGridIngresos.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.uGridIngresos.Cursor = System.Windows.Forms.Cursors.Default;
			this.uGridIngresos.DataSource = this.ultraDataSource1;
			appearance1.BackColor = System.Drawing.Color.White;
			this.uGridIngresos.DisplayLayout.Appearance = appearance1;
			ultraGridColumn1.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn1.Header.Caption = "...";
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn1.Width = 20;
			ultraGridColumn2.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridColumn2.Hidden = true;
			ultraGridColumn3.ButtonDisplayStyle = Infragistics.Win.UltraWinGrid.ButtonDisplayStyle.Always;
			ultraGridColumn3.Header.Caption = "Destino";
			ultraGridColumn3.Header.VisiblePosition = 2;
			ultraGridColumn3.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownList;
			ultraGridColumn3.Width = 150;
			ultraGridColumn4.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn4.Header.VisiblePosition = 3;
			ultraGridColumn4.Width = 150;
			ultraGridColumn5.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn5.Header.VisiblePosition = 4;
			ultraGridColumn5.Width = 100;
			ultraGridColumn6.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn6.Header.Caption = "Ingresado";
			ultraGridColumn6.Header.VisiblePosition = 5;
			ultraGridColumn6.Width = 70;
			ultraGridColumn7.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn7.Header.Caption = "Reparado";
			ultraGridColumn7.Header.VisiblePosition = 6;
			ultraGridColumn7.Width = 70;
			ultraGridColumn8.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn8.Header.Caption = "Cliente";
			ultraGridColumn8.Header.VisiblePosition = 7;
			ultraGridColumn8.Width = 220;
			ultraGridColumn9.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn9.Header.VisiblePosition = 8;
			ultraGridColumn9.Hidden = true;
			ultraGridColumn10.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn10.Header.VisiblePosition = 9;
			ultraGridColumn10.Width = 180;
			ultraGridColumn11.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn11.Header.VisiblePosition = 11;
			ultraGridColumn11.Hidden = true;
			ultraGridColumn12.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn12.Header.VisiblePosition = 10;
			ultraGridColumn12.Width = 140;
			ultraGridColumn13.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn13.Header.VisiblePosition = 13;
			ultraGridColumn13.Width = 180;
			ultraGridColumn14.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn14.Header.Caption = "Notas Reparación";
			ultraGridColumn14.Header.VisiblePosition = 12;
			ultraGridColumn14.Width = 500;
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
																										 ultraGridColumn14});
			this.uGridIngresos.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			appearance2.ForeColor = System.Drawing.Color.Black;
			appearance2.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridIngresos.DisplayLayout.Override.ActiveRowAppearance = appearance2;
			this.uGridIngresos.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
			this.uGridIngresos.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.uGridIngresos.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance3.BackColor = System.Drawing.Color.Transparent;
			this.uGridIngresos.DisplayLayout.Override.CardAreaAppearance = appearance3;
			appearance4.ForeColor = System.Drawing.Color.Black;
			appearance4.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridIngresos.DisplayLayout.Override.GroupByRowAppearance = appearance4;
			appearance5.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance5.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance5.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance5.FontData.BoldAsString = "True";
			appearance5.FontData.Name = "Arial";
			appearance5.FontData.SizeInPoints = 8.5F;
			appearance5.ForeColor = System.Drawing.Color.White;
			appearance5.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridIngresos.DisplayLayout.Override.HeaderAppearance = appearance5;
			this.uGridIngresos.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortSingle;
			appearance6.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance6.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance6.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridIngresos.DisplayLayout.Override.RowAlternateAppearance = appearance6;
			appearance7.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance7.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance7.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridIngresos.DisplayLayout.Override.RowSelectorAppearance = appearance7;
			appearance8.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance8.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance8.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridIngresos.DisplayLayout.Override.SelectedRowAppearance = appearance8;
			this.uGridIngresos.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridIngresos.Location = new System.Drawing.Point(8, 80);
			this.uGridIngresos.Name = "uGridIngresos";
			this.uGridIngresos.Size = new System.Drawing.Size(1076, 272);
			this.uGridIngresos.TabIndex = 662;
			this.uGridIngresos.KeyDown += new System.Windows.Forms.KeyEventHandler(this.uGridIngresos_KeyDown);
			this.uGridIngresos.AfterCellUpdate += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.uGridIngresos_AfterCellUpdate);
			// 
			// ultraDataSource1
			// 
			ultraDataColumn1.DataType = typeof(bool);
			ultraDataColumn1.DefaultValue = false;
			ultraDataColumn2.DataType = typeof(int);
			ultraDataColumn3.DataType = typeof(int);
			ultraDataColumn6.DataType = typeof(System.DateTime);
			ultraDataColumn7.DataType = typeof(System.DateTime);
			ultraDataColumn9.DataType = typeof(int);
			ultraDataColumn11.DataType = typeof(int);
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
																																 ultraDataColumn14});
			// 
			// label1
			// 
			this.label1.AutoSize = true;
			this.label1.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label1.Location = new System.Drawing.Point(8, 40);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(83, 17);
			this.label1.TabIndex = 700;
			this.label1.Text = "Bodega Destino";
			this.label1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// cmbBodega
			// 
			this.cmbBodega.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbBodega.Cursor = System.Windows.Forms.Cursors.Default;
			ultraGridColumn15.Header.VisiblePosition = 2;
			ultraGridColumn16.Header.VisiblePosition = 1;
			ultraGridColumn16.Width = 180;
			ultraGridColumn17.Header.VisiblePosition = 0;
			ultraGridBand2.Columns.AddRange(new object[] {
																										 ultraGridColumn15,
																										 ultraGridColumn16,
																										 ultraGridColumn17});
			this.cmbBodega.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			this.cmbBodega.DisplayMember = "Nombre";
			this.cmbBodega.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbBodega.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbBodega.Location = new System.Drawing.Point(104, 40);
			this.cmbBodega.MaxDropDownItems = 30;
			this.cmbBodega.Name = "cmbBodega";
			this.cmbBodega.Size = new System.Drawing.Size(208, 22);
			this.cmbBodega.TabIndex = 0;
			this.cmbBodega.ValueMember = "Bodega";
			this.cmbBodega.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cmbBodega_KeyDown);
			this.cmbBodega.ValueChanged += new System.EventHandler(this.cmbBodega_ValueChanged);
			// 
			// btnBuscar
			// 
			this.btnBuscar.CausesValidation = false;
			this.btnBuscar.Image = ((System.Drawing.Image)(resources.GetObject("btnBuscar.Image")));
			this.btnBuscar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnBuscar.Location = new System.Drawing.Point(320, 40);
			this.btnBuscar.Name = "btnBuscar";
			this.btnBuscar.Size = new System.Drawing.Size(80, 23);
			this.btnBuscar.TabIndex = 1;
			this.btnBuscar.Text = "Buscar";
			this.btnBuscar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnBuscar.Click += new System.EventHandler(this.btnBuscar_Click);
			// 
			// lblContador
			// 
			this.lblContador.AutoSize = true;
			this.lblContador.Font = new System.Drawing.Font("Tahoma", 9F);
			this.lblContador.ForeColor = System.Drawing.Color.Firebrick;
			this.lblContador.Location = new System.Drawing.Point(424, 40);
			this.lblContador.Name = "lblContador";
			this.lblContador.Size = new System.Drawing.Size(0, 18);
			this.lblContador.TabIndex = 702;
			this.lblContador.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// btnTransferir
			// 
			this.btnTransferir.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btnTransferir.CausesValidation = false;
			this.btnTransferir.Enabled = false;
			this.btnTransferir.Image = ((System.Drawing.Image)(resources.GetObject("btnTransferir.Image")));
			this.btnTransferir.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnTransferir.Location = new System.Drawing.Point(1000, 32);
			this.btnTransferir.Name = "btnTransferir";
			this.btnTransferir.Size = new System.Drawing.Size(80, 23);
			this.btnTransferir.TabIndex = 4;
			this.btnTransferir.Text = "Transferir";
			this.btnTransferir.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnTransferir.Click += new System.EventHandler(this.btnTransferir_Click);
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
			// lblNumero
			// 
			this.lblNumero.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.lblNumero.AutoSize = true;
			this.lblNumero.BackColor = System.Drawing.Color.Transparent;
			this.lblNumero.Location = new System.Drawing.Point(744, 35);
			this.lblNumero.Name = "lblNumero";
			this.lblNumero.Size = new System.Drawing.Size(94, 16);
			this.lblNumero.TabIndex = 707;
			this.lblNumero.Text = "Guia de Remision";
			// 
			// label2
			// 
			this.label2.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.label2.AutoSize = true;
			this.label2.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label2.Location = new System.Drawing.Point(744, 11);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(70, 17);
			this.label2.TabIndex = 705;
			this.label2.Text = "Transportista";
			this.label2.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtNumero
			// 
			this.txtNumero.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			appearance9.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtNumero.Appearance = appearance9;
			this.txtNumero.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNumero.Enabled = false;
			this.txtNumero.Location = new System.Drawing.Point(856, 33);
			this.txtNumero.MaxLength = 10;
			this.txtNumero.Name = "txtNumero";
			this.txtNumero.Size = new System.Drawing.Size(104, 21);
			this.txtNumero.TabIndex = 3;
			this.txtNumero.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtNumero_KeyDown);
			this.txtNumero.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtNumero_KeyPress);
			this.txtNumero.Validated += new System.EventHandler(this.txtNumero_Validated);
			// 
			// cmbTransportista
			// 
			this.cmbTransportista.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.cmbTransportista.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbTransportista.Cursor = System.Windows.Forms.Cursors.Default;
			ultraGridColumn18.Header.VisiblePosition = 0;
			ultraGridColumn18.Hidden = true;
			ultraGridColumn19.Header.VisiblePosition = 1;
			ultraGridColumn19.Width = 250;
			ultraGridBand3.Columns.AddRange(new object[] {
																										 ultraGridColumn18,
																										 ultraGridColumn19});
			this.cmbTransportista.DisplayLayout.BandsSerializer.Add(ultraGridBand3);
			this.cmbTransportista.DisplayMember = "Nombre";
			this.cmbTransportista.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbTransportista.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbTransportista.Enabled = false;
			this.cmbTransportista.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbTransportista.Location = new System.Drawing.Point(856, 8);
			this.cmbTransportista.MaxDropDownItems = 30;
			this.cmbTransportista.Name = "cmbTransportista";
			this.cmbTransportista.Size = new System.Drawing.Size(224, 22);
			this.cmbTransportista.TabIndex = 2;
			this.cmbTransportista.ValueMember = "idPersonal";
			this.cmbTransportista.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cmbTransportista_KeyDown);
			this.cmbTransportista.ValueChanged += new System.EventHandler(this.cmbTransportista_ValueChanged);
			this.cmbTransportista.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.cmbTransportista_InitializeLayout);
			// 
			// groupBox1
			// 
			this.groupBox1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.groupBox1.Location = new System.Drawing.Point(0, 64);
			this.groupBox1.Name = "groupBox1";
			this.groupBox1.Size = new System.Drawing.Size(1096, 8);
			this.groupBox1.TabIndex = 708;
			this.groupBox1.TabStop = false;
			// 
			// cmbBodegaDestino
			// 
			this.cmbBodegaDestino.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbBodegaDestino.DataSource = this.ultraDataSource2;
			ultraGridColumn20.Header.VisiblePosition = 0;
			ultraGridColumn20.Hidden = true;
			ultraGridColumn21.Header.VisiblePosition = 1;
			ultraGridColumn21.Width = 60;
			ultraGridColumn22.Header.VisiblePosition = 2;
			ultraGridColumn22.Width = 200;
			ultraGridBand4.Columns.AddRange(new object[] {
																										 ultraGridColumn20,
																										 ultraGridColumn21,
																										 ultraGridColumn22});
			this.cmbBodegaDestino.DisplayLayout.BandsSerializer.Add(ultraGridBand4);
			this.cmbBodegaDestino.DisplayMember = "Nombre";
			this.cmbBodegaDestino.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbBodegaDestino.Location = new System.Drawing.Point(272, 136);
			this.cmbBodegaDestino.Name = "cmbBodegaDestino";
			this.cmbBodegaDestino.Size = new System.Drawing.Size(264, 112);
			this.cmbBodegaDestino.TabIndex = 709;
			this.cmbBodegaDestino.ValueMember = "Bodega";
			this.cmbBodegaDestino.Visible = false;
			// 
			// ultraDataSource2
			// 
			ultraDataColumn15.DataType = typeof(int);
			this.ultraDataSource2.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn15,
																																 ultraDataColumn16,
																																 ultraDataColumn17});
			// 
			// cmbBodegaOrigen
			// 
			this.cmbBodegaOrigen.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbBodegaOrigen.Cursor = System.Windows.Forms.Cursors.Default;
			ultraGridColumn23.Header.VisiblePosition = 2;
			ultraGridColumn24.Header.VisiblePosition = 1;
			ultraGridColumn24.Width = 180;
			ultraGridColumn25.Header.VisiblePosition = 0;
			ultraGridBand5.Columns.AddRange(new object[] {
																										 ultraGridColumn23,
																										 ultraGridColumn24,
																										 ultraGridColumn25});
			this.cmbBodegaOrigen.DisplayLayout.BandsSerializer.Add(ultraGridBand5);
			this.cmbBodegaOrigen.DisplayMember = "Nombre";
			this.cmbBodegaOrigen.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbBodegaOrigen.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbBodegaOrigen.Location = new System.Drawing.Point(104, 8);
			this.cmbBodegaOrigen.MaxDropDownItems = 30;
			this.cmbBodegaOrigen.Name = "cmbBodegaOrigen";
			this.cmbBodegaOrigen.Size = new System.Drawing.Size(208, 22);
			this.cmbBodegaOrigen.TabIndex = 710;
			this.cmbBodegaOrigen.ValueMember = "Bodega";
			this.cmbBodegaOrigen.ValueChanged += new System.EventHandler(this.cmbBodegaOrigen_ValueChanged);
			// 
			// label3
			// 
			this.label3.AutoSize = true;
			this.label3.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label3.Location = new System.Drawing.Point(8, 11);
			this.label3.Name = "label3";
			this.label3.Size = new System.Drawing.Size(78, 17);
			this.label3.TabIndex = 711;
			this.label3.Text = "Bodega Origen";
			this.label3.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label4
			// 
			this.label4.AutoSize = true;
			this.label4.Font = new System.Drawing.Font("Tahoma", 9F);
			this.label4.ForeColor = System.Drawing.Color.Firebrick;
			this.label4.Location = new System.Drawing.Point(424, 34);
			this.label4.Name = "label4";
			this.label4.Size = new System.Drawing.Size(0, 18);
			this.label4.TabIndex = 712;
			this.label4.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// frmDTTransferencia
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(1092, 366);
			this.Controls.Add(this.label4);
			this.Controls.Add(this.label3);
			this.Controls.Add(this.cmbBodegaOrigen);
			this.Controls.Add(this.cmbBodegaDestino);
			this.Controls.Add(this.groupBox1);
			this.Controls.Add(this.lblNumero);
			this.Controls.Add(this.label2);
			this.Controls.Add(this.label1);
			this.Controls.Add(this.txtNumero);
			this.Controls.Add(this.cmbTransportista);
			this.Controls.Add(this.btnTransferir);
			this.Controls.Add(this.btnBuscar);
			this.Controls.Add(this.lblContador);
			this.Controls.Add(this.cmbBodega);
			this.Controls.Add(this.uGridIngresos);
			this.KeyPreview = true;
			this.Name = "frmDTTransferencia";
			this.Text = "Transferencia de mercaderia de servicio tecnico a locales";
			this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.frmDTTransferencia_KeyDown);
			this.Load += new System.EventHandler(this.frmDTTransferencia_Load);
			((System.ComponentModel.ISupportInitialize)(this.uGridIngresos)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbBodega)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNumero)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbTransportista)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbBodegaDestino)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbBodegaOrigen)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		int idPersonal = 0;

		private void frmDTTransferencia_Load(object sender, System.EventArgs e)
		{
			idPersonal = Funcion.iEscalarSQL(cdsSeteoF, string.Format("If Exists(Select idPersonal From Nomina Where Codigo In (Select Nombre From Usuario Where idUsuario = {0})) Select idPersonal From Nomina Where Codigo In (Select Nombre From Usuario Where idUsuario = {0}) Else Select 0", MenuLatinium.IdUsuario));

			FuncionesProcedimientos.ListaBodegas(this.cmbBodegaOrigen, MenuLatinium.IdUsuario, 57, DateTime.Today, false, cdsSeteoF);
						
			this.cmbTransportista.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Exec DT_Personal 28");
		}

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);
		}

		private void Consulta()
		{
			int iBodega = 0;
			if (this.cmbBodega.ActiveRow != null) iBodega = (int)this.cmbBodega.Value;

			string sSQL = string.Format("Exec DTListaMercaderiaTransferencia {0}", iBodega);
			FuncionesProcedimientos.dsGeneral(sSQL, this.uGridIngresos);

			this.lblContador.Text = this.uGridIngresos.Rows.Count + " REGISTROS ENCONTRADOS";
		}

		private void btnBuscar_Click(object sender, System.EventArgs e)
		{
			Consulta();
		}

		private void ImprimeGuiaDeRemision(int idGuiaRemision)
		{
			ParameterFields paramFields = new ParameterFields ();

			ParameterField Compra = new ParameterField ();
			ParameterDiscreteValue IdCompra = new ParameterDiscreteValue ();
			//Compra.ParameterFieldName = "@idMaestroGuiaRemision";
			Compra.ParameterFieldName = "@idCPA";
			IdCompra.Value = idGuiaRemision;
			Compra.CurrentValues.Add (IdCompra);
			paramFields.Add (Compra);
			
			//Reporte miRep = new Reporte("GuiaDeRemisionCarpa.rpt", "");
			Reporte miRep = new Reporte("Doc_GR.rpt", "");
			miRep.MdiParent = this.MdiParent;
			miRep.Titulo("Guia de Remisión");
			miRep.crVista.ParameterFieldInfo = paramFields;
			miRep.Show();
			
			Funcion.EjecutaSQL(cdsSeteoF, string.Format("Update MaestroGuiaRemision Set Impreso = 1 Where idMaestroGuiaRemision = {0}", idGuiaRemision));
		}

		private void btnTransferir_Click(object sender, System.EventArgs e)
		{
			if (this.uGridIngresos.Rows.Count == 0)
			{
				MessageBox.Show("No hay ingresos para transferir", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.cmbBodega.Focus();
				return;
			}

			if (!Validacion.vbComboVacio(this.cmbBodegaOrigen, "Seleccione la bodega origen")) return;

			if (!Validacion.vbComboVacio(this.cmbBodega, "Seleccione una bodega destino")) return;
						
			if (!Validacion.vbComboVacio(this.cmbTransportista, "Seleccione la persona que transporta la mercaderia")) return;
			
			if ((int)this.cmbTransportista.Value == 1013)
				if (!Validacion.vbTexto(this.txtNumero, 10, 10, "Guia de remision de Servientrega")) return;
		
			int iCont = 0;
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridIngresos.Rows.All)			
				if ((bool)dr.Cells["Sel"].Value) iCont++;

			if (iCont == 0)
			{
				MessageBox.Show("Seleccione al menos un ingreso para crear la transferencia", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.cmbBodega.Focus();
				return;
			}
			
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridIngresos.Rows.All)			
			{
				if ((bool)dr.Cells["Sel"].Value)
				{
					if (dr.Cells["BodegaDestino"].Value == System.DBNull.Value)
					{
						MessageBox.Show("Seleccione el Local Destino", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
						this.uGridIngresos.ActiveRow = dr;
						this.uGridIngresos.ActiveRow.Selected = true;
						return;
					}
				}
			}

			if (DialogResult.Yes == MessageBox.Show(string.Format("Esta seguro de generar la transferencia a la bodega '{0}'", this.cmbBodega.Text.ToString().Trim()), "Point Technology", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2)) 
			{	
				this.Cursor = Cursors.WaitCursor;
				try
				{					
					#region Guia de remision
					int idGuiaRemision = 0;
					DateTime dtInicio = DateTime.Today;
					DateTime dtFin = DateTime.Today.AddDays(3);
					string sNumeroGR = this.txtNumero.Text.ToString();
					
					if ((int)this.cmbTransportista.Value != 1013 && ((int)this.cmbBodega.Value != 36 || (int)this.cmbBodega.Value != 12))
					{
						bool bAutomatico = Funcion.bEscalarSQL(cdsSeteoF, string.Format("Select AutGuiaRemision From Bodega Where Bodega = {0}", 12), true);		
						
						int iBodega = 12;
						if ((int)this.cmbBodegaOrigen.Value == 95) iBodega = 15;

						if (bAutomatico) sNumeroGR = Funcion.sEscalarSQL(cdsSeteoF, string.Format("Exec NumeracionLocales 30, {0}, 0", iBodega));
										
						string sSQL = string.Format("Exec GGuiaRemision {0}, {1}, '{2}', '{3}', {4}, '{5}', {6}, {7}, {8}, '{9}', '{10}', {11}, '{12}', '{13}'", 
							0, 0, dtInicio.ToString("yyyyMMdd"), dtFin.ToString("yyyyMMdd"), (int)this.cmbBodegaOrigen.Value, sNumeroGR, 42, 5, 
							(int)this.cmbTransportista.Value, "", "DOCUMENTO GENERADO POR TRANSFERENCIA DE PRODUCTOS DE SERVICIO TECNICO POINT", (int)this.cmbBodega.Value, this.cmbBodega.Text.ToString(), 0);				
						idGuiaRemision = Funcion.iEscalarSQL(cdsSeteoF, sSQL);
					}
					else sNumeroGR = this.txtNumero.Text.ToString();
					#endregion Guia de remision

					#region Detalle Guia de Remision - Movimiento
					foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridIngresos.Rows.All)
					{
						if ((bool)dr.Cells["Sel"].Value) 
						{
							#region Detalle	Guia de Remision
							if ((int)this.cmbTransportista.Value != 1013)
							{
								string sSQLDetalle = string.Format("Exec GuardaGuiaDeRemision {0}, {1}, {2}, {3}",
									0, idGuiaRemision, (int)dr.Cells["idArticulo"].Value, (int)dr.Cells["idSerial"].Value);
								Funcion.EjecutaSQL(cdsSeteoF, sSQLDetalle, true);
							}
							#endregion Detalle Guia de Remision
							
							#region Guarda Movimiento
							string sSQLProcesos = string.Format("Exec GuardaDT_Procesos {0}, {1}, 26, {2}, 'ENVIO DE SERVICIO TECNICO POINT A ALMACEN {5}', {3}, '{4}', 0, 0, 0", 
								(int)dr.Cells["idIngreso"].Value, (int)this.cmbBodegaOrigen.Value, idPersonal, idGuiaRemision, sNumeroGR, this.cmbBodega.Text.ToString().Trim());
							Funcion.EjecutaSQL(cdsSeteoF, sSQLProcesos, true);
							#endregion Guarda Movimiento
						}					
					}
					#endregion Detalle Guia de Remision - Movimiento			
						
					if ((int)this.cmbTransportista.Value != 1013)
					{
						string sClaveAcceso = Funcion.sEscalarSQL(cdsSeteoF, string.Format("Exec FE_ClaveDeAcceso {0}, 6", idGuiaRemision));

						string cadInvertida = Facturacion.invertirCadena(sClaveAcceso.Substring(0, 48));
						string ValM11 = Facturacion.Modulo11(cadInvertida).ToString().Trim();
						string sClaveSRI = sClaveAcceso.ToString().Substring(0, 48) + ValM11;

						Funcion.EjecutaSQL(cdsSeteoF, string.Format("Update MaestroGuiaRemision Set Autorizacion = '{0}' Where idMaestroGuiaRemision = {1}", sClaveSRI, idGuiaRemision));

						ImprimeGuiaDeRemision(idGuiaRemision);
					}
								
					MessageBox.Show("Transferencia procesada correctamente", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);					
			
					Consulta();			
				}
				catch (Exception ex)
				{
					MessageBox.Show(string.Format("Exception Type: {0} {1}", ex.GetType(), ex.Message), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);							
				}
				this.Cursor = Cursors.Default;
			}			
		}

		private void cmbBodega_ValueChanged(object sender, System.EventArgs e)
		{
			this.cmbTransportista.Enabled = false;
			this.txtNumero.Enabled = false;
			this.btnTransferir.Enabled = false;
			this.uGridIngresos.DisplayLayout.Bands[0].Columns["Sel"].CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;

			if (this.cmbBodega.ActiveRow != null) 
			{
				this.cmbTransportista.Enabled = true;
				this.txtNumero.Enabled = true;
				this.btnTransferir.Enabled = true; 
				this.uGridIngresos.DisplayLayout.Bands[0].Columns["Sel"].CellActivation = Infragistics.Win.UltraWinGrid.Activation.AllowEdit;
			}

			Consulta();
		}

		private void frmDTTransferencia_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Escape) this.Close();
		}

		private void cmbTransportista_ValueChanged(object sender, System.EventArgs e)
		{
			this.txtNumero.Text = "";
			this.txtNumero.Enabled = false;

			if (this.cmbTransportista.ActiveRow != null)				
				if ((int)this.cmbTransportista.Value == 1013)
					this.txtNumero.Enabled = true;			
		}

		private void cmbTransportista_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.txtNumero.Focus();
		}

		private void txtNumero_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.btnTransferir.Focus();
		}

		private void txtNumero_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			Funcion.CamposNumericos(sender, e);
		}

		private void cmbBodega_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.btnBuscar_Click(sender, e);
		}

		private void txtNumero_Validated(object sender, System.EventArgs e)
		{
			if (this.txtNumero.Text.ToString().Trim().Length > 0)
			{
				char Pad = '0';
				int iCar = 9;

				if ((int)this.cmbTransportista.Value != 1013) iCar = 10;

				this.txtNumero.Text = this.txtNumero.Value.ToString().PadLeft(iCar, Pad);
			}
		}

		private void uGridIngresos_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			FuncionesProcedimientos.DesplazamientoGrids(sender, e, this.uGridIngresos);
		}

		private void uGridIngresos_AfterCellUpdate(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			if (e.Cell.Column.ToString() == "BodegaDestino")
			{
				string sSQL = string.Format("Update DTIngreso Set BodegaDestino = {0} Where idIngreso = {1}", 
					(int)e.Cell.Row.Cells["BodegaDestino"].Value, (int)e.Cell.Row.Cells["idIngreso"].Value);
				Funcion.EjecutaSQL(cdsSeteoF, sSQL, true);
			}
		}

		private void cmbTransportista_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}

		private void cmbBodegaOrigen_ValueChanged(object sender, System.EventArgs e)
		{
			if (this.cmbBodegaOrigen.ActiveRow != null)
			{
				int iRegion = Funcion.iEscalarSQL(cdsSeteoF, string.Format("Select Region From Bodega Where Bodega = {0}", (int)this.cmbBodegaOrigen.Value));

				this.cmbBodega.DataSource = Funcion.dvProcedimiento(cdsSeteoF, string.Format("Exec DTListaBodegasSalida {0}", iRegion));

				this.cmbBodegaDestino.DataSource = Funcion.dvProcedimiento(cdsSeteoF, string.Format("Exec DTListaBodegasSalida {0}", iRegion));
			}
		}
	}
}
