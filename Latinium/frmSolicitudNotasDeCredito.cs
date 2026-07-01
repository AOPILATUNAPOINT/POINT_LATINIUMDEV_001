using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;
using System.Data.SqlClient;
using Latinium.Models;
using Latinium.Services;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de frmSolicitudNotasDeCredito.
	/// </summary>
	public class frmSolicitudNotasDeCredito : System.Windows.Forms.Form
	{
		public Infragistics.Win.UltraWinGrid.UltraGrid uGridFacturas;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource1;
		private Infragistics.Win.Misc.UltraLabel lblLocal;
		private Infragistics.Win.Misc.UltraLabel ultraLabel1;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbBodega;
		private System.Windows.Forms.Button btnBuscar;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtBusqueda;
		private Infragistics.Win.UltraWinEditors.UltraOptionSet optEstados;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbMotivos;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource2;
		private C1.Data.C1DataSet cdsSeteoF;
		private System.Windows.Forms.Label label3;
		private System.Windows.Forms.GroupBox groupBox1;
		private System.Windows.Forms.Button btnExcel;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public frmSolicitudNotasDeCredito()
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
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCompra");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idFormaPago");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FormaPago");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bodega");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Almacen");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Numero");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Fecha");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Identificacion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cliente");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Total");
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn11 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idMotivo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn12 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Motivo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn13 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("EstadoSolNC");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn14 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Estado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn15 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaSolicitudNC");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn16 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaAprobacionNC");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn17 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("UsuarioAutoriza");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn18 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("btnAutorizar", 0);
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmSolicitudNotasDeCredito));
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCompra");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idFormaPago");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn3 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("FormaPago");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn4 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Bodega");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn5 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Almacen");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn6 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Numero");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn7 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Fecha");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn8 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Identificacion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn9 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Cliente");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn10 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Total");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn11 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idMotivo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn12 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Motivo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn13 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("EstadoSolNC");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn14 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Estado");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn15 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("FechaSolicitudNC");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn16 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("FechaAprobacionNC");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn17 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("UsuarioAutoriza");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn19 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bodega");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn20 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn21 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo", 0);
			Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance11 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance12 = new Infragistics.Win.Appearance();
			Infragistics.Win.ValueListItem valueListItem1 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem2 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem3 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem4 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem5 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.Appearance appearance13 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand3 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn22 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idMotivo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn23 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Motivo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn18 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idMotivo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn19 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Motivo");
			this.uGridFacturas = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource1 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.lblLocal = new Infragistics.Win.Misc.UltraLabel();
			this.ultraLabel1 = new Infragistics.Win.Misc.UltraLabel();
			this.cmbBodega = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.btnBuscar = new System.Windows.Forms.Button();
			this.txtBusqueda = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.optEstados = new Infragistics.Win.UltraWinEditors.UltraOptionSet();
			this.cmbMotivos = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.ultraDataSource2 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.label3 = new System.Windows.Forms.Label();
			this.groupBox1 = new System.Windows.Forms.GroupBox();
			this.btnExcel = new System.Windows.Forms.Button();
			((System.ComponentModel.ISupportInitialize)(this.uGridFacturas)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbBodega)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtBusqueda)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.optEstados)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbMotivos)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			this.SuspendLayout();
			// 
			// uGridFacturas
			// 
			this.uGridFacturas.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.uGridFacturas.Cursor = System.Windows.Forms.Cursors.Default;
			this.uGridFacturas.DataSource = this.ultraDataSource1;
			appearance1.BackColor = System.Drawing.Color.White;
			appearance1.ForeColor = System.Drawing.Color.Black;
			appearance1.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridFacturas.DisplayLayout.Appearance = appearance1;
			ultraGridColumn1.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn1.Width = 10;
			ultraGridColumn2.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn2.Header.VisiblePosition = 2;
			ultraGridColumn2.Hidden = true;
			ultraGridColumn3.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn3.Header.VisiblePosition = 3;
			ultraGridColumn3.Hidden = true;
			ultraGridColumn3.Width = 87;
			ultraGridColumn4.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn4.Header.VisiblePosition = 5;
			ultraGridColumn4.Hidden = true;
			ultraGridColumn5.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn5.Header.VisiblePosition = 6;
			ultraGridColumn5.Width = 60;
			ultraGridColumn6.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn6.Header.VisiblePosition = 7;
			ultraGridColumn6.Width = 80;
			ultraGridColumn7.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn7.Header.VisiblePosition = 4;
			ultraGridColumn7.Width = 75;
			ultraGridColumn8.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn8.Header.VisiblePosition = 8;
			ultraGridColumn8.Width = 90;
			ultraGridColumn9.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn9.Header.VisiblePosition = 9;
			ultraGridColumn9.Width = 200;
			ultraGridColumn10.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance2.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn10.CellAppearance = appearance2;
			ultraGridColumn10.Format = "#,##0.00";
			ultraGridColumn10.Header.VisiblePosition = 10;
			ultraGridColumn10.Width = 80;
			ultraGridColumn11.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn11.Header.VisiblePosition = 13;
			ultraGridColumn11.Hidden = true;
			ultraGridColumn12.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn12.Header.VisiblePosition = 11;
			ultraGridColumn12.Width = 130;
			ultraGridColumn13.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn13.Header.VisiblePosition = 16;
			ultraGridColumn13.Hidden = true;
			ultraGridColumn14.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn14.Header.VisiblePosition = 12;
			ultraGridColumn14.Width = 110;
			ultraGridColumn15.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn15.Format = "dd/MM/yyyy HH:mm";
			ultraGridColumn15.Header.Caption = "Solicitud";
			ultraGridColumn15.Header.VisiblePosition = 14;
			ultraGridColumn15.Width = 110;
			ultraGridColumn16.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn16.Format = "dd/MM/yyyy HH:mm";
			ultraGridColumn16.Header.Caption = "Aprobacion";
			ultraGridColumn16.Header.VisiblePosition = 15;
			ultraGridColumn16.Width = 110;
			ultraGridColumn17.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn17.Header.Caption = "Usuario Autoriza";
			ultraGridColumn17.Header.VisiblePosition = 17;
			ultraGridColumn17.Width = 110;
			ultraGridColumn18.ButtonDisplayStyle = Infragistics.Win.UltraWinGrid.ButtonDisplayStyle.Always;
			appearance3.Image = ((object)(resources.GetObject("appearance3.Image")));
			appearance3.ImageHAlign = Infragistics.Win.HAlign.Center;
			appearance3.ImageVAlign = Infragistics.Win.VAlign.Middle;
			ultraGridColumn18.CellButtonAppearance = appearance3;
			ultraGridColumn18.Header.Caption = "A";
			ultraGridColumn18.Header.VisiblePosition = 1;
			ultraGridColumn18.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.Button;
			ultraGridColumn18.Width = 40;
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
																										 ultraGridColumn18});
			ultraGridBand1.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti;
			this.uGridFacturas.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			appearance4.ForeColor = System.Drawing.Color.Black;
			appearance4.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridFacturas.DisplayLayout.Override.ActiveRowAppearance = appearance4;
			appearance5.BackColor = System.Drawing.Color.Transparent;
			this.uGridFacturas.DisplayLayout.Override.CardAreaAppearance = appearance5;
			appearance6.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance6.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance6.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance6.FontData.BoldAsString = "True";
			appearance6.FontData.Name = "Arial";
			appearance6.FontData.SizeInPoints = 9F;
			appearance6.ForeColor = System.Drawing.Color.White;
			appearance6.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridFacturas.DisplayLayout.Override.HeaderAppearance = appearance6;
			this.uGridFacturas.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti;
			appearance7.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance7.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance7.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridFacturas.DisplayLayout.Override.RowAlternateAppearance = appearance7;
			appearance8.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance8.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance8.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridFacturas.DisplayLayout.Override.RowSelectorAppearance = appearance8;
			appearance9.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance9.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance9.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridFacturas.DisplayLayout.Override.SelectedRowAppearance = appearance9;
			this.uGridFacturas.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridFacturas.Location = new System.Drawing.Point(8, 128);
			this.uGridFacturas.Name = "uGridFacturas";
			this.uGridFacturas.Size = new System.Drawing.Size(1136, 232);
			this.uGridFacturas.TabIndex = 7;
			this.uGridFacturas.ClickCellButton += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.uGridFacturas_ClickCellButton);
			this.uGridFacturas.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.uGridFacturas_InitializeLayout);
			// 
			// ultraDataSource1
			// 
			ultraDataColumn1.DataType = typeof(int);
			ultraDataColumn2.DataType = typeof(int);
			ultraDataColumn4.DataType = typeof(int);
			ultraDataColumn7.DataType = typeof(System.DateTime);
			ultraDataColumn10.DataType = typeof(System.Decimal);
			ultraDataColumn11.DataType = typeof(int);
			ultraDataColumn13.DataType = typeof(int);
			ultraDataColumn15.DataType = typeof(System.DateTime);
			ultraDataColumn16.DataType = typeof(System.DateTime);
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
																																 ultraDataColumn17});
			// 
			// lblLocal
			// 
			this.lblLocal.AutoSize = true;
			this.lblLocal.BackColor = System.Drawing.Color.Transparent;
			this.lblLocal.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.lblLocal.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(50)), ((System.Byte)(50)), ((System.Byte)(50)));
			this.lblLocal.Location = new System.Drawing.Point(8, 40);
			this.lblLocal.Name = "lblLocal";
			this.lblLocal.Size = new System.Drawing.Size(30, 15);
			this.lblLocal.TabIndex = 269;
			this.lblLocal.Text = "Local";
			// 
			// ultraLabel1
			// 
			this.ultraLabel1.AutoSize = true;
			this.ultraLabel1.BackColor = System.Drawing.Color.Transparent;
			this.ultraLabel1.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.ultraLabel1.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(50)), ((System.Byte)(50)), ((System.Byte)(50)));
			this.ultraLabel1.Location = new System.Drawing.Point(8, 75);
			this.ultraLabel1.Name = "ultraLabel1";
			this.ultraLabel1.Size = new System.Drawing.Size(37, 15);
			this.ultraLabel1.TabIndex = 268;
			this.ultraLabel1.Text = "Buscar";
			// 
			// cmbBodega
			// 
			this.cmbBodega.CausesValidation = false;
			this.cmbBodega.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbBodega.Cursor = System.Windows.Forms.Cursors.Default;
			ultraGridColumn19.Header.VisiblePosition = 2;
			ultraGridColumn20.Header.VisiblePosition = 1;
			ultraGridColumn20.Width = 180;
			ultraGridColumn21.Header.VisiblePosition = 0;
			ultraGridBand2.Columns.AddRange(new object[] {
																										 ultraGridColumn19,
																										 ultraGridColumn20,
																										 ultraGridColumn21});
			this.cmbBodega.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			this.cmbBodega.DisplayMember = "Nombre";
			this.cmbBodega.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbBodega.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbBodega.Enabled = false;
			this.cmbBodega.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbBodega.Location = new System.Drawing.Point(64, 40);
			this.cmbBodega.MaxDropDownItems = 30;
			this.cmbBodega.Name = "cmbBodega";
			this.cmbBodega.Size = new System.Drawing.Size(280, 21);
			this.cmbBodega.TabIndex = 266;
			this.cmbBodega.ValueMember = "Bodega";
			this.cmbBodega.ValueChanged += new System.EventHandler(this.cmbBodega_ValueChanged);
			this.cmbBodega.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.cmbBodega_InitializeLayout);
			// 
			// btnBuscar
			// 
			this.btnBuscar.CausesValidation = false;
			this.btnBuscar.Location = new System.Drawing.Point(272, 71);
			this.btnBuscar.Name = "btnBuscar";
			this.btnBuscar.TabIndex = 265;
			this.btnBuscar.Text = "&Buscar";
			this.btnBuscar.Click += new System.EventHandler(this.btnBuscar_Click);
			// 
			// txtBusqueda
			// 
			appearance10.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtBusqueda.Appearance = appearance10;
			this.txtBusqueda.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtBusqueda.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtBusqueda.Location = new System.Drawing.Point(64, 72);
			this.txtBusqueda.Name = "txtBusqueda";
			this.txtBusqueda.Size = new System.Drawing.Size(192, 22);
			this.txtBusqueda.TabIndex = 270;
			this.txtBusqueda.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtBusqueda_KeyDown);
			// 
			// optEstados
			// 
			appearance11.ForeColorDisabled = System.Drawing.Color.Black;
			this.optEstados.Appearance = appearance11;
			this.optEstados.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.optEstados.CheckedIndex = 0;
			this.optEstados.ItemAppearance = appearance12;
			this.optEstados.ItemOrigin = new System.Drawing.Point(10, 0);
			valueListItem1.DataValue = 0;
			valueListItem1.DisplayText = "Pendiente de Solicitud";
			valueListItem2.DataValue = 1;
			valueListItem2.DisplayText = "Pendiente Aprobación";
			valueListItem3.DataValue = 2;
			valueListItem3.DisplayText = "Aprobados";
			valueListItem4.DataValue = 3;
			valueListItem4.DisplayText = "Rechazados";
			valueListItem5.DataValue = 4;
			valueListItem5.DisplayText = "Procesados";
			this.optEstados.Items.Add(valueListItem1);
			this.optEstados.Items.Add(valueListItem2);
			this.optEstados.Items.Add(valueListItem3);
			this.optEstados.Items.Add(valueListItem4);
			this.optEstados.Items.Add(valueListItem5);
			this.optEstados.ItemSpacingVertical = 10;
			this.optEstados.Location = new System.Drawing.Point(424, 8);
			this.optEstados.Name = "optEstados";
			this.optEstados.Size = new System.Drawing.Size(568, 24);
			this.optEstados.TabIndex = 696;
			this.optEstados.Text = "Pendiente de Solicitud";
			this.optEstados.ValueChanged += new System.EventHandler(this.optEstados_ValueChanged);
			// 
			// cmbMotivos
			// 
			appearance13.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbMotivos.Appearance = appearance13;
			this.cmbMotivos.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbMotivos.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbMotivos.DataSource = this.ultraDataSource2;
			ultraGridColumn22.Header.VisiblePosition = 0;
			ultraGridColumn22.Hidden = true;
			ultraGridColumn23.Header.VisiblePosition = 1;
			ultraGridColumn23.Width = 272;
			ultraGridBand3.Columns.AddRange(new object[] {
																										 ultraGridColumn22,
																										 ultraGridColumn23});
			this.cmbMotivos.DisplayLayout.BandsSerializer.Add(ultraGridBand3);
			this.cmbMotivos.DisplayMember = "Motivo";
			this.cmbMotivos.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbMotivos.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbMotivos.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbMotivos.Location = new System.Drawing.Point(64, 8);
			this.cmbMotivos.Name = "cmbMotivos";
			this.cmbMotivos.Size = new System.Drawing.Size(280, 21);
			this.cmbMotivos.TabIndex = 697;
			this.cmbMotivos.ValueMember = "idMotivo";
			this.cmbMotivos.ValueChanged += new System.EventHandler(this.cmbMotivos_ValueChanged);
			// 
			// ultraDataSource2
			// 
			ultraDataColumn18.DataType = typeof(int);
			this.ultraDataSource2.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn18,
																																 ultraDataColumn19});
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
			// label3
			// 
			this.label3.AutoSize = true;
			this.label3.Location = new System.Drawing.Point(8, 8);
			this.label3.Name = "label3";
			this.label3.Size = new System.Drawing.Size(37, 17);
			this.label3.TabIndex = 698;
			this.label3.Text = "Mótivo";
			this.label3.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// groupBox1
			// 
			this.groupBox1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.groupBox1.Location = new System.Drawing.Point(0, 112);
			this.groupBox1.Name = "groupBox1";
			this.groupBox1.Size = new System.Drawing.Size(1152, 8);
			this.groupBox1.TabIndex = 699;
			this.groupBox1.TabStop = false;
			// 
			// btnExcel
			// 
			this.btnExcel.Image = ((System.Drawing.Image)(resources.GetObject("btnExcel.Image")));
			this.btnExcel.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnExcel.Location = new System.Drawing.Point(1064, 72);
			this.btnExcel.Name = "btnExcel";
			this.btnExcel.Size = new System.Drawing.Size(75, 24);
			this.btnExcel.TabIndex = 978;
			this.btnExcel.Text = "Excel";
			this.btnExcel.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnExcel.Visible = false;
			this.btnExcel.Click += new System.EventHandler(this.btnExcel_Click);
			// 
			// frmSolicitudNotasDeCredito
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 14);
			this.ClientSize = new System.Drawing.Size(1152, 366);
			this.Controls.Add(this.btnExcel);
			this.Controls.Add(this.groupBox1);
			this.Controls.Add(this.label3);
			this.Controls.Add(this.optEstados);
			this.Controls.Add(this.txtBusqueda);
			this.Controls.Add(this.lblLocal);
			this.Controls.Add(this.ultraLabel1);
			this.Controls.Add(this.cmbBodega);
			this.Controls.Add(this.btnBuscar);
			this.Controls.Add(this.uGridFacturas);
			this.Controls.Add(this.cmbMotivos);
			this.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.Name = "frmSolicitudNotasDeCredito";
			this.Text = "Solicitud De Notas De Crédito";
			this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.frmSolicitudNotasDeCredito_KeyDown);
			this.Load += new System.EventHandler(this.frmSolicitudNotasDeCredito_Load);
			((System.ComponentModel.ISupportInitialize)(this.uGridFacturas)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbBodega)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtBusqueda)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.optEstados)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbMotivos)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion


		DateTime dtFecha = DateTime.Today;
		string sClaveAcceso = "";

		bool bConsultas = false;

		private Acceso miAcceso;
		private void UnloadMe()
		{
			this.Close();
		}

		private void frmSolicitudNotasDeCredito_Load(object sender, System.EventArgs e)
		{
			miAcceso = new Acceso(cdsSeteoF, "0217");

			if (!Funcion.Permiso("1214", cdsSeteoF))
			{				
				MessageBox.Show("No tiene Acceso a Solicitudes de N/C", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				BeginInvoke(new MethodInvoker(UnloadMe));
				return;
			}

			dtFecha = FuncionesProcedimientos.dtRetornaFechaServidor(cdsSeteoF);

			FuncionesProcedimientos.ListaBodegas(this.cmbBodega, MenuLatinium.IdUsuario, 3, dtFecha, false, cdsSeteoF);
			
			int iBodPred = Funcion.iEscalarSQL(cdsSeteoF, string.Format("Exec BodegaAsignadaPorUsuario {0}", MenuLatinium.IdUsuario));

			if (iBodPred == 0) 
			{
				this.cmbBodega.Enabled = true;
				this.cmbBodega.Value = 1; 				
			}
			else this.cmbBodega.Value = iBodPred;

			if (miAcceso.Buscar)
				this.cmbMotivos.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Select idMotivo, Motivo From MotivosDevolucion Where idTipoFactura = 5 And idMotivo In (30, 45, 4) Order By Motivo");
			else
				this.cmbMotivos.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Select idMotivo, Motivo From MotivosDevolucion Where idTipoFactura = 5 And idMotivo In (30, 45) Order By Motivo");

			if (miAcceso.BSolNCAFD) 
			{
				this.cmbMotivos.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Select idMotivo, Motivo From MotivosDevolucion Where idTipoFactura = 5 And idMotivo In (30, 45, 5) Order By Motivo");
			}
			if (miAcceso.BExportar) this.btnExcel.Visible  = true;
			this.cmbMotivos.Value = 45;
		}

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);		
		}

		private void Consulta()
		{
			if(bConsultas)
			{
				int iBodega = 0;
				if (this.cmbBodega.ActiveRow != null) iBodega = (int)this.cmbBodega.Value;

				int idMotivo = 0;
				if (this.cmbMotivos.ActiveRow != null) idMotivo = (int)this.cmbMotivos.Value;
			
				if(miAcceso.Buscar)
				{
					if(idMotivo == 4)
					{
						string sSQL = string.Format("Exec SolicitudNotasDeCredito '{0}', {1}, {2}, {3}, '{4}', 0, 1, {5}", 
							dtFecha.ToString("yyyyMMdd"), 0, (int)this.optEstados.Value, idMotivo, this.txtBusqueda.Text.ToString(), MenuLatinium.iNivel);
						FuncionesProcedimientos.dsGeneral(sSQL, this.uGridFacturas);
					}
					else
					{
						string sSQL = string.Format("Exec SolicitudNotasDeCredito '{0}', {1}, {2}, {3}, '{4}', 0, 0, 1", 
							dtFecha.ToString("yyyyMMdd"), iBodega, (int)this.optEstados.Value, idMotivo, this.txtBusqueda.Text.ToString());
						FuncionesProcedimientos.dsGeneral(sSQL, this.uGridFacturas);
					}
				}
				else
				{
					string sSQL = string.Format("Exec SolicitudNotasDeCredito '{0}', {1}, {2}, {3}, '{4}', 0, 0, 1", 
						dtFecha.ToString("yyyyMMdd"), iBodega, (int)this.optEstados.Value, idMotivo, this.txtBusqueda.Text.ToString());
					FuncionesProcedimientos.dsGeneral(sSQL, this.uGridFacturas);
				}
			}
		}

		private void btnBuscar_Click(object sender, System.EventArgs e)
		{
			bConsultas = true;
			Consulta();
			bConsultas = false;
		}

		string sNumeroNC = "";

		private int NotaDeCredito(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			DateTime dtHoy = FuncionesProcedimientos.dtRetornaFechaServidor(cdsSeteoF);
		
			int idNotaDeCredito = 0;

			if (DialogResult.Yes == MessageBox.Show("Este proceso es irreversible.\n\nEsta seguro de generar una nota de credito este documento?", "Point Technology", MessageBoxButtons.YesNo, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2)) 
			{			
				using (SqlConnection oConexion = new SqlConnection("Password=Bl45o6$9;Persist Security Info=True;User ID= " + MenuLatinium.stUsuario + ";Initial Catalog=" + MenuLatinium.stDirFacturacion + ";Data Source=" + MenuLatinium.stDirServidor + "; max pool size=2048"))
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

						oCmdActualiza.CommandText = string.Format("Exec SolicitudNotasDeCreditoEstados {0}, {1}, '{2}', {3}", 
							4, (int)this.cmbMotivos.Value, "NOTA DE CREDITO GENERADA", (int)e.Cell.Row.Cells["idCompra"].Value);
						oCmdActualiza.ExecuteNonQuery	();

						bool bFiscal = true;						

						oCmdActualiza.CommandText = string.Format("Exec NumeracionLocales 5, 1, {0}", bFiscal);
						sNumeroNC = oCmdActualiza.ExecuteScalar().ToString();

						oCmdActualiza.CommandText = string.Format("Exec CreaNotaDeCreditoCambioProducto {0}, '{1}', {2}, '{3}'", 
							(int)e.Cell.Row.Cells["idCompra"].Value, sNumeroNC, (int)this.cmbMotivos.Value, "");
						idNotaDeCredito = (int)oCmdActualiza.ExecuteScalar();
						
						if (bFiscal)
						{
							oCmdActualiza.CommandText = string.Format("Exec FE_ClaveDeAcceso {0}, 5", idNotaDeCredito);;
							string sClaveAcceso = oCmdActualiza.ExecuteScalar().ToString();

							string cadInvertida = Facturacion.invertirCadena(sClaveAcceso.Substring(0, 48));
							string ValM11 = Facturacion.Modulo11(cadInvertida).ToString().Trim();
							string sClaveSRI = sClaveAcceso.ToString().Substring(0, 48) + ValM11;

							oCmdActualiza.CommandText = string.Format("Update Compra Set AutFactura = '{0}' Where idCompra = {1}", sClaveSRI, idNotaDeCredito);
							oCmdActualiza.ExecuteNonQuery();
						}
																		
						oTransaction.Commit();	

						oConexion.Close();

						oTransaction.Dispose();

						//Registrar el regisro de nota de crédito por medio de la idCompra que va hacer 
						#region FamilyFriends
						ConsultarEstadoFamilyFriends((int)e.Cell.Row.Cells["idCompra"].Value);
						FamilyFriendsCompraNotaCredito(FamilyFriendsService.EstadoFFCompraNotaCredito, (int)e.Cell.Row.Cells["idCompra"].Value, idNotaDeCredito);
						#endregion FamilyFriends

						MessageBox.Show("Nota de Crédito generada correctamente", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
					}
					catch(Exception Exc)
					{
						oTransaction.Rollback();

						if (idBloqueaTransacciones > 0)
						{
							oCmdActualiza.CommandText = string.Format("Exec RegistraBloqueoTransacciones {0}, 0, {1}, 1", idBloqueaTransacciones, (int)e.Cell.Row.Cells["idCompra"].Value);
							oCmdActualiza.ExecuteNonQuery	();
						}

						MessageBox.Show(string.Format("Commit Exception Type : {0} {1}", Exc.GetType(), Exc.Message), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);

						idNotaDeCredito = 0;
					}								
				}
			}

			return idNotaDeCredito;	
		}

		// Solo una variable para todo el estado FF
		private FamilyFriendsEstadoModel estadoFF = new FamilyFriendsEstadoModel();

		private void ConsultarEstadoFamilyFriends(int idCompra)
		{
			try
			{
				FamilyFriendsService service =
					new FamilyFriendsService(frmRequerimiento.sconexionPoint);
				FamilyFriendsEstadoModel temp = service.ObtenerEstadoFF(idCompra);
				if (temp != null && temp.ExisteRegistro)
				{
					estadoFF = temp;
				}
			}
			catch
			{
			}
		}


		private void FamilyFriendsCompraNotaCredito(int estado, int idCompra, int idNotaCredito)
		{
			string nombreClase = this.GetType().Name;
			string nombreMetodo = "FamilyFriendsCompraNotaCredito";

			// =====================================================
			// 1) CONSULTA FINAL DE ESTADO FF (SIN EFECTOS COLATERALES)
			// =====================================================
			ConsultarEstadoFamilyFriends(idCompra);

			// =====================================================
			// Log de estadoFF luego de consultar BD
			// =====================================================
			Funcion.LogSistema(frmRequerimiento.sconexionInicio, nombreClase + ": " + nombreMetodo,
				string.Format(
				"EstadoFF post-consulta | EsFF: {0}, RechazoPromocion: {1}, IdWEB_FF_Invitado: {2}",
				estadoFF.EsFF,
				estadoFF.RechazoPromocion,
				estadoFF.IdWEB_FF_Invitado
				),
				nombreMetodo,
				Funcion.slogtipoinformacion
				);

			// =====================================================
			// 2) SI NO EXISTE REGISTRO FF  NO PROCESAR
			// =====================================================
			if (!estadoFF.ExisteRegistro)
			{
				Funcion.LogSistema(frmRequerimiento.sconexionInicio, nombreClase + ": " + nombreMetodo,
					"No existe registro FF en BD, no se procesa Nota de Crédito",
					nombreMetodo,
					Funcion.slogtipoinformacion
					);
				return;
			}

			// =====================================================
			// 3) VALIDACIONES EXISTENTES (NO SE TOCAN)
			// =====================================================
			if (!estadoFF.EsFF && !estadoFF.RechazoPromocion)
			{
				Funcion.LogSistema(frmRequerimiento.sconexionInicio, nombreClase + ": " + nombreMetodo,
					"No se procesa la compra: no es FF ni rechazo explícito",
					nombreMetodo,
					Funcion.slogtipoinformacion
					);
				return;
			}

			if (estado != FamilyFriendsService.EstadoFFCompraNotaCredito)
			{
				Funcion.LogSistema(frmRequerimiento.sconexionInicio, nombreClase + ": " + nombreMetodo,
					"No se procesa la Nota de Crédito: estado diferente de EstadoFFCompraNotaCredito",
					nombreMetodo,
					Funcion.slogtipoinformacion
					);
				return;
			}

			// =====================================================
			// 4) EJECUTAR ACTUALIZACIÓN (IGUAL QUE ANTES)
			// =====================================================
			FamilyFriendsService service =
				new FamilyFriendsService(frmRequerimiento.sconexionPoint);

			Guid idResultado = service.ActualizarNotaCreditoFF(idCompra, idNotaCredito);

			// =====================================================
			// 5) LOG FINAL
			// =====================================================
			Funcion.LogSistema(frmRequerimiento.sconexionInicio, nombreClase + ": " + nombreMetodo,
				string.Format(
				"ActualizarNotaCreditoFF completado | idCompra: {0}, idNotaCredito: {1}, idWEB_FF_Compra retornado: {2}",
				idCompra,
				idNotaCredito,
				idResultado
				),
				nombreMetodo,
				Funcion.slogtipoinformacion
				);

			Console.WriteLine("idWEB_FF_Compra retornado: " + idResultado);
		}


		int idBloqueaTransacciones = 0;

		private void uGridFacturas_ClickCellButton(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{			
			if (e.Cell.Column.ToString() == "btnAutorizar")
			{
				int iEstado = Funcion.iEscalarSQL(cdsSeteoF, string.Format("Select ISNULL(EstadoSolNC,0) From Compra Where idCompra = {0}", (int)e.Cell.Row.Cells["idCompra"].Value));

				if (iEstado == 4)
				{
					MessageBox.Show("Ya se creo la nota de credito solicitada para este documento", "Point Technology");
					this.Consulta();
					return;
				}

				#region Validacion Venta de Cartera
				int iVentaCartera = Funcion.iEscalarSQL(cdsSeteoF, string.Format("SELECT COUNT(*)FROM Compra WHERE idCompra = {0} AND idFormaPago = 9 AND ISNULL(NumeroLote,0) > 0 AND ISNULL(Confirmado,0) = 0",(int)e.Cell.Row.Cells["idCompra"].Value));
				if (iVentaCartera > 0)
				{
					MessageBox.Show("LA FACTURA ESTA EN TRANSITO PARA VENTA DE CARTERA\n\nCOMUNIQUESE CON EL AREA DE CARTERA", "Point Technology",MessageBoxButtons.OK, MessageBoxIcon.Stop );
					return;
				}
				#endregion Validacion Venta de Cartera

				if ((int)this.cmbMotivos.Value == 30)
				{
					int iContPagos = Funcion.iEscalarSQL(cdsSeteoF, string.Format("Select COUNT(*) From Cre_TablaDeAmortizacion Where Estado = 2 And idCompra = {0}", (int)e.Cell.Row.Cells["idCompra"].Value));
					if (iContPagos > 0)
					{
						MessageBox.Show(string.Format("La factura tiene {0} pagos registrados en la tabla de amortizacion.\n\nComuniquese con Cobranzas", iContPagos), "Point Technology");
						return;
					}
				}
				
				if (Funcion.ValidaBloqueo((int)e.Cell.Row.Cells["idCompra"].Value, 1, cdsSeteoF)) return;
				
				if ((int)e.Cell.Row.Cells["EstadoSolNC"].Value == 3) return;
				 
				if ((int)e.Cell.Row.Cells["EstadoSolNC"].Value == 4) return;

				string sSQLBloqueo = string.Format("Exec RegistraBloqueoTransacciones 0, 1, {0}, 1", (int)e.Cell.Row.Cells["idCompra"].Value);
				int idBloqueaTransacciones = Funcion.iEscalarSQL(cdsSeteoF, sSQLBloqueo, true);

				DateTime dtHoy = FuncionesProcedimientos.dtRetornaFechaServidor(cdsSeteoF);		

				int iDias = Funcion.iEscalarSQL(cdsSeteoF, 
					string.Format("Select DATEDIFF(day, '{0}', '{1}') As DiasEstancia", 
					Convert.ToDateTime(e.Cell.Row.Cells["Fecha"].Value).ToString("yyyyMMdd"), Convert.ToDateTime(dtHoy).ToString("yyyyMMdd")));		

				#region Retiro de mercaderia
				if ((int)this.cmbMotivos.Value == 4)
				{
					#region Viene de Solicitud De Cobranzas
					if ((int)e.Cell.Row.Cells["EstadoSolNC"].Value == 0) 
					{
						using (frmJustAnulaVenta Nota = new frmJustAnulaVenta(20))
						{
							if (DialogResult.OK == Nota.ShowDialog())
							{
								Funcion.EjecutaSQL(cdsSeteoF, string.Format("Exec SolicitudNotasDeCreditoEstados {0}, {1}, '{2}', {3}", 
									1, (int)this.cmbMotivos.Value, Nota.txtNotas.Text.ToString(), (int)e.Cell.Row.Cells["idCompra"].Value));
							}
						}

						if (idBloqueaTransacciones > 0)
						{
							Funcion.EjecutaSQL(cdsSeteoF, string.Format("Exec RegistraBloqueoTransacciones {0}, 0, {1}, 1", 
								idBloqueaTransacciones, (int)e.Cell.Row.Cells["idCompra"].Value));	
						}
						return;
					}
					#endregion Viene de Solicitud De Cobranzas

					#region Crea Nota De Credito
					if ((int)e.Cell.Row.Cells["EstadoSolNC"].Value == 1)
					{						
						using (frmNCListaArticulos LA = new frmNCListaArticulos((int)e.Cell.Row.Cells["idCompra"].Value, e.Cell.Row.Cells["Notas"].Value.ToString(), (int)this.cmbMotivos.Value, Convert.ToDecimal(e.Cell.Row.Cells["Total"].Value)))
						{
							if (DialogResult.OK == LA.ShowDialog())
							{	
								#region Solicitud rechazada
								if ((int)LA.optEstado.Value == 3)
								{
									Funcion.EjecutaSQL(cdsSeteoF, string.Format("Exec SolicitudNotasDeCreditoEstados {0}, {1}, '{2}', {3}", 
										(int)LA.optEstado.Value, (int)this.cmbMotivos.Value, LA.txtNotasDT.Text.ToString(), (int)e.Cell.Row.Cells["idCompra"].Value));

									if (idBloqueaTransacciones > 0)
										Funcion.EjecutaSQL(cdsSeteoF, string.Format("Exec RegistraBloqueoTransacciones {0}, 0, {1}, 1", 
											idBloqueaTransacciones, (int)e.Cell.Row.Cells["idCompra"].Value));	
									
									Consulta();
									
									return;
								}
								#endregion Solicitud rechazada

								decimal dValorND = 0.00m;
								dValorND = (decimal)LA.txtValor.Value - (decimal)LA.dTotal;

								string sSQLNDR = string.Format("Exec CreaNotaDeDebitoRetirosMercaderia {0}, {1}",
									(int)e.Cell.Row.Cells["idCompra"].Value, dValorND);	
								Funcion.EjecutaSQL(cdsSeteoF, sSQLNDR);

								string sNumero = Funcion.sEscalarSQL(cdsSeteoF, "Exec NumeracionLocales 5, 1, 1");

								string sSQLNCR = string.Format("Exec CreaNotaDeCreditoRetiroMercaderia {0}, '{1}', '{2}', {3}, {4}",
									(int)e.Cell.Row.Cells["idCompra"].Value, sNumero, LA.txtNotasDT.Text.ToString(), (int)LA.cmbBodega.Value, (decimal)LA.txtValor.Value);	
								int idCompraNC = Funcion.iEscalarSQL(cdsSeteoF, sSQLNCR);

								string sClaveAcceso = Funcion.sEscalarSQL(cdsSeteoF, string.Format("Exec FE_ClaveDeAcceso {0}, 5", idCompraNC));

								string cadInvertida = Facturacion.invertirCadena(sClaveAcceso.Substring(0, 48));
								string ValM11 = Facturacion.Modulo11(cadInvertida).ToString().Trim();
								string sClaveSRI = sClaveAcceso.ToString().Substring(0, 48) + ValM11;

								Funcion.EjecutaSQL(cdsSeteoF, string.Format("Update Compra Set AutFactura = '{0}' Where idCompra = {1}", sClaveSRI, idCompraNC));														

								#region Envio a remates
								string sSQLDT = "";
								foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in LA.ultraGrid1.Rows.All)
								{			
									int idSerial = 0;

									if (dr.Cells["idSerial"].Value != null && dr.Cells["idSerial"].Value.ToString() != "")
									{
										idSerial = (int)dr.Cells["idSerial"].Value;
									}

									string serial = dr.Cells["Serial"].Value != null ? dr.Cells["Serial"].Value.ToString() : "";
									string notas = LA.txtNotasDT.Text != null ? LA.txtNotasDT.Text.ToString() : "";

									if ((bool)dr.Cells["Sel"].Value)
									{	
									
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

												sSQLDT = string.Format(
													"Exec DTGuardaArticuloRematesParaTransferencia {0}, {1}, {2}, '{3}', {4}, '{5}', {6}, {7}",
													(int)e.Cell.Row.Cells["idCompra"].Value,
													(int)dr.Cells["idArticulo"].Value,
													idSerial,
													serial, 
													0,
													notas,   
													Convert.ToDecimal(dr.Cells["Valores"].Value),
													idCompraNC
													);
												
												oCmdActualiza.CommandText = sSQLDT;
												oCmdActualiza.ExecuteNonQuery	();

												oTransaction.Commit();
												oConexion.Close();												
												Cursor = Cursors.Default;															
											}
											catch(Exception ex)
											{
												oTransaction.Rollback();													
												oConexion.Close();
												MessageBox.Show(string.Format("Commit Exception Type : {0} {1}", ex.GetType(), ex.Message), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);						
											}
											finally
											{
												oConexion.Close();
											}
										}										
									}									
								}
								#endregion Envio a remates
																			
								Funcion.EjecutaSQL(cdsSeteoF, string.Format("Exec SolicitudNotasDeCreditoEstados {0}, {1}, '{2}', {3}", 
									(int)LA.optEstado.Value, (int)this.cmbMotivos.Value, LA.txtNotasDT.Text.ToString(), (int)e.Cell.Row.Cells["idCompra"].Value));								
                								
								if (idBloqueaTransacciones > 0)
									Funcion.EjecutaSQL(cdsSeteoF, string.Format("Exec RegistraBloqueoTransacciones {0}, 0, {1}, 1", 
										idBloqueaTransacciones, (int)e.Cell.Row.Cells["idCompra"].Value));

								Consulta();

								Facturacion.Impresion(idCompraNC, false, 5, true, true, false, false, 5, 0.00m, 0, 0, "", DateTime.Today, 0, "", cdsSeteoF);
							}
							else 
							{
								if (idBloqueaTransacciones > 0)
								{
									Funcion.EjecutaSQL(cdsSeteoF, string.Format("Exec RegistraBloqueoTransacciones {0}, 0, {1}, 1", 
										idBloqueaTransacciones, (int)e.Cell.Row.Cells["idCompra"].Value));	
								}
							}
						}					
					}
					#endregion Crea Nota De Credito

					return;
				}
				#endregion Retiro de mercaderia

				#region Anula Facturas Del Dia
				if ((int)this.cmbMotivos.Value == 5)
				{
					if(iDias == 0) //************
					{
						if (miAcceso.BSolNCAFD && !miAcceso.BAnulaFacturaMayor0) 
						{
							//if ((int) MenuLatinium.IdUsuario == 461)
							//{
							this.NotaDeCredito(sender, e);
							
							Funcion.EjecutaSQL(cdsSeteoF, string.Format("Exec RegistraBloqueoTransacciones {0}, 0, {1}, 1", 
									idBloqueaTransacciones, (int)e.Cell.Row.Cells["idCompra"].Value));	
							
							Consulta();
							
							return;							
							//}
						}
					}

					if (miAcceso.BAnulaFacturaMayor0) 
					{
						this.NotaDeCredito(sender, e);
						
						Funcion.EjecutaSQL(cdsSeteoF, string.Format("Exec RegistraBloqueoTransacciones {0}, 0, {1}, 1", 
								idBloqueaTransacciones, (int)e.Cell.Row.Cells["idCompra"].Value));	
						
						Consulta();
							
						return;
					}
				}			
				#endregion Anula Facturas Del Dia
				      
				#region Estado Pendiente
				if ((int)e.Cell.Row.Cells["EstadoSolNC"].Value == 0)
				{
					#region Genera Nota de Credito El Mismo Dia
					if (iDias == 0)
					{						
						if ((int)this.cmbMotivos.Value == 45)
						{
							using(frmVenta Venta = new frmVenta((int)e.Cell.Row.Cells["idCompra"].Value, true, (int)this.cmbMotivos.Value, Convert.ToDecimal(e.Cell.Row.Cells["Total"].Value)))
							{
								if (DialogResult.OK == Venta.ShowDialog())
								{
									Funcion.EjecutaSQL(cdsSeteoF, string.Format("Exec SolicitudNotasDeCreditoEstados {0}, {1}, '{2}', {3}", 
										4, (int)this.cmbMotivos.Value, "NOTA DE CREDITO GENERADA", (int)e.Cell.Row.Cells["idCompra"].Value));									
								}
							}
						}

						if ((int)this.cmbMotivos.Value == 30)
							this.NotaDeCredito(sender, e);
						
						Funcion.EjecutaSQL(cdsSeteoF, string.Format("Exec RegistraBloqueoTransacciones {0}, 0, {1}, 1", 
								idBloqueaTransacciones, (int)e.Cell.Row.Cells["idCompra"].Value));	
						
						Consulta();
						
						return;
					}
					#endregion Genera Nota de Credito El Mismo Dia

					#region Solicita Nota De Credito dias anteriores
					if (iDias > 0)
					{
						using (frmJustAnulaVenta Nota = new frmJustAnulaVenta(20))
						{
							if (DialogResult.OK == Nota.ShowDialog())
							{
								Funcion.EjecutaSQL(cdsSeteoF, string.Format("Exec SolicitudNotasDeCreditoEstados {0}, {1}, '{2}', {3}", 
									1, (int)this.cmbMotivos.Value, Nota.txtNotas.Text.ToString(), (int)e.Cell.Row.Cells["idCompra"].Value));
							}
						}
						if (idBloqueaTransacciones > 0)
						{
							Funcion.EjecutaSQL(cdsSeteoF, string.Format("Exec RegistraBloqueoTransacciones {0}, 0, {1}, 1", 
								idBloqueaTransacciones, (int)e.Cell.Row.Cells["idCompra"].Value));	
						}
						Consulta();
						
						return;
					}
					#endregion Solicita Nota De Credito dias anteriores
				
					return;
				}
				#endregion Estado Pendiente

				#region Estado Pendiente de Aprobacion
				if ((int)e.Cell.Row.Cells["EstadoSolNC"].Value == 1)
				{
					#region Aprobacion Solicitud Hasta 3 dias de la factura
					#region Tarjeta de credito debito efectivo
					if (iDias > 0 && iDias <= 3 && (int)this.cmbMotivos.Value == 45)
					{
						if (!miAcceso.B30DiasCobro)
						{
							MessageBox.Show("No Tiene Acceso a Aprobar Notas De Credito, solicite ayuda a contabilidad", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);

							Funcion.EjecutaSQL(cdsSeteoF, string.Format("Exec RegistraBloqueoTransacciones {0}, 0, {1}, 1", 
								idBloqueaTransacciones, (int)e.Cell.Row.Cells["idCompra"].Value));

							return;
						}

						using (frmNCAR Nota = new frmNCAR())
						{
							if (DialogResult.OK == Nota.ShowDialog())
							{	
								Funcion.EjecutaSQL(cdsSeteoF, string.Format("Exec SolicitudNotasDeCreditoEstados {0}, {1}, '{2}', {3}", 
									(int)Nota.optEstado.Value, (int)this.cmbMotivos.Value, Nota.txtNotas.Text.ToString(), (int)e.Cell.Row.Cells["idCompra"].Value));							
							}
						}

						Funcion.EjecutaSQL(cdsSeteoF, string.Format("Exec RegistraBloqueoTransacciones {0}, 0, {1}, 1", 
								idBloqueaTransacciones, (int)e.Cell.Row.Cells["idCompra"].Value));
	
						Consulta();

						return;
					}
					#endregion Tarjeta de credito debito efectivo

					#region CrediPoint
					if (iDias > 0 && iDias <= 5 && (int)this.cmbMotivos.Value == 30)
					{
						if (!miAcceso.B30DiasCobro)
						{
							MessageBox.Show("No Tiene Acceso a Aprobar Notas De Credito, solicite ayuda a contabilidad", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);

							Funcion.EjecutaSQL(cdsSeteoF, string.Format("Exec RegistraBloqueoTransacciones {0}, 0, {1}, 1", 
								idBloqueaTransacciones, (int)e.Cell.Row.Cells["idCompra"].Value));

							return;
						}

						using (frmNCAR Nota = new frmNCAR())
						{
							if (DialogResult.OK == Nota.ShowDialog())
							{	
								Funcion.EjecutaSQL(cdsSeteoF, string.Format("Exec SolicitudNotasDeCreditoEstados {0}, {1}, '{2}', {3}", 
									(int)Nota.optEstado.Value, (int)this.cmbMotivos.Value, Nota.txtNotas.Text.ToString(), (int)e.Cell.Row.Cells["idCompra"].Value));								
							}							
						}
	
						Funcion.EjecutaSQL(cdsSeteoF, string.Format("Exec RegistraBloqueoTransacciones {0}, 0, {1}, 1", 
							idBloqueaTransacciones, (int)e.Cell.Row.Cells["idCompra"].Value));
	
						Consulta();

						return;
					}
					#endregion CrediPoint
					#endregion Aprobacion Solicitud Hasta 5 dias de la factura

					#region Aprobacion Solicitud mayor a 5 dias de la factura
					if ((iDias > 3 && (int)this.cmbMotivos.Value == 45) || (iDias > 5 && (int)this.cmbMotivos.Value == 30))
					{
						if (!miAcceso.B60DiasCobro)
						{
							MessageBox.Show("No Tiene Acceso a Aprobar Notas De Credito, solicite ayuda a contabilidad", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop); 
							
							Funcion.EjecutaSQL(cdsSeteoF, string.Format("Exec RegistraBloqueoTransacciones {0}, 0, {1}, 1", 
								idBloqueaTransacciones, (int)e.Cell.Row.Cells["idCompra"].Value));

							return;
						}

						using (frmNCAR LA = new frmNCAR())
						{
							if (DialogResult.OK == LA.ShowDialog())
							{	
								Funcion.EjecutaSQL(cdsSeteoF, string.Format("Exec SolicitudNotasDeCreditoEstados {0}, {1}, '{2}', {3}", 
									(int)LA.optEstado.Value, (int)this.cmbMotivos.Value, LA.txtNotas.Text.ToString(), (int)e.Cell.Row.Cells["idCompra"].Value));
							}
						}

						Funcion.EjecutaSQL(cdsSeteoF, string.Format("Exec RegistraBloqueoTransacciones {0}, 0, {1}, 1", 
								idBloqueaTransacciones, (int)e.Cell.Row.Cells["idCompra"].Value));
						
						Consulta();
						
						return;
					}
					#endregion Aprobacion Solicitud mayor a 5 dias de la factura												
				}
				#endregion Estado Pendiente de Aprobacion

				#region Genera Notas De Credito Aprobadas
				if ((int)e.Cell.Row.Cells["EstadoSolNC"].Value == 2 && miAcceso.BProcesar)
				{	
					#region Por cambio de producto
					if ((int)this.cmbMotivos.Value == 45)
					{
						bool bRemates = Funcion.bEscalarSQL(cdsSeteoF, string.Format("Exec SolNCRetonaTipoBodega {0}", (int)e.Cell.Row.Cells["idCompra"].Value), true);

						if (!bRemates)
						{
							using(frmVenta Venta = new frmVenta((int)e.Cell.Row.Cells["idCompra"].Value, true, (int)this.cmbMotivos.Value, Convert.ToDecimal(e.Cell.Row.Cells["Total"].Value)))
							{
								if (DialogResult.OK == Venta.ShowDialog())
								{
									Funcion.EjecutaSQL(cdsSeteoF, string.Format("Exec SolicitudNotasDeCreditoEstados {0}, {1}, '{2}', {3}", 
										4, (int)this.cmbMotivos.Value, "NOTA DE CREDITO GENERADA", (int)e.Cell.Row.Cells["idCompra"].Value));	
								}
							}
						}
						else
						{
							using(frmVentaRemates Venta = new frmVentaRemates((int)e.Cell.Row.Cells["idCompra"].Value, true, (int)this.cmbMotivos.Value, Convert.ToDecimal(e.Cell.Row.Cells["Total"].Value)))
							{
								if (DialogResult.OK == Venta.ShowDialog())
								{
									Funcion.EjecutaSQL(cdsSeteoF, string.Format("Exec SolicitudNotasDeCreditoEstados {0}, {1}, '{2}', {3}", 
										4, (int)this.cmbMotivos.Value, "NOTA DE CREDITO GENERADA", (int)e.Cell.Row.Cells["idCompra"].Value));	
								}
							}
						}
					}
					#endregion Por cambio de producto

					#region Por anulacion de credito
					if ((int)this.cmbMotivos.Value == 30)
						this.NotaDeCredito(sender, e);
					#endregion Por anulacion de credito

					Funcion.EjecutaSQL(cdsSeteoF, string.Format("Exec RegistraBloqueoTransacciones {0}, 0, {1}, 1", 
							idBloqueaTransacciones, (int)e.Cell.Row.Cells["idCompra"].Value));	
					
					Consulta();

					return;
				}
				#endregion Genera Notas De Credito Aprobadas

				Funcion.EjecutaSQL(cdsSeteoF, string.Format("Exec RegistraBloqueoTransacciones {0}, 0, {1}, 1", 
					idBloqueaTransacciones, (int)e.Cell.Row.Cells["idCompra"].Value));	
			}
		}

		private void uGridFacturas_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}

		private void optEstados_ValueChanged(object sender, System.EventArgs e)
		{
			if (this.uGridFacturas.Rows.Count > 0)
			{
				bConsultas = true;
				Consulta();
				bConsultas = false;
			}
		}

		private void cmbBodega_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}

		private void cmbBodega_ValueChanged(object sender, System.EventArgs e)
		{
			if (this.uGridFacturas.Rows.Count > 0)
			{
				bConsultas = true;
				Consulta();
				bConsultas = false;
			}
		}

		private void cmbMotivos_ValueChanged(object sender, System.EventArgs e)
		{
			if (this.uGridFacturas.Rows.Count > 0)
			{
				bConsultas = true;
				Consulta();
				bConsultas = false;
			}
		}

		private void txtBusqueda_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) Consulta();
		}

		private void frmSolicitudNotasDeCredito_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Escape) this.Close();
		}

		private void btnExcel_Click(object sender, System.EventArgs e)
		{
			FuncionesProcedimientos.ExportaExcel(this.uGridFacturas);
		}
	}
}
