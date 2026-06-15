using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using DevExpress.XtraEditors;

namespace Latinium
{
	/// <summary>
	/// Summary description for frmConfiguraRetencionesIVA.
	/// </summary>
	public class frmConfiguraRetencionesIVA : DevExpress.XtraEditors.XtraForm
	{
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource1;
		private Infragistics.Win.UltraWinGrid.UltraDropDown cmbSujetoRetencion;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource2;
		public Infragistics.Win.UltraWinGrid.UltraGrid ultraGrid1;
		private C1.Data.C1DataSet cdsSeteoF;
		private System.Windows.Forms.Button btnNuevo;
		private Infragistics.Win.UltraWinGrid.UltraDropDown cmbAgenteRetencion;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource3;
		private Infragistics.Win.UltraWinEditors.UltraOptionSet optAgente;
		/// <summary>
		/// Required designer variable.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public frmConfiguraRetencionesIVA()
		{
			//
			// Required for Windows Form Designer support
			//
			InitializeComponent();

			//
			// TODO: Add any constructor code after InitializeComponent call
			//
		}

		/// <summary>
		/// Clean up any resources being used.
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

		#region Windows Form Designer generated code
		/// <summary>
		/// Required method for Designer support - do not modify
		/// the contents of this method with the code editor.
		/// </summary>
		private void InitializeComponent()
		{
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idConfRetencionesIVA");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idSujetoRetencion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn3 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idAgenteRetencion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn4 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Bienes");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn5 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Bienes100");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn6 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Servicios");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn7 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Servicios100");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn8 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Servicios100Prof");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn9 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Servicios100Arrend");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn10 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("PorcIVA");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn11 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("PorcServicios");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn12 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Desde");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn13 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Hasta");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idSujetoRetencion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn14 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idSujetoRetencion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn15 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.Appearance appearance1 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idConfRetencionesIVA");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idSujetoRetencion", -1, "cmbSujetoRetencion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idAgenteRetencion", -1, "cmbAgenteRetencion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bienes");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bienes100");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Servicios");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Servicios100");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Servicios100Prof");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn11 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Servicios100Arrend");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn12 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("PorcIVA");
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn13 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("PorcServicios");
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn14 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Desde");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn15 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Hasta");
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmConfiguraRetencionesIVA));
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand3 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn16 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idAgenteRetencion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn17 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn16 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idAgenteRetencion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn17 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance11 = new Infragistics.Win.Appearance();
			Infragistics.Win.ValueListItem valueListItem1 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem2 = new Infragistics.Win.ValueListItem();
			this.ultraDataSource1 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.cmbSujetoRetencion = new Infragistics.Win.UltraWinGrid.UltraDropDown();
			this.ultraDataSource2 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.ultraGrid1 = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.btnNuevo = new System.Windows.Forms.Button();
			this.cmbAgenteRetencion = new Infragistics.Win.UltraWinGrid.UltraDropDown();
			this.ultraDataSource3 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.optAgente = new Infragistics.Win.UltraWinEditors.UltraOptionSet();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbSujetoRetencion)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbAgenteRetencion)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource3)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.optAgente)).BeginInit();
			this.SuspendLayout();
			// 
			// ultraDataSource1
			// 
			ultraDataColumn1.DataType = typeof(int);
			ultraDataColumn2.DataType = typeof(int);
			ultraDataColumn3.DataType = typeof(int);
			ultraDataColumn4.DataType = typeof(bool);
			ultraDataColumn4.DefaultValue = false;
			ultraDataColumn5.DataType = typeof(bool);
			ultraDataColumn5.DefaultValue = false;
			ultraDataColumn6.DataType = typeof(bool);
			ultraDataColumn6.DefaultValue = false;
			ultraDataColumn7.DataType = typeof(bool);
			ultraDataColumn7.DefaultValue = false;
			ultraDataColumn8.DataType = typeof(bool);
			ultraDataColumn8.DefaultValue = false;
			ultraDataColumn9.DataType = typeof(bool);
			ultraDataColumn9.DefaultValue = false;
			ultraDataColumn10.DataType = typeof(System.Decimal);
			ultraDataColumn10.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
			ultraDataColumn11.DataType = typeof(System.Decimal);
			ultraDataColumn11.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
			ultraDataColumn12.DataType = typeof(System.DateTime);
			ultraDataColumn13.DataType = typeof(System.DateTime);
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
																																 ultraDataColumn13});
			// 
			// cmbSujetoRetencion
			// 
			this.cmbSujetoRetencion.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbSujetoRetencion.DataSource = this.ultraDataSource2;
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn1.Width = 300;
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridColumn2.Width = 300;
			ultraGridBand1.Columns.AddRange(new object[] {
																										 ultraGridColumn1,
																										 ultraGridColumn2});
			this.cmbSujetoRetencion.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.cmbSujetoRetencion.DisplayMember = "Nombre";
			this.cmbSujetoRetencion.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbSujetoRetencion.Location = new System.Drawing.Point(136, 208);
			this.cmbSujetoRetencion.Name = "cmbSujetoRetencion";
			this.cmbSujetoRetencion.Size = new System.Drawing.Size(256, 112);
			this.cmbSujetoRetencion.TabIndex = 338;
			this.cmbSujetoRetencion.ValueMember = "idSujetoRetencion";
			this.cmbSujetoRetencion.Visible = false;
			// 
			// ultraDataSource2
			// 
			ultraDataColumn14.DataType = typeof(int);
			this.ultraDataSource2.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn14,
																																 ultraDataColumn15});
			// 
			// ultraGrid1
			// 
			this.ultraGrid1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.ultraGrid1.Cursor = System.Windows.Forms.Cursors.Default;
			this.ultraGrid1.DataSource = this.ultraDataSource1;
			appearance1.BackColor = System.Drawing.Color.White;
			this.ultraGrid1.DisplayLayout.Appearance = appearance1;
			this.ultraGrid1.DisplayLayout.AutoFitColumns = true;
			ultraGridColumn3.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn3.Header.VisiblePosition = 0;
			ultraGridColumn3.Hidden = true;
			ultraGridColumn3.Width = 137;
			ultraGridColumn4.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn4.Header.Caption = "Sujeto De Retención";
			ultraGridColumn4.Header.VisiblePosition = 1;
			ultraGridColumn4.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownList;
			ultraGridColumn4.Width = 160;
			ultraGridColumn5.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn5.Header.Caption = "Agente De Retencion";
			ultraGridColumn5.Header.VisiblePosition = 2;
			ultraGridColumn5.Hidden = true;
			ultraGridColumn5.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownList;
			ultraGridColumn5.Width = 180;
			ultraGridColumn6.Header.VisiblePosition = 3;
			ultraGridColumn6.Width = 77;
			ultraGridColumn7.Header.Caption = "B-100";
			ultraGridColumn7.Header.VisiblePosition = 4;
			ultraGridColumn7.Width = 77;
			ultraGridColumn8.Header.VisiblePosition = 5;
			ultraGridColumn8.Width = 77;
			ultraGridColumn9.Header.Caption = "S-100";
			ultraGridColumn9.Header.VisiblePosition = 6;
			ultraGridColumn9.Width = 77;
			ultraGridColumn10.Header.Caption = "S-100-P";
			ultraGridColumn10.Header.VisiblePosition = 7;
			ultraGridColumn10.Width = 77;
			ultraGridColumn11.Header.Caption = "S-100-A";
			ultraGridColumn11.Header.VisiblePosition = 8;
			ultraGridColumn11.Width = 77;
			appearance2.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn12.CellAppearance = appearance2;
			ultraGridColumn12.Format = "#,##0.00";
			ultraGridColumn12.Header.Caption = "% IVA";
			ultraGridColumn12.Header.VisiblePosition = 9;
			ultraGridColumn12.Width = 87;
			appearance3.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn13.CellAppearance = appearance3;
			ultraGridColumn13.Format = "#,##0.00";
			ultraGridColumn13.Header.Caption = "% Servicios";
			ultraGridColumn13.Header.VisiblePosition = 10;
			ultraGridColumn13.Width = 87;
			ultraGridColumn14.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn14.Header.VisiblePosition = 11;
			ultraGridColumn14.Width = 77;
			ultraGridColumn15.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn15.Header.VisiblePosition = 12;
			ultraGridColumn15.Width = 77;
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
																										 ultraGridColumn12,
																										 ultraGridColumn13,
																										 ultraGridColumn14,
																										 ultraGridColumn15});
			this.ultraGrid1.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			appearance4.ForeColor = System.Drawing.Color.Black;
			appearance4.ForeColorDisabled = System.Drawing.Color.Black;
			this.ultraGrid1.DisplayLayout.Override.ActiveRowAppearance = appearance4;
			this.ultraGrid1.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
			this.ultraGrid1.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.ultraGrid1.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance5.BackColor = System.Drawing.Color.Transparent;
			this.ultraGrid1.DisplayLayout.Override.CardAreaAppearance = appearance5;
			appearance6.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance6.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance6.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance6.FontData.BoldAsString = "True";
			appearance6.FontData.Name = "Arial";
			appearance6.FontData.SizeInPoints = 8.5F;
			appearance6.ForeColor = System.Drawing.Color.White;
			appearance6.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.ultraGrid1.DisplayLayout.Override.HeaderAppearance = appearance6;
			appearance7.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance7.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance7.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.RowAlternateAppearance = appearance7;
			appearance8.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance8.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance8.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.RowSelectorAppearance = appearance8;
			appearance9.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance9.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance9.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.SelectedRowAppearance = appearance9;
			this.ultraGrid1.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.ultraGrid1.Location = new System.Drawing.Point(8, 40);
			this.ultraGrid1.Name = "ultraGrid1";
			this.ultraGrid1.Size = new System.Drawing.Size(971, 328);
			this.ultraGrid1.TabIndex = 339;
			this.ultraGrid1.CellChange += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.ultraGrid1_CellChange);
			this.ultraGrid1.AfterCellUpdate += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.ultraGrid1_AfterCellUpdate);
			this.ultraGrid1.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.ultraGrid1_InitializeLayout);
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
			// btnNuevo
			// 
			this.btnNuevo.CausesValidation = false;
			this.btnNuevo.Image = ((System.Drawing.Image)(resources.GetObject("btnNuevo.Image")));
			this.btnNuevo.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnNuevo.Location = new System.Drawing.Point(8, 8);
			this.btnNuevo.Name = "btnNuevo";
			this.btnNuevo.Size = new System.Drawing.Size(72, 23);
			this.btnNuevo.TabIndex = 340;
			this.btnNuevo.Text = "Nuevo";
			this.btnNuevo.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnNuevo.Click += new System.EventHandler(this.btnNuevo_Click);
			// 
			// cmbAgenteRetencion
			// 
			this.cmbAgenteRetencion.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbAgenteRetencion.DataSource = this.ultraDataSource3;
			ultraGridColumn16.Header.VisiblePosition = 0;
			ultraGridColumn16.Hidden = true;
			ultraGridColumn17.Header.VisiblePosition = 1;
			ultraGridColumn17.Width = 300;
			ultraGridBand3.Columns.AddRange(new object[] {
																										 ultraGridColumn16,
																										 ultraGridColumn17});
			this.cmbAgenteRetencion.DisplayLayout.BandsSerializer.Add(ultraGridBand3);
			this.cmbAgenteRetencion.DisplayMember = "Nombre";
			this.cmbAgenteRetencion.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbAgenteRetencion.Location = new System.Drawing.Point(416, 208);
			this.cmbAgenteRetencion.Name = "cmbAgenteRetencion";
			this.cmbAgenteRetencion.Size = new System.Drawing.Size(256, 112);
			this.cmbAgenteRetencion.TabIndex = 341;
			this.cmbAgenteRetencion.ValueMember = "idAgenteRetencion";
			this.cmbAgenteRetencion.Visible = false;
			// 
			// ultraDataSource3
			// 
			ultraDataColumn16.DataType = typeof(int);
			this.ultraDataSource3.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn16,
																																 ultraDataColumn17});
			// 
			// optAgente
			// 
			appearance10.ForeColorDisabled = System.Drawing.Color.Black;
			this.optAgente.Appearance = appearance10;
			this.optAgente.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.optAgente.CheckedIndex = 0;
			this.optAgente.ItemAppearance = appearance11;
			this.optAgente.ItemOrigin = new System.Drawing.Point(10, 0);
			valueListItem1.DataValue = 0;
			valueListItem1.DisplayText = "Agente";
			valueListItem2.DataValue = 1;
			valueListItem2.DisplayText = "Sujeto";
			this.optAgente.Items.Add(valueListItem1);
			this.optAgente.Items.Add(valueListItem2);
			this.optAgente.ItemSpacingVertical = 10;
			this.optAgente.Location = new System.Drawing.Point(104, 8);
			this.optAgente.Name = "optAgente";
			this.optAgente.Size = new System.Drawing.Size(136, 26);
			this.optAgente.TabIndex = 342;
			this.optAgente.Text = "Agente";
			this.optAgente.ValueChanged += new System.EventHandler(this.optAgente_ValueChanged);
			// 
			// frmConfiguraRetencionesIVA
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 14);
			this.ClientSize = new System.Drawing.Size(987, 374);
			this.Controls.Add(this.optAgente);
			this.Controls.Add(this.cmbAgenteRetencion);
			this.Controls.Add(this.cmbSujetoRetencion);
			this.Controls.Add(this.btnNuevo);
			this.Controls.Add(this.ultraGrid1);
			this.KeyPreview = true;
			this.Name = "frmConfiguraRetencionesIVA";
			this.Text = "Configura Retenciones I. V. A. ";
			this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.frmConfiguraRetencionesIVA_KeyDown);
			this.Load += new System.EventHandler(this.frmConfiguraRetencionesIVA_Load);
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbSujetoRetencion)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbAgenteRetencion)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource3)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.optAgente)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);		
		}

		private void Consulta()
		{
			string sSQL = string.Format("Exec ConsultaConfRetencionesIVA {0}", (int)this.optAgente.Value);
			this.ultraGrid1.DataSource = Funcion.dvProcedimiento(cdsSeteoF, sSQL);

			if ((int)this.optAgente.Value == 0) 
			{
				this.ultraGrid1.DisplayLayout.Bands[0].Columns["idSujetoRetencion"].Hidden = false;
				this.ultraGrid1.DisplayLayout.Bands[0].Columns["idAgenteRetencion"].Hidden = true;
			}
			else
			{
				this.ultraGrid1.DisplayLayout.Bands[0].Columns["idSujetoRetencion"].Hidden = true;
				this.ultraGrid1.DisplayLayout.Bands[0].Columns["idAgenteRetencion"].Hidden = false;
			}
		}

		private void frmConfiguraRetencionesIVA_Load(object sender, System.EventArgs e)
		{
			this.cmbAgenteRetencion.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Select idAgenteRetencion, Nombre From AgenteRetencion Order By idAgenteRetencion");
			this.cmbSujetoRetencion.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Select idSujetoRetencion, Nombre From SujetoRetencion Order By idSujetoRetencion");

			Consulta();
		}

		private void ultraGrid1_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}

		private void ActualizaDatos(Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			DateTime dtDesde = (DateTime)e.Cell.Row.Cells["Desde"].Value;
			DateTime dtHasta = (DateTime)e.Cell.Row.Cells["Hasta"].Value;

			string sSQL = string.Format("Exec ActualizaConfiguracionRetencionIVA {0}, {1}, {2}, {3}, {4}, {5}, {6}, {7}, {8}, '{9}', '{10}'",
				(int)e.Cell.Row.Cells["idConfRetencionesIVA"].Value,
				(bool)e.Cell.Row.Cells["Bienes"].Value,
				(bool)e.Cell.Row.Cells["Bienes100"].Value,
				(bool)e.Cell.Row.Cells["Servicios"].Value, 
				(bool)e.Cell.Row.Cells["Servicios100"].Value, 
				(bool)e.Cell.Row.Cells["Servicios100Prof"].Value,
				(bool)e.Cell.Row.Cells["Servicios100Arrend"].Value,
				(decimal)e.Cell.Row.Cells["PorcIVA"].Value,
				(decimal)e.Cell.Row.Cells["PorcServicios"].Value,
				dtDesde.ToString("yyyyMMdd"), dtHasta.ToString("yyyyMMdd"));
			Funcion.EjecutaSQL(cdsSeteoF, sSQL, true);
		}

		private void ultraGrid1_CellChange(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			//c.Bienes, c.Bienes100, c.Servicios, c.Servicios100, c.Servicios100Prof, c.Servicios100Arrend, PorcIVA, PorcServicios, Desde, Hasta 
			
			#region Bienes
			if (e.Cell.Column.ToString() == "Bienes")
			{
				this.ultraGrid1.UpdateData();
				ActualizaDatos(e);
			}
			#endregion Bienes

			#region Bienes 100
			if (e.Cell.Column.ToString() == "Bienes100")
			{
				this.ultraGrid1.UpdateData();
				ActualizaDatos(e);
			}
			#endregion Bienes 100

			#region Servicios
			if (e.Cell.Column.ToString() == "Servicios")
			{
				this.ultraGrid1.UpdateData();
				ActualizaDatos(e);
			}
			#endregion Servicios

			#region Servicios 100
			if (e.Cell.Column.ToString() == "Servicios100")
			{
				this.ultraGrid1.UpdateData();
				ActualizaDatos(e);
			}
			#endregion Servicios 100

			#region Servicios 100 Profesionales
			if (e.Cell.Column.ToString() == "Servicios100Prof")
			{
				this.ultraGrid1.UpdateData();
				ActualizaDatos(e);
			}
			#endregion Servicios 100 Profesionales

			#region Servicios 100 Arrendamiento
			if (e.Cell.Column.ToString() == "Servicios100Arrend")
			{
				this.ultraGrid1.UpdateData();
				ActualizaDatos(e);
			}
			#endregion Servicios 100 Arrendamiento

			#region Porcentaje Bienes
			if (e.Cell.Column.ToString() == "PorcIVA")
			{
				this.ultraGrid1.UpdateData();
				ActualizaDatos(e);
			}
			#endregion Porcentaje Bienes

			#region Porcentaje Servicios
			if (e.Cell.Column.ToString() == "PorcServicios")
			{
				this.ultraGrid1.UpdateData();
				ActualizaDatos(e);
			}
			#endregion Porcentaje Servicios

			#region Desde
			if (e.Cell.Column.ToString() == "Desde")
			{
				this.ultraGrid1.UpdateData();
				ActualizaDatos(e);
			}
			#endregion Desde

			#region Hasta
			if (e.Cell.Column.ToString() == "Hasta")
			{
				this.ultraGrid1.UpdateData();
				ActualizaDatos(e);
			}
			#endregion Hasta
		}

		private void btnNuevo_Click(object sender, System.EventArgs e)
		{
			using (frmRegistraConfigRetenciones Nuevo = new frmRegistraConfigRetenciones())
			{
				if (DialogResult.OK == Nuevo.ShowDialog())
				{					
					string sSQL = string.Format("Exec GuardaConfigRetencionesIVA '{0}', '{1}', {2}, {3}",
						Convert.ToDecimal(Nuevo.dtDesde.Value).ToString("yyyyMMdd"), Convert.ToDecimal(Nuevo.dtHasta.Value).ToString("yyyyMMdd"), 
						Convert.ToDecimal(Nuevo.txtPorcBienes.Value), Convert.ToDecimal(Nuevo.txtPorcServicios.Value));
					Funcion.EjecutaSQL(cdsSeteoF, sSQL, true);

					Consulta();
				}
			}
		}

		private void optAgente_ValueChanged(object sender, System.EventArgs e)
		{
			Consulta();
		}

		private void frmConfiguraRetencionesIVA_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Escape) this.Close();
		}

		private void ultraGrid1_AfterCellUpdate(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			if (e.Cell.Column.ToString() == "PorcIVA")
			{
				string sSQL = string.Format("UPDATE ConfigRetencionesIva SET PorcIVA = {0} WHERE idConfRetencionesIVA = {1}", 
					(decimal)e.Cell.Row.Cells["PorcIVA"].Value, (int)e.Cell.Row.Cells["idConfRetencionesIVA"].Value);
				Funcion.EjecutaSQL(cdsSeteoF, sSQL);
			}

			if (e.Cell.Column.ToString() == "PorcServicios")
			{
				string sSQL = string.Format("UPDATE ConfigRetencionesIva SET PorcServicios = {0} WHERE idConfRetencionesIVA = {1}", 
					(decimal)e.Cell.Row.Cells["PorcServicios"].Value, (int)e.Cell.Row.Cells["idConfRetencionesIVA"].Value);
				Funcion.EjecutaSQL(cdsSeteoF, sSQL);
			}
		}
	}
}

