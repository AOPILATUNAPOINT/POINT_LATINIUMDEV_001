using System;
using System.Globalization;
using System.Data;
using System.Data.SqlClient;
using C1.Data;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using Infragistics.Win.UltraWinGrid;
using System.Web.Mail;
using CrystalDecisions.Shared;
using System.Runtime.InteropServices;
using System.Diagnostics;
using System.IO;
using CrystalDecisions;
using CrystalDecisions.CrystalReports.Engine;
using CrystalDecisions.ReportSource;
using System.Text;
using System.Net;
using System.Web;
using System.Xml;
using System.Xml.Schema;
using System.Security.Cryptography.Xml;
using System.Security.Cryptography;
using System.Security.Cryptography.X509Certificates;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de frmCobroPrendas.
	/// </summary>
	public class frmCobroPrendas : System.Windows.Forms.Form
	{
		private System.Windows.Forms.Label label12;
		private DevExpress.XtraEditors.SimpleButton btBuscar;
		private C1.Data.C1DataSet cdsSeteo;
		private System.Windows.Forms.TextBox txtCliente;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource1;
		private Infragistics.Win.UltraWinGrid.UltraGrid ugPrendas;
		private System.Windows.Forms.Label lblBodega;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbBodega;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbCajero;
		private System.Windows.Forms.Label label19;
		private System.Windows.Forms.Label label11;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbFormaPago;
		private DevExpress.XtraEditors.SimpleButton simpleButton1;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtFecha;
		private System.Windows.Forms.Label label1;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public frmCobroPrendas()
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
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idDetFormaPago");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCliente");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Numero");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Fecha");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Ruc");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("DescripcionPrenda");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Valor");
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Estado");
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCompra");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idDetFormaPago");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn3 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCliente");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn4 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Numero");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn5 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Fecha");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn6 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Ruc");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn7 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn8 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("DescripcionPrenda");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn9 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Valor");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn10 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Estado");
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmCobroPrendas));
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn11 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idBodega");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn12 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bodega");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn13 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand3 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn14 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPersonal");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn15 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn16 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand4 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn17 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idFormaPago");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn18 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FormaPago");
			Infragistics.Win.Appearance appearance11 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton1 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			this.label12 = new System.Windows.Forms.Label();
			this.txtCliente = new System.Windows.Forms.TextBox();
			this.ugPrendas = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource1 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.btBuscar = new DevExpress.XtraEditors.SimpleButton();
			this.cdsSeteo = new C1.Data.C1DataSet();
			this.lblBodega = new System.Windows.Forms.Label();
			this.cmbBodega = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.cmbCajero = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.label19 = new System.Windows.Forms.Label();
			this.label11 = new System.Windows.Forms.Label();
			this.cmbFormaPago = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.simpleButton1 = new DevExpress.XtraEditors.SimpleButton();
			this.dtFecha = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.label1 = new System.Windows.Forms.Label();
			((System.ComponentModel.ISupportInitialize)(this.ugPrendas)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbBodega)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCajero)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbFormaPago)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFecha)).BeginInit();
			this.SuspendLayout();
			// 
			// label12
			// 
			this.label12.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label12.Location = new System.Drawing.Point(160, 8);
			this.label12.Name = "label12";
			this.label12.Size = new System.Drawing.Size(88, 16);
			this.label12.TabIndex = 232;
			this.label12.Text = "# FACTURA";
			// 
			// txtCliente
			// 
			this.txtCliente.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtCliente.Location = new System.Drawing.Point(160, 25);
			this.txtCliente.Name = "txtCliente";
			this.txtCliente.Size = new System.Drawing.Size(96, 20);
			this.txtCliente.TabIndex = 233;
			this.txtCliente.Text = "";
			this.txtCliente.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
			this.txtCliente.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtCliente_KeyDown);
			this.txtCliente.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtCliente_KeyPress);
			// 
			// ugPrendas
			// 
			this.ugPrendas.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.ugPrendas.Cursor = System.Windows.Forms.Cursors.Default;
			this.ugPrendas.DataSource = this.ultraDataSource1;
			appearance1.BackColor = System.Drawing.Color.White;
			this.ugPrendas.DisplayLayout.Appearance = appearance1;
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridColumn2.Hidden = true;
			ultraGridColumn3.Header.VisiblePosition = 2;
			ultraGridColumn3.Hidden = true;
			ultraGridColumn4.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
			ultraGridColumn4.Header.VisiblePosition = 3;
			ultraGridColumn4.Width = 100;
			ultraGridColumn5.Header.VisiblePosition = 4;
			ultraGridColumn5.Width = 100;
			ultraGridColumn6.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
			ultraGridColumn6.Header.VisiblePosition = 5;
			ultraGridColumn6.Width = 100;
			ultraGridColumn7.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
			ultraGridColumn7.Header.VisiblePosition = 6;
			ultraGridColumn7.Width = 250;
			ultraGridColumn8.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
			ultraGridColumn8.Header.Caption = "Descripción";
			ultraGridColumn8.Header.VisiblePosition = 7;
			ultraGridColumn8.Width = 300;
			ultraGridColumn9.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
			appearance2.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn9.CellAppearance = appearance2;
			ultraGridColumn9.Format = "#,##0.00";
			ultraGridColumn9.Header.VisiblePosition = 8;
			ultraGridColumn9.Width = 70;
			ultraGridColumn10.Header.VisiblePosition = 9;
			ultraGridColumn10.Width = 118;
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
																										 ultraGridColumn10});
			this.ugPrendas.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			appearance3.ForeColor = System.Drawing.Color.Black;
			appearance3.ForeColorDisabled = System.Drawing.Color.Black;
			this.ugPrendas.DisplayLayout.Override.ActiveRowAppearance = appearance3;
			this.ugPrendas.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.TemplateOnBottom;
			this.ugPrendas.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.ugPrendas.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.False;
			appearance4.BackColor = System.Drawing.Color.Transparent;
			this.ugPrendas.DisplayLayout.Override.CardAreaAppearance = appearance4;
			appearance5.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance5.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance5.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance5.FontData.BoldAsString = "True";
			appearance5.FontData.Name = "Arial";
			appearance5.FontData.SizeInPoints = 10F;
			appearance5.ForeColor = System.Drawing.Color.White;
			appearance5.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.ugPrendas.DisplayLayout.Override.HeaderAppearance = appearance5;
			appearance6.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance6.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance6.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ugPrendas.DisplayLayout.Override.RowAlternateAppearance = appearance6;
			appearance7.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance7.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance7.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ugPrendas.DisplayLayout.Override.RowSelectorAppearance = appearance7;
			appearance8.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance8.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance8.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ugPrendas.DisplayLayout.Override.SelectedRowAppearance = appearance8;
			this.ugPrendas.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.ugPrendas.Location = new System.Drawing.Point(8, 48);
			this.ugPrendas.Name = "ugPrendas";
			this.ugPrendas.Size = new System.Drawing.Size(1064, 360);
			this.ugPrendas.TabIndex = 234;
			this.ugPrendas.DoubleClick += new System.EventHandler(this.ugPrendas_DoubleClick);
			this.ugPrendas.KeyDown += new System.Windows.Forms.KeyEventHandler(this.ugPrendas_KeyDown);
			// 
			// ultraDataSource1
			// 
			ultraDataColumn1.DataType = typeof(int);
			ultraDataColumn2.DataType = typeof(int);
			ultraDataColumn3.DataType = typeof(int);
			ultraDataColumn5.DataType = typeof(System.DateTime);
			ultraDataColumn9.DataType = typeof(System.Decimal);
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
			// btBuscar
			// 
			this.btBuscar.Image = ((System.Drawing.Image)(resources.GetObject("btBuscar.Image")));
			this.btBuscar.Location = new System.Drawing.Point(272, 8);
			this.btBuscar.Name = "btBuscar";
			this.btBuscar.Size = new System.Drawing.Size(80, 32);
			this.btBuscar.TabIndex = 237;
			this.btBuscar.Text = "&Buscar";
			this.btBuscar.Click += new System.EventHandler(this.btBuscar_Click);
			// 
			// cdsSeteo
			// 
			this.cdsSeteo.BindingContextCtrl = this;
			this.cdsSeteo.DataLibrary = "LibContabilidad";
			this.cdsSeteo.DataLibraryUrl = "";
			this.cdsSeteo.DataSetDef = "dsSeteo";
			this.cdsSeteo.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsSeteo.Name = "cdsSeteo";
			this.cdsSeteo.SchemaClassName = "LibContabilidad.DataClass";
			this.cdsSeteo.SchemaDef = null;
			// 
			// lblBodega
			// 
			this.lblBodega.AutoSize = true;
			this.lblBodega.Location = new System.Drawing.Point(8, 8);
			this.lblBodega.Name = "lblBodega";
			this.lblBodega.Size = new System.Drawing.Size(41, 16);
			this.lblBodega.TabIndex = 253;
			this.lblBodega.Text = "LOCAL";
			this.lblBodega.TextAlign = System.Drawing.ContentAlignment.TopRight;
			// 
			// cmbBodega
			// 
			this.cmbBodega.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbBodega.Cursor = System.Windows.Forms.Cursors.Default;
			ultraGridColumn11.Header.VisiblePosition = 0;
			ultraGridColumn11.Hidden = true;
			ultraGridColumn12.Header.VisiblePosition = 1;
			ultraGridColumn12.Hidden = true;
			ultraGridColumn13.Header.VisiblePosition = 2;
			ultraGridColumn13.Width = 144;
			ultraGridBand2.Columns.AddRange(new object[] {
																										 ultraGridColumn11,
																										 ultraGridColumn12,
																										 ultraGridColumn13});
			this.cmbBodega.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			this.cmbBodega.DisplayMember = "Nombre";
			this.cmbBodega.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbBodega.Enabled = false;
			this.cmbBodega.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbBodega.Location = new System.Drawing.Point(8, 24);
			this.cmbBodega.MaxDropDownItems = 30;
			this.cmbBodega.Name = "cmbBodega";
			this.cmbBodega.Size = new System.Drawing.Size(144, 21);
			this.cmbBodega.TabIndex = 252;
			this.cmbBodega.ValueMember = "idBodega";
			this.cmbBodega.ValueChanged += new System.EventHandler(this.cmbBodega_ValueChanged);
			// 
			// cmbCajero
			// 
			appearance9.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbCajero.Appearance = appearance9;
			this.cmbCajero.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbCajero.Cursor = System.Windows.Forms.Cursors.Default;
			ultraGridColumn14.Header.VisiblePosition = 0;
			ultraGridColumn14.Hidden = true;
			ultraGridColumn15.Header.VisiblePosition = 1;
			ultraGridColumn15.Width = 335;
			ultraGridColumn16.Header.VisiblePosition = 2;
			ultraGridBand3.Columns.AddRange(new object[] {
																										 ultraGridColumn14,
																										 ultraGridColumn15,
																										 ultraGridColumn16});
			this.cmbCajero.DisplayLayout.BandsSerializer.Add(ultraGridBand3);
			this.cmbCajero.DisplayMember = "Nombre";
			this.cmbCajero.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbCajero.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbCajero.Location = new System.Drawing.Point(476, 24);
			this.cmbCajero.Name = "cmbCajero";
			this.cmbCajero.Size = new System.Drawing.Size(200, 21);
			this.cmbCajero.TabIndex = 254;
			this.cmbCajero.ValueMember = "idPersonal";
			// 
			// label19
			// 
			this.label19.AutoSize = true;
			this.label19.Location = new System.Drawing.Point(476, 8);
			this.label19.Name = "label19";
			this.label19.Size = new System.Drawing.Size(50, 16);
			this.label19.TabIndex = 255;
			this.label19.Text = "CAJERO";
			// 
			// label11
			// 
			this.label11.Location = new System.Drawing.Point(680, 8);
			this.label11.Name = "label11";
			this.label11.Size = new System.Drawing.Size(88, 16);
			this.label11.TabIndex = 257;
			this.label11.Text = "Forma de Pago";
			// 
			// cmbFormaPago
			// 
			appearance10.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbFormaPago.Appearance = appearance10;
			this.cmbFormaPago.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbFormaPago.Cursor = System.Windows.Forms.Cursors.Default;
			ultraGridColumn17.Header.VisiblePosition = 0;
			ultraGridColumn17.Hidden = true;
			ultraGridColumn18.Header.VisiblePosition = 1;
			ultraGridColumn18.Width = 152;
			ultraGridBand4.Columns.AddRange(new object[] {
																										 ultraGridColumn17,
																										 ultraGridColumn18});
			this.cmbFormaPago.DisplayLayout.BandsSerializer.Add(ultraGridBand4);
			this.cmbFormaPago.DisplayMember = "FormaPago";
			this.cmbFormaPago.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbFormaPago.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbFormaPago.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbFormaPago.Location = new System.Drawing.Point(680, 24);
			this.cmbFormaPago.Name = "cmbFormaPago";
			this.cmbFormaPago.Size = new System.Drawing.Size(144, 21);
			this.cmbFormaPago.TabIndex = 256;
			this.cmbFormaPago.ValueMember = "idFormaPago";
			// 
			// simpleButton1
			// 
			this.simpleButton1.Enabled = false;
			this.simpleButton1.Image = ((System.Drawing.Image)(resources.GetObject("simpleButton1.Image")));
			this.simpleButton1.Location = new System.Drawing.Point(848, 8);
			this.simpleButton1.Name = "simpleButton1";
			this.simpleButton1.Size = new System.Drawing.Size(80, 32);
			this.simpleButton1.TabIndex = 258;
			this.simpleButton1.Text = "&Cobrar";
			this.simpleButton1.Visible = false;
			this.simpleButton1.Click += new System.EventHandler(this.simpleButton1_Click);
			// 
			// dtFecha
			// 
			appearance11.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtFecha.Appearance = appearance11;
			this.dtFecha.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton1.Caption = "Today";
			this.dtFecha.DateButtons.Add(dateButton1);
			this.dtFecha.Format = "dd/MM/yyyy";
			this.dtFecha.Location = new System.Drawing.Point(368, 24);
			this.dtFecha.Name = "dtFecha";
			this.dtFecha.NonAutoSizeHeight = 23;
			this.dtFecha.Size = new System.Drawing.Size(104, 21);
			this.dtFecha.SpinButtonsVisible = true;
			this.dtFecha.TabIndex = 259;
			this.dtFecha.Value = ((object)(resources.GetObject("dtFecha.Value")));
			// 
			// label1
			// 
			this.label1.AutoSize = true;
			this.label1.Location = new System.Drawing.Point(368, 8);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(36, 16);
			this.label1.TabIndex = 260;
			this.label1.Text = "Fecha";
			// 
			// frmCobroPrendas
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(1080, 413);
			this.Controls.Add(this.label1);
			this.Controls.Add(this.label19);
			this.Controls.Add(this.lblBodega);
			this.Controls.Add(this.txtCliente);
			this.Controls.Add(this.dtFecha);
			this.Controls.Add(this.simpleButton1);
			this.Controls.Add(this.label11);
			this.Controls.Add(this.cmbFormaPago);
			this.Controls.Add(this.cmbCajero);
			this.Controls.Add(this.cmbBodega);
			this.Controls.Add(this.btBuscar);
			this.Controls.Add(this.ugPrendas);
			this.Controls.Add(this.label12);
			this.Name = "frmCobroPrendas";
			this.Text = "Estado Prenda";
			this.Load += new System.EventHandler(this.frmCobroPrendas_Load);
			((System.ComponentModel.ISupportInitialize)(this.ugPrendas)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbBodega)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCajero)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbFormaPago)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFecha)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		int idClte = 0;
		string ruc = "";
		string nombr = "";
		int idDetFP = 0;
		string niUsuario = "";

		private void frmCobroPrendas_Load(object sender, System.EventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteo.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);

			string stExecBodega = string.Format("Select idBodega, Bodega, Nombre From Bodega Where Activo = 1");
			this.cmbBodega.DataSource = Funcion.dvProcedimiento(cdsSeteo, stExecBodega);

			int codBodHabtcn = Funcion.iEscalarSQL(cdsSeteo, "SELECT Bodega FROM UsuariosBodegas WHERE idUsuario="+MenuLatinium.IdUsuario.ToString());
			int BodaBusk = Funcion.iEscalarSQL(cdsSeteo, "SELECT idBodega FROM Bodega WHERE Bodega="+codBodHabtcn.ToString());
			this.cmbBodega.Value = BodaBusk;

			niUsuario = MenuLatinium.iNivel.ToString().Trim();
			if(niUsuario.Equals("1") || niUsuario.Equals("2") || niUsuario.Equals("3") || niUsuario.Equals("4") || niUsuario.Equals("15") || niUsuario.Equals("16")) 
			{
				this.cmbBodega.Enabled = true;
			}

			this.cmbFormaPago.DataSource = Funcion.dvProcedimiento(cdsSeteo, "Exec CompraFPagoCarga 1");
			this.dtFecha.Value = DateTime.Now;
		}

		private void btBuscar_Click(object sender, System.EventArgs e)
		{
			if(this.cmbBodega.ActiveRow == null)
			{
				MessageBox.Show("No se a seleccionado Guardia?","Aviso....!!!", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.cmbBodega.Focus();
			}
			else 
			{
				#region Busca FC
				string sSQL = string.Format("Exec Busk_Cliente {0}", this.cmbBodega.Value);
				ugPrendas.DataSource = Funcion.dvProcedimiento(cdsSeteo, sSQL);
				#endregion Busca FC
			}
		}

		private void txtCliente_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
		}

		private void txtCliente_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (e.KeyChar == 13) this.btBuscar_Click(sender, e);
		}

		private void cmbBodega_ValueChanged(object sender, System.EventArgs e)
		{
			this.cmbCajero.DataSource = Funcion.dvProcedimiento(cdsSeteo, string.Format("Exec PersonalLotte_Cajero {0}", (int)this.cmbBodega.Value));
		}

		private void ugPrendas_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (ugPrendas.ActiveCell != null)
			{
				try
				{				
					switch(e.KeyValue)
					{
							#region Enter
						case (int) Keys.Enter:
							ugPrendas.PerformAction(UltraGridAction.ExitEditMode, false, false);
							ugPrendas.PerformAction(UltraGridAction.NextCellByTab, false, false);
							e.Handled = true;
							ugPrendas.PerformAction(UltraGridAction.EnterEditMode, false, false);
							//							if(ultraGrid1.ActiveCell.DroppedDown == false) ultraGrid1.ActiveCell.DroppedDown = true;
							//						ultraGrid1.PerformAction(UltraGridAction.EnterEditModeAndDropdown, false, false);
							break;
							#endregion Enter
					}
				}
				catch(System.Exception ex)
				{
					MessageBox.Show(ex.Message, "Error del sistema.");
				}
			}
		}

		private void simpleButton1_Click(object sender, System.EventArgs e)
		{
			if(this.cmbCajero.ActiveRow == null)
			{
				MessageBox.Show("No se a seleccionado CAJERO?","Aviso....!!!", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.cmbCajero.Focus();
			}
			else if(this.cmbFormaPago.ActiveRow == null)
			{
				MessageBox.Show("No se a seleccionado FORMA DE PAGO?","Aviso....!!!", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.cmbFormaPago.Focus();
			}
			else if(this.dtFecha.Value == DBNull.Value)
			{
				MessageBox.Show("No se a seleccionado FECHA?","Aviso....!!!", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.dtFecha.Focus();
			}
			else if(idDetFP==0) 
			{
				MessageBox.Show("No a seleccionado Documento de PRENDA?","Aviso....!!!", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.ugPrendas.Focus();
			}
			else
			{
				#region Grabar
				#region Efectivo - Tarjeta 
				int idCpaFC = (int)ugPrendas.ActiveRow.Cells["idCompra"].Value;
				decimal ValorCobro = (decimal)ugPrendas.ActiveRow.Cells["Valor"].Value;
				int idClint = (int)ugPrendas.ActiveRow.Cells["idCliente"].Value;
				int idDFP = (int)this.cmbFormaPago.Value;
				int idCajro = (int)this.cmbCajero.Value;
				DateTime FeCobr = (DateTime)this.dtFecha.Value;
				string nrofac = "";
				if(this.txtCliente.Text.ToString().Length>0) nrofac=this.txtCliente.Text.ToString();
				else nrofac=ugPrendas.ActiveRow.Cells["Numero"].Value.ToString();
				DateTime FechaAct = DateTime.Today;

				using (VentaPagoLotPrenda miCobro = new VentaPagoLotPrenda(idCpaFC, (int)this.cmbFormaPago.Value, ValorCobro, nrofac, false, 0, 0, false, 0, false, 0, false, 0, 0, 6, idClint, 0, 0, 0, 0))
				{
					if (DialogResult.OK == miCobro.ShowDialog())
					{
						using (SqlConnection oConexion = new SqlConnection("Password=Bl45o6$9;Persist Security Info=True;User ID= " + MenuLatinium.stUsuario + ";Initial Catalog=" + MenuLatinium.stDirFacturacion + ";Data Source=" + MenuLatinium.stDirServidor))
						{
							oConexion.Open();
							SqlTransaction oTransaction = null;
							SqlCommand oCmdActualiza = oConexion.CreateCommand();
							oCmdActualiza.Connection = oConexion;
							oCmdActualiza.CommandType = CommandType.Text;
							oCmdActualiza.CommandTimeout = 0;
						
							try
							{
								oTransaction = oConexion.BeginTransaction();
								oCmdActualiza.Transaction = oTransaction;

								#region Guarda Cobro de la factura	
								if (miCobro.ultraGrid1.Rows.Count > 0)
								{
									foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in miCobro.ultraGrid1.Rows.All)
									{
										//int idDetFormaPago = 0;
										int idBanco = 0;
										int idTarjeta = 0;
										int idCuenta = 0;
										int idPlan = 0;
										int idPlazo = 0;
										decimal dComTarjeta = 0.00m;
										int idDocumento = 0;
										string DesPrenda = "";
										DateTime dtFecCaducidad = DateTime.Today;
							
										//if (dr.Cells["idDetFormaPago"].Value != System.DBNull.Value) idDetFormaPago = (int)dr.Cells["idDetFormaPago"].Value;
										if (dr.Cells["idBanco"].Value != System.DBNull.Value) idBanco = (int)dr.Cells["idBanco"].Value;
										if (dr.Cells["idTarjeta"].Value != System.DBNull.Value) idTarjeta = (int)dr.Cells["idTarjeta"].Value;
										if (dr.Cells["idCuenta"].Value != System.DBNull.Value) idCuenta = (int)dr.Cells["idCuenta"].Value;
										if (dr.Cells["idPlan"].Value != System.DBNull.Value) idPlan = (int)dr.Cells["idPlan"].Value;							
										if (dr.Cells["idPlazo"].Value != System.DBNull.Value) idPlazo = (int)dr.Cells["idPlazo"].Value;
										if ((int)dr.Cells["idFormaPago"].Value == 3) dtFecCaducidad = (DateTime)dr.Cells["FechaCaducidad"].Value;
										if (dr.Cells["ValorUsoTarjeta"].Value != System.DBNull.Value) dComTarjeta = (decimal)dr.Cells["ValorUsoTarjeta"].Value;
										if (dr.Cells["idDocumento"].Value != System.DBNull.Value) idDocumento = (int)dr.Cells["idDocumento"].Value;
										if (dr.Cells["DescripcionPrenda"].Value != System.DBNull.Value) DesPrenda = dr.Cells["DescripcionPrenda"].Value.ToString();
								
										string sSQLCobro = string.Format("Exec GuardaDetFP_Prenda {0}, {1}, {2}, {3}, {4}, {5}, {6}, '{7}', '{8}', '{9}', '{10}', '{11}', {12}, '{13}', '{14}', {15}, {16}, {17}, {18}, '{19}',{20}", 
											0, idCpaFC, 1, (int)dr.Cells["idFormaPago"].Value, (decimal)dr.Cells["Valor"].Value,
											idBanco, idTarjeta, dr.Cells["CuentaCorriente"].Value.ToString(), dr.Cells["Numero"].Value.ToString(),
											FeCobr.ToString("yyyyMMdd"), dr.Cells["NumVoucher"].Value.ToString(), dr.Cells["Autorizacion"].Value.ToString(),
											idCuenta, dr.Cells["Lote"].Value.ToString(), dr.Cells["Referencia"].Value.ToString(), 0, 0, 0, 0, "", (int)this.cmbCajero.Value);
										oCmdActualiza.CommandText = sSQLCobro;
										oCmdActualiza.ExecuteNonQuery();									

										/* VALOR DE ANTICIPOS */
										//if ((int)dr.Cells["idFormaPago"].Value == 8) dAnticipo = dAnticipo + Convert.ToDecimal(dr.Cells["Valor"].Value);

										#region actualiza cobro
										string actCobr = string.Format("Exec CobroPrenda {0},{1}",idDetFP,ValorCobro);
										oCmdActualiza.CommandText = actCobr;
										oCmdActualiza.ExecuteNonQuery();
										#endregion actualiza cobro
										/* ACTUALIZA SALDO DE NOTA DE CREDITO */
										if (idDocumento > 0)
										{
											oCmdActualiza.CommandText = string.Format("Exec ActualizaSaldoFacturas {0}, 5, 4", idDocumento, 1, 4);
											oCmdActualiza.ExecuteNonQuery();
										}
									}								
								}
								#endregion Guarda Cobro de la factura

								oTransaction.Commit();
							}
							catch (Exception ex)
							{
								try
								{
									oTransaction.Rollback();

									MessageBox.Show(string.Format("Commit Exception Type : {0} {1}", ex.GetType(), ex.Message), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
								}
								catch (Exception ex2)
								{
									oTransaction.Rollback();
									MessageBox.Show(string.Format("Rollback Exception Type: {0} {1}", ex2.GetType(), ex2.Message), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);							
								}
							}
							finally
							{
								oConexion.Close();
							}
							this.btBuscar_Click(sender, e);
						}
					}
					//else bEdicion = true;
					this.simpleButton1.Enabled = false;
				}
				#endregion Efectivo - Tarjeta 
				#endregion Grabar
			}
		}

		private void ugPrendas_DoubleClick(object sender, System.EventArgs e)
		{
			// obtener la fecha de factura
			try
			{
				//this.dtFecha.Value = ugPrendas.ActiveRow.Cells["Fecha"].Value;
				this.dtFecha.Value = DateTime.Now;
				idDetFP = (int)ugPrendas.ActiveRow.Cells["idDetFormaPago"].Value;
			}
			catch (Exception ex2) 
			{
				MessageBox.Show(string.Format("Error: {0} {1}", ex2.GetType(), ex2.Message), "Aviso...!!!", MessageBoxButtons.OK, MessageBoxIcon.Error);
			}
		}
	}
}
