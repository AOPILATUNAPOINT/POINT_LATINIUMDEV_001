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
	/// Descripción breve de frmNotasEntrega.
	/// </summary>
	public class frmNotasEntrega : System.Windows.Forms.Form
	{
		private System.Windows.Forms.Label lblBodega;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbBodega;
		private DevExpress.XtraEditors.SimpleButton btBuscar;
		private System.Windows.Forms.Label label12;
		private Infragistics.Win.UltraWinGrid.UltraGrid ugArticulos;
		private System.Windows.Forms.TextBox txtBskArticulo;
		private System.Windows.Forms.ToolTip toolTip1;
		private C1.Data.C1DataSet cdsSeteo;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource3;
		private System.Windows.Forms.RadioButton rb2;
		private System.Windows.Forms.RadioButton rb1;
		private System.Windows.Forms.Label label15;
		private System.Windows.Forms.Label lblIdentificacion;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbTipoRuc;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbCajero;
		private System.Windows.Forms.Label label19;
		private System.Windows.Forms.Label label10;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtFecha;
		private Infragistics.Win.UltraWinGrid.UltraGrid ugDetalleFactura;
		private System.Windows.Forms.TextBox txtNombre;
		private System.Windows.Forms.Label label3;
		private System.Windows.Forms.TextBox txtNumCedula;
		private System.Windows.Forms.Label label2;
		private System.Windows.Forms.Button button1;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource4;
		private System.Windows.Forms.Button bntGrabar;
		private System.Windows.Forms.Label label13;
		private System.Windows.Forms.Label label1;
		private System.Windows.Forms.Button btnNuevo;
		private System.Windows.Forms.Button btnBuscar;
		private System.Windows.Forms.Button btnImprimir;
		private System.Windows.Forms.Button btnAnular;
		private System.Windows.Forms.Button btnCancelar;
		private System.ComponentModel.IContainer components;

		public frmNotasEntrega()
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
		/// 

		int NumrHab = 0;
		string NomrBodga = "";
		int NBodga = 0;
		int idNhab = 0;
		int idClint = 0;
		int idCpaFC = 0;
		int idProyct = 0;
		decimal PorcentajeIva = 0.00m;
		int idAsient = 0;
		bool bImpreso = false;
		string NumFactura = "";
		private Acceso miAcceso;
		public int RegistroFactura = 0;
		bool bEliminaAlValidar = false;
		int idDetFormaPago = 0;
		string niUsuario = "";
		string nFacHab = "";
		int swCambio = 0;

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
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idBodega");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bodega");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmNotasEntrega));
			Infragistics.Win.Appearance appearance1 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idArticulo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idGrupoArticulo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idSubGrupo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Articulo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Precio1");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Precio2");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("UnboundColumn1", 0);
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idArticulo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idGrupoArticulo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn3 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idSubGrupo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn4 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Articulo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn5 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Precio1");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn6 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Precio2");
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand3 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn11 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idTipoRuc");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn12 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("TipoRuc");
			Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand4 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn13 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPersonal");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn14 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn15 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.Appearance appearance11 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton1 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.Appearance appearance12 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand5 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn16 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idDetCompra");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn17 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCompra");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn18 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idArticulo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn19 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Articulo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn20 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cantidad");
			Infragistics.Win.Appearance appearance13 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn21 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Notas");
			Infragistics.Win.Appearance appearance14 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance15 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance16 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance17 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance18 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance19 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn7 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idDetCompra");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn8 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCompra");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn9 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idArticulo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn10 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Articulo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn11 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Cantidad");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn12 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Notas");
			this.lblBodega = new System.Windows.Forms.Label();
			this.cmbBodega = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.btBuscar = new DevExpress.XtraEditors.SimpleButton();
			this.label12 = new System.Windows.Forms.Label();
			this.ugArticulos = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource3 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.txtBskArticulo = new System.Windows.Forms.TextBox();
			this.toolTip1 = new System.Windows.Forms.ToolTip(this.components);
			this.button1 = new System.Windows.Forms.Button();
			this.btnNuevo = new System.Windows.Forms.Button();
			this.btnBuscar = new System.Windows.Forms.Button();
			this.btnImprimir = new System.Windows.Forms.Button();
			this.bntGrabar = new System.Windows.Forms.Button();
			this.btnAnular = new System.Windows.Forms.Button();
			this.btnCancelar = new System.Windows.Forms.Button();
			this.cdsSeteo = new C1.Data.C1DataSet();
			this.rb2 = new System.Windows.Forms.RadioButton();
			this.rb1 = new System.Windows.Forms.RadioButton();
			this.label15 = new System.Windows.Forms.Label();
			this.lblIdentificacion = new System.Windows.Forms.Label();
			this.cmbTipoRuc = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.cmbCajero = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.label19 = new System.Windows.Forms.Label();
			this.label10 = new System.Windows.Forms.Label();
			this.dtFecha = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.ugDetalleFactura = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource4 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.txtNombre = new System.Windows.Forms.TextBox();
			this.label3 = new System.Windows.Forms.Label();
			this.txtNumCedula = new System.Windows.Forms.TextBox();
			this.label2 = new System.Windows.Forms.Label();
			this.label13 = new System.Windows.Forms.Label();
			this.label1 = new System.Windows.Forms.Label();
			((System.ComponentModel.ISupportInitialize)(this.cmbBodega)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ugArticulos)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource3)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbTipoRuc)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCajero)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFecha)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ugDetalleFactura)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource4)).BeginInit();
			this.SuspendLayout();
			// 
			// lblBodega
			// 
			this.lblBodega.AutoSize = true;
			this.lblBodega.Location = new System.Drawing.Point(8, 8);
			this.lblBodega.Name = "lblBodega";
			this.lblBodega.Size = new System.Drawing.Size(41, 16);
			this.lblBodega.TabIndex = 257;
			this.lblBodega.Text = "LOCAL";
			this.lblBodega.TextAlign = System.Drawing.ContentAlignment.TopRight;
			// 
			// cmbBodega
			// 
			this.cmbBodega.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbBodega.Cursor = System.Windows.Forms.Cursors.Default;
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridColumn3.Header.VisiblePosition = 2;
			ultraGridBand1.Columns.AddRange(new object[] {
																										 ultraGridColumn1,
																										 ultraGridColumn2,
																										 ultraGridColumn3});
			this.cmbBodega.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.cmbBodega.DisplayMember = "Nombre";
			this.cmbBodega.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbBodega.Enabled = false;
			this.cmbBodega.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbBodega.Location = new System.Drawing.Point(56, 6);
			this.cmbBodega.MaxDropDownItems = 30;
			this.cmbBodega.Name = "cmbBodega";
			this.cmbBodega.Size = new System.Drawing.Size(112, 21);
			this.cmbBodega.TabIndex = 256;
			this.cmbBodega.ValueMember = "idBodega";
			this.cmbBodega.ValueChanged += new System.EventHandler(this.cmbBodega_ValueChanged);
			// 
			// btBuscar
			// 
			this.btBuscar.Image = ((System.Drawing.Image)(resources.GetObject("btBuscar.Image")));
			this.btBuscar.Location = new System.Drawing.Point(312, 56);
			this.btBuscar.Name = "btBuscar";
			this.btBuscar.Size = new System.Drawing.Size(72, 22);
			this.btBuscar.TabIndex = 255;
			this.btBuscar.Text = "&Buscar";
			this.btBuscar.Click += new System.EventHandler(this.btBuscar_Click);
			// 
			// label12
			// 
			this.label12.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label12.Location = new System.Drawing.Point(8, 40);
			this.label12.Name = "label12";
			this.label12.Size = new System.Drawing.Size(64, 16);
			this.label12.TabIndex = 253;
			this.label12.Text = "ARTICULO";
			// 
			// ugArticulos
			// 
			this.ugArticulos.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.ugArticulos.Cursor = System.Windows.Forms.Cursors.Default;
			this.ugArticulos.DataSource = this.ultraDataSource3;
			appearance1.BackColor = System.Drawing.Color.White;
			this.ugArticulos.DisplayLayout.Appearance = appearance1;
			ultraGridColumn4.Header.VisiblePosition = 0;
			ultraGridColumn4.Hidden = true;
			ultraGridColumn5.Header.VisiblePosition = 1;
			ultraGridColumn5.Hidden = true;
			ultraGridColumn6.Header.VisiblePosition = 4;
			ultraGridColumn6.Hidden = true;
			ultraGridColumn7.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn7.Header.VisiblePosition = 3;
			ultraGridColumn7.Width = 330;
			ultraGridColumn8.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn8.Format = "#,##0.00";
			ultraGridColumn8.Header.Caption = "Precio";
			ultraGridColumn8.Header.VisiblePosition = 5;
			ultraGridColumn8.Width = 80;
			ultraGridColumn9.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn9.Format = "#,##0.00";
			ultraGridColumn9.Header.Caption = "Reserva";
			ultraGridColumn9.Header.VisiblePosition = 6;
			ultraGridColumn9.Hidden = true;
			ultraGridColumn9.Width = 80;
			ultraGridColumn10.ButtonDisplayStyle = Infragistics.Win.UltraWinGrid.ButtonDisplayStyle.Always;
			appearance2.Image = ((object)(resources.GetObject("appearance2.Image")));
			ultraGridColumn10.CellButtonAppearance = appearance2;
			ultraGridColumn10.Header.Caption = "Seleccionar";
			ultraGridColumn10.Header.VisiblePosition = 2;
			ultraGridColumn10.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.Button;
			ultraGridColumn10.Width = 20;
			ultraGridBand2.Columns.AddRange(new object[] {
																										 ultraGridColumn4,
																										 ultraGridColumn5,
																										 ultraGridColumn6,
																										 ultraGridColumn7,
																										 ultraGridColumn8,
																										 ultraGridColumn9,
																										 ultraGridColumn10});
			this.ugArticulos.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			appearance3.ForeColor = System.Drawing.Color.Black;
			appearance3.ForeColorDisabled = System.Drawing.Color.Black;
			this.ugArticulos.DisplayLayout.Override.ActiveRowAppearance = appearance3;
			this.ugArticulos.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.TemplateOnBottom;
			this.ugArticulos.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.ugArticulos.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.False;
			appearance4.BackColor = System.Drawing.Color.Transparent;
			this.ugArticulos.DisplayLayout.Override.CardAreaAppearance = appearance4;
			appearance5.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance5.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance5.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance5.FontData.BoldAsString = "True";
			appearance5.FontData.Name = "Arial";
			appearance5.FontData.SizeInPoints = 10F;
			appearance5.ForeColor = System.Drawing.Color.White;
			appearance5.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.ugArticulos.DisplayLayout.Override.HeaderAppearance = appearance5;
			appearance6.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance6.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance6.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ugArticulos.DisplayLayout.Override.RowAlternateAppearance = appearance6;
			appearance7.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance7.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance7.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ugArticulos.DisplayLayout.Override.RowSelectorAppearance = appearance7;
			appearance8.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance8.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance8.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ugArticulos.DisplayLayout.Override.SelectedRowAppearance = appearance8;
			this.ugArticulos.Font = new System.Drawing.Font("Microsoft Sans Serif", 13F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.ugArticulos.Location = new System.Drawing.Point(8, 80);
			this.ugArticulos.Name = "ugArticulos";
			this.ugArticulos.Size = new System.Drawing.Size(464, 416);
			this.ugArticulos.TabIndex = 252;
			this.ugArticulos.ClickCellButton += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.ugArticulos_ClickCellButton);
			this.ugArticulos.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.ugArticulos_InitializeLayout);
			// 
			// ultraDataSource3
			// 
			ultraDataColumn1.DataType = typeof(int);
			ultraDataColumn2.DataType = typeof(int);
			ultraDataColumn3.DataType = typeof(int);
			ultraDataColumn5.DataType = typeof(System.Decimal);
			ultraDataColumn6.DataType = typeof(System.Decimal);
			this.ultraDataSource3.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn1,
																																 ultraDataColumn2,
																																 ultraDataColumn3,
																																 ultraDataColumn4,
																																 ultraDataColumn5,
																																 ultraDataColumn6});
			// 
			// txtBskArticulo
			// 
			this.txtBskArticulo.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtBskArticulo.Location = new System.Drawing.Point(8, 56);
			this.txtBskArticulo.Name = "txtBskArticulo";
			this.txtBskArticulo.Size = new System.Drawing.Size(296, 20);
			this.txtBskArticulo.TabIndex = 254;
			this.txtBskArticulo.Text = "";
			this.txtBskArticulo.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtBskArticulo_KeyPress);
			// 
			// button1
			// 
			this.button1.Image = ((System.Drawing.Image)(resources.GetObject("button1.Image")));
			this.button1.Location = new System.Drawing.Point(1112, 22);
			this.button1.Name = "button1";
			this.button1.Size = new System.Drawing.Size(24, 23);
			this.button1.TabIndex = 275;
			this.toolTip1.SetToolTip(this.button1, "Registrar Número de Factura");
			this.button1.Click += new System.EventHandler(this.button1_Click);
			// 
			// btnNuevo
			// 
			this.btnNuevo.Image = ((System.Drawing.Image)(resources.GetObject("btnNuevo.Image")));
			this.btnNuevo.Location = new System.Drawing.Point(480, 456);
			this.btnNuevo.Name = "btnNuevo";
			this.btnNuevo.Size = new System.Drawing.Size(75, 40);
			this.btnNuevo.TabIndex = 301;
			this.toolTip1.SetToolTip(this.btnNuevo, "Nueva Nota de Entrega");
			this.btnNuevo.Click += new System.EventHandler(this.button3_Click);
			// 
			// btnBuscar
			// 
			this.btnBuscar.ForeColor = System.Drawing.Color.Coral;
			this.btnBuscar.Image = ((System.Drawing.Image)(resources.GetObject("btnBuscar.Image")));
			this.btnBuscar.Location = new System.Drawing.Point(640, 456);
			this.btnBuscar.Name = "btnBuscar";
			this.btnBuscar.Size = new System.Drawing.Size(75, 40);
			this.btnBuscar.TabIndex = 300;
			this.toolTip1.SetToolTip(this.btnBuscar, "Buscar Nota de Entrega");
			this.btnBuscar.Click += new System.EventHandler(this.button2_Click);
			// 
			// btnImprimir
			// 
			this.btnImprimir.ForeColor = System.Drawing.Color.Coral;
			this.btnImprimir.Image = ((System.Drawing.Image)(resources.GetObject("btnImprimir.Image")));
			this.btnImprimir.Location = new System.Drawing.Point(720, 456);
			this.btnImprimir.Name = "btnImprimir";
			this.btnImprimir.Size = new System.Drawing.Size(75, 40);
			this.btnImprimir.TabIndex = 299;
			this.toolTip1.SetToolTip(this.btnImprimir, "Imprimir Nota de Entrega");
			this.btnImprimir.Click += new System.EventHandler(this.btnAnular_Click);
			// 
			// bntGrabar
			// 
			this.bntGrabar.Image = ((System.Drawing.Image)(resources.GetObject("bntGrabar.Image")));
			this.bntGrabar.Location = new System.Drawing.Point(560, 456);
			this.bntGrabar.Name = "bntGrabar";
			this.bntGrabar.Size = new System.Drawing.Size(75, 40);
			this.bntGrabar.TabIndex = 298;
			this.toolTip1.SetToolTip(this.bntGrabar, "Grabar Nota de Entrega");
			this.bntGrabar.Click += new System.EventHandler(this.bntGrabar_Click);
			// 
			// btnAnular
			// 
			this.btnAnular.ForeColor = System.Drawing.Color.Coral;
			this.btnAnular.Image = ((System.Drawing.Image)(resources.GetObject("btnAnular.Image")));
			this.btnAnular.Location = new System.Drawing.Point(880, 456);
			this.btnAnular.Name = "btnAnular";
			this.btnAnular.Size = new System.Drawing.Size(75, 40);
			this.btnAnular.TabIndex = 304;
			this.toolTip1.SetToolTip(this.btnAnular, "Anular Nota de Entrega");
			this.btnAnular.Visible = false;
			this.btnAnular.Click += new System.EventHandler(this.btnAnular_Click_1);
			// 
			// btnCancelar
			// 
			this.btnCancelar.ForeColor = System.Drawing.Color.Coral;
			this.btnCancelar.Image = ((System.Drawing.Image)(resources.GetObject("btnCancelar.Image")));
			this.btnCancelar.Location = new System.Drawing.Point(800, 456);
			this.btnCancelar.Name = "btnCancelar";
			this.btnCancelar.Size = new System.Drawing.Size(75, 40);
			this.btnCancelar.TabIndex = 305;
			this.toolTip1.SetToolTip(this.btnCancelar, "Cancelar Nota de Entrega");
			this.btnCancelar.Click += new System.EventHandler(this.btnCancelar_Click);
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
			this.cdsSeteo.BeforeFill += new C1.Data.FillEventHandler(this.cdsSeteo_BeforeFill);
			// 
			// rb2
			// 
			this.rb2.Location = new System.Drawing.Point(816, 48);
			this.rb2.Name = "rb2";
			this.rb2.Size = new System.Drawing.Size(32, 16);
			this.rb2.TabIndex = 297;
			this.rb2.Text = "2";
			this.rb2.Visible = false;
			// 
			// rb1
			// 
			this.rb1.Checked = true;
			this.rb1.Location = new System.Drawing.Point(784, 48);
			this.rb1.Name = "rb1";
			this.rb1.Size = new System.Drawing.Size(32, 16);
			this.rb1.TabIndex = 296;
			this.rb1.TabStop = true;
			this.rb1.Text = "1";
			this.rb1.Visible = false;
			// 
			// label15
			// 
			this.label15.AutoSize = true;
			this.label15.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label15.Location = new System.Drawing.Point(728, 48);
			this.label15.Name = "label15";
			this.label15.Size = new System.Drawing.Size(46, 16);
			this.label15.TabIndex = 295;
			this.label15.Text = "TURNO";
			this.label15.Visible = false;
			// 
			// lblIdentificacion
			// 
			this.lblIdentificacion.AutoSize = true;
			this.lblIdentificacion.Location = new System.Drawing.Point(560, 8);
			this.lblIdentificacion.Name = "lblIdentificacion";
			this.lblIdentificacion.Size = new System.Drawing.Size(87, 16);
			this.lblIdentificacion.TabIndex = 292;
			this.lblIdentificacion.Text = "Tipo Documento";
			// 
			// cmbTipoRuc
			// 
			appearance9.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbTipoRuc.Appearance = appearance9;
			this.cmbTipoRuc.CausesValidation = false;
			this.cmbTipoRuc.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbTipoRuc.Cursor = System.Windows.Forms.Cursors.Default;
			ultraGridColumn11.Header.VisiblePosition = 0;
			ultraGridColumn11.Hidden = true;
			ultraGridColumn12.Header.VisiblePosition = 1;
			ultraGridColumn12.Width = 233;
			ultraGridBand3.Columns.AddRange(new object[] {
																										 ultraGridColumn11,
																										 ultraGridColumn12});
			this.cmbTipoRuc.DisplayLayout.BandsSerializer.Add(ultraGridBand3);
			this.cmbTipoRuc.DisplayMember = "TipoRuc";
			this.cmbTipoRuc.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbTipoRuc.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbTipoRuc.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbTipoRuc.Location = new System.Drawing.Point(560, 24);
			this.cmbTipoRuc.Name = "cmbTipoRuc";
			this.cmbTipoRuc.Size = new System.Drawing.Size(128, 21);
			this.cmbTipoRuc.TabIndex = 291;
			this.cmbTipoRuc.ValueMember = "idTipoRuc";
			// 
			// cmbCajero
			// 
			appearance10.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbCajero.Appearance = appearance10;
			this.cmbCajero.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbCajero.Cursor = System.Windows.Forms.Cursors.Default;
			ultraGridColumn13.Header.VisiblePosition = 0;
			ultraGridColumn13.Hidden = true;
			ultraGridColumn14.Header.VisiblePosition = 1;
			ultraGridColumn14.Width = 335;
			ultraGridColumn15.Header.VisiblePosition = 2;
			ultraGridBand4.Columns.AddRange(new object[] {
																										 ultraGridColumn13,
																										 ultraGridColumn14,
																										 ultraGridColumn15});
			this.cmbCajero.DisplayLayout.BandsSerializer.Add(ultraGridBand4);
			this.cmbCajero.DisplayMember = "Nombre";
			this.cmbCajero.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbCajero.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbCajero.Location = new System.Drawing.Point(560, 46);
			this.cmbCajero.Name = "cmbCajero";
			this.cmbCajero.Size = new System.Drawing.Size(152, 21);
			this.cmbCajero.TabIndex = 290;
			this.cmbCajero.ValueMember = "idPersonal";
			// 
			// label19
			// 
			this.label19.AutoSize = true;
			this.label19.Location = new System.Drawing.Point(520, 48);
			this.label19.Name = "label19";
			this.label19.Size = new System.Drawing.Size(38, 16);
			this.label19.TabIndex = 289;
			this.label19.Text = "Cajero";
			// 
			// label10
			// 
			this.label10.AutoSize = true;
			this.label10.Font = new System.Drawing.Font("Microsoft Sans Serif", 8F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label10.Location = new System.Drawing.Point(448, 8);
			this.label10.Name = "label10";
			this.label10.Size = new System.Drawing.Size(54, 16);
			this.label10.TabIndex = 288;
			this.label10.Text = "Fecha FC";
			// 
			// dtFecha
			// 
			appearance11.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtFecha.Appearance = appearance11;
			this.dtFecha.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton1.Caption = "Today";
			this.dtFecha.DateButtons.Add(dateButton1);
			this.dtFecha.Format = "dd/MM/yyyy";
			this.dtFecha.Location = new System.Drawing.Point(448, 24);
			this.dtFecha.Name = "dtFecha";
			this.dtFecha.NonAutoSizeHeight = 23;
			this.dtFecha.Size = new System.Drawing.Size(104, 21);
			this.dtFecha.SpinButtonsVisible = true;
			this.dtFecha.TabIndex = 287;
			this.dtFecha.Value = ((object)(resources.GetObject("dtFecha.Value")));
			// 
			// ugDetalleFactura
			// 
			this.ugDetalleFactura.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.ugDetalleFactura.Cursor = System.Windows.Forms.Cursors.Default;
			this.ugDetalleFactura.DataSource = this.ultraDataSource4;
			appearance12.BackColor = System.Drawing.Color.White;
			this.ugDetalleFactura.DisplayLayout.Appearance = appearance12;
			ultraGridColumn16.Header.VisiblePosition = 0;
			ultraGridColumn16.Hidden = true;
			ultraGridColumn17.Header.VisiblePosition = 1;
			ultraGridColumn17.Hidden = true;
			ultraGridColumn18.Header.VisiblePosition = 2;
			ultraGridColumn18.Hidden = true;
			ultraGridColumn19.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
			ultraGridColumn19.Header.VisiblePosition = 3;
			ultraGridColumn19.Width = 300;
			appearance13.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn20.CellAppearance = appearance13;
			ultraGridColumn20.Format = "#,##0";
			ultraGridColumn20.Header.VisiblePosition = 4;
			ultraGridColumn20.Width = 40;
			ultraGridColumn21.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn21.Header.VisiblePosition = 5;
			ultraGridColumn21.Width = 400;
			ultraGridBand5.Columns.AddRange(new object[] {
																										 ultraGridColumn16,
																										 ultraGridColumn17,
																										 ultraGridColumn18,
																										 ultraGridColumn19,
																										 ultraGridColumn20,
																										 ultraGridColumn21});
			this.ugDetalleFactura.DisplayLayout.BandsSerializer.Add(ultraGridBand5);
			appearance14.ForeColor = System.Drawing.Color.Black;
			appearance14.ForeColorDisabled = System.Drawing.Color.Black;
			this.ugDetalleFactura.DisplayLayout.Override.ActiveRowAppearance = appearance14;
			this.ugDetalleFactura.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.TemplateOnBottom;
			this.ugDetalleFactura.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.True;
			this.ugDetalleFactura.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance15.BackColor = System.Drawing.Color.Transparent;
			this.ugDetalleFactura.DisplayLayout.Override.CardAreaAppearance = appearance15;
			this.ugDetalleFactura.DisplayLayout.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.Edit;
			appearance16.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance16.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance16.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance16.FontData.BoldAsString = "True";
			appearance16.FontData.Name = "Arial";
			appearance16.FontData.SizeInPoints = 10F;
			appearance16.ForeColor = System.Drawing.Color.White;
			appearance16.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.ugDetalleFactura.DisplayLayout.Override.HeaderAppearance = appearance16;
			appearance17.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance17.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance17.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ugDetalleFactura.DisplayLayout.Override.RowAlternateAppearance = appearance17;
			appearance18.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance18.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance18.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ugDetalleFactura.DisplayLayout.Override.RowSelectorAppearance = appearance18;
			appearance19.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance19.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance19.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ugDetalleFactura.DisplayLayout.Override.SelectedRowAppearance = appearance19;
			this.ugDetalleFactura.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.ugDetalleFactura.Location = new System.Drawing.Point(480, 80);
			this.ugDetalleFactura.Name = "ugDetalleFactura";
			this.ugDetalleFactura.Size = new System.Drawing.Size(664, 336);
			this.ugDetalleFactura.TabIndex = 280;
			this.ugDetalleFactura.BeforeRowsDeleted += new Infragistics.Win.UltraWinGrid.BeforeRowsDeletedEventHandler(this.ugDetalleFactura_BeforeRowsDeleted);
			this.ugDetalleFactura.KeyDown += new System.Windows.Forms.KeyEventHandler(this.ugDetalleFactura_KeyDown);
			this.ugDetalleFactura.AfterCellUpdate += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.ugDetalleFactura_AfterCellUpdate);
			// 
			// ultraDataSource4
			// 
			ultraDataColumn7.DataType = typeof(int);
			ultraDataColumn8.DataType = typeof(int);
			ultraDataColumn9.DataType = typeof(int);
			ultraDataColumn11.DataType = typeof(int);
			this.ultraDataSource4.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn7,
																																 ultraDataColumn8,
																																 ultraDataColumn9,
																																 ultraDataColumn10,
																																 ultraDataColumn11,
																																 ultraDataColumn12});
			// 
			// txtNombre
			// 
			this.txtNombre.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtNombre.Location = new System.Drawing.Point(800, 24);
			this.txtNombre.Name = "txtNombre";
			this.txtNombre.Size = new System.Drawing.Size(296, 20);
			this.txtNombre.TabIndex = 279;
			this.txtNombre.Text = "";
			// 
			// label3
			// 
			this.label3.Location = new System.Drawing.Point(800, 8);
			this.label3.Name = "label3";
			this.label3.Size = new System.Drawing.Size(64, 16);
			this.label3.TabIndex = 278;
			this.label3.Text = "NOMBRES";
			// 
			// txtNumCedula
			// 
			this.txtNumCedula.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtNumCedula.Location = new System.Drawing.Point(696, 24);
			this.txtNumCedula.Name = "txtNumCedula";
			this.txtNumCedula.TabIndex = 277;
			this.txtNumCedula.Text = "";
			this.txtNumCedula.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
			this.txtNumCedula.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtNumCedula_KeyDown);
			this.txtNumCedula.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtNumCedula_KeyPress);
			this.txtNumCedula.Validating += new System.ComponentModel.CancelEventHandler(this.txtNumCedula_Validating);
			this.txtNumCedula.Validated += new System.EventHandler(this.txtNumCedula_Validated);
			// 
			// label2
			// 
			this.label2.Location = new System.Drawing.Point(696, 8);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(64, 16);
			this.label2.TabIndex = 276;
			this.label2.Text = "# CEDULA";
			// 
			// label13
			// 
			this.label13.AutoSize = true;
			this.label13.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label13.ForeColor = System.Drawing.Color.Red;
			this.label13.Location = new System.Drawing.Point(480, 424);
			this.label13.Name = "label13";
			this.label13.Size = new System.Drawing.Size(78, 22);
			this.label13.TabIndex = 302;
			this.label13.Text = "NORMAL";
			// 
			// label1
			// 
			this.label1.AutoSize = true;
			this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 14F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label1.Location = new System.Drawing.Point(216, 8);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(56, 25);
			this.label1.TabIndex = 303;
			this.label1.Text = "# NE:";
			// 
			// frmNotasEntrega
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(1152, 501);
			this.Controls.Add(this.btnCancelar);
			this.Controls.Add(this.btnAnular);
			this.Controls.Add(this.label1);
			this.Controls.Add(this.label13);
			this.Controls.Add(this.btnNuevo);
			this.Controls.Add(this.btnBuscar);
			this.Controls.Add(this.btnImprimir);
			this.Controls.Add(this.bntGrabar);
			this.Controls.Add(this.rb2);
			this.Controls.Add(this.rb1);
			this.Controls.Add(this.label15);
			this.Controls.Add(this.lblIdentificacion);
			this.Controls.Add(this.cmbTipoRuc);
			this.Controls.Add(this.cmbCajero);
			this.Controls.Add(this.label19);
			this.Controls.Add(this.label10);
			this.Controls.Add(this.dtFecha);
			this.Controls.Add(this.ugDetalleFactura);
			this.Controls.Add(this.txtNombre);
			this.Controls.Add(this.label3);
			this.Controls.Add(this.txtNumCedula);
			this.Controls.Add(this.label2);
			this.Controls.Add(this.button1);
			this.Controls.Add(this.lblBodega);
			this.Controls.Add(this.cmbBodega);
			this.Controls.Add(this.btBuscar);
			this.Controls.Add(this.label12);
			this.Controls.Add(this.ugArticulos);
			this.Controls.Add(this.txtBskArticulo);
			this.Name = "frmNotasEntrega";
			this.Text = "Notas de ENTREGA";
			this.Load += new System.EventHandler(this.frmNotasEntrega_Load);
			this.Closed += new System.EventHandler(this.frmNotasEntrega_Closed);
			((System.ComponentModel.ISupportInitialize)(this.cmbBodega)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ugArticulos)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource3)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbTipoRuc)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCajero)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFecha)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ugDetalleFactura)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource4)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void frmNotasEntrega_Load(object sender, System.EventArgs e)
		{
			/*Trabajar con idTipoFactura=51*/
			/* Porcentaje del IVA */
			string lcPiva = "SELECT Iva FROM SeteoF WHERE idSeteoF=1";
			PorcentajeIva = Funcion.decEscalarSQL(cdsSeteo, lcPiva); 

			this.dtFecha.Value = DateTime.Today;
			
			//Carga Vendedores, Cajero y Cobrador
			this.cmbCajero.DataSource = Funcion.dvProcedimiento(cdsSeteo, string.Format("Exec PersonalLotte_Cajero {0}", NBodga));
			string stExecBodega = string.Format("Select idBodega, Bodega, Nombre From Bodega Where Activo = 1");
      this.cmbBodega.DataSource = Funcion.dvProcedimiento(cdsSeteo, stExecBodega);
			
			int codBodHabtcn = Funcion.iEscalarSQL(cdsSeteo, "SELECT Bodega FROM UsuariosBodegas WHERE idUsuario="+MenuLatinium.IdUsuario.ToString());
			int BodaBusk = Funcion.iEscalarSQL(cdsSeteo, "SELECT idBodega FROM Bodega WHERE Bodega="+codBodHabtcn.ToString());
			this.cmbBodega.Value = BodaBusk;

			this.cmbTipoRuc.DataSource = Funcion.dvProcedimiento(cdsSeteo, "Exec ClienteTipoRucCarga 1");
			this.cmbTipoRuc.Value = 5;

			/*Habilitar boton de Anular nota de entrega*/
			niUsuario = MenuLatinium.iNivel.ToString().Trim();
			if(niUsuario.Equals("4") || niUsuario.Equals("1") || niUsuario.Equals("2") || niUsuario.Equals("3")) 
			{
				this.btnAnular.Visible = true;
				this.cmbBodega.Enabled = true;
			}
		}

		private void cdsSeteo_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteo.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);
		}

		private void txtBskArticulo_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (e.KeyChar == 13) this.btBuscar_Click(sender, e);
		}

		private void btBuscar_Click(object sender, System.EventArgs e)
		{
			this.Cursor = Cursors.WaitCursor;
			string stBusca = txtBskArticulo.Text.Trim();
			stBusca = stBusca.Replace("'", "''");
				
			//iBodega = (int)this.cmbBodega.Value;			

			DateTime dtFecha = DateTime.Today;

			string sSQL = string.Format("Exec ConsultaInv_Lotte '{0}'", stBusca);
			ugArticulos.DataSource = Funcion.dvProcedimiento(cdsSeteo, sSQL);
					
			ugArticulos.Focus();
			this.Cursor = Cursors.Default;
		}

		private void txtNumCedula_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.F3)
			{
				using (BuscaClientes miBusqueda = new BuscaClientes(this.txtNumCedula.Text.ToString(), 1, 1, 2, (int)idNhab, false))
				{
					if (DialogResult.OK == miBusqueda.ShowDialog())	
					{
						idClint = miBusqueda.idCliente;
						this.txtNumCedula.Text = miBusqueda.Ruc;
					}
				}
			}
		}

		private void txtNumCedula_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (e.KeyChar == 13) this.txtNombre.Focus();
		}

		private void txtNumCedula_Validated(object sender, System.EventArgs e)
		{
			#region Busca Cliente 
			/* Buscar Número de Bodega */
			string lcQryBd = "SELECT idProyecto From Bodega Where idBodega in (Select idSucursal From Habitaciones Where idHabitacion=5)";
			idProyct = Funcion.iEscalarSQL(cdsSeteo, lcQryBd);
				
			string sSQL = string.Format("Exec BuscaClientePorRUC '{0}', {1}, {2}, {3}", this.txtNumCedula.Text.ToString(), 1, 5, false);
			SqlDataReader drC = Funcion.rEscalarSQL(cdsSeteo, sSQL, true);
			drC.Read();
			if (drC.HasRows)
			{
				idClint = (int)drC.GetValue(0);
				
				this.txtNumCedula.Text = (string)drC.GetValue(1);
				this.txtNombre.Text = (string)drC.GetValue(2);
				if (!this.txtNumCedula.Text.ToString().Equals("9999999999"))
				{
					//this.txtDireccion.Text = (string)drC.GetValue(7);
					//this.txtTelefono.Text = (string)drC.GetValue(8);
				}
				drC.Close();
			}
			else
			{						
				using (VentaCliente miVCliente = new VentaCliente(this.txtNumCedula.Text, 0, 1))
				{
					if (DialogResult.OK == miVCliente.ShowDialog())
					{
						MessageBox.Show("Cliente registrado correctamente.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
						this.txtNombre.Text = miVCliente.txtNombre.Value.ToString().Trim();
						//this.txtDireccion.Text = miVCliente.txtDireccion.Value.ToString().Trim();
						//this.txtTelefono.Text = miVCliente.txtTelefono.Value.ToString().Trim();
					}
					else this.txtNumCedula.Focus();
				}			
			}
			drC.Close();		
			#endregion Busca Cliente
		}

		private void txtNumCedula_Validating(object sender, System.ComponentModel.CancelEventArgs e)
		{
			#region Valida identificación
			string VlCed = "";
			if(!this.txtNumCedula.Text.ToString().Equals("9999999999"))
			{
				string stExec = string.Format("Select dbo.ValidaIdentificacion ('{0}', {1})", this.txtNumCedula.Text.ToString(), this.cmbTipoRuc.Value);
				VlCed = Funcion.sEscalarSQL(cdsSeteo, stExec);
			}	
			if (VlCed.StartsWith("Error"))
			{
				MessageBox.Show(VlCed, "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.txtNumCedula.Focus();
				e.Cancel = true;			
			}
			#endregion Valida identificación
		}

		private void button1_Click(object sender, System.EventArgs e)
		{
			// Registrar nota de entrega
			try
			{
				//Validar Numero de Factura
				if(this.cmbBodega.ActiveRow == null)
				{
					MessageBox.Show("No se a seleccionado Bodega?","Aviso....!!!", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					this.cmbBodega.Focus();
				}
				else 
				{
					int idBod = (int)this.cmbBodega.Value;
					string SP_Qry = string.Format("Exec NotaEntrega_LotteServicio {0}", idBod);
					idCpaFC = (int)Funcion.iEscalarSQL(cdsSeteo,SP_Qry);
					string lcQry = "Select Numero From Compra Where idCompra="+idCpaFC.ToString();
					nFacHab = Funcion.sEscalarSQL(cdsSeteo, lcQry);
					this.label1.Text = "# NE:"+nFacHab;
					RegistroFactura = 1;
				}
				this.button1.Enabled = false;
				swCambio = 1;
			}
			catch (Exception ex2) 
			{
				MessageBox.Show(string.Format("Error: {0} {1}", ex2.GetType(), ex2.Message), "Aviso...!!!", MessageBoxButtons.OK, MessageBoxIcon.Error);
			}
		}

		private void button3_Click(object sender, System.EventArgs e)
		{
			/*Encerar variable de pantalla*/
			this.txtNumCedula.Text = "";
			this.dtFecha.Value = DateTime.Today;
			this.cmbCajero.Value = System.DBNull.Value;
			this.txtNombre.Text = "";
			this.bntGrabar.Enabled = true;
			this.label13.Text = "";
			this.label1.Text = "# NE:";
			this.btnAnular.Enabled = false;
			/*Grilla elimina datos*/
			string QryDetFc = "SELECT idDetCompra,idCompra,idArticulo,Articulo,Cantidad,Precio,SubTotal From DetCompra Where idCompra=0";
			ugDetalleFactura.DataSource = Funcion.dvProcedimiento(cdsSeteo, QryDetFc);
			this.ugDetalleFactura.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			this.ugDetalleFactura.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.True;
			this.button1.Enabled = true;
			swCambio = 1;
		}

		private void bntGrabar_Click(object sender, System.EventArgs e)
		{
			/* Grabar Factura Habitación */
			if(this.txtNumCedula.Text.ToString().Length==0)
			{
				MessageBox.Show("Debe Registrar CLIENTE para poder Grabar la FACTURA.?","Aviso....!!!", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.txtNumCedula.Focus();
			}
			else if(this.dtFecha.Value == DBNull.Value)
			{
				MessageBox.Show("Debe Seleccionar Factura y obtener la Fecha para Grabar la FACTURA.?","Aviso....!!!", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.dtFecha.Focus();
			}
			else if(this.cmbCajero.ActiveRow == null)
			{
				MessageBox.Show("No se a seleccionado Cajero?","Aviso....!!!", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.cmbCajero.Focus();
			}
			else
			{
				bool Manual = false;
				string sSQL = "";
				decimal dRetenido = 0;
				decimal dRetenidoIVA = 0;
				decimal dRetenidoRenta = 0;
				int idDetFormaPagoRetencion = 0;
				DateTime dtFechaCadR = DateTime.Today;

				#region Valida Articulos en el grid
				if (this.ugDetalleFactura.Rows.Count == 0)
				{
					MessageBox.Show("Ingrese los productos para crear Nota de Entrega", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
					this.txtBskArticulo.Focus();
					return;
				}
				#endregion Valida Articulos en el grid

				#region Grabar
				#region Efectivo - Tarjeta 
				DateTime dtFechaFC = (DateTime) this.dtFecha.Value;
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

						#region Graba Factura
						#region Maestro
						/*Tipo de FACTURA = 51 de nota de entrega*/
						sSQL = string.Format("Exec GrabaMaestroFacturaLotte {0}, {1}, {2}, {3}, {4}, {5}, {6}, {7}, {8}, {9}, {10}, {11}, '{12}', '{13}', {14}, {15}, '{16}', '{17}', '{18}', '{19}', '{20}', '{21}', {22}, '{23}', '{24}', {25}, '{26}', {27}, {28}, {29}, {30}, '{31}', {32}, {33}, {34}, {35}, {36}, {37}, {38}, {39}, {40}, {41}, {42}, 1, 0, 0, '', 0.0, {43}, {44}, {45}, {46}, {47}, {48}, {49}, {50}, {51}, {52}, {53},'', 0", 
							idCpaFC, 1,	0, 51, 15, 10, false, 1, 0, NBodga, idProyct, 0, nFacHab, "",
							5, idClint, this.txtNumCedula.Text.ToString(), this.txtNombre.Text.ToString(), dtFechaFC.ToString("yyyyMMdd"), 
							"", "", dtFechaFC.ToString("yyyyMMdd"), false, dtFechaFC.ToString("yyyyMMdd"), dtFechaFC.ToString("yyyyMMdd HH:mm"), 0,
							"", false, 0, 0, 0, dtFechaFC.ToString("yyyyMMdd"), 0, 0, 0, 
							0, 0, 0,	false, false, false, false, false, 0, 0, 0, 0, false, 0, 0, 0, 0, 0, 0);
						oCmdActualiza.CommandText = sSQL;
						idCpaFC = (int)oCmdActualiza.ExecuteScalar();
						#endregion Maestro
						#region cajero vendedor cobrador
						string lcQryCVC = string.Format("Exec ActualizaCajCob {0},{1},{2},0", 
							idCpaFC, 0, (int)this.cmbCajero.Value);
						oCmdActualiza.CommandText = lcQryCVC;
						oCmdActualiza.ExecuteNonQuery();
						#endregion cajero vendedor cobrador
						#region cajero vendedor cobrador
						string lcQCVC = string.Format("Exec GrabaVen_Caj_Cob_Turno {0},{1},{2},{3},0", 
							idCpaFC, 0, 0, (int)this.cmbCajero.Value);
						oCmdActualiza.CommandText = lcQCVC;
						oCmdActualiza.ExecuteNonQuery();
						#endregion cajero vendedor cobrador

						#region Detalle
						int iPosicion = 0;
						foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.ugDetalleFactura.Rows.All)
						{	
							iPosicion ++;
							string lcQryCAr = "SELECT Codigo From Articulo Where idArticulo="+dr.Cells["idArticulo"].Value.ToString();
							string CodigoArti = Funcion.sEscalarSQL(cdsSeteo, lcQryCAr);

							string sSQLDetalle = string.Format("Exec GuardaDetalleFacturaLotte {0}, {1}, {2}, {3}, '{4}', {5}, {6}, '{7}', '{8}', {9}, {10}, {11}, {12}, {13}, {14}, {15}, {16}, {17}, {18}, {19}, {20}, {21}, {22}, '{23}'",
								(int)dr.Cells["idDetCompra"].Value, 0, (int)dr.Cells["idCompra"].Value,	false, "", (int)dr.Cells["idArticulo"].Value, 0, 
								CodigoArti, dr.Cells["Articulo"].Value,	(int)dr.Cells["Cantidad"].Value, 0, 
								PorcentajeIva, 0, 0, 0,	NBodga, -1,	0, false, iPosicion, false, false, 0, dr.Cells["Notas"].Value);
							oCmdActualiza.CommandText = sSQLDetalle;
							oCmdActualiza.ExecuteNonQuery();
						}
						#endregion Detalle	
						#endregion Graba Factura

						#region Asiento Automatico
						if (false) //Hay que poner en verdadero cuando este listo el plan de cuentas
						{						
							string sSQLAsiento = string.Format("Exec CreaAsientoDeVenta {0}", idCpaFC.ToString());
							oCmdActualiza.CommandText = sSQLAsiento;
							oCmdActualiza.ExecuteNonQuery();

							string sSQLRetIdAsiento = string.Format("Select ISNULL(idAsiento, 0) From Compra Where idCompra = {0}", idCpaFC.ToString());
							oCmdActualiza.CommandText = sSQLRetIdAsiento; 
							idAsient = (int)oCmdActualiza.ExecuteScalar(); 

							this.label13.Text = "CONTABILIZADO";
								
							Cursor = Cursors.Default;								
						}
						#endregion Asiento Automatico

						oTransaction.Commit();
					
						#region Controles
						//this.cmbFormaPago.Enabled = false;
						this.bntGrabar.Enabled = false;
						#endregion Controles
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
				}
				#endregion Efectivo - Tarjeta 
				#endregion Grabar
				swCambio = 0;
			}
		}

		private void button2_Click(object sender, System.EventArgs e)
		{
			if(this.cmbBodega.ActiveRow == null)
			{
				MessageBox.Show("No se a seleccionado Bodega?","Aviso....!!!", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.cmbBodega.Focus();
			}
			else
			{
				using (frmBuscaFacturas miBusqueda = new frmBuscaFacturas(51, NBodga, miAcceso.BLimiteBusquedaFechas))
				{
					if (DialogResult.OK == miBusqueda.ShowDialog())
					{
						this.Consultar((int)miBusqueda.grdFacturas.ActiveRow.Cells["idCompra"].Value);
						this.btnAnular.Enabled = true;
					}
				}
			}
			swCambio = 0;
		}

		private void btnAnular_Click(object sender, System.EventArgs e)
		{
			/*Realizar impresión de nota de entrega*/
			#region Impresión
      Imprimir(0);
			#endregion Impresión
		}

		private void Imprimir(int iTipo)
		{
			#region Nombre de Reporte y Filtro
			Cursor = Cursors.WaitCursor;
			string stReporte = "Exec CompraReporteNombre " + idCpaFC.ToString();
			
			stReporte += ", 1";
			string stFiltro = "";
			stReporte = Funcion.sEscalarSQL(cdsSeteo, stReporte);
			if (iTipo==0)	stFiltro = "{Compra.idCompra} = " + idCpaFC.ToString();
			else stFiltro = "{RegistroPromociones.idCompra} = " + idCpaFC.ToString();
			#endregion Nombre de Reporte y Filtro

			#region Imprime el reporte
			//using (Reporte miReporte = new Reporte(stReporte, stFiltro))
			if (iTipo == 0) 
			{
				#region Impresion
				ParameterFields paramFields = new ParameterFields ();
				
				ParameterField pfRet = new ParameterField ();
				ParameterDiscreteValue idCompraRet = new ParameterDiscreteValue ();
				pfRet.ParameterFieldName = "@idCpa";
				idCompraRet.Value = idCpaFC;
				pfRet.CurrentValues.Add (idCompraRet);
				paramFields.Add (pfRet);
				
				Reporte miRepor = new Reporte("FacturaNotaEntrega.rpt", "");
				miRepor.MdiParent = this.MdiParent;
				miRepor.Titulo("Retenciones");
				miRepor.crVista.ParameterFieldInfo = paramFields;
				#endregion Impresion
				
				miRepor.Show();

				#region visualiza reporte
				#endregion visualiza reporte
			}
			if (iTipo == 1) 
			{
				using (Reporte miReporte = new Reporte("ImprPromocion", stFiltro))
				{		
					if (iTipo == 1) // Muestra en pantalla
						miReporte.ShowDialog();
					if (iTipo >= 2) // Imprime una o mas copias
					{
						string ImpresoraPredefinida = MenuLatinium.ImpresoraNombre;
						string stImpresora = Funcion.sEscalarSQL(cdsSeteo, "Exec CompraReporteImpresora '" + stReporte + "'");
						if (stImpresora.Length > 0)	MenuLatinium.ImpresoraNombre = stImpresora;
						miReporte.ImprimeRep(1);
						if (stImpresora.Length > 0)	MenuLatinium.ImpresoraNombre = ImpresoraPredefinida;
					}
					if (iTipo == -1) // Archivo PDF
					{
						miReporte.ImprimeRep(2);
					}
				}
			}
			#endregion Imprime el reporte

			#region Marca como impreso y muestra el tiempo
			string stGrabaImprime = "Update Compra Set Impreso = 1 Where idCompra = " + idCpaFC.ToString();
			Funcion.EjecutaSQL(cdsSeteo, stGrabaImprime);
			bImpreso = true;
			Cursor = Cursors.Default;
			#endregion Marca como impreso y muestra el tiempo
		}

		private void Consultar(int IdCompra)
		{
			try
			{	
				#region Maestro
				idCpaFC = IdCompra;
				int iBodega = 0;
				
				string sSQL = string.Format("Exec VentaConsultaIndividual {0}", idCpaFC);
				SqlDataReader dr = Funcion.rEscalarSQL(cdsSeteo, sSQL, true);
				dr.Read();
				if (dr.HasRows)
				{					
					this.label1.Text = "# NE:"+dr.GetString(13);
					nFacHab = dr.GetString(13);
					this.label13.Text = dr.GetString(6);
					iBodega = dr.GetInt32(10);
					idClint = dr.GetInt32(16);
					this.txtNumCedula.Text = dr.GetValue(17).ToString();
					this.txtNombre.Text = dr.GetValue(18).ToString();
					this.dtFecha.Value = dr.GetDateTime(19);	
				}
				dr.Close();

				#endregion Maestro
				#region adicionales
				string adSQL = string.Format("Exec DatosAd_Compra {0}",idCpaFC);
				SqlDataReader addr = Funcion.rEscalarSQL(cdsSeteo, adSQL, true);
				addr.Read();
				if (addr.HasRows)
				{					
					this.cmbCajero.Value = addr.GetInt32(3);
					idDetFormaPago = addr.GetInt32(4);
				}
				addr.Close();
				#endregion adicionales

				string QryDetFc = "SELECT idDetCompra,idCompra,idArticulo,Articulo,Cantidad,Notas From DetCompra Where idCompra="+idCpaFC.ToString();
				ugDetalleFactura.DataSource = Funcion.dvProcedimiento(cdsSeteo, QryDetFc);
			}
			catch(Exception Exc)
			{
				MessageBox.Show(string.Format("Error al Consultar: {0}", Exc), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
			}
		}

		private void ugArticulos_ClickCellButton(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			// Registrar articulo en detalle de factura
			try
			{
				if(idCpaFC==0)
				{
					MessageBox.Show(string.Format("Error: No se ha registrado nota de entrega.?"), "Aviso...!!!", MessageBoxButtons.OK, MessageBoxIcon.Error);
					this.txtNumCedula.Focus();
				}
				else
				{          
					//idCpaFC = (int)ugFacRelizadas.ActiveRow.Cells["idCompra"].Value;
					int idCltRuc = 0;
					if (this.txtNumCedula.Text.Length!=0) 
					{
						string lcQry = "SELECT idCliente FROM Cliente WHERE Ruc="+this.txtNumCedula.Text.ToString();
						idCltRuc = Funcion.iEscalarSQL(cdsSeteo, lcQry); 
					}
					int idClFc = 0;
					if(idCltRuc>0) idClFc = idCltRuc;
					int idArtFc = (int)ugArticulos.ActiveRow.Cells["idArticulo"].Value;
					decimal preFc = 0;
					preFc = (decimal)ugArticulos.ActiveRow.Cells["Precio1"].Value;
				
					//Validar existencia de Producto
					int Existencia = Funcion.iEscalarSQL(cdsSeteo, string.Format("Exec Existencia_Receta {0}, {1}", idArtFc, NBodga));

					int iCantidad = 0;
					if (iCantidad > Existencia)
					{					
						if (Existencia < 1)
						{
							MessageBox.Show("Este Articulo No Tiene Existencia", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
						}
						else
						{
							e.Cell.Value = 1;
							MessageBox.Show(string.Format("Existencia Actual : {0}", Existencia), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
						}				
					}
					else if(Existencia==0)
					{
						e.Cell.Value = 1;
						MessageBox.Show(string.Format("Existencia Actual : {0}", Existencia), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					}
					else
					{
						string SP_Insert = string.Format("Exec GrabaDetalle_FC_Habitacion {0},{1},{2},{3},{4}", idCpaFC,idClFc,idArtFc,1,preFc);
						Funcion.EjecutaSQL(cdsSeteo,SP_Insert);
						string QryDetFc = "SELECT idDetCompra,idCompra,idArticulo,Articulo,Cantidad,Notas From DetCompra Where idCompra="+idCpaFC.ToString();
						ugDetalleFactura.DataSource = Funcion.dvProcedimiento(cdsSeteo, QryDetFc);
					}
				}
			}
			catch (Exception ex2) 
			{
				MessageBox.Show(string.Format("Error: {0} {1}", ex2.GetType(), ex2.Message), "Aviso...!!!", MessageBoxButtons.OK, MessageBoxIcon.Error);
			}
		}

		private void cmbBodega_ValueChanged(object sender, System.EventArgs e)
		{
			int idBodg = (int)this.cmbBodega.Value;
			string nHoras = "SELECT Bodega From Bodega Where idBodega="+idBodg.ToString();
			NBodga = Funcion.iEscalarSQL(cdsSeteo, nHoras);
		}

		private void ugDetalleFactura_AfterCellUpdate(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			#region Cantidad
			if (e.Cell.Column.ToString().Equals("Cantidad"))
			{
				if ((e.Cell.Row.Cells["Cantidad"].Value == System.DBNull.Value))
				{
					e.Cell.Row.Cells["Cantidad"].Value = (int)e.Cell.OriginalValue;
				}
				if (true)
				{
					int idArticulo = (int)e.Cell.Row.Cells["idArticulo"].Value;
					int Existencia = Funcion.iEscalarSQL(cdsSeteo, string.Format("Select dbo.ExistenciasInventarioLotte({0}, {1})", idArticulo, NBodga));
					int iCantidad = 0;

					foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.ugDetalleFactura.Rows.All)
					{
						if (idArticulo == (int)dr.Cells["idArticulo"].Value) iCantidad = iCantidad + (int)dr.Cells["Cantidad"].Value;
					}

					if(idArticulo==3422 || idArticulo==3423) 
					{	
						if (e.Cell.Column.ToString() == "Cantidad") iCantidad=100;
					}
					else
					{
						if (iCantidad > Existencia)
						{					
							if (Existencia < 1)
							{
								MessageBox.Show("Este Articulo No Tiene Existencia", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
							}
							else
							{
								e.Cell.Value = 1;
								MessageBox.Show(string.Format("Existencia Actual : {0}", Existencia), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
							}				
						}
					}
				}
			}
			#endregion Cantidad
		}

		private void ugDetalleFactura_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (ugDetalleFactura.ActiveCell != null)
			{
				try
				{				
					switch(e.KeyValue)
					{
							#region Enter
						case (int) Keys.Enter:
							ugDetalleFactura.PerformAction(UltraGridAction.ExitEditMode, false, false);
							ugDetalleFactura.PerformAction(UltraGridAction.NextCellByTab, false, false);
							e.Handled = true;
							ugDetalleFactura.PerformAction(UltraGridAction.EnterEditMode, false, false);
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

		private void frmNotasEntrega_Closed(object sender, System.EventArgs e)
		{
			if (swCambio==1) Funcion.EjecutaSQL(cdsSeteo, string.Format("Exec Anular_NE {0},{1}", idCpaFC,0));
		}

		private void btnCancelar_Click(object sender, System.EventArgs e)
		{
			/*Cancelar Nota de Entrega*/
			if (DialogResult.Yes ==	MessageBox.Show("Desea Cancelar N/E.?", "Point Technology", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2)) 
			{
				Funcion.EjecutaSQL(cdsSeteo, string.Format("Exec Anular_NE {0},{1}", idCpaFC,0));
				swCambio = 0;
				this.btnCancelar.Enabled = false;
			}
		}

		private void btnAnular_Click_1(object sender, System.EventArgs e)
		{
			/*Cancelar Nota de Entrega*/
			if (DialogResult.Yes ==	MessageBox.Show("Desea Anular N/E.?", "Point Technology", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2)) 
			{
				Funcion.EjecutaSQL(cdsSeteo, string.Format("Exec Anular_NE {0},{1}", idCpaFC,1));
				swCambio = 0;
				this.btnAnular.Enabled = false;
			}
		}

		private void ugDetalleFactura_BeforeRowsDeleted(object sender, Infragistics.Win.UltraWinGrid.BeforeRowsDeletedEventArgs e)
		{
			for (int i = 0; i < e.Rows.Length; i++)
			{
				if ((int)e.Rows[i].Cells["idDetCompra"].Value > 0)
				{	
					if (DialogResult.Yes ==	MessageBox.Show("Desea Eliminar Registro.?", "Point Technology", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2)) 
					{
						Funcion.EjecutaSQL(cdsSeteo, string.Format("Delete From DetCompra Where idDetCompra = {0}", (int)e.Rows[i].Cells["idDetCompra"].Value));
						e.DisplayPromptMsg = false;
					}
					else e.Cancel = true;
				}
			}
		}

		private void ugArticulos_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}

	}
}
