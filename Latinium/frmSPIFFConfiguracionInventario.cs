using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using DevExpress.XtraEditors;
using System.Data;
using System.Data.SqlClient;
using CrystalDecisions.Shared;
using System.Net;
using System.Web;
using System.Xml;
using System.Xml.Schema;
using System.IO;
using System.Diagnostics;
using Infragistics.Shared;
using Infragistics.Win;
using Infragistics.Win.UltraWinGrid;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de frmSPIFFConfiguracionInventario.
	/// </summary>
	public class frmSPIFFConfiguracionInventario : System.Windows.Forms.Form
	{
		private System.Windows.Forms.GroupBox groupBox1;
		private System.Windows.Forms.GroupBox groupBox5;
		private System.Windows.Forms.Button btnConsultar;
		public System.Windows.Forms.Label label12;
		private System.Windows.Forms.Label lblContador;
		private System.Windows.Forms.GroupBox groupBox2;
		private System.Windows.Forms.Label lblBodega;
		private System.Windows.Forms.Label label53;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtDesde;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtHasta;
		private System.Windows.Forms.Button btnSalir;
		private System.Windows.Forms.Label label1;
		private System.Windows.Forms.Label label6;
		public Infragistics.Win.UltraWinEditors.UltraTextEditor txtNotas;
		private System.Windows.Forms.Label label5;
		private Infragistics.Win.UltraWinGrid.UltraGrid ugdSPIFFSubGrupoRangoCab;
		public Infragistics.Win.UltraWinGrid.UltraCombo cmbSPIFFEstado;
		private System.Windows.Forms.ToolTip toolTip1;
		private C1.Data.C1DataSet cdsSeteoF;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource udsSPIFFSubGrupoRangoCab;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource udsSPIFFEstado;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtBuscar;
		private System.Windows.Forms.Button btnDuplicar;
		public Infragistics.Win.UltraWinEditors.UltraTextEditor txtNombre;
		private System.ComponentModel.IContainer components;

		public frmSPIFFConfiguracionInventario()
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
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmSPIFFConfiguracionInventario));
			Infragistics.Win.Appearance appearance1 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idSPIFFSubGrupoRangoCab");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idSPIFFEstado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Desde");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Hasta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nota");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Seleccionar");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("sidSPIFFEstado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("btnGestion", 0);
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("btnProforma", 1);
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn11 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Row", 2);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn12 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("btnAnadir", 3);
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn13 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("btnAnular", 4);
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance11 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance12 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance13 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance14 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance15 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idSPIFFSubGrupoRangoCab");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idSPIFFEstado");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn3 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn4 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Desde");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn5 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Hasta");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn6 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nota");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn7 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Seleccionar");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn8 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("sidSPIFFEstado");
			Infragistics.Win.Appearance appearance16 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn14 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idSPIFFEstado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn15 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Descripcion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn9 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idSPIFFEstado");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn10 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Descripcion");
			Infragistics.Win.Appearance appearance17 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance18 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance19 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton1 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.Appearance appearance20 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton2 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			this.groupBox1 = new System.Windows.Forms.GroupBox();
			this.groupBox5 = new System.Windows.Forms.GroupBox();
			this.btnConsultar = new System.Windows.Forms.Button();
			this.label12 = new System.Windows.Forms.Label();
			this.txtBuscar = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.lblContador = new System.Windows.Forms.Label();
			this.ugdSPIFFSubGrupoRangoCab = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.udsSPIFFSubGrupoRangoCab = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.groupBox2 = new System.Windows.Forms.GroupBox();
			this.cmbSPIFFEstado = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.udsSPIFFEstado = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.label5 = new System.Windows.Forms.Label();
			this.label6 = new System.Windows.Forms.Label();
			this.txtNotas = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label1 = new System.Windows.Forms.Label();
			this.txtNombre = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.lblBodega = new System.Windows.Forms.Label();
			this.label53 = new System.Windows.Forms.Label();
			this.dtDesde = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.dtHasta = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.btnSalir = new System.Windows.Forms.Button();
			this.btnDuplicar = new System.Windows.Forms.Button();
			this.toolTip1 = new System.Windows.Forms.ToolTip(this.components);
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.groupBox1.SuspendLayout();
			this.groupBox5.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.txtBuscar)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ugdSPIFFSubGrupoRangoCab)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.udsSPIFFSubGrupoRangoCab)).BeginInit();
			this.groupBox2.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.cmbSPIFFEstado)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.udsSPIFFEstado)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNotas)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNombre)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtDesde)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtHasta)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			this.SuspendLayout();
			// 
			// groupBox1
			// 
			this.groupBox1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.groupBox1.Controls.Add(this.groupBox5);
			this.groupBox1.Controls.Add(this.ugdSPIFFSubGrupoRangoCab);
			this.groupBox1.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.groupBox1.Location = new System.Drawing.Point(16, 176);
			this.groupBox1.Name = "groupBox1";
			this.groupBox1.Size = new System.Drawing.Size(888, 304);
			this.groupBox1.TabIndex = 984;
			this.groupBox1.TabStop = false;
			this.groupBox1.Text = "Campañas";
			// 
			// groupBox5
			// 
			this.groupBox5.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.groupBox5.Controls.Add(this.btnConsultar);
			this.groupBox5.Controls.Add(this.label12);
			this.groupBox5.Controls.Add(this.txtBuscar);
			this.groupBox5.Controls.Add(this.lblContador);
			this.groupBox5.Location = new System.Drawing.Point(16, 24);
			this.groupBox5.Name = "groupBox5";
			this.groupBox5.Size = new System.Drawing.Size(848, 64);
			this.groupBox5.TabIndex = 1055;
			this.groupBox5.TabStop = false;
			this.groupBox5.Text = "Búsqueda";
			// 
			// btnConsultar
			// 
			this.btnConsultar.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(230)), ((System.Byte)(230)), ((System.Byte)(230)));
			this.btnConsultar.CausesValidation = false;
			this.btnConsultar.Image = ((System.Drawing.Image)(resources.GetObject("btnConsultar.Image")));
			this.btnConsultar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnConsultar.Location = new System.Drawing.Point(288, 24);
			this.btnConsultar.Name = "btnConsultar";
			this.btnConsultar.Size = new System.Drawing.Size(78, 24);
			this.btnConsultar.TabIndex = 1059;
			this.btnConsultar.Text = "&Consultar";
			this.btnConsultar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.toolTip1.SetToolTip(this.btnConsultar, "Buscar por nombre");
			this.btnConsultar.Click += new System.EventHandler(this.btnConsultar_Click);
			// 
			// label12
			// 
			this.label12.AutoSize = true;
			this.label12.Location = new System.Drawing.Point(16, 24);
			this.label12.Name = "label12";
			this.label12.Size = new System.Drawing.Size(39, 16);
			this.label12.TabIndex = 1058;
			this.label12.Text = "Buscar";
			this.label12.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtBuscar
			// 
			appearance1.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtBuscar.Appearance = appearance1;
			this.txtBuscar.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtBuscar.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtBuscar.Location = new System.Drawing.Point(72, 24);
			this.txtBuscar.Name = "txtBuscar";
			this.txtBuscar.Size = new System.Drawing.Size(200, 21);
			this.txtBuscar.TabIndex = 1057;
			// 
			// lblContador
			// 
			this.lblContador.AutoSize = true;
			this.lblContador.Font = new System.Drawing.Font("Tahoma", 10F);
			this.lblContador.ForeColor = System.Drawing.Color.Firebrick;
			this.lblContador.Location = new System.Drawing.Point(376, 24);
			this.lblContador.Name = "lblContador";
			this.lblContador.Size = new System.Drawing.Size(0, 20);
			this.lblContador.TabIndex = 1056;
			this.lblContador.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// ugdSPIFFSubGrupoRangoCab
			// 
			this.ugdSPIFFSubGrupoRangoCab.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.ugdSPIFFSubGrupoRangoCab.Cursor = System.Windows.Forms.Cursors.Default;
			this.ugdSPIFFSubGrupoRangoCab.DataSource = this.udsSPIFFSubGrupoRangoCab;
			appearance2.BackColor = System.Drawing.Color.White;
			this.ugdSPIFFSubGrupoRangoCab.DisplayLayout.Appearance = appearance2;
			ultraGridBand1.AutoPreviewEnabled = true;
			ultraGridColumn1.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn1.RowLayoutColumnInfo.PreferredCellSize = new System.Drawing.Size(27, 0);
			ultraGridColumn2.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn2.Header.VisiblePosition = 2;
			ultraGridColumn2.Hidden = true;
			ultraGridColumn3.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn3.Header.VisiblePosition = 4;
			ultraGridColumn3.RowLayoutColumnInfo.OriginX = 8;
			ultraGridColumn3.RowLayoutColumnInfo.OriginY = 0;
			ultraGridColumn3.RowLayoutColumnInfo.PreferredCellSize = new System.Drawing.Size(195, 0);
			ultraGridColumn3.RowLayoutColumnInfo.SpanX = 2;
			ultraGridColumn3.RowLayoutColumnInfo.SpanY = 2;
			ultraGridColumn4.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn4.Format = "dd/MM/yyyy";
			ultraGridColumn4.Header.VisiblePosition = 5;
			ultraGridColumn4.RowLayoutColumnInfo.OriginX = 12;
			ultraGridColumn4.RowLayoutColumnInfo.OriginY = 0;
			ultraGridColumn4.RowLayoutColumnInfo.SpanX = 2;
			ultraGridColumn4.RowLayoutColumnInfo.SpanY = 2;
			ultraGridColumn5.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn5.Format = "dd/MM/yyyy";
			ultraGridColumn5.Header.VisiblePosition = 6;
			ultraGridColumn5.RowLayoutColumnInfo.OriginX = 14;
			ultraGridColumn5.RowLayoutColumnInfo.OriginY = 0;
			ultraGridColumn5.RowLayoutColumnInfo.SpanX = 2;
			ultraGridColumn5.RowLayoutColumnInfo.SpanY = 2;
			ultraGridColumn6.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn6.Header.VisiblePosition = 7;
			ultraGridColumn6.RowLayoutColumnInfo.OriginX = 10;
			ultraGridColumn6.RowLayoutColumnInfo.OriginY = 0;
			ultraGridColumn6.RowLayoutColumnInfo.PreferredCellSize = new System.Drawing.Size(255, 0);
			ultraGridColumn6.RowLayoutColumnInfo.SpanX = 2;
			ultraGridColumn6.RowLayoutColumnInfo.SpanY = 2;
			ultraGridColumn7.Header.Caption = "...";
			ultraGridColumn7.Header.VisiblePosition = 8;
			ultraGridColumn7.RowLayoutColumnInfo.OriginX = 2;
			ultraGridColumn7.RowLayoutColumnInfo.OriginY = 0;
			ultraGridColumn7.RowLayoutColumnInfo.PreferredCellSize = new System.Drawing.Size(24, 0);
			ultraGridColumn7.RowLayoutColumnInfo.SpanX = 2;
			ultraGridColumn7.RowLayoutColumnInfo.SpanY = 2;
			ultraGridColumn8.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn8.Header.Caption = "Estado";
			ultraGridColumn8.Header.VisiblePosition = 9;
			ultraGridColumn8.RowLayoutColumnInfo.OriginX = 16;
			ultraGridColumn8.RowLayoutColumnInfo.OriginY = 0;
			ultraGridColumn8.RowLayoutColumnInfo.SpanX = 2;
			ultraGridColumn8.RowLayoutColumnInfo.SpanY = 2;
			ultraGridColumn9.ButtonDisplayStyle = Infragistics.Win.UltraWinGrid.ButtonDisplayStyle.Always;
			appearance3.Image = ((object)(resources.GetObject("appearance3.Image")));
			appearance3.ImageHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn9.CellButtonAppearance = appearance3;
			ultraGridColumn9.Header.Caption = "...";
			ultraGridColumn9.Header.VisiblePosition = 1;
			ultraGridColumn9.Hidden = true;
			ultraGridColumn9.RowLayoutColumnInfo.OriginX = 0;
			ultraGridColumn9.RowLayoutColumnInfo.OriginY = 0;
			ultraGridColumn9.RowLayoutColumnInfo.SpanX = 2;
			ultraGridColumn9.RowLayoutColumnInfo.SpanY = 2;
			ultraGridColumn9.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.Button;
			ultraGridColumn9.Width = 30;
			ultraGridColumn10.ButtonDisplayStyle = Infragistics.Win.UltraWinGrid.ButtonDisplayStyle.Always;
			appearance4.Image = ((object)(resources.GetObject("appearance4.Image")));
			appearance4.ImageHAlign = Infragistics.Win.HAlign.Center;
			appearance4.ImageVAlign = Infragistics.Win.VAlign.Middle;
			ultraGridColumn10.CellButtonAppearance = appearance4;
			ultraGridColumn10.Header.Caption = "Proforma";
			ultraGridColumn10.Header.VisiblePosition = 3;
			ultraGridColumn10.Hidden = true;
			ultraGridColumn10.RowLayoutColumnInfo.OriginX = 2;
			ultraGridColumn10.RowLayoutColumnInfo.OriginY = 0;
			ultraGridColumn10.RowLayoutColumnInfo.SpanX = 2;
			ultraGridColumn10.RowLayoutColumnInfo.SpanY = 2;
			ultraGridColumn10.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.Button;
			ultraGridColumn10.Width = 60;
			ultraGridColumn11.AutoEdit = false;
			ultraGridColumn11.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn11.Header.Caption = "N°";
			ultraGridColumn11.Header.VisiblePosition = 10;
			ultraGridColumn11.RowLayoutColumnInfo.OriginX = 0;
			ultraGridColumn11.RowLayoutColumnInfo.OriginY = 0;
			ultraGridColumn11.RowLayoutColumnInfo.PreferredCellSize = new System.Drawing.Size(26, 0);
			ultraGridColumn11.RowLayoutColumnInfo.SpanX = 2;
			ultraGridColumn11.RowLayoutColumnInfo.SpanY = 2;
			appearance5.Image = ((object)(resources.GetObject("appearance5.Image")));
			appearance5.ImageHAlign = Infragistics.Win.HAlign.Center;
			appearance5.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn12.CellAppearance = appearance5;
			appearance6.Image = ((object)(resources.GetObject("appearance6.Image")));
			appearance6.ImageHAlign = Infragistics.Win.HAlign.Center;
			appearance6.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn12.CellButtonAppearance = appearance6;
			ultraGridColumn12.Header.Caption = "...";
			ultraGridColumn12.Header.VisiblePosition = 11;
			ultraGridColumn12.RowLayoutColumnInfo.OriginX = 4;
			ultraGridColumn12.RowLayoutColumnInfo.OriginY = 0;
			ultraGridColumn12.RowLayoutColumnInfo.PreferredCellSize = new System.Drawing.Size(24, 0);
			ultraGridColumn12.RowLayoutColumnInfo.SpanX = 2;
			ultraGridColumn12.RowLayoutColumnInfo.SpanY = 2;
			ultraGridColumn12.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.Button;
			appearance7.Image = ((object)(resources.GetObject("appearance7.Image")));
			appearance7.ImageHAlign = Infragistics.Win.HAlign.Center;
			appearance7.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn13.CellAppearance = appearance7;
			appearance8.Image = ((object)(resources.GetObject("appearance8.Image")));
			appearance8.ImageHAlign = Infragistics.Win.HAlign.Center;
			appearance8.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn13.CellButtonAppearance = appearance8;
			ultraGridColumn13.Header.Caption = "...";
			ultraGridColumn13.Header.VisiblePosition = 12;
			ultraGridColumn13.RowLayoutColumnInfo.OriginX = 6;
			ultraGridColumn13.RowLayoutColumnInfo.OriginY = 0;
			ultraGridColumn13.RowLayoutColumnInfo.PreferredCellSize = new System.Drawing.Size(25, 0);
			ultraGridColumn13.RowLayoutColumnInfo.SpanX = 2;
			ultraGridColumn13.RowLayoutColumnInfo.SpanY = 2;
			ultraGridColumn13.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.Button;
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
																										 ultraGridColumn13});
			ultraGridBand1.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			ultraGridBand1.SummaryFooterCaption = "Sumatoria:";
			ultraGridBand1.UseRowLayout = true;
			this.ugdSPIFFSubGrupoRangoCab.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.ugdSPIFFSubGrupoRangoCab.DisplayLayout.GroupByBox.Hidden = true;
			appearance9.ForeColor = System.Drawing.Color.Black;
			appearance9.ForeColorDisabled = System.Drawing.Color.Black;
			this.ugdSPIFFSubGrupoRangoCab.DisplayLayout.Override.ActiveRowAppearance = appearance9;
			this.ugdSPIFFSubGrupoRangoCab.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.Yes;
			this.ugdSPIFFSubGrupoRangoCab.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.True;
			this.ugdSPIFFSubGrupoRangoCab.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance10.BackColor = System.Drawing.Color.Transparent;
			this.ugdSPIFFSubGrupoRangoCab.DisplayLayout.Override.CardAreaAppearance = appearance10;
			appearance11.ForeColor = System.Drawing.Color.Black;
			appearance11.ForeColorDisabled = System.Drawing.Color.Black;
			this.ugdSPIFFSubGrupoRangoCab.DisplayLayout.Override.CellAppearance = appearance11;
			appearance12.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance12.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance12.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance12.FontData.BoldAsString = "True";
			appearance12.FontData.Name = "Arial";
			appearance12.FontData.SizeInPoints = 8F;
			appearance12.ForeColor = System.Drawing.Color.White;
			appearance12.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.ugdSPIFFSubGrupoRangoCab.DisplayLayout.Override.HeaderAppearance = appearance12;
			this.ugdSPIFFSubGrupoRangoCab.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti;
			appearance13.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance13.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance13.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ugdSPIFFSubGrupoRangoCab.DisplayLayout.Override.RowAlternateAppearance = appearance13;
			appearance14.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance14.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance14.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ugdSPIFFSubGrupoRangoCab.DisplayLayout.Override.RowSelectorAppearance = appearance14;
			appearance15.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance15.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance15.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ugdSPIFFSubGrupoRangoCab.DisplayLayout.Override.SelectedRowAppearance = appearance15;
			this.ugdSPIFFSubGrupoRangoCab.DisplayLayout.ViewStyleBand = Infragistics.Win.UltraWinGrid.ViewStyleBand.OutlookGroupBy;
			this.ugdSPIFFSubGrupoRangoCab.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.ugdSPIFFSubGrupoRangoCab.Location = new System.Drawing.Point(16, 104);
			this.ugdSPIFFSubGrupoRangoCab.Name = "ugdSPIFFSubGrupoRangoCab";
			this.ugdSPIFFSubGrupoRangoCab.RightToLeft = System.Windows.Forms.RightToLeft.No;
			this.ugdSPIFFSubGrupoRangoCab.RowUpdateCancelAction = Infragistics.Win.UltraWinGrid.RowUpdateCancelAction.RetainDataAndActivation;
			this.ugdSPIFFSubGrupoRangoCab.Size = new System.Drawing.Size(848, 184);
			this.ugdSPIFFSubGrupoRangoCab.TabIndex = 981;
			this.ugdSPIFFSubGrupoRangoCab.ClickCellButton += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.ugdSPIFFSubGrupoRangoCab_ClickCellButton);
			this.ugdSPIFFSubGrupoRangoCab.AfterCellUpdate += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.ugdSPIFFSubGrupoRangoCab_AfterCellUpdate);
			// 
			// udsSPIFFSubGrupoRangoCab
			// 
			ultraDataColumn1.DataType = typeof(int);
			ultraDataColumn2.DataType = typeof(int);
			ultraDataColumn4.DataType = typeof(System.DateTime);
			ultraDataColumn5.DataType = typeof(System.DateTime);
			ultraDataColumn7.DataType = typeof(bool);
			this.udsSPIFFSubGrupoRangoCab.Band.Columns.AddRange(new object[] {
																																				 ultraDataColumn1,
																																				 ultraDataColumn2,
																																				 ultraDataColumn3,
																																				 ultraDataColumn4,
																																				 ultraDataColumn5,
																																				 ultraDataColumn6,
																																				 ultraDataColumn7,
																																				 ultraDataColumn8});
			// 
			// groupBox2
			// 
			this.groupBox2.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.groupBox2.Controls.Add(this.cmbSPIFFEstado);
			this.groupBox2.Controls.Add(this.label5);
			this.groupBox2.Controls.Add(this.label6);
			this.groupBox2.Controls.Add(this.txtNotas);
			this.groupBox2.Controls.Add(this.label1);
			this.groupBox2.Controls.Add(this.txtNombre);
			this.groupBox2.Controls.Add(this.lblBodega);
			this.groupBox2.Controls.Add(this.label53);
			this.groupBox2.Controls.Add(this.dtDesde);
			this.groupBox2.Controls.Add(this.dtHasta);
			this.groupBox2.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.groupBox2.Location = new System.Drawing.Point(16, 16);
			this.groupBox2.Name = "groupBox2";
			this.groupBox2.Size = new System.Drawing.Size(776, 152);
			this.groupBox2.TabIndex = 992;
			this.groupBox2.TabStop = false;
			this.groupBox2.Text = "Datos principales";
			this.groupBox2.Enter += new System.EventHandler(this.groupBox2_Enter);
			// 
			// cmbSPIFFEstado
			// 
			appearance16.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbSPIFFEstado.Appearance = appearance16;
			this.cmbSPIFFEstado.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbSPIFFEstado.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbSPIFFEstado.DataSource = this.udsSPIFFEstado;
			ultraGridColumn14.Header.VisiblePosition = 0;
			ultraGridColumn14.Hidden = true;
			ultraGridColumn15.Header.Caption = "Descripción";
			ultraGridColumn15.Header.VisiblePosition = 1;
			ultraGridColumn15.Width = 279;
			ultraGridBand2.Columns.AddRange(new object[] {
																										 ultraGridColumn14,
																										 ultraGridColumn15});
			this.cmbSPIFFEstado.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			this.cmbSPIFFEstado.DisplayMember = "Descripcion";
			this.cmbSPIFFEstado.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbSPIFFEstado.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbSPIFFEstado.Location = new System.Drawing.Point(480, 24);
			this.cmbSPIFFEstado.Name = "cmbSPIFFEstado";
			this.cmbSPIFFEstado.Size = new System.Drawing.Size(280, 21);
			this.cmbSPIFFEstado.TabIndex = 1040;
			this.cmbSPIFFEstado.ValueMember = "idSPIFFEstado";
			// 
			// udsSPIFFEstado
			// 
			ultraDataColumn9.DataType = typeof(int);
			this.udsSPIFFEstado.Band.Columns.AddRange(new object[] {
																															 ultraDataColumn9,
																															 ultraDataColumn10});
			// 
			// label5
			// 
			this.label5.AutoSize = true;
			this.label5.Location = new System.Drawing.Point(408, 24);
			this.label5.Name = "label5";
			this.label5.Size = new System.Drawing.Size(50, 16);
			this.label5.TabIndex = 1039;
			this.label5.Text = "* Estado:";
			this.label5.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label6
			// 
			this.label6.AutoSize = true;
			this.label6.Location = new System.Drawing.Point(8, 104);
			this.label6.Name = "label6";
			this.label6.Size = new System.Drawing.Size(39, 16);
			this.label6.TabIndex = 1038;
			this.label6.Text = "* Nota:";
			this.label6.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtNotas
			// 
			appearance17.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtNotas.Appearance = appearance17;
			this.txtNotas.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtNotas.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNotas.Location = new System.Drawing.Point(80, 88);
			this.txtNotas.Multiline = true;
			this.txtNotas.Name = "txtNotas";
			this.txtNotas.Size = new System.Drawing.Size(680, 48);
			this.txtNotas.TabIndex = 1037;
			// 
			// label1
			// 
			this.label1.AutoSize = true;
			this.label1.Location = new System.Drawing.Point(8, 24);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(55, 16);
			this.label1.TabIndex = 1036;
			this.label1.Text = "* Nombre:";
			this.label1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtNombre
			// 
			appearance18.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtNombre.Appearance = appearance18;
			this.txtNombre.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtNombre.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNombre.Location = new System.Drawing.Point(80, 24);
			this.txtNombre.Name = "txtNombre";
			this.txtNombre.Size = new System.Drawing.Size(320, 21);
			this.txtNombre.TabIndex = 1035;
			// 
			// lblBodega
			// 
			this.lblBodega.AutoSize = true;
			this.lblBodega.BackColor = System.Drawing.Color.Transparent;
			this.lblBodega.Location = new System.Drawing.Point(8, 64);
			this.lblBodega.Name = "lblBodega";
			this.lblBodega.Size = new System.Drawing.Size(41, 16);
			this.lblBodega.TabIndex = 993;
			this.lblBodega.Text = "* Inicio:";
			this.lblBodega.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label53
			// 
			this.label53.AutoSize = true;
			this.label53.BackColor = System.Drawing.Color.Transparent;
			this.label53.Location = new System.Drawing.Point(408, 64);
			this.label53.Name = "label53";
			this.label53.Size = new System.Drawing.Size(54, 16);
			this.label53.TabIndex = 991;
			this.label53.Text = "* Finaliza:";
			this.label53.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// dtDesde
			// 
			appearance19.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtDesde.Appearance = appearance19;
			this.dtDesde.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton1.Caption = "Today";
			this.dtDesde.DateButtons.Add(dateButton1);
			this.dtDesde.Format = "dd/MM/yyyy";
			this.dtDesde.Location = new System.Drawing.Point(80, 56);
			this.dtDesde.Name = "dtDesde";
			this.dtDesde.NonAutoSizeHeight = 23;
			this.dtDesde.Size = new System.Drawing.Size(320, 21);
			this.dtDesde.SpinButtonsVisible = true;
			this.dtDesde.TabIndex = 989;
			this.dtDesde.Value = ((object)(resources.GetObject("dtDesde.Value")));
			// 
			// dtHasta
			// 
			appearance20.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtHasta.Appearance = appearance20;
			this.dtHasta.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton2.Caption = "Today";
			this.dtHasta.DateButtons.Add(dateButton2);
			this.dtHasta.Format = "dd/MM/yyyy";
			this.dtHasta.Location = new System.Drawing.Point(480, 56);
			this.dtHasta.Name = "dtHasta";
			this.dtHasta.NonAutoSizeHeight = 23;
			this.dtHasta.Size = new System.Drawing.Size(280, 21);
			this.dtHasta.SpinButtonsVisible = true;
			this.dtHasta.TabIndex = 990;
			this.dtHasta.Value = ((object)(resources.GetObject("dtHasta.Value")));
			// 
			// btnSalir
			// 
			this.btnSalir.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btnSalir.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(230)), ((System.Byte)(230)), ((System.Byte)(230)));
			this.btnSalir.CausesValidation = false;
			this.btnSalir.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.btnSalir.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(50)), ((System.Byte)(50)), ((System.Byte)(50)));
			this.btnSalir.Image = ((System.Drawing.Image)(resources.GetObject("btnSalir.Image")));
			this.btnSalir.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnSalir.Location = new System.Drawing.Point(816, 64);
			this.btnSalir.Name = "btnSalir";
			this.btnSalir.Size = new System.Drawing.Size(80, 23);
			this.btnSalir.TabIndex = 991;
			this.btnSalir.Text = "&Salir";
			this.btnSalir.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnSalir.Click += new System.EventHandler(this.btnSalir_Click);
			// 
			// btnDuplicar
			// 
			this.btnDuplicar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btnDuplicar.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(230)), ((System.Byte)(230)), ((System.Byte)(230)));
			this.btnDuplicar.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(50)), ((System.Byte)(50)), ((System.Byte)(50)));
			this.btnDuplicar.Image = ((System.Drawing.Image)(resources.GetObject("btnDuplicar.Image")));
			this.btnDuplicar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnDuplicar.Location = new System.Drawing.Point(816, 24);
			this.btnDuplicar.Name = "btnDuplicar";
			this.btnDuplicar.Size = new System.Drawing.Size(82, 23);
			this.btnDuplicar.TabIndex = 990;
			this.btnDuplicar.Text = "&Duplicar";
			this.btnDuplicar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.toolTip1.SetToolTip(this.btnDuplicar, "Tiene el fin de que replicar la promoción con valores pre-cargados y no cargar to" +
				"do de nuevo.");
			this.btnDuplicar.Click += new System.EventHandler(this.btnDuplicar_Click);
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
			// frmSPIFFConfiguracionInventario
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(920, 494);
			this.Controls.Add(this.groupBox2);
			this.Controls.Add(this.btnSalir);
			this.Controls.Add(this.btnDuplicar);
			this.Controls.Add(this.groupBox1);
			this.Name = "frmSPIFFConfiguracionInventario";
			this.StartPosition = System.Windows.Forms.FormStartPosition.Manual;
			this.Text = "Configuración SPIFF inventario";
			this.Load += new System.EventHandler(this.frmSPIFFConfiguracionInventario_Load);
			this.groupBox1.ResumeLayout(false);
			this.groupBox5.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.txtBuscar)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ugdSPIFFSubGrupoRangoCab)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.udsSPIFFSubGrupoRangoCab)).EndInit();
			this.groupBox2.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.cmbSPIFFEstado)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.udsSPIFFEstado)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNotas)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNombre)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtDesde)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtHasta)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		static public int idSPIFFRango1 = 1;
		static public int idSPIFFRango2 = 2;
		static public int idSPIFFRango3 = 3;

		static public int idSPIFFEstadoEnProceso = 1;
		static public int idSPIFFEstadoFinaliza = 2;
		static public int idSPIFFEstadoPagado = 3;
		static public int idSPIFFEstadoAnulado = 4;

		private void groupBox2_Enter(object sender, System.EventArgs e)
		{
		
		}

		private void frmSPIFFConfiguracionInventario_Load(object sender, System.EventArgs e)
		{
			SeguridadAcceso();
			SPIFFEstadoCon();
			Consulta();
		}

		private void SeguridadAcceso()
		{
			string ambiente = "PRO";

			if(ambiente == "PRO")
			{
				if (!Funcion.Permiso("2080", cdsSeteoF))
				{
					MessageBox.Show("No puede ingresar a " + this.Text, "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					BeginInvoke(new MethodInvoker(UnloadMe));
					return;
				}
			}
		}

		private void UnloadMe()
		{
			this.Close();
		}

		private void SPIFFEstadoCon()
		{
			string squery = String.Format("Exec SPIFFEstadoCon");
			this.cmbSPIFFEstado.DataSource = Funcion.dvProcedimiento(cdsSeteoF, squery);
		}

		private void btnConsultar_Click(object sender, System.EventArgs e)
		{
			Consulta();
		}

		private void Consulta()
		{
			string sSQLPa = string.Format("Exec SPIFFSubGrupoRangoCabCon '{0}'", this.txtBuscar.Text); //this.txtBusqueda.Text.ToString()
			FuncionesProcedimientos.dsGeneralModulo(sSQLPa, this.ugdSPIFFSubGrupoRangoCab);
			//Label de registros
			this.lblContador.Text = this.ugdSPIFFSubGrupoRangoCab.Rows.Count + " REGISTROS ENCONTRADOS";
			//COntador de filas
			this.ugdSPIFFSubGrupoRangoCab = Funcion.ContadorFilas(ugdSPIFFSubGrupoRangoCab, "Row");
		}

		private void btnDuplicar_Click(object sender, System.EventArgs e)
		{
			SPIFFSubGrupoRangoCabDup();
		}

		private void SPIFFSubGrupoRangoCabDup()
		{
			#region Validación
			if(!Validacion.vbTexto(this.txtNombre, 1, 500, "Nombre"))return;
			if(!Validacion.vbComboVacio(this.cmbSPIFFEstado, "Seleccione un estado.")) return;
			if(!Validacion.vbValidaFechasDesdeHasta(this.dtDesde, this.dtHasta, "Inicio", "Finaliza")) return;
			if(!Validacion.vbTexto(this.txtNotas, 1, 2000, "Nota"))return;

			int iContador = 0;
			string squery = String.Format("Select Count(*) From SPIFFSubGrupoRangoCab where Seleccionar = 1");
			iContador = Funcion.iEscalarSQL(cdsSeteoF,squery);

			if(iContador == 0)
			{
				MessageBox.Show(string.Format("Se debe seleccionar un registro para poder duplicarlo."), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.ugdSPIFFSubGrupoRangoCab.Focus();				
				return;
			}
			else if(iContador > 1)
			{
				MessageBox.Show(string.Format("Se debe seleccionar solo un registro para poder duplicarlo."), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.ugdSPIFFSubGrupoRangoCab.Focus();				
				return;
			}

			// Se debe validar que no exista otro nombre similar, a otro SPIFF
			if(this.txtNombre.Text != "")
			{
				iContador = 0;
				squery = String.Format("EXEC SPIFFSubGrupoRangoCabNombreVal '{0}', {1}", this.txtNombre.Text, 0); //0 no tiene id todavía
				iContador = Funcion.iEscalarSQL(cdsSeteoF,squery);

				if(iContador > 0)
				{
					MessageBox.Show(string.Format("El nombre ya se encuentra registrado, intente con otra descripción."), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					this.txtNombre.Focus();				
					return;
				}
			}

			string FechaInicio = "";
			string FechaFinaliza = "";
			FechaInicio = Convert.ToDateTime(this.dtDesde.Value).ToString("yyyyMMdd");
			FechaFinaliza = Convert.ToDateTime(this.dtHasta.Value).ToString("yyyyMMdd");

			// Validar que la fecha de inicio, no se encuentra en otros peridos de SPIFF
			if(FechaInicio != "")
			{
				iContador = 0;
				squery = String.Format("EXEC SPIFFSubGrupoRangoCabFechaVal '{0}', {1}", FechaInicio, 0); //0 no tiene id todavía
				iContador = Funcion.iEscalarSQL(cdsSeteoF,squery);

				if(iContador > 0)
				{
					MessageBox.Show(string.Format("La fecha de inicio, ya corresponde a otro SPIFF generado."), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					this.dtDesde.Focus();				
					return;
				}
			}

			if(FechaFinaliza != "")
			{
				iContador = 0;
				squery = String.Format("EXEC SPIFFSubGrupoRangoCabFechaVal '{0}', {1}", FechaFinaliza, 0); //0 no tiene id todavía
				iContador = Funcion.iEscalarSQL(cdsSeteoF,squery);

				if(iContador > 0)
				{
					MessageBox.Show(string.Format("La fecha finaliza, ya corresponde a otro SPIFF generado."), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					this.dtHasta.Focus();				
					return;
				}
			}
			#endregion Validación
			// Mostrar el modal de "procesando"
			MenuLatinium.MostrarModalProcesando();
			#region Duplicación SPIFF Inventario
			using (SqlConnection oConexion = new SqlConnection("Password=Bl45o6$9;Persist Security Info=True;User ID= " + MenuLatinium.stUsuario + ";Initial Catalog=" + MenuLatinium.stDirFacturacion + ";Data Source=" + MenuLatinium.stDirServidor + "; max pool size=1024"))					
			{
				oConexion.Open();

				SqlTransaction oTransaction = null;
				SqlCommand oCmdActualiza = oConexion.CreateCommand();
				oCmdActualiza.Connection = oConexion;
				oCmdActualiza.CommandType = CommandType.Text;
				oCmdActualiza.CommandTimeout = 600;//120;
				try
				{		
					oTransaction = oConexion.BeginTransaction();
					oCmdActualiza.Transaction = oTransaction;

					#region Cobro seguro
					int idSPIFFSubGrupoRangoCab = 0;
					squery = String.Format("Select top 1 idSPIFFSubGrupoRangoCab From SPIFFSubGrupoRangoCab where Seleccionar = 1");
					idSPIFFSubGrupoRangoCab = Funcion.iEscalarSQL(cdsSeteoF,squery);

					if(idSPIFFSubGrupoRangoCab > 0)
					{
						string	sSQLg = string.Format("Exec SPIFFSubGrupoRangoCabDup {0}, '{1}', '{2}', '{3}', '{4}', {5}", 
							idSPIFFSubGrupoRangoCab, 
							this.txtNombre.Text,
							Convert.ToDateTime(this.dtDesde.Value).ToString("yyyyMMdd"), 
							Convert.ToDateTime(this.dtHasta.Value).ToString("yyyyMMdd"),
							this.txtNotas.Text,
							(int)this.cmbSPIFFEstado.Value);
						oCmdActualiza.CommandText = sSQLg;
						oCmdActualiza.ExecuteNonQuery();
					}
					else
					{
						MessageBox.Show(string.Format("Se debe seleccionar un registro para poder duplicarlo."), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
						this.ugdSPIFFSubGrupoRangoCab.Focus();				
						return;
					}
					#endregion Cobro seguro
					oTransaction.Commit();
					oConexion.Close();
					#region Mensaje exito
					MessageBox.Show("Duplicación registrada correctamente", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
					LimpiarCampos();
					Consulta();
					#endregion Mensaje exito
					
				}
				catch (Exception ex)
				{
					oTransaction.Rollback();
					MessageBox.Show(Funcion.smensajeerror, "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
					Funcion.LogSistema(frmRequerimiento.sconexionPoint, this.Name + ": "+ this.Text, string.Format("Commit Exception Type : {0} {1}", ex.GetType(), ex.Message), "", Funcion.slogtipoerror);
				}
				finally
				{
					oConexion.Close();
					// Ocultar el modal de "procesando" una vez que se haya completado la consulta
					MenuLatinium.OcultarModalProcesando();
				}

			}
			#endregion Duplicación SPIFF Inventario
		}

		private void LimpiarCampos()
		{
			this.txtNombre.Text = "";
			this.cmbSPIFFEstado.Value = DBNull.Value;
			this.dtDesde.Value = DBNull.Value;
			this.dtHasta.Value = DBNull.Value;
			this.txtNotas.Text = "";
		}

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);
		}

		private void ugdSPIFFSubGrupoRangoCab_AfterCellUpdate(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			if (e.Cell.Column.ToString() == "Seleccionar")
			{
				int idSPIFFSubGrupoRangoCab = (int)e.Cell.Row.Cells["idSPIFFSubGrupoRangoCab"].Value;
				bool Seleccionar = (bool)e.Cell.Row.Cells["Seleccionar"].Value;
				SPIFFSubGrupoRangoCabSel(idSPIFFSubGrupoRangoCab, Seleccionar);
			}
		}

		public void SPIFFSubGrupoRangoCabSel(int idSPIFFSubGrupoRangoCab, bool Seleccionar)
		{
			#region Guardar datos
			//Actualiza valores en la compra
			string squery = string.Format("EXEC SPIFFSubGrupoRangoCabSel {0}, {1}", 
				idSPIFFSubGrupoRangoCab, Seleccionar);
			Funcion.EjecutaSQL(cdsSeteoF, squery);
			#endregion Guardar datos
		}

		private void ugdSPIFFSubGrupoRangoCab_ClickCellButton(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			int idSPIFFSubGrupoRangoCab = 0;


			if (e.Cell.Column.ToString() == "btnAnadir")	
			{
				try
				{
					idSPIFFSubGrupoRangoCab = (int)e.Cell.Row.Cells["idSPIFFSubGrupoRangoCab"].Value;
					using (frmSPIFFConfiguracionInventarioAnadir frmS = new frmSPIFFConfiguracionInventarioAnadir(idSPIFFSubGrupoRangoCab))
					{		
						if (DialogResult.OK == frmS.ShowDialog())
						{
					 
						}
						Consulta();
					}
				}
				catch(Exception ex)
				{
					string sms = ex.Message;
					return;
				}
			}

			int idSPIFFEstado = 0;
			if (e.Cell.Column.ToString() == "btnAnular")	
			{
				idSPIFFEstado = (int)e.Cell.Row.Cells["idSPIFFEstado"].Value;
				if(idSPIFFEstado == idSPIFFEstadoEnProceso)
				{
					try
					{
						idSPIFFSubGrupoRangoCab = (int)e.Cell.Row.Cells["idSPIFFSubGrupoRangoCab"].Value;
						using (frmSPIFFConfiguracionInventarioDelete frmscid = new frmSPIFFConfiguracionInventarioDelete(idSPIFFSubGrupoRangoCab))
						{		
							if (DialogResult.OK == frmscid.ShowDialog())
							{
					 
							}
							Consulta();
						}
					}
					catch(Exception ex)
					{
						string sms = ex.Message;
						return;
					}
				}
				else
				{
					MessageBox.Show(string.Format("Este registro no puede ser anulado."), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					Consulta();
				}
			}
		}

		private void btnSalir_Click(object sender, System.EventArgs e)
		{
			this.Close();
		}
	}
}
