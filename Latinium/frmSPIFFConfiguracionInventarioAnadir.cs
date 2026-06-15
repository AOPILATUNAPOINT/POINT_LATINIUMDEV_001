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
	/// Descripción breve de frmSPIFFConfiguracionInventarioAnadir.
	/// </summary>
	public class frmSPIFFConfiguracionInventarioAnadir : System.Windows.Forms.Form
	{
		private System.Windows.Forms.GroupBox groupBox2;
		public Infragistics.Win.UltraWinGrid.UltraCombo cmbSPIFFEstado;
		private System.Windows.Forms.Label label5;
		private System.Windows.Forms.Label label6;
		public Infragistics.Win.UltraWinEditors.UltraTextEditor txtNotas;
		private System.Windows.Forms.Label label1;
		public Infragistics.Win.UltraWinEditors.UltraTextEditor txtNombre;
		private System.Windows.Forms.Label lblBodega;
		private System.Windows.Forms.Label label53;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtDesde;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtHasta;
		private System.Windows.Forms.Button btnSalir;
		private System.Windows.Forms.Button btnDuplicar;
		private System.Windows.Forms.GroupBox groupBox1;
		private System.Windows.Forms.GroupBox groupBox5;
		private System.Windows.Forms.Button btnConsultar;
		public System.Windows.Forms.Label label12;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtBuscar;
		private System.Windows.Forms.Label lblContador;
		private C1.Data.C1DataSet cdsSeteoF;
		private System.Windows.Forms.ToolTip toolTip1;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource udsSPIFFEstado;
		private System.ComponentModel.IContainer components;
		private Infragistics.Win.UltraWinGrid.UltraGrid ugdSPIFFSubGrupo;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource udsSPIFFSubGrupo;
		int idSPIFFSubGrupoRangoCab = 0;

		public frmSPIFFConfiguracionInventarioAnadir(int IdSPIFFSubGrupoRangoCab)
		{
			idSPIFFSubGrupoRangoCab = IdSPIFFSubGrupoRangoCab;
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
			Infragistics.Win.Appearance appearance1 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idSPIFFEstado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Descripcion");
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton1 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmSPIFFConfiguracionInventarioAnadir));
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton2 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idSPIFFSubGrupo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idSubGrupo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Descripcion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Valor1", -1, null, 0, Infragistics.Win.UltraWinGrid.SortIndicator.Ascending, false);
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idSPIFFSubGrupoRango1");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Valor2");
			Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance11 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idSPIFFSubGrupoRango2");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Valor3");
			Infragistics.Win.Appearance appearance12 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance13 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn11 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idSPIFFSubGrupoRango3");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn12 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("btnGestion", 0);
			Infragistics.Win.Appearance appearance14 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn13 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("btnProforma", 1);
			Infragistics.Win.Appearance appearance15 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn14 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Row", 2);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn15 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("btnAnadir", 3);
			Infragistics.Win.Appearance appearance16 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance17 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn16 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("btnAnular", 4);
			Infragistics.Win.Appearance appearance18 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance19 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings1 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Valor1", 3, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Valor1", 3, true);
			Infragistics.Win.Appearance appearance20 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings2 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Valor2", 5, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Valor2", 5, true);
			Infragistics.Win.Appearance appearance21 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings3 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Valor3", 7, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Valor3", 7, true);
			Infragistics.Win.Appearance appearance22 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance23 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance24 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance25 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance26 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance27 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance28 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance29 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idSPIFFSubGrupo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idSubGrupo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn3 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Descripcion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn4 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Valor1");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn5 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idSPIFFSubGrupoRango1");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn6 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Valor2");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn7 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idSPIFFSubGrupoRango2");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn8 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Valor3");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn9 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idSPIFFSubGrupoRango3");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn10 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idSPIFFEstado");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn11 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Descripcion");
			this.groupBox2 = new System.Windows.Forms.GroupBox();
			this.cmbSPIFFEstado = new Infragistics.Win.UltraWinGrid.UltraCombo();
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
			this.groupBox1 = new System.Windows.Forms.GroupBox();
			this.groupBox5 = new System.Windows.Forms.GroupBox();
			this.btnConsultar = new System.Windows.Forms.Button();
			this.label12 = new System.Windows.Forms.Label();
			this.txtBuscar = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.lblContador = new System.Windows.Forms.Label();
			this.ugdSPIFFSubGrupo = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.udsSPIFFSubGrupo = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.toolTip1 = new System.Windows.Forms.ToolTip(this.components);
			this.udsSPIFFEstado = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.groupBox2.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.cmbSPIFFEstado)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNotas)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNombre)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtDesde)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtHasta)).BeginInit();
			this.groupBox1.SuspendLayout();
			this.groupBox5.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.txtBuscar)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ugdSPIFFSubGrupo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.udsSPIFFSubGrupo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.udsSPIFFEstado)).BeginInit();
			this.SuspendLayout();
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
			this.groupBox2.TabIndex = 996;
			this.groupBox2.TabStop = false;
			this.groupBox2.Text = "Datos principales";
			// 
			// cmbSPIFFEstado
			// 
			appearance1.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbSPIFFEstado.Appearance = appearance1;
			this.cmbSPIFFEstado.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbSPIFFEstado.Cursor = System.Windows.Forms.Cursors.Default;
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn2.Header.Caption = "Descripción";
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridColumn2.Width = 279;
			ultraGridBand1.Columns.AddRange(new object[] {
																										 ultraGridColumn1,
																										 ultraGridColumn2});
			this.cmbSPIFFEstado.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.cmbSPIFFEstado.DisplayMember = "Descripcion";
			this.cmbSPIFFEstado.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbSPIFFEstado.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbSPIFFEstado.Location = new System.Drawing.Point(480, 24);
			this.cmbSPIFFEstado.Name = "cmbSPIFFEstado";
			this.cmbSPIFFEstado.Size = new System.Drawing.Size(280, 21);
			this.cmbSPIFFEstado.TabIndex = 1040;
			this.cmbSPIFFEstado.ValueMember = "idSPIFFEstado";
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
			appearance2.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtNotas.Appearance = appearance2;
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
			appearance3.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtNombre.Appearance = appearance3;
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
			appearance4.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtDesde.Appearance = appearance4;
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
			appearance5.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtHasta.Appearance = appearance5;
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
			this.btnSalir.Location = new System.Drawing.Point(832, 56);
			this.btnSalir.Name = "btnSalir";
			this.btnSalir.Size = new System.Drawing.Size(80, 23);
			this.btnSalir.TabIndex = 995;
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
			this.btnDuplicar.Location = new System.Drawing.Point(832, 24);
			this.btnDuplicar.Name = "btnDuplicar";
			this.btnDuplicar.Size = new System.Drawing.Size(80, 23);
			this.btnDuplicar.TabIndex = 994;
			this.btnDuplicar.Text = "&Guardar";
			this.btnDuplicar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnDuplicar.Click += new System.EventHandler(this.btnDuplicar_Click);
			// 
			// groupBox1
			// 
			this.groupBox1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.groupBox1.Controls.Add(this.groupBox5);
			this.groupBox1.Controls.Add(this.ugdSPIFFSubGrupo);
			this.groupBox1.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.groupBox1.Location = new System.Drawing.Point(16, 176);
			this.groupBox1.Name = "groupBox1";
			this.groupBox1.Size = new System.Drawing.Size(888, 328);
			this.groupBox1.TabIndex = 993;
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
			this.groupBox5.Size = new System.Drawing.Size(856, 64);
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
			appearance6.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtBuscar.Appearance = appearance6;
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
			// ugdSPIFFSubGrupo
			// 
			this.ugdSPIFFSubGrupo.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.ugdSPIFFSubGrupo.Cursor = System.Windows.Forms.Cursors.Default;
			this.ugdSPIFFSubGrupo.DataSource = this.udsSPIFFSubGrupo;
			appearance7.BackColor = System.Drawing.Color.White;
			this.ugdSPIFFSubGrupo.DisplayLayout.Appearance = appearance7;
			ultraGridBand2.AutoPreviewEnabled = true;
			ultraGridColumn3.Header.VisiblePosition = 0;
			ultraGridColumn3.Hidden = true;
			ultraGridColumn4.Header.VisiblePosition = 1;
			ultraGridColumn4.Hidden = true;
			ultraGridColumn4.RowLayoutColumnInfo.OriginX = 0;
			ultraGridColumn4.RowLayoutColumnInfo.OriginY = 0;
			ultraGridColumn4.RowLayoutColumnInfo.SpanX = 2;
			ultraGridColumn4.RowLayoutColumnInfo.SpanY = 2;
			ultraGridColumn5.AutoEdit = false;
			ultraGridColumn5.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
			ultraGridColumn5.Header.Caption = "Descripción";
			ultraGridColumn5.Header.VisiblePosition = 3;
			ultraGridColumn5.RowLayoutColumnInfo.OriginX = 2;
			ultraGridColumn5.RowLayoutColumnInfo.OriginY = 0;
			ultraGridColumn5.RowLayoutColumnInfo.PreferredCellSize = new System.Drawing.Size(371, 0);
			ultraGridColumn5.RowLayoutColumnInfo.PreferredLabelSize = new System.Drawing.Size(0, 21);
			ultraGridColumn5.RowLayoutColumnInfo.SpanX = 2;
			ultraGridColumn5.RowLayoutColumnInfo.SpanY = 2;
			appearance8.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn6.CellAppearance = appearance8;
			appearance9.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn6.CellButtonAppearance = appearance9;
			ultraGridColumn6.Format = "#,##0.00";
			ultraGridColumn6.Header.Caption = "121 A 150 DÍAS ($)";
			ultraGridColumn6.Header.VisiblePosition = 4;
			ultraGridColumn6.RowLayoutColumnInfo.OriginX = 4;
			ultraGridColumn6.RowLayoutColumnInfo.OriginY = 0;
			ultraGridColumn6.RowLayoutColumnInfo.PreferredCellSize = new System.Drawing.Size(142, 0);
			ultraGridColumn6.RowLayoutColumnInfo.PreferredLabelSize = new System.Drawing.Size(0, 21);
			ultraGridColumn6.RowLayoutColumnInfo.SpanX = 2;
			ultraGridColumn6.RowLayoutColumnInfo.SpanY = 2;
			ultraGridColumn7.Header.VisiblePosition = 6;
			ultraGridColumn7.Hidden = true;
			appearance10.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn8.CellAppearance = appearance10;
			appearance11.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn8.CellButtonAppearance = appearance11;
			ultraGridColumn8.Format = "#,##0.00";
			ultraGridColumn8.Header.Caption = "151 A 180 DÍAS  ($)";
			ultraGridColumn8.Header.VisiblePosition = 7;
			ultraGridColumn8.RowLayoutColumnInfo.OriginX = 6;
			ultraGridColumn8.RowLayoutColumnInfo.OriginY = 0;
			ultraGridColumn8.RowLayoutColumnInfo.PreferredCellSize = new System.Drawing.Size(130, 0);
			ultraGridColumn8.RowLayoutColumnInfo.PreferredLabelSize = new System.Drawing.Size(0, 21);
			ultraGridColumn8.RowLayoutColumnInfo.SpanX = 2;
			ultraGridColumn8.RowLayoutColumnInfo.SpanY = 2;
			ultraGridColumn9.Header.VisiblePosition = 8;
			ultraGridColumn9.Hidden = true;
			ultraGridColumn9.RowLayoutColumnInfo.OriginX = 10;
			ultraGridColumn9.RowLayoutColumnInfo.OriginY = 0;
			ultraGridColumn9.RowLayoutColumnInfo.SpanX = 2;
			ultraGridColumn9.RowLayoutColumnInfo.SpanY = 2;
			appearance12.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn10.CellAppearance = appearance12;
			appearance13.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn10.CellButtonAppearance = appearance13;
			ultraGridColumn10.Format = "#,##0.00";
			ultraGridColumn10.Header.Caption = "181 EN ADELANTE  ($)";
			ultraGridColumn10.Header.VisiblePosition = 9;
			ultraGridColumn10.RowLayoutColumnInfo.OriginX = 8;
			ultraGridColumn10.RowLayoutColumnInfo.OriginY = 0;
			ultraGridColumn10.RowLayoutColumnInfo.PreferredLabelSize = new System.Drawing.Size(0, 21);
			ultraGridColumn10.RowLayoutColumnInfo.SpanX = 2;
			ultraGridColumn10.RowLayoutColumnInfo.SpanY = 2;
			ultraGridColumn11.Header.VisiblePosition = 10;
			ultraGridColumn11.Hidden = true;
			ultraGridColumn12.ButtonDisplayStyle = Infragistics.Win.UltraWinGrid.ButtonDisplayStyle.Always;
			appearance14.Image = ((object)(resources.GetObject("appearance14.Image")));
			appearance14.ImageHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn12.CellButtonAppearance = appearance14;
			ultraGridColumn12.Header.Caption = "...";
			ultraGridColumn12.Header.VisiblePosition = 2;
			ultraGridColumn12.Hidden = true;
			ultraGridColumn12.RowLayoutColumnInfo.OriginX = 0;
			ultraGridColumn12.RowLayoutColumnInfo.OriginY = 0;
			ultraGridColumn12.RowLayoutColumnInfo.SpanX = 2;
			ultraGridColumn12.RowLayoutColumnInfo.SpanY = 2;
			ultraGridColumn12.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.Button;
			ultraGridColumn12.Width = 30;
			ultraGridColumn13.ButtonDisplayStyle = Infragistics.Win.UltraWinGrid.ButtonDisplayStyle.Always;
			appearance15.Image = ((object)(resources.GetObject("appearance15.Image")));
			appearance15.ImageHAlign = Infragistics.Win.HAlign.Center;
			appearance15.ImageVAlign = Infragistics.Win.VAlign.Middle;
			ultraGridColumn13.CellButtonAppearance = appearance15;
			ultraGridColumn13.Header.Caption = "Proforma";
			ultraGridColumn13.Header.VisiblePosition = 5;
			ultraGridColumn13.Hidden = true;
			ultraGridColumn13.RowLayoutColumnInfo.OriginX = 2;
			ultraGridColumn13.RowLayoutColumnInfo.OriginY = 0;
			ultraGridColumn13.RowLayoutColumnInfo.SpanX = 2;
			ultraGridColumn13.RowLayoutColumnInfo.SpanY = 2;
			ultraGridColumn13.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.Button;
			ultraGridColumn13.Width = 60;
			ultraGridColumn14.AutoEdit = false;
			ultraGridColumn14.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
			ultraGridColumn14.Header.Caption = "N°";
			ultraGridColumn14.Header.VisiblePosition = 11;
			ultraGridColumn14.RowLayoutColumnInfo.OriginX = 0;
			ultraGridColumn14.RowLayoutColumnInfo.OriginY = 0;
			ultraGridColumn14.RowLayoutColumnInfo.PreferredCellSize = new System.Drawing.Size(27, 0);
			ultraGridColumn14.RowLayoutColumnInfo.PreferredLabelSize = new System.Drawing.Size(0, 21);
			ultraGridColumn14.RowLayoutColumnInfo.SpanX = 2;
			ultraGridColumn14.RowLayoutColumnInfo.SpanY = 2;
			appearance16.Image = ((object)(resources.GetObject("appearance16.Image")));
			appearance16.ImageHAlign = Infragistics.Win.HAlign.Center;
			appearance16.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn15.CellAppearance = appearance16;
			appearance17.Image = ((object)(resources.GetObject("appearance17.Image")));
			appearance17.ImageHAlign = Infragistics.Win.HAlign.Center;
			appearance17.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn15.CellButtonAppearance = appearance17;
			ultraGridColumn15.Header.Caption = "...";
			ultraGridColumn15.Header.VisiblePosition = 12;
			ultraGridColumn15.Hidden = true;
			ultraGridColumn15.RowLayoutColumnInfo.OriginX = 8;
			ultraGridColumn15.RowLayoutColumnInfo.OriginY = 0;
			ultraGridColumn15.RowLayoutColumnInfo.PreferredCellSize = new System.Drawing.Size(24, 0);
			ultraGridColumn15.RowLayoutColumnInfo.SpanX = 2;
			ultraGridColumn15.RowLayoutColumnInfo.SpanY = 2;
			ultraGridColumn15.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.Button;
			appearance18.Image = ((object)(resources.GetObject("appearance18.Image")));
			appearance18.ImageHAlign = Infragistics.Win.HAlign.Center;
			appearance18.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn16.CellAppearance = appearance18;
			appearance19.Image = ((object)(resources.GetObject("appearance19.Image")));
			appearance19.ImageHAlign = Infragistics.Win.HAlign.Center;
			appearance19.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn16.CellButtonAppearance = appearance19;
			ultraGridColumn16.Header.Caption = "...";
			ultraGridColumn16.Header.VisiblePosition = 13;
			ultraGridColumn16.Hidden = true;
			ultraGridColumn16.RowLayoutColumnInfo.OriginX = 10;
			ultraGridColumn16.RowLayoutColumnInfo.OriginY = 0;
			ultraGridColumn16.RowLayoutColumnInfo.PreferredCellSize = new System.Drawing.Size(25, 0);
			ultraGridColumn16.RowLayoutColumnInfo.SpanX = 2;
			ultraGridColumn16.RowLayoutColumnInfo.SpanY = 2;
			ultraGridColumn16.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.Button;
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
																										 ultraGridColumn15,
																										 ultraGridColumn16});
			ultraGridBand2.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			appearance20.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings1.Appearance = appearance20;
			summarySettings1.DisplayFormat = "{0: #,##0.00}";
			summarySettings1.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance21.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings2.Appearance = appearance21;
			summarySettings2.DisplayFormat = "{0: #,##0.00}";
			summarySettings2.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance22.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings3.Appearance = appearance22;
			summarySettings3.DisplayFormat = "{0: #,##0.00}";
			summarySettings3.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			ultraGridBand2.Summaries.AddRange(new Infragistics.Win.UltraWinGrid.SummarySettings[] {
																																															summarySettings1,
																																															summarySettings2,
																																															summarySettings3});
			ultraGridBand2.SummaryFooterCaption = "Sumatoria:";
			ultraGridBand2.UseRowLayout = true;
			this.ugdSPIFFSubGrupo.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			this.ugdSPIFFSubGrupo.DisplayLayout.GroupByBox.Hidden = true;
			appearance23.ForeColor = System.Drawing.Color.Black;
			appearance23.ForeColorDisabled = System.Drawing.Color.Black;
			this.ugdSPIFFSubGrupo.DisplayLayout.Override.ActiveRowAppearance = appearance23;
			this.ugdSPIFFSubGrupo.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.Yes;
			this.ugdSPIFFSubGrupo.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.True;
			this.ugdSPIFFSubGrupo.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance24.BackColor = System.Drawing.Color.Transparent;
			this.ugdSPIFFSubGrupo.DisplayLayout.Override.CardAreaAppearance = appearance24;
			appearance25.ForeColor = System.Drawing.Color.Black;
			appearance25.ForeColorDisabled = System.Drawing.Color.Black;
			this.ugdSPIFFSubGrupo.DisplayLayout.Override.CellAppearance = appearance25;
			appearance26.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance26.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance26.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance26.FontData.BoldAsString = "True";
			appearance26.FontData.Name = "Arial";
			appearance26.FontData.SizeInPoints = 8F;
			appearance26.ForeColor = System.Drawing.Color.White;
			appearance26.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.ugdSPIFFSubGrupo.DisplayLayout.Override.HeaderAppearance = appearance26;
			this.ugdSPIFFSubGrupo.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti;
			appearance27.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance27.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance27.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ugdSPIFFSubGrupo.DisplayLayout.Override.RowAlternateAppearance = appearance27;
			appearance28.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance28.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance28.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ugdSPIFFSubGrupo.DisplayLayout.Override.RowSelectorAppearance = appearance28;
			appearance29.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance29.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance29.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ugdSPIFFSubGrupo.DisplayLayout.Override.SelectedRowAppearance = appearance29;
			this.ugdSPIFFSubGrupo.DisplayLayout.ViewStyleBand = Infragistics.Win.UltraWinGrid.ViewStyleBand.OutlookGroupBy;
			this.ugdSPIFFSubGrupo.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.ugdSPIFFSubGrupo.Location = new System.Drawing.Point(16, 96);
			this.ugdSPIFFSubGrupo.Name = "ugdSPIFFSubGrupo";
			this.ugdSPIFFSubGrupo.RightToLeft = System.Windows.Forms.RightToLeft.No;
			this.ugdSPIFFSubGrupo.RowUpdateCancelAction = Infragistics.Win.UltraWinGrid.RowUpdateCancelAction.RetainDataAndActivation;
			this.ugdSPIFFSubGrupo.Size = new System.Drawing.Size(856, 216);
			this.ugdSPIFFSubGrupo.TabIndex = 981;
			this.ugdSPIFFSubGrupo.AfterCellUpdate += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.ugdSPIFFSubGrupo_AfterCellUpdate);
			this.ugdSPIFFSubGrupo.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.ugdSPIFFSubGrupo_InitializeLayout);
			// 
			// udsSPIFFSubGrupo
			// 
			ultraDataColumn1.DataType = typeof(int);
			ultraDataColumn2.DataType = typeof(int);
			ultraDataColumn4.DataType = typeof(System.Decimal);
			ultraDataColumn5.DataType = typeof(int);
			ultraDataColumn6.DataType = typeof(System.Decimal);
			ultraDataColumn7.DataType = typeof(int);
			ultraDataColumn8.DataType = typeof(System.Decimal);
			ultraDataColumn9.DataType = typeof(int);
			this.udsSPIFFSubGrupo.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn1,
																																 ultraDataColumn2,
																																 ultraDataColumn3,
																																 ultraDataColumn4,
																																 ultraDataColumn5,
																																 ultraDataColumn6,
																																 ultraDataColumn7,
																																 ultraDataColumn8,
																																 ultraDataColumn9});
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
			// udsSPIFFEstado
			// 
			ultraDataColumn10.DataType = typeof(int);
			this.udsSPIFFEstado.Band.Columns.AddRange(new object[] {
																															 ultraDataColumn10,
																															 ultraDataColumn11});
			// 
			// frmSPIFFConfiguracionInventarioAnadir
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(928, 518);
			this.Controls.Add(this.groupBox2);
			this.Controls.Add(this.btnSalir);
			this.Controls.Add(this.btnDuplicar);
			this.Controls.Add(this.groupBox1);
			this.Name = "frmSPIFFConfiguracionInventarioAnadir";
			this.Text = "SPIFF añadir configuración ";
			this.Load += new System.EventHandler(this.frmSPIFFConfiguracionInventarioAnadir_Load);
			this.groupBox2.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.cmbSPIFFEstado)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNotas)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNombre)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtDesde)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtHasta)).EndInit();
			this.groupBox1.ResumeLayout(false);
			this.groupBox5.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.txtBuscar)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ugdSPIFFSubGrupo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.udsSPIFFSubGrupo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.udsSPIFFEstado)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void frmSPIFFConfiguracionInventarioAnadir_Load(object sender, System.EventArgs e)
		{
			SPIFFEstadoCon();
			SPIFFSubGrupoRangoCabModel();
			Consulta();
			//Inhabilitar campos, que se generaron, para el proceso, y que no pueden ser modificados.
			Control();
		}

		private void Control()
		{
			this.txtNombre.Enabled = false;
			this.dtDesde.Enabled = false;
			this.dtHasta.Enabled = false;
		}

		private void SPIFFEstadoCon()
		{
			string squery = String.Format("Exec SPIFFEstadoCon");
			this.cmbSPIFFEstado.DataSource = Funcion.dvProcedimiento(cdsSeteoF, squery);
		}

		private void SPIFFSubGrupoRangoCabModel()
		{
			if(idSPIFFSubGrupoRangoCab > 0)
			{
				this.txtNombre.Text = Funcion.sEscalarSQL(cdsSeteoF, String.Format("Exec SPIFFSubGrupoRangoCabModel '{0}', {1}", "ssgrc.Nombre", idSPIFFSubGrupoRangoCab));
			  this.cmbSPIFFEstado.Value = Funcion.iEscalarSQL(cdsSeteoF, String.Format("Exec SPIFFSubGrupoRangoCabModel '{0}', {1}", "ssgrc.idSPIFFEstado", idSPIFFSubGrupoRangoCab));
				this.dtDesde.Value = Funcion.dtEscalarSQL(cdsSeteoF, String.Format("Exec SPIFFSubGrupoRangoCabModel '{0}', {1}", "ssgrc.Desde", idSPIFFSubGrupoRangoCab));
				this.dtHasta.Value = Funcion.dtEscalarSQL(cdsSeteoF, String.Format("Exec SPIFFSubGrupoRangoCabModel '{0}', {1}", "ssgrc.Hasta", idSPIFFSubGrupoRangoCab));
				this.txtNotas.Text = Funcion.sEscalarSQL(cdsSeteoF, String.Format("Exec SPIFFSubGrupoRangoCabModel '{0}', {1}", "ssgrc.Nota", idSPIFFSubGrupoRangoCab));
			}
		}

		private void Consulta()
		{
			string sSQLPa = string.Format("Exec SPIFFSubGrupoCon {0}, '{1}'", idSPIFFSubGrupoRangoCab, this.txtBuscar.Text); //this.txtBusqueda.Text.ToString()
			FuncionesProcedimientos.dsGeneralModulo(sSQLPa, this.ugdSPIFFSubGrupo);
			//Label de registros
			this.lblContador.Text = this.ugdSPIFFSubGrupo.Rows.Count + " REGISTROS ENCONTRADOS";
			//COntador de filas
			this.ugdSPIFFSubGrupo = Funcion.ContadorFilas(ugdSPIFFSubGrupo, "Row");
		}

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);
		}

		private void ugdSPIFFSubGrupo_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
			FuncionesProcedimientos.FormatoGrid(e, "Valor1", 2);
			FuncionesProcedimientos.FormatoGrid(e, "Valor2", 2);
			FuncionesProcedimientos.FormatoGrid(e, "Valor3", 2);
		}

		private void ugdSPIFFSubGrupo_AfterCellUpdate(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			int idSPIFFSubGrupoRango = 0;
			decimal Valor = 0;
			int idSubGrupo = 0;

			if (e.Cell.Column.ToString() == "Valor1")
			{
				try{idSPIFFSubGrupoRango = (int)e.Cell.Row.Cells["idSPIFFSubGrupoRango1"].Value;}catch(Exception ex){string sms = ex.Message;}
				try{Valor = Convert.ToDecimal(e.Cell.Row.Cells["Valor1"].Value);}catch(Exception ex){string sms = ex.Message;}
				idSubGrupo = (int)e.Cell.Row.Cells["idSubGrupo"].Value;
				e.Cell.Row.Cells["idSPIFFSubGrupoRango1"].Value = SPIFFSubGrupoRangoAct(idSPIFFSubGrupoRango, Valor, frmSPIFFConfiguracionInventario.idSPIFFRango1, idSPIFFSubGrupoRangoCab, idSubGrupo);
			}
			if (e.Cell.Column.ToString() == "Valor2")
			{
				try{idSPIFFSubGrupoRango = (int)e.Cell.Row.Cells["idSPIFFSubGrupoRango2"].Value;}catch(Exception ex){string sms = ex.Message;}
				try{Valor = Convert.ToDecimal(e.Cell.Row.Cells["Valor2"].Value);}catch(Exception ex){string sms = ex.Message;}
				idSubGrupo = (int)e.Cell.Row.Cells["idSubGrupo"].Value;
				e.Cell.Row.Cells["idSPIFFSubGrupoRango2"].Value = SPIFFSubGrupoRangoAct(idSPIFFSubGrupoRango, Valor, frmSPIFFConfiguracionInventario.idSPIFFRango2, idSPIFFSubGrupoRangoCab, idSubGrupo);
			}
			if (e.Cell.Column.ToString() == "Valor3")
			{
				try{idSPIFFSubGrupoRango = (int)e.Cell.Row.Cells["idSPIFFSubGrupoRango3"].Value;}catch(Exception ex){string sms = ex.Message;}
				try{Valor = Convert.ToDecimal(e.Cell.Row.Cells["Valor3"].Value);}catch(Exception ex){string sms = ex.Message;}
				idSubGrupo = (int)e.Cell.Row.Cells["idSubGrupo"].Value;
				e.Cell.Row.Cells["idSPIFFSubGrupoRango3"].Value = SPIFFSubGrupoRangoAct(idSPIFFSubGrupoRango, Valor, frmSPIFFConfiguracionInventario.idSPIFFRango3, idSPIFFSubGrupoRangoCab, idSubGrupo);
			}
		}

		private int SPIFFSubGrupoRangoAct(int idSPIFFSubGrupoRango, decimal Valor, int idSPIFFRango, int idSPIFFSubGrupoRangoCab, int idSubGrupo)
		{
			int iResultado = 0;
			#region Guardar datos
			try
			{
				//Actualiza valores en la compra
				string squery = string.Format("EXEC SPIFFSubGrupoRangoAct {0}, {1}, {2}, {3}, {4}", 
					idSPIFFSubGrupoRango, Valor, idSPIFFRango, idSPIFFSubGrupoRangoCab, idSubGrupo);
				iResultado = Funcion.iEscalarSQL(cdsSeteoF, squery);
			}
			catch(Exception ex)
			{
				Funcion.LogSistema(frmRequerimiento.sconexionPoint, this.Name + ": "+ this.Text, string.Format("Commit Exception Type : {0} {1}", ex.GetType(), ex.Message), "", Funcion.slogtipoerror);
			}
			return iResultado;
			#endregion Guardar datos
		}

		private void btnDuplicar_Click(object sender, System.EventArgs e)
		{
			SPIFFSubGrupoRangoObjAct();
		
		}

		private void SPIFFSubGrupoRangoObjAct()
		{
			#region Validación
			if(!Validacion.vbTexto(this.txtNombre, 1, 500, "Nombre"))return;
			if(!Validacion.vbComboVacio(this.cmbSPIFFEstado, "Seleccione un estado.")) return;
			if(!Validacion.vbValidaFechasDesdeHasta(this.dtDesde, this.dtHasta, "Inicio", "Finaliza")) return;
			if(!Validacion.vbTexto(this.txtNotas, 1, 2000, "Nota"))return;

			int iContador = 0;
			string squery = "";

			// Se debe validar que no exista otro nombre similar, a otro SPIFF
			if(this.txtNombre.Text != "")
			{
				iContador = 0;
				squery = String.Format("EXEC SPIFFSubGrupoRangoCabNombreVal '{0}', {1}", this.txtNombre.Text, idSPIFFSubGrupoRangoCab); //si tiene id cab
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
				squery = String.Format("EXEC SPIFFSubGrupoRangoCabFechaVal '{0}', {1}", FechaInicio, idSPIFFSubGrupoRangoCab); //si tiene id cab
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
				squery = String.Format("EXEC SPIFFSubGrupoRangoCabFechaVal '{0}', {1}", FechaFinaliza, idSPIFFSubGrupoRangoCab); //si tiene id cab
				iContador = Funcion.iEscalarSQL(cdsSeteoF,squery);

				if(iContador > 0)
				{
					MessageBox.Show(string.Format("La fecha finaliza, ya corresponde a otro SPIFF generado."), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					this.dtHasta.Focus();				
					return;
				}
			}

			#endregion Validación

			#region Duplicación SPIFF Inventario
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

					#region Cobro seguro
					if(idSPIFFSubGrupoRangoCab > 0)
					{
						string	sSQLg = string.Format("Exec SPIFFSubGrupoRangoObjAct {0}, '{1}', '{2}', '{3}', '{4}', {5}", 
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
						MessageBox.Show(string.Format("Se debe comunicar con el departamento de sistemas."), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
						this.ugdSPIFFSubGrupo.Focus();				
						return;
					}
					#endregion Cobro seguro
					oTransaction.Commit();
					oConexion.Close();
					#region Mensaje exito
					MessageBox.Show("Información registrada correctamente", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
					this.DialogResult = DialogResult.OK;
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
				}
			}
			#endregion Duplicación SPIFF Inventario
		}

		private void btnSalir_Click(object sender, System.EventArgs e)
		{
			this.Close();
		}

		private void btnConsultar_Click(object sender, System.EventArgs e)
		{
			Consulta();
		}


	}
}
