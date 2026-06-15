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
	/// Descripción breve de frmConsultaGestionesCallCenter.
	/// </summary>
	public class frmConsultaGestionesCallCenter : System.Windows.Forms.Form
	{
		private System.Windows.Forms.Button btnVer;
		private System.Windows.Forms.Label label5;
		private System.Windows.Forms.Label label4;
		public Infragistics.Win.UltraWinEditors.UltraTextEditor txtBuscar;
		private System.Windows.Forms.Label label2;
		private System.Windows.Forms.Label label1;
		public Infragistics.Win.UltraWinGrid.UltraCombo cmbGestionCredito;
		private System.Windows.Forms.Label label3;
		private System.Windows.Forms.Label label6;
		public Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtDesde;
		public Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtHasta;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbCobrador_Operador;
		public Infragistics.Win.UltraWinGrid.UltraCombo cmbResultadoGestion;
		private Infragistics.Win.UltraWinGrid.UltraGrid uGridInformacion;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource1;
		private C1.Data.C1DataSet cdsSeteoF;
		private Infragistics.Win.UltraWinEditors.UltraOptionSet optEstados;
		private System.Windows.Forms.Label lblContador;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource2;
		private System.Windows.Forms.Button btnExcel;
		private System.Windows.Forms.GroupBox groupBox2;
		private Infragistics.Win.UltraWinEditors.UltraOptionSet optCobradorOperador;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public frmConsultaGestionesCallCenter()
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
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmConsultaGestionesCallCenter));
			Infragistics.Win.Appearance appearance1 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPersonal");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idPersonal");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton1 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCbo_ResultadoGestion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Resultado");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand3 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCbo_EstadoGestion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Estado");
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton2 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.ValueListItem valueListItem1 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem2 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem3 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand4 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Periodo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCbo_GestorDeCobranzas");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCbo_EstadoGestion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Estado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn11 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCbo_ResultadoGestion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn12 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Resultado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn13 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cliente");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn14 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cedula");
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn15 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idOperador");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn16 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Operadora");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn17 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCobrador");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn18 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cobrador");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn19 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Desde");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn20 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Hasta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn21 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Hora");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn22 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("cuota_vencida");
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn23 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("valor_cuota_vencida");
			Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn24 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Dias_Mora_Actual");
			Infragistics.Win.Appearance appearance11 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn25 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Observaciones");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn26 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Numero_Gestiones");
			Infragistics.Win.Appearance appearance12 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn27 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCompra");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn28 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Ir", 0);
			Infragistics.Win.Appearance appearance13 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance14 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance15 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance16 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance17 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance18 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn3 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Periodo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn4 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCbo_GestorDeCobranzas");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn5 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCbo_EstadoGestion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn6 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Estado");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn7 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCbo_ResultadoGestion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn8 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Resultado");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn9 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Cliente");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn10 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Cedula");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn11 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idOperador");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn12 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Operadora");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn13 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCobrador");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn14 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Cobrador");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn15 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Desde");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn16 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Hasta");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn17 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Hora");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn18 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("cuota_vencida");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn19 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("valor_cuota_vencida");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn20 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Dias_Mora_Actual");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn21 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Observaciones");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn22 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Numero_Gestiones");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn23 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCompra");
			Infragistics.Win.Appearance appearance19 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance20 = new Infragistics.Win.Appearance();
			Infragistics.Win.ValueListItem valueListItem4 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem5 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem6 = new Infragistics.Win.ValueListItem();
			this.btnVer = new System.Windows.Forms.Button();
			this.cmbCobrador_Operador = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.ultraDataSource2 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.label5 = new System.Windows.Forms.Label();
			this.label4 = new System.Windows.Forms.Label();
			this.dtDesde = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.txtBuscar = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label2 = new System.Windows.Forms.Label();
			this.cmbResultadoGestion = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.label1 = new System.Windows.Forms.Label();
			this.cmbGestionCredito = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.label3 = new System.Windows.Forms.Label();
			this.dtHasta = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.label6 = new System.Windows.Forms.Label();
			this.optEstados = new Infragistics.Win.UltraWinEditors.UltraOptionSet();
			this.uGridInformacion = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource1 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.lblContador = new System.Windows.Forms.Label();
			this.btnExcel = new System.Windows.Forms.Button();
			this.groupBox2 = new System.Windows.Forms.GroupBox();
			this.optCobradorOperador = new Infragistics.Win.UltraWinEditors.UltraOptionSet();
			((System.ComponentModel.ISupportInitialize)(this.cmbCobrador_Operador)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtDesde)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtBuscar)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbResultadoGestion)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbGestionCredito)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtHasta)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.optEstados)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridInformacion)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.optCobradorOperador)).BeginInit();
			this.SuspendLayout();
			// 
			// btnVer
			// 
			this.btnVer.Image = ((System.Drawing.Image)(resources.GetObject("btnVer.Image")));
			this.btnVer.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnVer.Location = new System.Drawing.Point(392, 134);
			this.btnVer.Name = "btnVer";
			this.btnVer.Size = new System.Drawing.Size(75, 24);
			this.btnVer.TabIndex = 970;
			this.btnVer.Text = "&Ver";
			this.btnVer.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnVer.Click += new System.EventHandler(this.btnVer_Click);
			// 
			// cmbCobrador_Operador
			// 
			appearance1.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbCobrador_Operador.Appearance = appearance1;
			this.cmbCobrador_Operador.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbCobrador_Operador.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbCobrador_Operador.DataSource = this.ultraDataSource2;
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn2.Header.Caption = "Operadora/Cobrador";
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridColumn2.Width = 248;
			ultraGridBand1.Columns.AddRange(new object[] {
																										 ultraGridColumn1,
																										 ultraGridColumn2});
			this.cmbCobrador_Operador.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.cmbCobrador_Operador.DisplayMember = "Nombre";
			this.cmbCobrador_Operador.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbCobrador_Operador.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbCobrador_Operador.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbCobrador_Operador.Location = new System.Drawing.Point(120, 40);
			this.cmbCobrador_Operador.Name = "cmbCobrador_Operador";
			this.cmbCobrador_Operador.Size = new System.Drawing.Size(248, 21);
			this.cmbCobrador_Operador.TabIndex = 969;
			this.cmbCobrador_Operador.ValueMember = "idPersonal";
			this.cmbCobrador_Operador.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cmbCobrador_Operador_KeyDown);
			// 
			// ultraDataSource2
			// 
			ultraDataColumn1.DataType = typeof(int);
			this.ultraDataSource2.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn1,
																																 ultraDataColumn2});
			// 
			// label5
			// 
			this.label5.AutoSize = true;
			this.label5.Location = new System.Drawing.Point(8, 40);
			this.label5.Name = "label5";
			this.label5.Size = new System.Drawing.Size(102, 16);
			this.label5.TabIndex = 968;
			this.label5.Text = "Operador/Cobrador";
			// 
			// label4
			// 
			this.label4.AutoSize = true;
			this.label4.Location = new System.Drawing.Point(8, 16);
			this.label4.Name = "label4";
			this.label4.Size = new System.Drawing.Size(37, 16);
			this.label4.TabIndex = 967;
			this.label4.Text = "Desde";
			this.label4.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// dtDesde
			// 
			appearance2.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtDesde.Appearance = appearance2;
			this.dtDesde.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton1.Caption = "Today";
			this.dtDesde.DateButtons.Add(dateButton1);
			this.dtDesde.Format = "dd/MM/yyyy ";
			this.dtDesde.Location = new System.Drawing.Point(72, 8);
			this.dtDesde.Name = "dtDesde";
			this.dtDesde.NonAutoSizeHeight = 23;
			this.dtDesde.Size = new System.Drawing.Size(104, 21);
			this.dtDesde.SpinButtonsVisible = true;
			this.dtDesde.TabIndex = 966;
			this.dtDesde.Value = ((object)(resources.GetObject("dtDesde.Value")));
			// 
			// txtBuscar
			// 
			this.txtBuscar.AcceptsReturn = true;
			appearance3.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtBuscar.Appearance = appearance3;
			this.txtBuscar.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtBuscar.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtBuscar.Location = new System.Drawing.Point(72, 136);
			this.txtBuscar.MaxLength = 1000;
			this.txtBuscar.Name = "txtBuscar";
			this.txtBuscar.Size = new System.Drawing.Size(296, 21);
			this.txtBuscar.TabIndex = 965;
			this.txtBuscar.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtBuscar_KeyDown);
			// 
			// label2
			// 
			this.label2.AutoSize = true;
			this.label2.Location = new System.Drawing.Point(8, 138);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(39, 16);
			this.label2.TabIndex = 964;
			this.label2.Text = "Buscar";
			this.label2.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// cmbResultadoGestion
			// 
			this.cmbResultadoGestion.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbResultadoGestion.Cursor = System.Windows.Forms.Cursors.Default;
			ultraGridColumn3.Header.VisiblePosition = 0;
			ultraGridColumn3.Hidden = true;
			ultraGridColumn4.Header.Caption = "Gestión";
			ultraGridColumn4.Header.VisiblePosition = 1;
			ultraGridColumn4.Width = 295;
			ultraGridBand2.Columns.AddRange(new object[] {
																										 ultraGridColumn3,
																										 ultraGridColumn4});
			this.cmbResultadoGestion.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			this.cmbResultadoGestion.DisplayMember = "Resultado";
			this.cmbResultadoGestion.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbResultadoGestion.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbResultadoGestion.Location = new System.Drawing.Point(72, 104);
			this.cmbResultadoGestion.MaxDropDownItems = 30;
			this.cmbResultadoGestion.Name = "cmbResultadoGestion";
			this.cmbResultadoGestion.Size = new System.Drawing.Size(296, 21);
			this.cmbResultadoGestion.TabIndex = 963;
			this.cmbResultadoGestion.ValueMember = "idCbo_ResultadoGestion";
			this.cmbResultadoGestion.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cmbResultadoGestion_KeyDown);
			// 
			// label1
			// 
			this.label1.AutoSize = true;
			this.label1.Location = new System.Drawing.Point(8, 72);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(43, 16);
			this.label1.TabIndex = 962;
			this.label1.Text = "Gestión";
			this.label1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// cmbGestionCredito
			// 
			this.cmbGestionCredito.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbGestionCredito.Cursor = System.Windows.Forms.Cursors.Default;
			ultraGridColumn5.Header.VisiblePosition = 0;
			ultraGridColumn5.Hidden = true;
			ultraGridColumn5.Width = 295;
			ultraGridColumn6.Header.VisiblePosition = 1;
			ultraGridColumn6.Width = 295;
			ultraGridBand3.Columns.AddRange(new object[] {
																										 ultraGridColumn5,
																										 ultraGridColumn6});
			this.cmbGestionCredito.DisplayLayout.BandsSerializer.Add(ultraGridBand3);
			this.cmbGestionCredito.DisplayMember = "Estado";
			this.cmbGestionCredito.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbGestionCredito.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbGestionCredito.Location = new System.Drawing.Point(72, 72);
			this.cmbGestionCredito.MaxDropDownItems = 30;
			this.cmbGestionCredito.Name = "cmbGestionCredito";
			this.cmbGestionCredito.Size = new System.Drawing.Size(296, 21);
			this.cmbGestionCredito.TabIndex = 961;
			this.cmbGestionCredito.ValueMember = "idCbo_EstadoGestion";
			this.cmbGestionCredito.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cmbGestionCredito_KeyDown);
			this.cmbGestionCredito.ValueChanged += new System.EventHandler(this.cmbGestionCredito_ValueChanged);
			// 
			// label3
			// 
			this.label3.AutoSize = true;
			this.label3.Location = new System.Drawing.Point(8, 104);
			this.label3.Name = "label3";
			this.label3.Size = new System.Drawing.Size(39, 16);
			this.label3.TabIndex = 960;
			this.label3.Text = "Estado";
			this.label3.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// dtHasta
			// 
			appearance4.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtHasta.Appearance = appearance4;
			this.dtHasta.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton2.Caption = "Today";
			this.dtHasta.DateButtons.Add(dateButton2);
			this.dtHasta.Format = "dd/MM/yyyy ";
			this.dtHasta.Location = new System.Drawing.Point(240, 8);
			this.dtHasta.Name = "dtHasta";
			this.dtHasta.NonAutoSizeHeight = 23;
			this.dtHasta.Size = new System.Drawing.Size(104, 21);
			this.dtHasta.SpinButtonsVisible = true;
			this.dtHasta.TabIndex = 974;
			this.dtHasta.Value = ((object)(resources.GetObject("dtHasta.Value")));
			// 
			// label6
			// 
			this.label6.AutoSize = true;
			this.label6.Location = new System.Drawing.Point(192, 16);
			this.label6.Name = "label6";
			this.label6.Size = new System.Drawing.Size(34, 16);
			this.label6.TabIndex = 975;
			this.label6.Text = "Hasta";
			this.label6.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// optEstados
			// 
			appearance5.ForeColorDisabled = System.Drawing.Color.Black;
			this.optEstados.Appearance = appearance5;
			this.optEstados.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.optEstados.CheckedIndex = 0;
			this.optEstados.ItemAppearance = appearance6;
			this.optEstados.ItemOrigin = new System.Drawing.Point(10, 0);
			valueListItem1.DataValue = 0;
			valueListItem1.DisplayText = "Todos";
			valueListItem2.DataValue = 1;
			valueListItem2.DisplayText = "Gestión";
			valueListItem3.DataValue = 2;
			valueListItem3.DisplayText = "sin Gestión";
			this.optEstados.Items.Add(valueListItem1);
			this.optEstados.Items.Add(valueListItem2);
			this.optEstados.Items.Add(valueListItem3);
			this.optEstados.ItemSpacingVertical = 10;
			this.optEstados.Location = new System.Drawing.Point(384, 40);
			this.optEstados.Name = "optEstados";
			this.optEstados.Size = new System.Drawing.Size(232, 26);
			this.optEstados.TabIndex = 976;
			this.optEstados.Text = "Todos";
			this.optEstados.ValueChanged += new System.EventHandler(this.optEstados_ValueChanged);
			// 
			// uGridInformacion
			// 
			this.uGridInformacion.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.uGridInformacion.Cursor = System.Windows.Forms.Cursors.Default;
			this.uGridInformacion.DataSource = this.ultraDataSource1;
			appearance7.BackColor = System.Drawing.Color.White;
			this.uGridInformacion.DisplayLayout.Appearance = appearance7;
			ultraGridColumn7.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn7.Header.VisiblePosition = 1;
			ultraGridColumn7.Width = 75;
			ultraGridColumn8.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn8.Header.VisiblePosition = 3;
			ultraGridColumn8.Hidden = true;
			ultraGridColumn9.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn9.Header.VisiblePosition = 4;
			ultraGridColumn9.Hidden = true;
			ultraGridColumn10.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn10.Header.Caption = "Gestión";
			ultraGridColumn10.Header.VisiblePosition = 9;
			ultraGridColumn10.Width = 150;
			ultraGridColumn11.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn11.Header.VisiblePosition = 10;
			ultraGridColumn11.Hidden = true;
			ultraGridColumn12.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn12.Header.Caption = "Estado";
			ultraGridColumn12.Header.VisiblePosition = 11;
			ultraGridColumn12.Width = 150;
			ultraGridColumn13.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn13.Header.VisiblePosition = 6;
			ultraGridColumn13.Width = 300;
			ultraGridColumn14.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance8.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn14.CellAppearance = appearance8;
			ultraGridColumn14.Header.VisiblePosition = 5;
			ultraGridColumn14.Width = 90;
			ultraGridColumn15.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn15.Header.VisiblePosition = 12;
			ultraGridColumn15.Hidden = true;
			ultraGridColumn16.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn16.Header.VisiblePosition = 7;
			ultraGridColumn16.Width = 250;
			ultraGridColumn17.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn17.Header.VisiblePosition = 13;
			ultraGridColumn17.Hidden = true;
			ultraGridColumn18.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn18.Header.VisiblePosition = 8;
			ultraGridColumn18.Width = 250;
			ultraGridColumn19.Format = "  yyy/MM/dd";
			ultraGridColumn19.Header.Caption = "Primera Gestión";
			ultraGridColumn19.Header.VisiblePosition = 14;
			ultraGridColumn19.Width = 100;
			ultraGridColumn20.Format = "  yyy/MM/dd ";
			ultraGridColumn20.Header.Caption = "Ultima Gestión";
			ultraGridColumn20.Header.VisiblePosition = 15;
			ultraGridColumn20.Width = 100;
			ultraGridColumn21.Format = "HH:mm:ss";
			ultraGridColumn21.Header.Caption = "Hora Gestión";
			ultraGridColumn21.Header.VisiblePosition = 16;
			ultraGridColumn21.Width = 100;
			appearance9.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn22.CellAppearance = appearance9;
			ultraGridColumn22.Header.Caption = "Cuota Vencida";
			ultraGridColumn22.Header.VisiblePosition = 17;
			ultraGridColumn22.Hidden = true;
			ultraGridColumn22.Width = 90;
			appearance10.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn23.CellAppearance = appearance10;
			ultraGridColumn23.Format = " #,##0.00";
			ultraGridColumn23.Header.Caption = "Valor Cuota Vencida";
			ultraGridColumn23.Header.VisiblePosition = 18;
			ultraGridColumn23.Width = 120;
			ultraGridColumn24.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance11.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn24.CellAppearance = appearance11;
			ultraGridColumn24.Header.Caption = "Dias Mora";
			ultraGridColumn24.Header.VisiblePosition = 20;
			ultraGridColumn24.Hidden = true;
			ultraGridColumn24.Width = 90;
			ultraGridColumn25.Header.VisiblePosition = 21;
			ultraGridColumn25.Width = 300;
			appearance12.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn26.CellAppearance = appearance12;
			ultraGridColumn26.Header.Caption = "# Gestiones";
			ultraGridColumn26.Header.VisiblePosition = 19;
			ultraGridColumn27.Header.VisiblePosition = 2;
			ultraGridColumn27.Hidden = true;
			ultraGridColumn28.ButtonDisplayStyle = Infragistics.Win.UltraWinGrid.ButtonDisplayStyle.Always;
			appearance13.Image = ((object)(resources.GetObject("appearance13.Image")));
			appearance13.ImageHAlign = Infragistics.Win.HAlign.Center;
			appearance13.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn28.CellButtonAppearance = appearance13;
			ultraGridColumn28.Header.Caption = "...";
			ultraGridColumn28.Header.VisiblePosition = 0;
			ultraGridColumn28.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.Button;
			ultraGridColumn28.Width = 30;
			ultraGridBand4.Columns.AddRange(new object[] {
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
																										 ultraGridColumn20,
																										 ultraGridColumn21,
																										 ultraGridColumn22,
																										 ultraGridColumn23,
																										 ultraGridColumn24,
																										 ultraGridColumn25,
																										 ultraGridColumn26,
																										 ultraGridColumn27,
																										 ultraGridColumn28});
			ultraGridBand4.Header.Caption = "Periodo A";
			ultraGridBand4.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
			ultraGridBand4.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			ultraGridBand4.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.False;
			ultraGridBand4.SummaryFooterCaption = "";
			this.uGridInformacion.DisplayLayout.BandsSerializer.Add(ultraGridBand4);
			this.uGridInformacion.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
			this.uGridInformacion.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.uGridInformacion.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.False;
			appearance14.BackColor = System.Drawing.Color.Transparent;
			this.uGridInformacion.DisplayLayout.Override.CardAreaAppearance = appearance14;
			appearance15.ForeColor = System.Drawing.Color.Black;
			appearance15.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridInformacion.DisplayLayout.Override.CellAppearance = appearance15;
			appearance16.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance16.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance16.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance16.FontData.BoldAsString = "True";
			appearance16.FontData.Name = "Arial";
			appearance16.FontData.SizeInPoints = 8F;
			appearance16.ForeColor = System.Drawing.Color.White;
			appearance16.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridInformacion.DisplayLayout.Override.HeaderAppearance = appearance16;
			this.uGridInformacion.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti;
			appearance17.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance17.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance17.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridInformacion.DisplayLayout.Override.RowSelectorAppearance = appearance17;
			appearance18.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance18.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance18.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridInformacion.DisplayLayout.Override.SelectedRowAppearance = appearance18;
			this.uGridInformacion.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridInformacion.Location = new System.Drawing.Point(8, 184);
			this.uGridInformacion.Name = "uGridInformacion";
			this.uGridInformacion.Size = new System.Drawing.Size(1208, 336);
			this.uGridInformacion.TabIndex = 977;
			this.uGridInformacion.ClickCellButton += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.uGridInformacion_ClickCellButton);
			// 
			// ultraDataSource1
			// 
			ultraDataColumn3.DataType = typeof(System.DateTime);
			ultraDataColumn4.DataType = typeof(int);
			ultraDataColumn5.DataType = typeof(int);
			ultraDataColumn7.DataType = typeof(int);
			ultraDataColumn11.DataType = typeof(int);
			ultraDataColumn13.DataType = typeof(int);
			ultraDataColumn15.DataType = typeof(System.DateTime);
			ultraDataColumn16.DataType = typeof(System.DateTime);
			ultraDataColumn17.DataType = typeof(System.DateTime);
			ultraDataColumn18.DataType = typeof(int);
			ultraDataColumn19.DataType = typeof(System.Decimal);
			ultraDataColumn20.DataType = typeof(int);
			ultraDataColumn22.DataType = typeof(int);
			ultraDataColumn23.DataType = typeof(int);
			this.ultraDataSource1.Band.Columns.AddRange(new object[] {
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
																																 ultraDataColumn20,
																																 ultraDataColumn21,
																																 ultraDataColumn22,
																																 ultraDataColumn23});
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
			// lblContador
			// 
			this.lblContador.AutoSize = true;
			this.lblContador.Font = new System.Drawing.Font("Tahoma", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblContador.ForeColor = System.Drawing.Color.Firebrick;
			this.lblContador.Location = new System.Drawing.Point(488, 135);
			this.lblContador.Name = "lblContador";
			this.lblContador.Size = new System.Drawing.Size(0, 23);
			this.lblContador.TabIndex = 978;
			this.lblContador.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// btnExcel
			// 
			this.btnExcel.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btnExcel.Enabled = false;
			this.btnExcel.Image = ((System.Drawing.Image)(resources.GetObject("btnExcel.Image")));
			this.btnExcel.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnExcel.Location = new System.Drawing.Point(1144, 8);
			this.btnExcel.Name = "btnExcel";
			this.btnExcel.Size = new System.Drawing.Size(72, 24);
			this.btnExcel.TabIndex = 979;
			this.btnExcel.Text = "&Excel";
			this.btnExcel.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnExcel.Click += new System.EventHandler(this.btnExcel_Click);
			// 
			// groupBox2
			// 
			this.groupBox2.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.groupBox2.Location = new System.Drawing.Point(0, 168);
			this.groupBox2.Name = "groupBox2";
			this.groupBox2.Size = new System.Drawing.Size(1224, 8);
			this.groupBox2.TabIndex = 980;
			this.groupBox2.TabStop = false;
			// 
			// optCobradorOperador
			// 
			appearance19.ForeColorDisabled = System.Drawing.Color.Black;
			this.optCobradorOperador.Appearance = appearance19;
			this.optCobradorOperador.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.optCobradorOperador.CheckedIndex = 0;
			this.optCobradorOperador.ItemAppearance = appearance20;
			this.optCobradorOperador.ItemOrigin = new System.Drawing.Point(10, 0);
			valueListItem4.DataValue = 0;
			valueListItem4.DisplayText = "General";
			valueListItem5.DataValue = 1;
			valueListItem5.DisplayText = "Operador";
			valueListItem6.DataValue = 2;
			valueListItem6.DisplayText = "Cobrador";
			this.optCobradorOperador.Items.Add(valueListItem4);
			this.optCobradorOperador.Items.Add(valueListItem5);
			this.optCobradorOperador.Items.Add(valueListItem6);
			this.optCobradorOperador.ItemSpacingVertical = 10;
			this.optCobradorOperador.Location = new System.Drawing.Point(384, 8);
			this.optCobradorOperador.Name = "optCobradorOperador";
			this.optCobradorOperador.Size = new System.Drawing.Size(232, 26);
			this.optCobradorOperador.TabIndex = 981;
			this.optCobradorOperador.Text = "General";
			this.optCobradorOperador.ValueChanged += new System.EventHandler(this.optCobradorOperador_ValueChanged);
			// 
			// frmConsultaGestionesCallCenter
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(1224, 526);
			this.Controls.Add(this.optCobradorOperador);
			this.Controls.Add(this.groupBox2);
			this.Controls.Add(this.btnExcel);
			this.Controls.Add(this.lblContador);
			this.Controls.Add(this.uGridInformacion);
			this.Controls.Add(this.optEstados);
			this.Controls.Add(this.label6);
			this.Controls.Add(this.label5);
			this.Controls.Add(this.label4);
			this.Controls.Add(this.label2);
			this.Controls.Add(this.label1);
			this.Controls.Add(this.label3);
			this.Controls.Add(this.dtHasta);
			this.Controls.Add(this.btnVer);
			this.Controls.Add(this.cmbCobrador_Operador);
			this.Controls.Add(this.dtDesde);
			this.Controls.Add(this.txtBuscar);
			this.Controls.Add(this.cmbResultadoGestion);
			this.Controls.Add(this.cmbGestionCredito);
			this.KeyPreview = true;
			this.Name = "frmConsultaGestionesCallCenter";
			this.Text = "Gestiones Call Center";
			this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.frmConsultaGestionesCallCenter_KeyDown);
			this.Load += new System.EventHandler(this.frmConsultaGestionesCallCenter_Load);
			((System.ComponentModel.ISupportInitialize)(this.cmbCobrador_Operador)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtDesde)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtBuscar)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbResultadoGestion)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbGestionCredito)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtHasta)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.optEstados)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridInformacion)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.optCobradorOperador)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

    int iGrupo = 0;
		private Acceso miAcceso;	

		private void UnloadMe()
		{
			this.Close();
		}

		private void frmConsultaGestionesCallCenter_Load(object sender, System.EventArgs e)
		{
			miAcceso = new Acceso(cdsSeteoF, "083604");
			if (!Funcion.Permiso("1192", cdsSeteoF))
			{		
				MessageBox.Show("No tiene Acceso a Gestion Call Center", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				BeginInvoke(new MethodInvoker(UnloadMe));
				return;
			}
			if (miAcceso.BExportar) this.btnExcel.Enabled = true;

			this.dtDesde.CalendarInfo.MinDate = DateTime.Parse("01/03/2019");	
			this.dtDesde.Value = DateTime.Today;

			this.dtHasta.CalendarInfo.MinDate = DateTime.Parse("01/03/2019");
			this.dtHasta.Value = DateTime.Today;

			
			int iPersonal = 0;

			string ssqlDo = string.Format ("Exec ConsultaCbo_OperadorCobrador {0}, 1", MenuLatinium.IdUsuario ); 
			SqlDataReader drDO = Funcion.rEscalarSQL(cdsSeteoF,ssqlDo,true);
			drDO.Read();
			if(drDO.HasRows)
			{
				iGrupo = drDO.GetInt32 (0);
			}
			drDO.Close();

			string sSqlID = string.Format ("Exec ConsultaCbo_OperadorCobrador {0}, 2", MenuLatinium.IdUsuario ); 
			SqlDataReader drId = Funcion.rEscalarSQL(cdsSeteoF,sSqlID,true);
			drId.Read();
			if(drId.HasRows)
			{
				iPersonal = drId.GetInt32 (0);
			}
			drId.Close();

			this.cmbGestionCredito.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Select idCbo_EstadoGestion, Estado from Cbo_EstadosGestion");
			int iHabilita =   Funcion.iEscalarSQL(cdsSeteoF, string.Format("Exec ConsultaCbo_OperadorCobrador {0}, 5,0,1", MenuLatinium.IdUsuario));	
			ListaBodegas(this.cmbCobrador_Operador, MenuLatinium.IdUsuario,iPersonal,4 ,cdsSeteoF);
			   
			this.cmbCobrador_Operador.Value =  Funcion.iEscalarSQL(cdsSeteoF, string.Format("Exec ConsultaCbo_OperadorCobrador {0}, 5", MenuLatinium.IdUsuario));	


			if (iHabilita == 0)
			{
				this.cmbCobrador_Operador.Enabled = false;
			}
		}
		private void ListaBodegas (Infragistics.Win.UltraWinGrid.UltraCombo uCombo, int Usuario,int iPersonal, int iTipo, C1.Data.C1DataSet ds)
		{
			uCombo.DataSource = Funcion.dvProcedimiento(ds, string.Format("Exec ConsultaCbo_OperadorCobrador {0}, 4, {1}", MenuLatinium.IdUsuario, iPersonal));
		}
	

		private void btnVer_Click(object sender, System.EventArgs e)
		{
			#region Valida Fechas
			if (this.dtDesde.Value == System.DBNull.Value)
			{
				MessageBox.Show("Seleccione la Fecha Inicial", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.dtDesde.Focus();
				return;
			}

			if (this.dtHasta.Value == System.DBNull.Value)
			{
				MessageBox.Show("Seleccione la Fecha Final", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.dtHasta.Focus();
				return;
			}

			if (Convert.ToDateTime(this.dtDesde.Value) > Convert.ToDateTime(this.dtHasta.Value))
			{
				MessageBox.Show("La Fecha Inicial no puede ser mayor a la Fecha Final", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.dtHasta.Focus();
				return;
			}
			#endregion Valida Fechas

			DateTime dtDesde = (DateTime)this.dtDesde.Value;
			DateTime dtHasta = (DateTime)this.dtHasta.Value;

			int iCobradorOperador = 0;
			int iGestionCredito = 0;
			int iResultadoGestion = 0;

			if (this.cmbCobrador_Operador.ActiveRow != null) iCobradorOperador = (int)this.cmbCobrador_Operador.Value;
			if (this.cmbGestionCredito.ActiveRow != null) iGestionCredito = (int)this.cmbGestionCredito.Value;
			if (this.cmbResultadoGestion.ActiveRow != null) iResultadoGestion = (int)this.cmbResultadoGestion.Value;

			string sSQL = string.Format("Exec ConsultaGestionesCallCenter '{0}', '{1}', {2}, {3}, {4}, {5}, '{6}',{7}", 
        dtDesde.ToString("yyyyMMdd"), dtHasta.ToString("yyyyMMdd"), (int)this.optEstados.Value, iCobradorOperador, iGestionCredito,
				iResultadoGestion, this.txtBuscar.Text.ToString (), (int)this.optCobradorOperador.Value);
			
			FuncionesProcedimientos.dsGeneral(sSQL, this.uGridInformacion);

			this.lblContador.Text = this.uGridInformacion.Rows.Count + " REGISTROS ENCONTRADOS";

			this.Cursor = Cursors.Default;

		}

		private void cmbGestionCredito_ValueChanged(object sender, System.EventArgs e)
		{
			if (this.cmbGestionCredito.ActiveRow != null)
			{
			
				if((int)this.cmbGestionCredito.Value == 1)
				{
					this.cmbResultadoGestion.DataSource = Funcion.dvProcedimiento(cdsSeteoF, string.Format("select idCbo_ResultadoGestion, Resultado from Cbo_ResultadoGestion where idCbo_EstadoGestion = {0}", (int) this.cmbGestionCredito.Value));
				
				}
				if((int)this.cmbGestionCredito.Value == 2)
				{
					this.cmbResultadoGestion.DataSource = Funcion.dvProcedimiento(cdsSeteoF, string.Format("select idCbo_ResultadoGestion, Resultado from Cbo_ResultadoGestion where idCbo_EstadoGestion = {0}", (int) this.cmbGestionCredito.Value));
			
				}
				if((int)this.cmbGestionCredito.Value == 3)
				{
					this.cmbResultadoGestion.DataSource = Funcion.dvProcedimiento(cdsSeteoF, string.Format("select idCbo_ResultadoGestion, Resultado from Cbo_ResultadoGestion where idCbo_EstadoGestion =  {0}", (int) this.cmbGestionCredito.Value));

				}

			}
		}

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);	
		}

		private void uGridInformacion_ClickCellButton(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			#region Operadora

			if (e.Cell.Column.ToString() == "Ir")
			{
				if (e.Cell.Row.Cells["idCompra"].Value != System.DBNull.Value  )
				{
					if ((int)e.Cell.Row.Cells["idCompra"].Value > 0 )
					{
						using (frmConsultaGeneralCallCenter Gestion = new frmConsultaGeneralCallCenter((int)e.Cell.Row.Cells["idCompra"].Value, 
										 e.Cell.Row.Cells["Cliente"].Value.ToString() ))
						{
							if (DialogResult.OK == Gestion.ShowDialog( ))
							{

	
							}
						}
					}
				}
			}

			#endregion Operadora
		}

		private void frmConsultaGestionesCallCenter_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
				if (e.KeyCode == Keys.Escape) this.Close();		
		}

		private void optEstados_ValueChanged(object sender, System.EventArgs e)
		{
			this.btnVer_Click (sender , e );
		}

		private void btnExcel_Click(object sender, System.EventArgs e)
		{
			FuncionesProcedimientos.ExportaExcel(this.uGridInformacion);
		}

		private void cmbCobrador_Operador_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.cmbGestionCredito.Focus();
		}

		private void cmbGestionCredito_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.cmbResultadoGestion.Focus();
		}

		private void cmbResultadoGestion_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.txtBuscar.Focus();
		}

		private void txtBuscar_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.btnVer_Click (sender, e);
		}

		private void optCobradorOperador_ValueChanged(object sender, System.EventArgs e)
		{
		this.btnVer_Click (sender , e );
		}
	}
}
