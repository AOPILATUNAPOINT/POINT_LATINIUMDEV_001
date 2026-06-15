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
using System.Threading;
//using CrystalDecisions;
using CrystalDecisions.CrystalReports.Engine;
using CrystalDecisions.ReportSource;
using CrystalDecisions.Shared;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de frmTranferenciaActivos.
	/// </summary>
	public class frmTranferenciaActivos : System.Windows.Forms.Form
	{
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource1;
		private System.Windows.Forms.Label lblUsuario;
		private System.Windows.Forms.Label label3;
		private System.Windows.Forms.Label label6;
		private System.Windows.Forms.Label label8;
		private System.Windows.Forms.Label label2;
		private System.Windows.Forms.Label label1;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtNotas;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbSolicita;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtidCompraTransf;
		private System.Windows.Forms.Button btnImprimir;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtEstado;
		private System.Windows.Forms.Button btnNuevo;
		private System.Windows.Forms.Button btnConsultar;
		private System.Windows.Forms.Button btnEditar;
		private System.Windows.Forms.Button btnGuardar;
		private System.Windows.Forms.Button btnAnular;
		private System.Windows.Forms.Button btnCancelar;
		private System.Windows.Forms.Button btnSalir;
		private System.Windows.Forms.GroupBox groupBox1;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbBodegaDestino;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtNumero;
		private Infragistics.Win.UltraWinGrid.UltraGrid ultraGrid1;
		private System.Windows.Forms.Label lblEstado;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtFecha;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtSerie;
		private System.Windows.Forms.Label label7;
		private C1.Data.C1DataSet cdsSeteoF;
		private Infragistics.Win.UltraWinGrid.UltraDropDown cmbCobradores;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource2;
		private System.Windows.Forms.Label lblVendedor;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbRealiza;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public frmTranferenciaActivos()
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
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idMovimientosActivos");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idActivo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn3 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idTipo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn4 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idPersonal");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn5 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idProyecto");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn6 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idPersonalRecibe");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn7 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idProyectoRecibe");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn8 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Estado");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn9 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Fecha");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn10 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCompraTrans");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn11 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idEstadoEquipo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn12 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Numero");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn13 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Codigo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn14 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Articulo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn15 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Serie");
			Infragistics.Win.Appearance appearance1 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPersonal");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmTranferenciaActivos));
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bodega");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand3 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idMovimientosActivos");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idActivo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idTipo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPersonal");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn11 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idProyecto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn12 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPersonalRecibe");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn13 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idProyectoRecibe");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn14 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Estado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn15 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Fecha");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn16 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCompraTrans");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn17 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idEstadoEquipo", -1, "cmbCobradores");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn18 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Numero");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn19 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn20 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Articulo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn21 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Serie");
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance11 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton1 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.Appearance appearance12 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand4 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn22 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idEstadoEquipo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn23 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn16 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idEstadoEquipo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn17 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.Appearance appearance13 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand5 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn24 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPersonal");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn25 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn26 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			this.ultraDataSource1 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.lblUsuario = new System.Windows.Forms.Label();
			this.label3 = new System.Windows.Forms.Label();
			this.label6 = new System.Windows.Forms.Label();
			this.label8 = new System.Windows.Forms.Label();
			this.label2 = new System.Windows.Forms.Label();
			this.label1 = new System.Windows.Forms.Label();
			this.txtNotas = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.cmbSolicita = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.txtidCompraTransf = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.btnImprimir = new System.Windows.Forms.Button();
			this.txtEstado = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.btnNuevo = new System.Windows.Forms.Button();
			this.btnConsultar = new System.Windows.Forms.Button();
			this.btnEditar = new System.Windows.Forms.Button();
			this.btnGuardar = new System.Windows.Forms.Button();
			this.btnAnular = new System.Windows.Forms.Button();
			this.btnCancelar = new System.Windows.Forms.Button();
			this.btnSalir = new System.Windows.Forms.Button();
			this.groupBox1 = new System.Windows.Forms.GroupBox();
			this.cmbBodegaDestino = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.txtNumero = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.ultraGrid1 = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.lblEstado = new System.Windows.Forms.Label();
			this.dtFecha = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.txtSerie = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label7 = new System.Windows.Forms.Label();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.cmbCobradores = new Infragistics.Win.UltraWinGrid.UltraDropDown();
			this.ultraDataSource2 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.lblVendedor = new System.Windows.Forms.Label();
			this.cmbRealiza = new Infragistics.Win.UltraWinGrid.UltraCombo();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNotas)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbSolicita)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtidCompraTransf)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtEstado)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbBodegaDestino)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNumero)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFecha)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtSerie)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCobradores)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbRealiza)).BeginInit();
			this.SuspendLayout();
			// 
			// ultraDataSource1
			// 
			ultraDataColumn1.DataType = typeof(int);
			ultraDataColumn2.DataType = typeof(int);
			ultraDataColumn3.DataType = typeof(int);
			ultraDataColumn4.DataType = typeof(int);
			ultraDataColumn5.DataType = typeof(int);
			ultraDataColumn6.DataType = typeof(int);
			ultraDataColumn7.DataType = typeof(int);
			ultraDataColumn8.DataType = typeof(int);
			ultraDataColumn9.DataType = typeof(System.DateTime);
			ultraDataColumn10.DataType = typeof(int);
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
																																 ultraDataColumn14,
																																 ultraDataColumn15});
			// 
			// lblUsuario
			// 
			this.lblUsuario.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.lblUsuario.AutoSize = true;
			this.lblUsuario.BackColor = System.Drawing.Color.Transparent;
			this.lblUsuario.Font = new System.Drawing.Font("Tahoma", 10F, System.Drawing.FontStyle.Bold);
			this.lblUsuario.Location = new System.Drawing.Point(528, 296);
			this.lblUsuario.Name = "lblUsuario";
			this.lblUsuario.Size = new System.Drawing.Size(0, 20);
			this.lblUsuario.TabIndex = 705;
			this.lblUsuario.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// label3
			// 
			this.label3.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.label3.AutoSize = true;
			this.label3.BackColor = System.Drawing.Color.Transparent;
			this.label3.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label3.Location = new System.Drawing.Point(16, 256);
			this.label3.Name = "label3";
			this.label3.Size = new System.Drawing.Size(33, 17);
			this.label3.TabIndex = 704;
			this.label3.Text = "Notas";
			this.label3.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label6
			// 
			this.label6.AutoSize = true;
			this.label6.Location = new System.Drawing.Point(376, 40);
			this.label6.Name = "label6";
			this.label6.Size = new System.Drawing.Size(46, 16);
			this.label6.TabIndex = 701;
			this.label6.Text = "Recepta";
			// 
			// label8
			// 
			this.label8.AutoSize = true;
			this.label8.BackColor = System.Drawing.Color.Transparent;
			this.label8.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label8.Location = new System.Drawing.Point(8, 40);
			this.label8.Name = "label8";
			this.label8.Size = new System.Drawing.Size(83, 17);
			this.label8.TabIndex = 686;
			this.label8.Text = "Bodega Destino";
			this.label8.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label2
			// 
			this.label2.AutoSize = true;
			this.label2.BackColor = System.Drawing.Color.Transparent;
			this.label2.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label2.Location = new System.Drawing.Point(520, 8);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(44, 17);
			this.label2.TabIndex = 684;
			this.label2.Text = "Número";
			this.label2.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label1
			// 
			this.label1.AutoSize = true;
			this.label1.BackColor = System.Drawing.Color.Transparent;
			this.label1.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label1.Location = new System.Drawing.Point(8, 8);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(34, 17);
			this.label1.TabIndex = 683;
			this.label1.Text = "Fecha";
			this.label1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtNotas
			// 
			this.txtNotas.AcceptsReturn = true;
			this.txtNotas.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			appearance1.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtNotas.Appearance = appearance1;
			this.txtNotas.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtNotas.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNotas.Enabled = false;
			this.txtNotas.Location = new System.Drawing.Point(16, 280);
			this.txtNotas.Multiline = true;
			this.txtNotas.Name = "txtNotas";
			this.txtNotas.Size = new System.Drawing.Size(497, 48);
			this.txtNotas.TabIndex = 703;
			// 
			// cmbSolicita
			// 
			appearance2.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbSolicita.Appearance = appearance2;
			this.cmbSolicita.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbSolicita.Cursor = System.Windows.Forms.Cursors.Default;
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridColumn2.Width = 335;
			ultraGridColumn3.Header.VisiblePosition = 2;
			ultraGridBand1.Columns.AddRange(new object[] {
																										 ultraGridColumn1,
																										 ultraGridColumn2,
																										 ultraGridColumn3});
			this.cmbSolicita.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.cmbSolicita.DisplayMember = "Nombre";
			this.cmbSolicita.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbSolicita.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbSolicita.Enabled = false;
			this.cmbSolicita.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbSolicita.Location = new System.Drawing.Point(440, 40);
			this.cmbSolicita.Name = "cmbSolicita";
			this.cmbSolicita.Size = new System.Drawing.Size(256, 21);
			this.cmbSolicita.TabIndex = 700;
			this.cmbSolicita.ValueMember = "idPersonal";
			// 
			// txtidCompraTransf
			// 
			this.txtidCompraTransf.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.txtidCompraTransf.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtidCompraTransf.Enabled = false;
			this.txtidCompraTransf.Location = new System.Drawing.Point(680, 232);
			this.txtidCompraTransf.Name = "txtidCompraTransf";
			this.txtidCompraTransf.PromptChar = ' ';
			this.txtidCompraTransf.Size = new System.Drawing.Size(16, 21);
			this.txtidCompraTransf.TabIndex = 698;
			this.txtidCompraTransf.Visible = false;
			// 
			// btnImprimir
			// 
			this.btnImprimir.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnImprimir.CausesValidation = false;
			this.btnImprimir.Enabled = false;
			this.btnImprimir.Image = ((System.Drawing.Image)(resources.GetObject("btnImprimir.Image")));
			this.btnImprimir.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnImprimir.Location = new System.Drawing.Point(256, 352);
			this.btnImprimir.Name = "btnImprimir";
			this.btnImprimir.Size = new System.Drawing.Size(77, 23);
			this.btnImprimir.TabIndex = 697;
			this.btnImprimir.Text = "&Imprimir";
			this.btnImprimir.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnImprimir.Click += new System.EventHandler(this.btnImprimir_Click);
			// 
			// txtEstado
			// 
			this.txtEstado.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.txtEstado.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtEstado.Enabled = false;
			this.txtEstado.Location = new System.Drawing.Point(680, 264);
			this.txtEstado.Name = "txtEstado";
			this.txtEstado.PromptChar = ' ';
			this.txtEstado.Size = new System.Drawing.Size(16, 21);
			this.txtEstado.TabIndex = 696;
			this.txtEstado.Visible = false;
			// 
			// btnNuevo
			// 
			this.btnNuevo.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnNuevo.CausesValidation = false;
			this.btnNuevo.Image = ((System.Drawing.Image)(resources.GetObject("btnNuevo.Image")));
			this.btnNuevo.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnNuevo.Location = new System.Drawing.Point(16, 352);
			this.btnNuevo.Name = "btnNuevo";
			this.btnNuevo.Size = new System.Drawing.Size(72, 23);
			this.btnNuevo.TabIndex = 695;
			this.btnNuevo.Text = "&Nuevo";
			this.btnNuevo.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnNuevo.Click += new System.EventHandler(this.btnNuevo_Click);
			// 
			// btnConsultar
			// 
			this.btnConsultar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnConsultar.CausesValidation = false;
			this.btnConsultar.Image = ((System.Drawing.Image)(resources.GetObject("btnConsultar.Image")));
			this.btnConsultar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnConsultar.Location = new System.Drawing.Point(96, 352);
			this.btnConsultar.Name = "btnConsultar";
			this.btnConsultar.Size = new System.Drawing.Size(78, 23);
			this.btnConsultar.TabIndex = 694;
			this.btnConsultar.Text = "&Consultar";
			this.btnConsultar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnConsultar.Click += new System.EventHandler(this.btnConsultar_Click);
			// 
			// btnEditar
			// 
			this.btnEditar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnEditar.CausesValidation = false;
			this.btnEditar.Enabled = false;
			this.btnEditar.Image = ((System.Drawing.Image)(resources.GetObject("btnEditar.Image")));
			this.btnEditar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnEditar.Location = new System.Drawing.Point(176, 352);
			this.btnEditar.Name = "btnEditar";
			this.btnEditar.Size = new System.Drawing.Size(76, 23);
			this.btnEditar.TabIndex = 693;
			this.btnEditar.Text = "&Editar";
			this.btnEditar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnEditar.Click += new System.EventHandler(this.btnEditar_Click);
			// 
			// btnGuardar
			// 
			this.btnGuardar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnGuardar.Enabled = false;
			this.btnGuardar.Image = ((System.Drawing.Image)(resources.GetObject("btnGuardar.Image")));
			this.btnGuardar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnGuardar.Location = new System.Drawing.Point(336, 352);
			this.btnGuardar.Name = "btnGuardar";
			this.btnGuardar.Size = new System.Drawing.Size(76, 23);
			this.btnGuardar.TabIndex = 692;
			this.btnGuardar.Text = "&Guardar";
			this.btnGuardar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnGuardar.Click += new System.EventHandler(this.btnGuardar_Click);
			// 
			// btnAnular
			// 
			this.btnAnular.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnAnular.CausesValidation = false;
			this.btnAnular.Enabled = false;
			this.btnAnular.Image = ((System.Drawing.Image)(resources.GetObject("btnAnular.Image")));
			this.btnAnular.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnAnular.Location = new System.Drawing.Point(416, 352);
			this.btnAnular.Name = "btnAnular";
			this.btnAnular.Size = new System.Drawing.Size(76, 23);
			this.btnAnular.TabIndex = 691;
			this.btnAnular.Text = "&Anular";
			this.btnAnular.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// btnCancelar
			// 
			this.btnCancelar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnCancelar.CausesValidation = false;
			this.btnCancelar.Enabled = false;
			this.btnCancelar.Image = ((System.Drawing.Image)(resources.GetObject("btnCancelar.Image")));
			this.btnCancelar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnCancelar.Location = new System.Drawing.Point(496, 352);
			this.btnCancelar.Name = "btnCancelar";
			this.btnCancelar.Size = new System.Drawing.Size(72, 23);
			this.btnCancelar.TabIndex = 690;
			this.btnCancelar.Text = "&Cancelar";
			this.btnCancelar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnCancelar.Click += new System.EventHandler(this.btnCancelar_Click);
			// 
			// btnSalir
			// 
			this.btnSalir.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnSalir.CausesValidation = false;
			this.btnSalir.Image = ((System.Drawing.Image)(resources.GetObject("btnSalir.Image")));
			this.btnSalir.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnSalir.Location = new System.Drawing.Point(576, 352);
			this.btnSalir.Name = "btnSalir";
			this.btnSalir.TabIndex = 689;
			this.btnSalir.Text = "&Salir";
			this.btnSalir.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnSalir.Click += new System.EventHandler(this.btnSalir_Click);
			// 
			// groupBox1
			// 
			this.groupBox1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.groupBox1.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.groupBox1.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(50)), ((System.Byte)(50)), ((System.Byte)(50)));
			this.groupBox1.Location = new System.Drawing.Point(16, 336);
			this.groupBox1.Name = "groupBox1";
			this.groupBox1.Size = new System.Drawing.Size(680, 8);
			this.groupBox1.TabIndex = 688;
			this.groupBox1.TabStop = false;
			// 
			// cmbBodegaDestino
			// 
			appearance3.ForeColor = System.Drawing.Color.Black;
			appearance3.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbBodegaDestino.Appearance = appearance3;
			this.cmbBodegaDestino.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbBodegaDestino.Cursor = System.Windows.Forms.Cursors.Default;
			ultraGridColumn4.Header.VisiblePosition = 0;
			ultraGridColumn4.Width = 250;
			ultraGridColumn5.Header.VisiblePosition = 1;
			ultraGridColumn5.Hidden = true;
			ultraGridColumn5.Width = 50;
			ultraGridColumn6.Header.VisiblePosition = 2;
			ultraGridColumn6.Hidden = true;
			ultraGridColumn6.Width = 50;
			ultraGridBand2.Columns.AddRange(new object[] {
																										 ultraGridColumn4,
																										 ultraGridColumn5,
																										 ultraGridColumn6});
			this.cmbBodegaDestino.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			this.cmbBodegaDestino.DisplayMember = "Nombre";
			this.cmbBodegaDestino.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbBodegaDestino.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbBodegaDestino.Enabled = false;
			this.cmbBodegaDestino.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbBodegaDestino.Location = new System.Drawing.Point(104, 40);
			this.cmbBodegaDestino.MaxDropDownItems = 30;
			this.cmbBodegaDestino.Name = "cmbBodegaDestino";
			this.cmbBodegaDestino.Size = new System.Drawing.Size(250, 21);
			this.cmbBodegaDestino.TabIndex = 680;
			this.cmbBodegaDestino.ValueMember = "Bodega";
			this.cmbBodegaDestino.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cmbBodegaDestino_KeyDown);
			this.cmbBodegaDestino.Validating += new System.ComponentModel.CancelEventHandler(this.cmbBodegaDestino_Validating);
			this.cmbBodegaDestino.ValueChanged += new System.EventHandler(this.cmbBodegaDestino_ValueChanged);
			// 
			// txtNumero
			// 
			appearance4.ForeColor = System.Drawing.Color.Black;
			appearance4.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtNumero.Appearance = appearance4;
			this.txtNumero.Enabled = false;
			this.txtNumero.Location = new System.Drawing.Point(584, 8);
			this.txtNumero.Name = "txtNumero";
			this.txtNumero.Size = new System.Drawing.Size(112, 21);
			this.txtNumero.TabIndex = 677;
			// 
			// ultraGrid1
			// 
			this.ultraGrid1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.ultraGrid1.Cursor = System.Windows.Forms.Cursors.Default;
			this.ultraGrid1.DataSource = this.ultraDataSource1;
			appearance5.BackColor = System.Drawing.Color.White;
			this.ultraGrid1.DisplayLayout.Appearance = appearance5;
			this.ultraGrid1.DisplayLayout.AutoFitColumns = true;
			ultraGridColumn7.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn7.Header.VisiblePosition = 0;
			ultraGridColumn7.Hidden = true;
			ultraGridColumn7.Width = 142;
			ultraGridColumn8.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn8.Header.VisiblePosition = 1;
			ultraGridColumn8.Hidden = true;
			ultraGridColumn8.Width = 32;
			ultraGridColumn9.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn9.Header.VisiblePosition = 2;
			ultraGridColumn9.Hidden = true;
			ultraGridColumn9.Width = 32;
			ultraGridColumn10.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn10.Header.VisiblePosition = 3;
			ultraGridColumn10.Hidden = true;
			ultraGridColumn10.Width = 37;
			ultraGridColumn11.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn11.Header.VisiblePosition = 4;
			ultraGridColumn11.Hidden = true;
			ultraGridColumn11.Width = 37;
			ultraGridColumn12.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn12.Header.VisiblePosition = 5;
			ultraGridColumn12.Hidden = true;
			ultraGridColumn12.Width = 56;
			ultraGridColumn13.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn13.Header.VisiblePosition = 6;
			ultraGridColumn13.Hidden = true;
			ultraGridColumn13.Width = 56;
			ultraGridColumn14.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn14.Header.VisiblePosition = 7;
			ultraGridColumn14.Hidden = true;
			ultraGridColumn14.Width = 129;
			ultraGridColumn15.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn15.Header.VisiblePosition = 8;
			ultraGridColumn15.Hidden = true;
			ultraGridColumn15.Width = 123;
			ultraGridColumn16.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn16.Header.VisiblePosition = 9;
			ultraGridColumn16.Hidden = true;
			ultraGridColumn16.Width = 113;
			ultraGridColumn17.ButtonDisplayStyle = Infragistics.Win.UltraWinGrid.ButtonDisplayStyle.Always;
			ultraGridColumn17.Header.Caption = "Estado";
			ultraGridColumn17.Header.VisiblePosition = 14;
			ultraGridColumn17.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDown;
			ultraGridColumn17.Width = 126;
			ultraGridColumn18.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn18.Header.VisiblePosition = 12;
			ultraGridColumn18.Hidden = true;
			ultraGridColumn18.Width = 117;
			ultraGridColumn19.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn19.Header.VisiblePosition = 10;
			ultraGridColumn19.Width = 163;
			ultraGridColumn20.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn20.Header.VisiblePosition = 11;
			ultraGridColumn20.Width = 172;
			ultraGridColumn21.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn21.Header.VisiblePosition = 13;
			ultraGridColumn21.Width = 206;
			ultraGridBand3.Columns.AddRange(new object[] {
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
																										 ultraGridColumn21});
			this.ultraGrid1.DisplayLayout.BandsSerializer.Add(ultraGridBand3);
			this.ultraGrid1.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.True;
			this.ultraGrid1.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance6.BackColor = System.Drawing.Color.Transparent;
			this.ultraGrid1.DisplayLayout.Override.CardAreaAppearance = appearance6;
			appearance7.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance7.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance7.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance7.FontData.BoldAsString = "True";
			appearance7.FontData.Name = "Arial";
			appearance7.FontData.SizeInPoints = 9F;
			appearance7.ForeColor = System.Drawing.Color.White;
			appearance7.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.ultraGrid1.DisplayLayout.Override.HeaderAppearance = appearance7;
			appearance8.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance8.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance8.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.RowAlternateAppearance = appearance8;
			appearance9.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance9.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance9.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.RowSelectorAppearance = appearance9;
			appearance10.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance10.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance10.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.SelectedRowAppearance = appearance10;
			this.ultraGrid1.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.ultraGrid1.Location = new System.Drawing.Point(8, 104);
			this.ultraGrid1.Name = "ultraGrid1";
			this.ultraGrid1.Size = new System.Drawing.Size(688, 144);
			this.ultraGrid1.TabIndex = 682;
			this.ultraGrid1.AfterRowInsert += new Infragistics.Win.UltraWinGrid.RowEventHandler(this.ultraGrid1_AfterRowInsert);
			// 
			// lblEstado
			// 
			this.lblEstado.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.lblEstado.AutoSize = true;
			this.lblEstado.BackColor = System.Drawing.Color.Transparent;
			this.lblEstado.Font = new System.Drawing.Font("Microsoft Sans Serif", 15F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblEstado.ForeColor = System.Drawing.Color.Firebrick;
			this.lblEstado.Location = new System.Drawing.Point(528, 248);
			this.lblEstado.Name = "lblEstado";
			this.lblEstado.Size = new System.Drawing.Size(0, 26);
			this.lblEstado.TabIndex = 687;
			this.lblEstado.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// dtFecha
			// 
			this.dtFecha.AllowMonthSelection = true;
			this.dtFecha.AllowWeekSelection = true;
			appearance11.ForeColor = System.Drawing.Color.Black;
			appearance11.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtFecha.Appearance = appearance11;
			this.dtFecha.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton1.Caption = "Today";
			this.dtFecha.DateButtons.Add(dateButton1);
			this.dtFecha.Enabled = false;
			this.dtFecha.Format = "dd/MM/yyyy";
			this.dtFecha.Location = new System.Drawing.Point(104, 8);
			this.dtFecha.Name = "dtFecha";
			this.dtFecha.NonAutoSizeHeight = 23;
			this.dtFecha.Size = new System.Drawing.Size(136, 21);
			this.dtFecha.SpinButtonsVisible = true;
			this.dtFecha.TabIndex = 678;
			this.dtFecha.Value = ((object)(resources.GetObject("dtFecha.Value")));
			this.dtFecha.KeyDown += new System.Windows.Forms.KeyEventHandler(this.dtFecha_KeyDown);
			this.dtFecha.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.dtFecha_KeyPress);
			this.dtFecha.Validating += new System.ComponentModel.CancelEventHandler(this.dtFecha_Validating);
			this.dtFecha.Validated += new System.EventHandler(this.dtFecha_Validated);
			this.dtFecha.ValueChanged += new System.EventHandler(this.dtFecha_ValueChanged);
			// 
			// txtSerie
			// 
			appearance12.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtSerie.Appearance = appearance12;
			this.txtSerie.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtSerie.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtSerie.Enabled = false;
			this.txtSerie.Location = new System.Drawing.Point(104, 72);
			this.txtSerie.MaxLength = 100;
			this.txtSerie.Name = "txtSerie";
			this.txtSerie.Size = new System.Drawing.Size(248, 21);
			this.txtSerie.TabIndex = 1037;
			this.txtSerie.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtSerie_KeyDown);
			this.txtSerie.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtSerie_KeyPress);
			// 
			// label7
			// 
			this.label7.AutoSize = true;
			this.label7.BackColor = System.Drawing.Color.Transparent;
			this.label7.Enabled = false;
			this.label7.Location = new System.Drawing.Point(8, 72);
			this.label7.Name = "label7";
			this.label7.Size = new System.Drawing.Size(31, 16);
			this.label7.TabIndex = 1038;
			this.label7.Text = "Serie";
			this.label7.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
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
			// cmbCobradores
			// 
			this.cmbCobradores.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbCobradores.DataSource = this.ultraDataSource2;
			ultraGridColumn22.Header.VisiblePosition = 0;
			ultraGridColumn22.Hidden = true;
			ultraGridColumn22.Width = 180;
			ultraGridColumn23.Header.VisiblePosition = 1;
			ultraGridBand4.Columns.AddRange(new object[] {
																										 ultraGridColumn22,
																										 ultraGridColumn23});
			this.cmbCobradores.DisplayLayout.BandsSerializer.Add(ultraGridBand4);
			this.cmbCobradores.DisplayMember = "Nombre";
			this.cmbCobradores.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbCobradores.Location = new System.Drawing.Point(300, 215);
			this.cmbCobradores.Name = "cmbCobradores";
			this.cmbCobradores.Size = new System.Drawing.Size(112, 56);
			this.cmbCobradores.TabIndex = 1039;
			this.cmbCobradores.ValueMember = "idEstadoEquipo";
			this.cmbCobradores.Visible = false;
			// 
			// ultraDataSource2
			// 
			ultraDataColumn16.DataType = typeof(int);
			this.ultraDataSource2.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn16,
																																 ultraDataColumn17});
			// 
			// lblVendedor
			// 
			this.lblVendedor.AutoSize = true;
			this.lblVendedor.Location = new System.Drawing.Point(376, 72);
			this.lblVendedor.Name = "lblVendedor";
			this.lblVendedor.Size = new System.Drawing.Size(42, 16);
			this.lblVendedor.TabIndex = 1040;
			this.lblVendedor.Text = "Realiza";
			// 
			// cmbRealiza
			// 
			appearance13.ForeColor = System.Drawing.Color.Black;
			appearance13.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbRealiza.Appearance = appearance13;
			this.cmbRealiza.CausesValidation = false;
			this.cmbRealiza.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbRealiza.Cursor = System.Windows.Forms.Cursors.Default;
			ultraGridColumn24.Header.VisiblePosition = 0;
			ultraGridColumn24.Hidden = true;
			ultraGridColumn25.Header.VisiblePosition = 1;
			ultraGridColumn26.Header.VisiblePosition = 2;
			ultraGridColumn26.Width = 300;
			ultraGridBand5.Columns.AddRange(new object[] {
																										 ultraGridColumn24,
																										 ultraGridColumn25,
																										 ultraGridColumn26});
			this.cmbRealiza.DisplayLayout.BandsSerializer.Add(ultraGridBand5);
			this.cmbRealiza.DisplayMember = "Nombre";
			this.cmbRealiza.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbRealiza.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbRealiza.Enabled = false;
			this.cmbRealiza.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbRealiza.Location = new System.Drawing.Point(440, 72);
			this.cmbRealiza.MaxDropDownItems = 30;
			this.cmbRealiza.Name = "cmbRealiza";
			this.cmbRealiza.Size = new System.Drawing.Size(256, 21);
			this.cmbRealiza.TabIndex = 1041;
			this.cmbRealiza.ValueMember = "idPersonal";
			// 
			// frmTranferenciaActivos
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(704, 382);
			this.Controls.Add(this.lblVendedor);
			this.Controls.Add(this.label7);
			this.Controls.Add(this.label3);
			this.Controls.Add(this.label6);
			this.Controls.Add(this.label8);
			this.Controls.Add(this.label2);
			this.Controls.Add(this.label1);
			this.Controls.Add(this.cmbRealiza);
			this.Controls.Add(this.cmbCobradores);
			this.Controls.Add(this.txtSerie);
			this.Controls.Add(this.lblUsuario);
			this.Controls.Add(this.txtNotas);
			this.Controls.Add(this.cmbSolicita);
			this.Controls.Add(this.txtidCompraTransf);
			this.Controls.Add(this.btnImprimir);
			this.Controls.Add(this.txtEstado);
			this.Controls.Add(this.btnNuevo);
			this.Controls.Add(this.btnConsultar);
			this.Controls.Add(this.btnEditar);
			this.Controls.Add(this.btnGuardar);
			this.Controls.Add(this.btnAnular);
			this.Controls.Add(this.btnCancelar);
			this.Controls.Add(this.btnSalir);
			this.Controls.Add(this.groupBox1);
			this.Controls.Add(this.cmbBodegaDestino);
			this.Controls.Add(this.txtNumero);
			this.Controls.Add(this.ultraGrid1);
			this.Controls.Add(this.lblEstado);
			this.Controls.Add(this.dtFecha);
			this.Name = "frmTranferenciaActivos";
			this.Text = "Acta De Entrega Activos";
			this.Load += new System.EventHandler(this.frmTranferenciaActivos_Load);
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNotas)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbSolicita)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtidCompraTransf)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtEstado)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbBodegaDestino)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNumero)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFecha)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtSerie)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCobradores)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbRealiza)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void uGridDetalle_AfterRowInsert(object sender, Infragistics.Win.UltraWinGrid.RowEventArgs e)
		{
			e.Row.Cells["idMovimientosActivos"].Value = 0;
			e.Row.Cells["idActivo"].Value = 0;
			e.Row.Cells["idTipo"].Value = 0;
			e.Row.Cells["idPersonal"].Value = 0;
			e.Row.Cells["idProyecto"].Value = 0;
			e.Row.Cells["idPersonalRecibe"].Value = 0;
			e.Row.Cells["idProyectoRecibe"].Value = 0;
			e.Row.Cells["idEstadoEquipo"].Value = 0;

		}

		private void btnNuevo_Click(object sender, System.EventArgs e)
		{
			this.btnCancelar_Click(sender, e);						
					
			if (miAcceso.CambiarFecha) this.dtFecha.Enabled = true;
						
			FuncionesProcedimientos.RetornaFechaServidor(this.dtFecha, cdsSeteoF, false);
									
			this.cmbBodegaDestino.Enabled = true;			
			FuncionesProcedimientos.EstadoGrids(true, this.ultraGrid1);
			this.txtNotas.Enabled = true;
			this.txtSerie.Enabled = true;
						
			bNuevo = true;
			bEdicion = true;

			if (iBodegaPredef > 0) 
			{
				this.txtEstado.Value = 1;
				this.lblEstado.Text = "PENDIENTE";
			}

			if (this.dtFecha.Value != System.DBNull.Value)
			{	

		
				this.cmbRealiza.Value = Funcion.iEscalarSQL(cdsSeteoF, string.Format("If Exists(Select idPersonal From Nomina Where Codigo In (Select Nombre From Usuario where idUsuario = {0})) Select idPersonal From Nomina Where Codigo In (Select Nombre From Usuario where idUsuario = {0}) Else Select 0", MenuLatinium.IdUsuario));					
			 


			}

			
			
			
			if (iBodegaPredef > 0)
			{
				if (Funcion.iEscalarSQL(cdsSeteoF, string.Format("Select COUNT(*) From CierreConteoInventario Where Bodega = {0} And CONVERT(DATE, Fecha) Between '20161001' And '{1}' And Estado = 1", 
					(int)this.cmbBodegaDestino.Value, Convert.ToDateTime(this.dtFecha.Value).ToString("yyyyMMdd"))) > 0)
				{
					MessageBox.Show("El local esta en Inventario no puede crear Transferencia", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					this.btnCancelar_Click(sender, e);						
					return;
				}
			}

			this.btnNuevo.Enabled = false;
			this.btnConsultar.Enabled = false;
			this.btnEditar.Enabled = false;
			this.btnImprimir.Enabled = false;
			this.btnGuardar.Enabled = true;
			this.btnAnular.Enabled = false;
			this.btnCancelar.Enabled = true;
			this.cmbSolicita.Enabled = true;		
			
			if (miAcceso.CambiarFecha) this.dtFecha.Focus(); 
			else if (this.cmbBodegaDestino.Enabled) this.cmbBodegaDestino.Focus(); else this.cmbBodegaDestino.Focus();
		}
		bool bNuevo = false;
		bool bEdicion = false;
		private Acceso miAcceso;
		bool bEliminaAlValidar = false;		
		int idBloqueaTransacciones = 0;
		int iBodegaPredef = 0;
		private bool bActualiza = true;
		bool bCombo = false;

		private void frmTranferenciaActivos_Load(object sender, System.EventArgs e)
		{
			miAcceso = new Acceso(cdsSeteoF, "2501");
			
			string sSQLBP = string.Format("Exec BodegaAsignadaPorUsuario {0}", MenuLatinium.IdUsuario);		
			iBodegaPredef = Funcion.iEscalarSQL(cdsSeteoF, sSQLBP);

			FuncionesProcedimientos.EstadoGrids(false, this.ultraGrid1);

			string sSQLE = string.Format("SELECT idEstadoEquipo, Nombre FROM EstadoEquipo ORDER BY Nombre Desc");
			this.cmbCobradores.DataSource = Funcion.dvProcedimiento(cdsSeteoF, sSQLE);

			this.cmbRealiza.Value = Funcion.iEscalarSQL(cdsSeteoF, string.Format("If Exists(Select idPersonal From Nomina Where Codigo In (Select Nombre From Usuario where idUsuario = {0})) Select idPersonal From Nomina Where Codigo In (Select Nombre From Usuario where idUsuario = {0}) Else Select 0", MenuLatinium.IdUsuario));					
			 

		}

		private void dtFecha_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Delete) e.Handled = true;
			if (e.KeyCode == Keys.Enter) this.cmbBodegaDestino.Focus();
		}

		private void dtFecha_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (e.KeyChar == 8) e.Handled = true;
		}

		private void dtFecha_Validated(object sender, System.EventArgs e)
		{
			if (this.dtFecha.Value == System.DBNull.Value) this.dtFecha.Value = DateTime.Today;
		}

		private void dtFecha_Validating(object sender, System.ComponentModel.CancelEventArgs e)
		{
			if (bEdicion)
			{
				if (this.dtFecha.Value == DBNull.Value)
				{
					MessageBox.Show("Ingrese la Fecha", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					this.dtFecha.Focus();
					e.Cancel = true;				
				}
				else if (Funcion.ValidaPeriodo((DateTime)this.dtFecha.Value, cdsSeteoF, "Bodega"))
				{
					DateTime dtFecha = (DateTime)this.dtFecha.Value;
					MessageBox.Show(string.Format("No puede ingresar Transferencias en esta Fecha, El Período '{0}' esta Cerrado, \n\n Consulte al Administrador", dtFecha.ToString("MMMM").ToUpper()), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					this.dtFecha.Focus();
					e.Cancel = true;					
				}
			}
		}

		private void dtFecha_ValueChanged(object sender, System.EventArgs e)
		{
			if (this.dtFecha.Value != System.DBNull.Value)
			{
				FuncionesProcedimientos.ListaBodegas(this.cmbBodegaDestino, MenuLatinium.IdUsuario, 42, (DateTime)this.dtFecha.Value, false, cdsSeteoF);	
				string sSQLEc = string.Format("Select idPersonal, (n.ApellidoPaterno + ' ' + n.ApellidoMaterno + ' ' + n.PrimerNombre + ' ' + n.SegundoNombre) As Nombre, n.Codigo From Nomina n Where  ISNULL((Select Top 1 FechaSalida From NominaEntrada Where idNomina = n.idNomina Order By FechaEntrada Desc), '20301231') > '{0}' Order By Nombre	", Convert.ToDateTime(this.dtFecha.Value).ToString("yyyyMMdd"));
				this.cmbRealiza.DataSource = Funcion.dvProcedimiento(cdsSeteoF, sSQLEc);

			}
		}

		private void ultraGrid1_AfterRowInsert(object sender, Infragistics.Win.UltraWinGrid.RowEventArgs e)
		{
			if (bEdicion)
			{
				e.Row.Cells["idMovimientosActivos"].Value = 0;
				e.Row.Cells["idActivo"].Value = 0;
				e.Row.Cells["idTipo"].Value = 1;
				e.Row.Cells["idPersonal"].Value = 0;
				e.Row.Cells["idProyecto"].Value = 0;	
				e.Row.Cells["idPersonalRecibe"].Value = 0;	
				e.Row.Cells["idProyectoRecibe"].Value = 0;
				e.Row.Cells["idCompraTrans"].Value = 0;	
				e.Row.Cells["idEstadoEquipo"].Value = 0;	
			}
		}

		private void btnCancelar_Click(object sender, System.EventArgs e)
		{
			if (bEdicion && !bNuevo)
			{
				string sSQLBloqueo = string.Format("Exec RegistraBloqueoTransacciones {0}, 0, {1}, 3", idBloqueaTransacciones, (int)this.txtidCompraTransf.Value);
				Funcion.EjecutaSQL(cdsSeteoF, sSQLBloqueo, true);
			}

			bEdicion = false;
			bNuevo = false;

			this.btnNuevo.Focus();

			this.txtidCompraTransf.Value = 0;
			this.txtEstado.Value = 0;
			this.txtNumero.Text = "";
			this.txtSerie.Text ="";
			this.dtFecha.Value = System.DBNull.Value;
			this.dtFecha.CalendarInfo.MaxDate = DateTime.Today;


			this.cmbBodegaDestino.Value = System.DBNull.Value;
			this.cmbSolicita.Value = System.DBNull.Value;
			this.cmbRealiza.Value = System.DBNull.Value;

			this.lblEstado.Text = "";
			this.lblUsuario.Text = "";
			this.txtNotas.Text = "";						
			
			this.ultraGrid1.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Exec ConsultaMovimientosActivosGrabaDetalle 0");

			this.dtFecha.Enabled = false;

			this.cmbBodegaDestino.Enabled = false;
      this.cmbRealiza.Enabled = false;
			this.cmbSolicita.Enabled = false;
		
			this.txtNotas.Enabled = false;
				this.txtSerie.Enabled = false;
			
			FuncionesProcedimientos.EstadoGrids(false, this.ultraGrid1);

			if (miAcceso.Nuevo) this.btnNuevo.Enabled = true;
			if (miAcceso.Buscar) this.btnConsultar.Enabled = true;
			this.btnEditar.Enabled = false;
			this.btnImprimir.Enabled = false;
			this.btnGuardar.Enabled = false;
			this.btnAnular.Enabled = false;
			this.btnCancelar.Enabled = false;

			idBloqueaTransacciones = 0;
		}

		private void btnEditar_Click(object sender, System.EventArgs e)
		{
			this.Consultar((int)this.txtidCompraTransf.Value);

			if ((int)this.txtEstado.Value == 5)
			{
				MessageBox.Show("La Transferencia esta Anulada", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);				
			}
		  if (Funcion.ValidaPeriodo((DateTime)this.dtFecha.Value, cdsSeteoF, "Bodega"))
			{
				DateTime dtFechaVP = (DateTime)this.dtFecha.Value;
				MessageBox.Show(string.Format("No puede Editar esta Transferencia, El Período '{0}' esta Cerrado,\n\n Consulte al Administrador", dtFechaVP.ToString("MMMM").ToUpper()), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
			}			
			
		 if ((int)this.txtEstado.Value == 4)
			{
				MessageBox.Show("No puede Editar esta  ya esta Procesada", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);				
			}
		 if (!Funcion.ValidaBloqueo((int) this.txtidCompraTransf.Value, 3, cdsSeteoF))
			{
				if ((int)this.txtEstado.Value ==1 )
					if (miAcceso.CambiarFecha) this.dtFecha.Enabled = true;

//					if ((int)this.txtEstado.Value == 0 || (int)this.txtEstado.Value == 1) this.chkActivar.Enabled = true;

//					if (iBodegaPredef > 0 && Funcion.iEscalarSQL(cdsSeteoF, string.Format("Select idCom_Empresa From Com_Empresa Where Bodega = {0}", (int)this.cmbBodegaOrigen.Value)) > 0) this.cmbRealiza.Enabled = true;
//					this.cmbSolicita.Enabled = true;

					FuncionesProcedimientos.EstadoGrids(true, this.ultraGrid1);

					this.txtNotas.Enabled = true;
			 this.cmbBodegaDestino.Enabled = true;
			 this.cmbSolicita.Enabled = true;
			 this.txtSerie.Enabled = true;
			 this.txtNotas.Enabled = true;
									
					this.btnNuevo.Enabled = false;
					this.btnConsultar.Enabled = false;
					this.btnEditar.Enabled = false;
					this.btnImprimir.Enabled = false;
					this.btnGuardar.Enabled = true;
					this.btnAnular.Enabled = false;
					this.btnCancelar.Enabled = true;

					string sSQLBloqueo = string.Format("Exec RegistraBloqueoTransacciones 0, 1, {0}, 3", (int)this.txtidCompraTransf.Value);
					idBloqueaTransacciones = Funcion.iEscalarSQL(cdsSeteoF, sSQLBloqueo, true);

					bNuevo = false;
					bEdicion = true;
				
			}
		}

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);		
	
		}

		private void txtSerie_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter || e.KeyCode == Keys.Tab)
			{
				if (!Validacion.vbComboVacio(this.cmbBodegaDestino, "Seleccion la Bodega de Destino")) return;
				if (!Validacion.vbComboVacio(this.cmbSolicita, "Seleccione el Personal")) return;
				int iContGrid = (int) this.ultraGrid1.Rows.Count;

				if (iContGrid > 0)
				{
					foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.ultraGrid1.Rows.All)
					{	
						if (Convert.ToString (dr.Cells["Serie"].Value) == Convert.ToString (this.txtSerie.Text))
						{
							MessageBox.Show("La serie ya fue registrada", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
							this.ultraGrid1.ActiveRow = dr;
							this.ultraGrid1.ActiveRow.Selected = true;
							return;
						}		
					}
				}
        

				#region Carga Grid de Seriales
				string sSQL = string.Format("Exec ValidaActivoFijo '{0}', 0", this.txtSerie.Text.ToString());
				int iCtaSerial = Funcion.iEscalarSQL(cdsSeteoF, sSQL);
        
				if (iCtaSerial > 0)
				{	
					int idMovimientosActivos= 0;
					int idActivo = 0;	
					int idTipo	= 0;
					int idPersonal = 0;	
					int idProyecto = 0;	 
					string Codigo = "";
					string 	Articulo = "";

					string sSQLIdSerial = string.Format("Exec ValidaActivoFijo '{0}',1 ", this.txtSerie.Text.ToString());
					SqlDataReader drIdSerial = Funcion.rEscalarSQL(cdsSeteoF, sSQLIdSerial, true);
					drIdSerial.Read();
					if (drIdSerial.HasRows)
					{
						idMovimientosActivos = drIdSerial.GetInt32(0);
						idActivo = drIdSerial.GetInt32(1);
						idTipo	= drIdSerial.GetInt32(2);
						idPersonal = drIdSerial.GetInt32(3);	
						idProyecto = drIdSerial.GetInt32(4);	 
						Codigo = drIdSerial.GetString(5);	 
						Articulo = drIdSerial.GetString(6);	 

					}
					drIdSerial.Close();												
					this.ultraGrid1.Rows.Band.AddNew();
					this.ultraGrid1.ActiveRow.Cells["idMovimientosActivos"].Value = idMovimientosActivos;
					this.ultraGrid1.ActiveRow.Cells["idActivo"].Value = idActivo;
					this.ultraGrid1.ActiveRow.Cells["idTipo"].Value = idTipo;
					this.ultraGrid1.ActiveRow.Cells["idPersonal"].Value = idPersonal;
					this.ultraGrid1.ActiveRow.Cells["idProyecto"].Value = idProyecto;	
					this.ultraGrid1.ActiveRow.Cells["Codigo"].Value = Codigo;	
					this.ultraGrid1.ActiveRow.Cells["Articulo"].Value = Articulo;
					this.ultraGrid1.ActiveRow.Cells["idPersonalRecibe"].Value = (int) this.cmbSolicita.Value ;
					this.ultraGrid1.ActiveRow.Cells["Fecha"].Value = Convert.ToDateTime(this.dtFecha.Value);
					this.ultraGrid1.ActiveRow.Cells["idProyectoRecibe"].Value = (int) this.cmbBodegaDestino.Value ;
					this.ultraGrid1.ActiveRow.Cells["Serie"].Value = (string) this.txtSerie.Value ;



				}
				else
				{
					MessageBox.Show("La serial no existe n\n\nComuniquese con el Administrador", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
					this.txtSerie.Focus();
					return;
				}
			
				this.txtSerie.Text = "";
				#endregion Carga Grid de Seriales

				this.txtSerie.Text = "";

				this.txtSerie.Focus();


			}
		}

		private void txtSerie_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (Char.IsControl(e.KeyChar))
			{
				e.Handled = false;
			}
			else if (Char.IsLetterOrDigit(e.KeyChar))
			{
				e.Handled = false;
			}			
			else
			{
				e.Handled = true;
			}
		}

		private void cmbBodegaDestino_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
				if (e.KeyCode == Keys.Enter) this.cmbSolicita.Focus();
		}

		private void cmbBodegaDestino_Validating(object sender, System.ComponentModel.CancelEventArgs e)
		{
			if (bEdicion)
			{
				if ( this.cmbBodegaDestino.ActiveRow == null) return;
			
				if (this.cmbBodegaDestino.ActiveRow == null)
				{
					MessageBox.Show("Seleccione la Bodega Destino", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					this.cmbBodegaDestino.Focus();
					e.Cancel = true;			
				}
				
			}
		}

		private void cmbBodegaDestino_ValueChanged(object sender, System.EventArgs e)
		{
			if (bEdicion) this.cmbSolicita.Value = System.DBNull.Value;

			if (this.cmbBodegaDestino.ActiveRow != null && this.dtFecha.Value != System.DBNull.Value)
				this.cmbSolicita.DataSource = Funcion.dvProcedimiento(cdsSeteoF, string.Format("Exec PersonalTransferencias '{0}', {1}, {2}, 1", 
					Convert.ToDateTime(this.dtFecha.Value).ToString("yyyyMMdd"), (int)this.cmbBodegaDestino.Value, MenuLatinium.IdUsuario));
	
		}

		private void btnConsultar_Click(object sender, System.EventArgs e)
		{
			this.btnCancelar_Click(sender, e);

			using (frmBuscaTransferenciaActivos Busqueda = new frmBuscaTransferenciaActivos(iBodegaPredef))
			{
				if (Busqueda.ShowDialog() == DialogResult.OK) 
					Consultar((int)Busqueda.ultraGrid1.ActiveRow.Cells["idCompraTransf"].Value);
			}
		}

		private void Consultar(int idTransferencia)
		{
			try
			{
				if (miAcceso.Nuevo) this.btnNuevo.Enabled = true;
				if (miAcceso.Buscar) this.btnConsultar.Enabled = true;
				this.btnEditar.Enabled = false;
				this.btnImprimir.Enabled = false;
				this.btnGuardar.Enabled = false;
				this.btnAnular.Enabled = false;
				this.btnCancelar.Enabled = false;

				DateTime dtFechaT = DateTime.Today;

				int iBodegaOrigen = 0;
				string sSQL = string.Format("Exec ConsultaIndividualTransferencias {0}", idTransferencia);
				SqlDataReader drTransferencia = Funcion.rEscalarSQL(cdsSeteoF, sSQL, true);
				drTransferencia.Read();
				if (drTransferencia.HasRows)
				{					
					this.txtidCompraTransf.Value = idTransferencia;
					this.txtNumero.Text = drTransferencia.GetValue(1).ToString();
					dtFechaT = Convert.ToDateTime(drTransferencia.GetValue(2));
					iBodegaOrigen = (int)drTransferencia.GetValue(3);
					this.cmbBodegaDestino.Value = (int)drTransferencia.GetValue(4);
					this.lblEstado.Text = drTransferencia.GetValue(5).ToString();
					this.txtEstado.Value = (int)drTransferencia.GetValue(6);
					this.txtNotas.Text = drTransferencia.GetValue(7).ToString();
						this.cmbRealiza.Value = (int)drTransferencia.GetValue(13);
					this.cmbSolicita.Value = (int)drTransferencia.GetValue(14);
					this.lblUsuario.Text = drTransferencia.GetValue(15).ToString();
				}
				drTransferencia.Close();
			
			this.ultraGrid1.DataSource = Funcion.dvProcedimiento(cdsSeteoF, string.Format("Exec [ConsultaMovimientosActivosGrabaDetalle] {0}", (int)this.txtidCompraTransf.Value));						
				
				this.dtFecha.Value = dtFechaT.Date;

				if (this.dtFecha.Value != System.DBNull.Value)
				{
					if (!Funcion.ValidaPeriodo((DateTime)this.dtFecha.Value, cdsSeteoF, "Bodega"))
					{
						if ((int)this.txtEstado.Value == 0) if (miAcceso.Editar) this.btnEditar.Enabled = true;
						if (miAcceso.Anular) 
						{
							if ((int)this.txtEstado.Value <= 1) this.btnAnular.Enabled = true;
							if ((int)this.txtEstado.Value == 3 && miAcceso.BAnularEnTransito) this.btnAnular.Enabled = true;
							if ((int)this.txtEstado.Value == 4 && miAcceso.BAnularTFProcesado) this.btnAnular.Enabled = true;
							this.btnImprimir.Enabled = true;
						}
					}
				}

				if (miAcceso.Imprimir && (int)this.txtEstado.Value < 5) this.btnImprimir.Enabled = true;

				this.btnCancelar.Enabled = true;
			}
			catch(Exception Exc)
			{
				MessageBox.Show(string.Format("Error al Consultar: {0}", Exc), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
			}
		}

		static public bool bValidaRepetidosTF(Infragistics.Win.UltraWinGrid.UltraGrid uGrid)
		{
			int iCont = 0;
			string sCodArticulo = "";
	
			#region Valida repetidos
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in uGrid.Rows.All)
			{					
				int idArticulo = (int)dr.Cells["idActivo"].Value;
				sCodArticulo = dr.Cells["Serie"].Value.ToString();

				iCont = 0;

				foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr1 in uGrid.Rows.All)
					if (idArticulo == (int)dr1.Cells["idActivo"].Value) iCont++;
			
				if (iCont > 1) break;
			}	
			#endregion Valida repetidos
						
			if (iCont > 1)
			{
				MessageBox.Show(string.Format("La serie '{0}' ha sido registrado {1} veces\n\nno puede repetir las serie ", sCodArticulo, iCont), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
				return false;
			}		

			return true;
		}

		private void btnGuardar_Click(object sender, System.EventArgs e)
		{
			bEliminaAlValidar = true;
			Validacion.vEliminaFilasVacias(this.ultraGrid1, "idActivo");
			bEliminaAlValidar = false;

			#region Validación
			if (!Validacion.vbFechaEnDocumentos(this.dtFecha, "Ingrese la Fecha del Documento", true, "BODEGA", cdsSeteoF)) return;
			if (!Validacion.vbComboVacio(this.cmbBodegaDestino, "Seleccione la Bodega Destino")) return;
			if (!Validacion.vbComboVacio(this.cmbSolicita, "Seleccione el nombre de la persona responsable de recibir la mercaderia")) return;				
			
			if (!Validacion.vNotas(this.txtNotas)) return;			
//			if (!vbGrids(this.ultraGrid1, "Codigo", 42)) return;
		 	if (!bValidaRepetidosTF(this.ultraGrid1)) return;
			#endregion Validación

			
			
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.ultraGrid1.Rows.All)
			{	
					

				if (dr.Cells["idEstadoEquipo"].Value == System.DBNull.Value)
				{
					MessageBox.Show("Seleccione el estado del equipo", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					this.ultraGrid1.ActiveRow = dr;
					this.ultraGrid1.ActiveRow.Selected = true;
					return;
				}		
				
				if ((int)dr.Cells["idEstadoEquipo"].Value == 0)
				{
					MessageBox.Show("Seleccione el estado del equipo", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					this.ultraGrid1.ActiveRow = dr;
					this.ultraGrid1.ActiveRow.Selected = true;
					return;
				}	

				if ((int)dr.Cells["idPersonal"].Value == (int) this.cmbSolicita.Value)
				{
					MessageBox.Show("No puede realizar la Entrega del Articulo al mismo Personal de Origen", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					this.ultraGrid1.ActiveRow = dr;
					this.ultraGrid1.ActiveRow.Selected = true;
					return;
				}	
			}

	


			if (DialogResult.Yes == MessageBox.Show("¿Desea generar la Entrega Del Equipo?, Si presiona SI ya no podra modificar o registrar mas Items.", "Point Technology", MessageBoxButtons.YesNo, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2))
				this.txtEstado.Value = 4;
		
			using (SqlConnection oConexion = new SqlConnection("Password=Bl45o6$9;Persist Security Info=True;User ID= " + MenuLatinium.stUsuario + ";Initial Catalog=" + MenuLatinium.stDirFacturacion + ";Data Source=" + MenuLatinium.stDirServidor + ";Max Pool Size = 1024;"))
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

					if (bNuevo)
					{
						string sSQLNumero = "Exec NumeracionLocales 70, 0, 0";
						oCmdActualiza.CommandText = sSQLNumero;
						this.txtNumero.Text = oCmdActualiza.ExecuteScalar().ToString();
					}	

					string sSQLGrabaMaestro = string.Format("Exec TransferenciasGrabaMaestro {0}, '{1}', '{2}', {3}, {4}, {5}, '{6}', '{7}', {8}, {9}, '{10}', '{11}', {12}, {13}, {14}", 
						(int)this.txtidCompraTransf.Value, this.txtNumero.Text.ToString(), Convert.ToDateTime(this.dtFecha.Value).ToString("yyyyMMdd"), 
						(int)this.cmbBodegaDestino.Value, (int)this.cmbBodegaDestino.Value, (int)this.txtEstado.Value, 
						this.txtNotas.Text.ToString(), "", 0, 0, 
						"", "", (int)this.cmbRealiza.Value, (int)this.cmbSolicita.Value, 2);

					oCmdActualiza.CommandText = sSQLGrabaMaestro;
					this.txtidCompraTransf.Value = (int)oCmdActualiza.ExecuteScalar();

					foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in ultraGrid1.Rows.All)
					{
						string sSQLGrabaDetalle = string.Format("Exec MovimientosActivosGrabaDetalle {0}, {1}, {2}, {3}, {4}, {5}, {6}, {7}, '{8}', {9}, {10}, '{11}'",
						  (int)dr.Cells["idMovimientosActivos"].Value ,	(int)dr.Cells["idActivo"].Value ,(int)dr.Cells["idTipo"].Value ,	
							(int)dr.Cells["idPersonal"].Value ,(int)dr.Cells["idProyecto"].Value , (int)dr.Cells["idPersonalRecibe"].Value ,
							(int)dr.Cells["idProyectoRecibe"].Value ,	(int)this.txtEstado.Value	, Convert.ToDateTime(dr.Cells["Fecha"].Value).ToString("yyyyMMdd")	, 
							 (int)this.txtidCompraTransf.Value ,
	             (int)dr.Cells["idEstadoEquipo"].Value  , (string) this.txtNumero.Text);
						
						oCmdActualiza.CommandText = sSQLGrabaDetalle;
						oCmdActualiza.ExecuteNonQuery();
					}

					string sSQLActDetalle = string.Format("Exec ConsultaMovimientosActivosGrabaDetalle {0}", (int)this.txtidCompraTransf.Value); 
					oCmdActualiza.CommandText = sSQLActDetalle;
					SqlDataAdapter oDADetalle = new SqlDataAdapter(oCmdActualiza);
					DataTable oDTDetalle = new DataTable();
					oDADetalle.Fill(oDTDetalle);
					this.ultraGrid1.DataSource = oDTDetalle;

					if (!bNuevo)
					{
						string sSQLBloqueo = string.Format("Exec RegistraBloqueoTransacciones {0}, 0, {1}, 3", idBloqueaTransacciones, (int)this.txtidCompraTransf.Value);
						oCmdActualiza.CommandText = sSQLBloqueo;
						oCmdActualiza.ExecuteNonQuery();
					}

					oTransaction.Commit();
					
					oConexion.Close();

					this.dtFecha.Enabled = false;
//					this.cmbBodegaOrigen.Enabled = false;
					this.cmbBodegaDestino.Enabled = false;
					this.cmbRealiza.Enabled = false;
					this.cmbSolicita.Enabled = false;
//					this.chkActivar.Enabled = false;					
					this.txtNotas.Enabled = false;
					this.txtSerie.Enabled = false;
										
					FuncionesProcedimientos.EstadoGrids(false, this.ultraGrid1);
										
					bNuevo = false;
					bEdicion = false;

					this.Consultar((int)this.txtidCompraTransf.Value);
				}
				catch (Exception ex)
				{
					oTransaction.Rollback();

					MessageBox.Show(string.Format("Commit Exception Type : {0} {1}", ex.GetType(), ex.Message), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
				}
				finally
				{
					oConexion.Close();
				}
			}
		}

		private void btnImprimir_Click(object sender, System.EventArgs e)
		{
			ParameterFields paramFields = new ParameterFields();
			
							
								ParameterField paramFieldCedula = new ParameterField ();
								ParameterDiscreteValue ValCedula = new ParameterDiscreteValue ();
								paramFieldCedula.ParameterFieldName = "@idCompraTransf";
								ValCedula.Value = (int) this.txtidCompraTransf.Value;
								paramFieldCedula.CurrentValues.Add (ValCedula);
								paramFields.Add (paramFieldCedula);
			
								Reporte miRep = new Reporte("ReporteActivosFijos.rpt", "");
								miRep.MdiParent = this.MdiParent;
								miRep.Titulo("Activos Fijos");
								miRep.crVista.ParameterFieldInfo = paramFields;
								miRep.Show();
		}

		private void btnSalir_Click(object sender, System.EventArgs e)
		{
			this.Close ();
		}
	}
}
