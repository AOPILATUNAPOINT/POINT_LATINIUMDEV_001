using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using DevExpress.XtraEditors;
using CrystalDecisions.Shared;
using Infragistics.Win.UltraWinGrid;
using System.Data;
using System.Data.SqlClient;
using System.Globalization;

namespace Latinium
{
	/// <summary>
	/// Summary description for frmTransferencias.
	/// </summary>
	public class frmTransferencias : DevExpress.XtraEditors.XtraForm
	{
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource1;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource2;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkActivar;
		private System.Windows.Forms.Label label8;
		private System.Windows.Forms.Label label10;
		private Infragistics.Win.UltraWinGrid.UltraGrid ultraGrid1;
		private System.Windows.Forms.Label label2;
		private System.Windows.Forms.Label label1;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtFecha;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtEstado;
		private System.Windows.Forms.Button btnNuevo;
		private System.Windows.Forms.Button btnConsultar;
		private System.Windows.Forms.Button btnEditar;
		private System.Windows.Forms.Button btnGuardar;
		private System.Windows.Forms.Button btnAnular;
		private System.Windows.Forms.Button btnCancelar;
		private System.Windows.Forms.Button btnSalir;
		private System.Windows.Forms.GroupBox groupBox1;
		private System.Windows.Forms.Button btnImprimir;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbBodegaDestino;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtNumero;
		private System.Windows.Forms.Label lblEstado;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource3;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtidCompraTransf;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource4;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource5;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbBodegaOrigen;
		private C1.Data.C1DataSet cdsSeteoF;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource6;
		private System.Windows.Forms.Label lblVendedor;
		private System.Windows.Forms.Label label6;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbSolicita;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbRealiza;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtNotas;
		private System.Windows.Forms.Label label3;
		private System.Windows.Forms.Label lblUsuario;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource7;
		/// <summary>
		/// Required designer variable.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public frmTransferencias()
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
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idTransporte");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("NombreTransporte");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn3 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idTipoTransporte");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn4 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("TipoTransporte");
			Infragistics.Win.Appearance appearance1 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bodega");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn5 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn6 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Bodega");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn7 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Codigo");
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idDetCompraTransf");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCompraTransf");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idArticulo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Articulo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cantidad");
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Borrar");
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn8 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idDetCompraTransf");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn9 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCompraTransf");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn10 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idArticulo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn11 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Codigo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn12 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Articulo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn13 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Cantidad");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn14 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Borrar");
			Infragistics.Win.Appearance appearance11 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton1 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmTransferencias));
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn15 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn16 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Bodega");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn17 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Codigo");
			Infragistics.Win.Appearance appearance12 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand3 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn11 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn12 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bodega");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn13 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn18 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idOrdenProd");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn19 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Numero");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn20 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Codigo");
			Infragistics.Win.Appearance appearance13 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand4 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn14 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPersonal");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn15 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn16 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.Appearance appearance14 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand5 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn17 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPersonal");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn18 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn19 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn21 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idPersonal");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn22 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Codigo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn23 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.Appearance appearance15 = new Infragistics.Win.Appearance();
			this.ultraDataSource1 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.ultraDataSource2 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.chkActivar = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.label8 = new System.Windows.Forms.Label();
			this.label10 = new System.Windows.Forms.Label();
			this.cmbBodegaDestino = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.ultraDataSource5 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.txtNumero = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.ultraGrid1 = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource3 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.lblEstado = new System.Windows.Forms.Label();
			this.label2 = new System.Windows.Forms.Label();
			this.label1 = new System.Windows.Forms.Label();
			this.dtFecha = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.txtEstado = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.btnNuevo = new System.Windows.Forms.Button();
			this.btnConsultar = new System.Windows.Forms.Button();
			this.btnEditar = new System.Windows.Forms.Button();
			this.btnGuardar = new System.Windows.Forms.Button();
			this.btnAnular = new System.Windows.Forms.Button();
			this.btnCancelar = new System.Windows.Forms.Button();
			this.btnSalir = new System.Windows.Forms.Button();
			this.groupBox1 = new System.Windows.Forms.GroupBox();
			this.btnImprimir = new System.Windows.Forms.Button();
			this.txtidCompraTransf = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.ultraDataSource4 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.cmbBodegaOrigen = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.ultraDataSource6 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.lblVendedor = new System.Windows.Forms.Label();
			this.label6 = new System.Windows.Forms.Label();
			this.cmbSolicita = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.cmbRealiza = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.ultraDataSource7 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.txtNotas = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label3 = new System.Windows.Forms.Label();
			this.lblUsuario = new System.Windows.Forms.Label();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbBodegaDestino)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource5)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNumero)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource3)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFecha)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtEstado)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtidCompraTransf)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource4)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbBodegaOrigen)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource6)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbSolicita)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbRealiza)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource7)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNotas)).BeginInit();
			this.SuspendLayout();
			// 
			// ultraDataSource1
			// 
			ultraDataColumn1.DataType = typeof(int);
			this.ultraDataSource1.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn1,
																																 ultraDataColumn2});
			// 
			// ultraDataSource2
			// 
			ultraDataColumn3.DataType = typeof(int);
			this.ultraDataSource2.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn3,
																																 ultraDataColumn4});
			// 
			// chkActivar
			// 
			appearance1.ForeColorDisabled = System.Drawing.Color.Black;
			this.chkActivar.Appearance = appearance1;
			this.chkActivar.BackColor = System.Drawing.Color.Transparent;
			this.chkActivar.Enabled = false;
			this.chkActivar.Location = new System.Drawing.Point(288, 8);
			this.chkActivar.Name = "chkActivar";
			this.chkActivar.Size = new System.Drawing.Size(64, 22);
			this.chkActivar.TabIndex = 4;
			this.chkActivar.Text = "Activar";
			this.chkActivar.KeyDown += new System.Windows.Forms.KeyEventHandler(this.chkActivar_KeyDown);
			this.chkActivar.CheckedChanged += new System.EventHandler(this.chkActivar_CheckedChanged);
			// 
			// label8
			// 
			this.label8.AutoSize = true;
			this.label8.BackColor = System.Drawing.Color.Transparent;
			this.label8.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label8.Location = new System.Drawing.Point(8, 74);
			this.label8.Name = "label8";
			this.label8.Size = new System.Drawing.Size(83, 17);
			this.label8.TabIndex = 168;
			this.label8.Text = "Bodega Destino";
			this.label8.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label10
			// 
			this.label10.AutoSize = true;
			this.label10.BackColor = System.Drawing.Color.Transparent;
			this.label10.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label10.Location = new System.Drawing.Point(8, 43);
			this.label10.Name = "label10";
			this.label10.Size = new System.Drawing.Size(78, 17);
			this.label10.TabIndex = 167;
			this.label10.Text = "Bodega Origen";
			this.label10.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// cmbBodegaDestino
			// 
			appearance2.ForeColor = System.Drawing.Color.Black;
			appearance2.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbBodegaDestino.Appearance = appearance2;
			this.cmbBodegaDestino.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbBodegaDestino.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbBodegaDestino.DataSource = this.ultraDataSource5;
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn1.Width = 250;
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridColumn2.Hidden = true;
			ultraGridColumn2.Width = 50;
			ultraGridColumn3.Header.VisiblePosition = 2;
			ultraGridColumn3.Hidden = true;
			ultraGridColumn3.Width = 50;
			ultraGridBand1.Columns.AddRange(new object[] {
																										 ultraGridColumn1,
																										 ultraGridColumn2,
																										 ultraGridColumn3});
			this.cmbBodegaDestino.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.cmbBodegaDestino.DisplayMember = "Nombre";
			this.cmbBodegaDestino.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbBodegaDestino.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbBodegaDestino.Enabled = false;
			this.cmbBodegaDestino.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbBodegaDestino.Location = new System.Drawing.Point(104, 72);
			this.cmbBodegaDestino.MaxDropDownItems = 30;
			this.cmbBodegaDestino.Name = "cmbBodegaDestino";
			this.cmbBodegaDestino.Size = new System.Drawing.Size(250, 21);
			this.cmbBodegaDestino.TabIndex = 3;
			this.cmbBodegaDestino.ValueMember = "Bodega";
			this.cmbBodegaDestino.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cmbBodegaDestino_KeyDown);
			this.cmbBodegaDestino.Validating += new System.ComponentModel.CancelEventHandler(this.cmbBodegaDestino_Validating);
			this.cmbBodegaDestino.ValueChanged += new System.EventHandler(this.cmbBodegaDestino_ValueChanged);
			this.cmbBodegaDestino.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.cmbBodegaDestino_InitializeLayout);
			// 
			// ultraDataSource5
			// 
			this.ultraDataSource5.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn5,
																																 ultraDataColumn6,
																																 ultraDataColumn7});
			// 
			// txtNumero
			// 
			appearance3.ForeColor = System.Drawing.Color.Black;
			appearance3.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtNumero.Appearance = appearance3;
			this.txtNumero.Enabled = false;
			this.txtNumero.Location = new System.Drawing.Point(584, 8);
			this.txtNumero.Name = "txtNumero";
			this.txtNumero.Size = new System.Drawing.Size(112, 22);
			this.txtNumero.TabIndex = 0;
			// 
			// ultraGrid1
			// 
			this.ultraGrid1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.ultraGrid1.Cursor = System.Windows.Forms.Cursors.Default;
			this.ultraGrid1.DataSource = this.ultraDataSource3;
			appearance4.BackColor = System.Drawing.Color.White;
			this.ultraGrid1.DisplayLayout.Appearance = appearance4;
			this.ultraGrid1.DisplayLayout.AutoFitColumns = true;
			ultraGridColumn4.Header.VisiblePosition = 0;
			ultraGridColumn4.Hidden = true;
			ultraGridColumn4.Width = 175;
			ultraGridColumn5.Header.VisiblePosition = 1;
			ultraGridColumn5.Hidden = true;
			ultraGridColumn5.Width = 145;
			ultraGridColumn6.Header.VisiblePosition = 2;
			ultraGridColumn6.Hidden = true;
			ultraGridColumn6.Width = 94;
			ultraGridColumn7.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn7.Header.VisiblePosition = 3;
			ultraGridColumn7.Width = 220;
			ultraGridColumn8.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn8.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn8.Header.VisiblePosition = 4;
			ultraGridColumn8.Width = 370;
			appearance5.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn9.CellAppearance = appearance5;
			ultraGridColumn9.Format = "0";
			ultraGridColumn9.Header.VisiblePosition = 5;
			ultraGridColumn9.Width = 80;
			ultraGridColumn10.Header.VisiblePosition = 6;
			ultraGridColumn10.Hidden = true;
			ultraGridColumn10.Width = 132;
			ultraGridBand2.Columns.AddRange(new object[] {
																										 ultraGridColumn4,
																										 ultraGridColumn5,
																										 ultraGridColumn6,
																										 ultraGridColumn7,
																										 ultraGridColumn8,
																										 ultraGridColumn9,
																										 ultraGridColumn10});
			this.ultraGrid1.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			this.ultraGrid1.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.TemplateOnBottom;
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
			this.ultraGrid1.Size = new System.Drawing.Size(691, 184);
			this.ultraGrid1.TabIndex = 163;
			this.ultraGrid1.AfterRowInsert += new Infragistics.Win.UltraWinGrid.RowEventHandler(this.ultraGrid1_AfterRowInsert);
			this.ultraGrid1.BeforeRowsDeleted += new Infragistics.Win.UltraWinGrid.BeforeRowsDeletedEventHandler(this.ultraGrid1_BeforeRowsDeleted);
			this.ultraGrid1.BeforeRowInsert += new Infragistics.Win.UltraWinGrid.BeforeRowInsertEventHandler(this.ultraGrid1_BeforeRowInsert);
			this.ultraGrid1.KeyDown += new System.Windows.Forms.KeyEventHandler(this.ultraGrid1_KeyDown);
			this.ultraGrid1.AfterCellUpdate += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.ultraGrid1_AfterCellUpdate);
			this.ultraGrid1.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.ultraGrid1_InitializeLayout);
			// 
			// ultraDataSource3
			// 
			ultraDataColumn8.DataType = typeof(int);
			ultraDataColumn9.DataType = typeof(int);
			ultraDataColumn10.DataType = typeof(int);
			ultraDataColumn13.DataType = typeof(int);
			ultraDataColumn13.DefaultValue = 0;
			ultraDataColumn14.DataType = typeof(bool);
			ultraDataColumn14.DefaultValue = false;
			this.ultraDataSource3.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn8,
																																 ultraDataColumn9,
																																 ultraDataColumn10,
																																 ultraDataColumn11,
																																 ultraDataColumn12,
																																 ultraDataColumn13,
																																 ultraDataColumn14});
			// 
			// lblEstado
			// 
			this.lblEstado.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.lblEstado.AutoSize = true;
			this.lblEstado.BackColor = System.Drawing.Color.Transparent;
			this.lblEstado.Font = new System.Drawing.Font("Microsoft Sans Serif", 15F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblEstado.ForeColor = System.Drawing.Color.Firebrick;
			this.lblEstado.Location = new System.Drawing.Point(520, 328);
			this.lblEstado.Name = "lblEstado";
			this.lblEstado.Size = new System.Drawing.Size(0, 26);
			this.lblEstado.TabIndex = 169;
			this.lblEstado.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label2
			// 
			this.label2.AutoSize = true;
			this.label2.BackColor = System.Drawing.Color.Transparent;
			this.label2.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label2.Location = new System.Drawing.Point(520, 11);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(44, 17);
			this.label2.TabIndex = 166;
			this.label2.Text = "Número";
			this.label2.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label1
			// 
			this.label1.AutoSize = true;
			this.label1.BackColor = System.Drawing.Color.Transparent;
			this.label1.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label1.Location = new System.Drawing.Point(8, 11);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(34, 17);
			this.label1.TabIndex = 165;
			this.label1.Text = "Fecha";
			this.label1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
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
			this.dtFecha.Location = new System.Drawing.Point(104, 9);
			this.dtFecha.Name = "dtFecha";
			this.dtFecha.NonAutoSizeHeight = 23;
			this.dtFecha.Size = new System.Drawing.Size(136, 21);
			this.dtFecha.SpinButtonsVisible = true;
			this.dtFecha.TabIndex = 1;
			this.dtFecha.Value = ((object)(resources.GetObject("dtFecha.Value")));
			this.dtFecha.KeyDown += new System.Windows.Forms.KeyEventHandler(this.dtFecha_KeyDown);
			this.dtFecha.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.dtFecha_KeyPress);
			this.dtFecha.Validating += new System.ComponentModel.CancelEventHandler(this.dtFecha_Validating);
			this.dtFecha.Validated += new System.EventHandler(this.dtFecha_Validated);
			this.dtFecha.BeforeDropDown += new System.ComponentModel.CancelEventHandler(this.dtFecha_BeforeDropDown);
			this.dtFecha.ValueChanged += new System.EventHandler(this.dtFecha_ValueChanged);
			// 
			// txtEstado
			// 
			this.txtEstado.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.txtEstado.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtEstado.Enabled = false;
			this.txtEstado.Location = new System.Drawing.Point(659, 432);
			this.txtEstado.Name = "txtEstado";
			this.txtEstado.PromptChar = ' ';
			this.txtEstado.Size = new System.Drawing.Size(16, 22);
			this.txtEstado.TabIndex = 304;
			this.txtEstado.Visible = false;
			// 
			// btnNuevo
			// 
			this.btnNuevo.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnNuevo.CausesValidation = false;
			this.btnNuevo.Image = ((System.Drawing.Image)(resources.GetObject("btnNuevo.Image")));
			this.btnNuevo.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnNuevo.Location = new System.Drawing.Point(8, 432);
			this.btnNuevo.Name = "btnNuevo";
			this.btnNuevo.Size = new System.Drawing.Size(72, 23);
			this.btnNuevo.TabIndex = 303;
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
			this.btnConsultar.Location = new System.Drawing.Point(83, 432);
			this.btnConsultar.Name = "btnConsultar";
			this.btnConsultar.Size = new System.Drawing.Size(78, 23);
			this.btnConsultar.TabIndex = 302;
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
			this.btnEditar.Location = new System.Drawing.Point(164, 432);
			this.btnEditar.Name = "btnEditar";
			this.btnEditar.Size = new System.Drawing.Size(76, 23);
			this.btnEditar.TabIndex = 301;
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
			this.btnGuardar.Location = new System.Drawing.Point(323, 432);
			this.btnGuardar.Name = "btnGuardar";
			this.btnGuardar.Size = new System.Drawing.Size(76, 23);
			this.btnGuardar.TabIndex = 300;
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
			this.btnAnular.Location = new System.Drawing.Point(402, 432);
			this.btnAnular.Name = "btnAnular";
			this.btnAnular.Size = new System.Drawing.Size(76, 23);
			this.btnAnular.TabIndex = 299;
			this.btnAnular.Text = "&Anular";
			this.btnAnular.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnAnular.Click += new System.EventHandler(this.btnAnular_Click);
			// 
			// btnCancelar
			// 
			this.btnCancelar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnCancelar.CausesValidation = false;
			this.btnCancelar.Enabled = false;
			this.btnCancelar.Image = ((System.Drawing.Image)(resources.GetObject("btnCancelar.Image")));
			this.btnCancelar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnCancelar.Location = new System.Drawing.Point(481, 432);
			this.btnCancelar.Name = "btnCancelar";
			this.btnCancelar.Size = new System.Drawing.Size(72, 23);
			this.btnCancelar.TabIndex = 298;
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
			this.btnSalir.Location = new System.Drawing.Point(556, 432);
			this.btnSalir.Name = "btnSalir";
			this.btnSalir.TabIndex = 297;
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
			this.groupBox1.Location = new System.Drawing.Point(0, 416);
			this.groupBox1.Name = "groupBox1";
			this.groupBox1.Size = new System.Drawing.Size(707, 8);
			this.groupBox1.TabIndex = 296;
			this.groupBox1.TabStop = false;
			// 
			// btnImprimir
			// 
			this.btnImprimir.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnImprimir.CausesValidation = false;
			this.btnImprimir.Enabled = false;
			this.btnImprimir.Image = ((System.Drawing.Image)(resources.GetObject("btnImprimir.Image")));
			this.btnImprimir.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnImprimir.Location = new System.Drawing.Point(243, 432);
			this.btnImprimir.Name = "btnImprimir";
			this.btnImprimir.Size = new System.Drawing.Size(77, 23);
			this.btnImprimir.TabIndex = 305;
			this.btnImprimir.Text = "&Imprimir";
			this.btnImprimir.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnImprimir.Click += new System.EventHandler(this.btnImprimir_Click);
			// 
			// txtidCompraTransf
			// 
			this.txtidCompraTransf.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.txtidCompraTransf.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtidCompraTransf.Enabled = false;
			this.txtidCompraTransf.Location = new System.Drawing.Point(683, 432);
			this.txtidCompraTransf.Name = "txtidCompraTransf";
			this.txtidCompraTransf.PromptChar = ' ';
			this.txtidCompraTransf.Size = new System.Drawing.Size(16, 22);
			this.txtidCompraTransf.TabIndex = 306;
			this.txtidCompraTransf.Visible = false;
			// 
			// ultraDataSource4
			// 
			this.ultraDataSource4.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn15,
																																 ultraDataColumn16,
																																 ultraDataColumn17});
			// 
			// cmbBodegaOrigen
			// 
			appearance12.ForeColor = System.Drawing.Color.Black;
			appearance12.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbBodegaOrigen.Appearance = appearance12;
			this.cmbBodegaOrigen.CausesValidation = false;
			this.cmbBodegaOrigen.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbBodegaOrigen.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbBodegaOrigen.DataSource = this.ultraDataSource4;
			ultraGridColumn11.Header.VisiblePosition = 1;
			ultraGridColumn11.Width = 250;
			ultraGridColumn12.Header.VisiblePosition = 2;
			ultraGridColumn12.Hidden = true;
			ultraGridColumn13.Header.VisiblePosition = 0;
			ultraGridColumn13.Hidden = true;
			ultraGridBand3.Columns.AddRange(new object[] {
																										 ultraGridColumn11,
																										 ultraGridColumn12,
																										 ultraGridColumn13});
			this.cmbBodegaOrigen.DisplayLayout.BandsSerializer.Add(ultraGridBand3);
			this.cmbBodegaOrigen.DisplayMember = "Nombre";
			this.cmbBodegaOrigen.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbBodegaOrigen.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbBodegaOrigen.Enabled = false;
			this.cmbBodegaOrigen.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbBodegaOrigen.Location = new System.Drawing.Point(104, 41);
			this.cmbBodegaOrigen.MaxDropDownItems = 30;
			this.cmbBodegaOrigen.Name = "cmbBodegaOrigen";
			this.cmbBodegaOrigen.Size = new System.Drawing.Size(250, 21);
			this.cmbBodegaOrigen.TabIndex = 2;
			this.cmbBodegaOrigen.ValueMember = "Bodega";
			this.cmbBodegaOrigen.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cmbBodegaOrigen_KeyDown);
			this.cmbBodegaOrigen.Validating += new System.ComponentModel.CancelEventHandler(this.cmbBodegaOrigen_Validating);
			this.cmbBodegaOrigen.ValueChanged += new System.EventHandler(this.cmbBodegaOrigen_ValueChanged);
			this.cmbBodegaOrigen.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.cmbBodegaOrigen_InitializeLayout);
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
			// ultraDataSource6
			// 
			ultraDataColumn18.DataType = typeof(int);
			this.ultraDataSource6.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn18,
																																 ultraDataColumn19,
																																 ultraDataColumn20});
			// 
			// lblVendedor
			// 
			this.lblVendedor.AutoSize = true;
			this.lblVendedor.Location = new System.Drawing.Point(376, 43);
			this.lblVendedor.Name = "lblVendedor";
			this.lblVendedor.Size = new System.Drawing.Size(39, 16);
			this.lblVendedor.TabIndex = 308;
			this.lblVendedor.Text = "Realiza";
			// 
			// label6
			// 
			this.label6.AutoSize = true;
			this.label6.Location = new System.Drawing.Point(376, 74);
			this.label6.Name = "label6";
			this.label6.Size = new System.Drawing.Size(39, 16);
			this.label6.TabIndex = 310;
			this.label6.Text = "Solicita";
			// 
			// cmbSolicita
			// 
			appearance13.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbSolicita.Appearance = appearance13;
			this.cmbSolicita.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbSolicita.Cursor = System.Windows.Forms.Cursors.Default;
			ultraGridColumn14.Header.VisiblePosition = 0;
			ultraGridColumn14.Hidden = true;
			ultraGridColumn15.Header.VisiblePosition = 1;
			ultraGridColumn15.Width = 335;
			ultraGridColumn16.Header.VisiblePosition = 2;
			ultraGridBand4.Columns.AddRange(new object[] {
																										 ultraGridColumn14,
																										 ultraGridColumn15,
																										 ultraGridColumn16});
			this.cmbSolicita.DisplayLayout.BandsSerializer.Add(ultraGridBand4);
			this.cmbSolicita.DisplayMember = "Nombre";
			this.cmbSolicita.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbSolicita.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbSolicita.Enabled = false;
			this.cmbSolicita.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbSolicita.Location = new System.Drawing.Point(440, 72);
			this.cmbSolicita.Name = "cmbSolicita";
			this.cmbSolicita.Size = new System.Drawing.Size(256, 21);
			this.cmbSolicita.TabIndex = 309;
			this.cmbSolicita.ValueMember = "idPersonal";
			this.cmbSolicita.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cmbSolicita_KeyDown);
			this.cmbSolicita.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.cmbSolicita_InitializeLayout);
			// 
			// cmbRealiza
			// 
			appearance14.ForeColor = System.Drawing.Color.Black;
			appearance14.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbRealiza.Appearance = appearance14;
			this.cmbRealiza.CausesValidation = false;
			this.cmbRealiza.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbRealiza.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbRealiza.DataSource = this.ultraDataSource7;
			ultraGridColumn17.Header.VisiblePosition = 0;
			ultraGridColumn17.Hidden = true;
			ultraGridColumn18.Header.VisiblePosition = 1;
			ultraGridColumn19.Header.VisiblePosition = 2;
			ultraGridColumn19.Width = 300;
			ultraGridBand5.Columns.AddRange(new object[] {
																										 ultraGridColumn17,
																										 ultraGridColumn18,
																										 ultraGridColumn19});
			this.cmbRealiza.DisplayLayout.BandsSerializer.Add(ultraGridBand5);
			this.cmbRealiza.DisplayMember = "Nombre";
			this.cmbRealiza.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbRealiza.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbRealiza.Enabled = false;
			this.cmbRealiza.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbRealiza.Location = new System.Drawing.Point(440, 41);
			this.cmbRealiza.MaxDropDownItems = 30;
			this.cmbRealiza.Name = "cmbRealiza";
			this.cmbRealiza.Size = new System.Drawing.Size(256, 21);
			this.cmbRealiza.TabIndex = 313;
			this.cmbRealiza.ValueMember = "idPersonal";
			// 
			// ultraDataSource7
			// 
			ultraDataColumn21.DataType = typeof(int);
			this.ultraDataSource7.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn21,
																																 ultraDataColumn22,
																																 ultraDataColumn23});
			// 
			// txtNotas
			// 
			this.txtNotas.AcceptsReturn = true;
			this.txtNotas.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			appearance15.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtNotas.Appearance = appearance15;
			this.txtNotas.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtNotas.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNotas.Enabled = false;
			this.txtNotas.Location = new System.Drawing.Point(8, 320);
			this.txtNotas.Multiline = true;
			this.txtNotas.Name = "txtNotas";
			this.txtNotas.Size = new System.Drawing.Size(497, 88);
			this.txtNotas.TabIndex = 314;
			// 
			// label3
			// 
			this.label3.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.label3.AutoSize = true;
			this.label3.BackColor = System.Drawing.Color.Transparent;
			this.label3.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label3.Location = new System.Drawing.Point(8, 296);
			this.label3.Name = "label3";
			this.label3.Size = new System.Drawing.Size(33, 17);
			this.label3.TabIndex = 315;
			this.label3.Text = "Notas";
			this.label3.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// lblUsuario
			// 
			this.lblUsuario.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.lblUsuario.AutoSize = true;
			this.lblUsuario.BackColor = System.Drawing.Color.Transparent;
			this.lblUsuario.Font = new System.Drawing.Font("Tahoma", 10F, System.Drawing.FontStyle.Bold);
			this.lblUsuario.Location = new System.Drawing.Point(520, 376);
			this.lblUsuario.Name = "lblUsuario";
			this.lblUsuario.Size = new System.Drawing.Size(0, 20);
			this.lblUsuario.TabIndex = 676;
			this.lblUsuario.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// frmTransferencias
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 14);
			this.CausesValidation = false;
			this.ClientSize = new System.Drawing.Size(707, 462);
			this.Controls.Add(this.lblUsuario);
			this.Controls.Add(this.label3);
			this.Controls.Add(this.label6);
			this.Controls.Add(this.lblVendedor);
			this.Controls.Add(this.label8);
			this.Controls.Add(this.label10);
			this.Controls.Add(this.label2);
			this.Controls.Add(this.label1);
			this.Controls.Add(this.txtNotas);
			this.Controls.Add(this.cmbRealiza);
			this.Controls.Add(this.cmbSolicita);
			this.Controls.Add(this.cmbBodegaOrigen);
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
			this.Controls.Add(this.chkActivar);
			this.KeyPreview = true;
			this.Name = "frmTransferencias";
			this.Text = "Transferencias";
			this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.frmTransferencias_KeyDown);
			this.Closing += new System.ComponentModel.CancelEventHandler(this.frmTransferencias_Closing);
			this.Load += new System.EventHandler(this.frmTransferencias_Load);
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbBodegaDestino)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource5)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNumero)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource3)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFecha)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtEstado)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtidCompraTransf)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource4)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbBodegaOrigen)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource6)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbSolicita)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbRealiza)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource7)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNotas)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		bool bNuevo = false;
		bool bEdicion = false;
		private Acceso miAcceso;
		bool bEliminaAlValidar = false;		
		int idBloqueaTransacciones = 0;
		int iBodegaPredef = 0;
		private bool bActualiza = true;
		bool bCombo = false;

		private void frmTransferencias_Load(object sender, System.EventArgs e)
		{							
			miAcceso = new Acceso(cdsSeteoF, "0315");
			
			string sSQLBP = string.Format("Exec BodegaAsignadaPorUsuario {0}", MenuLatinium.IdUsuario);		
			iBodegaPredef = Funcion.iEscalarSQL(cdsSeteoF, sSQLBP);

			FuncionesProcedimientos.EstadoGrids(false, this.ultraGrid1);
		}

		private void btnSalir_Click(object sender, System.EventArgs e)
		{
			this.Close();
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
			this.dtFecha.Value = System.DBNull.Value;
			this.dtFecha.CalendarInfo.MaxDate = DateTime.Today;

			this.cmbBodegaOrigen.Value = System.DBNull.Value;
			this.cmbRealiza.Value = System.DBNull.Value;
			this.cmbBodegaDestino.Value = System.DBNull.Value;
			this.cmbSolicita.Value = System.DBNull.Value;
			this.chkActivar.Checked = false;
			this.lblEstado.Text = "";
			this.lblUsuario.Text = "";
			this.txtNotas.Text = "";						
			
			this.ultraGrid1.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Exec TrasnferenciasConsulta 0");

			this.dtFecha.Enabled = false;
			this.cmbBodegaOrigen.Enabled = false;
			this.cmbBodegaDestino.Enabled = false;
			this.cmbRealiza.Enabled = false;
			this.cmbSolicita.Enabled = false;
			this.chkActivar.Enabled = false;			
			this.txtNotas.Enabled = false;
			
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

		private void btnNuevo_Click(object sender, System.EventArgs e)
		{
			this.btnCancelar_Click(sender, e);						
					
			if (miAcceso.CambiarFecha) this.dtFecha.Enabled = true;
						
			FuncionesProcedimientos.RetornaFechaServidor(this.dtFecha, cdsSeteoF, false);
									
			this.cmbBodegaDestino.Enabled = true;			
			FuncionesProcedimientos.EstadoGrids(true, this.ultraGrid1);
			this.txtNotas.Enabled = true;
						
			bNuevo = true;
			bEdicion = true;

			if (iBodegaPredef == 0) 
			{
				this.cmbBodegaOrigen.Enabled = true;
				this.chkActivar.Enabled = true;
			}
			else 
			{
				this.cmbBodegaOrigen.Value = iBodegaPredef;
				this.chkActivar.Checked = true;
				this.txtEstado.Value = 1;
				this.lblEstado.Text = "PENDIENTE";
			}

			if (iBodegaPredef > 0)
			{
				if (Funcion.iEscalarSQL(cdsSeteoF, string.Format("Select COUNT(*) From CierreConteoInventario Where Bodega = {0} And CONVERT(DATE, Fecha) Between '20161001' And '{1}' And Estado = 1", 
					(int)this.cmbBodegaOrigen.Value, Convert.ToDateTime(this.dtFecha.Value).ToString("yyyyMMdd"))) > 0)
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
			else if (this.cmbBodegaOrigen.Enabled) this.cmbBodegaOrigen.Focus(); else this.cmbRealiza.Focus();
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
								
				this.cmbBodegaOrigen.Value = iBodegaOrigen;

				if ((int)this.txtEstado.Value != 0 && (int)this.txtEstado.Value != 5) this.chkActivar.Checked = true;

				this.ultraGrid1.DataSource = Funcion.dvProcedimiento(cdsSeteoF, string.Format("Exec TrasnferenciasConsulta {0}", (int)this.txtidCompraTransf.Value));						
				
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

		private void btnConsultar_Click(object sender, System.EventArgs e)
		{
			this.btnCancelar_Click(sender, e);

			using (frmBuscaTransferencias Busqueda = new frmBuscaTransferencias(iBodegaPredef))
			{
				if (Busqueda.ShowDialog() == DialogResult.OK) 
					Consultar((int)Busqueda.ultraGrid1.ActiveRow.Cells["idCompraTransf"].Value);
			}
		}

		private void btnEditar_Click(object sender, System.EventArgs e)
		{				
			this.Consultar((int)this.txtidCompraTransf.Value);

			if ((int)this.txtEstado.Value == 5)
			{
				MessageBox.Show("La Transferencia esta Anulada", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);				
			}
			else if (Funcion.ValidaPeriodo((DateTime)this.dtFecha.Value, cdsSeteoF, "Bodega"))
			{
				DateTime dtFechaVP = (DateTime)this.dtFecha.Value;
				MessageBox.Show(string.Format("No puede Editar esta Transferencia, El Período '{0}' esta Cerrado,\n\n Consulte al Administrador", dtFechaVP.ToString("MMMM").ToUpper()), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
			}			
			else if ((int)this.txtEstado.Value == 1)
			{
				MessageBox.Show("Solo puede Editar Transferencias En Espera", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);				
			}
			else if ((int)this.txtEstado.Value == 3)
			{
				MessageBox.Show("No puede Editar esta Transferencia ya esta En Tránsito", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);				
			}
			else if ((int)this.txtEstado.Value == 4)
			{
				MessageBox.Show("No puede Editar esta Transferencia ya esta Procesada", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);				
			}
			else if (!Funcion.ValidaBloqueo((int) this.txtidCompraTransf.Value, 3, cdsSeteoF))
			{
				string sSQLValSeriales = string.Format("Exec ValidaSerialesIngresadasParaAnular {0}, 42", (int)this.txtidCompraTransf.Value);
				if ((int)this.txtEstado.Value > 0 && Funcion.iEscalarSQL(cdsSeteoF, sSQLValSeriales, true) > 0)
				{
					MessageBox.Show("No puede Editar Existen Seriales Ingresadas, \n\n Consulte al Administrador", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);					
				}
				else
				{
					if (miAcceso.CambiarFecha) this.dtFecha.Enabled = true;

					if ((int)this.txtEstado.Value == 0 || (int)this.txtEstado.Value == 1) this.chkActivar.Enabled = true;

					if (iBodegaPredef > 0 && Funcion.iEscalarSQL(cdsSeteoF, string.Format("Select idCom_Empresa From Com_Empresa Where Bodega = {0}", (int)this.cmbBodegaOrigen.Value)) > 0) this.cmbRealiza.Enabled = true;
					this.cmbSolicita.Enabled = true;

					FuncionesProcedimientos.EstadoGrids(true, this.ultraGrid1);

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
		}

		private void btnImprimir_Click(object sender, System.EventArgs e)
		{
			int idDBEmpresa = MenuLatinium.stIdDB;
			ParameterFields paramFields = new ParameterFields ();

			ParameterField paramFieldTransf = new ParameterField();
			ParameterDiscreteValue discreteValTransf = new ParameterDiscreteValue();
			paramFieldTransf.ParameterFieldName = "@idComprTransf";
			discreteValTransf.Value = int.Parse(this.txtidCompraTransf.Value.ToString());
			paramFieldTransf.CurrentValues.Add (discreteValTransf);
			paramFields.Add (paramFieldTransf);
			
			ParameterField paramFieldEmpresa = new ParameterField ();
			ParameterDiscreteValue discreteValEmpresa = new ParameterDiscreteValue ();
			paramFieldEmpresa.ParameterFieldName = "@idEmpresa";
			discreteValEmpresa.Value = idDBEmpresa; 
			paramFieldEmpresa.CurrentValues.Add (discreteValEmpresa);
			paramFields.Add (paramFieldEmpresa);

			Reporte miRep = new Reporte("Transfer2ConImg.rpt", "");
			miRep.MdiParent = this.MdiParent;
			miRep.Titulo("Transferencia");
			miRep.crVista.ParameterFieldInfo = paramFields;
			miRep.Show();
		}

		private void btnAnular_Click(object sender, System.EventArgs e)
		{	
			this.Consultar((int)this.txtidCompraTransf.Value);

			if ((int)this.txtEstado.Value == 5)
			{
				MessageBox.Show("La Transferencia esta Anulada", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return;
			}
			
			if (!miAcceso.Anular)
			{
				MessageBox.Show("No tiene Acceso a Anular", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);			
				return;
			}
			
			if (Funcion.ValidaPeriodo((DateTime)this.dtFecha.Value, cdsSeteoF, "Bodega"))
			{
				DateTime dtFechaVP = (DateTime)this.dtFecha.Value;
				MessageBox.Show(string.Format("No puede Anular esta Transferencia, El Período '{0}' esta Cerrado,\n\n Consulte al Administrador", dtFechaVP.ToString("MMMM").ToUpper()), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);				
				return;
			}
			
			if ((int)this.txtEstado.Value >= 3) 
				if (!FuncionesInventario.ValidaBodegaInventarios(this.ultraGrid1, (int)this.cmbBodegaOrigen.Value, (DateTime)this.dtFecha.Value, cdsSeteoF)) return;
			

			if ((int)this.txtEstado.Value >= 3)
				if (!FuncionesInventario.ValidaBodegaInventarios(this.ultraGrid1, (int)this.cmbBodegaDestino.Value, (DateTime)this.dtFecha.Value, cdsSeteoF)) return;

			if ((int)this.txtEstado.Value >= 3)
				if (!FuncionesInventario.ValidaBodegaEnInventario((int)this.cmbBodegaOrigen.Value, (DateTime)this.dtFecha.Value, cdsSeteoF)) return;

			if ((int)this.txtEstado.Value == 4)
				if (FuncionesInventario.ValidaBodegaEnInventario((int)this.cmbBodegaDestino.Value, (DateTime)this.dtFecha.Value, cdsSeteoF)) return;

			if ((int)this.txtEstado.Value == 3 && !miAcceso.BAnularEnTransito)
			{
				MessageBox.Show("La Transferencia esta En Transito, \n\nPara Anular Consulte al Administrador", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);			
				return;
			}
			
			if ((int)this.txtEstado.Value == 4 && !miAcceso.BAnularTFProcesado)
			{
				MessageBox.Show("La Transferencia esta Procesada, \n\nPara Anular Consulte al Administrador", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);			
				return;
			}
			
			if (Funcion.ValidaBloqueo((int) this.txtidCompraTransf.Value, 3, cdsSeteoF)) return;
			
			int idCompraE = Funcion.iEscalarSQL(cdsSeteoF, string.Format("Select idCompra From CompraTransf Where idCompraTransf = {0}", (int)this.txtidCompraTransf.Value));
			if (Funcion.ValidaBloqueoERMEnTransacciones(idCompraE, cdsSeteoF)) return;

			int idCompraI = Funcion.iEscalarSQL(cdsSeteoF, string.Format("Select idCompra1 From CompraTransf Where idCompraTransf = {0}", (int)this.txtidCompraTransf.Value));
			if (Funcion.ValidaBloqueoERMEnTransacciones(idCompraI, cdsSeteoF)) return;

			string sSQLValSeriales = string.Format("Exec ValidaSerialesIngresadasParaAnular {0}, 42", (int)this.txtidCompraTransf.Value);			
			if ((int)this.txtEstado.Value > 0 && !miAcceso.BAnularEnTransito && Funcion.iEscalarSQL(cdsSeteoF, sSQLValSeriales, true) > 0)
			{
				MessageBox.Show("No puede Anular Existen Seriales Ingresadas, \n\n Consulte al Administrador", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);					
				return;
			}
			
			if (DialogResult.Yes == MessageBox.Show("¿Está seguro que desea Anular el Registro?", "Point Technology", MessageBoxButtons.YesNo, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2)) 
			{
				using (frmJustAnulaVenta Nota = new frmJustAnulaVenta())
				{
					if (DialogResult.OK == Nota.ShowDialog())
					{
						try
						{
							string stBorrar = string.Format("Exec TransferenciaBorrar {0}, 0, '{1}'", (int)this.txtidCompraTransf.Value, Nota.txtNotas.Text.ToString());
							string stMensaje = Funcion.sEscalarSQL(cdsSeteoF, stBorrar, true);

							this.Consultar((int)this.txtidCompraTransf.Value);							
						}
						catch(Exception Exc)
						{
							MessageBox.Show(string.Format("Error al anular: {0}", Exc), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
						}
					}
				}
			}
		}

		private void dtFecha_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Delete) e.Handled = true;
			if (e.KeyCode == Keys.Enter) this.cmbBodegaOrigen.Focus();
		}

		private void cmbBodegaDestino_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.cmbSolicita.Focus();
		}

		private void dtFecha_Validated(object sender, System.EventArgs e)
		{
			if (this.dtFecha.Value == System.DBNull.Value) this.dtFecha.Value = DateTime.Today;
		}

		private void ultraGrid1_BeforeRowInsert(object sender, Infragistics.Win.UltraWinGrid.BeforeRowInsertEventArgs e)
		{
			if (ultraGrid1.Rows.Count == 25) 
			{				
				MessageBox.Show("No puede ingresar mas items.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop); 
				e.Cancel = true;
			}
		}

		private void ultraGrid1_BeforeRowsDeleted(object sender, Infragistics.Win.UltraWinGrid.BeforeRowsDeletedEventArgs e)
		{
			if (bEdicion)
			{
				if (!bEliminaAlValidar)
				{
					if (DialogResult.No == MessageBox.Show("¿Esta seguro de eliminar la linea seleccionada?", "Confirmación", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2)) e.Cancel = true;
					else
					{
						if (!bNuevo)
						{
							for (int i = 0; i < e.Rows.Length; i++)
							{
								if ((int)this.txtEstado.Value > 0)
								{
									MessageBox.Show("Solo puede eliminar filas de Transferencias EN ESPERA", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
									e.Cancel = true;
									return;
								}

								if ((int)this.txtEstado.Value == 0)
								{
									string sSQLVal = string.Format("Exec ValidaEliminaFilasTransferencias {0}, {1}, {2}", 
										(int)this.txtidCompraTransf.Value, (int) e.Rows[i].Cells["idDetCompraTransf"].Value, (int) e.Rows[i].Cells["idArticulo"].Value);
									string sMensaje = Funcion.sEscalarSQL(cdsSeteoF, sSQLVal);

									if (sMensaje.Length > 0)
									{
										MessageBox.Show(sMensaje, "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
										e.Cancel = true;
										return;
									}
									else
									{
										e.Rows[i].Cells["Borrar"].Value = true;
										e.Rows[i].Appearance.BackColor = Color.Red;
										e.Rows[i].Appearance.BackColor2 = Color.Red;
//										this.btnCancelar.Enabled = false;
										e.Cancel = true;
									}
								}
							}
						}
						else e.DisplayPromptMsg = false;						
					}
				}
				else e.DisplayPromptMsg = false;
			}
		}

		private void ultraGrid1_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (ultraGrid1.ActiveCell == null) return;

			if (e.KeyCode == Keys.F3)
				if (this.ultraGrid1.ActiveCell.Column.Key.ToString().ToUpper() == "CODIGO" && bEdicion)
					FuncionesInventario.AbreBusquedaArticulos(this.ultraGrid1, (int)this.cmbBodegaOrigen.Value, 42, 0, Convert.ToDateTime(this.dtFecha.Value), false);
			
			FuncionesProcedimientos.DesplazamientoGrids(sender, e, this.ultraGrid1);
		}

		private void btnGuardar_Click(object sender, System.EventArgs e)
		{
			bEliminaAlValidar = true;
			Validacion.vEliminaFilasVacias(this.ultraGrid1, "idArticulo");
			bEliminaAlValidar = false;

			#region Validación
			if (!Validacion.vbFechaEnDocumentos(this.dtFecha, "Ingrese la Fecha del Documento", true, "BODEGA", cdsSeteoF)) return;
			if (!Validacion.vbComboVacio(this.cmbBodegaOrigen, "Seleccione la Bodega Origén")) return;
			if (!Validacion.vbComboVacio(this.cmbRealiza, "Seleccione el nombre de la persona responsable del envio de la mercaderia")) return;
			if (!Validacion.vbComboVacio(this.cmbBodegaDestino, "Seleccione la Bodega Destino")) return;
			if (!Validacion.vbComboVacio(this.cmbSolicita, "Seleccione el nombre de la persona responsable de recibir la mercaderia")) return;				
			if ((int)cmbBodegaOrigen.Value == (int)cmbBodegaDestino.Value)
			{
				MessageBox.Show("La bodega de origen y destino no pueden ser las mismas","Información", MessageBoxButtons.OK,MessageBoxIcon.Information);
				this.cmbBodegaDestino.Focus();
				return;
			}			
			
			if (!FuncionesInventario.ValidaBodegaEnInventario((int)this.cmbBodegaOrigen.Value, Convert.ToDateTime(this.dtFecha.Value), cdsSeteoF)) 
			{
				this.btnCancelar_Click(sender, e);
				return;
			}
			
			if (!Validacion.vNotas(this.txtNotas)) return;			
			if (!Validacion.vbGrids(this.ultraGrid1, "Codigo", 42)) return;
			if (!Validacion.bValidaExistencias(this.ultraGrid1, (DateTime)this.dtFecha.Value, (int)this.txtidCompraTransf.Value, (int)this.cmbBodegaOrigen.Value, 42, 0, cdsSeteoF)) return;
			if (!Validacion.bValidaRepetidosTF(this.ultraGrid1)) return;
			#endregion Validación
		
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
						string sSQLNumero = "Exec NumeracionLocales 42, 0, 0";
						oCmdActualiza.CommandText = sSQLNumero;
						this.txtNumero.Text = oCmdActualiza.ExecuteScalar().ToString();
					}	

					string sSQLGrabaMaestro = string.Format("Exec TransferenciasGrabaMaestro {0}, '{1}', '{2}', {3}, {4}, {5}, '{6}', '{7}', {8}, {9}, '{10}', '{11}', {12}, {13}", 
						(int)this.txtidCompraTransf.Value, this.txtNumero.Text.ToString(), Convert.ToDateTime(this.dtFecha.Value).ToString("yyyyMMdd"), 
						(int)this.cmbBodegaOrigen.Value, (int)this.cmbBodegaDestino.Value, (int)this.txtEstado.Value, 
						this.txtNotas.Text.ToString(), "", 0, 0, 
						"", "", (int)this.cmbRealiza.Value, (int)this.cmbSolicita.Value);

					oCmdActualiza.CommandText = sSQLGrabaMaestro;
					this.txtidCompraTransf.Value = (int)oCmdActualiza.ExecuteScalar();

					foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in ultraGrid1.Rows.All)
					{
						string sSQLGrabaDetalle = string.Format("Exec TransferenciasGrabaDetalle {0}, {1}, {2}, {3}, {4}",
							(int)dr.Cells["idDetCompraTransf"].Value, (int)this.txtidCompraTransf.Value, (int)dr.Cells["idArticulo"].Value, (int)dr.Cells["Cantidad"].Value, (bool)dr.Cells["Borrar"].Value);
						oCmdActualiza.CommandText = sSQLGrabaDetalle;
						oCmdActualiza.ExecuteNonQuery();
					}

					string sSQLActDetalle = string.Format("Exec TrasnferenciasConsulta {0}", (int)this.txtidCompraTransf.Value); 
					oCmdActualiza.CommandText = sSQLActDetalle;
					SqlDataAdapter oDADetalle = new SqlDataAdapter(oCmdActualiza);
					DataTable oDTDetalle = new DataTable();
					oDADetalle.Fill(oDTDetalle);
					this.ultraGrid1.DataSource = oDTDetalle;
										
					string stExec = string.Format("Exec TransformaIE {0}", (int)this.txtidCompraTransf.Value);
					oCmdActualiza.CommandText = stExec;
					oCmdActualiza.ExecuteNonQuery();

					if ((int)this.txtEstado.Value == 0 || (int)this.txtEstado.Value == 1)
					{
						foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in ultraGrid1.Rows.All)
						{
							if ((bool)dr.Cells["Borrar"].Value)
							{
								string sSQLEliminaFilas = string.Format("Exec EliminaFilasTransferencia {0}", (int)dr.Cells["idDetCompraTransf"].Value);
								oCmdActualiza.CommandText = sSQLEliminaFilas;
								oCmdActualiza.ExecuteNonQuery();
							}
						}
					}	

					if (!bNuevo)
					{
						string sSQLBloqueo = string.Format("Exec RegistraBloqueoTransacciones {0}, 0, {1}, 3", idBloqueaTransacciones, (int)this.txtidCompraTransf.Value);
						oCmdActualiza.CommandText = sSQLBloqueo;
						oCmdActualiza.ExecuteNonQuery();
					}

					oTransaction.Commit();

					this.dtFecha.Enabled = false;
					this.cmbBodegaOrigen.Enabled = false;
					this.cmbBodegaDestino.Enabled = false;
					this.cmbRealiza.Enabled = false;
					this.cmbSolicita.Enabled = false;
					this.chkActivar.Enabled = false;					
					this.txtNotas.Enabled = false;
										
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

		private void CargaArticulos(Infragistics.Win.UltraWinGrid.CellEventArgs e, int idArticulo)
		{
			string sArticulo = "";
			string sSQLArticulo = string.Format("Select Articulo From Articulo Where idArticulo = {0}", idArticulo);
			SqlDataReader drArticulo = Funcion.rEscalarSQL(cdsSeteoF, sSQLArticulo, true);
			drArticulo.Read();
			if (drArticulo.HasRows)
			{
				sArticulo = drArticulo.GetString(0);
			}
			drArticulo.Close();

			e.Cell.Row.Cells["idArticulo"].Value = idArticulo;
			e.Cell.Row.Cells["Articulo"].Value = sArticulo;
		}

		private void ultraGrid1_AfterCellUpdate(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			if (!bEdicion) return;
			if (bCombo) return;
			if (!bActualiza) return;
						
			bActualiza = false;
			#region Codigo Articulo
			if (e.Cell.Column.ToString() == "Codigo")
			{
				bActualiza = false;
				FuncionesInventario.FacturacionCargaArticulos(sender, e, this.ultraGrid1, 1, 1, (int)this.cmbBodegaOrigen.Value, (int)this.cmbBodegaOrigen.Value, 0, 0, 0, 0, (DateTime)this.dtFecha.Value, 42, false, true, false, false, 0, 0, 0, 0, 0, false, false, 0, cdsSeteoF);		
				bActualiza = true;
			}
			#endregion Codigo Articulo

			#region Cantidad
			if (e.Cell.Column.ToString() == "Cantidad")
			{	
				FuncionesInventario.ValidaCamposGridEnteros(sender, e, "Cantidad");

				if ((int)this.txtEstado.Value > 0)
				{
					if (e.Cell.Row.Cells["idArticulo"].Value != System.DBNull.Value)
					{
						string sSQLVal = string.Format("Exec ValidaSerialesTransferencia {0}, {1}, {2}, {3}", 
							(int)this.txtidCompraTransf.Value, (int)e.Cell.Row.Cells["idDetCompraTransf"].Value, (int)e.Cell.Row.Cells["Cantidad"].Value, (int)e.Cell.Row.Cells["idArticulo"].Value);
						string sMensaje = Funcion.sEscalarSQL(cdsSeteoF, sSQLVal);

						if (sMensaje.Length > 0)
						{
							e.Cell.Row.Cells["Cantidad"].Value = (int)e.Cell.OriginalValue;
							MessageBox.Show(sMensaje, "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
							return;
						}
					}
				}
			}
			#endregion Cantidad

			bActualiza = true;
		}

		private void ultraGrid1_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{		
			#region Formato numerico
			CultureInfo cultureENUS = CultureInfo.CreateSpecificCulture("ja-jp");
			string stMoneda = "#,##0.";
			string stMascara = "nnn,nnn,nnn.";
			string stCeros = "0000000000";
			string stNumero = "nnnnnnnnnn";
			string stFormato = "";
			string stInput = "";
			#endregion Formato numerico
	
			#region Decimales para la cantidad
			stFormato = stMoneda + stCeros.Substring(0, 0);
			e.Layout.Bands[0].Columns["Cantidad"].Format = stFormato;
			stInput = stMascara + stNumero.Substring(0, 0);
			e.Layout.Bands[0].Columns["Cantidad"].MaskInput = stInput;
			#endregion Decimales para la cantidad
		}

		private void ultraGrid1_AfterRowInsert(object sender, Infragistics.Win.UltraWinGrid.RowEventArgs e)
		{
			if (bEdicion)
			{
				e.Row.Cells["idDetCompraTransf"].Value = 0;
				e.Row.Cells["idCompraTransf"].Value = 0;
				e.Row.Cells["Cantidad"].Value = 1;
				e.Row.Cells["idArticulo"].Value = 0;
				e.Row.Cells["Borrar"].Value = false;				
			}
		}

		private void cmbBodegaDestino_Validating(object sender, System.ComponentModel.CancelEventArgs e)
		{
			if (bEdicion)
			{
				if (this.cmbBodegaOrigen.ActiveRow == null || this.cmbBodegaDestino.ActiveRow == null) return;
			
				if (this.cmbBodegaDestino.ActiveRow == null)
				{
					MessageBox.Show("Seleccione la Bodega Destino", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					this.cmbBodegaDestino.Focus();
					e.Cancel = true;			
				}
				else if ((int)cmbBodegaOrigen.Value == (int)cmbBodegaDestino.Value)
				{
					MessageBox.Show("La Bodega de Origen y Destino no pueden ser las mismas", "Point Technology", MessageBoxButtons.OK,MessageBoxIcon.Stop);
					this.cmbBodegaDestino.Focus();
					e.Cancel = true;
				}
			}
		}

		private void chkActivar_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
		
		}

		private void chkActivar_CheckedChanged(object sender, System.EventArgs e)
		{
			if (bEdicion)
			{
				if ((int)this.txtEstado.Value == 0 || (int)this.txtEstado.Value == 1)
				{
					if (this.chkActivar.Checked)
					{
						this.txtEstado.Value = 1;
						this.lblEstado.Text = "PENDIENTE";
					}
					else 
					{
						this.txtEstado.Value = 0;
						this.lblEstado.Text = "EN ESPERA";
					}
				}
			}
		}

		private void cmbBodegaOrigen_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.cmbBodegaDestino.Focus();
		}

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);		
		}

		private void cmbBodegaOrigen_Validating(object sender, System.ComponentModel.CancelEventArgs e)
		{
			if (bEdicion)
			{
				if (this.cmbBodegaOrigen.ActiveRow == null)
				{
					MessageBox.Show("Seleccione la Bodega Origén", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
					this.cmbBodegaOrigen.Focus();
					e.Cancel = true;
				}

				if (Funcion.iEscalarSQL(cdsSeteoF, string.Format("Select COUNT(*) From CierreConteoInventario Where Bodega = '{0}' And CONVERT(DATE, Fecha) Between '20161001' And '{1}' And Estado = 1", 
					(int)this.cmbBodegaOrigen.Value, Convert.ToDateTime(this.dtFecha.Value).ToString("yyyyMMdd"))) > 0)
				{
					MessageBox.Show("El local esta en Inventario no puede crear Transferencia", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					e.Cancel = true;					
					return;
				}
			}
		}

		private void cmbBodegaDestino_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}

		private void dtFecha_BeforeDropDown(object sender, System.ComponentModel.CancelEventArgs e)
		{
		
		}

		private void cmbBodegaOrigen_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
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

		private void frmTransferencias_Closing(object sender, System.ComponentModel.CancelEventArgs e)
		{
			if (bEdicion)
			{
				MessageBox.Show("No puede salir mientras edita una Transferencia", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				e.Cancel = true;
			}
		}

		private void frmTransferencias_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (!bEdicion)
			{
				if (e.KeyCode == Keys.N) if (miAcceso.Nuevo) this.btnNuevo_Click(sender, e);
				if (e.KeyCode == Keys.B) if (miAcceso.Buscar) this.btnConsultar_Click(sender, e);
				if ((int)this.txtidCompraTransf.Value > 0)
				{
					if (e.KeyCode == Keys.E) this.btnEditar_Click(sender, e);
					if (e.KeyCode == Keys.I) this.btnImprimir_Click(sender, e);
					if (e.KeyCode == Keys.A) this.btnAnular_Click(sender, e);					
					if (e.KeyCode == Keys.F5)
					{
						miAcceso = new Acceso(cdsSeteoF, "0315");
						this.Consultar((int)this.txtidCompraTransf.Value);
					}
				}
				if (e.KeyCode == Keys.Escape) this.Close();					
			}
			else
			{
				if (e.KeyCode == Keys.Escape)
					if (DialogResult.Yes == MessageBox.Show("Esta seguro de Cancelar la Edición", "Point Technology", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2))
						this.btnCancelar_Click(sender, e);

				if (e.KeyCode == Keys.F12) 
					this.btnGuardar_Click(sender, e);
			}
		}

		private void dtFecha_ValueChanged(object sender, System.EventArgs e)
		{
			if (this.dtFecha.Value != System.DBNull.Value)
				FuncionesProcedimientos.ListaBodegas(this.cmbBodegaOrigen, MenuLatinium.IdUsuario, 42, (DateTime)this.dtFecha.Value, false, cdsSeteoF);	
		}

		private void cmbBodegaOrigen_ValueChanged(object sender, System.EventArgs e)
		{
			if (bEdicion)
			{
				this.cmbRealiza.Value = System.DBNull.Value;
				this.cmbBodegaDestino.Value = System.DBNull.Value;
				this.cmbSolicita.Value = System.DBNull.Value;

				if (this.ultraGrid1.Rows.Count > 0 && bEdicion) 
					this.ultraGrid1.DataSource = Funcion.dvProcedimiento(cdsSeteoF, string.Format("Exec TrasnferenciasConsulta {0}", (int)this.txtidCompraTransf.Value));				
			}

			if (this.cmbBodegaOrigen.ActiveRow != null && this.dtFecha.Value != System.DBNull.Value)
			{
				bool bConsignacion = Funcion.bEscalarSQL(cdsSeteoF, string.Format("Select Consignacion From Bodega Where Bodega = {0}", (int)this.cmbBodegaOrigen.Value), true);

				int iBodegaOrigen = (int)this.cmbBodegaOrigen.Value;
				if (bConsignacion)
					iBodegaOrigen = Funcion.iEscalarSQL(cdsSeteoF, string.Format("select Local from AsignaBodegasLocales Where Bodega = {0} And Consignacion = 1", (int)this.cmbBodegaOrigen.Value));

				this.cmbRealiza.DataSource = Funcion.dvProcedimiento(cdsSeteoF, string.Format("Exec PersonalTransferencias '{0}', {1}, {2}, {3}", 
					Convert.ToDateTime(this.dtFecha.Value).ToString("yyyyMMdd"), iBodegaOrigen, MenuLatinium.IdUsuario, iBodegaPredef));
				
				if (iBodegaPredef == 0 && bNuevo && (MenuLatinium.iNivel != 23 && MenuLatinium.iNivel != 25))
					this.cmbRealiza.Value = Funcion.iEscalarSQL(cdsSeteoF, string.Format("If Exists(Select idPersonal From Nomina Where Codigo In (Select Nombre From Usuario where idUsuario = {0})) Select idPersonal From Nomina Where Codigo In (Select Nombre From Usuario where idUsuario = {0}) Else Select 0", MenuLatinium.IdUsuario));					
				else
				{
					if (bEdicion) 
						this.cmbRealiza.Enabled = true;
				}

				this.cmbBodegaDestino.DataSource = Funcion.dvProcedimiento(cdsSeteoF, string.Format("Exec ListaDeBodegasActivas '{0}', {1}, {2}", 
					Convert.ToDateTime(this.dtFecha.Value).ToString("yyyyMMdd"), (int)this.cmbBodegaOrigen.Value, bConsignacion));
			}
		}

		private void cmbSolicita_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter)
			{				
				this.ultraGrid1.DisplayLayout.Bands[0].AddNew();
				this.ultraGrid1.ActiveCell = this.ultraGrid1.Rows[0].Cells["Codigo"];
			}
		}

		private void cmbBodegaDestino_ValueChanged(object sender, System.EventArgs e)
		{			
			if (bEdicion) this.cmbSolicita.Value = System.DBNull.Value;

			if (this.cmbBodegaDestino.ActiveRow != null && this.dtFecha.Value != System.DBNull.Value)
				this.cmbSolicita.DataSource = Funcion.dvProcedimiento(cdsSeteoF, string.Format("Exec PersonalTransferencias '{0}', {1}, {2}, 1", 
					Convert.ToDateTime(this.dtFecha.Value).ToString("yyyyMMdd"), (int)this.cmbBodegaDestino.Value, MenuLatinium.IdUsuario));
		}

		private void cmbSolicita_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}

		private void dtFecha_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (e.KeyChar == 8) e.Handled = true;
		}
	}
}

