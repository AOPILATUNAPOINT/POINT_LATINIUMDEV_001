using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;
using System.Data.SqlClient;
using DevExpress.XtraEditors;
using CrystalDecisions;
using CrystalDecisions.CrystalReports.Engine;
using CrystalDecisions.ReportSource;
using CrystalDecisions.Shared;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de frmRecibosManuales.
	/// </summary>
	public class frmRecibosManuales : System.Windows.Forms.Form
	{
		private System.Windows.Forms.Label label6;
		public Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtFecha;
		private System.Windows.Forms.Label label1;
		private System.Windows.Forms.Label label2;
		private System.Windows.Forms.Label label3;
		private System.Windows.Forms.Label label4;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource2;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource1;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource3;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource4;
		private System.Windows.Forms.Label label5;
		private System.Windows.Forms.Label label7;
		private System.Windows.Forms.Label label8;
		private System.Windows.Forms.Label label14;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtNumero;
		private System.Windows.Forms.Label label9;
		private System.Windows.Forms.GroupBox groupBox2;
		private System.Windows.Forms.Label label10;
		private System.Windows.Forms.Label label11;
		private System.Windows.Forms.GroupBox groupBox1;
		private System.Windows.Forms.Button btnSalir;
		private System.Windows.Forms.Button btnNuevo;
		private System.Windows.Forms.Button btnConsultar;
		private System.Windows.Forms.Button btnEditar;
		public System.Windows.Forms.Button btnGuardar;
		private System.Windows.Forms.Button btnCancelar;
		private System.Windows.Forms.Button btnAnular;
		public Infragistics.Win.UltraWinEditors.UltraNumericEditor txtEstado;
		public Infragistics.Win.UltraWinEditors.UltraNumericEditor txtidRecibosManuales;
		private System.Windows.Forms.Label lblEstado;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbEnvia;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbRecepta;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbBodega;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbLocal;
		public Infragistics.Win.UltraWinEditors.UltraTextEditor txtNota;
		public Infragistics.Win.UltraWinEditors.UltraTextEditor txtNotarecibe;
		public Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtRecibe;
		public Infragistics.Win.UltraWinEditors.UltraNumericEditor txtDesde;
		public Infragistics.Win.UltraWinEditors.UltraNumericEditor txtHasta;
		private C1.Data.C1DataSet cdsSeteoF;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource5;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbCodigo;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public frmRecibosManuales()
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
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton1 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmRecibosManuales));
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idNomina");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idNomina");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idNomina");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn3 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idNomina");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn4 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand3 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bodega");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn5 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Bodega");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn6 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand4 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bodega");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn7 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Bodega");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn8 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton2 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn9 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCodigoRecibo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn10 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand5 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCodigoRecibo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			this.label6 = new System.Windows.Forms.Label();
			this.dtFecha = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.cmbEnvia = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.ultraDataSource1 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.cmbRecepta = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.ultraDataSource4 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.cmbBodega = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.ultraDataSource2 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.cmbLocal = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.ultraDataSource3 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.label1 = new System.Windows.Forms.Label();
			this.label2 = new System.Windows.Forms.Label();
			this.label3 = new System.Windows.Forms.Label();
			this.label4 = new System.Windows.Forms.Label();
			this.label5 = new System.Windows.Forms.Label();
			this.txtDesde = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtHasta = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label7 = new System.Windows.Forms.Label();
			this.label8 = new System.Windows.Forms.Label();
			this.label14 = new System.Windows.Forms.Label();
			this.txtNumero = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtNota = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label9 = new System.Windows.Forms.Label();
			this.groupBox2 = new System.Windows.Forms.GroupBox();
			this.label10 = new System.Windows.Forms.Label();
			this.txtNotarecibe = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label11 = new System.Windows.Forms.Label();
			this.dtRecibe = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.groupBox1 = new System.Windows.Forms.GroupBox();
			this.btnSalir = new System.Windows.Forms.Button();
			this.btnNuevo = new System.Windows.Forms.Button();
			this.btnConsultar = new System.Windows.Forms.Button();
			this.btnEditar = new System.Windows.Forms.Button();
			this.btnGuardar = new System.Windows.Forms.Button();
			this.btnCancelar = new System.Windows.Forms.Button();
			this.btnAnular = new System.Windows.Forms.Button();
			this.txtEstado = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtidRecibosManuales = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.lblEstado = new System.Windows.Forms.Label();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.ultraDataSource5 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.cmbCodigo = new Infragistics.Win.UltraWinGrid.UltraCombo();
			((System.ComponentModel.ISupportInitialize)(this.dtFecha)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbEnvia)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbRecepta)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource4)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbBodega)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbLocal)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource3)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDesde)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtHasta)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNumero)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNota)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNotarecibe)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtRecibe)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtEstado)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtidRecibosManuales)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource5)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCodigo)).BeginInit();
			this.SuspendLayout();
			// 
			// label6
			// 
			this.label6.AutoSize = true;
			this.label6.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label6.Location = new System.Drawing.Point(8, 10);
			this.label6.Name = "label6";
			this.label6.Size = new System.Drawing.Size(36, 16);
			this.label6.TabIndex = 201;
			this.label6.Text = "Fecha";
			// 
			// dtFecha
			// 
			appearance1.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtFecha.Appearance = appearance1;
			this.dtFecha.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton1.Caption = "Today";
			this.dtFecha.DateButtons.Add(dateButton1);
			this.dtFecha.Enabled = false;
			this.dtFecha.Format = "dd/MM/yyyy";
			this.dtFecha.Location = new System.Drawing.Point(64, 8);
			this.dtFecha.Name = "dtFecha";
			this.dtFecha.NonAutoSizeHeight = 23;
			this.dtFecha.Size = new System.Drawing.Size(136, 21);
			this.dtFecha.SpinButtonsVisible = true;
			this.dtFecha.TabIndex = 200;
			this.dtFecha.Value = ((object)(resources.GetObject("dtFecha.Value")));
			this.dtFecha.ValueChanged += new System.EventHandler(this.dtFecha_ValueChanged);
			// 
			// cmbEnvia
			// 
			appearance2.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbEnvia.Appearance = appearance2;
			this.cmbEnvia.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbEnvia.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbEnvia.DataSource = this.ultraDataSource1;
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridColumn2.Width = 180;
			ultraGridBand1.Columns.AddRange(new object[] {
																										 ultraGridColumn1,
																										 ultraGridColumn2});
			this.cmbEnvia.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.cmbEnvia.DisplayMember = "Nombre";
			this.cmbEnvia.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbEnvia.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbEnvia.Enabled = false;
			this.cmbEnvia.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbEnvia.Location = new System.Drawing.Point(360, 40);
			this.cmbEnvia.Name = "cmbEnvia";
			this.cmbEnvia.Size = new System.Drawing.Size(256, 21);
			this.cmbEnvia.TabIndex = 840;
			this.cmbEnvia.ValueMember = "idNomina";
			this.cmbEnvia.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cmbEnvia_KeyDown);
			// 
			// ultraDataSource1
			// 
			ultraDataColumn1.DataType = typeof(int);
			this.ultraDataSource1.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn1,
																																 ultraDataColumn2});
			// 
			// cmbRecepta
			// 
			appearance3.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbRecepta.Appearance = appearance3;
			this.cmbRecepta.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbRecepta.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbRecepta.DataSource = this.ultraDataSource4;
			ultraGridColumn3.Header.VisiblePosition = 0;
			ultraGridColumn3.Hidden = true;
			ultraGridColumn4.Header.VisiblePosition = 1;
			ultraGridColumn4.Width = 180;
			ultraGridBand2.Columns.AddRange(new object[] {
																										 ultraGridColumn3,
																										 ultraGridColumn4});
			this.cmbRecepta.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			this.cmbRecepta.DisplayMember = "Nombre";
			this.cmbRecepta.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbRecepta.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbRecepta.Enabled = false;
			this.cmbRecepta.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbRecepta.Location = new System.Drawing.Point(360, 70);
			this.cmbRecepta.Name = "cmbRecepta";
			this.cmbRecepta.Size = new System.Drawing.Size(256, 21);
			this.cmbRecepta.TabIndex = 841;
			this.cmbRecepta.ValueMember = "idNomina";
			this.cmbRecepta.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cmbRecepta_KeyDown);
			// 
			// ultraDataSource4
			// 
			ultraDataColumn3.DataType = typeof(int);
			this.ultraDataSource4.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn3,
																																 ultraDataColumn4});
			// 
			// cmbBodega
			// 
			appearance4.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbBodega.Appearance = appearance4;
			this.cmbBodega.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbBodega.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbBodega.DataSource = this.ultraDataSource2;
			ultraGridColumn5.Header.VisiblePosition = 0;
			ultraGridColumn5.Hidden = true;
			ultraGridColumn6.Header.VisiblePosition = 1;
			ultraGridColumn6.Width = 275;
			ultraGridBand3.Columns.AddRange(new object[] {
																										 ultraGridColumn5,
																										 ultraGridColumn6});
			this.cmbBodega.DisplayLayout.BandsSerializer.Add(ultraGridBand3);
			this.cmbBodega.DisplayMember = "Nombre";
			this.cmbBodega.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbBodega.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbBodega.Enabled = false;
			this.cmbBodega.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbBodega.Location = new System.Drawing.Point(64, 40);
			this.cmbBodega.Name = "cmbBodega";
			this.cmbBodega.Size = new System.Drawing.Size(240, 21);
			this.cmbBodega.TabIndex = 842;
			this.cmbBodega.ValueMember = "Bodega";
			this.cmbBodega.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cmbBodega_KeyDown);
			// 
			// ultraDataSource2
			// 
			ultraDataColumn5.DataType = typeof(int);
			this.ultraDataSource2.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn5,
																																 ultraDataColumn6});
			// 
			// cmbLocal
			// 
			appearance5.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbLocal.Appearance = appearance5;
			this.cmbLocal.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbLocal.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbLocal.DataSource = this.ultraDataSource3;
			ultraGridColumn7.Header.VisiblePosition = 0;
			ultraGridColumn7.Hidden = true;
			ultraGridColumn8.Header.VisiblePosition = 1;
			ultraGridColumn8.Width = 270;
			ultraGridBand4.Columns.AddRange(new object[] {
																										 ultraGridColumn7,
																										 ultraGridColumn8});
			this.cmbLocal.DisplayLayout.BandsSerializer.Add(ultraGridBand4);
			this.cmbLocal.DisplayMember = "Nombre";
			this.cmbLocal.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbLocal.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbLocal.Enabled = false;
			this.cmbLocal.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbLocal.Location = new System.Drawing.Point(64, 72);
			this.cmbLocal.Name = "cmbLocal";
			this.cmbLocal.Size = new System.Drawing.Size(240, 21);
			this.cmbLocal.TabIndex = 843;
			this.cmbLocal.ValueMember = "Bodega";
			this.cmbLocal.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cmbLocal_KeyDown);
			this.cmbLocal.ValueChanged += new System.EventHandler(this.cmbLocal_ValueChanged);
			// 
			// ultraDataSource3
			// 
			ultraDataColumn7.DataType = typeof(int);
			this.ultraDataSource3.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn7,
																																 ultraDataColumn8});
			// 
			// label1
			// 
			this.label1.AutoSize = true;
			this.label1.Enabled = false;
			this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label1.Location = new System.Drawing.Point(312, 40);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(33, 16);
			this.label1.TabIndex = 844;
			this.label1.Text = "Envia";
			// 
			// label2
			// 
			this.label2.AutoSize = true;
			this.label2.Enabled = false;
			this.label2.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label2.Location = new System.Drawing.Point(312, 72);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(46, 16);
			this.label2.TabIndex = 845;
			this.label2.Text = "Recepta";
			// 
			// label3
			// 
			this.label3.AutoSize = true;
			this.label3.Enabled = false;
			this.label3.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label3.Location = new System.Drawing.Point(8, 40);
			this.label3.Name = "label3";
			this.label3.Size = new System.Drawing.Size(43, 16);
			this.label3.TabIndex = 846;
			this.label3.Text = "Bodega";
			// 
			// label4
			// 
			this.label4.AutoSize = true;
			this.label4.Enabled = false;
			this.label4.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label4.Location = new System.Drawing.Point(8, 72);
			this.label4.Name = "label4";
			this.label4.Size = new System.Drawing.Size(31, 16);
			this.label4.TabIndex = 847;
			this.label4.Text = "Local";
			// 
			// label5
			// 
			this.label5.AutoSize = true;
			this.label5.Enabled = false;
			this.label5.Location = new System.Drawing.Point(440, 106);
			this.label5.Name = "label5";
			this.label5.Size = new System.Drawing.Size(40, 16);
			this.label5.TabIndex = 848;
			this.label5.Text = "Codigo";
			this.label5.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// txtDesde
			// 
			this.txtDesde.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtDesde.Enabled = false;
			this.txtDesde.Location = new System.Drawing.Point(64, 102);
			this.txtDesde.Name = "txtDesde";
			this.txtDesde.PromptChar = ' ';
			this.txtDesde.Size = new System.Drawing.Size(104, 21);
			this.txtDesde.SpinButtonAlignment = Infragistics.Win.ButtonAlignment.Left;
			this.txtDesde.TabIndex = 865;
			this.txtDesde.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtDesde_KeyDown);
			this.txtDesde.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtDesde_KeyPress);
			this.txtDesde.Enter += new System.EventHandler(this.txtDesde_Enter);
			// 
			// txtHasta
			// 
			this.txtHasta.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtHasta.Enabled = false;
			this.txtHasta.Location = new System.Drawing.Point(240, 102);
			this.txtHasta.Name = "txtHasta";
			this.txtHasta.PromptChar = ' ';
			this.txtHasta.Size = new System.Drawing.Size(104, 21);
			this.txtHasta.SpinButtonAlignment = Infragistics.Win.ButtonAlignment.Left;
			this.txtHasta.TabIndex = 866;
			this.txtHasta.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtHasta_KeyDown);
			this.txtHasta.EnabledChanged += new System.EventHandler(this.txtHasta_EnabledChanged);
			this.txtHasta.Enter += new System.EventHandler(this.txtHasta_Enter);
			// 
			// label7
			// 
			this.label7.AutoSize = true;
			this.label7.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label7.Location = new System.Drawing.Point(8, 104);
			this.label7.Name = "label7";
			this.label7.Size = new System.Drawing.Size(37, 16);
			this.label7.TabIndex = 867;
			this.label7.Text = "Desde";
			// 
			// label8
			// 
			this.label8.AutoSize = true;
			this.label8.Enabled = false;
			this.label8.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label8.Location = new System.Drawing.Point(184, 104);
			this.label8.Name = "label8";
			this.label8.Size = new System.Drawing.Size(34, 16);
			this.label8.TabIndex = 868;
			this.label8.Text = "Hasta";
			// 
			// label14
			// 
			this.label14.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.label14.AutoSize = true;
			this.label14.BackColor = System.Drawing.Color.Transparent;
			this.label14.Enabled = false;
			this.label14.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label14.Location = new System.Drawing.Point(520, 8);
			this.label14.Name = "label14";
			this.label14.Size = new System.Drawing.Size(44, 17);
			this.label14.TabIndex = 985;
			this.label14.Text = "Número";
			this.label14.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtNumero
			// 
			this.txtNumero.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			appearance6.ForeColor = System.Drawing.Color.Black;
			appearance6.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtNumero.Appearance = appearance6;
			this.txtNumero.Enabled = false;
			this.txtNumero.Font = new System.Drawing.Font("Tahoma", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtNumero.Location = new System.Drawing.Point(576, 8);
			this.txtNumero.Name = "txtNumero";
			this.txtNumero.Size = new System.Drawing.Size(168, 23);
			this.txtNumero.TabIndex = 984;
			// 
			// txtNota
			// 
			appearance7.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtNota.Appearance = appearance7;
			this.txtNota.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtNota.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNota.Enabled = false;
			this.txtNota.Location = new System.Drawing.Point(64, 136);
			this.txtNota.Multiline = true;
			this.txtNota.Name = "txtNota";
			this.txtNota.Size = new System.Drawing.Size(680, 32);
			this.txtNota.TabIndex = 986;
			this.txtNota.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtNota_KeyDown);
			// 
			// label9
			// 
			this.label9.AutoSize = true;
			this.label9.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label9.Location = new System.Drawing.Point(8, 136);
			this.label9.Name = "label9";
			this.label9.Size = new System.Drawing.Size(28, 16);
			this.label9.TabIndex = 987;
			this.label9.Text = "Nota";
			// 
			// groupBox2
			// 
			this.groupBox2.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.groupBox2.Enabled = false;
			this.groupBox2.Location = new System.Drawing.Point(-48, 176);
			this.groupBox2.Name = "groupBox2";
			this.groupBox2.Size = new System.Drawing.Size(880, 8);
			this.groupBox2.TabIndex = 988;
			this.groupBox2.TabStop = false;
			// 
			// label10
			// 
			this.label10.AutoSize = true;
			this.label10.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label10.Location = new System.Drawing.Point(8, 224);
			this.label10.Name = "label10";
			this.label10.Size = new System.Drawing.Size(28, 16);
			this.label10.TabIndex = 990;
			this.label10.Text = "Nota";
			// 
			// txtNotarecibe
			// 
			appearance8.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtNotarecibe.Appearance = appearance8;
			this.txtNotarecibe.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtNotarecibe.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNotarecibe.Enabled = false;
			this.txtNotarecibe.Location = new System.Drawing.Point(64, 224);
			this.txtNotarecibe.Multiline = true;
			this.txtNotarecibe.Name = "txtNotarecibe";
			this.txtNotarecibe.Size = new System.Drawing.Size(680, 48);
			this.txtNotarecibe.TabIndex = 989;
			this.txtNotarecibe.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtNotarecibe_KeyDown);
			// 
			// label11
			// 
			this.label11.AutoSize = true;
			this.label11.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label11.Location = new System.Drawing.Point(8, 192);
			this.label11.Name = "label11";
			this.label11.Size = new System.Drawing.Size(39, 16);
			this.label11.TabIndex = 992;
			this.label11.Text = "Recibe";
			// 
			// dtRecibe
			// 
			appearance9.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtRecibe.Appearance = appearance9;
			this.dtRecibe.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton2.Caption = "Today";
			this.dtRecibe.DateButtons.Add(dateButton2);
			this.dtRecibe.Enabled = false;
			this.dtRecibe.Format = "dd/MM/yyyy";
			this.dtRecibe.Location = new System.Drawing.Point(64, 192);
			this.dtRecibe.Name = "dtRecibe";
			this.dtRecibe.NonAutoSizeHeight = 23;
			this.dtRecibe.Size = new System.Drawing.Size(136, 21);
			this.dtRecibe.SpinButtonsVisible = true;
			this.dtRecibe.TabIndex = 991;
			this.dtRecibe.Value = ((object)(resources.GetObject("dtRecibe.Value")));
			this.dtRecibe.KeyDown += new System.Windows.Forms.KeyEventHandler(this.dtRecibe_KeyDown);
			// 
			// groupBox1
			// 
			this.groupBox1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.groupBox1.Location = new System.Drawing.Point(-48, 280);
			this.groupBox1.Name = "groupBox1";
			this.groupBox1.Size = new System.Drawing.Size(880, 8);
			this.groupBox1.TabIndex = 993;
			this.groupBox1.TabStop = false;
			// 
			// btnSalir
			// 
			this.btnSalir.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnSalir.CausesValidation = false;
			this.btnSalir.Image = ((System.Drawing.Image)(resources.GetObject("btnSalir.Image")));
			this.btnSalir.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnSalir.Location = new System.Drawing.Point(488, 297);
			this.btnSalir.Name = "btnSalir";
			this.btnSalir.TabIndex = 1018;
			this.btnSalir.Text = "&Salir";
			this.btnSalir.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnSalir.Click += new System.EventHandler(this.btnSalir_Click);
			// 
			// btnNuevo
			// 
			this.btnNuevo.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnNuevo.CausesValidation = false;
			this.btnNuevo.Enabled = false;
			this.btnNuevo.Image = ((System.Drawing.Image)(resources.GetObject("btnNuevo.Image")));
			this.btnNuevo.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnNuevo.Location = new System.Drawing.Point(8, 296);
			this.btnNuevo.Name = "btnNuevo";
			this.btnNuevo.Size = new System.Drawing.Size(72, 24);
			this.btnNuevo.TabIndex = 1015;
			this.btnNuevo.Text = "&Nuevo";
			this.btnNuevo.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnNuevo.Click += new System.EventHandler(this.btnNuevo_Click);
			// 
			// btnConsultar
			// 
			this.btnConsultar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnConsultar.CausesValidation = false;
			this.btnConsultar.Enabled = false;
			this.btnConsultar.Image = ((System.Drawing.Image)(resources.GetObject("btnConsultar.Image")));
			this.btnConsultar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnConsultar.Location = new System.Drawing.Point(88, 296);
			this.btnConsultar.Name = "btnConsultar";
			this.btnConsultar.Size = new System.Drawing.Size(78, 24);
			this.btnConsultar.TabIndex = 1014;
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
			this.btnEditar.Location = new System.Drawing.Point(168, 296);
			this.btnEditar.Name = "btnEditar";
			this.btnEditar.Size = new System.Drawing.Size(76, 24);
			this.btnEditar.TabIndex = 1013;
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
			this.btnGuardar.Location = new System.Drawing.Point(248, 296);
			this.btnGuardar.Name = "btnGuardar";
			this.btnGuardar.Size = new System.Drawing.Size(76, 24);
			this.btnGuardar.TabIndex = 1012;
			this.btnGuardar.Text = "&Guardar";
			this.btnGuardar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnGuardar.Click += new System.EventHandler(this.btnGuardar_Click);
			// 
			// btnCancelar
			// 
			this.btnCancelar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnCancelar.CausesValidation = false;
			this.btnCancelar.Enabled = false;
			this.btnCancelar.Image = ((System.Drawing.Image)(resources.GetObject("btnCancelar.Image")));
			this.btnCancelar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnCancelar.Location = new System.Drawing.Point(408, 296);
			this.btnCancelar.Name = "btnCancelar";
			this.btnCancelar.Size = new System.Drawing.Size(72, 24);
			this.btnCancelar.TabIndex = 1011;
			this.btnCancelar.Text = "&Cancelar";
			this.btnCancelar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnCancelar.Click += new System.EventHandler(this.btnCancelar_Click);
			// 
			// btnAnular
			// 
			this.btnAnular.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnAnular.CausesValidation = false;
			this.btnAnular.Enabled = false;
			this.btnAnular.Image = ((System.Drawing.Image)(resources.GetObject("btnAnular.Image")));
			this.btnAnular.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnAnular.Location = new System.Drawing.Point(328, 297);
			this.btnAnular.Name = "btnAnular";
			this.btnAnular.Size = new System.Drawing.Size(72, 23);
			this.btnAnular.TabIndex = 1016;
			this.btnAnular.Text = "&Anular";
			this.btnAnular.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnAnular.Click += new System.EventHandler(this.btnAnular_Click);
			// 
			// txtEstado
			// 
			this.txtEstado.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtEstado.Enabled = false;
			this.txtEstado.Location = new System.Drawing.Point(664, 192);
			this.txtEstado.Name = "txtEstado";
			this.txtEstado.PromptChar = ' ';
			this.txtEstado.Size = new System.Drawing.Size(56, 21);
			this.txtEstado.SpinButtonAlignment = Infragistics.Win.ButtonAlignment.Left;
			this.txtEstado.TabIndex = 1020;
			this.txtEstado.Visible = false;
			// 
			// txtidRecibosManuales
			// 
			this.txtidRecibosManuales.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtidRecibosManuales.Enabled = false;
			this.txtidRecibosManuales.Location = new System.Drawing.Point(600, 192);
			this.txtidRecibosManuales.Name = "txtidRecibosManuales";
			this.txtidRecibosManuales.PromptChar = ' ';
			this.txtidRecibosManuales.Size = new System.Drawing.Size(56, 21);
			this.txtidRecibosManuales.SpinButtonAlignment = Infragistics.Win.ButtonAlignment.Left;
			this.txtidRecibosManuales.TabIndex = 1019;
			this.txtidRecibosManuales.Visible = false;
			// 
			// lblEstado
			// 
			this.lblEstado.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.lblEstado.AutoSize = true;
			this.lblEstado.Enabled = false;
			this.lblEstado.Font = new System.Drawing.Font("Microsoft Sans Serif", 15F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblEstado.ForeColor = System.Drawing.Color.Firebrick;
			this.lblEstado.Location = new System.Drawing.Point(376, 192);
			this.lblEstado.Name = "lblEstado";
			this.lblEstado.Size = new System.Drawing.Size(0, 26);
			this.lblEstado.TabIndex = 1021;
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
			// ultraDataSource5
			// 
			ultraDataColumn9.DataType = typeof(int);
			this.ultraDataSource5.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn9,
																																 ultraDataColumn10});
			// 
			// cmbCodigo
			// 
			appearance10.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbCodigo.Appearance = appearance10;
			this.cmbCodigo.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbCodigo.Cursor = System.Windows.Forms.Cursors.WaitCursor;
			this.cmbCodigo.DataSource = this.ultraDataSource5;
			ultraGridColumn9.Header.VisiblePosition = 0;
			ultraGridColumn10.Header.VisiblePosition = 1;
			ultraGridBand5.Columns.AddRange(new object[] {
																										 ultraGridColumn9,
																										 ultraGridColumn10});
			this.cmbCodigo.DisplayLayout.BandsSerializer.Add(ultraGridBand5);
			this.cmbCodigo.DisplayMember = "Nombre";
			this.cmbCodigo.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbCodigo.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbCodigo.Enabled = false;
			this.cmbCodigo.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbCodigo.Location = new System.Drawing.Point(488, 104);
			this.cmbCodigo.Name = "cmbCodigo";
			this.cmbCodigo.Size = new System.Drawing.Size(128, 21);
			this.cmbCodigo.TabIndex = 1023;
			this.cmbCodigo.ValueMember = "idCodigoRecibo";
			// 
			// frmRecibosManuales
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(752, 334);
			this.Controls.Add(this.cmbCodigo);
			this.Controls.Add(this.lblEstado);
			this.Controls.Add(this.txtEstado);
			this.Controls.Add(this.txtidRecibosManuales);
			this.Controls.Add(this.btnSalir);
			this.Controls.Add(this.btnNuevo);
			this.Controls.Add(this.btnConsultar);
			this.Controls.Add(this.btnEditar);
			this.Controls.Add(this.btnGuardar);
			this.Controls.Add(this.btnCancelar);
			this.Controls.Add(this.btnAnular);
			this.Controls.Add(this.groupBox1);
			this.Controls.Add(this.label11);
			this.Controls.Add(this.label10);
			this.Controls.Add(this.label9);
			this.Controls.Add(this.label14);
			this.Controls.Add(this.label8);
			this.Controls.Add(this.label7);
			this.Controls.Add(this.label5);
			this.Controls.Add(this.label4);
			this.Controls.Add(this.label3);
			this.Controls.Add(this.label2);
			this.Controls.Add(this.label1);
			this.Controls.Add(this.label6);
			this.Controls.Add(this.dtRecibe);
			this.Controls.Add(this.txtNotarecibe);
			this.Controls.Add(this.groupBox2);
			this.Controls.Add(this.txtNota);
			this.Controls.Add(this.txtNumero);
			this.Controls.Add(this.txtHasta);
			this.Controls.Add(this.txtDesde);
			this.Controls.Add(this.cmbLocal);
			this.Controls.Add(this.cmbBodega);
			this.Controls.Add(this.cmbRecepta);
			this.Controls.Add(this.cmbEnvia);
			this.Controls.Add(this.dtFecha);
			this.KeyPreview = true;
			this.MaximizeBox = false;
			this.Name = "frmRecibosManuales";
			this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
			this.Text = "Recibos Manuales";
			this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.frmRecibosManuales_KeyDown);
			this.Load += new System.EventHandler(this.frmRecibosManuales_Load);
			((System.ComponentModel.ISupportInitialize)(this.dtFecha)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbEnvia)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbRecepta)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource4)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbBodega)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbLocal)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource3)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDesde)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtHasta)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNumero)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNota)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNotarecibe)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtRecibe)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtEstado)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtidRecibosManuales)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource5)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCodigo)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private Acceso miAcceso;
		private void UnloadMe()
		{
			this.Close();
		}

		bool bNuevo = false;
		bool bEdicion = false;
		bool bAnulado = false;
		bool bEliminaAlValidar = false;


	

		private void frmRecibosManuales_Load(object sender, System.EventArgs e)
		{
			miAcceso = new Acceso(cdsSeteoF, "0839");

			if (!Funcion.Permiso("1436", cdsSeteoF))
			{				
				MessageBox.Show("No puede ingresar a Recibos Manuales ", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				BeginInvoke(new MethodInvoker(UnloadMe));
				return;
			}

			if (miAcceso.Nuevo) this.btnNuevo.Enabled = true;
			if (miAcceso.Buscar) this.btnConsultar.Enabled = true;

			this.cmbEnvia.DataSource = Funcion.dvProcedimiento(cdsSeteoF, string.Format("Exec RetornaNombres {0}", MenuLatinium.IdUsuario));
			this.cmbBodega.DataSource = Funcion.dvProcedimiento(cdsSeteoF, string.Format("EXEC [RetornaBodega]"));
			this.cmbCodigo.DataSource = Funcion.dvProcedimiento(cdsSeteoF, string.Format("EXEC Listas 91 , 0"));
		}

		private void dtFecha_ValueChanged(object sender, System.EventArgs e)
		{
			if (this.dtFecha.Value != System.DBNull.Value)
				this.cmbLocal.DataSource = Funcion.dvProcedimiento(cdsSeteoF, string.Format("Exec FacturacionListaBodegas {0}, 24, '{1}'", MenuLatinium.IdUsuario, Convert.ToDateTime(this.dtFecha.Value).ToString("yyyyMMdd")));
		
		}

		private void cmbLocal_ValueChanged(object sender, System.EventArgs e)
		{
			if (this.cmbLocal.ActiveRow != null)
			{
				if (this.cmbRecepta.ActiveRow != null) this.cmbRecepta.Value = System.DBNull.Value;
				//				this.cmbRecibidopor.DataSource = FuncionesProcedimientos.Lista_Vendedores(Convert.ToDateTime(this.dtFecha.Value), (int)this.cmbBodega.Value, cdsSeteoF);
				this.cmbRecepta.DataSource = FuncionesProcedimientos.dtGeneral(string.Format("Select  n.idNomina, (n.ApellidoPaterno + ' ' + n.ApellidoMaterno + ' ' + n.PrimerNombre + ' ' + n.SegundoNombre) As Nombre, n.Codigo From Com_AsignacionDeVendedores a Inner Join Nomina n On n.idPersonal = a.idPersonal Where a.idCargo in ( 2 ,1)And a.Bodega = {0} And '{1}' Between CONVERT(Date, a.Desde) And CONVERT(Date, a.Hasta)",
					this.cmbLocal.Value, Convert.ToDateTime(this.dtFecha.Value).ToString("yyyyMMdd")));	
	
//				if ((int) this.txtidRecibosManuales.Value == 0)
//				{
//					string sSQPer = string.Format(" Select Codigo from bodega Where bodega = {0} ", (int)this.cmbLocal.Value);	
//					this.txtCodigo.Value = Funcion.sEscalarSQL(cdsSeteoF, sSQPer);
//				}
				
			}
		}

		private void btnCancelar_Click(object sender, System.EventArgs e)
		{
			bNuevo = false;
			bEdicion = false;
			#region Vacia
			this.dtFecha.Value  = System.DBNull.Value;
			this.txtNumero.Text ="";
			this.cmbBodega.Value  = System.DBNull.Value;
			this.cmbLocal.Value  = System.DBNull.Value;
			this.cmbEnvia.Value  = System.DBNull.Value;
			this.cmbRecepta.Value  = System.DBNull.Value;
			this.cmbCodigo.Value  = System.DBNull.Value;
//      this.txtCodigo.Text ="";
			this.txtDesde.Value = 0;
			this.txtHasta.Value = 0;
			this.dtRecibe.Value  = System.DBNull.Value;
			this.txtNota.Text ="";
			this.txtNotarecibe.Text ="";
			this.txtEstado.Value = 0;
			this.txtidRecibosManuales.Value = 0;
			this.lblEstado.Text ="";
      #endregion Vacia

			#region Bloquea

			this.cmbBodega.Enabled = false;
			this.cmbLocal.Enabled = false;
			this.cmbEnvia.Enabled = false;
			this.cmbRecepta.Enabled = false;
			this.txtDesde.Enabled = false;
			this.txtHasta.Enabled = false;
			this.dtRecibe.Enabled = false;
			this.txtNota.Enabled = false;
			this.txtNotarecibe.Enabled = false;
			this.cmbCodigo.Enabled = false;


			if (miAcceso.Nuevo)this.btnNuevo.Enabled = true;
			if (miAcceso.Buscar)this.btnConsultar.Enabled = true;
			if (miAcceso.Editar) this.btnEditar.Enabled = false;
			this.btnGuardar.Enabled = false;
			this.btnCancelar.Enabled = false;
			this.btnAnular.Enabled = false;

			#endregion Bloquea


		}

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);
	
		}

		private void btnNuevo_Click(object sender, System.EventArgs e)
		{
			this.btnCancelar_Click (sender, e);
			FuncionesProcedimientos.RetornaFechaServidor(this.dtFecha, cdsSeteoF, false);

			#region Bloquea

			this.cmbBodega.Enabled = true;
			this.cmbLocal.Enabled = true;
			this.cmbEnvia.Enabled = true;
			this.cmbRecepta.Enabled = true;
			this.txtDesde.Enabled = true;
			this.txtHasta.Enabled = true;
			this.txtNota.Enabled = true;
			this.cmbCodigo.Enabled = true;

			this.txtEstado.Value =0;
			this.lblEstado.Text = "PENDIENTE";

			#endregion Bloquea
      
			this.btnNuevo.Enabled = false;
			this.btnConsultar.Enabled = false;
			this.btnEditar.Enabled = false;
			this.btnGuardar.Enabled = true;
			this.btnCancelar.Enabled = true;	

			bNuevo = true;
			bEdicion = true;

			this.cmbBodega.Focus();	


		}

		private void cmbBodega_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter) this.cmbEnvia.Focus();
		}

		private void cmbEnvia_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter) this.cmbLocal.Focus();
		}

		private void cmbLocal_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter) this.cmbRecepta.Focus();
		}

		private void cmbRecepta_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter) this.txtDesde.Focus();
		}

		private void txtDesde_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter) this.txtHasta.Focus();
		}

		private void txtHasta_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter) this.txtNota.Focus();
		}

		private void txtNota_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter) this.btnGuardar.Focus();
		}

		private void dtRecibe_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter) this.txtNotarecibe.Focus();
		}

		private void txtNotarecibe_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter) this.btnGuardar.Focus();
		}

		private void txtDesde_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
		
		}

		private void txtDesde_Enter(object sender, System.EventArgs e)
		{
			this.txtDesde.SelectAll();
		}

		private void txtHasta_EnabledChanged(object sender, System.EventArgs e)
		{
		
		}

		private void txtHasta_Enter(object sender, System.EventArgs e)
		{
			this.txtHasta.SelectAll();
		}

		private void btnGuardar_Click(object sender, System.EventArgs e)
		{
			
			#region Validacion			
			if (!Validacion.vbFechaEnDocumentos(this.dtFecha, "Seleccione la Fecha", false, "", cdsSeteoF)) return;
			if (!Validacion.vbComboVacio(this.cmbBodega, "Seleccione la Bodega")) return;
			if (!Validacion.vbComboVacio(this.cmbEnvia, "Seleccione la Persona que envia")) return;
			if (!Validacion.vbComboVacio(this.cmbLocal, "Seleccione el Local de Destino ")) return;
			if (!Validacion.vbComboVacio(this.cmbRecepta, "Seleccione la Persona que recibe ")) return;
			if (!Validacion.vbComboVacio(this.cmbCodigo, "Seleccione el Codigo ")) return;

			if (!Validacion.vbCampoEnteroVacio (this.txtDesde,"Ingrese la numeración Inicial ",1,10000)) return;	
			if (!Validacion.vbCampoEnteroVacio (this.txtHasta,"Ingrese la numeración Final ",1,10000)) return;
			if ((int) this.txtidRecibosManuales.Value == 0)
			{
				string sResp = Funcion.sEscalarSQL(cdsSeteoF, string.Format(" Exec [ValidaRecibo] {0}, {1}, {2}",
					(int) this.cmbCodigo.Value, (int) this.txtDesde.Value, (int) this.txtHasta.Value ));	

				if (sResp.Length > 0)
				{
					MessageBox.Show(string.Format(" '{0}'.", sResp ), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					return;
				}

			}
			if (!Validacion.vbTexto(this.txtNota, 10, 100, "Nota")) return;

			if ( (int) this.txtEstado.Value > 0)
			{
			
        if (!Validacion.vbFechaEnDocumentos(this.dtRecibe, "Seleccione la Fecha de Recepción", false, "", cdsSeteoF)) return;

				if (Convert.ToDateTime(this.dtFecha.Value) > Convert.ToDateTime(this.dtRecibe.Value))
				{
					MessageBox.Show("La Fecha Inicial no puede ser mayor a la Fecha de Recepeción", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					this.dtRecibe.Focus();
					return;
				}
				if (!Validacion.vbTexto(this.txtNotarecibe, 10, 100, "Nota")) return;
			}

			DateTime dtFechaRecepta = DateTime.Parse("01/01/2000");
			if (this.dtRecibe.Value != System.DBNull.Value)dtFechaRecepta=(DateTime)this.dtRecibe.Value;
	
			#endregion Validacion
			if (bNuevo || (int) this.txtEstado.Value == 0)
			{
				if (DialogResult.Yes == MessageBox.Show("¿Esta Seguro De Guardar y Enviar Los Recibos ?, ", "Point Technology", MessageBoxButtons.YesNo, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2))
				{
					this.txtEstado.Value = 1;
				}
				else
				{
					this.txtEstado.Value = 0;
					return;
				}
			}
			if ( (int) this.txtEstado.Value == 1 && (int) this.txtidRecibosManuales.Value >0  )
			{
				if (DialogResult.Yes == MessageBox.Show("¿Esta Seguro De Recibir los Recibos ?, ", "Point Technology", MessageBoxButtons.YesNo, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2))
				{
					this.txtEstado.Value = 2;
				}
				else
				{
					return;
				}
			}

			
				#region Guardar
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

						#region Numeracion
						if (bNuevo)
						{
							string ssSQL = string.Format("Exec NumeracionLocales 75, 0, 0");
							oCmdActualiza.CommandText = ssSQL;
							this.txtNumero.Text = oCmdActualiza.ExecuteScalar().ToString();						
						}
						#endregion Numeracion

						#region Graba Maestro
						int idBodega = 0;
						int idEnvia = 0;
						int idLocal = 0;
						int idRecepta = 0;
						if (this.cmbBodega.ActiveRow != null) idBodega = (int) this.cmbBodega.Value;
						if(this.cmbEnvia.ActiveRow != null) idEnvia = (int) this.cmbEnvia.Value;
						if (this.cmbLocal.ActiveRow != null) idLocal = (int) this.cmbLocal.Value;
						if(this.cmbRecepta.ActiveRow != null) idRecepta = (int) this.cmbRecepta.Value;

						string  sSQL1 = string.Format("Exec [GrabaReciboManuales] {0}, '{1}', '{2}', {3}, {4}, {5}, {6}, {7}, {8}, {9}, '{10}', '{11}', '{12}', {13}" ,
							(int)this.txtidRecibosManuales.Value,
							this.txtNumero.Text,
							Convert.ToDateTime(this.dtFecha.Value).ToString("yyyyMMdd"), 
							idBodega ,
							idEnvia ,
							idLocal ,
							idRecepta ,
							this.cmbCodigo.Value,
							(int) this.txtDesde.Value,
							(int) this.txtHasta.Value,
							Convert.ToDateTime(dtFechaRecepta).ToString("yyyyMMdd"), 
							this.txtNota.Text,
							this.txtNotarecibe.Text,
							(int) this.txtEstado.Value
							);
						oCmdActualiza.CommandText = sSQL1;
						this.txtidRecibosManuales.Value = (int)oCmdActualiza.ExecuteScalar();
						#endregion Graba Maestro					

						oTransaction.Commit();

						Cursor = Cursors.Default;	

						MessageBox.Show("Registro Guardado Correctamente", "Información", MessageBoxButtons.OK, MessageBoxIcon.Information);

						this.lblEstado.Text ="ENVIADO";

						#region Controles
					
						this.cmbBodega.Enabled = false;
						this.cmbLocal.Enabled = false;
						this.cmbEnvia.Enabled = false;
						this.cmbRecepta.Enabled = false;
						this.txtDesde.Enabled = false;
						this.txtHasta.Enabled = false;
						this.cmbCodigo.Enabled = false;
						this.dtRecibe.Enabled = false;
						this.txtNota.Enabled = false;
						this.txtNotarecibe.Enabled = false;


						if (miAcceso.Nuevo)this.btnNuevo.Enabled = true;
						if (miAcceso.Buscar)this.btnConsultar.Enabled = true;
						if (miAcceso.Editar)this.btnEditar.Enabled = false;

						this.btnGuardar.Enabled = false;
						this.btnCancelar.Enabled = true;
            this.Consulta((int)this.txtidRecibosManuales.Value);
						//					this.Consulta((int)this.txtidGastosviajes.Value);
						bNuevo = false;
						bEdicion = false;
						#endregion Controles			

					}
					catch(Exception ex)
					{
						try
						{ 
							oTransaction.Rollback();

							MessageBox.Show(string.Format("Commit Exception Type : {0} {1}", ex.GetType(), ex.Message), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
						}
						catch (Exception ex2)
						{
							MessageBox.Show(string.Format("Rollback Exception Type: {0} {1}", ex2.GetType(), ex2.Message), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);							
						}
					}
					finally
					{
						oConexion.Close();
					}
				}
				
				#endregion Guardar
			}
	

		private void btnEditar_Click(object sender, System.EventArgs e)
		{
			Consulta ((int)this.txtidRecibosManuales.Value);
			
			if ((int)this.txtEstado.Value == 3)
			{
				MessageBox.Show("El documento esta Anulado", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return;
			}


			if (!miAcceso.Editar)
			{
				MessageBox.Show("No tiene acceso a Editar", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return;
			}

			if ((int)this.txtEstado.Value == 0)
			{
				this.cmbBodega.Enabled = true;
				this.cmbLocal.Enabled = true;
				this.cmbEnvia.Enabled = true;
				this.cmbRecepta.Enabled = true;
				this.cmbCodigo.Enabled = true;
				this.txtDesde.Enabled = true;
				this.txtHasta.Enabled = true;
//			this.dtRecibe.Enabled = true;
				this.txtNota.Enabled = true;	
			}

			if ((int)this.txtEstado.Value == 1)
			{
				this.dtRecibe.Enabled = true;
				this.txtNotarecibe.Enabled = true;	
			}	
			this.btnNuevo.Enabled = false;
			this.btnConsultar.Enabled = false;
			this.btnEditar.Enabled = false;			
			this.btnGuardar.Enabled = true;			
			this.btnCancelar.Enabled = true;
			this.btnAnular.Enabled= false;

			bNuevo = false;
			bEdicion = true;
		}

		private void btnConsultar_Click(object sender, System.EventArgs e)
		{
			this.btnCancelar_Click(sender, e);

			using ( frmConsultaGeneralReciboManuales Busqueda = new frmConsultaGeneralReciboManuales())
			{
				if (DialogResult.OK == Busqueda.ShowDialog())
				{
					this.Consulta((int)Busqueda.uGridResumen.ActiveRow.Cells["idReciboManuales"].Value);
				}
			}
		}

		private void Consulta (int Id)
		{			
			string sSQL = string.Format("Exec [ConsultaIndividualReciboManuales] {0}", Id);
			SqlDataReader dr = Funcion.rEscalarSQL(cdsSeteoF, sSQL, true);
			dr.Read();
			if (dr.HasRows)
			{
		  this.txtidRecibosManuales.Value = Id;
			if (dr.GetValue(1) != System.DBNull.Value) this.txtNumero.Text = dr.GetString(1); 
			if (dr.GetValue(2) != System.DBNull.Value) this.dtFecha.Value = dr.GetDateTime(2);
			if (dr.GetValue(3) != System.DBNull.Value) this.cmbBodega.Value = dr.GetInt32(3);
			if (dr.GetValue(4) != System.DBNull.Value) this.cmbEnvia.Value = dr.GetInt32(4);
			if (dr.GetValue(5) != System.DBNull.Value) this.cmbLocal.Value = dr.GetInt32(5);
			if (dr.GetValue(6) != System.DBNull.Value) this.cmbRecepta.Value = dr.GetInt32(6);
   		if (dr.GetValue(7) != System.DBNull.Value) this.cmbCodigo.Value = dr.GetString(7);
			if (dr.GetValue(8) != System.DBNull.Value) this.txtDesde.Value = dr.GetInt32(8);
			if (dr.GetValue(9) != System.DBNull.Value) this.txtHasta.Value = dr.GetInt32(9);
			if (dr.GetValue(10) != System.DBNull.Value) this.dtRecibe.Value = dr.GetDateTime(10);
			if (dr.GetValue(11) != System.DBNull.Value) this.txtNota.Text = dr.GetString(11);
			if (dr.GetValue(12) != System.DBNull.Value) this.txtNotarecibe.Text = dr.GetString(12);
			if (dr.GetValue(13) != System.DBNull.Value) this.txtEstado.Value = dr.GetInt32(13);
			}
			dr.Close();

		  this.cmbEnvia.DataSource = Funcion.dvProcedimiento(cdsSeteoF, string.Format("Exec RetornaNombres {0}", MenuLatinium.IdUsuario));

			this.btnGuardar.Enabled = false;
			this.btnCancelar.Enabled = true;

			if (miAcceso.Nuevo) this.btnNuevo.Enabled = true;
			if (miAcceso.Buscar) this.btnConsultar.Enabled = true;
			if (miAcceso.Editar)
				if ((int)this.txtEstado.Value == 0 || (int)this.txtEstado.Value == 1  ) 
					this.btnEditar.Enabled = true;

			if ((int)this.txtEstado.Value == 0 || (int)this.txtEstado.Value == 1) // PENDIENTE - AUTORIZADO
			{
				if (miAcceso.Anular)this.btnAnular.Enabled = true;
			}

			if ((int)this.txtEstado.Value == 2) // PENDIENTE - AUTORIZADO
			{
				if (miAcceso.BAnularProcesado)this.btnAnular.Enabled = true;
			}

//			if ((int)this.txtEstado.Value == 2) // PENDIENTE - AUTORIZADO
//			{
//				if (miAcceso.Anular)this.btnAnular.Enabled = true;
//			}

			if ((int)this.txtEstado.Value == 0) this.lblEstado.Text ="PENDIENTE";
			if ((int)this.txtEstado.Value == 1) this.lblEstado.Text ="ENVIADO";
			if ((int)this.txtEstado.Value == 2) 	this.lblEstado.Text ="RECIBIDO";
			if ((int)this.txtEstado.Value == 3) 	this.lblEstado.Text ="ANULADO";
			if ((int) this.txtEstado.Value == 1)
			{
				this.btnEditar.Text ="Recibir";
			}
			else
			{
				this.btnEditar.Text ="Editar";
			}
	
		}

		private void btnAnular_Click(object sender, System.EventArgs e)
		{
			this.Consulta ((int) this.txtidRecibosManuales.Value);
			if ((int)this.txtEstado.Value == 3)
			{
				MessageBox.Show("El documento esta Anulado", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return;
			}

			if((int) this.txtEstado.Value < 2)
			{
				if (!miAcceso.Anular)
				{
					MessageBox.Show("No tiene acceso Anular", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					return;
				}
			}
			else
			{
				if (!miAcceso.BAnularProcesado)
				{
					MessageBox.Show("No tiene acceso Anular Recibos Receptados", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					return;
				}
			//	BAnularProcesado 487
			}

			if (DialogResult.Yes == MessageBox.Show("¿Esta Seguro De Anular ?, ", "Point Technology", MessageBoxButtons.YesNo, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2))
			{
				using (frmNotas Nota = new frmNotas( ))
				{
					if (DialogResult.OK == Nota.ShowDialog())
					{
						string notas = Nota.txtNotas.Text;

						string sSQLAsiEs = string.Format("update ReciboManuales set Estado = 3, Nota = Nota + '{0}'  where idReciboManuales = {1}", notas, (int)this.txtidRecibosManuales.Value);			
						Funcion.EjecutaSQL(cdsSeteoF, sSQLAsiEs);

						this.Consulta ((int) this.txtidRecibosManuales.Value);
					}
				}
			}
		}

		private void frmRecibosManuales_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (!bEdicion)
			{
				if (e.KeyCode == Keys.N) if (miAcceso.Nuevo) this.btnNuevo_Click(sender, e);
				if (e.KeyCode == Keys.B) if (miAcceso.Buscar) this.btnConsultar_Click(sender, e);
				if ((int)this.txtidRecibosManuales.Value > 0)
				{
					if (e.KeyCode == Keys.E) if (this.btnEditar.Enabled) this.btnEditar_Click(sender, e);					
					if (e.KeyCode == Keys.F5) this.Consulta((int)this.txtidRecibosManuales.Value);					
				}
				if (e.KeyCode == Keys.Escape) this.Close();				
			}
			else
			{
				if (e.KeyCode == Keys.Escape) 
					if (DialogResult.Yes == MessageBox.Show("Esta seguro de Cancelar la Edición", "Point Technology", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2))
						this.btnCancelar_Click(sender, e);

				if (e.KeyCode == Keys.F12) this.btnGuardar_Click(sender, e);				
			}
		}

		private void btnSalir_Click(object sender, System.EventArgs e)
		{
			this.Close();
		}

		private void txtCodigo_ValueChanged(object sender, System.EventArgs e)
		{
		
		}

	}
}
