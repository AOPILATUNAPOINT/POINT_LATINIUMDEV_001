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

namespace Latinium
{
	/// <summary>
	/// Descripción breve de frmLineaCreditoDestinoAnadir.
	/// </summary>
	public class frmLineaCreditoDestinoAnadir : System.Windows.Forms.Form
	{
		private System.Windows.Forms.GroupBox groupBox2;
		private System.Windows.Forms.Label label3;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbDestino;
		private System.Windows.Forms.Label label4;
		public System.Windows.Forms.Label label1;
		private System.Windows.Forms.TabControl tabRequerimiento;
		private System.Windows.Forms.TabPage tabPage3;
		private Infragistics.Win.UltraWinGrid.UltraGrid ugdDocumento;
		private System.Windows.Forms.Button btnSalir;
		private System.Windows.Forms.Button btnGuardar;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtMonto;
		private C1.Data.C1DataSet cdsSeteoF;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;
		bool bNuevo = false;
		bool bVer = false;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource udsDocumento;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource udsDestino;
		string txtDirectorioarchivo = "";
		int idLineaCreditoDestino = 0;
		int iTipo = 0;
		bool bAnadir = false;
		private System.Windows.Forms.GroupBox groupBox3;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor dtTiempo;
		bool bEliminar = false;

		public frmLineaCreditoDestinoAnadir(int IdLineaCreditoDestino, int ITipo)
		{

			iTipo = ITipo;
			idLineaCreditoDestino = IdLineaCreditoDestino;

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
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idLineaCreditoDestinoTipo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Descripcion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idLineaCreditoDestinoTipo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Descripcion");
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idLineaCreditoDestinoDoc");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idLineaCreditoDestino");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("UrlNube");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaIngreso");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Usuario");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Estacion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idLineaCreditoTransaccion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Estado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn11 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("sUrlNume");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn12 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("sCarpeta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn13 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("txtDirectorioarchivo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn14 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn15 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaEdicion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn16 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cargado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn17 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("btnDocumento", 0);
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmLineaCreditoDestinoAnadir));
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn18 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("btnEliminar", 1);
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn19 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("btnVer", 2);
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance11 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance12 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance13 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance14 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn3 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idLineaCreditoDestinoDoc");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn4 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idLineaCreditoDestino");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn5 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("UrlNube");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn6 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("FechaIngreso");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn7 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Usuario");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn8 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Estacion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn9 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idLineaCreditoTransaccion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn10 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Estado");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn11 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("sUrlNume");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn12 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("sCarpeta");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn13 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("txtDirectorioarchivo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn14 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn15 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("FechaEdicion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn16 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Cargado");
			this.groupBox2 = new System.Windows.Forms.GroupBox();
			this.dtTiempo = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtMonto = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label3 = new System.Windows.Forms.Label();
			this.cmbDestino = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.udsDestino = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.label4 = new System.Windows.Forms.Label();
			this.label1 = new System.Windows.Forms.Label();
			this.tabRequerimiento = new System.Windows.Forms.TabControl();
			this.tabPage3 = new System.Windows.Forms.TabPage();
			this.ugdDocumento = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.udsDocumento = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.btnSalir = new System.Windows.Forms.Button();
			this.btnGuardar = new System.Windows.Forms.Button();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.groupBox3 = new System.Windows.Forms.GroupBox();
			this.groupBox2.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.dtTiempo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtMonto)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbDestino)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.udsDestino)).BeginInit();
			this.tabRequerimiento.SuspendLayout();
			this.tabPage3.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.ugdDocumento)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.udsDocumento)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			this.SuspendLayout();
			// 
			// groupBox2
			// 
			this.groupBox2.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.groupBox2.Controls.Add(this.dtTiempo);
			this.groupBox2.Controls.Add(this.txtMonto);
			this.groupBox2.Controls.Add(this.label3);
			this.groupBox2.Controls.Add(this.cmbDestino);
			this.groupBox2.Controls.Add(this.label4);
			this.groupBox2.Controls.Add(this.label1);
			this.groupBox2.Location = new System.Drawing.Point(16, 16);
			this.groupBox2.Name = "groupBox2";
			this.groupBox2.Size = new System.Drawing.Size(624, 96);
			this.groupBox2.TabIndex = 1034;
			this.groupBox2.TabStop = false;
			this.groupBox2.Text = "Datos principales";
			this.groupBox2.Enter += new System.EventHandler(this.groupBox2_Enter);
			// 
			// dtTiempo
			// 
			appearance1.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtTiempo.Appearance = appearance1;
			this.dtTiempo.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.dtTiempo.FormatString = "";
			this.dtTiempo.Location = new System.Drawing.Point(136, 64);
			this.dtTiempo.MaskInput = "nn";
			this.dtTiempo.Name = "dtTiempo";
			this.dtTiempo.PromptChar = ' ';
			this.dtTiempo.Size = new System.Drawing.Size(152, 21);
			this.dtTiempo.SpinButtonDisplayStyle = Infragistics.Win.ButtonDisplayStyle.Always;
			this.dtTiempo.TabIndex = 1007;
			// 
			// txtMonto
			// 
			appearance2.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtMonto.Appearance = appearance2;
			this.txtMonto.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.txtMonto.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtMonto.FormatString = "#,##0.00";
			this.txtMonto.Location = new System.Drawing.Point(440, 24);
			this.txtMonto.Name = "txtMonto";
			this.txtMonto.NullText = "0.00";
			this.txtMonto.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtMonto.PromptChar = ' ';
			this.txtMonto.Size = new System.Drawing.Size(152, 19);
			this.txtMonto.TabIndex = 1006;
			// 
			// label3
			// 
			this.label3.Location = new System.Drawing.Point(8, 64);
			this.label3.Name = "label3";
			this.label3.Size = new System.Drawing.Size(104, 16);
			this.label3.TabIndex = 1003;
			this.label3.Text = "* Tiempo (meses):";
			this.label3.Click += new System.EventHandler(this.label3_Click);
			// 
			// cmbDestino
			// 
			appearance3.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbDestino.Appearance = appearance3;
			this.cmbDestino.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.cmbDestino.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbDestino.DataSource = this.udsDestino;
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn2.Header.Caption = "Descripción";
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridColumn2.Width = 165;
			ultraGridBand1.Columns.AddRange(new object[] {
																										 ultraGridColumn1,
																										 ultraGridColumn2});
			this.cmbDestino.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.cmbDestino.DisplayMember = "Descripcion";
			this.cmbDestino.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbDestino.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbDestino.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbDestino.Location = new System.Drawing.Point(136, 24);
			this.cmbDestino.Name = "cmbDestino";
			this.cmbDestino.Size = new System.Drawing.Size(152, 21);
			this.cmbDestino.TabIndex = 817;
			this.cmbDestino.ValueMember = "idLineaCreditoDestinoTipo";
			// 
			// udsDestino
			// 
			ultraDataColumn1.DataType = typeof(int);
			ultraDataColumn1.ReadOnly = Infragistics.Win.DefaultableBoolean.True;
			ultraDataColumn2.DataType = typeof(System.UInt32);
			this.udsDestino.Band.Columns.AddRange(new object[] {
																													 ultraDataColumn1,
																													 ultraDataColumn2});
			// 
			// label4
			// 
			this.label4.Location = new System.Drawing.Point(320, 24);
			this.label4.Name = "label4";
			this.label4.TabIndex = 814;
			this.label4.Text = "* Monto ($):";
			// 
			// label1
			// 
			this.label1.Location = new System.Drawing.Point(8, 24);
			this.label1.Name = "label1";
			this.label1.TabIndex = 812;
			this.label1.Text = "* Destino:";
			// 
			// tabRequerimiento
			// 
			this.tabRequerimiento.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.tabRequerimiento.Controls.Add(this.tabPage3);
			this.tabRequerimiento.Location = new System.Drawing.Point(16, 144);
			this.tabRequerimiento.Name = "tabRequerimiento";
			this.tabRequerimiento.SelectedIndex = 0;
			this.tabRequerimiento.Size = new System.Drawing.Size(720, 264);
			this.tabRequerimiento.TabIndex = 1035;
			// 
			// tabPage3
			// 
			this.tabPage3.Controls.Add(this.ugdDocumento);
			this.tabPage3.Location = new System.Drawing.Point(4, 22);
			this.tabPage3.Name = "tabPage3";
			this.tabPage3.Size = new System.Drawing.Size(712, 238);
			this.tabPage3.TabIndex = 2;
			this.tabPage3.Text = "Documento *";
			this.tabPage3.Visible = false;
			// 
			// ugdDocumento
			// 
			this.ugdDocumento.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.ugdDocumento.Cursor = System.Windows.Forms.Cursors.Default;
			this.ugdDocumento.DataSource = this.udsDocumento;
			appearance4.BackColor = System.Drawing.Color.White;
			this.ugdDocumento.DisplayLayout.Appearance = appearance4;
			this.ugdDocumento.DisplayLayout.AutoFitColumns = true;
			ultraGridColumn3.Header.VisiblePosition = 2;
			ultraGridColumn3.Hidden = true;
			ultraGridColumn3.Width = 61;
			ultraGridColumn4.Header.VisiblePosition = 3;
			ultraGridColumn4.Hidden = true;
			ultraGridColumn4.Width = 64;
			ultraGridColumn5.Header.VisiblePosition = 4;
			ultraGridColumn5.Hidden = true;
			ultraGridColumn5.Width = 58;
			ultraGridColumn6.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
			ultraGridColumn6.Format = "dd/MM/yyyy hh:mm:ss";
			ultraGridColumn6.Header.Caption = "Fecha";
			ultraGridColumn6.Header.VisiblePosition = 14;
			ultraGridColumn6.Width = 289;
			ultraGridColumn7.Header.VisiblePosition = 5;
			ultraGridColumn7.Hidden = true;
			ultraGridColumn7.Width = 127;
			ultraGridColumn8.Header.VisiblePosition = 6;
			ultraGridColumn8.Hidden = true;
			ultraGridColumn8.Width = 106;
			ultraGridColumn9.Header.VisiblePosition = 7;
			ultraGridColumn9.Hidden = true;
			ultraGridColumn9.Width = 132;
			ultraGridColumn10.Header.VisiblePosition = 8;
			ultraGridColumn10.Hidden = true;
			ultraGridColumn10.Width = 37;
			ultraGridColumn11.Header.VisiblePosition = 9;
			ultraGridColumn11.Hidden = true;
			ultraGridColumn11.Width = 151;
			ultraGridColumn12.Header.VisiblePosition = 10;
			ultraGridColumn12.Hidden = true;
			ultraGridColumn12.Width = 68;
			ultraGridColumn13.Header.VisiblePosition = 11;
			ultraGridColumn13.Hidden = true;
			ultraGridColumn13.Width = 70;
			ultraGridColumn14.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
			ultraGridColumn14.Header.VisiblePosition = 13;
			ultraGridColumn14.Width = 304;
			ultraGridColumn15.Header.VisiblePosition = 15;
			ultraGridColumn15.Hidden = true;
			ultraGridColumn15.Width = 53;
			ultraGridColumn16.Header.VisiblePosition = 16;
			ultraGridColumn16.Hidden = true;
			ultraGridColumn16.Width = 50;
			appearance5.Image = ((object)(resources.GetObject("appearance5.Image")));
			ultraGridColumn17.CellAppearance = appearance5;
			appearance6.Image = ((object)(resources.GetObject("appearance6.Image")));
			ultraGridColumn17.CellButtonAppearance = appearance6;
			ultraGridColumn17.Header.Caption = "...";
			ultraGridColumn17.Header.VisiblePosition = 0;
			ultraGridColumn17.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.Button;
			ultraGridColumn17.Width = 23;
			appearance7.Image = ((object)(resources.GetObject("appearance7.Image")));
			ultraGridColumn18.CellAppearance = appearance7;
			appearance8.Image = ((object)(resources.GetObject("appearance8.Image")));
			ultraGridColumn18.CellButtonAppearance = appearance8;
			ultraGridColumn18.Header.Caption = "...";
			ultraGridColumn18.Header.VisiblePosition = 1;
			ultraGridColumn18.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.Button;
			ultraGridColumn18.Width = 21;
			appearance9.Image = ((object)(resources.GetObject("appearance9.Image")));
			ultraGridColumn19.CellAppearance = appearance9;
			appearance10.Image = ((object)(resources.GetObject("appearance10.Image")));
			ultraGridColumn19.CellButtonAppearance = appearance10;
			ultraGridColumn19.Header.Caption = "...";
			ultraGridColumn19.Header.VisiblePosition = 12;
			ultraGridColumn19.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.Button;
			ultraGridColumn19.Width = 22;
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
																										 ultraGridColumn16,
																										 ultraGridColumn17,
																										 ultraGridColumn18,
																										 ultraGridColumn19});
			this.ugdDocumento.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			this.ugdDocumento.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.TemplateOnBottom;
			this.ugdDocumento.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.True;
			this.ugdDocumento.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance11.BackColor = System.Drawing.Color.Transparent;
			this.ugdDocumento.DisplayLayout.Override.CardAreaAppearance = appearance11;
			appearance12.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance12.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance12.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance12.FontData.BoldAsString = "True";
			appearance12.FontData.Name = "Arial";
			appearance12.FontData.SizeInPoints = 10F;
			appearance12.ForeColor = System.Drawing.Color.White;
			appearance12.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.ugdDocumento.DisplayLayout.Override.HeaderAppearance = appearance12;
			appearance13.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance13.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance13.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ugdDocumento.DisplayLayout.Override.RowSelectorAppearance = appearance13;
			appearance14.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance14.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance14.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ugdDocumento.DisplayLayout.Override.SelectedRowAppearance = appearance14;
			this.ugdDocumento.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.ugdDocumento.Location = new System.Drawing.Point(16, 16);
			this.ugdDocumento.Name = "ugdDocumento";
			this.ugdDocumento.Size = new System.Drawing.Size(680, 200);
			this.ugdDocumento.TabIndex = 810;
			this.ugdDocumento.ClickCellButton += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.ugdDocumento_ClickCellButton);
			this.ugdDocumento.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.ugdDocumento_InitializeLayout);
			// 
			// udsDocumento
			// 
			ultraDataColumn3.DataType = typeof(int);
			ultraDataColumn3.ReadOnly = Infragistics.Win.DefaultableBoolean.True;
			ultraDataColumn4.DataType = typeof(System.UInt32);
			ultraDataColumn6.DataType = typeof(System.DateTime);
			ultraDataColumn9.DataType = typeof(int);
			ultraDataColumn10.DataType = typeof(bool);
			ultraDataColumn16.DataType = typeof(bool);
			this.udsDocumento.Band.Columns.AddRange(new object[] {
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
																														 ultraDataColumn16});
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
			this.btnSalir.Location = new System.Drawing.Point(672, 64);
			this.btnSalir.Name = "btnSalir";
			this.btnSalir.Size = new System.Drawing.Size(80, 23);
			this.btnSalir.TabIndex = 1037;
			this.btnSalir.Text = "Salir";
			this.btnSalir.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnSalir.Click += new System.EventHandler(this.btnSalir_Click);
			// 
			// btnGuardar
			// 
			this.btnGuardar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btnGuardar.Image = ((System.Drawing.Image)(resources.GetObject("btnGuardar.Image")));
			this.btnGuardar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnGuardar.Location = new System.Drawing.Point(672, 24);
			this.btnGuardar.Name = "btnGuardar";
			this.btnGuardar.Size = new System.Drawing.Size(82, 23);
			this.btnGuardar.TabIndex = 1036;
			this.btnGuardar.Text = "&Guardar";
			this.btnGuardar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnGuardar.Click += new System.EventHandler(this.btnGuardar_Click);
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
			// 
			// groupBox3
			// 
			this.groupBox3.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.groupBox3.Location = new System.Drawing.Point(-1048, 120);
			this.groupBox3.Name = "groupBox3";
			this.groupBox3.Size = new System.Drawing.Size(2840, 8);
			this.groupBox3.TabIndex = 1044;
			this.groupBox3.TabStop = false;
			// 
			// frmLineaCreditoDestinoAnadir
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(760, 414);
			this.Controls.Add(this.groupBox3);
			this.Controls.Add(this.btnSalir);
			this.Controls.Add(this.btnGuardar);
			this.Controls.Add(this.tabRequerimiento);
			this.Controls.Add(this.groupBox2);
			this.Name = "frmLineaCreditoDestinoAnadir";
			this.Text = "Añadir destino";
			this.Load += new System.EventHandler(this.frmLineaCreditoDestinoAnadir_Load);
			this.groupBox2.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.dtTiempo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtMonto)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbDestino)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.udsDestino)).EndInit();
			this.tabRequerimiento.ResumeLayout(false);
			this.tabPage3.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.ugdDocumento)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.udsDocumento)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void frmLineaCreditoDestinoAnadir_Load(object sender, System.EventArgs e)
		{
			if(iTipo == frmLineaCreditoCrear.iNuevo)
			{
				bNuevo = true;
				this.Text = "Nuevo destino";
			}
			if(iTipo == frmLineaCreditoCrear.iVer)
			{
				bVer = true;
				this.Text = "Ver destino";
			}

			if(iTipo == frmLineaCreditoCrear.iAnadir)
			{
				bAnadir = true;
				this.Text = "Añadir destino";
			}

			if(iTipo == frmLineaCreditoCrear.iEliminar)
			{
				bEliminar = true;
				this.Text = "Eliminar destino";
			}

			//Listado destino 
			LineaCreditoDestinoTipoCon();
			//Listado documentos destino
			LineaCreditoDestinoDocCon();
			//Listado destino model
			LineaCreditoDestinoModel();
			ControlVista();
		}

		private void ControlVista()
		{
			if(bVer)
			{
					this.cmbDestino.Enabled = false;
					this.txtMonto.Enabled = false;
					this.dtTiempo.Enabled = false;
					this.btnGuardar.Enabled = false;
			}

			if(bEliminar)
			{
				this.cmbDestino.Enabled = false;
				this.txtMonto.Enabled = false;
				this.dtTiempo.Enabled = false;
				this.btnGuardar.Text = "Eliminar";
			}

		}

		private void LineaCreditoDestinoModel()
		{
			if(idLineaCreditoDestino != 0)
			{
				//Consultar valor para ser llenados
				this.cmbDestino.Value = Funcion.iEscalarSQL(cdsSeteoF, String.Format("EXEC LineaCreditoDestinoModel '{0}', {1}", "idLineaCreditoDestinoTipo", idLineaCreditoDestino));
				this.dtTiempo.Value = Funcion.iEscalarSQL(cdsSeteoF, String.Format("EXEC LineaCreditoDestinoModel '{0}', {1}", "Tiempo", idLineaCreditoDestino));
				this.txtMonto.Value = Funcion.decEscalarSQL(cdsSeteoF, String.Format("EXEC LineaCreditoDestinoModel '{0}', {1}", "Monto", idLineaCreditoDestino));
			}
		}

		private void LineaCreditoDestinoDocCon()
		{
				this.ugdDocumento.DataSource = Funcion.dvProcedimiento(cdsSeteoF, String.Format("EXEC LineaCreditoDestinoDocCon {0}, {1}", 
				frmLineaCreditoCrear.idLineaCreditoTransaccion, idLineaCreditoDestino));
		}

		private void LineaCreditoDestinoTipoCon()
		{
			this.cmbDestino.DataSource = Funcion.dvProcedimiento(cdsSeteoF, String.Format("EXEC LineaCreditoDestinoTipoCon"));
		}

		private void label3_Click(object sender, System.EventArgs e)
		{
		
		}

		private void ugdDocumento_ClickCellButton(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			string txtURL ="";
			long itamano = 0;
			bool bResultado = false;
			string sMensaje = "";
			DataSet dsConfiguracion = new DataSet();
			int idLineaCreditoDestinoDoc = 0;
			string gettxtDirectorioarchivo = "";

			if (e.Cell.Column.ToString() == "btnDocumento" && bNuevo)	
			{
					#region Cargar documento
					OpenFileDialog buscar = new OpenFileDialog ();
					if (buscar.ShowDialog () == DialogResult.OK)
					{
						//Guardar variable dirección del archivo
						txtDirectorioarchivo = buscar.FileName;

						#region Eliminar archivo anterior
						try {idLineaCreditoDestinoDoc = (int)e.Cell.Row.Cells["idLineaCreditoDestinoDoc"].Value; } catch(Exception ex){string sms = ex.Message;};
						try {gettxtDirectorioarchivo = (string)e.Cell.Row.Cells["txtDirectorioarchivo"].Value; } catch(Exception ex){string sms = ex.Message;};

						if(gettxtDirectorioarchivo != txtDirectorioarchivo && txtDirectorioarchivo != "" && idLineaCreditoDestinoDoc != 0)
						{
							//Eliminar anterior seleciconado
							try
							{
								//Si se guardo crear un registro por el nuevo guardado
								bool bCont =Funcion.bEjecucion(frmRequerimiento.sconexionPoint, string.Format("Exec LineaCreditoDestinoDocEliminar {0}", 
								idLineaCreditoDestinoDoc), false);
								if (bCont)
								{
									LineaCreditoDestinoDocCon();
								}
							}
							catch (Exception ex)
							{
								string sms = ex.Message;
								return;						
							}
						}
						#endregion Eliminar archivo anterior

						//Sacar el tamaño del archivo subido
						System.IO.FileInfo info = new System.IO.FileInfo(txtDirectorioarchivo);
						//El lenght está en tamaños bytes
						itamano = info.Length;
						decimal dtamano = Convert.ToDecimal(itamano);
						if(!frmLineaCredito.LineaCreditoConfgVal(cdsSeteoF,frmLineaCredito.iConfiguracionDocumento, dtamano, 1))return;

						//Si se guardo crear un registro por el nuevo guardado
						DataView dvLineaCreditoDestinoDocCrear = new DataView();
						dvLineaCreditoDestinoDocCrear = Funcion.dvProcedimiento(cdsSeteoF, string.Format("Exec LineaCreditoDestinoDocCrear '{0}', {1}, {2}, {3}", 
							txtDirectorioarchivo, 
							frmLineaCreditoCrear.idLineaCreditoTransaccion,
							0,
							0
							));
							  
						DataSet dsLineaCreditoDestinoDocCrear = new DataSet();
						dsLineaCreditoDestinoDocCrear = dvLineaCreditoDestinoDocCrear.Table.DataSet;
						bResultado = bool.Parse(dsLineaCreditoDestinoDocCrear.Tables[0].Rows[0]["bResultado"].ToString());
						sMensaje = dsLineaCreditoDestinoDocCrear.Tables[0].Rows[0]["sMensaje"].ToString();

						if(bResultado == false)
						{
							MessageBox.Show(string.Format(sMensaje), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
							return;
						}
						else
						{
							LineaCreditoDestinoDocCon();
						}
					}
					#endregion Cargar documento
			}

			#region Ver
			if (e.Cell.Column.ToString() == "btnVer")	
			{
				try
				{
					  bool Cargado = false;
					  try{Cargado = (bool)e.Cell.Row.Cells["Cargado"].Value;} catch(Exception ex){string sms = ex.Message;};

						if(!Cargado)
						{
							txtDirectorioarchivo = (string)e.Cell.Row.Cells["txtDirectorioarchivo"].Value;
						}
						else
						{
							txtDirectorioarchivo = (string)e.Cell.Row.Cells["sCarpeta"].Value;
						}
						if (txtDirectorioarchivo.Length > 10 )
						{
							Process.Start(txtDirectorioarchivo);
						}
				}
				catch (Exception ex)
				{
					string sms = ex.Message;
					return;						
				}
			}
			#endregion Ver

			#region Eliminar documento
			if (e.Cell.Column.ToString() == "btnEliminar")	
			{
				try
				{
					int itamanodocumentos = this.ugdDocumento.Rows.Count;
					if(itamanodocumentos == 0)
					{
						return;
					}
					//Si se guardo crear un registro por el nuevo guardado
					idLineaCreditoDestinoDoc = (int)e.Cell.Row.Cells["idLineaCreditoDestinoDoc"].Value;
					bool bCont =Funcion.bEjecucion(frmRequerimiento.sconexionPoint, string.Format("Exec LineaCreditoDestinoDocEliminar {0}", 
						idLineaCreditoDestinoDoc
						), false);
					if (bCont == true)
					{
						// Eliminar el path
						txtDirectorioarchivo = ((string)e.Cell.Row.Cells["txtDirectorioarchivo"].Value);
						txtURL =  Funcion.sEscalarSQL(cdsSeteoF, string.Format(" SELECT RIGHT('{0}', CHARINDEX('{1}', REVERSE('{0}')) - 1) AS Archivo",txtDirectorioarchivo,@"\"));
						//						System.IO.File.Delete(txtDirectorioarchivo) ;	
						MessageBox.Show(string.Format("El archivo '{0}' se eliminó correctamente. ", txtURL.Replace(" ", "_")), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
						//return;
					}
					LineaCreditoDestinoDocCon();
				}
				catch (Exception ex)
				{
					string sms = ex.Message;
					return;						
				}
			}
			#endregion Eliminar documento
		}

		private void btnGuardar_Click(object sender, System.EventArgs e)
		{
			if(bAnadir || bNuevo)
			{
				#region Valida campos  
				if (!Validacion.vbComboVacio(this.cmbDestino, "Ingrese un destino")) return;
				if (!Validacion.vbCampoDecimalVacio(this.txtMonto, "Ingrese el monto", 1, 999999999)) return;
				if (!Validacion.vbCampoEnteroVacio(this.dtTiempo, "Ingrese el tiempo", 1, 999999999)) return;
				#endregion Valida campos  
				#region guardar destino

				string squery = "";
				int idLineaCreditoDestino = 0;
				try
				{
					squery = String.Format("EXEC LineaCreditoDestinoCrear {0},{1},{2},'{3}'",
						frmLineaCreditoCrear.idLineaCreditoTransaccion,
						(int)this.cmbDestino.Value,
						Convert.ToDecimal(this.txtMonto.Value),
						(int)(this.dtTiempo.Value));
					idLineaCreditoDestino = Funcion.iEscalarSQL(cdsSeteoF, squery);
				}
				catch(Exception ex)
				{
					Funcion.LogSistema(cdsSeteoF, this.Name + ": "+ this.Text, string.Format("Commit Exception Type : {0} {1}", ex.GetType(), ex.Message), squery, Funcion.slogtipoerror);
				}

				#endregion guardar destino
				#region actualizar documentos

				int idLineaCreditoDestinoDoc = 0;

				foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.ugdDocumento.Rows.All)	
					if ((int)dr.Cells["idLineaCreditoDestinoDoc"].Value != 0)
					{
						//Guardar el documento
						DataView dvLineaCreditoDestinoDocCrear = new DataView();
						idLineaCreditoDestinoDoc = (int)dr.Cells["idLineaCreditoDestinoDoc"].Value;	
						string squery2 = "";
						try
						{
							squery2 = string.Format("Exec LineaCreditoDestinoDocCrear '{0}', {1}, {2}, {3}", 
								txtDirectorioarchivo, 
								frmLineaCreditoCrear.idLineaCreditoTransaccion,
								idLineaCreditoDestino,
								idLineaCreditoDestinoDoc
								);
							dvLineaCreditoDestinoDocCrear = Funcion.dvProcedimiento(cdsSeteoF, squery2);
						}
						catch(Exception ex)
						{
							Funcion.LogSistema(cdsSeteoF, this.Name + ": "+ this.Text, string.Format("Commit Exception Type : {0} {1}", ex.GetType(), ex.Message), squery2, Funcion.slogtipoerror);
						}
					}
				#endregion actualizar documentos
				MessageBox.Show(string.Format("Proceso Finalizado Correctamente, se creo el registro exitosamente."), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
				this.DialogResult = DialogResult.OK;
			}
			if(bEliminar)
			{
				string squery3 = "";
				try
				{
					squery3 = String.Format("EXEC LineaCreditoDestinoEliminar {0}", idLineaCreditoDestino);
					bool bResultadoEliminar = Funcion.bEscalarSQL(cdsSeteoF, squery3, false);
					if(bResultadoEliminar)
					{
						MessageBox.Show(string.Format("Proceso Finalizado Correctamente, se eliminó el registro exitosamente."), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
						this.DialogResult = DialogResult.OK;
					}
				}
				catch(Exception ex)
				{
					Funcion.LogSistema(cdsSeteoF, this.Name + ": "+ this.Text, string.Format("Commit Exception Type : {0} {1}", ex.GetType(), ex.Message), squery3, Funcion.slogtipoerror);
				}
			}
		}

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);
		}

		private void btnSalir_Click(object sender, System.EventArgs e)
		{
			this.DialogResult = DialogResult.No;
		}

		private void groupBox2_Enter(object sender, System.EventArgs e)
		{
		
		}

		private void ugdDocumento_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
			if(bVer)
			{
				e.Layout.Bands[0].Columns["btnDocumento"].CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
				e.Layout.Bands[0].Columns["btnEliminar"].CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			}

			if(bEliminar)
			{
				e.Layout.Bands[0].Columns["btnDocumento"].CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			}
		}
	}
}
