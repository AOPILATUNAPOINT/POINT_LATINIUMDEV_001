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
	/// Descripción breve de frmReportePresupuestoXLider.
	/// </summary>
	public class frmReportePresupuestoXLider : System.Windows.Forms.Form
	{
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtAnio;
		private Infragistics.Win.Misc.UltraLabel lblPeriodo;
		public Infragistics.Win.UltraWinGrid.UltraCombo cmbResponsable;
		private System.Windows.Forms.Label label1;
		private C1.Data.C1DataSet cdsSeteoF;
		private System.Windows.Forms.Label label2;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource1;
		public Infragistics.Win.UltraWinGrid.UltraCombo cmbCuenta;
		private System.Windows.Forms.Label label3;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource2;
		public Infragistics.Win.UltraWinGrid.UltraCombo cmbBodega;
		private System.Windows.Forms.Label label4;
		public Infragistics.Win.UltraWinGrid.UltraCombo cmbMes;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource3;
		private System.Windows.Forms.Button btnConsultar;
		private Infragistics.Win.UltraWinGrid.UltraGrid uGridInformacion;
		private System.Windows.Forms.Button btnExcel;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public frmReportePresupuestoXLider()
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
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idNomina");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCuenta");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Codigo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn3 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Cuenta");
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCuenta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cuenta");
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand3 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idProyecto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn4 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idProyecto");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn5 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand4 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idMes");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Mes");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn6 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idMes");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn7 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Mes");
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmReportePresupuestoXLider));
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance11 = new Infragistics.Win.Appearance();
			this.txtAnio = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.lblPeriodo = new Infragistics.Win.Misc.UltraLabel();
			this.cmbResponsable = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.label1 = new System.Windows.Forms.Label();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.label2 = new System.Windows.Forms.Label();
			this.ultraDataSource1 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.cmbCuenta = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.cmbBodega = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.ultraDataSource2 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.label3 = new System.Windows.Forms.Label();
			this.cmbMes = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.ultraDataSource3 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.label4 = new System.Windows.Forms.Label();
			this.btnConsultar = new System.Windows.Forms.Button();
			this.uGridInformacion = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.btnExcel = new System.Windows.Forms.Button();
			((System.ComponentModel.ISupportInitialize)(this.txtAnio)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbResponsable)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCuenta)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbBodega)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbMes)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource3)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridInformacion)).BeginInit();
			this.SuspendLayout();
			// 
			// txtAnio
			// 
			appearance1.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtAnio.Appearance = appearance1;
			this.txtAnio.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtAnio.FormatString = "";
			this.txtAnio.Location = new System.Drawing.Point(80, 8);
			this.txtAnio.MaskInput = "nnnn";
			this.txtAnio.Name = "txtAnio";
			this.txtAnio.PromptChar = ' ';
			this.txtAnio.Size = new System.Drawing.Size(96, 21);
			this.txtAnio.SpinButtonDisplayStyle = Infragistics.Win.ButtonDisplayStyle.Always;
			this.txtAnio.TabIndex = 761;
			this.txtAnio.ValueChanged += new System.EventHandler(this.txtAnio_ValueChanged);
			// 
			// lblPeriodo
			// 
			this.lblPeriodo.AutoSize = true;
			this.lblPeriodo.BackColor = System.Drawing.Color.Transparent;
			this.lblPeriodo.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.lblPeriodo.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(50)), ((System.Byte)(50)), ((System.Byte)(50)));
			this.lblPeriodo.Location = new System.Drawing.Point(16, 11);
			this.lblPeriodo.Name = "lblPeriodo";
			this.lblPeriodo.Size = new System.Drawing.Size(42, 15);
			this.lblPeriodo.TabIndex = 760;
			this.lblPeriodo.Text = "Periodo";
			// 
			// cmbResponsable
			// 
			appearance2.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbResponsable.Appearance = appearance2;
			this.cmbResponsable.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbResponsable.Cursor = System.Windows.Forms.Cursors.Default;
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridColumn2.Width = 250;
			ultraGridBand1.Columns.AddRange(new object[] {
																										 ultraGridColumn1,
																										 ultraGridColumn2});
			this.cmbResponsable.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.cmbResponsable.DisplayMember = "Nombre";
			this.cmbResponsable.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbResponsable.Enabled = false;
			this.cmbResponsable.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbResponsable.Location = new System.Drawing.Point(80, 38);
			this.cmbResponsable.Name = "cmbResponsable";
			this.cmbResponsable.Size = new System.Drawing.Size(256, 21);
			this.cmbResponsable.TabIndex = 1001;
			this.cmbResponsable.ValueMember = "idNomina";
			this.cmbResponsable.ValueChanged += new System.EventHandler(this.cmbResponsable_ValueChanged);
			// 
			// label1
			// 
			this.label1.AutoSize = true;
			this.label1.Location = new System.Drawing.Point(16, 40);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(49, 16);
			this.label1.TabIndex = 1000;
			this.label1.Text = "Personal";
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
			// label2
			// 
			this.label2.AutoSize = true;
			this.label2.Location = new System.Drawing.Point(16, 72);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(41, 16);
			this.label2.TabIndex = 1002;
			this.label2.Text = "Cuenta";
			// 
			// ultraDataSource1
			// 
			this.ultraDataSource1.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn1,
																																 ultraDataColumn2,
																																 ultraDataColumn3});
			// 
			// cmbCuenta
			// 
			appearance3.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbCuenta.Appearance = appearance3;
			this.cmbCuenta.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbCuenta.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbCuenta.DataSource = this.ultraDataSource1;
			ultraGridColumn3.Header.VisiblePosition = 0;
			ultraGridColumn3.Hidden = true;
			ultraGridColumn4.Header.VisiblePosition = 1;
			ultraGridColumn5.Header.VisiblePosition = 2;
			ultraGridColumn5.Width = 180;
			ultraGridBand2.Columns.AddRange(new object[] {
																										 ultraGridColumn3,
																										 ultraGridColumn4,
																										 ultraGridColumn5});
			this.cmbCuenta.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			this.cmbCuenta.DisplayMember = "Cuenta";
			this.cmbCuenta.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbCuenta.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbCuenta.Location = new System.Drawing.Point(80, 70);
			this.cmbCuenta.Name = "cmbCuenta";
			this.cmbCuenta.Size = new System.Drawing.Size(256, 21);
			this.cmbCuenta.TabIndex = 1004;
			this.cmbCuenta.ValueMember = "idCuenta";
			this.cmbCuenta.ValueChanged += new System.EventHandler(this.cmbCuenta_ValueChanged);
			// 
			// cmbBodega
			// 
			appearance4.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbBodega.Appearance = appearance4;
			this.cmbBodega.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbBodega.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbBodega.DataSource = this.ultraDataSource2;
			ultraGridColumn6.Header.VisiblePosition = 0;
			ultraGridColumn6.Hidden = true;
			ultraGridColumn7.Header.Caption = "Proyecto";
			ultraGridColumn7.Header.VisiblePosition = 1;
			ultraGridColumn7.Width = 256;
			ultraGridBand3.Columns.AddRange(new object[] {
																										 ultraGridColumn6,
																										 ultraGridColumn7});
			this.cmbBodega.DisplayLayout.BandsSerializer.Add(ultraGridBand3);
			this.cmbBodega.DisplayMember = "Nombre";
			this.cmbBodega.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbBodega.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbBodega.Location = new System.Drawing.Point(80, 104);
			this.cmbBodega.Name = "cmbBodega";
			this.cmbBodega.Size = new System.Drawing.Size(256, 21);
			this.cmbBodega.TabIndex = 1006;
			this.cmbBodega.ValueMember = "idProyecto";
			// 
			// ultraDataSource2
			// 
			ultraDataColumn4.DataType = typeof(int);
			this.ultraDataSource2.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn4,
																																 ultraDataColumn5});
			// 
			// label3
			// 
			this.label3.AutoSize = true;
			this.label3.Location = new System.Drawing.Point(16, 106);
			this.label3.Name = "label3";
			this.label3.Size = new System.Drawing.Size(49, 16);
			this.label3.TabIndex = 1005;
			this.label3.Text = "Proyecto";
			// 
			// cmbMes
			// 
			appearance5.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbMes.Appearance = appearance5;
			this.cmbMes.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbMes.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbMes.DataSource = this.ultraDataSource3;
			ultraGridColumn8.Header.VisiblePosition = 0;
			ultraGridColumn8.Hidden = true;
			ultraGridColumn9.Header.VisiblePosition = 1;
			ultraGridColumn9.Width = 130;
			ultraGridBand4.Columns.AddRange(new object[] {
																										 ultraGridColumn8,
																										 ultraGridColumn9});
			this.cmbMes.DisplayLayout.BandsSerializer.Add(ultraGridBand4);
			this.cmbMes.DisplayMember = "Mes";
			this.cmbMes.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbMes.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbMes.Location = new System.Drawing.Point(80, 136);
			this.cmbMes.Name = "cmbMes";
			this.cmbMes.Size = new System.Drawing.Size(136, 21);
			this.cmbMes.TabIndex = 1008;
			this.cmbMes.ValueMember = "idMes";
			// 
			// ultraDataSource3
			// 
			ultraDataColumn6.DataType = typeof(int);
			this.ultraDataSource3.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn6,
																																 ultraDataColumn7});
			// 
			// label4
			// 
			this.label4.AutoSize = true;
			this.label4.Location = new System.Drawing.Point(16, 138);
			this.label4.Name = "label4";
			this.label4.Size = new System.Drawing.Size(26, 16);
			this.label4.TabIndex = 1007;
			this.label4.Text = "Mes";
			// 
			// btnConsultar
			// 
			this.btnConsultar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnConsultar.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(247)), ((System.Byte)(245)), ((System.Byte)(241)));
			this.btnConsultar.CausesValidation = false;
			this.btnConsultar.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.btnConsultar.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(50)), ((System.Byte)(50)), ((System.Byte)(50)));
			this.btnConsultar.Image = ((System.Drawing.Image)(resources.GetObject("btnConsultar.Image")));
			this.btnConsultar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnConsultar.Location = new System.Drawing.Point(256, 136);
			this.btnConsultar.Name = "btnConsultar";
			this.btnConsultar.Size = new System.Drawing.Size(78, 23);
			this.btnConsultar.TabIndex = 1009;
			this.btnConsultar.Text = "Consultar";
			this.btnConsultar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnConsultar.Click += new System.EventHandler(this.btnConsultar_Click);
			// 
			// uGridInformacion
			// 
			this.uGridInformacion.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left)));
			this.uGridInformacion.Cursor = System.Windows.Forms.Cursors.Default;
			appearance6.BackColor = System.Drawing.Color.White;
			this.uGridInformacion.DisplayLayout.Appearance = appearance6;
			this.uGridInformacion.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
			this.uGridInformacion.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.uGridInformacion.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.False;
			appearance7.BackColor = System.Drawing.Color.Transparent;
			this.uGridInformacion.DisplayLayout.Override.CardAreaAppearance = appearance7;
			appearance8.ForeColor = System.Drawing.Color.Black;
			appearance8.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridInformacion.DisplayLayout.Override.CellAppearance = appearance8;
			appearance9.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance9.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance9.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance9.FontData.BoldAsString = "True";
			appearance9.FontData.Name = "Arial";
			appearance9.FontData.SizeInPoints = 8F;
			appearance9.ForeColor = System.Drawing.Color.White;
			appearance9.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridInformacion.DisplayLayout.Override.HeaderAppearance = appearance9;
			this.uGridInformacion.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti;
			appearance10.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance10.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance10.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridInformacion.DisplayLayout.Override.RowSelectorAppearance = appearance10;
			appearance11.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance11.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance11.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridInformacion.DisplayLayout.Override.SelectedRowAppearance = appearance11;
			this.uGridInformacion.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridInformacion.Location = new System.Drawing.Point(16, 176);
			this.uGridInformacion.Name = "uGridInformacion";
			this.uGridInformacion.Size = new System.Drawing.Size(1000, 288);
			this.uGridInformacion.TabIndex = 1010;
			this.uGridInformacion.ClickCellButton += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.uGridInformacion_ClickCellButton);
			// 
			// btnExcel
			// 
			this.btnExcel.Image = ((System.Drawing.Image)(resources.GetObject("btnExcel.Image")));
			this.btnExcel.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnExcel.Location = new System.Drawing.Point(936, 144);
			this.btnExcel.Name = "btnExcel";
			this.btnExcel.Size = new System.Drawing.Size(75, 24);
			this.btnExcel.TabIndex = 1011;
			this.btnExcel.Text = "Excel";
			this.btnExcel.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnExcel.Click += new System.EventHandler(this.btnExcel_Click);
			// 
			// frmReportePresupuestoXLider
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(1024, 470);
			this.Controls.Add(this.btnExcel);
			this.Controls.Add(this.uGridInformacion);
			this.Controls.Add(this.btnConsultar);
			this.Controls.Add(this.cmbMes);
			this.Controls.Add(this.label4);
			this.Controls.Add(this.cmbBodega);
			this.Controls.Add(this.label3);
			this.Controls.Add(this.cmbCuenta);
			this.Controls.Add(this.label2);
			this.Controls.Add(this.cmbResponsable);
			this.Controls.Add(this.label1);
			this.Controls.Add(this.txtAnio);
			this.Controls.Add(this.lblPeriodo);
			this.KeyPreview = true;
			this.Name = "frmReportePresupuestoXLider";
			this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
			this.Text = "Lider - Cuenta - Proyecto";
			this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.frmReportePresupuestoXLider_KeyDown);
			this.Load += new System.EventHandler(this.frmReportePresupuestoXLider_Load);
			((System.ComponentModel.ISupportInitialize)(this.txtAnio)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbResponsable)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCuenta)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbBodega)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbMes)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource3)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridInformacion)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void UnloadMe()
		{
			this.Close();
		}
		private Acceso miAcceso;

		private void frmReportePresupuestoXLider_Load(object sender, System.EventArgs e)
		{

			miAcceso = new Acceso(cdsSeteoF, "261006");
			
			if (!Funcion.Permiso("1876", cdsSeteoF))
			{				
				MessageBox.Show("No puede ingresar a LIDER-CUENTA-PROYECTO", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				BeginInvoke(new MethodInvoker(UnloadMe));
				return;
			}
			if (miAcceso.Buscar) this.cmbResponsable.Enabled = true;

			this.txtAnio.MinValue = 2019;
			this.txtAnio.MaxValue = DateTime.Today.Year + 1;
			this.txtAnio.Value = DateTime.Today.Year;
			this.cmbMes.DataSource = Funcion.dvProcedimiento(cdsSeteoF, string.Format("select idMes, upper(Mes) as Mes from Mes"));
		}

		private void txtAnio_ValueChanged(object sender, System.EventArgs e)
		{
			this.cmbResponsable.DataSource = Funcion.dvProcedimiento(cdsSeteoF, string.Format("select p.idNomina, (n.ApellidoPaterno + ' ' + n.ApellidoMaterno + ' '+ n.PrimerNombre + ' '+n.SegundoNombre) as Nombre from dbo.PresupuestoCCxPersonal p Inner Join dbo.Nomina n on n.idNomina = p.idNomina where Periodo = {0} group by p.idNomina, n.ApellidoPaterno ,n.ApellidoMaterno , n.PrimerNombre ,n.SegundoNombre", 
				(int) this.txtAnio.Value));

			if (!miAcceso.Buscar)
			{
				this.cmbResponsable.Value = Funcion.iEscalarSQL(cdsSeteoF, string.Format("If Exists(Select idNomina From Nomina Where Codigo In (Select Nombre From Usuario where idUsuario = {0})) Select idNomina From Nomina Where Codigo In (Select Nombre From Usuario where idUsuario = {0}) Else Select 0", MenuLatinium.IdUsuario));					
			}
			else
			{
				this.cmbResponsable.Enabled = true;
			}
		}

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);		
	
		}

		private void cmbResponsable_ValueChanged(object sender, System.EventArgs e)
		{
			if (this.cmbResponsable.Value !=  null) 
			{
				this.cmbCuenta.DataSource = Funcion.dvProcedimiento(cdsSeteoF, string.Format("  SELECT c.idCuenta, cu.Codigo, cu.Descripcion as Cuenta FROM dbo.PresupuestoCCxPersonal p Inner Join dbo.PresupuestoCCxCuenta c on c.idProyecto = p.idProyecto Inner join dbo.Cuenta cu on cu.idCuenta = c.idCuenta WHERE p.Periodo = {0} and idNomina = {1} group by c.idCuenta, cu.Codigo, cu.Descripcion order by cu.Codigo", 
					(int) this.txtAnio.Value, (int) this.cmbResponsable.Value ));
			}
			else
			{
				this.cmbCuenta.Value = System.DBNull.Value;
			}
		}

		private void cmbCuenta_ValueChanged(object sender, System.EventArgs e)
		{

			if (this.cmbResponsable.Value != null ) 
			{
				if ( this.cmbCuenta.Value != null) 
				{
					this.cmbBodega.DataSource = Funcion.dvProcedimiento(cdsSeteoF, string.Format("  SELECT cu.idProyecto, cu.Nombre from   dbo.PresupuestoCCxPersonal p Inner Join dbo.PresupuestoCCxCuenta c on c.idProyecto = p.idProyecto Inner join dbo.Proyecto cu on cu.idProyecto = c.idProyecto WHERE p.Periodo = {0} and idNomina = {1}  and c.idCuenta = {2} group by cu.idProyecto, cu.nombre order by cu.nombre", 
						(int) this.txtAnio.Value, (int) this.cmbResponsable.Value, (int) this.cmbCuenta.Value ));
				}
				else
				{
					this.cmbBodega.Value = System.DBNull.Value;
				}
			}
			else
			{
					this.cmbBodega.Value = System.DBNull.Value;
			}
		}
		
		private void btnConsultar_Click(object sender, System.EventArgs e)
		{
			if (!Validacion.vbComboVacio(this.cmbResponsable, "Seleccione el personal.")) return;

			if (this.txtAnio.Text.ToString().Trim().Length == 0)
			{
				MessageBox.Show("Seleccione un periodo", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.txtAnio.Focus();
				return;
			}
			int idPersonal  = 0;
			int idCuenta = 0;
			int idProyecto = 0;
			int idMes  = 0;
			

			if (this.cmbResponsable.ActiveRow != null) idPersonal = (int)this.cmbResponsable.Value;
			if (this.cmbCuenta.ActiveRow != null) idCuenta = (int)this.cmbCuenta.Value;
			if (this.cmbBodega.ActiveRow != null) idProyecto = (int)this.cmbBodega.Value;
			if (this.cmbMes.ActiveRow != null) idMes = (int)this.cmbMes.Value;
      this.uGridInformacion.DataSource = null;
			string sSQLPa = string.Format("Exec ReportePresupuestoXLider {0}, {1}, {2}, {3}, {4}", 
				(int) this.txtAnio.Value, idPersonal, idCuenta , idProyecto, idMes);
			this.uGridInformacion.DataSource = FuncionesProcedimientos.dtGeneral(sSQLPa);
     this.DiseñoGridSumatoria(this.uGridInformacion);

		}

		private void DiseñoGridSumatoria(Infragistics.Win.UltraWinGrid.UltraGrid uGrid)
		{
				this.uGridInformacion.DisplayLayout.Bands[0].Columns["idCuenta"].Hidden = true;
			this.uGridInformacion.DisplayLayout.Bands[0].Columns["Codigo"].Width = 150;
			this.uGridInformacion.DisplayLayout.Bands[0].Columns["Descripcion"].Width = 200;

			if (this.uGridInformacion.Rows.Count > 0)
			{				
				int iContColumnas = (int)this.uGridInformacion.DisplayLayout.Bands[0].Columns.Count;					
					
				string nombrCol = "";

				for (int j = 3; j < iContColumnas; j++)
				{
					nombrCol = this.uGridInformacion.DisplayLayout.Bands[0].Columns[j].Header.Caption;	
					this.uGridInformacion.DisplayLayout.Bands[0].Columns[nombrCol].Style = ColumnStyle.Button;
					this.uGridInformacion.DisplayLayout.Bands[0].Columns[nombrCol].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right;
					this.uGridInformacion.DisplayLayout.Bands[0].Columns[nombrCol].Width = 130;
					this.uGridInformacion.DisplayLayout.Bands[0].Columns[nombrCol].Format = "#,##0.00";
					this.uGridInformacion.DisplayLayout.Bands[0].Columns[nombrCol].Header.Caption = nombrCol;	
				}
			}

			#region Diseño Grid
			if (uGrid.Rows.Count > 0)
			{
				int iColIni = 3;

				
				if (uGrid.Rows.Count > 0)
				{
					string nombrCol = "";
							
					int iColumnas = (int)uGrid.DisplayLayout.Bands[0].Columns.Count;
				
					for (int j = iColIni; j < iColumnas; j++)
					{
							
						nombrCol = uGrid.DisplayLayout.Bands[0].Columns[j].Header.Caption;
										
						uGrid.DisplayLayout.Bands[0].Columns[nombrCol].CellActivation = Activation.ActivateOnly;
						uGrid.DisplayLayout.Bands[0].Columns[nombrCol].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right;
							
						uGrid.DisplayLayout.Bands[0].Columns[nombrCol].Width = 80;
						uGrid.DisplayLayout.Bands[0].Summaries.Add(nombrCol, Infragistics.Win.UltraWinGrid.SummaryType.Sum, uGrid.DisplayLayout.Bands[0].Columns[nombrCol], Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn);
						//							int a = Infragistics.Win.UltraWinGrid.SummaryType.Sum;			
						uGrid.DisplayLayout.Bands[0].Summaries[nombrCol].Appearance.TextHAlign = Infragistics.Win.HAlign.Right;
						uGrid.DisplayLayout.Bands[0].Summaries[nombrCol].DisplayFormat = "{0: #,##0.00}";
							
					}					
				}				
			}	
			#endregion Diseño Grid
		}

		private void uGridInformacion_ClickCellButton(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			string sColumna = e.Cell.Column.Key;
			int iMes = 0 ;

			if (sColumna == "ENERO"	) iMes = 1;
			if (sColumna == "FEBRERO"	) iMes = 2;
			if (sColumna == "MARZO"	) iMes = 3;
			if (sColumna == "ABRIL"	) iMes = 4;
			if (sColumna == "MAYO") iMes = 5;
			if (sColumna == "JUNIO"	) iMes = 6;
			if (sColumna == "JULIO"	) iMes = 7;
			if (sColumna == "AGOSTO") iMes = 8;
			if (sColumna == "SEPTIEMBRE") iMes = 9;
			if (sColumna == "OCTUBRE"	) iMes = 10;
			if (sColumna == "NOVIEMBRE"	) iMes = 11;
			if (sColumna == "DICIEMBRE"	) iMes = 12;

			if (e.Cell.Row.Cells[sColumna].Value == System.DBNull.Value)
			{
				MessageBox.Show(string.Format("No hay  registrado en la Columna : {0}", sColumna), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return;
			}
			else
			{
				int idPersonal  = 0;
				int idCuenta = 0;
				int idProyecto = 0;
				int idMes  = 0;
			

				if (this.cmbResponsable.ActiveRow != null) idPersonal = (int)this.cmbResponsable.Value;
				if (this.cmbCuenta.ActiveRow != null) idCuenta = (int)this.cmbCuenta.Value;
				if (this.cmbBodega.ActiveRow != null) idProyecto = (int)this.cmbBodega.Value;
				if (this.cmbMes.ActiveRow != null) idMes = (int)this.cmbMes.Value;

				string sMes =sColumna;
				sMes.Replace("Gasto_", "");
				using (frmReportePeriodoGastoDetalleComparacion CCJ = new frmReportePeriodoGastoDetalleComparacion ( (int) this.txtAnio.Value, idPersonal, (int)e.Cell.Row.Cells["idCuenta"].Value , sColumna,idProyecto, true))
				{
					if (DialogResult.OK == CCJ.ShowDialog())
					{
						
					}
				}
			}
		}

		private void btnExcel_Click(object sender, System.EventArgs e)
		{
			FuncionesProcedimientos.ExportaExcel(this.uGridInformacion);
		}

		private void frmReportePresupuestoXLider_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Escape) this.Close();
		}
	}
}
