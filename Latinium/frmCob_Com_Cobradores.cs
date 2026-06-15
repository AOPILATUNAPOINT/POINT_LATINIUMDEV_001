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
using System.Net;
using System.Web;
using System.Xml;
using System.Xml.Schema;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de frmCob_Com_Cobradores.
	/// </summary>
	public class frmCob_Com_Cobradores : System.Windows.Forms.Form
	{
		private System.Windows.Forms.GroupBox groupBox1;
		private System.Windows.Forms.Button button1;
		private System.Windows.Forms.Button button2;
		private Infragistics.Win.UltraWinGrid.UltraDropDown cmbCargos;
		public Infragistics.Win.UltraWinGrid.UltraGrid uGridDepositos;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource4;
		private C1.Data.C1DataSet cdsSeteoF;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource1;
		private System.Windows.Forms.Label label4;
		private System.Windows.Forms.DateTimePicker dtpPeriodo;
		private System.Windows.Forms.Button btnVer;
		private System.Windows.Forms.Button btnComandos;
		private System.Windows.Forms.ContextMenu mnuAcciones;
		private System.Windows.Forms.MenuItem mnuAccionesCopiarUltimoMes;
		private System.Windows.Forms.MenuItem mnuAccionesAgregarNuevoLocal;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public frmCob_Com_Cobradores()
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
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmCob_Com_Cobradores));
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idZonaGeografica");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idZonaGeografica");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.Appearance appearance1 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPresupuestoCobradores");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idZonaGeografica", -1, "cmbCargos");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("DiasDesde");
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("DiasHasta");
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("PorcDesde");
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("PorcHasta");
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bono");
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance11 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn3 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idPresupuestoCobradores");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn4 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idZonaGeografica");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn5 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("DiasDesde");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn6 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("DiasHasta");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn7 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("PorcDesde");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn8 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("PorcHasta");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn9 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Bono");
			this.groupBox1 = new System.Windows.Forms.GroupBox();
			this.button1 = new System.Windows.Forms.Button();
			this.button2 = new System.Windows.Forms.Button();
			this.cmbCargos = new Infragistics.Win.UltraWinGrid.UltraDropDown();
			this.ultraDataSource1 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.uGridDepositos = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource4 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.label4 = new System.Windows.Forms.Label();
			this.dtpPeriodo = new System.Windows.Forms.DateTimePicker();
			this.btnVer = new System.Windows.Forms.Button();
			this.btnComandos = new System.Windows.Forms.Button();
			this.mnuAcciones = new System.Windows.Forms.ContextMenu();
			this.mnuAccionesCopiarUltimoMes = new System.Windows.Forms.MenuItem();
			this.mnuAccionesAgregarNuevoLocal = new System.Windows.Forms.MenuItem();
			((System.ComponentModel.ISupportInitialize)(this.cmbCargos)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridDepositos)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource4)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			this.SuspendLayout();
			// 
			// groupBox1
			// 
			this.groupBox1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.groupBox1.Location = new System.Drawing.Point(0, 48);
			this.groupBox1.Name = "groupBox1";
			this.groupBox1.Size = new System.Drawing.Size(656, 8);
			this.groupBox1.TabIndex = 779;
			this.groupBox1.TabStop = false;
			// 
			// button1
			// 
			this.button1.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.button1.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(230)), ((System.Byte)(230)), ((System.Byte)(230)));
			this.button1.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.button1.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(50)), ((System.Byte)(50)), ((System.Byte)(50)));
			this.button1.Image = ((System.Drawing.Image)(resources.GetObject("button1.Image")));
			this.button1.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.button1.Location = new System.Drawing.Point(400, 8);
			this.button1.Name = "button1";
			this.button1.Size = new System.Drawing.Size(76, 23);
			this.button1.TabIndex = 777;
			this.button1.Text = "Guardar";
			this.button1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.button1.Visible = false;
			this.button1.Click += new System.EventHandler(this.button1_Click);
			// 
			// button2
			// 
			this.button2.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.button2.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(230)), ((System.Byte)(230)), ((System.Byte)(230)));
			this.button2.CausesValidation = false;
			this.button2.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.button2.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(50)), ((System.Byte)(50)), ((System.Byte)(50)));
			this.button2.Image = ((System.Drawing.Image)(resources.GetObject("button2.Image")));
			this.button2.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.button2.Location = new System.Drawing.Point(560, 8);
			this.button2.Name = "button2";
			this.button2.TabIndex = 778;
			this.button2.Text = "Salir";
			this.button2.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.button2.Click += new System.EventHandler(this.button2_Click);
			// 
			// cmbCargos
			// 
			this.cmbCargos.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbCargos.DataSource = this.ultraDataSource1;
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridColumn2.Width = 198;
			ultraGridBand1.Columns.AddRange(new object[] {
																										 ultraGridColumn1,
																										 ultraGridColumn2});
			this.cmbCargos.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.cmbCargos.DisplayMember = "Nombre";
			this.cmbCargos.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F);
			this.cmbCargos.Location = new System.Drawing.Point(84, 219);
			this.cmbCargos.Name = "cmbCargos";
			this.cmbCargos.Size = new System.Drawing.Size(200, 112);
			this.cmbCargos.TabIndex = 776;
			this.cmbCargos.ValueMember = "idZonaGeografica";
			this.cmbCargos.Visible = false;
			// 
			// ultraDataSource1
			// 
			ultraDataColumn1.DataType = typeof(int);
			this.ultraDataSource1.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn1,
																																 ultraDataColumn2});
			// 
			// uGridDepositos
			// 
			this.uGridDepositos.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.uGridDepositos.Cursor = System.Windows.Forms.Cursors.Default;
			this.uGridDepositos.DataSource = this.ultraDataSource4;
			appearance1.BackColor = System.Drawing.Color.White;
			this.uGridDepositos.DisplayLayout.Appearance = appearance1;
			ultraGridColumn3.Header.VisiblePosition = 0;
			ultraGridColumn3.Hidden = true;
			ultraGridColumn3.Width = 146;
			ultraGridColumn4.ButtonDisplayStyle = Infragistics.Win.UltraWinGrid.ButtonDisplayStyle.Always;
			ultraGridColumn4.Header.Caption = "Zona Geografica";
			ultraGridColumn4.Header.VisiblePosition = 1;
			ultraGridColumn4.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDown;
			ultraGridColumn4.Width = 180;
			appearance2.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn5.CellAppearance = appearance2;
			ultraGridColumn5.Header.Caption = "Dias Desde";
			ultraGridColumn5.Header.VisiblePosition = 2;
			ultraGridColumn5.Width = 70;
			appearance3.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn6.CellAppearance = appearance3;
			ultraGridColumn6.Header.Caption = "Dias Hasta";
			ultraGridColumn6.Header.VisiblePosition = 3;
			ultraGridColumn6.Width = 71;
			appearance4.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn7.CellAppearance = appearance4;
			ultraGridColumn7.Format = "#,##0.00";
			ultraGridColumn7.Header.Caption = "% Desde";
			ultraGridColumn7.Header.VisiblePosition = 4;
			ultraGridColumn7.Width = 82;
			appearance5.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn8.CellAppearance = appearance5;
			ultraGridColumn8.Format = "#,##0.00";
			ultraGridColumn8.Header.Caption = "% Hasta";
			ultraGridColumn8.Header.VisiblePosition = 5;
			ultraGridColumn8.Width = 82;
			appearance6.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn9.CellAppearance = appearance6;
			ultraGridColumn9.Format = "#,##0.00";
			ultraGridColumn9.Header.VisiblePosition = 6;
			ultraGridColumn9.Width = 82;
			ultraGridBand2.Columns.AddRange(new object[] {
																										 ultraGridColumn3,
																										 ultraGridColumn4,
																										 ultraGridColumn5,
																										 ultraGridColumn6,
																										 ultraGridColumn7,
																										 ultraGridColumn8,
																										 ultraGridColumn9});
			ultraGridBand2.Header.Caption = "Configuración";
			ultraGridBand2.HeaderVisible = true;
			ultraGridBand2.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.Edit;
			ultraGridBand2.SummaryFooterCaption = "Total";
			this.uGridDepositos.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			this.uGridDepositos.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.TemplateOnBottom;
			appearance7.BackColor = System.Drawing.Color.Transparent;
			this.uGridDepositos.DisplayLayout.Override.CardAreaAppearance = appearance7;
			appearance8.ForeColor = System.Drawing.Color.Black;
			appearance8.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridDepositos.DisplayLayout.Override.CellAppearance = appearance8;
			this.uGridDepositos.DisplayLayout.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.Edit;
			appearance9.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance9.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance9.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance9.FontData.BoldAsString = "True";
			appearance9.FontData.Name = "Arial";
			appearance9.FontData.SizeInPoints = 8F;
			appearance9.ForeColor = System.Drawing.Color.White;
			appearance9.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridDepositos.DisplayLayout.Override.HeaderAppearance = appearance9;
			appearance10.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance10.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance10.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridDepositos.DisplayLayout.Override.RowSelectorAppearance = appearance10;
			appearance11.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance11.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance11.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridDepositos.DisplayLayout.Override.SelectedRowAppearance = appearance11;
			this.uGridDepositos.Font = new System.Drawing.Font("Tahoma", 8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridDepositos.Location = new System.Drawing.Point(8, 64);
			this.uGridDepositos.Name = "uGridDepositos";
			this.uGridDepositos.Size = new System.Drawing.Size(616, 368);
			this.uGridDepositos.TabIndex = 775;
			this.uGridDepositos.AfterRowInsert += new Infragistics.Win.UltraWinGrid.RowEventHandler(this.uGridDepositos_AfterRowInsert);
			this.uGridDepositos.BeforeRowsDeleted += new Infragistics.Win.UltraWinGrid.BeforeRowsDeletedEventHandler(this.uGridDepositos_BeforeRowsDeleted);
			this.uGridDepositos.AfterCellUpdate += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.uGridDepositos_AfterCellUpdate);
			// 
			// ultraDataSource4
			// 
			ultraDataColumn3.DataType = typeof(int);
			ultraDataColumn4.DataType = typeof(int);
			ultraDataColumn5.DataType = typeof(int);
			ultraDataColumn6.DataType = typeof(int);
			ultraDataColumn7.DataType = typeof(System.Decimal);
			ultraDataColumn8.DataType = typeof(System.Decimal);
			ultraDataColumn9.DataType = typeof(System.Decimal);
			this.ultraDataSource4.Band.Columns.AddRange(new object[] {
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
			// label4
			// 
			this.label4.AutoSize = true;
			this.label4.BackColor = System.Drawing.Color.Transparent;
			this.label4.Location = new System.Drawing.Point(8, 11);
			this.label4.Name = "label4";
			this.label4.Size = new System.Drawing.Size(43, 16);
			this.label4.TabIndex = 781;
			this.label4.Text = "Periodo";
			this.label4.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// dtpPeriodo
			// 
			this.dtpPeriodo.CustomFormat = "MM/yyyy";
			this.dtpPeriodo.Format = System.Windows.Forms.DateTimePickerFormat.Custom;
			this.dtpPeriodo.Location = new System.Drawing.Point(64, 9);
			this.dtpPeriodo.MinDate = new System.DateTime(2016, 5, 1, 0, 0, 0, 0);
			this.dtpPeriodo.Name = "dtpPeriodo";
			this.dtpPeriodo.ShowUpDown = true;
			this.dtpPeriodo.Size = new System.Drawing.Size(96, 20);
			this.dtpPeriodo.TabIndex = 782;
			this.dtpPeriodo.Value = new System.DateTime(2016, 5, 1, 0, 0, 0, 0);
			this.dtpPeriodo.ValueChanged += new System.EventHandler(this.dtpPeriodo_ValueChanged);
			// 
			// btnVer
			// 
			this.btnVer.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(230)), ((System.Byte)(230)), ((System.Byte)(230)));
			this.btnVer.CausesValidation = false;
			this.btnVer.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.btnVer.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(50)), ((System.Byte)(50)), ((System.Byte)(50)));
			this.btnVer.Image = ((System.Drawing.Image)(resources.GetObject("btnVer.Image")));
			this.btnVer.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnVer.Location = new System.Drawing.Point(168, 8);
			this.btnVer.Name = "btnVer";
			this.btnVer.Size = new System.Drawing.Size(72, 23);
			this.btnVer.TabIndex = 783;
			this.btnVer.Text = "Consultar";
			this.btnVer.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnVer.Click += new System.EventHandler(this.btnVer_Click);
			// 
			// btnComandos
			// 
			this.btnComandos.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btnComandos.Image = ((System.Drawing.Image)(resources.GetObject("btnComandos.Image")));
			this.btnComandos.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnComandos.Location = new System.Drawing.Point(480, 8);
			this.btnComandos.Name = "btnComandos";
			this.btnComandos.TabIndex = 977;
			this.btnComandos.Text = "Acciones";
			this.btnComandos.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnComandos.Visible = false;
			this.btnComandos.Click += new System.EventHandler(this.btnComandos_Click);
			// 
			// mnuAcciones
			// 
			this.mnuAcciones.MenuItems.AddRange(new System.Windows.Forms.MenuItem[] {
																																								this.mnuAccionesCopiarUltimoMes,
																																								this.mnuAccionesAgregarNuevoLocal});
			// 
			// mnuAccionesCopiarUltimoMes
			// 
			this.mnuAccionesCopiarUltimoMes.Index = 0;
			this.mnuAccionesCopiarUltimoMes.Text = "Copiar Ultimo Mes";
			this.mnuAccionesCopiarUltimoMes.Click += new System.EventHandler(this.mnuAccionesCopiarUltimoMes_Click);
			// 
			// mnuAccionesAgregarNuevoLocal
			// 
			this.mnuAccionesAgregarNuevoLocal.Index = 1;
			this.mnuAccionesAgregarNuevoLocal.Text = "Exportar";
			this.mnuAccionesAgregarNuevoLocal.Click += new System.EventHandler(this.mnuAccionesAgregarNuevoLocal_Click);
			// 
			// frmCob_Com_Cobradores
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(640, 446);
			this.Controls.Add(this.btnComandos);
			this.Controls.Add(this.btnVer);
			this.Controls.Add(this.label4);
			this.Controls.Add(this.dtpPeriodo);
			this.Controls.Add(this.groupBox1);
			this.Controls.Add(this.button1);
			this.Controls.Add(this.button2);
			this.Controls.Add(this.cmbCargos);
			this.Controls.Add(this.uGridDepositos);
			this.KeyPreview = true;
			this.Name = "frmCob_Com_Cobradores";
			this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
			this.Text = "Configuración";
			this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.frmCob_Com_Cobradores_KeyDown);
			this.Load += new System.EventHandler(this.frmCob_Com_Cobradores_Load);
			((System.ComponentModel.ISupportInitialize)(this.cmbCargos)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridDepositos)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource4)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);		
	
		}

		private Acceso miAcceso;
		
		private void UnloadMe()
		{
			this.Close();
		}

		private void frmCob_Com_Cobradores_Load(object sender, System.EventArgs e)
		{
			
			miAcceso = new Acceso(cdsSeteoF, "083608");

			if (!Funcion.Permiso("1836", cdsSeteoF))
			{		
				MessageBox.Show("No tiene Acceso a Porcentaje de Cobradores", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				BeginInvoke(new MethodInvoker(UnloadMe));
				return;
			}

			if(miAcceso.BExportar)this.btnComandos.Visible = true;
			if(miAcceso.Nuevo)this.button1.Visible = true;


			this.dtpPeriodo.Value = DateTime.Parse("01/" + DateTime.Now.Month.ToString() + "/" + DateTime.Now.Year.ToString());
			this.cmbCargos.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Exec Listas 93, 0, 0");
		}
		private void Consulta()
		{
			if (this.dtpPeriodo.Text.ToString().Length == 0) return;			
			string sSQL = string.Format("Exec [consultaPresupuestoCobradores] '{0}'", 
				this.dtpPeriodo.Text.ToString());
			FuncionesProcedimientos.dsGeneral(sSQL, this.uGridDepositos);

		}

		private void uGridDepositos_AfterRowInsert(object sender, Infragistics.Win.UltraWinGrid.RowEventArgs e)
		{
			e.Row.Cells["idPresupuestoCobradores"].Value = 0;
      e.Row.Cells["idZonaGeografica"].Value = 0;
			e.Row.Cells["DiasDesde"].Value = 0;
			e.Row.Cells["DiasHasta"].Value = 0;
			e.Row.Cells["PorcDesde"].Value = 0.00m;			
			e.Row.Cells["PorcHasta"].Value = 0.00m;
			e.Row.Cells["Bono"].Value = 0.00m;
		}

			bool bActualiza = false;

		private void uGridDepositos_AfterCellUpdate(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			if (bActualiza) return;
			#region idZonaGeografica
			if (e.Cell.Column.ToString() == "idZonaGeografica")
			{	
				bActualiza = true;

				if (e.Cell.Row.Cells["idZonaGeografica"].Value == System.DBNull.Value)
				{					
					e.Cell.Row.Cells["idZonaGeografica"].Value = (int)e.Cell.OriginalValue;	
				}	
				bActualiza = false;
			}
      #endregion idZonaGeografica

			#region DiasDesde
			if (e.Cell.Column.ToString() == "DiasDesde")
			{	
				bActualiza = true;

				if (e.Cell.Row.Cells["DiasDesde"].Value == System.DBNull.Value)
				{					
					e.Cell.Row.Cells["DiasDesde"].Value = (int)e.Cell.OriginalValue;	
				}	
				if ((int ) e.Cell.Row.Cells["DiasDesde"].Value < 0 )
				{					
					e.Cell.Row.Cells["DiasDesde"].Value = (int)e.Cell.OriginalValue;	
				}
				bActualiza = false;
			}
      #endregion DiasHasta

			#region DiasDesde
			if (e.Cell.Column.ToString() == "DiasHasta")
			{	
				bActualiza = true;

				if (e.Cell.Row.Cells["DiasHasta"].Value == System.DBNull.Value)
				{					
					e.Cell.Row.Cells["DiasHasta"].Value = (int)e.Cell.OriginalValue;	
				}	
				if ((int ) e.Cell.Row.Cells["DiasHasta"].Value < 0 )
				{					
					e.Cell.Row.Cells["DiasHasta"].Value = (int)e.Cell.OriginalValue;	
				}
				bActualiza = false;
			}
			#endregion DiasHasta

			#region PorcDesde
			if (e.Cell.Column.ToString() == "PorcDesde")
			{	
				bActualiza = true;

				if (e.Cell.Row.Cells["PorcDesde"].Value == System.DBNull.Value)
				{					
					e.Cell.Row.Cells["PorcDesde"].Value = (decimal)e.Cell.OriginalValue;	
				}	
				if ((decimal ) e.Cell.Row.Cells["PorcDesde"].Value < 0 )
				{					
					e.Cell.Row.Cells["PorcDesde"].Value = (decimal)e.Cell.OriginalValue;	
				}
				bActualiza = false;
			}
     	#endregion PorcDesde
     
			#region PorcHasta
			if (e.Cell.Column.ToString() == "PorcHasta")
			{	
				bActualiza = true;

				if (e.Cell.Row.Cells["PorcHasta"].Value == System.DBNull.Value)
				{					
					e.Cell.Row.Cells["PorcHasta"].Value = (decimal)e.Cell.OriginalValue;	
				}	
				if ((decimal ) e.Cell.Row.Cells["PorcHasta"].Value < 0 )
				{					
					e.Cell.Row.Cells["PorcHasta"].Value = (decimal)e.Cell.OriginalValue;	
				}
				bActualiza = false;
			}
      #endregion PorcHasta

			#region Bono
			if (e.Cell.Column.ToString() == "Bono")
			{	
				bActualiza = true;

				if (e.Cell.Row.Cells["Bono"].Value == System.DBNull.Value)
				{					
					e.Cell.Row.Cells["Bono"].Value = (decimal)e.Cell.OriginalValue;	
				}	
				if ((decimal ) e.Cell.Row.Cells["Bono"].Value < 0 )
				{					
					e.Cell.Row.Cells["Bono"].Value = (decimal)e.Cell.OriginalValue;	
				}
				bActualiza = false;
			}
      #endregion Bono



		}

		private void button1_Click(object sender, System.EventArgs e)
		{
			if (this.dtpPeriodo.Text.ToString().Trim().Length == 0)
			{
				MessageBox.Show("Seleccione un periodo", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.dtpPeriodo.Focus();
				return;
			}
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridDepositos.Rows.All)
			{
				#region Cargo
				if (dr.Cells["idZonaGeografica"].Value == System.DBNull.Value)
				{
					MessageBox.Show("Seleccione una Zona Geografica", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);

					this.uGridDepositos.ActiveRow = dr;
					this.uGridDepositos.ActiveRow.Selected = true;

					return;
				}
				if ((int)dr.Cells["idZonaGeografica"].Value == 0)
				{
					MessageBox.Show("Seleccione una Zona Geografica", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);

					this.uGridDepositos.ActiveRow = dr;
					this.uGridDepositos.ActiveRow.Selected = true;

					return;
				}
				#endregion Cargo

				#region Desde
				if (dr.Cells["DiasDesde"].Value == System.DBNull.Value)
				{
					MessageBox.Show("Ingrese el dia Desde.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);

					this.uGridDepositos.ActiveRow = dr;
					this.uGridDepositos.ActiveRow.Selected = true;

					return;
				}
				#endregion Desde

				#region Hasta
				if (dr.Cells["DiasHasta"].Value == System.DBNull.Value)
				{
					MessageBox.Show("Ingrese el dia Hasta.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);

					this.uGridDepositos.ActiveRow = dr;
					this.uGridDepositos.ActiveRow.Selected = true;

					return;
				}
				#endregion Hasta

				#region PorcDesde

				if (dr.Cells["PorcDesde"].Value == System.DBNull.Value)
				{
					MessageBox.Show("Ingrese el % Desde.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);

					this.uGridDepositos.ActiveRow = dr;
					this.uGridDepositos.ActiveRow.Selected = true;

					return;
				}
				#endregion PorcDesde

				#region PorcHasta

				if (dr.Cells["PorcHasta"].Value == System.DBNull.Value)
				{
					MessageBox.Show("Ingrese el % Hasta.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);

					this.uGridDepositos.ActiveRow = dr;
					this.uGridDepositos.ActiveRow.Selected = true;

					return;
				}
				#endregion PorcHasta

				#region PorcHasta

				if (dr.Cells["Bono"].Value == System.DBNull.Value)
				{
					MessageBox.Show("Ingrese el Bono.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);

					this.uGridDepositos.ActiveRow = dr;
					this.uGridDepositos.ActiveRow.Selected = true;

					return;
				}
				#endregion PorcHasta

				#region Porcentajes Inicial Final
				if (Convert.ToInt32(dr.Cells["DiasDesde"].Value) > Convert.ToInt32(dr.Cells["DiasHasta"].Value))
				{
					MessageBox.Show("El dia Inicial no puede ser mayor al Final.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);

					this.uGridDepositos.ActiveRow = dr;
					this.uGridDepositos.ActiveRow.Selected = true;

					return;
				}
				#endregion Porcentajes Inicial Final

				#region % Inicial Final
				if (Convert.ToDecimal(dr.Cells["PorcDesde"].Value) > Convert.ToDecimal(dr.Cells["PorcHasta"].Value))
				{
					MessageBox.Show("El % Inicial no puede ser mayor al % Final.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);

					this.uGridDepositos.ActiveRow = dr;
					this.uGridDepositos.ActiveRow.Selected = true;

					return;
				}
				#endregion % Inicial Final

			}
			DateTime dtDesde = DateTime.Parse("01/" + this.dtpPeriodo.Text.ToString());
			DateTime dtHasta = DateTime.Parse(DateTime.DaysInMonth(Convert.ToInt32(this.dtpPeriodo.Text.Substring(3, 4)), Convert.ToInt32(this.dtpPeriodo.Text.Substring(0, 2))) + "/" + this.dtpPeriodo.Text.ToString());

			if (DialogResult.Yes == MessageBox.Show("¿Esta seguro de guardar.?", "Point Technology", MessageBoxButtons.YesNo, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2))
			{
				#region Guardar

				//			DateTime dtInicia = DateTime.Today;
				//			DateTime dtRetorna = DateTime.Today;
				//				dtInicia = (DateTime)this.dtInicia.Value;
				//				dtRetorna = (DateTime)this.dtRetorna.Value;

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

						#region Graba Maestro
						foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridDepositos.Rows.All)
						{
							string  sSQLDet = string.Format("Exec [GuardaPresupuestoCobradores] {0}, '{1}', {2}, {3}, {4}, {5}, {6}, {7}, '{8}', '{9}'" ,
								Convert.ToInt32 ( dr.Cells["idPresupuestoCobradores"].Value), 
								this.dtpPeriodo.Text.ToString(),
								Convert.ToInt32 ( dr.Cells["idZonaGeografica"].Value), 
								Convert.ToInt32 ( dr.Cells["DiasDesde"].Value), 
								Convert.ToInt32 ( dr.Cells["DiasHasta"].Value), 
								Convert.ToDecimal ( dr.Cells["PorcDesde"].Value),
								Convert.ToDecimal ( dr.Cells["PorcHasta"].Value),
								Convert.ToDecimal ( dr.Cells["Bono"].Value),
									Convert.ToDateTime(dtDesde).ToString("yyyyMMdd"),
									Convert.ToDateTime(dtHasta).ToString("yyyyMMdd")
								);
							oCmdActualiza.CommandText = sSQLDet;
							oCmdActualiza.ExecuteNonQuery();	
						}
						#endregion Graba Maestro

						oTransaction.Commit();

						Cursor = Cursors.Default;	

						MessageBox.Show("Registros Guardado Correctamente", "Información", MessageBoxButtons.OK, MessageBoxIcon.Information);

						#region Controles
				
					Consulta ();

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

		}

		private void dtpPeriodo_ValueChanged(object sender, System.EventArgs e)
		{
			Consulta ();
		}

		private void btnVer_Click(object sender, System.EventArgs e)
		{
			if (this.dtpPeriodo.Text.ToString().Trim().Length == 0)
			{
				MessageBox.Show("Seleccione un periodo", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.dtpPeriodo.Focus();
				return;
			}

			Consulta();
		}

		private void frmCob_Com_Cobradores_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Escape) this.Close();
		}

		private void button2_Click(object sender, System.EventArgs e)
		{
			this.Close();
		}

		private void btnExportar_Click(object sender, System.EventArgs e)
		{
		
		}

		private void mnuAccionesCopiarUltimoMes_Click(object sender, System.EventArgs e)
		{
			int iRetorna = 0;
			if (this.dtpPeriodo.Text.ToString().Trim().Length == 0)
			{
				MessageBox.Show("Seleccione un periodo", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.dtpPeriodo.Focus();
				return;
			}

		  string sPeriodo = Funcion.sEscalarSQL(cdsSeteoF, "Select top 1 Periodo from dbo.PresupuestoCobradores order by  idPresupuestoCobradores desc");

			if (DialogResult.Yes == MessageBox.Show("¿Esta Seguro De Copiar La Información del anterior Mes?", "Point Technology", MessageBoxButtons.YesNo, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2))
			{	

				iRetorna = Funcion.iEscalarSQL(cdsSeteoF, string.Format("Exec [CopiaAsignacioncobradoresPorcentaje] "));
			//	this.uGridDepositos.DataSource = FuncionesProcedimientos.dtGeneral(sSQLPa);
				if ( iRetorna == 1)
				{

					MessageBox.Show(string.Format("Se copio la Información del Periodo {0} correctamente", sPeriodo), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
					this.btnVer_Click(sender, e);
				}
				else
				{
					MessageBox.Show(string.Format("No existen Datos en el Mes anterior {0} o ya se genero una copia en el mes actual ", sPeriodo), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
					this.btnVer_Click(sender, e);

				}
			}
		}

		private void btnComandos_Click(object sender, System.EventArgs e)
		{
			this.mnuAcciones.Show(this.btnComandos, new Point(1, 25));
		}

		private void mnuAccionesAgregarNuevoLocal_Click(object sender, System.EventArgs e)
		{
		FuncionesProcedimientos.ExportaExcel(this.uGridDepositos);
		}

		private void uGridDepositos_BeforeRowsDeleted(object sender, Infragistics.Win.UltraWinGrid.BeforeRowsDeletedEventArgs e)
		{
			if (!miAcceso.BEliminaSeriales) 
			{
				e.Cancel = true;
			}
			else
			{
				if (DialogResult.Yes ==	MessageBox.Show("¿Esta seguro de Borrar las Líneas Seleccionadas?", "Point Technology", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2))				
				{
					for (int i = 0; i < e.Rows.Length; i++)
					{
						if (e.Rows[i].Cells["idPresupuestoCobradores"].Value != System.DBNull.Value)
						{									
							if ((int)e.Rows[i].Cells["idPresupuestoCobradores"].Value > 0)
							{
								string sSQL = string.Format("Delete From PresupuestoCobradores Where idPresupuestoCobradores = {0}", (int)e.Rows[i].Cells["idPresupuestoCobradores"].Value);
								Funcion.EjecutaSQL(cdsSeteoF, sSQL, true);
							}
						}			
					}
					e.DisplayPromptMsg = false;
				}
				else
				{
					e.Cancel = true;
				}
			
			}
		}
	}
}
