using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using DevExpress.XtraEditors;
using System.IO;
using System.Data.OleDb;
using System.Data;
using System.Data.SqlClient;
using System.Globalization;

namespace Latinium
{
	/// <summary>
	/// Summary description for ImportaBlasar.
	/// </summary>
	public class ImportaBlasar : DevExpress.XtraEditors.XtraForm
	{
		private DevExpress.XtraTab.XtraTabControl xtraTabControl1;
		private DevExpress.XtraTab.XtraTabPage xtraTabPage1;
		private DevExpress.XtraTab.XtraTabPage xtraTabPage2;
		private DevExpress.XtraEditors.LabelControl labelControl1;
		private DevExpress.XtraEditors.LabelControl labelControl2;
		private DevExpress.XtraEditors.LabelControl labelControl3;
		private DevExpress.XtraEditors.TextEdit txtArchivo;
		private DevExpress.XtraEditors.SimpleButton btArchivo;
		private DevExpress.XtraEditors.SimpleButton btLeer;
		private DevExpress.XtraEditors.SimpleButton btSubir;
		private DevExpress.XtraEditors.LookUpEdit cmbTabla;
		private DevExpress.XtraEditors.LookUpEdit cmbColumnas;
		private DevExpress.XtraTab.XtraTabPage xtraTabPage3;
		private DevExpress.XtraGrid.Views.Grid.GridView grvExcel;
		private System.Windows.Forms.OpenFileDialog openFileDialog;
		private DevExpress.XtraEditors.CheckEdit chkExcel2007;
		private C1.Data.C1DataSet cdsSeteoF;
		private DevExpress.XtraEditors.DateEdit cmbDesde;
		private DevExpress.XtraEditors.DateEdit cmbHasta;
		private DevExpress.XtraEditors.LabelControl labelControl4;
		private DevExpress.XtraEditors.LabelControl labelControl5;
		private DevExpress.XtraEditors.LabelControl labelControl6;
		private DevExpress.XtraEditors.SimpleButton btFiltro;
		private DevExpress.XtraGrid.GridControl grdErrores;
		private DevExpress.XtraGrid.Views.Grid.GridView grvErrores;
		private DevExpress.XtraEditors.SpinEdit spnIdSubida;
		private DevExpress.XtraEditors.CheckEdit chkEditarRegistros;
		private DevExpress.XtraEditors.CheckEdit chkMaestroDetalle;
		private DevExpress.XtraEditors.CheckEdit chkNoSubirNoExisteSecundario;
		private DevExpress.XtraEditors.CheckEdit chkCrearSecundarios;
		private DevExpress.XtraEditors.CheckEdit chkSubirVacios;
		private DevExpress.XtraEditors.CheckEdit chkBorrarDetalle;
		private DevExpress.XtraEditors.GroupControl groupControl1;
		private DevExpress.XtraEditors.GroupControl groupControl2;
		private DevExpress.XtraEditors.GroupControl groupControl3;
		private DevExpress.XtraGrid.GridControl grdExcel;
		/// <summary>
		/// Required designer variable.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public ImportaBlasar()
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
			this.xtraTabControl1 = new DevExpress.XtraTab.XtraTabControl();
			this.xtraTabPage1 = new DevExpress.XtraTab.XtraTabPage();
			this.btSubir = new DevExpress.XtraEditors.SimpleButton();
			this.btLeer = new DevExpress.XtraEditors.SimpleButton();
			this.cmbColumnas = new DevExpress.XtraEditors.LookUpEdit();
			this.cmbTabla = new DevExpress.XtraEditors.LookUpEdit();
			this.grdExcel = new DevExpress.XtraGrid.GridControl();
			this.grvExcel = new DevExpress.XtraGrid.Views.Grid.GridView();
			this.btArchivo = new DevExpress.XtraEditors.SimpleButton();
			this.txtArchivo = new DevExpress.XtraEditors.TextEdit();
			this.labelControl3 = new DevExpress.XtraEditors.LabelControl();
			this.labelControl2 = new DevExpress.XtraEditors.LabelControl();
			this.labelControl1 = new DevExpress.XtraEditors.LabelControl();
			this.xtraTabPage2 = new DevExpress.XtraTab.XtraTabPage();
			this.grdErrores = new DevExpress.XtraGrid.GridControl();
			this.grvErrores = new DevExpress.XtraGrid.Views.Grid.GridView();
			this.btFiltro = new DevExpress.XtraEditors.SimpleButton();
			this.spnIdSubida = new DevExpress.XtraEditors.SpinEdit();
			this.labelControl6 = new DevExpress.XtraEditors.LabelControl();
			this.labelControl5 = new DevExpress.XtraEditors.LabelControl();
			this.labelControl4 = new DevExpress.XtraEditors.LabelControl();
			this.cmbHasta = new DevExpress.XtraEditors.DateEdit();
			this.cmbDesde = new DevExpress.XtraEditors.DateEdit();
			this.xtraTabPage3 = new DevExpress.XtraTab.XtraTabPage();
			this.groupControl3 = new DevExpress.XtraEditors.GroupControl();
			this.chkMaestroDetalle = new DevExpress.XtraEditors.CheckEdit();
			this.chkBorrarDetalle = new DevExpress.XtraEditors.CheckEdit();
			this.groupControl2 = new DevExpress.XtraEditors.GroupControl();
			this.chkCrearSecundarios = new DevExpress.XtraEditors.CheckEdit();
			this.chkNoSubirNoExisteSecundario = new DevExpress.XtraEditors.CheckEdit();
			this.groupControl1 = new DevExpress.XtraEditors.GroupControl();
			this.chkSubirVacios = new DevExpress.XtraEditors.CheckEdit();
			this.chkEditarRegistros = new DevExpress.XtraEditors.CheckEdit();
			this.chkExcel2007 = new DevExpress.XtraEditors.CheckEdit();
			this.openFileDialog = new System.Windows.Forms.OpenFileDialog();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			((System.ComponentModel.ISupportInitialize)(this.xtraTabControl1)).BeginInit();
			this.xtraTabControl1.SuspendLayout();
			this.xtraTabPage1.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.cmbColumnas.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbTabla.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.grdExcel)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.grvExcel)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtArchivo.Properties)).BeginInit();
			this.xtraTabPage2.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.grdErrores)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.grvErrores)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.spnIdSubida.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbHasta.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbDesde.Properties)).BeginInit();
			this.xtraTabPage3.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.groupControl3)).BeginInit();
			this.groupControl3.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.chkMaestroDetalle.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.chkBorrarDetalle.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.groupControl2)).BeginInit();
			this.groupControl2.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.chkCrearSecundarios.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.chkNoSubirNoExisteSecundario.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.groupControl1)).BeginInit();
			this.groupControl1.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.chkSubirVacios.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.chkEditarRegistros.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.chkExcel2007.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			this.SuspendLayout();
			// 
			// xtraTabControl1
			// 
			this.xtraTabControl1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.xtraTabControl1.Location = new System.Drawing.Point(8, 8);
			this.xtraTabControl1.Name = "xtraTabControl1";
			this.xtraTabControl1.SelectedTabPage = this.xtraTabPage1;
			this.xtraTabControl1.Size = new System.Drawing.Size(768, 384);
			this.xtraTabControl1.TabIndex = 0;
			this.xtraTabControl1.TabPages.AddRange(new DevExpress.XtraTab.XtraTabPage[] {
																																										this.xtraTabPage1,
																																										this.xtraTabPage2,
																																										this.xtraTabPage3});
			this.xtraTabControl1.Text = "xtraTabControl1";
			// 
			// xtraTabPage1
			// 
			this.xtraTabPage1.Controls.Add(this.btSubir);
			this.xtraTabPage1.Controls.Add(this.btLeer);
			this.xtraTabPage1.Controls.Add(this.cmbColumnas);
			this.xtraTabPage1.Controls.Add(this.cmbTabla);
			this.xtraTabPage1.Controls.Add(this.grdExcel);
			this.xtraTabPage1.Controls.Add(this.btArchivo);
			this.xtraTabPage1.Controls.Add(this.txtArchivo);
			this.xtraTabPage1.Controls.Add(this.labelControl3);
			this.xtraTabPage1.Controls.Add(this.labelControl2);
			this.xtraTabPage1.Controls.Add(this.labelControl1);
			this.xtraTabPage1.Name = "xtraTabPage1";
			this.xtraTabPage1.Size = new System.Drawing.Size(759, 354);
			this.xtraTabPage1.Text = "Excel";
			// 
			// btSubir
			// 
			this.btSubir.Location = new System.Drawing.Point(496, 40);
			this.btSubir.Name = "btSubir";
			this.btSubir.Size = new System.Drawing.Size(48, 23);
			this.btSubir.TabIndex = 9;
			this.btSubir.Text = "Subir";
			this.btSubir.Click += new System.EventHandler(this.btSubir_Click);
			// 
			// btLeer
			// 
			this.btLeer.Location = new System.Drawing.Point(496, 16);
			this.btLeer.Name = "btLeer";
			this.btLeer.Size = new System.Drawing.Size(48, 23);
			this.btLeer.TabIndex = 8;
			this.btLeer.Text = "Leer";
			this.btLeer.Click += new System.EventHandler(this.btLeer_Click);
			// 
			// cmbColumnas
			// 
			this.cmbColumnas.Location = new System.Drawing.Point(312, 40);
			this.cmbColumnas.Name = "cmbColumnas";
			// 
			// cmbColumnas.Properties
			// 
			this.cmbColumnas.Properties.Buttons.AddRange(new DevExpress.XtraEditors.Controls.EditorButton[] {
																																																				new DevExpress.XtraEditors.Controls.EditorButton(DevExpress.XtraEditors.Controls.ButtonPredefines.Combo)});
			this.cmbColumnas.Properties.DisplayMember = "Columnas";
			this.cmbColumnas.Properties.NullText = "";
			this.cmbColumnas.Size = new System.Drawing.Size(168, 20);
			this.cmbColumnas.TabIndex = 7;
			// 
			// cmbTabla
			// 
			this.cmbTabla.Location = new System.Drawing.Point(72, 40);
			this.cmbTabla.Name = "cmbTabla";
			// 
			// cmbTabla.Properties
			// 
			this.cmbTabla.Properties.Buttons.AddRange(new DevExpress.XtraEditors.Controls.EditorButton[] {
																																																		 new DevExpress.XtraEditors.Controls.EditorButton(DevExpress.XtraEditors.Controls.ButtonPredefines.Combo)});
			this.cmbTabla.Properties.Columns.AddRange(new DevExpress.XtraEditors.Controls.LookUpColumnInfo[] {
																																																				 new DevExpress.XtraEditors.Controls.LookUpColumnInfo("Tabla", "Tablas", 20, DevExpress.Utils.FormatType.None, "", true, DevExpress.Utils.HorzAlignment.Default, DevExpress.Data.ColumnSortOrder.Ascending)});
			this.cmbTabla.Properties.DisplayMember = "Tabla";
			this.cmbTabla.Properties.ValueMember = "id";
			this.cmbTabla.Size = new System.Drawing.Size(168, 20);
			this.cmbTabla.TabIndex = 6;
			this.cmbTabla.Validated += new System.EventHandler(this.cmbTabla_Validated);
			// 
			// grdExcel
			// 
			this.grdExcel.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			// 
			// grdExcel.EmbeddedNavigator
			// 
			this.grdExcel.EmbeddedNavigator.Name = "";
			this.grdExcel.Location = new System.Drawing.Point(16, 72);
			this.grdExcel.MainView = this.grvExcel;
			this.grdExcel.Name = "grdExcel";
			this.grdExcel.Size = new System.Drawing.Size(736, 273);
			this.grdExcel.TabIndex = 5;
			this.grdExcel.ViewCollection.AddRange(new DevExpress.XtraGrid.Views.Base.BaseView[] {
																																														this.grvExcel});
			// 
			// grvExcel
			// 
			this.grvExcel.GridControl = this.grdExcel;
			this.grvExcel.Name = "grvExcel";
			this.grvExcel.OptionsView.ShowAutoFilterRow = true;
			// 
			// btArchivo
			// 
			this.btArchivo.Location = new System.Drawing.Point(456, 16);
			this.btArchivo.Name = "btArchivo";
			this.btArchivo.Size = new System.Drawing.Size(24, 23);
			this.btArchivo.TabIndex = 4;
			this.btArchivo.Text = "...";
			this.btArchivo.Click += new System.EventHandler(this.btArchivo_Click);
			// 
			// txtArchivo
			// 
			this.txtArchivo.Location = new System.Drawing.Point(72, 16);
			this.txtArchivo.Name = "txtArchivo";
			this.txtArchivo.Size = new System.Drawing.Size(376, 20);
			this.txtArchivo.TabIndex = 3;
			// 
			// labelControl3
			// 
			this.labelControl3.Location = new System.Drawing.Point(24, 40);
			this.labelControl3.Name = "labelControl3";
			this.labelControl3.Size = new System.Drawing.Size(40, 13);
			this.labelControl3.TabIndex = 2;
			this.labelControl3.Text = "Tablas";
			// 
			// labelControl2
			// 
			this.labelControl2.Location = new System.Drawing.Point(256, 40);
			this.labelControl2.Name = "labelControl2";
			this.labelControl2.Size = new System.Drawing.Size(48, 13);
			this.labelControl2.TabIndex = 1;
			this.labelControl2.Text = "Columnas";
			// 
			// labelControl1
			// 
			this.labelControl1.Location = new System.Drawing.Point(24, 16);
			this.labelControl1.Name = "labelControl1";
			this.labelControl1.Size = new System.Drawing.Size(40, 13);
			this.labelControl1.TabIndex = 0;
			this.labelControl1.Text = "Archivo";
			// 
			// xtraTabPage2
			// 
			this.xtraTabPage2.Controls.Add(this.grdErrores);
			this.xtraTabPage2.Controls.Add(this.btFiltro);
			this.xtraTabPage2.Controls.Add(this.spnIdSubida);
			this.xtraTabPage2.Controls.Add(this.labelControl6);
			this.xtraTabPage2.Controls.Add(this.labelControl5);
			this.xtraTabPage2.Controls.Add(this.labelControl4);
			this.xtraTabPage2.Controls.Add(this.cmbHasta);
			this.xtraTabPage2.Controls.Add(this.cmbDesde);
			this.xtraTabPage2.Name = "xtraTabPage2";
			this.xtraTabPage2.Size = new System.Drawing.Size(759, 353);
			this.xtraTabPage2.Text = "Mensajes";
			// 
			// grdErrores
			// 
			this.grdErrores.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			// 
			// grdErrores.EmbeddedNavigator
			// 
			this.grdErrores.EmbeddedNavigator.Name = "";
			this.grdErrores.Location = new System.Drawing.Point(8, 48);
			this.grdErrores.MainView = this.grvErrores;
			this.grdErrores.Name = "grdErrores";
			this.grdErrores.Size = new System.Drawing.Size(736, 296);
			this.grdErrores.TabIndex = 7;
			this.grdErrores.ViewCollection.AddRange(new DevExpress.XtraGrid.Views.Base.BaseView[] {
																																															this.grvErrores});
			// 
			// grvErrores
			// 
			this.grvErrores.GridControl = this.grdErrores;
			this.grvErrores.Name = "grvErrores";
			// 
			// btFiltro
			// 
			this.btFiltro.Location = new System.Drawing.Point(544, 16);
			this.btFiltro.Name = "btFiltro";
			this.btFiltro.TabIndex = 6;
			this.btFiltro.Text = "Filtro";
			this.btFiltro.Click += new System.EventHandler(this.btFiltro_Click);
			// 
			// spnIdSubida
			// 
			this.spnIdSubida.EditValue = new System.Decimal(new int[] {
																																	0,
																																	0,
																																	0,
																																	0});
			this.spnIdSubida.Location = new System.Drawing.Point(424, 16);
			this.spnIdSubida.Name = "spnIdSubida";
			// 
			// spnIdSubida.Properties
			// 
			this.spnIdSubida.Properties.Buttons.AddRange(new DevExpress.XtraEditors.Controls.EditorButton[] {
																																																				new DevExpress.XtraEditors.Controls.EditorButton()});
			this.spnIdSubida.TabIndex = 5;
			// 
			// labelControl6
			// 
			this.labelControl6.Location = new System.Drawing.Point(368, 16);
			this.labelControl6.Name = "labelControl6";
			this.labelControl6.Size = new System.Drawing.Size(40, 13);
			this.labelControl6.TabIndex = 4;
			this.labelControl6.Text = "Subida";
			// 
			// labelControl5
			// 
			this.labelControl5.Location = new System.Drawing.Point(184, 16);
			this.labelControl5.Name = "labelControl5";
			this.labelControl5.Size = new System.Drawing.Size(40, 13);
			this.labelControl5.TabIndex = 3;
			this.labelControl5.Text = "Hasta";
			// 
			// labelControl4
			// 
			this.labelControl4.Location = new System.Drawing.Point(16, 16);
			this.labelControl4.Name = "labelControl4";
			this.labelControl4.Size = new System.Drawing.Size(40, 13);
			this.labelControl4.TabIndex = 2;
			this.labelControl4.Text = "Desde";
			// 
			// cmbHasta
			// 
			this.cmbHasta.EditValue = null;
			this.cmbHasta.Location = new System.Drawing.Point(232, 16);
			this.cmbHasta.Name = "cmbHasta";
			// 
			// cmbHasta.Properties
			// 
			this.cmbHasta.Properties.Buttons.AddRange(new DevExpress.XtraEditors.Controls.EditorButton[] {
																																																		 new DevExpress.XtraEditors.Controls.EditorButton(DevExpress.XtraEditors.Controls.ButtonPredefines.Combo)});
			this.cmbHasta.TabIndex = 1;
			// 
			// cmbDesde
			// 
			this.cmbDesde.EditValue = null;
			this.cmbDesde.Location = new System.Drawing.Point(64, 16);
			this.cmbDesde.Name = "cmbDesde";
			// 
			// cmbDesde.Properties
			// 
			this.cmbDesde.Properties.Buttons.AddRange(new DevExpress.XtraEditors.Controls.EditorButton[] {
																																																		 new DevExpress.XtraEditors.Controls.EditorButton(DevExpress.XtraEditors.Controls.ButtonPredefines.Combo)});
			this.cmbDesde.TabIndex = 0;
			// 
			// xtraTabPage3
			// 
			this.xtraTabPage3.Controls.Add(this.groupControl3);
			this.xtraTabPage3.Controls.Add(this.groupControl2);
			this.xtraTabPage3.Controls.Add(this.groupControl1);
			this.xtraTabPage3.Name = "xtraTabPage3";
			this.xtraTabPage3.Size = new System.Drawing.Size(759, 353);
			this.xtraTabPage3.Text = "Opciones";
			// 
			// groupControl3
			// 
			this.groupControl3.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.groupControl3.Controls.Add(this.chkMaestroDetalle);
			this.groupControl3.Controls.Add(this.chkBorrarDetalle);
			this.groupControl3.Location = new System.Drawing.Point(16, 240);
			this.groupControl3.Name = "groupControl3";
			this.groupControl3.Size = new System.Drawing.Size(728, 88);
			this.groupControl3.TabIndex = 10;
			this.groupControl3.Text = "Maestro Detalle";
			// 
			// chkMaestroDetalle
			// 
			this.chkMaestroDetalle.Location = new System.Drawing.Point(16, 32);
			this.chkMaestroDetalle.Name = "chkMaestroDetalle";
			// 
			// chkMaestroDetalle.Properties
			// 
			this.chkMaestroDetalle.Properties.Caption = "Maestro Detalle";
			this.chkMaestroDetalle.Size = new System.Drawing.Size(176, 18);
			this.chkMaestroDetalle.TabIndex = 2;
			// 
			// chkBorrarDetalle
			// 
			this.chkBorrarDetalle.Location = new System.Drawing.Point(16, 56);
			this.chkBorrarDetalle.Name = "chkBorrarDetalle";
			// 
			// chkBorrarDetalle.Properties
			// 
			this.chkBorrarDetalle.Properties.Caption = "Borrar Tabla de Detalle al Subir";
			this.chkBorrarDetalle.Size = new System.Drawing.Size(240, 18);
			this.chkBorrarDetalle.TabIndex = 6;
			// 
			// groupControl2
			// 
			this.groupControl2.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.groupControl2.Controls.Add(this.chkCrearSecundarios);
			this.groupControl2.Controls.Add(this.chkNoSubirNoExisteSecundario);
			this.groupControl2.Location = new System.Drawing.Point(16, 136);
			this.groupControl2.Name = "groupControl2";
			this.groupControl2.Size = new System.Drawing.Size(728, 88);
			this.groupControl2.TabIndex = 9;
			this.groupControl2.Text = "Tablas Secundarias";
			// 
			// chkCrearSecundarios
			// 
			this.chkCrearSecundarios.EditValue = true;
			this.chkCrearSecundarios.Location = new System.Drawing.Point(16, 32);
			this.chkCrearSecundarios.Name = "chkCrearSecundarios";
			// 
			// chkCrearSecundarios.Properties
			// 
			this.chkCrearSecundarios.Properties.Caption = "Crear Registros Secundarios si no Existen";
			this.chkCrearSecundarios.Size = new System.Drawing.Size(240, 18);
			this.chkCrearSecundarios.TabIndex = 4;
			// 
			// chkNoSubirNoExisteSecundario
			// 
			this.chkNoSubirNoExisteSecundario.Location = new System.Drawing.Point(16, 56);
			this.chkNoSubirNoExisteSecundario.Name = "chkNoSubirNoExisteSecundario";
			// 
			// chkNoSubirNoExisteSecundario.Properties
			// 
			this.chkNoSubirNoExisteSecundario.Properties.Caption = "No subir si no existe registro secundario";
			this.chkNoSubirNoExisteSecundario.Size = new System.Drawing.Size(240, 18);
			this.chkNoSubirNoExisteSecundario.TabIndex = 3;
			// 
			// groupControl1
			// 
			this.groupControl1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.groupControl1.Controls.Add(this.chkSubirVacios);
			this.groupControl1.Controls.Add(this.chkEditarRegistros);
			this.groupControl1.Controls.Add(this.chkExcel2007);
			this.groupControl1.Location = new System.Drawing.Point(16, 8);
			this.groupControl1.Name = "groupControl1";
			this.groupControl1.Size = new System.Drawing.Size(728, 112);
			this.groupControl1.TabIndex = 8;
			this.groupControl1.Text = "Generales";
			// 
			// chkSubirVacios
			// 
			this.chkSubirVacios.EditValue = true;
			this.chkSubirVacios.Location = new System.Drawing.Point(16, 56);
			this.chkSubirVacios.Name = "chkSubirVacios";
			// 
			// chkSubirVacios.Properties
			// 
			this.chkSubirVacios.Properties.Caption = "Subir celdas que no tengan todos los campos llenos";
			this.chkSubirVacios.Size = new System.Drawing.Size(280, 18);
			this.chkSubirVacios.TabIndex = 5;
			// 
			// chkEditarRegistros
			// 
			this.chkEditarRegistros.Location = new System.Drawing.Point(16, 32);
			this.chkEditarRegistros.Name = "chkEditarRegistros";
			// 
			// chkEditarRegistros.Properties
			// 
			this.chkEditarRegistros.Properties.Caption = "Modificar Registros Existentes";
			this.chkEditarRegistros.Size = new System.Drawing.Size(176, 18);
			this.chkEditarRegistros.TabIndex = 1;
			// 
			// chkExcel2007
			// 
			this.chkExcel2007.Location = new System.Drawing.Point(16, 80);
			this.chkExcel2007.Name = "chkExcel2007";
			// 
			// chkExcel2007.Properties
			// 
			this.chkExcel2007.Properties.Caption = "Archivos de Excel 2007";
			this.chkExcel2007.Size = new System.Drawing.Size(176, 18);
			this.chkExcel2007.TabIndex = 0;
			// 
			// openFileDialog
			// 
			this.openFileDialog.DefaultExt = "xls";
			this.openFileDialog.FileName = "c:\\Latinium\\Articulo.xls";
			this.openFileDialog.Filter = "Excel (*.xls*)|*.xls;*.xlsx|Excel 2007 (*.xlsx)|*.xlsx|Excel 2003 (*.xls)|*.xls";
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
			// ImportaBlasar
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 14);
			this.ClientSize = new System.Drawing.Size(784, 406);
			this.Controls.Add(this.xtraTabControl1);
			this.Name = "ImportaBlasar";
			this.Text = "Importar desde Excel";
			this.Load += new System.EventHandler(this.ImportaBlasar_Load);
			((System.ComponentModel.ISupportInitialize)(this.xtraTabControl1)).EndInit();
			this.xtraTabControl1.ResumeLayout(false);
			this.xtraTabPage1.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.cmbColumnas.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbTabla.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.grdExcel)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.grvExcel)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtArchivo.Properties)).EndInit();
			this.xtraTabPage2.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.grdErrores)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.grvErrores)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.spnIdSubida.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbHasta.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbDesde.Properties)).EndInit();
			this.xtraTabPage3.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.groupControl3)).EndInit();
			this.groupControl3.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.chkMaestroDetalle.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.chkBorrarDetalle.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.groupControl2)).EndInit();
			this.groupControl2.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.chkCrearSecundarios.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.chkNoSubirNoExisteSecundario.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.groupControl1)).EndInit();
			this.groupControl1.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.chkSubirVacios.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.chkEditarRegistros.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.chkExcel2007.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void btArchivo_Click(object sender, System.EventArgs e)
		{
			openFileDialog.InitialDirectory = Application.StartupPath;
			if (chkExcel2007.Checked) openFileDialog.FilterIndex = 2;
			else openFileDialog.FilterIndex = 1;
			if (openFileDialog.ShowDialog() != DialogResult.OK) return;
			txtArchivo.Text = openFileDialog.FileName;
		}

		private void btLeer_Click(object sender, System.EventArgs e)
		{
			if (!File.Exists(txtArchivo.Text))
			{
				MessageBox.Show("No existe archivo " + txtArchivo.Text, "Informacion");
				return;
			}
			string strConn = "Provider=Microsoft.Jet.OLEDB.4.0;" +
				"Data Source=" + txtArchivo.Text + ";Extended Properties='Excel 8.0;HDR=Yes;IMEX=1;'";
			if (chkExcel2007.Checked)
				strConn = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + txtArchivo.Text
					+ ";Extended Properties=Excel 12.0";
			OleDbDataAdapter dataAdapter = new OleDbDataAdapter("SELECT * FROM [Hoja1$]", strConn);
			DataSet myDataSet = new DataSet();
			try
			{
				dataAdapter.Fill(myDataSet, "ExcelInfo");
				DataTable dataTable = myDataSet.Tables["ExcelInfo"];
				grdExcel.DataSource = (DataView)dataTable.DefaultView;
				grdExcel.MainView.PopulateColumns();
			}
			catch (Exception ex)
			{
				MessageBox.Show(ex.Message, "Error al Leer archivo de Excel");
				return;
			}

			string stTitulos = "";
			string stColNoExiste = "";
			for (int i = 0; i < grvExcel.Columns.Count; i++)
			{
				string stCampo = grvExcel.Columns[i].FieldName.ToString().Replace(" ", "");
				if (stCampo.IndexOf("-") > 0)
				{
					int iLong = stCampo.IndexOf("-", 0);
					int iFin = stCampo.IndexOf("-", iLong + 1);
					if (iFin < 0)
					{
						MessageBox.Show("Error en columna " + stCampo, "Informacion");
						continue;
					}
					string stTab = stCampo.Substring(0, iLong);
					string stCamp = stCampo.Substring(iLong + 1, iFin - iLong -1);
					string stExec1 = String.Format("Exec VerificaCampos '{0}', '{1}'", 
						stTab, stCamp);
					stColNoExiste += Funcion.sEscalarSQL(cdsSeteoF, stExec1);
				}
				else
					stTitulos += stCampo + (i == grvExcel.Columns.Count - 1 ? "" : ",");
			}
			string stTabla = cmbTabla.Text;
			string stExec = String.Format("Exec VerificaCampos '{0}', '{1}'", stTabla, stTitulos);
			stColNoExiste += Funcion.sEscalarSQL(cdsSeteoF, stExec);
			if (stColNoExiste.Trim().Length > 0)
				MessageBox.Show(String.Format("En la Tabla {0} No existe: {1}", stTabla, stColNoExiste), "Informacion");
		}

		private void btFiltro_Click(object sender, System.EventArgs e)
		{
			int IdSubida = int.Parse(spnIdSubida.EditValue.ToString());
			DateTime dtDesde = (DateTime)cmbDesde.EditValue;
			DateTime dtHasta = (DateTime)cmbHasta.EditValue;
			string stExec = String.Format("Exec ImportaErrorMostrar {0}, '{1}', '{2}'",
				IdSubida, dtDesde.ToString("yyyyMMdd HH:mm:ss"), dtHasta.ToString("yyyyMMdd HH:mm:ss"));
			grdErrores.DataSource = Funcion.dvProcedimiento(cdsSeteoF, stExec);

		}

		private void ImportaBlasar_Load(object sender, System.EventArgs e)
		{
			cmbDesde.EditValue = DateTime.Now;
			cmbHasta.EditValue = DateTime.Now;
			int idSubida = Funcion.iEscalarSQL(cdsSeteoF, "Select Count(*) From ImportaError");
			if (idSubida > 0)
				idSubida = Funcion.iEscalarSQL(cdsSeteoF, "Select Max(idSubida) From ImportaError");
			spnIdSubida.EditValue = idSubida;
			cmbTabla.Properties.DataSource = Funcion.dvProcedimiento(cdsSeteoF,
				"Select s.name as Tabla, s.id from sysobjects s where xtype = 'U' Order by name");
//			if (cmbTabla.Text.Trim().Length == 0)
				cmbTabla.Text = "Articulo";
			cmbTabla.Select();
		}

		private void cmbTabla_Validated(object sender, System.EventArgs e)
		{
			string stSelect = String.Format("Select s.name as Columnas From SysColumns s where s.id = {0} Order by s.name", 
				cmbTabla.EditValue);
			cmbColumnas.Properties.DataSource = Funcion.dvProcedimiento(cdsSeteoF, stSelect);

		}

		CultureInfo us = new CultureInfo("en-US");
		private string ObjetoToTexto(object o)
		{
			string stValor = "";
			if (o == null) return "null";
			if (o.GetType().Name == "String") stValor = String.Format("'{0}'", o.ToString().Replace("'", ""));
			if (o.GetType().Name == "Boolean") stValor = (bool)o == true ? "1" : "0";
			if (o.GetType().Name == "Double") stValor = ((double)o).ToString("0.000000", us);
			if (o.GetType().Name == "DateTime") stValor = String.Format("'{0:yyyyMMdd}'", (DateTime)o);
			return stValor;
		}

		private void btSubir_Click(object sender, System.EventArgs e)
		{
			DateTime dtInicio = DateTime.Now;
			Cursor = Cursors.WaitCursor;
			#region Verificaciones
			if (grvExcel.Columns.Count == 0)
			{
				MessageBox.Show("No existe columnas para subir", "Informacion");
				return;
			}
			#endregion Verificaciones
			#region Lee Columnas y Genera Insert
			string[] stColumnas = new string[grvExcel.Columns.Count];
			string[] stColumnasExt = new string[grvExcel.Columns.Count];
			string stTabla = cmbTabla.Text;
			string stInserta = "";
			for (int i = 0; i < grvExcel.Columns.Count; i++)
			{
				string stCampo = grvExcel.Columns[i].FieldName.Replace(" ", "");
				if (stCampo.IndexOf("-") > 0)
				{
					int iLong = stCampo.IndexOf("-", 0);
					int iFin = stCampo.IndexOf("-", iLong + 1);
					string stCamp = stCampo.Substring(iLong + 1, iFin - iLong - 1);
					stColumnas[i] = stCamp;
				}
				else
					stColumnas[i] = stCampo;

				stColumnasExt[i] = stCampo;
				stInserta += stColumnas[i] + ",";
			}
			stInserta = stInserta.Remove(stInserta.Length - 1,1);
			string stExec = String.Format("Exec VerificaTiposCampo '{0}', '{1}'", stTabla, stInserta);
			string[] stTipos = new string[grvExcel.Columns.Count];
			SqlDataReader dr = Funcion.rEscalarSQL(cdsSeteoF, stExec, true);
			int i1=0;
			while (dr.Read())
			{
				stTipos[i1++] = dr.GetString(0);
			}
			dr.Close();

			string stGrupoColumnas = stInserta;
			int idSubida = Funcion.iEscalarSQL(cdsSeteoF, "Select Count(*) From ImportaError");
			if (idSubida > 0)
				idSubida = Funcion.iEscalarSQL(cdsSeteoF, "Select Max(idSubida) From ImportaError");
			idSubida++;
			#endregion Lee Columnas
			#region Lectura de Valores e Insercion en SQL
			object vCelda=null;
			for (int i = 0; i < grvExcel.DataRowCount; i++)
			{
				stInserta = String.Format("Insert Into {0} ({1})", stTabla, stGrupoColumnas);
				#region Primera Columna Clave Primaria
				string stClaveP = ObjetoToTexto(grvExcel.GetRowCellValue(i, stColumnas[0]));
				if (stClaveP.Length == 0)
				{
					string stError = String.Format("Insert Into ImportaError (Fila, Columna, Mensaje, idSubida, Tipo) Values ({0}, '{1}', '{2}', {3}, '{4}')",
						i + 2, "A", "Primera columna no puede estar vacia en Excel", idSubida, "Falta Clave");
					Funcion.EjecutaSQL(cdsSeteoF, stError);
					continue;
				}
				string stVerifica1 = String.Format("Select Count(*) From {0} Where {1} = '{2}'",
					stTabla, stColumnas[0], stClaveP.Replace("'", ""));
				int iNumRep = 0;
				if (chkEditarRegistros.Checked) iNumRep = 1;
				bool bExiste = false;
				int iCuentaRep = Funcion.iEscalarSQL(cdsSeteoF, stVerifica1);
				if (iCuentaRep > iNumRep)
				{
					string stExiste = " YA Existe";
					if (iNumRep == 1) stExiste = " Existe MAS de una vez";
					string stError = String.Format("Insert Into ImportaError (Fila, Columna, Mensaje, idSubida, Tipo) Values ({0}, '{1}', '{2}', {3}, '{4}')",
						i + 2, "A", String.Format("En Tabla {0} Campo {1}{2}", stTabla, stClaveP.Replace("'", ""), stExiste), idSubida, "Clave ya Existe");
					Funcion.EjecutaSQL(cdsSeteoF, stError);
					continue;
				}
				if (iCuentaRep == 1)
				{
					stInserta = String.Format("Update {0} Set ", stTabla);
					bExiste = true;
				}
				#endregion Primera Columna Clave Primaria
				#region Segunda Columna Clave Primaria del Maestro
				string stClaveM = "";
				if (chkMaestroDetalle.Checked)
				{
					stClaveM = ObjetoToTexto(grvExcel.GetRowCellValue(i, stColumnas[1]));
					if (stClaveM.Length == 0)
					{
						string stError = String.Format("Insert Into ImportaError (Fila, Columna, Mensaje, idSubida, Tipo) Values ({0}, '{1}', '{2}', {3}, '{4}')",
							i + 2, "B", "Maestro detalle necesita 2 columnas en Excel", idSubida, "Falta Clave Maestro");
						Funcion.EjecutaSQL(cdsSeteoF, stError);
						continue;
					}
					string stTablaM = "Compra";
					if (cmbTabla.Text.ToUpper() == "ASIENTODETALLE")
					{
						stTablaM = "Asiento";
					}
					string stVerificaM = String.Format("Select Count(*) From {0} Where {1} = {2}",
						stTablaM, stColumnas[1], stClaveM);
					int iCuentaRepM = Funcion.iEscalarSQL(cdsSeteoF, stVerifica1);
					if (iCuentaRepM != 1)
					{
						string stError = String.Format("Insert Into ImportaError (Fila, Columna, Mensaje, idSubida, Tipo) Values ({0}, '{1}', '{2}', {3}, '{4}')",
							i + 2, "A", String.Format("En Tabla {0} Campo {1}{2}", stTablaM, stClaveM.Replace("'", ""), " Error en {0}"), idSubida, "Clave ya Existe");
						Funcion.EjecutaSQL(cdsSeteoF, stError);
						continue;
					}
					//          stInserta = stInserta.Remove(stInserta.Length - 2) + stColumnas[1] + ")";
					stInserta = String.Format("Update {0} Set ", stTabla, stColumnas[1]);
				}
				#endregion Primera Columna Clave Primaria
				#region Lectura de Celdas y generacion de instruccion SQL
				string stInsertaFila = stInserta;
				string stValor = "";
				if (!bExiste) stValor = "Values (";
				bool bError = false;
				for (int j = 0; j<stColumnas.Length; j++)
				{
					vCelda = grvExcel.GetRowCellValue(i, stColumnasExt[j]);
					#region Registros de otras Tablas
					if (stColumnasExt[j].IndexOf("-") > 0) // Es de otra tabla
					{
						vCelda = vCelda.ToString().Replace("'", "\"");
						string stExec1 = String.Format("Exec ImportaExterno '{0}', '{1}', {2}", 
							stColumnasExt[j], vCelda.ToString().Trim(), (chkCrearSecundarios.Checked == true? 1: 0));
						vCelda = (double) Funcion.iEscalarSQL(cdsSeteoF, stExec1);
						if ((double)vCelda == 0.00 && chkNoSubirNoExisteSecundario.Checked)
						{
							bError = true;
							string stError = String.Format("Insert Into ImportaError (Fila, Columna, Mensaje, idSubida, Tipo) Values ({0}, '{1}', '{2}', {3}, '{4}')",
								i + 2, (char)(j + 65), stTipos[j], idSubida, "Secundario Inexistente");
							Funcion.EjecutaSQL(cdsSeteoF, stError);
						}
					}
					#endregion Registros de otras Tablas
					if (vCelda.GetType().Name == "DBNull" || stTipos[j] == "No Soportado" || stTipos[j] == "Desconocido")
					{
						stInsertaFila = stInsertaFila.Replace(stColumnas[j] + ",", "");
						stInsertaFila = stInsertaFila.Replace("," + stColumnas[j], "");
						if (!chkSubirVacios.Checked) bError = true;
					}
					else
					{
						if (vCelda.GetType().Name == stTipos[j] || (vCelda.GetType().Name == "Double" && stTipos[j] == "String"))
						{
							string stCelda = ObjetoToTexto(vCelda);
							if (vCelda.GetType().Name == "Double" && stTipos[j] == "String")
								stCelda = String.Format("'{0}'", (double)vCelda);
							if (bExiste)
								stValor += String.Format("{0} = {1}, ", stColumnas[j], stCelda);
							else
								stValor += stCelda.Trim() + ", ";
						}
						else
						{
							bError = true;
							string stError = String.Format("Insert Into ImportaError (Fila, Columna, Mensaje, idSubida, Tipo) Values ({0}, '{1}', '{2}', {3}, '{4}')",
								i + 2, (char)(j + 65), stTipos[j], idSubida, "Tipo de Dato");
							Funcion.EjecutaSQL(cdsSeteoF, stError);
						}
					}
				}
				#endregion Lectura de Celdas y generacion de instruccion SQL
				#region Insercion de celdas
				if (!bError)
				{
					string stParentesis = ")";
					if (stClaveM.Length > 0)
						stParentesis = String.Format(", {0})", stClaveM.Replace("'", ""));

					if (bExiste) stParentesis = String.Format(" Where {0} = {1}", stColumnas[0], stClaveP);
					stValor = String.Format("{0}{1}{2} ", stInsertaFila, stValor.Remove(stValor.Length - 2, 2), stParentesis);
					try
					{
						Funcion.EjecutaSQL(cdsSeteoF, stValor);
					}
					catch(System.Exception ex)
					{
						string stError = String.Format("Insert Into ImportaError (Fila, Mensaje, idSubida, Tipo) Values ({0}, '{1}', {2}, '{3}')",
							i + 2, ex.Message.Replace("'", "") + stValor.Replace("'", "`"), idSubida, "Grabar");
						Funcion.EjecutaSQL(cdsSeteoF, stError);
					}
				}
				#endregion Insercion de celdas
			}
			#endregion Lectura de Valores
			#region Muestra los Mensajes o termina
			Funcion.EjecutaSQL(cdsSeteoF, "Exec ImportaTermina " + stTabla);
			int iErrores = Funcion.iEscalarSQL(cdsSeteoF, "Select Count(*) From ImportaError Where idSubida = " + idSubida);
			Cursor = Cursors.Default;
			TimeSpan ts = DateTime.Now.Subtract(dtInicio);
			if (iErrores > 0)
			{
				MessageBox.Show(String.Format("Tiempo: {0}:{1}. Existen {2} errores en la subida de datos. Haga clic en mensajes y en filtro para verlos", 
					ts.Minutes, ts.Seconds, iErrores), "Informacion");
				spnIdSubida.EditValue = idSubida;
			}
			else
			{
				MessageBox.Show(String.Format("Archivo Subido Perfectamente {0}:{1} [seg]", 
					ts.Minutes, ts.Seconds), "Informacion",
					MessageBoxButtons.OK, MessageBoxIcon.Information);
			}
			#endregion Fin de mosrtrar errores 
		}

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirConta);	
		}
	}
}

