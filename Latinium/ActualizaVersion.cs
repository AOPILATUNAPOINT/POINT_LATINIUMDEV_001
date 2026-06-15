using System;
using System.Data.SqlClient;
using System.IO;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de ActualizaVersion.
	/// </summary>
	public class ActualizaVersion : DevExpress.XtraEditors.XtraForm
	{
		private DevExpress.XtraEditors.SimpleButton btActualiza;
		private DevExpress.XtraEditors.SimpleButton btCopiar;
		private Infragistics.Win.UltraWinGrid.UltraGrid grdMaquinas;
		private DevExpress.XtraEditors.SimpleButton btProcedimiento;
		private DevExpress.XtraEditors.SimpleButton btTrigger;
		private C1.Data.C1DataSet cdsArticulo;
		private C1.Data.C1DataSet cdsMaquina;
		private System.Windows.Forms.HelpProvider helpProvider1;
		private DevExpress.XtraEditors.SimpleButton btReporte;
		private DevExpress.XtraEditors.SimpleButton ultraButton7;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkMostrarEjecucion;
		private DevExpress.XtraEditors.SimpleButton btActualizaTodo;
		private System.Windows.Forms.ToolTip toolTip1;
		private DevExpress.XtraEditors.SimpleButton btCopiaUsuarios;
		private DevExpress.XtraEditors.SimpleButton ultraButton1;
		private DevExpress.XtraEditors.SimpleButton ultraButton2;
		private DevExpress.XtraEditors.SimpleButton ultraButton3;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkPersonalizado;
		private DevExpress.XtraEditors.SimpleButton btGeneraArchivo;
		private DevExpress.XtraEditors.SimpleButton btRespaldo;
		private DevExpress.XtraTab.XtraTabControl xtraTabControl1;
		private DevExpress.XtraTab.XtraTabPage xtraTabPage1;
		private DevExpress.XtraTab.XtraTabPage xtraTabPage2;
		private DevExpress.XtraTab.XtraTabPage xtraTabPage3;
		private DevExpress.XtraEditors.PanelControl panelControl1;
		private DevExpress.XtraEditors.PanelControl panelControl2;
		private System.Windows.Forms.GroupBox groupBox1;
		private DevExpress.XtraEditors.CheckEdit chkRespaldo;
		private DevExpress.XtraGrid.Views.Grid.GridView grvProcedimientos;
		private DevExpress.XtraGrid.Views.Grid.GridView gridView4;
		private DevExpress.XtraGrid.GridControl grdProcedimiento;
		private DevExpress.XtraGrid.Views.Grid.GridView gridView1;
		private DevExpress.XtraGrid.Views.Grid.GridView gridView2;
		private DevExpress.XtraGrid.GridControl grdMensajes;
		private System.ComponentModel.IContainer components;

		public ActualizaVersion()
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
			Infragistics.Win.Appearance appearance1 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Maquina", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("IdMaquina");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Maquina");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Ruta");
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			this.btActualiza = new DevExpress.XtraEditors.SimpleButton();
			this.btCopiar = new DevExpress.XtraEditors.SimpleButton();
			this.grdMaquinas = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.cdsMaquina = new C1.Data.C1DataSet();
			this.btProcedimiento = new DevExpress.XtraEditors.SimpleButton();
			this.btTrigger = new DevExpress.XtraEditors.SimpleButton();
			this.cdsArticulo = new C1.Data.C1DataSet();
			this.helpProvider1 = new System.Windows.Forms.HelpProvider();
			this.btReporte = new DevExpress.XtraEditors.SimpleButton();
			this.ultraButton7 = new DevExpress.XtraEditors.SimpleButton();
			this.chkMostrarEjecucion = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.btActualizaTodo = new DevExpress.XtraEditors.SimpleButton();
			this.btCopiaUsuarios = new DevExpress.XtraEditors.SimpleButton();
			this.toolTip1 = new System.Windows.Forms.ToolTip(this.components);
			this.btGeneraArchivo = new DevExpress.XtraEditors.SimpleButton();
			this.ultraButton1 = new DevExpress.XtraEditors.SimpleButton();
			this.chkPersonalizado = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.ultraButton3 = new DevExpress.XtraEditors.SimpleButton();
			this.ultraButton2 = new DevExpress.XtraEditors.SimpleButton();
			this.btRespaldo = new DevExpress.XtraEditors.SimpleButton();
			this.xtraTabControl1 = new DevExpress.XtraTab.XtraTabControl();
			this.xtraTabPage1 = new DevExpress.XtraTab.XtraTabPage();
			this.xtraTabPage2 = new DevExpress.XtraTab.XtraTabPage();
			this.panelControl1 = new DevExpress.XtraEditors.PanelControl();
			this.xtraTabPage3 = new DevExpress.XtraTab.XtraTabPage();
			this.panelControl2 = new DevExpress.XtraEditors.PanelControl();
			this.groupBox1 = new System.Windows.Forms.GroupBox();
			this.chkRespaldo = new DevExpress.XtraEditors.CheckEdit();
			this.grdProcedimiento = new DevExpress.XtraGrid.GridControl();
			this.grvProcedimientos = new DevExpress.XtraGrid.Views.Grid.GridView();
			this.gridView4 = new DevExpress.XtraGrid.Views.Grid.GridView();
			this.grdMensajes = new DevExpress.XtraGrid.GridControl();
			this.gridView1 = new DevExpress.XtraGrid.Views.Grid.GridView();
			this.gridView2 = new DevExpress.XtraGrid.Views.Grid.GridView();
			((System.ComponentModel.ISupportInitialize)(this.grdMaquinas)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsMaquina)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsArticulo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.xtraTabControl1)).BeginInit();
			this.xtraTabControl1.SuspendLayout();
			this.xtraTabPage1.SuspendLayout();
			this.xtraTabPage2.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.panelControl1)).BeginInit();
			this.panelControl1.SuspendLayout();
			this.xtraTabPage3.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.panelControl2)).BeginInit();
			this.panelControl2.SuspendLayout();
			this.groupBox1.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.chkRespaldo.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.grdProcedimiento)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.grvProcedimientos)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.gridView4)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.grdMensajes)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.gridView1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.gridView2)).BeginInit();
			this.SuspendLayout();
			// 
			// btActualiza
			// 
			this.btActualiza.Location = new System.Drawing.Point(24, 78);
			this.btActualiza.Name = "btActualiza";
			this.btActualiza.Size = new System.Drawing.Size(136, 25);
			this.btActualiza.TabIndex = 0;
			this.btActualiza.Text = "Actualiza Estructura";
			this.btActualiza.Click += new System.EventHandler(this.btActualiza_Click);
			this.btActualiza.MouseUp += new System.Windows.Forms.MouseEventHandler(this.btActualiza_MouseUp);
			// 
			// btCopiar
			// 
			this.btCopiar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btCopiar.Location = new System.Drawing.Point(16, 8);
			this.btCopiar.Name = "btCopiar";
			this.btCopiar.Size = new System.Drawing.Size(112, 25);
			this.btCopiar.TabIndex = 1;
			this.btCopiar.Text = "Copiar Archivos";
			this.toolTip1.SetToolTip(this.btCopiar, "Clic Normal copia máquina seleccionda / Clic Secundario a Todas las máquinas de l" +
				"a lista");
			this.btCopiar.Click += new System.EventHandler(this.btCopiar_Click);
			// 
			// grdMaquinas
			// 
			this.grdMaquinas.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.grdMaquinas.Cursor = System.Windows.Forms.Cursors.Default;
			this.grdMaquinas.DataMember = "Maquina";
			this.grdMaquinas.DataSource = this.cdsMaquina;
			appearance1.BackColor = System.Drawing.Color.White;
			this.grdMaquinas.DisplayLayout.Appearance = appearance1;
			this.grdMaquinas.DisplayLayout.AutoFitColumns = true;
			ultraGridColumn1.Header.VisiblePosition = 2;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn2.Header.Caption = "Máquina";
			ultraGridColumn2.Header.VisiblePosition = 0;
			ultraGridColumn2.Width = 126;
			ultraGridColumn3.Header.VisiblePosition = 1;
			ultraGridColumn3.Width = 365;
			ultraGridBand1.Columns.AddRange(new object[] {
																										 ultraGridColumn1,
																										 ultraGridColumn2,
																										 ultraGridColumn3});
			this.grdMaquinas.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.grdMaquinas.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.TemplateOnBottom;
			this.grdMaquinas.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.True;
			this.grdMaquinas.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance2.BackColor = System.Drawing.Color.Transparent;
			this.grdMaquinas.DisplayLayout.Override.CardAreaAppearance = appearance2;
			appearance3.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance3.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance3.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance3.FontData.BoldAsString = "True";
			appearance3.FontData.Name = "Arial";
			appearance3.FontData.SizeInPoints = 10F;
			appearance3.ForeColor = System.Drawing.Color.White;
			appearance3.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.grdMaquinas.DisplayLayout.Override.HeaderAppearance = appearance3;
			appearance4.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance4.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance4.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.grdMaquinas.DisplayLayout.Override.RowSelectorAppearance = appearance4;
			appearance5.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance5.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance5.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.grdMaquinas.DisplayLayout.Override.SelectedRowAppearance = appearance5;
			this.grdMaquinas.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.grdMaquinas.Location = new System.Drawing.Point(8, 8);
			this.grdMaquinas.Name = "grdMaquinas";
			this.grdMaquinas.Size = new System.Drawing.Size(512, 256);
			this.grdMaquinas.TabIndex = 2;
			// 
			// cdsMaquina
			// 
			this.cdsMaquina.BindingContextCtrl = this;
			this.cdsMaquina.DataLibrary = "LibEmpresa";
			this.cdsMaquina.DataLibraryUrl = "";
			this.cdsMaquina.DataSetDef = "dsMaquina";
			this.cdsMaquina.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsMaquina.Name = "cdsMaquina";
			this.cdsMaquina.SchemaClassName = "LibEmpresa.DataClass";
			this.cdsMaquina.SchemaDef = null;
			// 
			// btProcedimiento
			// 
			this.btProcedimiento.Location = new System.Drawing.Point(24, 112);
			this.btProcedimiento.Name = "btProcedimiento";
			this.btProcedimiento.Size = new System.Drawing.Size(136, 25);
			this.btProcedimiento.TabIndex = 3;
			this.btProcedimiento.Text = "Actualiza Procedimiento";
			this.btProcedimiento.Click += new System.EventHandler(this.btProcedimiento_Click);
			// 
			// btTrigger
			// 
			this.btTrigger.Location = new System.Drawing.Point(24, 146);
			this.btTrigger.Name = "btTrigger";
			this.btTrigger.Size = new System.Drawing.Size(136, 25);
			this.btTrigger.TabIndex = 4;
			this.btTrigger.Text = "Actualiza Trigger";
			this.btTrigger.Click += new System.EventHandler(this.btTrigger_Click);
			this.btTrigger.MouseUp += new System.Windows.Forms.MouseEventHandler(this.btTrigger_MouseUp);
			// 
			// cdsArticulo
			// 
			this.cdsArticulo.BindingContextCtrl = this;
			this.cdsArticulo.DataLibrary = "LibFacturacion";
			this.cdsArticulo.DataLibraryUrl = "";
			this.cdsArticulo.DataSetDef = "dsRetazo";
			this.cdsArticulo.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsArticulo.Name = "cdsArticulo";
			this.cdsArticulo.SchemaClassName = "LibFacturacion.DataClass";
			this.cdsArticulo.SchemaDef = null;
			this.cdsArticulo.BeforeFill += new C1.Data.FillEventHandler(this.cdsArticulo_BeforeFill);
			// 
			// helpProvider1
			// 
			this.helpProvider1.HelpNamespace = "C:\\Latinium\\Latinium.chm";
			// 
			// btReporte
			// 
			this.btReporte.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btReporte.Location = new System.Drawing.Point(136, 8);
			this.btReporte.Name = "btReporte";
			this.btReporte.Size = new System.Drawing.Size(112, 25);
			this.btReporte.TabIndex = 7;
			this.btReporte.Text = "Copiar Reportes";
			this.btReporte.Click += new System.EventHandler(this.btReporte_Click);
			// 
			// ultraButton7
			// 
			this.ultraButton7.Location = new System.Drawing.Point(24, 78);
			this.ultraButton7.Name = "ultraButton7";
			this.ultraButton7.Size = new System.Drawing.Size(136, 25);
			this.ultraButton7.TabIndex = 0;
			this.ultraButton7.Text = "Actualiza Estructura";
			// 
			// chkMostrarEjecucion
			// 
			this.chkMostrarEjecucion.Location = new System.Drawing.Point(24, 9);
			this.chkMostrarEjecucion.Name = "chkMostrarEjecucion";
			this.chkMostrarEjecucion.Size = new System.Drawing.Size(128, 21);
			this.chkMostrarEjecucion.TabIndex = 11;
			this.chkMostrarEjecucion.Text = "Mostrar Ejecución";
			// 
			// btActualizaTodo
			// 
			this.btActualizaTodo.Location = new System.Drawing.Point(8, 16);
			this.btActualizaTodo.Name = "btActualizaTodo";
			this.btActualizaTodo.Size = new System.Drawing.Size(120, 25);
			this.btActualizaTodo.TabIndex = 12;
			this.btActualizaTodo.Text = "Actualiza Todo";
			this.btActualizaTodo.Click += new System.EventHandler(this.btActualizaTodo_Click);
			// 
			// btCopiaUsuarios
			// 
			this.btCopiaUsuarios.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btCopiaUsuarios.Location = new System.Drawing.Point(256, 8);
			this.btCopiaUsuarios.Name = "btCopiaUsuarios";
			this.btCopiaUsuarios.Size = new System.Drawing.Size(112, 25);
			this.btCopiaUsuarios.TabIndex = 15;
			this.btCopiaUsuarios.Text = "Copiar Usuarios";
			this.btCopiaUsuarios.Click += new System.EventHandler(this.btCopiaUsuarios_Click);
			// 
			// btGeneraArchivo
			// 
			this.btGeneraArchivo.Location = new System.Drawing.Point(24, 319);
			this.btGeneraArchivo.Name = "btGeneraArchivo";
			this.btGeneraArchivo.Size = new System.Drawing.Size(136, 25);
			this.btGeneraArchivo.TabIndex = 17;
			this.btGeneraArchivo.Text = "Genera Archivo Bat";
			this.toolTip1.SetToolTip(this.btGeneraArchivo, "Archivo Bat que permite actualizar con facilidad la versión");
			this.btGeneraArchivo.Click += new System.EventHandler(this.btGeneraArchivo_Click);
			// 
			// ultraButton1
			// 
			this.ultraButton1.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.ultraButton1.Location = new System.Drawing.Point(384, 8);
			this.ultraButton1.Name = "ultraButton1";
			this.ultraButton1.Size = new System.Drawing.Size(112, 25);
			this.ultraButton1.TabIndex = 16;
			this.ultraButton1.Text = "Grabar Máquinas";
			this.ultraButton1.Click += new System.EventHandler(this.ultraButton1_Click);
			// 
			// chkPersonalizado
			// 
			this.chkPersonalizado.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.chkPersonalizado.Location = new System.Drawing.Point(16, 8);
			this.chkPersonalizado.Name = "chkPersonalizado";
			this.chkPersonalizado.Size = new System.Drawing.Size(128, 22);
			this.chkPersonalizado.TabIndex = 16;
			this.chkPersonalizado.Text = "Solo Personalizados";
			// 
			// ultraButton3
			// 
			this.ultraButton3.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.ultraButton3.Location = new System.Drawing.Point(368, 8);
			this.ultraButton3.Name = "ultraButton3";
			this.ultraButton3.Size = new System.Drawing.Size(136, 25);
			this.ultraButton3.TabIndex = 15;
			this.ultraButton3.Text = "Grabar Procedimientos";
			this.ultraButton3.Click += new System.EventHandler(this.ultraButton3_Click);
			// 
			// ultraButton2
			// 
			this.ultraButton2.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.ultraButton2.Location = new System.Drawing.Point(168, 8);
			this.ultraButton2.Name = "ultraButton2";
			this.ultraButton2.Size = new System.Drawing.Size(136, 25);
			this.ultraButton2.TabIndex = 13;
			this.ultraButton2.Text = "Leer Procedimientos";
			this.ultraButton2.Click += new System.EventHandler(this.ultraButton2_Click);
			// 
			// btRespaldo
			// 
			this.btRespaldo.Location = new System.Drawing.Point(24, 43);
			this.btRespaldo.Name = "btRespaldo";
			this.btRespaldo.Size = new System.Drawing.Size(136, 26);
			this.btRespaldo.TabIndex = 18;
			this.btRespaldo.Text = "Respaldo Base Datos";
			this.btRespaldo.Click += new System.EventHandler(this.btRespaldo_Click);
			// 
			// xtraTabControl1
			// 
			this.xtraTabControl1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.xtraTabControl1.Location = new System.Drawing.Point(208, 16);
			this.xtraTabControl1.Name = "xtraTabControl1";
			this.xtraTabControl1.SelectedTabPage = this.xtraTabPage1;
			this.xtraTabControl1.Size = new System.Drawing.Size(536, 344);
			this.xtraTabControl1.TabIndex = 19;
			this.xtraTabControl1.TabPages.AddRange(new DevExpress.XtraTab.XtraTabPage[] {
																																										this.xtraTabPage1,
																																										this.xtraTabPage2,
																																										this.xtraTabPage3});
			this.xtraTabControl1.Text = "xtraTabControl1";
			// 
			// xtraTabPage1
			// 
			this.xtraTabPage1.Controls.Add(this.grdMensajes);
			this.xtraTabPage1.Name = "xtraTabPage1";
			this.xtraTabPage1.Size = new System.Drawing.Size(527, 313);
			this.xtraTabPage1.Text = "Mensajes y Errores";
			// 
			// xtraTabPage2
			// 
			this.xtraTabPage2.Controls.Add(this.panelControl1);
			this.xtraTabPage2.Controls.Add(this.grdMaquinas);
			this.xtraTabPage2.Name = "xtraTabPage2";
			this.xtraTabPage2.Size = new System.Drawing.Size(527, 313);
			this.xtraTabPage2.Text = "Copiar Archivos";
			// 
			// panelControl1
			// 
			this.panelControl1.Controls.Add(this.btCopiar);
			this.panelControl1.Controls.Add(this.btReporte);
			this.panelControl1.Controls.Add(this.btCopiaUsuarios);
			this.panelControl1.Controls.Add(this.ultraButton1);
			this.panelControl1.Dock = System.Windows.Forms.DockStyle.Bottom;
			this.panelControl1.Location = new System.Drawing.Point(0, 273);
			this.panelControl1.Name = "panelControl1";
			this.panelControl1.Size = new System.Drawing.Size(527, 40);
			this.panelControl1.TabIndex = 0;
			this.panelControl1.Text = "panelControl1";
			// 
			// xtraTabPage3
			// 
			this.xtraTabPage3.Controls.Add(this.grdProcedimiento);
			this.xtraTabPage3.Controls.Add(this.panelControl2);
			this.xtraTabPage3.Name = "xtraTabPage3";
			this.xtraTabPage3.Size = new System.Drawing.Size(527, 313);
			this.xtraTabPage3.Text = "Procedimientos";
			// 
			// panelControl2
			// 
			this.panelControl2.Controls.Add(this.chkPersonalizado);
			this.panelControl2.Controls.Add(this.ultraButton2);
			this.panelControl2.Controls.Add(this.ultraButton3);
			this.panelControl2.Dock = System.Windows.Forms.DockStyle.Bottom;
			this.panelControl2.Location = new System.Drawing.Point(0, 273);
			this.panelControl2.Name = "panelControl2";
			this.panelControl2.Size = new System.Drawing.Size(527, 40);
			this.panelControl2.TabIndex = 0;
			this.panelControl2.Text = "panelControl2";
			// 
			// groupBox1
			// 
			this.groupBox1.Controls.Add(this.chkRespaldo);
			this.groupBox1.Controls.Add(this.btActualizaTodo);
			this.groupBox1.Location = new System.Drawing.Point(24, 184);
			this.groupBox1.Name = "groupBox1";
			this.groupBox1.Size = new System.Drawing.Size(136, 88);
			this.groupBox1.TabIndex = 20;
			this.groupBox1.TabStop = false;
			// 
			// chkRespaldo
			// 
			this.chkRespaldo.EditValue = true;
			this.chkRespaldo.Location = new System.Drawing.Point(8, 56);
			this.chkRespaldo.Name = "chkRespaldo";
			// 
			// chkRespaldo.Properties
			// 
			this.chkRespaldo.Properties.Caption = "Incluir Respaldo";
			this.chkRespaldo.Size = new System.Drawing.Size(120, 19);
			this.chkRespaldo.TabIndex = 13;
			// 
			// grdProcedimiento
			// 
			this.grdProcedimiento.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			// 
			// grdBalances.EmbeddedNavigator
			// 
			this.grdProcedimiento.EmbeddedNavigator.Name = "";
			this.grdProcedimiento.Location = new System.Drawing.Point(8, 8);
			this.grdProcedimiento.MainView = this.grvProcedimientos;
			this.grdProcedimiento.Name = "grdProcedimiento";
			this.grdProcedimiento.Size = new System.Drawing.Size(504, 256);
			this.grdProcedimiento.TabIndex = 18;
			this.grdProcedimiento.ViewCollection.AddRange(new DevExpress.XtraGrid.Views.Base.BaseView[] {
																																																		this.grvProcedimientos,
																																																		this.gridView4});
			// 
			// grvProcedimientos
			// 
			this.grvProcedimientos.GridControl = this.grdProcedimiento;
			this.grvProcedimientos.Name = "grvProcedimientos";
			this.grvProcedimientos.OptionsSelection.MultiSelect = true;
			this.grvProcedimientos.OptionsSelection.MultiSelectMode = DevExpress.XtraGrid.Views.Grid.GridMultiSelectMode.CellSelect;
			this.grvProcedimientos.OptionsView.ShowAutoFilterRow = true;
			// 
			// gridView4
			// 
			this.gridView4.GridControl = this.grdProcedimiento;
			this.gridView4.Name = "gridView4";
			// 
			// grdMensajes
			// 
			this.grdMensajes.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			// 
			// gridControl1.EmbeddedNavigator
			// 
			this.grdMensajes.EmbeddedNavigator.Name = "";
			this.grdMensajes.Location = new System.Drawing.Point(11, 8);
			this.grdMensajes.MainView = this.gridView1;
			this.grdMensajes.Name = "grdMensajes";
			this.grdMensajes.Size = new System.Drawing.Size(504, 296);
			this.grdMensajes.TabIndex = 19;
			this.grdMensajes.ViewCollection.AddRange(new DevExpress.XtraGrid.Views.Base.BaseView[] {
																																															 this.gridView1,
																																															 this.gridView2});
			// 
			// gridView1
			// 
			this.gridView1.GridControl = this.grdMensajes;
			this.gridView1.Name = "gridView1";
			this.gridView1.OptionsSelection.MultiSelect = true;
			this.gridView1.OptionsSelection.MultiSelectMode = DevExpress.XtraGrid.Views.Grid.GridMultiSelectMode.CellSelect;
			this.gridView1.OptionsView.ShowAutoFilterRow = true;
			// 
			// gridView2
			// 
			this.gridView2.GridControl = this.grdMensajes;
			this.gridView2.Name = "gridView2";
			// 
			// ActualizaVersion
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 14);
			this.ClientSize = new System.Drawing.Size(752, 367);
			this.Controls.Add(this.groupBox1);
			this.Controls.Add(this.xtraTabControl1);
			this.Controls.Add(this.btRespaldo);
			this.Controls.Add(this.btGeneraArchivo);
			this.Controls.Add(this.chkMostrarEjecucion);
			this.Controls.Add(this.btTrigger);
			this.Controls.Add(this.btProcedimiento);
			this.Controls.Add(this.btActualiza);
			this.Controls.Add(this.ultraButton7);
			this.helpProvider1.SetHelpKeyword(this, "Actuaversion.htm");
			this.helpProvider1.SetHelpNavigator(this, System.Windows.Forms.HelpNavigator.Topic);
			this.MinimizeBox = false;
			this.Name = "ActualizaVersion";
			this.helpProvider1.SetShowHelp(this, true);
			this.Text = "Actualización de la Versión";
			this.Closing += new System.ComponentModel.CancelEventHandler(this.ActualizaVersion_Closing);
			this.Load += new System.EventHandler(this.ActualizaVersion_Load);
			((System.ComponentModel.ISupportInitialize)(this.grdMaquinas)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsMaquina)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsArticulo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.xtraTabControl1)).EndInit();
			this.xtraTabControl1.ResumeLayout(false);
			this.xtraTabPage1.ResumeLayout(false);
			this.xtraTabPage2.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.panelControl1)).EndInit();
			this.panelControl1.ResumeLayout(false);
			this.xtraTabPage3.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.panelControl2)).EndInit();
			this.panelControl2.ResumeLayout(false);
			this.groupBox1.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.chkRespaldo.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.grdProcedimiento)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.grvProcedimientos)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.gridView4)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.grdMensajes)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.gridView1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.gridView2)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void ProcesaArchivo(string stArchivo, int idTipo)
		{
			string stBase = MenuLatinium.stDirConta;
			Funcion.EjecutaSQL(cdsArticulo, "USE " + stBase);
			string stDir = MenuLatinium.stDirInicio + "\\";
			if (stDir.ToUpper().IndexOf("DEBUG") > 0 || stDir.ToUpper().EndsWith("RELEASE")) 
				stDir = MenuLatinium.stDirInicio + @"\";
			string stOrigen = stDir + stArchivo + ".Sql";
			string stTemp = stDir + stArchivo + ".Qry";
			if (!File.Exists(stOrigen))
			{
				MessageBox.Show("No existe archivo " + stOrigen);
				return;
			}
			string stLinea;
			StreamReader sr = new StreamReader(stOrigen);
			StreamWriter sw = new StreamWriter(stTemp);
			while ((stLinea = sr.ReadLine()) != null) 
			{
				if (idTipo == 1)
					stLinea = stLinea.Replace("LatiniumSQL", MenuLatinium.stDirConta);
				// Manejo de Procedmientos Almacenados
				if (idTipo == 2)
				{
					if (stLinea.ToUpper().Trim().StartsWith("CREATE PROCEDURE"))
						stLinea = stLinea.ToUpper().Replace("CREATE PROCEDURE", "ALTER PROCEDURE");
				}
				// Manejo de trigers
				if (idTipo == 3)
				{
					stLinea = stLinea.ToUpper().Replace("CREATE TRIGGER", "ALTER TRIGGER");
				}
				if (idTipo == 3)
				{
					string stCompara = stLinea.Trim().ToLower();
					if (stCompara.StartsWith("drop table")) stLinea = "";
					if (stCompara.IndexOf("isusertable") > 0) stLinea = "";
				}
				sw.WriteLine(stLinea);
			}
			sr.Close();
			sw.Close();
			try
			{
				string stParametro = "/U infoelect /P Bl45o6$9 /d ";
				stParametro += stBase;
				stParametro += " /S " + MenuLatinium.stDirServidor;
				stParametro += " /n /i \"" + stTemp + "\" /o \"" + stDir +"salida.txt\"";
				
				System.Diagnostics.Process.Start("osql", stParametro);
			}
			catch(System.Exception ex)
			{
				MessageBox.Show(ex.Message, "Comandos osql");
				return;
			}
			//			File.Delete(stTemp);

			if (DialogResult.Yes == MessageBox.Show(
				"¿Desea Ver Registro de Actividades?", 
				"Información", MessageBoxButtons.YesNo, MessageBoxIcon.Question))
				System.Diagnostics.Process.Start("NotePad.exe", stDir + "Salida.txt");
		}

		private void btActualiza_Click(object sender, System.EventArgs e)
		{
			if (DialogResult.No == MessageBox.Show("¿Desea Actualizar la estructura de las tablas?",
				"Confirmación", MessageBoxButtons.YesNo, MessageBoxIcon.Question)) return;
			
			DateTime dtIni = DateTime.Now;
			Cursor = Cursors.WaitCursor;
			string stSelect = "SELECT Count(*) FROM sys.sql_modules modu INNER JOIN sys.objects obj ON modu.object_id = obj.object_id "
				+ "WHERE obj.type = 'P' and obj.Name = 'ActualizaVersion'";
			int iCuenta = Funcion.iEscalarSQL(cdsArticulo, stSelect);

			stSelect = "SELECT modu.definition FROM bases.sys.sql_modules modu INNER JOIN bases.sys.objects obj ON modu.object_id = obj.object_id "
				+ "WHERE obj.type = 'P' and obj.Name = 'ActualizaVersion'";
			string stProc = Funcion.sEscalarSQL(cdsArticulo, stSelect);
			if (iCuenta == 1) stProc = "Alter " + stProc.Substring(7);
			Funcion.EjecutaSQL(cdsArticulo, stProc, true);

			int iEjecucion = 0;
			if (chkMostrarEjecucion.Checked) iEjecucion = 1;
			Funcion.EjecutaSQL(cdsArticulo, "Exec ActualizaVersion " + iEjecucion);
			grdMensajes.DataSource = Funcion.dvProcedimiento(cdsArticulo, "Select Mensaje, Proceso From ErrorSalida");
			Cursor = Cursors.Default;
			TimeSpan ts = DateTime.Now.Subtract(dtIni);
			MessageBox.Show("Estructura de Base de datos Actualizada en : " + ts.TotalSeconds.ToString("n2") + " seg.", 
				"Información", MessageBoxButtons.OK, MessageBoxIcon.Information);
		}

		private void btSalir_Click(object sender, System.EventArgs e)
		{
			try
			{
				cdsMaquina.Update();
			}
			catch(System.Exception ex)
			{
				MessageBox.Show(ex.Message, "Información");
				return;
			}
			this.Close();
		}

		private void btProcedimiento_Click(object sender, System.EventArgs e)
		{
			if (DialogResult.No == MessageBox.Show("¿Desea Crear Procedimientos Almacenados?",
				"Confirmación", MessageBoxButtons.YesNo, MessageBoxIcon.Question)) return;

			DateTime dtIni = DateTime.Now;
			#region SeteoGArbol
			Cursor = Cursors.WaitCursor;
			string stSelect = "SELECT Count(*) FROM sys.sql_modules modu INNER JOIN sys.objects obj ON modu.object_id = obj.object_id "
				+ "WHERE obj.type = 'P' and obj.Name = 'SeteoGArbol'";
			int iCuenta = Funcion.iEscalarSQL(cdsArticulo, stSelect);
			stSelect = "SELECT modu.definition FROM bases.sys.sql_modules modu INNER JOIN bases.sys.objects obj ON modu.object_id = obj.object_id "
				+ "WHERE obj.type = 'P' and obj.Name = 'SeteoGArbol'";
			string stProc = Funcion.sEscalarSQL(cdsArticulo, stSelect);
			if (iCuenta == 1) stProc = "Alter " + stProc.Substring(7);
			Funcion.EjecutaSQL(cdsArticulo, stProc, false);
			#endregion SeteoGArbol

			#region ActualizaProcedimiento
			stSelect = "SELECT Count(*) FROM sys.sql_modules modu INNER JOIN sys.objects obj ON modu.object_id = obj.object_id "
				+ "WHERE obj.type = 'P' and obj.Name = 'ActualizaProcedimiento'";
			iCuenta = Funcion.iEscalarSQL(cdsArticulo, stSelect);
			stSelect = "SELECT modu.definition FROM bases.sys.sql_modules modu INNER JOIN bases.sys.objects obj ON modu.object_id = obj.object_id "
				+ "WHERE obj.type = 'P' and obj.Name = 'ActualizaProcedimiento'";
			stProc = Funcion.sEscalarSQL(cdsArticulo, stSelect);
			if (iCuenta == 1) stProc = "Alter " + stProc.Substring(7);
			try
			{
				Funcion.EjecutaSQL(cdsArticulo, stProc, true);
			}
			catch (System.Exception ex)
			{
				MessageBox.Show(ex.Message, "Error Copiar Procedimiento");
				return;
			}
			int iEjecucion = 0;
			if (chkMostrarEjecucion.Checked) iEjecucion = 1;
			try
			{
				Funcion.EjecutaSQL(cdsArticulo, "Exec ActualizaProcedimiento " + iEjecucion);
			}
			catch (System.Exception ex)
			{
				MessageBox.Show(ex.Message, "Error Ejecutar Procedimiento");
			}

			grdMensajes.DataSource = Funcion.dvProcedimiento(cdsArticulo, "Select Mensaje, Proceso From ErrorSalida");
			Cursor = Cursors.Default;
			#endregion ActualizaProcedimiento

			TimeSpan ts = DateTime.Now.Subtract(dtIni);
			MessageBox.Show("Procedimientos Almacenados Creados en : " + ts.TotalSeconds.ToString("n2") + " seg.", 
				"Información", MessageBoxButtons.OK, MessageBoxIcon.Information);
		}

		private void btTrigger_Click(object sender, System.EventArgs e)
		{
			if (DialogResult.No == MessageBox.Show("¿Desea Crear Procedimientos Almacenados?",
				"Confirmación", MessageBoxButtons.YesNo, MessageBoxIcon.Question)) return;

			DateTime dtIni = DateTime.Now;

			#region ActualizaTrigger
			int iEjecucion = 10;
			if (chkMostrarEjecucion.Checked) iEjecucion = 11;
			Funcion.EjecutaSQL(cdsArticulo, "Exec ActualizaProcedimiento " + iEjecucion);

			grdMensajes.DataSource = Funcion.dvProcedimiento(cdsArticulo, "Select Mensaje, Proceso From ErrorSalida");
			Cursor = Cursors.Default;
			#endregion ActualizaTrigger	

			TimeSpan ts = DateTime.Now.Subtract(dtIni);
			MessageBox.Show("ActualizaTrigger Creados en : " + ts.TotalSeconds.ToString("n2") + " seg.", 
				"Información", MessageBoxButtons.OK, MessageBoxIcon.Information);
		}

		private void cdsArticulo_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			this.cdsArticulo.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);
		}

		private bool ExistenErrores()
		{
      int iCuenta = Funcion.iEscalarSQL(cdsArticulo, "Select Count(*) From ErrorSalida");
			if (iCuenta > 0)
			{
				grdMensajes.DataSource = Funcion.dvProcedimiento(cdsArticulo, "Select Mensaje, Proceso From ErrorSalida");
				xtraTabControl1.SelectedTabPage = xtraTabPage1;
				Cursor = Cursors.Default;
				MessageBox.Show("Existen errores en el proceso", 
					"Información", MessageBoxButtons.OK, MessageBoxIcon.Information);
				return true;
			}
			return false;
		}

		private void ActualizaVersion_Closing(object sender, System.ComponentModel.CancelEventArgs e)
		{
			cdsMaquina.Update();
			string stAudita = "Exec AuditaCrear 53, 7, 'Act. Version'";
			Funcion.EjecutaSQL(cdsArticulo, stAudita, true);
		}

		private void ActualizaVersion_Load(object sender, System.EventArgs e)
		{
			string stAudita = "Exec AuditaCrear 53, 6, 'Act. Version'";
			Funcion.EjecutaSQL(cdsArticulo, stAudita, true);
		}

		private void btTrigger_MouseUp(object sender, System.Windows.Forms.MouseEventArgs e)
		{
			if (e.Button != MouseButtons.Right) return;
      string stExec = "Select Name from SysObjects Where xtype='TR' And Category <> 2 Order by Name";
			string[] stTriggers = new string[1000];
			try 
			{
				SqlDataReader dr = Funcion.rEscalarSQL(cdsArticulo, stExec, true);
				int i=0;
				while (dr.Read())
					stTriggers[i++] = dr.GetString(0);
				dr.Close();

				StreamWriter sr = new StreamWriter("Trigger2008.sql");

				string stLinea = "";
				string DropTrigger = "IF OBJECT_ID ('AuditaAsientoI', 'TR') IS NOT NULL"
					+ "\n\tDROP TRIGGER AuditaAsientoI\nGo\n\n"
					+ "IF OBJECT_ID ('AuditaCompraNumeroInserta', 'TR') IS NOT NULL"
					+ "\n\tDROP TRIGGER AuditaCompraNumeroInserta\nGo\n\n"
					+ "IF OBJECT_ID ('AUDITANUEVACC', 'TR') IS NOT NULL"
					+ "\n\tDROP TRIGGER AUDITANUEVACC\nGo\n\n"
					+ "IF OBJECT_ID ('AUDITANUEVACUENTA', 'TR') IS NOT NULL"
					+ "\n\tDROP TRIGGER AUDITANUEVACUENTA\nGo\n\n"
					+ "IF OBJECT_ID ('AuditaOrdenMovInserta', 'TR') IS NOT NULL"
					+ "\n\tDROP TRIGGER AuditaOrdenMovInserta \nGo\n\n"
					+ "IF OBJECT_ID ('AUDITAPAGOLOTEINSERTA', 'TR') IS NOT NULL"
					+ "\n\tDROP TRIGGER AUDITAPAGOLOTEINSERTA \nGo\n\n"
					+ "IF OBJECT_ID ('AUDITAPERSONALINSERTA', 'TR') IS NOT NULL"
					+ "\n\tDROP TRIGGER AUDITAPERSONALINSERTA \nGo\n\n"
					+ "IF OBJECT_ID ('AuditaReservaInserta', 'TR') IS NOT NULL"
					+ "\n\tDROP TRIGGER AuditaReservaInserta \nGo\n\n"
					+ "IF OBJECT_ID ('AuditoriaAsientoDetI', 'TR') IS NOT NULL"
					+ "\n\tDROP TRIGGER AuditoriaAsientoDetI \nGo\n\n"
					+ "IF OBJECT_ID ('AUDITORIACLIENTEE', 'TR') IS NOT NULL"
					+ "\n\tDROP TRIGGER AUDITORIACLIENTEE \nGo\n\n"
					+ "IF OBJECT_ID ('AUDITORIACLIENTEGRUPOI', 'TR') IS NOT NULL"
					+ "\n\tDROP TRIGGER AUDITORIACLIENTEGRUPOI \nGo\n\n"
					+ "IF OBJECT_ID ('AUDITORIANUEVO', 'TR') IS NOT NULL"
					+ "\n\tDROP TRIGGER AUDITORIANUEVO \nGo\n\n"
					+ "IF OBJECT_ID ('AUDITORIANUEVO1', 'TR') IS NOT NULL"
					+ "\n\tDROP TRIGGER AUDITORIANUEVO1 \nGo\n\n"
					+ "IF OBJECT_ID ('AUDITORIANUEVOGRUPOART', 'TR') IS NOT NULL"
					+ "\n\tDROP TRIGGER AUDITORIANUEVOGRUPOART \nGo\n\n"
//					+ "IF OBJECT_ID ('AUDITORIABORRA1', 'TR') IS NOT NULL"
//					+ "\n\tDROP TRIGGER AUDITORIABORRA1 \nGo\n\n"
					+ "IF OBJECT_ID ('AUDITORIARolConceptoInserta', 'TR') IS NOT NULL"
					+ "\n\tDROP TRIGGER AUDITORIARolConceptoInserta \nGo\n\n"
					+ "IF OBJECT_ID ('AUDITORIARolPagoInserta', 'TR') IS NOT NULL"
					+ "\n\tDROP TRIGGER AUDITORIARolPagoInserta \nGo\n\n"
					+ "IF OBJECT_ID ('PagoInserta', 'TR') IS NOT NULL"
					+ "\n\tDROP TRIGGER PagoInserta \nGo\n\n"
					+ "IF OBJECT_ID ('RETENCIONINSERTA', 'TR') IS NOT NULL"
					+ "\n\tDROP TRIGGER RETENCIONINSERTA \nGo\n\n"
					+ "IF OBJECT_ID ('RETENCIONIVANUEVO', 'TR') IS NOT NULL"
					+ "\n\tDROP TRIGGER RETENCIONIVANUEVO \nGo\n\n";
	
				sr.Write(DropTrigger);
			
				for(int j=0; j<i; j++)
				{
					stExec = "Exec sp_helptext '" + stTriggers[j] + "'";
					dr = Funcion.rEscalarSQL(cdsArticulo, stExec, true);
//					sr.Write("-- Desarrollado por Infoelect\n");
					int iInicio = 0;
					while (dr.Read())
					{
						stLinea = dr.GetString(0);
						if (iInicio == 0 && stLinea == "\r\n") continue;
						if (iInicio == 0 && stLinea.StartsWith("\n"))
							stLinea = stLinea.Replace("\n", "");

						iInicio=1;
						if (stLinea.ToUpper().IndexOf("CREATE TRIGGER") >= 0)
							stLinea = stLinea.ToUpper().Replace("CREATE TRIGGER", "ALTER TRIGGER");
						sr.Write(stLinea);
					}
					if (stLinea.EndsWith("\r\n")) sr.WriteLine("GO\n");
					else sr.WriteLine("\nGO\n");
					dr.Close();
				}
				sr.Close();
			}
			catch(Exception ex)
			{
				Console.WriteLine(ex.Message, "Error al crear Trigger2008");
			}
		}

		private void btActualiza_MouseUp(object sender, System.Windows.Forms.MouseEventArgs e)
		{
			if (e.Button != MouseButtons.Right) return;

			ProcesaArchivo("ActualizaVersión", 1);
		}

		private void ultraButton4_Click(object sender, System.EventArgs e)
		{
			MessageBox.Show("Click");
		}

		private void ultraButton4_MouseUp(object sender, System.Windows.Forms.MouseEventArgs e)
		{
			if (e.Button != MouseButtons.Right) return;
			MessageBox.Show("Derecho");
		}

		private void btActualizaTodo_Click(object sender, System.EventArgs e)
		{
			#region Verificacion
			string stDir = "c:\\Latinium\\Respaldo\\";
			if (!Directory.Exists(stDir))
			{
				MessageBox.Show("Respaldo Cancelado porque NO existe Directorio: " + stDir,
					"Información", MessageBoxButtons.OK, MessageBoxIcon.Information);
				return;
			}
			
			if (DialogResult.No == MessageBox.Show("¿Desea Actualizar Toda la Base de Datos?",
				"Confirmación", MessageBoxButtons.YesNo, MessageBoxIcon.Question)) return;
			
			DateTime dtIni = DateTime.Now;
			Cursor = Cursors.WaitCursor;
			#endregion Verificacion

			#region Respaldo base de datos
			string stSelect = "SELECT Count(*) FROM sys.sql_modules modu INNER JOIN sys.objects obj ON modu.object_id = obj.object_id "
				+ "WHERE obj.type = 'P' and obj.Name = 'RespaldoDataBase'";
			
			int iCuenta = 0;
			string stProc;
			string stMensaje="";
			if (chkRespaldo.Checked)
			{
				iCuenta = Funcion.iEscalarSQL(cdsArticulo, stSelect);

				stSelect = "SELECT modu.definition FROM bases.sys.sql_modules modu INNER JOIN bases.sys.objects obj ON modu.object_id = obj.object_id "
					+ "WHERE obj.type = 'P' and obj.Name = 'RespaldoDataBase'";
				stProc = Funcion.sEscalarSQL(cdsArticulo, stSelect);
				if (iCuenta == 1) stProc = "Alter " + stProc.Substring(7);
				Funcion.EjecutaSQL(cdsArticulo, stProc, true);

				string stBase = MenuLatinium.stDirFacturacion; 
				DateTime dtHoy = DateTime.Now;
				stDir += stBase + "_" + dtHoy.ToString("dd-MMM-yyyy_HH_mm") + ".Bak";
				if (File.Exists(stDir))
				{
					if (DialogResult.No == MessageBox.Show("El Archivo: " + stDir
						+ " ya existe.\n¿Desea Sobrescribirlo Respaldo?",
						"Confirmación", MessageBoxButtons.YesNo, MessageBoxIcon.Warning)) 
					{

						// Respaldo de la base de datos
						stSelect = string.Format("Exec RespaldoDataBase '{0}', '{1}'", 
							MenuLatinium.stDirFacturacion, stDir);
						stMensaje = Funcion.sEscalarSQL(cdsArticulo, stSelect);
						if (stMensaje.Length > 0)
						{
							MessageBox.Show("No se puede obtener respaldo. \n"+stMensaje, "Informacion",
								MessageBoxButtons.OK, MessageBoxIcon.Warning);
							return;
						}
					}
				}
			}
			#endregion Respaldo base de datos

			#region Procedimientos ActualizaVersion y ActualizaProcedimiento Copiados
			stSelect = "SELECT Count(*) FROM sys.sql_modules modu INNER JOIN sys.objects obj ON modu.object_id = obj.object_id "
				+ "WHERE obj.type = 'P' and obj.Name = 'ActualizaVersion'";
			iCuenta = Funcion.iEscalarSQL(cdsArticulo, stSelect);

			stSelect = "SELECT modu.definition FROM bases.sys.sql_modules modu INNER JOIN bases.sys.objects obj ON modu.object_id = obj.object_id "
				+ "WHERE obj.type = 'P' and obj.Name = 'ActualizaVersion'";
			stProc = Funcion.sEscalarSQL(cdsArticulo, stSelect);
			if (iCuenta == 1) stProc = "Alter " + stProc.Substring(7);
			Funcion.EjecutaSQL(cdsArticulo, stProc, true);

			stSelect = "SELECT Count(*) FROM sys.sql_modules modu INNER JOIN sys.objects obj ON modu.object_id = obj.object_id "
				+ "WHERE obj.type = 'P' and obj.Name = 'ActualizaProcedimiento'";
			iCuenta = Funcion.iEscalarSQL(cdsArticulo, stSelect);
			stSelect = "SELECT modu.definition FROM bases.sys.sql_modules modu INNER JOIN bases.sys.objects obj ON modu.object_id = obj.object_id "
				+ "WHERE obj.type = 'P' and obj.Name = 'ActualizaProcedimiento'";
			stProc = Funcion.sEscalarSQL(cdsArticulo, stSelect);
			if (iCuenta == 1) stProc = "Alter " + stProc.Substring(7);
			try
			{
				Funcion.EjecutaSQL(cdsArticulo, stProc, true);
			}
			catch (System.Exception ex)
			{
				MessageBox.Show(ex.Message, "Error Copiar Procedimiento");
				return;
			}
			#endregion

			#region ejecucion
			stMensaje = " seg.";
			int iEjecucion = 0;
			if (chkMostrarEjecucion.Checked) iEjecucion = 1;

			Funcion.EjecutaSQL(cdsArticulo, string.Format("Exec ActualizaVersion {0}", iEjecucion));
			iCuenta = Funcion.iEscalarSQL(cdsArticulo, "Select Count(*) From ErrorSalida Where Proceso Like 'Error%'");
			if (iCuenta > 0) stMensaje += string.Format("\nExisten {0} Errores en Estructura", iCuenta);

			Funcion.EjecutaSQL(cdsArticulo, string.Format("Exec ActualizaProcedimiento {0}", iEjecucion));
			iCuenta = Funcion.iEscalarSQL(cdsArticulo, "Select Count(*) From ErrorSalida Where Proceso Like 'Error%'");
			if (iCuenta > 0) stMensaje += string.Format("\nExisten {0} Errores en Procedimientos", iCuenta);

			Funcion.EjecutaSQL(cdsArticulo, string.Format("Exec ActualizaProcedimiento {0}", iEjecucion + 10));
			iCuenta = Funcion.iEscalarSQL(cdsArticulo, "Select Count(*) From ErrorSalida Where Proceso Like 'Error%'");
			if (iCuenta > 0) stMensaje += string.Format("\nExisten {0} Errores en Triggers", iCuenta);
			
			Cursor = Cursors.Default;
			TimeSpan ts = DateTime.Now.Subtract(dtIni);
			System.Windows.Forms.MessageBoxIcon msgIcono = MessageBoxIcon.Information;
			if (stMensaje.Length > 5) msgIcono = MessageBoxIcon.Warning;
			#endregion ejecucion
			MessageBox.Show("Base de datos Actualizada en : " + ts.TotalSeconds.ToString("n2") + stMensaje, 
				"Información", MessageBoxButtons.OK, msgIcono);		
		}

		private void btCopiaUsuarios_Click(object sender, System.EventArgs e)
		{
			CopiarArchivos(MenuLatinium.stDirInicio, "", "Empresa.mdb", "Archivo de usuarios");
		}

		private void VerificaOrigen(string stOrigen)
		{
			if (!Directory.Exists(stOrigen))
			{
				string stSelect = string.Format("Insert Into ErrorSalida (Mensaje, Proceso) Values ('No existe Directorio: {0}', 'Verifica Directorio de Origen'",
					stOrigen);
				Funcion.EjecutaSQL(cdsArticulo, stSelect, true);
			}
		}
	
		private void VerificaDestinos()
		{
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in grdMaquinas.Selected.Rows)
			{
				string stDestino = (string) dr.Cells["Ruta"].Value;
				if (!stDestino.EndsWith(@"\")) stDestino += @"\";
				if (!Directory.Exists(stDestino))
				{
					string stSelect = string.Format("Insert Into ErrorSalida (Mensaje, Proceso) Values ('No existe Directorio: {0}', 'Verificación de directorio')",
						stDestino);
					Funcion.EjecutaSQL(cdsArticulo, stSelect);
				}
			}
		}

		private void CopiarArchivos(string stOrigen, string CarpetaDestino, string stArchivoOr, string stMensaje)
		{
			cdsMaquina.Update();
			string stSelect = "Delete ErrorSalida";
			Funcion.EjecutaSQL(cdsArticulo, stSelect, true);
			if (grdMaquinas.Selected.Rows.Count == 0) grdMaquinas.ActiveRow.Selected = true;

			VerificaOrigen(stOrigen);
			VerificaDestinos();
			if (ExistenErrores()) return;

			string stPlural = "s";
			if (grdMaquinas.Selected.Rows.Count == 1) stPlural = "";
			if (DialogResult.No == MessageBox.Show(string.Format("¿Desea copiar {0} a {1} máquina{2}?", 
				stMensaje, grdMaquinas.Selected.Rows.Count, stPlural), "Confirmación", MessageBoxButtons.YesNo, MessageBoxIcon.Question)) return;
			
			DateTime dtIni = DateTime.Now;
			Cursor = Cursors.WaitCursor;

			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in grdMaquinas.Selected.Rows)
			{
				string stDestino = (string) dr.Cells["Ruta"].Value;
				if (!stDestino.EndsWith(@"\")) stDestino += @"\";
				stDestino += CarpetaDestino;
				try
				{
					if (!Directory.Exists(stDestino)) Directory.CreateDirectory(stDestino);
				}
				catch(System.Exception ex)
				{
					stSelect = string.Format("Insert Into ErrorSalida (Mensaje, Proceso) Values ('No es posible crear Directorio: {0} : {2}', 'Crear Directorio')",
						stDestino, ex.Message);
					Funcion.EjecutaSQL(cdsArticulo, stSelect, true);
					break;
				}

				string [] files = Directory.GetFiles(stOrigen);
				if (stArchivoOr.Length > 0) files[0] = stOrigen + "\\"+ stArchivoOr;
				foreach(string stArchivo in files)
				{
					string stArch = stArchivo;
					string stDest = stDestino + stArchivo.Remove(0, stOrigen.Length);

					try
					{
						File.Copy(stArch, stDest, true);
					}
					catch(System.Exception ex)
					{
						stSelect = string.Format("Insert Into ErrorSalida (Mensaje, Proceso) Values ('No es posible copiar: {0} a: {1} Error: {2}', 'Copia Archivo')",
							stArch, stDest, ex.Message);
						Funcion.EjecutaSQL(cdsArticulo, stSelect, true);
						break;
					}
				}
			}
			if (ExistenErrores()) return; // Errores en copia de archivos

			Cursor = Cursors.Default;
			TimeSpan ts = DateTime.Now.Subtract(dtIni);
			MessageBox.Show("Copia de Archivos Terminada en : " + ts.TotalSeconds.ToString("n2") + " seg.", 
				"Información de " + stMensaje, MessageBoxButtons.OK, MessageBoxIcon.Information);
		}

		private void btCopiar_Click(object sender, System.EventArgs e)
		{
			CopiarArchivos(MenuLatinium.stDirInicio + @"\ActualizaLatinium\", "", "", "la actualización");
		}

		private void btReporte_Click(object sender, System.EventArgs e)
		{
			CopiarArchivos(MenuLatinium.stDirInicio + @"\Reportes\", "Reportes\\", "", "los reportes");
		}

		private void ultraButton1_Click(object sender, System.EventArgs e)
		{
			cdsMaquina.Update();
			MessageBox.Show("Máquinas Grabadas", 
				"Información", MessageBoxButtons.OK, MessageBoxIcon.Information);
		}

		private void ultraButton2_Click(object sender, System.EventArgs e)
		{
			string stSelect = "Insert Into Procedimientos (Nombre, Personalizado, Tipo) "
				+ "Select s.Name, 0, s.xtype from SysObjects s Left Join Procedimientos p On s.Name = p.Nombre "
				+ "Where s.xtype In ('P', 'TR', 'FN', 'TF') And s.Category <> 2 "
				+ "And s.name Not Like 'sp_%' And p.Nombre Is Null ";
			Funcion.EjecutaSQL(cdsArticulo, stSelect, true);
			string stFiltro = "";
			if (chkPersonalizado.Checked) stFiltro = "Where Personalizado = 1 ";
			stSelect = "Select Nombre, Personalizado, Detalle, idProcedimiento From Procedimientos "
				+ stFiltro + "Order by Nombre";
			grvProcedimientos.Columns.Clear();
			grdProcedimiento.DataSource = Funcion.dvProcedimiento(cdsArticulo, stSelect);
			DevExpress.XtraGrid.Views.Base.ColumnView View = 
				grdProcedimiento.MainView as DevExpress.XtraGrid.Views.Base.ColumnView;

			foreach (DevExpress.XtraGrid.Columns.GridColumn col in View.Columns)
			{
				if (col.FieldName == "Nombre") col.OptionsColumn.AllowEdit = false;
				if (col.FieldName.StartsWith("id"))
					col.Visible = false;
			}

			MessageBox.Show("Procedimientos Leidos", 
				"Información", MessageBoxButtons.OK, MessageBoxIcon.Information);
		}

		private void ultraButton3_Click(object sender, System.EventArgs e)
		{
			for (int i=0; i< grvProcedimientos.RowCount; i++)
			{
				int idProc = (int) grvProcedimientos.GetRowCellValue(i, "idProcedimiento");
        int iPersonalizado = 0;
				if ((bool) grvProcedimientos.GetRowCellValue(i, "Personalizado") == true) iPersonalizado = 1;
				string stDetalle = grvProcedimientos.GetRowCellValue(i, "Detalle").ToString();
				string stSelect = string.Format("Update Procedimientos Set Personalizado = {0}, Detalle = '{1}' Where idProcedimiento = {2} And (Personalizado != {0} Or IsNull(Detalle, '') != '{1}')",
					iPersonalizado, stDetalle, idProc);
				Funcion.EjecutaSQL(cdsArticulo, stSelect, true);
			}

			MessageBox.Show("Procedimientos Grabados", 
				"Información", MessageBoxButtons.OK, MessageBoxIcon.Information);
		}

		private void btGeneraArchivo_Click(object sender, System.EventArgs e)
		{
			if (DialogResult.No == MessageBox.Show("¿Desea Generar archivo ActualizaGenera.Bat para actualizar la versión?",
				"Confirmación", MessageBoxButtons.YesNo, MessageBoxIcon.Question)) return;

			string stDir = MenuLatinium.stDirInicio + "\\";
			if (stDir.ToUpper().IndexOf("DEBUG") > 0 || stDir.ToUpper().EndsWith("RELEASE")) 
				stDir = MenuLatinium.stDirInicio + @"\";
			string stOrigen = stDir + "ActualizaGenera.bat";
			if (!File.Exists(stOrigen))
			{
				MessageBox.Show("No existe archivo " + stOrigen);
				return;
			}
			
			string stLinea;
			StreamReader sr = new StreamReader(stOrigen);
			StreamWriter sw = new StreamWriter(stOrigen+".bak");
			while ((stLinea = sr.ReadLine()) != null) 
			{
				stLinea = stLinea.Replace("-S. ", "-S" +MenuLatinium.stDirServidor + " ");
				sw.WriteLine(stLinea);
			}
			sr.Close();
			sw.Close();
			File.Copy(stOrigen + ".bak", stOrigen, true);
			File.Delete(stOrigen + ".bak");
			MessageBox.Show("Archivo ActualizaGenera.Bat actualizado", 
				"Información", MessageBoxButtons.OK, MessageBoxIcon.Information);
		}

		private void btRespaldo_Click(object sender, System.EventArgs e)
		{
			string stDir = "c:\\Latinium\\Respaldo\\";
			if (!Directory.Exists(stDir))
			{
				try
				{
					Directory.CreateDirectory(stDir);
				}
				catch
				{
					MessageBox.Show("Respaldo Cancelado porque NO se puede crear Directorio: " + stDir,
						"Información", MessageBoxButtons.OK, MessageBoxIcon.Information);
					return;
				}
			}
			DateTime dtIni = DateTime.Now;
			string stSelect = "SELECT Count(*) FROM sys.sql_modules modu INNER JOIN sys.objects obj ON modu.object_id = obj.object_id "
				+ "WHERE obj.type = 'P' and obj.Name = 'RespaldoDataBase'";
			int iCuenta = Funcion.iEscalarSQL(cdsArticulo, stSelect);

			stSelect = "SELECT modu.definition FROM bases.sys.sql_modules modu INNER JOIN bases.sys.objects obj ON modu.object_id = obj.object_id "
				+ "WHERE obj.type = 'P' and obj.Name = 'RespaldoDataBase'";
			string stProc = Funcion.sEscalarSQL(cdsArticulo, stSelect);
			if (iCuenta == 1) stProc = "Alter " + stProc.Substring(7);
			Funcion.EjecutaSQL(cdsArticulo, stProc, true);

			string stBase = MenuLatinium.stDirFacturacion; 
			DateTime dtHoy = DateTime.Now;
			stDir += stBase + "_" + dtHoy.ToString("dd-MMM-yyyy_HH_mm") + ".Bak";
			if (DialogResult.No == MessageBox.Show("¿Desea Generar Respaldo en "
				+ stDir + "?", "Confirmación", MessageBoxButtons.YesNo,
				MessageBoxIcon.Question)) return;
			if (File.Exists(stDir))
			{
				if (DialogResult.No == MessageBox.Show("El Archivo: " + stDir
					+ " ya existe.\n¿Desea Sobrescribirlo?",
					"Confirmación", MessageBoxButtons.YesNo, MessageBoxIcon.Warning)) return;
			}

			Cursor = Cursors.WaitCursor;

			// Respaldo de la base de datos
			stSelect = string.Format("Exec RespaldoDataBase '{0}', '{1}'", 
				MenuLatinium.stDirFacturacion, stDir);
			string stMensaje = Funcion.sEscalarSQL(cdsArticulo, stSelect);
			if (stMensaje.Length > 0)
			{
				MessageBox.Show("No se puede obtener respaldo. "+stMensaje, "Informacion",
					MessageBoxButtons.OK, MessageBoxIcon.Warning);
				return;
			}
			Cursor = Cursors.Default;
			TimeSpan ts = DateTime.Now.Subtract(dtIni);
			MessageBox.Show("Respaldo " + stDir + "\nGenerado en : " + ts.TotalSeconds.ToString("n2") + " seg.", 
				"Información", MessageBoxButtons.OK, MessageBoxIcon.Information);
		}

	}
}
